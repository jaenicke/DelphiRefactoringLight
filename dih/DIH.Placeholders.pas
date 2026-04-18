(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Placeholders;

interface

uses
  System.SysUtils, System.Generics.Collections, DIH.Types;

type
  TDIHPlaceholderResolver = class
  private
    FValues: TDictionary<string, string>;
    FBaseDir: string;
    FBDSVersion: string;
    FBDSProfile: string;
    FBDSCommonDir: string;
    FBDSRootDir: string;
    function ExpandEnvironmentVars(const AValue: string): string;
    function ConvertDollarParenToPercent(const AValue: string): string;
    procedure ReadBDSDirsFromRegistry;
  public
    constructor Create(const ABaseDir, ABDSVersion, ABDSProfile: string);
    destructor Destroy; override;
    procedure SetValue(const AName, AValue: string);
    procedure SetDefaults(APlatform: TDIHPlatform; const ABuildConfig: string);
    procedure ApplyCustomValues(ACustom: TDictionary<string, string>);
    function Resolve(const AValue: string): string;
    function ResolveKeepEnvVars(const AValue: string): string;
    function GetRsVarsPath: string;
    property Values: TDictionary<string, string> read FValues;
    property BDSVersion: string read FBDSVersion;
    property BDSProfile: string read FBDSProfile;
    property BDSCommonDir: string read FBDSCommonDir;
    property BDSRootDir: string read FBDSRootDir;
  end;

implementation

uses
  Winapi.Windows, System.Win.Registry;

{ TDIHPlaceholderResolver }

constructor TDIHPlaceholderResolver.Create(const ABaseDir, ABDSVersion, ABDSProfile: string);
begin
  inherited Create;
  FValues := TDictionary<string, string>.Create;
  FBaseDir := ABaseDir;
  FBDSVersion := ABDSVersion;
  FBDSProfile := ABDSProfile;

  // Read BDS dirs from registry - profile determines the registry key name
  ReadBDSDirsFromRegistry;

  FValues.AddOrSetValue('BDSVersion', FBDSVersion);
  FValues.AddOrSetValue('BDSProfileName', FBDSProfile);
  FValues.AddOrSetValue('BDS', 'Software\Embarcadero\' + FBDSProfile + '\' + FBDSVersion);
end;

destructor TDIHPlaceholderResolver.Destroy;
begin
  FValues.Free;
  inherited;
end;

procedure TDIHPlaceholderResolver.SetValue(const AName, AValue: string);
begin
  FValues.AddOrSetValue(AName, AValue);
end;

procedure TDIHPlaceholderResolver.ReadBDSDirsFromRegistry;
var
  Reg: TRegistry;
  ProfileKey, BdsKey: string;
begin
  FBDSRootDir := '';
  FBDSCommonDir := '';
  // Profile determines the registry key: Software\Embarcadero\<Profile>\<Version>
  ProfileKey := 'Software\Embarcadero\' + FBDSProfile + '\' + FBDSVersion;
  // RootDir and CommonDir are always under the default BDS key
  BdsKey := 'Software\Embarcadero\BDS\' + FBDSVersion;

  Reg := TRegistry.Create(KEY_READ);
  try
    // Read RootDir from the profile key first, fall back to default BDS key
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(ProfileKey) then
    begin
      try
        FBDSRootDir := Reg.ReadString('RootDir');
      finally
        Reg.CloseKey;
      end;
    end;
    if FBDSRootDir.IsEmpty and (ProfileKey <> BdsKey) then
    begin
      if Reg.OpenKeyReadOnly(BdsKey) then
      begin
        try
          FBDSRootDir := Reg.ReadString('RootDir');
        finally
          Reg.CloseKey;
        end;
      end;
    end;

    // BDSCOMMONDIR is typically under HKLM\...\BDS\<Version>
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly(BdsKey) then
    begin
      try
        FBDSCommonDir := Reg.ReadString('CommonDir');
        if FBDSRootDir.IsEmpty then
          FBDSRootDir := Reg.ReadString('RootDir');
      finally
        Reg.CloseKey;
      end;
    end;
  finally
    Reg.Free;
  end;

  // Fallback: construct BDSCOMMONDIR from known pattern
  if FBDSCommonDir.IsEmpty then
    FBDSCommonDir := ExpandEnvironmentVars('%PUBLIC%') + '\Documents\Embarcadero\Studio\' + FBDSVersion;
end;

procedure TDIHPlaceholderResolver.SetDefaults(APlatform: TDIHPlatform; const ABuildConfig: string);
var
  PlatStr, DcuDir, BplDir, DcpDir, PlatSubDir: string;
begin
  PlatStr := APlatform.ToString;

  // Win32 uses Bpl/ and Dcp/ directly, Win64 uses Bpl/Win64/ and Dcp/Win64/
  if APlatform = dpWin32 then
    PlatSubDir := ''
  else
    PlatSubDir := PathDelim + PlatStr;

  DcuDir := IncludeTrailingPathDelimiter(FBaseDir) + 'DCU' + PathDelim + PlatStr + PathDelim + ABuildConfig;
  BplDir := FBDSCommonDir + PathDelim + 'Bpl' + PlatSubDir;
  DcpDir := FBDSCommonDir + PathDelim + 'Dcp' + PlatSubDir;

  FValues.AddOrSetValue('DcuTargetDir', DcuDir);
  FValues.AddOrSetValue('BplTargetDir', BplDir);
  FValues.AddOrSetValue('DcpTargetDir', DcpDir);
  FValues.AddOrSetValue('BDSCommonDir', FBDSCommonDir);
  FValues.AddOrSetValue('BDSRootDir', FBDSRootDir);
  FValues.AddOrSetValue('Platform', PlatStr);
  FValues.AddOrSetValue('Config', ABuildConfig);
  FValues.AddOrSetValue('BaseDir', FBaseDir);
end;

procedure TDIHPlaceholderResolver.ApplyCustomValues(ACustom: TDictionary<string, string>);
var
  Pair: TPair<string, string>;
begin
  for Pair in ACustom do
    FValues.AddOrSetValue(Pair.Key, Pair.Value);
end;

function TDIHPlaceholderResolver.GetRsVarsPath: string;
begin
  Result := IncludeTrailingPathDelimiter(FBDSRootDir) + 'bin' + PathDelim + 'rsvars.bat';
end;

function TDIHPlaceholderResolver.ExpandEnvironmentVars(const AValue: string): string;
var
  Buffer: array[0..4095] of Char;
  Len: DWORD;
begin
  Len := ExpandEnvironmentStrings(PChar(AValue), @Buffer[0], Length(Buffer));
  if Len > 0 then
    Result := String(Buffer)
  else
    Result := AValue;
end;

function TDIHPlaceholderResolver.ConvertDollarParenToPercent(const AValue: string): string;
var
  I, StartPos: Integer;
  VarName: string;
begin
  // Properly convert $(VAR) to %VAR% without destroying other parentheses
  Result := '';
  I := 1;
  while I <= Length(AValue) do
  begin
    if (I < Length(AValue)) and (AValue[I] = '$') and (AValue[I + 1] = '(') then
    begin
      // Found $( - look for matching )
      StartPos := I + 2;
      I := StartPos;
      while (I <= Length(AValue)) and (AValue[I] <> ')') do
        Inc(I);
      if I <= Length(AValue) then
      begin
        VarName := Copy(AValue, StartPos, I - StartPos);
        Result := Result + '%' + VarName + '%';
        Inc(I); // skip )
      end
      else
      begin
        // No matching ) found, keep original
        Result := Result + '$(';
        I := StartPos;
      end;
    end
    else
    begin
      Result := Result + AValue[I];
      Inc(I);
    end;
  end;
end;

function TDIHPlaceholderResolver.ResolveKeepEnvVars(const AValue: string): string;
var
  Pair: TPair<string, string>;
begin
  Result := AValue;

  // Only resolve {#...} placeholders, keep $(...) intact
  for Pair in FValues do
    Result := Result.Replace('{#' + Pair.Key + '}', Pair.Value, [rfReplaceAll, rfIgnoreCase]);
end;

function TDIHPlaceholderResolver.Resolve(const AValue: string): string;
var
  Pair: TPair<string, string>;
begin
  // First resolve {#...} placeholders
  Result := ResolveKeepEnvVars(AValue);

  // Resolve $(...) from known values (BDS-specific vars not in env)
  for Pair in FValues do
    Result := Result.Replace('$(' + Pair.Key + ')', Pair.Value, [rfReplaceAll, rfIgnoreCase]);

  // Resolve remaining $(...) via Windows environment variables
  if Result.Contains('$(') then
    Result := ExpandEnvironmentVars(ConvertDollarParenToPercent(Result));
end;

end.
