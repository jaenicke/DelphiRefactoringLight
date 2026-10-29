(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
  /// <summary>Can this account write into ADir? Injected so the rule below
  ///  is testable without an unwritable directory to hand.</summary>
  TDIHWriteProbe = reference to function(const ADir: string): Boolean;

/// <summary>Where a build may put its BPL and DCP, and what to say about it.
///  Delphi's installer writes the Documents folder of the INSTALLING account
///  into rsvars.bat and the registry, so on a machine where that was an admin
///  account, every other user's build dies with MSB3191 on a path it cannot
///  create (forum report 2026-10-15). The IDE loads a package by the full
///  path its registration names, so building into a directory of our own and
///  registering THAT is the only way a non-admin can install a package there
///  at all. ANote is empty when nothing is unusual - and never empty when
///  something was substituted, because a silent substitution is worse than
///  the error it replaces.</summary>
procedure DecideBuildOutputRoot(const ACommonDir, ABDSVersion,
  ALocalAppData: string; const ACanWrite: TDIHWriteProbe;
  out ARoot, ANote: string);

/// <summary>What {$LIBSUFFIX AUTO} appends to a package's file name, read
///  from the INSTALLATION rather than from the IDE version: the digits of
///  the RTL package the IDE ships (bin\rtl370.bpl -> '370'). Delphi 13 is
///  the one version where the registry version and the suffix agree
///  ("37.0" -> 370); Delphi 11 registers as 22.0 while its packages carry
///  280, and Delphi 12 is 23.0 with 290 - so a config that names a built
///  BPL cannot derive the suffix from the version. Falls back to the
///  version without its dot, which is what this resolver did before and
///  what Delphi 13 needs, so nothing changes there.</summary>
function ProductLibSuffix(const ARootDir, ABDSVersion: string): string;

type
  TDIHPlaceholderResolver = class
  private
    FValues: TDictionary<string, string>;
    FBaseDir: string;
    FBDSVersion: string;
    FBDSProfile: string;
    FBDSCommonDir: string;
    FBDSRootDir: string;
    FOutputRoot: string;
    FOutputNote: string;
    function ExpandEnvironmentVars(const AValue: string): string;
    function ConvertDollarParenToPercent(const AValue: string): string;
    procedure ReadBDSDirsFromRegistry;
    procedure DecideOutputRoot;
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
    /// <summary>Where the BPL and the DCP really go. The same as
    ///  BDSCommonDir unless THIS account cannot write there - which happens
    ///  when Delphi was installed by another user, because its installer
    ///  writes that user's Documents folder into rsvars.bat and the
    ///  registry (forum report 2026-10-15: C:\Users\ADMHUMMMA\Documents).
    ///  </summary>
    property OutputRoot: string read FOutputRoot;
    /// <summary>Empty when nothing is unusual; otherwise the sentence that
    ///  names what could not be written and where the output goes instead.
    ///  The caller logs it - a silent substitution is worse than the
    ///  error it replaces.</summary>
    property OutputNote: string read FOutputNote;
  end;

implementation

uses
  Winapi.Windows, System.IOUtils, System.Win.Registry;

function ProductLibSuffix(const ARootDir, ABDSVersion: string): string;
var
  SR: TSearchRec;
  Digits: string;
begin
  // The fallback first, so every exit path has an answer.
  Result := StringReplace(ABDSVersion, '.', '', [rfReplaceAll]);
  if ARootDir = '' then Exit;
  if FindFirst(IncludeTrailingPathDelimiter(ARootDir) + 'bin\rtl*.bpl',
    faAnyFile, SR) <> 0 then Exit;
  try
    repeat
      Digits := ChangeFileExt(SR.Name, '');
      Delete(Digits, 1, 3);                       // 'rtl'
      if (Digits <> '') and (Length(Digits) <= 4) then
      begin
        var AllDigits := True;
        for var C in Digits do
          if not CharInSet(C, ['0'..'9']) then AllDigits := False;
        if AllDigits then Exit(Digits);
      end;
    until FindNext(SR) <> 0;
  finally
    System.SysUtils.FindClose(SR);
  end;
end;

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
  // Where a per-user tool belongs. Read from the environment, so a
  // packaged (MSIX) host's redirected view is what a target resolves to -
  // the same folder that host's own processes see.
  FValues.AddOrSetValue('LocalAppData', GetEnvironmentVariable('LOCALAPPDATA'));
  // {$LIBSUFFIX AUTO} appends the PRODUCT suffix, which is the IDE
  // version without its dot for Delphi 13 and something else before it -
  // so it is read from the installation (see ProductLibSuffix).
  FValues.AddOrSetValue('LibSuffix',
    ProductLibSuffix(FBDSRootDir, FBDSVersion));
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

  DecideOutputRoot;
end;

// Can this account put a file into ADir? A missing directory is created -
// the build would do that anyway - and the probe file is removed again.
function CanWriteInto(const ADir: string): Boolean;
var
  Probe: string;
begin
  Result := False;
  if ADir.IsEmpty then Exit;
  try
    if not TDirectory.Exists(ADir) then
      if not ForceDirectories(ADir) then Exit;
    Probe := IncludeTrailingPathDelimiter(ADir) + '.dih-write-probe.tmp';
    TFile.WriteAllText(Probe, 'x');
    TFile.Delete(Probe);
    Result := True;
  except
    on E: Exception do Result := False;
  end;
end;

procedure DecideBuildOutputRoot(const ACommonDir, ABDSVersion,
  ALocalAppData: string; const ACanWrite: TDIHWriteProbe;
  out ARoot, ANote: string);
var
  Alt: string;
begin
  ARoot := ACommonDir;
  ANote := '';
  if not Assigned(ACanWrite) then Exit;
  if ACanWrite(ACommonDir) then Exit;

  Alt := IncludeTrailingPathDelimiter(ALocalAppData) + 'Embarcadero' +
    PathDelim + 'Studio' + PathDelim + ABDSVersion;
  if not ACanWrite(Alt) then
  begin
    // Both refused: the build is going to fail, and the reason is worth more
    // than the MSB3191 it will fail with.
    ANote := Format('WARNING: this account can write to neither "%s" nor ' +
      '"%s", so the build will fail. Delphi was installed by another user - ' +
      'its installer writes that path into rsvars.bat and the registry.',
      [ACommonDir, Alt]);
    Exit;
  end;
  ARoot := Alt;
  ANote := Format('Delphi''s common directory "%s" cannot be written by this ' +
    'account (Delphi was installed by another user), so the BPL and the DCP ' +
    'go to "%s" and are registered from there.', [ACommonDir, Alt]);
end;

procedure TDIHPlaceholderResolver.DecideOutputRoot;
begin
  // BDSCommonDir itself is deliberately left ALONE: an existing registration
  // spelled "$(BDSCOMMONDIR)\Bpl\X.bpl" means the IDE's value, and the
  // duplicate / stale-entry / suspend checks compare paths against it.
  DecideBuildOutputRoot(FBDSCommonDir, FBDSVersion,
    ExpandEnvironmentVars('%LOCALAPPDATA%'), CanWriteInto,
    FOutputRoot, FOutputNote);
end;

procedure TDIHPlaceholderResolver.SetDefaults(APlatform: TDIHPlatform; const ABuildConfig: string);
var
  PlatStr, DcuDir, BplDir, DcpDir, HppDir, PlatSubDir: string;
begin
  PlatStr := APlatform.ToString;

  // Win32 uses Bpl/ and Dcp/ directly, Win64 uses Bpl/Win64/ and Dcp/Win64/
  if APlatform = dpWin32 then
    PlatSubDir := ''
  else
    PlatSubDir := PathDelim + PlatStr;

  DcuDir := IncludeTrailingPathDelimiter(FBaseDir) + 'DCU' + PathDelim + PlatStr + PathDelim + ABuildConfig;
  // OutputRoot, not BDSCommonDir - see DecideOutputRoot.
  BplDir := FOutputRoot + PathDelim + 'Bpl' + PlatSubDir;
  DcpDir := FOutputRoot + PathDelim + 'Dcp' + PlatSubDir;
  // The C++ headers: the same place the IDE uses, so this changes nothing
  // where the common directory IS writable - and it follows the fallback
  // where it is not. CodeGear.Delphi.Targets creates this directory for
  // every project, whether or not it emits a header.
  HppDir := FOutputRoot + PathDelim + 'hpp' + PathDelim + PlatStr;

  FValues.AddOrSetValue('DcuTargetDir', DcuDir);
  FValues.AddOrSetValue('BplTargetDir', BplDir);
  FValues.AddOrSetValue('DcpTargetDir', DcpDir);
  FValues.AddOrSetValue('HppTargetDir', HppDir);
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
