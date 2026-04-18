(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.CommandLine;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, DIH.Types;

type
  TDIHCommandLine = class
  private
    FBDSVersion: string;
    FConfigFile: string;
    FPlatforms: TDIHPlatforms;
    FBuildConfigs: TStringList;
    FEntryIds: TStringList;
    FAction: TDIHAction;
    FUseBds: Boolean;
    FCustomPlaceholders: TDictionary<string, string>;
    FDcuDir: string;
    FBplDir: string;
    FDcpDir: string;
    FProfile: string;
    FVerboseTargets: TDIHVerboseTargets;
    procedure ShowUsage;
    function ParseVerbose(const AValue: string): TDIHVerboseTargets;
  public
    constructor Create;
    destructor Destroy; override;
    function Parse: Boolean;
    property BDSVersion: string read FBDSVersion;
    property ConfigFile: string read FConfigFile;
    property Platforms: TDIHPlatforms read FPlatforms;
    property BuildConfigs: TStringList read FBuildConfigs;
    property EntryIds: TStringList read FEntryIds;
    property Action: TDIHAction read FAction;
    property UseBds: Boolean read FUseBds;
    property CustomPlaceholders: TDictionary<string, string> read FCustomPlaceholders;
    property DcuDir: string read FDcuDir;
    property BplDir: string read FBplDir;
    property DcpDir: string read FDcpDir;
    property Profile: string read FProfile;
    property VerboseTargets: TDIHVerboseTargets read FVerboseTargets;
  end;

implementation

{ TDIHCommandLine }

constructor TDIHCommandLine.Create;
begin
  inherited;
  FBuildConfigs := TStringList.Create;
  FBuildConfigs.StrictDelimiter := True;
  FBuildConfigs.Delimiter := ',';
  FEntryIds := TStringList.Create;
  FEntryIds.StrictDelimiter := True;
  FEntryIds.Delimiter := ',';
  FCustomPlaceholders := TDictionary<string, string>.Create;
  FPlatforms := [dpWin32];
  FAction := daInstall;
  FUseBds := False;
  FProfile := 'BDS';
  FVerboseTargets := [vtOutput, vtLog]; // Default: both
  FBuildConfigs.Add('Release');
end;

destructor TDIHCommandLine.Destroy;
begin
  FBuildConfigs.Free;
  FEntryIds.Free;
  FCustomPlaceholders.Free;
  inherited;
end;

function TDIHCommandLine.ParseVerbose(const AValue: string): TDIHVerboseTargets;
var
  Parts: TArray<string>;
  Part: string;
begin
  Result := [];
  Parts := AValue.Split([',']);
  for Part in Parts do
  begin
    if SameText(Part.Trim, 'output') then
      Include(Result, vtOutput)
    else if SameText(Part.Trim, 'log') then
      Include(Result, vtLog);
  end;
  if Result = [] then
    Result := [vtOutput, vtLog]; // Fallback to both if invalid
end;

procedure TDIHCommandLine.ShowUsage;
begin
  Writeln('Delphi Install Helper - Usage:');
  Writeln('');
  Writeln('  delinst.exe <bdsversion> -config <xmlfile> [options]');
  Writeln('');
  Writeln('  <bdsversion>          BDS version number, e.g. 23.0, 24.0, 37.0 (required)');
  Writeln('');
  Writeln('Options:');
  Writeln('  -config <file>        XML configuration file (required)');
  Writeln('  -platforms <list>     Comma-separated platforms: Win32,Win64');
  Writeln('                        Default: Win32');
  Writeln('  -configs <list>       Comma-separated build configs: Release,Debug');
  Writeln('                        Default: Release');
  Writeln('  -entries <list>       Comma-separated entry IDs to process');
  Writeln('                        Default: all entries');
  Writeln('  -action <action>      install, uninstall, or build');
  Writeln('                        Default: install');
  Writeln('  -verbose <targets>    Where to show compiler output: output, log, or output,log');
  Writeln('                        Default: output,log');
  Writeln('  -usebds               Use bds.exe instead of msbuild');
  Writeln('  -profile <name>       BDS registry profile name (passed as -r to bds.exe)');
  Writeln('                        Default: BDS');
  Writeln('  -dcudir <path>        Override DCU output directory');
  Writeln('  -bpldir <path>        Override BPL output directory');
  Writeln('  -dcpdir <path>        Override DCP output directory');
  Writeln('  -D:<name>=<value>     Override placeholder value');
  Writeln('');
  Writeln('Examples:');
  Writeln('  delinst.exe 37.0 -config packages.xml');
  Writeln('  delinst.exe 37.0 -config packages.xml -verbose log');
  Writeln('  delinst.exe 23.0 -config packages.xml -platforms Win32,Win64 -configs Release');
  Writeln('  delinst.exe 37.0 -config packages.xml -entries vclui -action build');
end;

function TDIHCommandLine.Parse: Boolean;
var
  I: Integer;
  Param, Value, Key: string;
  EqPos: Integer;
begin
  Result := False;

  if ParamCount = 0 then
  begin
    ShowUsage;
    Exit;
  end;

  // First positional parameter: BDS version
  FBDSVersion := ParamStr(1);
  if FBDSVersion.StartsWith('-') or (Pos('.', FBDSVersion) = 0) then
  begin
    Writeln('Error: First parameter must be the BDS version (e.g. 23.0, 37.0)');
    Writeln('');
    ShowUsage;
    Exit;
  end;

  I := 2;
  while I <= ParamCount do
  begin
    Param := ParamStr(I);

    if SameText(Param, '-config') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -config requires a filename');
        Exit;
      end;
      FConfigFile := ParamStr(I);
    end
    else if SameText(Param, '-platforms') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -platforms requires a value');
        Exit;
      end;
      FPlatforms := ParsePlatforms(ParamStr(I));
    end
    else if SameText(Param, '-configs') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -configs requires a value');
        Exit;
      end;
      FBuildConfigs.Clear;
      FBuildConfigs.DelimitedText := ParamStr(I);
    end
    else if SameText(Param, '-entries') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -entries requires a value');
        Exit;
      end;
      FEntryIds.DelimitedText := ParamStr(I);
    end
    else if SameText(Param, '-action') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -action requires a value');
        Exit;
      end;
      Value := ParamStr(I);
      if SameText(Value, 'install') then
        FAction := daInstall
      else if SameText(Value, 'uninstall') then
        FAction := daUninstall
      else if SameText(Value, 'build') then
        FAction := daBuild
      else
      begin
        Writeln('Error: Unknown action: ' + Value);
        Exit;
      end;
    end
    else if SameText(Param, '-verbose') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -verbose requires a value (output, log, or output,log)');
        Exit;
      end;
      FVerboseTargets := ParseVerbose(ParamStr(I));
    end
    else if SameText(Param, '-usebds') then
      FUseBds := True
    else if SameText(Param, '-profile') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -profile requires a name');
        Exit;
      end;
      FProfile := ParamStr(I);
    end
    else if SameText(Param, '-dcudir') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -dcudir requires a path');
        Exit;
      end;
      FDcuDir := ParamStr(I);
    end
    else if SameText(Param, '-bpldir') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -bpldir requires a path');
        Exit;
      end;
      FBplDir := ParamStr(I);
    end
    else if SameText(Param, '-dcpdir') then
    begin
      Inc(I);
      if I > ParamCount then
      begin
        Writeln('Error: -dcpdir requires a path');
        Exit;
      end;
      FDcpDir := ParamStr(I);
    end
    else if Param.StartsWith('-D:', True) then
    begin
      Value := Param.Substring(3);
      EqPos := Value.IndexOf('=');
      if EqPos > 0 then
      begin
        Key := Value.Substring(0, EqPos);
        Value := Value.Substring(EqPos + 1);
        FCustomPlaceholders.AddOrSetValue(Key, Value);
      end
      else
      begin
        Writeln('Error: Invalid placeholder format: ' + Param);
        Exit;
      end;
    end
    else
    begin
      Writeln('Error: Unknown parameter: ' + Param);
      ShowUsage;
      Exit;
    end;

    Inc(I);
  end;

  if FConfigFile.IsEmpty then
  begin
    Writeln('Error: -config parameter is required');
    Writeln('');
    ShowUsage;
    Exit;
  end;

  if not FileExists(FConfigFile) then
  begin
    Writeln('Error: Configuration file not found: ' + FConfigFile);
    Exit;
  end;

  Result := True;
end;

end.
