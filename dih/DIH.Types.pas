(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Types;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TDIHPlatform = (dpWin32, dpWin64);
  TDIHPlatforms = set of TDIHPlatform;

  TDIHAction = (daInstall, daUninstall, daBuild);

  TDIHVerboseTarget = (vtOutput, vtLog);
  TDIHVerboseTargets = set of TDIHVerboseTarget;

  TDIHPathType = (ptNamespace, ptSearch, ptBrowsing);

  TDIHEntryResult = record
    EntryId: string;
    Description: string;
    Platform: string;
    Config: string;
    Success: Boolean;
    ErrorMsg: string;
  end;

  TDIHPathEntry = record
    PathType: TDIHPathType;
    Path: string;
    Recursive: Boolean;
    Platforms: TDIHPlatforms;
  end;

  TDIHFileEntry = record
    Source: string;
    Target: string;
    Platforms: TDIHPlatforms;
  end;

  TDIHPackageEntry = record
    BplPath: string;
    Description: string;
    Platforms: TDIHPlatforms;
  end;

  TDIHExpertEntry = record
    /// <summary>Name unter dem der Expert in der Registry eingetragen wird.
    ///  Falls leer, wird der BPL-Dateiname (ohne Endung) verwendet.</summary>
    Name: string;
    BplPath: string;
    Platforms: TDIHPlatforms;
  end;

  TDIHBuildProject = record
    ProjectPath: string;
    Platforms: TDIHPlatforms;
    ExtraParams: string;
  end;

  TDIHRegistryValue = record
    Path: string;
    Name: string;
    Value: string;
    ValueType: string; // 'String' or 'Integer'
  end;

  TDIHEventEntry = record
    Command: string;
    Description: string;
    Platforms: TDIHPlatforms;
  end;

  TDIHEntry = class
  private
    FId: string;
    FDescription: string;
    FSourceFile: string;
    FPaths: TList<TDIHPathEntry>;
    FFiles: TList<TDIHFileEntry>;
    FPackages: TList<TDIHPackageEntry>;
    FExperts: TList<TDIHExpertEntry>;
    FBuildProjects: TList<TDIHBuildProject>;
    FRegistryValues: TList<TDIHRegistryValue>;
    FPreEvents: TList<TDIHEventEntry>;
    FPostEvents: TList<TDIHEventEntry>;
  public
    constructor Create;
    destructor Destroy; override;
    property Id: string read FId write FId;
    property Description: string read FDescription write FDescription;
    property SourceFile: string read FSourceFile write FSourceFile;
    property Paths: TList<TDIHPathEntry> read FPaths;
    property Files: TList<TDIHFileEntry> read FFiles;
    property Packages: TList<TDIHPackageEntry> read FPackages;
    property Experts: TList<TDIHExpertEntry> read FExperts;
    property BuildProjects: TList<TDIHBuildProject> read FBuildProjects;
    property RegistryValues: TList<TDIHRegistryValue> read FRegistryValues;
    property PreEvents: TList<TDIHEventEntry> read FPreEvents;
    property PostEvents: TList<TDIHEventEntry> read FPostEvents;
  end;

  TDIHPlatformHelper = record helper for TDIHPlatform
    function ToString: string;
    class function FromString(const AValue: string): TDIHPlatform; static;
  end;

function ParsePlatforms(const AValue: string): TDIHPlatforms;
function PlatformsToStr(APlatforms: TDIHPlatforms): string;

const
  AllPlatforms: TDIHPlatforms = [dpWin32, dpWin64];

implementation

{ TDIHPlatformHelper }

function TDIHPlatformHelper.ToString: string;
begin
  case Self of
    dpWin32: Result := 'Win32';
    dpWin64: Result := 'Win64';
  else
    Result := 'Win32';
  end;
end;

class function TDIHPlatformHelper.FromString(const AValue: string): TDIHPlatform;
begin
  if SameText(AValue, 'Win64') then
    Result := dpWin64
  else
    Result := dpWin32;
end;

function ParsePlatforms(const AValue: string): TDIHPlatforms;
var
  Parts: TArray<string>;
  Part: string;
begin
  Result := [];
  if AValue.IsEmpty then
  begin
    Result := AllPlatforms;
    Exit;
  end;
  Parts := AValue.Split([',']);
  for Part in Parts do
  begin
    if SameText(Part.Trim, 'Win32') then
      Include(Result, dpWin32)
    else if SameText(Part.Trim, 'Win64') then
      Include(Result, dpWin64);
  end;
end;

function PlatformsToStr(APlatforms: TDIHPlatforms): string;
var
  SL: TStringList;
  P: TDIHPlatform;
begin
  SL := TStringList.Create;
  try
    SL.StrictDelimiter := True;
    SL.Delimiter := ',';
    for P in APlatforms do
      SL.Add(P.ToString);
    Result := SL.DelimitedText;
  finally
    SL.Free;
  end;
end;

{ TDIHEntry }

constructor TDIHEntry.Create;
begin
  inherited;
  FPaths := TList<TDIHPathEntry>.Create;
  FFiles := TList<TDIHFileEntry>.Create;
  FPackages := TList<TDIHPackageEntry>.Create;
  FExperts := TList<TDIHExpertEntry>.Create;
  FBuildProjects := TList<TDIHBuildProject>.Create;
  FRegistryValues := TList<TDIHRegistryValue>.Create;
  FPreEvents := TList<TDIHEventEntry>.Create;
  FPostEvents := TList<TDIHEventEntry>.Create;
end;

destructor TDIHEntry.Destroy;
begin
  FPreEvents.Free;
  FPostEvents.Free;
  FPaths.Free;
  FFiles.Free;
  FPackages.Free;
  FExperts.Free;
  FBuildProjects.Free;
  FRegistryValues.Free;
  inherited;
end;

end.
