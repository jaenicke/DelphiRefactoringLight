(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
    /// <summary>False (the default) = the file is an INPUT and is copied
    ///  before the build, which is the order DIH always had. True = it is
    ///  a build ARTEFACT (&lt;files stage="after-build"&gt;) and is copied
    ///  once the projects of this entry are built - otherwise an entry
    ///  that builds an exe and installs it would copy the PREVIOUS
    ///  build.</summary>
    AfterBuild: Boolean;
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

    /// <summary>True when this entry has anything to do for APlatform.
    ///  An entry may be limited to one platform - a helper exe that serves
    ///  both IDEs is built once - and reporting "OK" for the other one
    ///  claims a result for work that never happened.</summary>
    function HasWorkFor(APlatform: TDIHPlatform): Boolean;
  end;

  TDIHPlatformHelper = record helper for TDIHPlatform
    function ToString: string;
    class function FromString(const AValue: string): TDIHPlatform; static;
  end;

  /// <summary>A step of an install entry that comes after its build.</summary>
  TDIHAfterBuildStep = (abCopyArtefacts, abRegisterPackages, abRegisterExperts, abPostEvents);
  TDIHAfterBuildSteps = set of TDIHAfterBuildStep;

function ParsePlatforms(const AValue: string): TDIHPlatforms;
function PlatformsToStr(APlatforms: TDIHPlatforms): string;

/// <summary>What an install entry still does after its build - the one
///  answer to what a failed build means for the rest of the entry (#52).
///  A build that FAILED leaves the PREVIOUS build's files behind: copying
///  them installs a stale binary, registering the entry's packages and
///  experts registers the old ones as if they were new, and a post event
///  would verify the old file. So nothing follows a failed build. An entry
///  without a build counts as built. Asked once in InstallEntry, not once
///  per step: a per-step "if Built" is how steps 5, 6 and the post events
///  were missed. UninstallEntry does not ask it - removing what is
///  registered does not depend on a build.</summary>
function AfterBuildSteps(ABuilt: Boolean): TDIHAfterBuildSteps;

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

function AfterBuildSteps(ABuilt: Boolean): TDIHAfterBuildSteps;
begin
  if ABuilt then
    Result := [Low(TDIHAfterBuildStep)..High(TDIHAfterBuildStep)]
  else
    Result := [];
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

function TDIHEntry.HasWorkFor(APlatform: TDIHPlatform): Boolean;
var
  P: TDIHPathEntry;
  F: TDIHFileEntry;
  K: TDIHPackageEntry;
  X: TDIHExpertEntry;
  B: TDIHBuildProject;
  E: TDIHEventEntry;
begin
  // Registry values carry no platform of their own, so an entry that has
  // any is work on every platform.
  if FRegistryValues.Count > 0 then Exit(True);
  for P in FPaths do if APlatform in P.Platforms then Exit(True);
  for F in FFiles do if APlatform in F.Platforms then Exit(True);
  for K in FPackages do if APlatform in K.Platforms then Exit(True);
  for X in FExperts do if APlatform in X.Platforms then Exit(True);
  for B in FBuildProjects do if APlatform in B.Platforms then Exit(True);
  for E in FPreEvents do if APlatform in E.Platforms then Exit(True);
  for E in FPostEvents do if APlatform in E.Platforms then Exit(True);
  Result := False;
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
