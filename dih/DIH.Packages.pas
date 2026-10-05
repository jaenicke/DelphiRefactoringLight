(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Packages;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  /// <summary>A registry value that was taken out for the duration of a
  ///  build, so it can be put back exactly as it was.</summary>
  TDIHSuspendedValue = record
    Key: string;
    Name: string;
    Value: string;
  end;

  TDIHPackageManager = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    function GetKnownPackagesKey(APlatform: TDIHPlatform): string;
    function ExpandRegisteredPath(const AValueName: string): string;
    procedure RemoveStaleEntries(AReg: TRegistry; const AKeepName, ATargetFile: string);
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
    procedure RegisterPackages(const APackages: TArray<TDIHPackageEntry>; APlatform: TDIHPlatform);
    procedure UnregisterPackages(const APackages: TArray<TDIHPackageEntry>; APlatform: TDIHPlatform);
    /// <summary>Takes this entry's packages OUT of "Known Packages" and
    ///  returns what was there. For the bds.exe path only, and the reason
    ///  is the user's (2026-10-05): bds.exe IS the IDE, and the IDE loads
    ///  every registered design-time package at startup - so it reports
    ///  "[Fataler Fehler] Package X.bpl kann nicht geladen werden" when
    ///  the file is not there yet, and worse, it HOLDS the .bpl we are
    ///  about to link. A package that is not registered is not loaded.
    ///  Always pair with RestoreSuspended in a finally.</summary>
    function SuspendPackages(const APackages: TArray<TDIHPackageEntry>;
      APlatform: TDIHPlatform): TArray<TDIHSuspendedValue>;
    /// <summary>Writes suspended values back, whatever happened to the
    ///  build. Safe to call with an empty array.</summary>
    procedure RestoreSuspended(const AValues: TArray<TDIHSuspendedValue>);
  end;

  TDIHExpertManager = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    function GetExpertsKey(APlatform: TDIHPlatform): string;
    function ResolveExpertName(const AEntry: TDIHExpertEntry): string;
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
    procedure RegisterExperts(const AExperts: TArray<TDIHExpertEntry>; APlatform: TDIHPlatform);
    procedure UnregisterExperts(const AExperts: TArray<TDIHExpertEntry>; APlatform: TDIHPlatform);
    /// <summary>The same for experts: a registered expert DLL is loaded by
    ///  the IDE at startup too, so a bds.exe build would hold the file it
    ///  is supposed to replace.</summary>
    function SuspendExperts(const AExperts: TArray<TDIHExpertEntry>;
      APlatform: TDIHPlatform): TArray<TDIHSuspendedValue>;
  end;

implementation

{ TDIHPackageManager }

constructor TDIHPackageManager.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
end;

function TDIHPackageManager.GetKnownPackagesKey(APlatform: TDIHPlatform): string;
begin
  // The 64-bit IDE (bin64\bds.exe) keeps its design-time packages in a
  // PARALLEL registry branch with an " x64" suffix - registering a Win64
  // BPL under the 32-bit key would load it into the 32-bit IDE, where it
  // cannot load at all. Same scheme for "Known IDE Packages x64",
  // "Experts x64" and "Environment Variables x64".
  Result := FResolver.Resolve('{#BDS}') + '\Known Packages';
  if APlatform = dpWin64 then
    Result := Result + ' x64';
end;

// A Known Packages value name as the IDE reads it: $(BDSCOMMONDIR), $(BDS),
// $(BDSBIN) and environment variables expanded.
function TDIHPackageManager.ExpandRegisteredPath(const AValueName: string): string;
var
  Common, Root: string;
  Buffer: array[0..4095] of Char;
begin
  Common := FResolver.Resolve('{#BDSCommonDir}');
  Root := ExcludeTrailingPathDelimiter(FResolver.Resolve('{#BDSRootDir}'));
  Result := StringReplace(AValueName, '$(BDSCOMMONDIR)', Common, [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '$(BDSBIN)', Root + '\bin', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '$(BDS)', Root, [rfReplaceAll, rfIgnoreCase]);
  if Pos('%', Result) > 0 then
    if ExpandEnvironmentStrings(PChar(Result), @Buffer[0], Length(Buffer)) > 0 then
      Result := string(Buffer);
end;

// Package base name without the {$LIBSUFFIX} version digits:
// 'DelphiRefactoringLight370.bpl' and 'DelphiRefactoringLight.bpl' both
// give 'DelphiRefactoringLight'.
function PackageBaseName(const AFileName: string): string;
begin
  Result := ChangeFileExt(ExtractFileName(AFileName), '');
  while (Result <> '') and CharInSet(Result[Length(Result)], ['0'..'9']) do
    SetLength(Result, Length(Result) - 1);
end;

// Removes what would make the IDE load the package twice, or ask about a
// package that is gone ("... could not be loaded - load it next time?" at
// every start, which also blocks an unattended IDE start):
//  * DUPLICATES - another spelling of the SAME file ("$(BDSCOMMONDIR)\Bpl\X"
//    next to "C:\Users\Public\...\Bpl\X"),
//  * STALE VARIANTS - the same package base name with another (or no)
//    version suffix in the SAME folder, e.g. 'DelphiRefactoringLight.bpl'
//    left behind by a build without {$LIBSUFFIX AUTO}.
// Only the folder of the package being registered is touched, so packages of
// other vendors and other IDE versions are never affected.
procedure TDIHPackageManager.RemoveStaleEntries(AReg: TRegistry;
  const AKeepName, ATargetFile: string);
var
  Names: TStringList;
  Target, TargetDir, Base, Expanded: string;
begin
  Target := ExpandFileName(ATargetFile);
  TargetDir := ExtractFilePath(Target);
  Base := PackageBaseName(Target);
  if Base = '' then Exit;
  Names := TStringList.Create;
  try
    AReg.GetValueNames(Names);
    for var N in Names do
    begin
      if SameText(N, AKeepName) then Continue;
      Expanded := ExpandFileName(ExpandRegisteredPath(N));
      var Why := '';
      if SameText(Expanded, Target) then
        Why := 'duplicate entry for the same file'
      else if SameText(ExtractFilePath(Expanded), TargetDir) and
              SameText(PackageBaseName(Expanded), Base) and
              SameText(ExtractFileExt(Expanded), '.bpl') then
      begin
        if FileExists(Expanded) then
          Why := 'other version of the same package in the same folder'
        else
          Why := 'stale entry, the file no longer exists';
      end;
      if Why <> '' then
      begin
        AReg.DeleteValue(N);
        FLogger.Detail('Removed Known Packages entry %s (%s)', [N, Why]);
      end;
    end;
  finally
    Names.Free;
  end;
end;

// Deletes the value and remembers it. ONE helper for packages and experts:
// the only difference is which key and which value name.
function SuspendValues(ALogger: TDIHLogger; const AKey: string;
  const ANames: TArray<string>): TArray<TDIHSuspendedValue>;
var
  Reg: TRegistry;
  V: TDIHSuspendedValue;
begin
  Result := nil;
  if Length(ANames) = 0 then Exit;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    // False: a key that does not exist has nothing to suspend.
    if not Reg.OpenKey(AKey, False) then Exit;
    try
      for var Name in ANames do
      begin
        if Name = '' then Continue;
        if not Reg.ValueExists(Name) then Continue;
        V.Key := AKey;
        V.Name := Name;
        try
          V.Value := Reg.ReadString(Name);
        except
          V.Value := '';          // a wrong value type is still worth putting back
        end;
        Reg.DeleteValue(Name);
        Result := Result + [V];
        ALogger.Detail('Not loaded during the build: %s', [ExtractFileName(Name)]);
      end;
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

function TDIHPackageManager.SuspendPackages(const APackages: TArray<TDIHPackageEntry>;
  APlatform: TDIHPlatform): TArray<TDIHSuspendedValue>;
var
  Names: TArray<string>;
begin
  Names := nil;
  for var Pkg in APackages do
    if APlatform in Pkg.Platforms then
      // The very name RegisterPackages writes, or we would delete nothing.
      Names := Names + [FResolver.ResolveKeepEnvVars(Pkg.BplPath)];
  Result := SuspendValues(FLogger, GetKnownPackagesKey(APlatform), Names);
end;

procedure TDIHPackageManager.RestoreSuspended(const AValues: TArray<TDIHSuspendedValue>);
var
  Reg: TRegistry;
begin
  if Length(AValues) = 0 then Exit;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    for var V in AValues do
      if Reg.OpenKey(V.Key, True) then
      try
        Reg.WriteString(V.Name, V.Value);
      finally
        Reg.CloseKey;
      end
      else
        FLogger.Error('Could not put %s back into %s - register it again ' +
          'with install.cmd', [ExtractFileName(V.Name), V.Key]);
  finally
    Reg.Free;
  end;
end;

procedure TDIHPackageManager.RegisterPackages(const APackages: TArray<TDIHPackageEntry>;
  APlatform: TDIHPlatform);
var
  Reg: TRegistry;
  Pkg: TDIHPackageEntry;
  BplPath, RegKey: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    RegKey := GetKnownPackagesKey(APlatform);

    if not Reg.OpenKey(RegKey, True) then
    begin
      FLogger.Error('Failed to open Known Packages registry key');
      Exit;
    end;

    try
      for Pkg in APackages do
      begin
        if not (APlatform in Pkg.Platforms) then
          Continue;

        BplPath := FResolver.ResolveKeepEnvVars(Pkg.BplPath);
        RemoveStaleEntries(Reg, BplPath, FResolver.Resolve(Pkg.BplPath));
        Reg.WriteString(BplPath, Pkg.Description);
        FLogger.Detail('Registered package: %s (%s)', [ExtractFileName(BplPath), Pkg.Description]);
      end;
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TDIHPackageManager.UnregisterPackages(const APackages: TArray<TDIHPackageEntry>;
  APlatform: TDIHPlatform);
var
  Reg: TRegistry;
  Pkg: TDIHPackageEntry;
  BplPath, RegKey: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    RegKey := GetKnownPackagesKey(APlatform);

    if not Reg.OpenKey(RegKey, False) then
      Exit;

    try
      for Pkg in APackages do
      begin
        if not (APlatform in Pkg.Platforms) then
          Continue;

        BplPath := FResolver.ResolveKeepEnvVars(Pkg.BplPath);
        if Reg.ValueExists(BplPath) then
        begin
          Reg.DeleteValue(BplPath);
          FLogger.Detail('Unregistered package: %s', [ExtractFileName(BplPath)]);
        end;
      end;
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

{ TDIHExpertManager }

constructor TDIHExpertManager.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
end;

function TDIHExpertManager.GetExpertsKey(APlatform: TDIHPlatform): string;
begin
  // Experts are registered per IDE VERSION - but the 64-bit IDE reads its
  // own branch ("Experts x64"), just like it does for known packages.
  Result := FResolver.Resolve('{#BDS}') + '\Experts';
  if APlatform = dpWin64 then
    Result := Result + ' x64';
end;

function TDIHExpertManager.ResolveExpertName(const AEntry: TDIHExpertEntry): string;
var
  ResolvedBpl: string;
begin
  if not AEntry.Name.IsEmpty then
    Result := FResolver.Resolve(AEntry.Name)
  else
  begin
    ResolvedBpl := FResolver.Resolve(AEntry.BplPath);
    Result := ChangeFileExt(ExtractFileName(ResolvedBpl), '');
  end;
end;

function TDIHExpertManager.SuspendExperts(const AExperts: TArray<TDIHExpertEntry>;
  APlatform: TDIHPlatform): TArray<TDIHSuspendedValue>;
var
  Names: TArray<string>;
begin
  Names := nil;
  for var Expert in AExperts do
    if APlatform in Expert.Platforms then
      Names := Names + [ResolveExpertName(Expert)];
  Result := SuspendValues(FLogger, GetExpertsKey(APlatform), Names);
end;

procedure TDIHExpertManager.RegisterExperts(const AExperts: TArray<TDIHExpertEntry>;
  APlatform: TDIHPlatform);
var
  Reg: TRegistry;
  Expert: TDIHExpertEntry;
  BplPath, RegKey, ExpertName: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    RegKey := GetExpertsKey(APlatform);

    if not Reg.OpenKey(RegKey, True) then
    begin
      FLogger.Error('Failed to open Experts registry key');
      Exit;
    end;

    try
      for Expert in AExperts do
      begin
        if not (APlatform in Expert.Platforms) then
          Continue;

        BplPath := FResolver.ResolveKeepEnvVars(Expert.BplPath);
        ExpertName := ResolveExpertName(Expert);
        Reg.WriteString(ExpertName, BplPath);
        FLogger.Detail('Registered expert: %s -> %s', [ExpertName, ExtractFileName(BplPath)]);
      end;
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TDIHExpertManager.UnregisterExperts(const AExperts: TArray<TDIHExpertEntry>;
  APlatform: TDIHPlatform);
var
  Reg: TRegistry;
  Expert: TDIHExpertEntry;
  RegKey, ExpertName: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    RegKey := GetExpertsKey(APlatform);

    if not Reg.OpenKey(RegKey, False) then
      Exit;

    try
      for Expert in AExperts do
      begin
        if not (APlatform in Expert.Platforms) then
          Continue;

        ExpertName := ResolveExpertName(Expert);
        if Reg.ValueExists(ExpertName) then
        begin
          Reg.DeleteValue(ExpertName);
          FLogger.Detail('Unregistered expert: %s', [ExpertName]);
        end;
      end;
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

end.
