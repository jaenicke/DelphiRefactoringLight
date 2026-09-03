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
  System.SysUtils, System.Win.Registry, Winapi.Windows,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  TDIHPackageManager = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    function GetKnownPackagesKey(APlatform: TDIHPlatform): string;
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
    procedure RegisterPackages(const APackages: TArray<TDIHPackageEntry>; APlatform: TDIHPlatform);
    procedure UnregisterPackages(const APackages: TArray<TDIHPackageEntry>; APlatform: TDIHPlatform);
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
