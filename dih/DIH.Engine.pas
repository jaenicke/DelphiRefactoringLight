(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Engine;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Generics.Collections, Winapi.Windows,
  DIH.Types, DIH.Logger, DIH.CommandLine, DIH.Placeholders, DIH.XmlConfig, DIH.Registry, DIH.FileOps, DIH.PathManager,
  DIH.Builder, DIH.Packages;

type
  TDIHEngine = class
  private
    FCmdLine: TDIHCommandLine;
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
    FConfig: TDIHXmlConfig;
    FRegistryMgr: TDIHRegistryManager;
    FFileOps: TDIHFileOperations;
    FPathMgr: TDIHPathManager;
    FBuilderObj: TDIHBuilder;
    FPackageMgr: TDIHPackageManager;
    FExpertMgr: TDIHExpertManager;
    FErrorCount: Integer;
    FSuccessCount: Integer;
    FEntryResults: TList<TDIHEntryResult>;
    procedure ProcessEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
    procedure InstallEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
    procedure UninstallEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
    procedure BuildEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
    procedure AddEntryResult(const AEntryId, ADescription, APlatform, AConfig: string; ASuccess: Boolean;
      const AErrorMsg: string = '');
    procedure PrintSummary;
    function ExecuteEvents(const AEvents: TArray<TDIHEventEntry>; APlatform: TDIHPlatform; const APhase: string): Boolean;
    function GetLogFileName: string;
    function ShouldProcessEntry(AEntry: TDIHEntry): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
    function Run: Integer;
  end;

implementation

{ TDIHEngine }

constructor TDIHEngine.Create;
begin
  inherited;
  FCmdLine := TDIHCommandLine.Create;
  FEntryResults := TList<TDIHEntryResult>.Create;
end;

destructor TDIHEngine.Destroy;
begin
  FEntryResults.Free;
  FExpertMgr.Free;
  FPackageMgr.Free;
  FBuilderObj.Free;
  FPathMgr.Free;
  FFileOps.Free;
  FRegistryMgr.Free;
  FConfig.Free;
  FResolver.Free;
  FLogger.Free;
  FCmdLine.Free;
  inherited;
end;

function TDIHEngine.GetLogFileName: string;
begin
  Result := ChangeFileExt(FCmdLine.ConfigFile, '.log');
end;

function TDIHEngine.ShouldProcessEntry(AEntry: TDIHEntry): Boolean;
begin
  if FCmdLine.EntryIds.Count = 0 then
    Result := True
  else
    Result := FCmdLine.EntryIds.IndexOf(AEntry.Id) >= 0;
end;

procedure TDIHEngine.AddEntryResult(const AEntryId, ADescription, APlatform, AConfig: string; ASuccess: Boolean;
  const AErrorMsg: string);
var
  R: TDIHEntryResult;
begin
  R.EntryId := AEntryId;
  R.Description := ADescription;
  R.Platform := APlatform;
  R.Config := AConfig;
  R.Success := ASuccess;
  R.ErrorMsg := AErrorMsg;
  FEntryResults.Add(R);
end;

procedure TDIHEngine.PrintSummary;
var
  R: TDIHEntryResult;
  SuccessTotal, ErrorTotal: Integer;
  StatusStr: string;
begin
  FLogger.Separator;
  FLogger.Info('Summary:');
  FLogger.Info('');
  FLogger.Info('  %-20s %-8s %-10s %s', ['Entry', 'Platform', 'Config', 'Result']);
  FLogger.Info('  %-20s %-8s %-10s %s', ['--------------------', '--------', '----------', '----------']);

  SuccessTotal := 0;
  ErrorTotal := 0;

  for R in FEntryResults do
  begin
    if R.Success then
    begin
      StatusStr := 'OK';
      Inc(SuccessTotal);
      FLogger.Success('  %-20s %-8s %-10s %s', [R.EntryId, R.Platform, R.Config, StatusStr]);
    end
    else
    begin
      if R.ErrorMsg.IsEmpty then
        StatusStr := 'FAILED'
      else
        StatusStr := 'FAILED: ' + R.ErrorMsg;
      Inc(ErrorTotal);
      FLogger.Error('  %-20s %-8s %-10s %s', [R.EntryId, R.Platform, R.Config, StatusStr]);
    end;
  end;

  FLogger.Info('');
  FLogger.Separator;
  if ErrorTotal = 0 then
    FLogger.Success('All %d entries completed successfully.', [SuccessTotal])
  else
    FLogger.Error('%d of %d entries failed.', [ErrorTotal, SuccessTotal + ErrorTotal]);
end;

function TDIHEngine.Run: Integer;
var
  Platform: TDIHPlatform;
  BuildConfig: string;
  Entry: TDIHEntry;
  BaseDir: string;
  I: Integer;
  ActionStr: string;
begin
  FErrorCount := 0;
  FSuccessCount := 0;

  if not FCmdLine.Parse then
    Exit(1);

  // Initialize logger
  FLogger := TDIHLogger.Create(GetLogFileName);
  FLogger.VerboseTargets := FCmdLine.VerboseTargets;
  FLogger.Info('Delphi Install Helper v1.0');
  FLogger.Separator;

  // Determine action string for logging
  case FCmdLine.Action of
    daInstall:   ActionStr := 'Installing';
    daUninstall: ActionStr := 'Uninstalling';
    daBuild:     ActionStr := 'Building';
  end;
  FLogger.Info('Action: %s', [ActionStr]);
  FLogger.Info('BDS version: %s', [FCmdLine.BDSVersion]);
  if not SameText(FCmdLine.Profile, 'BDS') then
    FLogger.Info('BDS profile: %s', [FCmdLine.Profile]);
  FLogger.Info('Configuration: %s', [FCmdLine.ConfigFile]);
  FLogger.Info('Platforms: %s', [PlatformsToStr(FCmdLine.Platforms)]);
  FLogger.Info('Build configs: %s', [FCmdLine.BuildConfigs.DelimitedText]);
  if FCmdLine.EntryIds.Count > 0 then
    FLogger.Info('Selected entries: %s', [FCmdLine.EntryIds.DelimitedText]);
  FLogger.Separator;

  // Load configuration
  BaseDir := ExtractFilePath(ExpandFileName(FCmdLine.ConfigFile));
  FConfig := TDIHXmlConfig.Create;
  try
    FConfig.Load(FCmdLine.ConfigFile);
  except
    on E: Exception do
    begin
      FLogger.Error('Failed to load configuration: %s', [E.Message]);
      Exit(1);
    end;
  end;

  FLogger.Success('Configuration loaded: %d entries found', [FConfig.Entries.Count]);

  // Initialize components
  FResolver := TDIHPlaceholderResolver.Create(BaseDir, FCmdLine.BDSVersion, FCmdLine.Profile);
  FRegistryMgr := TDIHRegistryManager.Create(FLogger, FResolver);
  FFileOps := TDIHFileOperations.Create(FLogger, FResolver, BaseDir);
  FPathMgr := TDIHPathManager.Create(FLogger, FResolver);
  FBuilderObj := TDIHBuilder.Create(FLogger, FResolver, BaseDir, FCmdLine.UseBds);
  FPackageMgr := TDIHPackageManager.Create(FLogger, FResolver);
  FExpertMgr := TDIHExpertManager.Create(FLogger, FResolver);

  // Apply custom directory overrides
  if not FCmdLine.DcuDir.IsEmpty then
    FResolver.SetValue('DcuTargetDir', FCmdLine.DcuDir);
  if not FCmdLine.BplDir.IsEmpty then
    FResolver.SetValue('BplTargetDir', FCmdLine.BplDir);
  if not FCmdLine.DcpDir.IsEmpty then
    FResolver.SetValue('DcpTargetDir', FCmdLine.DcpDir);

  // Process each platform/config combination
  for Platform in FCmdLine.Platforms do
  begin
    for I := 0 to FCmdLine.BuildConfigs.Count - 1 do
    begin
      BuildConfig := FCmdLine.BuildConfigs[I];
      FLogger.Separator;
      FLogger.Info('Processing: Platform=%s, Config=%s', [Platform.ToString, BuildConfig]);
      FLogger.IncIndent;

      // Update placeholders for current platform/config
      FResolver.SetDefaults(Platform, BuildConfig);
      FResolver.ApplyCustomValues(FCmdLine.CustomPlaceholders);

      // Ensure DCU target directory exists
      var DcuDir := FResolver.Resolve('{#DcuTargetDir}');
      if not TDirectory.Exists(DcuDir) then
      begin
        TDirectory.CreateDirectory(DcuDir);
        FLogger.Detail('Created directory: %s', [DcuDir]);
      end;

      // Process entries
      for Entry in FConfig.Entries do
      begin
        if ShouldProcessEntry(Entry) then
          ProcessEntry(Entry, Platform, BuildConfig);
      end;

      FLogger.DecIndent;
    end;
  end;

  // Print entry summary
  PrintSummary;

  Result := FErrorCount;
end;

procedure TDIHEngine.ProcessEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
var
  EntrySuccess: Boolean;
  SavedErrors: Integer;
begin
  FLogger.Info('Entry: %s - %s', [AEntry.Id, AEntry.Description]);
  FLogger.IncIndent;

  SavedErrors := FErrorCount;
  EntrySuccess := True;

  try
    case FCmdLine.Action of
      daInstall:
        InstallEntry(AEntry, APlatform, ABuildConfig);
      daUninstall:
        UninstallEntry(AEntry, APlatform, ABuildConfig);
      daBuild:
        BuildEntry(AEntry, APlatform, ABuildConfig);
    end;

    // Check if errors were added during processing
    if FErrorCount > SavedErrors then
      EntrySuccess := False;
  except
    on E: Exception do
    begin
      FLogger.Error('Error processing entry "%s": %s', [AEntry.Id, E.Message]);
      Inc(FErrorCount);
      AddEntryResult(AEntry.Id, AEntry.Description, APlatform.ToString, ABuildConfig, False, E.Message);
      FLogger.DecIndent;
      Exit;
    end;
  end;

  AddEntryResult(AEntry.Id, AEntry.Description, APlatform.ToString, ABuildConfig, EntrySuccess);

  FLogger.DecIndent;
end;

function TDIHEngine.ExecuteEvents(const AEvents: TArray<TDIHEventEntry>; APlatform: TDIHPlatform;
  const APhase: string): Boolean;
var
  Event: TDIHEventEntry;
  Cmd, ResolvedCmd: string;
  SI: TStartupInfo;
  PI: TProcessInformation;
  ExitCode: DWORD;
begin
  Result := True;
  for Event in AEvents do
  begin
    if not (APlatform in Event.Platforms) then
      Continue;

    ResolvedCmd := FResolver.Resolve(Event.Command);

    if Event.Description <> '' then
      FLogger.Info('%s: %s', [APhase, Event.Description])
    else
      FLogger.Info('%s: %s', [APhase, ResolvedCmd]);

    Cmd := 'cmd.exe /c "' + ResolvedCmd + '"';

    ZeroMemory(@SI, SizeOf(SI));
    SI.cb := SizeOf(SI);
    SI.dwFlags := STARTF_USESHOWWINDOW;
    SI.wShowWindow := SW_HIDE;

    if not CreateProcess(nil, PChar(Cmd), nil, nil, False, CREATE_NO_WINDOW, nil, PChar(FConfig.BaseDir), SI, PI) then
    begin
      FLogger.Error('%s failed to execute: %s (Error: %d)', [APhase, ResolvedCmd, GetLastError]);
      Result := False;
      Inc(FErrorCount);
      Continue;
    end;

    WaitForSingleObject(PI.hProcess, INFINITE);
    GetExitCodeProcess(PI.hProcess, ExitCode);
    CloseHandle(PI.hProcess);
    CloseHandle(PI.hThread);

    if ExitCode <> 0 then
    begin
      FLogger.Error('%s failed (exit code %d): %s', [APhase, ExitCode, ResolvedCmd]);
      Result := False;
      Inc(FErrorCount);
    end
    else
      FLogger.Success('%s completed: %s', [APhase, ExtractFileName(ResolvedCmd)]);
  end;
end;

procedure TDIHEngine.InstallEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
var
  FileCount: Integer;
begin
  // Pre-events
  if AEntry.PreEvents.Count > 0 then
    ExecuteEvents(AEntry.PreEvents.ToArray, APlatform, 'Pre-build');

  // 1. Set registry values
  if AEntry.RegistryValues.Count > 0 then
  begin
    FLogger.Info('Setting registry values...');
    FRegistryMgr.SetValues(AEntry.RegistryValues.ToArray);
    FLogger.Success('Registry values set: %d', [AEntry.RegistryValues.Count]);
    Inc(FSuccessCount);
  end;

  // 2. Add paths
  if AEntry.Paths.Count > 0 then
  begin
    FLogger.Info('Adding library paths...');
    FPathMgr.AddPaths(AEntry.Paths.ToArray, APlatform);
    FLogger.Success('Library paths added');
    Inc(FSuccessCount);
  end;

  // 3. Copy files
  if AEntry.Files.Count > 0 then
  begin
    FLogger.Info('Copying files...');
    FileCount := FFileOps.CopyFiles(AEntry.Files.ToArray, APlatform);
    FLogger.Success('Files copied: %d', [FileCount]);
    Inc(FSuccessCount);
  end;

  // 4. Build projects
  if AEntry.BuildProjects.Count > 0 then
  begin
    FLogger.Info('Building projects...');
    if FBuilderObj.Build(AEntry.BuildProjects.ToArray, APlatform, ABuildConfig) then
    begin
      FLogger.Success('All projects built successfully');
      Inc(FSuccessCount);
    end
    else
    begin
      FLogger.Error('Some projects failed to build');
      Inc(FErrorCount);
    end;
  end;

  // 5. Register packages
  if AEntry.Packages.Count > 0 then
  begin
    FLogger.Info('Registering packages...');
    FPackageMgr.RegisterPackages(AEntry.Packages.ToArray, APlatform);
    FLogger.Success('Packages registered: %d', [AEntry.Packages.Count]);
    Inc(FSuccessCount);
  end;

  // 6. Register experts
  if AEntry.Experts.Count > 0 then
  begin
    FLogger.Info('Registering experts...');
    FExpertMgr.RegisterExperts(AEntry.Experts.ToArray, APlatform);
    FLogger.Success('Experts registered: %d', [AEntry.Experts.Count]);
    Inc(FSuccessCount);
  end;

  // Post-events
  if AEntry.PostEvents.Count > 0 then
    ExecuteEvents(AEntry.PostEvents.ToArray, APlatform, 'Post-install');
end;

procedure TDIHEngine.UninstallEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
var
  FileCount: Integer;
begin
  // Pre-events
  if AEntry.PreEvents.Count > 0 then
    ExecuteEvents(AEntry.PreEvents.ToArray, APlatform, 'Pre-uninstall');

  // Reverse order of install

  // 1a. Unregister experts
  if AEntry.Experts.Count > 0 then
  begin
    FLogger.Info('Unregistering experts...');
    FExpertMgr.UnregisterExperts(AEntry.Experts.ToArray, APlatform);
    FLogger.Success('Experts unregistered');
    Inc(FSuccessCount);
  end;

  // 1b. Unregister packages
  if AEntry.Packages.Count > 0 then
  begin
    FLogger.Info('Unregistering packages...');
    FPackageMgr.UnregisterPackages(AEntry.Packages.ToArray, APlatform);
    FLogger.Success('Packages unregistered');
    Inc(FSuccessCount);
  end;

  // 2. Delete copied files
  if AEntry.Files.Count > 0 then
  begin
    FLogger.Info('Removing copied files...');
    FileCount := FFileOps.DeleteCopiedFiles(AEntry.Files.ToArray, APlatform);
    FLogger.Success('Files removed: %d', [FileCount]);
    Inc(FSuccessCount);
  end;

  // 3. Remove paths
  if AEntry.Paths.Count > 0 then
  begin
    FLogger.Info('Removing library paths...');
    FPathMgr.RemovePaths(AEntry.Paths.ToArray, APlatform);
    FLogger.Success('Library paths removed');
    Inc(FSuccessCount);
  end;

  // 4. Delete registry values
  if AEntry.RegistryValues.Count > 0 then
  begin
    FLogger.Info('Removing registry values...');
    FRegistryMgr.DeleteValues(AEntry.RegistryValues.ToArray);
    FLogger.Success('Registry values removed');
    Inc(FSuccessCount);
  end;

  // Post-events
  if AEntry.PostEvents.Count > 0 then
    ExecuteEvents(AEntry.PostEvents.ToArray, APlatform, 'Post-uninstall');
end;

procedure TDIHEngine.BuildEntry(AEntry: TDIHEntry; APlatform: TDIHPlatform; const ABuildConfig: string);
begin
  // Pre-events
  if AEntry.PreEvents.Count > 0 then
    ExecuteEvents(AEntry.PreEvents.ToArray, APlatform, 'Pre-build');

  if AEntry.BuildProjects.Count > 0 then
  begin
    FLogger.Info('Building projects...');
    if FBuilderObj.Build(AEntry.BuildProjects.ToArray, APlatform, ABuildConfig) then
    begin
      FLogger.Success('All projects built successfully');
      Inc(FSuccessCount);
    end
    else
    begin
      FLogger.Error('Some projects failed to build');
      Inc(FErrorCount);
    end;
  end
  else
    FLogger.Warning('No build projects defined for entry "%s"', [AEntry.Id]);

  // Post-events
  if AEntry.PostEvents.Count > 0 then
    ExecuteEvents(AEntry.PostEvents.ToArray, APlatform, 'Post-build');
end;

end.
