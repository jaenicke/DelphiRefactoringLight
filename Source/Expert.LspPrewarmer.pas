(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.LspPrewarmer;

{
  Watches the IDE for project-open events and pre-warms our DelphiLSP
  instance in the background so the first refactoring action does not
  have to pay the cold-start cost.

  Mechanism:
    * TIdeFileNotifier (IOTAIDENotifier) listens for ofnFileOpened.
    * When a .dproj / .dpr is opened, schedule a background thread that
      - resolves the project's source files via TEditorHelper
      - calls TLspManager.Instance.GetClient(...)
      - calls EnsureProjectIndexed(...) so didOpen + diagnostics flow
    * Skips when the setting PrewarmLspOnProjectOpen is False.
    * Skips when a prewarm is already in flight.
}

interface

uses
  System.Classes, ToolsAPI;

type
  TLspPrewarmer = class;

  TPrewarmIdeNotifier = class(TNotifierObject, IOTANotifier, IOTAIDENotifier,
    IOTAIDENotifier50)
  private
    FOwner: TLspPrewarmer;
  public
    constructor Create(AOwner: TLspPrewarmer);
    // IOTAIDENotifier
    procedure FileNotification(NotifyCode: TOTAFileNotification;
      const FileName: string; var Cancel: Boolean);
    procedure BeforeCompile(const Project: IOTAProject; var Cancel: Boolean); overload;
    procedure AfterCompile(Succeeded: Boolean); overload;
    // IOTAIDENotifier50 - the IDE calls these when the interface is
    // implemented; IsCodeInsight separates real compiles from the
    // background code-insight compiles.
    procedure BeforeCompile(const Project: IOTAProject; IsCodeInsight: Boolean;
      var Cancel: Boolean); overload;
    procedure AfterCompile(Succeeded: Boolean; IsCodeInsight: Boolean); overload;
  end;

  TLspPrewarmer = class
  private
    FIdeNotifier: IOTAIDENotifier;
    FIdeNotifierIndex: Integer;
    FInFlight: Boolean;          // simple debounce flag
    FLastProject: string;
    procedure StartPrewarmFor(const AProjectFile: string);
  public
    constructor Create;
    destructor Destroy; override;

    procedure Install;
    procedure Uninstall;

    /// <summary>Called by TPrewarmIdeNotifier when the IDE opens a
    ///  project file. Kicks off the background prewarm.</summary>
    procedure HandleProjectOpened(const AProjectFile: string);
  end;

var
  LspPrewarmerInstance: TLspPrewarmer;

implementation

uses
  System.SysUtils, System.IOUtils, Winapi.Windows,
  Expert.PluginSettings, Expert.EditorHelperIntf, Expert.LspManager,
  Expert.UnitIndex, Expert.AutoImport;

{ TPrewarmIdeNotifier }

constructor TPrewarmIdeNotifier.Create(AOwner: TLspPrewarmer);
begin
  inherited Create;
  FOwner := AOwner;
end;

procedure TPrewarmIdeNotifier.FileNotification(NotifyCode: TOTAFileNotification;
  const FileName: string; var Cancel: Boolean);
var
  Ext: string;
begin
  if FOwner = nil then Exit;
  // Project switch: drop the live quick-fix results. They are keyed by
  // (file, content hash), so on an UNTOUCHED buffer nothing would ever
  // invalidate them - a stale fix survived even reopening the project.
  if NotifyCode = ofnActiveProjectChanged then
  begin
    LiveResetResults;
    Exit;
  end;
  if NotifyCode <> ofnFileOpened then Exit;
  if FileName = '' then Exit;
  Ext := LowerCase(ExtractFileExt(FileName));
  if (Ext <> '.dproj') and (Ext <> '.dpr') then Exit;
  LiveResetResults;   // new project context - re-analyse from scratch
  FOwner.HandleProjectOpened(FileName);
end;

procedure TPrewarmIdeNotifier.BeforeCompile(const Project: IOTAProject;
  var Cancel: Boolean);
begin
  // not needed
end;

procedure TPrewarmIdeNotifier.AfterCompile(Succeeded: Boolean);
begin
  // not needed (the IOTAIDENotifier50 overload is called instead)
end;

procedure TPrewarmIdeNotifier.BeforeCompile(const Project: IOTAProject;
  IsCodeInsight: Boolean; var Cancel: Boolean);
begin
  // not needed
end;

procedure TPrewarmIdeNotifier.AfterCompile(Succeeded: Boolean;
  IsCodeInsight: Boolean);
begin
  // A REAL compile just produced the complete diagnostic picture (incl.
  // hints the Structure view never shows). Re-arm the live quick-fix
  // checker so one LSP analysis of the active buffer picks them up.
  // Both outcomes matter: hints appear on success, errors on failure.
  if not IsCodeInsight then
    LiveRefreshAfterCompile;
end;

{ TLspPrewarmer }

constructor TLspPrewarmer.Create;
begin
  inherited Create;
  FIdeNotifierIndex := -1;
end;

destructor TLspPrewarmer.Destroy;
begin
  Uninstall;
  inherited;
end;

procedure TLspPrewarmer.Install;
var
  Services: IOTAServices;
begin
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  begin
    FIdeNotifier := TPrewarmIdeNotifier.Create(Self);
    FIdeNotifierIndex := Services.AddNotifier(FIdeNotifier);
  end;
end;

procedure TLspPrewarmer.Uninstall;
var
  Services: IOTAServices;
begin
  if (FIdeNotifierIndex >= 0)
    and Supports(BorlandIDEServices, IOTAServices, Services) then
  begin
    try
      Services.RemoveNotifier(FIdeNotifierIndex);
    except
      // ignorieren
    end;
  end;
  FIdeNotifierIndex := -1;
  FIdeNotifier := nil;
end;

procedure TLspPrewarmer.HandleProjectOpened(const AProjectFile: string);
var
  LProj: string;
  DoLsp: Boolean;
begin
  if AProjectFile = '' then Exit;
  // Debounce: ignore re-opens of the same project (the IDE fires
  // ofnFileOpened for the .dproj AND the .dpr when a project loads).
  if SameText(FLastProject, AProjectFile) then Exit;
  FLastProject := AProjectFile;
  LProj := AProjectFile;
  DoLsp := TPluginSettings.PrewarmLspOnProjectOpen and not FInFlight;
  // Defer to the next message-loop tick so the IDE has finished its
  // own project-load work first - some ToolsAPI calls (GetProjectRoot,
  // search-path access) can be flaky during construction.
  TThread.ForceQueue(nil,
    procedure
    begin
      // Add this project's scope to the identifier index (the global scope
      // is already warming from plugin startup). Independent of the LSP
      // prewarm setting.
      try TUnitIndex.Instance.RefreshSourcesFromEditor; except end;
      if DoLsp then StartPrewarmFor(LProj);
    end);
end;

procedure TLspPrewarmer.StartPrewarmFor(const AProjectFile: string);
var
  RootPath, DelphiLspJson: string;
  ScanFiles: TArray<string>;
begin
  if FInFlight then Exit;

  // Resolve config bits on the main thread (ToolsAPI is single-threaded).
  RootPath := Editor.GetProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(AProjectFile);
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then Exit;
  ScanFiles := Editor.GetProjectSourceFiles;
  if Length(ScanFiles) = 0 then Exit;

  FInFlight := True;
  TThread.CreateAnonymousThread(
    procedure
    begin
      try
        try
          // Lower the worker thread's priority so the LSP cold-start
          // (heavy CPU + disk) does not steal cycles from the IDE
          // foreground.
          SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_BELOW_NORMAL);
          TLspManager.Instance.GetClient(RootPath, AProjectFile, DelphiLspJson);
          TLspManager.Instance.EnsureProjectIndexed(ScanFiles, nil);
        except
          // Prewarm failures are non-fatal - the wizard will retry
          // when the user explicitly triggers a refactoring.
        end;
      finally
        FInFlight := False;
      end;
    end).Start;
end;

end.
