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
    /// <summary>Feeds the Messages window's compiler output into the live
    ///  quick-fix pipeline (third diagnostics source).</summary>
    procedure FeedCompilerMessages;
  end;

  TLspPrewarmer = class
  private
    FIdeNotifier: IOTAIDENotifier;
    FIdeNotifierIndex: Integer;
    FInFlight: Boolean;          // simple debounce flag
    // Audit #38, M4: FInFlight debounced, but it also DISCARDED a project
    // switch that arrived during the 10-30 s cold start - the prewarm
    // finished for the old project and nothing re-ran for the new one, so
    // the session stayed bound to the old one until some wizard happened
    // to call GetClient. The latest request is parked here instead and the
    // worker picks it up when it is done.
    FPending: string;
    // FPending came from RestartFor: run it even for the project that was
    // just warmed - the session it warmed is the one lsp_restart ended.
    FPendingRestart: Boolean;
    FLastProject: string;
    procedure StartPrewarmFor(const AProjectFile: string);
    function TryStartPrewarm(const AProjectFile: string;
      out AReason: string): Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Install;
    procedure Uninstall;

    /// <summary>Called by TPrewarmIdeNotifier when the IDE opens a
    ///  project file. Kicks off the background prewarm.</summary>
    procedure HandleProjectOpened(const AProjectFile: string);
    /// <summary>A project that was ALREADY open when this instance loaded
    ///  sends no "project opened" notification - which is what leaves a
    ///  reloaded package with an empty project index scope and no prewarmed
    ///  DelphiLSP, while the status window still says "ready". This does
    ///  exactly what the notification does, and nothing at all when the
    ///  notification has already been seen (FLastProject <> '') or when no
    ///  project is open (then it is still coming).
    ///  HandleProjectOpened deliberately, not RestartFor: only it refreshes
    ///  the identifier index as well, and only it honours
    ///  PrewarmLspOnProjectOpen. MAIN THREAD - its own ForceQueueSafe keeps
    ///  the ToolsAPI work off the caller.</summary>
    function AdoptAlreadyOpenProject(out AProject: string): Boolean;
    procedure HandleProjectSwitched(const AProjectFile: string);
    /// <summary>Starts a NEW session for AProjectFile after
    ///  TLspManager.Reset - main thread only. Unlike HandleProjectOpened it
    ///  is neither debounced nor subject to PrewarmLspOnProjectOpen: it is
    ///  an explicit request. True only when a start really began; while an
    ///  earlier start is still running the request is parked instead
    ///  (AQueued) and runs when that one finishes. AReason says why not.</summary>
    function RestartFor(const AProjectFile: string; out AQueued: Boolean;
      out AReason: string): Boolean;
  end;

var
  LspPrewarmerInstance: TLspPrewarmer;

implementation

uses
  System.SysUtils, System.IOUtils, Winapi.Windows, Lsp.Protocol,
  Expert.PluginSettings, Expert.EditorHelperIntf, Expert.LspManager,
  Expert.UnitIndex, Expert.AutoImport, Expert.MessagesReader,
  Expert.WorkerLatch;

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
    // ... and move our LSP session to the NEW active project. Without this
    // the live checker kept analysing in the old project's context and
    // published its E2003 again right after the reset (user, 2026-09-21).
    FOwner.HandleProjectSwitched(FileName);
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
  begin
    LiveRefreshAfterCompile;
    // THIRD diagnostics source: the compiler output in the Messages
    // window. It is the only one left when Error Insight / the IDE's LSP
    // have stopped answering - exactly the state the tester was in
    // (Structure view empty, our LSP silent, dcc32 errors on screen).
    FeedCompilerMessages;
    // Keep the probe log fresh: it documents what the reader sees and is
    // how we adapt if a future IDE renames a symbol.
    DumpMessagesWindow;
  end;
end;

// Reads the compiler messages for the ACTIVE buffer and hands them to the
// live quick-fix pipeline (same entry point the Structure view uses).
procedure TPrewarmIdeNotifier.FeedCompilerMessages;
var
  FileName, Content: string;
  Diags: TArray<TLspErrorDiag>;
begin
  try
    if Editor = nil then Exit;
    // Cheap getter - never GetCurrentContext here (it moves the caret).
    FileName := Editor.GetActiveFileName;
    if FileName = '' then Exit;
    Diags := ReadCompilerDiagnosticsFor(FileName);
    if Length(Diags) = 0 then Exit;
    if not Editor.ReadEditorContent(FileName, Content) then
      try
        Content := TFile.ReadAllText(FileName);
      except
        Exit;
      end;
    if Content = '' then Exit;
    LiveReportErrorDiags(FileName, Content, Diags, SrcCompiler);
  except
    // a diagnostics source must never disturb a build
  end;
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
  ForceQueueSafe(
    procedure
    begin
      // Add this project's scope to the identifier index (the global scope
      // is already warming from plugin startup). Independent of the LSP
      // prewarm setting.
      try TUnitIndex.Instance.RefreshSourcesFromEditor; except end;
      if DoLsp then StartPrewarmFor(LProj);
    end);
end;

function TLspPrewarmer.AdoptAlreadyOpenProject(out AProject: string): Boolean;
begin
  Result := False;
  AProject := '';
  // The notification has already been seen - nothing to catch up.
  if FLastProject <> '' then Exit;
  try
    AProject := Editor.GetCurrentProjectDproj;
  except
    AProject := '';
  end;
  // No project open: the notification is still ahead of us.
  if AProject = '' then Exit;
  HandleProjectOpened(AProject);
  Result := True;
end;

procedure TLspPrewarmer.HandleProjectSwitched(const AProjectFile: string);
var
  LProj: string;
begin
  LProj := AProjectFile;
  FLastProject := '';
  // State only here (notifier context); the ToolsAPI work runs queued.
  ForceQueueSafe(
    procedure
    var
      Proj: string;
    begin
      Proj := LProj;
      if not SameText(ExtractFileExt(Proj), '.dproj') then
        try Proj := Editor.GetCurrentProjectDproj; except Proj := ''; end;
      if Proj = '' then Exit;
      FLastProject := Proj;
      try TUnitIndex.Instance.RefreshSourcesFromEditor; except end;
      // A session that EXISTS is kept current even when the prewarm-on-open
      // setting is off - otherwise it would answer for the wrong project.
      if TPluginSettings.PrewarmLspOnProjectOpen
        or (TLspManager.Instance.PeekClient <> nil) then
        StartPrewarmFor(Proj);
    end);
end;

function TLspPrewarmer.RestartFor(const AProjectFile: string;
  out AQueued: Boolean; out AReason: string): Boolean;
begin
  Result := False;
  AQueued := False;
  AReason := '';
  if AProjectFile = '' then
  begin
    AReason := 'no project is open';
    Exit;
  end;
  // The notifier's debounce compares against this - keep it current, or
  // the next real open of the same project would be taken for a re-open.
  FLastProject := AProjectFile;
  if FInFlight then
  begin
    // A start begun before the reset is still running and may finish on
    // the session that was just ended - this one follows it.
    FPending := AProjectFile;
    FPendingRestart := True;
    AQueued := True;
    AReason := 'an earlier start is still running; the new session ' +
      'starts as soon as it has finished';
    Exit;
  end;
  Result := TryStartPrewarm(AProjectFile, AReason);
end;

procedure TLspPrewarmer.StartPrewarmFor(const AProjectFile: string);
var
  Reason: string;
begin
  TryStartPrewarm(AProjectFile, Reason);
end;

function TLspPrewarmer.TryStartPrewarm(const AProjectFile: string;
  out AReason: string): Boolean;
var
  RootPath, DelphiLspJson: string;
  ScanFiles: TArray<string>;
begin
  Result := False;
  AReason := '';
  if FInFlight then
  begin
    // not dropped: remembered, and run when the current one finishes
    FPending := AProjectFile;
    AReason := 'a start is already running; this one is queued behind it';
    Exit;
  end;
  FPending := '';
  FPendingRestart := False;

  // Resolve config bits on the main thread (ToolsAPI is single-threaded).
  RootPath := Editor.GetProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(AProjectFile);
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    AReason := 'the project has no .delphilsp.json';
    Exit;
  end;
  ScanFiles := Editor.GetProjectSourceFiles;
  if Length(ScanFiles) = 0 then
  begin
    AReason := 'the project lists no source files';
    Exit;
  end;

  FInFlight := True;
  // Counted worker (A7): the cold start takes 10-30 s, and unloading the
  // package meanwhile must wait for it instead of unmapping its code.
  if not StartWorker(
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
        // A switch that arrived while this ran is served now. On the MAIN
        // thread: StartPrewarmFor reads ToolsAPI (project root, the
        // .delphilsp.json, the source list), which is main-thread only.
        QueueSafe(
          procedure
          begin
            var Next := FPending;
            var Restart := FPendingRestart;
            FPending := '';
            FPendingRestart := False;
            // A parked RESTART always runs - it was asked for explicitly.
            // Anything else only when it really differs from what we just
            // warmed, and only while a session is wanted at all, the same
            // condition the notifier uses.
            if (Next <> '') and (Restart or (not SameText(Next, AProjectFile)
              and (TPluginSettings.PrewarmLspOnProjectOpen
                   or (TLspManager.Instance.PeekClient <> nil)))) then
              StartPrewarmFor(Next);
          end);
      end;
    end) then
  begin
    FInFlight := False;   // shutdown began - nothing was started
    FPending := '';
    FPendingRestart := False;
    AReason := 'the plugin is shutting down';
    Exit;
  end;
  Result := True;
end;

end.
