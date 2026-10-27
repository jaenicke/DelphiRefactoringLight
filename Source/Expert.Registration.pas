(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.Registration;

interface

procedure Register;

implementation

uses
  System.SysUtils, ToolsAPI,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod,
  Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard, Expert.KeyBinding, Expert.RestartHint,
  Expert.ContextMenu, Expert.UnitRenameWatcher, Expert.LspPrewarmer,
  Expert.LineEndingWatcher,
  Expert.WithRefactorWizard,
  Expert.UnitReferencesWizard, Expert.MoveToUnitWizard,
  Expert.Shortcuts, Expert.OptionsPage, Expert.UnitIndex, Expert.AutoImport,
  Expert.StructureErrors, Expert.QuickFixMarkers, Expert.StatusWindow,
  Expert.BlameGutter, Expert.PluginSettings,
  Expert.DialogHelper, Expert.ResourceMonitor, Expert.McpServer,
  Expert.PluginInfo, Expert.WorkerLatch, Expert.LspManager;

var
  /// <summary>Read in the finalization while the forms and the module list
  ///  still exist, and handed to the next instance through the marker file -
  ///  the finalization itself cannot declare locals, and by the time
  ///  RecordUnload runs (last statement, so that every shutdown has had its
  ///  say) the window and the buffer list are long gone.</summary>
  GDockAtUnload: TRlDockState;
  GHeadlessAtUnload: TArray<string>;

type
  TShortcutChangeHook = class
    class procedure HandleChanged;
  end;

class procedure TShortcutChangeHook.HandleChanged;
begin
  // Re-create the keyboard binding (the IDE caches its shortcut list)
  // and refresh the context-menu items.
  RebindKeyBinding;
  if ContextMenuInstance <> nil then
    ContextMenuInstance.RefreshShortcuts;
end;

procedure Register;
begin
  // Load configurable shortcut values from the registry first so all
  // installs below see the user's preferred bindings.
  TExpertsShortCut.LoadFromRegistry;
  TExpertsShortCut.AddListener(TShortcutChangeHook.HandleChanged);

  // Register the menu wizard (shows up under the Help menu)
  WizardInstance := TLspRenameWizard.Create;
  RegisterPackageWizard(WizardInstance);

  // Create the completion wizard
  CompletionWizardInstance := TLspCompletionWizard.Create;

  // Create the extract-method wizard
  ExtractMethodInstance := TLspExtractMethodWizard.Create;

  // Create the find-references wizard
  FindReferencesInstance := TLspFindReferencesWizard.Create;

  // Create the find-implementations wizard
  FindImplementationsInstance := TLspFindImplementationsWizard.Create;

  // Create the signature-check wizard
  SignatureCheckInstance := TLspSignatureCheckWizard.Create;

  // Create the project-wide remove-with wizard
  WithRefactorInstance := TLspWithRefactorWizard.Create;

  // Create the find-unit-references wizard (project-wide cross-references)
  UnitReferencesInstance := TLspFindUnitReferencesWizard.Create;

  // Create the move-to-unit wizard
  MoveToUnitInstance := TLspMoveToUnitWizard.Create;

  // Register keyboard shortcuts (defaults are Ctrl+Alt+Shift + R/Space/M/U/I/A,
  // user-configurable via Tools > Options > Refactoring Light).
  InstallKeyBinding;

  // Install the "Refactoring Light" submenu into the editor context menu
  ContextMenuInstance := TContextMenuInstaller.Create;
  ContextMenuInstance.Install;

  // Add the options page (Tools > Options > Third Party > Refactoring Light)
  RegisterOptionsPage;

  // Dockable status window. Registering here (inside Register) is what
  // lets the IDE restore it from a saved desktop layout.
  RegisterStatusWindow;
  // Help > About entry (the splash screen entry is added earlier, from
  // Expert.PluginInfo's initialization - see there).
  RegisterAboutBox;

  // Unit rename watcher: offers rename when a unit is renamed in the IDE
  // (File > Save As etc.).
  UnitRenameWatcherInstance := TUnitRenameWatcher.Create;
  UnitRenameWatcherInstance.Install;

  // Line endings: warns when a file the IDE just opened is not plain CRLF
  // and offers to correct it (user request, 2026-10-15).
  InstallLineEndingWatcher;

  // Pre-warm our DelphiLSP when a project is opened, so the first
  // refactoring action doesn't have to pay the cold-start cost.
  LspPrewarmerInstance := TLspPrewarmer.Create;
  LspPrewarmerInstance.Install;

  // Start building the GLOBAL identifier index (RTL/VCL/library paths) in
  // the background right away - it is project-independent, so it can warm
  // up before the user ever opens "Find Unit". The project scope is added
  // later when a project opens (see TLspPrewarmer) or when the dialog runs.
  TUnitIndex.Instance.StartGlobalIndex;

  // Live auto-import indicator: watches the active buffer in the background
  // and shows a small quick-fix button at the caret when an undeclared
  // identifier has a known unit (see Expert.AutoImport).
  StartAutoImportLive;
  // Warning dialog the "remove unused private member" fix asks
  // before it deletes a non-empty body.
  InstallRemovePrivateConfirm;
  // Primary diagnostics source in the IDE: the Structure view mirrors the
  // IDE's own Error Insight; its notifier feeds the live checker without
  // any LSP round-trip of our own (see Expert.StructureErrors).
  InstallStructureErrorSource;
  // Dotted-underline markers in the code editor for every fix location
  // (hints get no squiggle from the IDE itself).
  InstallQuickFixMarkers;
  // Live blame: an age stripe in the gutter and the annotation on the
  // caret line. Off unless the user switched it on (it runs git).
  InstallBlameGutter;
  // ApplyBlameSettings, NOT just SetBlameEnabled: the column width, the
  // offset and the content level live in the settings too, and only this
  // call pushes them into the painter. Switching on alone left the
  // painter on its built-in defaults until something opened the options
  // or the live adjuster - which is exactly how a tester found it ("the
  // offset was only honoured once I went into the mini dialog").
  ApplyBlameSettings;
  // GDI / USER objects and (32-bit) address space, sampled every 30 s
  // into %TEMP%\RefactoringLight-resources.log - the numbers an
  // "out of memory" report needs (see Expert.ResourceMonitor).
  StartResourceMonitor;
  // the only periodic heartbeat we have: it also retires the verification
  // LSP session after 10 idle minutes (it costs ~290 MB)
  ResourceTickHook :=
    procedure
    begin
      TLspManager.Instance.MaintainVerifySession;
    end;
  // Named pipe for the MCP bridge (RefactoringLightMcp.exe): diagnostics
  // and quick fixes for Claude Code & co, one pipe per IDE instance.
  StartMcpServer;

  // WHAT A RELOAD REALLY LOSES, and what of it can be put back. This runs
  // from the deferred tick - main thread, outside the install transaction -
  // so it may use ToolsAPI; it must not open a window (that is the tick's
  // job) and it must not raise.
  TRestartHint.SettleHook :=
    function: string
    var
      Facts: TReloadFacts;
      Reopen: Boolean;
      Proj: string;
    begin
      Facts := TRestartHint.Facts;
      if Facts.Reloaded then
      begin
        // 1. THE PROJECT. A project that was already open sends no
        //    "project opened" notification, so the identifier index has no
        //    project scope and DelphiLSP was never prewarmed - while the
        //    status window says "ready", because the global scope IS there.
        //    That is the gap a user meets as "it finds nothing any more
        //    since I installed".
        //    The answer is deliberately discarded: a catch-up that ran is a
        //    restore, and one that found nothing was not needed - neither is
        //    anything the box has to say.
        if LspPrewarmerInstance <> nil then
          LspPrewarmerInstance.AdoptAlreadyOpenProject(Proj);

        // 2. THE HEADLESS BUFFERS of the MCP bridge. The modules belong to
        //    the IDE and survive; only our list of them went with the old
        //    image - and without that list buffer_close refuses them as the
        //    user's own (audit #36, M42), so a modified module nobody can
        //    see would be unreachable for the rest of the session.
        McpAdoptHeadless(Facts.Headless);
      end;

      // 3. WHAT IS LEFT OVER. The rule is pure and tested (ReloadReportText);
      //    this only carries it out - which for a free-standing status window
      //    means putting it back, and for a docked one means saying so.
      Result := ReloadReportText(Facts, Reopen);
      if Reopen then
        ShowStatusWindowAt(Facts.DockBounds);
    end;
  // Arms that tick. FILE I/O AND STATE ONLY in here: Register runs inside
  // the IDE's install transaction, and a window there wedges it (2.6.11).
  TRestartHint.Check;
end;

initialization

finalization
  // FIRST, while everything still exists: what the next instance will have
  // to catch up. Both reads are ToolsAPI/VCL and answer nothing once the
  // teardown below has run.
  GDockAtUnload := StatusDockWindowState;
  GHeadlessAtUnload := McpHeadlessFiles;
  // And the tick may not call into the IDE-only units from here on - their
  // state is about to go. Clearing the hook is also what makes the ONE tick
  // that survives this finalization safe: the pinned case, where the image
  // can no longer be unmapped and the box is all that is left to do.
  TRestartHint.SettleHook := nil;
  // Our own windows next: a modeless result window or the quick-fix
  // popup that is still on screen would keep handlers - and a 250 ms
  // timer - in code this unload is about to unmap (audit #40, M31 /
  // #37, L3s).
  CloseTrackedPluginForms;
  // Then the dockable form back to the IDE - it must not outlive the BPL.
  UnregisterStatusWindow;
  UnregisterAboutBox;
  // Before the live checker: MCP requests read its data, and the server
  // waits for every running request before it returns.
  StopMcpServer;
  // Background workers (completion, signature help, prewarmer, find-unit
  // search, commit fetch): no new ones from here on, and the running ones
  // finish BEFORE any object they use is freed or the code is unmapped.
  // Ending DelphiLSP first wakes those waiting for an answer - their
  // requests fail at once instead of timing out.
  TLspManager.ShutdownIfRunning;
  // A worker that is still running when the image is unmapped executes code
  // that is no longer there. It is the one shutdown failure the user has to
  // hear about, so it is recorded rather than returned into nothing.
  if not ShutdownWorkersAndWait then
    NoteShutdownProblem('Background threads of the plugin did not finish in ' +
      'time. Code of the old version may still be running, which can crash ' +
      'this IDE - please restart it.');
  ResourceTickHook := nil;   // before the timer goes, and before the BPL unloads
  StopResourceMonitor;
  UninstallBlameGutter;
  UninstallQuickFixMarkers;
  UninstallStructureErrorSource;
  StopAutoImportLive;
  TExpertsShortCut.RemoveListener(TShortcutChangeHook.HandleChanged);
  UnregisterOptionsPage;
  UninstallLineEndingWatcher;
  FreeAndNil(LspPrewarmerInstance);
  FreeAndNil(UnitRenameWatcherInstance);
  FreeAndNil(ContextMenuInstance);
  FreeAndNil(MoveToUnitInstance);
  FreeAndNil(UnitReferencesInstance);
  FreeAndNil(WithRefactorInstance);
  FreeAndNil(SignatureCheckInstance);
  FreeAndNil(FindImplementationsInstance);
  FreeAndNil(FindReferencesInstance);
  FreeAndNil(ExtractMethodInstance);
  FreeAndNil(CompletionWizardInstance);
  UninstallKeyBinding;
  // LAST: every shutdown above has had its say, so this is the complete
  // picture. The marker file is the only channel to the next instance - the
  // one that knows is the one that is going away.
  TRestartHint.RecordUnload(ShutdownProblems, GHeadlessAtUnload,
    GDockAtUnload.Open, GDockAtUnload.Docked, GDockAtUnload.Bounds);
end.
