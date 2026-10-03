(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindReferencesWizard;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types, System.UITypes, System.Math, System.Generics.Collections,
  Vcl.Forms, Vcl.Dialogs, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF}  Expert.EditorHelperIntf, Expert.FindReferencesDialog, Expert.LspManager, Lsp.Uri, Lsp.Protocol,
  Lsp.Client, Delphi.FileEncoding, Expert.ScopeFiles, Expert.UnitIndex,
  Expert.IncludeExpansion, Expert.InterfaceLinks, Expert.ImplementationFinder,
  System.StrUtils, Expert.Version;

type
  TLspFindReferencesWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private
    // Bezeichnet nur den GERADE laufenden Search. Verschachtelte
    // Execute-Calls speichern den Vorgaengerwert auf dem Stack und
    // restaurieren ihn am Ende.
    FDialog: TFindReferencesDialog;
    /// <summary>A run WITHOUT a window: "Edit methods" verifies the
    ///  occurrences of the members it is about to move BEFORE it writes,
    ///  because afterwards those call sites are exactly the ones that no
    ///  longer compile and DelphiLSP cannot be trusted about them (user,
    ///  2026-10-04). Saved and restored per run like FDialog, so a nested
    ///  search cannot leave the outer one headless.</summary>
    FHeadless: Boolean;
    FOnStatus: TProc<string>;
    FOnAbort: TFunc<Boolean>;
    /// <summary>The result, kept whichever way the run was started - the
    ///  window path ignores it, the headless one reads it.</summary>
    FOutItems: TFindReferenceItems;
    FOutStatus: string;
    FContext: TEditorContext;
    // candidates decided from the sources, without a DelphiLSP request
    FPreSkipped: Integer;
    // what the second attempt for unanswered occurrences achieved
    FSecondPassNote: string;
    // requests DelphiLSP answered with an error (-32800 & co)
    FLspErrors: Integer;
    // "Copy report" (forum 2026-09-22: "wrong entries on the first run,
    // none on the second - what can I send you?"): how the result came
    // about, one line per decision, with the time since the search began
    FTrace: TStringList;
    FTraceT0: UInt64;
    /// <summary>Set while the declaration could not be resolved and the
    ///  caret is no declaration either: nothing may be filtered out then
    ///  (forum 2026-09-30).</summary>
    FNoAnchor: Boolean;
    procedure Trace(const AText: string);
    procedure DoGotoLocation(AItem: TFindReferenceItem);
    /// <summary>The three ways the search talks to its result: the window
    ///  when there is one, the callback / the stored result otherwise.
    ///  Deliberately NOT named like what they forward to - PR #23 was a
    ///  "Status" that called itself and overflowed the stack on the first
    ///  status line, which in a BPL kills the IDE.</summary>
    procedure Status(const AText: string);
    procedure Progress(ACurrent, ATotal: Integer);
    procedure PublishItems(const AItems: TFindReferenceItems);

    function FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
    function VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName: string;
      const ATargets: TLspSymbolTargets; ALinked: TLinkedTargets;
      AClient: TLspClient; AIncludes: TLspIncludeContext;
      AGraph: TTypeGraph; const AOwnerType: string): TFindReferenceItems;
    function ConvertLspLocations(const ALocations: TArray<TLspLocation>; const AOldName: string): TFindReferenceItems;
    /// <summary>The search must stop: the user closed the window, or the
    ///  IDE is shutting down. The scan runs on the MAIN thread, so one
    ///  that nobody watches any more keeps the IDE busy and blocks its
    ///  shutdown (tester, 2026-09-20).</summary>
    function Aborted: Boolean;

    procedure SearchAndShow;
  public
    /// <summary>Runs the SAME search without a window and hands the result
    ///  back. AOnAbort is polled where the window's close would be - the
    ///  caller must be able to stop a scan that asks DelphiLSP once per
    ///  candidate. False = the search could not run at all (AStatus says
    ///  why).</summary>
    function SearchHeadless(const ACtx: TEditorContext;
      const AOnStatus: TProc<string>; const AOnAbort: TFunc<Boolean>;
      out AItems: TFindReferenceItems; out AStatus: string): Boolean;

    {$IFNDEF STANDALONE_BUILD}

    // IOTAWizard / IOTAMenuWizard / IOTANotifier - IDE plugin only.
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    function GetMenuText: string;

    {$ENDIF}
    procedure Execute;
  end;

var
  FindReferencesInstance: TLspFindReferencesWizard;

implementation


uses
  Winapi.Windows, Expert.PascalScanner, Expert.SafeDeletePlan, Expert.AutoImport;

const
  /// <summary>How long the scan waits UP FRONT for DelphiLSP to name the
  ///  declaration. An anchor found here only saves requests later; when it
  ///  does not arrive, the scan derives one from its candidates' own answers
  ///  (DominantAnswer), so waiting out the whole project budget - 173 s in the
  ///  reported run, for nothing - buys the user nothing but a progress
  ///  bar.</summary>
  AnchorProbeMaxMs = 30000;

// Content for the kind classification: the editor buffer (main thread),
// else the disk
function EditorOrDiskContent: TFunc<string, string>;
begin
  Result :=
    function(AFile: string): string
    begin
      if not ((Editor <> nil) and Editor.ReadEditorContent(AFile, Result)) then
        Result := ReadDelphiFile(AFile);
    end;
end;

{$IFNDEF STANDALONE_BUILD}
{ TLspFindReferencesWizard - IOTAWizard / IOTAMenuWizard / IOTANotifier glue.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

procedure TLspFindReferencesWizard.AfterSave; begin end;
procedure TLspFindReferencesWizard.BeforeSave; begin end;
procedure TLspFindReferencesWizard.Destroyed; begin end;
procedure TLspFindReferencesWizard.Modified; begin end;

function TLspFindReferencesWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.FindReferencesWizard'; end;

function TLspFindReferencesWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Find References'; end;

function TLspFindReferencesWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspFindReferencesWizard.GetMenuText: string;
begin Result := 'Find references...'; end;
{$ENDIF}
procedure TLspFindReferencesWizard.Execute;
var
  PrevDialog: TFindReferencesDialog;
  PrevContext: TEditorContext;
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;

  if not Ctx.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak + 'Please place the cursor on an identifier.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // Save the fields so a nested Execute (e.g. user triggers Find References
  // again via shortcut while ProcessMessages is pumping) doesn't clobber
  // the currently-running search. After this Execute returns, the dialog
  // is detached (SetClosable) and continues to live on its own - that's
  // what gives us multiple-dialogs-at-once support.
  PrevDialog := FDialog;
  PrevContext := FContext;
  var PrevTrace := FTrace;
  var PrevT0 := FTraceT0;
  // ... and the four COUNTERS/VERDICT fields too (audit #40, M10): a
  // nested Execute resets them in SearchAndShow, so the outer search used
  // to finish with the inner one's numbers and its anchor verdict - a
  // summary that describes another search, and a "no anchor" mode that
  // may not apply.
  var PrevPreSkipped := FPreSkipped;
  var PrevSecondPassNote := FSecondPassNote;
  var PrevLspErrors := FLspErrors;
  var PrevNoAnchor := FNoAnchor;
  var PrevHeadless := FHeadless;
  // EXPLICIT types: the inference of an inline var CALLS a parameterless
  // function reference, so "var P := FOnAbort" made P a Boolean.
  var PrevOnStatus: TProc<string> := FOnStatus;
  var PrevOnAbort: TFunc<Boolean> := FOnAbort;
  FTrace := TStringList.Create;
  FTraceT0 := GetTickCount64;
  try
    FContext := Ctx;
    FHeadless := False;
    FOnStatus := nil;
    FOnAbort := nil;
    FDialog := TFindReferencesDialog.CreateDialog(Application.MainForm, Ctx.WordAtCursor);
    FDialog.OnGotoLocation := DoGotoLocation;
    // No OnDialogClose: each dialog manages its own free via
    // SetClosable + caFree once the search is done.
    TLspManager.Instance.ApplyStatusToCaption(FDialog);
    // Show non-modal so the user can interact with the editor (e.g.
    // jump to a found location and edit it) while the result list
    // stays open. Multiple dialogs may co-exist - each search is
    // tracked by its own FDialog/FContext during this Execute call.
    FDialog.Show;
    try
      Application.ProcessMessages;
      SearchAndShow;
    except
      on E: Exception do
        if FDialog <> nil then
          FDialog.SetStatus('Error: ' + E.Message);
    end;
    // Hand off ownership: from now on closing the dialog frees it. If
    // the user already clicked X / Close during the scan, this
    // releases it now.
    if (FDialog <> nil) and not FDialog.CloseRequested then
      FDialog.SetReport(FTrace.Text);
    if FDialog <> nil then
      FDialog.SetClosable;
  finally
    FTrace.Free;
    FTrace := PrevTrace;
    FTraceT0 := PrevT0;
    FDialog := PrevDialog;
    FContext := PrevContext;
    FPreSkipped := PrevPreSkipped;
    FSecondPassNote := PrevSecondPassNote;
    FLspErrors := PrevLspErrors;
    FNoAnchor := PrevNoAnchor;
    FHeadless := PrevHeadless;
    FOnStatus := PrevOnStatus;
    FOnAbort := PrevOnAbort;
  end;
end;

procedure TLspFindReferencesWizard.Trace(const AText: string);
begin
  if FTrace <> nil then
    FTrace.Add(Format('%7d ms  %s', [GetTickCount64 - FTraceT0, AText]));
end;

procedure TLspFindReferencesWizard.Status(const AText: string);
begin
  if FDialog <> nil then
    FDialog.SetStatus(AText);
  FOutStatus := AText;
  if Assigned(FOnStatus) then FOnStatus(AText);
end;

procedure TLspFindReferencesWizard.Progress(ACurrent, ATotal: Integer);
begin
  if FDialog <> nil then FDialog.SetProgress(ACurrent, ATotal);
end;

procedure TLspFindReferencesWizard.PublishItems(const AItems: TFindReferenceItems);
begin
  if FDialog <> nil then FDialog.SetItems(AItems);
  FOutItems := AItems;
end;

function TLspFindReferencesWizard.Aborted: Boolean;
begin
  // WITHOUT a window there is no window to close - the old test
  // ("FDialog = nil" means gone) would make a headless run abort before its
  // first candidate, silently and with an empty result.
  if FHeadless then
    Result := Application.Terminated or
      (Assigned(FOnAbort) and FOnAbort())
  else
    Result := (FDialog = nil) or FDialog.CloseRequested or Application.Terminated;
end;

function TLspFindReferencesWizard.SearchHeadless(const ACtx: TEditorContext;
  const AOnStatus: TProc<string>; const AOnAbort: TFunc<Boolean>;
  out AItems: TFindReferenceItems; out AStatus: string): Boolean;
begin
  AItems := nil;
  AStatus := '';
  // The same save / restore Execute does, for the same reason (audit #40,
  // M10): this can be called from a message pump of another search.
  var PrevDialog := FDialog;
  var PrevContext := FContext;
  var PrevTrace := FTrace;
  var PrevT0 := FTraceT0;
  var PrevPreSkipped := FPreSkipped;
  var PrevSecondPassNote := FSecondPassNote;
  var PrevLspErrors := FLspErrors;
  var PrevNoAnchor := FNoAnchor;
  var PrevHeadless := FHeadless;
  // EXPLICIT types: the inference of an inline var CALLS a parameterless
  // function reference, so "var P := FOnAbort" made P a Boolean.
  var PrevOnStatus: TProc<string> := FOnStatus;
  var PrevOnAbort: TFunc<Boolean> := FOnAbort;
  var PrevOutItems := FOutItems;
  var PrevOutStatus := FOutStatus;
  FTrace := TStringList.Create;
  FTraceT0 := GetTickCount64;
  try
    FDialog := nil;
    FHeadless := True;
    FOnStatus := AOnStatus;
    FOnAbort := AOnAbort;
    FOutItems := nil;
    FOutStatus := '';
    FContext := ACtx;
    Result := True;
    try
      SearchAndShow;
    except
      on E: Exception do
      begin
        FOutStatus := 'Error: ' + E.Message;
        Result := False;
      end;
    end;
    AItems := FOutItems;
    AStatus := FOutStatus;
  finally
    FTrace.Free;
    FTrace := PrevTrace;
    FTraceT0 := PrevT0;
    FDialog := PrevDialog;
    FContext := PrevContext;
    FPreSkipped := PrevPreSkipped;
    FSecondPassNote := PrevSecondPassNote;
    FLspErrors := PrevLspErrors;
    FNoAnchor := PrevNoAnchor;
    FHeadless := PrevHeadless;
    FOnStatus := PrevOnStatus;
    FOnAbort := PrevOnAbort;
    FOutItems := PrevOutItems;
    FOutStatus := PrevOutStatus;
  end;
end;

procedure TLspFindReferencesWizard.DoGotoLocation(AItem: TFindReferenceItem);
begin
  // Static-style: uses only the item's own location data. Safe to be
  // assigned as a callback to dialog instances that may outlive the
  // particular Execute call that opened them.
  Editor.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
end;

procedure TLspFindReferencesWizard.SearchAndShow;
var
  DelphiLspJson, RootPath, DefFilePath: string;
  ProjFiles: TArray<string>;
  Items: TFindReferenceItems;
  Client: TLspClient;
  LspLocations: TArray<TLspLocation>;
  LspLine, LspCol: Integer;
  DefLineOut: Integer;
begin
  FPreSkipped := 0;
  FSecondPassNote := '';
  FLspErrors := 0;
  Trace(Format('Find references - %s %s, %s', [PluginName, PluginVersion,
    FormatDateTime('yyyy-mm-dd hh:nn:ss', Now)]));
  Trace(Format('symbol "%s" at %s:%d:%d', [FContext.WordAtCursor, FContext.FileName,
    FContext.Line, FContext.Column]));
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    Status(LspConfigMissingHint);
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  // Save all editor changes
  Status('Saving all files...');
  Editor.SaveAllFiles;

  // Start LSP
  var WasRunning := TLspManager.Instance.IsAlive;
  if WasRunning then
    Status('LSP already running. Opening file...')
  else
    Status('Starting LSP server (one-time)...');

  Client := TLspManager.Instance.GetClient(
    RootPath, FContext.ProjectFile, DelphiLspJson);
  Trace(Format('session: %s, project %s, diagnostics pushed so far: %d, ' +
    'server busy: "%s", reports progress: %s', [IfThen(WasRunning, 'was running',
    'STARTED NOW'), ExtractFileName(FContext.ProjectFile), Client.GetDiagnosticsCount,
    Client.BusyWith, BoolToStr(Client.ReportsProgress, True)]));


  // The server may be BUSY (a big project takes 12-30 s to load). While it
  // is, the DelphiLSP controller aborts every request after 10 s, so asking
  // produces failures, not answers - waiting for its own "$/progress ...
  // end" is the honest readiness signal (issue #13).
  if Client.BusyWith <> '' then
  begin
    Status('DelphiLSP is busy (' + Client.BusyWith + ') - waiting for it...');
    Trace('waiting for the server to go idle (' + Client.BusyWith + ')');
    Client.WaitServerIdle(180000,
      function: Boolean
      begin
        Status('DelphiLSP is busy (' + Client.BusyWith + ') - waiting for it...');
        Application.ProcessMessages;
        Result := not Aborted;
      end);
  end;

  // send the unit only when its content changed since the last send, and
  // wait for its analysis then - a blind re-open restarts DelphiLSP's
  // analysis and every query before it is done answers null
  begin
    var StartBefore := Client.GetFileDiagnosticsVersion(FContext.FileName);
    // Wait also when the file did NOT have to be sent: right after the IDE
    // started, our session has the file but has not analysed anything yet -
    // every GotoDefinition then answers null and the whole result comes out
    // UNVERIFIED, which is what the tester saw on the first run ("LSP ready
    // (server did not publish diagnostics)" in the caption, correct rows on
    // the second run).
    var Sent := Client.SyncDocument(FContext.FileName);
    // 30 s are for a file we just SENT (it is being analysed). For one that
    // was merely never analysed, a short wait is enough - and if the
    // session stays silent for it, the client remembers that and the next
    // run does not wait at all (forum 2026-09-20: 20-30 s before every
    // single run).
    Trace(Format('start file: sent=%s, diagnostics version before=%d',
      [BoolToStr(Sent, True), StartBefore]));
    if (Sent or (StartBefore = 0)) and WasRunning then
      Client.WaitFileAnalysed(FContext.FileName, StartBefore,
        IfThen(Sent, 30000, 8000),
        function: Boolean
        begin
          if not Aborted then
            Status('Waiting for DelphiLSP to analyse ' +
              ExtractFileName(FContext.FileName) + '...');
          Application.ProcessMessages;
          Result := not Aborted;
        end);
  end;

  Trace(Format('start file: diagnostics version now=%d',
    [Client.GetFileDiagnosticsVersion(FContext.FileName)]));
  // A session that has never pushed a diagnostic may simply still be loading
  // the project - and then EVERY position request answers null, which the
  // scan used to take for "no declaration" (forum 2026-09-30: 29 of 327
  // references). documentSymbol is a readiness signal that does not depend on
  // diagnostics: the server can only answer it once it has parsed the unit.
  if Client.GetFileDiagnosticsVersion(FContext.FileName) = 0 then
  begin
    Status('DelphiLSP has not analysed ' +
      ExtractFileName(FContext.FileName) + ' yet - waiting for it...');
    if Client.WaitUnitParsed(FContext.FileName,
         LspReadinessBudgetMs(Length(ProjFiles)),
         function: Boolean
         begin
           Application.ProcessMessages;
           Result := not Aborted;
         end) then
      Trace('main session: ready (documentSymbol answered for the start file)')
    else
      Trace(Format('main session: NOT READY - documentSymbol for the start ' +
        'file went unanswered for %d s, so null answers below mean "not ' +
        'analysed", not "no declaration"',
        [LspReadinessBudgetMs(Length(ProjFiles)) div 1000]));
  end;
  if Aborted then Exit;
  LspLine := FContext.Line - 1;
  LspCol := FContext.Column - 1;

  // On first start wait until ready
  if not WasRunning then
  begin
    for var Retry := 1 to 30 do
    begin
      if Aborted then Exit;
      Status(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
      Application.ProcessMessages;
      Trace(Format('cold start: readiness probe %d', [Retry]));
      try
        var H := Client.GetHover(FContext.FileName, LspLine, LspCol);
        if H <> '' then Break;
        var D := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
        if Length(D) > 0 then Break;
      except end;
      Sleep(1000);
    end;
  end;

  if Aborted then Exit;
  // Strategy 1: try textDocument/references directly. DelphiLSP does NOT
  // offer it (no referencesProvider in any mode, measured 2026-09-21), so
  // for Delphi the text scan below is THE method, not a fallback - the
  // status only says "Fallback" when references really existed and failed.
  var Prefix := '';
  if Client.SupportsReferences then
  begin
    Prefix := 'Fallback: ';
    Status('Querying LSP server for references...');
    try
      LspLocations := Client.FindReferences(FContext.FileName,
        LspLine, LspCol, True);
    except
      on E: Exception do
      begin
        Status('LSP error on references: ' + E.Message
          + ' - switching to fallback...');
        SetLength(LspLocations, 0);
      end;
    end;

    if Length(LspLocations) > 0 then
    begin
      Items := ConvertLspLocations(LspLocations, FContext.WordAtCursor);
      AssignReferenceKinds(Items, FContext.WordAtCursor, '', -1, EditorOrDiskContent());
      PublishItems(Items);
      Status(Format('LSP: %d reference(s) found.', [Length(Items)]));
      Exit;
    end;
  end;

  // Strategy 2: text search + GotoDefinition verification
  Status(Prefix + 'Text search in project...');

  // Project + the caret's unit + the extras from the settings (see
  // Expert.ScopeFiles).
  ProjFiles := ProjectScopeFiles(FContext.FileName);

  var TextCandidates := FindCandidatesByText(FContext.WordAtCursor, ProjFiles);
  Trace(Format('text search: %d file(s) in scope, %d candidate(s)',
    [Length(ProjFiles), Length(TextCandidates)]));
  if Aborted then Exit;

  if Length(TextCandidates) = 0 then
  begin
    PublishItems(nil);
    Status('No occurrences found in the project.');
    Exit;
  end;

  // Positions inside {$I} include files are answered through the INCLUDING
  // unit, sent to DelphiLSP with the include expanded (Expert.IncludeExpansion);
  // freeing the context sends the original text again.
  // VERIFICATION runs through the agent session when it is available: it
  // answers textDocument/definition about twice as fast and the controller's
  // 10 s abort does not exist there (measured, issue #13). Identical answers
  // - and when it cannot be started, this IS the main client.
  var VClient := TLspManager.Instance.VerificationClient(Client, RootPath,
    FContext.ProjectFile, DelphiLspJson);
  if VClient <> Client then
  begin
    // it would read the start file from DISK otherwise - an unsaved buffer
    // would be verified against the wrong text
    var StartContent: string;
    if EditorOrDiskReader()(FContext.FileName, StartContent) then
      VClient.SyncDocumentWith(FContext.FileName, StartContent);
  end;
  if VClient <> Client then
  begin
    // THE SESSION THAT ANSWERS MUST BE THE ONE WE CHECK (forum 2026-09-30):
    // every readiness test above ran on the MAIN session, while a FRESHLY
    // STARTED agent session did the answering - and an agent pushes no
    // diagnostics at all, so nothing could have noticed that it was still
    // loading the project. In the report it answered null for the first
    // ~55 s; the declaration query fell into that window, the caret became
    // the anchor and 298 correctly resolved references were dropped as
    // "another symbol".
    //
    // CHOSEN BY EVIDENCE, not by a proxy signal: the question this scan needs
    // answered is the DECLARATION of the symbol, so that is what both
    // sessions are asked. MEASURED here, which is why it is not documentSymbol:
    // on a fresh agent session documentSymbol answered after 84 ms while
    // definitions were still unavailable - it is a WEAKER signal than the
    // thing we depend on. The budget scales with the project (DelphiLsp loads
    // all of it first) and the window's close cancels the wait.
    // MEASURED, and it is why this is no longer the project-scaled budget
    // (forum 2026-09-30, first log): waiting 173 s for the declaration
    // produced NOTHING, and twenty seconds later the scan's own queries were
    // answered fine. An anchor found up front only saves requests (the source
    // pre-check can skip candidates); when it does not come quickly the scan
    // derives it afterwards from the answers it collects anyway. So the user
    // does not sit in front of a progress bar for minutes for a "maybe".
    var Budget := Min(LspReadinessBudgetMs(Length(ProjFiles)), AnchorProbeMaxMs);
    var Dl := GetTickCount64 + Budget;
    var Probe: TArray<TLspLocation> := nil;
    var UsedMain := False;
    repeat
      Status(Format('Waiting for DelphiLSP to resolve the declaration ' +
        '(up to %d s - close this window to stop)...', [Budget div 1000]));
      Application.ProcessMessages;
      try Probe := VClient.GotoDefinition(FContext.FileName, LspLine, LspCol);
      except Probe := nil; end;
      if Length(Probe) > 0 then Break;
      // The MAIN session has been running with the IDE all along - when IT
      // answers, it is the one that can judge the candidates.
      if not Aborted then
        try
          Probe := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
          if Length(Probe) > 0 then
          begin
            UsedMain := True;
            Break;
          end;
        except
        end;
      if Aborted or (GetTickCount64 >= Dl) then Break;
      Sleep(500);
    until False;
    if UsedMain then
    begin
      Trace('verification: the agent session did not resolve the declaration, ' +
        'the MAIN session did - the main session verifies');
      VClient := Client;
    end
    else if Length(Probe) > 0 then
      Trace('verification: separate session (' + VClient.ServerType +
        ') - it resolved the declaration')
    else
      Trace(Format('verification: separate session (%s) - NEITHER session ' +
        'resolved the declaration within %d s (documentSymbol answers: %s); ' +
        'nothing will be filtered out below',
        [VClient.ServerType, Budget div 1000,
         BoolToStr(VClient.WaitUnitParsed(FContext.FileName, 1), True)]));
  end;
  if VClient = Client then
    Trace('verification: main session');
  var IncCtx := TLspIncludeContext.Create(VClient, EditorOrDiskReader());
  try
    IncCtx.RegisterFiles(ProjFiles);

    // Resolve the declaration (for verification comparison)
    Status('Finding declaration...');
    var DefLocs := IncCtx.Definition(FContext.FileName, LspLine, LspCol);
    // a caret ON a declaration: an answer in another file is a same-named
    // symbol elsewhere (see DeclarationAnswerIsForeign)
    if Length(DefLocs) > 0 then
    begin
      var CaretText: string;
      if EditorOrDiskReader()(FContext.FileName, CaretText) then
      begin
        var CL := CaretText.Replace(#13#10, #10).Split([#10]);
        if (LspLine <= High(CL)) and DeclarationAnswerIsForeign(CL[LspLine],
          FContext.WordAtCursor, FContext.FileName, TLspUri.FileUriToPath(DefLocs[0].Uri)) then
        begin
          Trace('declaration: DelphiLSP -> ' + TLspUri.FileUriToPath(DefLocs[0].Uri) +
            ' is another symbol of that name (the caret line declares it)');
          DefLocs := nil;
        end;
      end;
    end;
    // No answer AND the caret declares nothing: we have no anchor. The
    // caret is then a USE, and taking it as the declaration makes every
    // correctly resolved candidate look foreign (forum 2026-09-30: 29 of
    // 327 references, on a cold session).
    FNoAnchor := False;
    if Length(DefLocs) = 0 then
    begin
      var CaretText2: string;
      var CaretLine2 := '';
      if EditorOrDiskReader()(FContext.FileName, CaretText2) then
      begin
        var CL2 := CaretText2.Replace(#13#10, #10).Split([#10]);
        if LspLine <= High(CL2) then CaretLine2 := CL2[LspLine];
      end;
      FNoAnchor := DeclarationAnchorUnknown(False, CaretLine2, FContext.WordAtCursor);
      if FNoAnchor then
      begin
        // one more try - a session that is still analysing often answers a
        // few seconds later, and this single answer decides the whole result
        Status('Waiting for DelphiLSP to resolve the declaration...');
        var Dl := GetTickCount64 + 20000;
        while (Length(DefLocs) = 0) and (GetTickCount64 < Dl) and not Aborted do
        begin
          Sleep(500);
          Application.ProcessMessages;
          DefLocs := IncCtx.Definition(FContext.FileName, LspLine, LspCol);
        end;
        if Length(DefLocs) > 0 then
        begin
          FNoAnchor := False;
          Trace('declaration: answered on a second attempt');
        end;
      end;
    end;
    if Length(DefLocs) > 0 then
      Trace(Format('declaration: DelphiLSP -> %s:%d:%d', [TLspUri.FileUriToPath(DefLocs[0].Uri),
        DefLocs[0].Range.Start.Line + 1, DefLocs[0].Range.Start.Character + 1]))
    else if FNoAnchor then
      Trace('declaration: NO ANSWER and the caret line declares nothing - no ' +
        'anchor, so NOTHING is filtered out (every candidate is listed, marked)')
    else
      Trace('declaration: no answer - the caret is taken as the declaration');
    // The symbol = its declaration + implementation (see TLspSymbolTargets).
    var Targets: TLspSymbolTargets;
    var DefLine := LspLine;
    if Length(DefLocs) > 0 then
    begin
      DefFilePath := TLspUri.FileUriToPath(DefLocs[0].Uri);
      DefLine := DefLocs[0].Range.Start.Line;
      DefLineOut := DefLine;
      IncCtx.AddTargetWithPartner(Targets, DefFilePath, DefLocs[0].Range.Start.Line,
        DefLocs[0].Range.Start.Character, FContext.WordAtCursor);
    end
    else
    begin
      // null AT a declaration: the caret is the symbol
      DefFilePath := FContext.FileName;
      DefLineOut := LspLine;
      IncCtx.AddTargetWithPartner(Targets, FContext.FileName, LspLine, LspCol,
        FContext.WordAtCursor);
    end;

    // Interface <-> class (user request): the interface declaration a class
    // method implements is a use of it - even when the interface is never
    // called - and calls through the interface reach it; for an interface
    // method, the implementing class methods and the calls on them.
    var Linked := TLinkedTargets.Create;
    try
      var Links: TArray<TMemberLink> := nil;
      // the graph stays alive THROUGH the verification: an occurrence
      // DelphiLSP does not answer for is resolved through the declared
      // type of its qualifier instead of being listed as unverified noise
      var Graph := TTypeGraph.Create(ProjFiles, EditorOrDiskReader());
      try
        // SELF-CONSISTENCY ANCHOR (PsyPrax report, 2026-09-21): the source
        // pre-check below judges every candidate by where the SOURCES say it
        // leads, and drops the ones that lead elsewhere than the target set.
        // That is only safe while the target set really holds the symbol's
        // declaration - and when the partner query came back empty it did
        // not: all 8 calls of TGemTiFunctions.IsConnectorUnreachable
        // resolved correctly to its declaration and were thrown away as
        // "another type's member". So the START position is judged by the
        // same resolver: the declaration it finds there IS ours, and the
        // pre-check can never disagree with itself about it again.
        begin
          var StartContent: string;
          var StartLink: TMemberLink;
          if EditorOrDiskReader()(FContext.FileName, StartContent) and
             (ResolveMemberUse(Graph, FContext.FileName, StartContent, LspLine, LspCol,
               FContext.WordAtCursor, StartLink) = murResolved) then
            Targets.Add(StartLink.FilePath, StartLink.Line);
        end;
        var Owner := TImplementationFinder.FindContainingType(DefFilePath, DefLine);
        Links := CollectLinkedTargets(Graph, Owner, FContext.WordAtCursor, Linked);
        Trace('owner type: ' + IfThen(Owner = '', '(none)', Owner));
        Trace('symbol positions: ' + Targets.Text);
        if Linked.Count > 0 then Trace('linked positions: ' + Linked.Text);

        // Verify each candidate via GotoDefinition
        if FNoAnchor then
          Status('DelphiLSP did not resolve the declaration - NOTHING ' +
            'is filtered out, every occurrence is listed and marked...');
        Items := VerifyWithLsp(TextCandidates, FContext.WordAtCursor, Targets, Linked,
          VClient, IncCtx, Graph, Owner);
        // the window is gone: stop here, but let the finally blocks below
        // restore the include documents and free the graph
        if Aborted then Exit;
      finally
        Graph.Free;
      end;

      // a linked declaration outside the scanned files (an interface of a
      // library, say) has no text candidate - list it anyway
      for var L in Links do
      begin
        var Have := False;
        for var It in Items do
          if SameText(ExpandFileName(It.FilePath), ExpandFileName(L.FilePath)) and
             (It.Line = L.Line) then Have := True;
        if Have then Continue;
        var Extra: TFindReferenceItem;
        Extra.FilePath := L.FilePath;
        Extra.Line := L.Line;
        Extra.Col := L.Col;
        Extra.Length := System.Length(FContext.WordAtCursor);
        Extra.Preview := L.Text;
        Extra.Note := '';
        Extra.Relation := Linked.DeclLabel(L.FilePath, L.Line);
        Items := Items + [Extra];
      end;
    finally
      Linked.Free;
    end;
  finally
    IncCtx.Free;
  end;

  // how each hit uses the symbol (the "Kind" column)
  AssignReferenceKinds(Items, FContext.WordAtCursor, DefFilePath, DefLineOut, EditorOrDiskContent());
  var Unverified := 0;
  for var It in Items do
    if It.Note <> '' then Inc(Unverified);
  PublishItems(Items);
  Trace(Format('done: %d row(s), %d unverified, %d decided from the sources, ' +
    '%d aborted request(s), diagnostics pushed so far: %d', [Length(Items), Unverified,
    FPreSkipped, FLspErrors, Client.GetDiagnosticsCount]));
  // how many candidates never needed a DelphiLSP request (their qualifier's
  // declared type already said they belong to another type)
  var NotAnalysed := '';
  if (Unverified > 0) and (Client.GetDiagnosticsCount = 0) then
    NotAnalysed := ' DelphiLSP has not analysed this project yet (it published ' +
      'no diagnostics at all) - that is why they are unverified; running the ' +
      'search again in a moment should verify them.';
  var FromSource := '';
  if FPreSkipped > 0 then
    FromSource := Format(' %d were decided from the sources without asking ' +
      'DelphiLSP.', [FPreSkipped]);
  if FSecondPassNote <> '' then FromSource := FromSource + ' ' + FSecondPassNote + '.';
  if FLspErrors > 0 then
    FromSource := FromSource + Format(' %d request(s) were aborted by DelphiLSP ' +
      '(it was busy - those occurrences are marked, not dropped).', [FLspErrors]);
  if Unverified > 0 then
    Status(Prefix + Format('%d of %d candidate(s) verified, %d shown UNVERIFIED ' +
      '(see the Note column).%s%s', [Length(Items) - Unverified, Length(TextCandidates),
      Unverified, FromSource, NotAnalysed]))
  else
    Status(Prefix + Format('%d of %d candidate(s) verified.%s',
      [Length(Items), Length(TextCandidates), FromSource]));
end;

function TLspFindReferencesWizard.ConvertLspLocations(const ALocations: TArray<TLspLocation>;
  const AOldName: string): TFindReferenceItems;
var
  Item: TFindReferenceItem;
  Lines: TArray<string>;
  LastFile: string;
  ResultList: TList<TFindReferenceItem>;
begin
  ResultList := TList<TFindReferenceItem>.Create;
  try
    LastFile := '';
    SetLength(Lines, 0);
    for var Loc in ALocations do
    begin
      Item.FilePath := TLspUri.FileUriToPath(Loc.Uri);
      Item.Line := Loc.Range.Start.Line;
      Item.Col := Loc.Range.Start.Character;
      Item.Length := Loc.Range.End_.Character - Loc.Range.Start.Character;
      if Item.Length <= 0 then
        Item.Length := System.Length(AOldName);

      Item.Preview := '';
      if not SameText(LastFile, Item.FilePath) then
      begin
        try
          Lines := ReadDelphiFileLines(Item.FilePath);
          LastFile := Item.FilePath;
        except
          SetLength(Lines, 0);
        end;
      end;
      if (Item.Line >= 0) and (Item.Line < System.Length(Lines)) then
        Item.Preview := Trim(Lines[Item.Line]);

      ResultList.Add(Item);
    end;
    Result := ResultList.ToArray;
  finally
    ResultList.Free;
  end;
end;

{ Helper functions for text search (analogous to Rename wizard) }

function TLspFindReferencesWizard.FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
var
  CandidateList: TList<TFindReferenceItem>;
  F, Line, RawContent: string;
  Lines, Masked: TArray<string>;
  UpperOldName: string;
  LineIdx, SearchPos, FoundPos, AfterPos: Integer;
  BeforeOk, AfterOk: Boolean;
  Item: TFindReferenceItem;
begin
  UpperOldName := UpperCase(AOldName);
  CandidateList := TList<TFindReferenceItem>.Create;
  try
    Progress(0, System.Length(AFiles));
    // the bar alone says "something happens"; the text says how far, how
    // many hits so far, and which unit - a big project has thousands of
    // files (user request, 2026-09-21). Throttled: repainting the label for
    // every file would cost more than reading many of them.
    var LastText: UInt64 := 0;
    for var FileIdx := 0 to High(AFiles) do
    begin
      F := AFiles[FileIdx];
      if (FileIdx mod 5 = 0) then
      begin
        Progress(FileIdx + 1, System.Length(AFiles));
        if GetTickCount64 - LastText >= 150 then
        begin
          LastText := GetTickCount64;
          Status(Format('Text search: %d of %d file(s), %d candidate(s) so far - %s',
            [FileIdx + 1, System.Length(AFiles), CandidateList.Count, ExtractFileName(F)]));
        end;
        Application.ProcessMessages;
        if Aborted then Break;
      end;

      try
        RawContent := ReadDelphiFile(F);
        if Pos(UpperOldName, UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(F);
        // comment/string state carried ACROSS lines (multi-line { })
        Masked := MaskCommentsAndStrings(Lines);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        Line := Lines[LineIdx];
        SearchPos := 1;
        while SearchPos <= System.Length(Line) do
        begin
          FoundPos := Pos(UpperOldName, UpperCase(Copy(Line, SearchPos)));
          if FoundPos = 0 then Break;
          FoundPos := SearchPos + FoundPos - 1;

          BeforeOk := (FoundPos = 1) or
            not CharInSet(Line[FoundPos - 1], ['A'..'Z','a'..'z','0'..'9','_']);
          AfterPos := FoundPos + System.Length(AOldName);
          AfterOk := (AfterPos > System.Length(Line)) or
            not CharInSet(Line[AfterPos], ['A'..'Z','a'..'z','0'..'9','_']);

          if BeforeOk and AfterOk and (Masked[LineIdx][FoundPos] = Line[FoundPos]) then
          begin
            Item.FilePath := F;
            Item.Line := LineIdx;
            Item.Col := FoundPos - 1;
            Item.Length := System.Length(AOldName);
            Item.Preview := Trim(Line);
            CandidateList.Add(Item);
          end;
          SearchPos := FoundPos + System.Length(AOldName);
        end;
      end;
    end;
    Progress(System.Length(AFiles), System.Length(AFiles));
    Result := CandidateList.ToArray;
  finally
    CandidateList.Free;
  end;
end;

function TLspFindReferencesWizard.VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName: string;
  const ATargets: TLspSymbolTargets; ALinked: TLinkedTargets;
  AClient: TLspClient; AIncludes: TLspIncludeContext;
  AGraph: TTypeGraph; const AOwnerType: string): TFindReferenceItems;
var
  Verified: TList<TFindReferenceItem>;
  // index in Verified <-> index in ACandidates of every row that ended
  // UNVERIFIED, for the second attempt after the pass
  Retry: TList<TPair<Integer, Integer>>;
  // What DelphiLSP answered for every LISTED row, aligned with Verified.
  // A run without an anchor derives one from exactly this (see below).
  AnsFile: TList<string>;
  AnsLine: TList<Integer>;
  // the positions the post-pass judges against: ATargets, plus the anchor
  // a run without one derives from its own answers (it cannot be added to
  // ATargets itself - a var parameter cannot be captured by the closures)
  Anchors: TLspSymbolTargets;
  Synced: TDictionary<string, Boolean>;
  Contents: TDictionary<string, string>;
  Reader: TIncludeReader;
  I: Integer;
  C: TFindReferenceItem;

  // content of a candidate's file (buffer first), read once per file
  function FileContent(const AFile: string): string;
  begin
    if Contents.TryGetValue(UpperCase(AFile), Result) then Exit;
    if not Reader(AFile, Result) then Result := '';
    Contents.Add(UpperCase(AFile), Result);
  end;

  // the ONE place a row is listed, so its answer cannot get out of step
  procedure AddRow(const ARow: TFindReferenceItem; const AAnsFile: string;
    AAnsLine: Integer);
  begin
    Verified.Add(ARow);
    AnsFile.Add(AAnsFile);
    AnsLine.Add(AAnsLine);
  end;

begin
  Verified := TList<TFindReferenceItem>.Create;
  Retry := TList<TPair<Integer, Integer>>.Create;
  AnsFile := TList<string>.Create;
  AnsLine := TList<Integer>.Create;
  Synced := TDictionary<string, Boolean>.Create;
  Contents := TDictionary<string, string>.Create;
  Reader := EditorOrDiskReader();
  try
    Progress(0, System.Length(ACandidates));

    for I := 0 to High(ACandidates) do
    begin
      C := ACandidates[I];
      if Aborted then Break;
      var Where := Format('%s:%d:%d', [ExtractFileName(C.FilePath), C.Line + 1, C.Col + 1]);
      var How := '';
      Progress(I + 1, System.Length(ACandidates));
      if (I mod 3 = 0) then
      begin
        Status(Format('Verifying %d/%d...',
          [I + 1, System.Length(ACandidates)]));
        Application.ProcessMessages;
      end;

      // PRE-CHECK without DelphiLSP: a use site whose qualifier has a
      // declared type can be recognised as ANOTHER type's member from the
      // sources alone - and then no request is needed. DelphiLSP refuses a
      // second request while one is open ("Request removed", measured
      // 2026-09-20), so every saved request is saved waiting time.
      begin
        var PreLink: TMemberLink;
        if ClassifyUnansweredUse(AGraph, C.FilePath, FileContent(C.FilePath),
          C.Line, C.Col, AOldName,
          function(AFile: string; ALine: Integer): Boolean
          begin
            Result := ATargets.Contains(AFile, ALine) or ALinked.Contains(AFile, ALine);
          end,
          function(ATypeName: string): Boolean
          begin
            Result := SameText(ATypeName, AOwnerType) or ALinked.HasType(ATypeName);
          end, PreLink) = uuOtherSymbol then
        begin
          Inc(FPreSkipped);
          Trace(Where + '  pre-check: member of ' + PreLink.TypeName +
            ' -> skipped, no request');
          Continue;
        end;
      end;

      // Send each file once, only when its content changed, and wait for
      // the analysis then (the old re-open + 300 ms lost answers). An
      // include file (or a unit currently sent expanded) is served by the
      // include context - sending it here would undo that.
      var FirstInFile := not Synced.ContainsKey(UpperCase(C.FilePath));
      if FirstInFile then
      begin
        Synced.Add(UpperCase(C.FilePath), True);
        if not AIncludes.OwnsDocument(C.FilePath) then
        begin
          if AClient.ServerType <> '' then
          begin
            // AGENT session: it pushes no diagnostics, so there is nothing
            // to wait for - measured (21 candidates in 13 units): it answers
            // straight after the didOpen, with the same answers the main
            // session gives after its analysis wait.
            var SentA := AClient.SyncDocumentWith(C.FilePath, FileContent(C.FilePath));
            Trace(ExtractFileName(C.FilePath) + '  file: ' +
              IfThen(SentA, 'sent to the agent session', 'already current'));
          end
          else
          begin
            var Before := AClient.GetFileDiagnosticsVersion(C.FilePath);
            var SentM := AClient.SyncDocument(C.FilePath);
            Trace(Format('%s  file: %s, diagnostics version %d', [ExtractFileName(C.FilePath),
              IfThen(SentM, 'sent', 'already current'), Before]));
            if SentM then
            begin
              var Name := ExtractFileName(C.FilePath);
              AClient.WaitFileAnalysed(C.FilePath, Before, 30000,
                function: Boolean
                begin
                  if not Aborted then
                    Status('Waiting for DelphiLSP to analyse ' + Name + '...');
                  Application.ProcessMessages;
                  Result := not Aborted;   // a closed window waits for nothing
                end);
            end;
          end;
        end;
      end;

      var TReq := GetTickCount64;   // how long DelphiLSP took for this one
      var Matches := False;
      var ErrText := '';
      // Outside the try: the handler below clears it, and "no answer" is
      // derived from it once, after both paths (a flag assigned in the try
      // AND in the handler made the compiler call its initial value unused).
      var Defs: TArray<TLspLocation> := nil;
      try
        // A request the server ABORTED says nothing about the symbol, so it
        // is repeated (issue #13: the controller cancels after 10 s while
        // the project loads, which is exactly when the first search runs).
        for var Attempt := 1 to 3 do
        try
          Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
          ErrText := '';
          Break;
        except
          on E: Exception do
          begin
            ErrText := E.Message;
            Inc(FLspErrors);
            if (Attempt = 3) or Aborted then raise;
            Status(Format('DelphiLSP aborted a request (%d/3), retrying...',
              [Attempt]));
            var Until_ := GetTickCount64 + 700;
            while (GetTickCount64 < Until_) and not Aborted do
            begin
              Sleep(100);
              Application.ProcessMessages;
            end;
          end;
        end;
        // an EMPTY answer is retried briefly, a wrong one never
        if (System.Length(Defs) = 0) and not ATargets.Contains(C.FilePath, C.Line) and
           not ALinked.Contains(C.FilePath, C.Line) then
        begin
          var Dl := GetTickCount64 + UInt64(IfThen(FirstInFile, 3000, 600));
          while (System.Length(Defs) = 0) and (GetTickCount64 < Dl) and not Aborted do
          begin
            Sleep(150);
            Application.ProcessMessages;
            Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
          end;
        end;
        // The candidate IS one of the symbol's positions (declaration /
        // implementation - DelphiLSP answers those with null or with the
        // counterpart), or DelphiLSP takes it to one of them. The FILE
        // alone says nothing: several same-named methods in one unit.
        if ATargets.Contains(C.FilePath, C.Line) then
        begin
          Matches := True;
          How := 'is a position of the symbol';
        end
        else if ALinked.Contains(C.FilePath, C.Line) then
        begin
          // the declaration in the interface / the implementing class
          Matches := True;
          How := 'is a linked position';
          C.Relation := ALinked.DeclLabel(C.FilePath, C.Line);
        end
        else if System.Length(Defs) > 0 then
        begin
          var DF := TLspUri.FileUriToPath(Defs[0].Uri);
          var DL := Defs[0].Range.Start.Line;
          Matches := ATargets.Contains(DF, DL);
          if Matches then How := 'answer is a position of the symbol';
          if not Matches and ALinked.Contains(DF, DL) then
          begin
            Matches := True;
            How := 'answer is a linked position';
            C.Relation := ALinked.CallLabel(DF, DL);
          end;
        end;
      except
        on E: Exception do
        begin
          // AN ERROR IS NOT A NEGATIVE ANSWER (issue #13). While a big
          // project loads, the DelphiLSP controller aborts every request
          // after 10 s with -32800 "Request removed" - and this branch
          // used to treat that like "not a reference", so the occurrence
          // vanished from the result without a trace ("1 of 341
          // candidate(s) verified"). It counts as NO ANSWER now, which
          // means: resolved from the sources if possible, otherwise
          // listed as unverified - never silently dropped.
          Defs := nil;
          Matches := False;
          ErrText := E.Message;
          Inc(FLspErrors);
        end;
      end;
      var NoAnswer := System.Length(Defs) = 0;
      // kept for the anchor derivation after the pass
      var AnsF := '';
      var AnsL := -1;
      if not NoAnswer then
      begin
        AnsF := TLspUri.FileUriToPath(Defs[0].Uri);
        AnsL := Defs[0].Range.Start.Line;
      end;
      var Answer: string;
      if not NoAnswer then
        Answer := Format('DelphiLSP -> %s:%d', [ExtractFileName(TLspUri.FileUriToPath(Defs[0].Uri)),
          Defs[0].Range.Start.Line + 1])
      else if ErrText <> '' then
        Answer := 'DelphiLSP ERROR: ' + ErrText
      else
        Answer := 'DelphiLSP: no answer';
      Answer := Answer + Format(' (%d ms)', [GetTickCount64 - TReq]);

      // DelphiLSP said nothing: resolve the use site through the declared
      // type of its qualifier. Its NEGATIVE answer is the valuable one -
      // "TMyRec2.Init" is simply another symbol and drops out instead of
      // adding an unverified row the user has to judge (tester 2026-09-19).
      if NoAnswer and not Matches then
      begin
        var Link: TMemberLink;
        case ClassifyUnansweredUse(AGraph, C.FilePath, FileContent(C.FilePath),
          C.Line, C.Col, AOldName,
          function(AFile: string; ALine: Integer): Boolean
          begin
            Result := ATargets.Contains(AFile, ALine) or ALinked.Contains(AFile, ALine);
          end,
          function(ATypeName: string): Boolean
          begin
            Result := SameText(ATypeName, AOwnerType) or ALinked.HasType(ATypeName);
          end, Link) of
          uuOurs:
            begin
              Matches := True;
              How := 'sources: member of ' + Link.TypeName;
              C.Note := Format('verified via %s (no answer from DelphiLSP)',
                [Link.TypeName]);
            end;
          uuOtherSymbol:
            if FNoAnchor then
              // no anchor: "another type" is measured against a target set
              // we do not have - keep it, marked (forum 2026-09-30)
              C.Note := Format('UNVERIFIED - the sources place it in %s, but the ' +
                'declaration of the searched symbol is unknown', [Link.TypeName])
            else
            begin
              Trace(Where + '  ' + Answer + ' | sources: member of ' + Link.TypeName +
                ' -> dropped');
              Continue;      // belongs to another type - not a reference
            end;
          uuOverloaded:
            C.Note := Format('UNVERIFIED - overload of %s, DelphiLSP gave no answer',
              [Link.TypeName]);
        end;
      end;

      if Matches then
      begin
        Trace(Where + '  ' + Answer + ' -> LISTED (' + How + ')');
        AddRow(C, AnsF, AnsL);
      end
      else if NoAnswer and (C.Note <> '') then
      begin
        Trace(Where + '  ' + Answer + ' -> LISTED, ' + C.Note);
        AddRow(C, AnsF, AnsL);            // classified above (overload of our type)
      end
      else if NoAnswer and IsIncludeFile(C.FilePath) then
      begin
        // never drop a hit in an include file silently: DelphiLSP could not
        // tell (the including unit may not compile on its own)
        C.Note := 'UNVERIFIED - no answer inside this include file';
        Trace(Where + '  ' + Answer + ' -> LISTED UNVERIFIED (include file)');
        AddRow(C, AnsF, AnsL);
      end
      else if NoAnswer and not LineDeclaresName(C.Preview, AOldName) then
      begin
        // was silently dropped: an occurrence DelphiLSP does not resolve
        // (inactive {$IFDEF} branch, unit still in analysis) may well be a
        // reference - show it, marked. A declaration line without an
        // answer declares ANOTHER symbol and stays out.
        if ErrText <> '' then
          C.Note := 'UNVERIFIED - DelphiLSP reported an error: ' + ErrText
        else
          C.Note := 'UNVERIFIED - no answer from DelphiLSP';
        Trace(Where + '  ' + Answer + ' -> LISTED UNVERIFIED');
        Retry.Add(TPair<Integer, Integer>.Create(Verified.Count, I));
        AddRow(C, AnsF, AnsL);
      end
      else if NoAnswer then
        Trace(Where + '  ' + Answer + ' -> dropped (declares another symbol)')
      else if FNoAnchor then
      begin
        // The declaration of the searched symbol is unknown, so "leads
        // elsewhere" cannot be judged - dropping here is what turned 327
        // references into 29 on a cold session (forum 2026-09-30).
        C.Note := 'UNVERIFIED - ' + Answer + ', and the declaration of the ' +
          'searched symbol is unknown';
        Trace(Where + '  ' + Answer + ' -> LISTED UNVERIFIED (no anchor)');
        AddRow(C, AnsF, AnsL);
      end
      else if ForeignAnswerVerdict(not FNoAnchor, FLspErrors) = favKeepMarked then
      begin
        // THE SAME RULE the two post-passes apply. It matters here because a
        // session that has already aborted a request is degraded, and the
        // status line promises exactly this ("those occurrences are marked,
        // not dropped") - the main pass used to drop them anyway.
        // FLspErrors grows DURING the pass, so rows seen before the first
        // abort were judged on a healthy session; that is what the two
        // post-passes are for.
        C.Note := 'UNVERIFIED - ' + Answer + ', and DelphiLSP aborted ' +
          'request(s) in this run - not dropped on a degraded session';
        Trace(Where + '  ' + Answer + ' -> LISTED UNVERIFIED (session degraded)');
        AddRow(C, AnsF, AnsL);
      end
      else
        Trace(Where + '  ' + Answer + ' -> dropped (leads to another symbol)');
    end;

    // SECOND ATTEMPT for everything DelphiLSP stayed silent about. On a
    // cold session its first analysis of a unit takes many seconds, so the
    // first run used to end with a pile of UNVERIFIED rows that were
    // correct on the second run - the tester's "beim zweiten Mal ist das
    // Ergebnis richtig". By now those files have been sent and analysed,
    // so one more query each usually answers. Cheap: only the unanswered
    // ones, and only while the window is open.
    // everything the post-pass judges against; the derivation below may add to
    // it, the normal path uses it unchanged
    Anchors := ATargets;

    // Rows BOTH post-passes decided to remove. Collected here and deleted
    // ONCE at the end, because Retry holds INDICES into Verified - deleting
    // inside the derivation would shift them under the second attempt.
    var Dropped: TArray<Integer> := nil;
    var DerivedDropped := 0;

    // NO ANCHOR, BUT THE ANSWERS AGREE (forum 2026-09-30, first log): the
    // declaration query stayed unanswered for the whole budget, so every one
    // of the 341 hits was listed UNVERIFIED - while 326 of them had been
    // resolved by the very same session to ROM_Utils.pas:15589, exactly the
    // position the warm run took as the declaration two minutes later. The
    // scan's own answers ARE the evidence; throwing them away leaves the user
    // to run the search a second time.
    if FNoAnchor and (Verified.Count > 0) and not Aborted then
    begin
      var DerFile: string;
      var DerLine: Integer;
      var Agree := DominantAnswer(AnsFile.ToArray, AnsLine.ToArray, DerFile, DerLine);
      if Agree > 0 then
      begin
        // the same treatment an answered declaration gets: the whole header is
        // one position and the partner joins the set
        var DerContent: string;
        var DF1, DL1: Integer;
        if Reader(DerFile, DerContent) and
           DeclarationHeaderSpan(SplitContentLines(DerContent), DerLine, AOldName,
             DF1, DL1) then
        begin
          for var HL := DF1 to DL1 do Anchors.Add(DerFile, HL);
          AIncludes.AddTargetWithPartner(Anchors, DerFile, DF1,
            NameColumnOnLine(SplitContentLines(DerContent)[DF1], AOldName, 0),
            AOldName);
        end
        else
          Anchors.Add(DerFile, DerLine);
        Trace(Format('no anchor, but %d of %d listed rows answered %s:%d - taken ' +
          'as the declaration, the rows are judged against it',
          [Agree, Verified.Count, ExtractFileName(DerFile), DerLine + 1]));
        FNoAnchor := False;     // there IS an anchor now
        var Cleared := 0;
        var Elsewhere2 := 0;
        for var R := 0 to Verified.Count - 1 do
        begin
          if AnsFile[R] = '' then Continue;     // never answered: stays marked
          var Row := Verified[R];
          if Anchors.Contains(AnsFile[R], AnsLine[R]) or
             ALinked.Contains(AnsFile[R], AnsLine[R]) then
          begin
            Row.Note := '';
            Inc(Cleared);
          end
          else if ForeignAnswerVerdict(True, FLspErrors) = favDrop then
          begin
            // THE SAME RULE THE SECOND ATTEMPT APPLIES (forum 2026-09-30,
            // round three): the session answered this row clearly and named
            // another symbol's declaration, and it aborted nothing in this
            // run - so the row is foreign. The reported cold run kept three
            // such rows (GlobalConfig.Formulare.BTB -> UGlobalRomConfig.pas
            // :1586) while the second attempt dropped two more of exactly
            // that shape; the warm run lists none of the five.
            Dropped := Dropped + [R];
            Inc(DerivedDropped);
          end
          else
          begin
            // A degraded session (aborted requests): its answers come from
            // the state we do not trust, so the row is reported instead of
            // removed - losing a real reference is the worse error.
            Row.Note := Format('UNVERIFIED - DelphiLSP resolved it to %s:%d, ' +
              'which is not the declaration the other answers agree on',
              [ExtractFileName(AnsFile[R]), AnsLine[R] + 1]);
            Inc(Elsewhere2);
          end;
          Verified[R] := Row;
        end;
        Trace(Format('derived anchor: %d row(s) verified against it, %d dropped ' +
          '(the answer names another symbol), %d kept and marked (degraded ' +
          'session)', [Cleared, DerivedDropped, Elsewhere2]));
        Status(Format('DelphiLSP did not answer the declaration query - ' +
          'it was derived from %d agreeing answers (%s:%d): %d verified, %d ' +
          'dropped, %d lead elsewhere.', [Agree, ExtractFileName(DerFile),
          DerLine + 1, Cleared, DerivedDropped, Elsewhere2]));
      end;
    end;

    if (Retry.Count > 0) and not Aborted then
    begin
      var Fixed := 0;
      var Elsewhere := 0;
      begin
        for var R := 0 to Retry.Count - 1 do
        begin
          if Aborted then Break;
          Status(Format('Second attempt for unverified occurrences (%d/%d)...',
            [R + 1, Retry.Count]));
          Application.ProcessMessages;
          var VIdx := Retry[R].Key;
          var Cand := ACandidates[Retry[R].Value];
          var Defs2 := AIncludes.Definition(Cand.FilePath, Cand.Line, Cand.Col);
          var Where2 := Format('%s:%d:%d', [ExtractFileName(Cand.FilePath), Cand.Line + 1,
            Cand.Col + 1]);
          if System.Length(Defs2) = 0 then
          begin
            Trace(Where2 + '  second attempt: still no answer');
            Continue;
          end;
          var DF2 := TLspUri.FileUriToPath(Defs2[0].Uri);
          var DL2 := Defs2[0].Range.Start.Line;
          // The trace must say what really happens to the row - it used to
          // print "kept, marked" for rows the same pass then dropped.
          var Verdict2 := 'resolved';
          if not (Anchors.Contains(DF2, DL2) or ALinked.Contains(DF2, DL2)) then
            if ForeignAnswerVerdict(not FNoAnchor, FLspErrors) = favDrop then
              Verdict2 := 'elsewhere -> dropped (another symbol)'
            else
              Verdict2 := 'elsewhere (kept, marked)';
          Trace(Format('%s  second attempt: DelphiLSP -> %s:%d -> %s', [Where2,
            ExtractFileName(DF2), DL2 + 1, Verdict2]));
          var Row := Verified[VIdx];
          if Anchors.Contains(DF2, DL2) then
          begin
            Row.Note := '';
            Verified[VIdx] := Row;
            Inc(Fixed);
          end
          else if ALinked.Contains(DF2, DL2) then
          begin
            Row.Note := '';
            Row.Relation := ALinked.CallLabel(DF2, DL2);
            Verified[VIdx] := Row;
            Inc(Fixed);
          end
          else if ForeignAnswerVerdict(not FNoAnchor, FLspErrors) = favDrop then
            // DROPPED: the session answered every other request of this run
            // without a single abort, and it answers THIS one with another
            // symbol's declaration - the row is foreign. The forum log of
            // 2026-09-30 ends with exactly two such rows
            // ("GlobalConfig.Formulare.BTB" -> UGlobalRomConfig.pas:1586),
            // which the warm run does not list at all.
            Dropped := Dropped + [VIdx]
          else
          begin
            // KEPT (issue #13): with a degraded session (aborted requests)
            // this answer comes from exactly the state we do not trust, and
            // without an anchor there is nothing to compare it against. Say
            // where it led and let the user judge - removing a real
            // reference is the worse error.
            Row.Note := Format('UNVERIFIED - DelphiLSP resolved it to %s:%d',
              [ExtractFileName(DF2), DL2 + 1]);
            Verified[VIdx] := Row;
            Inc(Elsewhere);
          end;
        end;
        var SecondDropped := System.Length(Dropped) - DerivedDropped;
        if SecondDropped > 0 then
          Trace(Format('second attempt: %d occurrence(s) dropped - the answer ' +
            'names another symbol''s declaration and the session reported no ' +
            'aborted request', [SecondDropped]));
        if (Fixed > 0) or (Elsewhere > 0) or (SecondDropped > 0) then
          FSecondPassNote := Format(
            'second attempt: %d of %d unverified occurrence(s) resolved, %d ' +
            'pointed elsewhere (kept, marked), %d dropped (another symbol)',
            [Fixed, Retry.Count, Elsewhere, SecondDropped]);
      end;
    end;

    // BOTH post-passes removed rows, and their indices interleave - sort
    // DESCENDING and delete once, or a deletion shifts the ones still to come.
    if System.Length(Dropped) > 0 then
    begin
      var DropSet := TDictionary<Integer, Boolean>.Create;
      try
        for var D in Dropped do DropSet.AddOrSetValue(D, True);
        // walk the ROWS from the end: every deletion only shifts indices
        // above it, which are already done
        for var R := Verified.Count - 1 downto 0 do
          if DropSet.ContainsKey(R) then Verified.Delete(R);
      finally
        DropSet.Free;
      end;
    end;

    Progress(System.Length(ACandidates), System.Length(ACandidates));
    Result := Verified.ToArray;
  finally
    Contents.Free;
    Synced.Free;
    Retry.Free;
    AnsFile.Free;
    AnsLine.Free;
    Verified.Free;
  end;
end;

end.
