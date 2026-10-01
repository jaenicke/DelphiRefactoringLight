(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Standalone.MainForm;

// Standalone Refactoring Light application - main form.
//
// Layout is in Standalone.MainForm.dfm so the form can be edited in the
// Delphi form designer. Layout regions:
//   - Menu bar with File / Refactor
//   - Left: TTreeView of project source files
//   - Right: TMemo with the active file's content (v1; v2: TSynEdit)
//   - Bottom: TStatusBar
//
// Each Refactor menu item instantiates the matching wizard class,
// calls its Execute method, and frees it. The wizard reaches the
// editor only via the IEditorHelper interface, so it sees this
// form's project state and active file - no IDE required.

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.SyncObjs,
  Winapi.Messages,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.Menus, Vcl.Dialogs,
  Expert.EditorHelperIntf,
  Standalone.EditorHelper;

type
  TMainForm = class(TForm)
    Tree: TTreeView;
    Splitter1: TSplitter;
    Pages: TPageControl;
    TabEditor: TTabSheet;
    Memo: TMemo;
    TabLspLog: TTabSheet;
    LspLog: TMemo;
    LspLogToolbar: TPanel;
    BtnLspLogClear: TButton;
    ChkLspLogAutoscroll: TCheckBox;
    LblLspLogHint: TLabel;
    ProgressPanel: TPanel;
    ProgressLabel: TLabel;
    ProgressBar: TProgressBar;
    LspLogFlush: TTimer;
    Status: TStatusBar;
    MainMenu1: TMainMenu;

    // File menu
    MenuFile: TMenuItem;
    MenuFileOpen: TMenuItem;
    MenuFileSave: TMenuItem;
    MenuFileSep: TMenuItem;
    MenuFileExit: TMenuItem;

    // Refactor menu
    MenuRefactor: TMenuItem;
    MenuRename: TMenuItem;
    MenuFindRef: TMenuItem;
    MenuFindImp: TMenuItem;
    MenuAlignSig: TMenuItem;
    MenuExtractMethod: TMenuItem;
    MenuCompletion: TMenuItem;
    MenuSep1: TMenuItem;
    MenuRemoveWith: TMenuItem;
    MenuRwCursor: TMenuItem;
    MenuRwCurrent: TMenuItem;
    MenuRwSelected: TMenuItem;
    MenuRwProject: TMenuItem;
    MenuMoveToUnit: TMenuItem;
    MenuUnitRefs: TMenuItem;
    MenuFindUnit: TMenuItem;
    MenuShowFixes: TMenuItem;
    MenuAddUnitCursor: TMenuItem;
    MenuResolveMissing: TMenuItem;
    MenuUsesCleanup: TMenuItem;
    MenuFindOriginal: TMenuItem;
    MenuSep2: TMenuItem;
    MenuIface: TMenuItem;
    MenuIfaceExtract: TMenuItem;
    MenuIfaceAdd: TMenuItem;
    MenuIfaceImpl: TMenuItem;
    MenuSep3: TMenuItem;
    MenuSep4: TMenuItem;
    MenuChecks: TMenuItem;
    MenuCheckDfm: TMenuItem;
    MenuCheckGuids: TMenuItem;
    MenuCheckCircular: TMenuItem;
    MenuSemRep: TMenuItem;
    MenuSemRepCurrent: TMenuItem;
    MenuSemRepSelected: TMenuItem;
    MenuSemRepProject: TMenuItem;
    MenuSemRepEditRules: TMenuItem;

    LspPoll: TTimer;

    // Form events
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

    // File handlers
    procedure DoFileOpenProject(Sender: TObject);
    procedure DoFileSave(Sender: TObject);
    procedure DoFileExit(Sender: TObject);

    // Layout handlers
    procedure DoTreeClick(Sender: TObject);
    procedure DoMemoChange(Sender: TObject);
    procedure DoMemoClick(Sender: TObject);
    procedure DoMemoKeyPress(Sender: TObject; var Key: Char);
    procedure DoMemoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoMemoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoLspPollTick(Sender: TObject);
    procedure DoLspLogClear(Sender: TObject);
    procedure DoLspLogFlushTick(Sender: TObject);
    procedure DoPagesChange(Sender: TObject);

    // Refactor handlers
    procedure DoRefactorRename(Sender: TObject);
    procedure DoRefactorFindReferences(Sender: TObject);
    procedure DoRefactorFindImplementations(Sender: TObject);
    procedure DoRefactorAlignSignature(Sender: TObject);
    procedure DoRefactorExtractMethod(Sender: TObject);
    procedure DoRefactorCompletion(Sender: TObject);
    procedure DoRefactorRemoveWithCursor(Sender: TObject);
    procedure DoRefactorRemoveWithCurrent(Sender: TObject);
    procedure DoRefactorRemoveWithSelected(Sender: TObject);
    procedure DoRefactorRemoveWithProject(Sender: TObject);
    procedure DoRefactorMoveToUnit(Sender: TObject);
    procedure DoRefactorUnitRefs(Sender: TObject);
    procedure DoFindUnit(Sender: TObject);
    procedure DoAddUnitAtCursor(Sender: TObject);
    procedure DoShowQuickFixes(Sender: TObject);
    procedure DoResolveMissingUnits(Sender: TObject);
    procedure DoUsesCleanup(Sender: TObject);
    procedure DoFindOriginalSymbol(Sender: TObject);
    procedure DoRefactorExtractInterface(Sender: TObject);
    procedure DoRefactorAddToInterface(Sender: TObject);
    procedure DoRefactorAddIInterface(Sender: TObject);
    procedure DoRefactorSemanticEditRules(Sender: TObject);
    procedure DoCheckDfmEvents(Sender: TObject);
    procedure DoCheckInterfaceGuids(Sender: TObject);
    procedure DoCheckCircularRefs(Sender: TObject);
    procedure DoRefactorSemanticProject(Sender: TObject);
    procedure DoRefactorSemanticCurrent(Sender: TObject);
    procedure DoRefactorSemanticSelected(Sender: TObject);
  private
    FState: TStandaloneProjectState;
    /// <summary>The file whose text the Memo shows. Typing and Save go
    ///  to this file, not to FState.ActiveFile: the two used to drift
    ///  apart (a file switch, a GotoLocation) and the Memo's text then
    ///  landed in the other file's buffer or on its disk.</summary>
    FMemoFile: string;
    /// <summary>True while LoadFileIntoEditor puts a file into the Memo.
    ///  The load fires OnChange, which must not treat it as typing.</summary>
    FLoadingMemo: Boolean;
    /// <summary>Last printable character the user typed. Captured in
    ///  OnKeyPress, consumed (and cleared) in OnKeyUp; lets us
    ///  distinguish a real `.` keystroke from caret movement that
    ///  happens to land next to an existing `.` in the buffer.</summary>
    FLastKeyChar: Char;
    /// <summary>Re-entrancy guard for the auto-trigger. A completion
    ///  call shows a modal popup and ProcessMessages while it waits
    ///  for LSP - without the flag a second `.` typed during that
    ///  window would queue a recursive trigger.</summary>
    FCompletionInFlight: Boolean;
    /// <summary>One-shot suppression flag for the OnKeyPress that
    ///  follows the OnKeyDown we intercepted (e.g. Enter to accept a
    ///  completion). Setting Key := 0 in OnKeyDown is documented to
    ///  also suppress OnKeyPress, but TMemo's Enter handling in
    ///  particular sometimes still emits the #13 char; this flag is
    ///  the belt-and-suspenders that keeps the newline out of the
    ///  buffer.</summary>
    FConsumeNextChar: Boolean;
    /// <summary>True while the prewarm background thread is alive.
    ///  Polled by the status timer so the UI can show "starting" while
    ///  LSP is still cold.</summary>
    FLspPrewarming: Boolean;
    /// <summary>The prewarm worker, kept so FormDestroy can WAIT for it
    ///  (audit #33, M35). It was started and forgotten: closing the window
    ///  while a project was still indexing shut DelphiLsp down under the
    ///  worker, whose request then failed into an except branch that writes
    ///  the log - on a lock and a buffer FormDestroy was already freeing.
    ///  </summary>
    FPrewarmThread: TThread;
    /// <summary>A project opened while a prewarm was still running used to
    ///  be dropped by the FLspPrewarming guard - the same defect the IDE
    ///  side had (audit #38, M4). Parked here and started when the worker
    ///  finishes.</summary>
    FPrewarmPending: Boolean;
    /// <summary>The tree's file list. The node's Data used to be a
    ///  StrNew'd PChar that nothing ever disposed, so every RefreshTree -
    ///  one per refactoring - leaked the whole file list (audit #33,
    ///  L4-sa). Data is an INDEX into this array now (1-based, so nil
    ///  stays distinguishable).</summary>
    FTreeFiles: TArray<string>;
    /// <summary>Files the user typed in and has not saved. The README
    ///  claimed edits were flushed to disk on every change, which was
    ///  never true (audit #33, M34): they live in FState's buffer only, so
    ///  a file switch reloaded from disk and overwrote them. Now the
    ///  buffer is what a switch loads, and this set is what the close
    ///  prompt asks about.</summary>
    FDirty: TStringList;
    /// <summary>True after we have already hooked the LSP client's
    ///  OnLog callback in this session - prevents double-attaching when
    ///  StartLspPrewarm runs more than once per project.</summary>
    FLspLogHooked: Boolean;
    /// <summary>Producer buffer for LSP-log lines. The reader thread
    ///  (and the prewarm thread) push into this under a lock; the UI
    ///  thread drains it every 250 ms via DoLspLogFlushTick. Without
    ///  the batching, TThread.Queue + Memo.Lines.Add per JSON-RPC
    ///  message visibly slows the LSP cold start - a project-wide
    ///  index push triggers thousands of publishDiagnostics in a few
    ///  seconds.</summary>
    FLspLogBuffer: TStringList;
    FLspLogLock: TCriticalSection;
    /// <summary>True only while the user has the LSP Diagnostics tab
    ///  active. When false, AppendLspLog returns immediately and the
    ///  OnLog hook in AttachLspLog also short-circuits. The point: on
    ///  a large project the LSP server bursts thousands of
    ///  publishDiagnostics messages per second during cold start, and
    ///  formatting + buffering them ate enough main-thread budget to
    ///  visibly slow indexing and freeze the editor afterwards.</summary>
    FLspLogActive: Boolean;
    /// <summary>Pending progress update. The prewarm thread overwrites
    ///  the fields under FLspLogLock; the flush tick reads them and
    ///  pushes to the ProgressBar. Replaces a per-file TThread.Queue
    ///  call, which on a 1247-file sweep meant 1247 main-thread wakes.
    ///  FProgressDirty flips on every write, off when the tick
    ///  consumes the value.</summary>
    FProgressCur, FProgressTotal: Integer;
    FProgressFile: string;
    FProgressDirty, FProgressDone: Boolean;
    /// <summary>When the prewarm started. Used to compute elapsed-time
    ///  progress during the LSP-side index wait, when the worker thread
    ///  is blocked in a synchronous Client.GetDocumentSymbols call and
    ///  cannot fire the AProgress callback for 10-30 seconds at a
    ///  stretch. The poll tick reads this and bumps the bar visually.</summary>
    FPrewarmStart: TDateTime;
    /// <summary>Timestamp of the last AProgress callback received from
    ///  the worker. If no callback arrived in the last ~750 ms while
    ///  the prewarm is still in flight, the poll tick takes over and
    ///  drives the progress bar from FPrewarmStart instead.</summary>
    FLastProgressTime: TDateTime;
    /// <summary>Persistent completion wizard. The popup it shows is
    ///  non-modal: Execute returns BEFORE the user picks an item.
    ///  Freeing the wizard between trigger and pick would leave the
    ///  popup pointing at a dangling DoInsert. The IDE plugin
    ///  side-steps this because Expert.Registration keeps a singleton;
    ///  in standalone we keep it on the form for the same lifetime.</summary>
    FCompletionWizard: TObject;
    /// <summary>Same singleton pattern for the signature-help popup
    ///  (parameter hint on `(`). The hover roundtrip is async so the
    ///  wizard must outlive its trigger; reusing one instance also
    ///  means the popup itself persists between triggers (cheaper
    ///  than recreating the TForm every keystroke).</summary>
    FSignatureWizard: TObject;
    /// <summary>Marshals (current, total, filename) progress updates
    ///  from the prewarm thread to the UI. Replaces direct VCL access
    ///  from inside the callback - safe because the reference is
    ///  captured into the anonymous progress callback at start.</summary>
    procedure ApplyProgress(ACurrent, ATotal: Integer; const ACurrentFile: string);
    function TreeFileOf(ANode: TTreeNode): string;
    procedure MarkDirty(const AFile: string; AValue: Boolean);
    function DirtyList: string;
    procedure HandleCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure RefreshTree;
    procedure LoadFileIntoEditor(const AFile: string);
    procedure ReloadActiveFile;
    /// <summary>Selects ALine / ACol (1-based) in the Memo, plus
    ///  AHighlightLen characters, and scrolls it into view.</summary>
    procedure PlaceMemoCaret(ALine, ACol, AHighlightLen: Integer);
    /// <summary>FState.OnNavigate, i.e. a wizard's Editor.GotoLocation:
    ///  shows AFile in the Memo and places the caret there.</summary>
    procedure HandleNavigate(const AFile: string; ALine, ACol, AHighlightLen: Integer);
    procedure UpdateStatusBar;
    /// <summary>Returns the identifier word the caret sits inside / at
    ///  the end of - same logic as the completion wizard's
    ///  GetCurrentWordPrefix, but driven off the Memo so the host
    ///  form can update the popup filter live as the user types.</summary>
    function ComputeCurrentMemoWordPrefix: string;
    /// <summary>Screen-Y of the top pixel of line ALine (1-based) in
    ///  the Memo. Used by the signature wizard to position the popup
    ///  above the line that contains the call's `(`, not above
    ///  whichever argument line the caret currently sits in.</summary>
    function GetMemoLineScreenY(ALine: Integer): Integer;
    /// <summary>True when the completion popup is open AND the caret
    ///  is still in an identifier context. Hides the popup otherwise.</summary>
    procedure RefreshCompletionFilter;
    /// <summary>Kicks off a background thread that starts the LSP for
    ///  the currently loaded project and pushes every project source
    ///  file through textDocument/didOpen. Idempotent: a second call
    ///  while the prewarm is in flight is a no-op.</summary>
    procedure StartLspPrewarm;
    /// <summary>Hooks FClient.OnLog so every JSON-RPC message is
    ///  appended to the LSP Diagnostics tab. Called from the prewarm
    ///  thread but marshals its append back to the main thread via
    ///  TThread.Queue so VCL access stays single-threaded.</summary>
    procedure AttachLspLog;
    /// <summary>Appends a line to the LSP Diagnostics memo. Safe to
    ///  call from any thread. If the user has the autoscroll checkbox
    ///  ticked, scrolls to the bottom afterwards.</summary>
    procedure AppendLspLog(const ALine: string);
    /// <summary>Wraps a wizard invocation with a wait cursor + UI guard
    ///  and refreshes the editor / tree once the wizard returns. Without
    ///  the refresh, an Extract Interface (or any other write through
    ///  Editor.ReplaceFileContent) would not show up in the Memo until
    ///  the user clicks the tree node again. The wait cursor + Enabled
    ///  := False makes the (still synchronous) blocking phase visually
    ///  obvious and stops accidental double-clicks.</summary>
    procedure RunWizard(AProc: TProc);
  end;

var
  MainForm: TMainForm;

implementation

uses
  Winapi.Windows, System.DateUtils, System.UITypes,
  Lsp.Client, Expert.DialogHelper,
  Expert.SemanticReplaceWizard,
  Expert.RenameWizard,
  Expert.FindReferencesWizard,
  Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard,
  Expert.WithRefactorWizard,
  Expert.MoveToUnitWizard,
  Expert.UnitReferencesWizard,
  Expert.ExtractInterfaceWizard,
  Expert.ExtractMethod,
  Expert.CompletionWizard,
  Expert.SignatureHelpWizard,
  Expert.DfmEventCheckDialog,
  Expert.InterfaceGuidDialog,
  Expert.CircularRefsDialog,
  Expert.FindUnitDialog,
  Expert.AutoImport,
  Expert.FindOriginalSymbolWizard,
  Expert.UsesCleanup,
  Expert.LspManager;

{$R *.dfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FLspLogBuffer := TStringList.Create;
  FLspLogLock := TCriticalSection.Create;
  FDirty := TStringList.Create;
  FDirty.CaseSensitive := False;
  FDirty.Sorted := True;
  FDirty.Duplicates := dupIgnore;
  OnCloseQuery := HandleCloseQuery;
  FLspLogActive := Pages.ActivePage = TabLspLog;
  LspLogFlush.Enabled := True;
  FState := TStandaloneProjectState.Create;
  SetEditorImpl(TStandaloneEditorHelper.Create(FState));
  FCompletionWizard := TLspCompletionWizard.Create;
  FSignatureWizard := TLspSignatureHelpWizard.Create;
  // Reload the Memo whenever a wizard writes to the active file from
  // outside the form. Covers the async case: the completion popup is
  // non-modal, so the wizard's Execute returns BEFORE the user picks
  // an item; without this notification, the popup's DoInsert would
  // silently update the buffer and only the next tab-switch would
  // show the change.
  FState.OnExternalChange :=
    procedure(AFile: string)
    begin
      if (FState.ActiveFile <> '') and SameText(AFile, FState.ActiveFile) then
        ReloadActiveFile;
    end;

  // A wizard's Editor.GotoLocation (Circular References, Blame, Auto
  // Import, ...) moves the active file; show that file in the Memo too.
  FState.OnNavigate :=
    procedure(const AFile: string; ALine, ACol, AHighlightLen: Integer)
    begin
      HandleNavigate(AFile, ALine, ACol, AHighlightLen);
    end;

  // Install the selection callback the standalone IEditorHelper uses
  // to answer Editor.GetSelection. Extract Method calls into this.
  FState.GetSelectionFunc :=
    function: TStandaloneSelection
    var
      SelStart, SelLen, P: Integer;
      Text: string;
    begin
      Result := Default(TStandaloneSelection);
      if FState.ActiveFile = '' then Exit;
      SelStart := Memo.SelStart;
      SelLen := Memo.SelLength;
      if SelLen <= 0 then Exit;
      Result.HasSelection := True;
      Result.FileName := FState.ActiveFile;
      Text := Memo.Text;
      Result.Text := Copy(Text, SelStart + 1, SelLen);
      // Compute 1-based (line, col) for the start and end of the selection.
      Result.StartLine := Memo.Perform(EM_LINEFROMCHAR, SelStart, 0) + 1;
      P := Memo.Perform(EM_LINEINDEX, Result.StartLine - 1, 0);
      Result.StartCol := SelStart - P + 1;
      Result.EndLine := Memo.Perform(EM_LINEFROMCHAR, SelStart + SelLen, 0) + 1;
      P := Memo.Perform(EM_LINEINDEX, Result.EndLine - 1, 0);
      Result.EndCol := (SelStart + SelLen) - P + 1;
    end;

  // Live auto-import indicator (idles until an LSP client is up).
  StartAutoImportLive;
  // Warning dialog the "remove unused private member" fix asks
  // before it deletes a non-empty body.
  InstallRemovePrivateConfirm;

  UpdateStatusBar;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  StopAutoImportLive;
  LspPoll.Enabled := False;
  LspLogFlush.Enabled := False;
  try
    TLspManager.Instance.Shutdown;
  except
    // Best-effort shutdown - never throw out of FormDestroy.
  end;
  // AFTER the shutdown (which makes the worker's pending LSP call fail) and
  // BEFORE the log lock and buffer are freed: the worker's except branch
  // writes into both (audit #33, M35). Bounded - a worker that will not
  // come back must not keep the window open for ever.
  if FPrewarmThread <> nil then
  begin
    try
      FPrewarmThread.WaitFor;
    except
      // a thread that already ended raises nothing we can act on
    end;
    FreeAndNil(FPrewarmThread);
  end;
  FCompletionWizard.Free;
  FSignatureWizard.Free;
  SetEditorImpl(nil);
  FState.Free;
  FDirty.Free;
  FLspLogLock.Free;
  FLspLogBuffer.Free;
end;

procedure TMainForm.HandleCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  // There is no auto-flush (see FDirty): without this, closing the window
  // threw typed edits away without a word.
  CanClose := True;
  if FDirty.Count = 0 then Exit;
  case MessageDlg(Format('%d file(s) have unsaved changes:'#13#10#13#10 +
    '%s'#13#10#13#10'Save them before closing?', [FDirty.Count, DirtyList]),
    mtWarning, [mbYes, mbNo, mbCancel], 0) of
    mrYes:
      begin
        // the shown file first - its text lives in the Memo, not in a buffer
        DoFileSave(nil);
        for var I := FDirty.Count - 1 downto 0 do
        begin
          var F := FDirty[I];
          var Content := '';
          if FState.TryGetBuffer(F, Content) then
            try Editor.ReplaceFileContent(F, Content); except end;
        end;
        FDirty.Clear;
      end;
    mrCancel: CanClose := False;
  end;
end;

function TMainForm.TreeFileOf(ANode: TTreeNode): string;
var
  Idx: NativeInt;
begin
  Result := '';
  if ANode = nil then Exit;
  Idx := NativeInt(ANode.Data);          // 1-based, 0 = nothing stored
  if (Idx >= 1) and (Idx <= Length(FTreeFiles)) then
    Result := FTreeFiles[Idx - 1];
end;

procedure TMainForm.MarkDirty(const AFile: string; AValue: Boolean);
var
  Idx: Integer;
begin
  if AFile = '' then Exit;
  if AValue then
    FDirty.Add(AFile)
  else
  begin
    Idx := FDirty.IndexOf(AFile);
    if Idx >= 0 then FDirty.Delete(Idx);
  end;
  UpdateStatusBar;
end;

function TMainForm.DirtyList: string;
begin
  Result := '';
  for var I := 0 to FDirty.Count - 1 do
  begin
    if I = 8 then Exit(Result + '  ...');
    Result := Result + '  ' + ExtractFileName(FDirty[I]) + #13#10;
  end;
end;

procedure TMainForm.RefreshTree;
begin
  Tree.Items.BeginUpdate;
  try
    Tree.Items.Clear;
    // The INDEX goes into Data (1-based), never a StrNew'd pointer: nothing
    // ever disposed those, and RefreshTree runs after every refactoring.
    FTreeFiles := FState.SourceFiles;
    for var I := 0 to High(FTreeFiles) do
      Tree.Items.AddObject(nil, ExtractFileName(FTreeFiles[I]),
        Pointer(NativeInt(I + 1)));
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure TMainForm.LoadFileIntoEditor(const AFile: string);
var
  Buffered: string;
begin
  if not TFile.Exists(AFile) then Exit;
  // Before showing a file, SAVE what the Memo currently holds into its own
  // buffer: a switch away from a file the user typed in used to lose the
  // edits, because only disk was ever loaded back (audit #33, M34).
  if (FMemoFile <> '') and not SameText(FMemoFile, AFile) then
    FState.UpdateBuffer(FMemoFile, Memo.Lines.Text);
  // LoadFromFile fires OnChange while the state still names the previous
  // file, and DoMemoChange stored the NEW file's text as the PREVIOUS
  // file's buffer - the next wizard on that file wrote it to disk.
  FLoadingMemo := True;
  try
    // The BUFFER is the truth when it differs from disk - that is where an
    // unsaved edit lives, and a wizard works on it too.
    if (FDirty.IndexOf(AFile) >= 0) and FState.TryGetBuffer(AFile, Buffered) then
      Memo.Lines.Text := Buffered
    else
      Memo.Lines.LoadFromFile(AFile);
  finally
    FLoadingMemo := False;
  end;
  FMemoFile := AFile;
  FState.SetActiveFile(AFile, 1, 1);
  FState.UpdateBuffer(AFile, Memo.Lines.Text);
  UpdateStatusBar;
end;

procedure TMainForm.PlaceMemoCaret(ALine, ACol, AHighlightLen: Integer);
begin
  if (ALine < 1) or (ALine > Memo.Lines.Count) then Exit;
  Memo.SelStart := Memo.Perform(EM_LINEINDEX, ALine - 1, 0) + (ACol - 1);
  if AHighlightLen > 0 then
    Memo.SelLength := AHighlightLen
  else
    Memo.SelLength := 0;
  Memo.Perform(EM_SCROLLCARET, 0, 0);
end;

procedure TMainForm.HandleNavigate(const AFile: string; ALine, ACol, AHighlightLen: Integer);
var
  Line, Col: Integer;
begin
  if not SameText(AFile, FMemoFile) then
    LoadFileIntoEditor(AFile);
  if SameText(AFile, FMemoFile) then
  begin
    FState.SetActiveFile(FMemoFile, ALine, ACol);
    PlaceMemoCaret(ALine, ACol, AHighlightLen);
  end
  else
  begin
    // The target could not be loaded: point the state back at the file
    // the Memo shows, so the wizards and Save mean the same file.
    Line := Memo.Perform(EM_LINEFROMCHAR, Memo.SelStart, 0) + 1;
    Col := Memo.SelStart - Memo.Perform(EM_LINEINDEX, Line - 1, 0) + 1;
    FState.SetActiveFile(FMemoFile, Line, Col);
  end;
  UpdateStatusBar;
end;

procedure TMainForm.ReloadActiveFile;
// Re-syncs the Memo with the active file's current content. Source of
// truth is the standalone state's buffer (wizards write through that);
// if no buffer exists - e.g. the wizard skipped this file - we fall
// back to disk so an external edit still shows up.
//
// We preserve the cursor's logical (line, col) instead of restoring
// the raw SelStart, because the file's line count likely changed.
var
  NewText: string;
  SavedLine, SavedCol: Integer;
  LineOffset: Integer;
  TopLine: Integer;
begin
  if FState.ActiveFile = '' then Exit;
  if not FState.TryGetBuffer(FState.ActiveFile, NewText) then
  begin
    if not TFile.Exists(FState.ActiveFile) then Exit;
    NewText := TFile.ReadAllText(FState.ActiveFile);
  end;
  if Memo.Lines.Text = NewText then Exit;

  SavedLine := FState.ActiveLine;
  SavedCol := FState.ActiveCol;
  TopLine := Memo.Perform(EM_GETFIRSTVISIBLELINE, 0, 0);

  Memo.Lines.BeginUpdate;
  try
    Memo.Text := NewText;
  finally
    Memo.Lines.EndUpdate;
  end;

  // Restore caret + top-of-view (best effort: file may have grown or shrunk).
  if (SavedLine >= 1) and (SavedLine <= Memo.Lines.Count) then
  begin
    LineOffset := Memo.Perform(EM_LINEINDEX, SavedLine - 1, 0);
    Memo.SelStart := LineOffset + (SavedCol - 1);
    Memo.SelLength := 0;
  end;
  Memo.Perform(EM_LINESCROLL, 0,
    TopLine - Memo.Perform(EM_GETFIRSTVISIBLELINE, 0, 0));

  // Keep the state's buffer in sync with what's now in the editor; this
  // prevents the next click on the Memo (which fires OnChange via
  // UpdateBuffer) from overwriting the wizard's edit with a possibly
  // already-stale value.
  FState.UpdateBuffer(FState.ActiveFile, NewText);
end;

procedure TMainForm.DoFileOpenProject(Sender: TObject);
var
  Dlg: TOpenDialog;
begin
  Dlg := TOpenDialog.Create(Self);
  try
    Dlg.Filter := 'Delphi project (*.dproj)|*.dproj|Delphi project source (*.dpr)|*.dpr';
    Dlg.DefaultExt := 'dproj';
    if not Dlg.Execute then Exit;
    // A new project means the old LSP session is stale - drop it so the
    // next prewarm starts fresh against the new .delphilsp.json and the
    // new project root.
    TLspManager.Instance.Reset;
    FState.LoadFromDproj(Dlg.FileName);
    RefreshTree;
    if Length(FState.SourceFiles) > 0 then
      LoadFileIntoEditor(FState.SourceFiles[0]);
    UpdateStatusBar;
    StartLspPrewarm;
  finally
    Dlg.Free;
  end;
end;

procedure TMainForm.StartLspPrewarm;
var
  RootPath, DelphiLspJson, ProjectFile: string;
  ScanFiles: TArray<string>;
begin
  if FLspPrewarming then
  begin
    FPrewarmPending := True;   // not dropped - run when this one finishes
    Exit;
  end;
  RootPath := Editor.GetProjectRoot;
  if RootPath = '' then Exit;
  ProjectFile := Editor.GetCurrentProjectDproj;
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    Status.Panels[3].Text := 'LSP not started - no .delphilsp.json next to .dproj';
    AppendLspLog('[error] no .delphilsp.json found next to the .dproj. ' +
      'Place one (see DelphiRefactoringLight.delphilsp.json for an example) ' +
      'and re-open the project.');
    Exit;
  end;
  ScanFiles := Editor.GetProjectSourceFiles;
  if Length(ScanFiles) = 0 then Exit;

  FLspPrewarming := True;
  FPrewarmPending := False;
  // A worker of a PREVIOUS project may still be shutting down - its object
  // is ours to free, and FormDestroy waits on whatever is here.
  if (FPrewarmThread <> nil) and FPrewarmThread.Finished then
    FreeAndNil(FPrewarmThread);
  FLspLogHooked := False;
  FPrewarmStart := Now;
  FLastProgressTime := Now;
  LspPoll.Enabled := True;
  Status.Panels[3].Text := 'LSP: launching DelphiLsp.exe...';
  AppendLspLog(Format('[prewarm] launching DelphiLsp for %s (%d source files)',
    [ExtractFileName(ProjectFile), Length(ScanFiles)]));

  FPrewarmThread := TThread.CreateAnonymousThread(
    procedure
    var
      StartT: TDateTime;
    begin
      StartT := Now;
      try
        try
          // Lower priority so the LSP cold-start does not steal cycles
          // from the editor UI.
          SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_BELOW_NORMAL);
          TLspManager.Instance.GetClient(RootPath, ProjectFile, DelphiLspJson);
          AttachLspLog;
          AppendLspLog(Format('[prewarm] LSP client up after %.2fs, starting didOpen sweep...',
            [(Now - StartT) * SecsPerDay]));
          TLspManager.Instance.EnsureProjectIndexed(ScanFiles,
            procedure(ACurrent, ATotal: Integer; const ACurrentFile: string)
            begin
              ApplyProgress(ACurrent, ATotal, ACurrentFile);
            end);
          AppendLspLog(Format('[prewarm] done. ProjectIndexed=%s, total %.2fs',
            [BoolToStr(TLspManager.Instance.ProjectIndexed, True),
             (Now - StartT) * SecsPerDay]));
          TThread.Queue(nil,
            procedure
            begin
              if Assigned(ProgressPanel) then ProgressPanel.Visible := False;
            end);
        except
          on E: Exception do
          begin
            AppendLspLog(Format('[prewarm] FAILED after %.2fs: %s: %s',
              [(Now - StartT) * SecsPerDay, E.ClassName, E.Message]));
            // The panel is only hidden on the success path, so a failure
            // left it frozen on its last caption and the 500 ms poll timer
            // running for the rest of the session (audit #33, L4-sb).
            TThread.Queue(nil,
              procedure
              begin
                if Assigned(ProgressPanel) then ProgressPanel.Visible := False;
                if Assigned(LspPoll) then LspPoll.Enabled := False;
                Status.Panels[3].Text := 'LSP: prewarm failed - see the ' +
                  'LSP Diagnostics tab';
              end);
          end;
        end;
      finally
        FLspPrewarming := False;
        TThread.Queue(nil,
          procedure
          begin
            if FPrewarmPending then
            begin
              FPrewarmPending := False;
              StartLspPrewarm;    // the project that arrived meanwhile
            end;
          end);
      end;
    end);
  FPrewarmThread.FreeOnTerminate := False;   // FormDestroy waits on it
  FPrewarmThread.Start;
end;

procedure TMainForm.AttachLspLog;
var
  RootPath, ProjectFile, DelphiLspJson: string;
  C: TLspClient;
begin
  if FLspLogHooked then Exit;
  RootPath := Editor.GetProjectRoot;
  ProjectFile := Editor.GetCurrentProjectDproj;
  DelphiLspJson := Editor.FindDelphiLspJson;
  // GetClient is idempotent for the current project; this returns the
  // existing client without restarting it.
  try
    C := TLspManager.Instance.GetClient(RootPath, ProjectFile, DelphiLspJson);
  except
    Exit;
  end;
  if C = nil then Exit;
  C.Verbose := True;
  C.OnLog :=
    procedure(const ADirection, AMethod, ABody: string)
    var
      Body: string;
    begin
      // Cheap path: if the user is not looking at the LSP tab, do
      // nothing. Saves a Format + a Copy + potentially-long body
      // string operations per JSON-RPC message - thousands per
      // second on cold start.
      if not FLspLogActive then Exit;
      // Bodies can be large (hundreds of KB on a project-wide
      // documentSymbol response). Truncate to keep the memo usable
      // for browsing.
      Body := ABody;
      if Length(Body) > 800 then
        Body := Copy(Body, 1, 800) + '... [truncated, ' +
          IntToStr(Length(ABody)) + ' bytes total]';
      AppendLspLog(Format('%s  %s  %s', [ADirection, AMethod, Body]));
    end;
  FLspLogHooked := True;
end;

procedure TMainForm.AppendLspLog(const ALine: string);
// Always buffers - this is the path the prewarm thread uses for
// status notes ("[prewarm] launching", "[prewarm] done", error
// traces). The noisy high-volume path (OnLog from the LSP client)
// gates separately on FLspLogActive so we never pay the Format cost
// for traffic the user is not looking at.
//
// We cap the buffer at 2000 lines to bound memory if a long session
// stays on the Editor tab forever - the oldest entries get dropped.
var
  Line: string;
begin
  Line := '[' + FormatDateTime('hh:nn:ss.zzz', Now) + '] ' + ALine;
  FLspLogLock.Enter;
  try
    FLspLogBuffer.Add(Line);
    while FLspLogBuffer.Count > 2000 do
      FLspLogBuffer.Delete(0);
  finally
    FLspLogLock.Leave;
  end;
end;

procedure TMainForm.DoPagesChange(Sender: TObject);
begin
  FLspLogActive := Pages.ActivePage = TabLspLog;
end;

procedure TMainForm.DoLspLogFlushTick(Sender: TObject);
// Main-thread tick. Two jobs:
//   1) Drain whatever AppendLspLog buffered into LspLog.Lines, in one
//      BeginUpdate / EndUpdate block. Cap per tick at 500 lines so a
//      burst (DelphiLSP sometimes dumps 5000 publishDiagnostics in
//      one second) does not lock up the form for half a second.
//   2) Push the latest pending progress values into the bar / label.
const
  MaxPerTick = 500;
var
  Batch: TStringList;
  ToDrop, ProgCur, ProgTotal: Integer;
  ProgFile, Cap: string;
  ProgDirty, ProgDone: Boolean;
  BarPos: Integer;
begin
  Batch := TStringList.Create;
  try
    FLspLogLock.Enter;
    try
      // Drain at most MaxPerTick lines so a burst does not freeze us.
      // But: only if the LSP tab is visible. Otherwise leave the buffer
      // intact (capped to 2000 in AppendLspLog) so the user sees the
      // recent tail when they switch in.
      if FLspLogActive then
        while (Batch.Count < MaxPerTick) and (FLspLogBuffer.Count > 0) do
        begin
          Batch.Add(FLspLogBuffer[0]);
          FLspLogBuffer.Delete(0);
        end;
      ProgDirty := FProgressDirty;
      ProgDone := FProgressDone;
      ProgCur := FProgressCur;
      ProgTotal := FProgressTotal;
      ProgFile := FProgressFile;
      FProgressDirty := False;
      FProgressDone := False;
    finally
      FLspLogLock.Leave;
    end;

    // ---- progress flush ----
    if ProgDirty then
    begin
      if ProgDone then
        ProgressPanel.Visible := False
      else if ProgTotal > 0 then
      begin
        if Pos('Waiting', ProgFile) > 0 then
          Cap := Format('LSP: %s', [ProgFile])
        else if ProgCur < ProgTotal then
          Cap := Format('LSP indexing %d / %d  -  %s',
            [ProgCur, ProgTotal, ExtractFileName(ProgFile)])
        else
          Cap := Format('LSP: %s', [ProgFile]);
        if ProgCur > ProgTotal then BarPos := ProgTotal else BarPos := ProgCur;
        ProgressPanel.Visible := True;
        ProgressBar.Max := ProgTotal;
        ProgressBar.Position := BarPos;
        ProgressLabel.Caption := Cap;
      end;
    end;

    // ---- log flush ----
    if Batch.Count = 0 then Exit;
    LspLog.Lines.BeginUpdate;
    try
      LspLog.Lines.AddStrings(Batch);
      ToDrop := LspLog.Lines.Count - 5000;
      while ToDrop > 0 do
      begin
        LspLog.Lines.Delete(0);
        Dec(ToDrop);
      end;
    finally
      LspLog.Lines.EndUpdate;
    end;
    if ChkLspLogAutoscroll.Checked then
    begin
      LspLog.SelStart := Length(LspLog.Text);
      LspLog.Perform(EM_SCROLLCARET, 0, 0);
    end;
  finally
    Batch.Free;
  end;
end;

procedure TMainForm.DoLspLogClear(Sender: TObject);
begin
  FLspLogLock.Enter;
  try
    FLspLogBuffer.Clear;
  finally
    FLspLogLock.Leave;
  end;
  LspLog.Lines.Clear;
end;

procedure TMainForm.ApplyProgress(ACurrent, ATotal: Integer; const ACurrentFile: string);
// Called from the prewarm thread (potentially thousands of times
// during the didOpen sweep). We avoid TThread.Queue per call - on a
// 1247-file project that was 1247 main-thread wake-ups and visibly
// slowed indexing. Instead we just stash the latest values; the
// flush tick on the UI thread reads them at 250 ms cadence, which is
// enough for visual feedback.
begin
  FLspLogLock.Enter;
  try
    FProgressCur := ACurrent;
    FProgressTotal := ATotal;
    FProgressFile := ACurrentFile;
    FProgressDirty := True;
    FLastProgressTime := Now;
    if (ATotal > 0) and (ACurrent >= ATotal) and (ACurrentFile = '') then
      FProgressDone := True;
  finally
    FLspLogLock.Leave;
  end;
end;

procedure TMainForm.DoLspPollTick(Sender: TObject);
// Two jobs (every 500 ms while prewarm is in flight):
//   1) Pull the LSP manager's status line into the 4th status-bar
//      panel.
//   2) When the worker thread is blocked in a synchronous LSP call
//      (typical: GetDocumentSymbols hangs for 5-25 s while the LSP
//      builds its internal index), no AProgress callback fires for
//      that whole period and the bar would freeze at 100 %. We
//      detect "no fresh callback in the last second" and take over,
//      driving a time-based bar from FPrewarmStart.
const
  FallbackAfterSec = 1.0;  // worker silent this long -> we drive
  MaxFallbackSec   = 60.0; // bar saturates here so it does not look idle
var
  ElapsedSinceCallback, ElapsedSinceStart: Double;
  Sec: Integer;
begin
  Status.Panels[3].Text := 'LSP: ' + TLspManager.Instance.GetWarmupStatusLine;

  if FLspPrewarming and not TLspManager.Instance.ProjectIndexed then
  begin
    ElapsedSinceCallback := (Now - FLastProgressTime) * SecsPerDay;
    if ElapsedSinceCallback >= FallbackAfterSec then
    begin
      ElapsedSinceStart := (Now - FPrewarmStart) * SecsPerDay;
      Sec := Round(ElapsedSinceStart);
      if Sec > Round(MaxFallbackSec) then Sec := Round(MaxFallbackSec);
      ProgressPanel.Visible := True;
      ProgressBar.Max := Round(MaxFallbackSec);
      ProgressBar.Position := Sec;
      ProgressLabel.Caption := Format(
        'LSP: building index (%d s) - first refactor will wait for this', [Sec]);
    end;
  end;

  // Stop when the index is there - or when nobody is working on it any
  // more: a failed prewarm never sets ProjectIndexed, so the timer used to
  // tick for the rest of the session (audit #33, L4-sb).
  if TLspManager.Instance.ProjectIndexed
    or (not FLspPrewarming and not TLspManager.Instance.ProjectIndexed) then
    LspPoll.Enabled := False;
end;

procedure TMainForm.DoFileSave(Sender: TObject);
begin
  if FMemoFile = '' then Exit;
  Editor.ReplaceFileContent(FMemoFile, Memo.Lines.Text);
  MarkDirty(FMemoFile, False);
end;

procedure TMainForm.DoFileExit(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.DoTreeClick(Sender: TObject);
var
  N: TTreeNode;
  F: string;
begin
  N := Tree.Selected;
  if N = nil then Exit;
  F := TreeFileOf(N);
  if F = '' then Exit;
  LoadFileIntoEditor(F);
end;

procedure TMainForm.DoMemoChange(Sender: TObject);
begin
  if FLoadingMemo or (FMemoFile = '') then Exit;
  FState.UpdateBuffer(FMemoFile, Memo.Lines.Text);
  MarkDirty(FMemoFile, True);   // the standalone has no auto-flush
  // While the completion popup is visible, keep its filter in sync
  // with what the user is typing in the editor. The popup never has
  // focus, so we drive it from here.
  if TLspCompletionWizard(FCompletionWizard).IsPopupVisible then
    RefreshCompletionFilter;
end;

procedure TMainForm.DoMemoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
// Intercepts navigation / accept / cancel keys when the completion
// popup is showing items. Loading-phase popups do NOT intercept keys
// so the editor stays normally usable while the LSP is still working.
begin
  if not TLspCompletionWizard(FCompletionWizard).IsPopupActive then Exit;
  case Key of
    VK_UP:
      begin TLspCompletionWizard(FCompletionWizard).MoveSelection(-1); Key := 0; end;
    VK_DOWN:
      begin TLspCompletionWizard(FCompletionWizard).MoveSelection(+1); Key := 0; end;
    VK_PRIOR:
      begin TLspCompletionWizard(FCompletionWizard).MoveSelection(-8); Key := 0; end;
    VK_NEXT:
      begin TLspCompletionWizard(FCompletionWizard).MoveSelection(+8); Key := 0; end;
    VK_RETURN, VK_TAB:
      begin
        Key := 0;
        FConsumeNextChar := True;  // swallow the matching OnKeyPress #13 / #9
        TLspCompletionWizard(FCompletionWizard).InsertSelected;
      end;
    VK_ESCAPE:
      begin Key := 0; TLspCompletionWizard(FCompletionWizard).HidePopup; end;
  end;
end;

function TMainForm.GetMemoLineScreenY(ALine: Integer): Integer;
// Maps a 1-based line number to its top-of-line screen Y in pixels.
// Uses EM_LINEINDEX to find the line's first char, then EM_POSFROMCHAR
// to get its client coordinates, then ClientToScreen to convert.
var
  CharIdx: Integer;
  PosDw: DWORD;
  ClientPt: TPoint;
begin
  Result := 0;
  if (ALine < 1) or (ALine > Memo.Lines.Count) then Exit;
  CharIdx := Memo.Perform(EM_LINEINDEX, ALine - 1, 0);
  if CharIdx < 0 then Exit;
  PosDw := DWORD(Memo.Perform(EM_POSFROMCHAR, CharIdx, 0));
  ClientPt.X := SmallInt(LoWord(PosDw));
  ClientPt.Y := SmallInt(HiWord(PosDw));
  ClientPt := Memo.ClientToScreen(ClientPt);
  Result := ClientPt.Y;
end;

function TMainForm.ComputeCurrentMemoWordPrefix: string;
var
  LineNo, ColPos, Col, Start: Integer;
  LineText: string;
begin
  Result := '';
  if FState.ActiveFile = '' then Exit;
  if Memo.Lines.Count = 0 then Exit;
  LineNo := Memo.Perform(EM_LINEFROMCHAR, Memo.SelStart, 0) + 1;
  ColPos := Memo.SelStart - Memo.Perform(EM_LINEINDEX, LineNo - 1, 0) + 1;
  if (LineNo < 1) or (LineNo > Memo.Lines.Count) then Exit;
  LineText := Memo.Lines[LineNo - 1];
  Col := ColPos - 1;
  if Col > Length(LineText) then Col := Length(LineText);
  Start := Col;
  while (Start >= 1) and CharInSet(LineText[Start], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(Start);
  Result := Copy(LineText, Start + 1, Col - Start);
end;

procedure TMainForm.RefreshCompletionFilter;
var
  Prefix: string;
begin
  // The popup was opened anchored at the position right after a `.`
  // (or after a manual Ctrl+Space). The "filter" is everything the
  // user has typed since then - which is exactly the word at the
  // current caret. When that is empty, the user has typed a non-word
  // character (`(`, `;`, space, ...) and is no longer in the
  // completion context; close the popup.
  Prefix := ComputeCurrentMemoWordPrefix;
  if Prefix = '' then
    TLspCompletionWizard(FCompletionWizard).HidePopup
  else
    TLspCompletionWizard(FCompletionWizard).SetPrefix(Prefix);
end;

procedure TMainForm.DoMemoClick(Sender: TObject);
var
  Line, Col: Integer;
begin
  Line := Memo.Perform(EM_LINEFROMCHAR, Memo.SelStart, 0) + 1;
  Col := Memo.SelStart - Memo.Perform(EM_LINEINDEX, Line - 1, 0) + 1;
  FState.SetActiveFile(FState.ActiveFile, Line, Col);
  UpdateStatusBar;
  TLspSignatureHelpWizard(FSignatureWizard).UpdateCursor(Line, Col);
end;

procedure TMainForm.DoMemoKeyPress(Sender: TObject; var Key: Char);
begin
  // Suppression for keys we already intercepted in OnKeyDown (Enter
  // to accept a completion, Tab, Escape). Without this the Memo can
  // still insert the corresponding char into the buffer.
  if FConsumeNextChar then
  begin
    FConsumeNextChar := False;
    Key := #0;
    Exit;
  end;
  // Capture only - the char has not been inserted into the buffer yet.
  // Real handling happens in OnKeyUp, when the caret already sits past
  // the just-inserted character.
  FLastKeyChar := Key;
end;

procedure TMainForm.DoMemoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  TypedChar: Char;
  CaretPos: TPoint;
begin
  // Refresh the cursor position so the wizard's GetCurrentContext
  // sees what the user actually has on screen.
  var Line := Memo.Perform(EM_LINEFROMCHAR, Memo.SelStart, 0) + 1;
  var Col := Memo.SelStart - Memo.Perform(EM_LINEINDEX, Line - 1, 0) + 1;
  if FState.ActiveFile <> '' then
    FState.SetActiveFile(FState.ActiveFile, Line, Col);
  UpdateStatusBar;

  // Manual signature trigger: Ctrl+Shift+Space - LSP standard.
  // Walks back to the enclosing '(' and shows the hint as if it had
  // just been typed. The screen-Y of the open-paren line is resolved
  // via the GetMemoLineScreenY callback so the popup floats above
  // the function name, even when the user is several lines below
  // typing arguments.
  if (Key = VK_SPACE) and (ssCtrl in Shift) and (ssShift in Shift) then
  begin
    if Winapi.Windows.GetCaretPos(CaretPos) then
    begin
      CaretPos := Memo.ClientToScreen(CaretPos);
      TLspSignatureHelpWizard(FSignatureWizard).TriggerManually(
        CaretPos.X, GetMemoLineScreenY);
    end;
    FLastKeyChar := #0;
    Exit;
  end;

  // FLastKeyChar (set by OnKeyPress) is the real keystroke, not just
  // a cursor move - so arrow-keying past an existing `.` or `(`
  // does not retrigger.
  TypedChar := FLastKeyChar;
  FLastKeyChar := #0;

  case TypedChar of
    '.':
      // Code completion (member list).
      if not FCompletionInFlight then
      begin
        FCompletionInFlight := True;
        try DoRefactorCompletion(nil); finally FCompletionInFlight := False; end;
      end;
    '(':
      begin
        // Parameter hint above the current line.
        if Winapi.Windows.GetCaretPos(CaretPos) then
        begin
          CaretPos := Memo.ClientToScreen(CaretPos);
          TLspSignatureHelpWizard(FSignatureWizard).ExecuteAt(
            CaretPos.X, CaretPos.Y);
        end;
      end;
  end;

  // After every keystroke / arrow key, let the signature wizard
  // recompute the active parameter (or hide when the caret leaves
  // the tracked call). Cheap when no popup is active.
  TLspSignatureHelpWizard(FSignatureWizard).UpdateCursor(Line, Col);

  // Escape on the Memo dismisses the parameter hint.
  if Key = VK_ESCAPE then
    TLspSignatureHelpWizard(FSignatureWizard).HidePopup;
end;

{ ---------- Refactor handlers - each spins up the wizard, calls into
  it, frees it. The wizards reach this form's project state through
  the global Editor accessor installed in FormCreate. All go through
  RunWizard so the Memo + tree refresh and the wait cursor are
  guaranteed. ---------- }

procedure TMainForm.RunWizard(AProc: TProc);
var
  PrevCursor: TCursor;
  ActiveBefore: string;
begin
  if not Assigned(AProc) then Exit;
  ActiveBefore := FState.ActiveFile;
  PrevCursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  Self.Enabled := False;
  try
    // The wizard runs synchronously on the UI thread. Without a real
    // background thread we cannot keep the window responsive; the wait
    // cursor + Self.Enabled := False at least make it visually obvious
    // that the app is working, not crashed, and stops accidental
    // double-clicks while a long LSP query is in flight.
    try
      AProc();
    except
      on E: Exception do
        ShowMessage('Refactoring failed: ' + E.Message);
    end;
  finally
    Self.Enabled := True;
    Screen.Cursor := PrevCursor;
  end;

  // Post-run refresh: the wizard may have written through
  // Editor.ReplaceFileContent (Extract Interface, Rename, Remove with,
  // Semantic Replace, ...) which updates the standalone state's buffer
  // but does not push the new text into our Memo. It may also have
  // called Editor.AddFileToActiveProject (Extract Interface) which
  // mutated FState.SourceFiles. Refresh both.
  RefreshTree;
  if (FState.ActiveFile <> '') and (FState.ActiveFile = ActiveBefore) then
    ReloadActiveFile;
  UpdateStatusBar;
end;

procedure TMainForm.DoRefactorRename(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspRenameWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorFindReferences(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspFindReferencesWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorFindImplementations(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspFindImplementationsWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorAlignSignature(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspSignatureCheckWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorExtractMethod(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspExtractMethodWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorCompletion(Sender: TObject);
// Bypasses RunWizard on purpose:
//   * The completion popup is non-modal - Execute returns immediately
//     after showing it. RunWizard's Self.Enabled := False / True
//     toggle would steal focus from the popup the instant Execute
//     returns, triggering its OnDeactivate -> Close.
//   * The wizard must outlive the popup: the popup fires OnInsert
//     when the user picks an item, possibly seconds after Execute
//     returned. We reuse the persistent FCompletionWizard for that
//     reason (the IDE plugin does the same via its singleton).
// Buffer refresh after the insert still happens via
// FState.OnExternalChange -> ReloadActiveFile, set up in FormCreate.
begin
  TLspCompletionWizard(FCompletionWizard).Execute;
end;

procedure TMainForm.DoRefactorRemoveWithCursor(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspWithRefactorWizard.Create;
    try W.ExecuteAtCursor; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorRemoveWithCurrent(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspWithRefactorWizard.Create;
    try W.ExecuteCurrentUnit; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorRemoveWithSelected(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspWithRefactorWizard.Create;
    try W.ExecuteSelectedUnits; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorRemoveWithProject(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspWithRefactorWizard.Create;
    try W.ExecuteProjectWide; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorMoveToUnit(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspMoveToUnitWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorUnitRefs(Sender: TObject);
begin
  RunWizard(procedure begin
    var W := TLspFindUnitReferencesWizard.Create;
    try W.Execute; finally W.Free; end;
  end);
end;

procedure TMainForm.DoRefactorExtractInterface(Sender: TObject);
begin RunWizard(procedure begin ExtractInterfaceFromClass; end); end;

procedure TMainForm.DoRefactorAddToInterface(Sender: TObject);
begin RunWizard(procedure begin AddToExistingInterface; end); end;

procedure TMainForm.DoRefactorAddIInterface(Sender: TObject);
begin RunWizard(procedure begin DelegateInterfaceImplementation; end); end;

procedure TMainForm.DoRefactorSemanticEditRules(Sender: TObject);
begin RunWizard(procedure begin EditSemanticReplaceRules; end); end;

procedure TMainForm.DoCheckDfmEvents(Sender: TObject);
begin RunWizard(procedure begin CheckDfmEventHandlers; end); end;

procedure TMainForm.DoCheckInterfaceGuids(Sender: TObject);
begin RunWizard(procedure begin CheckInterfaceGuids; end); end;

procedure TMainForm.DoCheckCircularRefs(Sender: TObject);
begin RunWizard(procedure begin CheckCircularReferences; end); end;

procedure TMainForm.DoFindUnit(Sender: TObject);
begin RunWizard(procedure begin FindUnitForIdentifier; end); end;

procedure TMainForm.DoAddUnitAtCursor(Sender: TObject);
begin RunWizard(procedure begin AddUnitForIdentifierAtCursor; end); end;

procedure TMainForm.DoShowQuickFixes(Sender: TObject);
begin RunWizard(procedure begin ShowAllQuickFixes; end); end;

procedure TMainForm.DoResolveMissingUnits(Sender: TObject);
begin RunWizard(procedure begin ResolveMissingUnits; end); end;

procedure TMainForm.DoFindOriginalSymbol(Sender: TObject);
begin RunWizard(procedure begin FindOriginalSymbol; end); end;

procedure TMainForm.DoUsesCleanup(Sender: TObject);
begin RunWizard(procedure begin CleanupUsesCurrentUnit; end); end;

procedure TMainForm.DoRefactorSemanticProject(Sender: TObject);
begin RunWizard(procedure begin ApplySemanticReplacements_Project; end); end;

procedure TMainForm.DoRefactorSemanticCurrent(Sender: TObject);
begin RunWizard(procedure begin ApplySemanticReplacements_CurrentUnit; end); end;

procedure TMainForm.DoRefactorSemanticSelected(Sender: TObject);
begin RunWizard(procedure begin ApplySemanticReplacements_SelectedUnits; end); end;

procedure TMainForm.UpdateStatusBar;
begin
  if FState.ProjectRoot <> '' then
    Status.Panels[0].Text := 'Project: ' + FState.ProjectRoot
  else
    Status.Panels[0].Text := 'No project loaded';
  if FState.ActiveFile <> '' then
  begin
    Status.Panels[1].Text := 'File: ' + ExtractFileName(FState.ActiveFile);
    // The dirty marker the README used to say could not exist: edits are
    // kept in the buffer and only Ctrl+S writes them (audit #33, M34).
    if (FDirty <> nil) and (FDirty.IndexOf(FState.ActiveFile) >= 0) then
      Status.Panels[1].Text := Status.Panels[1].Text + ' *';
    if (FDirty <> nil) and (FDirty.Count > 1) then
      Status.Panels[1].Text := Status.Panels[1].Text +
        Format('  (%d unsaved)', [FDirty.Count]);
  end
  else
    Status.Panels[1].Text := '';
  Status.Panels[2].Text := Format('Line %d  Col %d',
    [FState.ActiveLine, FState.ActiveCol]);
end;

end.
