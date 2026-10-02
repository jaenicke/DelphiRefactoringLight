(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.EditorHelperIntf;

// IDE-agnostic editor abstraction.
//
// The wizards and engines that used to call `TEditorHelper.Foo` directly
// (which delegated to RAD Studio's ToolsAPI through BorlandIDEServices)
// now call `Editor.Foo` via this interface. Two concrete implementations
// fulfil the contract:
//
//   * TIDEEditorHelper (Expert.EditorHelper.pas) - the default,
//     ToolsAPI-backed implementation that talks to the running IDE.
//
//   * TStandaloneEditorHelper - a future implementation for the
//     standalone executable; it talks to the app's own file tree +
//     embedded editor instead of an IDE.
//
// The active implementation is installed via SetEditorImpl at startup.
// Until that call, Editor returns nil, so the interface unit itself has
// zero ToolsAPI / VCL dependencies and can be used by the standalone
// project without pulling in IDE references.

interface

uses
  // RTL only - the unit's point is to carry NO ToolsAPI / VCL dependency,
  // and System.SysUtils is what EOffMainThread needs.
  System.SysUtils;

type
  /// <summary>Cursor + project state at the moment a wizard is invoked.
  ///  Line and Column are 1-based (matching what the user sees in
  ///  status bars and dialogs).</summary>
  TEditorContext = record
    FileName: string;
    Line: Integer;
    Column: Integer;
    WordAtCursor: string;
    ProjectFile: string;
    ProjectRoot: string;
    IsValid: Boolean;
  end;

  /// <summary>One unit the IDE's form designer writes into a form unit's
  ///  interface uses BY ITSELF, with the component that asks for it
  ///  ('cxGrid1: TcxGrid'). Removing such an entry is a tug-of-war: the
  ///  IDE puts it back on the next save.</summary>
  TDesignerRequiredUnit = record
    UnitName: string;
    Reason: string;
  end;

  IEditorHelper = interface
    ['{1F6F4D86-5C8D-4A6E-9B0E-7BCE2C5F0A12}']
    // ---------- Cursor / project context ----------
    function GetCurrentContext: TEditorContext;
    /// <summary>File name of the topmost editor buffer - CHEAP and without
    ///  touching the caret/selection (unlike GetCurrentContext, which moves
    ///  the edit position). Safe to call from idle/timer handlers.
    ///  '' when no editor is active.</summary>
    function GetActiveFileName: string;

    /// <summary>Caret position (1-based line/column) of the active editor -
    ///  CHEAP and read-only like GetActiveFileName (never moves the edit
    ///  position). False when no editor is active.</summary>
    function GetCaretLineCol(out ALine, ACol: Integer): Boolean;
    /// <summary>The 1-based STRING index in ALINE that the IDE's 1-based
    ///  DISPLAY column ADISPLAYCOL points at. The two differ as soon as
    ///  the line contains a TAB: CursorPos.Col is tab-expanded, and using
    ///  it as a string index made the completion replace the wrong span
    ///  ("<Tab>Foo.Ba" + "Bar" became "<Tab>Foo.        Bar", audit #39,
    ///  M36a). The IDE converts it for us (IOTAEditBuffer.ConvertPos);
    ///  outside the IDE, and for a line without tabs, the answer is
    ///  ADISPLAYCOL itself.</summary>
    function RawColumn(const AFile: string; ALine, ADisplayCol: Integer): Integer;

    function GetCurrentProjectDproj: string;
    function GetProjectRoot: string;
    function GetProjectSearchPaths: string;
    function GetProjectSourceFiles: TArray<string>;
    /// <summary>Source files currently OPEN in the editor. Cheap (a
    ///  handful of modules) and the only place where UNSAVED changes
    ///  live - a caller that scans the project from disk must check
    ///  these buffers separately or it works on stale text.</summary>
    function GetOpenSourceFiles: TArray<string>;
    function BuildSearchPathFromProject(
      const ADprojPath, ARootPath: string): string;
    function FindDelphiLspJson: string;

    // ---------- File-level reads ----------
    /// <summary>Returns the live editor buffer for AFilePath (True) or
    ///  False when the file is not open in the editor; in the latter
    ///  case the caller should fall back to a disk read.</summary>
    function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;

    // ---------- File-level writes (undoable where possible) ----------
    /// <summary>Replaces the entire content of AFilePath. In the
    ///  IDE-backed implementation this goes through IOTAEditWriter, so
    ///  the change is undoable and visible without a manual reload. In
    ///  standalone, it just writes to disk.</summary>
    function ReplaceFileContent(const AFilePath: string;
      const ANewContent: string): Boolean;

    /// <summary>Replaces the (1-based) range [AStartLine:AStartCol,
    ///  AEndLine:AEndCol) with ANewText.</summary>
    function ReplaceSelection(const AFilePath: string;
      AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      const ANewText: string): Boolean;

    /// <summary>Replaces line ALine (1-based) wholesale.</summary>
    function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;

    /// <summary>Deletes line ALine (1-based).</summary>
    function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;

    /// <summary>Inserts AText at the very start of line ALine
    ///  (1-based), bypassing the IDE's auto-indent.</summary>
    function InsertTextAtLineStart(const AFilePath: string;
      ALine: Integer; const AText: string): Boolean;

    /// <summary>Replaces a specific token. ALine/ACol are 0-based.</summary>
    function ApplyEditViaEditor(const AFilePath: string;
      ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;

    // ---------- IDE-specific niceties (no-ops in standalone) ----------
    procedure SaveAllFiles;
    /// <summary>Saves the single module for AFilePath if it is open in the
    ///  IDE (so a caller can save one form at a time and show which unit is
    ///  being processed). No-op / True in standalone, where edits are
    ///  already written straight to disk.</summary>
    function SaveFile(const AFilePath: string): Boolean;
    procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    procedure NotifyClassStructureChanged(const AFilePath: string);
    /// <summary>True when the form belonging to APasFile is loaded in the
    ///  IDE's form designer. The designer then OWNS the form: a .dfm
    ///  changed on disk would simply be overwritten on the next save.
    ///  Always False in standalone.</summary>
    function IsFormInDesigner(const APasFile: string): Boolean;
    /// <summary>The units the IDE's form designer inserts by itself for the
    ///  form of APasFile: the unit of every component class AND of its
    ///  ancestors, plus whatever the registered selection editors ask for
    ///  (ISelectionEditor.RequiresUnits - DevExpress uses that heavily).
    ///  None of these appears as an identifier in the .pas, so a textual
    ///  analysis cannot see them.
    ///  False when the form is not loaded in the designer; AComplete is
    ///  False when a selection editor raised, so the list may MISS units -
    ///  the caller must then treat the answer as unverified rather than
    ///  complete. Main thread only. Always False in standalone.</summary>
    function GetDesignerRequiredUnits(const APasFile: string;
      out AUnits: TArray<TDesignerRequiredUnit>;
      out AComplete: Boolean): Boolean;
    /// <summary>Renames a component (AIsMethod = False) or an event
    ///  handler method (True) through the form designer of APasFile - the
    ///  path the Object Inspector takes, so the designer updates its own
    ///  bindings and the declaration in the source. False with a reason in
    ///  AMessage when the form is not in the designer or the designer
    ///  refused (e.g. an inherited component).</summary>
    function RenameInFormDesigner(const APasFile, AOldName, ANewName: string;
      AIsMethod: Boolean; out AMessage: string): Boolean;

    /// <summary>Opens AFilePath in the editor and positions the cursor
    ///  at (ALine, ACol). 0-based positions (LSP convention).
    ///  AHighlightLen > 0 selects that many characters from the cursor.</summary>
    function GotoLocation(const AFilePath: string;
      ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;

    /// <summary>Adds AFilePath to the currently active project. In the
    ///  IDE this goes through IOTAProject.AddFile so the .dproj is
    ///  updated in-memory and saved on next File > Save All. In
    ///  standalone this writes a new DCCReference entry into the
    ///  loaded .dproj XML directly. Idempotent: a no-op when the file
    ///  is already part of the project. Returns False if there is no
    ///  active project.</summary>
    function AddFileToActiveProject(const AFilePath: string): Boolean;

    /// <summary>Appends ADir to the active project's unit search path
    ///  (DCC_UnitSearchPath of the BASE configuration in the IDE).
    ///  Idempotent: a no-op returning True when the directory is already
    ///  listed. False when there is no active project or the host cannot
    ///  modify the project options (standalone).</summary>
    function AddProjectSearchPath(const ADir: string): Boolean;

    /// <summary>Returns the active editor's current selection.
    ///  Line/Col are 1-based, AEndLine/AEndCol point one past the last
    ///  character (LSP-range-end style).
    ///
    ///  Returns False when there is no selection (or no active editor)
    ///  - the caller should warn the user and abort.
    ///
    ///  Used by Extract Method, which needs the literal selected text
    ///  to extract; other wizards work off the cursor position alone
    ///  (see GetCurrentContext).</summary>
    function GetSelection(out AFilePath: string;
      out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      out AText: string): Boolean;
  end;

/// <summary>Returns the active IEditorHelper implementation. Nil if no
///  implementation has been installed yet (call SetEditorImpl in
///  initialization).</summary>
function Editor: IEditorHelper;

/// <summary>Installs the active implementation. Pass nil to clear (used
///  by tests).</summary>
procedure SetEditorImpl(const AImpl: IEditorHelper);

// ---------------------------------------------------------------------------
//  The main-thread rule, made visible
// ---------------------------------------------------------------------------
//
// ToolsAPI - editor buffers included - is MAIN THREAD ONLY. That rule has
// been broken silently more than once: the completion wizards broke it
// through a "read the live buffer" helper hidden inside the LSP client, and
// three MCP tools apply their edit AFTER their McpRunOnMain block has closed,
// i.e. from the pipe handler thread (fork audit, 2026-10). Such a race does
// not fail - it corrupts a buffer now and then, which is the worst way for a
// defect to behave.
// So every write path names itself, and a violation becomes VISIBLE (status
// window, get_status) instead of being a coin toss.

type
  EOffMainThread = class(Exception);

var
  /// <summary>Turn a violation into an exception instead of a counter.
  ///  OFF by default ON PURPOSE: three MCP tools are known to violate the
  ///  rule today (reported, fix pending), and a guard that turns a
  ///  working-by-luck tool into a hard error before the tool is fixed would
  ///  be a regression of its own. Switch it on once MainThreadViolations
  ///  stays 0 through a full round of the tools.</summary>
  StrictMainThread: Boolean = False;

/// <summary>True iff the caller runs on the main thread.</summary>
function OnMainThread: Boolean;

/// <summary>Records that AWhat ran off the main thread (counter + name).</summary>
procedure NoteOffMainThread(const AWhat: string);

/// <summary>What every write path calls first: records the violation and,
///  with StrictMainThread on, raises EOffMainThread instead of racing.</summary>
procedure RequireMainThread(const AWhat: string);

/// <summary>How many violations were recorded, and the name of the last one
///  ('ReplaceFileContent' ...) - for the status window and get_status.</summary>
function MainThreadViolations: Integer;
function LastOffMainThreadCall: string;

implementation

uses
  System.Classes, System.SyncObjs;

var
  GEditor: IEditorHelper;
  GViolations: Integer;
  GLastOffMain: string;
  GLastLock: TCriticalSection;

function Editor: IEditorHelper;
begin
  Result := GEditor;
end;

procedure SetEditorImpl(const AImpl: IEditorHelper);
begin
  GEditor := AImpl;
end;

function OnMainThread: Boolean;
begin
  Result := TThread.CurrentThread.ThreadID = MainThreadID;
end;

procedure NoteOffMainThread(const AWhat: string);
begin
  TInterlocked.Increment(GViolations);
  // A worker that is still running while the BPL unloads would otherwise
  // AV on a freed lock - the counter is what matters, the name is a bonus.
  if GLastLock = nil then Exit;
  GLastLock.Enter;
  try
    GLastOffMain := AWhat;
  finally
    GLastLock.Leave;
  end;
end;

procedure RequireMainThread(const AWhat: string);
begin
  if OnMainThread then Exit;
  NoteOffMainThread(AWhat);
  if StrictMainThread then
    raise EOffMainThread.CreateFmt('%s was called from a worker thread. ' +
      'ToolsAPI and the editor buffers are main thread only - run it through ' +
      'RunOnMain / McpRunOnMain.', [AWhat]);
end;

function MainThreadViolations: Integer;
begin
  Result := TInterlocked.CompareExchange(GViolations, 0, 0);
end;

function LastOffMainThreadCall: string;
begin
  if GLastLock = nil then Exit('');
  GLastLock.Enter;
  try
    Result := GLastOffMain;
  finally
    GLastLock.Leave;
  end;
end;

initialization
  GLastLock := TCriticalSection.Create;

finalization
  var L := GLastLock;
  GLastLock := nil;   // readers bail out instead of touching a freed lock
  L.Free;

end.
