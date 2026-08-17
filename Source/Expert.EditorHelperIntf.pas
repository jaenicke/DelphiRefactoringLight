(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
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

  IEditorHelper = interface
    ['{1F6F4D86-5C8D-4A6E-9B0E-7BCE2C5F0A12}']
    // ---------- Cursor / project context ----------
    function GetCurrentContext: TEditorContext;
    function GetCurrentProjectDproj: string;
    function GetProjectRoot: string;
    function GetProjectSearchPaths: string;
    function GetProjectSourceFiles: TArray<string>;
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
    procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    procedure NotifyClassStructureChanged(const AFilePath: string);

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

implementation

var
  GEditor: IEditorHelper;

function Editor: IEditorHelper;
begin
  Result := GEditor;
end;

procedure SetEditorImpl(const AImpl: IEditorHelper);
begin
  GEditor := AImpl;
end;

end.
