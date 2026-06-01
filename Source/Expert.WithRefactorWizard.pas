(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WithRefactorWizard;

{
  Orchestrates the project-wide "remove with" refactoring:

    1) Resolve project files via TEditorHelper.GetProjectSourceFiles.
    2) Save unsaved editor buffers to disk so the LSP and the file
       reads agree on content.
    3) Start LSP via TLspManager and ensure it has indexed the project.
    4) For each file, run TWithScanner.ScanFile.
    5) For each occurrence, run TWithRewriter.Rewrite to produce the
       before/after texts.
    6) Show TWithRefactorDialog modal. The user reviews the entries
       and clicks "Apply selected", "Apply all" or "Close".
    7) Apply the chosen edits via TEditorHelper.ApplyEditViaEditor.
       For multiple edits in the same file, apply bottom-up so earlier
       offsets stay stable.

  This wizard does NOT need an editor cursor position - it operates on
  the whole project. It can therefore be invoked even when no source
  file is currently focused.
}

interface

uses
  System.Classes, System.SysUtils, System.Generics.Collections,
  Vcl.Forms, ToolsAPI;

type
  /// <summary>Scope of a remove-with run. Determines which source
  ///  files are scanned for occurrences.</summary>
  TWithRefactorScope = (
    /// <summary>Every scannable project source file (the original
    ///  behaviour). Slow on large projects.</summary>
    wrsProject,
    /// <summary>Only the currently active editor file.</summary>
    wrsCurrentUnit,
    /// <summary>A user-selected subset of project source files,
    ///  chosen via a multi-select dialog.</summary>
    wrsSelectedUnits,
    /// <summary>Only the with-statement that encloses the editor
    ///  caret position. If no with-statement encloses the caret, the
    ///  whole current unit is scanned as a graceful fallback.</summary>
    wrsCursor
  );

  TLspWithRefactorWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  public
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    /// <summary>Legacy entry point - delegates to ExecuteAtCursor (the
    ///  fastest single-action default).</summary>
    procedure Execute;
    function GetMenuText: string;

    /// <summary>Scans the entire project. Slow on big code bases.</summary>
    procedure ExecuteProjectWide;
    /// <summary>Scans only the currently active editor file.</summary>
    procedure ExecuteCurrentUnit;
    /// <summary>Opens a multi-select dialog of project files and
    ///  scans only the chosen ones.</summary>
    procedure ExecuteSelectedUnits;
    /// <summary>Scans only the with-statement that encloses the
    ///  cursor (falls back to the current unit when no with is at
    ///  the cursor).</summary>
    procedure ExecuteAtCursor;
  private
    procedure RunWithScope(AScope: TWithRefactorScope;
      const AExtraFiles: TArray<string>);
  end;

var
  WithRefactorInstance: TLspWithRefactorWizard;

implementation

uses
  System.UITypes, System.IOUtils, System.Math, System.JSON,
  Vcl.Dialogs, Vcl.Controls,
  Expert.EditorHelper, Expert.LspManager, Lsp.Client,
  Expert.WithScanner, Expert.WithRewriter, Expert.WithRefactorDialog;

procedure TLspWithRefactorWizard.AfterSave; begin end;
procedure TLspWithRefactorWizard.BeforeSave; begin end;
procedure TLspWithRefactorWizard.Destroyed; begin end;
procedure TLspWithRefactorWizard.Modified; begin end;

function TLspWithRefactorWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.WithRefactorWizard';
end;

function TLspWithRefactorWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Remove with';
end;

function TLspWithRefactorWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspWithRefactorWizard.GetMenuText: string;
begin
  Result := 'Remove with (project-wide)...';
end;

{ ---------- Apply phase ---------- }

type
  /// <summary>One concrete text edit to apply via
  ///  TEditorHelper.ApplyEditViaEditor. Position is 1-based here; the
  ///  apply step converts to 0-based.</summary>
  TPlainEdit = record
    FileName: string;
    Line: Integer;
    Col: Integer;
    OldText: string;
    NewText: string;
  end;

/// <summary>Builds the var-section insertion text for a method that
///  ALREADY has a 'var' section. Returns a block of new lines that
///  should be inserted after the column-after-the-last-var-decl-line.
///  Each decl is prefixed with AIndent and terminated with ';' + CRLF.</summary>
function MakeVarLinesAppend(const AIndent: string;
  const ADecls: TArray<TClassicVarDecl>): string;
var
  I: Integer;
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    for I := 0 to High(ADecls) do
      SB.Append(#13#10).Append(AIndent)
        .Append(ADecls[I].Name).Append(': ').Append(ADecls[I].TypeName)
        .Append(';');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

/// <summary>Builds the new 'var' section to insert before the method's
///  'begin' when no var section exists yet. Each decl is on its own
///  line indented by AIndent. The block ends with a CRLF so the
///  following 'begin' starts on a fresh line.</summary>
function MakeVarSectionNew(const AIndent: string;
  const ABeginCol: Integer;
  const ADecls: TArray<TClassicVarDecl>): string;
var
  I: Integer;
  SB: TStringBuilder;
  HeaderIndent: string;
begin
  // 'var' keyword aligns with the column of the method's 'begin'.
  HeaderIndent := StringOfChar(' ', Max(0, ABeginCol - 1));
  SB := TStringBuilder.Create;
  try
    SB.Append('var');
    for I := 0 to High(ADecls) do
      SB.Append(#13#10).Append(AIndent)
        .Append(ADecls[I].Name).Append(': ').Append(ADecls[I].TypeName)
        .Append(';');
    SB.Append(#13#10).Append(HeaderIndent);
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

/// <summary>Aggregates per-item Classic data into one flat edit list.
///  Per-method var declarations and per-file uses additions are
///  combined so each method gets exactly one var-section edit and
///  each file gets at most one uses-clause edit.</summary>
procedure BuildClassicEdits(const AItems: TArray<TWithRewriteResult>;
  out AEdits: TArray<TPlainEdit>);
type
  TMethodAgg = record
    AnyItem: TWithRewriteResult;
    AllDecls: TArray<TClassicVarDecl>;
  end;
  TFileAgg = record
    AnyItem: TWithRewriteResult;
    UnitsToAdd: TArray<string>;
  end;
var
  ByMethod: TDictionary<string, TMethodAgg>;
  ByFile: TDictionary<string, TFileAgg>;
  Edits: TList<TPlainEdit>;
  Item: TWithRewriteResult;
  Edit: TPlainEdit;

  function HasUnit(const AArr: TArray<string>; const AUnit: string): Boolean;
  var K: Integer;
  begin
    for K := 0 to High(AArr) do
      if SameText(AArr[K], AUnit) then Exit(True);
    Result := False;
  end;

  procedure AppendDecl(var AArr: TArray<TClassicVarDecl>;
    const ADecl: TClassicVarDecl);
  var K: Integer;
  begin
    for K := 0 to High(AArr) do
      if SameText(AArr[K].Name, ADecl.Name) then Exit;
    SetLength(AArr, Length(AArr) + 1);
    AArr[High(AArr)] := ADecl;
  end;

  procedure AppendUnit(var AArr: TArray<string>; const AUnit: string);
  begin
    if HasUnit(AArr, AUnit) then Exit;
    SetLength(AArr, Length(AArr) + 1);
    AArr[High(AArr)] := AUnit;
  end;

var
  MA: TMethodAgg;
  FA: TFileAgg;
  PairM: TPair<string, TMethodAgg>;
  PairF: TPair<string, TFileAgg>;
  J: Integer;
  KeyF: string;
begin
  ByMethod := TDictionary<string, TMethodAgg>.Create;
  ByFile := TDictionary<string, TFileAgg>.Create;
  Edits := TList<TPlainEdit>.Create;
  try
    for Item in AItems do
    begin
      if not Item.Classic.Supported then Continue;

      // 1) Body-replacement edit.
      Edit.FileName := Item.FileName;
      Edit.Line := Item.ReplaceRange.StartPos.Line;
      Edit.Col := Item.ReplaceRange.StartPos.Col;
      Edit.OldText := Item.OriginalText;
      Edit.NewText := Item.Classic.BodyText;
      Edits.Add(Edit);

      // 2) Aggregate var-section additions per method.
      if Item.Classic.MethodKey <> '' then
      begin
        if not ByMethod.TryGetValue(Item.Classic.MethodKey, MA) then
          MA := Default(TMethodAgg);
        MA.AnyItem := Item;
        for J := 0 to High(Item.Classic.VarDecls) do
          AppendDecl(MA.AllDecls, Item.Classic.VarDecls[J]);
        ByMethod.AddOrSetValue(Item.Classic.MethodKey, MA);
      end;

      // 3) Aggregate uses additions per file.
      if Length(Item.Classic.AddUnits) > 0 then
      begin
        KeyF := Item.FileName;
        if not ByFile.TryGetValue(KeyF, FA) then
          FA := Default(TFileAgg);
        FA.AnyItem := Item;
        for J := 0 to High(Item.Classic.AddUnits) do
          AppendUnit(FA.UnitsToAdd, Item.Classic.AddUnits[J]);
        ByFile.AddOrSetValue(KeyF, FA);
      end;
    end;

    // Emit method var-section edits.
    for PairM in ByMethod do
    begin
      Item := PairM.Value.AnyItem;
      if PairM.Value.AllDecls = nil then Continue;
      if Item.Classic.HasVarSection
         and (Item.Classic.VarSectionLastLine > 0) then
      begin
        // Append after the last var-decl line. Insert at column 1 of the
        // NEXT line; OldText = '' so it's a pure insertion.
        Edit.FileName := Item.FileName;
        Edit.Line := Item.Classic.VarSectionLastLine + 1;
        Edit.Col := 1;
        Edit.OldText := '';
        Edit.NewText := Copy(MakeVarLinesAppend(Item.Classic.LocalIndent,
          PairM.Value.AllDecls), 3, MaxInt) + #13#10; // strip leading CRLF, end with CRLF
        Edits.Add(Edit);
      end
      else
      begin
        // No var section yet. Insert before the 'begin' line.
        Edit.FileName := Item.FileName;
        Edit.Line := Item.Classic.MethodBodyBeginLine;
        Edit.Col := 1;
        Edit.OldText := '';
        Edit.NewText := MakeVarSectionNew(Item.Classic.LocalIndent,
          Item.Classic.MethodBodyBeginCol, PairM.Value.AllDecls);
        Edits.Add(Edit);
      end;
    end;

    // Emit uses-clause edits.
    for PairF in ByFile do
    begin
      Item := PairF.Value.AnyItem;
      if PairF.Value.UnitsToAdd = nil then Continue;
      if Item.Classic.ImplUsesFound then
      begin
        // Append ", U1, U2, ..." after last unit name (before ';').
        var SB := TStringBuilder.Create;
        try
          for J := 0 to High(PairF.Value.UnitsToAdd) do
            SB.Append(', ').Append(PairF.Value.UnitsToAdd[J]);
          Edit.FileName := Item.FileName;
          Edit.Line := Item.Classic.ImplUsesLastLine;
          Edit.Col := Item.Classic.ImplUsesLastCol + 1;
          Edit.OldText := '';
          Edit.NewText := SB.ToString;
        finally
          SB.Free;
        end;
        Edits.Add(Edit);
      end
      else if Item.Classic.ImplKeywordLine > 0 then
      begin
        // No implementation uses; create a new clause right after the
        // 'implementation' keyword.
        var SB := TStringBuilder.Create;
        try
          SB.Append(#13#10).Append(#13#10).Append('uses').Append(#13#10)
            .Append('  ');
          for J := 0 to High(PairF.Value.UnitsToAdd) do
          begin
            if J > 0 then SB.Append(', ');
            SB.Append(PairF.Value.UnitsToAdd[J]);
          end;
          SB.Append(';');
          Edit.FileName := Item.FileName;
          Edit.Line := Item.Classic.ImplKeywordLine;
          Edit.Col := Item.Classic.ImplKeywordCol + Length('implementation');
          Edit.OldText := '';
          Edit.NewText := SB.ToString;
        finally
          SB.Free;
        end;
        Edits.Add(Edit);
      end
      else if Item.Classic.IntfUsesFound then
      begin
        // No implementation keyword, no implementation uses; fall back
        // to appending to interface uses (e.g. .dpr).
        var SB := TStringBuilder.Create;
        try
          for J := 0 to High(PairF.Value.UnitsToAdd) do
            SB.Append(', ').Append(PairF.Value.UnitsToAdd[J]);
          Edit.FileName := Item.FileName;
          Edit.Line := Item.Classic.IntfUsesLastLine;
          Edit.Col := Item.Classic.IntfUsesLastCol + 1;
          Edit.OldText := '';
          Edit.NewText := SB.ToString;
        finally
          SB.Free;
        end;
        Edits.Add(Edit);
      end;
    end;

    AEdits := Edits.ToArray;
  finally
    Edits.Free;
    ByFile.Free;
    ByMethod.Free;
  end;
end;

/// <summary>Sorts a flat edit list by file (asc) and within each file
///  by (line, col) DESCENDING so applying edits top-to-bottom does not
///  invalidate offsets of later edits in the same file.</summary>
procedure SortPlainEdits(var AEdits: TArray<TPlainEdit>);

  procedure SwapAt(I, J: Integer);
  var
    Tmp: TPlainEdit;
  begin
    Tmp := AEdits[I];
    AEdits[I] := AEdits[J];
    AEdits[J] := Tmp;
  end;

  function Less(const A, B: TPlainEdit): Boolean;
  var
    Cmp: Integer;
  begin
    Cmp := CompareText(A.FileName, B.FileName);
    if Cmp <> 0 then Exit(Cmp < 0);
    if A.Line <> B.Line then Exit(A.Line > B.Line);
    Result := A.Col > B.Col;
  end;

var
  I, J: Integer;
begin
  for I := 1 to High(AEdits) do
  begin
    J := I;
    while (J > 0) and Less(AEdits[J], AEdits[J - 1]) do
    begin
      SwapAt(J, J - 1);
      Dec(J);
    end;
  end;
end;

procedure ApplyPlainEdits(const AEdits: TArray<TPlainEdit>;
  out AOk, AFailed: Integer);
var
  Sorted: TArray<TPlainEdit>;
  I: Integer;
begin
  AOk := 0;
  AFailed := 0;
  Sorted := Copy(AEdits);
  SortPlainEdits(Sorted);
  for I := 0 to High(Sorted) do
    if TEditorHelper.ApplyEditViaEditor(
        Sorted[I].FileName,
        Sorted[I].Line - 1,
        Sorted[I].Col - 1,
        Sorted[I].OldText,
        Sorted[I].NewText) then
      Inc(AOk)
    else
      Inc(AFailed);
end;

/// <summary>Inline-mode apply: one edit per item, replacing the
///  with-block with the inline-var NewText.</summary>
procedure ApplyEditsInline(const AItems: TArray<TWithRewriteResult>;
  out AOk, AFailed: Integer);
var
  Edits: TArray<TPlainEdit>;
  Edit: TPlainEdit;
  Item: TWithRewriteResult;
  I: Integer;
begin
  SetLength(Edits, Length(AItems));
  I := 0;
  for Item in AItems do
  begin
    Edit.FileName := Item.FileName;
    Edit.Line := Item.ReplaceRange.StartPos.Line;
    Edit.Col := Item.ReplaceRange.StartPos.Col;
    Edit.OldText := Item.OriginalText;
    Edit.NewText := Item.NewText;
    Edits[I] := Edit;
    Inc(I);
  end;
  ApplyPlainEdits(Edits, AOk, AFailed);
end;

procedure ApplyEdits(const AItems: TArray<TWithRewriteResult>;
  AUseInlineVars: Boolean;
  out AOk, AFailed: Integer);
var
  Edits: TArray<TPlainEdit>;
begin
  if AUseInlineVars then
    ApplyEditsInline(AItems, AOk, AFailed)
  else
  begin
    BuildClassicEdits(AItems, Edits);
    ApplyPlainEdits(Edits, AOk, AFailed);
  end;
end;

{ ---------- Scan phase ---------- }

/// <summary>Returns True iff the file path looks like a Pascal source
///  we should scan (.pas / .dpr / .dpk). Other extensions are skipped.</summary>
function IsScannableSource(const APath: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(APath));
  Result := (Ext = '.pas') or (Ext = '.dpr') or (Ext = '.dpk');
end;

procedure TLspWithRefactorWizard.Execute;
begin
  // Default shortcut behaviour: cursor-only (fastest, most common
  // single-edit case). The submenu items in the editor context menu
  // expose the other scopes (current unit / selected units / project
  // wide). On very large projects, project-wide can take many
  // minutes and should be triggered explicitly.
  ExecuteAtCursor;
end;

procedure TLspWithRefactorWizard.ExecuteProjectWide;
begin
  RunWithScope(wrsProject, nil);
end;

procedure TLspWithRefactorWizard.ExecuteCurrentUnit;
begin
  RunWithScope(wrsCurrentUnit, nil);
end;

procedure TLspWithRefactorWizard.ExecuteSelectedUnits;
var
  AllFiles, Selected: TArray<string>;
begin
  AllFiles := TEditorHelper.GetProjectSourceFiles;
  if Length(AllFiles) = 0 then
  begin
    MessageDlg('No project source files found.', mtInformation, [mbOK], 0);
    Exit;
  end;
  // Filter to scannable files first so the user only sees what we
  // could actually process.
  var Filtered: TArray<string>;
  for var F in AllFiles do
    if IsScannableSource(F) and TFile.Exists(F) then
      Filtered := Filtered + [F];
  if Length(Filtered) = 0 then
  begin
    MessageDlg('No scannable .pas/.dpr/.dpk files in this project.',
      mtInformation, [mbOK], 0);
    Exit;
  end;
  if not TWithRefactorDialog.PickFiles(Application.MainForm, Filtered, Selected)
  then Exit;
  if Length(Selected) = 0 then Exit;
  RunWithScope(wrsSelectedUnits, Selected);
end;

procedure TLspWithRefactorWizard.ExecuteAtCursor;
var
  Ctx: TEditorContext;
begin
  Ctx := TEditorHelper.GetCurrentContext;
  if (Ctx.FileName = '') or not IsScannableSource(Ctx.FileName) then
  begin
    MessageDlg('No Pascal file is currently active.', mtInformation, [mbOK], 0);
    Exit;
  end;
  // RunWithScope detects wrsCursor and uses GetCurrentContext to find
  // which occurrence is at the caret.
  RunWithScope(wrsCursor, [Ctx.FileName]);
end;

procedure TLspWithRefactorWizard.RunWithScope(AScope: TWithRefactorScope;
  const AExtraFiles: TArray<string>);
var
  Dialog: TWithRefactorDialog;
  ProjFiles: TArray<string>;
  DelphiLspJson, RootPath, ProjFile: string;
  Client: TLspClient;
  Results: TList<TWithRewriteResult>;
  ScanFiles: TList<string>;
  I, FileIdx: Integer;
  Source: string;
  Occs: TArray<TWithOccurrence>;
  Occ: TWithOccurrence;
  Rewrite: TWithRewriteResult;
  AutoCount: Integer;
begin
  Dialog := TWithRefactorDialog.CreateDialog(Application.MainForm);
  try
    case AScope of
      wrsProject:
        begin
          Dialog.Caption := 'Remove with - project-wide review';
          Dialog.SetStatus('Resolving project source files (project-wide)...');
        end;
      wrsCurrentUnit:
        begin
          Dialog.Caption := 'Remove with - current unit';
          Dialog.SetStatus('Scanning current unit...');
        end;
      wrsSelectedUnits:
        begin
          Dialog.Caption := Format('Remove with - %d selected unit(s)',
            [Length(AExtraFiles)]);
          Dialog.SetStatus(Format('Scanning %d selected unit(s)...',
            [Length(AExtraFiles)]));
        end;
      wrsCursor:
        begin
          Dialog.Caption := 'Remove with - at cursor';
          Dialog.SetStatus('Scanning at cursor...');
        end;
    end;
    TLspManager.Instance.ApplyStatusToCaption(Dialog);
    Dialog.Show;
    Application.ProcessMessages;

    ScanFiles := TList<string>.Create;
    try
      case AScope of
        wrsProject:
          begin
            ProjFiles := TEditorHelper.GetProjectSourceFiles;
            for I := 0 to High(ProjFiles) do
              if IsScannableSource(ProjFiles[I]) and TFile.Exists(ProjFiles[I]) then
                ScanFiles.Add(ProjFiles[I]);
          end;
        wrsCurrentUnit, wrsCursor:
          begin
            var Ctx := TEditorHelper.GetCurrentContext;
            if (Ctx.FileName <> '') and IsScannableSource(Ctx.FileName)
              and TFile.Exists(Ctx.FileName) then
              ScanFiles.Add(Ctx.FileName);
          end;
        wrsSelectedUnits:
          begin
            for I := 0 to High(AExtraFiles) do
              if IsScannableSource(AExtraFiles[I]) and TFile.Exists(AExtraFiles[I]) then
                ScanFiles.Add(AExtraFiles[I]);
          end;
      end;

      if ScanFiles.Count = 0 then
      begin
        Dialog.SetStatus('No source files to scan for this scope.');
        Dialog.Hide;
        Dialog.ShowModal;
        Exit;
      end;

      // Save editors so on-disk content matches what the LSP sees.
      Dialog.SetStatus('Saving open editor buffers...');
      Application.ProcessMessages;
      TEditorHelper.SaveAllFiles;

      // Resolve LSP config.
      DelphiLspJson := TEditorHelper.FindDelphiLspJson;
      if DelphiLspJson = '' then
      begin
        Dialog.SetStatus('No .delphilsp.json found - cannot resolve target types. ' +
          'Enable Tools > Options > Editor > Language > Code Insight > "Generate LSP Config".');
        Dialog.Hide;
        Dialog.ShowModal;
        Exit;
      end;

      RootPath := TEditorHelper.GetProjectRoot;
      ProjFile := TEditorHelper.GetCurrentProjectDproj;
      if RootPath = '' then
        RootPath := ExtractFilePath(ProjFile);

      Dialog.SetStatus('Starting LSP server...');
      Application.ProcessMessages;
      try
        Client := TLspManager.Instance.GetClient(RootPath, ProjFile, DelphiLspJson);
      except
        on E: Exception do
        begin
          Dialog.SetStatus('LSP startup failed: ' + E.Message);
          Dialog.Hide;
          Dialog.ShowModal;
          Exit;
        end;
      end;

      Dialog.SetStatus('Indexing project...');
      Application.ProcessMessages;
      try
        TLspManager.Instance.EnsureProjectIndexed(ScanFiles.ToArray,
          procedure(ACurrent, ATotal: Integer; const ACurrentFile: string)
          begin
            Dialog.SetProgress(ACurrent, ATotal);
            if ACurrentFile <> '' then
              Dialog.SetStatus(Format('Indexing %d/%d: %s',
                [ACurrent, ATotal, ExtractFileName(ACurrentFile)]));
            Application.ProcessMessages;
          end);
      except
        on E: Exception do
          Dialog.SetStatus('Indexing partial: ' + E.Message);
      end;

      // Diagnostic-Aufwaermphase ist file-individuell: DelphiLSP pusht
      // publishDiagnostics in der Reihenfolge wie er Files analysiert
      // (Sekunden bis Minuten pro File). Wir warten hier nicht global -
      // stattdessen wartet die per-File-Verarbeitungsschleife unten
      // gezielt auf "Diagnostics fuer DIESES File angekommen?", bevor
      // sie IsLineInactive prueft.

      // Scan + rewrite.
      Results := TList<TWithRewriteResult>.Create;
      try
        Dialog.SetProgress(0, ScanFiles.Count);
        for FileIdx := 0 to ScanFiles.Count - 1 do
        begin
          Dialog.SetStatus(Format('Scanning %d/%d: %s',
            [FileIdx + 1, ScanFiles.Count, ExtractFileName(ScanFiles[FileIdx])]));
          Dialog.SetProgress(FileIdx + 1, ScanFiles.Count);
          // Caption-Status (im Title-Bar) reflektiert den aktuellen
          // LSP-Stand (Diagnostics-Count, inaktive Regionen). Wird jeden
          // File aktualisiert damit der User die Live-Entwicklung sieht.
          TLspManager.Instance.ApplyStatusToCaption(Dialog);
          if (FileIdx mod 3 = 0) then Application.ProcessMessages;

          try
            Source := TFile.ReadAllText(ScanFiles[FileIdx]);
          except
            Continue;
          end;

          Occs := TWithScanner.ScanSource(Source);
          if Length(Occs) = 0 then Continue;

          // In wrsCursor mode we filter Occs down to the single
          // with-statement that ENCLOSES the editor caret position. If
          // none does, we leave Occs untouched so the user at least
          // gets the whole current unit (graceful fallback).
          if AScope = wrsCursor then
          begin
            var Ctx := TEditorHelper.GetCurrentContext;
            // Only filter when we're scanning the file the caret is in
            // (otherwise leave Occs alone - same unit fallback).
            if SameText(Ctx.FileName, ScanFiles[FileIdx]) then
            begin
              var Picked: TArray<TWithOccurrence>;
              for var O in Occs do
                if (Ctx.Line >= O.KeywordPos.Line)
                  and (Ctx.Line <= O.BodyRange.EndPos.Line) then
                  Picked := Picked + [O];
              if Length(Picked) > 0 then
                Occs := Picked;
              // else: keep full Occs as fallback
            end;
          end;

          // Bevor wir Occurrences gegen inaktive Regionen pruefen,
          // braucht DelphiLSP eine durchgefuehrte Analyse des Files.
          //
          // Eine passive WaitForDiagnostics-Schleife reicht nicht: in
          // grossen Projekten serialisiert DelphiLSP die Analyse stark
          // (im Test ~22 s pro File). Wir koennen aber AKTIV blockieren
          // indem wir 'textDocument/documentSymbol' synchron anfordern -
          // LSP MUSS dafuer das File analysieren, sonst kommt keine
          // Antwort. Mit Timeout 90 s (30 s im Project-Scope).
          // Anschliessend kurz auf publishDiagnostics warten (die
          // kommen direkt nach der Analyse - max. 10 s).
          if not Client.HasReceivedDiagnostics(ScanFiles[FileIdx]) then
          begin
            Dialog.SetStatus(Format('Waiting for LSP to analyse: %s...',
              [ExtractFileName(ScanFiles[FileIdx])]));
            Application.ProcessMessages;
            try
              Client.RefreshDocument(ScanFiles[FileIdx]);
            except
              // Refresh-Fehler ist nicht kritisch
            end;
            // Aktive Analyse-Anfrage: documentSymbol BLOCKIERT bis LSP
            // antwortet. Wir verwerfen das Ergebnis - es geht uns nur
            // darum, dass LSP den File durchparst.
            var SymTimeoutMs: Cardinal;
            if AScope = wrsProject then SymTimeoutMs := 30000 else SymTimeoutMs := 120000;
            var SymJson: TJSONArray := nil;
            try
              SymJson := Client.GetDocumentSymbols(ScanFiles[FileIdx], SymTimeoutMs);
            except
              // Timeout / Fehler: wir versuchen trotzdem weiterzumachen.
            end;
            if SymJson <> nil then SymJson.Free;
            // Nach documentSymbol ist DelphiLSP soweit. Diagnostics
            // kommen typisch in den naechsten Sekunden hinterher.
            Client.WaitForDiagnostics(ScanFiles[FileIdx], 15000);
          end;

          // Wenn DelphiLSP fuer dieses File ueberhaupt keine
          // publishDiagnostics geliefert hat, koennen wir nicht
          // sicherstellen, dass die with-Statements nicht in einem
          // inaktiven {$IFDEF}/{$IF defined(...)}-Block stehen. Eigene
          // Textanalyse waere unzuverlaessig (verschachtelte $IF,
          // $IF defined, $IFOPT, projektspezifische Defines, ...).
          // Wir markieren in dem Fall alle Occurrences als skipped.
          var FileHasDiagnostics: Boolean := Client.HasReceivedDiagnostics(ScanFiles[FileIdx]);

          for Occ in Occs do
          begin
            if not FileHasDiagnostics then
            begin
              Rewrite := Default(TWithRewriteResult);
              Rewrite.FileName := ScanFiles[FileIdx];
              Rewrite.Occurrence := Occ;
              Include(Rewrite.Issues, wriLspNoDiagnostics);
              Rewrite.OriginalText :=
                'DelphiLSP delivered no diagnostics for this file - '
                + 'cannot determine whether the with-statement is inside '
                + 'an inactive {$IFDEF}/{$IF defined(...)} region. '
                + 'Skipped to avoid rewriting potentially dead code.';
              Results.Add(Rewrite);
              Continue;
            end;
            // Skip occurrences inside inactive {$IFDEF}-regions:
            // DelphiLSP pushes those as diagnostics with code H2655/H2656
            // and tag=Unnecessary. We pre-populate the LSP-client's
            // inactive-range table; if the with-keyword line falls into
            // one, the rewriter would produce noise on dead code.
            // KeywordPos.Line is 1-based, IsLineInactive expects 0-based.
            if Client.IsLineInactive(ScanFiles[FileIdx], Occ.KeywordPos.Line - 1) then
            begin
              Rewrite := Default(TWithRewriteResult);
              Rewrite.FileName := ScanFiles[FileIdx];
              Rewrite.Occurrence := Occ;
              Include(Rewrite.Issues, wriInactiveRegion);
              Rewrite.OriginalText :=
                'with-statement inside an inactive {$IFDEF}-region '
                + '(DelphiLSP H2655/H2656). Skipped.';
              Results.Add(Rewrite);
              Continue;
            end;
            try
              Rewrite := TWithRewriter.Rewrite(Client,
                ScanFiles[FileIdx], Source, Occ,
                TWithRewriteSettings.Defaults);
            except
              on E: Exception do
              begin
                Rewrite := Default(TWithRewriteResult);
                Rewrite.FileName := ScanFiles[FileIdx];
                Rewrite.Occurrence := Occ;
                Include(Rewrite.Issues, wriTypeUnresolved);
              end;
            end;
            Results.Add(Rewrite);
          end;
        end;

        // Compose status summary.
        AutoCount := 0;
        for I := 0 to Results.Count - 1 do
          if Results[I].IsAutoRewritable then Inc(AutoCount);

        Dialog.SetStatus(Format('Found %d with-statement(s) - %d auto-rewritable.',
          [Results.Count, AutoCount]));
        Dialog.SetProgress(0, 0);
        Dialog.SetItems(Results.ToArray);
        TLspManager.Instance.ApplyStatusToCaption(Dialog);

        // Non-modal review: wire double-click-goto + apply callback,
        // hand ownership to the dialog so it free's itself when the
        // user closes it. Execute returns; the dialog lives on its
        // own. Multiple Remove-with dialogs can be open at the same
        // time - each one captures its own copy of the relevant state.
        Dialog.OnGoto :=
          procedure(const AItem: TWithRewriteResult)
          begin
            // Jump to the with-keyword position in the source file.
            TEditorHelper.GotoLocation(
              AItem.FileName,
              AItem.Occurrence.KeywordPos.Line - 1,
              AItem.Occurrence.KeywordPos.Col - 1,
              0);
          end;
        Dialog.OnApply :=
          procedure(const ARequested: TArray<TWithRewriteResult>;
            AUseInlineVars: Boolean;
            out AApplied: TArray<TWithRewriteResult>)
          var
            ROk, RFailed: Integer;
          begin
            AApplied := nil;
            if Length(ARequested) = 0 then Exit;
            ApplyEdits(ARequested, AUseInlineVars, ROk, RFailed);
            if RFailed > 0 then
              MessageDlg(Format('Applied %d edit(s); %d failed.',
                [ROk, RFailed]), mtWarning, [mbOK], 0);
            // We don't actually know which individual items failed
            // vs succeeded here; the rewriter applies them in order
            // and bails on first error, so the first ROk items are
            // good and the rest stay in the dialog. Conservative: only
            // report success when nothing failed; otherwise the user
            // can re-trigger and we'll re-run on the remaining items.
            if RFailed = 0 then
              AApplied := ARequested;
          end;
        // Reuse vorhandener "hand off ownership" Mechanik aus Find Refs:
        // ab jetzt schliesst der User den Dialog und der Dialog free't
        // sich selbst.
        Dialog.SetClosable;
        Dialog.Show;
        // Lokale Dialog-Referenz NICHT freigeben - der Dialog managed
        // sich selbst.
        Dialog := nil;
      finally
        Results.Free;
      end;
    finally
      ScanFiles.Free;
    end;
  finally
    // Wenn Dialog nicht auf nil gesetzt wurde (frueher Exit / Fehler),
    // muessen wir trotzdem freigeben.
    if Dialog <> nil then
      Dialog.Free;
  end;
end;

end.
