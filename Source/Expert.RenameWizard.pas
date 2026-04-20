(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.RenameWizard;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types, System.UITypes, System.Math, System.StrUtils,
  System.Generics.Collections, System.Generics.Defaults, Vcl.Forms, Vcl.Dialogs, ToolsAPI, Expert.EditorHelper,
  Expert.RenameDialog, Expert.LspManager, Expert.ImplementationFinder, Expert.FindReferencesDialog, Lsp.Uri, Lsp.Protocol,
  Lsp.Client, Rename.WorkspaceEdit, Delphi.FileEncoding;

type
  TRenameCandidate = record
    FilePath: string;
    Line: Integer;      // 0-based
    Col: Integer;       // 0-based
    OldText: string;
  end;

  TLspRenameWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  private
    FDialog: TRenameDialog;
    FContext: TEditorContext;
    FEdit: TLspWorkspaceEdit;
    FDiagLog: string;
    /// <summary>True when the dialog was opened for a unit rename
    ///  (triggered by the IDE module notifier). In that mode the preview
    ///  skips LSP verification and just does a text-based scan.</summary>
    FUnitRenameMode: Boolean;
    procedure DoPreview(Sender: TObject);
    procedure DoPreviewForIdentifier;
    procedure DoPreviewForUnit;
    procedure ApplyFEdit;
    function BuildEditFromCandidates(const ACandidates: TArray<TRenameCandidate>;
      const AOldName, ANewName: string): TLspWorkspaceEdit;

    function FindCandidates(const AOldName: string; const AFiles: TArray<string>): TArray<TRenameCandidate>;
    function VerifyWithLsp(const ACandidates: TArray<TRenameCandidate>; const AOldName, ANewName, ADefFilePath: string;
      const AImplFiles: TArray<string>; AClient: TLspClient): TLspWorkspaceEdit;
    function IsInStringOrComment(const ALine: string; APos: Integer): Boolean;

    /// <summary>Finds interface/class method implementations via a
    ///  text + syntax scan over all project files. Returns candidates
    ///  like 'procedure TFoo.Bar', 'function TFoo.Bar', ...
    ///  AOwnerType (optional) is the type name (interface/class) that
    ///  declares the method; only classes that implement this type are
    ///  returned.</summary>
    function FindImplementations(const AOldName: string; const AProjectFiles: TArray<string>;
      const AOwnerType: string): TArray<TRenameCandidate>;

    /// <summary>Converts a WorkspaceEdit into a list of
    ///  TRenamePreviewItem: for each edit the affected line is read,
    ///  the original line and the preview line are set, and the kind
    ///  (Interface / Implementation / Call etc.) is determined.</summary>
    function BuildPreviewItems(const AEdit: TLspWorkspaceEdit; const ADefFilePath: string; ADefLine: Integer;
      const AImplFiles: TArray<string>): TRenamePreviewItems;
  public
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    procedure Execute;

    /// <summary>Triggered by the unit-rename watcher when the IDE renames
    ///  a unit (e.g. via File > Save As). Opens the standard rename dialog
    ///  pre-filled with (OldUnitName -> NewUnitName); the user sees the
    ///  preview list and confirms just like with a normal identifier
    ///  rename.</summary>
    procedure ExecuteForUnit(const AOldUnitName, ANewUnitName: string);

    function GetMenuText: string;
  end;

var
  WizardInstance: TLspRenameWizard;

implementation

{ TLspRenameWizard - IOTANotifier }

procedure TLspRenameWizard.AfterSave; begin end;
procedure TLspRenameWizard.BeforeSave; begin end;
procedure TLspRenameWizard.Destroyed; begin end;
procedure TLspRenameWizard.Modified; begin end;

function TLspRenameWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.RenameWizard';
end;

function TLspRenameWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Rename';
end;

function TLspRenameWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspRenameWizard.GetMenuText: string;
begin
  Result := 'Rename identifier...';
end;

procedure TLspRenameWizard.Execute;
begin
  FUnitRenameMode := False;
  FContext := TEditorHelper.GetCurrentContext;

  if not FContext.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak +
      'Please place the cursor on an identifier.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  FDialog := TRenameDialog.CreateDialog(Application.MainForm, FContext.WordAtCursor);
  try
    FDialog.OnPreviewRequested := DoPreview;
    FDialog.SetCheckContext(FContext.FileName, TEditorHelper.GetProjectSourceFiles);
    if FDialog.ShowModal = mrOk then
      ApplyFEdit;
  finally
    FDialog.Free;
    FDialog := nil;
  end;
end;

procedure TLspRenameWizard.ExecuteForUnit(const AOldUnitName, ANewUnitName: string);
begin
  if (AOldUnitName = '') or (ANewUnitName = '') or
     SameText(AOldUnitName, ANewUnitName) then Exit;

  FUnitRenameMode := True;
  try
    // Synthetic context: only WordAtCursor is actually used in unit mode.
    FContext := Default(TEditorContext);
    FContext.WordAtCursor := AOldUnitName;
    FContext.IsValid := True;

    FDialog := TRenameDialog.CreateDialog(Application.MainForm, AOldUnitName);
    try
      FDialog.OnPreviewRequested := DoPreview;
      FDialog.SetNewName(ANewUnitName);
      if FDialog.ShowModal = mrOk then
        ApplyFEdit;
    finally
      FDialog.Free;
      FDialog := nil;
    end;
  finally
    FUnitRenameMode := False;
  end;
end;

procedure TLspRenameWizard.ApplyFEdit;
begin
  if Length(FEdit.FileEdits) = 0 then Exit;

  var AppliedCount := 0;
  var FailedCount := 0;
  var AffectedFiles := TList<string>.Create;
  try
    // Apply via IDE editor API (undoable!)
    // Per file, sort edits line-descending so earlier edits do not
    // shift later edits' positions.
    for var FE in FEdit.FileEdits do
    begin
      var SortedEdits := Copy(FE.Edits);
      TArray.Sort<TLspTextEdit>(SortedEdits,
        TComparer<TLspTextEdit>.Construct(
          function(const L, R: TLspTextEdit): Integer
          begin
            Result := R.Range.Start.Line - L.Range.Start.Line;
            if Result = 0 then
              Result := R.Range.Start.Character - L.Range.Start.Character;
          end));

      for var Edit in SortedEdits do
      begin
        if TEditorHelper.ApplyEditViaEditor(FE.FilePath, Edit.Range.Start.Line, Edit.Range.Start.Character,
          FContext.WordAtCursor, Edit.NewText) then
          Inc(AppliedCount)
        else
          Inc(FailedCount);
      end;

      if not AffectedFiles.Contains(FE.FilePath) then
        AffectedFiles.Add(FE.FilePath);
    end;

    // Inform LSP about the changes (not needed in unit-rename mode
    // because LSP was not used for verification, but harmless).
    if TLspManager.Instance.IsAlive then
    begin
      try
        var Client := TLspManager.Instance.GetClient(FContext.ProjectRoot, FContext.ProjectFile, TEditorHelper.FindDelphiLspJson);
        for var F in AffectedFiles do
          Client.RefreshDocument(F);
      except
        // LSP refresh is best-effort
      end;
    end;

    if FailedCount = 0 then
      MessageDlg(Format('%d change(s) applied successfully (Ctrl+Z to undo).',
        [AppliedCount]), mtInformation, [mbOK], 0)
    else
      MessageDlg(Format('%d applied, %d failed.',
        [AppliedCount, FailedCount]), mtWarning, [mbOK], 0);
  finally
    AffectedFiles.Free;
  end;
end;

procedure TLspRenameWizard.DoPreview(Sender: TObject);
begin
  if FUnitRenameMode then
    DoPreviewForUnit
  else
    DoPreviewForIdentifier;
end;

procedure TLspRenameWizard.DoPreviewForUnit;
var
  NewName: string;
  ProjFiles: TArray<string>;
  Candidates: TArray<TRenameCandidate>;
  PreviewItems: TRenamePreviewItems;
  TotalEdits: Integer;
begin
  NewName := FDialog.GetNewName;
  if NewName = '' then
  begin
    MessageDlg('Please enter a new name.', mtWarning, [mbOK], 0);
    Exit;
  end;
  if SameText(NewName, FContext.WordAtCursor) then
  begin
    MessageDlg('The new name is identical to the old one.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // Save all dirty files so the text scan sees their current state
  TEditorHelper.SaveAllFiles;

  FDialog.SetBusy(True);
  FDiagLog := '';
  try
    ProjFiles := TEditorHelper.GetProjectSourceFiles;
    FDiagLog :=
      '=== Diagnostics (Unit Rename) ===' + sLineBreak +
      'Old unit name: ' + FContext.WordAtCursor + sLineBreak +
      'New unit name: ' + NewName + sLineBreak +
      'Project files: ' + IntToStr(Length(ProjFiles)) + sLineBreak +
      sLineBreak +
      'Text search (whole-word, skipping strings and comments)...' + sLineBreak;

    FDialog.SetStatus(Format('Scanning %d project file(s)...', [Length(ProjFiles)]));
    Candidates := FindCandidates(FContext.WordAtCursor, ProjFiles);
    FDiagLog := FDiagLog + 'Text candidates: ' + IntToStr(Length(Candidates)) + sLineBreak;

    if Length(Candidates) = 0 then
    begin
      FDialog.SetPreviewItems(nil);
      FDialog.SetDetailsText(Format('No references to unit "%s" found in the project.',
        [FContext.WordAtCursor]) + sLineBreak + sLineBreak + FDiagLog);
      FDialog.SetStatus('Done - no matches.');
      FDialog.SetBusy(False);
      Exit;
    end;

    // No LSP verification in unit-rename mode: a unit name in 'uses'
    // clauses (and qualified references like 'OldUnit.Something') is
    // unambiguous, and LSP's textDocument/definition would typically
    // point at the renamed file which no longer exists.
    FEdit := BuildEditFromCandidates(Candidates, FContext.WordAtCursor, NewName);

    // Preview items (Kind is always "Call" in unit-rename mode because
    // the uses-clause occurrences don't start with a method keyword).
    PreviewItems := BuildPreviewItems(FEdit, '', -1, nil);

    TotalEdits := 0;
    for var FE in FEdit.FileEdits do
      Inc(TotalEdits, Length(FE.Edits));

    FDialog.SetPreviewItems(PreviewItems);
    FDialog.SetDetailsText(FDiagLog);
    FDialog.EnableRename(True);
    FDialog.SetStatus(Format('Done: %d change(s) in %d file(s).', [TotalEdits, Length(FEdit.FileEdits)]));
  except
    on E: Exception do
    begin
      FDiagLog := FDiagLog + sLineBreak + 'EXCEPTION: ' + E.ClassName + ': ' + E.Message;
      FDialog.SetPreviewItems(nil);
      FDialog.SetDetailsText(FDiagLog);
      FDialog.SetStatus('An error occurred.');
    end;
  end;
  FDialog.SetBusy(False);
end;

function TLspRenameWizard.BuildEditFromCandidates(const ACandidates: TArray<TRenameCandidate>;
  const AOldName, ANewName: string): TLspWorkspaceEdit;
var
  FileMap: TDictionary<string, TList<TLspTextEdit>>;
  TE: TLspTextEdit;
  Idx: Integer;
begin
  FileMap := TDictionary<string, TList<TLspTextEdit>>.Create;
  try
    for var C in ACandidates do
    begin
      TE.Range.Start.Line := C.Line;
      TE.Range.Start.Character := C.Col;
      TE.Range.End_.Line := C.Line;
      TE.Range.End_.Character := C.Col + Length(AOldName);
      TE.NewText := ANewName;

      if not FileMap.ContainsKey(C.FilePath) then
        FileMap.Add(C.FilePath, TList<TLspTextEdit>.Create);
      FileMap[C.FilePath].Add(TE);
    end;

    SetLength(Result.FileEdits, FileMap.Count);
    Idx := 0;
    for var Pair in FileMap do
    begin
      Result.FileEdits[Idx].FilePath := Pair.Key;
      Result.FileEdits[Idx].Edits := Pair.Value.ToArray;
      Inc(Idx);
    end;

    for var Pair in FileMap do
      Pair.Value.Free;
  finally
    FileMap.Free;
  end;
end;

procedure TLspRenameWizard.DoPreviewForIdentifier;
var
  NewName, DelphiLspJson, RootPath, DefFilePath: string;
  ProjFiles: TArray<string>;
  Candidates, ImplCandidates: TArray<TRenameCandidate>;
  Client: TLspClient;
begin
  NewName := FDialog.GetNewName;
  if NewName = '' then
  begin
    MessageDlg('Please enter a new name.', mtWarning, [mbOK], 0);
    Exit;
  end;
  if SameText(NewName, FContext.WordAtCursor) then
  begin
    MessageDlg('The new name is identical to the old one.', mtWarning, [mbOK], 0);
    Exit;
  end;

  DelphiLspJson := TEditorHelper.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    MessageDlg('No .delphilsp.json found.' + sLineBreak +
      'Please enable: Tools > Options > Editor > Language > ' +
      'Code Insight > "Generate LSP Config".',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  // Save all modified files so the LSP sees current data
  TEditorHelper.SaveAllFiles;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  FDialog.SetBusy(True);
  FDiagLog := '';
  try
    // Save all unsaved files (so LSP sees the current state)
    FDialog.SetStatus('Saving all files...');
    TEditorHelper.SaveAllFiles;

    FDiagLog := '=== Diagnostics ===' + sLineBreak +
      'File: ' + FContext.FileName + sLineBreak +
      'Identifier: ' + FContext.WordAtCursor + sLineBreak +
      'Position: ' + IntToStr(FContext.Line) + ':' + IntToStr(FContext.Column) + sLineBreak +
      'delphilsp.json: ' + DelphiLspJson + sLineBreak + sLineBreak;

    // Get project files from the IDE
    ProjFiles := TEditorHelper.GetProjectSourceFiles;
    FDiagLog := FDiagLog + 'Project files: ' + IntToStr(Length(ProjFiles)) + sLineBreak + sLineBreak;

    // Phase 1: text search over project files
    FDialog.SetStatus('Phase 1: text search...');
    Candidates := FindCandidates(FContext.WordAtCursor, ProjFiles);
    FDiagLog := FDiagLog + 'Text candidates: ' + IntToStr(Length(Candidates)) + sLineBreak + sLineBreak;

    if Length(Candidates) = 0 then
    begin
      FDialog.SetPreviewItems(nil);
      FDialog.SetDetailsText('No occurrences found.' + sLineBreak + sLineBreak + FDiagLog);
      FDialog.SetStatus('Done - no matches.');
      FDialog.SetBusy(False);
      Exit;
    end;

    // Phase 2: start LSP (singleton - first call slow, later calls instant)
    var WasRunning := TLspManager.Instance.IsAlive;
    if WasRunning then
      FDialog.SetStatus('LSP already running. Opening file...')
    else
      FDialog.SetStatus('Starting LSP server (one-time)...');

    Client := TLspManager.Instance.GetClient(
      RootPath, FContext.ProjectFile, DelphiLspJson);

    // Refresh document (didClose+didOpen when LSP is already running,
    // just didOpen on first start - so LSP always has current file content)
    if WasRunning then
      Client.RefreshDocument(FContext.FileName)
    else
      Client.RefreshDocument(FContext.FileName);

    // On first start, wait for readiness
    if not WasRunning then
    begin
      var LspLine := FContext.Line - 1;
      var LspCol := FContext.Column - 1;      for var Retry := 1 to 30 do
      begin
        FDialog.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
        Application.ProcessMessages;
        try
          var H := Client.GetHover(FContext.FileName, LspLine, LspCol);
          if H <> '' then Break;
          var D := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
          if Length(D) > 0 then Break;
        except end;
        Sleep(1000);
      end;
    end
    else
      Sleep(500);

    // Phase 2b: find declaration
    FDialog.SetStatus('Finding declaration...');

    var LspLine := FContext.Line - 1;
    var LspCol := FContext.Column - 1;
    var DefLocs := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
    var DefLine := 0;
    var DefCol := 0;

    if Length(DefLocs) > 0 then
    begin
      DefFilePath := TLspUri.FileUriToPath(DefLocs[0].Uri);
      DefLine := DefLocs[0].Range.Start.Line;
      DefCol := DefLocs[0].Range.Start.Character;
    end
    else
      DefFilePath := FContext.FileName;

    FDiagLog := FDiagLog + 'Declaration: ' + DefFilePath + ':' + IntToStr(DefLine + 1) + ':' + IntToStr(DefCol + 1) + sLineBreak;

    // Phase 2c: find interface/class method implementations.
    // Text-based scan over all project files with syntax filter on lines
    // like 'procedure TClass.Method'. Only classes that implement the
    // container (owner) type of the method are kept.
    FDialog.SetStatus('Searching for interface implementations...');

    var OwnerType := TImplementationFinder.FindContainingType(DefFilePath, DefLine);
    FDiagLog := FDiagLog + 'Owner type for impl verification: ' +
      IfThen(OwnerType <> '', OwnerType, '(none - free procedure, impl scan skipped)') + sLineBreak;

    // Only do the class-method implementation scan when the declaration is
    // actually a class/interface method. For free procedures there are no
    // "TClass.Method" implementations to find; scanning with an empty owner
    // type would otherwise accept unrelated same-named class methods in
    // other units.
    if OwnerType <> '' then
      ImplCandidates := FindImplementations(FContext.WordAtCursor, ProjFiles, OwnerType)
    else
      ImplCandidates := nil;

    // Track impl files separately - used in VerifyWithLsp as additional
    // accepted GotoDefinition targets (DelphiLSP resolves class-bound
    // positions to the class level instead of the interface level).
    var ImplFilesList := TList<string>.Create;
    var ImplFilesArray: TArray<string>;
    try
      if Length(ImplCandidates) > 0 then
      begin
        FDiagLog := FDiagLog + 'Interface implementations: ' + IntToStr(Length(ImplCandidates)) + sLineBreak;
        for var IC in ImplCandidates do
        begin
          FDiagLog := FDiagLog + '  ' + ExtractFileName(IC.FilePath) + ':' + IntToStr(IC.Line + 1) + sLineBreak;
          if ImplFilesList.IndexOf(IC.FilePath) < 0 then
            ImplFilesList.Add(IC.FilePath);
        end;
        FDiagLog := FDiagLog + sLineBreak;

        // Add implementations to the candidates
        // (Dedup is handled later in VerifyWithLsp by line+column)
        var Combined := TList<TRenameCandidate>.Create;
        try
          for var C in Candidates do Combined.Add(C);
          for var IC in ImplCandidates do Combined.Add(IC);
          Candidates := Combined.ToArray;
        finally
          Combined.Free;
        end;
      end;

      // Phase 3: LSP verification
      FDialog.SetStatus(Format('%d candidate(s). Verifying...', [Length(Candidates)]));

      ImplFilesArray := ImplFilesList.ToArray;
      FEdit := VerifyWithLsp(Candidates, FContext.WordAtCursor, NewName, DefFilePath, ImplFilesArray, Client);
    finally
      ImplFilesList.Free;
    end;

    if Length(FEdit.FileEdits) = 0 then
    begin
      FDialog.SetPreviewItems(nil);
      FDialog.SetDetailsText(FDiagLog);
      FDialog.SetStatus('Done - no verified matches.');
      FDialog.SetBusy(False);
      Exit;
    end;

    // Build structured preview for the ListView
    var PreviewItems := BuildPreviewItems(FEdit, DefFilePath, DefLine, ImplFilesArray);

    // Count for the status line
    var TotalEdits := 0;
    for var FE in FEdit.FileEdits do
      Inc(TotalEdits, Length(FE.Edits));

    FDialog.SetPreviewItems(PreviewItems);
    FDialog.SetDetailsText(FDiagLog);
    FDialog.EnableRename(True);
    FDialog.SetStatus(Format('Done: %d change(s) in %d file(s).', [TotalEdits, Length(FEdit.FileEdits)]));
  except
    on E: Exception do
    begin
      FDiagLog := FDiagLog + sLineBreak + 'EXCEPTION: ' + E.ClassName + ': ' + E.Message;
      FDialog.SetPreviewItems(nil);
      FDialog.SetDetailsText(FDiagLog);
      FDialog.SetStatus('An error occurred.');
    end;
  end;
  FDialog.SetBusy(False);
end;

{ Helper functions }

function TLspRenameWizard.IsInStringOrComment(const ALine: string; APos: Integer): Boolean;
var
  I: Integer;
  InString: Boolean;
begin
  Result := False;
  for I := 1 to APos - 1 do
    if (I < Length(ALine)) and (ALine[I] = '/') and (ALine[I + 1] = '/') then
      Exit(True);
  var BraceDepth := 0;
  for I := 1 to APos - 1 do
  begin
    if ALine[I] = '{' then Inc(BraceDepth)
    else if ALine[I] = '}' then Dec(BraceDepth);
  end;
  if BraceDepth > 0 then Exit(True);
  var PSDepth := 0;
  for I := 1 to APos - 2 do
  begin
    if (ALine[I] = '(') and (ALine[I+1] = '*') then Inc(PSDepth)
    else if (ALine[I] = '*') and (ALine[I+1] = ')') then Dec(PSDepth);
  end;
  if PSDepth > 0 then Exit(True);
  InString := False;
  for I := 1 to APos - 1 do
    if ALine[I] = '''' then InString := not InString;
  if InString then Exit(True);
end;

{ Text search }

function TLspRenameWizard.FindCandidates(const AOldName: string; const AFiles: TArray<string>): TArray<TRenameCandidate>;
var
  CandidateList: TList<TRenameCandidate>;
  F, Line, RawContent: string;
  Lines: TArray<string>;
  UpperOldName: string;
  LineIdx, SearchPos, FoundPos, AfterPos: Integer;
  BeforeOk, AfterOk: Boolean;
  Candidate: TRenameCandidate;
begin
  UpperOldName := UpperCase(AOldName);
  CandidateList := TList<TRenameCandidate>.Create;
  try
    FDialog.SetProgress(0, Length(AFiles));

    for var FileIdx := 0 to High(AFiles) do
    begin
      F := AFiles[FileIdx];
      if (FileIdx mod 10 = 0) then
        FDialog.SetProgress(FileIdx + 1, Length(AFiles));

      try
        RawContent := ReadDelphiFile(F);
        if Pos(UpperOldName, UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(F);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        Line := Lines[LineIdx];
        SearchPos := 1;
        while SearchPos <= Length(Line) do
        begin
          FoundPos := Pos(UpperOldName, UpperCase(Copy(Line, SearchPos)));
          if FoundPos = 0 then Break;
          FoundPos := SearchPos + FoundPos - 1;

          BeforeOk := (FoundPos = 1) or
            not CharInSet(Line[FoundPos - 1], ['A'..'Z','a'..'z','0'..'9','_']);
          AfterPos := FoundPos + Length(AOldName);
          AfterOk := (AfterPos > Length(Line)) or
            not CharInSet(Line[AfterPos], ['A'..'Z','a'..'z','0'..'9','_']);

          if BeforeOk and AfterOk and not IsInStringOrComment(Line, FoundPos) then
          begin
            Candidate.FilePath := F;
            Candidate.Line := LineIdx;
            Candidate.Col := FoundPos - 1;
            Candidate.OldText := Copy(Line, FoundPos, Length(AOldName));
            CandidateList.Add(Candidate);
          end;
          SearchPos := FoundPos + Length(AOldName);
        end;
      end;
    end;

    FDialog.SetProgress(Length(AFiles), Length(AFiles));
    Result := CandidateList.ToArray;
  finally
    CandidateList.Free;
  end;
end;

{ Find interface implementations }

function TLspRenameWizard.FindImplementations(const AOldName: string; const AProjectFiles: TArray<string>;
  const AOwnerType: string): TArray<TRenameCandidate>;
var
  Items: TFindReferenceItems;
  ResultList: TList<TRenameCandidate>;
  Candidate: TRenameCandidate;
begin
  FDiagLog := FDiagLog + 'FindImplementations (text+syntax scan, owner=' +
    IfThen(AOwnerType <> '', AOwnerType, '(all)') + ')' + sLineBreak;

  // Shared finder (also used by the Find Implementations wizard).
  // Text scan over all project files, filtered by class method impl
  // syntax, with owner type verification.
  Items := TImplementationFinder.FindByProjectScan(AProjectFiles, AOldName, AOwnerType, nil);

  FDiagLog := FDiagLog + '  Result: ' + IntToStr(Length(Items)) + ' implementation(s)' + sLineBreak;

  ResultList := TList<TRenameCandidate>.Create;
  try
    for var Item in Items do
    begin
      Candidate.FilePath := Item.FilePath;
      Candidate.Line := Item.Line;
      Candidate.Col := Item.Col;
      Candidate.OldText := AOldName;
      ResultList.Add(Candidate);
      FDiagLog := FDiagLog + '    ' + ExtractFileName(Item.FilePath) + ':' + IntToStr(Item.Line + 1) + sLineBreak;
    end;
    FDiagLog := FDiagLog + sLineBreak;
    Result := ResultList.ToArray;
  finally
    ResultList.Free;
  end;
end;

{ Preview building }

function TLspRenameWizard.BuildPreviewItems(const AEdit: TLspWorkspaceEdit; const ADefFilePath: string; ADefLine: Integer;
  const AImplFiles: TArray<string>): TRenamePreviewItems;

  function IsKnownImplFile(const APath: string): Boolean;
  begin
    Result := False;
    for var F in AImplFiles do
      if SameText(ExpandFileName(APath), ExpandFileName(F)) then
        Exit(True);
  end;

  function LineStartsWithMethodKeyword(const ALine: string): Boolean;
  var
    Trimmed: string;
  begin
    Trimmed := LowerCase(TrimLeft(ALine));
    Result :=
      StartsStr('procedure ', Trimmed) or
      StartsStr('function ', Trimmed) or
      StartsStr('constructor ', Trimmed) or
      StartsStr('destructor ', Trimmed) or
      StartsStr('operator ', Trimmed) or
      StartsStr('class procedure ', Trimmed) or
      StartsStr('class function ', Trimmed) or
      StartsStr('class constructor ', Trimmed) or
      StartsStr('class destructor ', Trimmed) or
      StartsStr('class operator ', Trimmed);
  end;

  function DetermineKind(const AFilePath: string; ALine, ACol: Integer;
    const AOrigLine: string): string;
  var
    IsHeader, DotBefore: Boolean;
  begin
    IsHeader := LineStartsWithMethodKeyword(AOrigLine);
    DotBefore := (ACol > 0) and (ACol <= Length(AOrigLine)) and (AOrigLine[ACol] = '.');

    if IsHeader and DotBefore then
      Exit('Implementation');

    if IsHeader then
    begin
      if SameText(ExpandFileName(AFilePath), ExpandFileName(ADefFilePath)) and (ALine = ADefLine) then
        Exit('Interface declaration');
      if IsKnownImplFile(AFilePath) then
        Exit('Class declaration');
      Exit('Declaration');
    end;

    // Non-header line: either a call / use, or a declaration within an
    // interface/class block (e.g. 'property Bar: T read Bar;').
    if SameText(ExpandFileName(AFilePath), ExpandFileName(ADefFilePath)) then
      Exit('Interface reference');
    Exit('Call');
  end;

var
  List: TList<TRenamePreviewItem>;
  Item: TRenamePreviewItem;
  Lines: TArray<string>;
  LineNo, StartCol, EndCol: Integer;
  OrigLine: string;
begin
  List := TList<TRenamePreviewItem>.Create;
  try
    for var FE in AEdit.FileEdits do
    begin
      try
        Lines := ReadDelphiFileLines(FE.FilePath);
      except
        Continue;
      end;

      for var Edit in FE.Edits do
      begin
        LineNo := Edit.Range.Start.Line;
        if (LineNo < 0) or (LineNo >= Length(Lines)) then Continue;

        OrigLine := Lines[LineNo];
        StartCol := Edit.Range.Start.Character;
        EndCol   := Edit.Range.End_.Character;

        Item.FilePath := FE.FilePath;
        Item.Line := LineNo;
        Item.Col := StartCol;
        Item.OriginalLine := OrigLine;

        // Preview: swap the replaced section with NewText.
        // Col values are 0-based (LSP), Pascal strings are 1-based.
        if (StartCol >= 0) and (StartCol <= Length(OrigLine)) and (EndCol >= StartCol) and (EndCol <= Length(OrigLine)) then
          Item.PreviewLine := Copy(OrigLine, 1, StartCol) + Edit.NewText + Copy(OrigLine, EndCol + 1, MaxInt)
        else
          Item.PreviewLine := OrigLine;

        Item.Kind := DetermineKind(FE.FilePath, LineNo, StartCol, OrigLine);

        List.Add(Item);
      end;
    end;
    Result := List.ToArray;
  finally
    List.Free;
  end;
end;

{ LSP verification }

function TLspRenameWizard.VerifyWithLsp(const ACandidates: TArray<TRenameCandidate>;
  const AOldName, ANewName, ADefFilePath: string; const AImplFiles: TArray<string>; AClient: TLspClient): TLspWorkspaceEdit;

  /// <summary>Checks whether APath is one of the impl files.</summary>
  function IsImplFile(const APath: string): Boolean;
  begin
    Result := False;
    for var F in AImplFiles do
      if SameText(ExpandFileName(APath), ExpandFileName(F)) then
        Exit(True);
  end;

var
  FileMap: TDictionary<string, TList<TLspTextEdit>>;
  LastOpenedFile: string;
  VerifiedCount, SkippedCount, I: Integer;
  C: TRenameCandidate;
  TextEdit: TLspTextEdit;
begin
  FileMap := TDictionary<string, TList<TLspTextEdit>>.Create;
  try
    LastOpenedFile := '';
    VerifiedCount := 0;
    SkippedCount := 0;

    FDialog.SetProgress(0, Length(ACandidates));

    for I := 0 to High(ACandidates) do
    begin
      C := ACandidates[I];
      FDialog.SetProgress(I + 1, Length(ACandidates));
      if (I mod 3 = 0) then
        FDialog.SetStatus(Format('Verifying %d/%d (ok:%d skip:%d)', [I + 1, Length(ACandidates), VerifiedCount, SkippedCount]));

      // Open file on the LSP
      if not SameText(C.FilePath, LastOpenedFile) then
      begin
        AClient.RefreshDocument(C.FilePath);
        Sleep(300);
        LastOpenedFile := C.FilePath;
      end;

      var Matches := False;
      var DiagLine := Format('  [%d] %s:%d:%d => ', [I, ExtractFileName(C.FilePath), C.Line + 1, C.Col + 1]);

      try
        var Defs := AClient.GotoDefinition(C.FilePath, C.Line, C.Col);

        if Length(Defs) = 0 then
        begin
          // Null: either on the declaration itself, or LSP cannot
          // resolve. We accept if the candidate lies in the interface
          // file or in a known impl file.
          if SameText(ExpandFileName(C.FilePath), ExpandFileName(ADefFilePath)) then
          begin
            Matches := True;
            DiagLine := DiagLine + 'null -> MATCH (declaration file)';
          end
          else if IsImplFile(C.FilePath) then
          begin
            Matches := True;
            DiagLine := DiagLine + 'null -> MATCH (impl file)';
          end
          else
            DiagLine := DiagLine + 'null -> SKIP';
        end
        else
        begin
          var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
          DiagLine := DiagLine + ExtractFileName(DefPath) + ':' + IntToStr(Defs[0].Range.Start.Line + 1);
          if SameText(ExpandFileName(DefPath), ExpandFileName(ADefFilePath)) then
          begin
            Matches := True;
            DiagLine := DiagLine + ' -> MATCH';
          end
          else if IsImplFile(DefPath) then
          begin
            // DelphiLSP resolves class-bound positions (class method decl,
            // calls via class variable, impl body) to the class level
            // rather than the interface level. If the target lies in one
            // of our known impl files, the candidate is a valid reference
            // point for the interface method.
            Matches := True;
            DiagLine := DiagLine + ' -> MATCH (impl file)';
          end
          else
            DiagLine := DiagLine + ' -> SKIP (expected: ' +
              ExtractFileName(ADefFilePath) + ')';
        end;
      except
        on E: Exception do
          DiagLine := DiagLine + 'ERROR: ' + E.Message + ' -> SKIP';
      end;

      FDiagLog := FDiagLog + DiagLine + sLineBreak;

      if Matches then
      begin
        TextEdit.Range.Start.Line := C.Line;
        TextEdit.Range.Start.Character := C.Col;
        TextEdit.Range.End_.Line := C.Line;
        TextEdit.Range.End_.Character := C.Col + Length(AOldName);
        TextEdit.NewText := ANewName;

        if not FileMap.ContainsKey(C.FilePath) then
          FileMap.Add(C.FilePath, TList<TLspTextEdit>.Create);

        // Dedup: do not add the same (line, column) twice, otherwise a
        // position would be edited twice on apply. Happens when an impl
        // candidate already appears as a text candidate.
        var Exists := False;
        for var Existing in FileMap[C.FilePath] do
          if (Existing.Range.Start.Line = TextEdit.Range.Start.Line) and
             (Existing.Range.Start.Character = TextEdit.Range.Start.Character) then
          begin
            Exists := True;
            Break;
          end;

        if not Exists then
        begin
          FileMap[C.FilePath].Add(TextEdit);
          Inc(VerifiedCount);
        end
        else
          DiagLine := DiagLine + '  (dedup: already added)';
      end
      else
        Inc(SkippedCount);
    end;

    FDialog.SetProgress(Length(ACandidates), Length(ACandidates));

    SetLength(Result.FileEdits, FileMap.Count);
    var Idx := 0;
    for var Pair in FileMap do
    begin
      Result.FileEdits[Idx].FilePath := Pair.Key;
      Result.FileEdits[Idx].Edits := Pair.Value.ToArray;
      Inc(Idx);
    end;

    for var Pair in FileMap do
      Pair.Value.Free;
  finally
    FileMap.Free;
  end;
end;

end.
