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
  System.Generics.Collections, System.Generics.Defaults, System.RegularExpressions,
  Vcl.Forms, Vcl.Dialogs,
  {$IFNDEF STANDALONE_BUILD} ToolsAPI, {$ENDIF}
  Expert.EditorHelperIntf,
  Expert.RenameDialog, Expert.LspManager, Expert.ImplementationFinder, Expert.FindReferencesDialog,
  Expert.UnitIndex, Expert.UnitUsageProbe, Expert.ScopeFiles, Expert.DfmRename,
  Expert.InterfaceLinks,
  Expert.IncludeExpansion,
  Lsp.Uri, Lsp.Protocol,
  Lsp.Client, Rename.WorkspaceEdit, Delphi.FileEncoding, Expert.UsesEditor,
  Expert.SafeDeletePlan;

type
  TRenameCandidate = record
    FilePath: string;
    Line: Integer;      // 0-based
    Col: Integer;       // 0-based
    OldText: string;
  end;

  /// <summary>Form-file occurrences of the renamed symbol in ONE .dfm/.fmx.</summary>
  TRenameFormPlan = record
    FormFile: string;
    PasFile: string;
    Hits: TArray<TDfmHit>;
  end;

  /// <summary>What the rename pipeline needs from its front end. The
  ///  dialog is one (TDialogRenameHost); the MCP bridge runs the SAME
  ///  pipeline without a window through its own host - so preview, LSP
  ///  verification, implementations and form files stay identical.</summary>
  IRenameHost = interface
    ['{3C1F7B2A-94D6-4E0B-A8C5-6E2D19F07B41}']
    function GetNewName: string;
    /// <summary>Copy every affected file into a backup directory before
    ///  the rename writes anything (the dialog's "Create backup").</summary>
    function CreateBackup: Boolean;
    function Scope: TRenameScope;
    function SelectedUnits: TArray<string>;
    function IncludeOpenUnits: Boolean;
    function IncludeUsedUnits: Boolean;
    function ScanCancelled: Boolean;
    procedure SetBusy(ABusy: Boolean);
    procedure SetStatus(const AText: string);
    procedure SetProgress(AValue, AMax: Integer);
    procedure SetPreviewItems(const AItems: TRenamePreviewItems);
    procedure SetDetailsText(const AText: string);
    procedure EnableRename(AEnabled: Boolean);
    /// <summary>A message the dialog shows in a box (validation, outcome).
    ///  </summary>
    procedure Notify(const AText: string; AWarning: Boolean);
  end;

  TDialogRenameHost = class(TInterfacedObject, IRenameHost)
  private
    FDialog: TRenameDialog;
  public
    constructor Create(ADialog: TRenameDialog);
    function GetNewName: string;
    function CreateBackup: Boolean;
    function Scope: TRenameScope;
    function SelectedUnits: TArray<string>;
    function IncludeOpenUnits: Boolean;
    function IncludeUsedUnits: Boolean;
    function ScanCancelled: Boolean;
    procedure SetBusy(ABusy: Boolean);
    procedure SetStatus(const AText: string);
    procedure SetProgress(AValue, AMax: Integer);
    procedure SetPreviewItems(const AItems: TRenamePreviewItems);
    procedure SetDetailsText(const AText: string);
    procedure EnableRename(AEnabled: Boolean);
    procedure Notify(const AText: string; AWarning: Boolean);
  end;

  TLspRenameWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private
    FDialog: TRenameDialog;
    FHost: IRenameHost;
    // Form files (see CollectFormEdits). The LSP knows nothing about them,
    // so they are planned separately and applied BEFORE the source edits:
    // an open form is renamed through its designer, which also renames the
    // declaration in the source.
    FFormPlans: TArray<TRenameFormPlan>;
    FFormTarget: TDfmTargetKind;
    FFormDeclFile: string;
    FContext: TEditorContext;
    FEdit: TLspWorkspaceEdit;
    /// <summary>Hits in include files DelphiLSP gave no answer for: NOT
    ///  renamed, but listed in the preview (never dropped silently).</summary>
    FUnverified: TArray<TRenameCandidate>;
    FUnverifiedWhy: TArray<string>;   // parallel to FUnverified
    FDiagLog: string;
    /// <summary>True when the dialog was opened for a unit rename
    ///  (triggered by the IDE module notifier). In that mode the preview
    ///  skips LSP verification and just does a text-based scan.</summary>
    FUnitRenameMode: Boolean;
    procedure CollectFormEdits(const AFiles: TArray<string>;
      const ADefFile: string; ADefLine: Integer; const AOwnerType,
      AOldName, ANewName: string);
    function FormHitKindAt(const AFile: string; ALine, ACol: Integer;
      out AKind: string): Boolean;
    procedure DoPreview(Sender: TObject);
    procedure DoPreviewForIdentifier;
    procedure DoPreviewForUnit;
    procedure ApplyFEdit;
    function BuildEditFromCandidates(const ACandidates: TArray<TRenameCandidate>;
      const AOldName, ANewName: string): TLspWorkspaceEdit;

    function FindCandidates(const AOldName: string; const AFiles: TArray<string>): TArray<TRenameCandidate>;
    /// <summary>'' or a report of places where ANewName already exists
    ///  (files of the edit, a member of the owner type, a used unit).</summary>
    function NameConflictNote(const ANewName, AOwnerType, ADefFile: string): string;
    function VerifyWithLsp(const ACandidates: TArray<TRenameCandidate>; const AOldName, ANewName: string;
      AIncludes: TLspIncludeContext;
      const ATargets: TLspSymbolTargets; AClient: TLspClient): TLspWorkspaceEdit;

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

    /// <summary>Triggered by the unit-rename watcher when the IDE renames
    ///  a unit (e.g. via File > Save As). Opens the standard rename dialog
    ///  pre-filled with (OldUnitName -> NewUnitName); the user sees the
    ///  preview list and confirms just like with a normal identifier
    ///  rename.</summary>
    procedure ExecuteForUnit(const AOldUnitName, ANewUnitName: string);

    /// <summary>Identifier rename WITHOUT the dialog (MCP bridge): runs the
    ///  normal preview for AContext (file, 1-based line/column, the
    ///  identifier) against AHost. True when there is something to apply.
    ///  Main thread; saves all files first, like the dialog.</summary>
    function PreviewHeadless(const AContext: TEditorContext;
      const AHost: IRenameHost): Boolean;
    /// <summary>Applies the edits of the last PreviewHeadless; the outcome
    ///  arrives through AHost.Notify.</summary>
    procedure ApplyHeadless(const AHost: IRenameHost);
  end;

var
  WizardInstance: TLspRenameWizard;

implementation


uses
  Winapi.Windows, Expert.PascalScanner, Expert.IdentifierCheck;
// True when the identifier at (ALine0, ACol0) directly follows a declaring
// keyword - 'procedure X', 'class function X', 'constructor X',
// 'destructor X', 'property X' - i.e. the caret is ON a declaration.
function CaretOnDeclaration(const AFile: string; ALine0, ACol0: Integer): Boolean;
var
  Lines: TArray<string>;
  S, Before: string;
  P: Integer;
begin
  Result := False;
  try
    Lines := ReadDelphiFileLines(AFile);
  except
    Exit;
  end;
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  S := Lines[ALine0];
  // start of the identifier under the caret (1-based)
  P := ACol0 + 1;
  if (P < 1) or (P > Length(S) + 1) then Exit;
  while (P > 1) and (CharInSet(S[P - 1], ['A'..'Z', 'a'..'z', '0'..'9', '_'])) do
    Dec(P);
  Before := LowerCase(Trim(Copy(S, 1, P - 1)));
  for var K in ['procedure', 'function', 'constructor', 'destructor', 'property'] do
    if (Before = K) or EndsStr(' ' + K, Before) then
      Exit(True);
end;

{$IFNDEF STANDALONE_BUILD}
{ TLspRenameWizard - IOTANotifier / IOTAWizard / IOTAMenuWizard stubs.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

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
{$ENDIF}
procedure TLspRenameWizard.Execute;
begin
  FUnitRenameMode := False;
  FContext := Editor.GetCurrentContext;

  if not FContext.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak +
      'Please place the cursor on an identifier.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  FDialog := TRenameDialog.CreateDialog(Application.MainForm, FContext.WordAtCursor);
  FHost := TDialogRenameHost.Create(FDialog);
  try
    FDialog.OnPreviewRequested := DoPreview;
    FDialog.SetCheckContext(FContext.FileName, Editor.GetProjectSourceFiles);
    TLspManager.Instance.ApplyStatusToCaption(FDialog);
    if FDialog.ShowModal = mrOk then
      ApplyFEdit;
  finally
    FHost := nil;
    FDialog.Free;
    FDialog := nil;
  end;
end;

procedure TLspRenameWizard.ExecuteForUnit(const AOldUnitName, ANewUnitName: string);
begin
  if (AOldUnitName = '') or (ANewUnitName = '') or
     SameText(AOldUnitName, ANewUnitName) then Exit;

  // A unit NOBODY references has nothing to rename - the dialog would
  // only be in the way of a plain "Save as...". The probe answers from
  // raw bytes, in parallel, within a time budget; open editor buffers
  // are read live so unsaved references still count. uurUnknown (budget
  // spent, unreadable file) keeps the old behaviour - the dialog opens.
  var HitFile := '';
  var ProbeMs := 0;
  if ProbeUnitUsage(Editor.GetProjectSourceFiles, AOldUnitName,
    DefaultUnitUsageBudgetMs, HitFile, ProbeMs) = uurUnused then
    Exit;

  FUnitRenameMode := True;
  try
    // Synthetic context: only WordAtCursor is actually used in unit mode.
    FContext := Default(TEditorContext);
    FContext.WordAtCursor := AOldUnitName;
    FContext.IsValid := True;

    FDialog := TRenameDialog.CreateDialog(Application.MainForm, AOldUnitName);
    FHost := TDialogRenameHost.Create(FDialog);
    try
      FDialog.OnPreviewRequested := DoPreview;
      FDialog.SetNewName(ANewUnitName);
      TLspManager.Instance.ApplyStatusToCaption(FDialog);
      if FDialog.ShowModal = mrOk then
        ApplyFEdit;
    finally
      FHost := nil;
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

  // BACKUP first ("Create backup" in the dialog, on by default). Until
  // now the checkbox did nothing at all - a safety net that did not exist.
  // The copy is the state the rename starts from: the editor buffer for an
  // open file (unsaved changes included), the disk file otherwise. A
  // failing backup stops the rename before anything is written.
  var BackupDir := '';
  if FHost.CreateBackup then
  begin
    var Files := TList<string>.Create;
    try
      for var FE in FEdit.FileEdits do
        if not Files.Contains(FE.FilePath) then Files.Add(FE.FilePath);
      for var Plan in FFormPlans do
        if not Files.Contains(Plan.FormFile) then Files.Add(Plan.FormFile);
      try
        BackupDir := NewRenameBackupDir;
        for var F in Files do
        begin
          var Target := MirroredBackupPath(BackupDir, F);
          ForceDirectories(ExtractFilePath(Target));
          var Live: string;
          if Editor.ReadEditorContent(F, Live) then
            TDelphiFileEncoding.WriteAll(Target, Live, TDelphiFileEncoding.Detect(F))
          else if TFile.Exists(F) then
            TFile.Copy(F, Target, True);
        end;
      except
        on E: Exception do
        begin
          FHost.Notify('The backup could not be created, nothing was renamed:' +
            sLineBreak + E.Message, True);
          Exit;
        end;
      end;
    finally
      Files.Free;
    end;
  end;

  var AppliedCount := 0;
  var FailedCount := 0;
  var FormNotes := '';
  var AffectedFiles := TList<string>.Create;
  var FormFilesHandled := TList<string>.Create;
  try
    // ---- 1. forms open in the DESIGNER --------------------------------
    // The designer owns such a form; its .dfm on disk is rewritten from
    // memory on the next save. The declaring form is renamed THROUGH the
    // designer (the Object Inspector path), which also renames the
    // declaration in the source - so this runs before the source edits,
    // and ApplyEditViaEditor recognises the already-renamed positions.
    for var Plan in FFormPlans do
    begin
      if not Editor.IsFormInDesigner(Plan.PasFile) then Continue;
      FormFilesHandled.Add(Plan.FormFile);
      if (FFormTarget in [dtField, dtMethod])
        and SameText(ExpandFileName(Plan.PasFile), ExpandFileName(FFormDeclFile)) then
      begin
        var Msg: string;
        if Editor.RenameInFormDesigner(Plan.PasFile, FContext.WordAtCursor,
          FEdit.FileEdits[0].Edits[0].NewText, FFormTarget = dtMethod, Msg) then
          Inc(AppliedCount, Length(Plan.Hits))
        else
        begin
          Inc(FailedCount, Length(Plan.Hits));
          FormNotes := FormNotes + sLineBreak + ExtractFileName(Plan.FormFile) +
            ': the form designer refused - ' + Msg;
        end;
      end
      else
        FormNotes := FormNotes + sLineBreak + ExtractFileName(Plan.FormFile) +
          ': open in the form designer and not the declaring form - not ' +
          'changed by the rename (an inherited form normally follows its ' +
          'ancestor; please check, or close it and rename again)';
    end;

    // ---- 3. (below) form files NOT in a designer: edited as text ------
    // Apply via IDE editor API (undoable!)
    // Per file, sort edits line-descending so earlier edits do not
    // shift later edits' positions.
    for var FE in FEdit.FileEdits do
    begin
      // form files are handled in steps 1 and 3
      if SameText(ExtractFileExt(FE.FilePath), '.dfm')
        or SameText(ExtractFileExt(FE.FilePath), '.fmx') then
        Continue;
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
        if Editor.ApplyEditViaEditor(FE.FilePath, Edit.Range.Start.Line, Edit.Range.Start.Character,
          FContext.WordAtCursor, Edit.NewText) then
          Inc(AppliedCount)
        else
          Inc(FailedCount);
      end;

      if not AffectedFiles.Contains(FE.FilePath) then
        AffectedFiles.Add(FE.FilePath);
    end;

    // ---- 3. form files that are not in a designer ----------------------
    for var Plan in FFormPlans do
    begin
      if FormFilesHandled.Contains(Plan.FormFile) then Continue;
      try
        var Enc := TDelphiFileEncoding.Detect(Plan.FormFile);
        var Old := TDelphiFileEncoding.ReadAll(Plan.FormFile);
        var Skipped: Integer;
        var NewName := '';
        for var FE in FEdit.FileEdits do
          if Length(FE.Edits) > 0 then begin NewName := FE.Edits[0].NewText; Break; end;
        var Changed := ApplyDfmHits(Old, FContext.WordAtCursor, NewName,
          Plan.Hits, Skipped);
        var Buf: string;
        if Editor.ReadEditorContent(Plan.FormFile, Buf) then
          Editor.ReplaceFileContent(Plan.FormFile, Changed)   // standalone tab
        else
          TDelphiFileEncoding.WriteAll(Plan.FormFile, Changed, Enc);
        Inc(AppliedCount, Length(Plan.Hits) - Skipped);
        Inc(FailedCount, Skipped);
      except
        on E: Exception do
        begin
          Inc(FailedCount, Length(Plan.Hits));
          FormNotes := FormNotes + sLineBreak + ExtractFileName(Plan.FormFile) +
            ': ' + E.Message;
        end;
      end;
    end;

    // Inform LSP about the changes (not needed in unit-rename mode
    // because LSP was not used for verification, but harmless).
    if TLspManager.Instance.IsAlive then
    begin
      try
        var Client := TLspManager.Instance.GetClient(FContext.ProjectRoot, FContext.ProjectFile, Editor.FindDelphiLspJson);
        for var F in AffectedFiles do
          Client.SyncDocument(F);   // sends only what really changed
      except
        // LSP refresh is best-effort
      end;
    end;

    var BackupNote := '';
    if BackupDir <> '' then
      BackupNote := sLineBreak + 'Backup of the previous state: ' + BackupDir;
    if (FailedCount = 0) and (FormNotes = '') then
      FHost.Notify(Format('%d change(s) applied successfully (Ctrl+Z to undo ' +
        'in files open in the editor).%s', [AppliedCount, BackupNote]), False)
    else
      FHost.Notify(Format('%d applied, %d failed.%s%s',
        [AppliedCount, FailedCount, FormNotes, BackupNote]), True);
  finally
    FormFilesHandled.Free;
    AffectedFiles.Free;
  end;
end;

procedure TLspRenameWizard.CollectFormEdits(const AFiles: TArray<string>;
  const ADefFile: string; ADefLine: Integer; const AOwnerType, AOldName,
  ANewName: string);
var
  DeclLines: TArray<string>;
  Decl, Low: string;
  ImplLine, I, DeclIdx: Integer;
  Parents: TDictionary<string, string>;
  QualRoots: TDictionary<string, Boolean>;
  Texts: TDictionary<string, string>;
  Pairs: TList<TPair<string, string>>;   // (pas, form file)
begin
  FFormPlans := nil;
  FFormDeclFile := ADefFile;
  try
    DeclLines := ReadDelphiFileLines(ADefFile);
  except
    Exit;
  end;
  if (ADefLine < 0) or (ADefLine > High(DeclLines)) then Exit;
  DeclIdx := ADefLine;
  Decl := DeclLines[DeclIdx];

  // What kind of symbol is it? Members only count when declared INSIDE a
  // class body in the interface part: a local variable "Button1" in
  // TForm1.FormCreate has TForm1 as containing type too, and must never
  // rename the form's real Button1.
  ImplLine := MaxInt;
  for I := 0 to High(DeclLines) do
    if SameText(Trim(DeclLines[I]), 'implementation') then
    begin
      ImplLine := I;
      Break;
    end;

  // DelphiLSP frequently answers GotoDefinition for a METHOD with its
  // IMPLEMENTATION header ("procedure TDlgAbout.sbDbFilesPathClick") -
  // tester: renaming an event handler showed both source lines but no
  // form entry, because that line lies below 'implementation' and was
  // taken for a local. An implementation header of the owner type is not
  // a local: go to the member's declaration in the class body instead.
  if (AOwnerType <> '') and (DeclIdx >= ImplLine)
    and SameText(TImplementationFinder.OwnerTypeFromImplLine(Decl), AOwnerType) then
  begin
    var InClass := FindMemberDeclarationLine(string.Join(sLineBreak, DeclLines),
      AOwnerType, AOldName);
    if (InClass >= 0) and (InClass < ImplLine) then
    begin
      DeclIdx := InClass;
      Decl := DeclLines[DeclIdx];
      FDiagLog := FDiagLog + Format('Form check: definition was the ' +
        'implementation header, using the class declaration at line %d' +
        sLineBreak, [DeclIdx + 1]);
    end;
  end;
  Low := LowerCase(TrimLeft(Decl));

  if (AOwnerType <> '') and (DeclIdx < ImplLine) then
  begin
    if StartsStr('property ', Low) or StartsStr('class property ', Low) then
      FFormTarget := dtProperty
    else if StartsStr('procedure ', Low) or StartsStr('function ', Low)
      or StartsStr('class procedure ', Low) or StartsStr('class function ', Low) then
      FFormTarget := dtMethod
    else if Pos(':', Decl) > 0 then
      FFormTarget := dtField
    else
      Exit;
  end
  else if (AOwnerType = '') and TRegEx.IsMatch(Decl,
    '^\s*' + TRegEx.Escape(AOldName) + '\s*(<[^>]*>)?\s*=\s*(packed\s+)?class\b',
    [roIgnoreCase]) then
    FFormTarget := dtType
  else
    Exit;

  Parents := TDictionary<string, string>.Create;
  QualRoots := TDictionary<string, Boolean>.Create;
  Texts := TDictionary<string, string>.Create;
  Pairs := TList<TPair<string, string>>.Create;
  try
    // Form files next to the scanned units (plus the declaring unit).
    var Seen := TDictionary<string, Boolean>.Create;
    try
      for var F in AFiles + [ADefFile] do
      begin
        if not SameText(ExtractFileExt(F), '.pas') then Continue;
        if Seen.ContainsKey(UpperCase(F)) then Continue;
        Seen.Add(UpperCase(F), True);
        var Form := FormFileOf(F);
        if Form = '' then Continue;
        if not IsTextFormFile(Form) then
        begin
          FDiagLog := FDiagLog + 'Form file skipped (binary format): ' + Form + sLineBreak;
          Continue;
        end;
        try
          Texts.Add(Form, TDelphiFileEncoding.ReadAll(Form));
          Pairs.Add(TPair<string, string>.Create(F, Form));
        except
        end;
      end;
    finally
      Seen.Free;
    end;
    if Pairs.Count = 0 then Exit;

    // Class hierarchy from the scanned sources: an inherited form or a
    // frame class only matches through its parent chain.
    for var F in AFiles + [ADefFile] do
      if SameText(ExtractFileExt(F), '.pas') then
      try
        CollectClassParents(ReadDelphiFileLines(F), Parents);
      except
      end;

    var Owner := AOwnerType;
    var Match: TDfmClassMatch :=
      function(const AClass: string): Boolean
      begin
        Result := ClassMatchesType(Parents, AClass, Owner);
      end;

    // Roots whose class matches: "DataModule1.Button1" qualifiers.
    for var Pr in Pairs do
    begin
      var Root := ReadDfmRoot(Texts[Pr.Value]);
      if (Root.Name <> '') and Match(Root.ClassName) then
        QualRoots.AddOrSetValue(UpperCase(Root.Name), True);
    end;
    var Qual: TDfmQualifierMatch :=
      function(const AQualifier: string): Boolean
      begin
        Result := QualRoots.ContainsKey(UpperCase(AQualifier));
      end;

    for var Pr in Pairs do
    begin
      var Hits := FindDfmRenameHits(Texts[Pr.Value], AOldName, FFormTarget,
        Match, Qual);
      if Length(Hits) = 0 then Continue;

      var Plan: TRenameFormPlan;
      Plan.FormFile := Pr.Value;
      Plan.PasFile := Pr.Key;
      Plan.Hits := Hits;
      FFormPlans := FFormPlans + [Plan];

      // Same shape as the LSP edits, so preview, counting and the details
      // log treat form occurrences like any other.
      var FE: TLspFileEdits;
      FE.FilePath := Pr.Value;
      SetLength(FE.Edits, Length(Hits));
      for I := 0 to High(Hits) do
      begin
        FE.Edits[I].Range.Start.Line := Hits[I].Line;
        FE.Edits[I].Range.Start.Character := Hits[I].Col;
        FE.Edits[I].Range.End_.Line := Hits[I].Line;
        FE.Edits[I].Range.End_.Character := Hits[I].Col + Hits[I].Len;
        FE.Edits[I].NewText := ANewName;
      end;
      FEdit.FileEdits := FEdit.FileEdits + [FE];
      FDiagLog := FDiagLog + Format('Form file %s: %d occurrence(s)%s',
        [ExtractFileName(Pr.Value), Length(Hits),
         IfThen(Editor.IsFormInDesigner(Pr.Key), ' (open in the form designer)', '')]) +
        sLineBreak;
    end;
  finally
    Pairs.Free;
    Texts.Free;
    QualRoots.Free;
    Parents.Free;
  end;
end;

function TLspRenameWizard.FormHitKindAt(const AFile: string; ALine, ACol: Integer;
  out AKind: string): Boolean;
begin
  Result := False;
  for var Plan in FFormPlans do
    if SameText(Plan.FormFile, AFile) then
      for var H in Plan.Hits do
        if (H.Line = ALine) and (H.Col = ACol) then
        begin
          AKind := DfmHitKindText(H.Kind);
          if Editor.IsFormInDesigner(Plan.PasFile) then
            AKind := AKind + ' (via designer)';
          Exit(True);
        end;
end;

function TLspRenameWizard.PreviewHeadless(const AContext: TEditorContext;
  const AHost: IRenameHost): Boolean;
begin
  FUnitRenameMode := False;
  FContext := AContext;
  FEdit := Default(TLspWorkspaceEdit);
  FHost := AHost;
  DoPreview(nil);
  Result := Length(FEdit.FileEdits) > 0;
end;

procedure TLspRenameWizard.ApplyHeadless(const AHost: IRenameHost);
begin
  FHost := AHost;
  ApplyFEdit;
end;

{ TDialogRenameHost }

constructor TDialogRenameHost.Create(ADialog: TRenameDialog);
begin
  inherited Create;
  FDialog := ADialog;
end;

function TDialogRenameHost.GetNewName: string;
begin
  Result := FDialog.GetNewName;
end;

function TDialogRenameHost.CreateBackup: Boolean;
begin
  Result := FDialog.GetCreateBackup;
end;

function TDialogRenameHost.Scope: TRenameScope;
begin
  Result := FDialog.Scope;
end;

function TDialogRenameHost.SelectedUnits: TArray<string>;
begin
  Result := FDialog.SelectedUnits;
end;

function TDialogRenameHost.IncludeOpenUnits: Boolean;
begin
  Result := FDialog.IncludeOpenUnits;
end;

function TDialogRenameHost.IncludeUsedUnits: Boolean;
begin
  Result := FDialog.IncludeUsedUnits;
end;

function TDialogRenameHost.ScanCancelled: Boolean;
begin
  Result := FDialog.ScanCancelled;
end;

procedure TDialogRenameHost.SetBusy(ABusy: Boolean);
begin
  FDialog.SetBusy(ABusy);
end;

procedure TDialogRenameHost.SetStatus(const AText: string);
begin
  FDialog.SetStatus(AText);
end;

procedure TDialogRenameHost.SetProgress(AValue, AMax: Integer);
begin
  FDialog.SetProgress(AValue, AMax);
end;

procedure TDialogRenameHost.SetPreviewItems(const AItems: TRenamePreviewItems);
begin
  FDialog.SetPreviewItems(AItems);
end;

procedure TDialogRenameHost.SetDetailsText(const AText: string);
begin
  FDialog.SetDetailsText(AText);
end;

procedure TDialogRenameHost.EnableRename(AEnabled: Boolean);
begin
  FDialog.EnableRename(AEnabled);
end;

procedure TDialogRenameHost.Notify(const AText: string; AWarning: Boolean);
begin
  if AWarning then
    MessageDlg(AText, mtWarning, [mbOK], 0)
  else
    MessageDlg(AText, mtInformation, [mbOK], 0);
end;

procedure TLspRenameWizard.DoPreview(Sender: TObject);
begin
  // a previous preview's form plans must never be applied to this one
  FFormPlans := nil;
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
  NewName := FHost.GetNewName;
  if NewName = '' then
  begin
    FHost.Notify('Please enter a new name.', True);
    Exit;
  end;
  // Case-sensitive compare: a case-only change (foo -> Foo) is a
  // legitimate rename in Pascal and must go through the full pipeline.
  if NewName = FContext.WordAtCursor then
  begin
    FHost.Notify('The new name is identical to the old one.', True);
    Exit;
  end;

  // Save all dirty files so the text scan sees their current state
  Editor.SaveAllFiles;

  FHost.SetBusy(True);
  FDiagLog := '';
  try
    // Units outside the project that use the renamed unit need the new
    // name as well (settings: open units / units via uses).
    ProjFiles := ProjectScopeFiles(FContext.FileName);
    FDiagLog :=
      '=== Diagnostics (Unit Rename) ===' + sLineBreak +
      'Old unit name: ' + FContext.WordAtCursor + sLineBreak +
      'New unit name: ' + NewName + sLineBreak +
      'Project files: ' + IntToStr(Length(ProjFiles)) + sLineBreak +
      sLineBreak +
      'Text search (whole-word, skipping strings and comments)...' + sLineBreak;

    FHost.SetStatus(Format('Scanning %d project file(s)...', [Length(ProjFiles)]));
    Candidates := FindCandidates(FContext.WordAtCursor, ProjFiles);
    FDiagLog := FDiagLog + 'Text candidates: ' + IntToStr(Length(Candidates)) + sLineBreak;

    if Length(Candidates) = 0 then
    begin
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText(Format('No references to unit "%s" found in the project.',
        [FContext.WordAtCursor]) + sLineBreak + sLineBreak + FDiagLog);
      FHost.SetStatus('Done - no matches.');
      FHost.SetBusy(False);
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

    // A CANCELLED scan has only partial candidates - never present that
    // as a finished preview (applying it would rename some occurrences
    // and silently leave the rest behind).
    if FHost.ScanCancelled then
    begin
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText('Preview cancelled by the user.' + sLineBreak +
        sLineBreak + FDiagLog);
      FHost.SetStatus('Preview cancelled - nothing was changed.');
      FHost.EnableRename(False);
      FHost.SetBusy(False);
      Exit;
    end;

    TotalEdits := 0;
    for var FE in FEdit.FileEdits do
      Inc(TotalEdits, Length(FE.Edits));

    FHost.SetPreviewItems(PreviewItems);

    FHost.SetDetailsText(FDiagLog);
    FHost.EnableRename(True);
    FHost.SetStatus(Format('Done: %d change(s) in %d file(s).',
      [TotalEdits, Length(FEdit.FileEdits)]));
  except
    on E: Exception do
    begin
      FDiagLog := FDiagLog + sLineBreak + 'EXCEPTION: ' + E.ClassName + ': ' + E.Message;
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText(FDiagLog);
      FHost.SetStatus('An error occurred.');
    end;
  end;
  FHost.SetBusy(False);
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
  ScopeFirst, ScopeLast: Integer;   // "current method" line window (0-based)
  IncCtx: TLspIncludeContext;       // freed = expanded units restored
begin
  IncCtx := nil;
  NewName := FHost.GetNewName;
  if NewName = '' then
  begin
    FHost.Notify('Please enter a new name.', True);
    Exit;
  end;
  // Case-sensitive compare: a case-only change (foo -> Foo) is a
  // legitimate rename in Pascal and must go through the full pipeline.
  if NewName = FContext.WordAtCursor then
  begin
    FHost.Notify('The new name is identical to the old one.', True);
    Exit;
  end;
  // A reserved word ('string', 'begin' ...) is no symbol - and every
  // occurrence in the project would become a candidate to verify, one
  // DelphiLSP round trip each, on the main thread (a misplaced caret on
  // 'string' kept the IDE busy for minutes).
  if TIdentifierChecker.IsPascalKeyword(FContext.WordAtCursor) then
  begin
    FHost.Notify(Format('"%s" is a reserved word - place the caret on the ' +
      'identifier to rename.', [FContext.WordAtCursor]), True);
    Exit;
  end;

  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    FHost.Notify('No .delphilsp.json found.' + sLineBreak +
      'Please enable: Tools > Options > Editor > Language > ' +
      'Code Insight > "Generate LSP Config".', True);
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  FHost.SetBusy(True);
  FDiagLog := '';
  try
    // Save all unsaved files (so LSP sees the current state).
    FHost.SetStatus('Saving all files...');
    Editor.SaveAllFiles;

    FDiagLog := '=== Diagnostics ===' + sLineBreak +
      'File: ' + FContext.FileName + sLineBreak +
      'Identifier: ' + FContext.WordAtCursor + sLineBreak +
      'Position: ' + IntToStr(FContext.Line) + ':' + IntToStr(FContext.Column) + sLineBreak +
      'delphilsp.json: ' + DelphiLspJson + sLineBreak + sLineBreak;

    // ---- SCOPE ---------------------------------------------------------
    // The dialog decides how far the rename reaches. Limiting the FILE
    // SET here also makes the scan itself much faster, because every
    // later phase (text scan, implementation scan, LSP verification)
    // works on it.
    ScopeFirst := -1;
    ScopeLast := -1;
    case FHost.Scope of
      rscCurrentUnit:
        begin
          ProjFiles := [FContext.FileName];
          FHost.SetStatus('Scope: current unit.');
        end;
      rscCurrentMethod:
        begin
          ProjFiles := [FContext.FileName];
          var MethContent: string;
          if not Editor.ReadEditorContent(FContext.FileName, MethContent) then
            MethContent := '';
          if not FindEnclosingRoutineRange(MethContent, FContext.Line - 1,
            ScopeFirst, ScopeLast) then
          begin
            FHost.Notify('The caret is not inside a method body - ' +
              'choose another scope.', True);
            FHost.SetBusy(False);
            Exit;
          end;
          FHost.SetStatus(Format('Scope: current method (lines %d-%d).',
            [ScopeFirst + 1, ScopeLast + 1]));
        end;
      rscSelectedUnits:
        begin
          ProjFiles := FHost.SelectedUnits;
          if Length(ProjFiles) = 0 then
          begin
            FHost.Notify('No units selected - press "Select..." first.', True);
            FHost.SetBusy(False);
            Exit;
          end;
          // The declaration lives where the caret is - keep that file in
          // scope, otherwise the rename could never touch it.
          var HaveCur := False;
          for var PF in ProjFiles do
            if SameText(PF, FContext.FileName) then HaveCur := True;
          if not HaveCur then ProjFiles := ProjFiles + [FContext.FileName];
          FHost.SetStatus(Format('Scope: %d selected unit(s).',
            [Length(ProjFiles)]));
        end;
    else
      // The project's sources, the caret's own unit ALWAYS (forum: a
      // rename started in a unit outside the project changed every
      // project unit but not the one it was started in), plus the
      // extras ticked in the dialog.
      begin
        var Extra: TScopeExtra;
        ProjFiles := ProjectScopeFiles(FContext.FileName,
          FHost.IncludeOpenUnits, FHost.IncludeUsedUnits, Extra);
        FHost.SetStatus('Scope: ' + ScopeExtraText(Length(ProjFiles), Extra) + '.');
        FDiagLog := FDiagLog + 'Scope: ' +
          ScopeExtraText(Length(ProjFiles), Extra) + sLineBreak;
      end;
    end;
    FDiagLog := FDiagLog + 'Project files: ' + IntToStr(Length(ProjFiles)) + sLineBreak + sLineBreak;

    // Phase 1: text search over project files
    FHost.SetStatus('Phase 1: text search...');
    Candidates := FindCandidates(FContext.WordAtCursor, ProjFiles);
    // "In current method": drop everything outside the routine's lines.
    if ScopeFirst >= 0 then
    begin
      var InScope: TArray<TRenameCandidate> := nil;
      for var C in Candidates do
        if (C.Line >= ScopeFirst) and (C.Line <= ScopeLast) then
          InScope := InScope + [C];
      FDiagLog := FDiagLog + Format(
        'Scope filter (method lines %d-%d): %d of %d candidate(s) kept' +
        sLineBreak, [ScopeFirst + 1, ScopeLast + 1, Length(InScope),
        Length(Candidates)]);
      Candidates := InScope;
    end;
    FDiagLog := FDiagLog + 'Text candidates: ' + IntToStr(Length(Candidates)) + sLineBreak + sLineBreak;

    if Length(Candidates) = 0 then
    begin
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText('No occurrences found.' + sLineBreak + sLineBreak + FDiagLog);
      FHost.SetStatus('Done - no matches.');
      FHost.SetBusy(False);
      Exit;
    end;

    // Phase 2: start LSP (singleton - first call slow, later calls instant)
    var WasRunning := TLspManager.Instance.IsAlive;
    if WasRunning then
      FHost.SetStatus('LSP already running. Opening file...')
    else
      FHost.SetStatus('Starting LSP server (one-time)...');

    Client := TLspManager.Instance.GetClient(
      RootPath, FContext.ProjectFile, DelphiLspJson);

    // Positions inside {$I} include files are answered through the
    // INCLUDING unit, sent expanded (Expert.IncludeExpansion); freeing the
    // context sends the original text again.
    FUnverified := nil;
    FUnverifiedWhy := nil;
    IncCtx := TLspIncludeContext.Create(Client, EditorOrDiskReader());
    IncCtx.RegisterFiles(ProjFiles);

    // Hand the current content to DelphiLSP when it changed since the last
    // send, and wait for the unit's analysis then (see VerifyWithLsp for
    // why a blind re-open + fixed sleep loses answers). An include file is
    // no unit - the include context serves it.
    if not IncCtx.OwnsDocument(FContext.FileName) then
    begin
      var StartBefore := Client.GetFileDiagnosticsVersion(FContext.FileName);
      if Client.SyncDocument(FContext.FileName) and WasRunning then
        Client.WaitFileAnalysed(FContext.FileName, StartBefore, 30000,
          function: Boolean
          begin
            FHost.SetStatus('Waiting for DelphiLSP to analyse ' +
              ExtractFileName(FContext.FileName) + '...');
            Result := not FHost.ScanCancelled;
          end);
    end;

    // On first start, wait for readiness
    if not WasRunning then
    begin
      var LspLine := FContext.Line - 1;
      var LspCol := FContext.Column - 1;      for var Retry := 1 to 30 do
      begin
        FHost.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
        Application.ProcessMessages;
        try
          var H := Client.GetHover(FContext.FileName, LspLine, LspCol);
          if H <> '' then Break;
          var D := IncCtx.Definition(FContext.FileName, LspLine, LspCol);
          if Length(D) > 0 then Break;
        except end;
        Sleep(1000);
      end;
    end;

    // Phase 2b: find declaration
    FHost.SetStatus('Finding declaration...');

    var LspLine := FContext.Line - 1;
    var LspCol := FContext.Column - 1;
    var DefLocs := IncCtx.Definition(FContext.FileName, LspLine, LspCol);
    var DefLine := 0;
    var DefCol := 0;

    if Length(DefLocs) > 0 then
    begin
      DefFilePath := TLspUri.FileUriToPath(DefLocs[0].Uri);
      DefLine := DefLocs[0].Range.Start.Line;
      DefCol := DefLocs[0].Range.Start.Character;
    end
    else
    begin
      DefFilePath := FContext.FileName;
      // DelphiLSP answers GotoDefinition AT a declaration with null. When
      // the caret sits on one ('procedure X', 'function X', 'property X'
      // ...), the caret IS the declaration - without this the wizard took
      // "<file>:1:1", found no owner type and skipped the implementing
      // classes in other units (renaming IRenameHost.SetStatus left
      // THeadlessRenameHost.SetStatus behind - did not compile).
      if CaretOnDeclaration(FContext.FileName, LspLine, LspCol) then
      begin
        DefLine := LspLine;
        DefCol := LspCol;
        FDiagLog := FDiagLog + 'LSP gave no definition - the caret is on a ' +
          'declaration, using it.' + sLineBreak;
      end;
    end;

    FDiagLog := FDiagLog + 'Declaration: ' + DefFilePath + ':' + IntToStr(DefLine + 1) + ':' + IntToStr(DefCol + 1) + sLineBreak;

    // Declared in the RAD Studio installation (RTL/VCL): not ours to rename.
    // Refuse BEFORE verifying - such names ('Integer', 'Create', 'Free')
    // occur thousands of times, each one a DelphiLSP round trip.
    var BdsRootDir := FindBdsRoot;
    if (BdsRootDir <> '') and (DefFilePath <> '') and
       UpperCase(ExpandFileName(DefFilePath)).StartsWith(
         UpperCase(IncludeTrailingPathDelimiter(ExpandFileName(BdsRootDir)))) then
    begin
      FDiagLog := FDiagLog + 'Refused: the declaration is part of the RAD Studio ' +
        'installation.' + sLineBreak;
      FreeAndNil(IncCtx);
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText(FDiagLog);
      FHost.SetStatus(Format('"%s" is declared in %s (RAD Studio installation) - ' +
        'it cannot be renamed.', [FContext.WordAtCursor, ExtractFileName(DefFilePath)]));
      FHost.EnableRename(False);
      FHost.SetBusy(False);
      Exit;
    end;

    // The DECLARATION may live in a unit none of the scanned files is -
    // renaming every use but not the declaration does not compile. For
    // the whole-project scope, scan that file too (never inside the
    // RAD Studio installation: the RTL/VCL is not ours to rename).
    if (FHost.Scope = rscProject) and (DefFilePath <> '')
      and TFile.Exists(DefFilePath) then
    begin
      var HaveDecl := False;
      for var PF in ProjFiles do
        if SameText(ExpandFileName(PF), ExpandFileName(DefFilePath)) then
          HaveDecl := True;
      var Bds := FindBdsRoot;
      if (Bds <> '') and UpperCase(ExpandFileName(DefFilePath)).StartsWith(
        UpperCase(IncludeTrailingPathDelimiter(ExpandFileName(Bds)))) then
        HaveDecl := True;   // never rename inside the installation
      if not HaveDecl then
      begin
        ProjFiles := ProjFiles + [DefFilePath];
        Candidates := Candidates + FindCandidates(FContext.WordAtCursor, [DefFilePath]);
        FDiagLog := FDiagLog + 'Declaration unit added to the scan: ' +
          DefFilePath + sLineBreak;
      end;
    end;

    // Phase 2c: find interface/class method implementations.
    // Text-based scan over all project files with syntax filter on lines
    // like 'procedure TClass.Method'. Only classes that implement the
    // container (owner) type of the method are kept.
    FHost.SetStatus('Searching for interface implementations...');

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
      FHost.SetStatus(Format('%d candidate(s). Verifying...', [Length(Candidates)]));

      ImplFilesArray := ImplFilesList.ToArray;
      // The symbol's own positions: its declaration + implementation, and
      // those of every implementing class (interface / virtual methods).
      var Targets: TLspSymbolTargets;
      if Length(DefLocs) > 0 then
        IncCtx.AddTargetWithPartner(Targets, DefFilePath, DefLine, DefCol)
      else
        IncCtx.AddTargetWithPartner(Targets, FContext.FileName, LspLine, LspCol);
      // The implementation scan is for OTHER types (interface implementers,
      // overrides in descendants). A header of the OWNER type itself that is
      // not the symbol already is a sibling OVERLOAD - forum report: renaming
      // Init(const xBoolean: Boolean) also renamed the parameterless Init.
      var ImplLinesOf := TDictionary<string, TArray<string>>.Create;
      try
        for var IC in ImplCandidates do
        begin
          var ICLines: TArray<string>;
          if not ImplLinesOf.TryGetValue(UpperCase(IC.FilePath), ICLines) then
          begin
            try
              ICLines := ReadDelphiFileLines(IC.FilePath);
            except
              ICLines := nil;
            end;
            ImplLinesOf.Add(UpperCase(IC.FilePath), ICLines);
          end;
          if (OwnerType <> '') and (IC.Line >= 0) and (IC.Line <= High(ICLines))
            and SameText(TImplementationFinder.OwnerTypeFromImplLine(ICLines[IC.Line]), OwnerType)
            and not Targets.Contains(IC.FilePath, IC.Line) then
          begin
            FDiagLog := FDiagLog + 'Sibling overload skipped: ' +
              ExtractFileName(IC.FilePath) + ':' + IntToStr(IC.Line + 1) + sLineBreak;
            Continue;
          end;
          IncCtx.AddTargetWithPartner(Targets, IC.FilePath, IC.Line, IC.Col);
        end;
      finally
        ImplLinesOf.Free;
      end;
      FDiagLog := FDiagLog + 'Symbol positions (' + IntToStr(Targets.Count) + '):' +
        sLineBreak + Targets.Text + sLineBreak;
      FEdit := VerifyWithLsp(Candidates, FContext.WordAtCursor, NewName, IncCtx, Targets, Client);
    finally
      ImplFilesList.Free;
    end;
    if IncCtx.Activations > 0 then
      FDiagLog := FDiagLog + 'Include files:' + sLineBreak + IncCtx.NotesText;
    FreeAndNil(IncCtx);   // sends the original text of expanded units again

    // Form files: component names, event handlers, component references.
    // Only when the LSP actually told us where the declaration is, and not
    // for the "current method" scope (locals never appear in a form).
    FFormPlans := nil;
    if (Length(DefLocs) > 0) and (FHost.Scope <> rscCurrentMethod)
      and not FHost.ScanCancelled then
    begin
      FHost.SetStatus('Checking form files...');
      CollectFormEdits(ProjFiles, DefFilePath, DefLine, OwnerType,
        FContext.WordAtCursor, NewName);
    end;

    if (Length(FEdit.FileEdits) = 0) and (Length(FUnverified) = 0) then
    begin
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText(FDiagLog);
      FHost.SetStatus('Done - no verified matches.');
      FHost.SetBusy(False);
      Exit;
    end;

    // Build structured preview for the ListView
    var PreviewItems := BuildPreviewItems(FEdit, DefFilePath, DefLine, ImplFilesArray);
    for var U in FUnverified do
    begin
      var PI := Default(TRenamePreviewItem);
      PI.FilePath := U.FilePath;
      PI.Line := U.Line;
      PI.Col := U.Col;
      PI.Kind := 'UNVERIFIED - not renamed';
      for var K := 0 to High(FUnverified) do
        if (FUnverified[K].FilePath = U.FilePath) and (FUnverified[K].Line = U.Line) and
           (FUnverified[K].Col = U.Col) and (K <= High(FUnverifiedWhy)) then
          PI.Kind := 'UNVERIFIED - not renamed (' + FUnverifiedWhy[K] + ')';
      try
        var UL := ReadDelphiFileLines(U.FilePath);
        if (U.Line >= 0) and (U.Line <= High(UL)) then PI.OriginalLine := UL[U.Line];
      except
        PI.OriginalLine := '';
      end;
      PI.PreviewLine := PI.OriginalLine;
      PreviewItems := PreviewItems + [PI];
    end;

    // Count for the status line
    // A CANCELLED scan has only partial candidates - never present that
    // as a finished preview (applying it would rename some occurrences
    // and silently leave the rest behind).
    if FHost.ScanCancelled then
    begin
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText('Preview cancelled by the user.' + sLineBreak +
        sLineBreak + FDiagLog);
      FHost.SetStatus('Preview cancelled - nothing was changed.');
      FHost.EnableRename(False);
      FHost.SetBusy(False);
      Exit;
    end;

    var TotalEdits := 0;
    for var FE in FEdit.FileEdits do
      Inc(TotalEdits, Length(FE.Edits));

    FHost.SetPreviewItems(PreviewItems);
    // A LIMITED scope can leave the declaration out - the result would
    // not compile. The preview must say that plainly; it is still a
    // legitimate choice (renaming a local variable, or a staged rename).
    var ScopeWarn := '';
    if (FHost.Scope <> rscProject) and (DefFilePath <> '') then
    begin
      var DeclCovered := False;
      for var FE in FEdit.FileEdits do
        if SameText(FE.FilePath, DefFilePath) then
          for var Ed in FE.Edits do
            if (ScopeFirst < 0)
              or ((Ed.Range.Start.Line >= ScopeFirst)
                  and (Ed.Range.Start.Line <= ScopeLast)) then
              DeclCovered := True;
      if not DeclCovered then
        ScopeWarn := '  WARNING: the declaration is OUTSIDE the selected ' +
          'scope and stays unchanged.';
    end;
    // CONFLICT CHECK (issue #11): the new name may already exist where the
    // renamed symbol lives or is used. Such a rename can compile and still
    // mean something different - a local hiding a field, a method hiding
    // an ancestor member, a global shadowed by a used unit. Reported, not
    // refused: the user decides.
    var Conflict := NameConflictNote(NewName, OwnerType, DefFilePath);
    if Conflict <> '' then
      ScopeWarn := ScopeWarn + '  NOTE: "' + NewName + '" already exists - ' +
        'see the details tab.';
    FHost.SetDetailsText(TrimLeft(ScopeWarn) + sLineBreak + Conflict + sLineBreak +
      sLineBreak + FDiagLog);
    if Length(FUnverified) > 0 then
      ScopeWarn := ScopeWarn + Format('  WARNING: %d occurrence(s) could not be ' +
        'verified by DelphiLSP and are NOT renamed - check them (UNVERIFIED rows).',
        [Length(FUnverified)]);
    FHost.EnableRename(Length(FEdit.FileEdits) > 0);
    FHost.SetStatus(Format('Done: %d change(s) in %d file(s).%s',
      [TotalEdits, Length(FEdit.FileEdits), ScopeWarn]));
  except
    on E: Exception do
    begin
      FDiagLog := FDiagLog + sLineBreak + 'EXCEPTION: ' + E.ClassName + ': ' + E.Message;
      FHost.SetPreviewItems(nil);
      FHost.SetDetailsText(FDiagLog);
      FHost.SetStatus('An error occurred.');
    end;
  end;
  FreeAndNil(IncCtx);   // after an exception: restore the expanded units
  FHost.SetBusy(False);
end;

{ Helper functions }

function TLspRenameWizard.NameConflictNote(const ANewName, AOwnerType,
  ADefFile: string): string;
const
  MaxExamples = 5;
var
  Lines: TArray<string>;
  Content, DefContent: string;
  Examples: TStringList;
  Total: Integer;
begin
  Result := '';
  Examples := TStringList.Create;
  try
    Total := 0;
    // 1. the new name already occurs in the files this rename touches
    for var FE in FEdit.FileEdits do
    begin
      if not Editor.ReadEditorContent(FE.FilePath, Content) then
        try
          Content := ReadDelphiFile(FE.FilePath);
        except
          Continue;
        end;
      if SameText(FE.FilePath, ADefFile) then DefContent := Content;
      Lines := Content.Split([#13#10, #10]);
      for var L in CodeWordLines(Lines, ANewName) do
      begin
        Inc(Total);
        if Examples.Count < MaxExamples then
          Examples.Add(Format('  %s:%d  %s', [ExtractFileName(FE.FilePath), L + 1,
            Trim(Lines[L])]));
      end;
    end;
    if Total > 0 then
      Result := Result + Format('"%s" already occurs %d time(s) in the files ' +
        'this rename touches:', [ANewName, Total]) + sLineBreak +
        Examples.Text;
    if (DefContent = '') and (ADefFile <> '') then
      try
        if not Editor.ReadEditorContent(ADefFile, DefContent) then
          DefContent := ReadDelphiFile(ADefFile);
      except
        DefContent := '';
      end;
    // 2. a member of that name in the owner type (hides / is hidden)
    if (AOwnerType <> '') and (DefContent <> '') then
    begin
      var ML := FindMemberDeclarationLine(DefContent, AOwnerType, ANewName);
      if ML >= 0 then
        Result := Result + Format('%s already declares a member "%s" (%s:%d).',
          [AOwnerType, ANewName, ExtractFileName(ADefFile), ML + 1]) + sLineBreak;
    end;
    // 3. a unit in the declaring file's uses clause declares it globally
    if DefContent <> '' then
    begin
      var Units := '';
      for var Hit in TUnitIndex.Instance.Lookup(ANewName) do
        if UnitInUsesText(DefContent, Hit.UnitName) and (Pos(Hit.UnitName, Units) = 0) then
          Units := Units + IfThen(Units <> '', ', ', '') + Hit.UnitName;
      if Units <> '' then
        Result := Result + Format('"%s" is also declared in unit(s) used here: %s.',
          [ANewName, Units]) + sLineBreak;
    end;
    if Result <> '' then
      Result := 'NAME CONFLICT CHECK - the rename can compile and still change ' +
        'what a name refers to:' + sLineBreak + Result;
  finally
    Examples.Free;
  end;
end;

{ Text search }

function TLspRenameWizard.FindCandidates(const AOldName: string; const AFiles: TArray<string>): TArray<TRenameCandidate>;
var
  CandidateList: TList<TRenameCandidate>;
  F, Line, RawContent: string;
  Lines, Masked: TArray<string>;
  UpperOldName: string;
  LineIdx, SearchPos, FoundPos, AfterPos: Integer;
  BeforeOk, AfterOk: Boolean;
  Candidate: TRenameCandidate;
begin
  UpperOldName := UpperCase(AOldName);
  CandidateList := TList<TRenameCandidate>.Create;
  try
    FHost.SetProgress(0, Length(AFiles));

    for var FileIdx := 0 to High(AFiles) do
    begin
      F := AFiles[FileIdx];
      if (FileIdx mod 10 = 0) then
        FHost.SetProgress(FileIdx + 1, Length(AFiles));
      if FHost.ScanCancelled then Break;

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

          if BeforeOk and AfterOk and (Masked[LineIdx][FoundPos] = Line[FoundPos]) then
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

    FHost.SetProgress(Length(AFiles), Length(AFiles));
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
  // Progress callback matters beyond the bar: it is what pumps the
  // message queue during this full-project scan, so Stop stays clickable.
  Items := TImplementationFinder.FindByProjectScan(AProjectFiles, AOldName, AOwnerType,
    procedure(ACurrent, ATotal: Integer)
    begin
      FHost.SetProgress(ACurrent, ATotal);
      if (ACurrent mod 10 = 0) or (ACurrent = ATotal) then
        FHost.SetStatus(Format('Phase 2c: scanning implementations (%d/%d)...',
          [ACurrent, ATotal]));
    end);

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

  // 0-based line of the 'implementation' keyword, -1 when there is none
  // (program files, include files).
  function ImplementationLineOf(const ALines: TArray<string>): Integer;
  begin
    for var I := 0 to High(ALines) do
      if SameText(Trim(ALines[I]), 'implementation') then
        Exit(I);
    Result := -1;
  end;

  // Kind of the type whose body contains ALine: 'interface', 'class'
  // (also record / object) or '' for a free routine. Walks up to the
  // nearest type opener; a bare 'end;' or a section keyword on the way
  // means the line is outside any type body.
  function EnclosingTypeKind(const ALines: TArray<string>; ALine: Integer): string;
  begin
    Result := '';
    for var I := ALine - 1 downto 0 do
    begin
      var S := LowerCase(Trim(ALines[I]));
      if (S = 'end;') or (S = 'interface') or (S = 'implementation') or
         (S = 'type') then
        Exit;
      var P := Pos('=', S);
      if P < 2 then Continue;
      var Rhs := Trim(Copy(S, P + 1, MaxInt));
      if StartsStr('packed ', Rhs) then Rhs := Trim(Copy(Rhs, 8, MaxInt));
      // 'TFoo = class;' is a forward declaration, 'class of' a metaclass
      if EndsStr(';', Rhs) and not StartsStr('record', Rhs) then Continue;
      if StartsStr('interface', Rhs) or StartsStr('dispinterface', Rhs) then
        Exit('interface');
      if (StartsStr('class', Rhs) and not StartsStr('class of', Rhs)) or
         StartsStr('record', Rhs) or StartsStr('object', Rhs) then
        Exit('class');
    end;
  end;

  // The kind is decided by the SECTION, not by comparing with the LSP's
  // definition line: DelphiLSP answers GotoDefinition for a METHOD with its
  // declaration in the class, but for a FREE routine with its implementation
  // header - comparing with that line labelled the implementation of a free
  // routine "Interface declaration" and its interface line "Declaration".
  function DetermineKind(const AFilePath: string; ALine, ACol: Integer;
    const AOrigLine: string; AImplLine: Integer; const ALines: TArray<string>): string;
  var
    IsHeader, DotBefore, InImplementation, InDefFile: Boolean;
  begin
    IsHeader := LineStartsWithMethodKeyword(AOrigLine);
    DotBefore := (ACol > 0) and (ACol <= Length(AOrigLine)) and (AOrigLine[ACol] = '.');
    InImplementation := (AImplLine >= 0) and (ALine > AImplLine);
    InDefFile := SameText(ExpandFileName(AFilePath), ExpandFileName(ADefFilePath));

    // Below 'implementation' an UNINDENTED header is a routine's
    // implementation; an indented one is a method declared in a class body
    // that lives in the implementation section.
    if IsHeader and (DotBefore or (InImplementation and (AOrigLine <> '') and
       not CharInSet(AOrigLine[1], [' ', #9]))) then
      Exit('Implementation');

    if IsHeader then
    begin
      // any other header is a declaration - of an interface member, a
      // class/record member, or a free routine
      var TypeKind := EnclosingTypeKind(ALines, ALine);
      if TypeKind = 'interface' then
        Exit('Interface declaration');
      if (TypeKind = 'class') or IsKnownImplFile(AFilePath) then
        Exit('Class declaration');
      Exit('Declaration');
    end;

    // Non-header line: a use in code, or a declaration-like use in the
    // interface section (e.g. 'property Bar: T read Bar;').
    if InDefFile and not InImplementation then
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
      var ImplLine := ImplementationLineOf(Lines);

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

        var FormKind: string;
        if FormHitKindAt(FE.FilePath, LineNo, StartCol, FormKind) then
          Item.Kind := FormKind
        else
          Item.Kind := DetermineKind(FE.FilePath, LineNo, StartCol, OrigLine, ImplLine, Lines);

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
  const AOldName, ANewName: string; AIncludes: TLspIncludeContext;
  const ATargets: TLspSymbolTargets; AClient: TLspClient): TLspWorkspaceEdit;
var
  FileMap: TDictionary<string, TList<TLspTextEdit>>;
  Synced: TDictionary<string, Boolean>;
  LineCache: TDictionary<string, TArray<string>>;
  Contents: TDictionary<string, string>;
  Graph: TTypeGraph;
  LastOpenedFile: string;
  VerifiedCount, SkippedCount, PreSkipped, I: Integer;
  C: TRenameCandidate;
  TextEdit: TLspTextEdit;

  function CandidateLine(const AC: TRenameCandidate): string;
  var
    L: TArray<string>;
  begin
    Result := '';
    if not LineCache.TryGetValue(UpperCase(AC.FilePath), L) then
    begin
      var Content: string;
      if not ((Editor <> nil) and Editor.ReadEditorContent(AC.FilePath, Content)) then
        try
          Content := ReadDelphiFile(AC.FilePath);
        except
          Content := '';
        end;
      L := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
      LineCache.Add(UpperCase(AC.FilePath), L);
    end;
    if (AC.Line >= 0) and (AC.Line <= High(L)) then Result := L[AC.Line];
  end;

  procedure Unverified(const AC: TRenameCandidate; const AWhy: string);
  begin
    FUnverified := FUnverified + [AC];
    FUnverifiedWhy := FUnverifiedWhy + [AWhy];
  end;

  // whole file (buffer first), read once - the type resolution below needs
  // more than the candidate's own line
  function FileContent(const AFile: string): string;
  begin
    if Contents.TryGetValue(UpperCase(AFile), Result) then Exit;
    if not ((Editor <> nil) and Editor.ReadEditorContent(AFile, Result)) then
      try
        Result := ReadDelphiFile(AFile);
      except
        Result := '';
      end;
    Contents.Add(UpperCase(AFile), Result);
  end;

begin
  FileMap := TDictionary<string, TList<TLspTextEdit>>.Create;
  Synced := TDictionary<string, Boolean>.Create;
  LineCache := TDictionary<string, TArray<string>>.Create;
  Contents := TDictionary<string, string>.Create;
  // types of the candidates' files; anything else is loaded on demand
  // through the identifier index
  var GraphFiles: TArray<string> := nil;
  for var GC in ACandidates do
  begin
    var Known := False;
    for var GF in GraphFiles do
      if SameText(GF, GC.FilePath) then Known := True;
    if not Known then GraphFiles := GraphFiles + [GC.FilePath];
  end;
  Graph := TTypeGraph.Create(GraphFiles, EditorOrDiskReader());
  try
    LastOpenedFile := '';
    VerifiedCount := 0;
    SkippedCount := 0;
    PreSkipped := 0;

    FHost.SetProgress(0, Length(ACandidates));

    for I := 0 to High(ACandidates) do
    begin
      if FHost.ScanCancelled then Break;
      C := ACandidates[I];
      FHost.SetProgress(I + 1, Length(ACandidates));
      if (I mod 3 = 0) then
        FHost.SetStatus(Format('Verifying %d/%d (ok:%d skip:%d)', [I + 1, Length(ACandidates), VerifiedCount, SkippedCount]));

      // PRE-CHECK without DelphiLSP: when the qualifier's declared type
      // says this occurrence is a member of ANOTHER type, no request is
      // needed at all. That is the only speed lever there is - DelphiLSP
      // answers "Request removed" to a second request while one is still
      // open, so verification is strictly one round trip per candidate
      // (measured 2026-09-20).
      begin
        var PreLink: TMemberLink;
        if ClassifyUnansweredUse(Graph, FileContent(C.FilePath), C.Line, C.Col,
          AOldName,
          function(AFile: string; ALine: Integer): Boolean
          begin
            Result := ATargets.Contains(AFile, ALine);
          end,
          function(AFile: string): Boolean
          begin
            Result := ATargets.ContainsFile(AFile);
          end, PreLink) = uuOtherSymbol then
        begin
          Inc(SkippedCount);
          Inc(PreSkipped);
          FDiagLog := FDiagLog + Format('  [%d] %s:%d:%d => member of %s -> ' +
            'SKIP (decided from the sources, no request)' + sLineBreak,
            [I, ExtractFileName(C.FilePath), C.Line + 1, C.Col + 1, PreLink.TypeName]);
          Continue;
        end;
      end;

      // Hand the file to DelphiLSP - ONLY when its content changed since it
      // was last sent, and then WAIT until the unit is analysed. The old
      // code re-opened every file on each switch and waited a fixed 300 ms:
      // a re-open restarts DelphiLSP's analysis (seconds for a big unit),
      // every query before it is done answers null, and null was SKIP -
      // real occurrences silently dropped out of the rename.
      // An include file / a unit sent expanded is served by the include
      // context instead.
      var FirstInFile := not Synced.ContainsKey(UpperCase(C.FilePath));
      if FirstInFile then
      begin
        Synced.Add(UpperCase(C.FilePath), True);
        if not AIncludes.OwnsDocument(C.FilePath) then
        begin
          var Before := AClient.GetFileDiagnosticsVersion(C.FilePath);
          if AClient.SyncDocument(C.FilePath) then
          begin
            var Name := ExtractFileName(C.FilePath);
            FHost.SetStatus('Waiting for DelphiLSP to analyse ' + Name + '...');
            if not AClient.WaitFileAnalysed(C.FilePath, Before, 30000,
              function: Boolean
              begin
                FHost.SetStatus('Waiting for DelphiLSP to analyse ' + Name + '...');
                Result := not FHost.ScanCancelled;
              end) then
              FDiagLog := FDiagLog + '  ' + Name + ': no analysis result within 30 s' + sLineBreak;
          end;
        end;
      end;
      LastOpenedFile := C.FilePath;

      var Matches := False;
      var DiagLine := Format('  [%d] %s:%d:%d => ', [I, ExtractFileName(C.FilePath), C.Line + 1, C.Col + 1]);

      try
        var Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
        // an EMPTY answer is retried briefly (the unit may still be in
        // analysis), a wrong one never
        if (Length(Defs) = 0) and not ATargets.Contains(C.FilePath, C.Line) then
        begin
          var Dl := GetTickCount64 + UInt64(IfThen(FirstInFile, 3000, 600));
          while (Length(Defs) = 0) and (GetTickCount64 < Dl) and not FHost.ScanCancelled do
          begin
            Sleep(150);
            FHost.SetStatus(Format('Verifying %d/%d (ok:%d skip:%d) - waiting for an answer',
              [I + 1, Length(ACandidates), VerifiedCount, SkippedCount]));
            Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
          end;
        end;

        // The candidate IS one of the symbol's positions (declaration /
        // implementation of the symbol or of an implementing class - see
        // TLspSymbolTargets), or DelphiLSP takes it to one of them. The
        // FILE alone is not enough: with several same-named methods in one
        // unit, every one of them was renamed (forum report 2026-09).
        if ATargets.Contains(C.FilePath, C.Line) then
        begin
          Matches := True;
          DiagLine := DiagLine + '(symbol position) -> MATCH';
        end
        else if (Length(Defs) = 0) and IsIncludeFile(C.FilePath) then
        begin
          DiagLine := DiagLine + 'null in an include file -> UNVERIFIED (listed, not renamed)';
          Unverified(C, 'include file');
        end
        else if (Length(Defs) = 0) and LineDeclaresName(CandidateLine(C), AOldName) then
          // DelphiLSP answers nothing AT a declaration; this one is not a
          // position of the symbol, so it declares ANOTHER symbol
          DiagLine := DiagLine + 'null at another declaration -> SKIP (other symbol)'
        else if Length(Defs) = 0 then
        begin
          // DelphiLSP said nothing. A DOTTED use site can still be decided
          // from the sources: resolve the qualifier's declared type and
          // look the member up there (the only way for the private/public
          // overload pair of RSS-5463, and it also covers occurrences in
          // an inactive {$IFDEF} branch, which used to stay unrenamed).
          var Link: TMemberLink;
          var Cls := ClassifyUnansweredUse(Graph, FileContent(C.FilePath),
            C.Line, C.Col, AOldName,
            function(AFile: string; ALine: Integer): Boolean
            begin
              Result := ATargets.Contains(AFile, ALine);
            end,
            function(AFile: string): Boolean
            begin
              Result := ATargets.ContainsFile(AFile);
            end, Link);
          case Cls of
            uuOurs:
              begin
                Matches := True;
                DiagLine := DiagLine + Format('null, resolved via %s -> MATCH',
                  [Link.TypeName]);
              end;
            uuOtherSymbol:
              DiagLine := DiagLine + Format('null, resolved via %s -> SKIP (other symbol)',
                [Link.TypeName]);
            uuOverloaded:
              begin
                DiagLine := DiagLine + Format('null, %s declares overloads -> UNVERIFIED',
                  [Link.TypeName]);
                Unverified(C, Format('overload of %s - no answer from DelphiLSP',
                  [Link.TypeName]));
              end;
          else
            // was a silent SKIP: a real occurrence DelphiLSP did not resolve
            // (inactive {$IFDEF} branch, unit still in analysis) must be
            // SHOWN, never dropped without a trace
            DiagLine := DiagLine + 'null -> UNVERIFIED (listed, not renamed)';
            Unverified(C, 'no answer from DelphiLSP');
          end;
        end
        else
        begin
          var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
          DiagLine := DiagLine + ExtractFileName(DefPath) + ':' + IntToStr(Defs[0].Range.Start.Line + 1);
          if ATargets.Contains(DefPath, Defs[0].Range.Start.Line) then
          begin
            Matches := True;
            DiagLine := DiagLine + ' -> MATCH';
          end
          else
            DiagLine := DiagLine + ' -> SKIP (another symbol)';
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

    FHost.SetProgress(Length(ACandidates), Length(ACandidates));
    if PreSkipped > 0 then
      FDiagLog := FDiagLog + Format(
        '%d of %d candidate(s) were decided from the sources - that many ' +
        'DelphiLSP requests saved.' + sLineBreak,
        [PreSkipped, Length(ACandidates)]);

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
    Graph.Free;
    Contents.Free;
    LineCache.Free;
    Synced.Free;
    FileMap.Free;
  end;
end;

end.
