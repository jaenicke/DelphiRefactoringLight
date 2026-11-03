(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnusedUnitsDialog;

// The IDE half of the project check "Unused project units" (issue #47):
// gathering the facts the pure planner judges, and the result window.
//
// REPORT ONLY, by the user's decision ("bei #47 reicht die Information,
// selbst entfernen ist nicht noetig"): removing a unit edits the .dproj,
// which no Ctrl+Z reaches - so this check says what it found and the
// project manager does the removing.
//
// THE TWO PASSES, and the measurement behind them: reading every project
// file and scanning its uses clauses is cheap, running the production
// parser over it is not (15.7 ms per file, measured over a real project -
// half a minute for 1822 units). So pass 1 reads the files and asks only
// the uses question, and pass 2 parses the ~119 units that survived it.

interface

uses
  System.SysUtils,                 // TProc for the progress callback
  Vcl.Forms,                       // TForm, for the render seam below
  Expert.UnusedUnits;

/// <summary>Scans the project and opens the result window.</summary>
procedure CheckUnusedUnits;

/// <summary>The facts the planner needs: one entry per project unit, plus
///  one per unit that is merely reachable through a uses clause - a unit
///  on the search path is not in the project list and can still name a
///  project unit, and then that unit is used.
///  AError is set (and the result empty) when the check cannot apply at
///  all, which is what a package project is: the units it contains ARE
///  its purpose, so "nothing uses them" is normal there.
///  MAIN THREAD - it reads the project and the editor buffers.</summary>
function GatherUnusedUnitFacts(out AError: string;
  const AProgress: TProc<Integer, Integer, string> = nil): TArray<TUnitFacts>;

/// <summary>The expensive half, for the planner's callback: the
///  production parser over one file (initialization section + exported
///  identifiers). Any thread - it reads from disk.</summary>
function UnusedUnitDetailsReader: TUnusedUnitDetails;

/// <summary>Builds the result window without showing it. Exported so the
///  layout can be RENDERED headless (scratchpad unusedrender), which is
///  how four defects of the Edit-methods dialog were found before anyone
///  clicked it. Caller owns the form.</summary>
function CreateUnusedUnitsDialog(const AIssues: TArray<TUnusedUnitIssue>;
  const ANotes: TArray<string>; const ASummary: string): TForm;

implementation

uses
  System.Classes, System.IOUtils,
  System.Generics.Collections, System.UITypes,
  Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.ExtCtrls, Vcl.ClipBrd,
  Expert.DialogHelper, Expert.EditorHelperIntf, Expert.IdeThemes,
  Expert.ListViewSort, Expert.PascalScanner, Expert.ScopeFiles,
  Expert.UnitIndex, Expert.UsesGraph, Delphi.FileEncoding;

{ ---------- gathering ---------- }

function ReadUnitText(const AFile: string; out AText: string): Boolean;
begin
  // The editor buffer is the truth while a file is open - a uses entry the
  // user has just added or deleted is not on disk yet.
  if Editor.ReadEditorContent(AFile, AText) then Exit(True);
  AText := '';
  try
    Result := TFile.Exists(AFile);
    if Result then AText := TDelphiFileEncoding.ReadAll(AFile);
  except
    Result := False;
  end;
end;

function UnitNameOfFile(const AFile: string): string;
begin
  // Delphi requires a unit's name to match its file name, so the file name
  // IS the unit name - including the dotted form.
  Result := ChangeFileExt(ExtractFileName(AFile), '');
end;

function UsesNamesOf(const AText: string): TArray<string>;
var
  E: TUnitUsesEntry;
  L: TList<string>;
begin
  // ParseUsesEntries masks comments and strings itself and reports every
  // entry of both clauses - and it reads EVERY conditional branch, which
  // is what makes the answer hold for all configurations.
  L := TList<string>.Create;
  try
    for E in TUsesGraphAnalyzer.ParseUsesEntries(AText) do
      if E.UnitName <> '' then L.Add(E.UnitName);
    Result := L.ToArray;
  finally
    L.Free;
  end;
end;

function MainBodyOf(const AText: string): string;
var
  Entries: TArray<TUnitUsesEntry>;
  Lines: TArray<string>;
  LastLine, I: Integer;
  SB: TStringBuilder;
begin
  // The body = everything after the project's uses clause. That clause is
  // the unit LIST, and what matters here is what the body REACHES - the
  // Application.CreateForm(TForm1, Form1) line names the form's class, so
  // the body is matched against each candidate's exported identifiers.
  Entries := TUsesGraphAnalyzer.ParseUsesEntries(AText);
  LastLine := 0;
  for I := 0 to High(Entries) do
    if Entries[I].Line > LastLine then LastLine := Entries[I].Line;

  Lines := MaskCommentsAndStrings(SplitEditorLines(AText));
  SB := TStringBuilder.Create;
  try
    for I := LastLine to High(Lines) do     // Line is 1-based, so this skips it
      SB.AppendLine(Lines[I]);
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function UnusedUnitDetailsReader: TUnusedUnitDetails;
begin
  Result :=
    function (const AFile: string; out AHasInitCode: Boolean;
      out AExportedIdents: TArray<string>): Boolean
    var
      UName: string;
    begin
      AHasInitCode := False;
      AExportedIdents := nil;
      try
        AExportedIdents := ParseUnit(AFile, UName, AHasInitCode);
        Result := True;
      except
        Result := False;
      end;
    end;
end;

function ResolveListedUnit(const AName, AMainSource: string;
  out APath: string): Boolean;
var
  Snap: IUnitSnapshot;
begin
  // The identifier index knows the library and the project search paths,
  // which is where a unit listed without "in '<file>'" has to come from.
  // AMAINSOURCE is the .dpr whose list this is, so a group in which two
  // projects have a same-named unit resolves to the right one.
  APath := '';
  Snap := TUnitIndex.Instance.Snapshot;
  Result := (Snap <> nil) and Snap.TryGetUnitPathFor(AName, AMainSource, APath)
    and (APath <> '');
end;

/// <summary>The fact that answers a name AS THE PROJECT LISTS IT, or -1.
///  Exact first; then the last dotted segment, because a .dpr may list
///  "Forms" while the unit is "Vcl.Forms" (unit scope names). That loose
///  match errs towards SILENCE on purpose - taking two units for one can
///  only suppress a finding, never invent one.</summary>
function IndexOfListedUnit(const AFacts: TList<TUnitFacts>;
  const AByName: TDictionary<string, Integer>;
  const AName: string): Integer;
var
  Last, N: string;
  I: Integer;
begin
  if AByName.TryGetValue(LowerCase(AName), Result) then Exit;
  Last := AName;
  if Last.Contains('.') then
    Last := Last.Substring(Last.LastIndexOf('.') + 1);
  for I := 0 to AFacts.Count - 1 do
  begin
    N := AFacts[I].UnitName;
    if N.Contains('.') then N := N.Substring(N.LastIndexOf('.') + 1);
    if SameText(N, Last) then Exit(I);
  end;
  Result := -1;
end;

function GatherUnusedUnitFacts(out AError: string;
  const AProgress: TProc<Integer, Integer, string>): TArray<TUnitFacts>;
var
  Facts: TList<TUnitFacts>;
  Seen: TDictionary<string, Boolean>;
  ByName: TDictionary<string, Integer>;
  Dproj, MainFile, Path: string;
  ProjectFiles, ScopeAll, MainUses: TArray<string>;
  Extra: TScopeExtra;
  I, J: Integer;
  F: TUnitFacts;
  Text: string;
begin
  AError := '';
  Result := nil;

  Dproj := Editor.GetCurrentProjectDproj;
  if Dproj = '' then
  begin
    AError := 'No project is open.';
    Exit;
  end;

  // A PACKAGE is out of scope, and saying so is the honest answer: the
  // units a .dpk contains are what it exists for, so "no uses clause names
  // them" is the normal state and every one of them would be reported.
  MainFile := ChangeFileExt(Dproj, '.dpk');
  if TFile.Exists(MainFile) then
  begin
    AError := 'This is a package (' + ExtractFileName(MainFile) + '). The ' +
      'units a package contains are what it exists for, so this check ' +
      'cannot say anything useful about them.';
    Exit;
  end;
  MainFile := ChangeFileExt(Dproj, '.dpr');
  if not TFile.Exists(MainFile) then MainFile := '';

  if Assigned(AProgress) then
    AProgress(0, 0, 'reading the project''s unit list...');
  ProjectFiles := Editor.GetProjectSourceFiles;
  if Length(ProjectFiles) = 0 then
  begin
    AError := 'The project lists no source files.';
    Exit;
  end;

  // Reference sources BEYOND the project: the uses closure. A unit on the
  // search path never appears in the .dpr and can still name a project
  // unit - scanning only project files turns that into a false finding.
  // THE expensive step, and it has no progress of its own: on a project
  // of 1800 units the closure takes seconds, and before issue #47 nothing
  // said so - the window simply stood there.
  if Assigned(AProgress) then
    AProgress(0, 0, Format('following the uses clauses of %d project ' +
      'unit(s) - this can take a few seconds...', [Length(ProjectFiles)]));
  ScopeAll := ProjectScopeFiles(Editor.GetActiveFileName, True, True, Extra);

  Facts := TList<TUnitFacts>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  try
    // The main source first, so the planner has the body before anything
    // else. Its own uses clause is the project's unit list and the planner
    // ignores it by itself.
    if MainFile <> '' then
    begin
      F := Default(TUnitFacts);
      F.UnitName := UnitNameOfFile(MainFile);
      F.FileName := MainFile;
      F.InProject := True;
      F.IsMainUnit := True;
      F.Readable := ReadUnitText(MainFile, Text);
      if F.Readable then
      begin
        F.UsesNames := UsesNamesOf(Text);
        F.MainBody := MainBodyOf(Text);
      end;
      MainUses := F.UsesNames;    // the project's membership list
      Facts.Add(F);
      Seen.AddOrSetValue(LowerCase(MainFile), True);
    end;

    for I := 0 to High(ProjectFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(I + 1, Length(ProjectFiles) + Length(ScopeAll),
          ExtractFileName(ProjectFiles[I]));
      if Seen.ContainsKey(LowerCase(ProjectFiles[I])) then Continue;
      Seen.AddOrSetValue(LowerCase(ProjectFiles[I]), True);

      F := Default(TUnitFacts);
      F.UnitName := UnitNameOfFile(ProjectFiles[I]);
      F.FileName := ProjectFiles[I];
      F.InProject := True;
      F.Readable := ReadUnitText(ProjectFiles[I], Text);
      if F.Readable then F.UsesNames := UsesNamesOf(Text);
      Facts.Add(F);
    end;

    // Everything else of the closure is a reference source only.
    for I := 0 to High(ScopeAll) do
    begin
      if Assigned(AProgress) then
        AProgress(Length(ProjectFiles) + I + 1,
          Length(ProjectFiles) + Length(ScopeAll),
          ExtractFileName(ScopeAll[I]));
      if Seen.ContainsKey(LowerCase(ScopeAll[I])) then Continue;
      Seen.AddOrSetValue(LowerCase(ScopeAll[I]), True);
      if not SameText(ExtractFileExt(ScopeAll[I]), '.pas') then Continue;

      F := Default(TUnitFacts);
      F.UnitName := UnitNameOfFile(ScopeAll[I]);
      F.FileName := ScopeAll[I];
      F.InProject := False;
      F.Readable := ReadUnitText(ScopeAll[I], Text);
      if F.Readable then F.UsesNames := UsesNamesOf(Text);
      Facts.Add(F);
    end;

    // A unit the .dpr lists WITHOUT an "in '<file>'" clause carries no
    // file name in the project's module list, so not one of the loops
    // above ever saw it. MEASURED in a user's IDE (2026-10-17): the
    // project listed Unit7, no such file existed anywhere, and the check
    // answered "every unit this project lists is named somewhere". The
    // MAIN SOURCE'S OWN uses clause is that list, so it closes the gap.
    ByName := TDictionary<string, Integer>.Create;
    try
      for I := 0 to Facts.Count - 1 do
        ByName.AddOrSetValue(LowerCase(Facts[I].UnitName), I);

      for var Listed in MainUses do
      begin
        if Listed = '' then Continue;
        J := IndexOfListedUnit(Facts, ByName, Listed);
        if J >= 0 then
        begin
          // It came in as a reference source of the closure (Vcl.Forms and
          // its kind arrive that way). The project LISTS it, so it is a
          // candidate like every other entry.
          if not Facts[J].InProject then
          begin
            F := Facts[J];
            F.InProject := True;
            Facts[J] := F;
          end;
          Continue;
        end;

        F := Default(TUnitFacts);
        F.UnitName := Listed;
        F.InProject := True;
        // When nothing resolves the name there is nothing to read - and
        // THAT is the finding, not a reason to drop the entry.
        if ResolveListedUnit(Listed, MainFile, Path) and
           not Seen.ContainsKey(LowerCase(Path)) then
        begin
          Seen.AddOrSetValue(LowerCase(Path), True);
          F.FileName := Path;
          F.Readable := ReadUnitText(Path, Text);
          if F.Readable then F.UsesNames := UsesNamesOf(Text);
        end;
        ByName.AddOrSetValue(LowerCase(Listed), Facts.Count);
        Facts.Add(F);
      end;
    finally
      ByName.Free;
    end;

    Result := Facts.ToArray;
  finally
    Seen.Free;
    Facts.Free;
  end;
end;

{ ---------- the result window ---------- }

type
  TUnusedUnitsForm = class(TThemedToolForm)
  private
    FList: TListView;
    FDetail: TMemo;
    FIssues: TArray<TUnusedUnitIssue>;
    FNotes: TArray<string>;
    procedure FillRows;
    procedure DoSelect(Sender: TObject; Item: TListItem; Change: TItemChange);
    procedure DoGoto(Sender: TObject);
    /// <summary>A button of a MODELESS form closes it by calling Close -
    ///  ModalResult is a plain field that only ShowModal's loop reads, so
    ///  on this window the click (and Escape, through Cancel) did nothing
    ///  at all (issue #47).</summary>
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure DoCopy(Sender: TObject);
    function SelectedIssue(out AIssue: TUnusedUnitIssue): Boolean;
  public
    constructor CreateFor(const AIssues: TArray<TUnusedUnitIssue>;
      const ANotes: TArray<string>; const ASummary: string); reintroduce;
  end;

constructor TUnusedUnitsForm.CreateFor(const AIssues: TArray<TUnusedUnitIssue>;
  const ANotes: TArray<string>; const ASummary: string);
var
  Bar: TPanel;
  Btn: TButton;
begin
  inherited CreateNew(Application);
  OnClose := DoFormClose;
  FIssues := AIssues;
  FNotes := ANotes;

  Caption := 'Unused project units';
  Width := 980;
  Height := 620;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  var Head := TLabel.Create(Self);
  Head.Parent := Self;
  Head.Align := alTop;
  Head.AutoSize := False;
  Head.WordWrap := True;
  Head.Height := 40;
  Head.Caption := ASummary;

  // Two alBottom siblings are ordered by their TOP, not by the order
  // they were created in - swapping the creation order changed
  // nothing in the headless render. So both say where they belong.
  FDetail := TMemo.Create(Self);
  FDetail.Parent := Self;
  FDetail.Top := 10000;
  FDetail.Align := alBottom;
  FDetail.Height := 120;
  FDetail.ReadOnly := True;
  FDetail.ScrollBars := ssVertical;
  FDetail.WordWrap := True;

  Bar := TPanel.Create(Self);
  Bar.Parent := Self;
  Bar.Top := 20000;                 // below the detail pane
  Bar.Align := alBottom;
  Bar.BevelOuter := bvNone;
  Bar.Height := 44;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Go to';
  Btn.SetBounds(8, 8, 90, 28);
  Btn.OnClick := DoGoto;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Copy report';
  Btn.SetBounds(104, 8, 110, 28);
  Btn.OnClick := DoCopy;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Close';
  Btn.SetBounds(Bar.Width - 98, 8, 90, 28);
  Btn.Anchors := [akRight, akTop];
  Btn.OnClick := DoCloseClick;
  Btn.Cancel := True;          // Escape clicks it, so it runs OnClick too

  FList := TThemedListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.HideSelection := False;
  FList.OnChange := DoSelect;
  FList.OnDblClick := DoGoto;
  with FList.Columns.Add do begin Caption := 'Unit'; Width := 260; end;
  with FList.Columns.Add do begin Caption := 'What was found'; Width := 190; end;
  with FList.Columns.Add do begin Caption := 'Folder'; Width := 480; end;

  FillRows;
  EnableListViewSorting(FList);
  // The notes are the first thing in the detail pane: what limits the
  // answer must be readable without clicking a row.
  FDetail.Text := string.Join(sLineBreak, ANotes);
  PrepareDialog(Self, Application);
  EnableThemes(Self);
end;

procedure TUnusedUnitsForm.FillRows;
var
  I: Integer;
  It: TListItem;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for I := 0 to High(FIssues) do
    begin
      It := FList.Items.Add;
      It.Caption := FIssues[I].UnitName;
      It.SubItems.Add(UnusedUnitKindText(FIssues[I].Kind));
      It.SubItems.Add(ExtractFilePath(FIssues[I].FileName));
      // Rows map to the array through Data, never through Item.Index - the
      // list is sortable.
      It.Data := Pointer(NativeInt(I));
    end;
  finally
    FList.Items.EndUpdate;
  end;
end;

function TUnusedUnitsForm.SelectedIssue(out AIssue: TUnusedUnitIssue): Boolean;
var
  Idx: Integer;
begin
  AIssue := Default(TUnusedUnitIssue);
  Result := False;
  if FList.Selected = nil then Exit;
  Idx := NativeInt(FList.Selected.Data);
  if (Idx < 0) or (Idx > High(FIssues)) then Exit;
  AIssue := FIssues[Idx];
  Result := True;
end;

procedure TUnusedUnitsForm.DoSelect(Sender: TObject; Item: TListItem;
  Change: TItemChange);
var
  Issue: TUnusedUnitIssue;
begin
  if Change <> ctState then Exit;
  if not SelectedIssue(Issue) then Exit;
  FDetail.Text := Issue.FileName + sLineBreak + sLineBreak + Issue.Reason;
end;

procedure TUnusedUnitsForm.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TUnusedUnitsForm.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  // A result window of its own - every run opens a new one, so the closed
  // one has to go rather than wait for the package to unload.
  Action := caFree;
end;

procedure TUnusedUnitsForm.DoGoto(Sender: TObject);
var
  Issue: TUnusedUnitIssue;
  Dproj, Main, Text: string;
  L, C: Integer;
begin
  if not SelectedIssue(Issue) then Exit;

  if Issue.FileName <> '' then
  begin
    // The unit's first line, and nothing highlighted - there is no symbol
    // this check could point at.
    if not Editor.GotoLocation(Issue.FileName, 0, 0, 0) then
      ShowThemedMessage('Could not open ' + Issue.FileName + '.');
    Exit;
  end;

  // A finding with no file of its own: its PLACE is the ENTRY that lists
  // it, in the project's own source - which is also exactly what has to
  // go. The button used to return here without a word (user, 2026-10-17:
  // "Goto bei unused units macht gar nichts").
  Dproj := Editor.GetCurrentProjectDproj;
  Main := '';
  if Dproj <> '' then Main := ChangeFileExt(Dproj, '.dpr');
  if (Main <> '') and ReadUnitText(Main, Text) and
     UsesEntryPosition(Text, Issue.UnitName, L, C) then
  begin
    if not Editor.GotoLocation(Main, L, C, Length(Issue.UnitName)) then
      ShowThemedMessage('Could not open ' + Main + '.');
    Exit;
  end;

  // Never a silent no-op: say which of the two it is.
  if Main = '' then
    ShowThemedMessage(Issue.UnitName + ' has no file to open, and no ' +
      'project is open either - there is nothing to go to.')
  else
    ShowThemedMessage(Format('%s has no file to open: the project lists ' +
      'it without one. Its entry should be in the uses clause of %s, but ' +
      'that clause does not name it (any more) - the list has changed ' +
      'since the check ran.', [Issue.UnitName, ExtractFileName(Main)]));
end;

procedure TUnusedUnitsForm.DoCopy(Sender: TObject);
var
  SB: TStringBuilder;
  I: TUnusedUnitIssue;
  N: string;
begin
  SB := TStringBuilder.Create;
  try
    SB.AppendLine('Unused project units');
    SB.AppendLine;
    for I in FIssues do
      SB.AppendLine(Format('%-12s %-40s %s',
        [UnusedUnitKindText(I.Kind), I.UnitName, I.FileName]));
    SB.AppendLine;
    for N in FNotes do SB.AppendLine(N);
    Clipboard.AsText := SB.ToString;
  finally
    SB.Free;
  end;
end;

function CreateUnusedUnitsDialog(const AIssues: TArray<TUnusedUnitIssue>;
  const ANotes: TArray<string>; const ASummary: string): TForm;
begin
  Result := TUnusedUnitsForm.CreateFor(AIssues, ANotes, ASummary);
end;

{ ---------- entry point ---------- }

procedure CheckUnusedUnits;
var
  Facts: TArray<TUnitFacts>;
  Issues: TArray<TUnusedUnitIssue>;
  Notes: TArray<string>;
  Err, Summary: string;
  Prog: TCheckProgressWindow;
  Safe, Hint, NoSrc, Unknown, InProject: Integer;
  Form: TUnusedUnitsForm;
  I: TUnusedUnitIssue;
begin
  Prog := CreateCheckProgress('Unused project units', nil);
  try
    Facts := GatherUnusedUnitFacts(Err,
      procedure (ACur, ATotal: Integer; AWhat: string)
      begin
        Prog.Step(ACur, ATotal, AWhat);
      end);
    if Err <> '' then
    begin
      Prog.Free;
      Prog := nil;
      ShowThemedMessage(Err);
      Exit;
    end;

    Prog.Step(0, 0, 'parsing the units nothing names...');
    Issues := CollectUnusedProjectUnits(Facts, Notes, UnusedUnitDetailsReader());
  finally
    if Prog <> nil then Prog.Free;
  end;

  Safe := 0;
  Hint := 0;
  NoSrc := 0;
  Unknown := 0;
  for I in Issues do
    case I.Kind of
      uukSafe: Inc(Safe);
      uukInitCode: Inc(Hint);
      uukNoSource: Inc(NoSrc);
      uukUnreadable: Inc(Unknown);
    end;

  // Length(Facts) is project units PLUS the reference sources of the
  // closure - as a denominator that reads like a project twice the size
  // (measured live: 4270 for a project of 1822 units).
  InProject := ProjectUnitCount(Facts);
  Summary := Format('%d of %d project unit(s) are named by no uses clause: ' +
    '%d without an initialization section, %d with one (check those), ' +
    '%d with no source found, %d not checked. %d further unit(s) were ' +
    'read as reference sources.',
    [Length(Issues), InProject, Safe, Hint, NoSrc, Unknown,
     Length(Facts) - InProject]);

  if Length(Issues) = 0 then
  begin
    ShowThemedMessage('Every unit this project lists is named somewhere. ' +
      sLineBreak + sLineBreak + string.Join(sLineBreak, Notes));
    Exit;
  end;

  Form := TUnusedUnitsForm.CreateFor(Issues, Notes, Summary);
  Form.Show;                     // non-modal: it is a result list
end;

end.
