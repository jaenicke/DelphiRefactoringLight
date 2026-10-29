(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ConstParamsDialog;

// The IDE half of the project check "Missing const (string parameters)":
// the scan over the project scope, the result window and the write.
//
// WHY A LIST and not a sweep: the rules in Expert.ConstParams can prove that
// nothing writes the parameter and that both its headers are known, but not
// that no caller releases the variable it is passed - a const string
// parameter aliases the caller's reference without raising its count, and
// that case dies with EInvalidPointer and no compiler warning (measured).
// So the check proposes with a reason per row and the user decides, which is
// also what the user asked for: "Das muss immer eine Liste sein."

interface

uses
  System.SysUtils,                 // TProc for the progress callback
  Vcl.Forms,                       // TForm, for the render seam below
  Expert.ConstParams;

/// <summary>Scans the project and opens the result window.</summary>
procedure CheckConstParams;

/// <summary>Every candidate of the project scope, with the refused ones when
///  AWITHREFUSED is set (the dialog counts them; the MCP answer can name
///  them, because "not proposed" and "never looked at" must not read the
///  same). ASUMMARY is the window's head line, ANOTES what the scan could
///  not do.
///  MAIN THREAD: it reads the project and the editor buffers.</summary>
function GatherConstParamFindings(AWithRefused: Boolean;
  out ASummary: string; out ANotes: TArray<string>;
  const AProgress: TProc<Integer, Integer, string> = nil): TArray<TConstParamFinding>;

/// <summary>Writes the const of ONE finding. It re-runs the check on the
///  file as it is NOW and takes the positions from that run, so a row from a
///  scan of ten minutes ago cannot write to a line that has moved - the
///  reason says so instead. MAIN THREAD (it writes through the editor).</summary>
function ApplyConstParamFinding(const AFinding: TConstParamFinding;
  out AWhy: string): Boolean;

/// <summary>Builds the result window without showing it. Exported so the
///  layout can be RENDERED headless before anyone clicks it - which is how
///  four defects of the Edit-methods dialog were found. Caller owns the
///  form.</summary>
function CreateConstParamsDialog(const AFindings: TArray<TConstParamFinding>;
  const ANotes: TArray<string>; const ASummary: string): TForm;

implementation

uses
  System.Classes, System.IOUtils, System.Math, System.Types,
  System.Generics.Collections, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.Graphics, Vcl.ClipBrd,
  Expert.EditorHelperIntf, Expert.DialogHelper, Expert.IdeThemes,
  Expert.ListViewSort, Expert.ScopeFiles, Expert.UsesEditor,
  Expert.InterfaceLinks, Expert.PascalScanner, Expert.DebugConsistency,
  Expert.AutoImport;                 // JoinContentLines

// ---- the scan -------------------------------------------------------------

// A class method's header says nothing about an interface that declares the
// same member - and const makes it stop satisfying it (measured: E2291). The
// project's type graph is the only thing that knows, so it answers for the
// pure collector.
function BoundLookupFor(AGraph: TTypeGraph): TConstBoundLookup;
begin
  Result :=
    function(const AOwnerType, AMember: string; out AWhy: string): Boolean
    begin
      AWhy := '';
      Result := False;
      if (AGraph = nil) or (AOwnerType = '') or (AMember = '') then Exit;
      try
        var Links := AGraph.InterfaceMethodsImplementedBy(AOwnerType, AMember);
        Result := Length(Links) > 0;
        if Result then
          AWhy := Format('%s implements %s, which declares %s - with const ' +
            'the class would stop satisfying it (E2291)',
            [AOwnerType, Links[0].TypeName, AMember]);
      except
        // A type graph that cannot answer must never cost the whole scan;
        // the row then simply keeps its other verdict.
        Result := False;
      end;
    end;
end;

function GatherConstParamFindings(AWithRefused: Boolean;
  out ASummary: string; out ANotes: TArray<string>;
  const AProgress: TProc<Integer, Integer, string>): TArray<TConstParamFinding>;
var
  Graph: TTypeGraph;
  Extra: TScopeExtra;
  Refused, Unreadable: Integer;
begin
  Result := nil;
  ANotes := nil;
  ASummary := '';
  Refused := 0;
  Unreadable := 0;

  if Assigned(AProgress) then AProgress(0, 1, 'collecting the project''s units...');
  var Files := ProjectScopeFiles('', True, False, Extra);
  if Length(Files) = 0 then
  begin
    ASummary := 'No project source files were found.';
    Exit;
  end;

  if Assigned(AProgress) then
    AProgress(0, Length(Files), 'reading the type graph...');
  Graph := TTypeGraph.Create(Files, nil, True);
  try
    var Bound := BoundLookupFor(Graph);
    for var I := 0 to High(Files) do
    begin
      if Assigned(AProgress) then
        AProgress(I, Length(Files), ExtractFileName(Files[I]));
      var Lines := ReadUnitLines(Files[I]);
      if Length(Lines) = 0 then
      begin
        Inc(Unreadable);
        Continue;
      end;
      try
        for var F in CollectConstParamCandidates(Files[I], Lines, Bound, True) do
          if F.Refusal = cprNone then
            Result := Result + [F]
          else
          begin
            Inc(Refused);
            if AWithRefused then Result := Result + [F];
          end;
      except
        on E: Exception do
          ANotes := ANotes + [Format('%s could not be analysed: %s',
            [ExtractFileName(Files[I]), E.Message])];
      end;
    end;
  finally
    Graph.Free;
  end;

  var Candidates := 0;
  for var F in Result do
    if F.Refusal = cprNone then Inc(Candidates);
  ASummary := Format('%d string parameter group(s) can take const. %s; %d ' +
    'const-less string parameter(s) were refused for a named reason.',
    [Candidates, ScopeExtraText(Length(Files), Extra), Refused]);
  if Unreadable > 0 then
    ANotes := ANotes + [Format('%d file(s) could not be read, so nothing ' +
      'about them was decided.', [Unreadable])];
  ANotes := ANotes + ['A const string parameter aliases the caller''s ' +
    'variable without raising its reference count. If the routine - or ' +
    'anything it calls - releases that variable, the parameter dangles, ' +
    'which no compiler warning catches. That is why these are proposals.'];
end;

function ApplyConstParamFinding(const AFinding: TConstParamFinding;
  out AWhy: string): Boolean;
var
  New: TArray<string>;
begin
  Result := False;
  AWhy := '';
  var Lines := ReadUnitLines(AFinding.FileName);
  if Length(Lines) = 0 then
  begin
    AWhy := 'the file could not be read';
    Exit;
  end;
  // Re-run the check on the text as it is NOW: the row may be minutes old,
  // and its positions are lines of the file it was found in.
  var Fresh := Default(TConstParamFinding);
  var Found := False;
  for var F in CollectConstParamCandidates(AFinding.FileName, Lines, nil) do
    if SameText(F.Id, AFinding.Id) then
    begin
      Fresh := F;
      Found := True;
    end;
  if not Found then
  begin
    AWhy := 'this parameter is no longer a candidate - run the check again';
    Exit;
  end;
  if not PlanConstParamEdit(Lines, Fresh, New, AWhy) then Exit;
  var SL := TStringList.Create;
  try
    SL.Text := JoinContentLines(New, string.Join(#13#10, Lines));
    Result := ApplyLinesMinimal(AFinding.FileName, SL,
      string.Join(#13#10, Lines));
    if not Result then
      AWhy := 'the write was refused - the buffer changed since it was read';
  finally
    SL.Free;
  end;
end;

// ---- the window -----------------------------------------------------------

type
  TConstParamsForm = class(TThemedToolForm)
  private
    FList: TListView;
    FDetail: TMemo;
    FBtnApply: TButton;
    FFindings: TArray<TConstParamFinding>;
    FNotes: TArray<string>;
    FBusy: Boolean;
    procedure FillRows;
    procedure DoSelect(Sender: TObject; Item: TListItem; Change: TItemChange);
    procedure DoGoto(Sender: TObject);
    procedure DoApply(Sender: TObject);
    procedure DoCopy(Sender: TObject);
    /// <summary>A button of a MODELESS form closes it by calling Close -
    ///  ModalResult is a plain field only ShowModal's loop reads (issue
    ///  #47).</summary>
    /// <summary>What the scan could not do - the detail pane's content
    ///  while no row is picked.</summary>
    procedure ShowNotes;
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    function SelectedFinding(out AFinding: TConstParamFinding): Boolean;
  public
    constructor CreateFor(const AFindings: TArray<TConstParamFinding>;
      const ANotes: TArray<string>; const ASummary: string); reintroduce;
  end;

constructor TConstParamsForm.CreateFor(const AFindings: TArray<TConstParamFinding>;
  const ANotes: TArray<string>; const ASummary: string);
var
  Bar: TPanel;
  Btn: TButton;
begin
  inherited CreateNew(Application);
  OnClose := DoFormClose;
  FFindings := AFindings;
  FNotes := ANotes;

  Caption := 'Missing const (string parameters)';
  Width := 980;
  Height := 620;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  var Head := TLabel.Create(Self);
  Head.Parent := Self;
  Head.Align := alTop;
  Head.AlignWithMargins := True;
  Head.Margins.SetBounds(10, 8, 10, 6);
  Head.AutoSize := False;
  Head.WordWrap := True;
  Head.Caption := ASummary;
  // The height comes from the TEXT, not from a number: the summary says how
  // many were found and how many refused, and a fixed height cuts the
  // second line off as soon as a project has more of either.
  Head.Height := TextHeightFor(Font, ASummary, Width - 40);

  // Two alBottom siblings are ordered by their TOP, not by the order they
  // were created in, so both say where they belong.
  FDetail := TMemo.Create(Self);
  FDetail.Parent := Self;
  FDetail.Top := 10000;
  FDetail.Align := alBottom;
  FDetail.Height := 110;
  FDetail.ReadOnly := True;
  FDetail.ScrollBars := ssVertical;
  FDetail.WordWrap := True;

  Bar := TPanel.Create(Self);
  Bar.Parent := Self;
  Bar.Top := 20000;
  Bar.Align := alBottom;
  Bar.BevelOuter := bvNone;
  Bar.Height := 44;

  FBtnApply := TButton.Create(Self);
  FBtnApply.Parent := Bar;
  FBtnApply.Caption := 'Add const to &checked';
  FBtnApply.SetBounds(8, 8, 160, 28);
  FBtnApply.OnClick := DoApply;
  // Never the Default button: Enter while browsing the list must not write
  // code (audit #41, M40e).

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Go to';
  Btn.SetBounds(176, 8, 90, 28);
  Btn.OnClick := DoGoto;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := 'Cop&y report';
  Btn.SetBounds(272, 8, 110, 28);
  Btn.OnClick := DoCopy;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Close';
  Btn.SetBounds(Bar.Width - 98, 8, 90, 28);
  Btn.Anchors := [akRight, akTop];
  Btn.OnClick := DoCloseClick;
  Btn.Cancel := True;

  FList := TThemedListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.HideSelection := False;
  FList.Checkboxes := True;
  FList.OnChange := DoSelect;
  FList.OnDblClick := DoGoto;
  // Measured against a real project's names: a qualified method name is the
  // longest cell by far ("TGemTiFunctions.IsConnectorUnreachable" was cut
  // at 250 in the render).
  with FList.Columns.Add do begin Caption := 'Routine'; Width := 300; end;
  with FList.Columns.Add do begin Caption := 'Parameter'; Width := 130; end;
  with FList.Columns.Add do begin Caption := 'Type'; Width := 70; end;
  with FList.Columns.Add do begin Caption := 'Unit'; Width := 200; end;
  with FList.Columns.Add do begin Caption := 'Verdict'; Width := 240; end;

  FillRows;
  EnableListViewSorting(FList);
  ShowNotes;
  PrepareDialog(Self, Application);
  EnableThemes(Self);
  TrackPluginForm(Self);
end;

procedure TConstParamsForm.FillRows;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var I := 0 to High(FFindings) do
    begin
      var F := FFindings[I];
      var It := FList.Items.Add;
      It.Caption := F.Routine;
      It.SubItems.Add(string.Join(', ', F.ParamNames));
      It.SubItems.Add(F.ParamType);
      It.SubItems.Add(ExtractFileName(F.FileName));
      It.SubItems.Add(ConstParamRefusalText(F.Refusal));
      // sortable rows map to the array ONLY via Data, never via Index
      It.Data := Pointer(NativeInt(I));
      It.Checked := F.Refusal = cprNone;
    end;
  finally
    FList.Items.EndUpdate;
  end;
end;

function TConstParamsForm.SelectedFinding(out AFinding: TConstParamFinding): Boolean;
begin
  Result := False;
  if FList.Selected = nil then Exit;
  var Idx := NativeInt(FList.Selected.Data);
  if (Idx < 0) or (Idx > High(FFindings)) then Exit;
  AFinding := FFindings[Idx];
  Result := True;
end;

procedure TConstParamsForm.ShowNotes;
var
  SL: TStringList;
begin
  // Until a row is picked the detail pane carries what the scan could NOT
  // do. A file nobody could read is not a file without a candidate, and an
  // empty pane would have said nothing at all (found by rendering it).
  SL := TStringList.Create;
  try
    SL.Add('Pick a row to see what was found and why.');
    if Length(FNotes) > 0 then
    begin
      SL.Add('');
      for var N in FNotes do SL.Add('Note: ' + N);
    end;
    FDetail.Text := SL.Text;
  finally
    SL.Free;
  end;
end;

procedure TConstParamsForm.DoSelect(Sender: TObject; Item: TListItem;
  Change: TItemChange);
var
  F: TConstParamFinding;
begin
  if Change <> ctState then Exit;
  if not SelectedFinding(F) then begin ShowNotes; Exit; end;
  FDetail.Text := Format('%s(%d): %s'#13#10#13#10'%s',
    [ExtractFileName(F.FileName), F.Line + 1,
     string.Join(', ', F.ParamNames) + ': ' + F.ParamType, F.Reason]);
end;

procedure TConstParamsForm.DoGoto(Sender: TObject);
var
  F: TConstParamFinding;
begin
  if not SelectedFinding(F) then Exit;
  if Editor = nil then Exit;
  // The column of the NAME, never the line start - a highlight counted from
  // column 0 marks the indentation instead (1.23.2).
  if not Editor.GotoLocation(F.FileName, F.Line, F.Col,
    Length(F.ParamNames[0])) then
    ShowThemedMessage('Could not open ' + F.FileName);
end;

procedure TConstParamsForm.DoApply(Sender: TObject);
var
  Ids: TArray<string>;
  Done, Failed: Integer;
  Report: TStringList;
begin
  if FBusy then Exit;                 // the progress pumps messages
  Ids := nil;
  for var It in FList.Items do
    if It.Checked then
    begin
      var Idx := NativeInt(It.Data);
      if (Idx >= 0) and (Idx <= High(FFindings))
        and (FFindings[Idx].Refusal = cprNone) then
        Ids := Ids + [FFindings[Idx].Id];
    end;
  if Length(Ids) = 0 then
  begin
    ShowThemedMessage('Nothing is checked that can take const.');
    Exit;
  end;
  if not AskThemedConfirm(
    Format('Write "const" in front of %d parameter group(s)?'#13#10#13#10 +
    'Both headers of each routine are changed. A const string parameter ' +
    'aliases the caller''s variable: if the routine releases it, the ' +
    'parameter dangles - which no compiler warning catches.', [Length(Ids)]),
    'Add const') then
    Exit;

  FBusy := True;
  Done := 0;
  Failed := 0;
  // Where the write goes is asked BEFORE it happens, and of a file that is
  // really in this run: a project-wide check reports mostly on files the
  // user never opened, so "changed in the editor" would be untrue for them
  // (user report 2026-10-05).
  var Note := '';
  for var C in FFindings do
    if SameText(C.Id, Ids[0]) then
      Note := MinimalWriteNote(MinimalWriteGoesToBuffer(C.FileName));
  Report := TStringList.Create;
  var Prog := CreateCheckProgress('Adding const', Self, 'Writing...');
  try
    // A SNAPSHOT of the ids, and the list disabled for the run: the progress
    // pumps messages, so a row must not move under the loop (audit #37, M22).
    FList.Enabled := False;
    FBtnApply.Enabled := False;
    for var I := 0 to High(Ids) do
    begin
      var F := Default(TConstParamFinding);
      for var C in FFindings do
        if SameText(C.Id, Ids[I]) then F := C;
      Prog.Step(I, Length(Ids), ExtractFileName(F.FileName));
      var Why := '';
      if ApplyConstParamFinding(F, Why) then
      begin
        Inc(Done);
        Report.Add(Format('%s %s: const added', [F.Routine,
          string.Join(', ', F.ParamNames)]));
      end
      else
      begin
        Inc(Failed);
        Report.Add(Format('%s %s: NOT changed - %s', [F.Routine,
          string.Join(', ', F.ParamNames), Why]));
      end;
    end;
  finally
    Prog.Free;
    FList.Enabled := True;
    FBtnApply.Enabled := True;
    FBusy := False;
  end;
  FDetail.Text := Report.Text;
  Report.Free;
  ShowThemedMessage(Format('%d group(s) changed, %d not changed.'#13#10 +
    '%s'#13#10#13#10'Run the check again to see the current state.',
    [Done, Failed, Note]));
end;

procedure TConstParamsForm.DoCopy(Sender: TObject);
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    for var F in FFindings do
      SL.Add(Format('%s(%d): %s %s: %s - %s', [F.FileName, F.Line + 1,
        F.Routine, string.Join(', ', F.ParamNames), F.ParamType, F.Reason]));
    for var N in FNotes do SL.Add('note: ' + N);
    Clipboard.AsText := SL.Text;
  finally
    SL.Free;
  end;
end;

procedure TConstParamsForm.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TConstParamsForm.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;      // every run opens a new one
end;

function CreateConstParamsDialog(const AFindings: TArray<TConstParamFinding>;
  const ANotes: TArray<string>; const ASummary: string): TForm;
begin
  Result := TConstParamsForm.CreateFor(AFindings, ANotes, ASummary);
end;

procedure CheckConstParams;
var
  Findings: TArray<TConstParamFinding>;
  Notes: TArray<string>;
  Summary: string;
begin
  var Prog := CreateCheckProgress('Missing const (string parameters)', nil,
    'collecting the project''s units...');
  try
    Findings := GatherConstParamFindings(False, Summary, Notes,
      procedure(ACur, ATotal: Integer; AText: string)
      begin
        Prog.Step(ACur, ATotal, AText);
      end);
  finally
    Prog.Free;
  end;
  if Length(Findings) = 0 then
  begin
    var Msg := 'No string parameter was found that can take const.';
    if Summary <> '' then Msg := Msg + #13#10#13#10 + Summary;
    ShowThemedMessage(Msg);
    Exit;
  end;
  CreateConstParamsDialog(Findings, Notes, Summary).Show;
end;

end.
