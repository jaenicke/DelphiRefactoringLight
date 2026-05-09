(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WithRefactorDialog;

{
  Modal review dialog for the project-wide "remove with" refactoring.

  Layout
  ------
    +-------------------------------------------------------------+
    | Status: scanned 23 units, found 47 with-statements          |
    | [progress bar]                                              |
    +-------------------------------------------------------------+
    | ListView: file | line | target | status                     |
    |  - Refactor.Foo.pas    142   FFoo            ok             |
    |  - Refactor.Foo.pas    287   GetClient(...)  ok (inline-var)|
    |  - Refactor.Bar.pas     33   A, B            needs review   |
    | ...                                                         |
    +---------------------------- splitter -----------------------+
    | Before                       | After                        |
    | [TMemo, monospaced, ro]      | [TMemo, monospaced, ro]      |
    |                              |                              |
    +-------------------------------------------------------------+
    |                       [Apply selected] [Apply all] [Close]  |
    +-------------------------------------------------------------+

  The dialog is purely presentational: it does NOT touch the file
  system. The orchestrator (Expert.WithRefactorWizard) decides what to
  do based on which apply-button was pressed and which row is selected.

  Modal results:
    mrCancel  - user closed without applying anything
    mrOk      - user pressed "Apply all" (all auto-rewritable)
    100       - user pressed "Apply selected" (read SelectedItem)
}

interface

uses
  System.Classes, System.SysUtils, System.UITypes, System.Generics.Collections,
  Winapi.Windows, Winapi.Messages,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Graphics,
  Expert.WithRewriter;

const
  /// <summary>ModalResult value used when the user pressed
  ///  "Apply selected". Distinct from mrOk so the caller can tell
  ///  the two intents apart.</summary>
  mrApplySelected = 100;

type
  TWithRefactorDialog = class(TForm)
  private
    FStatusLabel: TLabel;
    FProgress: TProgressBar;
    FListView: TListView;
    FBeforeLabel: TLabel;
    FAfterLabel: TLabel;
    FBeforeMemo: TMemo;
    FAfterMemo: TMemo;
    FDebugMemo: TMemo;
    FSplitterMain: TSplitter;
    FSplitterPreview: TSplitter;
    FBtnApplySelected: TButton;
    FBtnApplyAll: TButton;
    FBtnClose: TButton;
    FPreviewPanel: TPanel;
    FPageControl: TPageControl;
    FTabDiff: TTabSheet;
    FTabDebug: TTabSheet;
    FDiffPanel: TPanel;
    FBeforePanel: TPanel;
    FAfterPanel: TPanel;
    FBtnPanel: TPanel;
    FItems: TArray<TWithRewriteResult>;
    FAutoRewritableCount: Integer;
    FCommonPathPrefix: string;

    procedure CreateControls;
    procedure DoListSelect(Sender: TObject; AItem: TListItem; Selected: Boolean);
    procedure DoBtnApplySelectedClick(Sender: TObject);
    procedure DoBtnApplyAllClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure UpdatePreview(AIndex: Integer);
    procedure UpdateDebug(AIndex: Integer);
    procedure UpdateApplyButtons;
    function ComputeCommonPathPrefix: string;
    function StatusTextOf(const AItem: TWithRewriteResult): string;
    function TargetSummaryOf(const AItem: TWithRewriteResult): string;
    function BuildDebugText(const AItem: TWithRewriteResult): string;
  public
    constructor CreateDialog(AOwner: TComponent); reintroduce;

    /// <summary>Replaces the current item list and updates the UI.</summary>
    procedure SetItems(const AItems: TArray<TWithRewriteResult>);

    /// <summary>Updates the status text shown above the list.</summary>
    procedure SetStatus(const AText: string);

    /// <summary>Updates the progress bar (0..ATotal).</summary>
    procedure SetProgress(ACurrent, ATotal: Integer);

    /// <summary>Returns the index in the items array of the currently
    ///  selected ListView row, or -1 if none.</summary>
    function SelectedIndex: Integer;

    /// <summary>The full item list (read-only).</summary>
    property Items: TArray<TWithRewriteResult> read FItems;
  end;

implementation

uses
  System.IOUtils;

{ TWithRefactorDialog }

constructor TWithRefactorDialog.CreateDialog(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  Caption := 'Remove with - project-wide review';
  Position := poOwnerFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  Width := 1000;
  Height := 700;
  Constraints.MinWidth := 700;
  Constraints.MinHeight := 460;
  KeyPreview := True;
  OnKeyDown := DoFormKeyDown;

  CreateControls;
  UpdateApplyButtons;
end;

procedure TWithRefactorDialog.CreateControls;
var
  Col: TListColumn;
begin
  // Status
  FStatusLabel := TLabel.Create(Self);
  FStatusLabel.Parent := Self;
  FStatusLabel.Align := alTop;
  FStatusLabel.AlignWithMargins := True;
  FStatusLabel.Margins.SetBounds(8, 8, 8, 4);
  FStatusLabel.Caption := 'Scanning project...';

  // Progress
  FProgress := TProgressBar.Create(Self);
  FProgress.Parent := Self;
  FProgress.Align := alTop;
  FProgress.AlignWithMargins := True;
  FProgress.Margins.SetBounds(8, 0, 8, 4);
  FProgress.Height := 16;
  FProgress.Min := 0;
  FProgress.Max := 100;
  FProgress.Position := 0;

  // Bottom: buttons
  FBtnPanel := TPanel.Create(Self);
  FBtnPanel.Parent := Self;
  FBtnPanel.Align := alBottom;
  FBtnPanel.Height := 44;
  FBtnPanel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := FBtnPanel;
  FBtnClose.Caption := 'Close';
  FBtnClose.Width := 110;
  FBtnClose.Height := 30;
  FBtnClose.Top := 8;
  FBtnClose.Anchors := [akTop, akRight];
  FBtnClose.Left := FBtnPanel.Width - FBtnClose.Width - 8;
  FBtnClose.OnClick := DoBtnCloseClick;
  FBtnClose.Cancel := True;

  FBtnApplyAll := TButton.Create(Self);
  FBtnApplyAll.Parent := FBtnPanel;
  FBtnApplyAll.Caption := 'Apply all';
  FBtnApplyAll.Width := 130;
  FBtnApplyAll.Height := 30;
  FBtnApplyAll.Top := 8;
  FBtnApplyAll.Anchors := [akTop, akRight];
  FBtnApplyAll.Left := FBtnClose.Left - FBtnApplyAll.Width - 6;
  FBtnApplyAll.OnClick := DoBtnApplyAllClick;

  FBtnApplySelected := TButton.Create(Self);
  FBtnApplySelected.Parent := FBtnPanel;
  FBtnApplySelected.Caption := 'Apply selected';
  FBtnApplySelected.Width := 130;
  FBtnApplySelected.Height := 30;
  FBtnApplySelected.Top := 8;
  FBtnApplySelected.Anchors := [akTop, akRight];
  FBtnApplySelected.Left := FBtnApplyAll.Left - FBtnApplySelected.Width - 6;
  FBtnApplySelected.OnClick := DoBtnApplySelectedClick;
  FBtnApplySelected.Default := True;

  // Bottom-half: preview panel hosts a TPageControl with two tabs:
  //   - "Diff"  : Before / After memos side by side
  //   - "Debug" : raw target / LSP info per body identifier
  FPreviewPanel := TPanel.Create(Self);
  FPreviewPanel.Parent := Self;
  FPreviewPanel.Align := alBottom;
  FPreviewPanel.Height := 320;
  FPreviewPanel.BevelOuter := bvNone;

  // Splitter between list and preview
  FSplitterMain := TSplitter.Create(Self);
  FSplitterMain.Parent := Self;
  FSplitterMain.Align := alBottom;
  FSplitterMain.Top := FPreviewPanel.Top - 1;
  FSplitterMain.Height := 4;
  FSplitterMain.MinSize := 120;
  FSplitterMain.ResizeStyle := rsUpdate;

  FPageControl := TPageControl.Create(Self);
  FPageControl.Parent := FPreviewPanel;
  FPageControl.Align := alClient;

  FTabDiff := TTabSheet.Create(FPageControl);
  FTabDiff.PageControl := FPageControl;
  FTabDiff.Caption := 'Diff';

  FTabDebug := TTabSheet.Create(FPageControl);
  FTabDebug.PageControl := FPageControl;
  FTabDebug.Caption := 'Debug';

  // ---- Diff tab: Before / After ----
  FDiffPanel := TPanel.Create(Self);
  FDiffPanel.Parent := FTabDiff;
  FDiffPanel.Align := alClient;
  FDiffPanel.BevelOuter := bvNone;

  // Before-panel (left)
  FBeforePanel := TPanel.Create(Self);
  FBeforePanel.Parent := FDiffPanel;
  FBeforePanel.Align := alLeft;
  FBeforePanel.Width := FDiffPanel.Width div 2;
  FBeforePanel.BevelOuter := bvNone;

  FBeforeLabel := TLabel.Create(Self);
  FBeforeLabel.Parent := FBeforePanel;
  FBeforeLabel.Align := alTop;
  FBeforeLabel.AlignWithMargins := True;
  FBeforeLabel.Margins.SetBounds(8, 4, 8, 2);
  FBeforeLabel.Caption := 'Before';
  FBeforeLabel.Font.Style := [fsBold];

  FBeforeMemo := TMemo.Create(Self);
  FBeforeMemo.Parent := FBeforePanel;
  FBeforeMemo.Align := alClient;
  FBeforeMemo.AlignWithMargins := True;
  FBeforeMemo.Margins.SetBounds(8, 0, 4, 4);
  FBeforeMemo.ReadOnly := True;
  FBeforeMemo.ScrollBars := ssBoth;
  FBeforeMemo.WordWrap := False;
  FBeforeMemo.Font.Name := 'Consolas';
  FBeforeMemo.Font.Size := 10;
  FBeforeMemo.Color := clBtnFace;

  // Splitter between Before and After
  FSplitterPreview := TSplitter.Create(Self);
  FSplitterPreview.Parent := FDiffPanel;
  FSplitterPreview.Align := alLeft;
  FSplitterPreview.Width := 4;
  FSplitterPreview.MinSize := 120;
  FSplitterPreview.ResizeStyle := rsUpdate;

  // After-panel (right, fills remaining space)
  FAfterPanel := TPanel.Create(Self);
  FAfterPanel.Parent := FDiffPanel;
  FAfterPanel.Align := alClient;
  FAfterPanel.BevelOuter := bvNone;

  FAfterLabel := TLabel.Create(Self);
  FAfterLabel.Parent := FAfterPanel;
  FAfterLabel.Align := alTop;
  FAfterLabel.AlignWithMargins := True;
  FAfterLabel.Margins.SetBounds(8, 4, 8, 2);
  FAfterLabel.Caption := 'After';
  FAfterLabel.Font.Style := [fsBold];

  FAfterMemo := TMemo.Create(Self);
  FAfterMemo.Parent := FAfterPanel;
  FAfterMemo.Align := alClient;
  FAfterMemo.AlignWithMargins := True;
  FAfterMemo.Margins.SetBounds(4, 0, 8, 4);
  FAfterMemo.ReadOnly := True;
  FAfterMemo.ScrollBars := ssBoth;
  FAfterMemo.WordWrap := False;
  FAfterMemo.Font.Name := 'Consolas';
  FAfterMemo.Font.Size := 10;

  // ---- Debug tab: raw resolution info ----
  FDebugMemo := TMemo.Create(Self);
  FDebugMemo.Parent := FTabDebug;
  FDebugMemo.Align := alClient;
  FDebugMemo.AlignWithMargins := True;
  FDebugMemo.Margins.SetBounds(8, 8, 8, 8);
  FDebugMemo.ReadOnly := True;
  FDebugMemo.ScrollBars := ssBoth;
  FDebugMemo.WordWrap := False;
  FDebugMemo.Font.Name := 'Consolas';
  FDebugMemo.Font.Size := 9;
  FDebugMemo.Color := clBtnFace;

  // ListView fills the rest
  FListView := TListView.Create(Self);
  FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.AlignWithMargins := True;
  FListView.Margins.SetBounds(8, 4, 8, 0);
  FListView.ViewStyle := vsReport;
  FListView.RowSelect := True;
  FListView.ReadOnly := True;
  FListView.HideSelection := False;
  FListView.GridLines := True;
  FListView.OnSelectItem := DoListSelect;

  Col := FListView.Columns.Add;
  Col.Caption := 'File';
  Col.Width := 280;

  Col := FListView.Columns.Add;
  Col.Caption := 'Line';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Target';
  Col.Width := 240;

  Col := FListView.Columns.Add;
  Col.Caption := 'Status';
  Col.Width := 200;
end;

function TWithRefactorDialog.ComputeCommonPathPrefix: string;
var
  I, J, MaxLen: Integer;
  P: string;
begin
  Result := '';
  if Length(FItems) = 0 then Exit;
  Result := ExtractFilePath(FItems[0].FileName);
  for I := 1 to High(FItems) do
  begin
    P := ExtractFilePath(FItems[I].FileName);
    MaxLen := Length(Result);
    if Length(P) < MaxLen then MaxLen := Length(P);
    J := 1;
    while (J <= MaxLen) and (UpCase(Result[J]) = UpCase(P[J])) do
      Inc(J);
    Result := Copy(Result, 1, J - 1);
    if Result = '' then Exit;
  end;
end;

function TWithRefactorDialog.StatusTextOf(const AItem: TWithRewriteResult): string;
begin
  if AItem.IsAutoRewritable then
  begin
    Result := 'ok';
    if Pos('__with', AItem.NewText) > 0 then
      Result := 'ok (inline-var)';
  end
  else if wriMultipleTargets in AItem.Issues then
    Result := 'multi-target - manual review'
  else if wriTypeUnresolved in AItem.Issues then
    Result := 'type unresolved'
  else if wriClassRangeUnknown in AItem.Issues then
    Result := 'class range unknown'
  else if wriNameClash in AItem.Issues then
    Result := 'name clash'
  else
    Result := 'no rewrite';
end;

function TWithRefactorDialog.TargetSummaryOf(const AItem: TWithRewriteResult): string;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to High(AItem.Occurrence.Targets) do
  begin
    if I > 0 then Result := Result + ', ';
    Result := Result + AItem.Occurrence.Targets[I].Expression;
  end;
end;

procedure TWithRefactorDialog.SetItems(const AItems: TArray<TWithRewriteResult>);
var
  I: Integer;
  LI: TListItem;
  DisplayPath: string;
begin
  FItems := AItems;
  FCommonPathPrefix := ComputeCommonPathPrefix;
  FAutoRewritableCount := 0;

  FListView.Items.BeginUpdate;
  try
    FListView.Clear;
    for I := 0 to High(AItems) do
    begin
      LI := FListView.Items.Add;
      DisplayPath := AItems[I].FileName;
      if (FCommonPathPrefix <> '') and DisplayPath.StartsWith(FCommonPathPrefix, True) then
        DisplayPath := Copy(DisplayPath, Length(FCommonPathPrefix) + 1, MaxInt);
      LI.Caption := DisplayPath;
      LI.SubItems.Add(IntToStr(AItems[I].Occurrence.KeywordPos.Line));
      LI.SubItems.Add(TargetSummaryOf(AItems[I]));
      LI.SubItems.Add(StatusTextOf(AItems[I]));

      if AItems[I].IsAutoRewritable then
        Inc(FAutoRewritableCount);
    end;
    if FListView.Items.Count > 0 then
    begin
      FListView.Items[0].Selected := True;
      FListView.Items[0].Focused := True;
      UpdatePreview(0);
    end
    else
    begin
      FBeforeMemo.Clear;
      FAfterMemo.Clear;
      FDebugMemo.Clear;
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  UpdateApplyButtons;
end;

procedure TWithRefactorDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TWithRefactorDialog.SetProgress(ACurrent, ATotal: Integer);
begin
  if ATotal <= 0 then
  begin
    FProgress.Position := 0;
    Exit;
  end;
  FProgress.Max := ATotal;
  if ACurrent > ATotal then ACurrent := ATotal;
  FProgress.Position := ACurrent;
end;

function TWithRefactorDialog.SelectedIndex: Integer;
begin
  if Assigned(FListView.Selected) then
    Result := FListView.Selected.Index
  else
    Result := -1;
end;

procedure TWithRefactorDialog.UpdatePreview(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex > High(FItems)) then
  begin
    FBeforeMemo.Clear;
    FAfterMemo.Clear;
    UpdateDebug(AIndex);
    Exit;
  end;
  FBeforeMemo.Lines.Text := FItems[AIndex].OriginalText;
  if FItems[AIndex].NewText <> '' then
    FAfterMemo.Lines.Text := FItems[AIndex].NewText
  else
    FAfterMemo.Lines.Text := '<no rewrite available - ' +
      StatusTextOf(FItems[AIndex]) + '>';
  UpdateDebug(AIndex);
end;

procedure TWithRefactorDialog.UpdateDebug(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex > High(FItems)) then
  begin
    FDebugMemo.Clear;
    Exit;
  end;
  FDebugMemo.Lines.Text := BuildDebugText(FItems[AIndex]);
end;

function TWithRefactorDialog.BuildDebugText(const AItem: TWithRewriteResult): string;
var
  SB: TStringBuilder;
  I, J: Integer;
  T: TWithDebugTargetInfo;
  R: TWithDebugRefInfo;
  MatchSrc, TargetIdx: string;
const
  MatchNames: array[TWithDebugMatch] of string = (
    'none (kept as-is)', 'member (parsed)', 'LSP');
begin
  SB := TStringBuilder.Create;
  try
    SB.Append('File: ').AppendLine(AItem.FileName);
    SB.Append('with-keyword line: ')
      .Append(AItem.Occurrence.KeywordPos.Line)
      .Append(', col ').AppendLine(IntToStr(AItem.Occurrence.KeywordPos.Col));
    SB.AppendLine;

    // Targets
    SB.Append('Targets (').Append(Length(AItem.Debug.Targets)).AppendLine('):');
    if Length(AItem.Debug.Targets) = 0 then
      SB.AppendLine('  <none resolved>');
    for I := 0 to High(AItem.Debug.Targets) do
    begin
      T := AItem.Debug.Targets[I];
      SB.Append('  [').Append(I).Append('] ').AppendLine(T.Expression);
      SB.Append('       resolved      : ').AppendLine(BoolToStr(T.Resolved, True));
      SB.Append('       type file     : ').AppendLine(T.TypeFile);
      SB.Append('       class lines   : ')
        .Append(T.ClassStartLine).Append('..').AppendLine(IntToStr(T.ClassEndLine));
      SB.Append('       inline-var    : ');
      if T.InlineVarName = '' then
        SB.AppendLine('(none, prefix used directly)')
      else
        SB.AppendLine(T.InlineVarName);
      SB.Append('       qualify-prefix: ').AppendLine(T.QualifyPrefix);
      SB.Append('       direct members (').Append(Length(T.Members)).Append('): ');
      if Length(T.Members) = 0 then
        SB.AppendLine('<none / parser found nothing>')
      else
      begin
        for J := 0 to High(T.Members) do
        begin
          if J > 0 then SB.Append(', ');
          SB.Append(T.Members[J]);
        end;
        SB.AppendLine;
      end;
    end;

    SB.AppendLine;

    // Body identifiers
    SB.Append('Body identifiers (').Append(Length(AItem.Debug.Refs)).AppendLine('):');
    if Length(AItem.Debug.Refs) = 0 then
      SB.AppendLine('  <none scanned>');
    for I := 0 to High(AItem.Debug.Refs) do
    begin
      R := AItem.Debug.Refs[I];
      MatchSrc := MatchNames[R.MatchSource];
      if R.MatchedTargetIdx >= 0 then
        TargetIdx := Format(' -> target [%d]', [R.MatchedTargetIdx])
      else
        TargetIdx := '';
      SB.Append(Format('  %-20s @L%d:C%d  match=%s%s',
        [R.Name, R.Line, R.Col, MatchSrc, TargetIdx])).AppendLine;
      if R.AppliedPrefix <> '' then
        SB.Append('       applied prefix: "').Append(R.AppliedPrefix).AppendLine('"');
      if R.LspHadResult then
      begin
        SB.Append('       LSP result    : ').Append(R.LspFile)
          .Append(' line ').AppendLine(IntToStr(R.LspLine));
      end
      else if R.MatchSource <> dmMember then
      begin
        SB.AppendLine('       LSP result    : <none / not queried>');
      end;
    end;

    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure TWithRefactorDialog.UpdateApplyButtons;
var
  Idx: Integer;
begin
  Idx := SelectedIndex;
  FBtnApplySelected.Enabled := (Idx >= 0) and (Idx <= High(FItems))
    and FItems[Idx].IsAutoRewritable;
  FBtnApplyAll.Enabled := FAutoRewritableCount > 0;
  if FAutoRewritableCount > 0 then
    FBtnApplyAll.Caption := Format('Apply all (%d)', [FAutoRewritableCount])
  else
    FBtnApplyAll.Caption := 'Apply all';
end;

procedure TWithRefactorDialog.DoListSelect(Sender: TObject; AItem: TListItem; Selected: Boolean);
begin
  if Selected and Assigned(AItem) then
    UpdatePreview(AItem.Index);
  UpdateApplyButtons;
end;

procedure TWithRefactorDialog.DoBtnApplySelectedClick(Sender: TObject);
begin
  if (SelectedIndex >= 0) and FItems[SelectedIndex].IsAutoRewritable then
    ModalResult := mrApplySelected;
end;

procedure TWithRefactorDialog.DoBtnApplyAllClick(Sender: TObject);
begin
  if FAutoRewritableCount > 0 then
    ModalResult := mrOk;
end;

procedure TWithRefactorDialog.DoBtnCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TWithRefactorDialog.DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    ModalResult := mrCancel;
    Key := 0;
  end;
end;

end.
