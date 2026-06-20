(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SemanticReplaceDialogs;

// Dialogs for the Semantic Replace workflow:
//   - TSemanticReplaceRulesListDialog: list / add / edit / delete rules
//   - TSemanticReplaceRuleEditDialog:  edit a single rule
//   - TSemanticReplaceUnitsDialog:     multi-pick of project .pas files
//   - TSemanticReplacePreviewDialog:   show the planned edits + apply

interface

uses
  System.SysUtils, System.Classes, System.Types,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.CheckLst, Vcl.ComCtrls, Vcl.Dialogs, Vcl.Menus,
  Expert.SemanticReplace;

type
  TSemanticReplaceRuleEditDialog = class(TForm)
  private
    FEdtFind, FEdtReplace, FEdtUses: TEdit;
    FEdtVarName, FEdtVarType, FEdtVarValue, FEdtVarReplace: TEdit;
    FBtnOk, FBtnCancel: TButton;
    procedure BuildLayout;
  public
    constructor CreateDialog(AOwner: TComponent;
      const ARule: TSemanticReplaceRule);
    function GetRule: TSemanticReplaceRule;
    class function Edit(AOwner: TComponent;
      var ARule: TSemanticReplaceRule): Boolean;
  end;

  TSemanticReplaceRulesListDialog = class(TForm)
  private
    FRules: TArray<TSemanticReplaceRule>;
    FListView: TListView;
    FBtnAdd, FBtnEdit, FBtnDelete, FBtnOk, FBtnCancel: TButton;
    procedure BuildLayout;
    procedure RefreshList;
    procedure DoAdd(Sender: TObject);
    procedure DoEdit(Sender: TObject);
    procedure DoDelete(Sender: TObject);
    procedure DoDblClick(Sender: TObject);
  public
    constructor CreateDialog(AOwner: TComponent;
      const ARules: TArray<TSemanticReplaceRule>);
    property Rules: TArray<TSemanticReplaceRule> read FRules;
    class function Edit(AOwner: TComponent;
      var ARules: TArray<TSemanticReplaceRule>): Boolean;
  end;

  TSemanticReplaceUnitsDialog = class(TForm)
  private
    FAllFiles: TArray<string>;
    FCheckList: TCheckListBox;
    FEdtFilter: TEdit;
    FBtnAll, FBtnNone, FBtnOk, FBtnCancel: TButton;
    procedure BuildLayout;
    procedure RefreshList;
    procedure DoFilterChange(Sender: TObject);
    procedure DoAll(Sender: TObject);
    procedure DoNone(Sender: TObject);
  public
    constructor CreateDialog(AOwner: TComponent;
      const AAllFiles: TArray<string>);
    function SelectedFiles: TArray<string>;
    class function Choose(AOwner: TComponent;
      const AAllFiles: TArray<string>;
      out AChosen: TArray<string>): Boolean;
  end;

  TSemanticReplacePreviewDialog = class(TForm)
  private
    FMemo: TMemo;
    FBtnApply, FBtnCancel: TButton;
    FLblSummary: TLabel;
    procedure BuildLayout;
  public
    constructor CreateDialog(AOwner: TComponent;
      const ASummary, APreview: string);
    class function Confirm(AOwner: TComponent;
      const ASummary, APreview: string): Boolean;
  end;

implementation

uses
  System.UITypes, System.IOUtils, System.StrUtils;

{ ---------- TSemanticReplaceRuleEditDialog ---------- }

constructor TSemanticReplaceRuleEditDialog.CreateDialog(AOwner: TComponent;
  const ARule: TSemanticReplaceRule);
begin
  inherited CreateNew(AOwner);
  Caption := 'Semantic replace rule';
  Position := poMainFormCenter;
  BorderStyle := bsDialog;
  ClientWidth := 700;
  ClientHeight := 380;
  BuildLayout;

  FEdtFind.Text := ARule.Find;
  FEdtReplace.Text := ARule.Replace;
  FEdtUses.Text := string.Join(', ', ARule.UsesToAdd);
  FEdtVarName.Text := ARule.LocalVarName;
  FEdtVarType.Text := ARule.LocalVarType;
  FEdtVarValue.Text := ARule.LocalVarValue;
  FEdtVarReplace.Text := ARule.ReplaceWhenLocalVar;
end;

procedure TSemanticReplaceRuleEditDialog.BuildLayout;
const
  LeftCol = 12;
  EditCol = 180;
  EditW = 500;
  RowH = 26;
  RowGap = 6;
var
  Y: Integer;
  Lbl: TLabel;
  procedure Row(const ALabel: string; out AEdit: TEdit);
  begin
    Lbl := TLabel.Create(Self); Lbl.Parent := Self;
    Lbl.SetBounds(LeftCol, Y + 4, 160, 16);
    Lbl.Caption := ALabel;
    AEdit := TEdit.Create(Self); AEdit.Parent := Self;
    AEdit.SetBounds(EditCol, Y, EditW, RowH);
    Inc(Y, RowH + RowGap);
  end;
begin
  Y := 12;
  Row('Find:',                              FEdtFind);
  Row('Replace:',                           FEdtReplace);
  Row('Uses (comma-separated):',            FEdtUses);
  Inc(Y, RowGap);

  Lbl := TLabel.Create(Self); Lbl.Parent := Self;
  Lbl.SetBounds(LeftCol, Y, EditCol + EditW - LeftCol, 16);
  Lbl.Caption := 'Local var (optional - introduced when the rule fires >=2 times in the same routine):';
  Inc(Y, 18 + RowGap);

  Row('  Name:',                            FEdtVarName);
  Row('  Type:',                            FEdtVarType);
  Row('  Value:',                           FEdtVarValue);
  Row('  Replace (when var hoisted):',      FEdtVarReplace);

  FBtnOk := TButton.Create(Self); FBtnOk.Parent := Self;
  FBtnOk.SetBounds(ClientWidth - 200, ClientHeight - 40, 90, 28);
  FBtnOk.Caption := 'OK'; FBtnOk.Default := True; FBtnOk.ModalResult := mrOk;
  FBtnCancel := TButton.Create(Self); FBtnCancel.Parent := Self;
  FBtnCancel.SetBounds(ClientWidth - 100, ClientHeight - 40, 90, 28);
  FBtnCancel.Caption := 'Cancel'; FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;
end;

function TSemanticReplaceRuleEditDialog.GetRule: TSemanticReplaceRule;
var
  U: TArray<string>;
  S: string;
begin
  Result := Default(TSemanticReplaceRule);
  Result.Find := Trim(FEdtFind.Text);
  Result.Replace := Trim(FEdtReplace.Text);
  for S in string(FEdtUses.Text).Split([',']) do
    if Trim(S) <> '' then U := U + [Trim(S)];
  Result.UsesToAdd := U;
  Result.LocalVarName := Trim(FEdtVarName.Text);
  Result.LocalVarType := Trim(FEdtVarType.Text);
  Result.LocalVarValue := Trim(FEdtVarValue.Text);
  Result.ReplaceWhenLocalVar := Trim(FEdtVarReplace.Text);
end;

class function TSemanticReplaceRuleEditDialog.Edit(AOwner: TComponent;
  var ARule: TSemanticReplaceRule): Boolean;
var
  Dlg: TSemanticReplaceRuleEditDialog;
begin
  Dlg := TSemanticReplaceRuleEditDialog.CreateDialog(AOwner, ARule);
  try
    Result := Dlg.ShowModal = mrOk;
    if Result then ARule := Dlg.GetRule;
  finally
    Dlg.Free;
  end;
end;

{ ---------- TSemanticReplaceRulesListDialog ---------- }

constructor TSemanticReplaceRulesListDialog.CreateDialog(AOwner: TComponent;
  const ARules: TArray<TSemanticReplaceRule>);
begin
  inherited CreateNew(AOwner);
  Caption := 'Semantic replace rules';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 900;
  ClientHeight := 480;
  Constraints.MinWidth := 700;
  Constraints.MinHeight := 360;
  FRules := ARules;
  BuildLayout;
  RefreshList;
end;

procedure TSemanticReplaceRulesListDialog.BuildLayout;
var
  Col: TListColumn;
begin
  FListView := TListView.Create(Self); FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.ViewStyle := vsReport;
  FListView.ReadOnly := True;
  FListView.RowSelect := True;
  FListView.HideSelection := False;
  FListView.OnDblClick := DoDblClick;
  Col := FListView.Columns.Add; Col.Caption := 'Find'; Col.Width := 250;
  Col := FListView.Columns.Add; Col.Caption := 'Replace'; Col.Width := 320;
  Col := FListView.Columns.Add; Col.Caption := 'Uses'; Col.Width := 180;
  Col := FListView.Columns.Add; Col.Caption := 'Local var?'; Col.Width := 80;

  var Pnl: TPanel := TPanel.Create(Self);
  Pnl.Parent := Self;
  Pnl.Align := alBottom;
  Pnl.Height := 44;
  Pnl.BevelOuter := bvNone;

  FBtnAdd := TButton.Create(Self); FBtnAdd.Parent := Pnl;
  FBtnAdd.SetBounds(8, 8, 90, 28); FBtnAdd.Caption := 'Add...';
  FBtnAdd.OnClick := DoAdd;
  FBtnEdit := TButton.Create(Self); FBtnEdit.Parent := Pnl;
  FBtnEdit.SetBounds(108, 8, 90, 28); FBtnEdit.Caption := 'Edit...';
  FBtnEdit.OnClick := DoEdit;
  FBtnDelete := TButton.Create(Self); FBtnDelete.Parent := Pnl;
  FBtnDelete.SetBounds(208, 8, 90, 28); FBtnDelete.Caption := 'Delete';
  FBtnDelete.OnClick := DoDelete;

  FBtnOk := TButton.Create(Self); FBtnOk.Parent := Pnl;
  FBtnOk.SetBounds(ClientWidth - 200, 8, 90, 28);
  FBtnOk.Anchors := [akRight, akTop];
  FBtnOk.Caption := 'OK'; FBtnOk.Default := True; FBtnOk.ModalResult := mrOk;
  FBtnCancel := TButton.Create(Self); FBtnCancel.Parent := Pnl;
  FBtnCancel.SetBounds(ClientWidth - 100, 8, 90, 28);
  FBtnCancel.Anchors := [akRight, akTop];
  FBtnCancel.Caption := 'Cancel'; FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;
end;

procedure TSemanticReplaceRulesListDialog.RefreshList;
var
  I: Integer;
  R: TSemanticReplaceRule;
  Item: TListItem;
  HasVar: string;
begin
  FListView.Items.BeginUpdate;
  try
    FListView.Items.Clear;
    for I := 0 to High(FRules) do
    begin
      R := FRules[I];
      Item := FListView.Items.Add;
      Item.Caption := R.Find;
      Item.SubItems.Add(R.Replace);
      Item.SubItems.Add(string.Join(', ', R.UsesToAdd));
      if (R.LocalVarName <> '') and (R.LocalVarType <> '') and
         (R.LocalVarValue <> '') and (R.ReplaceWhenLocalVar <> '') then
        HasVar := 'yes (' + R.LocalVarName + ')'
      else
        HasVar := 'no';
      Item.SubItems.Add(HasVar);
    end;
  finally
    FListView.Items.EndUpdate;
  end;
end;

procedure TSemanticReplaceRulesListDialog.DoAdd(Sender: TObject);
var
  R: TSemanticReplaceRule;
begin
  R := Default(TSemanticReplaceRule);
  if TSemanticReplaceRuleEditDialog.Edit(Self, R) then
    if Trim(R.Find) <> '' then
    begin
      FRules := FRules + [R];
      RefreshList;
    end;
end;

procedure TSemanticReplaceRulesListDialog.DoEdit(Sender: TObject);
var
  Idx: Integer;
  R: TSemanticReplaceRule;
begin
  Idx := FListView.ItemIndex;
  if (Idx < 0) or (Idx > High(FRules)) then Exit;
  R := FRules[Idx];
  if TSemanticReplaceRuleEditDialog.Edit(Self, R) then
  begin
    FRules[Idx] := R;
    RefreshList;
    FListView.ItemIndex := Idx;
  end;
end;

procedure TSemanticReplaceRulesListDialog.DoDelete(Sender: TObject);
var
  Idx, I: Integer;
  NewRules: TArray<TSemanticReplaceRule>;
begin
  Idx := FListView.ItemIndex;
  if (Idx < 0) or (Idx > High(FRules)) then Exit;
  if MessageDlg('Delete this rule?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
  SetLength(NewRules, Length(FRules) - 1);
  for I := 0 to Idx - 1 do NewRules[I] := FRules[I];
  for I := Idx + 1 to High(FRules) do NewRules[I - 1] := FRules[I];
  FRules := NewRules;
  RefreshList;
end;

procedure TSemanticReplaceRulesListDialog.DoDblClick(Sender: TObject);
begin
  DoEdit(nil);
end;

class function TSemanticReplaceRulesListDialog.Edit(AOwner: TComponent;
  var ARules: TArray<TSemanticReplaceRule>): Boolean;
var
  Dlg: TSemanticReplaceRulesListDialog;
begin
  Dlg := TSemanticReplaceRulesListDialog.CreateDialog(AOwner, ARules);
  try
    Result := Dlg.ShowModal = mrOk;
    if Result then ARules := Dlg.Rules;
  finally
    Dlg.Free;
  end;
end;

{ ---------- TSemanticReplaceUnitsDialog ---------- }

constructor TSemanticReplaceUnitsDialog.CreateDialog(AOwner: TComponent;
  const AAllFiles: TArray<string>);
begin
  inherited CreateNew(AOwner);
  Caption := 'Apply semantic replacements - pick units';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 720;
  ClientHeight := 480;
  Constraints.MinWidth := 520;
  Constraints.MinHeight := 360;
  FAllFiles := AAllFiles;
  BuildLayout;
  RefreshList;
end;

procedure TSemanticReplaceUnitsDialog.BuildLayout;
var
  PnlTop, PnlBottom: TPanel;
  Lbl: TLabel;
begin
  PnlTop := TPanel.Create(Self); PnlTop.Parent := Self;
  PnlTop.Align := alTop;
  PnlTop.Height := 56; PnlTop.BevelOuter := bvNone;
  Lbl := TLabel.Create(Self); Lbl.Parent := PnlTop;
  Lbl.SetBounds(8, 10, 80, 16); Lbl.Caption := 'Filter:';
  FEdtFilter := TEdit.Create(Self); FEdtFilter.Parent := PnlTop;
  FEdtFilter.SetBounds(56, 6, ClientWidth - 64, 24);
  FEdtFilter.Anchors := [akLeft, akTop, akRight];
  FEdtFilter.OnChange := DoFilterChange;
  FBtnAll := TButton.Create(Self); FBtnAll.Parent := PnlTop;
  FBtnAll.SetBounds(8, 32, 80, 22); FBtnAll.Caption := 'Select all';
  FBtnAll.OnClick := DoAll;
  FBtnNone := TButton.Create(Self); FBtnNone.Parent := PnlTop;
  FBtnNone.SetBounds(94, 32, 80, 22); FBtnNone.Caption := 'Clear';
  FBtnNone.OnClick := DoNone;

  PnlBottom := TPanel.Create(Self); PnlBottom.Parent := Self;
  PnlBottom.Align := alBottom;
  PnlBottom.Height := 44; PnlBottom.BevelOuter := bvNone;
  FBtnOk := TButton.Create(Self); FBtnOk.Parent := PnlBottom;
  FBtnOk.SetBounds(ClientWidth - 200, 8, 90, 28);
  FBtnOk.Anchors := [akRight, akTop];
  FBtnOk.Caption := 'OK'; FBtnOk.Default := True; FBtnOk.ModalResult := mrOk;
  FBtnCancel := TButton.Create(Self); FBtnCancel.Parent := PnlBottom;
  FBtnCancel.SetBounds(ClientWidth - 100, 8, 90, 28);
  FBtnCancel.Anchors := [akRight, akTop];
  FBtnCancel.Caption := 'Cancel'; FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;

  FCheckList := TCheckListBox.Create(Self); FCheckList.Parent := Self;
  FCheckList.Align := alClient;
end;

procedure TSemanticReplaceUnitsDialog.RefreshList;
var
  Filter: string;
  F: string;
begin
  Filter := UpperCase(Trim(FEdtFilter.Text));
  FCheckList.Items.BeginUpdate;
  try
    FCheckList.Items.Clear;
    for F in FAllFiles do
      if (Filter = '') or (Pos(Filter, UpperCase(ExtractFileName(F))) > 0) then
        FCheckList.Items.AddObject(ExtractFileName(F) + '  -  ' + F, TObject(0));
  finally
    FCheckList.Items.EndUpdate;
  end;
end;

procedure TSemanticReplaceUnitsDialog.DoFilterChange(Sender: TObject);
begin
  RefreshList;
end;

procedure TSemanticReplaceUnitsDialog.DoAll(Sender: TObject);
var I: Integer;
begin
  for I := 0 to FCheckList.Items.Count - 1 do FCheckList.Checked[I] := True;
end;

procedure TSemanticReplaceUnitsDialog.DoNone(Sender: TObject);
var I: Integer;
begin
  for I := 0 to FCheckList.Items.Count - 1 do FCheckList.Checked[I] := False;
end;

function TSemanticReplaceUnitsDialog.SelectedFiles: TArray<string>;
var
  I, P: Integer;
  S: string;
  Res: TArray<string>;
begin
  for I := 0 to FCheckList.Items.Count - 1 do
    if FCheckList.Checked[I] then
    begin
      S := FCheckList.Items[I];
      P := Pos('-  ', S);
      if P > 0 then Res := Res + [Trim(Copy(S, P + 3, MaxInt))];
    end;
  Result := Res;
end;

class function TSemanticReplaceUnitsDialog.Choose(AOwner: TComponent;
  const AAllFiles: TArray<string>; out AChosen: TArray<string>): Boolean;
var
  Dlg: TSemanticReplaceUnitsDialog;
begin
  Dlg := TSemanticReplaceUnitsDialog.CreateDialog(AOwner, AAllFiles);
  try
    Result := Dlg.ShowModal = mrOk;
    if Result then AChosen := Dlg.SelectedFiles;
  finally
    Dlg.Free;
  end;
end;

{ ---------- TSemanticReplacePreviewDialog ---------- }

constructor TSemanticReplacePreviewDialog.CreateDialog(AOwner: TComponent;
  const ASummary, APreview: string);
begin
  inherited CreateNew(AOwner);
  Caption := 'Semantic replace - preview';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 1000;
  ClientHeight := 640;
  Constraints.MinWidth := 700;
  Constraints.MinHeight := 420;
  BuildLayout;
  FLblSummary.Caption := ASummary;
  FMemo.Lines.Text := APreview;
end;

procedure TSemanticReplacePreviewDialog.BuildLayout;
var
  PnlTop, PnlBottom: TPanel;
begin
  PnlTop := TPanel.Create(Self); PnlTop.Parent := Self;
  PnlTop.Align := alTop; PnlTop.Height := 36; PnlTop.BevelOuter := bvNone;
  FLblSummary := TLabel.Create(Self); FLblSummary.Parent := PnlTop;
  FLblSummary.SetBounds(8, 10, ClientWidth - 16, 18);
  FLblSummary.Anchors := [akLeft, akTop, akRight];

  PnlBottom := TPanel.Create(Self); PnlBottom.Parent := Self;
  PnlBottom.Align := alBottom; PnlBottom.Height := 44; PnlBottom.BevelOuter := bvNone;
  FBtnApply := TButton.Create(Self); FBtnApply.Parent := PnlBottom;
  FBtnApply.SetBounds(ClientWidth - 220, 8, 110, 28);
  FBtnApply.Anchors := [akRight, akTop];
  FBtnApply.Caption := 'Apply'; FBtnApply.Default := True;
  FBtnApply.ModalResult := mrOk;
  FBtnCancel := TButton.Create(Self); FBtnCancel.Parent := PnlBottom;
  FBtnCancel.SetBounds(ClientWidth - 100, 8, 90, 28);
  FBtnCancel.Anchors := [akRight, akTop];
  FBtnCancel.Caption := 'Cancel'; FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;

  FMemo := TMemo.Create(Self); FMemo.Parent := Self;
  FMemo.Align := alClient;
  FMemo.ReadOnly := True;
  FMemo.ScrollBars := ssBoth;
  FMemo.WordWrap := False;
  FMemo.Font.Name := 'Consolas';
  FMemo.Font.Size := 10;
end;

class function TSemanticReplacePreviewDialog.Confirm(AOwner: TComponent;
  const ASummary, APreview: string): Boolean;
var
  Dlg: TSemanticReplacePreviewDialog;
begin
  Dlg := TSemanticReplacePreviewDialog.CreateDialog(AOwner, ASummary, APreview);
  try
    Result := Dlg.ShowModal = mrOk;
  finally
    Dlg.Free;
  end;
end;

end.
