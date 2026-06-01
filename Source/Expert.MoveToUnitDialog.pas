(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MoveToUnitDialog;

{
  Modal dialog for "Move identifier to other unit".
  - Lists every project .pas unit (except the source unit), filterable.
  - Shows a preview memo describing the planned edits.
  - Returns the chosen target file path.
}

interface

uses
  System.SysUtils, System.Classes,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls,
  Expert.MoveToUnit;

type
  TMoveToUnitDialog = class(TForm)
  private
    FLblIdent: TLabel;
    FEdtIdent: TEdit;
    FLblFilter: TLabel;
    FEdtFilter: TEdit;
    FLstUnits: TListBox;
    FMemoPreview: TMemo;
    FBtnOk: TButton;
    FBtnCancel: TButton;
    FSymbol: string;
    FSourceFile: string;
    FAllFiles: TArray<string>;
    procedure DoFilterChange(Sender: TObject);
    procedure DoUnitChange(Sender: TObject);
    procedure RefreshPreview;
    procedure RefreshList;
  public
    constructor CreateDialog(AOwner: TComponent;
      const ASymbol, ASourceFile: string;
      const AProjectFiles: TArray<string>);
    function SelectedFile: string;

    class function Choose(AOwner: TComponent;
      const ASymbol, ASourceFile: string;
      const AProjectFiles: TArray<string>;
      out ATargetFile: string): Boolean;
  end;

implementation

uses
  System.UITypes, Expert.DialogHelper, Expert.LspManager;

{ TMoveToUnitDialog }

constructor TMoveToUnitDialog.CreateDialog(AOwner: TComponent;
  const ASymbol, ASourceFile: string;
  const AProjectFiles: TArray<string>);
var
  F: string;
  Filtered: TArray<string>;
  Count: Integer;
begin
  inherited CreateNew(AOwner);
  Caption := 'Move to unit...';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 720;
  ClientHeight := 480;
  Constraints.MinWidth := 520;
  Constraints.MinHeight := 360;
  KeyPreview := True;

  FSymbol := ASymbol;
  FSourceFile := ASourceFile;

  // Filter the project file list: only .pas files, exclude the source.
  Count := 0;
  SetLength(Filtered, Length(AProjectFiles));
  for F in AProjectFiles do
  begin
    if SameText(F, ASourceFile) then Continue;
    if not SameText(ExtractFileExt(F), '.pas') then Continue;
    Filtered[Count] := F;
    Inc(Count);
  end;
  SetLength(Filtered, Count);
  FAllFiles := Filtered;

  FLblIdent := TLabel.Create(Self);
  FLblIdent.Parent := Self;
  FLblIdent.Left := 12;
  FLblIdent.Top := 8;
  FLblIdent.Caption := 'Identifier to move:';

  FEdtIdent := TEdit.Create(Self);
  FEdtIdent.Parent := Self;
  FEdtIdent.Left := 12;
  FEdtIdent.Top := 26;
  FEdtIdent.Width := ClientWidth - 24;
  FEdtIdent.Anchors := [akLeft, akTop, akRight];
  FEdtIdent.Text := ASymbol;
  FEdtIdent.ReadOnly := True;
  FEdtIdent.Color := $F0F0F0;

  FLblFilter := TLabel.Create(Self);
  FLblFilter.Parent := Self;
  FLblFilter.Left := 12;
  FLblFilter.Top := 58;
  FLblFilter.Caption := 'Target unit (filter):';

  FEdtFilter := TEdit.Create(Self);
  FEdtFilter.Parent := Self;
  FEdtFilter.Left := 12;
  FEdtFilter.Top := 76;
  FEdtFilter.Width := ClientWidth - 24;
  FEdtFilter.Anchors := [akLeft, akTop, akRight];
  FEdtFilter.OnChange := DoFilterChange;

  FLstUnits := TListBox.Create(Self);
  FLstUnits.Parent := Self;
  FLstUnits.Left := 12;
  FLstUnits.Top := 108;
  FLstUnits.Width := 280;
  FLstUnits.Height := ClientHeight - 108 - 48;
  FLstUnits.Anchors := [akLeft, akTop, akBottom];
  FLstUnits.OnClick := DoUnitChange;

  FMemoPreview := TMemo.Create(Self);
  FMemoPreview.Parent := Self;
  FMemoPreview.Left := 300;
  FMemoPreview.Top := 108;
  FMemoPreview.Width := ClientWidth - 300 - 12;
  FMemoPreview.Height := FLstUnits.Height;
  FMemoPreview.Anchors := [akLeft, akTop, akRight, akBottom];
  FMemoPreview.ReadOnly := True;
  FMemoPreview.ScrollBars := ssBoth;
  FMemoPreview.Font.Name := 'Consolas';
  FMemoPreview.Font.Size := 9;
  FMemoPreview.WordWrap := False;

  FBtnCancel := TButton.Create(Self);
  FBtnCancel.Parent := Self;
  FBtnCancel.Caption := 'Cancel';
  FBtnCancel.Width := 90;
  FBtnCancel.Height := 28;
  FBtnCancel.Top := ClientHeight - 38;
  FBtnCancel.Left := ClientWidth - 90 - 12;
  FBtnCancel.Anchors := [akRight, akBottom];
  FBtnCancel.ModalResult := mrCancel;
  FBtnCancel.Cancel := True;

  FBtnOk := TButton.Create(Self);
  FBtnOk.Parent := Self;
  FBtnOk.Caption := 'Move';
  FBtnOk.Width := 90;
  FBtnOk.Height := 28;
  FBtnOk.Top := ClientHeight - 38;
  FBtnOk.Left := FBtnCancel.Left - 96;
  FBtnOk.Anchors := [akRight, akBottom];
  FBtnOk.ModalResult := mrOk;
  FBtnOk.Default := True;
  FBtnOk.Enabled := False;

  RefreshList;

  PrepareDialog(Self, AOwner);
end;

procedure TMoveToUnitDialog.RefreshList;
var
  F, U, Flt: string;
begin
  FLstUnits.Items.BeginUpdate;
  try
    FLstUnits.Items.Clear;
    Flt := AnsiUpperCase(Trim(FEdtFilter.Text));
    for F in FAllFiles do
    begin
      U := ChangeFileExt(ExtractFileName(F), '');
      if (Flt = '') or (Pos(Flt, AnsiUpperCase(U)) > 0) then
        FLstUnits.Items.Add(U);
    end;
  finally
    FLstUnits.Items.EndUpdate;
  end;
  FBtnOk.Enabled := False;
  FMemoPreview.Clear;
end;

procedure TMoveToUnitDialog.DoFilterChange(Sender: TObject);
begin
  RefreshList;
end;

procedure TMoveToUnitDialog.DoUnitChange(Sender: TObject);
begin
  FBtnOk.Enabled := FLstUnits.ItemIndex >= 0;
  RefreshPreview;
end;

function TMoveToUnitDialog.SelectedFile: string;
var
  Sel, F: string;
begin
  Result := '';
  if FLstUnits.ItemIndex < 0 then Exit;
  Sel := FLstUnits.Items[FLstUnits.ItemIndex];
  for F in FAllFiles do
    if SameText(ChangeFileExt(ExtractFileName(F), ''), Sel) then
      Exit(F);
end;

procedure TMoveToUnitDialog.RefreshPreview;
var
  Plan: TMovePlan;
  Sel: string;
  E: TMoveEdit;
  SB: TStringBuilder;
begin
  Sel := SelectedFile;
  FMemoPreview.Clear;
  if Sel = '' then Exit;
  SB := TStringBuilder.Create;
  try
    SB.Append('Symbol: ').Append(FSymbol).AppendLine;
    SB.Append('Source: ').Append(ExtractFileName(FSourceFile)).AppendLine;
    SB.Append('Target: ').Append(ExtractFileName(Sel)).AppendLine;
    SB.AppendLine;
    if TLspMoveToUnit.BuildPlan(FSymbol, FSourceFile, Sel, Plan) then
    begin
      SB.Append('Planned edits:').AppendLine;
      for E in Plan.Edits do
        SB.Append('  - ').Append(ExtractFileName(E.FilePath))
          .Append(': ').Append(E.Description).AppendLine;
      if Length(Plan.Consumers) = 0 then
      begin
        SB.AppendLine;
        SB.Append('No consumer files reference this symbol.').AppendLine;
      end;
      FBtnOk.Enabled := True;
    end
    else
    begin
      SB.AppendLine;
      SB.Append('Cannot move: ').Append(Plan.ProblemDetail).AppendLine;
      FBtnOk.Enabled := False;
    end;
    FMemoPreview.Lines.Text := SB.ToString;
  finally
    SB.Free;
  end;
end;

class function TMoveToUnitDialog.Choose(AOwner: TComponent;
  const ASymbol, ASourceFile: string;
  const AProjectFiles: TArray<string>;
  out ATargetFile: string): Boolean;
var
  Dlg: TMoveToUnitDialog;
begin
  Result := False;
  ATargetFile := '';
  Dlg := TMoveToUnitDialog.CreateDialog(AOwner, ASymbol, ASourceFile, AProjectFiles);
  try
    TLspManager.Instance.ApplyStatusToCaption(Dlg);
    if Dlg.ShowModal = mrOk then
    begin
      ATargetFile := Dlg.SelectedFile;
      Result := ATargetFile <> '';
    end;
  finally
    Dlg.Free;
  end;
end;

initialization
  RegisterDialogClass(TMoveToUnitDialog);

end.
