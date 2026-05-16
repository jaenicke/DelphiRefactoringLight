(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SignatureCheckDialog;

{
  Interactive dialog for the "Align method signature" feature. Shows all
  found declarations / implementations for the method under the cursor
  side-by-side, marks the ones that differ from the majority, and lets
  the user jump to each location by double-click / Enter / "Go to".
}

interface

uses
  System.SysUtils, System.Classes, System.UITypes, System.Generics.Collections,
  Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Graphics,
  Expert.SignatureCheck;

type
  TSignatureCheckDialog = class(TForm)
  private
    FStatusLabel: TLabel;
    FListView: TListView;
    FBtnGoto: TButton;
    FBtnClose: TButton;
    FEntries: TSignatureEntries;
    FOnGotoLocation: TProc<TSignatureEntry>;
    FReferenceNormalized: string;

    procedure CreateControls;
    procedure DoListDblClick(Sender: TObject);
    procedure DoBtnGotoClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoListCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure GotoSelected;
    function PickReference(const AEntries: TSignatureEntries): string;
  public
    constructor CreateDialog(AOwner: TComponent; const AMethodName: string); reintroduce;

    procedure SetEntries(const AEntries: TSignatureEntries);
    procedure SetStatus(const AText: string);

    property OnGotoLocation: TProc<TSignatureEntry> read FOnGotoLocation write FOnGotoLocation;
  end;

implementation

uses
  Winapi.UxTheme, Expert.IdeThemes;

{ TSignatureCheckDialog }

constructor TSignatureCheckDialog.CreateDialog(AOwner: TComponent;
  const AMethodName: string);
begin
  inherited CreateNew(AOwner);
  Caption := 'Align method signature: ' + AMethodName;
  Position := poOwnerFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  Width := 900;
  Height := 400;
  Constraints.MinWidth := 520;
  Constraints.MinHeight := 240;
  KeyPreview := True;
  OnKeyDown := DoFormKeyDown;

  CreateControls;
  Expert.IdeThemes.EnableThemes(Self);
end;

procedure TSignatureCheckDialog.CreateControls;
var
  Col: TListColumn;
begin
  FStatusLabel := TLabel.Create(Self);
  FStatusLabel.Parent := Self;
  FStatusLabel.Align := alTop;
  FStatusLabel.AlignWithMargins := True;
  FStatusLabel.Margins.SetBounds(8, 8, 8, 4);
  FStatusLabel.Caption := 'Collecting signatures...';

  var BtnPanel := TPanel.Create(Self);
  BtnPanel.Parent := Self;
  BtnPanel.Align := alBottom;
  BtnPanel.Height := 40;
  BtnPanel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := BtnPanel;
  FBtnClose.Caption := 'Close';
  FBtnClose.Width := 100;
  FBtnClose.Height := 28;
  FBtnClose.Top := 6;
  FBtnClose.Anchors := [akTop, akRight];
  FBtnClose.Left := BtnPanel.Width - FBtnClose.Width - 8;
  FBtnClose.OnClick := DoBtnCloseClick;
  FBtnClose.Cancel := True;

  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := BtnPanel;
  FBtnGoto.Caption := 'Go To';
  FBtnGoto.Width := 100;
  FBtnGoto.Height := 28;
  FBtnGoto.Top := 6;
  FBtnGoto.Anchors := [akTop, akRight];
  FBtnGoto.Left := FBtnClose.Left - FBtnGoto.Width - 6;
  FBtnGoto.OnClick := DoBtnGotoClick;
  FBtnGoto.Default := True;
  FBtnGoto.Enabled := False;

  FListView := TListView.Create(Self);
  FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.AlignWithMargins := True;
  FListView.Margins.SetBounds(8, 4, 8, 4);
  FListView.ViewStyle := vsReport;
  FListView.RowSelect := True;
  FListView.ReadOnly := True;
  FListView.HideSelection := False;
  FListView.GridLines := True;
  FListView.HotTrack := False;
  FListView.HandleNeeded;
  SetWindowTheme(FListView.Handle, '', '');
  FListView.OnDblClick := DoListDblClick;
  FListView.OnKeyDown := DoListKeyDown;
  FListView.OnCustomDrawItem := DoListCustomDrawItem;

  Col := FListView.Columns.Add;
  Col.Caption := 'Role';
  Col.Width := 120;

  Col := FListView.Columns.Add;
  Col.Caption := 'Container';
  Col.Width := 140;

  Col := FListView.Columns.Add;
  Col.Caption := 'File';
  Col.Width := 180;

  Col := FListView.Columns.Add;
  Col.Caption := 'Line';
  Col.Width := 50;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Match';
  Col.Width := 60;

  Col := FListView.Columns.Add;
  Col.Caption := 'Signature';
  Col.Width := 320;
end;

function TSignatureCheckDialog.PickReference(const AEntries: TSignatureEntries): string;
var
  Counts: TDictionary<string, Integer>;
  Best: string;
  BestCount: Integer;
  Pair: TPair<string, Integer>;
begin
  Result := '';
  if Length(AEntries) = 0 then Exit;
  Counts := TDictionary<string, Integer>.Create;
  try
    for var E in AEntries do
      if Counts.ContainsKey(E.Normalized) then
        Counts[E.Normalized] := Counts[E.Normalized] + 1
      else
        Counts.Add(E.Normalized, 1);
    Best := AEntries[0].Normalized;
    BestCount := 0;
    for Pair in Counts do
      if Pair.Value > BestCount then
      begin
        Best := Pair.Key;
        BestCount := Pair.Value;
      end;
    Result := Best;
  finally
    Counts.Free;
  end;
end;

procedure TSignatureCheckDialog.SetEntries(const AEntries: TSignatureEntries);
var
  LI: TListItem;
  IsMatch: Boolean;
begin
  FEntries := AEntries;
  FReferenceNormalized := PickReference(AEntries);

  FListView.Items.BeginUpdate;
  try
    FListView.Clear;
    for var E in AEntries do
    begin
      LI := FListView.Items.Add;
      LI.Caption := TSignatureChecker.RoleToString(E.Role);
      LI.SubItems.Add(E.Container);
      LI.SubItems.Add(ExtractFileName(E.FilePath));
      LI.SubItems.Add(IntToStr(E.Line + 1));
      IsMatch := E.Normalized = FReferenceNormalized;
      if IsMatch then
        LI.SubItems.Add('yes')
      else
        LI.SubItems.Add('NO');
      LI.SubItems.Add(E.RawSignature);
    end;
    if FListView.Items.Count > 0 then
    begin
      FListView.Items[0].Selected := True;
      FListView.Items[0].Focused := True;
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  FBtnGoto.Enabled := Length(AEntries) > 0;
end;

procedure TSignatureCheckDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TSignatureCheckDialog.DoListCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  if (Item.Index >= 0) and (Item.Index <= High(FEntries)) then
  begin
    if FEntries[Item.Index].Normalized <> FReferenceNormalized then
      Sender.Canvas.Brush.Color := RGB(255, 230, 230)  // light red
    else
      Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
  end;
  DefaultDraw := True;
end;

procedure TSignatureCheckDialog.GotoSelected;
var
  Idx: Integer;
begin
  if not Assigned(FListView.Selected) then Exit;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx > High(FEntries)) then Exit;
  if Assigned(FOnGotoLocation) then
    FOnGotoLocation(FEntries[Idx]);
end;

procedure TSignatureCheckDialog.DoListDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TSignatureCheckDialog.DoBtnGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TSignatureCheckDialog.DoBtnCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TSignatureCheckDialog.DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    ModalResult := mrCancel;
    Key := 0;
  end;
end;

procedure TSignatureCheckDialog.DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    GotoSelected;
    ModalResult := mrOk;
    Key := 0;
  end;
end;

end.
