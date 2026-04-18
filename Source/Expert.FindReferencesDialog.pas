(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindReferencesDialog;

interface

uses
  System.SysUtils, System.Classes, System.UITypes, System.Generics.Collections, Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls;

type
  TFindReferenceItem = record
    FilePath: string;
    Line: Integer;     // 0-based (LSP)
    Col: Integer;      // 0-based
    Length: Integer;   // Length of the identifier
    Preview: string;   // Line containing the match
  end;

  TFindReferenceItems = TArray<TFindReferenceItem>;

  /// <summary>Modal dialog with a ListView of all occurrences of an identifier.
  ///  Double-click = jump to location (dialog stays open).
  ///  ENTER on a line = jump to location and close dialog.
  ///  ESC / Close button = close dialog.</summary>
  TFindReferencesDialog = class(TForm)
  private
    FStatusLabel: TLabel;
    FProgress: TProgressBar;
    FListView: TListView;
    FBtnGoto: TButton;
    FBtnClose: TButton;
    FItems: TFindReferenceItems;
    FOnGotoLocation: TProc<TFindReferenceItem>;

    procedure CreateControls;
    procedure DoListDblClick(Sender: TObject);
    procedure DoBtnGotoClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function CommonPathPrefix(const AItems: TFindReferenceItems): string;
    procedure GotoSelected;
  public
    constructor CreateDialog(AOwner: TComponent; const AIdentifier: string;
      const ATitlePrefix: string = 'References'); reintroduce;

    /// <summary>Sets the matches and fills the ListView. Replaced previous items.</summary>
    procedure SetItems(const AItems: TFindReferenceItems);

    procedure SetStatus(const AText: string);
    procedure SetProgress(ACurrent, ATotal: Integer);

    property OnGotoLocation: TProc<TFindReferenceItem> read FOnGotoLocation write FOnGotoLocation;
  end;

implementation

uses
  System.IOUtils;

{ TFindReferencesDialog }

constructor TFindReferencesDialog.CreateDialog(AOwner: TComponent; const AIdentifier: string; const ATitlePrefix: string);
begin
  inherited CreateNew(AOwner);
  Caption := ATitlePrefix + ': ' + AIdentifier;
  Position := poOwnerFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  Width := 800;
  Height := 500;
  Constraints.MinWidth := 480;
  Constraints.MinHeight := 280;
  KeyPreview := True;
  OnKeyDown := DoFormKeyDown;

  CreateControls;
end;

procedure TFindReferencesDialog.CreateControls;
var
  Col: TListColumn;
begin
  // Status line at top
  FStatusLabel := TLabel.Create(Self);
  FStatusLabel.Parent := Self;
  FStatusLabel.Align := alTop;
  FStatusLabel.AlignWithMargins := True;
  FStatusLabel.Margins.SetBounds(8, 8, 8, 4);
  FStatusLabel.Caption := 'Searching...';

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

  // Button panel at bottom
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

  // ListView in the middle
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
  FListView.OnDblClick := DoListDblClick;
  FListView.OnKeyDown := DoListKeyDown;

  Col := FListView.Columns.Add;
  Col.Caption := 'File';
  Col.Width := 220;

  Col := FListView.Columns.Add;
  Col.Caption := 'Line';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Column';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Preview';
  Col.Width := 420;
end;

function TFindReferencesDialog.CommonPathPrefix(
  const AItems: TFindReferenceItems): string;
var
  P: string;
  I, J, MaxLen: Integer;
begin
  Result := '';
  if Length(AItems) = 0 then Exit;
  Result := ExtractFilePath(AItems[0].FilePath);
  for I := 1 to High(AItems) do
  begin
    P := ExtractFilePath(AItems[I].FilePath);
    MaxLen := Length(Result);
    if Length(P) < MaxLen then MaxLen := Length(P);
    J := 1;
    while (J <= MaxLen) and (UpCase(Result[J]) = UpCase(P[J])) do
      Inc(J);
    Result := Copy(Result, 1, J - 1);
    if Result = '' then Exit;
  end;
end;

procedure TFindReferencesDialog.SetItems(const AItems: TFindReferenceItems);
var
  Item: TFindReferenceItem;
  LI: TListItem;
  Prefix, DisplayPath: string;
begin
  FItems := AItems;
  Prefix := CommonPathPrefix(AItems);

  FListView.Items.BeginUpdate;
  try
    FListView.Clear;
    for Item in AItems do
    begin
      LI := FListView.Items.Add;
      DisplayPath := Item.FilePath;
      if (Prefix <> '') and DisplayPath.StartsWith(Prefix, True) then
        DisplayPath := Copy(DisplayPath, Length(Prefix) + 1, MaxInt);
      LI.Caption := DisplayPath;
      LI.SubItems.Add(IntToStr(Item.Line + 1));
      LI.SubItems.Add(IntToStr(Item.Col + 1));
      LI.SubItems.Add(Item.Preview);
    end;
    if FListView.Items.Count > 0 then
    begin
      FListView.Items[0].Selected := True;
      FListView.Items[0].Focused := True;
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  FBtnGoto.Enabled := Length(AItems) > 0;
end;

procedure TFindReferencesDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TFindReferencesDialog.SetProgress(ACurrent, ATotal: Integer);
begin
  if ATotal <= 0 then
  begin
    FProgress.Position := 0;
    Exit;
  end;
  FProgress.Max := ATotal;
  if ACurrent > ATotal then
    ACurrent := ATotal;
  FProgress.Position := ACurrent;
end;

procedure TFindReferencesDialog.GotoSelected;
var
  Idx: Integer;
begin
  if not Assigned(FListView.Selected) then Exit;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if Assigned(FOnGotoLocation) then
    FOnGotoLocation(FItems[Idx]);
end;

procedure TFindReferencesDialog.DoListDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TFindReferencesDialog.DoBtnGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TFindReferencesDialog.DoBtnCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TFindReferencesDialog.DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    ModalResult := mrCancel;
    Key := 0;
  end;
end;

procedure TFindReferencesDialog.DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    GotoSelected;
    ModalResult := mrOk;
    Key := 0;
  end;
end;

end.
