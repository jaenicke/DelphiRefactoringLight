(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
    FBtnAlign: TButton;
    FBtnClose: TButton;
    FOnAlign: TFunc<TSignatureEntry, TSignatureEntry, string>;
    FEntries: TSignatureEntries;
    FOnGotoLocation: TProc<TSignatureEntry>;
    FReferenceNormalized: string;
    FAllowFree: Boolean;
    FCloseRequested: Boolean;

    procedure CreateControls;
    procedure DoListDblClick(Sender: TObject);
    procedure DoBtnGotoClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoListCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure GotoSelected;
    function PickReference(const AEntries: TSignatureEntries): string;
    procedure DoBtnAlignClick(Sender: TObject);
    procedure DoListSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    function SelectedIndex: Integer;
    function ReferenceEntry(out AEntry: TSignatureEntry): Boolean;
    function AlignBlocker(AIdx: Integer): string;
    procedure UpdateAlignButton;
  public
    constructor CreateDialog(AOwner: TComponent; const AMethodName: string); reintroduce;

    procedure SetEntries(const AEntries: TSignatureEntries);
    procedure SetStatus(const AText: string);

    /// <summary>Hands the dialog ownership over to itself: closing
    ///  (X / Close / Esc) frees the dialog. Before SetClosable is
    ///  called, close requests are deferred so the still-running
    ///  search loop doesn't operate on a freed dialog.</summary>
    procedure SetClosable;

    property OnGotoLocation: TProc<TSignatureEntry> read FOnGotoLocation write FOnGotoLocation;
    /// <summary>Aligns the first entry with the second (the reference).
    ///  Returns '' on success, else the reason it was refused.</summary>
    property OnAlign: TFunc<TSignatureEntry, TSignatureEntry, string> read FOnAlign write FOnAlign;
  end;

implementation

uses
  Winapi.UxTheme, Expert.DialogHelper, Expert.IdeThemes, Expert.ListViewSort;

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
  OnClose := DoFormClose;

  CreateControls;
  Expert.IdeThemes.EnableThemes(Self);

  PrepareDialog(Self, AOwner);
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

  // Aligns the selected divergent row with the majority signature - the
  // fix for what the list shows (it used to be diagnostic only).
  FBtnAlign := TButton.Create(Self);
  FBtnAlign.Parent := BtnPanel;
  FBtnAlign.Caption := 'Align';
  FBtnAlign.Hint := 'Rewrite the selected declaration / implementation to the ' +
    'majority signature (implementations keep their parameter names)';
  FBtnAlign.ShowHint := True;
  FBtnAlign.Width := 100;
  FBtnAlign.Height := 28;
  FBtnAlign.Top := 6;
  FBtnAlign.Anchors := [akTop, akRight];
  FBtnAlign.Left := FBtnGoto.Left - FBtnAlign.Width - 6;
  FBtnAlign.OnClick := DoBtnAlignClick;
  FBtnAlign.Enabled := False;

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
  FListView.OnSelectItem := DoListSelectItem;

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

  EnableListViewSorting(FListView);
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
    for var I := 0 to High(AEntries) do
    begin
      var E := AEntries[I];
      LI := FListView.Items.Add;
      LI.Data := Pointer(NativeInt(I));  // entry index; survives sorting
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
  UpdateAlignButton;
end;

function TSignatureCheckDialog.SelectedIndex: Integer;
begin
  Result := -1;
  if Assigned(FListView.Selected) then
    Result := NativeInt(FListView.Selected.Data);
  if (Result < 0) or (Result > High(FEntries)) then Result := -1;
end;

// The entry the others are aligned to: one carrying the majority signature,
// a DECLARATION preferred (its text has no "TClass." qualifier and its
// parameter names are the documented ones).
function TSignatureCheckDialog.ReferenceEntry(out AEntry: TSignatureEntry): Boolean;
begin
  Result := False;
  for var Pass := 0 to 1 do
    for var E in FEntries do
      if (E.Normalized = FReferenceNormalized)
        and ((Pass = 1) or (E.Role in [srInterfaceDecl, srClassDecl])) then
      begin
        AEntry := E;
        Exit(True);
      end;
end;

// '' when the row at AIdx can be aligned, else why not.
function TSignatureCheckDialog.AlignBlocker(AIdx: Integer): string;
var
  Ref: TSignatureEntry;
begin
  if AIdx < 0 then Exit('select a row');
  if FEntries[AIdx].Normalized = FReferenceNormalized then
    Exit('this row already matches');
  if not ReferenceEntry(Ref) then Exit('no reference signature');
  if FEntries[AIdx].Role = srImplementation then
    // the implementation is aligned with the class declaration of ITS unit
    // - which must be the right one first
    for var E in FEntries do
      if (E.Role = srClassDecl) and SameText(E.FilePath, FEntries[AIdx].FilePath)
        and SameText(E.Container, FEntries[AIdx].Container)
        and (E.Normalized <> FReferenceNormalized) then
        Exit('align the class declaration first');
  Result := '';
end;

procedure TSignatureCheckDialog.UpdateAlignButton;
begin
  FBtnAlign.Enabled := Assigned(FOnAlign) and (AlignBlocker(SelectedIndex) = '');
end;

procedure TSignatureCheckDialog.DoListSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  UpdateAlignButton;
end;

procedure TSignatureCheckDialog.DoBtnAlignClick(Sender: TObject);
var
  Idx: Integer;
  Ref: TSignatureEntry;
  Why: string;
begin
  Idx := SelectedIndex;
  Why := AlignBlocker(Idx);
  if Why = '' then
    if not ReferenceEntry(Ref) then Why := 'no reference signature';
  if Why = '' then
    Why := FOnAlign(FEntries[Idx], Ref);
  if Why <> '' then
  begin
    SetStatus('Not aligned: ' + Why);
    Exit;
  end;
  // The buffer is changed (not saved); the row now carries the reference.
  FEntries[Idx].Normalized := FReferenceNormalized;
  if FListView.Selected <> nil then
    FListView.Selected.SubItems[3] := 'aligned';
  FListView.Invalidate;
  SetStatus(Format('%s in %s, line %d aligned (not saved - Ctrl+Z undoes it).',
    [TSignatureChecker.RoleToString(FEntries[Idx].Role),
     ExtractFileName(FEntries[Idx].FilePath), FEntries[Idx].Line + 1]));
  UpdateAlignButton;
end;

procedure TSignatureCheckDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TSignatureCheckDialog.DoListCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
var
  Idx: NativeInt;
begin
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
  Sender.Canvas.Font.Color := GetThemedColor(clWindowText);
  Idx := NativeInt(Item.Data);
  if (Idx >= 0) and (Idx <= High(FEntries)) then
  begin
    if FEntries[Idx].Normalized <> FReferenceNormalized then
    begin
      Sender.Canvas.Brush.Color := RGB(255, 230, 230);  // light red
      // explicit black: the light-red brush stays light in dark mode, a
      // themed (light) font color would be unreadable on it
      Sender.Canvas.Font.Color := clBlack;
    end;
  end;
  DefaultDraw := True;
end;

procedure TSignatureCheckDialog.GotoSelected;
var
  Idx: NativeInt;
begin
  if not Assigned(FListView.Selected) then Exit;
  Idx := NativeInt(FListView.Selected.Data);
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
  // Defensive: hide first so the user sees immediate feedback even
  // while the wizard's synchronous search loop is still pumping
  // messages. The actual free happens via DoFormClose when FAllowFree
  // is True (SetClosable already called).
  FCloseRequested := True;
  Hide;
  if FAllowFree then Release;
end;

procedure TSignatureCheckDialog.DoFormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not FAllowFree then
  begin
    FCloseRequested := True;
    Hide;
    Action := caNone;
    Exit;
  end;
  Action := caFree;
end;

procedure TSignatureCheckDialog.SetClosable;
begin
  FAllowFree := True;
  if FCloseRequested then
    Close;
end;

procedure TSignatureCheckDialog.DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Close;
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

initialization
  RegisterDialogClass(TSignatureCheckDialog);

end.
