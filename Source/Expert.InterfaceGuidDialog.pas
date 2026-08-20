(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.InterfaceGuidDialog;

// List dialog for the interface-GUID check. Shows every interface
// with its GUID; duplicates sort to the top and render in red.
// Double-click / "Go to" jumps to the declaration.

interface

procedure CheckInterfaceGuids;

implementation

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.Generics.Defaults, System.UITypes,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.InterfaceGuidCheck, Expert.DialogHelper;

// Small helper: bottom-aligned button panel. Kept local to avoid
// pulling a whole layout framework into this dialog.
function CreateButtonRow(AOwner: TForm): TWinControl;
var
  P: TPanel;
begin
  P := TPanel.Create(AOwner);
  P.Parent := AOwner;
  P.Align := alBottom;
  P.Height := 40;
  P.BevelOuter := bvNone;
  Result := P;
end;

type
  TInterfaceGuidDialog = class(TForm)
  private
    FListView: TListView;
    FLblSummary: TLabel;
    FBtnClose: TButton;
    FBtnGoto: TButton;
    FEntries: TArray<TInterfaceGuidEntry>;
    procedure DoCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure GotoSelected;
    procedure FillList;
  public
    constructor CreateDialog(AOwner: TComponent;
      const AEntries: TArray<TInterfaceGuidEntry>);
  end;

constructor TInterfaceGuidDialog.CreateDialog(AOwner: TComponent;
  const AEntries: TArray<TInterfaceGuidEntry>);
var
  Col: TListColumn;
  DupCount, NoGuidCount: Integer;
  E: TInterfaceGuidEntry;
begin
  inherited CreateNew(AOwner);
  Caption := 'Interface GUIDs';
  Width := 900;
  Height := 600;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FEntries := AEntries;

  FLblSummary := TLabel.Create(Self);
  FLblSummary.Parent := Self;
  FLblSummary.Align := alTop;
  FLblSummary.AlignWithMargins := True;
  FLblSummary.Margins.SetBounds(8, 8, 8, 4);

  FListView := TListView.Create(Self);
  FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.AlignWithMargins := True;
  FListView.Margins.SetBounds(8, 4, 8, 4);
  FListView.ViewStyle := vsReport;
  FListView.ReadOnly := True;
  FListView.RowSelect := True;
  FListView.OnCustomDrawItem := DoCustomDrawItem;
  FListView.OnDblClick := DoDblClick;

  Col := FListView.Columns.Add; Col.Caption := 'GUID';      Col.Width := 290;
  Col := FListView.Columns.Add; Col.Caption := 'Interface'; Col.Width := 220;
  Col := FListView.Columns.Add; Col.Caption := 'Unit';      Col.Width := 260;
  Col := FListView.Columns.Add; Col.Caption := 'Line';      Col.Width := 60;

  var Panel := CreateButtonRow(Self);
  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := Panel;
  FBtnGoto.Caption := '&Go to';
  FBtnGoto.Align := alRight;
  FBtnGoto.AlignWithMargins := True;
  FBtnGoto.OnClick := DoGotoClick;
  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Align := alRight;
  FBtnClose.AlignWithMargins := True;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  FillList;

  DupCount := 0; NoGuidCount := 0;
  for E in FEntries do
  begin
    if E.IsDuplicate then Inc(DupCount);
    if not E.HasGuid then Inc(NoGuidCount);
  end;
  FLblSummary.Caption := Format(
    '%d interface(s), %d with duplicate GUIDs (shown red, on top), %d without GUID.',
    [Length(FEntries), DupCount, NoGuidCount]);

  PrepareDialog(Self, AOwner);
end;

procedure TInterfaceGuidDialog.FillList;
var
  Sorted: TArray<TInterfaceGuidEntry>;
  E: TInterfaceGuidEntry;
  Item: TListItem;
begin
  // Duplicates first (grouped by GUID so pairs sit together), then the
  // rest alphabetically by interface name.
  Sorted := Copy(FEntries);
  TArray.Sort<TInterfaceGuidEntry>(Sorted,
    TComparer<TInterfaceGuidEntry>.Construct(
      function(const L, R: TInterfaceGuidEntry): Integer
      begin
        if L.IsDuplicate <> R.IsDuplicate then
        begin
          if L.IsDuplicate then Exit(-1) else Exit(1);
        end;
        if L.IsDuplicate then
        begin
          Result := CompareText(L.Guid, R.Guid);
          if Result <> 0 then Exit;
        end;
        Result := CompareText(L.InterfaceName, R.InterfaceName);
      end));
  FEntries := Sorted;

  FListView.Items.BeginUpdate;
  try
    FListView.Items.Clear;
    for E in FEntries do
    begin
      Item := FListView.Items.Add;
      if E.HasGuid then Item.Caption := E.Guid
      else Item.Caption := '(no GUID)';
      Item.SubItems.Add(E.InterfaceName);
      Item.SubItems.Add(ExtractFileName(E.FileName));
      Item.SubItems.Add(IntToStr(E.Line));
    end;
  finally
    FListView.Items.EndUpdate;
  end;
end;

procedure TInterfaceGuidDialog.DoCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  if (Item.Index >= 0) and (Item.Index < Length(FEntries)) then
  begin
    if FEntries[Item.Index].IsDuplicate then
      Sender.Canvas.Font.Color := clRed
    else if not FEntries[Item.Index].HasGuid then
      Sender.Canvas.Font.Color := clGrayText
    else
      Sender.Canvas.Font.Color := clWindowText;
  end;
end;

procedure TInterfaceGuidDialog.GotoSelected;
var
  Idx: Integer;
begin
  if FListView.Selected = nil then Exit;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FEntries)) then Exit;
  // GotoLocation expects 0-based line/col (LSP convention).
  Editor.GotoLocation(FEntries[Idx].FileName, FEntries[Idx].Line - 1, 0,
    Length(FEntries[Idx].InterfaceName));
end;

procedure TInterfaceGuidDialog.DoDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TInterfaceGuidDialog.DoGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TInterfaceGuidDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure CheckInterfaceGuids;
var
  Files: TArray<string>;
  Entries: TArray<TInterfaceGuidEntry>;
  Dlg: TInterfaceGuidDialog;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowMessage('No project loaded / no source files found.');
    Exit;
  end;
  Entries := TInterfaceGuidChecker.Scan(Files);
  if Length(Entries) = 0 then
  begin
    ShowMessage('No interface declarations found in the project.');
    Exit;
  end;
  Dlg := TInterfaceGuidDialog.CreateDialog(Application.MainForm, Entries);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
