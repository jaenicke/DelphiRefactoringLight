(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitReferencesDialog;

{
  Non-modal dialog that lists every cross-reference of a Delphi unit:
  for each project file that uses the target unit, every identifier
  from the target that appears in that file is shown as a row in
  a TListView (Identifier, File, Line, Column, Preview).

  Double-click / Enter -> jump to the location (uses Editor.GotoLocation).
}

interface

uses
  System.SysUtils, System.Classes, System.UITypes, System.Generics.Collections,
  Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Dialogs;

type
  TUnitRefItem = record
    Identifier: string;  // the identifier from the target unit
    FilePath: string;    // file in which the reference was found
    Line: Integer;       // 0-based (LSP); ignored when IsDead
    Col: Integer;        // 0-based;       ignored when IsDead
    Length: Integer;     // length of the identifier token
    Preview: string;     // full source line at Line, or info text when IsDead
    /// <summary>True when this row represents a "dead" uses entry: the
    ///  target unit is listed in the file's uses clause but no
    ///  symbols of it are actually referenced. Line/Col are not
    ///  meaningful in that case; goto opens the file at the top.</summary>
    IsDead: Boolean;
  end;

  TUnitRefItems = TArray<TUnitRefItem>;

  /// <summary>Modal dialog with a ListView of all cross-references of a unit.
  ///  Double-click = jump to location (dialog stays open).
  ///  ENTER on a row = jump to location and close dialog.
  ///  ESC / Close = close dialog.</summary>
  TUnitReferencesDialog = class(TForm)
  private
    FStatusLabel: TLabel;
    FProgress: TProgressBar;
    FListView: TListView;
    FBtnGoto: TButton;
    FBtnExportCsv: TButton;
    FBtnClose: TButton;
    FItems: TUnitRefItems;
    FOnGotoLocation: TProc<TUnitRefItem>;
    FOnDialogClose: TNotifyEvent;
    FAllowFree: Boolean;
    FCloseRequested: Boolean;
    FUnitName: string;

    procedure CreateControls;
    procedure DoListDblClick(Sender: TObject);
    procedure DoBtnGotoClick(Sender: TObject);
    procedure DoBtnExportCsvClick(Sender: TObject);
    procedure DoBtnCloseClick(Sender: TObject);
    procedure DoFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure DoListKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function CommonPathPrefix(const AItems: TUnitRefItems): string;
    procedure GotoSelected;
    procedure ExportToCsv(const AFileName: string);
  public
    /// <summary>Creates the dialog. AUnitName is the bare unit name without
    ///  extension; it is used for the caption.</summary>
    constructor CreateDialog(AOwner: TComponent; const AUnitName: string); reintroduce;

    /// <summary>Replaces the current rows with AItems.</summary>
    procedure SetItems(const AItems: TUnitRefItems);

    procedure SetStatus(const AText: string);
    procedure SetProgress(ACurrent, ATotal: Integer);

    /// <summary>Switches the dialog into "review mode": from now on,
    ///  closing the dialog actually frees it. Called by the wizard
    ///  after the search has completed. If the user already requested
    ///  close during the scan, the dialog closes immediately.</summary>
    procedure SetClosable;

    /// <summary>True after the user clicked Close (or pressed ESC) while
    ///  the scan was still running. The wizard polls this to abort
    ///  long-running operations early.</summary>
    property CloseRequested: Boolean read FCloseRequested;

    property OnGotoLocation: TProc<TUnitRefItem> read FOnGotoLocation write FOnGotoLocation;
    /// <summary>Fired right before the dialog is destroyed. The wizard
    ///  uses this to detach LSP trace hooks and clear its own reference.</summary>
    property OnDialogClose: TNotifyEvent read FOnDialogClose write FOnDialogClose;
  end;

implementation

uses
  System.StrUtils, Expert.DialogHelper;

{ TUnitReferencesDialog }

constructor TUnitReferencesDialog.CreateDialog(AOwner: TComponent; const AUnitName: string);
begin
  inherited CreateNew(AOwner);
  FUnitName := AUnitName;
  Caption := 'Unit references: ' + AUnitName;
  Position := poOwnerFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  Width := 900;
  Height := 540;
  Constraints.MinWidth := 540;
  Constraints.MinHeight := 300;
  KeyPreview := True;
  OnKeyDown := DoFormKeyDown;
  OnClose := DoFormClose;

  CreateControls;

  PrepareDialog(Self, AOwner);
end;

procedure TUnitReferencesDialog.CreateControls;
var
  Col: TListColumn;
  BtnPanel: TPanel;
begin
  // Status line
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

  // Buttons
  BtnPanel := TPanel.Create(Self);
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

  FBtnExportCsv := TButton.Create(Self);
  FBtnExportCsv.Parent := BtnPanel;
  FBtnExportCsv.Caption := 'Export CSV...';
  FBtnExportCsv.Width := 120;
  FBtnExportCsv.Height := 28;
  FBtnExportCsv.Top := 6;
  FBtnExportCsv.Anchors := [akTop, akRight];
  FBtnExportCsv.Left := FBtnGoto.Left - FBtnExportCsv.Width - 6;
  FBtnExportCsv.OnClick := DoBtnExportCsvClick;
  FBtnExportCsv.Enabled := False;

  // ListView
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
  Col.Caption := 'Identifier';
  Col.Width := 180;

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
  Col.Width := 360;
end;

function TUnitReferencesDialog.CommonPathPrefix(const AItems: TUnitRefItems): string;
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

procedure TUnitReferencesDialog.SetItems(const AItems: TUnitRefItems);
var
  Item: TUnitRefItem;
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
      if Item.IsDead then
        LI.Caption := '(unused)'
      else
        LI.Caption := Item.Identifier;
      DisplayPath := Item.FilePath;
      if (Prefix <> '') and DisplayPath.StartsWith(Prefix, True) then
        DisplayPath := Copy(DisplayPath, Length(Prefix) + 1, MaxInt);
      LI.SubItems.Add(DisplayPath);
      if Item.IsDead then
      begin
        LI.SubItems.Add('');
        LI.SubItems.Add('');
      end
      else
      begin
        LI.SubItems.Add(IntToStr(Item.Line + 1));
        LI.SubItems.Add(IntToStr(Item.Col + 1));
      end;
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
  FBtnExportCsv.Enabled := Length(AItems) > 0;
end;

procedure TUnitReferencesDialog.SetStatus(const AText: string);
begin
  FStatusLabel.Caption := AText;
end;

procedure TUnitReferencesDialog.SetProgress(ACurrent, ATotal: Integer);
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

procedure TUnitReferencesDialog.GotoSelected;
var
  Idx: Integer;
begin
  if not Assigned(FListView.Selected) then Exit;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if Assigned(FOnGotoLocation) then
    FOnGotoLocation(FItems[Idx]);
end;

procedure TUnitReferencesDialog.DoListDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TUnitReferencesDialog.DoBtnGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TUnitReferencesDialog.DoBtnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TUnitReferencesDialog.DoFormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Close;
    Key := 0;
  end;
end;

procedure TUnitReferencesDialog.DoFormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not FAllowFree then
  begin
    // Search is still in progress - don't free yet. Hide the dialog and
    // remember that the user wanted to close. The wizard calls
    // SetClosable when the search is done; that triggers the actual
    // close.
    FCloseRequested := True;
    Hide;
    Action := caNone;
    Exit;
  end;
  if Assigned(FOnDialogClose) then FOnDialogClose(Self);
  Action := caFree;
end;

procedure TUnitReferencesDialog.SetClosable;
begin
  FAllowFree := True;
  if FCloseRequested then
    Close;
end;

procedure TUnitReferencesDialog.DoListKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    GotoSelected;
    // Stay open - the user wants to keep reviewing.
    Key := 0;
  end;
end;

{ ----- CSV export ----- }

function CsvEscape(const AField: string): string;
var
  NeedsQuotes: Boolean;
  I: Integer;
begin
  NeedsQuotes := False;
  for I := 1 to Length(AField) do
    if CharInSet(AField[I], [',', '"', #13, #10, ';']) then
    begin
      NeedsQuotes := True;
      Break;
    end;
  if NeedsQuotes then
    Result := '"' + StringReplace(AField, '"', '""', [rfReplaceAll]) + '"'
  else
    Result := AField;
end;

procedure TUnitReferencesDialog.ExportToCsv(const AFileName: string);
var
  SL: TStringList;
  Item: TUnitRefItem;
  LineStr, ColStr: string;
begin
  SL := TStringList.Create;
  try
    SL.Add('Identifier,File,Line,Column,Preview');
    for Item in FItems do
    begin
      if Item.IsDead then
      begin
        LineStr := '';
        ColStr := '';
      end
      else
      begin
        LineStr := IntToStr(Item.Line + 1);
        ColStr := IntToStr(Item.Col + 1);
      end;
      SL.Add(
        CsvEscape(IfThen(Item.IsDead, '(unused)', Item.Identifier)) + ',' +
        CsvEscape(Item.FilePath) + ',' +
        LineStr + ',' +
        ColStr + ',' +
        CsvEscape(Item.Preview));
    end;
    SL.SaveToFile(AFileName, TEncoding.UTF8);
  finally
    SL.Free;
  end;
end;

procedure TUnitReferencesDialog.DoBtnExportCsvClick(Sender: TObject);
var
  Dlg: TSaveDialog;
begin
  Dlg := TSaveDialog.Create(Self);
  try
    Dlg.Title := 'Export unit references to CSV';
    Dlg.Filter := 'CSV file (*.csv)|*.csv|All files (*.*)|*.*';
    Dlg.DefaultExt := 'csv';
    Dlg.FileName := FUnitName + '-references.csv';
    Dlg.Options := Dlg.Options + [ofOverwritePrompt, ofPathMustExist];
    if Dlg.Execute then
    try
      ExportToCsv(Dlg.FileName);
      SetStatus('Exported ' + IntToStr(Length(FItems)) + ' row(s) to ' + Dlg.FileName);
    except
      on E: Exception do
        MessageDlg('Export failed: ' + E.Message, mtError, [mbOK], 0);
    end;
  finally
    Dlg.Free;
  end;
end;

initialization
  RegisterDialogClass(TUnitReferencesDialog);

end.
