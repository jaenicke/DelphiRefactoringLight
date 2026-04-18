(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.RenameDialog;

interface

uses
  System.SysUtils, System.Classes, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls;

type
  /// <summary>A single entry for the rename preview list.</summary>
  TRenamePreviewItem = record
    FilePath: string;
    Line: Integer;           // 0-based
    Col: Integer;            // 0-based
    /// <summary>Match kind, e.g. "Interface declaration", "Implementation",
    ///  "Class declaration", "Call".</summary>
    Kind: string;
    /// <summary>Original line content.</summary>
    OriginalLine: string;
    /// <summary>Line content after replacement (preview).</summary>
    PreviewLine: string;
  end;

  TRenamePreviewItems = TArray<TRenamePreviewItem>;

  TRenameDialog = class(TForm)
  private
    FLblOldName: TLabel;
    FEdtOldName: TEdit;
    FLblNewName: TLabel;
    FEdtNewName: TEdit;
    FLblNameCheck: TLabel;
    FChkBackup: TCheckBox;
    FBtnPreview: TButton;
    FBtnRename: TButton;
    FBtnCancel: TButton;
    FProgressBar: TProgressBar;
    FLblStatus: TLabel;
    FPanelBottom: TPanel;
    FPanelTop: TPanel;
    FPageControl: TPageControl;
    FTabChanges: TTabSheet;
    FTabDetails: TTabSheet;
    FListView: TListView;
    FMemoDetails: TMemo;
    FCheckTimer: TTimer;
    FOldName: string;
    FCurrentFile: string;
    FCurrentFileText: string;
    FProjectFiles: TArray<string>;
    procedure DoFormShow(Sender: TObject);
    procedure DoBtnPreviewClick(Sender: TObject);
    procedure DoNewNameChange(Sender: TObject);
    procedure DoCheckTimer(Sender: TObject);
    procedure RunIdentifierCheck;
    function CommonPathPrefix(const AItems: TRenamePreviewItems): string;
  public
    OnPreviewRequested: TNotifyEvent;

    constructor CreateDialog(AOwner: TComponent; const AOldName: string);

    /// <summary>Configure the live identifier check. ACurrentFile is the
    ///  active unit (excluded from the project-wide scan). AProjectFiles is
    ///  the full list of project source files. When set, a debounced
    ///  validity + collision check runs on every keystroke and the status
    ///  is displayed below the "New name" edit.</summary>
    procedure SetCheckContext(const ACurrentFile: string; const AProjectFiles: TArray<string>);

    function GetNewName: string;
    function GetCreateBackup: Boolean;

    /// <summary>Pre-fill the "New name" edit with a value (used e.g. by
    ///  the unit-rename flow to propose the new unit name).</summary>
    procedure SetNewName(const AName: string);

    /// <summary>Populates the list view with preview entries. Paths are
    ///  shortened by their common prefix, line content is trimmed for
    ///  display (full content stays inside the record).</summary>
    procedure SetPreviewItems(const AItems: TRenamePreviewItems);

    /// <summary>Sets the text for the Details tab (diagnostic log).</summary>
    procedure SetDetailsText(const AText: string);

    procedure SetProgress(AValue, AMax: Integer);
    procedure SetStatus(const AText: string);
    procedure EnableRename(AEnabled: Boolean);
    procedure SetBusy(ABusy: Boolean);
  end;

implementation

uses
  System.UITypes, System.IOUtils, Vcl.Graphics, Expert.IdentifierCheck;

constructor TRenameDialog.CreateDialog(AOwner: TComponent; const AOldName: string);
var
  Col: TListColumn;
begin
  inherited CreateNew(AOwner);
  Caption := 'Rename Identifier';
  Position := poMainFormCenter;
  BorderStyle := bsSizeable;
  BorderIcons := [biSystemMenu];
  ClientWidth := 880;
  ClientHeight := 560;
  Constraints.MinWidth := 640;
  Constraints.MinHeight := 380;
  OnShow := DoFormShow;

  // --- Top panel ---
  FPanelTop := TPanel.Create(Self);
  FPanelTop.Parent := Self;
  FPanelTop.Align := alTop;
  FPanelTop.Height := 160;
  FPanelTop.BevelOuter := bvNone;

  FOldName := AOldName;

  FLblOldName := TLabel.Create(Self);
  FLblOldName.Parent := FPanelTop;
  FLblOldName.Left := 12;
  FLblOldName.Top := 8;
  FLblOldName.Caption := 'Current identifier:';

  FEdtOldName := TEdit.Create(Self);
  FEdtOldName.Parent := FPanelTop;
  FEdtOldName.Left := 12;
  FEdtOldName.Top := 26;
  FEdtOldName.Width := ClientWidth - 24;
  FEdtOldName.Anchors := [akLeft, akTop, akRight];
  FEdtOldName.ReadOnly := True;
  FEdtOldName.Text := AOldName;
  FEdtOldName.Color := $F0F0F0;
  FEdtOldName.Font.Size := 10;

  FLblNewName := TLabel.Create(Self);
  FLblNewName.Parent := FPanelTop;
  FLblNewName.Left := 12;
  FLblNewName.Top := 56;
  FLblNewName.Caption := 'New name:';

  FEdtNewName := TEdit.Create(Self);
  FEdtNewName.Parent := FPanelTop;
  FEdtNewName.Left := 12;
  FEdtNewName.Top := 74;
  FEdtNewName.Width := ClientWidth - 24;
  FEdtNewName.Anchors := [akLeft, akTop, akRight];
  FEdtNewName.Text := AOldName;
  FEdtNewName.Font.Size := 10;
  FEdtNewName.OnChange := DoNewNameChange;

  FLblNameCheck := TLabel.Create(Self);
  FLblNameCheck.Parent := FPanelTop;
  FLblNameCheck.Left := 12;
  FLblNameCheck.Top := 102;
  FLblNameCheck.AutoSize := False;
  FLblNameCheck.Width := ClientWidth - 24;
  FLblNameCheck.Height := 16;
  FLblNameCheck.Anchors := [akLeft, akTop, akRight];
  FLblNameCheck.Caption := '';

  FCheckTimer := TTimer.Create(Self);
  FCheckTimer.Enabled := False;
  FCheckTimer.Interval := 350;
  FCheckTimer.OnTimer := DoCheckTimer;

  FChkBackup := TCheckBox.Create(Self);
  FChkBackup.Parent := FPanelTop;
  FChkBackup.Left := 12;
  FChkBackup.Top := 128;
  FChkBackup.Width := 200;
  FChkBackup.Caption := 'Create backup';
  FChkBackup.Checked := True;

  FBtnCancel := TButton.Create(Self);
  FBtnCancel.Parent := FPanelTop;
  FBtnCancel.Width := 90;
  FBtnCancel.Height := 28;
  FBtnCancel.Top := 124;
  FBtnCancel.Anchors := [akTop, akRight];
  FBtnCancel.Left := ClientWidth - FBtnCancel.Width - 12;
  FBtnCancel.Caption := 'Cancel';
  FBtnCancel.ModalResult := mrCancel;
  FBtnCancel.Cancel := True;

  FBtnRename := TButton.Create(Self);
  FBtnRename.Parent := FPanelTop;
  FBtnRename.Width := 96;
  FBtnRename.Height := 28;
  FBtnRename.Top := 124;
  FBtnRename.Anchors := [akTop, akRight];
  FBtnRename.Left := FBtnCancel.Left - FBtnRename.Width - 6;
  FBtnRename.Caption := 'Rename';
  FBtnRename.ModalResult := mrOk;
  FBtnRename.Enabled := False;
  FBtnRename.Default := True;

  FBtnPreview := TButton.Create(Self);
  FBtnPreview.Parent := FPanelTop;
  FBtnPreview.Width := 90;
  FBtnPreview.Height := 28;
  FBtnPreview.Top := 124;
  FBtnPreview.Anchors := [akTop, akRight];
  FBtnPreview.Left := FBtnRename.Left - FBtnPreview.Width - 6;
  FBtnPreview.Caption := 'Preview';
  FBtnPreview.OnClick := DoBtnPreviewClick;

  // --- Bottom panel (status + progress) ---
  FPanelBottom := TPanel.Create(Self);
  FPanelBottom.Parent := Self;
  FPanelBottom.Align := alBottom;
  FPanelBottom.Height := 40;
  FPanelBottom.BevelOuter := bvNone;

  FProgressBar := TProgressBar.Create(Self);
  FProgressBar.Parent := FPanelBottom;
  FProgressBar.Left := 12;
  FProgressBar.Top := 2;
  FProgressBar.Width := ClientWidth - 24;
  FProgressBar.Height := 14;
  FProgressBar.Anchors := [akLeft, akTop, akRight];
  FProgressBar.Visible := False;

  FLblStatus := TLabel.Create(Self);
  FLblStatus.Parent := FPanelBottom;
  FLblStatus.Left := 12;
  FLblStatus.Top := 20;
  FLblStatus.AutoSize := False;
  FLblStatus.Width := ClientWidth - 24;
  FLblStatus.Anchors := [akLeft, akTop, akRight];
  FLblStatus.Caption := 'Ready.';

  // --- PageControl fills the middle ---
  FPageControl := TPageControl.Create(Self);
  FPageControl.Parent := Self;
  FPageControl.Align := alClient;
  FPageControl.AlignWithMargins := True;
  FPageControl.Margins.SetBounds(8, 4, 8, 4);

  FTabChanges := TTabSheet.Create(FPageControl);
  FTabChanges.PageControl := FPageControl;
  FTabChanges.Caption := 'Changes';

  FListView := TListView.Create(Self);
  FListView.Parent := FTabChanges;
  FListView.Align := alClient;
  FListView.ViewStyle := vsReport;
  FListView.RowSelect := True;
  FListView.ReadOnly := True;
  FListView.HideSelection := False;
  FListView.GridLines := True;

  Col := FListView.Columns.Add;
  Col.Caption := 'File';
  Col.Width := 200;

  Col := FListView.Columns.Add;
  Col.Caption := 'Line';
  Col.Width := 60;
  Col.Alignment := taRightJustify;

  Col := FListView.Columns.Add;
  Col.Caption := 'Kind';
  Col.Width := 160;

  Col := FListView.Columns.Add;
  Col.Caption := 'Original';
  Col.Width := 220;

  Col := FListView.Columns.Add;
  Col.Caption := 'Preview';
  Col.Width := 220;

  FTabDetails := TTabSheet.Create(FPageControl);
  FTabDetails.PageControl := FPageControl;
  FTabDetails.Caption := 'Details';

  FMemoDetails := TMemo.Create(Self);
  FMemoDetails.Parent := FTabDetails;
  FMemoDetails.Align := alClient;
  FMemoDetails.ReadOnly := True;
  FMemoDetails.ScrollBars := ssBoth;
  FMemoDetails.Font.Name := 'Consolas';
  FMemoDetails.Font.Size := 9;
  FMemoDetails.WordWrap := False;

  FPageControl.ActivePage := FTabChanges;
end;

procedure TRenameDialog.DoFormShow(Sender: TObject);
begin
  FEdtNewName.SetFocus;
  FEdtNewName.SelectAll;
end;

procedure TRenameDialog.DoBtnPreviewClick(Sender: TObject);
begin
  if Assigned(OnPreviewRequested) then
    OnPreviewRequested(Self);
end;

procedure TRenameDialog.DoNewNameChange(Sender: TObject);
begin
  FListView.Clear;
  FMemoDetails.Clear;
  FBtnRename.Enabled := False;
  FLblStatus.Caption := 'Ready.';

  // Debounce the live identifier check
  FCheckTimer.Enabled := False;
  if FCurrentFile <> '' then
  begin
    FLblNameCheck.Font.Color := clGrayText;
    FLblNameCheck.Caption := 'Checking...';
    FCheckTimer.Enabled := True;
  end;
end;

procedure TRenameDialog.DoCheckTimer(Sender: TObject);
begin
  FCheckTimer.Enabled := False;
  RunIdentifierCheck;
end;

procedure TRenameDialog.RunIdentifierCheck;
var
  Res: TIdentifierCheckResult;
begin
  Res := TIdentifierChecker.Check(FEdtNewName.Text, FOldName,
    FCurrentFileText, FProjectFiles, FCurrentFile);

  case Res.Status of
    icsOk:        FLblNameCheck.Font.Color := clGreen;
    icsUnchanged: FLblNameCheck.Font.Color := clGrayText;
    icsInProject: FLblNameCheck.Font.Color := $00008CFF; // orange-ish
  else
    FLblNameCheck.Font.Color := clRed;
  end;
  FLblNameCheck.Font.Style := [fsBold];
  FLblNameCheck.Caption := Res.Message;
end;

procedure TRenameDialog.SetCheckContext(const ACurrentFile: string;
  const AProjectFiles: TArray<string>);
begin
  FCurrentFile := ACurrentFile;
  FProjectFiles := AProjectFiles;
  FCurrentFileText := '';
  if (ACurrentFile <> '') and TFile.Exists(ACurrentFile) then
  begin
    try
      FCurrentFileText := TFile.ReadAllText(ACurrentFile);
    except
    end;
  end;
  // Run an initial check right away
  FCheckTimer.Enabled := False;
  RunIdentifierCheck;
end;

function TRenameDialog.GetNewName: string;
begin
  Result := Trim(FEdtNewName.Text);
end;

function TRenameDialog.GetCreateBackup: Boolean;
begin
  Result := FChkBackup.Checked;
end;

procedure TRenameDialog.SetNewName(const AName: string);
begin
  FEdtNewName.Text := AName;
  FEdtNewName.SelectAll;
end;

function TRenameDialog.CommonPathPrefix(const AItems: TRenamePreviewItems): string;
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

procedure TRenameDialog.SetPreviewItems(const AItems: TRenamePreviewItems);
var
  Item: TRenamePreviewItem;
  LI: TListItem;
  Prefix, DisplayPath: string;
begin
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
      LI.SubItems.Add(Item.Kind);
      LI.SubItems.Add(Trim(Item.OriginalLine));
      LI.SubItems.Add(Trim(Item.PreviewLine));
    end;
    if FListView.Items.Count > 0 then
    begin
      FListView.Items[0].Selected := True;
      FListView.Items[0].Focused := True;
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  FBtnRename.Enabled := Length(AItems) > 0;
end;

procedure TRenameDialog.SetDetailsText(const AText: string);
begin
  FMemoDetails.Lines.Text := AText;
end;

procedure TRenameDialog.SetProgress(AValue, AMax: Integer);
begin
  FProgressBar.Visible := AMax > 0;
  FProgressBar.Max := AMax;
  FProgressBar.Position := AValue;
  Application.ProcessMessages;
end;

procedure TRenameDialog.SetStatus(const AText: string);
begin
  FLblStatus.Caption := AText;
  Application.ProcessMessages;
end;

procedure TRenameDialog.EnableRename(AEnabled: Boolean);
begin
  FBtnRename.Enabled := AEnabled;
end;

procedure TRenameDialog.SetBusy(ABusy: Boolean);
begin
  FBtnPreview.Enabled := not ABusy;
  if ABusy then
    FBtnRename.Enabled := False;
  FEdtNewName.Enabled := not ABusy;
  if ABusy then
    Screen.Cursor := crHourGlass
  else
    Screen.Cursor := crDefault;
  Application.ProcessMessages;
end;

end.
