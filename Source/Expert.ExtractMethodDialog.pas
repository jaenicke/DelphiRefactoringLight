(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ExtractMethodDialog;

interface

uses
  System.SysUtils, System.Classes, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Expert.IdentifierCheck;

type
  TExtractMethodDialog = class(TForm)
  private
    FPanelTop: TPanel;
    FPanelBottom: TPanel;
    FLblMethodName: TLabel;
    FEdtMethodName: TEdit;
    FLblNameCheck: TLabel;
    FProgressBar: TProgressBar;
    FLblStatus: TLabel;
    FMemoPreview: TMemo;
    FBtnExtract: TButton;
    FBtnCancel: TButton;
    FCheckTimer: TTimer;
    FCurrentFile: string;
    FCurrentFileText: string;
    FIndex: TProjectTextIndex;
    procedure DoMethodNameChange(Sender: TObject);
    procedure DoFormShow(Sender: TObject);
    procedure DoCheckTimer(Sender: TObject);
    procedure RunIdentifierCheck;
  public
    OnNameChanged: TNotifyEvent;
    constructor CreateDialog(AOwner: TComponent; const ADefaultName: string);
    destructor Destroy; override;

    /// <summary>Configure the live identifier check. See
    ///  TRenameDialog.SetCheckContext.</summary>
    procedure SetCheckContext(const ACurrentFile: string; const AProjectFiles: TArray<string>);

    function GetMethodName: string;
    procedure SetStatus(const AText: string);
    procedure SetProgress(AValue, AMax: Integer);
    procedure SetPreviewText(const AText: string);
    procedure EnableExtract(AEnabled: Boolean);
    procedure SetBusy(ABusy: Boolean);
  end;

implementation

uses
  System.IOUtils, Vcl.Graphics;

constructor TExtractMethodDialog.CreateDialog(AOwner: TComponent; const ADefaultName: string);
begin
  inherited CreateNew(AOwner);
  Caption := 'Extract Method';
  ClientWidth := 700;
  ClientHeight := 520;
  Position := poMainFormCenter;
  BorderStyle := bsDialog;
  OnShow := DoFormShow;

  // --- Top area ---
  FPanelTop := TPanel.Create(Self);
  FPanelTop.Parent := Self;
  FPanelTop.Align := alTop;
  FPanelTop.Height := 94;
  FPanelTop.BevelOuter := bvNone;

  FLblMethodName := TLabel.Create(Self);
  FLblMethodName.Parent := FPanelTop;
  FLblMethodName.Left := 12;
  FLblMethodName.Top := 8;
  FLblMethodName.Caption := 'Name of the new method:';

  FEdtMethodName := TEdit.Create(Self);
  FEdtMethodName.Parent := FPanelTop;
  FEdtMethodName.Left := 12;
  FEdtMethodName.Top := 26;
  FEdtMethodName.Width := 460;
  FEdtMethodName.Font.Size := 10;
  FEdtMethodName.Text := ADefaultName;
  FEdtMethodName.OnChange := DoMethodNameChange;

  FLblNameCheck := TLabel.Create(Self);
  FLblNameCheck.Parent := FPanelTop;
  FLblNameCheck.Left := 12;
  FLblNameCheck.Top := 56;
  FLblNameCheck.AutoSize := False;
  FLblNameCheck.Width := 676;
  FLblNameCheck.Height := 16;
  FLblNameCheck.Caption := '';

  FCheckTimer := TTimer.Create(Self);
  FCheckTimer.Enabled := False;
  FCheckTimer.Interval := 350;
  FCheckTimer.OnTimer := DoCheckTimer;

  FBtnExtract := TButton.Create(Self);
  FBtnExtract.Parent := FPanelTop;
  FBtnExtract.Left := 490;
  FBtnExtract.Top := 24;
  FBtnExtract.Width := 96;
  FBtnExtract.Height := 28;
  FBtnExtract.Caption := 'Extract';
  FBtnExtract.ModalResult := mrOk;
  FBtnExtract.Enabled := False;
  FBtnExtract.Default := True;

  FBtnCancel := TButton.Create(Self);
  FBtnCancel.Parent := FPanelTop;
  FBtnCancel.Left := 592;
  FBtnCancel.Top := 24;
  FBtnCancel.Width := 96;
  FBtnCancel.Height := 28;
  FBtnCancel.Caption := 'Cancel';
  FBtnCancel.ModalResult := mrCancel;
  FBtnCancel.Cancel := True;

  // --- Bottom area (status + progress) ---
  FPanelBottom := TPanel.Create(Self);
  FPanelBottom.Parent := Self;
  FPanelBottom.Align := alBottom;
  FPanelBottom.Height := 38;
  FPanelBottom.BevelOuter := bvNone;

  FProgressBar := TProgressBar.Create(Self);
  FProgressBar.Parent := FPanelBottom;
  FProgressBar.Left := 12;
  FProgressBar.Top := 2;
  FProgressBar.Width := 676;
  FProgressBar.Height := 14;
  FProgressBar.Visible := False;

  FLblStatus := TLabel.Create(Self);
  FLblStatus.Parent := FPanelBottom;
  FLblStatus.Left := 12;
  FLblStatus.Top := 19;
  FLblStatus.AutoSize := False;
  FLblStatus.Width := 676;
  FLblStatus.Caption := 'Ready.';

  // --- Preview Memo ---
  FMemoPreview := TMemo.Create(Self);
  FMemoPreview.Parent := Self;
  FMemoPreview.Align := alClient;
  FMemoPreview.AlignWithMargins := True;
  FMemoPreview.Margins.Left := 12;
  FMemoPreview.Margins.Right := 12;
  FMemoPreview.Margins.Top := 4;
  FMemoPreview.Margins.Bottom := 4;
  FMemoPreview.ReadOnly := True;
  FMemoPreview.ScrollBars := ssBoth;
  FMemoPreview.Font.Name := 'Consolas';
  FMemoPreview.Font.Size := 9;
  FMemoPreview.WordWrap := False;

  FIndex := TProjectTextIndex.Create;
end;

destructor TExtractMethodDialog.Destroy;
begin
  FIndex.Free;
  inherited;
end;

procedure TExtractMethodDialog.DoFormShow(Sender: TObject);
begin
  // SetFocus can crash when the form is not yet fully initialized (e.g.
  // when shown again after Hide). Try/Except guards against that.
  try
    if FEdtMethodName.CanFocus then
    begin
      FEdtMethodName.SetFocus;
      FEdtMethodName.SelectAll;
    end;
  except
  end;
end;

procedure TExtractMethodDialog.DoMethodNameChange(Sender: TObject);
begin
  if Assigned(OnNameChanged) then
    OnNameChanged(Self);

  FCheckTimer.Enabled := False;
  if FCurrentFile <> '' then
  begin
    FLblNameCheck.Font.Color := clGrayText;
    FLblNameCheck.Caption := 'Checking...';
    FCheckTimer.Enabled := True;
  end;
end;

procedure TExtractMethodDialog.DoCheckTimer(Sender: TObject);
begin
  FCheckTimer.Enabled := False;
  if not FIndex.PollReady then
  begin
    FLblNameCheck.Font.Color := clGrayText;
    FLblNameCheck.Font.Style := [];
    FLblNameCheck.Caption := 'Indexing project for collision check...';
    FCheckTimer.Interval := 200;
    FCheckTimer.Enabled := True;
    Exit;
  end;
  FCheckTimer.Interval := 350;
  RunIdentifierCheck;
end;

procedure TExtractMethodDialog.RunIdentifierCheck;
var
  Res: TIdentifierCheckResult;
begin
  Res := TIdentifierChecker.Check(FEdtMethodName.Text, '',
    FCurrentFileText, FIndex.OtherContents);

  case Res.Status of
    icsOk:        FLblNameCheck.Font.Color := clGreen;
    icsInProject: FLblNameCheck.Font.Color := $00008CFF;
  else
    FLblNameCheck.Font.Color := clRed;
  end;
  FLblNameCheck.Font.Style := [fsBold];
  FLblNameCheck.Caption := Res.Message;
end;

procedure TExtractMethodDialog.SetCheckContext(const ACurrentFile: string;
  const AProjectFiles: TArray<string>);
begin
  FCurrentFile := ACurrentFile;
  FCurrentFileText := '';
  if (ACurrentFile <> '') and TFile.Exists(ACurrentFile) then
  begin
    try
      FCurrentFileText := TFile.ReadAllText(ACurrentFile);
    except
    end;
  end;
  FIndex.Build(AProjectFiles, ACurrentFile);
  FCheckTimer.Enabled := False;
  DoCheckTimer(nil);
end;

function TExtractMethodDialog.GetMethodName: string;
begin
  Result := Trim(FEdtMethodName.Text);
end;

procedure TExtractMethodDialog.SetStatus(const AText: string);
begin
  FLblStatus.Caption := AText;
  FLblStatus.Update;
end;

procedure TExtractMethodDialog.SetProgress(AValue, AMax: Integer);
begin
  FProgressBar.Visible := AMax > 0;
  FProgressBar.Max := AMax;
  FProgressBar.Position := AValue;
  FProgressBar.Update;
end;

procedure TExtractMethodDialog.SetPreviewText(const AText: string);
begin
  FMemoPreview.Lines.Text := AText;
end;

procedure TExtractMethodDialog.EnableExtract(AEnabled: Boolean);
begin
  FBtnExtract.Enabled := AEnabled;
end;

procedure TExtractMethodDialog.SetBusy(ABusy: Boolean);
begin
  FEdtMethodName.Enabled := not ABusy;
  if ABusy then
  begin
    FBtnExtract.Enabled := False;
    Screen.Cursor := crHourGlass;
  end
  else
    Screen.Cursor := crDefault;
end;

end.
