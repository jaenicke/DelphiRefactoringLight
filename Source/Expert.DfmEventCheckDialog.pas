(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DfmEventCheckDialog;

// List dialog for the DFM event-handler check. Shows every event
// reference whose handler is missing (red - crashes on form load) or
// whose signature deviates from the expected event type (orange -
// stack corruption when the event fires). Double-click / "Go to"
// jumps to the .pas declaration (mismatch) or the .dfm line (missing).

interface

procedure CheckDfmEventHandlers;

implementation

uses
  System.SysUtils, System.Classes, System.UITypes,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.DfmEventCheck, Expert.DialogHelper;

type
  TDfmEventCheckDialog = class(TForm)
  private
    FListView: TListView;
    FLblSummary: TLabel;
    FBtnClose: TButton;
    FBtnGoto: TButton;
    FIssues: TArray<TDfmEventIssue>;
    procedure DoCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure GotoSelected;
  public
    constructor CreateDialog(AOwner: TComponent;
      const AIssues: TArray<TDfmEventIssue>);
  end;

constructor TDfmEventCheckDialog.CreateDialog(AOwner: TComponent;
  const AIssues: TArray<TDfmEventIssue>);
var
  Col: TListColumn;
  Issue: TDfmEventIssue;
  Item: TListItem;
  MissCount, SigCount: Integer;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  Caption := 'DFM event handler check';
  Width := 1000;
  Height := 560;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FIssues := AIssues;

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

  Col := FListView.Columns.Add; Col.Caption := 'Problem';   Col.Width := 150;
  Col := FListView.Columns.Add; Col.Caption := 'Form';      Col.Width := 170;
  Col := FListView.Columns.Add; Col.Caption := 'Component'; Col.Width := 150;
  Col := FListView.Columns.Add; Col.Caption := 'Event';     Col.Width := 110;
  Col := FListView.Columns.Add; Col.Caption := 'Handler';   Col.Width := 150;
  Col := FListView.Columns.Add; Col.Caption := 'Details';   Col.Width := 240;

  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 40;
  Panel.BevelOuter := bvNone;

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

  MissCount := 0; SigCount := 0;
  FListView.Items.BeginUpdate;
  try
    for Issue in FIssues do
    begin
      Item := FListView.Items.Add;
      case Issue.Kind of
        eikMissingHandler:
          begin
            Item.Caption := 'Missing handler';
            Inc(MissCount);
          end;
        eikSignatureMismatch:
          begin
            Item.Caption := 'Signature mismatch';
            Inc(SigCount);
          end;
      end;
      Item.SubItems.Add(ExtractFileName(Issue.DfmFile));
      Item.SubItems.Add(Issue.ComponentName + ': ' + Issue.ComponentType);
      Item.SubItems.Add(Issue.EventName);
      Item.SubItems.Add(Issue.HandlerName);
      if Issue.Kind = eikSignatureMismatch then
        Item.SubItems.Add('expected ' + Issue.Expected + ', found ' + Issue.Actual)
      else
        Item.SubItems.Add('crashes with "Method not found" at form load');
    end;
  finally
    FListView.Items.EndUpdate;
  end;

  FLblSummary.Caption := Format(
    '%d issue(s): %d missing handler(s), %d signature mismatch(es).',
    [Length(FIssues), MissCount, SigCount]);

  PrepareDialog(Self, AOwner);
end;

procedure TDfmEventCheckDialog.DoCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  if (Item.Index >= 0) and (Item.Index < Length(FIssues)) then
  begin
    case FIssues[Item.Index].Kind of
      eikMissingHandler:     Sender.Canvas.Font.Color := clRed;
      eikSignatureMismatch:  Sender.Canvas.Font.Color := $00008CFF; // orange
    end;
  end;
end;

procedure TDfmEventCheckDialog.GotoSelected;
var
  Idx: Integer;
begin
  if FListView.Selected = nil then Exit;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FIssues)) then Exit;
  // Signature mismatch -> jump to the .pas declaration; missing
  // handler -> jump to the offending .dfm line so the user sees which
  // event reference to remove (or which method to create).
  if (FIssues[Idx].Kind = eikSignatureMismatch) and (FIssues[Idx].PasLine > 0) then
    Editor.GotoLocation(FIssues[Idx].PasFile, FIssues[Idx].PasLine - 1, 0,
      Length(FIssues[Idx].HandlerName))
  else
    Editor.GotoLocation(FIssues[Idx].DfmFile, FIssues[Idx].DfmLine - 1, 0,
      Length(FIssues[Idx].HandlerName));
end;

procedure TDfmEventCheckDialog.DoDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TDfmEventCheckDialog.DoGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TDfmEventCheckDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure CheckDfmEventHandlers;
var
  Files: TArray<string>;
  Issues: TArray<TDfmEventIssue>;
  Dlg: TDfmEventCheckDialog;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowMessage('No project loaded / no source files found.');
    Exit;
  end;
  Screen.Cursor := crHourGlass;
  try
    Issues := TDfmEventChecker.CheckProject(Files);
  finally
    Screen.Cursor := crDefault;
  end;
  if Length(Issues) = 0 then
  begin
    ShowMessage('No problems found - every DFM event reference has a matching handler.');
    Exit;
  end;
  Dlg := TDfmEventCheckDialog.CreateDialog(Application.MainForm, Issues);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
