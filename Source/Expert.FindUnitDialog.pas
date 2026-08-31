(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindUnitDialog;

// "Find Unit": type an identifier, see which unit(s) declare it (from the
// background identifier index over the whole search path), and add the
// chosen unit to the current file's uses clause.

interface

procedure FindUnitForIdentifier;

implementation

uses
  System.SysUtils, System.Classes, System.UITypes, System.SyncObjs,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Dialogs,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.UsesEditor,
  Expert.DialogHelper, Expert.IdeThemes, Expert.ListViewSort;

type
  // Reference-counted hand-off from the search thread back to the UI poll
  // timer. Outlives both the thread and (if it closes early) the form, so a
  // finishing search never touches a freed form.
  ISearchInbox = interface
    procedure Post(AGen: Integer; const AResults: TArray<TFindUnitHit>);
    function Take(out AGen: Integer; out AResults: TArray<TFindUnitHit>): Boolean;
  end;

  TSearchInbox = class(TInterfacedObject, ISearchInbox)
  private
    FLock: TCriticalSection;
    FGen: Integer;
    FHasNew: Boolean;
    FResults: TArray<TFindUnitHit>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Post(AGen: Integer; const AResults: TArray<TFindUnitHit>);
    function Take(out AGen: Integer; out AResults: TArray<TFindUnitHit>): Boolean;
  end;

  TFindUnitDialog = class(TForm)
  private
    FEdit: TEdit;
    FList: TListView;
    FStatus: TLabel;
    FBtnIntf: TButton;
    FBtnImpl: TButton;
    FBtnGoto: TButton;
    FBtnClose: TButton;
    FDebounce: TTimer;
    FPoll: TTimer;
    FTargetFile: string;   // the unit whose uses we edit
    FHits: TArray<TFindUnitHit>;
    FInbox: ISearchInbox;
    FSearchGen: Integer;
    FIndexWasReady: Boolean;
    procedure DoEditChange(Sender: TObject);
    procedure DoDebounce(Sender: TObject);
    procedure DoPoll(Sender: TObject);
    procedure DoListDblClick(Sender: TObject);
    procedure DoAddIntf(Sender: TObject);
    procedure DoAddImpl(Sender: TObject);
    procedure DoGoto(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure LaunchSearch;
    procedure FillResults(const AHits: TArray<TFindUnitHit>);
    procedure AddSelected(ASection: TUsesSection);
    function SelectedHit(out AHit: TFindUnitHit): Boolean;
  public
    constructor CreateDialog(AOwner: TComponent; const AInitial, ATargetFile: string);
  end;

constructor TSearchInbox.Create;
begin
  inherited Create;
  FLock := TCriticalSection.Create;
end;

destructor TSearchInbox.Destroy;
begin
  FLock.Free;
  inherited;
end;

procedure TSearchInbox.Post(AGen: Integer; const AResults: TArray<TFindUnitHit>);
begin
  FLock.Enter;
  try FGen := AGen; FResults := AResults; FHasNew := True; finally FLock.Leave; end;
end;

function TSearchInbox.Take(out AGen: Integer; out AResults: TArray<TFindUnitHit>): Boolean;
begin
  FLock.Enter;
  try
    Result := FHasNew;
    if Result then
    begin AGen := FGen; AResults := FResults; FHasNew := False; end;
  finally
    FLock.Leave;
  end;
end;

constructor TFindUnitDialog.CreateDialog(AOwner: TComponent;
  const AInitial, ATargetFile: string);
var
  Col: TListColumn;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  Caption := 'Find Unit';
  Width := 720;
  Height := 460;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  OnClose := DoFormClose;
  FTargetFile := ATargetFile;

  FEdit := TEdit.Create(Self);
  FEdit.Parent := Self;
  FEdit.Align := alTop;
  FEdit.AlignWithMargins := True;
  FEdit.Margins.SetBounds(8, 8, 8, 4);
  FEdit.TextHint := 'Type an identifier (>= 2 chars)...';
  FEdit.OnChange := DoEditChange;

  FStatus := TLabel.Create(Self);
  FStatus.Parent := Self;
  FStatus.Align := alTop;
  FStatus.AlignWithMargins := True;
  FStatus.Margins.SetBounds(8, 0, 8, 4);
  FStatus.Caption := 'Starting index...';

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.Margins.SetBounds(8, 4, 8, 4);
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.OnDblClick := DoListDblClick;
  Col := FList.Columns.Add; Col.Caption := 'Identifier'; Col.Width := 200;
  Col := FList.Columns.Add; Col.Caption := 'Unit';       Col.Width := 190;
  Col := FList.Columns.Add; Col.Caption := 'File';       Col.Width := 290;
  EnableListViewSorting(FList);

  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 40;
  Panel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Align := alRight;
  FBtnClose.AlignWithMargins := True;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := Panel;
  FBtnGoto.Caption := '&Go to';
  FBtnGoto.Align := alRight;
  FBtnGoto.AlignWithMargins := True;
  FBtnGoto.OnClick := DoGoto;

  FBtnIntf := TButton.Create(Self);
  FBtnIntf.Parent := Panel;
  FBtnIntf.Caption := 'Add to &interface uses';
  FBtnIntf.Align := alLeft;
  FBtnIntf.Width := 160;
  FBtnIntf.AlignWithMargins := True;
  FBtnIntf.Default := True;
  FBtnIntf.OnClick := DoAddIntf;

  FBtnImpl := TButton.Create(Self);
  FBtnImpl.Parent := Panel;
  FBtnImpl.Caption := 'Add to i&mplementation uses';
  FBtnImpl.Align := alLeft;
  FBtnImpl.Width := 190;
  FBtnImpl.AlignWithMargins := True;
  FBtnImpl.OnClick := DoAddImpl;

  FInbox := TSearchInbox.Create;

  FDebounce := TTimer.Create(Self);
  FDebounce.Enabled := False;
  FDebounce.Interval := 200;
  FDebounce.OnTimer := DoDebounce;

  FPoll := TTimer.Create(Self);
  FPoll.Enabled := True;
  FPoll.Interval := 150;   // status + pick up background search results
  FPoll.OnTimer := DoPoll;

  FEdit.Text := AInitial;
  FEdit.SelectAll;

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TFindUnitDialog.DoEditChange(Sender: TObject);
begin
  FDebounce.Enabled := False;
  FDebounce.Enabled := True;   // restart debounce
end;

procedure TFindUnitDialog.DoDebounce(Sender: TObject);
begin
  FDebounce.Enabled := False;
  LaunchSearch;
end;

procedure TFindUnitDialog.LaunchSearch;
var
  LGen: Integer;
  LText: string;
  LInbox: ISearchInbox;
begin
  Inc(FSearchGen);
  LGen := FSearchGen;
  LText := FEdit.Text;
  LInbox := FInbox;   // captured by value (interface) - survives form close
  // The scan runs against the immutable snapshot, so it needs no lock and
  // cannot block the worker. Only the (bounded) result set crosses back.
  TThread.CreateAnonymousThread(
    procedure
    begin
      // Cap hard: with only a couple of letters a substring match hits huge
      // numbers of identifiers. Filling thousands of rows into the (non
      // owner-data) list would stall the UI - the user just types more.
      LInbox.Post(LGen, TUnitIndex.Instance.Search(LText, 200));
    end).Start;
end;

procedure TFindUnitDialog.DoPoll(Sender: TObject);
var
  Gen: Integer;
  Res: TArray<TFindUnitHit>;
begin
  FStatus.Caption := TUnitIndex.Instance.StatusLine;
  // Pick up a finished background search (ignore stale generations).
  if FInbox.Take(Gen, Res) and (Gen = FSearchGen) then
    FillResults(Res);
  // Kick a first search the moment the index becomes usable.
  if TUnitIndex.Instance.Ready and not FIndexWasReady then
  begin
    FIndexWasReady := True;
    LaunchSearch;
  end;
end;

procedure TFindUnitDialog.FillResults(const AHits: TArray<TFindUnitHit>);
var
  I: Integer;
  Item: TListItem;
begin
  FHits := AHits;
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for I := 0 to High(FHits) do
    begin
      Item := FList.Items.Add;
      if FHits[I].IsGeneric then
        Item.Caption := FHits[I].Identifier + '<>'   // generic declaration
      else
        Item.Caption := FHits[I].Identifier;
      Item.SubItems.Add(FHits[I].UnitName);
      Item.SubItems.Add(FHits[I].Path);
      Item.Data := Pointer(NativeInt(I));   // survives click-to-sort reorder
    end;
    if FList.Items.Count > 0 then
      FList.Items[0].Selected := True;
  finally
    FList.Items.EndUpdate;
  end;
end;

function TFindUnitDialog.SelectedHit(out AHit: TFindUnitHit): Boolean;
var
  Idx: NativeInt;
begin
  Result := False;
  if FList.Selected = nil then Exit;
  Idx := NativeInt(FList.Selected.Data);   // fill index, not row index (sortable)
  if (Idx < 0) or (Idx >= Length(FHits)) then Exit;
  AHit := FHits[Idx];
  Result := True;
end;

procedure TFindUnitDialog.AddSelected(ASection: TUsesSection);
var
  H: TFindUnitHit;
  SecName: string;
begin
  if not SelectedHit(H) then
  begin
    ShowThemedMessage('Select a unit in the list first.');
    Exit;
  end;
  if FTargetFile = '' then
  begin
    ShowThemedMessage('No active editor file to add the uses entry to.');
    Exit;
  end;
  if SameText(ChangeFileExt(ExtractFileName(FTargetFile), ''), H.UnitName) then
  begin
    ShowThemedMessage('That is the current unit itself.');
    Exit;
  end;
  if ASection = usInterface then SecName := 'interface' else SecName := 'implementation';
  if AddUnitToUses(FTargetFile, H.UnitName, ASection) then
    ShowThemedMessage(Format('Added "%s" to the %s uses of %s.',
      [H.UnitName, SecName, ExtractFileName(FTargetFile)]))
  else
    ShowThemedMessage(Format('"%s" is already reachable from %s (or could not be written).',
      [H.UnitName, ExtractFileName(FTargetFile)]));
end;

procedure TFindUnitDialog.DoAddIntf(Sender: TObject);
begin
  AddSelected(usInterface);
end;

procedure TFindUnitDialog.DoAddImpl(Sender: TObject);
begin
  AddSelected(usImplementation);
end;

procedure TFindUnitDialog.DoListDblClick(Sender: TObject);
begin
  AddSelected(usInterface);
end;

procedure TFindUnitDialog.DoGoto(Sender: TObject);
var
  H: TFindUnitHit;
begin
  if not SelectedHit(H) then Exit;
  if Editor <> nil then
    Editor.GotoLocation(H.Path, 0, 0, 0);
end;

procedure TFindUnitDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFindUnitDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure FindUnitForIdentifier;
var
  Ctx: TEditorContext;
  Dlg: TFindUnitDialog;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  // Kick / refresh the background index for the current project.
  TUnitIndex.Instance.RefreshSourcesFromEditor;
  Dlg := TFindUnitDialog.CreateDialog(Application.MainForm, Ctx.WordAtCursor, Ctx.FileName);
  Dlg.Show;
end;

end.
