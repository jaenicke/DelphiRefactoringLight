(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
  System.Generics.Defaults, System.UITypes, System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.InterfaceGuidCheck, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort, Delphi.FileEncoding,
  // SplitEditorLines (whose trailing empty element IS the file's final line
  // break) and ApplyLinesMinimal - the "Assign new GUID" button writes
  // through the same two as every other edit of this plugin.
  Expert.PascalScanner, Expert.UsesEditor;

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
    FBtnNewGuid: TButton;
    FEntries: TArray<TInterfaceGuidEntry>;
    /// <summary>Live-refresh state. The dialog is non-modal; while it
    ///  is open a timer watches the ACTIVE editor unit. When its
    ///  content changes, only that file is rescanned, its entries are
    ///  swapped in, and duplicates are recomputed - so a fixed (or
    ///  newly introduced) GUID shows up without rerunning the check.</summary>
    FWatchTimer: TTimer;
    FWatchFile: string;
    FWatchContent: string;
    /// <summary>The files the CHECK looked at. A file outside this set
    ///  never changes the result, however often the user edits it
    ///  (audit #37, L3e). nil = no set given, then everything counts.</summary>
    FScanned: TStringList;
    // Header-click sort state (-1 = default duplicates-first ordering).
    // Virtual list: sorting means reordering FEntries + Invalidate.
    FSortCol: Integer;
    FSortAsc: Boolean;
    function CompareByColumn(const L, R: TInterfaceGuidEntry): Integer;
    procedure ApplyColumnSort;
    procedure DoColumnClick(Sender: TObject; Column: TListColumn);
    procedure DoWatchTick(Sender: TObject);
    procedure RefreshFileEntries(const AFile: string);
    function WasScanned(const AFile: string): Boolean;
    procedure UpdateSummary;
    procedure DoCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoData(Sender: TObject; Item: TListItem);
    procedure DoDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    procedure DoNewGuidClick(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure GotoSelected;
    procedure FillList;
  public
    constructor CreateDialog(AOwner: TComponent;
      const AEntries: TArray<TInterfaceGuidEntry>;
      const AScannedFiles: TArray<string> = nil);
    destructor Destroy; override;
  end;

// Only a file the CHECK itself looked at may change the result (audit
// #37, L3e): the watcher follows the active editor file, and a Ctrl+click
// into an RTL or vendor unit used to pull ITS interfaces into the list,
// where they appeared as duplicates of the project's own.
function TInterfaceGuidDialog.WasScanned(const AFile: string): Boolean;
begin
  Result := (FScanned = nil) or (FScanned.IndexOf(AFile) >= 0);
end;

destructor TInterfaceGuidDialog.Destroy;
begin
  FScanned.Free;
  inherited;
end;

constructor TInterfaceGuidDialog.CreateDialog(AOwner: TComponent;
  const AEntries: TArray<TInterfaceGuidEntry>;
  const AScannedFiles: TArray<string>);
var
  Col: TListColumn;
begin
  inherited CreateNew(AOwner);
  Caption := 'Interface GUIDs';
  Width := 900;
  Height := 600;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  // Non-modal: the user jumps to entries via "Go to" and needs to
  // navigate / edit in the editor while the list stays open.
  OnClose := DoFormClose;

  FEntries := AEntries;
  if Length(AScannedFiles) > 0 then
  begin
    FScanned := TStringList.Create;
    FScanned.CaseSensitive := False;
    FScanned.Sorted := True;
    FScanned.AddStrings(AScannedFiles);
  end;

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
  // Virtual mode: with thousands of interfaces (type-library imports!)
  // populating the list item-by-item froze the UI for many seconds.
  // OwnerData renders rows on demand from FEntries instead.
  FListView.OwnerData := True;
  FListView.OnData := DoData;
  FListView.OnCustomDrawItem := DoCustomDrawItem;
  FListView.OnDblClick := DoDblClick;
  FListView.OnColumnClick := DoColumnClick;
  FSortCol := -1;

  Col := FListView.Columns.Add; Col.Caption := 'GUID';      Col.Width := 290;
  Col := FListView.Columns.Add; Col.Caption := 'Interface'; Col.Width := 220;
  Col := FListView.Columns.Add; Col.Caption := 'Unit';      Col.Width := 260;
  Col := FListView.Columns.Add; Col.Caption := 'Line';      Col.Width := 60;

  var Panel := CreateButtonRow(Self);
  // Writes code, so it is never the Default button (audit #41, M40e) and it
  // asks first - WHICH declaration keeps the GUID is the decision here.
  FBtnNewGuid := TButton.Create(Self);
  FBtnNewGuid.Parent := Panel;
  FBtnNewGuid.Caption := 'Assign &new GUID';
  FBtnNewGuid.Width := 130;
  FBtnNewGuid.Align := alLeft;
  FBtnNewGuid.AlignWithMargins := True;
  FBtnNewGuid.OnClick := DoNewGuidClick;
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
  UpdateSummary;

  // Watch the active editor unit for changes and refresh its entries
  // live. 1 s cadence: reading + comparing one file's buffer is cheap.
  FWatchTimer := TTimer.Create(Self);
  FWatchTimer.Interval := 1000;
  FWatchTimer.OnTimer := DoWatchTick;
  FWatchTimer.Enabled := True;

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TInterfaceGuidDialog.UpdateSummary;
var
  DupCount, NoGuidCount: Integer;
  E: TInterfaceGuidEntry;
begin
  DupCount := 0; NoGuidCount := 0;
  for E in FEntries do
  begin
    if E.IsDuplicate then Inc(DupCount);
    if not E.HasGuid then Inc(NoGuidCount);
  end;
  FLblSummary.Caption := Format(
    '%d interface(s), %d with duplicate GUIDs (shown red, on top), %d without GUID.',
    [Length(FEntries), DupCount, NoGuidCount]);
end;

// Effective content of AFile as the compiler / a fresh scan would see
// it: the live editor buffer when the file is open, the disk state
// otherwise (e.g. right after the user closed the unit WITHOUT saving
// - the buffer is gone and the disk state is authoritative again).
function ReadEffectiveContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := True;
  if Editor.ReadEditorContent(AFile, AContent) then Exit;
  AContent := '';
  if not TFile.Exists(AFile) then Exit(False);
  try
    AContent := TDelphiFileEncoding.ReadAll(AFile);
  except
    Result := False;
  end;
end;

procedure TInterfaceGuidDialog.DoWatchTick(Sender: TObject);
var
  FileName, Content, OldEffective: string;
begin
  // Timer context: only the cheap GetActiveFileName is allowed here -
  // GetCurrentContext moves the IDE caret (see TAutoImportLive).
  FileName := Editor.GetActiveFileName;
  if (FileName = '') or not SameText(ExtractFileExt(FileName), '.pas') then
    Exit;
  if not ReadEffectiveContent(FileName, Content) then
    Exit;
  if not SameText(FileName, FWatchFile) then
  begin
    // Active unit switched. Before re-baselining, give the PREVIOUS
    // watch file one final check: if the user closed it without
    // saving, its effective content reverted to the disk state and
    // the list would keep showing the discarded edits.
    if (FWatchFile <> '') and WasScanned(FWatchFile)
       and ReadEffectiveContent(FWatchFile, OldEffective)
       and (OldEffective <> FWatchContent) then
      RefreshFileEntries(FWatchFile);
    // Baseline for the new file; refresh on the NEXT change so merely
    // viewing a file does not churn the list.
    FWatchFile := FileName;
    FWatchContent := Content;
    Exit;
  end;
  if Content = FWatchContent then Exit;
  FWatchContent := Content;
  if WasScanned(FileName) then RefreshFileEntries(FileName);
end;

procedure TInterfaceGuidDialog.RefreshFileEntries(const AFile: string);
var
  Kept: TList<TInterfaceGuidEntry>;
  E: TInterfaceGuidEntry;
  SelName, SelFile: string;
  I: Integer;
begin
  // Remember the selection so the refill can restore it.
  SelName := ''; SelFile := '';
  if (FListView.Selected <> nil) and (FListView.Selected.Index < Length(FEntries)) then
  begin
    SelName := FEntries[FListView.Selected.Index].InterfaceName;
    SelFile := FEntries[FListView.Selected.Index].FileName;
  end;

  Kept := TList<TInterfaceGuidEntry>.Create;
  try
    for E in FEntries do
      if not SameText(E.FileName, AFile) then
        Kept.Add(E);
    Kept.AddRange(TInterfaceGuidChecker.ScanSingleFile(AFile));
    FEntries := Kept.ToArray;
  finally
    Kept.Free;
  end;
  TInterfaceGuidChecker.MarkDuplicates(FEntries);
  FillList;
  UpdateSummary;

  if SelName <> '' then
    for I := 0 to High(FEntries) do
      if SameText(FEntries[I].InterfaceName, SelName)
         and SameText(FEntries[I].FileName, SelFile) then
      begin
        FListView.ItemIndex := I;
        FListView.Selected.MakeVisible(False);
        Break;
      end;
end;

procedure TInterfaceGuidDialog.FillList;
var
  Sorted: TArray<TInterfaceGuidEntry>;
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

  // A header-click sort (if any) overrides the default ordering, also
  // after a live refresh re-sorted the array above.
  ApplyColumnSort;

  // Virtual list: no item creation - just announce the row count and
  // let DoData serve rows as they scroll into view.
  FListView.Items.Count := Length(FEntries);
  FListView.Invalidate;
end;

// Column-to-field mapping mirrors DoData.
function TInterfaceGuidDialog.CompareByColumn(const L, R: TInterfaceGuidEntry): Integer;
begin
  case FSortCol of
    0: Result := CompareText(L.Guid, R.Guid);
    1: Result := CompareText(L.InterfaceName, R.InterfaceName);
    2: Result := CompareText(ExtractFileName(L.FileName), ExtractFileName(R.FileName));
    3: Result := L.Line - R.Line;
  else
    Result := 0;
  end;
  if Result = 0 then
    Result := CompareText(L.InterfaceName, R.InterfaceName);
  if not FSortAsc then
    Result := -Result;
end;

procedure TInterfaceGuidDialog.ApplyColumnSort;
begin
  if FSortCol < 0 then Exit;
  TArray.Sort<TInterfaceGuidEntry>(FEntries,
    TComparer<TInterfaceGuidEntry>.Construct(CompareByColumn));
end;

procedure TInterfaceGuidDialog.DoColumnClick(Sender: TObject; Column: TListColumn);
begin
  if FSortCol = Column.Index then
    FSortAsc := not FSortAsc
  else
  begin
    FSortCol := Column.Index;
    FSortAsc := True;
  end;
  ApplyColumnSort;
  FListView.Invalidate;
  SetListViewSortArrow(FListView, FSortCol, FSortAsc);
end;

procedure TInterfaceGuidDialog.DoData(Sender: TObject; Item: TListItem);
var
  E: TInterfaceGuidEntry;
begin
  if (Item.Index < 0) or (Item.Index >= Length(FEntries)) then Exit;
  E := FEntries[Item.Index];
  if E.HasGuid then Item.Caption := E.Guid
  else Item.Caption := '(no GUID)';
  if E.IsDispInterface then
    Item.SubItems.Add(E.InterfaceName + '  (dispinterface)')
  else
    Item.SubItems.Add(E.InterfaceName);
  Item.SubItems.Add(ExtractFileName(E.FileName));
  Item.SubItems.Add(IntToStr(E.Line));
end;

procedure TInterfaceGuidDialog.DoCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  // Touching the canvas in a custom-draw handler makes the native control
  // use Brush.Color as the row background - without setting it, themed
  // (dark) list views paint these rows on WHITE.
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
  if (Item.Index >= 0) and (Item.Index < Length(FEntries)) then
  begin
    if FEntries[Item.Index].IsDuplicate then
      Sender.Canvas.Font.Color := clRed
    else if not FEntries[Item.Index].HasGuid then
      Sender.Canvas.Font.Color := GetThemedColor(clGrayText)
    else
      Sender.Canvas.Font.Color := GetThemedColor(clWindowText);
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

// The selected declaration gets a FRESH GUID. Only the selected one: for a
// duplicate, the declaration that is NOT touched keeps the GUID - and which
// of them owns it (the oldest, the one a registry entry names, ...) is
// exactly the decision this dialog leaves to the user.
procedure TInterfaceGuidDialog.DoNewGuidClick(Sender: TObject);
var
  Idx: Integer;
  Old, Content, Msg, NewGuid: string;
  Lines: TArray<string>;
  At: TArray<Integer>;
  Plan: TGuidEditPlan;
  SL: TStringList;
  ToBuffer: Boolean;
begin
  if FListView.Selected = nil then
  begin
    ShowThemedMessage('Select the declaration that should get a new GUID.');
    Exit;
  end;
  Idx := FListView.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FEntries)) then Exit;
  var E := FEntries[Idx];

  // Buffer first, disk otherwise - the same reader the live watcher uses.
  if not ReadEffectiveContent(E.FileName, Content) or (Content = '') then
  begin
    ShowThemedMessage('The unit ' + E.FileName + ' could not be read.');
    Exit;
  end;
  Lines := SplitEditorLines(Content);
  // By NAME in the CURRENT text - the line this list shows can have moved.
  At := InterfaceDeclLines(Lines, E.InterfaceName);
  if Length(At) = 0 then
  begin
    ShowThemedMessage(E.InterfaceName + ' is no longer declared in ' +
      ExtractFileName(E.FileName) + ' - run the check again.');
    Exit;
  end;
  if Length(At) > 1 then
  begin
    ShowThemedMessage(E.InterfaceName + ' is declared ' + IntToStr(Length(At)) +
      ' times in ' + ExtractFileName(E.FileName) + ' (a conditional ' +
      'declaration). Assign the GUID by hand - which branch you mean is not ' +
      'decidable here.');
    Exit;
  end;

  if E.HasGuid then
    Msg := E.InterfaceName + ' carries ' + E.Guid + '.' + sLineBreak +
      'Give THIS declaration a new GUID?' + sLineBreak + sLineBreak +
      'Another declaration on the same GUID keeps it - so do this on the ' +
      'copy, not on the original.'
  else
    Msg := E.InterfaceName + ' has no GUID.' + sLineBreak +
      'Insert a fresh one below its declaration?';
  if not AskThemedConfirm(Msg, 'Assign new GUID') then Exit;

  NewGuid := UpperCase(GUIDToString(TGUID.NewGuid));
  Plan := PlanInterfaceGuidEdit(Lines, At[0], NewGuid);
  if not Plan.Ok then
  begin
    ShowThemedMessage('Not changed: ' + Plan.Problem);
    Exit;
  end;
  // BEFORE the write: this check lists every project source, so most of
  // its files are NOT open - ApplyLinesMinimal then writes them to disk
  // and no Ctrl+Z can reach that (user report 2026-10-05).
  ToBuffer := MinimalWriteGoesToBuffer(E.FileName);
  Old := Content;
  var LB := #13#10;
  if (Pos(#13#10, Old) = 0) and (Pos(#10, Old) > 0) then LB := #10;
  SL := TStringList.Create;
  try
    // Joined, not Add-ed: the trailing empty element of the split IS the
    // file's final line break and TStringList.Text appends one (1.18.7).
    SL.Text := string.Join(LB, Plan.Lines);
    if not ApplyLinesMinimal(E.FileName, SL, Old) then
    begin
      ShowThemedMessage('The unit changed while the dialog was open - nothing ' +
        'was written. Run the check again.');
      Exit;
    end;
  finally
    SL.Free;
  end;

  // The list must show what the file now says, and the duplicate flags are
  // computed over the whole set - so refresh that file's entries.
  RefreshFileEntries(E.FileName);
  if Plan.Inserted then
    ShowThemedMessage(E.InterfaceName + ' now carries ' + NewGuid +
      ' (inserted on line ' + IntToStr(Plan.Line + 1) + '). ' +
      MinimalWriteNote(ToBuffer))
  else
    ShowThemedMessage(E.InterfaceName + ': ' + Plan.OldGuid + ' -> ' + NewGuid +
      '. ' + MinimalWriteNote(ToBuffer));
end;

procedure TInterfaceGuidDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TInterfaceGuidDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  // Non-modal dialog owns its own lifetime.
  Action := caFree;
end;

procedure CheckInterfaceGuids;
var
  Files: TArray<string>;
  Entries: TArray<TInterfaceGuidEntry>;
  Dlg: TInterfaceGuidDialog;
  ProgressForm: TCheckProgressWindow;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowThemedMessage('No project loaded / no source files found.');
    Exit;
  end;

  // Small always-on-top progress window; the scan runs synchronously
  // on the UI thread, so we pump messages every few files to keep the
  // label painting. Without this, big projects (many hundred files,
  // possibly on a network share) look like a silent hang.
  ProgressForm := CreateCheckProgress('Interface GUIDs', nil);
  try
    Screen.Cursor := crHourGlass;
    try
      Entries := TInterfaceGuidChecker.Scan(Files,
        procedure(ACurrent, ATotal: Integer; AFile: string)
        begin
          if (ACurrent mod 10 = 0) or (ACurrent = ATotal) then
            ProgressForm.Step(ACurrent, ATotal, ExtractFileName(AFile));
        end);
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    ProgressForm.Free;
  end;

  if Length(Entries) = 0 then
  begin
    ShowThemedMessage('No interface declarations found in the project.');
    Exit;
  end;
  // Non-modal (frees itself on close) so the user can keep navigating
  // and editing while the list stays open.
  Dlg := TInterfaceGuidDialog.CreateDialog(Application.MainForm, Entries, Files);
  Dlg.Show;
end;

end.
