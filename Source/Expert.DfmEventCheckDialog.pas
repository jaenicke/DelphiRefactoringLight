(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DfmEventCheckDialog;

// List dialog for the DFM event-handler check. Shows every event
// reference whose handler is missing (crashes on form load) or whose
// signature deviates from the expected event type (stack corruption
// when the event fires). Signature mismatches highlight the differing
// parameter types in bold in the Details column; auto-fixable rows are
// flagged in the "Fix" column and pre-ticked. Double-click / "Go to"
// jumps to the .pas declaration (mismatch) or the .dfm line (missing).

interface

procedure CheckDfmEventHandlers;

implementation

uses
  System.SysUtils, System.Classes, System.UITypes, System.IOUtils,
  System.StrUtils, System.Win.Registry, System.Generics.Collections,
  System.Generics.Defaults, Winapi.Windows, Winapi.Messages, Winapi.CommCtrl,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.DfmEventCheck, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort;

type
  // ListView that re-applies native double buffering on every handle
  // (re)creation. Toggling GroupView/Checkboxes recreates the handle, so
  // setting LVS_EX_DOUBLEBUFFER once from outside would not stick;
  // CreateWnd runs after each recreation and makes it permanent.
  // It also OWNER-DRAWS the group headers: the native painter uses a
  // hard-coded dark-blue accent (and ignores clrText for groups), which
  // is unreadable on the IDE's dark theme.
  TBufferedListView = class(TListView)
  protected
    procedure CreateWnd; override;
    procedure CNNotify(var Message: TWMNotify); message CN_NOTIFY;
  end;

  TDfmEventCheckDialog = class(TForm)
  private
    FListView: TListView;
    FLblSummary: TLabel;
    FLblStatus: TLabel;    // shows which unit is being processed right now
    FEdtFilter: TEdit;
    FBtnClose: TButton;
    FBtnGoto: TButton;
    FBtnGotoType: TButton;
    FBtnFix: TButton;
    FChkIde: TCheckBox;
    FChkAliasDiff: TCheckBox;
    FIssues: TArray<TDfmEventIssue>;
    FExpTok, FActTok: TArray<TArray<string>>;   // pre-split for fast painting
    /// <summary>Check state per ISSUE (not per row): the list is sortable
    ///  and filterable, so rows come and go - ticks must survive both.</summary>
    FChecked: TArray<Boolean>;
    FFillingRows: Boolean;
    procedure FillRows;
    procedure DoFilterChange(Sender: TObject);
    procedure DoItemChecked(Sender: TObject; Item: TListItem);
    procedure DoAdvancedDrawSubItem(Sender: TCustomListView; Item: TListItem;
      SubItem: Integer; State: TCustomDrawState; Stage: TCustomDrawStage;
      var DefaultDraw: Boolean);
    procedure DoAliasDiffClick(Sender: TObject);
    function GroupKeyOf(const AIssue: TDfmEventIssue): string;
    function NameColumn(const AFile: string; ALine1: Integer;
      const AName: string): Integer;
    procedure DoDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    procedure DoGotoTypeClick(Sender: TObject);
    procedure DoFixClick(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure GotoSelected;
  public
    constructor CreateDialog(AOwner: TComponent;
      const AIssues: TArray<TDfmEventIssue>);
  end;

procedure TBufferedListView.CreateWnd;
begin
  inherited;
  SendMessage(Handle, LVM_SETEXTENDEDLISTVIEWSTYLE,
    LVS_EX_DOUBLEBUFFER, LVS_EX_DOUBLEBUFFER);
end;

procedure TBufferedListView.CNNotify(var Message: TWMNotify);
var
  Cd: PNMLVCustomDraw;
begin
  if Message.NMHdr.code = NM_CUSTOMDRAW then
  begin
    Cd := PNMLVCustomDraw(Message.NMHdr);
    // Group headers get their own PREPAINT cycle with
    // dwItemType = LVCDI_GROUP (empirically verified; they never arrive
    // as ITEMPREPAINT, and the native painter ignores clrText) - draw
    // them ourselves in theme-aware colors.
    if (Cd.nmcd.dwDrawStage = CDDS_PREPAINT)
      and (Cd.dwItemType = LVCDI_GROUP) then
    begin
      var Grp: TLVGroup;
      var Buf: array[0..511] of Char;
      FillChar(Grp, SizeOf(Grp), 0);
      FillChar(Buf, SizeOf(Buf), 0);
      Grp.cbSize := SizeOf(Grp);
      Grp.mask := LVGF_HEADER;
      Grp.pszHeader := @Buf[0];
      Grp.cchHeader := Length(Buf);
      SendMessage(Handle, LVM_GETGROUPINFO, Cd.nmcd.dwItemSpec, LPARAM(@Grp));

      var R: TRect;
      R.Top := LVGGR_HEADER;
      R.Left := 0; R.Right := 0; R.Bottom := 0;
      SendMessage(Handle, LVM_GETGROUPRECT, Cd.nmcd.dwItemSpec, LPARAM(@R));

      var Hdc := Cd.nmcd.hdc;
      var Br := CreateSolidBrush(ColorToRGB(GetThemedColor(clWindow)));
      FillRect(Hdc, R, Br);
      DeleteObject(Br);
      var OldFont := SelectObject(Hdc, Font.Handle);
      SetBkMode(Hdc, TRANSPARENT);
      SetTextColor(Hdc, ColorToRGB(GetThemedColor(clWindowText)));
      var TxtR := R;
      Inc(TxtR.Left, 8);
      var Txt: string := Grp.pszHeader;
      DrawText(Hdc, PChar(Txt), Length(Txt), TxtR,
        DT_SINGLELINE or DT_VCENTER or DT_LEFT or DT_END_ELLIPSIS);
      // Separator line from the text to the right edge, like the native look.
      var Sz: TSize;
      GetTextExtentPoint32(Hdc, PChar(Txt), Length(Txt), Sz);
      var Pen := CreatePen(PS_SOLID, 1, ColorToRGB(GetThemedColor(clGrayText)));
      var OldPen := SelectObject(Hdc, Pen);
      var Y := (R.Top + R.Bottom) div 2;
      MoveToEx(Hdc, TxtR.Left + Sz.cx + 8, Y, nil);
      LineTo(Hdc, R.Right - 4, Y);
      SelectObject(Hdc, OldPen);
      DeleteObject(Pen);
      SelectObject(Hdc, OldFont);

      Message.Result := CDRF_SKIPDEFAULT;
      Exit;
    end;
  end;
  inherited;
end;

constructor TDfmEventCheckDialog.CreateDialog(AOwner: TComponent;
  const AIssues: TArray<TDfmEventIssue>);
var
  Col: TListColumn;
  MissCount, SigCount, N: Integer;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  Caption := 'DFM event handler check';
  Width := 1000;
  Height := 560;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  // Non-modal: fixing the findings requires editing while the list
  // stays open.
  OnClose := DoFormClose;

  FIssues := AIssues;

  // Pre-split the normalized signatures once so painting the Details
  // column never allocates (this list is custom-drawn on every scroll).
  SetLength(FExpTok, Length(FIssues));
  SetLength(FActTok, Length(FIssues));
  SetLength(FChecked, Length(FIssues));
  for N := 0 to High(FIssues) do
  begin
    FExpTok[N] := FIssues[N].ExpectedNorm.Split(['|']);
    FActTok[N] := FIssues[N].ActualNorm.Split(['|']);
    // Signature mismatches are corrected in place and PRE-TICKED; missing
    // handlers are fixable but NOT pre-ticked - on inherited forms the
    // "missing" handler may exist in the (out-of-project) ancestor unit,
    // and generating an empty override would silently disable it.
    FChecked[N] := (FIssues[N].Kind = eikSignatureMismatch)
      and (FIssues[N].ExpectedRawParams <> '');
  end;

  FLblSummary := TLabel.Create(Self);
  FLblSummary.Parent := Self;
  FLblSummary.Align := alTop;
  FLblSummary.AlignWithMargins := True;
  FLblSummary.Margins.SetBounds(8, 8, 8, 4);

  FLblStatus := TLabel.Create(Self);
  FLblStatus.Parent := Self;
  FLblStatus.Align := alTop;
  FLblStatus.AlignWithMargins := True;
  FLblStatus.Margins.SetBounds(8, 0, 8, 4);
  FLblStatus.EllipsisPosition := epPathEllipsis;
  FLblStatus.Font.Style := [fsBold];
  FLblStatus.Caption := '';

  FEdtFilter := TEdit.Create(Self);
  FEdtFilter.Parent := Self;
  FEdtFilter.Top := FLblStatus.Top + FLblStatus.Height + 1;  // below the labels
  FEdtFilter.Align := alTop;
  FEdtFilter.AlignWithMargins := True;
  FEdtFilter.Margins.SetBounds(8, 0, 8, 4);
  FEdtFilter.TextHint := 'Filter (form / component / event / handler / details)...';
  FEdtFilter.OnChange := DoFilterChange;

  FListView := TBufferedListView.Create(Self);
  FListView.Parent := Self;
  FListView.Align := alClient;
  FListView.AlignWithMargins := True;
  FListView.Margins.SetBounds(8, 4, 8, 4);
  FListView.ViewStyle := vsReport;
  FListView.ReadOnly := True;
  FListView.RowSelect := True;
  FListView.Checkboxes := True;   // tick the rows to auto-fix
  FListView.OnAdvancedCustomDrawSubItem := DoAdvancedDrawSubItem;
  FListView.OnDblClick := DoDblClick;
  FListView.OnItemChecked := DoItemChecked;

  Col := FListView.Columns.Add; Col.Caption := 'Problem';   Col.Width := 140;
  Col := FListView.Columns.Add; Col.Caption := 'Auto-fix';  Col.Width := 56;
    Col.Alignment := taCenter;
  Col := FListView.Columns.Add; Col.Caption := 'Form';      Col.Width := 165;
  Col := FListView.Columns.Add; Col.Caption := 'Component'; Col.Width := 150;
  Col := FListView.Columns.Add; Col.Caption := 'Event';     Col.Width := 110;
  Col := FListView.Columns.Add; Col.Caption := 'Handler';   Col.Width := 150;
  Col := FListView.Columns.Add; Col.Caption := 'Details';   Col.Width := 320;

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
  FBtnGotoType := TButton.Create(Self);
  FBtnGotoType.Parent := Panel;
  FBtnGotoType.Caption := 'Go to event &type';
  FBtnGotoType.Align := alRight;
  FBtnGotoType.Width := 130;
  FBtnGotoType.AlignWithMargins := True;
  FBtnGotoType.OnClick := DoGotoTypeClick;
  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Align := alRight;
  FBtnClose.AlignWithMargins := True;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  // Auto-fix controls on the left.
  FBtnFix := TButton.Create(Self);
  FBtnFix.Parent := Panel;
  FBtnFix.Caption := 'Fix &checked signatures';
  FBtnFix.Align := alLeft;
  FBtnFix.Width := 170;
  FBtnFix.AlignWithMargins := True;
  FBtnFix.OnClick := DoFixClick;
  FChkIde := TCheckBox.Create(Self);
  FChkIde.Parent := Panel;
  FChkIde.Caption := 'Apply via IDE (open form + save)';
  FChkIde.Align := alLeft;
  FChkIde.Width := 230;
  FChkIde.AlignWithMargins := True;
  FChkIde.ShowHint := True;
  FChkIde.Hint :=
    'Off (recommended): the corrected .pas files are written straight to disk.'#13#10 +
    'On: each corrected form is opened and saved so the IDE re-streams the DFM.'#13#10 +
    'Note: saving makes the IDE re-validate ALL event bindings on the form, and'#13#10 +
    'its validator mishandles IFDEF''d parameter types - it may then ask to remove'#13#10 +
    'a reference for a handler that is actually correct. Answer "No" in that case.';
  FChkAliasDiff := TCheckBox.Create(Self);
  FChkAliasDiff.Parent := Panel;
  FChkAliasDiff.Caption := 'Ignore Integer/LongInt differences';
  FChkAliasDiff.Align := alLeft;
  FChkAliasDiff.Width := 230;
  FChkAliasDiff.AlignWithMargins := True;
  FChkAliasDiff.Checked := False;   // show alias differences (bold) by default
  FChkAliasDiff.OnClick := DoAliasDiffClick;

  FillRows;
  EnableListViewSorting(FListView);

  MissCount := 0; SigCount := 0;
  for N := 0 to High(FIssues) do
    if FIssues[N].Kind = eikMissingHandler then Inc(MissCount)
    else Inc(SigCount);
  FLblSummary.Caption := Format(
    '%d issue(s): %d missing handler(s), %d signature mismatch(es).',
    [Length(FIssues), MissCount, SigCount]);

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

// (Re)builds groups + rows for the issues matching the filter box. Rows
// carry the ISSUE index in Item.Data (the list is sortable - Item.Index
// is meaningless for mapping); check states live in FChecked.
procedure TDfmEventCheckDialog.FillRows;
var
  Issue: TDfmEventIssue;
  Item: TListItem;
  N, IssueIdx: Integer;
  GroupCounts, GroupIds: TDictionary<string, Integer>;
  GroupKeys: TArray<string>;
  Key, Filter: string;
  Grp: TListGroup;

  function Matches(const AIssue: TDfmEventIssue): Boolean;
  begin
    Result := (Filter = '')
      or (Pos(Filter, UpperCase(ExtractFileName(AIssue.DfmFile))) > 0)
      or (Pos(Filter, UpperCase(AIssue.ComponentName)) > 0)
      or (Pos(Filter, UpperCase(AIssue.ComponentType)) > 0)
      or (Pos(Filter, UpperCase(AIssue.EventName)) > 0)
      or (Pos(Filter, UpperCase(AIssue.HandlerName)) > 0)
      or (Pos(Filter, UpperCase(AIssue.Expected)) > 0)
      or (Pos(Filter, UpperCase(AIssue.Actual)) > 0);
  end;

begin
  Filter := UpperCase(Trim(FEdtFilter.Text));
  FFillingRows := True;
  try
    // Group identical event types together. Build the group headers first
    // (sorted, with per-group counts over the FILTERED rows) so rows can
    // be assigned to them.
    GroupCounts := TDictionary<string, Integer>.Create;
    GroupIds := TDictionary<string, Integer>.Create;
    try
      for Issue in FIssues do
      begin
        if not Matches(Issue) then Continue;
        Key := GroupKeyOf(Issue);
        GroupCounts.TryGetValue(Key, N);
        GroupCounts.AddOrSetValue(Key, N + 1);
      end;
      GroupKeys := GroupCounts.Keys.ToArray;
      TArray.Sort<string>(GroupKeys, TComparer<string>.Construct(
        function(const L, R: string): Integer
        begin
          Result := CompareText(L, R);
        end));

      FListView.GroupView := True;
      // Force the window handle to exist before adding groups and ticking
      // rows. On the FIRST handle creation VCL only replays check states
      // when a saved stream exists (see TCustomListView.CreateWnd), so a
      // pre-tick done while the handle is unallocated is cached but never
      // written to the native control - and Checked would then read back
      // False. With the handle live, Item.Checked writes through.
      FListView.HandleNeeded;
      FListView.Items.BeginUpdate;
      try
        FListView.Items.Clear;
        FListView.Groups.Clear;
        for Key in GroupKeys do
        begin
          Grp := FListView.Groups.Add;
          Grp.Header := Format('%s  (%d)', [Key, GroupCounts[Key]]);
          GroupIds.Add(Key, Grp.GroupID);
        end;

        for IssueIdx := 0 to High(FIssues) do
        begin
          Issue := FIssues[IssueIdx];
          if not Matches(Issue) then Continue;
          Item := FListView.Items.Add;
          Item.Data := Pointer(NativeInt(IssueIdx));
          Item.GroupID := GroupIds[GroupKeyOf(Issue)];
          case Issue.Kind of
            eikMissingHandler:    Item.Caption := 'Missing handler';
            eikSignatureMismatch: Item.Caption := 'Signature mismatch';
          end;
          // Fix column: check-mark for auto-fixable rows.
          if Issue.ExpectedRawParams <> '' then
            Item.SubItems.Add(#$2713)   // check mark
          else
            Item.SubItems.Add('');
          Item.Checked := FChecked[IssueIdx];
          Item.SubItems.Add(ExtractFileName(Issue.DfmFile));
          Item.SubItems.Add(Issue.ComponentName + ': ' + Issue.ComponentType);
          Item.SubItems.Add(Issue.EventName);
          Item.SubItems.Add(Issue.HandlerName);
          if Issue.Kind = eikSignatureMismatch then
            Item.SubItems.Add('expected ' + Issue.Expected + ', found ' + Issue.Actual)
          else if Issue.ExpectedRawParams <> '' then
            Item.SubItems.Add('crashes with "Method not found" at form load - ' +
              'fix generates an empty handler (' + Issue.ExpectedRawParams + ')')
          else
            Item.SubItems.Add('crashes with "Method not found" at form load');
        end;
      finally
        FListView.Items.EndUpdate;
      end;
    finally
      GroupCounts.Free;
      GroupIds.Free;
    end;
  finally
    FFillingRows := False;
  end;
end;

procedure TDfmEventCheckDialog.DoFilterChange(Sender: TObject);
begin
  FillRows;
end;

procedure TDfmEventCheckDialog.DoItemChecked(Sender: TObject; Item: TListItem);
begin
  if FFillingRows or (Item = nil) then Exit;
  var Idx := NativeInt(Item.Data);
  if (Idx >= 0) and (Idx <= High(FChecked)) then
    FChecked[Idx] := Item.Checked;
end;

// Canonical name for interchangeable integer types, so e.g. LongInt and
// Integer (both 32-bit) are treated as the same type. Only the bare type
// name is folded; parameter modifiers (var/const/out) stay intact.
function CanonType(const S: string): string;
begin
  Result := UpperCase(Trim(S));
  if (Result = 'INTEGER') or (Result = 'LONGINT') or (Result = 'FIXEDINT') then
    Result := 'INTEGER'
  else if (Result = 'CARDINAL') or (Result = 'LONGWORD') or (Result = 'FIXEDUINT')
       or (Result = 'DWORD') then
    Result := 'CARDINAL';
end;

// Two parameter types match when they are literally equal, or (unless the
// user opted to see them) when they are interchangeable integer aliases.
function TypesMatch(const A, B: string; AFoldAliases: Boolean): Boolean;
begin
  Result := SameText(Trim(A), Trim(B));
  if (not Result) and AFoldAliases then
    Result := CanonType(A) = CanonType(B);
end;

// Draw one text segment at AX/ATop, then advance AX by its width. Bold
// marks a differing parameter type. Text is clipped by the caller's DC
// clip region so it never bleeds into neighbouring columns.
procedure DrawSeg(ACanvas: TCanvas; var AX: Integer; ATop: Integer;
  const AText: string; ABold: Boolean);
begin
  // Only touch Font.Style when the weight actually changes: each
  // assignment forces GDI to select a fresh font handle, which is the
  // dominant cost when painting many segments per row while scrolling.
  if ABold <> (fsBold in ACanvas.Font.Style) then
    if ABold then
      ACanvas.Font.Style := [fsBold]
    else
      ACanvas.Font.Style := [];
  ACanvas.TextOut(AX, ATop, AText);
  Inc(AX, ACanvas.TextWidth(AText));
end;

procedure TDfmEventCheckDialog.DoAdvancedDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  Stage: TCustomDrawStage; var DefaultDraw: Boolean);
const
  DetailsCol = 6;
var
  Issue: TDfmEventIssue;
  Exp, Act: TArray<string>;
  R: TRect;
  X, Top, Saved: Integer;
  I: Integer;
  Diff, Fold: Boolean;
  Cv: TCanvas;
begin
  DefaultDraw := True;
  if Stage <> cdPrePaint then Exit;
  if SubItem <> DetailsCol then Exit;
  var IssueIdx := NativeInt(Item.Data);   // sortable list - never Item.Index
  if (IssueIdx < 0) or (IssueIdx >= Length(FIssues)) then Exit;
  Issue := FIssues[IssueIdx];
  if (Issue.Kind <> eikSignatureMismatch) or (Issue.ExpectedNorm = '') then Exit;

  Exp := FExpTok[IssueIdx];
  Act := FActTok[IssueIdx];

  // Bounds of the Details subitem cell.
  R.Left := LVIR_LABEL;
  R.Top := SubItem;
  if SendMessage(FListView.Handle, LVM_GETSUBITEMRECT, Item.Index, LPARAM(@R)) = 0 then
    Exit;

  Cv := Sender.Canvas;
  if Item.Selected then
  begin
    Cv.Brush.Color := GetThemedColor(clHighlight);
    Cv.Font.Color := GetThemedColor(clHighlightText);
  end
  else
  begin
    Cv.Brush.Color := GetThemedColor(clWindow);
    Cv.Font.Color := GetThemedColor(clWindowText);
  end;
  Cv.FillRect(R);
  SetBkMode(Cv.Handle, TRANSPARENT);

  // Clip everything we draw to the cell so long signatures don't spill
  // past the Details column.
  Saved := SaveDC(Cv.Handle);
  try
    IntersectClipRect(Cv.Handle, R.Left, R.Top, R.Right, R.Bottom);
    X := R.Left + 4;
    Top := R.Top + (R.Bottom - R.Top - Cv.TextHeight('Wg')) div 2;

    // Fold integer aliases while "Ignore Integer/LongInt differences" is on.
    Fold := FChkAliasDiff.Checked;
    DrawSeg(Cv, X, Top, 'expected (', False);
    for I := 0 to High(Exp) do
    begin
      Diff := (I > High(Act)) or not TypesMatch(Exp[I], Act[I], Fold);
      DrawSeg(Cv, X, Top, Exp[I], Diff);
      if I < High(Exp) then DrawSeg(Cv, X, Top, '; ', False);
    end;
    DrawSeg(Cv, X, Top, '), found (', False);
    for I := 0 to High(Act) do
    begin
      Diff := (I > High(Exp)) or not TypesMatch(Exp[I], Act[I], Fold);
      DrawSeg(Cv, X, Top, Act[I], Diff);
      if I < High(Act) then DrawSeg(Cv, X, Top, '; ', False);
    end;
    DrawSeg(Cv, X, Top, ')', False);
  finally
    RestoreDC(Cv.Handle, Saved);
  end;
  Cv.Font.Style := [];

  DefaultDraw := False;
end;

procedure TDfmEventCheckDialog.DoAliasDiffClick(Sender: TObject);
begin
  FListView.Invalidate;   // re-evaluate the bold diffs
end;

// Group key: identical event types share one group. Handlers with no
// resolved event type fall back to the event name.
function TDfmEventCheckDialog.GroupKeyOf(const AIssue: TDfmEventIssue): string;
begin
  if AIssue.EventTypeName <> '' then
    Result := AIssue.EventTypeName
  else if AIssue.Kind = eikMissingHandler then
    Result := 'Missing handler: ' + AIssue.EventName
  else
    Result := AIssue.EventName;
end;

// 0-based column of AName on 1-based line ALine1 of AFile (case-insensitive),
// so the editor highlights the handler name itself rather than a fixed-width
// span from column 0 (which would land on "procedure ..." instead).
function TDfmEventCheckDialog.NameColumn(const AFile: string; ALine1: Integer;
  const AName: string): Integer;
var
  Content, Line: string;
  Lines: TArray<string>;
  P: Integer;
begin
  Result := 0;
  if (AName = '') or (ALine1 <= 0) then Exit;
  if (Editor = nil) or not Editor.ReadEditorContent(AFile, Content) then
  begin
    if not TFile.Exists(AFile) then Exit;
    try Content := TFile.ReadAllText(AFile); except Exit; end;
  end;
  Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  if (ALine1 - 1) > High(Lines) then Exit;
  Line := Lines[ALine1 - 1];
  P := Pos(UpperCase(AName), UpperCase(Line));
  if P > 0 then Result := P - 1;
end;

procedure TDfmEventCheckDialog.GotoSelected;
var
  Idx, Col: Integer;
begin
  if FListView.Selected = nil then Exit;
  Idx := NativeInt(FListView.Selected.Data);
  if (Idx < 0) or (Idx >= Length(FIssues)) then Exit;
  // Signature mismatch -> jump to the .pas declaration; missing
  // handler -> jump to the offending .dfm line so the user sees which
  // event reference to remove (or which method to create).
  if (FIssues[Idx].Kind = eikSignatureMismatch) and (FIssues[Idx].PasLine > 0) then
  begin
    Col := NameColumn(FIssues[Idx].PasFile, FIssues[Idx].PasLine,
      FIssues[Idx].HandlerName);
    Editor.GotoLocation(FIssues[Idx].PasFile, FIssues[Idx].PasLine - 1, Col,
      Length(FIssues[Idx].HandlerName));
  end
  else
  begin
    Col := NameColumn(FIssues[Idx].DfmFile, FIssues[Idx].DfmLine,
      FIssues[Idx].HandlerName);
    Editor.GotoLocation(FIssues[Idx].DfmFile, FIssues[Idx].DfmLine - 1, Col,
      Length(FIssues[Idx].HandlerName));
  end;
end;

procedure TDfmEventCheckDialog.DoDblClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TDfmEventCheckDialog.DoGotoClick(Sender: TObject);
begin
  GotoSelected;
end;

procedure TDfmEventCheckDialog.DoGotoTypeClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FListView.Selected = nil then Exit;
  Idx := NativeInt(FListView.Selected.Data);
  if (Idx < 0) or (Idx >= Length(FIssues)) then Exit;
  if (FIssues[Idx].EventTypeFile = '') or (FIssues[Idx].EventTypeLine <= 0) then
  begin
    ShowThemedMessage('The event type location is only known for mismatches ' +
      'resolved from source. This row has no resolved event type.');
    Exit;
  end;
  Editor.GotoLocation(FIssues[Idx].EventTypeFile,
    FIssues[Idx].EventTypeLine - 1, 0, Length(FIssues[Idx].EventTypeName));
end;

procedure TDfmEventCheckDialog.DoFixClick(Sender: TObject);
var
  I, CheckedCount, Fixed, NotFixable, Failed, Total, Done: Integer;
  Forms: TStringList;   // affected .pas files (for the IDE open+save step)
  Reason: string;
  FirstReasons: TStringList;   // a few example failure reasons for the user
  Ctx: TFixContext;     // shared LSP client + type->unit cache for the batch
  Prog: TCheckProgressWindow;
begin
  CheckedCount := 0; Fixed := 0; NotFixable := 0; Failed := 0; Done := 0;
  Total := 0;
  for I := 0 to FListView.Items.Count - 1 do
    if FListView.Items[I].Checked then Inc(Total);
  if Total = 0 then
  begin
    ShowThemedMessage('No rows checked. Tick the signature mismatches you want to correct.');
    Exit;
  end;

  Forms := TStringList.Create;
  FirstReasons := TStringList.Create;
  Ctx := TFixContext.Create;

  // Progress window: applying fixes runs on the main thread (LSP calls,
  // file writes) - without it the IDE looks frozen on large batches.
  Prog := CreateCheckProgress('Applying signature fixes', Self, 'Applying...');

  try
    Forms.Duplicates := dupIgnore;
    Forms.Sorted := True;
    Screen.Cursor := crHourGlass;
    try
      try
      // Rows map to issues through Item.Data (sortable/filterable list).
      for I := 0 to FListView.Items.Count - 1 do
      begin
        if not FListView.Items[I].Checked then Continue;
        var Idx := NativeInt(FListView.Items[I].Data);
        Inc(CheckedCount);
        Inc(Done);
        if (Idx >= 0) and (Idx < Length(FIssues)) then
          Prog.Step(Done, Total, ExtractFileName(FIssues[Idx].PasFile))
        else
          Prog.Step(Done, Total, '');
        if (Idx < 0) or (Idx >= Length(FIssues)) then Continue;
        if FIssues[Idx].ExpectedRawParams = '' then
        begin
          Inc(NotFixable);
          Continue;
        end;
        if TDfmEventChecker.ApplyFix(FIssues[Idx], Reason, Ctx) then
        begin
          Inc(Fixed);
          Forms.Add(FIssues[Idx].PasFile);
          FListView.Items[I].SubItems[FListView.Items[I].SubItems.Count - 1] :=
            'FIXED -> (' + FIssues[Idx].ExpectedRawParams + ')';
          FListView.Items[I].Checked := False;
          FChecked[Idx] := False;
        end
        else
        begin
          Inc(Failed);   // checked + fixable, but ApplyFix could not apply
          if FirstReasons.Count < 4 then FirstReasons.Add(Reason);
        end;
      end;
      finally
        Prog.Hide;   // the progress popup only covers the fix loop
      end;

      // Optional IDE apply phase: open each corrected form and save it so
      // the IDE re-streams the DFM. The current unit is shown in the
      // (non-topmost) main window - Delphi's own "remove reference?"
      // prompt names only the event method, not the unit, so this lets
      // the user tell which form each prompt belongs to.
      if FChkIde.Checked and (Forms.Count > 0) then
      begin
        for var F in Forms.ToStringArray do
        begin
          FLblStatus.Caption := 'Saving via IDE: ' + ExtractFileName(F) +
            '   (' + F + ')';
          FLblStatus.Update;
          Application.ProcessMessages;
          Editor.GotoLocation(F, 0, 0, 0);
          Editor.SaveFile(F);
        end;
        FLblStatus.Caption := '';
      end;
    finally
      Screen.Cursor := crDefault;
    end;

    FListView.Invalidate;
    if CheckedCount = 0 then
      ShowThemedMessage('No rows checked. Tick the signature mismatches you want to correct.')
    else
      ShowThemedMessage(Format('%d signature(s) corrected (of %d checked).%s%s',
        [Fixed, CheckedCount,
         IfThen(NotFixable > 0,
           sLineBreak + Format('%d checked row(s) are not auto-fixable ' +
             '(missing handler, or signature not resolvable from source).',
             [NotFixable]),
           ''),
         IfThen(Failed > 0,
           sLineBreak + Format('%d checked row(s) could not be applied. Examples:',
             [Failed]) + sLineBreak + '  ' +
             StringReplace(FirstReasons.Text, sLineBreak, sLineBreak + '  ',
               [rfReplaceAll]),
           '')]));
  finally
    Forms.Free;
    FirstReasons.Free;
    Ctx.Free;
    Prog.Free;
  end;
end;

procedure TDfmEventCheckDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TDfmEventCheckDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

/// <summary>Highest-installed BDS root directory (via registry), '' if
///  none found. Used to locate global library/browsing paths and to
///  exclude the built-in Studio source tree from scanning.</summary>
function FindBdsRoot: string;

  function TryHive(ARootKey: HKEY): string;
  var
    Reg: TRegistry;
    Versions: TStringList;
    V, Best: string;
    BestNum, Num: Double;
    FS: TFormatSettings;
  begin
    Result := '';
    Best := ''; BestNum := -1;
    FS := TFormatSettings.Invariant;
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := ARootKey;
      if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS') then
      begin
        Versions := TStringList.Create;
        try
          Reg.GetKeyNames(Versions);
          for V in Versions do
            if TryStrToFloat(V, Num, FS) and (Num > BestNum) then
            begin BestNum := Num; Best := V; end;
        finally
          Versions.Free;
        end;
        Reg.CloseKey;
      end;
      if (Best <> '')
         and Reg.OpenKeyReadOnly('Software\Embarcadero\BDS\' + Best) then
      begin
        if Reg.ValueExists('RootDir') then
          Result := Reg.ReadString('RootDir');
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
  end;

begin
  Result := TryHive(HKEY_CURRENT_USER);
  if Result = '' then Result := TryHive(HKEY_LOCAL_MACHINE);
end;

/// <summary>Reads the global Library "Search Path" + "Browsing Path"
///  for Win32 from the highest BDS, expands the common macros, and
///  returns the existing directories that are NOT under the BDS root
///  (the built-in RTL/VCL/FMX source is huge and already covered by
///  the signature table, so we skip it).</summary>
function GatherGlobalLibraryDirs(const ABdsRoot: string): TArray<string>;

  function ReadPath(const ASubKey, AValue: string): string;
  var Reg: TRegistry;
  begin
    Result := '';
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly(ASubKey) and Reg.ValueExists(AValue) then
        Result := Reg.ReadString(AValue);
      Reg.CloseKey;
    finally
      Reg.Free;
    end;
  end;

  function Expand(const S: string): string;
  begin
    Result := S;
    Result := StringReplace(Result, '$(BDSLIB)',
      IncludeTrailingPathDelimiter(ABdsRoot) + 'lib', [rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result, '$(BDSINCLUDE)',
      IncludeTrailingPathDelimiter(ABdsRoot) + 'include', [rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result, '$(BDS)', ABdsRoot, [rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result, '$(Platform)', 'Win32', [rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result, '$(Config)', 'Release', [rfReplaceAll, rfIgnoreCase]);
  end;

var
  Raw, Dir: string;
  Dirs: TList<string>;
  RootPref: string;
begin
  Result := nil;
  if ABdsRoot = '' then Exit;
  RootPref := IncludeTrailingPathDelimiter(ABdsRoot);
  Dirs := TList<string>.Create;
  try
    // Locate the versioned Library\Win32 subkey again.
    var Reg := TRegistry.Create(KEY_READ);
    var VerKey := '';
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS') then
      begin
        var Vers := TStringList.Create;
        try
          Reg.GetKeyNames(Vers);
          var BestNum: Double := -1; var FS := TFormatSettings.Invariant;
          for var V in Vers do
          begin
            var N: Double;
            if TryStrToFloat(V, N, FS) and (N > BestNum) then
            begin BestNum := N; VerKey := V; end;
          end;
        finally
          Vers.Free;
        end;
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
    if VerKey = '' then Exit;

    for Raw in [ReadPath('Software\Embarcadero\BDS\' + VerKey + '\Library\Win32', 'Search Path'),
                ReadPath('Software\Embarcadero\BDS\' + VerKey + '\Library\Win32', 'Browsing Path')] do
      for Dir in Raw.Split([';'], TStringSplitOptions.ExcludeEmpty) do
      begin
        var D := Expand(Trim(Dir));
        if Pos('$(', D) > 0 then Continue;                  // unresolved macro
        D := ExcludeTrailingPathDelimiter(D);
        if not TDirectory.Exists(D) then Continue;
        if StartsText(RootPref, IncludeTrailingPathDelimiter(D)) then Continue; // built-in
        if not Dirs.Contains(D) then Dirs.Add(D);
      end;
    Result := Dirs.ToArray;
  finally
    Dirs.Free;
  end;
end;

/// <summary>Reads DCC_UnitSearchPath directories from the .dproj,
///  resolved relative to the .dproj directory.</summary>
function GatherDprojSearchDirs: TArray<string>;
var
  Dproj, Content, Inner, DprojDir, Ent, D: string;
  Dirs: TList<string>;
  P1, P2: Integer;
begin
  Result := nil;
  Dproj := Editor.GetCurrentProjectDproj;
  if (Dproj = '') or not TFile.Exists(Dproj) then Exit;
  DprojDir := ExtractFilePath(Dproj);
  try Content := TFile.ReadAllText(Dproj); except Exit; end;
  Dirs := TList<string>.Create;
  try
    // Union of every <DCC_UnitSearchPath>...</DCC_UnitSearchPath> block.
    P1 := Pos('<DCC_UnitSearchPath>', Content);
    while P1 > 0 do
    begin
      P2 := PosEx('</DCC_UnitSearchPath>', Content, P1);
      if P2 = 0 then Break;
      Inner := Copy(Content, P1 + Length('<DCC_UnitSearchPath>'),
                    P2 - P1 - Length('<DCC_UnitSearchPath>'));
      for Ent in Inner.Split([';'], TStringSplitOptions.ExcludeEmpty) do
      begin
        var E := Trim(Ent);
        if (E = '') or E.StartsWith('$(') then Continue;   // skip macro roots
        if TPath.IsRelativePath(E) then
          D := TPath.GetFullPath(TPath.Combine(DprojDir, E))
        else
          D := E;
        D := ExcludeTrailingPathDelimiter(D);
        if TDirectory.Exists(D) and not Dirs.Contains(D) then
          Dirs.Add(D);
      end;
      P1 := PosEx('<DCC_UnitSearchPath>', Content, P2);
    end;
    Result := Dirs.ToArray;
  finally
    Dirs.Free;
  end;
end;

/// <summary>Gathers component-source .pas from every search path a
///  general project can reach: the project's own search paths, the
///  .dproj DCC_UnitSearchPath, and the IDE's global library / browsing
///  paths (minus the built-in Studio source tree). Used to resolve
///  third-party event signatures (VirtualTrees, ...) from their real
///  .pas - version-correct, no hardcoded tables.</summary>
function GatherSignatureFiles: TArray<string>;
var
  AllDirs: TStringList;
  BdsRoot, Dir, F: string;

  procedure AddDirs(const ADirs: TArray<string>);
  var D: string;
  begin
    for D in ADirs do
      if D <> '' then AllDirs.Add(ExcludeTrailingPathDelimiter(D));
  end;

const
  MaxSigFiles = 6000;   // 32-bit IDE process guard
var
  Files: TStringList;
begin
  Result := nil;
  AllDirs := TStringList.Create;
  Files := TStringList.Create;
  try
    AllDirs.Duplicates := dupIgnore;
    AllDirs.Sorted := True;
    // Project dirs + the IDE global library/browsing path (where visual
    // components with DFM events live: VirtualTrees, TRichView, ...).
    // The .dproj DCC_UnitSearchPath is DELIBERATELY excluded: it holds
    // non-visual compile libraries (mORMot, FHIR - thousands of files)
    // that never declare DFM component events, and scanning them would
    // exhaust the 32-bit IDE process.
    AddDirs(Editor.GetProjectSearchPaths.Split([';'], TStringSplitOptions.ExcludeEmpty));
    BdsRoot := FindBdsRoot;
    AddDirs(GatherGlobalLibraryDirs(BdsRoot));

    Files.Duplicates := dupIgnore;
    Files.Sorted := True;
    for Dir in AllDirs do
    begin
      if Files.Count >= MaxSigFiles then Break;
      if TDirectory.Exists(Dir) then
        try
          for F in TDirectory.GetFiles(Dir, '*.pas') do
          begin
            Files.Add(F);
            if Files.Count >= MaxSigFiles then Break;
          end;
        except
          // unreadable dir - skip
        end;
    end;
    Result := Files.ToStringArray;
  finally
    Files.Free;
    AllDirs.Free;
  end;
end;

procedure CheckDfmEventHandlers;
var
  Files, SigFiles: TArray<string>;
  Issues: TArray<TDfmEventIssue>;
  Dlg: TDfmEventCheckDialog;
  ProgressForm: TCheckProgressWindow;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowThemedMessage('No project loaded / no source files found.');
    Exit;
  end;
  SigFiles := GatherSignatureFiles;

  // Progress window: the check reads every project file once for the
  // class index, then walks each form - visible feedback instead of a
  // silent multi-second hang on large projects.
  ProgressForm := CreateCheckProgress('DFM event handler check', nil);
  try
    Screen.Cursor := crHourGlass;
    try
      Issues := TDfmEventChecker.CheckProject(Files,
        procedure(ACurrent, ATotal: Integer; AFile: string)
        begin
          ProgressForm.Step(ACurrent, ATotal, ExtractFileName(AFile));
        end,
        SigFiles);

      // Build the results dialog while the progress window is STILL up
      // (populating a big list can take a moment - no silent gap).
      Dlg := nil;
      if Length(Issues) > 0 then
      begin
        ProgressForm.Bar.Position := ProgressForm.Bar.Max;
        ProgressForm.Step(0, 0, Format('Building list (%d issues)...', [Length(Issues)]));
        // Non-modal (frees itself on close) so the user can fix findings
        // in the editor while the list stays open.
        Dlg := TDfmEventCheckDialog.CreateDialog(Application.MainForm, Issues);
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    ProgressForm.Free;
  end;

  if Length(Issues) = 0 then
    ShowThemedMessage('No problems found - every DFM event reference has a matching handler.')
  else
    Dlg.Show;
end;

end.
