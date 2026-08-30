(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.CircularRefsDialog;

// Circular-unit-reference viewer. Five tabs:
//
// "Cycles": master list of every uses edge inside a cycle group (top).
//   Selecting a row shows, below, the shortest concrete cycle running
//   through that edge - each hop is "UnitA uses UnitB (section, line)",
//   and the last hop closes back to the first unit, so the ring is
//   readable top-to-bottom. Columns: Group, Unit, Section, uses, Line,
//   Group size.
//
// "Groups": one row per strongly-connected group with its metrics -
//   Units in group, Shortest cycle (hop count), Max span (longest simple
//   cycle in the group), Edges. Double-click switches to the Cycles tab
//   and selects that group's first edge, so its shortest cycle is shown.
//
// "Edge levers": each unique (from -> to) uses dependency ranked by
//   "Units freed" = how many units drop out of ANY cycle if you remove
//   that one uses entry (Baseline - CountCycleNodes with the edge cut).
//   The top rows are the highest-impact single edges to break. Columns:
//   Units freed, Unit, uses, Section, Group.
//
// "All cycles": a plain-text, complete enumeration of every simple cycle
//   (UnitA -> UnitB -> ... -> UnitA), like the original plugin offers -
//   truncated with a note if the total is too large to list in full.
//
// "Hotspots": units ranked by how many cycle edges touch them
//   (Connections). The units at the top are involved in the most cycles
//   and are the best targets for breaking dependencies.
//
// In the edge/path lists interface-section edges render red,
// implementation edges orange. Double-click / "Go to" jumps to the
// exact uses entry (column computed so the used unit name is selected,
// not the line start).

interface

procedure CheckCircularReferences;

implementation

uses
  System.SysUtils, System.Classes, System.UITypes, System.IOUtils, System.StrUtils,
  System.Generics.Defaults, System.Generics.Collections,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.UsesGraph, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort;

type
  TCircularRefsDialog = class(TForm)
  private
    FResult: TUsesCycleResult;
    FHotspots: TArray<TUnitHotspot>;
    FGroups: TArray<TCycleGroupInfo>;
    FLevers: TArray<TEdgeLever>;
    FCurPath: TArray<TCycleHop>;
    FEdgeSortCol: Integer;    // -1 = unsorted (FEdgeList is virtual)
    FEdgeSortAsc: Boolean;

    FPages: TPageControl;
    FTabCycles: TTabSheet;
    FTabGroups: TTabSheet;
    FTabLevers: TTabSheet;
    FTabAll: TTabSheet;
    FTabHot: TTabSheet;
    FTabHelp: TTabSheet;
    FHelpMemo: TMemo;       // static help text
    FEdgeList: TListView;   // owner-data master (can be large)
    FPathList: TListView;   // filled per selection (small)
    FPathLbl: TLabel;
    FGroupList: TListView;  // filled once
    FLeverList: TListView;  // filled once
    FAllMemo: TMemo;        // filled once (text dump)
    FHotList: TListView;    // filled once
    FLblSummary: TLabel;
    FBtnClose: TButton;
    FBtnGoto: TButton;

    procedure BuildLayout;
    procedure FillHotspots;
    procedure FillGroups;
    procedure FillLevers;
    procedure FillAllCycles;
    procedure DoGroupDblClick(Sender: TObject);
    procedure DoLeverDblClick(Sender: TObject);
    procedure DoEdgeColumnClick(Sender: TObject; Column: TListColumn);
    procedure DoEdgeData(Sender: TObject; Item: TListItem);
    procedure DoEdgeSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure DoEdgeDraw(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoPathDraw(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoEdgeDblClick(Sender: TObject);
    procedure DoPathDblClick(Sender: TObject);
    procedure DoHotDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    /// <summary>0-based column of AUnit on 1-based line ALine1 of AFile, so
    ///  the editor highlights the used unit itself rather than column 0.</summary>
    function UnitColumn(const AFile: string; ALine1: Integer;
      const AUnit: string): Integer;
    procedure DoCloseClick(Sender: TObject);
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
  public
    constructor CreateDialog(AOwner: TComponent; AResult: TUsesCycleResult);
  end;

const
  clOrange = TColor($00008CFF);

constructor TCircularRefsDialog.CreateDialog(AOwner: TComponent;
  AResult: TUsesCycleResult);
var
  Groups, IntfEdges: Integer;
  E: TCycleEdge;
begin
  inherited CreateNew(AOwner);
  Caption := 'Circular unit references';
  Width := 1000;
  Height := 640;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  OnClose := DoFormClose;

  FResult := AResult;
  FHotspots := FResult.Hotspots;
  FGroups := FResult.GroupInfos;
  FLevers := FResult.EdgeLevers;
  FEdgeSortCol := -1;

  BuildLayout;

  FEdgeList.Items.Count := Length(FResult.Edges);
  FillGroups;
  FillLevers;
  FillAllCycles;
  FillHotspots;
  EnableListViewSorting(FGroupList);
  EnableListViewSorting(FLeverList);
  EnableListViewSorting(FHotList);

  Groups := 0; IntfEdges := 0;
  for E in FResult.Edges do
  begin
    if E.Group > Groups then Groups := E.Group;
    if E.InInterface then Inc(IntfEdges);
  end;
  FLblSummary.Caption := Format(
    '%d cycle group(s), %d edge(s), %d via interface uses.  ' +
    'Cycles tab: select an edge to see how it closes. Groups tab: shortest/longest cycle lengths. ' +
    'Hotspots tab: the most entangled units.',
    [Groups, Length(FResult.Edges), IntfEdges]);

  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TCircularRefsDialog.BuildLayout;
var
  Col: TListColumn;
  Panel, PathPanel: TPanel;
  Splitter: TSplitter;
begin
  FLblSummary := TLabel.Create(Self);
  FLblSummary.Parent := Self;
  FLblSummary.Align := alTop;
  FLblSummary.AlignWithMargins := True;
  FLblSummary.Margins.SetBounds(8, 8, 8, 4);

  FPages := TPageControl.Create(Self);
  FPages.Parent := Self;
  FPages.Align := alClient;
  FPages.AlignWithMargins := True;
  FPages.Margins.SetBounds(6, 2, 6, 2);

  FTabCycles := TTabSheet.Create(FPages);
  FTabCycles.PageControl := FPages;
  FTabCycles.Caption := 'Cycles';

  FTabGroups := TTabSheet.Create(FPages);
  FTabGroups.PageControl := FPages;
  FTabGroups.Caption := 'Groups';

  FTabLevers := TTabSheet.Create(FPages);
  FTabLevers.PageControl := FPages;
  FTabLevers.Caption := 'Edge levers';

  FTabAll := TTabSheet.Create(FPages);
  FTabAll.PageControl := FPages;
  FTabAll.Caption := 'All cycles';

  FTabHot := TTabSheet.Create(FPages);
  FTabHot.PageControl := FPages;
  FTabHot.Caption := 'Hotspots';

  FTabHelp := TTabSheet.Create(FPages);
  FTabHelp.PageControl := FPages;
  FTabHelp.Caption := 'Help';

  // ---- Cycles tab: master edge list (top) + path list (bottom) ----
  FEdgeList := TListView.Create(Self);
  FEdgeList.Parent := FTabCycles;
  FEdgeList.Align := alTop;
  FEdgeList.Height := 280;
  FEdgeList.ViewStyle := vsReport;
  FEdgeList.ReadOnly := True;
  FEdgeList.RowSelect := True;
  FEdgeList.OwnerData := True;
  FEdgeList.OnData := DoEdgeData;
  FEdgeList.OnSelectItem := DoEdgeSelect;
  FEdgeList.OnCustomDrawItem := DoEdgeDraw;
  FEdgeList.OnDblClick := DoEdgeDblClick;
  FEdgeList.OnColumnClick := DoEdgeColumnClick;
  Col := FEdgeList.Columns.Add; Col.Caption := 'Group';    Col.Width := 70;
  Col := FEdgeList.Columns.Add; Col.Caption := 'Unit';     Col.Width := 250;
  Col := FEdgeList.Columns.Add; Col.Caption := 'Section';  Col.Width := 110;
  Col := FEdgeList.Columns.Add; Col.Caption := 'uses';     Col.Width := 250;
  Col := FEdgeList.Columns.Add; Col.Caption := 'Line';     Col.Width := 60;
  Col := FEdgeList.Columns.Add; Col.Caption := 'Group size'; Col.Width := 80;

  Splitter := TSplitter.Create(Self);
  Splitter.Parent := FTabCycles;
  Splitter.Align := alTop;
  Splitter.Height := 5;

  PathPanel := TPanel.Create(Self);
  PathPanel.Parent := FTabCycles;
  PathPanel.Align := alClient;
  PathPanel.BevelOuter := bvNone;

  FPathLbl := TLabel.Create(Self);
  FPathLbl.Parent := PathPanel;
  FPathLbl.Align := alTop;
  FPathLbl.AlignWithMargins := True;
  FPathLbl.Margins.SetBounds(2, 4, 2, 2);
  FPathLbl.Caption := 'Select an edge above to see the cycle it closes.';

  FPathList := TListView.Create(Self);
  FPathList.Parent := PathPanel;
  FPathList.Align := alClient;
  FPathList.ViewStyle := vsReport;
  FPathList.ReadOnly := True;
  FPathList.RowSelect := True;
  FPathList.OnCustomDrawItem := DoPathDraw;
  FPathList.OnDblClick := DoPathDblClick;
  Col := FPathList.Columns.Add; Col.Caption := 'Step';    Col.Width := 50;
  Col := FPathList.Columns.Add; Col.Caption := 'Unit';    Col.Width := 250;
  Col := FPathList.Columns.Add; Col.Caption := 'Section'; Col.Width := 110;
  Col := FPathList.Columns.Add; Col.Caption := 'uses';    Col.Width := 250;
  Col := FPathList.Columns.Add; Col.Caption := 'Line';    Col.Width := 60;

  // ---- Groups tab ----
  var GroupInfo := TLabel.Create(Self);
  GroupInfo.Parent := FTabGroups;
  GroupInfo.Align := alTop;
  GroupInfo.AlignWithMargins := True;
  GroupInfo.Margins.SetBounds(4, 6, 4, 4);
  GroupInfo.Caption :=
    'Shortest cycle = exact girth (fewest units). Max span = units in the ' +
    'component (the largest a cycle could be; exact longest is NP-hard). ' +
    'Double-click jumps to the shortest cycle in the Cycles tab.';

  FGroupList := TListView.Create(Self);
  FGroupList.Parent := FTabGroups;
  FGroupList.Align := alClient;
  FGroupList.ViewStyle := vsReport;
  FGroupList.ReadOnly := True;
  FGroupList.RowSelect := True;
  FGroupList.OnDblClick := DoGroupDblClick;
  Col := FGroupList.Columns.Add; Col.Caption := 'Group';          Col.Width := 90;
  Col := FGroupList.Columns.Add; Col.Caption := 'Units in group'; Col.Width := 120;
  Col := FGroupList.Columns.Add; Col.Caption := 'Shortest cycle'; Col.Width := 120;
  Col := FGroupList.Columns.Add; Col.Caption := 'Max span';       Col.Width := 100;
  Col := FGroupList.Columns.Add; Col.Caption := 'Edges';          Col.Width := 80;

  // ---- Edge levers tab ----
  var LeverInfo := TLabel.Create(Self);
  LeverInfo.Parent := FTabLevers;
  LeverInfo.Align := alTop;
  LeverInfo.AlignWithMargins := True;
  LeverInfo.Margins.SetBounds(4, 6, 4, 4);
  LeverInfo.Caption :=
    'Each dependency ranked by how many units leave the cycle if you remove ' +
    'that one uses entry. The top rows are the biggest levers. Double-click to jump.';

  FLeverList := TListView.Create(Self);
  FLeverList.Parent := FTabLevers;
  FLeverList.Align := alClient;
  FLeverList.ViewStyle := vsReport;
  FLeverList.ReadOnly := True;
  FLeverList.RowSelect := True;
  FLeverList.OnDblClick := DoLeverDblClick;
  Col := FLeverList.Columns.Add; Col.Caption := 'Units freed'; Col.Width := 90;
  Col := FLeverList.Columns.Add; Col.Caption := 'Unit';        Col.Width := 250;
  Col := FLeverList.Columns.Add; Col.Caption := 'uses';        Col.Width := 250;
  Col := FLeverList.Columns.Add; Col.Caption := 'Section';     Col.Width := 110;
  Col := FLeverList.Columns.Add; Col.Caption := 'Group';       Col.Width := 70;

  // ---- All cycles tab (text dump) ----
  FAllMemo := TMemo.Create(Self);
  FAllMemo.Parent := FTabAll;
  FAllMemo.Align := alClient;
  FAllMemo.ReadOnly := True;
  FAllMemo.ScrollBars := ssBoth;
  FAllMemo.WordWrap := False;
  FAllMemo.Font.Name := 'Consolas';
  FAllMemo.Font.Size := 9;

  // ---- Hotspots tab ----
  FHotList := TListView.Create(Self);
  FHotList.Parent := FTabHot;
  FHotList.Align := alClient;
  FHotList.ViewStyle := vsReport;
  FHotList.ReadOnly := True;
  FHotList.RowSelect := True;
  FHotList.OnDblClick := DoHotDblClick;
  Col := FHotList.Columns.Add; Col.Caption := 'Unit';        Col.Width := 300;
  Col := FHotList.Columns.Add; Col.Caption := 'Group';       Col.Width := 80;
  Col := FHotList.Columns.Add; Col.Caption := 'Connections'; Col.Width := 110;
  Col := FHotList.Columns.Add; Col.Caption := 'Units in cycle'; Col.Width := 110;

  // ---- Help tab (static explanation of the other tabs) ----
  FHelpMemo := TMemo.Create(Self);
  FHelpMemo.Parent := FTabHelp;
  FHelpMemo.Align := alClient;
  FHelpMemo.ReadOnly := True;
  FHelpMemo.ScrollBars := ssVertical;
  FHelpMemo.WordWrap := True;
  FHelpMemo.BorderStyle := bsNone;
  FHelpMemo.Lines.Text :=
    'Circular unit references' + sLineBreak +
    'A cycle is a set of units that (directly or indirectly) use each ' +
    'other, so no clean compile/link order exists. This tool finds every ' +
    'such cycle in the project''s uses graph and shows several views of it.' + sLineBreak +
    sLineBreak +
    'Cycles' + sLineBreak +
    'Master list of every uses edge that lies inside a cycle group. ' +
    'Selecting a row shows, below it, the shortest concrete cycle running ' +
    'through that edge - each hop reads "UnitA uses UnitB (section, line)", ' +
    'and the last hop closes back to the first unit, so the ring is ' +
    'readable top to bottom.' + sLineBreak +
    sLineBreak +
    'Groups' + sLineBreak +
    'One row per strongly-connected group (a maximal set of mutually ' +
    'reachable units). Columns: Units in group; Shortest cycle (number of ' +
    'hops in the smallest ring); Max span (length of the longest simple ' +
    'cycle in the group); Edges (uses entries inside the group). ' +
    'Double-click switches to the Cycles tab and selects that group''s ' +
    'first edge.' + sLineBreak +
    sLineBreak +
    'Edge levers' + sLineBreak +
    'Every unique "UnitA uses UnitB" dependency inside a cycle, ranked by ' +
    '"Units freed" = how many units would no longer be part of ANY cycle ' +
    'if you removed that single uses entry. The top rows are therefore the ' +
    'highest-impact edges to break first. One removed edge can dissolve ' +
    'many cycles at once, so this counts freed UNITS, not cycles.' + sLineBreak +
    sLineBreak +
    'All cycles' + sLineBreak +
    'A plain-text, complete enumeration of every simple cycle ' +
    '(UnitA -> UnitB -> ... -> UnitA). If the total is too large to list ' +
    'in full it is truncated with a note.' + sLineBreak +
    sLineBreak +
    'Hotspots' + sLineBreak +
    'Units ranked by how many cycle edges touch them (Connections). The ' +
    'units at the top are involved in the most cycles and are usually the ' +
    'best targets for restructuring.' + sLineBreak +
    sLineBreak +
    'Colours & navigation' + sLineBreak +
    'In the edge and path lists, interface-section uses render red and ' +
    'implementation-section uses render orange (implementation cycles are ' +
    'the easier ones to break). Double-click a row, or use "Go to", to ' +
    'jump straight to that uses entry in the editor.';

  // ---- buttons ----
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
end;

procedure TCircularRefsDialog.FillHotspots;
var
  H: TUnitHotspot;
  Item: TListItem;
begin
  FHotList.Items.BeginUpdate;
  try
    FHotList.Items.Clear;
    for H in FHotspots do
    begin
      Item := FHotList.Items.Add;
      Item.Caption := H.UnitName;
      Item.SubItems.Add('Group ' + IntToStr(H.Group));
      Item.SubItems.Add(IntToStr(H.EdgeCount));
      Item.SubItems.Add(IntToStr(H.GroupSize));
      Item.Data := Pointer(NativeInt(FHotList.Items.Count - 1));
    end;
  finally
    FHotList.Items.EndUpdate;
  end;
end;

procedure TCircularRefsDialog.FillGroups;
var
  G: TCycleGroupInfo;
  Item: TListItem;
begin
  FGroupList.Items.BeginUpdate;
  try
    FGroupList.Items.Clear;
    for G in FGroups do
    begin
      Item := FGroupList.Items.Add;
      Item.Caption := 'Group ' + IntToStr(G.Group);
      Item.SubItems.Add(IntToStr(G.UnitCount));
      Item.SubItems.Add(IntToStr(G.ShortestCycle) + ' units');
      Item.SubItems.Add('up to ' + IntToStr(G.UnitCount));
      Item.SubItems.Add(IntToStr(G.EdgeCount));
      Item.Data := Pointer(NativeInt(G.Group));
    end;
  finally
    FGroupList.Items.EndUpdate;
  end;
end;

procedure TCircularRefsDialog.FillLevers;
var
  L: TEdgeLever;
  Item: TListItem;
  I: Integer;
begin
  FLeverList.Items.BeginUpdate;
  try
    FLeverList.Items.Clear;
    for I := 0 to High(FLevers) do
    begin
      L := FLevers[I];
      Item := FLeverList.Items.Add;
      Item.Caption := IntToStr(L.UnitsFreed);
      Item.SubItems.Add(L.FromUnit);
      Item.SubItems.Add(L.ToUnit);
      Item.SubItems.Add(L.Section);
      Item.SubItems.Add('Group ' + IntToStr(L.Group));
      Item.Data := Pointer(NativeInt(I));
    end;
  finally
    FLeverList.Items.EndUpdate;
  end;
end;

procedure TCircularRefsDialog.FillAllCycles;
const
  MaxCycles = 20000;
var
  Cycles: TArray<TCyclePath>;
  Truncated: Boolean;
  SB: TStringBuilder;
  CP: TCyclePath;
  I: Integer;
begin
  Cycles := FResult.EnumerateCycles(MaxCycles, Truncated);
  SB := TStringBuilder.Create;
  try
    if Truncated then
      SB.AppendLine(Format('Showing the first %d simple cycles (more exist - ' +
        'the total is too large to enumerate fully).', [Length(Cycles)]))
    else
      SB.AppendLine(Format('%d simple cycle(s), complete list:', [Length(Cycles)]));
    SB.AppendLine;
    for CP in Cycles do
    begin
      for I := 0 to High(CP.Units) do
      begin
        SB.Append(CP.Units[I]);
        SB.Append(' -> ');
      end;
      SB.Append(CP.Units[0]);   // close the ring
      SB.Append('   (').Append(Length(CP.Units)).Append(' units)');
      SB.AppendLine;
    end;
    FAllMemo.Text := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TCircularRefsDialog.UnitColumn(const AFile: string; ALine1: Integer;
  const AUnit: string): Integer;
var
  Content, Line: string;
  Lines: TArray<string>;
  P: Integer;
begin
  Result := 0;
  if (AUnit = '') or (ALine1 <= 0) then Exit;
  if (Editor = nil) or not Editor.ReadEditorContent(AFile, Content) then
  begin
    if not TFile.Exists(AFile) then Exit;
    try Content := TFile.ReadAllText(AFile); except Exit; end;
  end;
  Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  if (ALine1 - 1) > High(Lines) then Exit;
  Line := Lines[ALine1 - 1];
  // Whole-token match (case-insensitive): a unit name in a uses clause is
  // bounded by non-identifier characters, so "uDMDruck" never matches
  // inside "uDMDruckExtra". Dotted names (Vcl.Controls) are passed whole.
  var U := UpperCase(Line);
  var Needle := UpperCase(AUnit);
  P := Pos(Needle, U);
  while P > 0 do
  begin
    var OkBefore := (P = 1) or
      not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_', '.']);
    var AfterIdx := P + Length(Needle);
    var OkAfter := (AfterIdx > Length(U)) or
      not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_', '.']);
    if OkBefore and OkAfter then
    begin
      Result := P - 1;
      Exit;
    end;
    P := PosEx(Needle, U, P + 1);
  end;
end;

procedure TCircularRefsDialog.DoLeverDblClick(Sender: TObject);
var
  Idx, Col: Integer;
begin
  if FLeverList.Selected = nil then Exit;
  Idx := NativeInt(FLeverList.Selected.Data);
  if (Idx < 0) or (Idx >= Length(FLevers)) then Exit;
  Col := UnitColumn(FLevers[Idx].FromFile, FLevers[Idx].Line, FLevers[Idx].ToUnit);
  Editor.GotoLocation(FLevers[Idx].FromFile, FLevers[Idx].Line - 1, Col,
    Length(FLevers[Idx].ToUnit));
end;

procedure TCircularRefsDialog.DoGroupDblClick(Sender: TObject);
var
  Grp, I: Integer;
begin
  if FGroupList.Selected = nil then Exit;
  Grp := NativeInt(FGroupList.Selected.Data);
  // Find the first edge of this group, select it in the Cycles tab and
  // switch there so the shortest-cycle path is shown.
  for I := 0 to High(FResult.Edges) do
    if FResult.Edges[I].Group = Grp then
    begin
      FPages.ActivePage := FTabCycles;
      FEdgeList.ItemIndex := I;
      if FEdgeList.Items[I] <> nil then
        FEdgeList.Items[I].MakeVisible(False);
      FEdgeList.SetFocus;
      Break;
    end;
end;

procedure TCircularRefsDialog.DoEdgeColumnClick(Sender: TObject;
  Column: TListColumn);
var
  Col: Integer;
  Asc: Boolean;
  A: TArray<TCycleEdge>;
begin
  Col := Column.Index;
  if FEdgeSortCol = Col then
    FEdgeSortAsc := not FEdgeSortAsc
  else
  begin
    FEdgeSortCol := Col;
    FEdgeSortAsc := True;
  end;
  Asc := FEdgeSortAsc;
  // FEdgeList is virtual: sort the backing array itself. All consumers
  // (DoEdgeData/Select/Draw/DblClick, group jump) index FResult.Edges,
  // so they stay consistent after the sort.
  A := FResult.Edges;
  TArray.Sort<TCycleEdge>(A, TComparer<TCycleEdge>.Construct(
    function(const L, R: TCycleEdge): Integer
    begin
      case Col of
        0: Result := L.Group - R.Group;
        1: Result := CompareText(L.FromUnit, R.FromUnit);
        2: Result := Ord(L.InInterface) - Ord(R.InInterface);
        3: Result := CompareText(L.ToUnit, R.ToUnit);
        4: Result := L.Line - R.Line;
      else Result := L.GroupSize - R.GroupSize;
      end;
      if Result = 0 then
        Result := CompareText(L.FromUnit, R.FromUnit);
      if not Asc then
        Result := -Result;
    end));
  FEdgeList.ClearSelection;   // old selected index = a different edge now
  FEdgeList.Invalidate;
  SetListViewSortArrow(FEdgeList, Col, Asc);
end;

procedure TCircularRefsDialog.DoEdgeData(Sender: TObject; Item: TListItem);
var
  E: TCycleEdge;
begin
  if (Item.Index < 0) or (Item.Index >= Length(FResult.Edges)) then Exit;
  E := FResult.Edges[Item.Index];
  Item.Caption := 'Group ' + IntToStr(E.Group);
  Item.SubItems.Add(E.FromUnit);
  if E.InInterface then Item.SubItems.Add('interface')
  else Item.SubItems.Add('implementation');
  Item.SubItems.Add(E.ToUnit);
  Item.SubItems.Add(IntToStr(E.Line));
  Item.SubItems.Add(IntToStr(E.GroupSize));
end;

procedure TCircularRefsDialog.DoEdgeSelect(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  Hop: TCycleHop;
  LI: TListItem;
  Step: Integer;
begin
  if not Selected then Exit;
  if (Item.Index < 0) or (Item.Index >= Length(FResult.Edges)) then Exit;

  FCurPath := FResult.ShortestCyclePath(FResult.Edges[Item.Index]);
  FPathLbl.Caption := Format(
    'Cycle through %s (%d hops back to the start):',
    [FResult.Edges[Item.Index].FromUnit, Length(FCurPath)]);

  FPathList.Items.BeginUpdate;
  try
    FPathList.Items.Clear;
    Step := 1;
    for Hop in FCurPath do
    begin
      LI := FPathList.Items.Add;
      LI.Caption := IntToStr(Step);
      LI.SubItems.Add(Hop.FromUnit);
      if Hop.InInterface then LI.SubItems.Add('interface')
      else LI.SubItems.Add('implementation');
      LI.SubItems.Add(Hop.ToUnit);
      LI.SubItems.Add(IntToStr(Hop.Line));
      Inc(Step);
    end;
  finally
    FPathList.Items.EndUpdate;
  end;
end;

procedure TCircularRefsDialog.DoEdgeDraw(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  if (Item.Index >= 0) and (Item.Index < Length(FResult.Edges)) then
    if FResult.Edges[Item.Index].InInterface then
      Sender.Canvas.Font.Color := clRed
    else
      Sender.Canvas.Font.Color := clOrange;
end;

procedure TCircularRefsDialog.DoPathDraw(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  if (Item.Index >= 0) and (Item.Index < Length(FCurPath)) then
    if FCurPath[Item.Index].InInterface then
      Sender.Canvas.Font.Color := clRed
    else
      Sender.Canvas.Font.Color := clOrange;
end;

procedure TCircularRefsDialog.DoEdgeDblClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FEdgeList.Selected = nil then Exit;
  Idx := FEdgeList.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FResult.Edges)) then Exit;
  var Col := UnitColumn(FResult.Edges[Idx].FromFile, FResult.Edges[Idx].Line,
    FResult.Edges[Idx].ToUnit);
  Editor.GotoLocation(FResult.Edges[Idx].FromFile, FResult.Edges[Idx].Line - 1, Col,
    Length(FResult.Edges[Idx].ToUnit));
end;

procedure TCircularRefsDialog.DoPathDblClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FPathList.Selected = nil then Exit;
  Idx := FPathList.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FCurPath)) then Exit;
  if FCurPath[Idx].Line <= 0 then Exit;
  var Col := UnitColumn(FCurPath[Idx].FromFile, FCurPath[Idx].Line,
    FCurPath[Idx].ToUnit);
  Editor.GotoLocation(FCurPath[Idx].FromFile, FCurPath[Idx].Line - 1, Col,
    Length(FCurPath[Idx].ToUnit));
end;

procedure TCircularRefsDialog.DoHotDblClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FHotList.Selected = nil then Exit;
  Idx := NativeInt(FHotList.Selected.Data);
  if (Idx < 0) or (Idx >= Length(FHotspots)) then Exit;
  // Jump to the top of the unit; the cycle edges are what matter, but
  // this at least opens the file.
  Editor.GotoLocation(FHotspots[Idx].FileName, 0, 0, 0);
end;

procedure TCircularRefsDialog.DoGotoClick(Sender: TObject);
begin
  if FPages.ActivePage = FTabHot then
    DoHotDblClick(nil)
  else if FPages.ActivePage = FTabLevers then
    DoLeverDblClick(nil)
  else if FPages.ActivePage = FTabGroups then
    DoGroupDblClick(nil)
  else if (FPathList.Selected <> nil) then
    DoPathDblClick(nil)
  else
    DoEdgeDblClick(nil);
end;

procedure TCircularRefsDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TCircularRefsDialog.DoFormClose(Sender: TObject; var Action: TCloseAction);
begin
  FResult.Free;   // dialog owns the analysis result
  Action := caFree;
end;

procedure CheckCircularReferences;
var
  Files: TArray<string>;
  Res: TUsesCycleResult;
  Dlg: TCircularRefsDialog;
  ProgressForm: TForm;
  ProgressLbl: TLabel;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowMessage('No project loaded / no source files found.');
    Exit;
  end;

  ProgressForm := TForm.CreateNew(nil);
  try
    ProgressForm.Caption := 'Circular unit references';
    ProgressForm.BorderStyle := bsToolWindow;
    ProgressForm.FormStyle := fsStayOnTop;
    ProgressForm.Position := poScreenCenter;
    ProgressForm.ClientWidth := 420;
    ProgressForm.ClientHeight := 48;
    ProgressLbl := TLabel.Create(ProgressForm);
    ProgressLbl.Parent := ProgressForm;
    ProgressLbl.Align := alClient;
    ProgressLbl.Alignment := taCenter;
    ProgressLbl.Layout := tlCenter;
    ProgressLbl.Caption := 'Scanning...';
    PrepareDialog(ProgressForm, nil);
    ProgressForm.Show;
    Screen.Cursor := crHourGlass;
    try
      Res := TUsesGraphAnalyzer.Analyze(Files,
        procedure(ACurrent, ATotal: Integer; AFile: string)
        begin
          if (ACurrent mod 10 = 0) or (ACurrent = ATotal) then
          begin
            ProgressLbl.Caption := Format('Scanning %d / %d  -  %s',
              [ACurrent, ATotal, ExtractFileName(AFile)]);
            Application.ProcessMessages;
          end;
        end);
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    ProgressForm.Free;
  end;

  if Length(Res.Edges) = 0 then
  begin
    Res.Free;
    ShowMessage('No circular unit references found.');
    Exit;
  end;
  Dlg := TCircularRefsDialog.CreateDialog(Application.MainForm, Res);
  Dlg.Show;
end;

end.
