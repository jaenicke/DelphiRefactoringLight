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
// "Path analysis": the point-to-point query. Pick two units and get the
//   shortest dependency chain between them, over the FULL project graph -
//   not only over edges that already lie in a cycle. That distinction is
//   the whole point: when the compiler says
//     UnitA.pas(20): F2047 Circular unit reference to 'UnitB'
//   the edge UnitA -> UnitB is the one being ADDED, so it does not exist
//   in the analysed sources and neither unit need be in any cycle yet.
//   What the user needs is the path that already runs from UnitB back to
//   UnitA - enter B and A, tick "interface uses only" (that is the
//   relation F2047 checks) and the intermediate stations are listed.
//   Entering the same unit twice asks for the shortest cycle through it.
//
// "All cycles": a plain-text enumeration of every simple cycle
//   (UnitA -> UnitB -> ... -> UnitA). The count limit is editable (0 =
//   no limit, guarded by a time budget), the list can be filtered to the
//   cycles through ONE unit, and the complete result can be exported to
//   a file - a big legacy project easily has more cycles than any list
//   can usefully show.
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

uses
  Vcl.Forms, Expert.UsesGraph;

procedure CheckCircularReferences;

/// <summary>Builds the viewer for an already computed analysis; the
///  dialog then OWNS AResult and frees it when it closes.
///  APrefillTarget seeds the "to unit" field of the path query.
///  Separated from CheckCircularReferences so the layout can be
///  rendered headless in a test - the tabs are built in code, and a
///  screenshot is the only way to see what the user sees.</summary>
function CreateCircularRefsDialog(AResult: TUsesCycleResult;
  const APrefillTarget: string = ''): TForm;

implementation

uses
  System.SysUtils, System.Classes, System.UITypes, System.IOUtils, System.StrUtils,
  System.Types, System.Math, System.DateUtils,
  System.Generics.Defaults, System.Generics.Collections,
  Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics,
  Vcl.Dialogs, Vcl.ExtCtrls,
  Expert.EditorHelperIntf, Expert.DialogHelper,
  Expert.IdeThemes, Expert.ListViewSort;

type
  TCircularRefsDialog = class(TForm)
  private
    FResult: TUsesCycleResult;
    FHotspots: TArray<TUnitHotspot>;
    FGroups: TArray<TCycleGroupInfo>;
    FLevers: TArray<TEdgeLever>;
    FCurPath: TArray<TCycleHop>;
    FQueryPath: TArray<TCycleHop>;
    FAllCycles: TArray<TCyclePath>;
    FAllTruncated: Boolean;
    FEdgeSortCol: Integer;    // -1 = unsorted (FEdgeList is virtual)
    FEdgeSortAsc: Boolean;

    FPages: TPageControl;
    FTabCycles: TTabSheet;
    FTabGroups: TTabSheet;
    FTabLevers: TTabSheet;
    FTabAll: TTabSheet;
    FTabPath: TTabSheet;
    FTabHot: TTabSheet;
    FTabHelp: TTabSheet;
    FHelpMemo: TMemo;       // static help text
    FEdgeList: TListView;   // owner-data master (can be large)
    FPathList: TListView;   // filled per selection (small)
    FPathLbl: TLabel;
    FGroupList: TListView;  // filled once
    FLeverList: TListView;  // filled once
    FAllMemo: TMemo;        // text dump, refilled on demand
    FAllFilter: TComboBox;  // '' = whole project, else one unit
    FAllLimit: TEdit;       // 0 = no count limit
    FAllStatus: TLabel;
    FBtnAllRefresh: TButton;
    FBtnAllExport: TButton;
    FCbFrom: TComboBox;
    FCbTo: TComboBox;
    FChkIntfOnly: TCheckBox;
    FQueryList: TListView;
    FQueryLbl: TLabel;
    FHotList: TListView;    // filled once
    FLblSummary: TLabel;
    FBtnClose: TButton;
    FBtnGoto: TButton;

    procedure BuildLayout;
    procedure FillHotspots;
    procedure FillGroups;
    procedure FillLevers;
    procedure FillAllCycles;
    procedure BuildPathTab;
    procedure BuildAllCyclesTab;
    procedure ShowQueryPath(const AHops: TArray<TCycleHop>; const ACaption: string);
    procedure DoFindPathClick(Sender: TObject);
    procedure DoSwapClick(Sender: TObject);
    procedure DoCyclesThroughClick(Sender: TObject);
    procedure DoQueryDraw(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoQueryDblClick(Sender: TObject);
    procedure DoAllRefreshClick(Sender: TObject);
    procedure DoAllExportClick(Sender: TObject);
    function CycleReport(AForFile: Boolean): string;
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
    /// <summary>Opens the dialog on the Path analysis tab, pre-filled -
    ///  the entry point for "the compiler just told me A -> B is
    ///  circular".</summary>
    procedure StartWithPathQuery(const AFrom, ATo: string);
  end;

const
  clOrange = TColor($00008CFF);
  /// <summary>Default cap for the cycle enumeration. Editable in the
  ///  dialog (0 = no limit) - a big legacy project can have far more
  ///  than this, which is exactly what a tester reported.</summary>
  DefaultCycleLimit = 20000;
  /// <summary>Wall-clock guard for an uncapped enumeration. The number
  ///  of simple cycles is exponential in the worst case, so "no limit"
  ///  needs a way back.</summary>
  CycleBudgetMs = 20000;

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
  if Length(FResult.Edges) = 0 then
    FLblSummary.Caption :=
      'No circular unit references in this project.  The Path analysis tab ' +
      'still works: it answers "which chain leads from unit X to unit Y?" ' +
      'over the whole uses graph, cycle or not.'
  else
    FLblSummary.Caption := Format(
      '%d cycle group(s), %d edge(s), %d via interface uses.  ' +
      'Cycles tab: select an edge to see how it closes. Path analysis tab: the chain between two units. ' +
      'Hotspots tab: the most entangled units.',
      [Groups, Length(FResult.Edges), IntfEdges]);
  if Length(FResult.Edges) = 0 then
    FPages.ActivePage := FTabPath;

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

  FTabPath := TTabSheet.Create(FPages);
  FTabPath.PageControl := FPages;
  FTabPath.Caption := 'Path analysis';

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

  // ---- All cycles tab ----
  BuildAllCyclesTab;

  // ---- Path analysis tab ----
  BuildPathTab;

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
    'A plain-text enumeration of every simple cycle ' +
    '(UnitA -> UnitB -> ... -> UnitA). "Max" caps the number of cycles; ' +
    'set it to 0 for no limit (a time budget then stops a runaway ' +
    'enumeration - the number of simple cycles grows exponentially). ' +
    '"Only cycles containing" restricts the enumeration to the cycles ' +
    'that pass through one unit, which is usually the question you ' +
    'actually have. "Export to file" writes the complete result, ' +
    'however long it is.' + sLineBreak +
    sLineBreak +
    'Path analysis' + sLineBreak +
    'The point-to-point query, and the answer to a compiler error like ' +
    '"UnitA.pas(20): F2047 Circular unit reference to UnitB". ' +
    'That edge UnitA -> UnitB is the one you are trying to ADD, so it is ' +
    'not in the sources and the other tabs cannot show it. Enter UnitB ' +
    'as "from" and UnitA as "to", tick "interface uses only" (that is ' +
    'the relation the compiler rejects) and you get the shortest chain ' +
    'UnitB -> ... -> UnitA that already exists - every intermediate ' +
    'station listed, double-click to jump to the uses entry. Entering ' +
    'the same unit twice gives the shortest cycle through it.' + sLineBreak +
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

procedure TCircularRefsDialog.BuildAllCyclesTab;
var
  Bar: TPanel;
  L: TLabel;
begin
  Bar := TPanel.Create(Self);
  Bar.Parent := FTabAll;
  Bar.Align := alTop;
  Bar.Height := 62;
  Bar.BevelOuter := bvNone;

  L := TLabel.Create(Self);
  L.Parent := Bar;
  L.SetBounds(6, 10, 130, 16);
  L.Caption := 'Only cycles containing:';

  FAllFilter := TComboBox.Create(Self);
  FAllFilter.Parent := Bar;
  FAllFilter.SetBounds(142, 6, 240, 24);
  FAllFilter.Style := csDropDown;      // typing is faster than scrolling
  FAllFilter.AutoComplete := True;
  FAllFilter.Sorted := False;
  FAllFilter.Items.Add('');            // '' = the whole project
  for var UName in FResult.UnitNames do
    FAllFilter.Items.Add(UName);
  FAllFilter.ItemIndex := 0;

  L := TLabel.Create(Self);
  L.Parent := Bar;
  L.SetBounds(398, 10, 130, 16);
  L.Caption := 'Max (0 = no limit):';

  FAllLimit := TEdit.Create(Self);
  FAllLimit.Parent := Bar;
  FAllLimit.SetBounds(514, 6, 70, 24);
  FAllLimit.Text := IntToStr(DefaultCycleLimit);

  FBtnAllRefresh := TButton.Create(Self);
  FBtnAllRefresh.Parent := Bar;
  FBtnAllRefresh.SetBounds(596, 5, 90, 26);
  FBtnAllRefresh.Caption := 'Re&fresh';
  FBtnAllRefresh.OnClick := DoAllRefreshClick;

  FBtnAllExport := TButton.Create(Self);
  FBtnAllExport.Parent := Bar;
  FBtnAllExport.SetBounds(694, 5, 130, 26);
  FBtnAllExport.Caption := '&Export to file...';
  FBtnAllExport.OnClick := DoAllExportClick;

  FAllStatus := TLabel.Create(Self);
  FAllStatus.Parent := Bar;
  FAllStatus.SetBounds(6, 38, 800, 16);

  FAllMemo := TMemo.Create(Self);
  FAllMemo.Parent := FTabAll;
  FAllMemo.Align := alClient;
  FAllMemo.ReadOnly := True;
  FAllMemo.ScrollBars := ssBoth;
  FAllMemo.WordWrap := False;
  FAllMemo.Font.Name := 'Consolas';
  FAllMemo.Font.Size := 9;
end;

procedure TCircularRefsDialog.BuildPathTab;
var
  Bar: TPanel;
  L: TLabel;
  Info: TLabel;
  Btn: TButton;
  Col: TListColumn;
  Names: TArray<string>;
begin
  Info := TLabel.Create(Self);
  Info.Parent := FTabPath;
  Info.Align := alTop;
  Info.AlignWithMargins := True;
  Info.Margins.SetBounds(6, 6, 6, 2);
  // VCL pitfall: AutoSize + WordWrap + alTop wraps against the DEFAULT
  // width and produces an invisible giant. Fixed height, no AutoSize.
  Info.AutoSize := False;
  Info.WordWrap := True;
  Info.Height := 48;
  Info.Caption :=
    'The shortest dependency chain between two units, over the whole uses ' +
    'graph - the edge does not have to be part of a cycle. After ' +
    '"F2047 Circular unit reference to UnitB" in UnitA: enter UnitB -> ' +
    'UnitA with "interface uses only" ticked to see the chain that ' +
    'already leads back. Same unit twice = the shortest cycle through it.';

  Bar := TPanel.Create(Self);
  Bar.Parent := FTabPath;
  Bar.Align := alTop;
  Bar.Height := 104;
  Bar.BevelOuter := bvNone;

  Names := FResult.UnitNames;

  L := TLabel.Create(Self);
  L.Parent := Bar;
  L.SetBounds(6, 11, 70, 16);
  L.Caption := 'From unit:';

  FCbFrom := TComboBox.Create(Self);
  FCbFrom.Parent := Bar;
  FCbFrom.SetBounds(80, 7, 240, 24);
  FCbFrom.Style := csDropDown;
  FCbFrom.AutoComplete := True;
  for var UName in Names do
    FCbFrom.Items.Add(UName);

  L := TLabel.Create(Self);
  L.Parent := Bar;
  L.SetBounds(334, 11, 60, 16);
  L.Caption := 'to unit:';

  FCbTo := TComboBox.Create(Self);
  FCbTo.Parent := Bar;
  FCbTo.SetBounds(392, 7, 240, 24);
  FCbTo.Style := csDropDown;
  FCbTo.AutoComplete := True;
  for var UName in Names do
    FCbTo.Items.Add(UName);

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.SetBounds(642, 6, 90, 26);
  Btn.Caption := '&Swap';
  Btn.OnClick := DoSwapClick;

  FChkIntfOnly := TCheckBox.Create(Self);
  FChkIntfOnly.Parent := Bar;
  FChkIntfOnly.SetBounds(80, 40, 420, 20);
  FChkIntfOnly.Caption := 'interface uses only (the relation F2047 rejects)';
  FChkIntfOnly.Checked := True;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.SetBounds(80, 66, 110, 26);
  Btn.Caption := 'Find &path';
  Btn.Default := True;
  Btn.OnClick := DoFindPathClick;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.SetBounds(200, 66, 230, 26);
  Btn.Caption := 'All &cycles through the first unit';
  Btn.OnClick := DoCyclesThroughClick;

  FQueryLbl := TLabel.Create(Self);
  FQueryLbl.Parent := FTabPath;
  // alTop stacks by the CURRENT Top, and a freshly created control sits
  // at 0 - which put the status line ABOVE the input row. Push it down
  // so it lands where it belongs: directly over the result list.
  FQueryLbl.Top := Bar.Top + Bar.Height + 1;
  FQueryLbl.Align := alTop;
  FQueryLbl.AlignWithMargins := True;
  FQueryLbl.Margins.SetBounds(6, 4, 6, 4);
  FQueryLbl.Caption := 'No query run yet.';

  FQueryList := TListView.Create(Self);
  FQueryList.Parent := FTabPath;
  FQueryList.Align := alClient;
  FQueryList.ViewStyle := vsReport;
  FQueryList.ReadOnly := True;
  FQueryList.RowSelect := True;
  FQueryList.OnCustomDrawItem := DoQueryDraw;
  FQueryList.OnDblClick := DoQueryDblClick;
  Col := FQueryList.Columns.Add; Col.Caption := 'Step';    Col.Width := 50;
  Col := FQueryList.Columns.Add; Col.Caption := 'Unit';    Col.Width := 250;
  Col := FQueryList.Columns.Add; Col.Caption := 'Section'; Col.Width := 110;
  Col := FQueryList.Columns.Add; Col.Caption := 'uses';    Col.Width := 250;
  Col := FQueryList.Columns.Add; Col.Caption := 'Line';    Col.Width := 60;
end;

procedure TCircularRefsDialog.ShowQueryPath(const AHops: TArray<TCycleHop>;
  const ACaption: string);
var
  Hop: TCycleHop;
  LI: TListItem;
  Step: Integer;
begin
  FQueryPath := AHops;
  FQueryLbl.Caption := ACaption;
  FQueryList.Items.BeginUpdate;
  try
    FQueryList.Items.Clear;
    Step := 1;
    for Hop in FQueryPath do
    begin
      LI := FQueryList.Items.Add;
      LI.Caption := IntToStr(Step);
      LI.SubItems.Add(Hop.FromUnit);
      if Hop.InInterface then LI.SubItems.Add('interface')
      else LI.SubItems.Add('implementation');
      LI.SubItems.Add(Hop.ToUnit);
      LI.SubItems.Add(IntToStr(Hop.Line));
      Inc(Step);
    end;
  finally
    FQueryList.Items.EndUpdate;
  end;
end;

procedure TCircularRefsDialog.DoFindPathClick(Sender: TObject);
var
  FromU, ToU: string;
  Hops: TArray<TCycleHop>;
  Sect: string;
begin
  FromU := Trim(FCbFrom.Text);
  ToU := Trim(FCbTo.Text);
  if (FromU = '') or (ToU = '') then
  begin
    ShowThemedMessage('Please name both units.');
    Exit;
  end;
  // Naming a unit the analysis never saw is the most likely mistake
  // (a library unit, a typo) - say which one, not just "no path".
  if not FResult.KnowsUnit(FromU) then
  begin
    ShowThemedMessage(Format('%s is not one of the analysed project units.', [FromU]));
    Exit;
  end;
  if not FResult.KnowsUnit(ToU) then
  begin
    ShowThemedMessage(Format('%s is not one of the analysed project units.', [ToU]));
    Exit;
  end;

  if FChkIntfOnly.Checked then Sect := 'interface uses only'
  else Sect := 'interface + implementation uses';
  Hops := FResult.FindPath(FromU, ToU, FChkIntfOnly.Checked);
  if Length(Hops) = 0 then
  begin
    ShowQueryPath(nil, Format(
      'No path %s -> %s (%s). ' +
      'With "interface uses only" ticked this means the compiler would ' +
      'NOT reject %s in the interface uses of %s.',
      [FromU, ToU, Sect, FromU, ToU]));
    Exit;
  end;
  if SameText(FromU, ToU) then
    ShowQueryPath(Hops, Format('Shortest cycle through %s: %d hop(s), %s.',
      [FromU, Length(Hops), Sect]))
  else
    ShowQueryPath(Hops, Format(
      'Shortest path %s -> %s: %d hop(s), %s. ' +
      'Adding %s to the %s uses of %s would close this ring.',
      [FromU, ToU, Length(Hops), Sect, FromU,
       System.StrUtils.IfThen(FChkIntfOnly.Checked, 'interface', ''), ToU]));
end;

procedure TCircularRefsDialog.DoSwapClick(Sender: TObject);
var
  S: string;
begin
  S := FCbFrom.Text;
  FCbFrom.Text := FCbTo.Text;
  FCbTo.Text := S;
end;

procedure TCircularRefsDialog.DoCyclesThroughClick(Sender: TObject);
var
  U: string;
begin
  U := Trim(FCbFrom.Text);
  if U = '' then
  begin
    ShowThemedMessage('Please name a unit in the "From unit" field.');
    Exit;
  end;
  if not FResult.KnowsUnit(U) then
  begin
    ShowThemedMessage(Format('%s is not one of the analysed project units.', [U]));
    Exit;
  end;
  FAllFilter.Text := U;
  FPages.ActivePage := FTabAll;
  FillAllCycles;
end;

procedure TCircularRefsDialog.DoQueryDraw(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  DefaultDraw := True;
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
  if (Item.Index >= 0) and (Item.Index < Length(FQueryPath)) then
    if FQueryPath[Item.Index].InInterface then
      Sender.Canvas.Font.Color := clRed
    else
      Sender.Canvas.Font.Color := clOrange;
end;

procedure TCircularRefsDialog.DoQueryDblClick(Sender: TObject);
var
  Idx: Integer;
begin
  if FQueryList.Selected = nil then Exit;
  Idx := FQueryList.Selected.Index;
  if (Idx < 0) or (Idx >= Length(FQueryPath)) then Exit;
  if FQueryPath[Idx].Line <= 0 then Exit;
  var Col := UnitColumn(FQueryPath[Idx].FromFile, FQueryPath[Idx].Line,
    FQueryPath[Idx].ToUnit);
  Editor.GotoLocation(FQueryPath[Idx].FromFile, FQueryPath[Idx].Line - 1, Col,
    Length(FQueryPath[Idx].ToUnit));
end;

procedure TCircularRefsDialog.StartWithPathQuery(const AFrom, ATo: string);
begin
  FCbFrom.Text := AFrom;
  FCbTo.Text := ATo;
  // Only jump to the tab (and answer) when both ends are actually known.
  // Called with just the caret's unit as the target, this merely
  // pre-fills the field the user would otherwise type first.
  if FResult.KnowsUnit(AFrom) and FResult.KnowsUnit(ATo) then
  begin
    FPages.ActivePage := FTabPath;
    DoFindPathClick(nil);
  end;
end;

procedure TCircularRefsDialog.DoAllRefreshClick(Sender: TObject);
begin
  FillAllCycles;
end;

function TCircularRefsDialog.CycleReport(AForFile: Boolean): string;
// The memo and the export share one renderer; only the amount differs.
// A memo with a million lines is unusable (and slow to fill), so the
// display is cut at MaxMemoCycles while the file always gets everything
// that was enumerated - that is what "export the full result" means.
const
  MaxMemoCycles = 20000;
var
  SB: TStringBuilder;
  CP: TCyclePath;
  I, Shown, Limit: Integer;
begin
  SB := TStringBuilder.Create;
  try
    if AForFile then Limit := Length(FAllCycles)
    else Limit := Min(MaxMemoCycles, Length(FAllCycles));

    if Trim(FAllFilter.Text) <> '' then
      SB.AppendLine(Format('Cycles through %s: %d', [Trim(FAllFilter.Text),
        Length(FAllCycles)]))
    else
      SB.AppendLine(Format('Simple cycles in the project: %d', [Length(FAllCycles)]));
    if FAllTruncated then
      SB.AppendLine('INCOMPLETE - the enumeration hit the limit; more cycles exist.');
    if Limit < Length(FAllCycles) then
      SB.AppendLine(Format('Listing the first %d here - use "Export to file" ' +
        'for the complete list.', [Limit]));
    SB.AppendLine;

    Shown := 0;
    for CP in FAllCycles do
    begin
      if Shown >= Limit then Break;
      for I := 0 to High(CP.Units) do
      begin
        SB.Append(CP.Units[I]);
        SB.Append(' -> ');
      end;
      SB.Append(CP.Units[0]);   // close the ring
      SB.Append('   (').Append(Length(CP.Units)).Append(' units)');
      SB.AppendLine;
      Inc(Shown);
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure TCircularRefsDialog.FillAllCycles;
var
  Filter: string;
  Limit: Integer;
  Started: TDateTime;
begin
  Filter := Trim(FAllFilter.Text);
  if (Filter <> '') and not FResult.KnowsUnit(Filter) then
  begin
    ShowThemedMessage(Format('%s is not one of the analysed project units.', [Filter]));
    Exit;
  end;
  if not TryStrToInt(Trim(FAllLimit.Text), Limit) or (Limit < 0) then
    Limit := DefaultCycleLimit;

  Screen.Cursor := crHourGlass;
  Started := Now;
  try
    if Filter <> '' then
      FAllCycles := FResult.EnumerateCyclesThrough(Filter, Limit, FAllTruncated,
        CycleBudgetMs)
    else
      FAllCycles := FResult.EnumerateCycles(Limit, FAllTruncated, CycleBudgetMs);
  finally
    Screen.Cursor := crDefault;
  end;

  FAllMemo.Text := CycleReport(False);
  FAllStatus.Caption := Format('%d cycle(s)%s, %.1f s.  %s',
    [Length(FAllCycles),
     System.StrUtils.IfThen(FAllTruncated, ' (limit reached - more exist)', ''),
     MilliSecondsBetween(Now, Started) / 1000,
     System.StrUtils.IfThen(FAllTruncated,
       'Raise "Max" (0 = no limit) or filter by a unit for the complete picture.',
       'Complete.')]);
end;

procedure TCircularRefsDialog.DoAllExportClick(Sender: TObject);
var
  Dlg: TSaveDialog;
  Base: string;
begin
  if Length(FAllCycles) = 0 then
  begin
    ShowThemedMessage('Nothing to export - run the enumeration first.');
    Exit;
  end;
  Dlg := TSaveDialog.Create(nil);
  try
    Dlg.Filter := 'Text file (*.txt)|*.txt|All files (*.*)|*.*';
    Dlg.DefaultExt := 'txt';
    Base := Trim(FAllFilter.Text);
    if Base <> '' then Dlg.FileName := 'cycles-' + Base + '.txt'
    else Dlg.FileName := 'cycles.txt';
    Dlg.Options := Dlg.Options + [ofOverwritePrompt, ofPathMustExist];
    if not Dlg.Execute(Handle) then Exit;
    try
      TFile.WriteAllText(Dlg.FileName, CycleReport(True), TEncoding.UTF8);
    except
      on E: Exception do
      begin
        ShowThemedMessage('Could not write the file: ' + E.Message);
        Exit;
      end;
    end;
    FAllStatus.Caption := Format('%d cycle(s) written to %s',
      [Length(FAllCycles), Dlg.FileName]);
  finally
    Dlg.Free;
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
  // Custom-draw rows take their background from Brush.Color - set the
  // themed one, or dark-themed list views paint these rows on WHITE.
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
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
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
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
  if FPages.ActivePage = FTabPath then
    DoQueryDblClick(nil)
  else if FPages.ActivePage = FTabHot then
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

function CreateCircularRefsDialog(AResult: TUsesCycleResult;
  const APrefillTarget: string): TForm;
var
  Dlg: TCircularRefsDialog;
begin
  Dlg := TCircularRefsDialog.CreateDialog(Application.MainForm, AResult);
  Dlg.StartWithPathQuery('', APrefillTarget);
  Result := Dlg;
end;

procedure CheckCircularReferences;
var
  Files: TArray<string>;
  Res: TUsesCycleResult;
  Dlg: TForm;
  ProgressForm: TCheckProgressWindow;
begin
  Files := Editor.GetProjectSourceFiles;
  if Length(Files) = 0 then
  begin
    ShowThemedMessage('No project loaded / no source files found.');
    Exit;
  end;

  ProgressForm := CreateCheckProgress('Circular unit references', nil);
  try
    Screen.Cursor := crHourGlass;
    try
      Res := TUsesGraphAnalyzer.Analyze(Files,
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

  // NOT an early exit any more when nothing is circular: the path
  // analysis answers "which chain leads from X to Y" for a project with
  // no cycle at all - and that is precisely the situation after the
  // compiler refused the edge you wanted to add.
  // The unit in the editor is almost always the one the compiler
  // complained about, i.e. the TARGET of the path query ("...back to
  // UnitA"), so pre-fill it and leave the source to the user.
  Dlg := CreateCircularRefsDialog(Res,
    ChangeFileExt(ExtractFileName(Editor.GetActiveFileName), ''));
  Dlg.Show;
end;

end.
