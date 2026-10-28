(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.StatusWindow;

// Dockable "Refactoring Light Status" window (IDE only) - one place to
// see what the plugin is currently doing: identifier index, LSP session,
// live quick-fix checker, which diagnostics source answered, and how the
// menus got installed.
//
// Built on the OFFICIAL ToolsAPI docking API:
//   INTAServices.RegisterDockableForm(INTACustomDockableForm) at load
//   time (so the IDE can restore the window from a saved desktop) and
//   CreateDockableForm to show it. The interface hands the IDE a FRAME
//   CLASS - the IDE owns the surrounding form, we only fill the frame.
//   MANDATORY: unregister before the BPL unloads.
//
// HARD RULE (see CLAUDE.md): the refresh runs on a plain WM_TIMER tick
// and must never call GetCurrentContext (that moves the caret) - only
// cheap state getters are used.

interface

uses
  System.Types,        // TRect - the window's own geometry is part of its state
  Expert.IdeThemes,    // TStatusLevel: a row's colour is part of its data
  Expert.RestartHint;  // TDockPlace: the hand-over and this window have to
                       // mean the SAME record by where a window was docked -
                       // a second copy of it here is a copy that can drift.

/// <summary>Registers the dockable form with the IDE. Call from Register.</summary>
procedure RegisterStatusWindow;
/// <summary>Unregisters and closes it. Call before the BPL unloads.</summary>
procedure UnregisterStatusWindow;
const
  /// <summary>The identifier INTACustomDockableForm.GetIdentifier answers -
  ///  and therefore the window's SECTION NAME in the desktop file, which is
  ///  also how a dock site lists it ("TabDockClients=MessageView,
  ///  RefactoringLightStatus"). ONE place: the registration, the desktop
  ///  lookup and the .dsk all have to mean the same string.</summary>
  StatusWindowIdentifier = 'RefactoringLightStatus';
  McpToolsWindowIdentifier = 'RefactoringLightMcpTools';

type
  /// <summary>What a dockable window of ours was doing at the unload.</summary>
  TRlDockState = record
    Open: Boolean;
    Docked: Boolean;
    /// <summary>SCREEN coordinates, and filled whatever the window was doing.
    ///
    ///  THE CONVERSION IS THE POINT, and 2.8.1 solved the same problem the
    ///  other way round: BoundsRect is relative to the PARENT, and a docked
    ///  form's parent is its dock site, so Left/Top are near zero - handing
    ///  that on as a screen position is what put the window in the top left
    ///  corner after an install. 2.8.1 answered it by recording no bounds at
    ///  all for a docked window, which left the window in the MIDDLE of the
    ///  screen when the re-dock did not work (reported 2026-10-28).
    ///  So the bounds are converted through the parent instead
    ///  (ScreenBoundsOf), which makes them true in both states - and the
    ///  guard moves to where the rect is READ rather than to the shape of
    ///  the restore: a docked window whose dock sites are all gone goes back
    ///  to the place it occupied, floating, which is not a restore and is not
    ///  reported as one.</summary>
    Bounds: TRect;
    /// <summary>The dock chain it sat in - empty while it is
    ///  free-standing.</summary>
    Host: TDockPlace;
  end;

/// <summary>What the status window is doing right now - read at the unload,
///  while the form still exists, and handed to the next instance through the
///  marker.
///  MEASURED in the user's own project desktop: the docking position is NOT
///  stored with our window ([RefactoringLightStatus] holds geometry and
///  Docked=1 only) but with the dock SITE ("[DockSite5] HostDockSite=
///  DockBottomCenterPanel / TabDockClients=MessageView,
///  RefactoringLightStatus"). What that file cannot say, the LIVE form can -
///  and it can say more than one link of it, which is what ReadDockTarget
///  walks: the site, the form that owns it, the site THAT form is docked in,
///  and so on up to the IDE's main window.</summary>
function StatusDockWindowState: TRlDockState;

/// <summary>That state in one line, for the status window's own row - which
///  is where one looks to find out whether a reload will be able to put the
///  window back.</summary>
function StatusDockWindowStateText: string;

/// <summary>Opens the status window free-standing at ABounds (screen
///  coordinates). True when the window really is back where it was; the
///  caller reports a False, because a box that sends the user to the desktop
///  menu must not appear while the window is already sitting in its
///  place.
///  It is also the LAST RESORT after a docked restore failed - and there the
///  caller deliberately discards the answer: the window is at the place it
///  occupied, but free-standing, and that is not the layout it had.</summary>
function ShowStatusWindowAt(const ABounds: TRect): Boolean;

/// <summary>Opens the status window and hands it back to the dock chain it
///  came from - the host that held it if that host is still there, and
///  otherwise the site that host was itself docked in, which is the same
///  place on screen.
///  A link's site name may be empty (an older marker, or the saved desktop as
///  the fallback) - a tab host can be docked into through its own form.
///
///  DELIBERATELY NO BOUNDS HERE, and that is the whole reason this is a
///  second entry point rather than an argument: putting a window at a
///  remembered rectangle and docking it into a site are two different
///  outcomes, and only the second one is a restore. The caller does the
///  first one afterwards, on purpose, when this answered anything but
///  rwoRestored.
///  AND THE ANSWER IS THREE-VALUED since 2.9.0: rwoPlaceOnly means the window
///  is a client of the right dock site and no longer a TAB beside the window
///  it shared a tab group with - the state a user reported on 2026-10-28 and
///  which the old Boolean reported as a success.</summary>
function ShowStatusWindowDocked(const APlace: TDockPlace): TReloadWindowOutcome;

/// <summary>Menu entry point: shows (and focuses) the status window.</summary>
procedure ShowStatusWindow;
/// <summary>Menu entry point: the MCP tools window - every tool the bridge
///  offers, whether this IDE handles it, and what its calls did.</summary>
procedure ShowMcpToolsWindow;

type
  TStatusRow = record
    Caption, Value, Detail: string;
    /// <summary>How the row reads at a glance - the frame paints the status
    ///  cell in the matching colour (user request 2026-09-30). Rows that are
    ///  plain information stay slNeutral.</summary>
    Level: TStatusLevel;
  end;

/// <summary>The rows the status window shows, collected NOW - whether the
///  window is open or not (the MCP tool get_status). MAIN THREAD ONLY, like
///  every ToolsAPI read in there.</summary>
function StatusSnapshot: TArray<TStatusRow>;

implementation

uses
  Expert.WorkerLatch,
  Expert.ResourceMonitor, Expert.CompletionWizard, Expert.LineEndingWatcher,
  System.SysUtils, System.Classes, System.IniFiles, System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Graphics,
  Vcl.ActnList, Vcl.ImgList, Vcl.Menus,
  ToolsAPI, DesignIntf,   // DesignIntf: TEditState / TEditAction
  // THE IDE'S OWN DOCKING UNITS. They ship without source but live in
  // designide.dcp, which this package already requires - verified by
  // compiling against them (scratchpad mmx\DockProbe): DeskForm gives
  // TDesktopForm with DeskSection, DockForm gives TDockableForm, TabDock
  // gives TTabDockHostForm.ShowDockClient and DeskUtil gives
  // ShowDockableForm(TCustomForm) + FocusWindow(TCustomForm).
  DeskForm, DockForm, TabDock, DeskUtil,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.LspManager, Expert.PluginSettings,
  Expert.AutoImport, Expert.ContextMenu, Expert.DialogHelper,   // IdeThemes: interface uses
  Expert.BlameGutter, Expert.BlameDialogs, Expert.VcsBlame,
  Expert.MessagesReader, Expert.StructureErrors, Expert.McpServer, Lsp.Client,
  Expert.Version, Expert.ListViewSort, Mcp.Protocol, Expert.DebugConsistency,
  System.JSON, System.StrUtils, Vcl.StdCtrls, Winapi.Windows;

// TCustomFrame.Create does InitInheritedComponent(Self, TFrame) for every
// descendant and RAISES EResNotFound when NO class of its chain has a DFM
// resource - so even a fully code-built frame needs one. Expert.StatusWindow
// .dfm is an empty frame of the common BASE class TRlDockFrame: one
// resource serves every dockable window of this unit (all controls are
// still created in the constructors).
{$R *.dfm}

type
  TRlDockFrame = class(TFrame)
  end;

  // Collects the rows. Separate from the frame, so the MCP bridge can ask
  // for the same rows while no status window is open (StatusSnapshot).
  TStatusCollector = class
  private
    // Collected rows of the CURRENT tick. The row SET is fixed (same
    // count, same order, always) - only the cell texts change, so the
    // refresh can write single cells instead of rebuilding the list.
    // Rebuilding made the view flicker once a second and dropped the
    // user's selection every time.
    FRows: TArray<TStatusRow>;
    FRowCount: Integer;
    FProbeSeen: Boolean;   // messages log found once - stop stat()ing it
    FTicks: Integer;
    FMsgCount: Integer;    // last compiler-message count (see Collect)
    FMsgCountTick: Integer;
    FResSample: TResourceSample;   // refreshed every 5th tick (VA walk)
    FResValid: Boolean;
    FMemText, FMemDetail: string;  // refreshed every 10th tick (walks caches)
    FMemValid: Boolean;
    FLspDcuDir: string;    // the DCU output that configuration names
    FLspDcuCount: Integer; // .dcu lying there - how many, against
                           // FLspDcuUnits: a COMPLETE set is what makes
                           // cross-unit answers work (measured)
    FLspDcuValid: Boolean; // ...looked up at all yet? The SNAPSHOT
                           // collector ticks once per get_status call,
                           // so "every tenth tick" alone answered
                           // nine times out of ten from nothing.
    FLspJson: string;      // .delphilsp.json of the current project,
    FLspJsonValid: Boolean;  // looked up every 10th tick (disk + modules)
    FLspDcuUnits: Integer;   // units the configuration lists
    procedure Row(const ACaption, AValue, ADetail: string;
      ALevel: TStatusLevel = slNeutral);
  public
    /// <summary>One refresh round: the expensive rows re-sample every Nth
    ///  call, so calling it about once a second is what it expects.</summary>
    procedure Collect;
  end;

  // The frame the IDE embeds into its dockable form.
  TStatusFrame = class(TRlDockFrame)
  private
    FList: TListView;
    // The painter runs from the control's own message loop, not from our
    // tick, so it must not touch the collector - it reads this copy, which
    // Apply keeps in step with the visible rows.
    FLevels: TArray<TStatusLevel>;
    FPopup: TPopupMenu;
    FMniAdjust: TMenuItem;
    FTimer: TTimer;
    FC: TStatusCollector;
    procedure DoDrawSubItem(Sender: TCustomListView; AItem: TListItem;
      ASubItem: Integer; AState: TCustomDrawState; var ADefaultDraw: Boolean);
    procedure DoTick(Sender: TObject);
    procedure DoListDblClick(Sender: TObject);
    procedure DoAdjustBlameClick(Sender: TObject);
    procedure DoPopup(Sender: TObject);
    procedure Apply;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

  // The MCP tools window: one row per tool of the bridge's list.
  TMcpToolsFrame = class(TRlDockFrame)
  private type
    TToolDef = record
      Name, Description: string;
      Bridge: Boolean;   // handled by the bridge itself, never reaches an IDE
    end;
  private
    FHeader: TLabel;
    FList: TListView;
    FDetail: TMemo;
    FTimer: TTimer;
    FDefs: TArray<TToolDef>;
    FDetailFor: string;    // tool + stats shown in the memo (skip rewrites)
    FThemeTick: Integer;   // re-apply the theme every 10th tick, not every one
    procedure LoadDefs;
    procedure DoTick(Sender: TObject);
    procedure DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure RefreshRows;
    procedure ShowDetail;
    function StatOf(const AName: string; const AStats: TArray<TMcpToolStat>;
      out AStat: TMcpToolStat): Boolean;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

  // One class for both windows - they differ only in caption, identifier
  // and frame.
  TStatusDockable = class(TInterfacedObject, INTACustomDockableForm)
  private
    FCaption, FIdent: string;
    FFrameClass: TCustomFrameClass;
  public
    constructor Create(const ACaption, AIdent: string; AFrameClass: TCustomFrameClass);
    function GetCaption: string;
    function GetIdentifier: string;
    function GetFrameClass: TCustomFrameClass;
    procedure FrameCreated(AFrame: TCustomFrame);
    function GetMenuActionList: TCustomActionList;
    function GetMenuImageList: TCustomImageList;
    procedure CustomizePopupMenu(PopupMenu: TPopupMenu);
    function GetToolBarActionList: TCustomActionList;
    function GetToolBarImageList: TCustomImageList;
    procedure CustomizeToolBar(ToolBar: TToolBar);
    procedure SaveWindowState(Desktop: TCustomIniFile; const Section: string;
      IsProject: Boolean);
    procedure LoadWindowState(Desktop: TCustomIniFile; const Section: string);
    function GetEditState: TEditState;
    function EditAction(Action: TEditAction): Boolean;
  end;

var
  /// <summary>The dock CHAIN we last saw the status window in - the host
  ///  that held it and the site that host sat in. Filled by
  ///  StatusDockWindowState, which the status row asks once a second, so it
  ///  is known long before the window is closed, and it is what lets a reopen
  ///  put back a window the IDE handed back undocked. It is deliberately NOT
  ///  cleared when the user floats the window: ShowFormDockAware only uses it
  ///  for a window that was not already on screen, which is the case it is
  ///  for.</summary>
  GLastStatusDock: TDockPlace;
  GDockable: INTACustomDockableForm;
  GForm: TCustomForm;
  GRegistered: Boolean;
  GToolsDockable: INTACustomDockableForm;
  GToolsForm: TCustomForm;
  GToolsRegistered: Boolean;

/// <summary>Records / clears the dock window that hosts a frame
///  (audit #37, M30). AForm is the global the dock pair uses; the frame's
///  own Owner chain gives the host form.</summary>
procedure NoteDockHost(AFrame: TFrame; var AForm: TCustomForm;
  AAttached: Boolean);
var
  Host: TCustomForm;
begin
  Host := nil;
  var C: TComponent := AFrame;
  while (C <> nil) and (Host = nil) do
  begin
    if C is TCustomForm then Host := TCustomForm(C);
    C := C.Owner;
  end;
  if AAttached then
  begin
    if Host <> nil then AForm := Host;
  end
  else
    // the IDE may free the window itself - never leave a dangling pointer
    if (Host = nil) or (AForm = Host) then AForm := nil;
end;

{ TStatusFrame }

constructor TStatusFrame.Create(AOwner: TComponent);
begin
  inherited;
  Name := '';   // the IDE names the embedded instance
  NoteDockHost(Self, GForm, True);
  FC := TStatusCollector.Create;

  FList := TThemedListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.GridLines := True;
  // The blame column has to be lined up against whatever else draws in
  // the gutter, and pixels are not something to guess in an options page:
  // a double-click (or the context menu) on the "Live blame" row opens the
  // live adjuster.
  FList.OnDblClick := DoListDblClick;
  // Colour is the whole point of the request ("mit Farben und einer einfach
  // verstaendlichen Statuszeile"): the STATUS cell is drawn in the level's
  // colour, the rest stays as the theme wants it.
  FList.OnCustomDrawSubItem := DoDrawSubItem;
  FPopup := TPopupMenu.Create(Self);
  FPopup.OnPopup := DoPopup;
  FMniAdjust := TMenuItem.Create(FPopup);
  FMniAdjust.Caption := 'Adjust live blame column...';
  FMniAdjust.OnClick := DoAdjustBlameClick;
  FPopup.Items.Add(FMniAdjust);
  FList.PopupMenu := FPopup;
  FList.Columns.Add.Caption := 'Item';
  FList.Columns[0].Width := 160;
  FList.Columns.Add.Caption := 'Status';
  FList.Columns[1].Width := 180;
  FList.Columns.Add.Caption := 'Details';
  FList.Columns[2].Width := 460;

  FTimer := TTimer.Create(Self);
  FTimer.Interval := 1000;
  FTimer.OnTimer := DoTick;
  FTimer.Enabled := True;
  // The IDE themes the form chrome around us but not this frame's children -
  // measured: the item column stayed white while the custom-drawn cells were
  // dark. So the colours are ours to set.
  ApplyThemeToControls(Self);
  FC.Collect;
  Apply;
end;

destructor TStatusFrame.Destroy;
begin
  NoteDockHost(Self, GForm, False);
  FreeAndNil(FTimer);
  inherited;
  FreeAndNil(FC);
end;

procedure TStatusCollector.Row(const ACaption, AValue, ADetail: string;
  ALevel: TStatusLevel);
begin
  if FRowCount >= Length(FRows) then
    SetLength(FRows, FRowCount + 8);
  FRows[FRowCount].Caption := ACaption;
  FRows[FRowCount].Value := AValue;
  FRows[FRowCount].Detail := ADetail;
  FRows[FRowCount].Level := ALevel;
  Inc(FRowCount);
end;

function TStatusFrame_IsBlameRow(AItem: TListItem): Boolean;
begin
  Result := (AItem <> nil) and SameText(Trim(AItem.Caption), 'Live blame');
end;

procedure TStatusFrame.DoListDblClick(Sender: TObject);
begin
  if TStatusFrame_IsBlameRow(FList.Selected) then
    AdjustBlameColumn
  else if (FList.Selected <> nil) and
    (SameText(Trim(FList.Selected.Caption), 'MCP bridge endpoint') or
     SameText(Trim(FList.Selected.Caption), 'Claude Code')) then
    ShowMcpToolsWindow;
end;

procedure TStatusFrame.DoAdjustBlameClick(Sender: TObject);
begin
  AdjustBlameColumn;
end;

procedure TStatusFrame.DoPopup(Sender: TObject);
begin
  // Only offer it where it means something.
  FMniAdjust.Enabled := TStatusFrame_IsBlameRow(FList.Selected);
end;

// Writes the collected rows into the list view - and ONLY what actually
// changed. In the normal case (nothing moved since the last tick) not a
// single assignment happens, so the control never repaints and the
// selection survives.
// HARD RULE of this code base: a handler that touches the canvas MUST set
// Brush.Color from the theme - the native control then takes the row
// background from it, and the unthemed default is WHITE (that is how the
// GUID and circular-reference lists ended up white-on-dark).
procedure TStatusFrame.DoDrawSubItem(Sender: TCustomListView; AItem: TListItem;
  ASubItem: Integer; AState: TCustomDrawState; var ADefaultDraw: Boolean);
var
  Lvl: TStatusLevel;
  Want: TColor;
begin
  ADefaultDraw := True;
  if cdsSelected in AState then Exit;   // the selection owns background AND text
  // The control's OWN colour, not a freshly asked themed one: the columns this
  // handler does not touch are painted by the control with exactly that, so
  // taking it from there is the only way the row cannot end up two-toned.
  Sender.Canvas.Brush.Color := TListView(Sender).Color;
  // The VCL resets its Canvas.Font per subitem, but the control only
  // RE-SELECTS the font into the device context when something really changed.
  // Assigning the colour the canvas already holds is therefore a no-op, and
  // the DC keeps what the STATUS cell selected - which is why the colour bled
  // into the details column since 1.15.0. Forcing one different value first is
  // what makes the reset arrive. (Found by rendering the list and measuring
  // the text colour - the code reads correct either way.)
  Want := TListView(Sender).Font.Color;
  if Sender.Canvas.Font.Color = Want then
    Sender.Canvas.Font.Color := TColor(Cardinal(Want) xor $010101);
  Sender.Canvas.Font.Color := Want;
  Sender.Canvas.Font.Style := Sender.Canvas.Font.Style - [fsBold];
  if (AItem = nil) or (AItem.Index < 0) or (AItem.Index > High(FLevels)) then Exit;
  // column 1 is "Status" - the one word the user should be able to read
  // without reading anything else
  if ASubItem <> 1 then Exit;
  Lvl := FLevels[AItem.Index];
  if Lvl = slNeutral then Exit;
  Sender.Canvas.Font.Color := StatusLevelColor(Lvl,
    TListView(Sender).Color, TListView(Sender).Font.Color);
  Sender.Canvas.Font.Style := Sender.Canvas.Font.Style + [fsBold];
end;

procedure TStatusFrame.Apply;
var
  I: Integer;
  It: TListItem;
begin
  // the painter's copy of the levels, and a repaint ONLY when one changed
  // (the tick must not touch windows otherwise - the hint's CPU storm)
  var LevelsChanged := Length(FLevels) <> FC.FRowCount;
  if LevelsChanged then SetLength(FLevels, FC.FRowCount);
  for I := 0 to FC.FRowCount - 1 do
    if FLevels[I] <> FC.FRows[I].Level then
    begin
      FLevels[I] := FC.FRows[I].Level;
      LevelsChanged := True;
    end;

  if FList.Items.Count <> FC.FRowCount then
  begin
    // Structural change (should not happen - the row set is fixed).
    FList.Items.BeginUpdate;
    try
      FList.Items.Clear;
      for I := 0 to FC.FRowCount - 1 do
      begin
        It := FList.Items.Add;
        It.Caption := FC.FRows[I].Caption;
        It.SubItems.Add(FC.FRows[I].Value);
        It.SubItems.Add(FC.FRows[I].Detail);
      end;
    finally
      FList.Items.EndUpdate;
    end;
    Exit;
  end;

  for I := 0 to FC.FRowCount - 1 do
  begin
    It := FList.Items[I];
    if It.Caption <> FC.FRows[I].Caption then
      It.Caption := FC.FRows[I].Caption;
    if It.SubItems.Count > 0 then
    begin
      if It.SubItems[0] <> FC.FRows[I].Value then
        It.SubItems[0] := FC.FRows[I].Value;
      if (It.SubItems.Count > 1) and (It.SubItems[1] <> FC.FRows[I].Detail) then
        It.SubItems[1] := FC.FRows[I].Detail;
    end;
  end;
  if LevelsChanged then FList.Invalidate;
end;

// What THIS plugin keeps in memory, per consumer. The plugin shares the
// IDE's memory manager, so this cannot be measured - it is estimated from
// the data structures (Expert.ResourceMonitor's heap-size helpers). Asked
// for by a user whose IDE needs 2 GB for a large project on its own and
// then runs out of memory with the plugin loaded.
function PluginMemoryText(out ADetail: string): string;
var
  Idx: TIndexMemory;
  Lsp, Blame, Gutter, Live, IdxLib, IdxPrj, Total: Int64;
  BlFiles, BlLines: Integer;
  Client: TLspClient;
begin
  Idx := TUnitIndex.Instance.MemoryInfo;
  Lsp := 0;
  if TLspManager.Instance.IsAlive then
  begin
    Client := TLspManager.Instance.PeekClient;
    if Client <> nil then
      try
        Lsp := Client.EstimateRetainedBytes;
      except
        Lsp := 0;
      end;
  end;
  Blame := BlameCacheBytes(BlFiles, BlLines);
  Gutter := BlameGutterBytes;
  Live := LiveMemoryBytes;
  IdxLib := Idx.GlobalRawBytes + Idx.GlobalLayerBytes;
  IdxPrj := Idx.ProjectRawBytes + Idx.ProjectLayerBytes;
  Total := IdxLib + IdxPrj + Lsp + Blame + Gutter + Live;
  Result := Format('~%s MB: index library %s, project %s | LSP diagnostics %s | ' +
    'blame %s | live checker %s',
    [MBText(Total), MBText(IdxLib), MBText(IdxPrj), MBText(Lsp),
     MBText(Blame + Gutter), MBText(Live)]);
  ADetail := Format('ESTIMATE from our data structures (MB), refreshed every 10 s. ' +
    'Library index: %d units, %d identifiers = identifier lists %s + lookup %s ' +
    '(a library rescan builds the lookup a second time for a moment). ' +
    'Project index: %d units, lists %s + lookup %s. Blame: %d files, %d lines. ' +
    'DelphiLSP itself runs as a separate process and is NOT included.',
    [Idx.GlobalUnits, Idx.GlobalIdents, MBText(Idx.GlobalRawBytes),
     MBText(Idx.GlobalLayerBytes), Idx.ProjectUnits, MBText(Idx.ProjectRawBytes),
     MBText(Idx.ProjectLayerBytes), BlFiles, BlLines]);
end;

procedure TStatusCollector.Collect;
var
  Client: TLspClient;
  LiveFile, S, Detail, DiagCodes: string;
  Analysing, Resolving, FromLsp, Fresh: Boolean;
  FixCount, DiagCount, DiagFiles, DiagSeen, DiagHandled: Integer;
  StrFires, StrNodes, StrDiags: Integer;
  StrInstalled: Boolean;
  StrReason, SrcStructure, SrcLsp: string;
begin
  // FIXED row set: every branch below fills the same rows in the same
  // order, so Apply never has to restructure the list.
  FRowCount := 0;

  // ---- the plugin itself ---------------------------------------------------
  // First row: which build is actually loaded answers half of all "does not
  // work for me" reports.
  try
    Detail := GetModuleName(HInstance) + ', built ' +
      FormatDateTime('yyyy-mm-dd hh:nn', TFile.GetLastWriteTime(GetModuleName(HInstance)));
  except
    Detail := GetModuleName(HInstance);
  end;
  Row(PluginName, 'version ' + PluginVersion, Detail);

  // ---- process resources --------------------------------------------------
  // First, because it is what the next "out of memory" report needs:
  // EPNGOutMemory means CreateDIBSection failed - out of GDI objects, or a
  // 32-bit address space too fragmented to map the bitmap.
  if not FResValid or (FTicks mod 5 = 0) then
  begin
    FResSample := SampleResources(True);
    FResValid := True;
  end;
  Row('Process resources',
    Format('GDI %d, USER %d, private %d MB',
      [FResSample.GdiObjects, FResSample.UserObjects, FResSample.PrivateMB]),
    Format('peaks: GDI %d, USER %d - Windows allows 10,000 of each per ' +
      'process; trend log: %%TEMP%%\RefactoringLight-resources.log',
      [FResSample.GdiPeak, FResSample.UserPeak]));
  if FResSample.LargestFreeMB >= 0 then
    Row('  address space (32-bit)',
      Format('%d MB free, largest block %d MB',
        [FResSample.FreeTotalMB, FResSample.LargestFreeMB]),
      'large allocations (bitmaps, big arrays) fail once the largest ' +
      'block gets small - the IDE then reports "out of memory"')
  else
    Row('  address space', '64-bit process', 'not a limiting factor');
  if not FMemValid or (FTicks mod 10 = 0) then
  begin
    FMemText := PluginMemoryText(FMemDetail);
    FMemValid := True;
  end;
  Row('  memory held by this plugin', FMemText, FMemDetail);
  Row('  GDI balance of this plugin', GdiBalanceText,
    'objects our own ticks / paint handlers created and did not release ' +
    'since the IDE started - a number that keeps growing is a leak there');

  // ---- code completion: generated entries ---------------------------------
  // Why the last completion call did (not) offer an event handler / an
  // anonymous method - otherwise "no entry" looks exactly like "broken".
  Row('Completion: generated entries', CompletionGenerationNote,
    'last code completion call; history in %TEMP%\RefactoringLight-completion.log');
  Row('MCP bridge endpoint', McpServerStatus,
    McpServerPipe + ' - used by RefactoringLightMcp.exe (Claude Code)');
  S := McpConnectionStatus(Detail);
  if McpBridgeProblem <> '' then
    Row('  Claude Code', S, Detail, slBad)
  else
    Row('  Claude Code', S, Detail);

  // ---- identifier index ---------------------------------------------------
  if TUnitIndex.Instance.Ready then S := 'ready' else S := 'building...';
  Row('Identifier index', S, TUnitIndex.Instance.StatusLine);
  Row('  scan cycles', IntToStr(TUnitIndex.Instance.ScanCycle),
    'one per completed worker pass (project scope re-scans every 30 s)');

  // Unresolved IDE path variables mean whole libraries are missing from
  // the index (a DevExpress "$(DXVCL)\..." entry cost the tester every
  // DevExpress symbol) - surface it instead of failing silently.
  S := UnresolvedPathVars;
  if S = '' then
    Row('  path variables', 'all resolved', '')
  else
    Row('  path variables', 'UNRESOLVED: ' + S,
      'directories behind these are NOT indexed - define them under ' +
      'Tools > Options > IDE > Environment Variables');

  // A LIBRARY SCOPE OF ZERO UNITS IS A FAILURE, and it has to say so: every
  // RTL / VCL / third-party identifier is then unknown, so find unit, the
  // add-unit fixes, "did you mean" and the uses cleanup only ever know this
  // project - while the row above still reads "ready". That is exactly what
  // happened between 1.16.24 and 1.16.30 (a wrong registry key, found by
  // reading the memory row rather than this one).
  var Mem := TUnitIndex.Instance.MemoryInfo;
  if Mem.GlobalUnits > 0 then
    Row('  library scope', Format('%d unit(s), %d identifier(s)',
      [Mem.GlobalUnits, Mem.GlobalIdents]),
      'RTL, VCL and third-party sources from the IDE Library and Browsing paths')
  else if TUnitIndex.Instance.Ready then
    Row('  library scope', 'EMPTY - not one library unit is indexed',
      'the IDE Library / Browsing paths gave no directory, so RTL, VCL and ' +
      'third-party identifiers are missing: find unit, "add unit" and the ' +
      'uses cleanup only know this project''s own units', slBad)
  else
    Row('  library scope', 'building...', '');

  // ---- LSP session --------------------------------------------------------
  // ONE LINE IN PLAIN WORDS FIRST (user request 2026-09-30): what the
  // session means for the features, not what it consists of. The rows below
  // keep the numbers for a closer look. Cheap getters only - this runs once
  // a second and must never send a request.
  Client := nil;   // a local object reference is NOT zero-initialised
  DiagCount := 0;
  DiagFiles := 0;
  if TLspManager.Instance.IsAlive then
  begin
    Client := TLspManager.Instance.PeekClient;
    if Client <> nil then
      try
        DiagCount := Client.GetDiagnosticsCount;
        DiagFiles := Client.GetDiagnosticFileCount;
      except
      end;
  end;
  var Busy := '';
  if Client <> nil then
    try Busy := Client.BusyWith; except end;
  var Answered := 0;
  if Client <> nil then
    try Answered := Client.GetAnsweredCount; except end;

  if Client = nil then
    Row('DelphiLSP', 'not started yet',
      'It starts with the first search, rename, completion or quick fix - ' +
      'nothing is wrong here.', slNeutral)
  else if not Client.IsConnected then
    Row('DelphiLSP', 'CONNECTION LOST',
      'DelphiLsp.exe is gone. The next request starts a new session; if it ' +
      'keeps happening, switch the LSP log on in the options and send the ' +
      'log.', slBad)
  else if Busy <> '' then
    Row('DelphiLSP', 'BUSY - loading the project',
      Format('It is working on "%s". A search started now would have to wait ' +
        'for it: while the project loads, requests are answered with nothing.',
        [Busy]), slWait)
  else if Answered = 0 then
    Row('DelphiLSP', 'starting up - not answering yet',
      'The session is running but has not answered a single request. That is ' +
      'the state right after the IDE starts: give it a moment, a search ' +
      'started now waits for it.', slWait)
  else if DiagCount = 0 then
    Row('DelphiLSP', 'ready for searches, no diagnostics',
      'Searches, rename and navigation work. Only the compiler HINTS ' +
      '(H2443 "unit needed", unused variables, ...) are missing - our session ' +
      'has never pushed one; errors still come from the Structure view and ' +
      'the compiler.', slWait)
  else
    Row('DelphiLSP', 'ready',
      Format('Everything works: %d answer(s) given, %d diagnostic push(es) ' +
        'received for %d file(s).', [Answered, DiagCount, DiagFiles]), slGood);

  // WITHOUT THAT FILE NOTHING ABOVE CAN EVER WORK, and the user cannot guess
  // it: the IDE generates it per project, but only while the Code Insight
  // option is on (forum screenshot 2026-10-02 - a dialog that said no more
  // than "No .delphilsp.json found."). The lookup reads the project and the
  // disk, so it runs every 10th tick like the memory row.
  if not FLspJsonValid or (FTicks mod 10 = 0) then
  begin
    FLspJson := '';
    try
      if Editor <> nil then FLspJson := Editor.FindDelphiLspJson;
    except
      FLspJson := '';
    end;
    FLspJsonValid := True;
  end;
  if FLspJson <> '' then
    Row('  project configuration', 'found', FLspJson)
  else
    Row('  project configuration', '.delphilsp.json is MISSING',
      LspConfigMissingHintLong, slBad);

  // WHERE THE SESSION GETS THE OTHER UNITS FROM. Measured 2026-10-20,
  // eight states with one variable changed each: a unit the LSP loads from
  // a .dcu DOES answer with its source position - so the number of .dcu
  // here is not the problem. What broke go to definition was this
  // directory NOT MATCHING the project: with one unit's .dcu missing from
  // it, that unit came from another build and the answer was empty
  // ("Kibitz definition found in file:///(-1:0)" in the session log);
  // putting the file back made it answer again. One directory listing and
  // one read of the configuration, on the same 10th tick as above.
  if FLspJson <> '' then
  begin
    if not FLspDcuValid or (FTicks mod 10 = 0) then
    begin
      FLspDcuValid := True;
      FLspDcuCount := 0;
      FLspDcuUnits := 0;
      try
        FLspDcuDir := LspDcuDirOfConfig(FLspJson,
          ExcludeTrailingPathDelimiter(ExtractFilePath(FLspJson)));
        if (FLspDcuDir <> '') and TDirectory.Exists(FLspDcuDir) then
          FLspDcuCount := Length(TDirectory.GetFiles(FLspDcuDir, '*.dcu'));
        FLspDcuUnits := LspProjectUnitCount(FLspJson);
      except
        FLspDcuDir := '';
        FLspDcuCount := 0;
        FLspDcuUnits := 0;
      end;
    end;
    if FLspDcuDir = '' then
      Row('  DCU output of that configuration', 'none named',
        'The session resolves the project''s other units from the search ' +
        'path or from the sources.', slNeutral)
    else if FLspDcuCount = 0 then
      Row('  DCU output of that configuration', 'empty',
        FLspDcuDir + ' holds no .dcu, so every other unit comes from the ' +
        'same place - the search path or the sources. That is consistent; ' +
        'a build of this configuration fills the directory.', slNeutral)
    else if (FLspDcuUnits > 0) and (FLspDcuCount < FLspDcuUnits) then
      Row('  DCU output of that configuration',
        Format('INCOMPLETE - %d .dcu for %d unit(s)',
          [FLspDcuCount, FLspDcuUnits]),
        FLspDcuDir + ' - the LSP searches it the way the compiler does. ' +
        'MEASURED: a unit MISSING here comes from another build instead ' +
        '(a copy on the library path, the installed package''s .dcp), and ' +
        'then go to definition, find references and rename can answer ' +
        'nothing for its symbols. Build this configuration - removing ' +
        '.dcu files makes it worse, not better. Project checks > Debug ' +
        'consistency names what is missing.', slBad)
    else
      Row('  DCU output of that configuration',
        Format('%d .dcu for %d unit(s)', [FLspDcuCount, FLspDcuUnits]),
        FLspDcuDir + ' - the LSP reads it the way the compiler does, and ' +
        'it matches the project: measured, that is the state in which go ' +
        'to definition across units answers with the source position. ' +
        'Keep it built; a .dcu older than its source is what Project ' +
        'checks > Debug consistency reports.', slGood);
  end;

  // The VERIFICATION session is the one that answers the scans since 1.10.0,
  // and it pushes no diagnostics at all - so its own line, or a freshly
  // started one looks "fine" while it answers nothing (forum 2026-09-30).
  var VClient := TLspManager.Instance.PeekVerifyClient;
  if VClient = nil then
    Row('  verification session', 'not started yet',
      'A separate, faster session for "find references" and rename; it ' +
      'starts with the first one of them.', slNeutral)
  else
  begin
    var VAnswered := 0;
    var VBusy := '';
    try
      VAnswered := VClient.GetAnsweredCount;
      VBusy := VClient.BusyWith;
    except
    end;
    if not VClient.IsConnected then
      Row('  verification session', 'CONNECTION LOST',
        'The scans fall back to the session above.', slBad)
    else if VBusy <> '' then
      Row('  verification session', 'BUSY - loading the project',
        Format('Working on "%s". A search waits for it and says so.', [VBusy]), slWait)
    else if VAnswered = 0 then
      Row('  verification session', 'starting up - not answering yet',
        'It has not answered anything yet. A search started now waits for it ' +
        'first (up to 2 minutes) and then lets the main session verify.', slWait)
    else
      Row('  verification session', 'ready',
        Format('%d answer(s) given. It carries the per-occurrence checks of ' +
          '"find references" and rename.', [VAnswered]), slGood);
  end;
  if Client <> nil then
  begin
    if DiagCount = 0 then
      Row('  diagnostics', 'none received',
        'our session has never pushed one - hints (H2443, ...) are ' +
        'unavailable; errors come from the Structure view / the compiler',
        slWait)
    else
      Row('  diagnostics', Format('%d push(es), %d file(s)',
        [DiagCount, DiagFiles]),
        'held for ALL files the session has seen, not just this one', slGood);
  end
  else
    Row('  diagnostics', '-', '');
  // Inactive {$IFDEF} regions ride along with those pushes. Remove-with
  // skips occurrences inside them, and the uses cleanup judges a conditional
  // clause per entry with them - so "0" means both fall back to their
  // conservative behaviour, and that is worth seeing.
  if Client <> nil then
  begin
    var Inact := 0;
    try Inact := Client.GetInactiveRangesTotal; except end;
    if Inact = 0 then
      Row('  inactive regions', 'none known',
        '{$IFDEF} branches the compiler skips are reported with the ' +
        'diagnostics; without them remove-with and the uses cleanup stay ' +
        'conservative')
    else
      Row('  inactive regions', Format('%d range(s)', [Inact]),
        'branches the compiler skips - remove-with leaves those occurrences ' +
        'alone, the uses cleanup never judges an entry on such a line',
        slGood);
  end
  else
    Row('  inactive regions', '-', '');
  if TLspManager.Instance.ProjectIndexed then S := 'yes' else S := 'no';
  Row('  project indexed', S, 'the LSP has seen this project once');
  // What the server is doing RIGHT NOW: while it loads a project (12-30 s
  // for a big one) the controller aborts every request after 10 s, so this
  // row explains a search that seems to find nothing (issue #13).
  S := '';
  if Client <> nil then
    try S := Client.BusyWith; except end;
  if S <> '' then
    Row('  server busy', S,
      'requests are aborted while this runs - the scans wait for it')
  else if (Client <> nil) and not Client.ReportsProgress then
    Row('  server busy', 'no (it reports no progress)',
      'this server does not send $/progress, so "busy" cannot be seen')
  else
    Row('  server busy', 'no', 'it answers requests');
  S := '';
  // The STATE of that session is the plain-language row further up - this one
  // only says how it is CONFIGURED (two rows of the same name looked like a
  // bug when the window was first seen with a real project, 2026-09-30).
  if TPluginSettings.VerifySession then
    Row('  verification setting', 'on',
      'a second DelphiLsp process answers the candidate checks of find ' +
      'references / rename - about twice as fast, never aborted after 10 s, ' +
      'shut down again after 10 idle minutes')
  else
    Row('  verification setting', 'off',
      'the main session verifies - slower, and it can be aborted while a ' +
      'project loads');
  if TPluginSettings.LspLogging then
    Row('  session log', 'on',
      TPath.Combine(TPath.Combine(TPath.GetTempPath, 'DelphiLSP'),
        'RefactoringLight*.log') + ' - every request with its duration')
  else
    Row('  session log', 'off', 'switch it on in the options to diagnose a slow session');

  // ToolsAPI is MAIN THREAD ONLY, and breaking that rule does not fail -
  // it corrupts a buffer now and then. Every write path reports itself, so
  // a violation is a NUMBER here instead of a rare mystery (fork audit,
  // 2026-10: three MCP tools apply their edit from the pipe thread).
  if MainThreadViolations = 0 then
    Row('main-thread rule', 'kept', 'no editor write came from a worker thread')
  else
    Row('main-thread rule', Format('%d violation(s)', [MainThreadViolations]),
      Format('last: %s - ToolsAPI and the editor buffers are main thread ' +
      'only; this is a DEFECT in the plugin, please report it',
      [LastOffMainThreadCall]), slBad);

  // What QueueSafe swallowed. It HAS to swallow - an exception escaping a
  // queued closure makes CheckSynchronize open a modal box inside whatever
  // IDE transaction pumped the queue - but swallowing is only defensible
  // with a channel, and this is it.
  if QueuedErrorCount = 0 then
    Row('queued callbacks', 'clean', 'no deferred callback of this plugin ' +
      'has raised')
  else
    Row('queued callbacks', Format('%d raised', [QueuedErrorCount]),
      Format('last: %s - swallowed on purpose (an escaping exception would ' +
      'open a modal box inside an IDE transaction), but it is a DEFECT in ' +
      'the plugin: please report it', [LastQueuedError]), slBad);

  // THIS WINDOW ITSELF, and the row exists because the answer decides what a
  // reload can restore. It names the CHAIN, not one host: "host unknown"
  // means the hand-over could read nothing, not that nothing can be restored
  // - the saved desktop is asked next, and a remembered rectangle after that.
  Row('this window', StatusDockWindowStateText,
    'when the package is installed again, a docked window goes back into the ' +
    'host that held it; when that host is gone - a tab host the user made by ' +
    'dragging dissolves - the TAB GROUP is rebuilt beside the neighbour named ' +
    'here, and failing that the site that host was docked in, or where the ' +
    'saved desktop says, or free-standing over the place it occupied. The ' +
    'report after an install names it when only the place came back (see ' +
    '"Installieren")');

  // Line endings at file open: the check corrects files on disk, so what it
  // did (and whether it is asking at all) must be readable somewhere - this
  // is where one looks when a file was silently rewritten, or was not.
  Row('line endings at open', LineEndingWatcherStatus,
    'LF-only or mixed line endings are reported when the IDE opens a source ' +
    'file; "Line endings" in the options page switches it and the automatic ' +
    'correction');

  // ---- live quick-fix checker --------------------------------------------
  LiveStatusInfo(LiveFile, Analysing, Resolving, FromLsp, Fresh, FixCount);
  // A checker that stopped itself says so, in red (audit #37, M28).
  var Stopped := LiveStoppedReason;
  if Stopped <> '' then
  begin
    Row('Live checker', 'STOPPED after an error', Stopped, slBad);
    Row('  diagnostics from', '-', 'nothing is being analysed any more');
  end
  else if LiveFile = '' then
  begin
    Row('Live checker', 'idle', 'no source buffer active');
    Row('  diagnostics from', '-', '');
  end
  else
  begin
    if Analysing then S := 'analysing...'
    else if Resolving then S := 'resolving...'
    else if Fresh then S := Format('%d fix(es)', [FixCount])
    else S := 'waiting for analysis';
    Row('Live checker', S, ExtractFileName(LiveFile));
    if not Fresh then
      Row('  diagnostics from', '-',
        'published results belong to an older buffer state')
    else if FromLsp then
      Row('  diagnostics from', 'own LSP session',
        'errors AND hints (H2443, H2164, ...)')
    else
      Row('  diagnostics from', 'Structure view',
        'errors only - hints follow after the LSP pass or a compile');
  end;

  // What the LAST resolution actually worked on - turns a bare
  // "0 fix(es)" into an answer to why (no diagnostics at all? codes we
  // have no provider for? or providers that declined?).
  LiveDiagStats(DiagSeen, DiagHandled, DiagCodes);
  if DiagSeen = 0 then
    Row('  last resolution', 'no diagnostics FOR THIS FILE',
      'the session-wide push counter above says nothing about this ' +
      'buffer - no source delivered a diagnostic for it')
  else
    Row('  last resolution',
      Format('%d diag, %d fixable', [DiagSeen, DiagHandled]),
      'codes: ' + DiagCodes);

  // A diagnostic we HAVE a provider for that still produced nothing: the
  // provider says why (no declaring unit, stale, already reachable, ...).
  S := LiveDeclineNote;
  if S = '' then
    Row('  declined', '-', '')
  else
    Row('  declined', 'see details', S);

  // Which window (if any) currently keeps the fix hint hidden.
  S := LiveHintBlocker;
  if S = '' then
    Row('  hint', 'free', 'nothing covers the caret area')
  else
    Row('  hint', 'yielding to ' + S,
      'the hint stays hidden while another popup sits at the caret');

  LiveSourceStats(SrcStructure, SrcLsp);
  if SrcLsp = '' then
    Row('  from LSP pass', 'never completed', '')
  else
    Row('  from LSP pass', SrcLsp, 'last analysis of our own LSP session');
  if SrcStructure = '' then
    Row('  from Structure', 'never resolved', '')
  else
    Row('  from Structure', SrcStructure, 'last payload of the Structure view');

  // ---- Structure view (second diagnostics source) ------------------------
  StructureSourceStats(StrInstalled, StrFires, StrNodes, StrDiags, StrReason);
  if not StrInstalled then
    Row('Structure source', 'NOT installed',
      'IOTAStructureView notifier could not be registered')
  else if StrFires = 0 then
    Row('Structure source', 'installed, never fired',
      'the IDE has not reported a structure change yet')
  else if StrReason <> '' then
    Row('Structure source', Format('%d fire(s), last: nothing', [StrFires]),
      StrReason)
  else
    Row('Structure source', Format('%d fire(s), last: %d diag', [StrFires, StrDiags]),
      Format('%d node(s) walked', [StrNodes]));

  // ---- compiler output (third diagnostics source) ------------------------
  // CompilerMessageCount walks the IDE's INTERNAL message model through
  // raw exported method pointers - that must not run once a second. The
  // IDE-Logger caught the price: a first-chance ACCESS_VIOLATION every
  // few seconds from exactly this path (TTimer -> DoTick -> Collect ->
  // WalkMessageLines). It is a diagnostic number, so a 10 s refresh is
  // plenty; the compile path (FeedCompilerMessages) reads it for real.
  S := MessagesReaderProblem;
  if S <> '' then
    Row('Compiler messages', 'unavailable', S)
  else
  begin
    if (FMsgCountTick = 0) or (FTicks - FMsgCountTick >= 10) then
    begin
      FMsgCountTick := FTicks;
      try
        FMsgCount := CompilerMessageCount;
      except
        FMsgCount := 0;
      end;
    end;
    Row('Compiler messages', Format('%d line(s) readable', [FMsgCount]),
      'read from the Messages window after each compile');
  end;

  // ---- live blame ---------------------------------------------------------
  if BlameEnabled then
    Row('Live blame', BlameGutterStatus,
      'viewer: ' + BlameViewerName)
  else
    Row('Live blame', 'off',
      'switch it on via the Refactoring Light menu (runs "git blame")');

  // ---- Messages-window read probe ----------------------------------------
  S := TPath.Combine(TPath.GetTempPath, 'RefactoringLight-messages.log');
  // Touch the disk at most every 5th tick, and never again once found.
  if not FProbeSeen and (FTicks mod 5 = 0) then
    try
      FProbeSeen := TFile.Exists(S);
    except
    end;
  if FProbeSeen then
    Row('Messages probe', 'log written', S)
  else
    Row('Messages probe', 'no log yet', 'compile once to produce ' + S);

  // ---- menu integration ---------------------------------------------------
  if ContextMenuInstance <> nil then
    Row('Menu integration', ContextMenuInstance.MenuStatus,
      'editor popup path + IDE Refactor menu')
  else
    Row('Menu integration', 'not installed', '');

  // ---- active buffer ------------------------------------------------------
  S := '(none)';
  Detail := '';
  if Editor <> nil then
  begin
    // Cheap getter on purpose - GetCurrentContext would move the caret.
    Detail := Editor.GetActiveFileName;
    if Detail <> '' then S := ExtractFileName(Detail) else S := '(none)';
  end;
  Row('Active buffer', S, Detail);
end;

procedure TStatusFrame.DoTick(Sender: TObject);
begin
  // GDI objects this call leaves behind are booked per subsystem -
  // the status window shows the balance (Expert.ResourceMonitor).
  var GdiG := GdiGuard(gsStatusTick);
  // Window/state refresh only from a plain WM_TIMER tick (deadlock rule).
  if not Visible then Exit;
  Inc(FC.FTicks);
  try
    // a theme switch happens while we are open; re-applying only assigns what
    // really differs, so in the steady state this touches no window at all
    if FC.FTicks mod 10 = 0 then ApplyThemeToControls(Self);
    FC.Collect;
    Apply;   // writes only what changed - no flicker, keeps the selection
  except
    // a status display must never disturb the IDE
  end;
end;

var
  GSnapshot: TStatusCollector = nil;

function StatusSnapshot: TArray<TStatusRow>;
begin
  // Its own collector: the window's one ticks with the window, this one
  // with the requests - each keeps its own "every Nth round" rhythm.
  if GSnapshot = nil then GSnapshot := TStatusCollector.Create;
  Inc(GSnapshot.FTicks);
  GSnapshot.Collect;
  Result := Copy(GSnapshot.FRows, 0, GSnapshot.FRowCount);
end;

{ TMcpToolsFrame }

const
  // column indexes of the tools list (0 = caption)
  tcHandled = 0; tcCalls = 1; tcErrors = 2; tcLast = 3; tcLastMs = 4;
  tcAvgMax = 5; tcResult = 6; tcDescription = 7;

constructor TMcpToolsFrame.Create(AOwner: TComponent);

  procedure Col(const ACaption: string; AWidth: Integer; ARight: Boolean = False);
  begin
    var C := FList.Columns.Add;
    C.Caption := ACaption;
    C.Width := AWidth;
    if ARight then C.Alignment := taRightJustify;
  end;

begin
  inherited;
  Name := '';
  NoteDockHost(Self, GToolsForm, True);   // audit #37, M30
  FHeader := TLabel.Create(Self);
  FHeader.Parent := Self;
  FHeader.Align := alTop;
  FHeader.AutoSize := False;   // see CLAUDE.md: AutoSize + alTop misjudges
  FHeader.Height := 22;
  FHeader.Layout := tlCenter;
  FHeader.Transparent := True;
  FHeader.AlignWithMargins := True;
  FHeader.Margins.SetBounds(6, 2, 6, 0);

  FDetail := TMemo.Create(Self);
  FDetail.Parent := Self;
  FDetail.Align := alBottom;
  FDetail.Height := 90;
  FDetail.ReadOnly := True;
  FDetail.ScrollBars := ssVertical;
  FDetail.WordWrap := True;

  var Split := TSplitter.Create(Self);
  Split.Parent := Self;
  Split.Align := alBottom;
  Split.Top := FDetail.Top - 1;   // above the memo, not below it

  FList := TThemedListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.GridLines := True;
  FList.HideSelection := False;
  FList.OnSelectItem := DoSelect;
  Col('Tool', 150);
  Col('Handled by', 70);
  Col('Calls', 50, True);
  Col('Errors', 50, True);
  Col('Last call', 70);
  Col('Last ms', 60, True);
  Col('Avg / max ms', 90, True);
  Col('Last result', 200);
  Col('Description', 500);
  EnableListViewSorting(FList);

  LoadDefs;
  FTimer := TTimer.Create(Self);
  FTimer.Interval := 1000;
  FTimer.OnTimer := DoTick;
  FTimer.Enabled := True;
  ApplyThemeToControls(Self);   // the frame's children are ours to theme
  RefreshRows;
end;

destructor TMcpToolsFrame.Destroy;
begin
  NoteDockHost(Self, GToolsForm, False);   // audit #37, M30
  inherited;
end;

procedure TMcpToolsFrame.LoadDefs;
var
  Arr: TJSONArray;
begin
  // The bridge's list is compiled into this plugin as well (Mcp.Protocol) -
  // the same definitions the IDE serves to the bridge.
  Arr := McpToolDefinitions;
  try
    SetLength(FDefs, Arr.Count);
    for var I := 0 to Arr.Count - 1 do
    begin
      var O := Arr.Items[I] as TJSONObject;
      FDefs[I].Name := O.GetValue<string>('name', '');
      FDefs[I].Description := O.GetValue<string>('description', '');
      FDefs[I].Bridge := not McpHandlesTool(FDefs[I].Name);
    end;
  finally
    Arr.Free;
  end;
end;

function TMcpToolsFrame.StatOf(const AName: string;
  const AStats: TArray<TMcpToolStat>; out AStat: TMcpToolStat): Boolean;
begin
  for var S in AStats do
    if S.Name = AName then
    begin
      AStat := S;
      Exit(True);
    end;
  AStat := Default(TMcpToolStat);
  Result := False;
end;

procedure TMcpToolsFrame.DoTick(Sender: TObject);
begin
  // plain WM_TIMER tick - state reads only, cells written when changed
  if not Visible then Exit;
  Inc(FThemeTick);
  try
    if FThemeTick mod 10 = 0 then ApplyThemeToControls(Self);
    RefreshRows;
  except
    // a status display must never disturb the IDE
  end;
end;

procedure TMcpToolsFrame.DoSelect(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  ShowDetail;
end;

procedure TMcpToolsFrame.RefreshRows;

  procedure SetCell(AItem: TListItem; ACol: Integer; const AText: string);
  begin
    while AItem.SubItems.Count <= ACol do AItem.SubItems.Add('');
    if AItem.SubItems[ACol] <> AText then AItem.SubItems[ACol] := AText;
  end;

var
  Stats: TArray<TMcpToolStat>;
  St: TMcpToolStat;
  Detail, Conn: string;
  TotalCalls, TotalErrors: Integer;
begin
  Stats := McpToolStats;
  TotalCalls := 0;
  TotalErrors := 0;
  for var S in Stats do
  begin
    Inc(TotalCalls, S.Calls);
    Inc(TotalErrors, S.Errors);
  end;
  Conn := McpConnectionStatus(Detail);
  var H := Format('%d tools  |  %d call(s), %d error(s) this session  |  ' +
    'Claude Code: %s', [Length(FDefs), TotalCalls, TotalErrors, Conn]);
  if FHeader.Caption <> H then FHeader.Caption := H;

  if FList.Items.Count <> Length(FDefs) then
  begin
    FList.Items.BeginUpdate;
    try
      FList.Items.Clear;
      for var I := 0 to High(FDefs) do
      begin
        var It := FList.Items.Add;
        It.Caption := FDefs[I].Name;
        // rows map to FDefs through Data, never Index (sortable list)
        It.Data := Pointer(NativeInt(I));
      end;
    finally
      FList.Items.EndUpdate;
    end;
  end;

  for var K := 0 to FList.Items.Count - 1 do
  begin
    var It := FList.Items[K];
    var D := FDefs[NativeInt(It.Data)];
    if D.Bridge then SetCell(It, tcHandled, 'bridge')
    else SetCell(It, tcHandled, 'IDE');
    if StatOf(D.Name, Stats, St) then
    begin
      SetCell(It, tcCalls, IntToStr(St.Calls));
      SetCell(It, tcErrors, IfThen(St.Errors > 0, IntToStr(St.Errors), ''));
      SetCell(It, tcLast, FormatDateTime('hh:nn:ss', St.LastTime));
      if St.Running > 0 then
      begin
        SetCell(It, tcLastMs, 'running');
        SetCell(It, tcResult, 'running...');
      end
      else
      begin
        SetCell(It, tcLastMs, IntToStr(St.LastMs));
        if St.LastOk then SetCell(It, tcResult, 'ok')
        else SetCell(It, tcResult, 'ERROR: ' + St.LastError);
      end;
      var Done := St.Calls - St.Running;
      if Done > 0 then
        SetCell(It, tcAvgMax, Format('%d / %d', [St.TotalMs div Done, St.MaxMs]))
      else
        SetCell(It, tcAvgMax, '');
    end
    else
    begin
      SetCell(It, tcCalls, IfThen(D.Bridge, '-', '0'));
      SetCell(It, tcErrors, '');
      SetCell(It, tcLast, '');
      SetCell(It, tcLastMs, '');
      SetCell(It, tcAvgMax, '');
      SetCell(It, tcResult, IfThen(D.Bridge,
        'answered by the bridge, not counted here', ''));
    end;
    SetCell(It, tcDescription, D.Description);
  end;
  ShowDetail;
end;

procedure TMcpToolsFrame.ShowDetail;
var
  St: TMcpToolStat;
  Text: string;
begin
  if FList.Selected = nil then
  begin
    Text := 'Select a tool to see its full description and its last error.' +
      sLineBreak + 'The tools are used by Claude Code through the MCP bridge ' +
      '(RefactoringLightMcp.exe); "IDE" tools run in this IDE, "bridge" tools ' +
      'in the bridge itself (choosing the IDE).';
  end
  else
  begin
    var D := FDefs[NativeInt(FList.Selected.Data)];
    Text := D.Name + sLineBreak + D.Description;
    if StatOf(D.Name, McpToolStats, St) then
    begin
      Text := Text + sLineBreak + sLineBreak + Format('%d call(s), %d error(s), ' +
        'last at %s', [St.Calls, St.Errors, FormatDateTime('hh:nn:ss', St.LastTime)]);
      if St.LastError <> '' then
        Text := Text + sLineBreak + 'Last error: ' + St.LastError;
    end;
  end;
  // only when it changed - rewriting would reset the user's scroll/selection
  if Text <> FDetailFor then
  begin
    FDetailFor := Text;
    FDetail.Text := Text;
  end;
end;

{ TStatusDockable }

constructor TStatusDockable.Create(const ACaption, AIdent: string;
  AFrameClass: TCustomFrameClass);
begin
  inherited Create;
  FCaption := ACaption;
  FIdent := AIdent;
  FFrameClass := AFrameClass;
end;

function TStatusDockable.GetCaption: string;
begin
  Result := FCaption;
end;

function TStatusDockable.GetIdentifier: string;
begin
  // Section name in the desktop state file - do not translate or change.
  Result := FIdent;
end;

function TStatusDockable.GetFrameClass: TCustomFrameClass;
begin
  Result := FFrameClass;
end;

procedure TStatusDockable.FrameCreated(AFrame: TCustomFrame);
var
  Svc: IOTAIDEThemingServices;
begin
  // The IDE themes its own form chrome; the FRAME is ours, so hand it to
  // the theming service directly (EnableThemes takes a TCustomForm).
  try
    if (AFrame <> nil)
      and Supports(BorlandIDEServices, IOTAIDEThemingServices, Svc)
      and Svc.IDEThemingEnabled then
      Svc.ApplyTheme(AFrame);
  except
  end;
end;

function TStatusDockable.GetMenuActionList: TCustomActionList;
begin
  Result := nil;
end;

function TStatusDockable.GetMenuImageList: TCustomImageList;
begin
  Result := nil;
end;

procedure TStatusDockable.CustomizePopupMenu(PopupMenu: TPopupMenu);
begin
  // the default menu (Stay on top / Dockable) is enough
end;

function TStatusDockable.GetToolBarActionList: TCustomActionList;
begin
  Result := nil;
end;

function TStatusDockable.GetToolBarImageList: TCustomImageList;
begin
  Result := nil;
end;

procedure TStatusDockable.CustomizeToolBar(ToolBar: TToolBar);
begin
  // no toolbar
end;

procedure TStatusDockable.SaveWindowState(Desktop: TCustomIniFile;
  const Section: string; IsProject: Boolean);
begin
  // Nothing of our own: this window has no state besides its geometry, and
  // the rows rebuild themselves from the IDE on the next tick.
  // WHAT THE IDE WRITES, measured in the user's own project desktop rather
  // than assumed: our [RefactoringLightStatus] section is key for key the
  // same as the IDE's own [MessageView] and [StructureView] (Docked=1, the
  // geometry, Dockable, StayOnTop), and WHERE we were docked is recorded with
  // the dock SITE ("[DockSite5] TabDockClients=MessageView,
  // RefactoringLightStatus"). So the IDE persists our docking exactly as it
  // persists its own - but ONLY a saved desktop does: a package reload
  // neither saves nor loads one, which is why the unload hands the host
  // section to the next instance instead of relying on this.
end;

procedure TStatusDockable.LoadWindowState(Desktop: TCustomIniFile;
  const Section: string);
begin
end;

function TStatusDockable.GetEditState: TEditState;
begin
  Result := [];
end;

function TStatusDockable.EditAction(Action: TEditAction): Boolean;
begin
  Result := False;
end;

{ registration }

const
  /// <summary>How deep a dock chain is walked. MEASURED: a window in a tab
  ///  host in the main window is two links, and the IDE's own layout nests no
  ///  further - this is the runaway guard, not a limit anybody should
  ///  meet.</summary>
  MaxDockChainLinks = 8;
  /// <summary>How many clients of one host are recorded - a runaway guard in
  ///  the spirit of MaxDockChainLinks, not a limit: the biggest tab group in
  ///  the desktop files on this machine holds six.</summary>
  MaxDockClients = 16;

function ScreenBoundsOf(AForm: TCustomForm): TRect;
var
  TopLeft: TPoint;
begin
  // BoundsRect is relative to the PARENT, and docking REPARENTS the form into
  // its dock site - which is why the raw rectangle of a docked window reads
  // as the top left corner of the screen. Through the parent's client origin
  // it is the rectangle the window really occupies, in both states.
  //
  // AND IT MAY NOT RAISE: the one caller that matters runs in the package's
  // FINALIZATION (Expert.Registration asks for this state first, while the
  // window still exists), and an exception there costs the rest of the
  // teardown. BoundsRect on its own cannot raise, which is why the old
  // version needed no guard; ClientToScreen goes through Parent.Handle, and
  // asking a control for a handle it has not got CREATES one - during a
  // teardown that is the last thing to do. Hence HandleAllocated, and a
  // net around it: a window whose rectangle we could not convert keeps the
  // parent-relative one, and the restore refuses to place it anyway
  // (ShowStatusWindowAt wants a non-empty rect).
  Result := Default(TRect);
  if AForm = nil then Exit;
  try
    Result := AForm.BoundsRect;
    if (AForm.Parent <> nil) and AForm.Parent.HandleAllocated then
    begin
      TopLeft := AForm.Parent.ClientToScreen(Result.TopLeft);
      Result := TRect.Create(TopLeft, Result.Width, Result.Height);
    end;
  except
  end;
end;

function ReadDockTarget(AForm: TCustomForm): TDockPlace;
var
  Walk, HostForm: TCustomForm;
  Host: TWinControl;
  Link: TDockLink;
begin
  // THE WHOLE CHAIN, from the host that holds the window outwards to the form
  // that is not itself docked (the IDE's main window). Each link is a dock
  // SITE control plus the form that owns it.
  //
  // TWO MEASUREMENTS SHAPED THIS, both from the live IDE:
  //  * 2026-10-27: GetParentForm defaults to TopForm = True and walks to the
  //    OUTERMOST form, so a window in a tab host answered "Main Window" - the
  //    tab host is itself docked in the main window. False is the immediate
  //    parent form, i.e. the form that really owns the site. That fix stands;
  //    this walk simply takes the step on purpose instead of by accident.
  //  * 2026-10-28: the host was a TTabDockHostForm named "TabDockHostForm",
  //    created when the user DRAGGED the window to dock it, and its
  //    DeskSection is EMPTY - so a section cannot be the key. The component
  //    NAME is there either way (ide_introspect lists it), and measured on
  //    the live IDE every one of the 42 named forms has a distinct name, so
  //    it is what finds the same live form again after our package was loaded
  //    a second time.
  // AND THE CHAIN RATHER THAN ONE HOST, because whether the host survives our
  // unload depends on how the user made it. Measured in designide370.bpl:
  // TTabDockHostForm.DoRemoveDockClient posts $506 to itself as soon as at
  // most ONE client is left, $506 is TBaseDockHostForm.UMFinalUndock in the
  // class's own message table, and that hands the dock place to the remaining
  // client and calls TCustomForm.Release - but one flag byte gates it, written
  // only by TDesktopForm.LoadDockStream, so a host the DESKTOP created is
  // exempt and a DRAG-created one dissolves. Live: DockSite0..4 are still
  // there and the saved desktop's DockSite5 is not. The chain covers both
  // without having to know which: the host while it is there, and otherwise
  // the site it was itself docked in - where the surviving client goes too.
  Result := Default(TDockPlace);
  if AForm = nil then Exit;
  try
    Walk := AForm;
    // The step counter bounds the walk, not the number of links: a link that
    // cannot be named is SKIPPED and the walk goes on. Screen.Forms really
    // does hold nameless forms (measured: index 7 is a TForm with Name = ''),
    // and breaking there would throw away the link one step further out -
    // which is the one the whole chain design rests on, because it is the one
    // that survives our unload.
    for var Step := 1 to MaxDockChainLinks do
    begin
      Host := Walk.HostDockSite;
      if Host = nil then Break;          // the main window: nothing above it
      HostForm := GetParentForm(Host, False);
      if (HostForm = nil) or (HostForm = Walk) then Break;
      Link := Default(TDockLink);
      Link.SiteName := Host.Name;
      Link.FormName := HostForm.Name;
      Link.FormClass := HostForm.ClassName;
      if HostForm is TDesktopForm then
        Link.Section := TDesktopForm(HostForm).DeskSection;
      // AND WHETHER THIS HOST'S NAME WILL STILL MEAN IT. A host with no
      // DeskSection was created at run time, and its name is the class
      // default every TTabDockHostForm carries - so it is an identity only
      // while the host itself is alive. The measured dissolve rule
      // (UMFinalUndock fires when at most ONE client is left) says when that
      // ends: at most two clients now, ours among them. DockClientCount is
      // asked of the dock SITE, which is the control the clients really
      // registered with, whichever of the two it is; a count that does not
      // even include our own window is not something to trust either, so 0
      // counts as transient too.
      Link.Transient := (Link.Section = '') and (Host.DockClientCount <= 2);
      // AND WHO ELSE IS IN THIS HOST - the tab SIBLINGS, which is the half a
      // place cannot carry (user, 2026-10-28: the window came back docked
      // "rechts daneben" instead of as a tab beside the Messages view).
      // DockClientCount / DockClients are the public, VCL-maintained pair and
      // are asked of the dock SITE - the control the clients really
      // registered with, whichever of the two HostDockSite is - and the list
      // is in DOCK ORDER, ours among them. The cap is a runaway guard in the
      // spirit of MaxDockChainLinks: the biggest tab group in the desktop
      // files on this machine holds six clients.
      for var CI := 0 to Host.DockClientCount - 1 do
      begin
        if Length(Link.Clients) >= MaxDockClients then Break;
        var C := Host.DockClients[CI];
        if not (C is TCustomForm) then Continue;   // a client need not be one
        var Ref := Default(TDockClientRef);
        Ref.Ordinal := CI;
        Ref.IsSelf := (C = AForm);
        Ref.FormName := TCustomForm(C).Name;
        Ref.FormClass := C.ClassName;
        if C is TDesktopForm then
          Ref.Section := TDesktopForm(C).DeskSection;
        Link.Clients := Link.Clients + [Ref];
      end;
      if Link.Known then
        Result.Links := Result.Links + [Link];
      Walk := HostForm;
    end;
  except
    Result := Default(TDockPlace);
  end;
end;

function FindDockSite(AHost: TCustomForm; const AName: string): TWinControl;

  function Scan(AParent: TWinControl): TWinControl;
  begin
    Result := nil;
    for var I := 0 to AParent.ControlCount - 1 do
    begin
      if not (AParent.Controls[I] is TWinControl) then Continue;
      var C := TWinControl(AParent.Controls[I]);
      if SameText(C.Name, AName) then Exit(C);
      Result := Scan(C);
      if Result <> nil then Exit;
    end;
  end;

begin
  // The site control inside the host form. FindComponent answers for
  // anything the form owns, which is every panel the IDE streamed from its
  // own DFM; the control TREE is asked as well, because a panel created at
  // run time can be owned by something else entirely.
  Result := nil;
  if (AHost = nil) or (AName = '') then Exit;
  try
    if AHost.FindComponent(AName) is TWinControl then
      Exit(TWinControl(AHost.FindComponent(AName)));
    Result := Scan(AHost);
  except
    Result := nil;
  end;
end;

function FindDockHostForm(const ALink: TDockLink): TCustomForm;
begin
  // A dock site belongs to the IDE, so it is still on screen after our
  // package was unloaded and loaded again - which is the whole reason this
  // works at all.
  //
  // BY NAME FIRST, unless that name is a class default the IDE will hand
  // out again (Transient): a run-time host has no DeskSection (measured
  // 2026-10-28), so the section cannot be the key, and its own name is only
  // an identity while the host lives. The
  // section is still asked for afterwards, because that is all an older
  // marker and the saved desktop file carry - and for a host the desktop
  // created the two strings are the same anyway ("DockSite5").
  Result := nil;
  try
    // A TRANSIENT LINK IS NOT LOOKED UP BY NAME, because that name is a
    // class default which the next form the IDE streams will take back (see
    // TDockLink.Transient). Docking into whatever carries it now would put
    // the status window into a stranger's tab group AND report a successful
    // restore - the class test cannot catch it, since every run-time host is
    // a TTabDockHostForm. Its section is empty by definition, so the lookup
    // answers nothing and the restore takes the next link of the chain.
    if ALink.FormName <> '' then
      if ALink.Transient then
        Exit(nil)
      else
      for var I := 0 to Screen.FormCount - 1 do
      begin
        var F := Screen.Forms[I];
        if not SameText(F.Name, ALink.FormName) then Continue;
        // A name the IDE handed to a form of ANOTHER CLASS is not our
        // host. That is all this test can do, and it is not the case the
        // name key was introduced for: every run-time dock host is a
        // TTabDockHostForm, so a recycled name of the same class passes
        // here. Transient is what rejects that one, above.
        if (ALink.FormClass <> '') and
           not SameText(F.ClassName, ALink.FormClass) then Continue;
        Exit(F);
      end;
    if ALink.Section <> '' then
      for var I := 0 to Screen.FormCount - 1 do
      begin
        var F := Screen.Forms[I];
        if (F is TDesktopForm) and
           SameText(TDesktopForm(F).DeskSection, ALink.Section) then
          Exit(F);
      end;
    // AND A SITE NAME WITH NO OWNER, which is all the saved desktop file
    // gives: the form that OWNS a control of that name. EVIDENCE ONLY - with
    // two such forms (a second edit window has its own
    // DockBottomCenterPanel) the file does not say which, so this refuses
    // instead of picking one. The same rule everything else here follows:
    // a wrong host moves the window somewhere it never was, while no answer
    // leaves the next fallback its turn.
    if (ALink.FormName = '') and (ALink.Section = '') and
       (ALink.SiteName <> '') then
    begin
      var Owner: TCustomForm := nil;
      for var I := 0 to Screen.FormCount - 1 do
        if FindDockSite(Screen.Forms[I], ALink.SiteName) <> nil then
        begin
          if Owner <> nil then Exit(nil);   // ambiguous - say nothing
          Owner := Screen.Forms[I];
        end;
      Exit(Owner);
    end;
  except
    Result := nil;
  end;
end;

function DockedInto(AForm, AHost: TCustomForm): Boolean;
var
  Host: TWinControl;
begin
  // IS THIS FORM A CLIENT OF THAT HOST - asked about the two form objects,
  // never about their names. Screen.Forms really does hold nameless forms
  // (measured: index 7 is a TForm with Name = ''), so comparing names would
  // let two empty ones read as "already docked there" and verify a dock that
  // never happened.
  //
  // AND BY CONTAINMENT, not by the immediate host: dropping a window onto one
  // of the IDE's dock panels is what created the user's TabDockHostForm in
  // the first place, and ManualDock ends in that panel's own DockDrop - the
  // same path a drag takes - so the IDE may well wrap our window in a FRESH
  // tab host again. Then the immediate host form is that new host and not the
  // form we docked into, and comparing only one step would call a dock that
  // landed in exactly the right place a failure: the box would tell the user
  // to pick their desktop again, and the next fallback would move the window
  // away from where it had just correctly landed.
  Result := False;
  if (AForm = nil) or (AHost = nil) then Exit;
  var Walk: TCustomForm := AForm;
  for var Step := 1 to MaxDockChainLinks do
  begin
    Host := Walk.HostDockSite;
    if Host = nil then Exit(False);
    var HostForm := GetParentForm(Host, False);
    if HostForm = nil then Exit(False);
    if HostForm = AHost then Exit(True);
    if HostForm = Walk then Exit(False);
    Walk := HostForm;
  end;
end;

function FindSiblingForm(const ARef: TDockClientRef): TCustomForm;
begin
  // BY DESKSECTION, and by nothing else - the opposite of FindDockHostForm,
  // which may use a component name because a host that answers for one is
  // still that host. A sibling is looked up in order to be DOCKED INTO, so a
  // wrong answer would splice the status window into a stranger's tab group
  // AND report a restore; the component name is the DFM default (see
  // TDockLink.Transient), so it cannot carry that weight.
  Result := nil;
  if not ARef.Known then Exit;
  try
    for var I := 0 to Screen.FormCount - 1 do
    begin
      var F := Screen.Forms[I];
      if not (F is TDesktopForm) then Continue;
      if not SameText(TDesktopForm(F).DeskSection, ARef.Section) then Continue;
      if (ARef.FormClass <> '') and
         not SameText(F.ClassName, ARef.FormClass) then Continue;
      // ONLY A DOCKABLE FORM CAN BE SPLICED INTO A TAB HOST: the host takes
      // the sibling's own dock zone, which a form that is not docked through
      // the IDE's dockable-form machinery does not have. Found and then
      // refused is the safe direction - the relation is simply not restored.
      if not (F is TDockableForm) then Continue;
      Exit(F);
    end;
  except
    Result := nil;
  end;
end;

function MakeTabBeside(AForm: TCustomForm; ASibling: TCustomForm): Boolean;
var
  Site: TWinControl;
  Host: TTabDockHostForm;
begin
  // A TAB IS A RELATION, AND THIS IS THE ONLY WAY TO BUILD ONE: there is no
  // single API call for "make these two a tab group", and the IDE has no
  // helper for it either. TDockableForm.CreateDockParent is the one routine
  // in designide that creates a TTabDockHostForm at run time, and this IS its
  // body - measured byte for byte (research pass plus a skeptic over the same
  // bytes, 2026-10-28): construct the host with Application as its owner, set
  // its Parent to the dock site, make it visible, then call its own public
  // ManualDockClient TWICE - once for the client that is already docked with
  // Replace = True, once for the newcomer with False - and finish with Show +
  // FocusWindow.
  //
  // WHY THE SIBLING AND NOT WE ARE THE REPLACE PARTNER, and this is structural
  // rather than a preference: ManualDockClient(C, D, A, True) decodes to
  // Host.ReplaceDockedControl(Control := C, NewDockSite := Host, ...), and
  // Vcl.Controls.TControl.ReplaceDockedControl runs ManualDock on SELF - the
  // HOST - with the REPLACED control's old site. So it is the host that takes
  // the sibling's zone in the dock tree, which is exactly the inverse of the
  // UMFinalUndock that dissolved our old host and handed its zone to that same
  // sibling. With the arguments the other way round - our own window, whose
  // HostDockSite is nil after a reload - ManualDock(nil, ...) would FLOAT the
  // host instead, carrying every client it already holds. (CLAUDE.md recorded
  // that hazard with the wrong mechanism until this round: it said ManualDock
  // runs on the control being replaced. It does not; it runs on Self.)
  //
  // THE GUARD IS THE SAFETY RULE: ReplaceDockedControl only takes the branch
  // above when the replaced control's host has a dock MANAGER - otherwise it
  // ends in TWinControl.DockReplaceDockClient, which is a do-nothing that
  // returns False. Asking for it up front is what keeps a sibling that is not
  // really docked from floating the new host.
  //
  // NOT DONE, and named rather than forgotten: TEditorDockPanel
  // .ResizeSuspended, which CreateDockParent sets around the two calls. It
  // would avoid one re-layout and costs a dependency on IDEDockPanel; a
  // flicker is not a defect.
  Result := False;
  if (AForm = nil) or (ASibling = nil) or (AForm = ASibling) then Exit;
  Site := ASibling.HostDockSite;
  if (Site = nil) or not Site.UseDockManager or (Site.DockManager = nil) then
    Exit;
  try
    Host := TTabDockHostForm.Create(Application);
    Host.Parent := Site;
    // VISIBLE BEFORE DOCKED, the rule DockIntoLink states: TDockTree
    // .InsertControl opens with "if not Control.Visible then Exit" AFTER
    // CM_DOCKCLIENT has already re-parented the control.
    Host.Visible := True;
    if not AForm.Visible then
      try AForm.Show; except end;
    Host.ManualDockClient(ASibling, nil, alNone, True);
    Host.ManualDockClient(AForm, nil, alNone, False);
    Host.Show;
    // VERIFIED BY READING IT BACK, and here the test really is the RELATION:
    // both windows have to be clients of the same new host, which is what a
    // tab group IS. Containment in the old site would be true for the state
    // the user reported as a defect.
    Result := (AForm.HostDockSite <> nil) and not AForm.Floating and
      (GetParentForm(AForm.HostDockSite, False) = Host) and
      (ASibling.HostDockSite <> nil) and
      (GetParentForm(ASibling.HostDockSite, False) = Host);
    if Result then
      try
        Host.ShowDockClient(AForm);
        FocusWindow(Host);
      except
      end
    else
      // AND A HOST THAT DID NOT TAKE ITS CLIENTS MUST NOT STAY ON SCREEN.
      // Found by reading this routine rather than by a report: every exit
      // above leaves a constructed, VISIBLE TTabDockHostForm behind, and if
      // the dock did not happen it is an empty tab window the user never
      // asked for. Release and not Free, because this can run from inside
      // message handling - and only while it really is empty, so a half-done
      // dock that DID take the sibling is left for the IDE's own rule to
      // settle (DoRemoveDockClient -> UMFinalUndock hands the zone back).
      try
        if Host.DockClientCount = 0 then
          Host.Release;
      except
      end;
  except
    Result := False;
  end;
end;

function TabBesideRecorded(AForm: TCustomForm;
  const ASiblings: TArray<TDockClientRef>): Boolean;
begin
  // IN DOCK ORDER, and the first sibling that is still there wins: they were
  // all in one tab group, so any of them puts the window back into it.
  Result := False;
  for var Ref in ASiblings do
  begin
    var F := FindSiblingForm(Ref);
    if (F <> nil) and MakeTabBeside(AForm, F) then
      Exit(True);
  end;
end;

function DockIntoLink(AForm: TCustomForm; const ALink: TDockLink): Boolean;
var
  HostForm: TCustomForm;
  Site: TWinControl;
begin
  // ONE PRIMITIVE PER KIND OF HOST, and which one is not a matter of taste -
  // each of these was read either in the shipped VCL source or out of
  // designide370.bpl, twice (a research pass and then a skeptic over the same
  // bytes, who refuted part of the first answer):
  //  * A TAB HOST IS DOCKED INTO THROUGH ITS OWN FORM, with the IDE's
  //    wrapper: ManualDockClient(Client, nil, alClient, False) decodes to
  //    Client.ManualDock(HostForm, nil, alClient), so the dock goes through
  //    the host's CM_DOCKCLIENT and its DoAddDockClient - which is where the
  //    whole TTabDockInfo bookkeeping lives (the tab, the icon, the client
  //    list). Docking into its DockPanel instead looks more precise and is
  //    not: that panel has no dock manager, so per Vcl.Controls the client is
  //    never added to DockClients and no tab is ever made - an unmanaged
  //    child that our own check would still call a restore.
  //  * NOT THE FOUR-ARGUMENT TRUE, which is what this code passed until
  //    2.8.2. It selects Host.ReplaceDockedControl(Client, Host, nil,
  //    alClient), and Vcl.Controls.TControl.ReplaceDockedControl takes, for a
  //    client whose HostDockSite is NIL - every form after a package reload -
  //    the branch that calls "ManualDock(OldDockSite, nil, alTop)" with
  //    OldDockSite = nil ON THE CONTROL BEING REPLACED. ManualDock with a nil
  //    site FLOATS. So a "replace" against the user's tab host, which also
  //    held the Messages view, would have floated that out of his layout and
  //    then reported success. A restore must not move anything but its own
  //    window.
  //  * ANY OTHER HOST IS DOCKED INTO THROUGH ITS SITE CONTROL: a window in
  //    one of the IDE main form's panels has TAppBuilder as its host form,
  //    and there the panel is the dock site (every HostDockSite= line of
  //    every desktop file on this machine names one). TControl.ManualDock
  //    ends in NewDockSite.DockDrop, i.e. exactly the path a drag and drop
  //    takes.
  //  * ShowDockClient is the second step and never the first:
  //    GetDockClientInfo(Client), and - decoded, because our own comment
  //    claimed otherwise until 2.8.2 - a "test eax,eax / jz" before
  //    TTabDockInfo.ShowClient, so a client the host does not know simply
  //    returns. It brings an EXISTING client's tab to the front.
  Result := False;
  if (AForm = nil) or not ALink.Known then Exit;
  HostForm := FindDockHostForm(ALink);
  if HostForm = nil then Exit;
  try
    if not DockedInto(AForm, HostForm) then
    begin
      // VISIBLE FIRST, and that is not housekeeping: Vcl.Controls
      // TDockTree.InsertControl opens with "if not Control.Visible then Exit"
      // - after CMDockClient has already re-parented the control. So docking
      // a hidden form leaves it a child of the dock site with no zone in the
      // dock manager, and HostDockSite is set, so the check below would
      // report a restore that nobody can see. The callers show the window
      // before they get here (ShowStatusWindow does), which is exactly why
      // this is one line and not a long argument.
      if not AForm.Visible then
        try AForm.Show; except end;
      if HostForm is TTabDockHostForm then
        TTabDockHostForm(HostForm).ManualDockClient(AForm, nil, alClient, False)
      else
      begin
        Site := FindDockSite(HostForm, ALink.SiteName);
        // AND THE HOST FORM CAN BE ITS OWN SITE CONTROL
        // (TBaseDockHostForm.GetDockSiteControl is a bare ret, i.e. Self), so
        // a recorded site name that IS the form's name means the form.
        if (Site = nil) and SameText(ALink.SiteName, HostForm.Name) then
          Site := HostForm;
        if Site = nil then Exit;
        AForm.ManualDock(Site, nil, alClient);
      end;
    end;
    // VERIFIED BY READING IT BACK, against the form we actually found rather
    // than against the name we were given: the two differ exactly when the
    // link was matched by section, and what counts is where the window IS.
    Result := DockedInto(AForm, HostForm) and not AForm.Floating;
    // Now that it IS a client of this host, its tab may still be behind
    // another one. This is what that call is for.
    if Result and (HostForm is TTabDockHostForm) then
      try TTabDockHostForm(HostForm).ShowDockClient(AForm); except end;
  except
    Result := False;
  end;
end;

function RedockFormInto(AForm: TCustomForm;
  const APlace: TDockPlace): TReloadWindowOutcome;
var
  HadRelation: Boolean;
begin
  // THE DEEPEST LINK THAT STILL EXISTS. The host that held the window is the
  // first try, and for a host the desktop file created (DockSite0..N, kept
  // alive by its other clients) it is the answer. A tab host created by a
  // drag releases itself when our client leaves - measured, see
  // ReadDockTarget - and then the next link is the site THAT host was docked
  // in: the same place on screen.
  //
  // WHAT THAT SECOND LINK PRODUCES is the ORDINARY case, and this comment
  // said the opposite until an adversarial pass read the desktop files
  // properly. It is true that every HostDockSite= line sits in a DockSite<N>
  // section - but that key is not where a plain window's membership of a
  // panel is recorded: the PANEL'S OWN client list is. Measured in the files
  // on this machine: "LeftPanelClients=PropertyInspector,DockSite3,
  // MMXCodeExplorer" and "BottomPanelClients=MessageView,DockSite1", each of
  // those windows with its own section carrying Docked=1 and no
  // HostDockSite= and no TabDockClients= at all. MMXCodeExplorer's form is
  // OTADockForm_2 - a TOTADockForm, exactly our own kind. So a plain window
  // IS a normal direct client of a main-window panel, and docking into one
  // needs no apology. AND WHAT SUCH A DROP PRODUCES IS MEASURED now (the
  // user named the rule on 2026-10-28, and the binaries agree): a panel is
  // never wrapped in a fresh tab host. TPercentageDockTree.InsertControl -
  // the DefaultDockTreeClass designide installs - only adds a dictionary
  // entry and tail-calls the inherited one, with no tab set anywhere in it,
  // so Vcl.Controls' "if InsertAt in [alClient, alNone] then InsertAt :=
  // alRight" stands and the window lands BESIDE its old neighbour. A tab
  // needs a dockable FORM as the target (TDockableForm.CreateDockParent,
  // selected by a point inside GetTabDockRect's middle zone), or - when the
  // target is a tab host whose bar is visible - a point ON THAT BAR, which
  // then even picks the tab's position. So this link restores the PLACE and
  // not the tab relation; the restore verifies by containment, and the
  // status row says where the window ended up.
  // AND THE ANSWER IS THREE-VALUED, because the reported defect lives between
  // its last two values: the window came back INSIDE its old dock site and
  // the restore said "restored", while the user could see it was no longer a
  // TAB beside the Messages view. A Boolean cannot carry that, so the box
  // said nothing and he had to find it by looking.
  Result := rwoNotDone;
  HadRelation := False;
  for var Link in APlace.Links do
    if Length(Link.Siblings) > 0 then
      HadRelation := True;

  for var I := 0 to High(APlace.Links) do
  begin
    if DockIntoLink(AForm, APlace.Links[I]) then
    begin
      // THE INNERMOST LINK IS THE HOST THAT HELD THE WINDOW, so landing
      // there restores the place and every tab with it. Any link further out
      // is the place ALONE - and only worth saying when there was a relation
      // to lose: a marker older than 2.9.0 records no clients, and then
      // "place" is all we can check, which is exactly the old behaviour.
      if (I = 0) or not HadRelation then
        Exit(rwoRestored);
      Exit(rwoPlaceOnly);
    end;
    // THIS link's host is gone - a drag-created one dissolves when our client
    // leaves it. Its tab GROUP can still be rebuilt, and that beats falling
    // through to the next link, which can only give the place back.
    if (Length(APlace.Links[I].Siblings) > 0) and
       TabBesideRecorded(AForm, APlace.Links[I].Siblings) then
      Exit(rwoRestored);
  end;
end;

procedure ShowFormDockAware(AForm: TCustomForm; const AFallback: TDockPlace);
var
  WasOnScreen: Boolean;
begin
  // THE DEFECT THIS REPLACES: Show + BringToFront on a form that is DOCKED
  // (merely hidden, because the user closed it with the X) leaves it hidden
  // where it is. DeskUtil.FocusWindow is the IDE's own entry point for showing
  // a dockable window, and it takes exactly the TCustomForm that
  // INTAServices.CreateDockableForm hands us.
  //
  // MEASURED, by disassembling both routines of designide370.bpl rather than
  // trusting the names: FocusWindow IS ShowDockableForm plus SetWindowState,
  // BringToFront and the focus walk, so calling it alone is the superset.
  // ShowDockableForm reads HostDockSite, and when the form IS docked it makes
  // the form AND its host form visible (it returns the host) - which is
  // exactly the hidden-dock-client case. What it does NOT do is DOCK: for a
  // form with no host it degenerates to Visible := True, with no write to
  // HostDockSite and no ManualDock anywhere in its 216 bytes.
  // Hence the fallback below: whether the IDE leaves a closed dock client
  // docked or undocks it is not measurable without a live IDE, so the show is
  // followed by a CHECK, and a window that did not come back docked is handed
  // to the host we last saw it in. Right under either behaviour.
  if AForm = nil then Exit;
  WasOnScreen := (AForm.HandleAllocated and IsWindowVisible(AForm.Handle)) or
    AForm.Visible;
  try
    FocusWindow(AForm);
  except
    // Never lose the window over this: the old behaviour is still a window
    // on screen, just not in its dock site.
    try
      AForm.Show;
      AForm.BringToFront;
    except
    end;
  end;
  // AND ONLY FOR A WINDOW THAT WAS NOT ALREADY ON SCREEN. The fallback is
  // for the form the IDE has just handed back - freshly created after a
  // reload, or hidden because the user closed it with the X. A window that is
  // OPEN and free-standing is where the USER put it: re-docking it because
  // the menu entry was pressed would undo their own drag, and the remembered
  // place is never cleared when they float it. Read BEFORE the show, because
  // the show is what changes it - and through the handle, because a window
  // shown via the WinAPI desyncs the VCL Visible flag (the rule the
  // completion popup established).
  if (not WasOnScreen) and AFallback.Known and
     ((AForm.HostDockSite = nil) or AForm.Floating) then
    // The outcome is not read here: this path repairs a window the IDE handed
    // back undocked, and whatever it achieves is better than leaving it
    // floating. The reload REPORT asks the same function and does read it.
    RedockFormInto(AForm, AFallback);
end;

procedure RegisterDock(var ADock: INTACustomDockableForm; var ARegistered: Boolean;
  const ACaption, AIdent: string; AFrameClass: TCustomFrameClass);
var
  Svc: INTAServices;
begin
  if ARegistered then Exit;
  if not Supports(BorlandIDEServices, INTAServices, Svc) then Exit;
  ADock := TStatusDockable.Create(ACaption, AIdent, AFrameClass);
  try
    // Registering (rather than only creating) lets the IDE restore the
    // window from a saved desktop layout.
    Svc.RegisterDockableForm(ADock);
    ARegistered := True;
  except
    ADock := nil;
  end;
end;

procedure UnregisterDock(var ADock: INTACustomDockableForm; var AForm: TCustomForm;
  var ARegistered: Boolean);
var
  Svc: INTAServices;
  WasRegistered: Boolean;
begin
  // MANDATORY before the BPL unloads - the IDE would otherwise hold a
  // reference to code that is no longer mapped.
  //
  // THE ORDER IS THE POINT, and it is measured: the body of
  // INTAServices.UnregisterDockableForm is FreeAndNil on the IDE's OWN copy
  // of this very form (coreide370.bpl, @Otadockableform@
  // TOTADockableFormManager@UnregisterDockableForm, rva 00405C68: FreeAndNil
  // then IntfClear). Freeing it ourselves first was therefore a DOUBLE FREE
  // on every unload with the window open - which is exactly the
  // "Installieren" case this whole area is about, and a double free corrupts
  // the heap rather than raising, so the try/except below would not have
  // caught it. The IDE frees the form; we only let go of our pointer.
  WasRegistered := ARegistered;
  try
    if WasRegistered and Supports(BorlandIDEServices, INTAServices, Svc) then
      Svc.UnregisterDockableForm(ADock);
  except
  end;
  // Only when the registration never happened can the form still be ours -
  // then nobody else frees it. (ShowDock creates none without it, so this is
  // belt and braces.)
  if not WasRegistered then
    try
      AForm.Free;
    except
    end;
  AForm := nil;
  ARegistered := False;
  ADock := nil;
end;

procedure ShowDock(var ADock: INTACustomDockableForm; var AForm: TCustomForm;
  const AWhat: string; const AFallback: TDockPlace);
var
  Svc: INTAServices;
  Why: string;
begin
  Why := '';
  try
    if ADock = nil then
      RegisterStatusWindow;
    if ADock = nil then
      Why := 'the dockable form could not be registered (INTAServices missing?)'
    else
    begin
      if AForm = nil then
      begin
        if Supports(BorlandIDEServices, INTAServices, Svc) then
          AForm := Svc.CreateDockableForm(ADock)
        else
          Why := 'INTAServices not available';
      end;
      if (Why = '') and (AForm = nil) then
        Why := 'CreateDockableForm returned nil';
    end;
    if AForm <> nil then
      ShowFormDockAware(AForm, AFallback);
  except
    on E: Exception do
      Why := E.ClassName + ': ' + E.Message;
  end;
  // A menu entry that does nothing at all is the worst outcome - say why.
  if Why <> '' then
    ShowThemedMessage('The ' + AWhat + ' could not be opened.'#13#10#13#10 + Why);
end;

procedure RegisterStatusWindow;
begin
  RegisterDock(GDockable, GRegistered, 'Refactoring Light Status ' + PluginVersion,
    StatusWindowIdentifier, TStatusFrame);
  RegisterDock(GToolsDockable, GToolsRegistered, 'Refactoring Light MCP Tools',
    McpToolsWindowIdentifier, TMcpToolsFrame);
end;

procedure UnregisterStatusWindow;
begin
  FreeAndNil(GSnapshot);
  UnregisterDock(GToolsDockable, GToolsForm, GToolsRegistered);
  UnregisterDock(GDockable, GForm, GRegistered);
end;

function StatusDockWindowState: TRlDockState;
begin
  Result := Default(TRlDockState);
  if GForm = nil then Exit;
  Result.Open := True;
  // NOT Floating alone: the VCL makes it False for a window with no host at
  // all, so a free-standing window would read as docked.
  // THE HOST SITE DECIDES: a window with no host site is not docked. The
  // Floating term is belt and braces and not the reason - this is already
  // False without a host - and that is all this line claims now. (It used to
  // argue from TControl.GetFloating, i.e. "HostDockSite is
  // FloatingDockSiteClass"; since a form sets FloatingDockSiteClass to its
  // own ClassType, that reasoning cuts both ways and nobody measured which
  // way round it falls for a TOTADockForm.)
  Result.Docked := (GForm.HostDockSite <> nil) and not GForm.Floating;
  // THE RECTANGLE ALWAYS, THE SITE ONLY WHILE IT IS DOCKED - and the two
  // halves are asymmetric for different reasons.
  //  * The rectangle used to be recorded for a free-standing window only,
  //    because the raw BoundsRect of a docked one is its dock site's client
  //    coordinates and reads as the top left corner of the screen (2.8.1).
  //    Converted through the parent it is true in both states, and then it is
  //    worth having for a docked window too: when every dock site of the
  //    chain is gone, the place the window occupied is the best thing left -
  //    far better than the middle of the screen, which is what a user got
  //    on 2026-10-28.
  //  * The site is still read only while the window really is docked. The VCL
  //    gives a FLOATING control a host site of its own, and remembering that
  //    as "where it was docked" would hand the next reopen a target nobody
  //    ever docked into.
  Result.Bounds := ScreenBoundsOf(GForm);
  if Result.Docked then
    Result.Host := ReadDockTarget(GForm);
  // REMEMBERED WHILE IT IS TRUE. The status row asks this once a second, so
  // by the time the user closes the window we know which host held it - and
  // a host we have seen is the only thing that can put a window back that
  // the IDE handed back undocked.
  if Result.Host.Known then
    GLastStatusDock := Result.Host;
end;

function StatusDockWindowStateText: string;
var
  S: TRlDockState;
  Names: string;
begin
  S := StatusDockWindowState;
  if not S.Open then
    Exit('not open');
  if not S.Docked then
    Exit('open, free-standing');
  if not S.Host.Known then
    // STILL POSSIBLE, and it is what the reading of 2.8.1 answered for every
    // host the user had created by a drag: nothing of the chain could be
    // read. Then a reload really has nothing to put the window back into.
    Exit('open and docked, host unknown');
  Result := 'open and docked in ' + S.Host.Describe;
  // AND WITH WHOM, which is the half THREE rounds of this defect were about:
  // 2.8.1 and 2.8.2 were each diagnosed in one call because this row named
  // the host, and the round after each of them was about the window coming
  // back in the wrong RELATION - something the row could not say. A tab is a
  // relation to another client, so the row names the clients.
  Names := '';
  for var C in S.Host.Innermost.Siblings do
  begin
    if Names <> '' then Names := Names + ', ';
    if C.Section <> '' then
      Names := Names + C.Section
    else
      Names := Names + C.FormName;
  end;
  if Names <> '' then
    Result := Result + ', as a tab beside ' + Names;
end;

function ShowStatusWindowAt(const ABounds: TRect): Boolean;
begin
  ShowStatusWindow;
  Result := False;
  if GForm = nil then Exit;
  // A WINDOW THAT CAME BACK DOCKED IS LEFT ALONE, whoever asked: BoundsRect
  // is relative to the PARENT, and a docked form's parent is its dock site,
  // so writing a screen rectangle into it would resize the window inside
  // that site. It is open and the IDE placed it, which is more than this
  // call had to do - and it is what keeps the last resort after a failed
  // re-dock from making things worse.
  if (GForm.HostDockSite <> nil) and not GForm.Floating then Exit(True);
  if (ABounds.Width > 0) and (ABounds.Height > 0) then
    try
      GForm.BoundsRect := ABounds;
      Result := True;
    except
    end;
end;

function ShowStatusWindowDocked(const APlace: TDockPlace): TReloadWindowOutcome;
begin
  ShowStatusWindow;
  // AND A DOCK THAT WAS ASKED FOR AND DID NOT HAPPEN IS NOT A RESTORE, nor is
  // one that got the place but lost the tab - that distinction IS this
  // round's defect. Reporting more than happened would be the one thing this
  // project is against.
  if GForm = nil then
    Result := rwoNotDone
  else
    Result := RedockFormInto(GForm, APlace);
end;

procedure ShowStatusWindow;
begin
  ShowDock(GDockable, GForm, 'status window', GLastStatusDock);
  // AND THE SHOW ITSELF REMEMBERS WHERE IT LANDED, instead of leaving that
  // to the status ROW, which fills it only because it asks once a second -
  // an indirect coupling a later edit to that row would switch off in
  // silence. It asks the same function the row asks, whose documented side
  // effect that is, so the two can never disagree about what "docked, and
  // where" means; the answer is not needed here, only the note it leaves.
  StatusDockWindowState;
end;

procedure ShowMcpToolsWindow;
begin
  // No remembered dock site for this one: only the status window is restored
  // after a reload, and saying so explicitly is better than a default
  // argument that hides which window has that behaviour.
  ShowDock(GToolsDockable, GToolsForm, 'MCP tools window',
    Default(TDockPlace));
end;

end.
