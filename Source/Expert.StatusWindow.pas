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
  Expert.IdeThemes;   // TStatusLevel: a row's colour is part of its data

/// <summary>Registers the dockable form with the IDE. Call from Register.</summary>
procedure RegisterStatusWindow;
/// <summary>Unregisters and closes it. Call before the BPL unloads.</summary>
procedure UnregisterStatusWindow;
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
  Expert.ResourceMonitor, Expert.CompletionWizard,
  System.SysUtils, System.Classes, System.IniFiles, System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Graphics,
  Vcl.ActnList, Vcl.ImgList, Vcl.Menus,
  ToolsAPI, DesignIntf,   // DesignIntf: TEditState / TEditAction
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.LspManager, Expert.PluginSettings,
  Expert.AutoImport, Expert.ContextMenu, Expert.DialogHelper,   // IdeThemes: interface uses
  Expert.BlameGutter, Expert.BlameDialogs, Expert.VcsBlame,
  Expert.MessagesReader, Expert.StructureErrors, Expert.McpServer, Lsp.Client,
  Expert.Version, Expert.ListViewSort, Mcp.Protocol,
  System.JSON, System.StrUtils, Vcl.StdCtrls;

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
    procedure LoadDefs;
    procedure DoTick(Sender: TObject);
    procedure DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure RefreshRows;
    procedure ShowDetail;
    function StatOf(const AName: string; const AStats: TArray<TMcpToolStat>;
      out AStat: TMcpToolStat): Boolean;
  public
    constructor Create(AOwner: TComponent); override;
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
  GDockable: INTACustomDockableForm;
  GForm: TCustomForm;
  GRegistered: Boolean;
  GToolsDockable: INTACustomDockableForm;
  GToolsForm: TCustomForm;
  GToolsRegistered: Boolean;

{ TStatusFrame }

constructor TStatusFrame.Create(AOwner: TComponent);
begin
  inherited;
  Name := '';   // the IDE names the embedded instance
  FC := TStatusCollector.Create;

  FList := TListView.Create(Self);
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
  FC.Collect;
  Apply;
end;

destructor TStatusFrame.Destroy;
begin
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
begin
  ADefaultDraw := True;
  Sender.Canvas.Brush.Color := GetThemedColor(clWindow);
  if (AItem = nil) or (AItem.Index < 0) or (AItem.Index > High(FLevels)) then Exit;
  // column 1 is "Status" - the one word the user should be able to read
  // without reading anything else
  if ASubItem <> 1 then Exit;
  Lvl := FLevels[AItem.Index];
  if Lvl = slNeutral then Exit;
  if cdsSelected in AState then Exit;   // the selection owns its colours
  Sender.Canvas.Font.Color := StatusLevelColor(Lvl,
    GetThemedColor(clWindow), GetThemedColor(clWindowText));
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

  // ---- live quick-fix checker --------------------------------------------
  LiveStatusInfo(LiveFile, Analysing, Resolving, FromLsp, Fresh, FixCount);
  if LiveFile = '' then
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

  FList := TListView.Create(Self);
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
  RefreshRows;
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
  try
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
  // nothing of our own - the IDE persists size and dock position
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
begin
  // MANDATORY before the BPL unloads - the IDE would otherwise hold a
  // reference to code that is no longer mapped.
  try
    if AForm <> nil then
    begin
      AForm.Free;
      AForm := nil;
    end;
  except
    AForm := nil;
  end;
  try
    if ARegistered and Supports(BorlandIDEServices, INTAServices, Svc) then
      Svc.UnregisterDockableForm(ADock);
  except
  end;
  ARegistered := False;
  ADock := nil;
end;

procedure ShowDock(var ADock: INTACustomDockableForm; var AForm: TCustomForm;
  const AWhat: string);
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
    begin
      AForm.Show;
      AForm.BringToFront;
    end;
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
    'RefactoringLightStatus', TStatusFrame);
  RegisterDock(GToolsDockable, GToolsRegistered, 'Refactoring Light MCP Tools',
    'RefactoringLightMcpTools', TMcpToolsFrame);
end;

procedure UnregisterStatusWindow;
begin
  FreeAndNil(GSnapshot);
  UnregisterDock(GToolsDockable, GToolsForm, GToolsRegistered);
  UnregisterDock(GDockable, GForm, GRegistered);
end;

procedure ShowStatusWindow;
begin
  ShowDock(GDockable, GForm, 'status window');
end;

procedure ShowMcpToolsWindow;
begin
  ShowDock(GToolsDockable, GToolsForm, 'MCP tools window');
end;

end.
