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

/// <summary>Registers the dockable form with the IDE. Call from Register.</summary>
procedure RegisterStatusWindow;
/// <summary>Unregisters and closes it. Call before the BPL unloads.</summary>
procedure UnregisterStatusWindow;
/// <summary>Menu entry point: shows (and focuses) the status window.</summary>
procedure ShowStatusWindow;

implementation

uses
  System.SysUtils, System.Classes, System.IniFiles, System.IOUtils,
  Vcl.Forms, Vcl.Controls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.ActnList, Vcl.ImgList, Vcl.Menus,
  ToolsAPI, DesignIntf,   // DesignIntf: TEditState / TEditAction
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.LspManager,
  Expert.AutoImport, Expert.ContextMenu, Expert.IdeThemes, Expert.DialogHelper,
  Expert.MessagesReader, Expert.StructureErrors, Lsp.Client;

// TCustomFrame.Create does InitInheritedComponent(Self, TFrame) for every
// descendant and RAISES EResNotFound when the class has no DFM resource -
// so even a fully code-built frame needs one (Expert.StatusWindow.dfm is
// an empty frame; all controls are still created in the constructor).
{$R *.dfm}

type
  TStatusRow = record
    Caption, Value, Detail: string;
  end;

  // The frame the IDE embeds into its dockable form.
  TStatusFrame = class(TFrame)
  private
    FList: TListView;
    FTimer: TTimer;
    // Collected rows of the CURRENT tick. The row SET is fixed (same
    // count, same order, always) - only the cell texts change, so the
    // refresh can write single cells instead of rebuilding the list.
    // Rebuilding made the view flicker once a second and dropped the
    // user's selection every time.
    FRows: TArray<TStatusRow>;
    FRowCount: Integer;
    FProbeSeen: Boolean;   // messages log found once - stop stat()ing it
    FTicks: Integer;
    procedure DoTick(Sender: TObject);
    procedure Row(const ACaption, AValue, ADetail: string);
    procedure Collect;
    procedure Apply;
  public
    constructor Create(AOwner: TComponent); override;
  end;

  TStatusDockable = class(TInterfacedObject, INTACustomDockableForm)
  public
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

{ TStatusFrame }

constructor TStatusFrame.Create(AOwner: TComponent);
begin
  inherited;
  Name := '';   // the IDE names the embedded instance

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.GridLines := True;
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
  Collect;
  Apply;
end;

procedure TStatusFrame.Row(const ACaption, AValue, ADetail: string);
begin
  if FRowCount >= Length(FRows) then
    SetLength(FRows, FRowCount + 8);
  FRows[FRowCount].Caption := ACaption;
  FRows[FRowCount].Value := AValue;
  FRows[FRowCount].Detail := ADetail;
  Inc(FRowCount);
end;

// Writes the collected rows into the list view - and ONLY what actually
// changed. In the normal case (nothing moved since the last tick) not a
// single assignment happens, so the control never repaints and the
// selection survives.
procedure TStatusFrame.Apply;
var
  I: Integer;
  It: TListItem;
begin
  if FList.Items.Count <> FRowCount then
  begin
    // Structural change (should not happen - the row set is fixed).
    FList.Items.BeginUpdate;
    try
      FList.Items.Clear;
      for I := 0 to FRowCount - 1 do
      begin
        It := FList.Items.Add;
        It.Caption := FRows[I].Caption;
        It.SubItems.Add(FRows[I].Value);
        It.SubItems.Add(FRows[I].Detail);
      end;
    finally
      FList.Items.EndUpdate;
    end;
    Exit;
  end;

  for I := 0 to FRowCount - 1 do
  begin
    It := FList.Items[I];
    if It.Caption <> FRows[I].Caption then
      It.Caption := FRows[I].Caption;
    if It.SubItems.Count > 0 then
    begin
      if It.SubItems[0] <> FRows[I].Value then
        It.SubItems[0] := FRows[I].Value;
      if (It.SubItems.Count > 1) and (It.SubItems[1] <> FRows[I].Detail) then
        It.SubItems[1] := FRows[I].Detail;
    end;
  end;
end;

procedure TStatusFrame.Collect;
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
  if not TLspManager.Instance.IsAlive then
    Row('DelphiLSP session', 'not started',
      'starts on the first request (rename, completion, quick fixes)')
  else
  begin
    DiagCount := 0;
    DiagFiles := 0;
    Client := TLspManager.Instance.PeekClient;
    if Client <> nil then
      try
        DiagCount := Client.GetDiagnosticsCount;
        DiagFiles := Client.GetDiagnosticFileCount;
      except
      end;
    if DiagCount = 0 then
      Row('DelphiLSP session', 'running, no diagnostics',
        'our session has never pushed one - hints (H2443, ...) are ' +
        'unavailable; errors come from the Structure view / the compiler')
    else
      Row('DelphiLSP session', 'running',
        Format('%d push(es) received, diagnostics held for %d file(s) ' +
          '(ALL files, not just this one)',
          [DiagCount, DiagFiles]));
  end;
  if TLspManager.Instance.ProjectIndexed then S := 'yes' else S := 'no';
  Row('  project indexed', S, 'the LSP has seen this project once');

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
  S := MessagesReaderProblem;
  if S <> '' then
    Row('Compiler messages', 'unavailable', S)
  else
    Row('Compiler messages', Format('%d line(s) readable', [CompilerMessageCount]),
      'read from the Messages window after each compile');

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
  // Window/state refresh only from a plain WM_TIMER tick (deadlock rule).
  if not Visible then Exit;
  Inc(FTicks);
  try
    Collect;
    Apply;   // writes only what changed - no flicker, keeps the selection
  except
    // a status display must never disturb the IDE
  end;
end;

{ TStatusDockable }

function TStatusDockable.GetCaption: string;
begin
  Result := 'Refactoring Light Status';
end;

function TStatusDockable.GetIdentifier: string;
begin
  // Section name in the desktop state file - do not translate or change.
  Result := 'RefactoringLightStatus';
end;

function TStatusDockable.GetFrameClass: TCustomFrameClass;
begin
  Result := TStatusFrame;
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

procedure RegisterStatusWindow;
var
  Svc: INTAServices;
begin
  if GRegistered then Exit;
  if not Supports(BorlandIDEServices, INTAServices, Svc) then Exit;
  GDockable := TStatusDockable.Create;
  try
    // Registering (rather than only creating) lets the IDE restore the
    // window from a saved desktop layout.
    Svc.RegisterDockableForm(GDockable);
    GRegistered := True;
  except
    GDockable := nil;
  end;
end;

procedure UnregisterStatusWindow;
var
  Svc: INTAServices;
begin
  // MANDATORY before the BPL unloads - the IDE would otherwise hold a
  // reference to code that is no longer mapped.
  try
    if GForm <> nil then
    begin
      GForm.Free;
      GForm := nil;
    end;
  except
    GForm := nil;
  end;
  try
    if GRegistered and Supports(BorlandIDEServices, INTAServices, Svc) then
      Svc.UnregisterDockableForm(GDockable);
  except
  end;
  GRegistered := False;
  GDockable := nil;
end;

procedure ShowStatusWindow;
var
  Svc: INTAServices;
  Why: string;
begin
  Why := '';
  try
    if GDockable = nil then
      RegisterStatusWindow;
    if GDockable = nil then
      Why := 'the dockable form could not be registered (INTAServices missing?)'
    else
    begin
      if GForm = nil then
      begin
        if Supports(BorlandIDEServices, INTAServices, Svc) then
          GForm := Svc.CreateDockableForm(GDockable)
        else
          Why := 'INTAServices not available';
      end;
      if (Why = '') and (GForm = nil) then
        Why := 'CreateDockableForm returned nil';
    end;
    if GForm <> nil then
    begin
      GForm.Show;
      GForm.BringToFront;
    end;
  except
    on E: Exception do
      Why := E.ClassName + ': ' + E.Message;
  end;
  // A menu entry that does nothing at all is the worst outcome - say why.
  if Why <> '' then
    ShowThemedMessage('The status window could not be opened.'#13#10#13#10 + Why);
end;

end.
