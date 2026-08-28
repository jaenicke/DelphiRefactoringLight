(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ContextMenu;

{
  Hooks a "Refactoring Light" submenu into the Delphi IDE code editor's
  context menu (the popup that appears when right-clicking in the code
  editor).

  IMPORTANT - why we hook OnPopup instead of just appending items:

  Permanently keeping our TMenuItems inside EditorLocalMenu.Items
  triggers the IDE's "A component named X already exists" error when
  the popup is opened. The IDE's own popup-open logic walks the menu
  and re-registers component names; foreign items confuse it.

  CnPack solves this in CnMenuHook.pas with the same pattern we use
  here: hook the popup's OnPopup, and on every popup:
    1. remove our items so the menu looks pristine,
    2. run the original OnPopup (the IDE bookkeeping sees only its
       own items),
    3. add our items back so the user can click them.

  We also re-hook periodically because other IDE plugins may override
  OnPopup later and would otherwise displace us.

  The popup may not exist yet when Register runs, so the install is
  retried via TTimer until the menu is found or the retry budget is
  exhausted.
}

interface

uses
  System.Classes, System.Generics.Collections, Vcl.Menus, Vcl.ExtCtrls,
  Expert.Shortcuts;

type
  TContextMenuInstaller = class
  private
    FItems: TArray<TMenuItem>;
    FSeparator: TMenuItem;
    FPopupMenu: TPopupMenu;
    FOldOnPopup: TNotifyEvent;
    FHooked: Boolean;
    // Main-menu integration: a copy of our actions hung directly into the
    // IDE's top-level Refactor menu.
    FMainMenuOwner: TComponent;        // owns the main-menu item tree
    FMainMenuAdded: TArray<TMenuItem>; // items added to the Refactor menu (incl. separator)
    FMainDumped: Boolean;              // wrote the debug dump already
    // The Refactor menu is disabled by its action's OnUpdate; we hook that
    // update to keep the menu openable (our items stay reachable).
    FRefactorAction: TBasicAction;
    FRefOldUpdate: TNotifyEvent;
    FHidden: TArray<TMenuItem>;   // IDE's own Refactor items we hid (to restore)
    FItemReq: TDictionary<TMenuItem, Integer>;   // main-menu item -> context requirement
    FLastStateTick: Cardinal;   // throttle for UpdateMainItemStates
    FRetryTimer: TTimer;
    FRetryCount: Integer;
    FSyncTimer: TTimer;
    FInPopup: Boolean;
    procedure UpdateMainItemStates;
    procedure OnRename(Sender: TObject);
    procedure OnFindReferences(Sender: TObject);
    procedure OnFindImplementations(Sender: TObject);
    procedure OnExtractMethod(Sender: TObject);
    procedure OnCompletion(Sender: TObject);
    procedure OnSignatureCheck(Sender: TObject);
    procedure OnRemoveWithProjectWide(Sender: TObject);
    procedure OnRemoveWithCurrentUnit(Sender: TObject);
    procedure OnRemoveWithSelectedUnits(Sender: TObject);
    procedure OnRemoveWithAtCursor(Sender: TObject);
    procedure OnUnitRefs(Sender: TObject);
    procedure OnFindUnit(Sender: TObject);
    procedure OnMoveToUnit(Sender: TObject);
    procedure OnFindOriginalSymbol(Sender: TObject);
    procedure OnExtractInterface(Sender: TObject);
    procedure OnAddToExistingInterface(Sender: TObject);
    procedure OnDelegateInterface(Sender: TObject);
    procedure OnSemanticReplaceCurrent(Sender: TObject);
    procedure OnSemanticReplaceSelected(Sender: TObject);
    procedure OnSemanticReplaceProject(Sender: TObject);
    procedure OnSemanticReplaceEditRules(Sender: TObject);
    procedure OnCheckDfmEvents(Sender: TObject);
    procedure OnCheckInterfaceGuids(Sender: TObject);
    procedure OnCheckCircularRefs(Sender: TObject);
    procedure OnRetryTimer(Sender: TObject);
    procedure OnSyncTimer(Sender: TObject);
    procedure DoOnPopup(Sender: TObject);
    function IsOurHandler(const AHandler: TNotifyEvent): Boolean;
    function FindEditorPopupMenu: TPopupMenu;
    function MakeItem(AOwner: TComponent; const ACaption: string;
      AOnClick: TNotifyEvent; AKind: TShortcutKind; ATrackShortcut: Boolean): TMenuItem;
    function BuildMenuTree(AOwner: TComponent; ATrackShortcut: Boolean): TMenuItem;
    function FindRefactorMenu: TMenuItem;
    procedure DumpMainMenu(const AReason: string);
    procedure HookRefactorAction(AParent: TMenuItem);
    procedure DoRefactorActionUpdate(Sender: TObject);
    function IsOurMainItem(AItem: TMenuItem): Boolean;
    procedure HideExistingEntries(AParent: TMenuItem);
    procedure InstallIntoMainMenu;
    procedure RemoveFromMainMenu;
    procedure BuildItems;
    procedure AddOurItems;
    procedure RemoveOurItems;
    procedure HookPopup;
    procedure UnhookPopup;
    procedure TryInstall;
  public
    destructor Destroy; override;
    procedure Install;
    procedure Uninstall;
    /// <summary>Re-applies the shortcuts from Expert.Shortcuts to all
    ///  menu items. Call after the user changed the shortcut settings.</summary>
    procedure RefreshShortcuts;
  end;

var
  ContextMenuInstance: TContextMenuInstaller;

implementation

uses
  System.SysUtils, System.UITypes, System.IOUtils, System.StrUtils,
  System.Actions, Winapi.Windows, ToolsAPI,
  Vcl.Forms, Vcl.Controls, Vcl.ActnList,
  Expert.EditorHelperIntf,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod,
  Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard, Expert.WithRefactorWizard, Expert.UnitReferencesWizard,
  Expert.MoveToUnitWizard, Expert.ExtractInterfaceWizard,
  Expert.SemanticReplaceWizard, Expert.DfmEventCheckDialog,
  Expert.InterfaceGuidDialog, Expert.CircularRefsDialog,
  Expert.FindUnitDialog, Expert.FindOriginalSymbolWizard;

const
  /// <summary>Maximum retry attempts when the editor popup is not yet
  ///  available. With a 500 ms interval, 40 retries give up to 20 s.</summary>
  MaxRetries = 40;

  /// <summary>Period at which we verify our OnPopup hook is still in
  ///  place. Other plugins may install their own OnPopup later and
  ///  unintentionally displace ours.</summary>
  SyncIntervalMs = 2000;

  // Context requirement for a main-menu entry.
  REQ_PROJECT = 1;   // needs an open project
  REQ_EDITOR  = 2;   // needs an active source editor (cursor context)

{ TContextMenuInstaller }

destructor TContextMenuInstaller.Destroy;
begin
  Uninstall;
  FreeAndNil(FItemReq);
  inherited;
end;

function TContextMenuInstaller.FindEditorPopupMenu: TPopupMenu;
var
  I, J: Integer;
  Form: TCustomForm;
  Comp: TComponent;
begin
  Result := nil;
  for I := 0 to Screen.CustomFormCount - 1 do
  begin
    Form := Screen.CustomForms[I];
    for J := 0 to Form.ComponentCount - 1 do
    begin
      Comp := Form.Components[J];
      if (Comp is TPopupMenu) and
         SameText(Comp.Name, 'EditorLocalMenu') then
        Exit(TPopupMenu(Comp));
    end;
  end;
end;

function TContextMenuInstaller.MakeItem(AOwner: TComponent; const ACaption: string;
  AOnClick: TNotifyEvent; AKind: TShortcutKind; ATrackShortcut: Boolean): TMenuItem;
begin
  Result := TMenuItem.Create(AOwner);
  Result.Caption := ACaption;
  Result.OnClick := AOnClick;
  if ATrackShortcut then
  begin
    // Tag stores the shortcut kind so RefreshShortcuts can find this item.
    // Encoded as Ord+1 so 0 means "no shortcut tracked".
    Result.Tag := Ord(AKind) + 1;
    Result.ShortCut := TExpertsShortCut.Shortcuts[AKind];
  end;
end;

procedure TContextMenuInstaller.RefreshShortcuts;
var
  Item: TMenuItem;
  Kind: TShortcutKind;
begin
  for Item in FItems do
  begin
    if Item = nil then Continue;
    if (Item.Tag >= 1) and (Item.Tag <= Ord(High(TShortcutKind)) + 1) then
    begin
      Kind := TShortcutKind(Item.Tag - 1);
      Item.ShortCut := TExpertsShortCut.Shortcuts[Kind];
    end;
  end;
end;

// Builds the whole "Refactoring Light" submenu tree owned by AOwner and
// returns its root item. ATrackShortcut controls whether the leaf items
// carry the configurable shortcuts (True for the editor popup, False for
// the copy hung into the IDE's main Refactor menu).
function TContextMenuInstaller.BuildMenuTree(AOwner: TComponent;
  ATrackShortcut: Boolean): TMenuItem;

  // Records the item's context requirement (only for the main-menu copy;
  // the popup already only appears in an editor context).
  procedure Req(AItem: TMenuItem; AReq: Integer);
  begin
    if (not ATrackShortcut) and (FItemReq <> nil) then
      FItemReq.AddOrSetValue(AItem, AReq);
  end;

  function Leaf(AParent: TMenuItem; const ACaption: string; AOnClick: TNotifyEvent;
    AKind: TShortcutKind; AReq: Integer): TMenuItem;
  begin
    Result := MakeItem(AOwner, ACaption, AOnClick, AKind, ATrackShortcut);
    AParent.Add(Result);
    Req(Result, AReq);
  end;

  function Plain(AParent: TMenuItem; const ACaption: string; AOnClick: TNotifyEvent;
    AReq: Integer): TMenuItem;
  begin
    Result := TMenuItem.Create(AOwner);
    Result.Caption := ACaption;
    Result.OnClick := AOnClick;
    AParent.Add(Result);
    Req(Result, AReq);
  end;

  function Sub(AParent: TMenuItem; const ACaption: string): TMenuItem;
  begin
    Result := TMenuItem.Create(AOwner);
    Result.Caption := ACaption;
    AParent.Add(Result);
    Req(Result, REQ_PROJECT);   // a group is usable whenever a project is open
  end;

var
  Root, RemoveWithSub, IfaceSub, SemSub, ChecksSub: TMenuItem;
begin
  if (not ATrackShortcut) and (FItemReq = nil) then
    FItemReq := TDictionary<TMenuItem, Integer>.Create;

  Root := TMenuItem.Create(AOwner);
  Root.Caption := 'Refactoring Light';

  Leaf(Root, 'Rename...',                 OnRename,               skRename,     REQ_EDITOR);
  Leaf(Root, 'Find References',           OnFindReferences,       skFindRef,    REQ_EDITOR);
  Leaf(Root, 'Find Implementations',      OnFindImplementations,  skFindImp,    REQ_EDITOR);
  Leaf(Root, 'Find Original Symbol',      OnFindOriginalSymbol,   skFindOriginalSymbol, REQ_EDITOR);
  Leaf(Root, 'Extract Method',            OnExtractMethod,        skExtract,    REQ_EDITOR);
  Leaf(Root, 'Align method signature...', OnSignatureCheck,       skAlign,      REQ_EDITOR);
  Leaf(Root, 'Code Completion',           OnCompletion,           skCompletion, REQ_EDITOR);

  RemoveWithSub := Sub(Root, 'Remove with');
  Leaf(RemoveWithSub, 'At cursor only',     OnRemoveWithAtCursor,      skRemoveWith, REQ_EDITOR);
  Plain(RemoveWithSub, 'In current unit',    OnRemoveWithCurrentUnit,   REQ_EDITOR);
  Plain(RemoveWithSub, 'In selected units...', OnRemoveWithSelectedUnits, REQ_PROJECT);
  Plain(RemoveWithSub, 'In whole project...', OnRemoveWithProjectWide,   REQ_PROJECT);

  Leaf(Root, 'Move to unit...',           OnMoveToUnit,           skMoveToUnit, REQ_EDITOR);
  Leaf(Root, 'Find unit references...',   OnUnitRefs,             skUnitRefs,   REQ_EDITOR);
  Plain(Root, 'Find unit for identifier...', OnFindUnit,          REQ_PROJECT);

  IfaceSub := Sub(Root, 'Extract / extend interface');
  Plain(IfaceSub, 'Extract new interface from class...', OnExtractInterface,       REQ_EDITOR);
  Plain(IfaceSub, 'Add to existing interface...',        OnAddToExistingInterface, REQ_EDITOR);
  Plain(IfaceSub, 'Add IInterface support to class...',  OnDelegateInterface,      REQ_EDITOR);

  SemSub := Sub(Root, 'Semantic replace');
  Plain(SemSub, 'In current unit',    OnSemanticReplaceCurrent,   REQ_EDITOR);
  Plain(SemSub, 'In selected units...', OnSemanticReplaceSelected, REQ_PROJECT);
  Plain(SemSub, 'In whole project...', OnSemanticReplaceProject,   REQ_PROJECT);
  Plain(SemSub, 'Edit rules...',       OnSemanticReplaceEditRules, REQ_PROJECT);

  ChecksSub := Sub(Root, 'Project checks');
  Plain(ChecksSub, 'DFM event handlers...',        OnCheckDfmEvents,       REQ_PROJECT);
  Plain(ChecksSub, 'Interface GUIDs...',           OnCheckInterfaceGuids,  REQ_PROJECT);
  Plain(ChecksSub, 'Circular unit references...',  OnCheckCircularRefs,    REQ_PROJECT);

  Result := Root;
end;

procedure TContextMenuInstaller.BuildItems;
var
  Submenu: TMenuItem;
begin
  if Length(FItems) > 0 then Exit;

  Submenu := BuildMenuTree(FPopupMenu, True);

  FSeparator := TMenuItem.Create(FPopupMenu);
  FSeparator.Caption := '-';

  // FItems holds exactly the top-level items we add to FPopupMenu.Items
  // each time the popup opens (and remove again before the next open).
  FItems := [FSeparator, Submenu];
end;

// Strips the '&' accelerator markers and a trailing ellipsis from a menu
// caption and upper-cases it, for language-tolerant caption matching.
function NormCaption(const S: string): string;
begin
  Result := UpperCase(Trim(StringReplace(S, '&', '', [rfReplaceAll])));
  Result := Result.TrimRight(['.', #$2026, ' ']);
end;

function TContextMenuInstaller.FindRefactorMenu: TMenuItem;
const
  // Captions of the IDE's Refactor menu across the shipped IDE languages
  // (accelerators/ellipsis already stripped by NormCaption).
  RefactorCaptions: array[0..6] of string =
    ('REFACTOR', 'REFACTORING', 'REFAKTORIEREN', 'REFACTORISER',
     'RIFATTORIZZA', 'REFATORAR', 'REFACTORIZAR');
var
  NTA: INTAServices;
  Menu: TMainMenu;
  I, K: Integer;
  It, Cand: TMenuItem;
  Cap: string;
begin
  Result := nil;
  if not Supports(BorlandIDEServices, INTAServices, NTA) then Exit;
  Menu := NTA.MainMenu;
  if Menu = nil then Exit;

  // 1) By component Name - language-independent.
  for I := 0 to Menu.Items.Count - 1 do
  begin
    It := Menu.Items[I];
    if SameText(It.Name, 'RefactorMenu') or SameText(It.Name, 'RefactoringMenu') then
      Exit(It);
  end;

  // 2) By (normalized) caption against the known translations.
  for I := 0 to Menu.Items.Count - 1 do
  begin
    It := Menu.Items[I];
    Cap := NormCaption(It.Caption);
    for K := Low(RefactorCaptions) to High(RefactorCaptions) do
      if Cap = RefactorCaptions[K] then Exit(It);
  end;

  // 3) Fallback heuristic: the sole top-level item that is an empty
  //    placeholder - a real caption but no sub-items and no click handler.
  Cand := nil;
  for I := 0 to Menu.Items.Count - 1 do
  begin
    It := Menu.Items[I];
    if (It.Count = 0) and (It.Caption <> '') and (It.Caption <> '-')
       and not Assigned(It.OnClick) then
    begin
      if Cand = nil then Cand := It else begin Cand := nil; Break; end;
    end;
  end;
  Result := Cand;
end;

procedure TContextMenuInstaller.DumpMainMenu(const AReason: string);

  function ActInfo(It: TMenuItem): string;
  begin
    if It.Action <> nil then
      Result := Format('Action="%s"(En=%s)',
        [It.Action.Name, BoolToStr((It.Action as TContainedAction).Enabled, True)])
    else
      Result := 'Action=-';
  end;

var
  NTA: INTAServices;
  Menu: TMainMenu;
  SB: TStringBuilder;
  I: Integer;
  It, Found: TMenuItem;
  Path: string;
begin
  try
    SB := TStringBuilder.Create;
    try
      SB.AppendLine('--- Refactoring Light main-menu dump (' + AReason + ') ---');
      if not Supports(BorlandIDEServices, INTAServices, NTA) then
        SB.AppendLine('INTAServices NOT available')
      else
      begin
        Menu := NTA.MainMenu;
        if Menu = nil then
          SB.AppendLine('INTAServices.MainMenu is nil')
        else
        begin
          Found := FindRefactorMenu;
          SB.AppendLine(Format('MainMenu has %d top-level items:', [Menu.Items.Count]));
          for I := 0 to Menu.Items.Count - 1 do
          begin
            It := Menu.Items[I];
            SB.AppendLine(Format('  [%2d] Name="%s"  Caption="%s"  Sub=%d  Enabled=%s  %s%s',
              [I, It.Name, It.Caption, It.Count, BoolToStr(It.Enabled, True),
               ActInfo(It), IfThen(It = Found, '   <== matched', '')]));
          end;
          if Found = nil then SB.AppendLine('=> No Refactor menu matched.');
        end;
      end;
      Path := TPath.Combine(TPath.GetTempPath, 'RefactoringLight-mainmenu.log');
      TFile.AppendAllText(Path, SB.ToString + sLineBreak);
    finally
      SB.Free;
    end;
  except
    // diagnostics are best-effort
  end;
end;

// True if a source editor is on top. Deliberately does NOT use
// Editor.GetCurrentContext: that one walks the caret via IOTAEditPosition
// to read the word under the cursor, which COLLAPSES an active selection.
// Since this runs on every action update (IDE idle), it must be cheap and
// must never touch the caret / selection. TopBuffer is a read-only query.
function HasActiveSourceEditor: Boolean;
var
  ES: IOTAEditorServices;
begin
  Result := Supports(BorlandIDEServices, IOTAEditorServices, ES)
    and (ES.TopBuffer <> nil);
end;

procedure TContextMenuInstaller.UpdateMainItemStates;
var
  HasProject, HasEditor: Boolean;
  Pair: TPair<TMenuItem, Integer>;
  En: Boolean;
  Tick: Cardinal;
begin
  if (FItemReq = nil) or (Editor = nil) then Exit;
  // Throttle: this fires on every action-list update (IDE idle). Recompute
  // the states at most a few times per second.
  Tick := GetTickCount;
  if (FLastStateTick <> 0) and (Tick - FLastStateTick < 300) then Exit;
  FLastStateTick := Tick;

  HasProject := Editor.GetCurrentProjectDproj <> '';
  HasEditor := HasActiveSourceEditor;
  for Pair in FItemReq do
  begin
    case Pair.Value of
      REQ_EDITOR:  En := HasProject and HasEditor;
    else           // REQ_PROJECT (and groups)
      En := HasProject;
    end;
    try Pair.Key.Enabled := En; except end;
  end;
end;

procedure TContextMenuInstaller.DoRefactorActionUpdate(Sender: TObject);
begin
  // Let the IDE run its own update first (it disables the menu when there
  // is no refactoring context), then force it back on so the menu - and
  // thus OUR entries - can always be opened.
  if Assigned(FRefOldUpdate) then
    try FRefOldUpdate(Sender); except end;
  if Sender is TContainedAction then
    TContainedAction(Sender).Enabled := True;
  // Gate OUR entries by context (no project -> disabled, etc.).
  UpdateMainItemStates;
end;

procedure TContextMenuInstaller.HookRefactorAction(AParent: TMenuItem);
var
  Act: TBasicAction;
begin
  if AParent = nil then Exit;
  Act := AParent.Action;
  if Act = nil then
  begin
    AParent.Enabled := True;   // no action drives it: a plain enable sticks
    Exit;
  end;
  if Act = FRefactorAction then Exit;   // already hooked this action
  // A different action (first time, or the IDE recreated it): hook OnUpdate.
  FRefactorAction := Act;
  FRefOldUpdate := (Act as TContainedAction).OnUpdate;
  (Act as TContainedAction).OnUpdate := DoRefactorActionUpdate;
  (Act as TContainedAction).Enabled := True;
end;

procedure TContextMenuInstaller.InstallIntoMainMenu;
var
  Parent, Root, Child: TMenuItem;
  FirstTime: Boolean;
begin
  FirstTime := not FMainDumped;
  if FirstTime then
  begin
    FMainDumped := True;
    try TFile.Delete(TPath.Combine(TPath.GetTempPath, 'RefactoringLight-mainmenu.log')); except end;
    DumpMainMenu('before add');
  end;

  Parent := FindRefactorMenu;   // always re-query - the IDE may recreate it
  if Parent = nil then Exit;

  if FMainMenuOwner = nil then
    FMainMenuOwner := TComponent.Create(nil);

  // Build our tree once, then hang its children DIRECTLY into the Refactor
  // menu. No leading separator - we replace the menu's contents entirely.
  if Length(FMainMenuAdded) = 0 then
  begin
    Root := BuildMenuTree(FMainMenuOwner, False);   // no shortcuts here
    while Root.Count > 0 do
    begin
      Child := Root.Items[0];
      Root.Remove(Child);
      FMainMenuAdded := FMainMenuAdded + [Child];
    end;
    Root.Free;   // empty throwaway container
  end;

  // Hide the IDE's own (contextually disabled) Refactor entries, add ours.
  HideExistingEntries(Parent);
  for Child in FMainMenuAdded do
    if Parent.IndexOf(Child) < 0 then Parent.Add(Child);

  // Keep the menu openable against the IDE's action-driven disabling.
  HookRefactorAction(Parent);

  if FirstTime then DumpMainMenu('after add');
end;

function TContextMenuInstaller.IsOurMainItem(AItem: TMenuItem): Boolean;
var
  It: TMenuItem;
begin
  for It in FMainMenuAdded do
    if It = AItem then Exit(True);
  Result := False;
end;

procedure TContextMenuInstaller.HideExistingEntries(AParent: TMenuItem);
var
  I: Integer;
  It: TMenuItem;
  Known: Boolean;
  H: TMenuItem;
begin
  for I := 0 to AParent.Count - 1 do
  begin
    It := AParent.Items[I];
    if IsOurMainItem(It) then Continue;   // never hide our own entries
    if not It.Visible then Continue;
    It.Visible := False;
    Known := False;
    for H in FHidden do if H = It then begin Known := True; Break; end;
    if not Known then FHidden := FHidden + [It];
  end;
end;

procedure TContextMenuInstaller.RemoveFromMainMenu;
var
  Parent, Child: TMenuItem;
  Idx: Integer;
begin
  try
    // Restore the IDE's own entries we hid.
    for Child in FHidden do
      if Child <> nil then
        try Child.Visible := True; except end;

    // Un-hook the Refactor action's OnUpdate.
    if (FRefactorAction <> nil) then
      try (FRefactorAction as TContainedAction).OnUpdate := FRefOldUpdate; except end;

    Parent := FindRefactorMenu;
    if Parent <> nil then
      for Child in FMainMenuAdded do
        if Child <> nil then
        begin
          Idx := Parent.IndexOf(Child);
          if Idx >= 0 then Parent.Remove(Child);
        end;
  except
    // menu already torn down - the owner free below still cleans up
  end;
  FHidden := nil;
  FRefactorAction := nil;
  FRefOldUpdate := nil;
  FMainMenuAdded := nil;
  FreeAndNil(FMainMenuOwner);   // frees the whole main-menu item tree
end;

procedure TContextMenuInstaller.AddOurItems;
var
  Item: TMenuItem;
begin
  if FPopupMenu = nil then Exit;
  for Item in FItems do
    if (Item <> nil) and (FPopupMenu.Items.IndexOf(Item) < 0) then
      FPopupMenu.Items.Add(Item);
end;

procedure TContextMenuInstaller.RemoveOurItems;
var
  Item: TMenuItem;
  Idx: Integer;
begin
  if FPopupMenu = nil then Exit;
  for Item in FItems do
  begin
    if Item = nil then Continue;
    try
      Idx := FPopupMenu.Items.IndexOf(Item);
      if Idx >= 0 then
        FPopupMenu.Items.Remove(Item);
    except
      // popup torn down beneath us - fine, the item is effectively gone
    end;
  end;
end;

procedure TContextMenuInstaller.DoOnPopup(Sender: TObject);
begin
  // Re-entrancy / cycle guard.
  //
  // Other plugins (Parnassus, CnPack) install OnPopup hooks the same
  // way we do. When their SyncTimers run in an unfavorable order the
  // saved FOldOnPopup chain forms a CYCLE:
  //   ours -> Parnassus -> CnPack -> ours -> ...
  // Forwarding to FOldOnPopup again here would restart the cycle and
  // produce a stack overflow. So on re-entry we MUST NOT forward; we
  // simply return. The outer DoOnPopup invocation continues normally
  // and AddOurItems still runs at the end.
  if FInPopup then
    Exit;

  FInPopup := True;
  try
    // 1. Take our items out so the IDE's popup bookkeeping (which
    //    re-registers component names) sees a pristine menu.
    RemoveOurItems;

    // 2. Run the original handler the IDE installed.
    //    Guard against a self-reference (FOldOnPopup pointing back to
    //    our own DoOnPopup) - that would also recurse.
    try
      if Assigned(FOldOnPopup) and not IsOurHandler(FOldOnPopup) then
        FOldOnPopup(Sender);
    except
      // never let an IDE handler exception escape into VCL's popup loop
    end;

    // 3. Put our items back so the user can click them.
    AddOurItems;
  finally
    FInPopup := False;
  end;
end;

procedure TContextMenuInstaller.HookPopup;
begin
  if (FPopupMenu = nil) or FHooked then Exit;
  FOldOnPopup := FPopupMenu.OnPopup;
  FPopupMenu.OnPopup := DoOnPopup;
  FHooked := True;
end;

function TContextMenuInstaller.IsOurHandler(const AHandler: TNotifyEvent): Boolean;
var
  Ours: TNotifyEvent;
begin
  Ours := DoOnPopup;
  Result := (TMethod(AHandler).Code = TMethod(Ours).Code) and
            (TMethod(AHandler).Data = TMethod(Ours).Data);
end;

procedure TContextMenuInstaller.UnhookPopup;
begin
  if (FPopupMenu = nil) or not FHooked then Exit;
  // Only restore if we are still the active handler. If someone else
  // chained on top of us, leave their handler alone - touching it
  // would break their plugin.
  if IsOurHandler(FPopupMenu.OnPopup) then
    FPopupMenu.OnPopup := FOldOnPopup;
  FOldOnPopup := nil;
  FHooked := False;
end;

procedure TContextMenuInstaller.Install;
begin
  if FHooked then Exit; // already installed

  // Try once immediately; if the popup is not yet available, start a
  // timer that keeps retrying until it shows up.
  TryInstall;
  if FHooked then Exit;

  if FRetryTimer = nil then
  begin
    FRetryTimer := TTimer.Create(nil);
    FRetryTimer.Interval := 500;
    FRetryTimer.OnTimer := OnRetryTimer;
  end;
  FRetryCount := 0;
  FRetryTimer.Enabled := True;
end;

procedure TContextMenuInstaller.OnRetryTimer(Sender: TObject);
begin
  TryInstall;
  if FHooked or (FRetryCount >= MaxRetries) then
    FRetryTimer.Enabled := False
  else
    Inc(FRetryCount);
end;

procedure TContextMenuInstaller.OnSyncTimer(Sender: TObject);
var
  Current: TNotifyEvent;
begin
  // Another plugin may have replaced OnPopup with its own handler
  // after we installed ours. Detect and re-hook on top.
  //
  // CAREFUL: if their handler internally chains to ours (e.g. they
  // saved our DoOnPopup as their FOldOnPopup), then naively saving
  // their handler as our new FOldOnPopup would create a cycle:
  // we'd call them, they'd call us, we'd call them, ...
  // The guard in DoOnPopup (re-entry returns immediately, and we
  // never forward to a self-reference) is what actually breaks the
  // cycle at runtime; but we still try to keep FOldOnPopup honest.
  // Re-assert our main-menu entry too - the IDE may rebuild the Refactor
  // menu contextually and drop our submenu.
  try InstallIntoMainMenu; except end;

  if (FPopupMenu = nil) or not FHooked then Exit;
  Current := FPopupMenu.OnPopup;
  if not Assigned(Current) then
  begin
    // Someone cleared the handler entirely - just put ours back.
    FPopupMenu.OnPopup := DoOnPopup;
    Exit;
  end;
  if IsOurHandler(Current) then Exit; // already in place

  // Save the new top-of-chain as our old, install ourselves on top.
  // The DoOnPopup guards prevent runaway recursion if Current
  // transitively chains back to us.
  FOldOnPopup := Current;
  FPopupMenu.OnPopup := DoOnPopup;
end;

procedure TContextMenuInstaller.TryInstall;
begin
  if FHooked then Exit;

  FPopupMenu := FindEditorPopupMenu;
  if FPopupMenu = nil then Exit;

  BuildItems;
  HookPopup;

  // Also hang a copy of our menu into the IDE's top-level Refactor menu.
  try InstallIntoMainMenu; except end;

  // Periodic re-sync so other plugins don't accidentally evict us.
  if FSyncTimer = nil then
  begin
    FSyncTimer := TTimer.Create(nil);
    FSyncTimer.Interval := SyncIntervalMs;
    FSyncTimer.OnTimer := OnSyncTimer;
    FSyncTimer.Enabled := True;
  end;
end;

procedure TContextMenuInstaller.Uninstall;
var
  I: Integer;
begin
  if FRetryTimer <> nil then
  begin
    FRetryTimer.Enabled := False;
    FreeAndNil(FRetryTimer);
  end;
  if FSyncTimer <> nil then
  begin
    FSyncTimer.Enabled := False;
    FreeAndNil(FSyncTimer);
  end;

  // Remove our copy from the IDE main Refactor menu (frees that tree).
  RemoveFromMainMenu;

  // Unhook BEFORE removing items so an in-flight popup that is mid
  // OnPopup callback doesn't see a partially-disassembled state.
  try
    UnhookPopup;
  except
    FHooked := False;
  end;

  if FPopupMenu <> nil then
  begin
    try
      RemoveOurItems;
    except
      // popup gone - fine
    end;
  end;

  // Items are owned by FPopupMenu - DO NOT free them here. FPopupMenu
  // frees its Components when it is itself destroyed by the IDE at
  // shutdown. Freeing them earlier would crash the IDE if it still
  // holds references through Items[] or Action lists.
  for I := 0 to High(FItems) do
    FItems[I] := nil;
  FItems := nil;
  FSeparator := nil;

  FPopupMenu := nil;
end;

procedure TContextMenuInstaller.OnRename(Sender: TObject);
begin
  if WizardInstance <> nil then
    WizardInstance.Execute;
end;

procedure TContextMenuInstaller.OnFindReferences(Sender: TObject);
begin
  if FindReferencesInstance <> nil then
    FindReferencesInstance.Execute;
end;

procedure TContextMenuInstaller.OnFindImplementations(Sender: TObject);
begin
  if FindImplementationsInstance <> nil then
    FindImplementationsInstance.Execute;
end;

procedure TContextMenuInstaller.OnExtractMethod(Sender: TObject);
begin
  if ExtractMethodInstance <> nil then
    ExtractMethodInstance.Execute;
end;

procedure TContextMenuInstaller.OnCompletion(Sender: TObject);
begin
  if CompletionWizardInstance <> nil then
    CompletionWizardInstance.Execute;
end;

procedure TContextMenuInstaller.OnSignatureCheck(Sender: TObject);
begin
  if SignatureCheckInstance <> nil then
    SignatureCheckInstance.Execute;
end;

procedure TContextMenuInstaller.OnRemoveWithAtCursor(Sender: TObject);
begin
  if WithRefactorInstance <> nil then
    WithRefactorInstance.ExecuteAtCursor;
end;

procedure TContextMenuInstaller.OnRemoveWithCurrentUnit(Sender: TObject);
begin
  if WithRefactorInstance <> nil then
    WithRefactorInstance.ExecuteCurrentUnit;
end;

procedure TContextMenuInstaller.OnRemoveWithSelectedUnits(Sender: TObject);
begin
  if WithRefactorInstance <> nil then
    WithRefactorInstance.ExecuteSelectedUnits;
end;

procedure TContextMenuInstaller.OnRemoveWithProjectWide(Sender: TObject);
begin
  if WithRefactorInstance <> nil then
    WithRefactorInstance.ExecuteProjectWide;
end;

procedure TContextMenuInstaller.OnUnitRefs(Sender: TObject);
begin
  if UnitReferencesInstance <> nil then
    UnitReferencesInstance.Execute;
end;

procedure TContextMenuInstaller.OnFindUnit(Sender: TObject);
begin
  Expert.FindUnitDialog.FindUnitForIdentifier;
end;

procedure TContextMenuInstaller.OnMoveToUnit(Sender: TObject);
begin
  if MoveToUnitInstance <> nil then
    MoveToUnitInstance.Execute;
end;

procedure TContextMenuInstaller.OnFindOriginalSymbol(Sender: TObject);
begin
  if FindOriginalSymbolInstance <> nil then
    FindOriginalSymbolInstance.Execute;
end;

procedure TContextMenuInstaller.OnExtractInterface(Sender: TObject);
begin
  Expert.ExtractInterfaceWizard.ExtractInterfaceFromClass;
end;

procedure TContextMenuInstaller.OnAddToExistingInterface(Sender: TObject);
begin
  Expert.ExtractInterfaceWizard.AddToExistingInterface;
end;

procedure TContextMenuInstaller.OnDelegateInterface(Sender: TObject);
begin
  Expert.ExtractInterfaceWizard.DelegateInterfaceImplementation;
end;

procedure TContextMenuInstaller.OnSemanticReplaceCurrent(Sender: TObject);
begin
  Expert.SemanticReplaceWizard.ApplySemanticReplacements_CurrentUnit;
end;

procedure TContextMenuInstaller.OnSemanticReplaceSelected(Sender: TObject);
begin
  Expert.SemanticReplaceWizard.ApplySemanticReplacements_SelectedUnits;
end;

procedure TContextMenuInstaller.OnSemanticReplaceProject(Sender: TObject);
begin
  Expert.SemanticReplaceWizard.ApplySemanticReplacements_Project;
end;

procedure TContextMenuInstaller.OnSemanticReplaceEditRules(Sender: TObject);
begin
  Expert.SemanticReplaceWizard.EditSemanticReplaceRules;
end;

procedure TContextMenuInstaller.OnCheckDfmEvents(Sender: TObject);
begin
  Expert.DfmEventCheckDialog.CheckDfmEventHandlers;
end;

procedure TContextMenuInstaller.OnCheckInterfaceGuids(Sender: TObject);
begin
  Expert.InterfaceGuidDialog.CheckInterfaceGuids;
end;

procedure TContextMenuInstaller.OnCheckCircularRefs(Sender: TObject);
begin
  Expert.CircularRefsDialog.CheckCircularReferences;
end;

end.
