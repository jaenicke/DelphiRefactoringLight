(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ContextMenu;

{
  Puts a "Refactoring Light" submenu into the Delphi IDE code editor's
  context menu, and a copy of the same tree into the IDE's top-level
  Refactor menu.

  TWO PATHS, in this order:

  1. PREFERRED - the official ToolsAPI (Delphi 12+):
     IOTAEditorServices.GetEditorLocalMenu -> INTAEditorLocalMenu
     .RegisterActionList(<our TActionList>, <own category>,
                         InsertAfter = cEdMenuCatRefactor).
     The IDE rebuilds the local menu on every use and calls each
     action's OnUpdate first, so nothing of ours stays inside the
     popup between uses: no OnPopup hooking, no component-name
     clashes, and the IDE's own (dynamically added) entries are never
     touched. The registration MUST be undone before the BPL unloads
     (RemoveLocalMenu).

  2. FALLBACK (older IDEs / registration refused) - hook the popup's
     OnPopup, like CnPack's CnMenuHook does:
       a. remove our items so the menu looks pristine,
       b. run the original handler (the IDE bookkeeping sees only its
          own items - permanently parked TMenuItems otherwise trigger
          "A component named X already exists"),
       c. add our items back.
     HARD RULE on this path: capture the forward target ONCE and never
     re-take the top of the chain later - overwriting it drops the
     IDE's own handler out of the chain, and with it every entry the
     IDE adds per popup ("Find declaration" & co).

  The editor popup / the IDE services may not exist yet when Register
  runs, so the install is retried via TTimer until one path succeeds or
  the retry budget is exhausted.
}

interface

uses
  System.Classes, System.Generics.Collections, Vcl.Menus, Vcl.ExtCtrls,
  Vcl.ActnList,
  Expert.Shortcuts;

type
  /// <summary>Owns a menu item's real handler and gates it through the
  ///  shortcut debounce (see TExpertsShortCut.AllowAction).</summary>
  TMenuActionProxy = class(TComponent)
  private
    FKind: TShortcutKind;
    FInner: TNotifyEvent;
  public
    procedure Click(Sender: TObject);
  end;

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
    // Editor-popup takeover: the IDE's own "Refactoring" popup entry we
    // hid for the current popup round (re-found and re-hidden per popup,
    // restored in RemoveOurItems / on uninstall).
    FPopupTakenOver: TMenuItem;
    FPopupDumped: Boolean;
    FPopupNoMatchDumped: Boolean;
    // OFFICIAL editor-popup integration (Delphi 12+): an action list
    // registered through INTAEditorLocalMenu. When this works we never
    // touch the popup's OnPopup chain at all.
    FLocalMenuOwner: TComponent;      // owns the action list + source tree
    FLocalMenuActions: TActionList;
    FLocalMenuOk: Boolean;
    FActReq: TDictionary<TObject, Integer>;      // action -> requirement
    FLocalReq: TDictionary<TMenuItem, Integer>;  // local tree -> requirement
    FLocalLastSep: Boolean;   // last emitted local-menu action was a separator
    // Context cache: ReqState runs per menu item / per action, and
    // GetCurrentProjectDproj walks every open module - query at most a
    // few times per second instead.
    FCtxTick: Cardinal;
    FCtxProject: Boolean;
    FCtxEditor: Boolean;
    procedure MenuContext(out AHasProject, AHasEditor: Boolean);
    /// <summary>Registers our tree via INTAEditorLocalMenu. False when
    ///  the IDE does not expose it (pre-12) or registration fails - the
    ///  caller then falls back to the OnPopup hook.</summary>
    function TryInstallLocalMenu: Boolean;
    procedure BuildLocalMenuActions(AParent: TMenuItem;
      const ACat, APrefix: string);
    procedure DoLocalActionUpdate(Sender: TObject);
    /// <summary>The submenu ROOT must stay openable no matter what the
    ///  context is - otherwise a disabled parent hides every entry (the
    ///  IDE greys a parent whose action carries no handler / whose
    ///  children are all disabled). Same reasoning as HookRefactorAction
    ///  for the main menu.</summary>
    procedure DoLocalRootUpdate(Sender: TObject);
    procedure DoLocalRootExecute(Sender: TObject);
    procedure RemoveLocalMenu;
    /// <summary>Enabled state (and annotated caption) for a context
    ///  requirement - shared by the main-menu updater and the editor
    ///  local-menu actions.</summary>
    function ReqState(AReq: Integer; const ABase: string;
      out ACaption: string): Boolean;
    procedure UpdateMainItemStates;
    procedure OnRename(Sender: TObject);
    procedure OnFindReferences(Sender: TObject);
    procedure OnFindImplementations(Sender: TObject);
    procedure OnFindOriginal(Sender: TObject);
    procedure OnExtractMethod(Sender: TObject);
    procedure OnCompletion(Sender: TObject);
    procedure OnSignatureCheck(Sender: TObject);
    procedure OnRemoveWithProjectWide(Sender: TObject);
    procedure OnRemoveWithCurrentUnit(Sender: TObject);
    procedure OnRemoveWithSelectedUnits(Sender: TObject);
    procedure OnRemoveWithAtCursor(Sender: TObject);
    procedure OnUnitRefs(Sender: TObject);
    procedure OnFindUnit(Sender: TObject);
    procedure OnAddUnitAtCursor(Sender: TObject);
    procedure OnShowQuickFixes(Sender: TObject);
    procedure OnResolveMissingUnits(Sender: TObject);
    procedure OnUsesCleanup(Sender: TObject);
    procedure OnMoveToUnit(Sender: TObject);
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
      AOnClick: TNotifyEvent; AKind: TShortcutKind): TMenuItem;
    /// <summary>Re-applies the ShortCut property of AItem and all its
    ///  children from the current settings (Tag carries the kind).</summary>
    procedure RefreshShortcutsIn(AItem: TMenuItem);
    /// <summary>Builds the menu tree owned by AOwner. AReqMap (when
    ///  assigned) receives each item's context requirement - pass the map
    ///  belonging to THAT tree: mixing trees into one map means updating
    ///  invisible items, and freeing one tree leaves dangling keys.</summary>
    function BuildMenuTree(AOwner: TComponent;
      AReqMap: TDictionary<TMenuItem, Integer>): TMenuItem;
    function FindRefactorMenu: TMenuItem;
    function FindPopupRefactorItem: TMenuItem;
    procedure DumpMainMenu(const AReason: string);
    procedure DumpEditorPopup(const AReason: string);
    procedure LogPopupEvent(const AText: string);
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
  Vcl.Forms, Vcl.Controls,
  Expert.EditorHelperIntf,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod,
  Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard, Expert.WithRefactorWizard, Expert.UnitReferencesWizard,
  Expert.MoveToUnitWizard, Expert.ExtractInterfaceWizard,
  Expert.SemanticReplaceWizard, Expert.DfmEventCheckDialog,
  Expert.InterfaceGuidDialog, Expert.CircularRefsDialog,
  Expert.FindUnitDialog, Expert.AutoImport, Expert.FindOriginalSymbolWizard,
  Expert.UsesCleanup;

const
  /// <summary>Maximum retry attempts when the editor popup is not yet
  ///  available. With a 500 ms interval, 40 retries give up to 20 s.</summary>
  MaxRetries = 40;

  /// <summary>Period at which we verify our OnPopup hook is still in
  ///  place. Other plugins may install their own OnPopup later and
  ///  unintentionally displace ours.</summary>
  SyncIntervalMs = 2000;

  // Context requirement for a main-menu entry.
  REQ_PROJECT  = 1;  // needs an open project
  REQ_EDITOR   = 2;  // needs an active source editor (cursor context)
  REQ_LIVEALL  = 4;  // like REQ_EDITOR; caption annotated with the TOTAL
                     // live quick-fix count ("Show all quick fixes (N)")

  // Base caption of the REQ_LIVEALL entry - annotated with the count by
  // ReqState, so both menu copies and the action list stay in sync.
  CapShowAllFixes = 'Show all quick fixes...';

  // Our category in the editor local menu (INTAEditorLocalMenu).
  LocalMenuCategory = 'RefactoringLight';
  REQ_LIVEFIX = 3;   // like REQ_EDITOR, but additionally reflects the live
                     // auto-import check: disabled when the checker KNOWS
                     // there is nothing to fix; annotated with the count
                     // when it knows there is.

{ TContextMenuInstaller }

destructor TContextMenuInstaller.Destroy;
begin
  Uninstall;
  FreeAndNil(FActReq);
  FreeAndNil(FLocalReq);
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

{ TMenuActionProxy }

procedure TMenuActionProxy.Click(Sender: TObject);
begin
  // One keystroke can reach us twice (menu accelerator + keyboard
  // binding) - run the action once.
  if not TExpertsShortCut.AllowAction(FKind) then Exit;
  if Assigned(FInner) then FInner(Sender);
end;

function TContextMenuInstaller.MakeItem(AOwner: TComponent; const ACaption: string;
  AOnClick: TNotifyEvent; AKind: TShortcutKind): TMenuItem;
begin
  Result := TMenuItem.Create(AOwner);
  Result.Caption := ACaption;
  // Tag stores the shortcut kind so the refreshers can find this item.
  // Encoded as Ord+1 so 0 means "no shortcut tracked".
  Result.Tag := Ord(AKind) + 1;
  // Both copies show the key natively (the IDE menu paints the ShortCut
  // property; a caption with #9 is NOT rendered right-aligned there).
  // Popup shortcuts are display-only in the VCL, main-MENU ones really
  // dispatch - so the click goes through a proxy that debounces against
  // our IOTAKeyboardBinding firing for the same keystroke.
  Result.ShortCut := TExpertsShortCut.Shortcuts[AKind];
  var Proxy := TMenuActionProxy.Create(AOwner);
  Proxy.FKind := AKind;
  Proxy.FInner := AOnClick;
  Result.OnClick := Proxy.Click;
end;

procedure TContextMenuInstaller.RefreshShortcutsIn(AItem: TMenuItem);
var
  I: Integer;
begin
  if AItem = nil then Exit;
  if (AItem.Tag >= 1) and (AItem.Tag <= Ord(High(TShortcutKind)) + 1) then
    AItem.ShortCut := TExpertsShortCut.Shortcuts[TShortcutKind(AItem.Tag - 1)];
  for I := 0 to AItem.Count - 1 do
    RefreshShortcutsIn(AItem.Items[I]);
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
  // The main-menu copy is a separate tree - refresh it recursively.
  for Item in FMainMenuAdded do
    RefreshShortcutsIn(Item);
  // ... and so are the editor local-menu ACTIONS: they copied the key at
  // build time, so without this the popup advertises the old shortcut
  // until the IDE restarts.
  if FLocalMenuActions <> nil then
    for var I := 0 to FLocalMenuActions.ActionCount - 1 do
      if FLocalMenuActions.Actions[I] is TCustomAction then
      begin
        var Act := TCustomAction(FLocalMenuActions.Actions[I]);
        if (Act.Tag >= 1) and (Act.Tag <= Ord(High(TShortcutKind)) + 1) then
          Act.ShortCut := TExpertsShortCut.Shortcuts[TShortcutKind(Act.Tag - 1)];
      end;
end;

// Builds the whole "Refactoring Light" submenu tree owned by AOwner and
// returns its root item. Every leaf carries its shortcut (display) and
// routes its click through the debounce proxy; AReqMap collects the
// context requirements for whoever updates THAT tree.
function TContextMenuInstaller.BuildMenuTree(AOwner: TComponent;
  AReqMap: TDictionary<TMenuItem, Integer>): TMenuItem;

  procedure Req(AItem: TMenuItem; AReq: Integer);
  begin
    if AReqMap <> nil then AReqMap.AddOrSetValue(AItem, AReq);
  end;

  function Leaf(AParent: TMenuItem; const ACaption: string; AOnClick: TNotifyEvent;
    AKind: TShortcutKind; AReq: Integer): TMenuItem;
  begin
    Result := MakeItem(AOwner, ACaption, AOnClick, AKind);
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

  procedure Sep(AParent: TMenuItem);
  begin
    var Item := TMenuItem.Create(AOwner);
    Item.Caption := '-';
    AParent.Add(Item);          // no Req: separators are never dis/enabled
  end;

var
  Root, RemoveWithSub, IfaceSub, SemSub, ChecksSub: TMenuItem;
begin
  Root := TMenuItem.Create(AOwner);
  Root.Caption := 'Refactoring Light';

  // ---- Refactor code at the cursor ----------------------------------------
  Leaf(Root, 'Rename...',                 OnRename,               skRename,     REQ_EDITOR);
  Leaf(Root, 'Extract Method',            OnExtractMethod,        skExtract,    REQ_EDITOR);
  Leaf(Root, 'Align method signature...', OnSignatureCheck,       skAlign,      REQ_EDITOR);
  Leaf(Root, 'Move to unit...',           OnMoveToUnit,           skMoveToUnit, REQ_EDITOR);

  RemoveWithSub := Sub(Root, 'Remove with');
  Leaf(RemoveWithSub, 'At cursor only',     OnRemoveWithAtCursor,      skRemoveWith, REQ_EDITOR);
  Plain(RemoveWithSub, 'In current unit',    OnRemoveWithCurrentUnit,   REQ_EDITOR);
  Plain(RemoveWithSub, 'In selected units...', OnRemoveWithSelectedUnits, REQ_PROJECT);
  Plain(RemoveWithSub, 'In whole project...', OnRemoveWithProjectWide,   REQ_PROJECT);

  IfaceSub := Sub(Root, 'Interfaces');
  Plain(IfaceSub, 'Extract new interface from class...', OnExtractInterface,       REQ_EDITOR);
  Plain(IfaceSub, 'Add to existing interface...',        OnAddToExistingInterface, REQ_EDITOR);
  Plain(IfaceSub, 'Add IInterface support to class...',  OnDelegateInterface,      REQ_EDITOR);

  // ---- Find / navigate ----------------------------------------------------
  Sep(Root);
  Leaf(Root, 'Find References',           OnFindReferences,       skFindRef,    REQ_EDITOR);
  Leaf(Root, 'Find Implementations',      OnFindImplementations,  skFindImp,    REQ_EDITOR);
  Leaf(Root, 'Find original symbol',      OnFindOriginal,         skFindOriginal, REQ_EDITOR);
  Leaf(Root, 'Find unit references...',   OnUnitRefs,             skUnitRefs,   REQ_EDITOR);

  // ---- Uses clause / missing units ----------------------------------------
  Sep(Root);
  Plain(Root, CapShowAllFixes,            OnShowQuickFixes,       REQ_LIVEALL);
  Plain(Root, 'Add unit for identifier at cursor', OnAddUnitAtCursor, REQ_LIVEFIX);
  Plain(Root, 'Resolve missing units...', OnResolveMissingUnits,  REQ_EDITOR);
  Plain(Root, 'Find unit for identifier...', OnFindUnit,          REQ_PROJECT);
  Plain(Root, 'Uses cleanup (current unit)...', OnUsesCleanup,    REQ_EDITOR);

  // ---- Project-wide tools & checks ----------------------------------------
  Sep(Root);
  SemSub := Sub(Root, 'Semantic replace');
  Plain(SemSub, 'In current unit',    OnSemanticReplaceCurrent,   REQ_EDITOR);
  Plain(SemSub, 'In selected units...', OnSemanticReplaceSelected, REQ_PROJECT);
  Plain(SemSub, 'In whole project...', OnSemanticReplaceProject,   REQ_PROJECT);
  Plain(SemSub, 'Edit rules...',       OnSemanticReplaceEditRules, REQ_PROJECT);

  ChecksSub := Sub(Root, 'Project checks');
  Plain(ChecksSub, 'DFM event handlers...',        OnCheckDfmEvents,       REQ_PROJECT);
  Plain(ChecksSub, 'Interface GUIDs...',           OnCheckInterfaceGuids,  REQ_PROJECT);
  Plain(ChecksSub, 'Circular unit references...',  OnCheckCircularRefs,    REQ_PROJECT);

  Sep(Root);
  Leaf(Root, 'Code Completion',           OnCompletion,           skCompletion, REQ_EDITOR);

  Result := Root;
end;

procedure TContextMenuInstaller.BuildItems;
var
  Submenu: TMenuItem;
begin
  if Length(FItems) > 0 then Exit;

  Submenu := BuildMenuTree(FPopupMenu, nil);   // popup: no state map

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

const
  // Captions of the IDE's Refactor menu / popup entry across the shipped
  // IDE languages (accelerators/ellipsis already stripped by NormCaption).
  RefactorCaptions: array[0..6] of string =
    ('REFACTOR', 'REFACTORING', 'REFAKTORIEREN', 'REFACTORISER',
     'RIFATTORIZZA', 'REFATORAR', 'REFACTORIZAR');

function IsRefactorCaption(const ACaption: string): Boolean;
var
  Cap: string;
  K: Integer;
begin
  Cap := NormCaption(ACaption);
  for K := Low(RefactorCaptions) to High(RefactorCaptions) do
    if Cap = RefactorCaptions[K] then Exit(True);
  Result := False;
end;

function TContextMenuInstaller.FindRefactorMenu: TMenuItem;
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

// The IDE's own "Refactoring" entry inside the editor popup - found per
// popup round (the IDE rebuilds/updates the popup in its own OnPopup).
// Matched language-independently: component name containing 'refactor',
// else normalized caption against the shipped translations. Our own
// submenu (which may carry the same caption after the takeover) is
// excluded via the FItems check.
function TContextMenuInstaller.FindPopupRefactorItem: TMenuItem;

  function IsOurs(AItem: TMenuItem): Boolean;
  var
    Own: TMenuItem;
  begin
    for Own in FItems do
      if Own = AItem then Exit(True);
    Result := False;
  end;

var
  I: Integer;
  It: TMenuItem;
begin
  Result := nil;
  if FPopupMenu = nil then Exit;
  for I := 0 to FPopupMenu.Items.Count - 1 do
  begin
    It := FPopupMenu.Items[I];
    if IsOurs(It) then Continue;
    if (Pos('REFACTOR', UpperCase(It.Name)) > 0) or IsRefactorCaption(It.Caption) then
      Exit(It);
  end;
end;

procedure TContextMenuInstaller.DumpEditorPopup(const AReason: string);
var
  SB: TStringBuilder;
  I: Integer;
  It: TMenuItem;
  Path, ActName: string;
begin
  try
    SB := TStringBuilder.Create;
    try
      SB.AppendLine('--- Refactoring Light editor-popup dump (' + AReason + ') ---');
      if FPopupMenu = nil then
        SB.AppendLine('FPopupMenu is nil')
      else
      begin
        SB.AppendLine(Format('Popup "%s" has %d top-level items:',
          [FPopupMenu.Name, FPopupMenu.Items.Count]));
        for I := 0 to FPopupMenu.Items.Count - 1 do
        begin
          It := FPopupMenu.Items[I];
          if It.Action <> nil then ActName := It.Action.Name else ActName := '-';
          SB.AppendLine(Format('  [%2d] Name="%s"  Caption="%s"  Sub=%d  Enabled=%s  Visible=%s  Action=%s',
            [I, It.Name, It.Caption, It.Count, BoolToStr(It.Enabled, True),
             BoolToStr(It.Visible, True), ActName]));
        end;
      end;
      Path := TPath.Combine(TPath.GetTempPath, 'RefactoringLight-editorpopup.log');
      TFile.AppendAllText(Path, SB.ToString + sLineBreak);
    finally
      SB.Free;
    end;
  except
    // diagnostics are best-effort
  end;
end;

procedure TContextMenuInstaller.LogPopupEvent(const AText: string);
begin
  try
    TFile.AppendAllText(
      TPath.Combine(TPath.GetTempPath, 'RefactoringLight-editorpopup.log'),
      FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + '  ' + AText + sLineBreak);
  except
    // diagnostics are best-effort
  end;
end;

// Owner class of a TNotifyEvent's method pointer - identifies WHOSE
// handler sits in the OnPopup chain (IDE window vs another plugin).
function HandlerOwnerClass(const AHandler: TNotifyEvent): string;
begin
  Result := 'nil';
  if not Assigned(AHandler) then Exit;
  try
    if TMethod(AHandler).Data <> nil then
      Result := TObject(TMethod(AHandler).Data).ClassName
    else
      Result := 'no-data';
  except
    Result := 'unknown';
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

procedure TContextMenuInstaller.MenuContext(out AHasProject, AHasEditor: Boolean);
var
  Tick: Cardinal;
begin
  // GetCurrentProjectDproj walks every open module - too expensive to
  // repeat per menu item (ReqState runs once per entry, and the IDE
  // rebuilds the local menu action by action).
  Tick := GetTickCount;
  if (FCtxTick = 0) or (Tick - FCtxTick > 250) then
  begin
    FCtxTick := Tick;
    FCtxProject := (Editor <> nil) and (Editor.GetCurrentProjectDproj <> '');
    FCtxEditor := HasActiveSourceEditor;
  end;
  AHasProject := FCtxProject;
  AHasEditor := FCtxEditor;
end;

function TContextMenuInstaller.ReqState(AReq: Integer; const ABase: string;
  out ACaption: string): Boolean;
var
  HasProject, HasEditor: Boolean;
  LiveCount: Integer;
  AllFixes: TArray<TQuickFix>;
begin
  ACaption := ABase;
  MenuContext(HasProject, HasEditor);
  case AReq of
    REQ_EDITOR: Result := HasProject and HasEditor;
    REQ_LIVEFIX:
      begin
        // Fresh live data with zero missing units -> nothing to do at the
        // caret. No count annotation here: it is file-wide, this action is
        // not (that read as a broken promise).
        Result := HasProject and HasEditor;
        if Result and LiveFreshInfo(Editor.GetActiveFileName, LiveCount) then
          Result := LiveCount > 0;
      end;
    REQ_LIVEALL:
      begin
        // File-wide overview: annotate with the TOTAL fix count.
        Result := HasProject and HasEditor;
        if Result and LiveAllFixes(Editor.GetActiveFileName, AllFixes) then
        begin
          Result := Length(AllFixes) > 0;
          if Result then
            ACaption := Format('%s (%d found)', [ABase, Length(AllFixes)]);
        end;
      end;
  else          // REQ_PROJECT (and group headers)
    Result := HasProject;
  end;
end;

function TContextMenuInstaller.TryInstallLocalMenu: Boolean;
// The OFFICIAL way to extend the editor popup (Delphi 12+): register a
// TActionList under our own category, inserted after the IDE's Refactor
// category. The IDE rebuilds the menu on every popup and calls each
// action's OnUpdate first - so no OnPopup hooking, no re-entrancy guard,
// and the IDE's own dynamic entries (Find declaration & co) are never
// disturbed.
var
  ES: IOTAEditorServices;
  LocalMenu: INTAEditorLocalMenu;
  Root: TMenuItem;
  RootAct: TAction;
begin
  Result := False;
  if FLocalMenuOk then Exit(True);
  if not Supports(BorlandIDEServices, IOTAEditorServices, ES) then Exit;
  try
    LocalMenu := ES.GetEditorLocalMenu;
  except
    LocalMenu := nil;
  end;
  if LocalMenu = nil then Exit;

  if FActReq = nil then FActReq := TDictionary<TObject, Integer>.Create;
  if FLocalReq = nil then FLocalReq := TDictionary<TMenuItem, Integer>.Create;
  FLocalMenuOwner := TComponent.Create(nil);
  FLocalMenuActions := TActionList.Create(FLocalMenuOwner);

  // Source of truth stays BuildMenuTree - the action list mirrors it, so
  // popup and main menu can never drift apart. The tree stays alive: the
  // actions call its (debounced) click handlers.
  Root := BuildMenuTree(FLocalMenuOwner, FLocalReq);
  Root.Caption := 'Refactoring Light';

  RootAct := TAction.Create(FLocalMenuOwner);
  RootAct.ActionList := FLocalMenuActions;
  RootAct.Caption := Root.Caption;
  RootAct.Category := LocalMenuCategory;
  RootAct.OnUpdate := DoLocalRootUpdate;
  RootAct.OnExecute := DoLocalRootExecute;
  RootAct.Enabled := True;
  // Sub-items must FOLLOW their parent and carry the parent's category
  // plus '.<name>' (INTAEditorLocalMenu contract).
  FLocalLastSep := True;   // nothing emitted yet -> no leading separator
  BuildLocalMenuActions(Root, LocalMenuCategory + '.Items', '');
  // A group at the very end would leave a dangling separator.
  if (FLocalMenuActions.ActionCount > 0)
    and (TCustomAction(FLocalMenuActions.Actions[
           FLocalMenuActions.ActionCount - 1]).Caption = '-') then
    FLocalMenuActions.Actions[FLocalMenuActions.ActionCount - 1].Free;

  try
    LocalMenu.RegisterActionList(FLocalMenuActions, LocalMenuCategory,
      cEdMenuCatRefactor);
    FLocalMenuOk := True;
    Result := True;
    LogPopupEvent('editor local menu registered via INTAEditorLocalMenu');
  except
    on E: Exception do
    begin
      LogPopupEvent('INTAEditorLocalMenu registration failed: ' + E.Message);
      FLocalMenuActions := nil;
      // The owner frees the whole mirror tree AND its actions - their
      // pointers must not stay behind as dictionary keys (the updaters
      // would then write into freed memory).
      FreeAndNil(FLocalMenuOwner);
      FLocalReq.Clear;
      FActReq.Clear;
    end;
  end;
end;

procedure TContextMenuInstaller.BuildLocalMenuActions(AParent: TMenuItem;
  const ACat, APrefix: string);
// The editor local menu renders exactly ONE nesting level below the
// registered category (verified empirically: a second '.<name>' level is
// NOT turned into a submenu, its actions just land flat next to their
// parent - which then sits there as a dead row). Our tree has two levels,
// so a GROUP is flattened into a separator-delimited block whose entries
// carry the group name as a caption prefix ("Remove with > In current
// unit"). The main-menu copy keeps the real submenus.
var
  I, Req: Integer;
  Item: TMenuItem;
  Act: TAction;

  procedure EmitSeparator;
  var
    Sep: TAction;
  begin
    if FLocalLastSep then Exit;   // never two in a row / none at the top
    Sep := TAction.Create(FLocalMenuOwner);
    Sep.ActionList := FLocalMenuActions;
    Sep.Caption := '-';
    Sep.Category := ACat;
    Sep.Enabled := False;
    FLocalLastSep := True;
  end;

begin
  for I := 0 to AParent.Count - 1 do
  begin
    Item := AParent.Items[I];

    if Item.Count > 0 then
    begin
      // A group: fence it off and prefix its entries with the group name.
      EmitSeparator;
      BuildLocalMenuActions(Item, ACat, Item.Caption + ' > ');
      EmitSeparator;
      Continue;
    end;

    if Item.IsLine then
    begin
      EmitSeparator;
      Continue;
    end;

    Act := TAction.Create(FLocalMenuOwner);
    Act.ActionList := FLocalMenuActions;
    Act.Caption := APrefix + Item.Caption;
    Act.Category := ACat;
    Act.Tag := Item.Tag;
    Act.OnExecute := Item.OnClick;    // the debounce proxy from MakeItem
    Act.ShortCut := Item.ShortCut;
    Act.OnUpdate := DoLocalActionUpdate;
    if (FLocalReq <> nil) and FLocalReq.TryGetValue(Item, Req) then
      FActReq.AddOrSetValue(Act, Req);
    FLocalLastSep := False;
  end;
end;

procedure TContextMenuInstaller.DoLocalRootUpdate(Sender: TObject);
begin
  if Sender is TCustomAction then
    TCustomAction(Sender).Enabled := True;
end;

procedure TContextMenuInstaller.DoLocalRootExecute(Sender: TObject);
begin
  // A parent item only opens its submenu - but an action WITHOUT an
  // execute handler is greyed out by the action machinery.
end;

procedure TContextMenuInstaller.DoLocalActionUpdate(Sender: TObject);
var
  Req: Integer;
  Cap: string;
  Act: TCustomAction;
begin
  if not (Sender is TCustomAction) then Exit;
  Act := TCustomAction(Sender);
  if (FActReq = nil) or not FActReq.TryGetValue(Act, Req) then Exit;
  try
    Act.Enabled := ReqState(Req, CapShowAllFixes, Cap);
    if (Req = REQ_LIVEALL) and (Act.Caption <> Cap) then Act.Caption := Cap;
    // (the LIVEALL entry is never inside a group, so the base caption
    //  needs no prefix handling)
  except
    // never let an exception escape into the IDE's menu build
  end;
end;

procedure TContextMenuInstaller.RemoveLocalMenu;
var
  ES: IOTAEditorServices;
  LocalMenu: INTAEditorLocalMenu;
begin
  if not FLocalMenuOk then Exit;
  FLocalMenuOk := False;
  // MANDATORY before the BPL unloads (documented in the ToolsAPI).
  try
    if Supports(BorlandIDEServices, IOTAEditorServices, ES) then
    begin
      LocalMenu := ES.GetEditorLocalMenu;
      if LocalMenu <> nil then
        LocalMenu.UnregisterActionList(LocalMenuCategory);
    end;
  except
    // IDE shutting down - nothing left to unregister
  end;
  FLocalMenuActions := nil;
  FreeAndNil(FLocalMenuOwner);   // frees the mirror tree + its actions
  if FActReq <> nil then FActReq.Clear;
  if FLocalReq <> nil then FLocalReq.Clear;
end;

procedure TContextMenuInstaller.UpdateMainItemStates;
var
  Pair: TPair<TMenuItem, Integer>;
  En: Boolean;
  Cap: string;
  Tick: Cardinal;
begin
  if (FItemReq = nil) or (Editor = nil) then Exit;
  // Throttle: this fires on every action-list update (IDE idle). Recompute
  // the states at most a few times per second.
  Tick := GetTickCount;
  if (FLastStateTick <> 0) and (Tick - FLastStateTick < 300) then Exit;
  FLastStateTick := Tick;

  for Pair in FItemReq do
  begin
    En := ReqState(Pair.Value, CapShowAllFixes, Cap);
    // Only write on an actual change - TMenuItem setters trigger
    // MenuChanged up to the menu bar, which repaints it.
    try
      if (Pair.Value = REQ_LIVEALL) and (Pair.Key.Caption <> Cap) then
        Pair.Key.Caption := Cap;
    except
    end;
    try
      if Pair.Key.Enabled <> En then Pair.Key.Enabled := En;
    except
    end;
  end;
end;

procedure TContextMenuInstaller.DoRefactorActionUpdate(Sender: TObject);
begin
  // Deliberately do NOT chain the IDE's own update handler. Its only
  // visible effect is disabling the action when there is no IDE
  // refactoring context - and since we re-enable right afterwards, the
  // False->True toggle repainted the top-level Refactor item on EVERY
  // action-update cycle (which fires on mouse moves): the menu entry
  // visibly flickered. The IDE's own entries under this menu are hidden
  // while we are installed, so nothing depends on that handler running.
  // Setting Enabled to an unchanged value is a no-op (no repaint).
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
    if FItemReq = nil then FItemReq := TDictionary<TMenuItem, Integer>.Create;
    Root := BuildMenuTree(FMainMenuOwner, FItemReq);
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
  Item, IdeRef, Submenu: TMenuItem;
  Idx: Integer;
begin
  if FPopupMenu = nil then Exit;
  if not FPopupDumped then
  begin
    FPopupDumped := True;
    DumpEditorPopup('first popup');
  end;

  // FItems = [Separator, Submenu] - see BuildItems.
  Submenu := nil;
  if Length(FItems) > 1 then Submenu := FItems[1];

  // Take over the IDE's own "Refactoring" popup entry (analogous to the
  // main-menu Refactor takeover): hide the IDE's entry and show our tree
  // at the SAME position under the SAME (localized) caption. The entry is
  // re-found on every popup because the IDE rebuilds its popup content.
  IdeRef := FindPopupRefactorItem;
  if (IdeRef <> nil) and (Submenu <> nil) then
  begin
    FPopupTakenOver := IdeRef;
    try
      IdeRef.Visible := False;
      Submenu.Caption := IdeRef.Caption;
      Idx := FPopupMenu.Items.IndexOf(IdeRef);
      if Idx < 0 then Idx := 0;
      if FPopupMenu.Items.IndexOf(Submenu) < 0 then
        FPopupMenu.Items.Insert(Idx, Submenu);
      Exit;   // takeover mode: no trailing separator + submenu
    except
      // popup rebuilt beneath us - fall through to the append fallback
      FPopupTakenOver := nil;
    end;
  end
  else if (IdeRef = nil) and not FPopupNoMatchDumped then
  begin
    FPopupNoMatchDumped := True;
    DumpEditorPopup('no Refactoring item matched - using append fallback');
  end;

  // Fallback (no IDE Refactoring entry found): the classic appended
  // separator + "Refactoring Light" submenu at the end of the popup.
  if Submenu <> nil then Submenu.Caption := 'Refactoring Light';
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

  // Restore the IDE's popup entry we hid in the previous round, so the
  // IDE's own popup bookkeeping sees its menu unmodified.
  if FPopupTakenOver <> nil then
  begin
    try
      FPopupTakenOver.Visible := True;
    except
      // item freed by the IDE between popups - nothing to restore
    end;
    FPopupTakenOver := nil;
  end;

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
  // FOldOnPopup is our permanent forward target: at hook time the chain
  // below us is intact and ends at the IDE's own handler (which adds the
  // DYNAMIC popup entries - "Deklaration suchen" & co). It must never be
  // overwritten later.
  LogPopupEvent('hooked OnPopup; forward target owner: ' +
    HandlerOwnerClass(FOldOnPopup));
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
  if FHooked or FLocalMenuOk then Exit; // already installed

  // Try once immediately; if neither path is available yet (no editor
  // popup, services not up), start a timer that keeps retrying.
  TryInstall;
  if FHooked or FLocalMenuOk then Exit;

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
  if FHooked or FLocalMenuOk or (FRetryCount >= MaxRetries) then
    FRetryTimer.Enabled := False
  else
    Inc(FRetryCount);
end;

procedure TContextMenuInstaller.OnSyncTimer(Sender: TObject);
var
  Current: TNotifyEvent;
begin
  // Re-assert our main-menu entry - the IDE may rebuild the Refactor
  // menu contextually and drop our submenu.
  try InstallIntoMainMenu; except end;

  if (FPopupMenu = nil) or not FHooked then Exit;
  Current := FPopupMenu.OnPopup;
  if not Assigned(Current) then
  begin
    // Someone cleared the handler entirely - just put ours back.
    LogPopupEvent('sync: OnPopup was nil - reinstalled ours');
    FPopupMenu.OnPopup := DoOnPopup;
    Exit;
  end;

  // A FOREIGN handler on top is FINE and must be left alone: whoever
  // hooked after us saved OUR handler as their forward target, so we
  // still run through their chain. The previous behavior re-took the
  // top and saved the foreign handler as our new FOldOnPopup - that
  // OVERWROTE our pointer to the IDE's ORIGINAL handler and formed the
  // cycle ours -> Parnassus -> ours: the guard in DoOnPopup stopped the
  // stack overflow, but the IDE handler had dropped out of the chain
  // entirely, so its DYNAMIC popup entries ("Deklaration suchen" & co,
  // added on every popup by the IDE itself) vanished for good.
end;

procedure TContextMenuInstaller.TryInstall;
begin
  if FHooked or FLocalMenuOk then Exit;

  // PREFERRED: the official editor-menu API (Delphi 12+). Only when it is
  // unavailable do we fall back to hooking the popup's OnPopup chain -
  // that path has to hide/restore the IDE's own Refactoring entry on
  // every popup and must never break the handler chain.
  try
    if TryInstallLocalMenu then
    begin
      try InstallIntoMainMenu; except end;
      if FSyncTimer = nil then
      begin
        FSyncTimer := TTimer.Create(nil);
        FSyncTimer.Interval := SyncIntervalMs;
        FSyncTimer.OnTimer := OnSyncTimer;
        FSyncTimer.Enabled := True;   // keeps the MAIN menu entry alive
      end;
      Exit;
    end;
  except
    // fall through to the legacy hook
  end;

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

  // Official editor-menu registration must go before the BPL unloads.
  RemoveLocalMenu;

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

procedure TContextMenuInstaller.OnFindOriginal(Sender: TObject);
begin
  FindOriginalSymbol;
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

procedure TContextMenuInstaller.OnAddUnitAtCursor(Sender: TObject);
begin
  Expert.AutoImport.AddUnitForIdentifierAtCursor;
end;

procedure TContextMenuInstaller.OnShowQuickFixes(Sender: TObject);
begin
  Expert.AutoImport.ShowAllQuickFixes;
end;

procedure TContextMenuInstaller.OnResolveMissingUnits(Sender: TObject);
begin
  Expert.AutoImport.ResolveMissingUnits;
end;

procedure TContextMenuInstaller.OnUsesCleanup(Sender: TObject);
begin
  CleanupUsesCurrentUnit;
end;

procedure TContextMenuInstaller.OnMoveToUnit(Sender: TObject);
begin
  if MoveToUnitInstance <> nil then
    MoveToUnitInstance.Execute;
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
