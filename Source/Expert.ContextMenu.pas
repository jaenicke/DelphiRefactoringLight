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
    // Editor-popup takeover: the IDE's own "Refactoring" popup entry we
    // hid for the current popup round (re-found and re-hidden per popup,
    // restored in RemoveOurItems / on uninstall).
    FPopupTakenOver: TMenuItem;
    FPopupDumped: Boolean;
    FPopupNoMatchDumped: Boolean;
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
      AOnClick: TNotifyEvent; AKind: TShortcutKind; ATrackShortcut: Boolean): TMenuItem;
    function BuildMenuTree(AOwner: TComponent; ATrackShortcut: Boolean): TMenuItem;
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
  Vcl.Forms, Vcl.Controls, Vcl.ActnList,
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
  REQ_LIVEFIX = 3;   // like REQ_EDITOR, but additionally reflects the live
                     // auto-import check: disabled when the checker KNOWS
                     // there is nothing to fix; annotated with the count
                     // when it knows there is.

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

  procedure Sep(AParent: TMenuItem);
  begin
    var Item := TMenuItem.Create(AOwner);
    Item.Caption := '-';
    AParent.Add(Item);          // no Req: separators are never dis/enabled
  end;

var
  Root, RemoveWithSub, IfaceSub, SemSub, ChecksSub: TMenuItem;
begin
  if (not ATrackShortcut) and (FItemReq = nil) then
    FItemReq := TDictionary<TMenuItem, Integer>.Create;

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

  IfaceSub := Sub(Root, 'Extract / extend interface');
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
  Plain(Root, 'Show all quick fixes...', OnShowQuickFixes,        REQ_LIVEALL);
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
      REQ_LIVEFIX:
        begin
          En := HasProject and HasEditor;
          // When the live auto-import check holds FRESH results for the
          // active buffer, reflect them: no missing units -> disabled.
          // (No count annotation here - the count is file-wide while this
          // action only works AT the cursor, which read as a broken
          // promise; the count lives on "Show all quick fixes" now.)
          var LiveCount: Integer;
          if En and LiveFreshInfo(Editor.GetActiveFileName, LiveCount) then
            En := LiveCount > 0;
        end;
      REQ_LIVEALL:
        begin
          En := HasProject and HasEditor;
          // File-wide overview: annotate with the TOTAL number of live
          // quick fixes. Fresh data with zero fixes -> disabled; without
          // fresh data the entry stays enabled (the action waits for the
          // next analysis itself).
          var AllFixes: TArray<TQuickFix>;
          var Cap := 'Show all quick fixes...';
          if En and LiveAllFixes(Editor.GetActiveFileName, AllFixes) then
          begin
            En := Length(AllFixes) > 0;
            if Length(AllFixes) > 0 then
              Cap := Format('Show all quick fixes... (%d found)', [Length(AllFixes)]);
          end;
          // Only write on an actual change - TMenuItem setters trigger
          // MenuChanged up to the menu bar, which repaints it.
          try
            if Pair.Key.Caption <> Cap then Pair.Key.Caption := Cap;
          except
          end;
        end;
    else           // REQ_PROJECT (and groups)
      En := HasProject;
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
