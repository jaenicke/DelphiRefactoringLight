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
  System.Classes, Vcl.Menus, Vcl.ExtCtrls, Expert.Shortcuts;

type
  TContextMenuInstaller = class
  private
    FItems: TArray<TMenuItem>;
    FSeparator: TMenuItem;
    FPopupMenu: TPopupMenu;
    FOldOnPopup: TNotifyEvent;
    FHooked: Boolean;
    FRetryTimer: TTimer;
    FRetryCount: Integer;
    FSyncTimer: TTimer;
    FInPopup: Boolean;
    procedure OnRename(Sender: TObject);
    procedure OnFindReferences(Sender: TObject);
    procedure OnFindImplementations(Sender: TObject);
    procedure OnExtractMethod(Sender: TObject);
    procedure OnCompletion(Sender: TObject);
    procedure OnSignatureCheck(Sender: TObject);
    procedure OnRemoveWith(Sender: TObject);
    procedure OnRetryTimer(Sender: TObject);
    procedure OnSyncTimer(Sender: TObject);
    procedure DoOnPopup(Sender: TObject);
    function IsOurHandler(const AHandler: TNotifyEvent): Boolean;
    function FindEditorPopupMenu: TPopupMenu;
    function CreateItem(const ACaption: string; AKind: TShortcutKind;
      AOnClick: TNotifyEvent): TMenuItem;
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
  System.SysUtils, System.UITypes, Winapi.Windows,
  Vcl.Forms, Vcl.Controls,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod,
  Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard, Expert.WithRefactorWizard;

const
  /// <summary>Maximum retry attempts when the editor popup is not yet
  ///  available. With a 500 ms interval, 40 retries give up to 20 s.</summary>
  MaxRetries = 40;

  /// <summary>Period at which we verify our OnPopup hook is still in
  ///  place. Other plugins may install their own OnPopup later and
  ///  unintentionally displace ours.</summary>
  SyncIntervalMs = 2000;

{ TContextMenuInstaller }

destructor TContextMenuInstaller.Destroy;
begin
  Uninstall;
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

function TContextMenuInstaller.CreateItem(const ACaption: string;
  AKind: TShortcutKind; AOnClick: TNotifyEvent): TMenuItem;
begin
  // Owner = nil: we manage these items ourselves and free them in
  // Uninstall. They are not permanent children of FPopupMenu.
  Result := TMenuItem.Create(nil);
  Result.Caption := ACaption;
  // Tag stores the shortcut kind so RefreshShortcuts can find this item.
  // Encoded as Ord+1 so 0 means "no shortcut tracked".
  Result.Tag := Ord(AKind) + 1;
  Result.ShortCut := TExpertsShortCut.Shortcuts[AKind];
  Result.OnClick := AOnClick;
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

procedure TContextMenuInstaller.BuildItems;
var
  Submenu: TMenuItem;
begin
  if Length(FItems) > 0 then Exit;

  // The visible top-level entry: a submenu carrying all our actions.
  Submenu := TMenuItem.Create(nil);
  Submenu.Caption := 'Refactoring Light';
  Submenu.Tag := 0; // no shortcut tracked for the parent

  Submenu.Add(CreateItem('Rename...',                       skRename,     OnRename));
  Submenu.Add(CreateItem('Find References',                 skFindRef,    OnFindReferences));
  Submenu.Add(CreateItem('Find Implementations',            skFindImp,    OnFindImplementations));
  Submenu.Add(CreateItem('Extract Method',                  skExtract,    OnExtractMethod));
  Submenu.Add(CreateItem('Align method signature...',       skAlign,      OnSignatureCheck));
  Submenu.Add(CreateItem('Code Completion',                 skCompletion, OnCompletion));
  Submenu.Add(CreateItem('Remove with (project-wide)...',   skRemoveWith, OnRemoveWith));

  FSeparator := TMenuItem.Create(nil);
  FSeparator.Caption := '-';

  // FItems holds exactly the top-level items we add to FPopupMenu.Items
  // each time the popup opens (and remove again before the next open).
  FItems := [FSeparator, Submenu];
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
  // Re-entrancy guard: a forwarded OnPopup must not loop back into us.
  if FInPopup then
  begin
    if Assigned(FOldOnPopup) then
      FOldOnPopup(Sender);
    Exit;
  end;

  FInPopup := True;
  try
    // 1. Take our items out so the IDE's popup bookkeeping (which
    //    re-registers component names) sees a pristine menu.
    RemoveOurItems;

    // 2. Run the original handler the IDE installed.
    try
      if Assigned(FOldOnPopup) then
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
begin
  // Another plugin may have replaced OnPopup with its own handler
  // after we installed ours. Detect and re-hook on top.
  if (FPopupMenu = nil) or not FHooked then Exit;
  if not IsOurHandler(FPopupMenu.OnPopup) then
  begin
    FOldOnPopup := FPopupMenu.OnPopup;
    FPopupMenu.OnPopup := DoOnPopup;
  end;
end;

procedure TContextMenuInstaller.TryInstall;
begin
  if FHooked then Exit;

  FPopupMenu := FindEditorPopupMenu;
  if FPopupMenu = nil then Exit;

  BuildItems;
  HookPopup;

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
  Item: TMenuItem;
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

  // Free our items (Owner = nil, so we own them).
  for I := 0 to High(FItems) do
  begin
    Item := FItems[I];
    if Item = nil then Continue;
    try
      Item.Free;
    except
      // ignore
    end;
    FItems[I] := nil;
  end;
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

procedure TContextMenuInstaller.OnRemoveWith(Sender: TObject);
begin
  if WithRefactorInstance <> nil then
    WithRefactorInstance.Execute;
end;

end.
