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

  The IDE exposes this popup as a VCL TPopupMenu component on one of the
  running forms (typically named "EditorLocalMenu"). We locate it via
  Screen.CustomForms at package load time and append our submenu.

  Because the popup may not exist yet when the package's Register
  procedure runs (the editor is not always initialized when design-time
  packages load), the install attempt is retried via a TTimer until the
  menu is found or a timeout is reached.
}

interface

uses
  System.Classes, Vcl.Menus, Vcl.ExtCtrls, Expert.Shortcuts;

type
  TContextMenuInstaller = class
  private
    FSubmenu: TMenuItem;
    FSeparator: TMenuItem;
    FPopupMenu: TPopupMenu;
    FRetryTimer: TTimer;
    FRetryCount: Integer;
    procedure OnRename(Sender: TObject);
    procedure OnFindReferences(Sender: TObject);
    procedure OnFindImplementations(Sender: TObject);
    procedure OnExtractMethod(Sender: TObject);
    procedure OnCompletion(Sender: TObject);
    procedure OnSignatureCheck(Sender: TObject);
    procedure OnRemoveWith(Sender: TObject);
    procedure OnRetryTimer(Sender: TObject);
    function FindEditorPopupMenu: TPopupMenu;
    function CreateItem(AParent: TMenuItem; const ACaption: string;
      AKind: TShortcutKind; AOnClick: TNotifyEvent): TMenuItem;
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

function TContextMenuInstaller.CreateItem(AParent: TMenuItem;
  const ACaption: string; AKind: TShortcutKind;
  AOnClick: TNotifyEvent): TMenuItem;
begin
  // Owner = nil so the item does NOT pollute the IDE's component
  // bookkeeping (EditorLocalMenu.Components / its Owner.Components).
  // Some IDE features iterate that list per popup-open and choke when
  // they see foreign components in there. We free our items manually
  // in Uninstall.
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
  I: Integer;
  Item: TMenuItem;
  Kind: TShortcutKind;
begin
  if FSubmenu = nil then Exit;
  for I := 0 to FSubmenu.Count - 1 do
  begin
    Item := FSubmenu.Items[I];
    if (Item.Tag >= 1) and (Item.Tag <= Ord(High(TShortcutKind)) + 1) then
    begin
      Kind := TShortcutKind(Item.Tag - 1);
      Item.ShortCut := TExpertsShortCut.Shortcuts[Kind];
    end;
  end;
end;

procedure TContextMenuInstaller.Install;
begin
  if FSubmenu <> nil then Exit; // already installed

  // Try once immediately; if the popup is not yet available, start a
  // timer that keeps retrying until it shows up.
  TryInstall;
  if FSubmenu <> nil then Exit;

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
  if (FSubmenu <> nil) or (FRetryCount >= MaxRetries) then
    FRetryTimer.Enabled := False
  else
    Inc(FRetryCount);
end;

procedure TContextMenuInstaller.TryInstall;
begin
  if FSubmenu <> nil then Exit;

  FPopupMenu := FindEditorPopupMenu;
  if FPopupMenu = nil then Exit;

  // Build our submenu. Owner = nil — see CreateItem for why.
  FSubmenu := TMenuItem.Create(nil);
  FSubmenu.Caption := 'Refactoring Light';

  FSubmenu.Add(CreateItem(FSubmenu, 'Rename...',
    skRename, OnRename));
  FSubmenu.Add(CreateItem(FSubmenu, 'Find References',
    skFindRef, OnFindReferences));
  FSubmenu.Add(CreateItem(FSubmenu, 'Find Implementations',
    skFindImp, OnFindImplementations));
  FSubmenu.Add(CreateItem(FSubmenu, 'Extract Method',
    skExtract, OnExtractMethod));
  FSubmenu.Add(CreateItem(FSubmenu, 'Align method signature...',
    skAlign, OnSignatureCheck));
  FSubmenu.Add(CreateItem(FSubmenu, 'Code Completion',
    skCompletion, OnCompletion));
  FSubmenu.Add(CreateItem(FSubmenu, 'Remove with (project-wide)...',
    skRemoveWith, OnRemoveWith));

  // Separator above our submenu so it is visually grouped.
  // Owner = nil — see CreateItem for why.
  FSeparator := TMenuItem.Create(nil);
  FSeparator.Caption := '-';

  FPopupMenu.Items.Add(FSeparator);
  FPopupMenu.Items.Add(FSubmenu);
end;

procedure TContextMenuInstaller.Uninstall;
begin
  if FRetryTimer <> nil then
  begin
    FRetryTimer.Enabled := False;
    FreeAndNil(FRetryTimer);
  end;

  // During IDE shutdown the editor popup may have been cleared or torn
  // down before our finalization runs. Items.Remove then raises
  // "Submenu is not in menu" (EMenuError). Treat any failure as
  // "already gone" - the IDE is destroying everything anyway.
  if FPopupMenu <> nil then
  begin
    if FSubmenu <> nil then
    begin
      try
        if FPopupMenu.Items.IndexOf(FSubmenu) >= 0 then
          FPopupMenu.Items.Remove(FSubmenu);
      except
        // popup already gone
      end;
      try
        FreeAndNil(FSubmenu);
      except
        FSubmenu := nil;
      end;
    end;
    if FSeparator <> nil then
    begin
      try
        if FPopupMenu.Items.IndexOf(FSeparator) >= 0 then
          FPopupMenu.Items.Remove(FSeparator);
      except
        // popup already gone
      end;
      try
        FreeAndNil(FSeparator);
      except
        FSeparator := nil;
      end;
    end;
  end
  else
  begin
    // Popup was never hooked; just release our orphan items.
    try FreeAndNil(FSubmenu);   except FSubmenu := nil;   end;
    try FreeAndNil(FSeparator); except FSeparator := nil; end;
  end;

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
