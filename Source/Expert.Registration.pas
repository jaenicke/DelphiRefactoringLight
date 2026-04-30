(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.Registration;

interface

procedure Register;

implementation

uses
  System.SysUtils, ToolsAPI, Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod, Expert.FindReferencesWizard,
  Expert.FindImplementationsWizard, Expert.SignatureCheckWizard, Expert.KeyBinding, Expert.RestartHint, Expert.ContextMenu,
  Expert.UnitRenameWatcher;

var
  KeyBindingIndex: Integer = -1;

procedure Register;
var
  Services: IOTAKeyboardServices;
begin
  // Register the menu wizard (shows up under the Help menu)
  WizardInstance := TLspRenameWizard.Create;
  RegisterPackageWizard(WizardInstance);

  // Create the completion wizard
  CompletionWizardInstance := TLspCompletionWizard.Create;

  // Create the extract-method wizard
  ExtractMethodInstance := TLspExtractMethodWizard.Create;

  // Create the find-references wizard
  FindReferencesInstance := TLspFindReferencesWizard.Create;

  // Create the find-implementations wizard
  FindImplementationsInstance := TLspFindImplementationsWizard.Create;

  // Create the signature-check wizard
  SignatureCheckInstance := TLspSignatureCheckWizard.Create;

  // Register keyboard shortcuts
  // Ctrl+Alt+Shift+R     -> Rename
  // Ctrl+Alt+Shift+Space -> Completion
  // Ctrl+Alt+Shift+M     -> Extract Method
  // Ctrl+Alt+Shift+U     -> Find References (Usages)
  // Ctrl+Alt+Shift+I     -> Find Implementations
  // Ctrl+Alt+Shift+A     -> Align method signature
  if Supports(BorlandIDEServices, IOTAKeyboardServices, Services) then
    KeyBindingIndex := Services.AddKeyboardBinding(TLspKeyBinding.Create);

  // Install the "Refactoring Light" submenu into the editor context menu
  ContextMenuInstance := TContextMenuInstaller.Create;
  ContextMenuInstance.Install;

  // Unit rename watcher: offers rename when a unit is renamed in the IDE
  // (File > Save As etc.).
  UnitRenameWatcherInstance := TUnitRenameWatcher.Create;
  UnitRenameWatcherInstance.Install;

  // On manual (re-)install inside a running IDE: show the restart hint.
  TRestartHint.Check;
end;

initialization

finalization
  FreeAndNil(UnitRenameWatcherInstance);
  FreeAndNil(ContextMenuInstance);
  FreeAndNil(SignatureCheckInstance);
  FreeAndNil(FindImplementationsInstance);
  FreeAndNil(FindReferencesInstance);
  FreeAndNil(ExtractMethodInstance);
  FreeAndNil(CompletionWizardInstance);

  if KeyBindingIndex >= 0 then
  begin
    var Services: IOTAKeyboardServices;
    if Supports(BorlandIDEServices, IOTAKeyboardServices, Services) then
      Services.RemoveKeyboardBinding(KeyBindingIndex);
  end;
end.
