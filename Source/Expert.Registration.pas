(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
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
  System.SysUtils, ToolsAPI,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod,
  Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard, Expert.KeyBinding, Expert.RestartHint,
  Expert.ContextMenu, Expert.UnitRenameWatcher, Expert.WithRefactorWizard,
  Expert.Shortcuts, Expert.OptionsPage;

type
  TShortcutChangeHook = class
    class procedure HandleChanged;
  end;

class procedure TShortcutChangeHook.HandleChanged;
begin
  // Re-create the keyboard binding (the IDE caches its shortcut list)
  // and refresh the context-menu items.
  RebindKeyBinding;
  if ContextMenuInstance <> nil then
    ContextMenuInstance.RefreshShortcuts;
end;

procedure Register;
begin
  // Load configurable shortcut values from the registry first so all
  // installs below see the user's preferred bindings.
  TExpertsShortCut.LoadFromRegistry;
  TExpertsShortCut.AddListener(TShortcutChangeHook.HandleChanged);

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

  // Create the project-wide remove-with wizard
  WithRefactorInstance := TLspWithRefactorWizard.Create;

  // Register keyboard shortcuts (defaults are Ctrl+Alt+Shift + R/Space/M/U/I/A,
  // user-configurable via Tools > Options > Refactoring Light).
  InstallKeyBinding;

  // Install the "Refactoring Light" submenu into the editor context menu
  ContextMenuInstance := TContextMenuInstaller.Create;
  ContextMenuInstance.Install;

  // Add the options page (Tools > Options > Third Party > Refactoring Light)
  RegisterOptionsPage;

  // Unit rename watcher: offers rename when a unit is renamed in the IDE
  // (File > Save As etc.).
  UnitRenameWatcherInstance := TUnitRenameWatcher.Create;
  UnitRenameWatcherInstance.Install;

  // On manual (re-)install inside a running IDE: show the restart hint.
  TRestartHint.Check;
end;

initialization

finalization
  TExpertsShortCut.RemoveListener(TShortcutChangeHook.HandleChanged);
  UnregisterOptionsPage;
  FreeAndNil(UnitRenameWatcherInstance);
  FreeAndNil(ContextMenuInstance);
  FreeAndNil(WithRefactorInstance);
  FreeAndNil(SignatureCheckInstance);
  FreeAndNil(FindImplementationsInstance);
  FreeAndNil(FindReferencesInstance);
  FreeAndNil(ExtractMethodInstance);
  FreeAndNil(CompletionWizardInstance);
  UninstallKeyBinding;
end.
