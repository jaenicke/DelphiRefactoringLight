(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.KeyBinding;

interface

uses
  System.SysUtils, System.Classes, Winapi.Windows, Vcl.Menus, ToolsAPI;

type
  TLspKeyBinding = class(TNotifierObject, IOTAKeyboardBinding)
  private
    procedure RenameKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure CompletionKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure ExtractMethodKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure FindReferencesKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure FindImplementationsKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
    procedure SignatureCheckKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
  public
    function GetBindingType: TBindingType;
    function GetDisplayName: string;
    function GetName: string;
    procedure BindKeyboard(const BindingServices: IOTAKeyBindingServices);
  end;

implementation

uses
  Expert.Shortcuts,
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod, Expert.FindReferencesWizard, Expert.FindImplementationsWizard,
  Expert.SignatureCheckWizard;

{ TLspKeyBinding }

procedure TLspKeyBinding.RenameKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if WizardInstance <> nil then
    WizardInstance.Execute;
end;

procedure TLspKeyBinding.CompletionKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if CompletionWizardInstance <> nil then
    CompletionWizardInstance.Execute;
end;

procedure TLspKeyBinding.ExtractMethodKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if ExtractMethodInstance <> nil then
    ExtractMethodInstance.Execute;
end;

procedure TLspKeyBinding.FindReferencesKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if FindReferencesInstance <> nil then
    FindReferencesInstance.Execute;
end;

procedure TLspKeyBinding.FindImplementationsKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if FindImplementationsInstance <> nil then
    FindImplementationsInstance.Execute;
end;

procedure TLspKeyBinding.SignatureCheckKeyProc(const Context: IOTAKeyContext; KeyCode: TShortCut; var BindingResult: TKeyBindingResult);
begin
  BindingResult := krHandled;
  if SignatureCheckInstance <> nil then
    SignatureCheckInstance.Execute;
end;

function TLspKeyBinding.GetBindingType: TBindingType;
begin
  Result := btPartial;
end;

function TLspKeyBinding.GetDisplayName: string;
begin
  Result := 'Delphi Refactoring Light';
end;

function TLspKeyBinding.GetName: string;
begin
  Result := 'DelphiRefactoringLight.KeyBinding';
end;

procedure TLspKeyBinding.BindKeyboard(
  const BindingServices: IOTAKeyBindingServices);
begin
  // Ctrl+Alt+Shift+R -> Rename
  BindingServices.AddKeyBinding([TExpertsShortCut.scRename], RenameKeyProc, nil);

  // Ctrl+Alt+Shift+Space -> Completion
  BindingServices.AddKeyBinding([TExpertsShortCut.scCompletion], CompletionKeyProc, nil);

  // Ctrl+Alt+Shift+M -> Extract Method
  BindingServices.AddKeyBinding([TExpertsShortCut.scExtract], ExtractMethodKeyProc, nil);

  // Ctrl+Alt+Shift+U -> Find References (Usages)
  BindingServices.AddKeyBinding([TExpertsShortCut.scFindRef], FindReferencesKeyProc, nil);

  // Ctrl+Alt+Shift+I -> Find Implementations
  BindingServices.AddKeyBinding([TExpertsShortCut.scFindImp], FindImplementationsKeyProc, nil);

  // Ctrl+Alt+Shift+A -> Align method signature
  BindingServices.AddKeyBinding([TExpertsShortCut.scAlign], SignatureCheckKeyProc, nil);
end;

end.
