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
  public
    function GetBindingType: TBindingType;
    function GetDisplayName: string;
    function GetName: string;
    procedure BindKeyboard(const BindingServices: IOTAKeyBindingServices);
  end;

implementation

uses
  Expert.RenameWizard, Expert.CompletionWizard, Expert.ExtractMethod, Expert.FindReferencesWizard, Expert.FindImplementationsWizard;

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
  // Ctrl+Shift+R -> Rename
  BindingServices.AddKeyBinding([ShortCut(Ord('R'), [ssCtrl, ssShift])], RenameKeyProc, nil);

  // Ctrl+Shift+Space -> Completion
  BindingServices.AddKeyBinding([ShortCut(VK_SPACE, [ssCtrl, ssShift])], CompletionKeyProc, nil);

  // Ctrl+Shift+M -> Extract Method
  BindingServices.AddKeyBinding([ShortCut(Ord('M'), [ssCtrl, ssShift])], ExtractMethodKeyProc, nil);

  // Ctrl+Shift+U -> Find References (Usages)
  BindingServices.AddKeyBinding([ShortCut(Ord('U'), [ssCtrl, ssShift])], FindReferencesKeyProc, nil);

  // Ctrl+Shift+I -> Find Implementations
  BindingServices.AddKeyBinding([ShortCut(Ord('I'), [ssCtrl, ssShift])], FindImplementationsKeyProc, nil);
end;

end.
