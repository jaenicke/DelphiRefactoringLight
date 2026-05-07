(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.OptionsPage;

{
  Registers a "Refactoring Light" page under Tools > Options > Third
  Party. The page hosts TLspOptionsFrame (Expert.OptionsFrame). When the
  user closes the dialog with OK, the new shortcut values are persisted
  and listeners (KeyBinding, ContextMenu) are notified to rebind/refresh.
}

interface

uses
  System.Classes, Vcl.Forms, ToolsAPI;

type
  TLspOptionsAddIn = class(TInterfacedObject, INTAAddInOptions)
  private
    FFrame: TCustomFrame;
  public
    // INTAAddInOptions
    function GetArea: string;
    function GetCaption: string;
    function GetFrameClass: TCustomFrameClass;
    procedure FrameCreated(AFrame: TCustomFrame);
    procedure DialogClosed(Accepted: Boolean);
    function ValidateContents: Boolean;
    function GetHelpContext: Integer;
    function IncludeInIDEInsight: Boolean;
  end;

procedure RegisterOptionsPage;
procedure UnregisterOptionsPage;

implementation

uses
  Winapi.Windows, System.SysUtils,
  Expert.OptionsFrame, Expert.Shortcuts;

var
  OptionsAddIn: INTAAddInOptions;
  RegistrationAttempted: Boolean = False;

procedure TraceRegistration(const Msg: string);
begin
  // Visible in DebugView / IDE Event Log so the user can confirm whether
  // the options page registered successfully.
  OutputDebugString(PChar('[DelphiRefactoringLight] ' + Msg));
end;

{ TLspOptionsAddIn }

function TLspOptionsAddIn.GetArea: string;
begin
  // Empty string places the page under the default "Third Party" group.
  Result := '';
end;

function TLspOptionsAddIn.GetCaption: string;
begin
  Result := 'Refactoring Light';
end;

function TLspOptionsAddIn.GetFrameClass: TCustomFrameClass;
begin
  Result := TLspOptionsFrame;
end;

procedure TLspOptionsAddIn.FrameCreated(AFrame: TCustomFrame);
begin
  FFrame := AFrame;
  if AFrame is TLspOptionsFrame then
    TLspOptionsFrame(AFrame).LoadFromSettings;
end;

procedure TLspOptionsAddIn.DialogClosed(Accepted: Boolean);
begin
  if Accepted and (FFrame is TLspOptionsFrame) then
  begin
    TLspOptionsFrame(FFrame).StoreToSettings;
    TExpertsShortCut.SaveToRegistry;
    TExpertsShortCut.NotifyChanged;
  end;
  FFrame := nil;
end;

function TLspOptionsAddIn.ValidateContents: Boolean;
begin
  Result := True;
end;

function TLspOptionsAddIn.GetHelpContext: Integer;
begin
  Result := 0;
end;

function TLspOptionsAddIn.IncludeInIDEInsight: Boolean;
begin
  Result := True;
end;

{ Registration }

procedure DoRegister;
var
  Services: INTAEnvironmentOptionsServices;
begin
  if OptionsAddIn <> nil then Exit;
  if Supports(BorlandIDEServices, INTAEnvironmentOptionsServices, Services) then
  begin
    OptionsAddIn := TLspOptionsAddIn.Create;
    Services.RegisterAddInOptions(OptionsAddIn);
    TraceRegistration('Options page registered (Tools > Options > Third Party > Refactoring Light)');
  end
  else
    TraceRegistration('INTAEnvironmentOptionsServices not available yet');
end;

procedure RegisterOptionsPage;
begin
  if RegistrationAttempted then Exit;
  RegistrationAttempted := True;

  // Try immediately. If the service interface isn't available yet (can
  // happen when other addins are still initialising), retry once after
  // the IDE has finished its startup tick.
  DoRegister;
  if OptionsAddIn = nil then
    TThread.ForceQueue(nil,
      procedure
      begin
        DoRegister;
        if OptionsAddIn = nil then
          TraceRegistration('Options page registration failed - INTAEnvironmentOptionsServices missing');
      end);
end;

procedure UnregisterOptionsPage;
var
  Services: INTAEnvironmentOptionsServices;
begin
  RegistrationAttempted := False;
  if OptionsAddIn = nil then Exit;
  if Supports(BorlandIDEServices, INTAEnvironmentOptionsServices, Services) then
  try
    Services.UnregisterAddInOptions(OptionsAddIn);
  except
    // ignore - IDE may already be shutting down
  end;
  OptionsAddIn := nil;
end;

end.
