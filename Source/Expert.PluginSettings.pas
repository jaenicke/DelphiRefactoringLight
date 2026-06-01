(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.PluginSettings;

{
  Lightweight registry-backed settings (non-shortcut) for the plugin.

  Currently holds:
    PrewarmLspOnProjectOpen  - when True, the plugin starts its own
                               DelphiLSP instance and pre-indexes the
                               just-opened project in the background.
                               First refactoring action is then
                               instant. Default True.

  Stored under the IDE's base registry key,
    sub-key 'RefactoringLight\Settings'.
}

interface

type
  TPluginSettings = class
  strict private
    class var FPrewarmLspOnProjectOpen: Boolean;
    class var FLoaded: Boolean;
    class function RegistryKey: string; static;
  public
    /// <summary>Reads the settings from the registry. Called automatically
    ///  by the property getters on first access; can be called manually to
    ///  re-read after an external change.</summary>
    class procedure Load; static;
    /// <summary>Writes the current values back to the registry.</summary>
    class procedure Save; static;

    /// <summary>When True, the plugin starts its own DelphiLSP and
    ///  pre-indexes the project as soon as it is opened, so the first
    ///  refactoring action doesn't have to wait for the LSP cold-start.
    ///  Costs ~one extra DelphiLSP process while the IDE is open.</summary>
    class property PrewarmLspOnProjectOpen: Boolean
      read FPrewarmLspOnProjectOpen write FPrewarmLspOnProjectOpen;

    class function DefaultPrewarm: Boolean; static;
  end;

implementation

uses
  System.SysUtils, System.Win.Registry, Winapi.Windows, ToolsAPI;

{ TPluginSettings }

class function TPluginSettings.RegistryKey: string;
var
  Services: IOTAServices;
  BaseKey: string;
begin
  BaseKey := '';
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  try
    BaseKey := Services.GetBaseRegistryKey;
  except
    BaseKey := '';
  end;
  if BaseKey = '' then
    BaseKey := 'Software\Embarcadero\BDS\37.0';
  Result := BaseKey + '\RefactoringLight\Settings';
end;

class function TPluginSettings.DefaultPrewarm: Boolean;
begin
  Result := True;
end;

class procedure TPluginSettings.Load;
var
  Reg: TRegistry;
begin
  FPrewarmLspOnProjectOpen := DefaultPrewarm;
  FLoaded := True;

  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(RegistryKey) then
    try
      if Reg.ValueExists('PrewarmLspOnProjectOpen') then
        FPrewarmLspOnProjectOpen := Reg.ReadBool('PrewarmLspOnProjectOpen');
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

class procedure TPluginSettings.Save;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(RegistryKey, True) then
    try
      Reg.WriteBool('PrewarmLspOnProjectOpen', FPrewarmLspOnProjectOpen);
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

initialization
  TPluginSettings.Load;

end.
