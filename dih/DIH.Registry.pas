(*
 * Copyright (c) 2026 Sebastian J�nicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Registry;

interface

uses
  System.SysUtils, System.Win.Registry, Winapi.Windows,
  DIH.Types, DIH.Logger, DIH.Placeholders;

type
  TDIHRegistryManager = class
  private
    FLogger: TDIHLogger;
    FResolver: TDIHPlaceholderResolver;
  public
    constructor Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
    procedure SetValues(const AValues: TArray<TDIHRegistryValue>);
    procedure DeleteValues(const AValues: TArray<TDIHRegistryValue>);
  end;

implementation

{ TDIHRegistryManager }

constructor TDIHRegistryManager.Create(ALogger: TDIHLogger; AResolver: TDIHPlaceholderResolver);
begin
  inherited Create;
  FLogger := ALogger;
  FResolver := AResolver;
end;

procedure TDIHRegistryManager.SetValues(const AValues: TArray<TDIHRegistryValue>);
var
  Reg: TRegistry;
  RegVal: TDIHRegistryValue;
  Path: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    for RegVal in AValues do
    begin
      Path := FResolver.Resolve(RegVal.Path);
      if Reg.OpenKey(Path, True) then
      begin
        try
          if SameText(RegVal.ValueType, 'Integer') then
            Reg.WriteInteger(RegVal.Name, StrToIntDef(RegVal.Value, 0))
          else
            Reg.WriteString(RegVal.Name, RegVal.Value);
          FLogger.Detail('Registry: %s\%s = %s', [Path, RegVal.Name, RegVal.Value]);
        finally
          Reg.CloseKey;
        end;
      end
      else
        FLogger.Error('Failed to open registry key: %s', [Path]);
    end;
  finally
    Reg.Free;
  end;
end;

procedure TDIHRegistryManager.DeleteValues(const AValues: TArray<TDIHRegistryValue>);
var
  Reg: TRegistry;
  RegVal: TDIHRegistryValue;
  Path: string;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    for RegVal in AValues do
    begin
      Path := FResolver.Resolve(RegVal.Path);
      if Reg.OpenKey(Path, False) then
      begin
        try
          if Reg.ValueExists(RegVal.Name) then
          begin
            Reg.DeleteValue(RegVal.Name);
            FLogger.Detail('Registry: Deleted %s\%s', [Path, RegVal.Name]);
          end;
        finally
          Reg.CloseKey;
        end;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

end.
