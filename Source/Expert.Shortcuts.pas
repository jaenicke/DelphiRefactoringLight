(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.Shortcuts;

{
  Central, user-configurable shortcut settings for all six refactoring
  features. Values are persisted under the IDE's base registry key
  (..\Embarcadero\BDS\<ver>\DelphiRefactoringLight\Shortcuts) so each
  Delphi version keeps its own bindings. The Tools > Options page
  "Refactoring Light" (see Expert.OptionsPage) edits these values; on
  Apply, registered listeners are notified so KeyBinding and the editor
  context menu can update themselves live.
}

interface

uses
  Winapi.Windows, System.UITypes, System.Classes, System.Generics.Collections,
  Vcl.Menus;

type
  TShortcutKind = (skRename, skCompletion, skExtract, skFindRef, skFindImp, skAlign, skRemoveWith);

  TShortcutChangedProc = procedure of object;

  TExpertsShortCut = class
  strict private
    class var FShortcuts: array[TShortcutKind] of TShortCut;
    class var FListeners: TList<TMethod>;
    class function GetShortcut(Kind: TShortcutKind): TShortCut; static;
    class procedure SetShortcut(Kind: TShortcutKind; Value: TShortCut); static;
    class function RegistryKey: string; static;
  public
    class constructor Create;
    class destructor Destroy;

    /// <summary>Default shortcut for a given kind (Ctrl+Alt+Shift+&lt;letter&gt;).</summary>
    class function Default(Kind: TShortcutKind): TShortCut; static;
    /// <summary>Human-readable name for the options page, e.g. "Rename".</summary>
    class function DisplayName(Kind: TShortcutKind): string; static;

    /// <summary>Reads all shortcut values from the registry. Missing
    ///  entries fall back to the default.</summary>
    class procedure LoadFromRegistry; static;
    /// <summary>Persists all shortcut values to the registry.</summary>
    class procedure SaveToRegistry; static;
    /// <summary>Restores the default values in memory (does not save).</summary>
    class procedure ResetToDefaults; static;
    /// <summary>Fires all registered change listeners. Call after edits
    ///  so KeyBinding and ContextMenu can refresh.</summary>
    class procedure NotifyChanged; static;

    class procedure AddListener(AProc: TShortcutChangedProc); static;
    class procedure RemoveListener(AProc: TShortcutChangedProc); static;

    class property Shortcuts[Kind: TShortcutKind]: TShortCut read GetShortcut
      write SetShortcut;

    // Backwards-compatible shorthand accessors used by existing call sites.
    class function scRename: TShortCut; static;
    class function scCompletion: TShortCut; static;
    class function scExtract: TShortCut; static;
    class function scFindRef: TShortCut; static;
    class function scFindImp: TShortCut; static;
    class function scAlign: TShortCut; static;
    class function scRemoveWith: TShortCut; static;
  end;

implementation

uses
  System.Win.Registry, System.SysUtils, ToolsAPI;

const
  DefaultShortcuts: array[TShortcutKind] of TShortCut = (
    TShortCut(vkR     or scAlt or scCtrl or scShift),
    TShortCut(vkSpace or scAlt or scCtrl or scShift),
    TShortCut(vkM     or scAlt or scCtrl or scShift),
    TShortCut(vkU     or scAlt or scCtrl or scShift),
    TShortCut(vkI     or scAlt or scCtrl or scShift),
    TShortCut(vkA     or scAlt or scCtrl or scShift),
    TShortCut(vkW     or scAlt or scCtrl or scShift)
  );

  ValueNames: array[TShortcutKind] of string = (
    'Rename', 'Completion', 'ExtractMethod',
    'FindReferences', 'FindImplementations', 'AlignSignature',
    'RemoveWith'
  );

  DisplayNames: array[TShortcutKind] of string = (
    'Rename',
    'Code completion',
    'Extract method',
    'Find references',
    'Find implementations',
    'Align method signature',
    'Remove with (project-wide)'
  );

{ TExpertsShortCut }

class constructor TExpertsShortCut.Create;
var
  K: TShortcutKind;
begin
  FListeners := TList<TMethod>.Create;
  for K := Low(TShortcutKind) to High(TShortcutKind) do
    FShortcuts[K] := DefaultShortcuts[K];
end;

class destructor TExpertsShortCut.Destroy;
begin
  FListeners.Free;
end;

class function TExpertsShortCut.Default(Kind: TShortcutKind): TShortCut;
begin
  Result := DefaultShortcuts[Kind];
end;

class function TExpertsShortCut.DisplayName(Kind: TShortcutKind): string;
begin
  Result := DisplayNames[Kind];
end;

class function TExpertsShortCut.GetShortcut(Kind: TShortcutKind): TShortCut;
begin
  Result := FShortcuts[Kind];
end;

class procedure TExpertsShortCut.SetShortcut(Kind: TShortcutKind;
  Value: TShortCut);
begin
  FShortcuts[Kind] := Value;
end;

class function TExpertsShortCut.RegistryKey: string;
var
  Services: IOTAServices;
  Base: string;
begin
  Base := 'Software\Embarcadero\BDS';
  if Supports(BorlandIDEServices, IOTAServices, Services) then
    Base := Services.GetBaseRegistryKey;
  // GetBaseRegistryKey may already be relative to HKCU; ensure no leading slash.
  while (Base <> '') and (Base[1] = '\') do
    Delete(Base, 1, 1);
  Result := Base + '\DelphiRefactoringLight\Shortcuts';
end;

class procedure TExpertsShortCut.LoadFromRegistry;
var
  Reg: TRegistry;
  K: TShortcutKind;
begin
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(RegistryKey) then
    try
      for K := Low(TShortcutKind) to High(TShortcutKind) do
        if Reg.ValueExists(ValueNames[K]) then
          FShortcuts[K] := TShortCut(Reg.ReadInteger(ValueNames[K]))
        else
          FShortcuts[K] := DefaultShortcuts[K];
    finally
      Reg.CloseKey;
    end
    else
    begin
      for K := Low(TShortcutKind) to High(TShortcutKind) do
        FShortcuts[K] := DefaultShortcuts[K];
    end;
  finally
    Reg.Free;
  end;
end;

class procedure TExpertsShortCut.SaveToRegistry;
var
  Reg: TRegistry;
  K: TShortcutKind;
begin
  Reg := TRegistry.Create(KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(RegistryKey, True) then
    try
      for K := Low(TShortcutKind) to High(TShortcutKind) do
        Reg.WriteInteger(ValueNames[K], Integer(FShortcuts[K]));
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

class procedure TExpertsShortCut.ResetToDefaults;
var
  K: TShortcutKind;
begin
  for K := Low(TShortcutKind) to High(TShortcutKind) do
    FShortcuts[K] := DefaultShortcuts[K];
end;

class procedure TExpertsShortCut.AddListener(AProc: TShortcutChangedProc);
begin
  FListeners.Add(TMethod(AProc));
end;

class procedure TExpertsShortCut.RemoveListener(AProc: TShortcutChangedProc);
var
  M: TMethod;
  I: Integer;
begin
  M := TMethod(AProc);
  for I := FListeners.Count - 1 downto 0 do
    if (FListeners[I].Code = M.Code) and (FListeners[I].Data = M.Data) then
      FListeners.Delete(I);
end;

class procedure TExpertsShortCut.NotifyChanged;
var
  I: Integer;
  Proc: TShortcutChangedProc;
begin
  for I := 0 to FListeners.Count - 1 do
  begin
    TMethod(Proc) := FListeners[I];
    if Assigned(Proc) then
      try
        Proc;
      except
        // listener errors must not break the options dialog
      end;
  end;
end;

class function TExpertsShortCut.scRename: TShortCut;
begin
  Result := FShortcuts[skRename];
end;

class function TExpertsShortCut.scCompletion: TShortCut;
begin
  Result := FShortcuts[skCompletion];
end;

class function TExpertsShortCut.scExtract: TShortCut;
begin
  Result := FShortcuts[skExtract];
end;

class function TExpertsShortCut.scFindRef: TShortCut;
begin
  Result := FShortcuts[skFindRef];
end;

class function TExpertsShortCut.scFindImp: TShortCut;
begin
  Result := FShortcuts[skFindImp];
end;

class function TExpertsShortCut.scAlign: TShortCut;
begin
  Result := FShortcuts[skAlign];
end;

class function TExpertsShortCut.scRemoveWith: TShortCut;
begin
  Result := FShortcuts[skRemoveWith];
end;

end.
