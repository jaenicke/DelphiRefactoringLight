(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
  TShortcutKind = (skRename, skCompletion, skExtract, skFindRef, skFindImp, skAlign, skRemoveWith, skUnitRefs, skMoveToUnit, skFindOriginal);

  TShortcutChangedProc = procedure of object;

  TExpertsShortCut = class
  strict private
    class var FShortcuts: array[TShortcutKind] of TShortCut;
    class var FListeners: TList<TMethod>;
    /// <summary>Tick of the last dispatch per kind - see AllowAction.</summary>
    class var FLastFired: array[TShortcutKind] of Cardinal;
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

    /// <summary>Debounce for actions reachable through TWO dispatch paths
    ///  for the same keystroke: the main-menu accelerator
    ///  (TMenuItem.ShortCut, which the VCL dispatches via
    ///  TCustomForm.IsShortCut) and our IOTAKeyboardBinding. Normally
    ///  only one of them ever sees the key - this makes "both fired"
    ///  harmless instead of running the action twice. Checks AND stamps.</summary>
    class function AllowAction(Kind: TShortcutKind): Boolean; static;

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
    class function scUnitRefs: TShortCut; static;
    class function scMoveToUnit: TShortCut; static;
    class function scFindOriginal: TShortCut; static;
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
    TShortCut(vkW     or scAlt or scCtrl or scShift),
    TShortCut(vkF     or scAlt or scCtrl or scShift),
    TShortCut(vkM     or scCtrl or scShift),
    // Deliberately NOT plain Ctrl+G (the PR's original choice): our
    // key binding captures the chord globally and would shadow the IDE's
    // own Ctrl+G (goto line). Rebind in Tools > Options if preferred.
    TShortCut(vkG     or scAlt or scCtrl or scShift)
  );

  ValueNames: array[TShortcutKind] of string = (
    'Rename', 'Completion', 'ExtractMethod',
    'FindReferences', 'FindImplementations', 'AlignSignature',
    'RemoveWith', 'UnitReferences', 'MoveToUnit', 'FindOriginalSymbol'
  );

  DisplayNames: array[TShortcutKind] of string = (
    'Rename',
    'Code completion',
    'Extract method',
    'Find references',
    'Find implementations',
    'Align method signature',
    'Remove with (project-wide)',
    'Find unit references (project-wide)',
    'Move to unit (project-wide)',
    'Find original symbol (go to declaration)'
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

class function TExpertsShortCut.AllowAction(Kind: TShortcutKind): Boolean;
const
  DebounceMs = 400;
var
  Now_: Cardinal;
begin
  Now_ := GetTickCount;
  Result := (FLastFired[Kind] = 0) or (Now_ - FLastFired[Kind] > DebounceMs);
  if Result then FLastFired[Kind] := Now_;
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

class function TExpertsShortCut.scUnitRefs: TShortCut;
begin
  Result := FShortcuts[skUnitRefs];
end;

class function TExpertsShortCut.scMoveToUnit: TShortCut;
begin
  Result := FShortcuts[skMoveToUnit];
end;

class function TExpertsShortCut.scFindOriginal: TShortCut;
begin
  Result := FShortcuts[skFindOriginal];
end;

end.
