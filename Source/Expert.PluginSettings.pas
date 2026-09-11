(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
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
    class var FLiveBlame: Boolean;
    class var FBlameColumnWidth: Integer;
    class var FBlameInfo: Integer;
    class var FBlameColumnOffset: Integer;
    class var FBlameUseTortoise: Boolean;
    class var FScopeIncludeOpenUnits: Boolean;
    class var FScopeIncludeUsedUnits: Boolean;
    class var FLoaded: Boolean;
    class function BaseRegistryKey: string; static;
    class function RegistryKey: string; static;
    class function LegacyRegistryKey: string; static;
    class procedure MigrateLegacyKey; static;
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

    /// <summary>Live blame in the editor gutter. OFF by default: it runs
    ///  "git blame" per file, and not every project is a git working
    ///  copy.</summary>
    class property LiveBlame: Boolean read FLiveBlame write FLiveBlame;

    /// <summary>Pixels the editor gutter is WIDENED by while live blame is
    ///  on. 0 = do not touch the gutter at all (only the age stripe and
    ///  the caret-line annotation) - the polite setting when another
    ///  add-on already uses that space.</summary>
    class property BlameColumnWidth: Integer
      read FBlameColumnWidth write FBlameColumnWidth;

    /// <summary>What the gutter column shows: 0 = revision, 1 = revision +
    ///  author, 2 = revision + author + age.</summary>
    class property BlameInfo: Integer read FBlameInfo write FBlameInfo;

    /// <summary>Pixels between the start of the gutter's data area and OUR
    ///  column. The IDE does not arbitrate gutter space, and other add-ons
    ///  draw there too (the bundled Parnassus Navigator puts its marks at
    ///  the very left), so the position has to be adjustable.</summary>
    class property BlameColumnOffset: Integer
      read FBlameColumnOffset write FBlameColumnOffset;

    /// <summary>Use TortoiseGit / TortoiseSVN for the commit and blame
    ///  views when they are installed - their windows are richer than
    ///  anything we would rebuild, and they are what most people here
    ///  already know. Falls back to the built-in views when the client is
    ///  missing or refuses to start.</summary>
    class property BlameUseTortoise: Boolean
      read FBlameUseTortoise write FBlameUseTortoise;

    /// <summary>Project-wide scans (rename, find references, find
    ///  implementations, find unit references) also look at units that
    ///  are OPEN in the editor but not part of the project. The caret's
    ///  own unit is always included, regardless of this switch.</summary>
    class property ScopeIncludeOpenUnits: Boolean
      read FScopeIncludeOpenUnits write FScopeIncludeOpenUnits;
    /// <summary>... and units reachable through uses clauses (resolved via
    ///  the identifier index, never below the RAD Studio installation).
    ///  Off by default: with large libraries this can be thousands of
    ///  files.</summary>
    class property ScopeIncludeUsedUnits: Boolean
      read FScopeIncludeUsedUnits write FScopeIncludeUsedUnits;

    class function DefaultPrewarm: Boolean; static;
  end;

implementation

uses
  System.SysUtils, System.Win.Registry, Winapi.Windows
  {$IFNDEF STANDALONE_BUILD}, ToolsAPI{$ENDIF};

{ TPluginSettings }

class function TPluginSettings.BaseRegistryKey: string;
var
  {$IFNDEF STANDALONE_BUILD}
  Services: IOTAServices;
  {$ENDIF}
  BaseKey: string;
begin
  BaseKey := '';
  // The standalone exe has no IDE services and reads the same branch the
  // IDE plugin uses, so both share one configuration.
  {$IFNDEF STANDALONE_BUILD}
  if Supports(BorlandIDEServices, IOTAServices, Services) then
  try
    BaseKey := Services.GetBaseRegistryKey;
  except
    BaseKey := '';
  end;
  {$ENDIF}
  if BaseKey = '' then
    BaseKey := 'Software\Embarcadero\BDS\37.0';
  // GetBaseRegistryKey can come back with a leading slash - same
  // normalisation TExpertsShortCut does.
  while (BaseKey <> '') and (BaseKey[1] = '\') do
    Delete(BaseKey, 1, 1);
  Result := BaseKey;
end;

// ONE branch for the whole plugin. The settings used to live under
// '...\RefactoringLight\Settings' while the shortcuts were (and are)
// under '...\DelphiRefactoringLight\Shortcuts' - two keys for one
// plugin, which a tester rightly called out. The PACKAGE name wins,
// and MigrateLegacyKey moves an existing configuration over exactly
// once, so nobody has to set their options up again.
class function TPluginSettings.RegistryKey: string;
begin
  Result := BaseRegistryKey + '\DelphiRefactoringLight\Settings';
end;

class function TPluginSettings.LegacyRegistryKey: string;
begin
  Result := BaseRegistryKey + '\RefactoringLight\Settings';
end;

class procedure TPluginSettings.MigrateLegacyKey;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    // Only when there is nothing new yet AND something old to take over.
    if Reg.KeyExists(RegistryKey) then Exit;
    if not Reg.KeyExists(LegacyRegistryKey) then Exit;
    if Reg.OpenKeyReadOnly(LegacyRegistryKey) then
    try
      if Reg.ValueExists('PrewarmLspOnProjectOpen') then
        FPrewarmLspOnProjectOpen := Reg.ReadBool('PrewarmLspOnProjectOpen');
      if Reg.ValueExists('LiveBlame') then
        FLiveBlame := Reg.ReadBool('LiveBlame');
      if Reg.ValueExists('BlameColumnWidth') then
        FBlameColumnWidth := Reg.ReadInteger('BlameColumnWidth');
      if Reg.ValueExists('BlameInfo') then
        FBlameInfo := Reg.ReadInteger('BlameInfo');
      if Reg.ValueExists('BlameColumnOffset') then
        FBlameColumnOffset := Reg.ReadInteger('BlameColumnOffset');
      if Reg.ValueExists('BlameUseTortoise') then
        FBlameUseTortoise := Reg.ReadBool('BlameUseTortoise');
    finally
      Reg.CloseKey;
    end;
    Save;                        // write everything to the new place
    // Then take the old branch away: it is ours, now redundant, and
    // leaving it invites the next 'which one is real?' question.
    try
      Reg.DeleteKey(LegacyRegistryKey);
      Reg.DeleteKey(BaseRegistryKey + '\RefactoringLight');   // if empty
    except
      // a leftover key is harmless - never fail a load over it
    end;
  finally
    Reg.Free;
  end;
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
  FLiveBlame := False;
  FBlameColumnWidth := 150;
  FBlameInfo := 1;
  // 17 px: measured against the Parnassus Navigator, which ships WITH
  // Delphi now and draws its marks at the very left of the same area.
  FBlameColumnOffset := 17;
  FBlameUseTortoise := True;
  FScopeIncludeOpenUnits := True;
  FScopeIncludeUsedUnits := False;
  FLoaded := True;

  MigrateLegacyKey;

  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(RegistryKey) then
    try
      if Reg.ValueExists('PrewarmLspOnProjectOpen') then
        FPrewarmLspOnProjectOpen := Reg.ReadBool('PrewarmLspOnProjectOpen');
      if Reg.ValueExists('LiveBlame') then
        FLiveBlame := Reg.ReadBool('LiveBlame');
      if Reg.ValueExists('BlameColumnWidth') then
        FBlameColumnWidth := Reg.ReadInteger('BlameColumnWidth');
      if Reg.ValueExists('BlameInfo') then
        FBlameInfo := Reg.ReadInteger('BlameInfo');
      if Reg.ValueExists('BlameColumnOffset') then
        FBlameColumnOffset := Reg.ReadInteger('BlameColumnOffset');
      if Reg.ValueExists('BlameUseTortoise') then
        FBlameUseTortoise := Reg.ReadBool('BlameUseTortoise');
      if Reg.ValueExists('ScopeIncludeOpenUnits') then
        FScopeIncludeOpenUnits := Reg.ReadBool('ScopeIncludeOpenUnits');
      if Reg.ValueExists('ScopeIncludeUsedUnits') then
        FScopeIncludeUsedUnits := Reg.ReadBool('ScopeIncludeUsedUnits');
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
      Reg.WriteBool('LiveBlame', FLiveBlame);
      Reg.WriteInteger('BlameColumnWidth', FBlameColumnWidth);
      Reg.WriteInteger('BlameInfo', FBlameInfo);
      Reg.WriteInteger('BlameColumnOffset', FBlameColumnOffset);
      Reg.WriteBool('BlameUseTortoise', FBlameUseTortoise);
      Reg.WriteBool('ScopeIncludeOpenUnits', FScopeIncludeOpenUnits);
      Reg.WriteBool('ScopeIncludeUsedUnits', FScopeIncludeUsedUnits);
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
