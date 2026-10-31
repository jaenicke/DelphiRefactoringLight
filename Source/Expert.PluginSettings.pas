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
  /// <summary>What kind of value a setting holds - all a caller needs to
  ///  know to write one as text.</summary>
  TPluginSettingKind = (pskBool, pskInt, pskString);

  /// <summary>One setting, described well enough to be listed, validated
  ///  and written by NAME. The same array drives Save(), so a setting
  ///  cannot be in one and missing from the other - the shape that made
  ///  the settings migration lose five values (audit #37, L4e).</summary>
  TPluginSettingInfo = record
    Name: string;        // the registry value name, and the tool's name
    Kind: TPluginSettingKind;
    DefaultText: string;
    Min, Max: Integer;   // pskInt only (Min = Max = 0: no range)
    Choices: string;     // pskInt with named values: "0 = revision, ..."
    Summary: string;     // one line - what it does
  end;

  TPluginSettings = class
  strict private
    class var FPrewarmLspOnProjectOpen: Boolean;
    class var FLiveBlame: Boolean;
    class var FBlameColumnWidth: Integer;
    class var FBlameInfo: Integer;
    class var FBlameColumnOffset: Integer;
    class var FBlameUseTortoise: Boolean;
    class var FIntfVisualizerViewer: Boolean;
    class var FIntfVisualizerLine: Boolean;
    class var FLspLogging: Boolean;
    class var FVerifySession: Boolean;
    class var FScopeIncludeOpenUnits: Boolean;
    class var FScopeIncludeUsedUnits: Boolean;
    class var FRenameScope: Integer;
    class var FRenameBackup: Boolean;
    class var FUsesCleanupKeepUnits: string;
    class var FLineEndingCheck: Boolean;
    class var FLineEndingAutoFix: Boolean;
    class var FLoaded: Boolean;
    // Session-only overrides: the names, and the text the registry
    // still holds for them (see MarkSessionOnly).
    class var FSessionNames: TArray<string>;
    class var FSessionValues: TArray<string>;
    class function ReadPersistedText(
      const AInfo: TPluginSettingInfo): string; static;
    class function LegacyRegistryKey: string; static;
    class procedure MigrateLegacyKey; static;
  public
    /// <summary>Our own registry branch under the key of the IDE this
    ///  plugin runs in (the IDE's own version key plus our branch). Public
    ///  because it is the ONE place that knows WHICH Delphi is hosting
    ///  us - Expert.UnitIndex derives the version key from it instead of
    ///  guessing the newest installed one (audit #37, L3c).</summary>
    class function BaseRegistryKey: string; static;
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
    /// <summary>Offer the interface visualizer in the debugger: the window
    ///  that resolves an interface reference to its object and lists the
    ///  class hierarchy with the interfaces each level implements. On by
    ///  default - it is opt-in per use (the debugger offers it in its own
    ///  list for the type) and it writes nothing.</summary>
    class property IntfVisualizerViewer: Boolean
      read FIntfVisualizerViewer write FIntfVisualizerViewer;
    /// <summary>...and replace the VALUE SHOWN for an interface reference
    ///  in every tooltip, watch, local view and inspector with one line
    ///  naming the class behind it. OFF by default: the IDE allows only
    ///  one such replacer per type, so this takes over the display of
    ///  every interface variable in the debugger.</summary>
    class property IntfVisualizerLine: Boolean
      read FIntfVisualizerLine write FIntfVisualizerLine;
    /// <summary>Starts our own DelphiLsp.exe with '-LogModes 255', so
    ///  %TEMP%\DelphiLSP\RefactoringLight*.log records every request and
    ///  its duration. Off by default (the log grows); asked for in issue
    ///  #13, because without it a slow session cannot be diagnosed.</summary>
    class property LspLogging: Boolean read FLspLogging write FLspLogging;
    /// <summary>Verify candidates through a SECOND DelphiLsp session in
    ///  'agent' mode: measured twice as fast and immune to the controller's
    ///  10 s abort, at the price of one more process (~290 MB) that is shut
    ///  down again after 10 idle minutes. See issue #13.</summary>
    class property VerifySession: Boolean read FVerifySession write FVerifySession;
    /// <summary>Where these settings live - for the status window and the
    ///  tests.</summary>
    class function RegistryKey: string; static;

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
    /// <summary>The rename dialog opens with the choices of the previous
    ///  rename (issue: renaming several things in one unit meant switching
    ///  the scope and unticking the backup EVERY time). RenameScope is the
    ///  combo index (0 = whole project .. 3 = selected units).</summary>
    class property RenameScope: Integer read FRenameScope write FRenameScope;
    class property RenameBackup: Boolean read FRenameBackup write FRenameBackup;

    /// <summary>Units the uses cleanup must never offer for removal, as
    /// masks separated by ';' ('dxSkin*;MyCompany.Components.*'). The
    /// designer query (issue #20) covers what ISelectionEditor.RequiresUnits
    /// reports; a unit some other IDE expert writes on save is not in that
    /// answer, and this is where such a false positive is turned off. Empty
    /// by default - a mask like 'cx*' would also protect cx units that
    /// really are unused.</summary>
    class property UsesCleanupKeepUnits: string
      read FUsesCleanupKeepUnits write FUsesCleanupKeepUnits;

    /// <summary>Check a unit's line endings when the IDE opens it and warn
    ///  when they are not plain CRLF. Delphi copes badly with LF-only and
    ///  worst with MIXED endings (breakpoints on the wrong line, a
    ///  half-converted file after a checkout), and the warning is the only
    ///  moment at which nobody has edited the buffer yet. On by default:
    ///  it reads the file the IDE has just read anyway.</summary>
    class property LineEndingCheck: Boolean
      read FLineEndingCheck write FLineEndingCheck;
    /// <summary>...and fix it without asking. Off by default - rewriting
    ///  every line break of a file is a change the user should have said
    ///  yes to once. The dialog's "always" option is what turns it on.</summary>
    class property LineEndingAutoFix: Boolean
      read FLineEndingAutoFix write FLineEndingAutoFix;

    /// <summary>Every setting, in the order the options page shows
    ///  them. The ONE list: Save writes exactly these.</summary>
    class function SettingInfos: TArray<TPluginSettingInfo>; static;
    /// <summary>The description of one setting, by name (case-insensitive).
    ///  </summary>
    class function FindSetting(const AName: string;
      out AInfo: TPluginSettingInfo): Boolean; static;
    /// <summary>The current value of one setting as text ('true' / 'false'
    ///  for a switch), '' plus False for an unknown name.</summary>
    class function GetValueText(const AName: string;
      out AText: string): Boolean; static;
    /// <summary>Writes one setting from text, in MEMORY - it takes effect
    ///  at once and the caller decides whether Save() makes it permanent.
    ///  Rejects an unknown name, a value that is not of the right kind and
    ///  an integer outside its range, each with a reason.</summary>
    class function SetValueText(const AName, AText: string;
      out AError: string): Boolean; static;

    /// <summary>Marks a setting as "this IDE instance only": Save() then
    ///  writes the value the REGISTRY had, not the current one, so an
    ///  unrelated Save elsewhere (the rename dialog stores its scope, the
    ///  uses cleanup its keep list) cannot make a temporary change
    ///  permanent behind the user''s back. Clearing the mark makes the
    ///  current value ordinary again, and the next Save persists it.
    ///  </summary>
    class procedure MarkSessionOnly(const AName: string;
      AOn: Boolean); static;
    /// <summary>True while that setting is a session-only override, with
    ///  the text the registry still holds.</summary>
    class function IsSessionOnly(const AName: string;
      out APersisted: string): Boolean; static;
    /// <summary>The names that are session-only right now.</summary>
    class function SessionOnlyNames: TArray<string>; static;

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

// A registry value of the WRONG TYPE used to raise out of Load, and the IDE
// then refused to load the ENTIRE package: "Package ...370.bpl kann nicht
// geladen werden. Ungueltiger Datentyp fuer 'VerifySession'." (seen
// 2026-09-21, after a script wrote the value as a string). One stray value -
// from an older build, another tool or a hand edit - must never cost the
// user the whole plugin. Every read goes through these two now: wrong type
// or unreadable = the default, and a string value is still understood.
function ReadBoolDef(AReg: TRegistry; const AName: string; ADefault: Boolean): Boolean;
var
  S: string;
begin
  Result := ADefault;
  try
    if not AReg.ValueExists(AName) then Exit;
    case AReg.GetDataType(AName) of
      rdInteger: Result := AReg.ReadInteger(AName) <> 0;
      rdString, rdExpandString:
        begin
          S := Trim(AReg.ReadString(AName));
          Result := (S <> '') and (S <> '0') and not SameText(S, 'false');
        end;
    end;
  except
    Result := ADefault;
  end;
end;

function ReadIntDef(AReg: TRegistry; const AName: string; ADefault: Integer): Integer;
begin
  Result := ADefault;
  try
    if not AReg.ValueExists(AName) then Exit;
    case AReg.GetDataType(AName) of
      rdInteger: Result := AReg.ReadInteger(AName);
      rdString, rdExpandString: Result := StrToIntDef(Trim(AReg.ReadString(AName)), ADefault);
    end;
  except
    Result := ADefault;
  end;
end;

function ReadStrDef(AReg: TRegistry; const AName: string;
  const ADefault: string): string;
begin
  Result := ADefault;
  try
    if not AReg.ValueExists(AName) then Exit;
    case AReg.GetDataType(AName) of
      rdString, rdExpandString: Result := AReg.ReadString(AName);
      rdInteger: Result := IntToStr(AReg.ReadInteger(AName));
    end;
  except
    Result := ADefault;
  end;
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
      FPrewarmLspOnProjectOpen := ReadBoolDef(Reg, 'PrewarmLspOnProjectOpen', FPrewarmLspOnProjectOpen);
      FLiveBlame := ReadBoolDef(Reg, 'LiveBlame', FLiveBlame);
      FBlameColumnWidth := ReadIntDef(Reg, 'BlameColumnWidth', FBlameColumnWidth);
      FBlameInfo := ReadIntDef(Reg, 'BlameInfo', FBlameInfo);
      FBlameColumnOffset := ReadIntDef(Reg, 'BlameColumnOffset', FBlameColumnOffset);
      FBlameUseTortoise := ReadBoolDef(Reg, 'BlameUseTortoise', FBlameUseTortoise);
      FIntfVisualizerViewer := ReadBoolDef(Reg, 'IntfVisualizerViewer',
        FIntfVisualizerViewer);
      FIntfVisualizerLine := ReadBoolDef(Reg, 'IntfVisualizerLine',
        FIntfVisualizerLine);
      FLspLogging := ReadBoolDef(Reg, 'LspLogging', FLspLogging);
      FVerifySession := ReadBoolDef(Reg, 'VerifySession', FVerifySession);
      // ALL of them (audit #37, L4e): these five were missing, so Save
      // wrote their DEFAULTS to the new key and DeleteKey then took the
      // user's values away with the old branch - a migration that loses
      // settings is worse than none.
      FScopeIncludeOpenUnits := ReadBoolDef(Reg, 'ScopeIncludeOpenUnits',
        FScopeIncludeOpenUnits);
      FScopeIncludeUsedUnits := ReadBoolDef(Reg, 'ScopeIncludeUsedUnits',
        FScopeIncludeUsedUnits);
      FRenameScope := ReadIntDef(Reg, 'RenameScope', FRenameScope);
      FRenameBackup := ReadBoolDef(Reg, 'RenameBackup', FRenameBackup);
      FUsesCleanupKeepUnits := ReadStrDef(Reg, 'UsesCleanupKeepUnits',
        FUsesCleanupKeepUnits);
      FLineEndingCheck := ReadBoolDef(Reg, 'LineEndingCheck', FLineEndingCheck);
      FLineEndingAutoFix := ReadBoolDef(Reg, 'LineEndingAutoFix',
        FLineEndingAutoFix);
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
  // The VIEWER is opt-in per use anyway (the debugger lists it for the
  // type and the user picks it), so it may default to on. The LINE is a
  // takeover - one replacer per type for the whole debugger - so it does
  // not.
  FIntfVisualizerViewer := True;
  FIntfVisualizerLine := False;
  FLspLogging := False;
  FVerifySession := True;
  FScopeIncludeOpenUnits := True;
  FScopeIncludeUsedUnits := False;
  FRenameScope := 0;
  FRenameBackup := True;
  FUsesCleanupKeepUnits := '';
  FLineEndingCheck := True;
  FLineEndingAutoFix := False;
  FLoaded := True;

  try
    MigrateLegacyKey;
  except
    // a broken legacy branch is not worth a failed package load either
  end;

  try
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly(RegistryKey) then
    try
      FPrewarmLspOnProjectOpen := ReadBoolDef(Reg, 'PrewarmLspOnProjectOpen', FPrewarmLspOnProjectOpen);
      FLiveBlame := ReadBoolDef(Reg, 'LiveBlame', FLiveBlame);
      FBlameColumnWidth := ReadIntDef(Reg, 'BlameColumnWidth', FBlameColumnWidth);
      FBlameInfo := ReadIntDef(Reg, 'BlameInfo', FBlameInfo);
      FBlameColumnOffset := ReadIntDef(Reg, 'BlameColumnOffset', FBlameColumnOffset);
      FBlameUseTortoise := ReadBoolDef(Reg, 'BlameUseTortoise', FBlameUseTortoise);
      FIntfVisualizerViewer := ReadBoolDef(Reg, 'IntfVisualizerViewer',
        FIntfVisualizerViewer);
      FIntfVisualizerLine := ReadBoolDef(Reg, 'IntfVisualizerLine',
        FIntfVisualizerLine);
      FLspLogging := ReadBoolDef(Reg, 'LspLogging', FLspLogging);
      FVerifySession := ReadBoolDef(Reg, 'VerifySession', FVerifySession);
      FScopeIncludeOpenUnits := ReadBoolDef(Reg, 'ScopeIncludeOpenUnits', FScopeIncludeOpenUnits);
      FScopeIncludeUsedUnits := ReadBoolDef(Reg, 'ScopeIncludeUsedUnits', FScopeIncludeUsedUnits);
      FRenameScope := ReadIntDef(Reg, 'RenameScope', FRenameScope);
      if (FRenameScope < 0) or (FRenameScope > 3) then FRenameScope := 0;
      FRenameBackup := ReadBoolDef(Reg, 'RenameBackup', FRenameBackup);
      FUsesCleanupKeepUnits := ReadStrDef(Reg, 'UsesCleanupKeepUnits',
        FUsesCleanupKeepUnits);
      FLineEndingCheck := ReadBoolDef(Reg, 'LineEndingCheck', FLineEndingCheck);
      FLineEndingAutoFix := ReadBoolDef(Reg, 'LineEndingAutoFix',
        FLineEndingAutoFix);
    finally
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
  except
    // LAST net: the defaults above are already in place. Settings are a
    // convenience - they must never decide whether the plugin loads.
  end;
end;

class function TPluginSettings.SettingInfos: TArray<TPluginSettingInfo>;

  function B(const AName, ADefault, ASummary: string): TPluginSettingInfo;
  begin
    Result := Default(TPluginSettingInfo);
    Result.Name := AName;
    Result.Kind := pskBool;
    Result.DefaultText := ADefault;
    Result.Summary := ASummary;
  end;

  function I(const AName, ADefault: string; AMin, AMax: Integer;
    const AChoices, ASummary: string): TPluginSettingInfo;
  begin
    Result := Default(TPluginSettingInfo);
    Result.Name := AName;
    Result.Kind := pskInt;
    Result.DefaultText := ADefault;
    Result.Min := AMin;
    Result.Max := AMax;
    Result.Choices := AChoices;
    Result.Summary := ASummary;
  end;

  function S(const AName, ADefault, ASummary: string): TPluginSettingInfo;
  begin
    Result := Default(TPluginSettingInfo);
    Result.Name := AName;
    Result.Kind := pskString;
    Result.DefaultText := ADefault;
    Result.Summary := ASummary;
  end;

begin
  Result := [
    B('PrewarmLspOnProjectOpen', BoolToStr(DefaultPrewarm, True),
      'Start our own DelphiLSP and index the project as soon as it opens, ' +
      'so the first refactoring does not wait for a cold start. Costs one ' +
      'more DelphiLsp process while the IDE is open.'),
    B('VerifySession', 'True',
      'Verify the candidates of find references / rename through a SECOND ' +
      'DelphiLsp session in agent mode: about twice as fast and never ' +
      'aborted after 10 s, at the price of one more process (~290 MB) that ' +
      'is shut down again after 10 idle minutes.'),
    B('LspLogging', 'False',
      'Start our DelphiLsp with -LogModes 255, so %TEMP%\DelphiLSP\' +
      'RefactoringLight*.log records every request and its duration. The ' +
      'only way to diagnose a session that answers nothing; the log grows.'),
    B('ScopeIncludeOpenUnits', 'True',
      'Project-wide scans also look at units that are OPEN in the editor ' +
      'but not part of the project. The unit at the caret is always ' +
      'included, whatever this says.'),
    B('ScopeIncludeUsedUnits', 'False',
      '...and at units reachable through uses clauses, resolved via the ' +
      'identifier index and never below the RAD Studio installation. Off ' +
      'by default: with big libraries that is thousands of files.'),
    I('RenameScope', '0', 0, 3, '0 = whole project, 1 = current unit, ' +
      '2 = current method, 3 = selected units',
      'The scope the rename dialog opens with - it remembers the last one.'),
    B('RenameBackup', 'True',
      'The rename dialog''s "create backup" box, remembered. A backup is ' +
      'written to %LOCALAPPDATA%\DelphiRefactoringLight\backup\<stamp>\ ' +
      'before the first file is changed.'),
    S('UsesCleanupKeepUnits', '',
      'Units the uses cleanup must never offer for removal, as masks ' +
      'separated by ";" (dxSkin*;MyCompany.Components.*). For a false ' +
      'positive the designer query cannot cover. Empty by default - a mask ' +
      'like cx* would also protect cx units that really are unused.'),
    B('LineEndingCheck', 'True',
      'Check a unit''s line endings when the IDE opens it and warn when ' +
      'they are not plain CRLF. Delphi copes badly with LF-only and worst ' +
      'with MIXED endings, and the open is the moment at which nobody has ' +
      'edited the buffer yet.'),
    B('LineEndingAutoFix', 'False',
      '...and correct it without asking. Off by default: rewriting every ' +
      'line break is a change somebody should have said yes to once.'),
    B('LiveBlame', 'False',
      'Live blame in the editor gutter (git and svn). Off by default: it ' +
      'runs blame per file, and not every project is a working copy.'),
    I('BlameInfo', '1', 0, 2,
      '0 = revision, 1 = revision + author, 2 = revision + author + age',
      'What the blame column shows.'),
    I('BlameColumnWidth', '150', 0, 600, '',
      'Pixels the gutter is WIDENED by while live blame is on. 0 = do not ' +
      'touch the gutter at all (only the age stripe and the caret line) - ' +
      'the polite setting when another add-on uses that space.'),
    I('BlameColumnOffset', '17', 0, 600, '',
      'Pixels between the start of the gutter''s data area and our column. ' +
      '17 is measured against the Parnassus Navigator, which ships with ' +
      'Delphi and draws at the very left of the same area.'),
    B('IntfVisualizerViewer', 'True',
      'Offer the interface visualizer while debugging: a window that ' +
      'resolves an interface reference to the object behind it and lists ' +
      'the class hierarchy with the interfaces each level implements, read ' +
      'from the RTTI bytes. Opt-in per use - the debugger offers it in its ' +
      'own list for the type.'),
    B('IntfVisualizerLine', 'False',
      '...and replace the value SHOWN for an interface reference in every ' +
      'tooltip, watch, local view and inspector with one line naming the ' +
      'class behind it. Off by default: the IDE allows only one value ' +
      'replacer per type, so this takes over the display of every ' +
      'interface variable in the debugger.'),
    B('BlameUseTortoise', 'True',
      'Use TortoiseGit / TortoiseSVN for the commit and blame views when ' +
      'they are installed, and fall back to the built-in ones when they ' +
      'are not.')
  ];
end;

class function TPluginSettings.FindSetting(const AName: string;
  out AInfo: TPluginSettingInfo): Boolean;
begin
  AInfo := Default(TPluginSettingInfo);
  for var Info in SettingInfos do
    if SameText(Info.Name, AName) then
    begin
      AInfo := Info;
      Exit(True);
    end;
  Result := False;
end;

// The ONE place that maps a name to its storage. Every caller - the tool,
// Save(), the session-only bookkeeping - goes through these two, so a new
// setting is wired up once.
class function TPluginSettings.GetValueText(const AName: string;
  out AText: string): Boolean;
begin
  if not FLoaded then Load;
  Result := True;
  if SameText(AName, 'PrewarmLspOnProjectOpen') then AText := BoolToStr(FPrewarmLspOnProjectOpen, True)
  else if SameText(AName, 'VerifySession') then AText := BoolToStr(FVerifySession, True)
  else if SameText(AName, 'LspLogging') then AText := BoolToStr(FLspLogging, True)
  else if SameText(AName, 'ScopeIncludeOpenUnits') then AText := BoolToStr(FScopeIncludeOpenUnits, True)
  else if SameText(AName, 'ScopeIncludeUsedUnits') then AText := BoolToStr(FScopeIncludeUsedUnits, True)
  else if SameText(AName, 'RenameScope') then AText := IntToStr(FRenameScope)
  else if SameText(AName, 'RenameBackup') then AText := BoolToStr(FRenameBackup, True)
  else if SameText(AName, 'UsesCleanupKeepUnits') then AText := FUsesCleanupKeepUnits
  else if SameText(AName, 'LineEndingCheck') then AText := BoolToStr(FLineEndingCheck, True)
  else if SameText(AName, 'LineEndingAutoFix') then AText := BoolToStr(FLineEndingAutoFix, True)
  else if SameText(AName, 'LiveBlame') then AText := BoolToStr(FLiveBlame, True)
  else if SameText(AName, 'BlameInfo') then AText := IntToStr(FBlameInfo)
  else if SameText(AName, 'BlameColumnWidth') then AText := IntToStr(FBlameColumnWidth)
  else if SameText(AName, 'BlameColumnOffset') then AText := IntToStr(FBlameColumnOffset)
  else if SameText(AName, 'BlameUseTortoise') then AText := BoolToStr(FBlameUseTortoise, True)
  else if SameText(AName, 'IntfVisualizerViewer') then
    AText := BoolToStr(FIntfVisualizerViewer, True)
  else if SameText(AName, 'IntfVisualizerLine') then
    AText := BoolToStr(FIntfVisualizerLine, True)
  else
  begin
    AText := '';
    Result := False;
  end;
end;

class function TPluginSettings.SetValueText(const AName, AText: string;
  out AError: string): Boolean;
var
  Info: TPluginSettingInfo;
  B: Boolean;
  N: Integer;
  T: string;
begin
  AError := '';
  Result := False;
  if not FLoaded then Load;
  if not FindSetting(AName, Info) then
  begin
    AError := '"' + AName + '" is not a setting of this plugin - call ' +
      'settings without arguments for the list';
    Exit;
  end;

  T := Trim(AText);
  B := False;
  N := 0;
  case Info.Kind of
    pskBool:
      begin
        // The shapes a caller may plausibly send; anything else is a
        // mistake and is named rather than read as False.
        if SameText(T, 'true') or SameText(T, '1') or SameText(T, 'on') or
           SameText(T, 'yes') then B := True
        else if SameText(T, 'false') or SameText(T, '0') or
                SameText(T, 'off') or SameText(T, 'no') then B := False
        else
        begin
          AError := Info.Name + ' is a switch - pass true or false, not "' +
            AText + '"';
          Exit;
        end;
      end;
    pskInt:
      begin
        if not TryStrToInt(T, N) then
        begin
          AError := Info.Name + ' is a number - "' + AText + '" is not one';
          Exit;
        end;
        if (Info.Max > Info.Min) and ((N < Info.Min) or (N > Info.Max)) then
        begin
          AError := Format('%s must be between %d and %d, not %d',
            [Info.Name, Info.Min, Info.Max, N]);
          Exit;
        end;
      end;
  end;

  if SameText(AName, 'PrewarmLspOnProjectOpen') then FPrewarmLspOnProjectOpen := B
  else if SameText(AName, 'VerifySession') then FVerifySession := B
  else if SameText(AName, 'LspLogging') then FLspLogging := B
  else if SameText(AName, 'ScopeIncludeOpenUnits') then FScopeIncludeOpenUnits := B
  else if SameText(AName, 'ScopeIncludeUsedUnits') then FScopeIncludeUsedUnits := B
  else if SameText(AName, 'RenameScope') then FRenameScope := N
  else if SameText(AName, 'RenameBackup') then FRenameBackup := B
  else if SameText(AName, 'UsesCleanupKeepUnits') then FUsesCleanupKeepUnits := Trim(AText)
  else if SameText(AName, 'LineEndingCheck') then FLineEndingCheck := B
  else if SameText(AName, 'LineEndingAutoFix') then FLineEndingAutoFix := B
  else if SameText(AName, 'LiveBlame') then FLiveBlame := B
  else if SameText(AName, 'BlameInfo') then FBlameInfo := N
  else if SameText(AName, 'BlameColumnWidth') then FBlameColumnWidth := N
  else if SameText(AName, 'BlameColumnOffset') then FBlameColumnOffset := N
  else if SameText(AName, 'BlameUseTortoise') then FBlameUseTortoise := B
  else if SameText(AName, 'IntfVisualizerViewer') then FIntfVisualizerViewer := B
  else if SameText(AName, 'IntfVisualizerLine') then FIntfVisualizerLine := B
  else
  begin
    // FindSetting said yes, so this can only be a setting that was added
    // to the descriptor and nowhere else - and silence would make the
    // answer a lie.
    AError := Info.Name + ' is described but not wired up - please report this';
    Exit;
  end;
  Result := True;
end;

class procedure TPluginSettings.MarkSessionOnly(const AName: string;
  AOn: Boolean);
var
  Info: TPluginSettingInfo;
  Persisted: string;
  Keep: TArray<string>;
  KeepVal: TArray<string>;
begin
  if not FindSetting(AName, Info) then Exit;
  Keep := nil;
  KeepVal := nil;
  for var I := 0 to High(FSessionNames) do
    if not SameText(FSessionNames[I], Info.Name) then
    begin
      Keep := Keep + [FSessionNames[I]];
      KeepVal := KeepVal + [FSessionValues[I]];
    end;
  if AOn then
  begin
    // What the registry holds right NOW is what Save must keep writing.
    Persisted := ReadPersistedText(Info);
    Keep := Keep + [Info.Name];
    KeepVal := KeepVal + [Persisted];
  end;
  FSessionNames := Keep;
  FSessionValues := KeepVal;
end;

class function TPluginSettings.IsSessionOnly(const AName: string;
  out APersisted: string): Boolean;
begin
  APersisted := '';
  for var I := 0 to High(FSessionNames) do
    if SameText(FSessionNames[I], AName) then
    begin
      APersisted := FSessionValues[I];
      Exit(True);
    end;
  Result := False;
end;

class function TPluginSettings.SessionOnlyNames: TArray<string>;
begin
  Result := Copy(FSessionNames);
end;

// What the registry holds for one setting, as text - its default when the
// value is not there (or is of the wrong type, which the defensive readers
// turn into the default anyway).
class function TPluginSettings.ReadPersistedText(
  const AInfo: TPluginSettingInfo): string;
var
  Reg: TRegistry;
begin
  Result := AInfo.DefaultText;
  try
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly(RegistryKey) then
      try
        case AInfo.Kind of
          pskBool:
            Result := BoolToStr(ReadBoolDef(Reg, AInfo.Name,
              SameText(AInfo.DefaultText, 'True')), True);
          pskInt:
            Result := IntToStr(ReadIntDef(Reg, AInfo.Name,
              StrToIntDef(AInfo.DefaultText, 0)));
          pskString:
            Result := ReadStrDef(Reg, AInfo.Name, AInfo.DefaultText);
        end;
      finally
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
  except
    // Same rule as Load: a setting must never cost more than its value.
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
      // Driven by the DESCRIPTOR, so a setting cannot be in the list the
      // tool and the options page see and missing here (two hand-kept
      // lists are what made the legacy migration drop five values -
      // audit #37, L4e). A session-only override writes the text the
      // registry already had, never the temporary value.
      for var Info in SettingInfos do
      begin
        var Text: string;
        if not IsSessionOnly(Info.Name, Text) then
          if not GetValueText(Info.Name, Text) then Continue;
        case Info.Kind of
          pskBool: Reg.WriteBool(Info.Name, SameText(Trim(Text), 'True'));
          pskInt: Reg.WriteInteger(Info.Name, StrToIntDef(Trim(Text),
            StrToIntDef(Info.DefaultText, 0)));
          pskString: Reg.WriteString(Info.Name, Text);
        end;
      end;
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
