(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ScopeChooser;

// One menu entry per multi-scope tool instead of a submenu: the editor's
// local menu cannot nest (INTAEditorLocalMenu renders one level below our
// root), so "Remove with", "Expand include files" and "Semantic replace"
// used to take twelve flat rows. Each now opens ChooseThemedOption with
// its scopes; the menu entry and the keyboard shortcut both land here, so
// they behave the same. The last choice per tool is remembered for the
// session.

interface

/// <summary>"Remove with...": at cursor / current unit / selected units /
///  whole project.</summary>
procedure ChooseRemoveWith;

/// <summary>"Expand include files...": current unit / selected units / a
///  directory / whole project.</summary>
procedure ChooseExpandIncludes;

/// <summary>"Semantic replace...": current unit / selected units / whole
///  project / edit the rules.</summary>
procedure ChooseSemanticReplace;

implementation

uses
  System.SysUtils,
  Expert.DialogHelper, Expert.EditorHelperIntf,
  Expert.WithRefactorWizard, Expert.IncludeExpander, Expert.SemanticReplaceWizard;

var
  GLastRemoveWith: Integer = 0;
  GLastIncludes: Integer = 0;
  GLastSemantic: Integer = 0;

const
  NoProject = 'needs an open project';
  NoEditor = 'needs an open unit in the editor';

function Choice(const ACaption, AHint: string; AEnabled: Boolean;
  const AWhy: string): TThemedChoice;
begin
  Result.Caption := ACaption;
  Result.Hint := AHint;
  Result.Enabled := AEnabled;
  Result.DisabledWhy := AWhy;
end;

procedure Context(out AHasProject, AHasEditor: Boolean);
begin
  AHasProject := False;
  AHasEditor := False;
  if Editor = nil then Exit;
  AHasProject := Editor.GetCurrentProjectDproj <> '';
  var F := Editor.GetActiveFileName;
  AHasEditor := SameText(ExtractFileExt(F), '.pas') or SameText(ExtractFileExt(F), '.dpr');
end;

procedure ChooseRemoveWith;
var
  P, E: Boolean;
begin
  if WithRefactorInstance = nil then Exit;
  Context(P, E);
  var Idx := GLastRemoveWith;
  if not ChooseThemedOption('Remove with', 'Rewrite "with" statements as inline ' +
    'variables and qualified accesses - where?', [
    Choice('At the cursor', 'only the with statement the cursor is in', E, NoEditor),
    Choice('In the current unit', 'every with statement of the unit in the editor', E, NoEditor),
    Choice('In selected units...', 'pick the units in the next step', P, NoProject),
    Choice('In the whole project', 'every unit of the project', P, NoProject)], Idx) then
    Exit;
  GLastRemoveWith := Idx;
  case Idx of
    0: WithRefactorInstance.ExecuteAtCursor;
    1: WithRefactorInstance.ExecuteCurrentUnit;
    2: WithRefactorInstance.ExecuteSelectedUnits;
    3: WithRefactorInstance.ExecuteProjectWide;
  end;
end;

procedure ChooseExpandIncludes;
var
  P, E: Boolean;
begin
  Context(P, E);
  var Idx := GLastIncludes;
  if not ChooseThemedOption('Expand include files', 'Write the content of the ' +
    '{$I} files into their units (for debugging; marker comments keep the ' +
    'directive) - where?', [
    Choice('In the current unit', 'the unit in the editor', E, NoEditor),
    Choice('In selected units...', 'pick the units in the next step', P, NoProject),
    Choice('In a directory...', 'every .pas / .dpr / .dpk below a folder you pick', True, ''),
    Choice('In the whole project', 'every unit of the project', P, NoProject)], Idx) then
    Exit;
  GLastIncludes := Idx;
  case Idx of
    0: ExpandIncludesCurrentUnit;
    1: ExpandIncludesSelectedUnits;
    2: ExpandIncludesInDirectory;
    3: ExpandIncludesProjectWide;
  end;
end;

procedure ChooseSemanticReplace;
var
  P, E: Boolean;
begin
  Context(P, E);
  var Idx := GLastSemantic;
  if not ChooseThemedOption('Semantic replace', 'Apply the rules of ' +
    'semantic-replace.json (every match verified by DelphiLSP) - where?', [
    Choice('In the current unit', 'the unit in the editor', E and P, NoEditor),
    Choice('In selected units...', 'pick the units in the next step', P, NoProject),
    Choice('In the whole project', 'every unit of the project', P, NoProject),
    Choice('Edit the rules...', 'the rules editor for semantic-replace.json', P, NoProject)],
    Idx) then
    Exit;
  GLastSemantic := Idx;
  case Idx of
    0: ApplySemanticReplacements_CurrentUnit;
    1: ApplySemanticReplacements_SelectedUnits;
    2: ApplySemanticReplacements_Project;
    3: EditSemanticReplaceRules;
  end;
end;

end.
