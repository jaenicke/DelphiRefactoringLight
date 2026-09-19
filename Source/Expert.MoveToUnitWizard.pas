(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MoveToUnitWizard;

{
  IOTAWizard / IOTAMenuWizard glue for the "Move to unit (project-wide)"
  refactoring. Opens a unit picker, then hands the chosen target off to
  Expert.MoveToUnit.TLspMoveToUnit.Execute.
}

interface

uses
  System.SysUtils, System.UITypes, System.StrUtils,
  Vcl.Forms, Vcl.Dialogs, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF} 
  Expert.EditorHelperIntf, Expert.MoveToUnit, Expert.MoveToUnitDialog;

type
  TLspMoveToUnitWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  public
    {$IFNDEF STANDALONE_BUILD}

    // IOTAWizard / IOTAMenuWizard / IOTANotifier - IDE plugin only.
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    function GetMenuText: string;

    {$ENDIF}
    procedure Execute;
  end;

var
  MoveToUnitInstance: TLspMoveToUnitWizard;

/// <summary>Editor entry point "Move to new unit...": asks for the new
///  unit's name (created next to the current unit; a path is accepted)
///  and moves the identifier at the cursor there.</summary>
procedure MoveToNewUnitAtCursor;

implementation

uses
  Expert.DialogHelper;

{$IFNDEF STANDALONE_BUILD}
{ TLspMoveToUnitWizard - IOTAWizard / IOTAMenuWizard / IOTANotifier glue.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

procedure TLspMoveToUnitWizard.AfterSave; begin end;
procedure TLspMoveToUnitWizard.BeforeSave; begin end;
procedure TLspMoveToUnitWizard.Destroyed; begin end;
procedure TLspMoveToUnitWizard.Modified; begin end;

function TLspMoveToUnitWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.MoveToUnitWizard'; end;

function TLspMoveToUnitWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Move To Unit'; end;

function TLspMoveToUnitWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspMoveToUnitWizard.GetMenuText: string;
begin Result := 'Move to unit...'; end;
{$ENDIF}
procedure TLspMoveToUnitWizard.Execute;
var
  Ctx: TEditorContext;
  Target: string;
begin
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or
     not SameText(ExtractFileExt(Ctx.FileName), '.pas') then
  begin
    MessageDlg('Please open a Delphi unit (.pas) first.',
      mtWarning, [mbOK], 0);
    Exit;
  end;
  if Ctx.WordAtCursor = '' then
  begin
    MessageDlg('Place the cursor on an identifier first.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  if not TMoveToUnitDialog.Choose(Application.MainForm,
       Ctx.WordAtCursor, Ctx.FileName,
       Editor.GetProjectSourceFiles, Target) then
    Exit;

  if Target = '' then Exit;

  TLspMoveToUnit.Execute(Ctx.WordAtCursor, Ctx.FileName, Target, Ctx);
end;

procedure MoveToNewUnitAtCursor;
var
  Ctx: TEditorContext;
  Name, Err: string;
begin
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or not SameText(ExtractFileExt(Ctx.FileName), '.pas') then
  begin
    ShowThemedMessage('Please open a Delphi unit (.pas) first.');
    Exit;
  end;
  if Ctx.WordAtCursor = '' then
  begin
    ShowThemedMessage('Place the cursor on the identifier to move first.');
    Exit;
  end;
  // "TCustomerList" -> "CustomerList": the usual one-class-per-unit name
  Name := Ctx.WordAtCursor;
  if (Length(Name) > 1) and (Name[1] = 'T') and CharInSet(Name[2], ['A'..'Z']) then
    Name := Copy(Name, 2, MaxInt);
  var Dir := ExtractFilePath(Ctx.FileName);
  if not AskThemedText('Move to new unit',
    Format('Move %s into a NEW unit. Unit name (created in %s; a full path is ' +
      'accepted, too):', [Ctx.WordAtCursor, Dir]), Name,
    function(AValue: string): string
    begin
      var V := Trim(AValue);
      if SameText(ExtractFileExt(V), '.pas') then V := ChangeFileExt(V, '');
      Result := CheckNewUnitName(ExtractFileName(V));
      if (Result = '') and FileExists(IfThen(ExtractFilePath(V) <> '', V, Dir + V) + '.pas') then
        Result := 'that unit already exists';
    end) then
    Exit;
  Name := Trim(Name);
  if SameText(ExtractFileExt(Name), '.pas') then Name := ChangeFileExt(Name, '');
  var NewFile := IfThen(ExtractFilePath(Name) <> '', Name, Dir + Name) + '.pas';
  if TLspMoveToUnit.ExecuteToNewUnit(Ctx.WordAtCursor, Ctx.FileName, NewFile, Err) then
  begin
    if Err <> '' then ShowThemedMessage(Err);
    Editor.GotoLocation(NewFile, 0, 0);
  end
  else
    ShowThemedMessage('Move to new unit: ' + Err);
end;

end.
