(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
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
  System.SysUtils, System.UITypes,
  Vcl.Forms, Vcl.Dialogs, ToolsAPI,
  Expert.EditorHelper, Expert.MoveToUnit, Expert.MoveToUnitDialog;

type
  TLspMoveToUnitWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  public
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    procedure Execute;
    function GetMenuText: string;
  end;

var
  MoveToUnitInstance: TLspMoveToUnitWizard;

implementation

procedure TLspMoveToUnitWizard.AfterSave; begin end;
procedure TLspMoveToUnitWizard.BeforeSave; begin end;
procedure TLspMoveToUnitWizard.Destroyed; begin end;
procedure TLspMoveToUnitWizard.Modified; begin end;

function TLspMoveToUnitWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.MoveToUnitWizard';
end;

function TLspMoveToUnitWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Move To Unit';
end;

function TLspMoveToUnitWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspMoveToUnitWizard.GetMenuText: string;
begin
  Result := 'Move to unit...';
end;

procedure TLspMoveToUnitWizard.Execute;
var
  Ctx: TEditorContext;
  Target: string;
begin
  Ctx := TEditorHelper.GetCurrentContext;
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
       TEditorHelper.GetProjectSourceFiles, Target) then
    Exit;

  if Target = '' then Exit;

  TLspMoveToUnit.Execute(Ctx.WordAtCursor, Ctx.FileName, Target, Ctx);
end;

end.
