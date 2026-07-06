(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindOriginalSymbolWizard;

interface

uses
  System.SysUtils, System.Classes, Vcl.Forms, Vcl.Dialogs, Vcl.Controls,
  {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF}
  Expert.EditorHelperIntf, Expert.LspManager, Lsp.Uri, Lsp.Protocol, Lsp.Client;

type
  TLspFindOriginalSymbolWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private
    FContext: TEditorContext;
    procedure SearchAndGo;
  public
    {$IFNDEF STANDALONE_BUILD}
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
  FindOriginalSymbolInstance: TLspFindOriginalSymbolWizard;

implementation

{$IFNDEF STANDALONE_BUILD}
procedure TLspFindOriginalSymbolWizard.AfterSave; begin end;
procedure TLspFindOriginalSymbolWizard.BeforeSave; begin end;
procedure TLspFindOriginalSymbolWizard.Destroyed; begin end;
procedure TLspFindOriginalSymbolWizard.Modified; begin end;

function TLspFindOriginalSymbolWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.FindOriginalSymbolWizard'; end;

function TLspFindOriginalSymbolWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Find Original Symbol'; end;

function TLspFindOriginalSymbolWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspFindOriginalSymbolWizard.GetMenuText: string;
begin Result := 'Find Original Symbol...'; end;
{$ENDIF}

procedure TLspFindOriginalSymbolWizard.Execute;
var
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;

  if not Ctx.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak +
      'Please place the cursor on an identifier.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  FContext := Ctx;
  SearchAndGo;
end;

procedure TLspFindOriginalSymbolWizard.SearchAndGo;
var
  DelphiLspJson, RootPath: string;
  Client: TLspClient;
  Defs: TArray<TLspLocation>;
  WasRunning: Boolean;
begin
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    MessageDlg('No .delphilsp.json found.' + sLineBreak +
      'Enable Tools > Options > Editor > Language > Code Insight > "Generate LSP Config".',
      mtError, [mbOK], 0);
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  Editor.SaveAllFiles;

  Screen.Cursor := crHourGlass;
  try
    WasRunning := TLspManager.Instance.IsAlive;
    Client := TLspManager.Instance.GetClient(RootPath, FContext.ProjectFile, DelphiLspJson);
    Client.RefreshDocument(FContext.FileName);

    if not WasRunning then
    begin
      for var Retry := 1 to 30 do
      begin
        try
          var H := Client.GetHover(FContext.FileName, FContext.Line - 1, FContext.Column - 1);
          if H <> '' then Break;
        except end;
        Sleep(1000);
      end;
    end;

    Defs := Client.GotoDefinition(FContext.FileName, FContext.Line - 1, FContext.Column - 1);
    if Length(Defs) > 0 then
    begin
      var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
      Editor.GotoLocation(DefPath,
        Defs[0].Range.Start.Line, Defs[0].Range.Start.Character,
        System.Length(FContext.WordAtCursor));
    end
    else
      MessageDlg('Original symbol not found.' + sLineBreak +
        'Could not determine the definition of "' + FContext.WordAtCursor + '".',
        mtInformation, [mbOK], 0);
  finally
    Screen.Cursor := crDefault;
  end;
end;

end.
