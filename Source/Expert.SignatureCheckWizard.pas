(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SignatureCheckWizard;

interface

uses
  System.SysUtils, System.Classes, System.UITypes,
  Vcl.Forms, Vcl.Dialogs, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF} 
  Expert.EditorHelperIntf, Expert.LspManager, Expert.SignatureCheck,
  Expert.SignatureCheckDialog, Lsp.Client;

type
  TLspSignatureCheckWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private
    FDialog: TSignatureCheckDialog;
    FContext: TEditorContext;
    procedure DoGotoLocation(AEntry: TSignatureEntry);
    procedure CollectAndShow;
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
  SignatureCheckInstance: TLspSignatureCheckWizard;

implementation

{$IFNDEF STANDALONE_BUILD}
{ TLspSignatureCheckWizard - IOTAWizard / IOTAMenuWizard / IOTANotifier glue.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

procedure TLspSignatureCheckWizard.AfterSave; begin end;
procedure TLspSignatureCheckWizard.BeforeSave; begin end;
procedure TLspSignatureCheckWizard.Destroyed; begin end;
procedure TLspSignatureCheckWizard.Modified; begin end;

function TLspSignatureCheckWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.SignatureCheckWizard'; end;

function TLspSignatureCheckWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Align Method Signature'; end;

function TLspSignatureCheckWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspSignatureCheckWizard.GetMenuText: string;
begin Result := 'Align method signature...'; end;
{$ENDIF}
procedure TLspSignatureCheckWizard.DoGotoLocation(AEntry: TSignatureEntry);
begin
  Editor.GotoLocation(AEntry.FilePath, AEntry.Line, AEntry.Col,
    Length(AEntry.Name));
end;

procedure TLspSignatureCheckWizard.Execute;
var
  PrevDialog: TSignatureCheckDialog;
  PrevContext: TEditorContext;
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;
  if not Ctx.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak +
      'Please place the cursor on a method name.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // Save the wizard's fields so a nested Execute (user retriggers via
  // shortcut while the search is pumping) doesn't clobber the running
  // search. After return the dialog detaches via SetClosable and lives
  // on its own.
  PrevDialog := FDialog;
  PrevContext := FContext;
  try
    FContext := Ctx;
    FDialog := TSignatureCheckDialog.CreateDialog(Application.MainForm, Ctx.WordAtCursor);
    FDialog.OnGotoLocation := DoGotoLocation;
    TLspManager.Instance.ApplyStatusToCaption(FDialog);
    FDialog.Show;
    try
      Application.ProcessMessages;
      CollectAndShow;
    except
      on E: Exception do
        if FDialog <> nil then
          FDialog.SetStatus('Error: ' + E.Message);
    end;
    if FDialog <> nil then
      FDialog.SetClosable;
  finally
    FDialog := PrevDialog;
    FContext := PrevContext;
  end;
end;

procedure TLspSignatureCheckWizard.CollectAndShow;
var
  DelphiLspJson, RootPath: string;
  Client: TLspClient;
  Entries: TSignatureEntries;
begin
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    FDialog.SetStatus('No .delphilsp.json found - enable Tools > Options > '
      + 'Editor > Language > Code Insight > "Generate LSP Config".');
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  FDialog.SetStatus('Saving all files...');
  Editor.SaveAllFiles;

  var WasRunning := TLspManager.Instance.IsAlive;
  if WasRunning then
    FDialog.SetStatus('LSP already running. Opening file...')
  else
    FDialog.SetStatus('Starting LSP server (one-time)...');

  Client := TLspManager.Instance.GetClient(
    RootPath, FContext.ProjectFile, DelphiLspJson);
  Client.RefreshDocument(FContext.FileName);

  if not WasRunning then
  begin
    for var Retry := 1 to 30 do
    begin
      FDialog.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
      Application.ProcessMessages;
      try
        var H := Client.GetHover(FContext.FileName,
          FContext.Line - 1, FContext.Column - 1);
        if H <> '' then Break;
      except end;
      Sleep(1000);
    end;
  end
  else
    Sleep(300);

  FDialog.SetStatus('Querying document symbols...');
  Application.ProcessMessages;

  try
    Entries := TSignatureChecker.Collect(Client, FContext.FileName, FContext.WordAtCursor);
  except
    on E: Exception do
    begin
      FDialog.SetStatus('LSP error: ' + E.Message);
      Exit;
    end;
  end;

  FDialog.SetEntries(Entries);

  if Length(Entries) = 0 then
    FDialog.SetStatus(Format(
      'No declaration / implementation found for "%s". '
      + 'Place the cursor on the method name.', [FContext.WordAtCursor]))
  else if Length(Entries) = 1 then
    FDialog.SetStatus(Format(
      'Only one entry found for "%s" - no comparison possible. '
      + '(LSP may not have resolved the counterpart.)',
      [FContext.WordAtCursor]))
  else if TSignatureChecker.AllEqual(Entries) then
    FDialog.SetStatus(Format(
      'All %d signatures for "%s" are identical.',
      [Length(Entries), FContext.WordAtCursor]))
  else
    FDialog.SetStatus(Format(
      'Mismatch: %d signatures for "%s" do not all match. '
      + 'Rows marked "NO" differ from the majority.',
      [Length(Entries), FContext.WordAtCursor]));
end;

end.
