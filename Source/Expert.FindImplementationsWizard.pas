(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindImplementationsWizard;

{
  "Find implementations" - shows all class implementations of an interface
  method or a virtual method.

  Strategy analogous to the Rename wizard:
  1. Text search across all project files for the identifier (word boundaries).
  2. Syntactic filter: only lines that look like a class method implementation
     (procedure TClass.Method etc.) are kept.

  This makes the search independent of the LSP index and provides reliable
  results. Uses the same result dialog as Find References.
}

interface

uses
  System.SysUtils, System.Classes, Vcl.Forms, Vcl.Dialogs, ToolsAPI, Expert.EditorHelper, Expert.FindReferencesDialog,
  Expert.LspManager, Expert.ImplementationFinder, Lsp.Uri, Lsp.Client;

type
  TLspFindImplementationsWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  private
    FDialog: TFindReferencesDialog;
    FContext: TEditorContext;
    procedure DoGotoLocation(AItem: TFindReferenceItem);
    procedure SearchAndShow;
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
  FindImplementationsInstance: TLspFindImplementationsWizard;

implementation

{ TLspFindImplementationsWizard - IOTANotifier }

procedure TLspFindImplementationsWizard.AfterSave; begin end;
procedure TLspFindImplementationsWizard.BeforeSave; begin end;
procedure TLspFindImplementationsWizard.Destroyed; begin end;
procedure TLspFindImplementationsWizard.Modified; begin end;

function TLspFindImplementationsWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.FindImplementationsWizard';
end;

function TLspFindImplementationsWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Find Implementations';
end;

function TLspFindImplementationsWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspFindImplementationsWizard.GetMenuText: string;
begin
  Result := 'Find implementations...';
end;

procedure TLspFindImplementationsWizard.Execute;
begin
  FContext := TEditorHelper.GetCurrentContext;

  if not FContext.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak +
      'Please place the cursor on an interface method, or on a call to an ' +
      'interface or virtual method.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  FDialog := TFindReferencesDialog.CreateDialog(Application.MainForm,
    FContext.WordAtCursor, 'Implementations');
  try
    FDialog.OnGotoLocation := DoGotoLocation;
    // Show dialog + run search + switch dialog to modal
    FDialog.Show;
    try
      Application.ProcessMessages;
      SearchAndShow;
      FDialog.Hide;
      FDialog.ShowModal;
    except
      on E: Exception do
      begin
        FDialog.SetStatus('Error: ' + E.Message);
        FDialog.Hide;
        FDialog.ShowModal;
      end;
    end;
  finally
    FDialog.Free;
    FDialog := nil;
  end;
end;

procedure TLspFindImplementationsWizard.DoGotoLocation(AItem: TFindReferenceItem);
begin
  TEditorHelper.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
end;

procedure TLspFindImplementationsWizard.SearchAndShow;
var
  ProjFiles: TArray<string>;
  Items: TFindReferenceItems;
  OwnerType: string;
begin
  // Flush editor changes to the file system, otherwise the text scan would
  // see stale content.
  FDialog.SetStatus('Saving all files...');
  Application.ProcessMessages;
  TEditorHelper.SaveAllFiles;

  ProjFiles := TEditorHelper.GetProjectSourceFiles;
  if System.Length(ProjFiles) = 0 then
  begin
    FDialog.SetStatus('No project context found.');
    Exit;
  end;

  // Determine the container type (interface or class) in which the method
  // is DECLARED (not the enclosing method at the cursor).
  //
  // LSP-FIRST: the cursor may sit on a call like 'aa.Bar' inside
  // 'TXyz.Test'. A pure backward text scan from the cursor would
  // wrongly yield TXyz here. Instead we use LSP GotoDefinition to
  // jump to the method declaration, and determine the container
  // type THERE.
  OwnerType := '';
  FDialog.SetStatus('Resolving method declaration via LSP...');
  Application.ProcessMessages;

  try
    var DelphiLspJson := TEditorHelper.FindDelphiLspJson;
    if DelphiLspJson <> '' then
    begin
      var RootPath := FContext.ProjectRoot;
      if RootPath = '' then
        RootPath := ExtractFilePath(FContext.FileName);

      var WasRunning := TLspManager.Instance.IsAlive;
      if not WasRunning then
        FDialog.SetStatus('Starting LSP server (one-time)...');
      Application.ProcessMessages;

      var Client: TLspClient := TLspManager.Instance.GetClient(RootPath, FContext.ProjectFile, DelphiLspJson);
      Client.RefreshDocument(FContext.FileName);

      // On first start wait until ready
      if not WasRunning then
      begin
        for var Retry := 1 to 30 do
        begin
          FDialog.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
          Application.ProcessMessages;
          try
            var H := Client.GetHover(FContext.FileName, FContext.Line - 1, FContext.Column - 1);
            if H <> '' then Break;
          except end;
          Sleep(1000);
        end;
      end;

      var Defs := Client.GotoDefinition(FContext.FileName, FContext.Line - 1, FContext.Column - 1);
      if System.Length(Defs) > 0 then
      begin
        var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
        var DefLine := Defs[0].Range.Start.Line;
        OwnerType := TImplementationFinder.FindContainingType(DefPath, DefLine);
      end;
    end;
  except
    // Tolerate LSP errors - fallback follows below
  end;

  // Fallback: if LSP could not deliver anything, try a text scan from
  // the cursor. Works reliably when the cursor is on the declaration
  // itself (e.g. 'procedure Bar;' inside an interface block). For calls
  // in a method body the result here is potentially wrong - in that
  // case we prefer showing unverified candidates over incorrectly
  // filtered ones (see status message).
  if OwnerType = '' then
    OwnerType := TImplementationFinder.FindContainingType(
      FContext.FileName, FContext.Line - 1);

  if OwnerType <> '' then
    FDialog.SetStatus(Format('Searching implementations of %s.%s in %d file(s)...',
      [OwnerType, FContext.WordAtCursor, System.Length(ProjFiles)]))
  else
    FDialog.SetStatus(Format('Owner type not resolvable, searching all implementations of "%s"...', [FContext.WordAtCursor]));
  Application.ProcessMessages;

  // Text+syntax scan with verification:
  // For each line like 'procedure TClass.Method' we check whether
  // TClass implements the container type (directly or via inheritance).
  Items := TImplementationFinder.FindByProjectScan(ProjFiles, FContext.WordAtCursor, OwnerType,
    procedure(ACurrent, ATotal: Integer)
    begin
      FDialog.SetProgress(ACurrent, ATotal);
      if (ACurrent mod 5 = 0) or (ACurrent = ATotal) then
      begin
        FDialog.SetStatus(Format('Scanning project (%d/%d)...', [ACurrent, ATotal]));
        Application.ProcessMessages;
      end;
    end);

  FDialog.SetItems(Items);
  FDialog.SetProgress(0, 0);
  if System.Length(Items) = 0 then
  begin
    if OwnerType <> '' then
      FDialog.SetStatus(Format('No implementations of %s.%s found.', [OwnerType, FContext.WordAtCursor]))
    else
      FDialog.SetStatus(Format('No implementations of "%s" found.', [FContext.WordAtCursor]));
  end
  else
  begin
    if OwnerType <> '' then
      FDialog.SetStatus(Format('%d implementation(s) of %s.%s found.', [System.Length(Items), OwnerType, FContext.WordAtCursor]))
    else
      FDialog.SetStatus(Format('%d impl candidate(s) for "%s" (unverified).', [System.Length(Items), FContext.WordAtCursor]));
  end;
end;

end.
