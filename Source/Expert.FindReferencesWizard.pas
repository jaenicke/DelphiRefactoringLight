(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.FindReferencesWizard;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Types, System.UITypes, System.Math, System.Generics.Collections,
  Vcl.Forms, Vcl.Dialogs, ToolsAPI, Expert.EditorHelper, Expert.FindReferencesDialog, Expert.LspManager, Lsp.Uri, Lsp.Protocol,
  Lsp.Client, Delphi.FileEncoding;

type
  TLspFindReferencesWizard = class(TNotifierObject, IOTAWizard, IOTAMenuWizard)
  private
    FDialog: TFindReferencesDialog;
    FContext: TEditorContext;
    procedure DoGotoLocation(AItem: TFindReferenceItem);

    function FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
    function VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName, ADefFilePath: string;
      AClient: TLspClient): TFindReferenceItems;
    function ConvertLspLocations(const ALocations: TArray<TLspLocation>; const AOldName: string): TFindReferenceItems;
    function IsInStringOrComment(const ALine: string; APos: Integer): Boolean;

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
  FindReferencesInstance: TLspFindReferencesWizard;

implementation

{ TLspFindReferencesWizard - IOTANotifier }

procedure TLspFindReferencesWizard.AfterSave; begin end;
procedure TLspFindReferencesWizard.BeforeSave; begin end;
procedure TLspFindReferencesWizard.Destroyed; begin end;
procedure TLspFindReferencesWizard.Modified; begin end;

function TLspFindReferencesWizard.GetIDString: string;
begin
  Result := 'DelphiRefactoringLight.FindReferencesWizard';
end;

function TLspFindReferencesWizard.GetName: string;
begin
  Result := 'Delphi Refactoring Light - Find References';
end;

function TLspFindReferencesWizard.GetState: TWizardState;
begin
  Result := [wsEnabled];
end;

function TLspFindReferencesWizard.GetMenuText: string;
begin
  Result := 'Find references...';
end;

procedure TLspFindReferencesWizard.Execute;
begin
  FContext := TEditorHelper.GetCurrentContext;

  if not FContext.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak + 'Please place the cursor on an identifier.', mtWarning, [mbOK], 0);
    Exit;
  end;

  FDialog := TFindReferencesDialog.CreateDialog(Application.MainForm, FContext.WordAtCursor);
  try
    FDialog.OnGotoLocation := DoGotoLocation;
    // Run the search right after showing - the dialog loop is kept alive
    // via ProcessMessages while we search.
    FDialog.Show;
    try
      Application.ProcessMessages;
      SearchAndShow;
      // After the search switch to modal so the dialog blocks
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

procedure TLspFindReferencesWizard.DoGotoLocation(AItem: TFindReferenceItem);
begin
  TEditorHelper.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
end;

procedure TLspFindReferencesWizard.SearchAndShow;
var
  DelphiLspJson, RootPath, DefFilePath: string;
  ProjFiles: TArray<string>;
  Items: TFindReferenceItems;
  Client: TLspClient;
  LspLocations: TArray<TLspLocation>;
  LspLine, LspCol: Integer;
begin
  DelphiLspJson := TEditorHelper.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    FDialog.SetStatus('No .delphilsp.json found - enable Tools > Options > '
      + 'Editor > Language > Code Insight > "Generate LSP Config".');
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  // Save all editor changes
  FDialog.SetStatus('Saving all files...');
  TEditorHelper.SaveAllFiles;

  // Start LSP
  var WasRunning := TLspManager.Instance.IsAlive;
  if WasRunning then
    FDialog.SetStatus('LSP already running. Opening file...')
  else
    FDialog.SetStatus('Starting LSP server (one-time)...');

  Client := TLspManager.Instance.GetClient(
    RootPath, FContext.ProjectFile, DelphiLspJson);

  Client.RefreshDocument(FContext.FileName);

  LspLine := FContext.Line - 1;
  LspCol := FContext.Column - 1;

  // On first start wait until ready
  if not WasRunning then
  begin
    for var Retry := 1 to 30 do
    begin
      FDialog.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
      Application.ProcessMessages;
      try
        var H := Client.GetHover(FContext.FileName, LspLine, LspCol);
        if H <> '' then Break;
        var D := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
        if Length(D) > 0 then Break;
      except end;
      Sleep(1000);
    end;
  end
  else
    Sleep(300);

  // Strategy 1: try textDocument/references directly
  if Client.SupportsReferences then
  begin
    FDialog.SetStatus('Querying LSP server for references...');
    try
      LspLocations := Client.FindReferences(FContext.FileName,
        LspLine, LspCol, True);
    except
      on E: Exception do
      begin
        FDialog.SetStatus('LSP error on references: ' + E.Message
          + ' - switching to fallback...');
        SetLength(LspLocations, 0);
      end;
    end;

    if Length(LspLocations) > 0 then
    begin
      Items := ConvertLspLocations(LspLocations, FContext.WordAtCursor);
      FDialog.SetItems(Items);
      FDialog.SetStatus(Format('LSP: %d reference(s) found.', [Length(Items)]));
      Exit;
    end;
  end;

  // Strategy 2: fallback - text search + GotoDefinition verification
  FDialog.SetStatus('Fallback: text search in project...');

  ProjFiles := TEditorHelper.GetProjectSourceFiles;

  var TextCandidates := FindCandidatesByText(FContext.WordAtCursor, ProjFiles);

  if Length(TextCandidates) = 0 then
  begin
    FDialog.SetItems(nil);
    FDialog.SetStatus('No occurrences found in the project.');
    Exit;
  end;

  // Resolve the declaration (for verification comparison)
  FDialog.SetStatus('Finding declaration...');
  var DefLocs := Client.GotoDefinition(FContext.FileName, LspLine, LspCol);
  if Length(DefLocs) > 0 then
    DefFilePath := TLspUri.FileUriToPath(DefLocs[0].Uri)
  else
    DefFilePath := FContext.FileName;

  // Verify each candidate via GotoDefinition
  Items := VerifyWithLsp(TextCandidates, FContext.WordAtCursor,
    DefFilePath, Client);

  FDialog.SetItems(Items);
  FDialog.SetStatus(Format('Fallback: %d of %d candidate(s) verified.',
    [Length(Items), Length(TextCandidates)]));
end;

function TLspFindReferencesWizard.ConvertLspLocations(const ALocations: TArray<TLspLocation>;
  const AOldName: string): TFindReferenceItems;
var
  Item: TFindReferenceItem;
  Lines: TArray<string>;
  LastFile: string;
  ResultList: TList<TFindReferenceItem>;
begin
  ResultList := TList<TFindReferenceItem>.Create;
  try
    LastFile := '';
    SetLength(Lines, 0);
    for var Loc in ALocations do
    begin
      Item.FilePath := TLspUri.FileUriToPath(Loc.Uri);
      Item.Line := Loc.Range.Start.Line;
      Item.Col := Loc.Range.Start.Character;
      Item.Length := Loc.Range.End_.Character - Loc.Range.Start.Character;
      if Item.Length <= 0 then
        Item.Length := System.Length(AOldName);

      Item.Preview := '';
      if not SameText(LastFile, Item.FilePath) then
      begin
        try
          Lines := ReadDelphiFileLines(Item.FilePath);
          LastFile := Item.FilePath;
        except
          SetLength(Lines, 0);
        end;
      end;
      if (Item.Line >= 0) and (Item.Line < System.Length(Lines)) then
        Item.Preview := Trim(Lines[Item.Line]);

      ResultList.Add(Item);
    end;
    Result := ResultList.ToArray;
  finally
    ResultList.Free;
  end;
end;

{ Helper functions for text search (analogous to Rename wizard) }

function TLspFindReferencesWizard.IsInStringOrComment(const ALine: string; APos: Integer): Boolean;
var
  I: Integer;
  InString: Boolean;
begin
  Result := False;
  for I := 1 to APos - 1 do
    if (I < System.Length(ALine)) and (ALine[I] = '/') and (ALine[I + 1] = '/') then
      Exit(True);
  var BraceDepth := 0;
  for I := 1 to APos - 1 do
  begin
    if ALine[I] = '{' then Inc(BraceDepth)
    else if ALine[I] = '}' then Dec(BraceDepth);
  end;
  if BraceDepth > 0 then Exit(True);
  var PSDepth := 0;
  for I := 1 to APos - 2 do
  begin
    if (ALine[I] = '(') and (ALine[I+1] = '*') then Inc(PSDepth)
    else if (ALine[I] = '*') and (ALine[I+1] = ')') then Dec(PSDepth);
  end;
  if PSDepth > 0 then Exit(True);
  InString := False;
  for I := 1 to APos - 1 do
    if ALine[I] = '''' then InString := not InString;
  if InString then Exit(True);
end;

function TLspFindReferencesWizard.FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
var
  CandidateList: TList<TFindReferenceItem>;
  F, Line, RawContent: string;
  Lines: TArray<string>;
  UpperOldName: string;
  LineIdx, SearchPos, FoundPos, AfterPos: Integer;
  BeforeOk, AfterOk: Boolean;
  Item: TFindReferenceItem;
begin
  UpperOldName := UpperCase(AOldName);
  CandidateList := TList<TFindReferenceItem>.Create;
  try
    FDialog.SetProgress(0, System.Length(AFiles));
    for var FileIdx := 0 to High(AFiles) do
    begin
      F := AFiles[FileIdx];
      if (FileIdx mod 5 = 0) then
      begin
        FDialog.SetProgress(FileIdx + 1, System.Length(AFiles));
        Application.ProcessMessages;
      end;

      try
        RawContent := ReadDelphiFile(F);
        if Pos(UpperOldName, UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(F);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        Line := Lines[LineIdx];
        SearchPos := 1;
        while SearchPos <= System.Length(Line) do
        begin
          FoundPos := Pos(UpperOldName, UpperCase(Copy(Line, SearchPos)));
          if FoundPos = 0 then Break;
          FoundPos := SearchPos + FoundPos - 1;

          BeforeOk := (FoundPos = 1) or
            not CharInSet(Line[FoundPos - 1], ['A'..'Z','a'..'z','0'..'9','_']);
          AfterPos := FoundPos + System.Length(AOldName);
          AfterOk := (AfterPos > System.Length(Line)) or
            not CharInSet(Line[AfterPos], ['A'..'Z','a'..'z','0'..'9','_']);

          if BeforeOk and AfterOk and not IsInStringOrComment(Line, FoundPos) then
          begin
            Item.FilePath := F;
            Item.Line := LineIdx;
            Item.Col := FoundPos - 1;
            Item.Length := System.Length(AOldName);
            Item.Preview := Trim(Line);
            CandidateList.Add(Item);
          end;
          SearchPos := FoundPos + System.Length(AOldName);
        end;
      end;
    end;
    FDialog.SetProgress(System.Length(AFiles), System.Length(AFiles));
    Result := CandidateList.ToArray;
  finally
    CandidateList.Free;
  end;
end;

function TLspFindReferencesWizard.VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName, ADefFilePath: string;
  AClient: TLspClient): TFindReferenceItems;
var
  Verified: TList<TFindReferenceItem>;
  LastOpenedFile: string;
  I: Integer;
  C: TFindReferenceItem;
begin
  Verified := TList<TFindReferenceItem>.Create;
  try
    LastOpenedFile := '';
    FDialog.SetProgress(0, System.Length(ACandidates));

    for I := 0 to High(ACandidates) do
    begin
      C := ACandidates[I];
      FDialog.SetProgress(I + 1, System.Length(ACandidates));
      if (I mod 3 = 0) then
      begin
        FDialog.SetStatus(Format('Verifying %d/%d...',
          [I + 1, System.Length(ACandidates)]));
        Application.ProcessMessages;
      end;

      if not SameText(C.FilePath, LastOpenedFile) then
      begin
        AClient.RefreshDocument(C.FilePath);
        Sleep(300);
        LastOpenedFile := C.FilePath;
      end;

      var Matches := False;
      try
        var Defs := AClient.GotoDefinition(C.FilePath, C.Line, C.Col);

        if System.Length(Defs) = 0 then
        begin
          // null = cursor is on the declaration itself
          if SameText(ExpandFileName(C.FilePath), ExpandFileName(ADefFilePath)) then
            Matches := True;
        end
        else
        begin
          var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
          if SameText(ExpandFileName(DefPath), ExpandFileName(ADefFilePath)) then
            Matches := True;
        end;
      except
        // Error -> location skipped
        Matches := False;
      end;

      if Matches then
        Verified.Add(C);
    end;

    FDialog.SetProgress(System.Length(ACandidates), System.Length(ACandidates));
    Result := Verified.ToArray;
  finally
    Verified.Free;
  end;
end;

end.
