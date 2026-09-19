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
  Vcl.Forms, Vcl.Dialogs, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF}  Expert.EditorHelperIntf, Expert.FindReferencesDialog, Expert.LspManager, Lsp.Uri, Lsp.Protocol,
  Lsp.Client, Delphi.FileEncoding, Expert.ScopeFiles, Expert.UnitIndex,
  Expert.IncludeExpansion, Expert.InterfaceLinks, Expert.ImplementationFinder;

type
  TLspFindReferencesWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private
    // Bezeichnet nur den GERADE laufenden Search. Verschachtelte
    // Execute-Calls speichern den Vorgaengerwert auf dem Stack und
    // restaurieren ihn am Ende.
    FDialog: TFindReferencesDialog;
    FContext: TEditorContext;
    // candidates decided from the sources, without a DelphiLSP request
    FPreSkipped: Integer;
    procedure DoGotoLocation(AItem: TFindReferenceItem);

    function FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
    function VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName: string;
      const ATargets: TLspSymbolTargets; ALinked: TLinkedTargets;
      AClient: TLspClient; AIncludes: TLspIncludeContext;
      AGraph: TTypeGraph): TFindReferenceItems;
    function ConvertLspLocations(const ALocations: TArray<TLspLocation>; const AOldName: string): TFindReferenceItems;

    procedure SearchAndShow;
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
  FindReferencesInstance: TLspFindReferencesWizard;

implementation


uses
  Winapi.Windows, Expert.PascalScanner, Expert.SafeDeletePlan;

// Content for the kind classification: the editor buffer (main thread),
// else the disk
function EditorOrDiskContent: TFunc<string, string>;
begin
  Result :=
    function(AFile: string): string
    begin
      if not ((Editor <> nil) and Editor.ReadEditorContent(AFile, Result)) then
        Result := ReadDelphiFile(AFile);
    end;
end;

{$IFNDEF STANDALONE_BUILD}
{ TLspFindReferencesWizard - IOTAWizard / IOTAMenuWizard / IOTANotifier glue.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

procedure TLspFindReferencesWizard.AfterSave; begin end;
procedure TLspFindReferencesWizard.BeforeSave; begin end;
procedure TLspFindReferencesWizard.Destroyed; begin end;
procedure TLspFindReferencesWizard.Modified; begin end;

function TLspFindReferencesWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.FindReferencesWizard'; end;

function TLspFindReferencesWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Find References'; end;

function TLspFindReferencesWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspFindReferencesWizard.GetMenuText: string;
begin Result := 'Find references...'; end;
{$ENDIF}
procedure TLspFindReferencesWizard.Execute;
var
  PrevDialog: TFindReferencesDialog;
  PrevContext: TEditorContext;
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;

  if not Ctx.IsValid then
  begin
    MessageDlg('No identifier found at the cursor.' + sLineBreak + 'Please place the cursor on an identifier.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // Save the fields so a nested Execute (e.g. user triggers Find References
  // again via shortcut while ProcessMessages is pumping) doesn't clobber
  // the currently-running search. After this Execute returns, the dialog
  // is detached (SetClosable) and continues to live on its own - that's
  // what gives us multiple-dialogs-at-once support.
  PrevDialog := FDialog;
  PrevContext := FContext;
  try
    FContext := Ctx;
    FDialog := TFindReferencesDialog.CreateDialog(Application.MainForm, Ctx.WordAtCursor);
    FDialog.OnGotoLocation := DoGotoLocation;
    // No OnDialogClose: each dialog manages its own free via
    // SetClosable + caFree once the search is done.
    TLspManager.Instance.ApplyStatusToCaption(FDialog);
    // Show non-modal so the user can interact with the editor (e.g.
    // jump to a found location and edit it) while the result list
    // stays open. Multiple dialogs may co-exist - each search is
    // tracked by its own FDialog/FContext during this Execute call.
    FDialog.Show;
    try
      Application.ProcessMessages;
      SearchAndShow;
    except
      on E: Exception do
        if FDialog <> nil then
          FDialog.SetStatus('Error: ' + E.Message);
    end;
    // Hand off ownership: from now on closing the dialog frees it. If
    // the user already clicked X / Close during the scan, this
    // releases it now.
    if FDialog <> nil then
      FDialog.SetClosable;
  finally
    FDialog := PrevDialog;
    FContext := PrevContext;
  end;
end;

procedure TLspFindReferencesWizard.DoGotoLocation(AItem: TFindReferenceItem);
begin
  // Static-style: uses only the item's own location data. Safe to be
  // assigned as a callback to dialog instances that may outlive the
  // particular Execute call that opened them.
  Editor.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
end;

procedure TLspFindReferencesWizard.SearchAndShow;
var
  DelphiLspJson, RootPath, DefFilePath: string;
  ProjFiles: TArray<string>;
  Items: TFindReferenceItems;
  Client: TLspClient;
  LspLocations: TArray<TLspLocation>;
  LspLine, LspCol: Integer;
  DefLineOut: Integer;
begin
  FPreSkipped := 0;
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

  // Save all editor changes
  FDialog.SetStatus('Saving all files...');
  Editor.SaveAllFiles;

  // Start LSP
  var WasRunning := TLspManager.Instance.IsAlive;
  if WasRunning then
    FDialog.SetStatus('LSP already running. Opening file...')
  else
    FDialog.SetStatus('Starting LSP server (one-time)...');

  Client := TLspManager.Instance.GetClient(
    RootPath, FContext.ProjectFile, DelphiLspJson);

  // send the unit only when its content changed since the last send, and
  // wait for its analysis then - a blind re-open restarts DelphiLSP's
  // analysis and every query before it is done answers null
  begin
    var StartBefore := Client.GetFileDiagnosticsVersion(FContext.FileName);
    if Client.SyncDocument(FContext.FileName) and WasRunning then
      Client.WaitFileAnalysed(FContext.FileName, StartBefore, 30000,
        function: Boolean
        begin
          FDialog.SetStatus('Waiting for DelphiLSP to analyse ' +
            ExtractFileName(FContext.FileName) + '...');
          Application.ProcessMessages;
          Result := True;
        end);
  end;

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
  end;

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
      AssignReferenceKinds(Items, FContext.WordAtCursor, '', -1, EditorOrDiskContent());
      FDialog.SetItems(Items);
      FDialog.SetStatus(Format('LSP: %d reference(s) found.', [Length(Items)]));
      Exit;
    end;
  end;

  // Strategy 2: fallback - text search + GotoDefinition verification
  FDialog.SetStatus('Fallback: text search in project...');

  // Project + the caret's unit + the extras from the settings (see
  // Expert.ScopeFiles).
  ProjFiles := ProjectScopeFiles(FContext.FileName);

  var TextCandidates := FindCandidatesByText(FContext.WordAtCursor, ProjFiles);

  if Length(TextCandidates) = 0 then
  begin
    FDialog.SetItems(nil);
    FDialog.SetStatus('No occurrences found in the project.');
    Exit;
  end;

  // Positions inside {$I} include files are answered through the INCLUDING
  // unit, sent to DelphiLSP with the include expanded (Expert.IncludeExpansion);
  // freeing the context sends the original text again.
  var IncCtx := TLspIncludeContext.Create(Client, EditorOrDiskReader());
  try
    IncCtx.RegisterFiles(ProjFiles);

    // Resolve the declaration (for verification comparison)
    FDialog.SetStatus('Finding declaration...');
    var DefLocs := IncCtx.Definition(FContext.FileName, LspLine, LspCol);
    // The symbol = its declaration + implementation (see TLspSymbolTargets).
    var Targets: TLspSymbolTargets;
    var DefLine := LspLine;
    if Length(DefLocs) > 0 then
    begin
      DefFilePath := TLspUri.FileUriToPath(DefLocs[0].Uri);
      DefLine := DefLocs[0].Range.Start.Line;
      DefLineOut := DefLine;
      IncCtx.AddTargetWithPartner(Targets, DefFilePath, DefLocs[0].Range.Start.Line,
        DefLocs[0].Range.Start.Character);
    end
    else
    begin
      // null AT a declaration: the caret is the symbol
      DefFilePath := FContext.FileName;
      DefLineOut := LspLine;
      IncCtx.AddTargetWithPartner(Targets, FContext.FileName, LspLine, LspCol);
    end;

    // Interface <-> class (user request): the interface declaration a class
    // method implements is a use of it - even when the interface is never
    // called - and calls through the interface reach it; for an interface
    // method, the implementing class methods and the calls on them.
    var Linked := TLinkedTargets.Create;
    try
      var Links: TArray<TMemberLink> := nil;
      // the graph stays alive THROUGH the verification: an occurrence
      // DelphiLSP does not answer for is resolved through the declared
      // type of its qualifier instead of being listed as unverified noise
      var Graph := TTypeGraph.Create(ProjFiles, EditorOrDiskReader());
      try
        var Owner := TImplementationFinder.FindContainingType(DefFilePath, DefLine);
        Links := CollectLinkedTargets(Graph, Owner, FContext.WordAtCursor, Linked);

        // Verify each candidate via GotoDefinition
        Items := VerifyWithLsp(TextCandidates, FContext.WordAtCursor, Targets, Linked,
          Client, IncCtx, Graph);
      finally
        Graph.Free;
      end;

      // a linked declaration outside the scanned files (an interface of a
      // library, say) has no text candidate - list it anyway
      for var L in Links do
      begin
        var Have := False;
        for var It in Items do
          if SameText(ExpandFileName(It.FilePath), ExpandFileName(L.FilePath)) and
             (It.Line = L.Line) then Have := True;
        if Have then Continue;
        var Extra: TFindReferenceItem;
        Extra.FilePath := L.FilePath;
        Extra.Line := L.Line;
        Extra.Col := L.Col;
        Extra.Length := System.Length(FContext.WordAtCursor);
        Extra.Preview := L.Text;
        Extra.Note := '';
        Extra.Relation := Linked.DeclLabel(L.FilePath, L.Line);
        Items := Items + [Extra];
      end;
    finally
      Linked.Free;
    end;
  finally
    IncCtx.Free;
  end;

  // how each hit uses the symbol (the "Kind" column)
  AssignReferenceKinds(Items, FContext.WordAtCursor, DefFilePath, DefLineOut, EditorOrDiskContent());
  var Unverified := 0;
  for var It in Items do
    if It.Note <> '' then Inc(Unverified);
  FDialog.SetItems(Items);
  // how many candidates never needed a DelphiLSP request (their qualifier's
  // declared type already said they belong to another type)
  var FromSource := '';
  if FPreSkipped > 0 then
    FromSource := Format(' %d were decided from the sources without asking ' +
      'DelphiLSP.', [FPreSkipped]);
  if Unverified > 0 then
    FDialog.SetStatus(Format('Fallback: %d of %d candidate(s) verified, %d shown UNVERIFIED ' +
      '(see the Note column).%s', [Length(Items) - Unverified, Length(TextCandidates),
      Unverified, FromSource]))
  else
    FDialog.SetStatus(Format('Fallback: %d of %d candidate(s) verified.%s',
      [Length(Items), Length(TextCandidates), FromSource]));
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

function TLspFindReferencesWizard.FindCandidatesByText(const AOldName: string; const AFiles: TArray<string>): TFindReferenceItems;
var
  CandidateList: TList<TFindReferenceItem>;
  F, Line, RawContent: string;
  Lines, Masked: TArray<string>;
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
        // comment/string state carried ACROSS lines (multi-line { })
        Masked := MaskCommentsAndStrings(Lines);
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

          if BeforeOk and AfterOk and (Masked[LineIdx][FoundPos] = Line[FoundPos]) then
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

function TLspFindReferencesWizard.VerifyWithLsp(const ACandidates: TFindReferenceItems; const AOldName: string;
  const ATargets: TLspSymbolTargets; ALinked: TLinkedTargets;
  AClient: TLspClient; AIncludes: TLspIncludeContext;
  AGraph: TTypeGraph): TFindReferenceItems;
var
  Verified: TList<TFindReferenceItem>;
  Synced: TDictionary<string, Boolean>;
  Contents: TDictionary<string, string>;
  Reader: TIncludeReader;
  I: Integer;
  C: TFindReferenceItem;

  // content of a candidate's file (buffer first), read once per file
  function FileContent(const AFile: string): string;
  begin
    if Contents.TryGetValue(UpperCase(AFile), Result) then Exit;
    if not Reader(AFile, Result) then Result := '';
    Contents.Add(UpperCase(AFile), Result);
  end;

begin
  Verified := TList<TFindReferenceItem>.Create;
  Synced := TDictionary<string, Boolean>.Create;
  Contents := TDictionary<string, string>.Create;
  Reader := EditorOrDiskReader();
  try
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

      // PRE-CHECK without DelphiLSP: a use site whose qualifier has a
      // declared type can be recognised as ANOTHER type's member from the
      // sources alone - and then no request is needed. DelphiLSP refuses a
      // second request while one is open ("Request removed", measured
      // 2026-09-20), so every saved request is saved waiting time.
      begin
        var PreLink: TMemberLink;
        if ClassifyUnansweredUse(AGraph, FileContent(C.FilePath), C.Line, C.Col,
          AOldName,
          function(AFile: string; ALine: Integer): Boolean
          begin
            Result := ATargets.Contains(AFile, ALine) or ALinked.Contains(AFile, ALine);
          end,
          function(AFile: string): Boolean
          begin
            Result := ATargets.ContainsFile(AFile) or ALinked.ContainsFile(AFile);
          end, PreLink) = uuOtherSymbol then
        begin
          Inc(FPreSkipped);
          Continue;
        end;
      end;

      // Send each file once, only when its content changed, and wait for
      // the analysis then (the old re-open + 300 ms lost answers). An
      // include file (or a unit currently sent expanded) is served by the
      // include context - sending it here would undo that.
      var FirstInFile := not Synced.ContainsKey(UpperCase(C.FilePath));
      if FirstInFile then
      begin
        Synced.Add(UpperCase(C.FilePath), True);
        if not AIncludes.OwnsDocument(C.FilePath) then
        begin
          var Before := AClient.GetFileDiagnosticsVersion(C.FilePath);
          if AClient.SyncDocument(C.FilePath) then
          begin
            var Name := ExtractFileName(C.FilePath);
            AClient.WaitFileAnalysed(C.FilePath, Before, 30000,
              function: Boolean
              begin
                FDialog.SetStatus('Waiting for DelphiLSP to analyse ' + Name + '...');
                Application.ProcessMessages;
                Result := True;
              end);
          end;
        end;
      end;

      var Matches := False;
      var NoAnswer := False;
      try
        var Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
        // an EMPTY answer is retried briefly, a wrong one never
        if (System.Length(Defs) = 0) and not ATargets.Contains(C.FilePath, C.Line) and
           not ALinked.Contains(C.FilePath, C.Line) then
        begin
          var Dl := GetTickCount64 + UInt64(IfThen(FirstInFile, 3000, 600));
          while (System.Length(Defs) = 0) and (GetTickCount64 < Dl) do
          begin
            Sleep(150);
            Application.ProcessMessages;
            Defs := AIncludes.Definition(C.FilePath, C.Line, C.Col);
          end;
        end;
        NoAnswer := System.Length(Defs) = 0;

        // The candidate IS one of the symbol's positions (declaration /
        // implementation - DelphiLSP answers those with null or with the
        // counterpart), or DelphiLSP takes it to one of them. The FILE
        // alone says nothing: several same-named methods in one unit.
        if ATargets.Contains(C.FilePath, C.Line) then
          Matches := True
        else if ALinked.Contains(C.FilePath, C.Line) then
        begin
          // the declaration in the interface / the implementing class
          Matches := True;
          C.Relation := ALinked.DeclLabel(C.FilePath, C.Line);
        end
        else if System.Length(Defs) > 0 then
        begin
          var DF := TLspUri.FileUriToPath(Defs[0].Uri);
          var DL := Defs[0].Range.Start.Line;
          Matches := ATargets.Contains(DF, DL);
          if not Matches and ALinked.Contains(DF, DL) then
          begin
            Matches := True;
            C.Relation := ALinked.CallLabel(DF, DL);
          end;
        end;
      except
        // Error -> location skipped
        Matches := False;
      end;

      // DelphiLSP said nothing: resolve the use site through the declared
      // type of its qualifier. Its NEGATIVE answer is the valuable one -
      // "TMyRec2.Init" is simply another symbol and drops out instead of
      // adding an unverified row the user has to judge (tester 2026-09-19).
      if NoAnswer and not Matches then
      begin
        var Link: TMemberLink;
        case ClassifyUnansweredUse(AGraph, FileContent(C.FilePath), C.Line, C.Col,
          AOldName,
          function(AFile: string; ALine: Integer): Boolean
          begin
            Result := ATargets.Contains(AFile, ALine) or ALinked.Contains(AFile, ALine);
          end,
          function(AFile: string): Boolean
          begin
            Result := ATargets.ContainsFile(AFile) or ALinked.ContainsFile(AFile);
          end, Link) of
          uuOurs:
            begin
              Matches := True;
              C.Note := Format('verified via %s (no answer from DelphiLSP)',
                [Link.TypeName]);
            end;
          uuOtherSymbol:
            Continue;      // belongs to another type - not a reference
          uuOverloaded:
            C.Note := Format('UNVERIFIED - overload of %s, DelphiLSP gave no answer',
              [Link.TypeName]);
        end;
      end;

      if Matches then
        Verified.Add(C)
      else if NoAnswer and (C.Note <> '') then
        Verified.Add(C)            // classified above (overload of our type)
      else if NoAnswer and IsIncludeFile(C.FilePath) then
      begin
        // never drop a hit in an include file silently: DelphiLSP could not
        // tell (the including unit may not compile on its own)
        C.Note := 'UNVERIFIED - no answer inside this include file';
        Verified.Add(C);
      end
      else if NoAnswer and not LineDeclaresName(C.Preview, AOldName) then
      begin
        // was silently dropped: an occurrence DelphiLSP does not resolve
        // (inactive {$IFDEF} branch, unit still in analysis) may well be a
        // reference - show it, marked. A declaration line without an
        // answer declares ANOTHER symbol and stays out.
        C.Note := 'UNVERIFIED - no answer from DelphiLSP';
        Verified.Add(C);
      end;
    end;

    FDialog.SetProgress(System.Length(ACandidates), System.Length(ACandidates));
    Result := Verified.ToArray;
  finally
    Contents.Free;
    Synced.Free;
    Verified.Free;
  end;
end;

end.
