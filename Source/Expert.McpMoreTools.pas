(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.McpMoreTools;

// The plugin's refactoring and analysis features as MCP tools (IDE-only):
// identifier index, uses clause, implementations, references, unit
// dependencies, debug consistency, blame and rename. Registered with
// Expert.McpServer from the initialization section; the tool DEFINITIONS
// live in Mcp.Protocol.
//
// Same threading rules as the server: ToolsAPI only through McpRunOnMain,
// everything that reads files, waits for DelphiLSP or runs git/svn stays
// on the pipe handler thread and watches AStop.
//
// SCANS READ FILES FROM DISK (implementations, references fallback, uses
// graph for units not open in the editor): unsaved changes in open buffers
// are seen only where noted. Rename saves all files first - exactly what
// the rename dialog does.

interface

implementation

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.JSON, System.IOUtils,
  System.Generics.Collections, System.Generics.Defaults, System.StrUtils, System.Math,
  System.Character,
  Vcl.Forms,
  Expert.McpServer, Expert.McpLspTools, Expert.EditorHelperIntf, Expert.UnitIndex,
  Expert.UnitAvailability, Expert.UsesEditor, Expert.UsesCleanup,
  Expert.ImplementationFinder, Expert.FindReferencesDialog, Expert.ScopeFiles,
  Expert.UsesGraph, Expert.DebugConsistency, Expert.DebugConsistencyDialog,
  Expert.VcsBlame, Expert.RenameWizard, Expert.RenameDialog, Expert.LspManager,
  Expert.PluginSettings, Expert.IncludeExpansion, Expert.InterfaceLinks,
  Expert.SemanticReplace, Expert.SemanticReplaceWizard, Expert.MoveToUnit, Expert.SafeDelete,
  Expert.SafeDeletePlan, Expert.StatementRefactor, Expert.WithRefactorWizard,
  Expert.WithRewriter, Expert.WithScanner, Expert.FindOriginalSymbolWizard,
  Expert.UnitReferencesWizard, Expert.UnitReferencesDialog,
  Expert.InterfaceGuidCheck, Expert.DfmEventCheck, Expert.DfmEventCheckDialog,
  Expert.DfmRename,
  Expert.SignatureCheck, Expert.SignatureCheckWizard,
  Expert.ExtractInterface, Expert.ExtractInterfaceWizard, Expert.ExtractMethod,
  Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Delphi.FileEncoding, Expert.PascalScanner, Expert.McpTools;

// ---------------------------------------------------------------------------
//  Helpers
// ---------------------------------------------------------------------------

function ArgStr(AArgs: TJSONObject; const AName: string; const ADefault: string = ''): string;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<string>(AName, ADefault);
end;

function ArgInt(AArgs: TJSONObject; const AName: string; ADefault: Integer = 0): Integer;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<Integer>(AName, ADefault);
end;

function ArgBool(AArgs: TJSONObject; const AName: string; ADefault: Boolean = False): Boolean;
begin
  Result := ADefault;
  if AArgs <> nil then Result := AArgs.GetValue<Boolean>(AName, ADefault);
end;

function SplitLines(const S: string): TArray<string>;
begin
  Result := S.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

// The identifier at (0-based line, 0-based col) - the caret may also sit
// directly behind it.
function IdentifierAt(const AContent: string; ALine0, ACol0: Integer;
  out AStartCol0: Integer): string;
var
  Lines: TArray<string>;
  S: string;
  P, Q: Integer;
begin
  Result := '';
  AStartCol0 := -1;
  Lines := SplitLines(AContent);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  S := Lines[ALine0];
  P := ACol0 + 1;   // 1-based
  if (P > Length(S)) or not IsIdentChar(S[P]) then
    if (P - 1 >= 1) and (P - 1 <= Length(S)) and IsIdentChar(S[P - 1]) then
      Dec(P)
    else
      Exit;
  Q := P;
  while (P > 1) and IsIdentChar(S[P - 1]) do Dec(P);
  while (Q < Length(S)) and IsIdentChar(S[Q + 1]) do Inc(Q);
  Result := Copy(S, P, Q - P + 1);
  if (Result <> '') and Result[1].IsDigit then Exit('');
  AStartCol0 := P - 1;
end;

function RequireFilePos(AArgs: TJSONObject; out AFile: string;
  out ALine1, ACol1: Integer; out AError: string): Boolean;
begin
  AFile := ArgStr(AArgs, 'file');
  ALine1 := ArgInt(AArgs, 'line');
  ACol1 := ArgInt(AArgs, 'column');
  Result := (AFile <> '') and (ALine1 > 0) and (ACol1 > 0);
  if Result then
    AFile := ExpandFileName(AFile)
  else
    AError := 'arguments "file", "line" and "column" (1-based) are required';
end;

type
  TPosContext = record
    FileName: string;
    Content: string;
    Identifier: string;
    IdentCol0: Integer;
    ProjectFile: string;
    ProjectRoot: string;
    ScopeFiles: TArray<string>;
    Client: TLspClient;
  end;

// Buffer, identifier, project and scope files for a file position - one
// trip to the main thread.
function GatherPosContext(const AFile: string; ALine1, ACol1: Integer;
  AWantScope: Boolean; AStop: THandle; out ACtx: TPosContext;
  out AError: string): Boolean;
var
  Ctx: TPosContext;
  Found: Boolean;
begin
  Result := False;
  Ctx := Default(TPosContext);
  Ctx.FileName := AFile;
  Found := False;
  if not McpRunOnMain(
    procedure
    begin
      Found := McpReadContent(AFile, Ctx.Content);
      if Editor <> nil then
      begin
        Ctx.ProjectFile := Editor.GetCurrentProjectDproj;
        Ctx.ProjectRoot := Editor.GetProjectRoot;
        if AWantScope then
          Ctx.ScopeFiles := ProjectScopeFiles(AFile);
      end;
      Ctx.Client := TLspManager.Instance.PeekClient;
    end, True, AStop, AError) then Exit;
  if not Found then
  begin
    AError := 'file not found: ' + AFile;
    Exit;
  end;
  Ctx.Identifier := IdentifierAt(Ctx.Content, ALine1 - 1, ACol1 - 1, Ctx.IdentCol0);
  if Ctx.Identifier = '' then
  begin
    AError := Format('no identifier at %d:%d', [ALine1, ACol1]);
    Exit;
  end;
  ACtx := Ctx;
  Result := True;
end;

function LineOfFile(const AFile: string; ALine0: Integer;
  ACache: TDictionary<string, TArray<string>>): string;
var
  L: TArray<string>;
begin
  Result := '';
  if not ACache.TryGetValue(UpperCase(AFile), L) then
  begin
    try
      L := ReadDelphiFileLines(AFile);
    except
      L := nil;
    end;
    ACache.Add(UpperCase(AFile), L);
  end;
  if (ALine0 >= 0) and (ALine0 <= High(L)) then Result := Trim(L[ALine0]);
end;

function RefItemsToJson(const AItems: TFindReferenceItems): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var It in AItems do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', It.FilePath);
    O.AddPair('line', TJSONNumber.Create(It.Line + 1));
    O.AddPair('column', TJSONNumber.Create(It.Col + 1));
    O.AddPair('text', Trim(It.Preview));
    if It.Kind <> '' then O.AddPair('kind', It.Kind);
    if It.Relation <> '' then O.AddPair('relation', It.Relation);
    if It.Note <> '' then O.AddPair('note', It.Note);
    Result.Add(O);
  end;
end;

// ---------------------------------------------------------------------------
//  Identifier index + uses clause
// ---------------------------------------------------------------------------

function AvailName(const AUnit, APath: string): string;
begin
  case CheckUnitAvailability(AUnit, APath) of
    uaInProject: Result := 'in_project';
    uaOnSearchPath: Result := 'search_path';
    uaDcuAvailable: Result := 'dcu';
  else
    Result := 'browsing_only';
  end;
end;

function ToolFindUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  Ident, Err: string;
  Hits: TArray<TFindUnitHit>;
  Avail: TArray<string>;
begin
  Ident := Trim(ArgStr(AArgs, 'identifier'));
  if Ident = '' then Exit(McpErr('argument "identifier" is required'));
  var Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or not TUnitIndex.Instance.Ready then
    Exit(McpErr('the identifier index is not ready yet: ' + TUnitIndex.Instance.StatusLine));
  if ArgBool(AArgs, 'partial') then
    Hits := Snap.Search(Ident, ArgInt(AArgs, 'max', 50))
  else
    Hits := Snap.Lookup(Ident);
  SetLength(Avail, Length(Hits));
  // compile-visibility needs the project's search paths (ToolsAPI)
  McpRunOnMain(
    procedure
    begin
      for var I := 0 to High(Hits) do
        Avail[I] := AvailName(Hits[I].UnitName, Hits[I].Path);
    end, True, AStop, Err);
  var Arr := TJSONArray.Create;
  for var I := 0 to High(Hits) do
  begin
    var O := TJSONObject.Create;
    O.AddPair('identifier', Hits[I].Identifier);
    O.AddPair('unit', Hits[I].UnitName);
    O.AddPair('path', Hits[I].Path);
    O.AddPair('generic', TJSONBool.Create(Hits[I].IsGeneric));
    if Avail[I] <> '' then O.AddPair('availability', Avail[I]);
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('hits', Arr);
  Res.AddPair('note', 'The index holds top-level INTERFACE declarations of ' +
    'the library/browsing paths and the project (class members are not ' +
    'indexed). availability browsing_only = the compiler cannot see the unit.');
  Result := McpOk(Res);
end;

// apply=false (the default since 2026-09-24) answers with the CHANGES the
// edit would make and a token; apply=true writes, and with the token only
// when the buffer still looks the way the preview described it.
function ToolAddUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, U, Sec, Err, Msg, NewContent, Token: string;
  Ok, DoApply: Boolean;
  Changes: TArray<TPreviewChange>;
  Total: Integer;
begin
  F := ArgStr(AArgs, 'file');
  U := Trim(ArgStr(AArgs, 'unit'));
  Sec := ArgStr(AArgs, 'section', 'interface');
  if (F = '') or (U = '') then Exit(McpErr('arguments "file" and "unit" are required'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Token := ArgStr(AArgs, 'token');
  Msg := '';
  Ok := False;
  Changes := nil;
  Total := 0;
  if not McpRunOnMain(
    procedure
    var
      Path, C: string;
    begin
      if not McpReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      var Snap := TUnitIndex.Instance.Snapshot;
      if (Snap <> nil) and Snap.TryGetUnitPath(U, Path) and
         (CheckUnitAvailability(U, Path) = uaBrowsingOnly) then
      begin
        Msg := U + ' is only on the IDE''s browsing path (' + Path + ') - the ' +
          'compiler would not find it. Add it in the IDE (project or search path).';
        Exit;
      end;
      var Section := usInterface;
      if SameText(Sec, 'implementation') then Section := usImplementation;
      if not PlanAddUnitToUsesText(C, U, Section, NewContent) then
      begin
        Msg := U + ' was not added - it is already reachable from that section, ' +
          'or the uses clause could not be edited safely (comments inside it)';
        Exit;
      end;
      Changes := DiffToChanges(F, C, NewContent, 40, Total);
      if not DoApply then
      begin
        Token := NewPreviewToken('add_unit', [F], [C]);
        Ok := True;
        Exit;
      end;
      if (Token <> '') and not CheckPreviewToken(Token, 'add_unit',
        function(AFile: string): string
        begin
          if not McpReadContent(AFile, Result) then Result := '';
        end, Msg) then Exit;
      Ok := AddUnitToUses(F, U, Section);
      if not Ok then
        Msg := U + ' could not be added (the buffer changed?)';
    end, not DoApply, AStop, Err) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('unit', U);
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('changes', ChangesToJson(Changes));
  if Total > Length(Changes) then
    Res.AddPair('changesTruncated', TJSONNumber.Create(Total));
  if DoApply then
    Res.AddPair('note', 'Edited in the IDE buffer when the file is open (not ' +
      'saved), otherwise on disk.')
  else
  begin
    Res.AddPair('token', Token);
    Res.AddPair('note', 'Nothing was written. Call again with apply=true (and ' +
      'this token) to make the change.');
  end;
  Result := McpOk(Res);
end;

function ToolRemoveUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, U, Err, Msg, NewContent, Token: string;
  Ok, DoApply: Boolean;
  Changes: TArray<TPreviewChange>;
  Total: Integer;
begin
  F := ArgStr(AArgs, 'file');
  U := Trim(ArgStr(AArgs, 'unit'));
  if (F = '') or (U = '') then Exit(McpErr('arguments "file" and "unit" are required'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Token := ArgStr(AArgs, 'token');
  Ok := False;
  Msg := '';
  Changes := nil;
  Total := 0;
  if not McpRunOnMain(
    procedure
    var
      C: string;
    begin
      if not McpReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      if not PlanRemoveUnitFromUsesText(C, U, NewContent) then
      begin
        Msg := U + ' could not be removed (not in a uses clause of ' + F + '?)';
        Exit;
      end;
      Changes := DiffToChanges(F, C, NewContent, 40, Total);
      if not DoApply then
      begin
        Token := NewPreviewToken('remove_unit', [F], [C]);
        Ok := True;
        Exit;
      end;
      if (Token <> '') and not CheckPreviewToken(Token, 'remove_unit',
        function(AFile: string): string
        begin
          if not McpReadContent(AFile, Result) then Result := '';
        end, Msg) then Exit;
      Ok := RemoveUnitFromUses(F, U);
      if not Ok then Msg := U + ' could not be removed (the buffer changed?)';
    end, not DoApply, AStop, Err) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('unit', U);
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('changes', ChangesToJson(Changes));
  if Total > Length(Changes) then
    Res.AddPair('changesTruncated', TJSONNumber.Create(Total));
  if not DoApply then
  begin
    Res.AddPair('token', Token);
    Res.AddPair('note', 'Nothing was written. Call again with apply=true (and ' +
      'this token) to make the change.');
  end;
  Result := McpOk(Res);
end;

const
  // Appended, never reordered - a consumer keyed on 'unused' stays right.
  VerdictNames: array[TUsesVerdict] of string = ('used', 'unused', 'movable',
    'unknown', 'init_code', 'ide_managed', 'kept_by_user', 'unverified');

type
  // Issue #20: what the IDE's form designer would write into this unit's
  // uses by itself. Only answerable ON THE MAIN THREAD and only for a form
  // unit whose form the IDE has loaded, so both tools fill it inside their
  // McpRunOnMain block and report the state per file.
  TDesignerState = record
    Required: TArray<TDesignerRequiredUnit>;
    Complete: Boolean;    // no selection editor raised while asking
    Answered: Boolean;    // there WAS a designer to ask
    IsForm: Boolean;
    function Verified: Boolean;
    function StateText: string;
  end;

function TDesignerState.Verified: Boolean;
begin
  Result := Answered and Complete;
end;

function TDesignerState.StateText: string;
begin
  if not IsForm then
    Result := 'not a form unit - nothing re-adds units here'
  else if Verified then
    Result := Format('verified: %d unit(s) required by the form designer',
      [Length(Required)])
  else if Answered then
    Result := 'INCOMPLETE: a selection editor raised, so units may be missing'
  else
    Result := 'form not loaded in the designer - the IDE may re-add entries';
end;

// Must run on the main thread (ToolsAPI + designer).
procedure QueryDesigner(const AFile: string; var AState: TDesignerState);
begin
  AState := Default(TDesignerState);
  AState.IsForm := FormFileOf(AFile) <> '';
  if not AState.IsForm or (Editor = nil) then Exit;
  AState.Answered := Editor.GetDesignerRequiredUnits(AFile, AState.Required,
    AState.Complete);
end;

// The three injected lookups are the same for both tools.
function AnalyzeUsesWithDesigner(const AContent, AFile: string;
  const ASnap: IUnitSnapshot; const AState: TDesignerState;
  const AKeepList: string): TArray<TUsesEntryInfo>;
begin
  Result := AnalyzeUses(AContent,
    function(const AIdent: string): TArray<string>
    begin
      Result := nil;
      for var H in ASnap.Lookup(AIdent) do Result := Result + [H.UnitName];
    end,
    function(const AUnitName: string): Boolean
    begin
      Result := ASnap.HasUnit(AUnitName);
    end,
    function(const AUnitName: string): Boolean
    begin
      Result := ASnap.HasInitCode(AUnitName);
    end,
    DesignerRequiredLookup(AState.Required),
    function(const AUnitName: string): Boolean
    begin
      Result := MatchesKeepList(AUnitName, AKeepList);
    end,
    AState.IsForm and not AState.Verified,
    // which clause lines the compiler does not see right now - the same
    // question the dialog asks (user, 2026-10-01)
    InactiveLookup(AFile));
end;

function ToolAnalyzeUses(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, C: string;
  Found: Boolean;
  Cycle0: Integer;
  St: TDesignerState;
begin
  F := ArgStr(AArgs, 'file');
  if F = '' then Exit(McpErr('argument "file" is required'));
  F := ExpandFileName(F);
  Found := False;
  Cycle0 := TUnitIndex.Instance.ScanCycle;
  if not McpRunOnMain(
    procedure
    begin
      Found := McpReadContent(F, C);
      // the index parses from DISK - one fresh cycle, like the dialog does
      TUnitIndex.Instance.RefreshSourcesFromEditor;
      // the designer answers only here, on the main thread (issue #20)
      QueryDesigner(F, St);
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Found then Exit(McpErr('file not found: ' + F));
  var Waited := 0;
  while (TUnitIndex.Instance.ScanCycle = Cycle0) and (Waited < 5000) do
  begin
    if WaitForSingleObject(AStop, 50) = WAIT_OBJECT_0 then Exit(McpErr('shutting down'));
    Inc(Waited, 50);
  end;
  var Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or (Snap.IdentCount = 0) then
    Exit(McpErr('the identifier index is not ready yet'));
  var Entries := AnalyzeUsesWithDesigner(C, F, Snap, St,
    TPluginSettings.UsesCleanupKeepUnits);
  var Arr := TJSONArray.Create;
  for var E in Entries do
  begin
    var O := TJSONObject.Create;
    O.AddPair('unit', E.UnitName);
    if E.Section = usImplementation then O.AddPair('section', 'implementation')
    else O.AddPair('section', 'interface');
    O.AddPair('verdict', VerdictNames[E.Verdict]);
    O.AddPair('usages', TJSONNumber.Create(E.UsageCount));
    if E.FirstUseLine >= 0 then O.AddPair('firstUseLine', TJSONNumber.Create(E.FirstUseLine + 1));
    if E.Reason <> '' then O.AddPair('reason', E.Reason);
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('designerVerified', TJSONBool.Create(St.Verified));
  Res.AddPair('designerState', St.StateText);
  Res.AddPair('entries', Arr);
  Res.AddPair('note', 'unused = no identifier of the unit is used (remove_unit); ' +
    'movable = only used in the implementation (remove + add_unit with ' +
    'section implementation); init_code / ide_managed / kept_by_user / ' +
    'unknown are KEPT deliberately - ide_managed means the form designer ' +
    'writes that entry itself (see "reason"), so removing it is undone on the ' +
    'next save. unverified = a FORM unit whose form is not loaded in the IDE: ' +
    'nothing could confirm what the designer would re-add, so treat those as ' +
    'unknown rather than unused. Class helpers, operators and initialization ' +
    'side effects are invisible to this textual analysis.');
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Implementations + references
// ---------------------------------------------------------------------------

function ToolFindImplementations(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, OwnerType, Note: string;
  L1, C1: Integer;
  Ctx: TPosContext;
  Items: TFindReferenceItems;
begin
  if not RequireFilePos(AArgs, F, L1, C1, Err) then Exit(McpErr(Err));
  if not GatherPosContext(F, L1, C1, True, AStop, Ctx, Err) then Exit(McpErr(Err));
  // The type that DECLARES the member: from the LSP declaration when
  // possible (the caret may sit on a call inside another class' method).
  OwnerType := '';
  if Ctx.Client <> nil then
    try
      var Defs := Ctx.Client.GotoDefinition(F, L1 - 1, Ctx.IdentCol0);
      if Length(Defs) > 0 then
        OwnerType := TImplementationFinder.FindContainingType(
          TLspUri.FileUriToPath(Defs[0].Uri), Defs[0].Range.Start.Line);
    except
    end;
  if OwnerType = '' then
    OwnerType := TImplementationFinder.FindContainingTypeInLines(SplitLines(Ctx.Content), L1 - 1);

  Note := '';
  if SameText(OwnerType, Ctx.Identifier) then
  begin
    Items := TImplementationFinder.FindTypeImplementations(Ctx.ScopeFiles, Ctx.Identifier, nil);
    Note := 'classes implementing / descending from ' + Ctx.Identifier;
  end
  else
  begin
    Items := TImplementationFinder.FindByProjectScan(Ctx.ScopeFiles, Ctx.Identifier, OwnerType, nil);
    if (Length(Items) = 0) and (OwnerType <> '') then
    begin
      Items := TImplementationFinder.FindByProjectScan(Ctx.ScopeFiles, Ctx.Identifier, '', nil);
      if Length(Items) > 0 then
        Note := 'UNVERIFIED: no implementation matched the owner type ' + OwnerType +
          ' - these are all methods named ' + Ctx.Identifier;
    end;
    if Length(Items) = 0 then
    begin
      Items := TImplementationFinder.FindPropertyImplementations(Ctx.ScopeFiles, Ctx.Identifier, nil);
      if Length(Items) > 0 then Note := 'property declarations and their accessor implementations';
    end;
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('identifier', Ctx.Identifier);
  Res.AddPair('ownerType', OwnerType);
  Res.AddPair('filesScanned', TJSONNumber.Create(Length(Ctx.ScopeFiles)));
  Res.AddPair('implementations', RefItemsToJson(Items));
  if Note <> '' then Res.AddPair('note', Note);
  Res.AddPair('basis', 'files on disk - save open buffers first for unsaved changes');
  Result := McpOk(Res);
end;

// Whole-word occurrences outside comments and strings (per line; block
// comments across lines are tracked).
procedure CollectWordHits(const AFile: string; const AWord: string;
  AList: TList<TFindReferenceItem>; AMax: Integer);
var
  Lines: TArray<string>;
  InBrace, InParen: Boolean;
begin
  try
    Lines := ReadDelphiFileLines(AFile);
  except
    Exit;
  end;
  InBrace := False;
  InParen := False;
  for var L := 0 to High(Lines) do
  begin
    var S := Lines[L];
    var I := 1;
    while I <= Length(S) do
    begin
      if AList.Count >= AMax then Exit;
      if InBrace then
      begin
        if S[I] = '}' then InBrace := False;
        Inc(I);
        Continue;
      end;
      if InParen then
      begin
        if (S[I] = '*') and (I < Length(S)) and (S[I + 1] = ')') then
        begin
          InParen := False;
          Inc(I);
        end;
        Inc(I);
        Continue;
      end;
      case S[I] of
        '{': begin InBrace := True; Inc(I); Continue; end;
        '''':
          begin
            Inc(I);
            while (I <= Length(S)) and (S[I] <> '''') do Inc(I);
            Inc(I);
            Continue;
          end;
        '/':
          if (I < Length(S)) and (S[I + 1] = '/') then Break;
        '(':
          if (I < Length(S)) and (S[I + 1] = '*') then
          begin
            InParen := True;
            Inc(I, 2);
            Continue;
          end;
      end;
      if IsIdentChar(S[I]) and ((I = 1) or not IsIdentChar(S[I - 1])) then
      begin
        var J := I;
        while (J <= Length(S)) and IsIdentChar(S[J]) do Inc(J);
        if SameText(Copy(S, I, J - I), AWord) then
        begin
          var It: TFindReferenceItem;
          It.FilePath := AFile;
          It.Line := L;
          It.Col := I - 1;
          It.Length := Length(AWord);
          It.Preview := S;
          AList.Add(It);
        end;
        I := J;
        Continue;
      end;
      Inc(I);
    end;
  end;
end;

function ToolFindReferences(AArgs: TJSONObject; AStop: THandle): string;
const
  MaxCandidates = 400;
var
  F, Err, Method: string;
  L1, C1: Integer;
  Ctx: TPosContext;
  Items: TFindReferenceItems;
  Rejected: TJSONArray;
  DeclFileOut: string;
  DeclLineOut: Integer;
begin
  Rejected := nil;
  DeclFileOut := '';
  DeclLineOut := -1;
  if not RequireFilePos(AArgs, F, L1, C1, Err) then Exit(McpErr(Err));
  if not GatherPosContext(F, L1, C1, True, AStop, Ctx, Err) then Exit(McpErr(Err));
  if Ctx.Client = nil then
    Exit(McpErr('the plugin''s DelphiLSP session is not running yet (it starts with ' +
      'the first opened project)'));
  // Only when changed - every didOpen restarts DelphiLSP's analysis of the
  // unit, and until it is done the unit answers nothing. After a send,
  // wait for the analysis (the per-file diagnostics push).
  // (an include file is no unit - the include context below serves it)
  // VERIFICATION through the agent session when it can be had: twice as
  // fast and free of the controller's 10 s abort (issue #13). Falls back to
  // the main client, and then nothing changes.
  var VClient := TLspManager.Instance.VerificationClient(Ctx.Client);
  if not IsIncludeFile(F) then
  begin
    if VClient <> Ctx.Client then
    begin
      VClient.SyncDocumentWith(F, Ctx.Content);
      // An agent session pushes neither diagnostics nor progress (both
      // measured), so nothing says whether it is ready - and a freshly
      // started one answers every request with null while it loads the
      // project (forum 2026-09-30). Ask BOTH sessions the question this tool
      // depends on - the declaration - and use the one that answers it.
      var PBudget := LspReadinessBudgetMs(Length(Ctx.ScopeFiles));
      var PDl := GetTickCount64 + PBudget;
      repeat
        var PProbe: TArray<TLspLocation> := nil;
        try PProbe := VClient.GotoDefinition(F, L1 - 1, Ctx.IdentCol0);
        except PProbe := nil; end;
        if Length(PProbe) > 0 then Break;
        try PProbe := Ctx.Client.GotoDefinition(F, L1 - 1, Ctx.IdentCol0);
        except PProbe := nil; end;
        if Length(PProbe) > 0 then
        begin
          VClient := Ctx.Client;
          Break;
        end;
        if (WaitForSingleObject(AStop, 500) = WAIT_OBJECT_0)
          or (GetTickCount64 >= PDl) then Break;
      until False;
    end;
    if VClient = Ctx.Client then
    begin
      var StartBefore := Ctx.Client.GetFileDiagnosticsVersion(F);
      if McpSyncLspContent(Ctx.Client, F, Ctx.Content) then
        McpWaitLspAnalysed(Ctx.Client, F, StartBefore, AStop);
    end;
  end;
  Items := nil;
  Method := '';
  if Ctx.Client.SupportsReferences then
    try
      var Locs := Ctx.Client.FindReferences(F, L1 - 1, Ctx.IdentCol0, True);
      var Cache := TDictionary<string, TArray<string>>.Create;
      try
        for var Loc in Locs do
        begin
          var It: TFindReferenceItem;
          It.FilePath := TLspUri.FileUriToPath(Loc.Uri);
          It.Line := Loc.Range.Start.Line;
          It.Col := Loc.Range.Start.Character;
          It.Length := Length(Ctx.Identifier);
          It.Preview := LineOfFile(It.FilePath, It.Line, Cache);
          Items := Items + [It];
        end;
      finally
        Cache.Free;
      end;
      Method := 'DelphiLSP textDocument/references';
    except
      Items := nil;
    end;

  if Length(Items) = 0 then
  begin
    // Text scan over the project scope, every hit verified by asking the
    // LSP where it leads - the approach of the Find References dialog.
    // Positions inside {$I} include files are answered through the
    // INCLUDING unit, sent expanded; freeing the context restores it.
    var ContentOf: TDictionary<string, string> := nil;
    var Linked: TLinkedTargets := nil;
    var Links: TArray<TMemberLink> := nil;
    var IncCtx := TLspIncludeContext.Create(VClient,
      function(const APath: string; out AContent: string): Boolean
      begin
        if (ContentOf <> nil) and ContentOf.TryGetValue(UpperCase(APath), AContent) then
          Exit(True);
        try
          AContent := ReadDelphiFile(APath);
          Result := True;
        except
          Result := False;
        end;
      end);
    try
    IncCtx.RegisterFiles(Ctx.ScopeFiles);
    var Decl := IncCtx.Definition(F, L1 - 1, Ctx.IdentCol0);
    if Length(Decl) > 0 then
    begin
      var CL0 := Ctx.Content.Replace(#13#10, #10).Split([#10]);
      if (L1 - 1 <= High(CL0)) and DeclarationAnswerIsForeign(CL0[L1 - 1], Ctx.Identifier, F,
        TLspUri.FileUriToPath(Decl[0].Uri)) then
        Decl := nil;   // another symbol of that name - the caret is the declaration
    end;
    var DeclFile := '';
    var DeclLine: Integer;
    var DeclCol := 0;
    if Length(Decl) > 0 then
    begin
      DeclFile := ExpandFileName(TLspUri.FileUriToPath(Decl[0].Uri));
      DeclLine := Decl[0].Range.Start.Line;
      DeclCol := Decl[0].Range.Start.Character;
    end
    else
    begin
      // DelphiLSP answers null AT a declaration - then the position IS it
      var CL := Ctx.Content.Replace(#13#10, #10).Split([#10]);
      var CaretDeclLine := '';
      if L1 - 1 <= High(CL) then CaretDeclLine := CL[L1 - 1];
      if not DeclarationAnchorUnknown(False, CaretDeclLine, Ctx.Identifier) then
      begin
        DeclFile := F;
        DeclLine := L1 - 1;
        DeclCol := Ctx.IdentCol0;
      end
      else
      begin
        // No anchor: a caret on a USE whose declaration DelphiLSP did not
        // resolve. Guessing the caret would make every correctly resolved
        // candidate look foreign - the forum case of 2026-09-30 (29 of 327
        // references on a cold session). One more attempt, then refuse.
        var Deadline := GetTickCount64 + 20000;
        while (Length(Decl) = 0) and (GetTickCount64 < Deadline) do
        begin
          if WaitForSingleObject(AStop, 500) = WAIT_OBJECT_0 then Break;
          Decl := IncCtx.Definition(F, L1 - 1, Ctx.IdentCol0);
        end;
        if Length(Decl) > 0 then
        begin
          DeclFile := TLspUri.FileUriToPath(Decl[0].Uri);
          DeclLine := Decl[0].Range.Start.Line;
          DeclCol := Decl[0].Range.Start.Character;
        end
        else
          Exit(McpErr('DelphiLSP did not resolve the declaration of ' +
            Ctx.Identifier + ' at that position, and the line declares nothing - ' +
            'the session is probably still analysing the project (get_status ' +
            'shows whether it is busy). Without the declaration every ' +
            'occurrence would be judged against a guess, so nothing is ' +
            'reported; try again in a moment.'));
      end;
    end;
    DeclFileOut := DeclFile;
    DeclLineOut := DeclLine;
    // THE SYMBOL'S POSITIONS, not just one: declaration AND implementation
    // (DelphiLSP answers a use with either of them, and a use resolved from
    // the sources lands on the DECLARATION while DeclLine may be the
    // implementation). With only one of the two, "Self.Init(...)" inside the
    // record was taken for another symbol and dropped (forum 2026-09-20) -
    // the same shape as the linked-targets gap the day before.
    var Targets: TLspSymbolTargets;
    Targets := Default(TLspSymbolTargets);
    // the COLUMN matters: asked at the line start, DelphiLSP answers
    // nothing and the partner is lost
    IncCtx.AddTargetWithPartner(Targets, DeclFile, DeclLine, DeclCol, Ctx.Identifier);
    // the caret itself, when it sits on a declaration of the name (the
    // partner query can fail - a declaration must never be missing here)
    begin
      var CaretLines := Ctx.Content.Replace(#13#10, #10).Split([#10]);
      if (L1 - 1 <= High(CaretLines)) and LineDeclaresName(CaretLines[L1 - 1], Ctx.Identifier) then
        Targets.Add(F, L1 - 1);
    end;
    // Interface <-> class (user request): the interface declaration a class
    // method implements counts as a use, calls through the interface reach
    // it; for an interface method, the implementing class methods and the
    // calls on them.
    Linked := TLinkedTargets.Create;
    var OwnerTypeName := TImplementationFinder.FindContainingType(DeclFile, DeclLine);
    // kept alive through the verification below: an occurrence DelphiLSP
    // does not answer for is resolved through the declared type of its
    // qualifier instead of being rejected blindly
    var Graph := TTypeGraph.Create(Ctx.ScopeFiles, nil);
    // SELF-CONSISTENCY ANCHOR (PsyPrax report, 2026-09-21): the source
    // pre-check below judges every candidate by where the SOURCES say it
    // leads, and drops the ones that lead elsewhere than the target set.
    // That is only safe while the target set really holds the symbol's
    // declaration - and when the partner query came back empty it did
    // not: all 8 calls of TGemTiFunctions.IsConnectorUnreachable
    // resolved correctly to its declaration and were thrown away as
    // "another type's member". So the START position is judged by the
    // same resolver: the declaration it finds there IS ours, and the
    // pre-check can never disagree with itself about it again.
    begin
      var StartLink: TMemberLink;
      if ResolveMemberUse(Graph, F, Ctx.Content, L1 - 1, Ctx.IdentCol0,
           Ctx.Identifier, StartLink) = murResolved then
        Targets.Add(StartLink.FilePath, StartLink.Line);
    end;
    Links := CollectLinkedTargets(Graph, OwnerTypeName, Ctx.Identifier, Linked);
    var PreSkipped := 0;
    var LspErrors := 0;
    var Cands := TList<TFindReferenceItem>.Create;
    try
      for var SF in Ctx.ScopeFiles do
      begin
        if WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0 then Exit(McpErr('shutting down'));
        CollectWordHits(SF, Ctx.Identifier, Cands, MaxCandidates);
      end;
      // The candidate files must be in the session with their CURRENT
      // content (buffer if open, else disk - read on the main thread in one
      // trip); a file is sent only when that content changed, and after a
      // send we wait for its diagnostics push. (Switching between files is
      // safe: TLspClient.AutoCompleteUnits restores the previous unit before
      // the first query in another one.)
      var Paths: TArray<string> := nil;
      var Seen := TDictionary<string, Boolean>.Create;
      try
        for var Cd in Cands do
          if not Seen.ContainsKey(UpperCase(Cd.FilePath)) then
          begin
            Seen.Add(UpperCase(Cd.FilePath), True);
            Paths := Paths + [Cd.FilePath];
          end;
      finally
        Seen.Free;
      end;
      var Contents: TArray<string>;
      var Found: TArray<Boolean>;
      SetLength(Contents, Length(Paths));
      SetLength(Found, Length(Paths));
      if not McpRunOnMain(
        procedure
        begin
          for var I := 0 to High(Paths) do
            Found[I] := McpReadContent(Paths[I], Contents[I]);
        end, True, AStop, Err) then
        Exit(McpErr(Err));
      ContentOf := TDictionary<string, string>.Create;
      for var I := 0 to High(Paths) do
        if Found[I] then ContentOf.AddOrSetValue(UpperCase(Paths[I]), Contents[I]);
      var Synced := TDictionary<string, Boolean>.Create;    // files checked this call
      var SentCount := 0;
      var TimedOut := 0;
      var Verified: TArray<TFindReferenceItem> := nil;
      // Every candidate that did NOT verify is reported with the answer it
      // got - a silently shorter list is indistinguishable from "no more
      // references", and the answer tells which of LSP / scan / sync failed.
      Rejected := TJSONArray.Create;
      var Answered := TDictionary<string, Boolean>.Create;   // file answered once
      try
      for var Cd in Cands do
      begin
        if WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0 then Exit(McpErr('shutting down'));
        var Answer := '';
        try
          var Key := UpperCase(Cd.FilePath);

          // PRE-CHECK without DelphiLSP: the qualifier's declared type can
          // already say this is another type's member - then no request is
          // needed (the server takes one at a time, so that is real time).
          begin
            var PreContent: string;
            var PreLink: TMemberLink;
            // The target test MUST include the LINKED positions (the
            // implementing classes and their implementation headers).
            // Without them the pre-check dropped "TDialogRenameHost
            // .SetStatus" as a foreign member - caught by comparing the
            // result against the run before the filter existed.
            if ContentOf.TryGetValue(Key, PreContent) and
               (ClassifyUnansweredUse(Graph, Cd.FilePath, PreContent, Cd.Line, Cd.Col,
                 Ctx.Identifier,
                 function(AFile: string; ALine: Integer): Boolean
                 begin
                   Result := Targets.Contains(AFile, ALine) or Linked.Contains(AFile, ALine);
                 end,
                 function(ATypeName: string): Boolean
                 begin
                   Result := SameText(ATypeName, OwnerTypeName) or Linked.HasType(ATypeName);
                 end, PreLink) = uuOtherSymbol) then
            begin
              Inc(PreSkipped);
              Continue;
            end;
          end;

          if not Synced.ContainsKey(Key) then
          begin
            Synced.Add(Key, True);
            var C: string;
            if not IncCtx.OwnsDocument(Cd.FilePath) and ContentOf.TryGetValue(Key, C) then
            begin
              if VClient <> Ctx.Client then
              begin
                // the agent pushes no diagnostics - measured: it answers
                // right after the didOpen
                if VClient.SyncDocumentWith(Cd.FilePath, C) then Inc(SentCount);
              end
              else
              begin
                var Before := Ctx.Client.GetFileDiagnosticsVersion(Cd.FilePath);
                if McpSyncLspContent(Ctx.Client, Cd.FilePath, C) then
                begin
                  Inc(SentCount);
                  if not McpWaitLspAnalysed(Ctx.Client, Cd.FilePath, Before, AStop) then
                    Inc(TimedOut);
                end;
              end;
            end;
          end;
          var D := IncCtx.Definition(Cd.FilePath, Cd.Line, Cd.Col);
          // An EMPTY answer is retried briefly, a WRONG one never: 3 s for
          // the first query in a file sent a moment ago, 0.5 s otherwise.
          // (A unit that answers NOTHING at all is usually not slow: DelphiLSP
          // took it or a unit it uses from a precompiled DCU - see the hint
          // below and CLAUDE.md.)
          var Deadline := GetTickCount64 + 500;
          if not Answered.ContainsKey(Key) then
          begin
            var Ago := McpLspSentAgoMs(Ctx.Client, Cd.FilePath);
            if (Ago >= 0) and (Ago < 10000) then
              Deadline := GetTickCount64 + 3000;
          end;
          while (Length(D) = 0) and (GetTickCount64 < Deadline) and
                (WaitForSingleObject(AStop, 500) <> WAIT_OBJECT_0) do
            D := IncCtx.Definition(Cd.FilePath, Cd.Line, Cd.Col);
          // answered, or had its one longer wait - short retries from now on
          Answered.AddOrSetValue(Key, True);

          // No answer? Then the sources decide: resolve the qualifier's
          // declared type and look the member up there (ancestors
          // included). "Belongs to another type" is the valuable half -
          // it turns an unexplained rejection into a reason.
          var TypeAnswer := '';
          var TypeIsRef := False;
          if Length(D) = 0 then
          begin
            var Content: string;
            if ContentOf.TryGetValue(Key, Content) then
            begin
              var Link: TMemberLink;
              case ClassifyUnansweredUse(Graph, Cd.FilePath, Content, Cd.Line, Cd.Col,
                Ctx.Identifier,
                function(AFile: string; ALine: Integer): Boolean
                begin
                  Result := Targets.Contains(AFile, ALine) or Linked.Contains(AFile, ALine);
                end,
                function(ATypeName: string): Boolean
                begin
                  Result := SameText(ATypeName, OwnerTypeName) or Linked.HasType(ATypeName);
                end, Link) of
                uuOurs: TypeIsRef := True;
                uuOtherSymbol:
                  TypeAnswer := Format('member of %s (resolved through the declared ' +
                    'type - DelphiLSP gave no answer)', [Link.TypeName]);
                uuOverloaded:
                  TypeAnswer := Format('an overload of %s - DelphiLSP gave no answer, ' +
                    'so which one cannot be decided', [Link.TypeName]);
              end;
            end;
          end;
          if (Length(D) > 0) and
             Targets.Contains(TLspUri.FileUriToPath(D[0].Uri), D[0].Range.Start.Line) then
            Verified := Verified + [Cd]
          else if Targets.Contains(Cd.FilePath, Cd.Line) then
            Verified := Verified + [Cd]   // a declaration / implementation itself
          else if Linked.Contains(Cd.FilePath, Cd.Line) then
          begin
            var U := Cd;
            U.Relation := Linked.DeclLabel(Cd.FilePath, Cd.Line);
            Verified := Verified + [U];
          end
          else if (Length(D) > 0) and Linked.Contains(TLspUri.FileUriToPath(D[0].Uri),
            D[0].Range.Start.Line) then
          begin
            var U := Cd;
            U.Relation := Linked.CallLabel(TLspUri.FileUriToPath(D[0].Uri), D[0].Range.Start.Line);
            Verified := Verified + [U];
          end
          else if (Length(D) = 0) and IsIncludeFile(Cd.FilePath) then
          begin
            // never dropped silently: listed, marked unverified
            var U := Cd;
            U.Note := 'UNVERIFIED - no answer inside this include file';
            Verified := Verified + [U];
          end
          else if TypeIsRef then
            // decided from the sources: the qualifier's declared type says
            // this IS our member (DelphiLSP stays silent for a private/
            // public overload pair since 13.1 - RSS-5463)
            Verified := Verified + [Cd]
          else if TypeAnswer <> '' then
            Answer := TypeAnswer
          else if (Length(D) = 0) and LineDeclaresName(Cd.Preview, Ctx.Identifier) then
            // DelphiLSP answers nothing AT a declaration - one that is no
            // position of the symbol declares another symbol (not a sign of
            // a silent unit, so no DCU hint for it)
            Answer := 'another declaration of the name (DelphiLSP answers nothing at declarations)'
          else if Length(D) = 0 then
            Answer := 'no answer from DelphiLSP'
          else
            Answer := Format('leads to %s:%d', [TLspUri.FileUriToPath(D[0].Uri),
              D[0].Range.Start.Line + 1]);
        except
          on E: Exception do
          begin
            // an ERROR is not a negative answer (issue #13) - it is counted
            // so the caller can see a degraded session instead of guessing
            Answer := 'NO ANSWER, the server reported an error: ' + E.Message;
            Inc(LspErrors);
          end;
        end;
        if Answer <> '' then
        begin
          var O := TJSONObject.Create;
          O.AddPair('file', Cd.FilePath);
          O.AddPair('line', TJSONNumber.Create(Cd.Line + 1));
          O.AddPair('column', TJSONNumber.Create(Cd.Col + 1));
          O.AddPair('text', Trim(Cd.Preview));
          O.AddPair('definition', Answer);
          Rejected.Add(O);
        end;
      end;
      finally
        Answered.Free;
        FreeAndNil(ContentOf);
        Synced.Free;
      end;
      // a linked declaration outside the scanned files (an interface of a
      // library, say) has no text candidate - list it anyway
      for var L in Links do
      begin
        var Have := False;
        for var It in Verified do
          if SameText(ExpandFileName(It.FilePath), ExpandFileName(L.FilePath)) and
             (It.Line = L.Line) then Have := True;
        if Have then Continue;
        var Extra := Default(TFindReferenceItem);
        Extra.FilePath := L.FilePath;
        Extra.Line := L.Line;
        Extra.Col := L.Col;
        Extra.Length := Length(Ctx.Identifier);
        Extra.Preview := L.Text;
        Extra.Relation := Linked.DeclLabel(L.FilePath, L.Line);
        Verified := Verified + [Extra];
      end;
      Items := Verified;
      Method := Format('text scan of %d file(s), %d candidate(s) verified via ' +
        'GotoDefinition', [Length(Ctx.ScopeFiles), Cands.Count]);
      if PreSkipped > 0 then
        Method := Method + Format('; %d of them decided from the sources ' +
          '(another type''s member), so that many requests were saved',
          [PreSkipped]);
      if SentCount > 0 then
        Method := Method + Format('; %d file(s) (re)sent to DelphiLSP and ' +
          'waited for', [SentCount]);
      if LspErrors > 0 then
        Method := Method + Format('; WARNING: %d request(s) came back as an ' +
          'ERROR (the server was busy - those occurrences are unverified, ' +
          'not absent; run again)', [LspErrors]);
      if TimedOut > 0 then
        Method := Method + Format(' - %d analysis wait(s) TIMED OUT', [TimedOut]);
      if Cands.Count >= MaxCandidates then
        Method := Method + Format(' (STOPPED at %d candidates)', [MaxCandidates]);
      if IncCtx.Activations > 0 then
        Method := Method + '; include files: ' +
          Trim(IncCtx.NotesText).Replace(sLineBreak, '; ');
    finally
      Cands.Free;
      Graph.Free;
    end;
    finally
      Linked.Free;
      IncCtx.Free;
    end;
  end;

  // how each hit uses the symbol ("kind"); buffers are read on the main
  // thread, one call per file
  AssignReferenceKinds(Items, Ctx.Identifier, DeclFileOut, DeclLineOut,
    function(AFile: string): string
    var
      C, E: string;
    begin
      C := '';
      if not McpRunOnMain(
        procedure
        begin
          if not McpReadContent(AFile, C) then C := '';
        end, True, AStop, E) then C := '';
      Result := C;
    end);
  var Res := TJSONObject.Create;
  Res.AddPair('identifier', Ctx.Identifier);
  Res.AddPair('method', Method);
  Res.AddPair('references', RefItemsToJson(Items));
  if Rejected <> nil then
  begin
    Res.AddPair('declaration', Format('%s:%d', [DeclFileOut, DeclLineOut + 1]));
    Res.AddPair('rejected', Rejected);
    if Pos('no answer from DelphiLSP', Rejected.ToJSON) > 0 then
      Res.AddPair('hint', 'DelphiLSP answered NOTHING for some candidates. When ' +
        'a unit stays silent (try lsp_hover / lsp_definition in it), the usual ' +
        'cause is that a unit (or one it uses) is taken from a precompiled .dcu ' +
        'instead of its source - e.g. DCUs of the project''s own units on ' +
        'the IDE library path or in the DCU output directory. To see it: ' +
        'start the IDE with REFACTORINGLIGHT_LSP_ARGS=-LogModes 255 and read ' +
        '%TEMP%\DelphiLSP\Agent*.log.');
  end;
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Unit dependencies
// ---------------------------------------------------------------------------

function AnalyzeProjectGraph(AStop: THandle; out AResult: TUsesCycleResult;
  out AError: string): Boolean;
var
  Files: TArray<string>;
  Buffers: TDictionary<string, string>;
begin
  Result := False;
  AResult := nil;
  Buffers := TDictionary<string, string>.Create;
  try
    // Open buffers are read on the main thread (unsaved edits count), the
    // rest from disk on this thread.
    if not McpRunOnMain(
      procedure
      var
        C: string;
      begin
        if Editor = nil then Exit;
        Files := Editor.GetProjectSourceFiles;
        for var OF_ in Editor.GetOpenSourceFiles do
          if Editor.ReadEditorContent(OF_, C) then
            Buffers.AddOrSetValue(UpperCase(ExpandFileName(OF_)), C);
      end, True, AStop, AError) then Exit;
    if Length(Files) = 0 then
    begin
      AError := 'no project is open in the IDE';
      Exit;
    end;
    AResult := TUsesGraphAnalyzer.Analyze(Files, nil,
      function(AFile: string): string
      begin
        if not Buffers.TryGetValue(UpperCase(ExpandFileName(AFile)), Result) then
          try
            Result := TDelphiFileEncoding.ReadAll(AFile);
          except
            Result := '';
          end;
      end);
    Result := True;
  finally
    Buffers.Free;
  end;
end;

function ToolUsesPath(AArgs: TJSONObject; AStop: THandle): string;
var
  FromU, ToU, Err: string;
  R: TUsesCycleResult;
begin
  FromU := Trim(ArgStr(AArgs, 'from_unit'));
  ToU := Trim(ArgStr(AArgs, 'to_unit'));
  if (FromU = '') or (ToU = '') then
    Exit(McpErr('arguments "from_unit" and "to_unit" are required'));
  if not AnalyzeProjectGraph(AStop, R, Err) then Exit(McpErr(Err));
  try
    if not R.KnowsUnit(FromU) then Exit(McpErr(FromU + ' is not a unit of the project'));
    if not R.KnowsUnit(ToU) then Exit(McpErr(ToU + ' is not a unit of the project'));
    var Hops := R.FindPath(FromU, ToU, ArgBool(AArgs, 'interface_only', True));
    var Arr := TJSONArray.Create;
    for var H in Hops do
    begin
      var O := TJSONObject.Create;
      O.AddPair('from', H.FromUnit);
      O.AddPair('to', H.ToUnit);
      if H.InInterface then O.AddPair('section', 'interface')
      else O.AddPair('section', 'implementation');
      O.AddPair('file', H.FromFile);
      O.AddPair('line', TJSONNumber.Create(H.Line));
      Arr.Add(O);
    end;
    var Res := TJSONObject.Create;
    Res.AddPair('from', FromU);
    Res.AddPair('to', ToU);
    Res.AddPair('found', TJSONBool.Create(Length(Hops) > 0));
    Res.AddPair('path', Arr);
    Res.AddPair('note', 'Shortest chain of uses entries from from_unit to ' +
      'to_unit. With interface_only (default) only interface-section uses count - ' +
      'the relation behind F2047 "circular unit reference": adding to_unit -> ' +
      'from_unit to an interface uses clause would close this chain.');
    Result := McpOk(Res);
  finally
    R.Free;
  end;
end;

function ToolUsesCycles(AArgs: TJSONObject; AStop: THandle): string;
var
  Err: string;
  R: TUsesCycleResult;
  Truncated: Boolean;
  Paths: TArray<TCyclePath>;
begin
  if not AnalyzeProjectGraph(AStop, R, Err) then Exit(McpErr(Err));
  try
    var U := Trim(ArgStr(AArgs, 'unit'));
    var MaxN := ArgInt(AArgs, 'max', 50);
    if U <> '' then
    begin
      if not R.KnowsUnit(U) then Exit(McpErr(U + ' is not a unit of the project'));
      Paths := R.EnumerateCyclesThrough(U, MaxN, Truncated, 10000);
    end
    else
      Paths := R.EnumerateCycles(MaxN, Truncated, 10000);
    var Arr := TJSONArray.Create;
    for var P in Paths do
      Arr.Add(string.Join(' -> ', P.Units));
    var Levers := TJSONArray.Create;
    var LeverList := R.EdgeLevers;
    for var I := 0 to Min(9, High(LeverList)) do
    begin
      var O := TJSONObject.Create;
      O.AddPair('from', LeverList[I].FromUnit);
      O.AddPair('to', LeverList[I].ToUnit);
      O.AddPair('section', LeverList[I].Section);
      O.AddPair('file', LeverList[I].FromFile);
      O.AddPair('line', TJSONNumber.Create(LeverList[I].Line));
      O.AddPair('unitsFreed', TJSONNumber.Create(LeverList[I].UnitsFreed));
      Levers.Add(O);
    end;
    var Res := TJSONObject.Create;
    Res.AddPair('units', TJSONNumber.Create(Length(R.UnitNames)));
    Res.AddPair('cycles', Arr);
    Res.AddPair('truncated', TJSONBool.Create(Truncated));
    Res.AddPair('bestLevers', Levers);
    Res.AddPair('note', 'Cycles over ALL uses (interface and implementation) of ' +
      'the project units. bestLevers: the uses entries whose removal breaks the ' +
      'most cycles.');
    Result := McpOk(Res);
  finally
    R.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Debug consistency
// ---------------------------------------------------------------------------

function ToolDebugConsistency(AArgs: TJSONObject; AStop: THandle): string;
var
  Input: TDebugCheckInput;
  Err, GErr: string;
  Ok: Boolean;
begin
  Ok := False;
  if not McpRunOnMain(
    procedure
    begin
      Ok := GatherDebugCheckInput(Input, GErr);
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(GErr));
  var Issues := RunDebugConsistencyCheck(Input,
    function(ACurrent, ATotal: Integer; const AText: string): Boolean
    begin
      Result := WaitForSingleObject(AStop, 0) <> WAIT_OBJECT_0;
    end);
  TArray.Sort<TDebugIssue>(Issues, TComparer<TDebugIssue>.Construct(
    function(const A, B: TDebugIssue): Integer
    begin
      Result := Ord(B.Severity) - Ord(A.Severity);
      if Result = 0 then Result := Ord(A.Kind) - Ord(B.Kind);
      if Result = 0 then Result := CompareText(A.FileName, B.FileName);
    end));
  // A real project easily yields 100+ rows of the same kind (stray DCUs
  // of a whole library output dir) - summarise by kind, then list at most
  // "max" rows so the answer stays readable.
  var MaxRows := ArgInt(AArgs, 'max', 60);
  if MaxRows <= 0 then MaxRows := MaxInt;
  var Counts := TDictionary<string, Integer>.Create;
  var Summary := TJSONObject.Create;
  try
    for var I in Issues do
    begin
      var K := DebugIssueKindName(I.Kind);
      var N: Integer;
      if not Counts.TryGetValue(K, N) then N := 0;
      Counts.AddOrSetValue(K, N + 1);
    end;
    for var P in Counts do
      Summary.AddPair(P.Key, TJSONNumber.Create(P.Value));
  finally
    Counts.Free;
  end;
  var Arr := TJSONArray.Create;
  for var I in Issues do
  begin
    if Arr.Count >= MaxRows then Break;
    var O := TJSONObject.Create;
    case I.Severity of
      dsProblem: O.AddPair('severity', 'problem');
      dsWarning: O.AddPair('severity', 'warning');
    else
      O.AddPair('severity', 'note');
    end;
    O.AddPair('kind', DebugIssueKindName(I.Kind));
    O.AddPair('file', I.FileName);
    if I.Line > 0 then O.AddPair('line', TJSONNumber.Create(I.Line));
    O.AddPair('reason', I.Reason);
    if I.Hint <> '' then O.AddPair('hint', I.Hint);
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('project', Input.ProjectFile);
  Res.AddPair('configuration', Input.ConfigName);
  Res.AddPair('total', TJSONNumber.Create(Length(Issues)));
  Res.AddPair('by_kind', Summary);
  Res.AddPair('issues', Arr);
  if Arr.Count < Length(Issues) then
    Res.AddPair('note', Format('%d of %d issues listed (problems first) - pass ' +
      '"max" (0 = all) for more', [Arr.Count, Length(Issues)]));
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Blame
// ---------------------------------------------------------------------------

function WaitForBlame(const AFile: string; AStop: THandle; out ALines: TBlameLines;
  out AError: string): Boolean;
begin
  Result := False;
  if DetectVcs(AFile) = vcsNone then
  begin
    AError := AFile + ' is not in a git or svn working copy';
    Exit;
  end;
  if BlameForFile(AFile, ALines) then Exit(True);
  RequestBlame(AFile);
  for var I := 1 to 600 do   // up to 60 s - svn can be slow
  begin
    if WaitForSingleObject(AStop, 100) = WAIT_OBJECT_0 then
    begin
      AError := 'shutting down';
      Exit;
    end;
    if BlameForFile(AFile, ALines) then Exit(True);
    if BlameLoadFailed(AFile) then
    begin
      // typically an UNTRACKED file (git: "no such path in HEAD")
      AError := 'no blame for ' + ExtractFileName(AFile) + ' - ' + BlameStatus +
        ' (is the file under version control and committed?)';
      Exit;
    end;
  end;
  AError := 'blame did not finish within 60 s (' + BlameStatus + ')';
end;

function ToolBlame(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  Lines: TBlameLines;
begin
  F := ArgStr(AArgs, 'file');
  if F = '' then Exit(McpErr('argument "file" is required'));
  F := ExpandFileName(F);
  if not WaitForBlame(F, AStop, Lines, Err) then Exit(McpErr(Err));
  var First := Max(1, ArgInt(AArgs, 'start_line', 1));
  var Last := ArgInt(AArgs, 'end_line', 0);
  if (Last <= 0) or (Last > Length(Lines)) then Last := Length(Lines);
  if Last - First > 499 then Last := First + 499;
  var Src := ReadDelphiFileLines(F);
  var Arr := TJSONArray.Create;
  for var L := First to Last do
  begin
    var B := Lines[L - 1];
    var O := TJSONObject.Create;
    O.AddPair('line', TJSONNumber.Create(L));
    if B.IsUncommitted then
      O.AddPair('revision', 'uncommitted')
    else
    begin
      O.AddPair('revision', B.ShortHash);
      O.AddPair('author', B.Author);
      O.AddPair('date', FormatDateTime('yyyy-mm-dd hh:nn', B.AuthorTime));
      O.AddPair('summary', B.Summary);
    end;
    if L - 1 <= High(Src) then O.AddPair('text', Src[L - 1]);
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('lines', Arr);
  Res.AddPair('note', 'Blame of the file ON DISK (unsaved buffer changes are not ' +
    'part of it). At most 500 lines per call - use start_line / end_line.');
  Result := McpOk(Res);
end;

function ToolCommitInfo(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err: string;
  Lines: TBlameLines;
  Info: TCommitInfo;
begin
  F := ArgStr(AArgs, 'file');
  var L := ArgInt(AArgs, 'line');
  if (F = '') or (L <= 0) then Exit(McpErr('arguments "file" and "line" are required'));
  F := ExpandFileName(F);
  if not WaitForBlame(F, AStop, Lines, Err) then Exit(McpErr(Err));
  if L > Length(Lines) then Exit(McpErr('the file has only ' + IntToStr(Length(Lines)) + ' line(s)'));
  if Lines[L - 1].IsUncommitted then Exit(McpErr('line ' + IntToStr(L) + ' is not committed yet'));
  if not GetCommitInfo(F, Lines[L - 1], Info) then
    Exit(McpErr('the commit details could not be read'));
  var WithDiff := ArgStr(AArgs, 'diff_file');
  var Arr := TJSONArray.Create;
  for var CF in Info.Files do
  begin
    var O := TJSONObject.Create;
    O.AddPair('path', CF.Path);
    O.AddPair('action', CF.Action);
    if (WithDiff <> '') and (Pos(UpperCase(WithDiff), UpperCase(CF.Path)) > 0) then
      O.AddPair('diff', Copy(CF.Diff, 1, 60000));
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('revision', Info.Revision);
  Res.AddPair('author', Info.Author);
  Res.AddPair('date', Info.DateStr);
  Res.AddPair('subject', Info.Subject);
  if Info.Body <> '' then Res.AddPair('body', Info.Body);
  Res.AddPair('files', Arr);
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Rename (the rename dialog's pipeline without the dialog)
// ---------------------------------------------------------------------------

type
  THeadlessRenameHost = class(TInterfacedObject, IRenameHost)
  public
    NewName: string;
    ScopeValue: TRenameScope;
    Units: TArray<string>;
    IncOpen, IncUsed: Boolean;
    Stop: THandle;
    Items: TRenamePreviewItems;
    Status, Details: string;
    Enabled: Boolean;
    Notes: string;
    function GetNewName: string;
    function CreateBackup: Boolean;
    function Scope: TRenameScope;
    function SelectedUnits: TArray<string>;
    function IncludeOpenUnits: Boolean;
    function IncludeUsedUnits: Boolean;
    function ScanCancelled: Boolean;
    procedure SetBusy(ABusy: Boolean);
    procedure SetStatus(const AText: string);
    procedure SetProgress(AValue, AMax: Integer);
    procedure SetPreviewItems(const AItems: TRenamePreviewItems);
    procedure SetDetailsText(const AText: string);
    procedure EnableRename(AEnabled: Boolean);
    procedure Notify(const AText: string; AWarning: Boolean);
  end;

function THeadlessRenameHost.GetNewName: string; begin Result := NewName; end;
// no backup copies for remote renames - the caller works on a VCS
// checkout and gets the full edit list back
function THeadlessRenameHost.CreateBackup: Boolean; begin Result := False; end;
function THeadlessRenameHost.Scope: TRenameScope; begin Result := ScopeValue; end;
function THeadlessRenameHost.SelectedUnits: TArray<string>; begin Result := Units; end;
function THeadlessRenameHost.IncludeOpenUnits: Boolean; begin Result := IncOpen; end;
function THeadlessRenameHost.IncludeUsedUnits: Boolean; begin Result := IncUsed; end;
function THeadlessRenameHost.ScanCancelled: Boolean;
begin
  Result := (Stop <> 0) and (WaitForSingleObject(Stop, 0) = WAIT_OBJECT_0);
end;
procedure THeadlessRenameHost.SetBusy(ABusy: Boolean); begin end;
procedure THeadlessRenameHost.SetStatus(const AText: string); begin Status := AText; end;
procedure THeadlessRenameHost.SetProgress(AValue, AMax: Integer); begin end;
procedure THeadlessRenameHost.SetPreviewItems(const AItems: TRenamePreviewItems);
begin
  Items := AItems;
end;
procedure THeadlessRenameHost.SetDetailsText(const AText: string); begin Details := AText; end;
procedure THeadlessRenameHost.EnableRename(AEnabled: Boolean); begin Enabled := AEnabled; end;
procedure THeadlessRenameHost.Notify(const AText: string; AWarning: Boolean);
begin
  if Notes <> '' then Notes := Notes + sLineBreak;
  Notes := Notes + AText;
end;

var
  // The last preview, waiting for rename_apply. One at a time: a new
  // preview replaces it (main thread only).
  GRenameWizard: TLspRenameWizard = nil;
  GRenameHost: IRenameHost = nil;          // keeps the host alive ...
  GRenameHostObj: THeadlessRenameHost = nil; // ... and this reads it
  GRenameToken: string = '';

procedure DropPendingRename;
begin
  GRenameHostObj := nil;
  GRenameHost := nil;
  FreeAndNil(GRenameWizard);
  GRenameToken := '';
end;

const
  RenameTimeoutMs = 300000;

function ToolRenamePreview(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, NewName, Msg: string;
  L1, C1: Integer;
  Res: TJSONObject;
begin
  if not RequireFilePos(AArgs, F, L1, C1, Err) then Exit(McpErr(Err));
  NewName := Trim(ArgStr(AArgs, 'new_name'));
  if not IsValidIdent(NewName) then Exit(McpErr('argument "new_name" must be a valid identifier'));
  var ScopeArg := LowerCase(ArgStr(AArgs, 'scope', 'project'));
  Res := nil;
  Msg := '';
  if not McpRunOnMain(
    procedure
    var
      C: string;
      StartCol: Integer;
    begin
      if not McpReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      var Ident := IdentifierAt(C, L1 - 1, C1 - 1, StartCol);
      if Ident = '' then
      begin
        Msg := Format('no identifier at %d:%d', [L1, C1]);
        Exit;
      end;
      var Ctx := Default(TEditorContext);
      Ctx.FileName := F;
      Ctx.Line := L1;
      Ctx.Column := StartCol + 1;
      Ctx.WordAtCursor := Ident;
      Ctx.ProjectFile := Editor.GetCurrentProjectDproj;
      Ctx.ProjectRoot := Editor.GetProjectRoot;
      Ctx.IsValid := True;

      var Host := THeadlessRenameHost.Create;
      Host.NewName := NewName;
      if ScopeArg = 'unit' then Host.ScopeValue := rscCurrentUnit
      else if ScopeArg = 'method' then Host.ScopeValue := rscCurrentMethod
      else Host.ScopeValue := rscProject;
      Host.IncOpen := ArgBool(AArgs, 'include_open_units', TPluginSettings.ScopeIncludeOpenUnits);
      Host.IncUsed := ArgBool(AArgs, 'include_used_units', TPluginSettings.ScopeIncludeUsedUnits);
      Host.Stop := AStop;

      DropPendingRename;
      var HostRef: IRenameHost := Host;
      GRenameWizard := TLspRenameWizard.Create;
      var Any := GRenameWizard.PreviewHeadless(Ctx, HostRef);
      if Any and Host.Enabled then
      begin
        GRenameHost := HostRef;
        GRenameHostObj := Host;
        GRenameToken := TGUID.NewGuid.ToString;
      end
      else
        DropPendingRename;

      Res := TJSONObject.Create;
      Res.AddPair('identifier', Ident);
      Res.AddPair('newName', NewName);
      Res.AddPair('status', Host.Status);
      if Host.Notes <> '' then Res.AddPair('message', Host.Notes);
      var Arr := TJSONArray.Create;
      for var It in Host.Items do
      begin
        var O := TJSONObject.Create;
        O.AddPair('file', It.FilePath);
        O.AddPair('line', TJSONNumber.Create(It.Line + 1));
        O.AddPair('column', TJSONNumber.Create(It.Col + 1));
        O.AddPair('kind', It.Kind);
        O.AddPair('before', Trim(It.OriginalLine));
        O.AddPair('after', Trim(It.PreviewLine));
        Arr.Add(O);
      end;
      Res.AddPair('changes', Arr);
      if GRenameToken <> '' then
        Res.AddPair('token', GRenameToken)
      else
        Res.AddPair('note', 'Nothing to apply.');
      Res.AddPair('details', Copy(Host.Details, 1, 20000));
    end, False, AStop, Err, RenameTimeoutMs) then Exit(McpErr(Err));
  if Msg <> '' then
  begin
    Res.Free;
    Exit(McpErr(Msg));
  end;
  Result := McpOk(Res);
end;

function ToolRenameApply(AArgs: TJSONObject; AStop: THandle): string;
var
  Token, Err, Msg, Outcome: string;
begin
  Token := ArgStr(AArgs, 'token');
  if Token = '' then Exit(McpErr('argument "token" (from rename_preview) is required'));
  Msg := '';
  if not McpRunOnMain(
    procedure
    begin
      if (GRenameWizard = nil) or (Token <> GRenameToken) then
      begin
        Msg := 'unknown or expired token - run rename_preview again';
        Exit;
      end;
      GRenameHostObj.Notes := '';
      GRenameWizard.ApplyHeadless(GRenameHost);
      Outcome := GRenameHostObj.Notes;
      DropPendingRename;
    end, False, AStop, Err, RenameTimeoutMs) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('outcome', Outcome);
  Res.AddPair('note', 'Applied through the IDE editor (undoable with Ctrl+Z in ' +
    'each file); the files are NOT saved. Form files open in the designer were ' +
    'renamed through the designer.');
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Semantic replace (rules from <project>\semantic-replace.json)
// ---------------------------------------------------------------------------

function VerdictName(AVerdict: TMatchVerdict): string;
begin
  case AVerdict of
    mvVerified: Result := 'verified';
    mvOtherSymbol: Result := 'other_symbol';
    mvWrongUnit: Result := 'wrong_unit';
  else
    Result := 'no_answer';
  end;
end;

function ToolSemanticReplace(AArgs: TJSONObject; AStop: THandle): string;
var
  Err, RulesPath, LoadErr: string;
  Rules: TArray<TSemanticReplaceRule>;
  Files: TArray<string>;
  Plans: TArray<TSemanticFilePlan>;
  Dominant: TArray<string>;
  Client: TLspClient;
begin
  var DoApply := AArgs.GetValue<Boolean>('apply', False);
  var IncludeUnverified := AArgs.GetValue<Boolean>('include_unverified', False);
  var MaxRows := AArgs.GetValue<Integer>('max', 100);
  Files := nil;
  if AArgs.GetValue('files') is TJSONArray then
    for var V in TJSONArray(AArgs.GetValue('files')) do
      Files := Files + [ExpandFileName(V.Value)];
  Plans := nil;
  Client := nil;
  if not McpRunOnMain(
    procedure
    begin
      RulesPath := SemanticRulesPath;
      if (RulesPath = '') or not FileExists(RulesPath) then Exit;
      Rules := TSemanticReplaceEngine.LoadRules(RulesPath, LoadErr);
      if Length(Files) = 0 then Files := Editor.GetProjectSourceFiles;
      for var F in Files do
      begin
        var P := Default(TSemanticFilePlan);
        P.FileName := F;
        if not McpReadContent(F, P.Original) then Continue;
        P.Matches := TSemanticReplaceEngine.FindAllMatches(P.Original, Rules);
        if Length(P.Matches) > 0 then Plans := Plans + [P];
      end;
      Client := TLspManager.Instance.PeekClient;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if (RulesPath = '') or not FileExists(RulesPath) then
    Exit(McpErr('no semantic-replace.json in the project root (' + RulesPath +
      ') - create the rules in the IDE (Refactoring Light > Semantic replace > Edit rules)'));
  if LoadErr <> '' then Exit(McpErr('the rules file is invalid: ' + LoadErr));

  var Note := '';
  if Client = nil then
    Note := 'the plugin''s DelphiLSP session is not running - matches are NOT verified'
  else if not VerifySemanticPlans(Client, Plans, Rules,
    function(ACur, ATotal: Integer; const AText: string): Boolean
    begin
      Result := WaitForSingleObject(AStop, 0) <> WAIT_OBJECT_0;
    end, Dominant) then
    Exit(McpErr('cancelled'));

  var Res := TJSONObject.Create;
  Res.AddPair('rules_file', RulesPath);
  if Note <> '' then Res.AddPair('note', Note);
  var RA := TJSONArray.Create;
  for var R := 0 to High(Rules) do
  begin
    var O := TJSONObject.Create;
    O.AddPair('find', Rules[R].Find);
    O.AddPair('replace', Rules[R].Replace);
    if Rules[R].DeclaredIn <> '' then O.AddPair('declaredIn', Rules[R].DeclaredIn);
    if (R <= High(Dominant)) and (Dominant[R] <> '') then O.AddPair('symbol', Dominant[R]);
    RA.Add(O);
  end;
  Res.AddPair('rules', RA);
  var Counts: array[TMatchVerdict] of Integer;
  for var V := Low(TMatchVerdict) to High(TMatchVerdict) do Counts[V] := 0;
  var MA := TJSONArray.Create;
  var Rows := 0;
  for var P in Plans do
    for var I := 0 to High(P.Matches) do
    begin
      var Verdict := mvVerified;
      if Length(P.Verdicts) > 0 then Verdict := P.Verdicts[I];
      Inc(Counts[Verdict]);
      if Rows >= MaxRows then Continue;
      Inc(Rows);
      var L, C: Integer;
      TSemanticReplaceEngine.OffsetToLineCol(P.Original, P.Matches[I].Offset, L, C);
      var O := TJSONObject.Create;
      O.AddPair('file', P.FileName);
      O.AddPair('line', TJSONNumber.Create(L));
      O.AddPair('column', TJSONNumber.Create(C));
      O.AddPair('text', Trim(TSemanticReplaceEngine.LineAtOffset(P.Original, P.Matches[I].Offset)));
      O.AddPair('rule', Rules[P.Matches[I].RuleIdx].Find);
      if Length(P.Verdicts) > 0 then
      begin
        O.AddPair('verdict', VerdictName(Verdict));
        O.AddPair('detail', SemanticVerdictText(Verdict, P.Targets[I],
          Rules[P.Matches[I].RuleIdx]));
      end;
      MA.Add(O);
    end;
  Res.AddPair('verified', TJSONNumber.Create(Counts[mvVerified]));
  Res.AddPair('other_symbol', TJSONNumber.Create(Counts[mvOtherSymbol] + Counts[mvWrongUnit]));
  Res.AddPair('no_answer', TJSONNumber.Create(Counts[mvNoAnswer]));
  Res.AddPair('matches', MA);
  if DoApply then
  begin
    var Replaced := 0;
    var Changed := '';
    if not McpRunOnMain(
      procedure
      begin
        for var P in Plans do
        begin
          var Cur: string;
          if not McpReadContent(P.FileName, Cur) or (Cur <> P.Original) then
          begin
            Changed := Changed + ExtractFileName(P.FileName) + ' ';
            Continue;
          end;
          Inc(Replaced, ApplySemanticPlan(P, Rules, IncludeUnverified));
        end;
      end, False, AStop, Err, 120000) then
    begin
      Res.Free;
      Exit(McpErr(Err));
    end;
    Res.AddPair('replaced', TJSONNumber.Create(Replaced));
    if Changed <> '' then
      Res.AddPair('skipped_changed_files', Trim(Changed));
  end;
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Move to new unit
// ---------------------------------------------------------------------------

function MovePlanToJson(const APlan: TMovePlan): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('source', APlan.SourceFile);
  Result.AddPair('target', APlan.TargetFile);
  Result.AddPair('declaration', TJSONObject.Create
    .AddPair('fromLine', TJSONNumber.Create(APlan.DeclStartLine))
    .AddPair('toLine', TJSONNumber.Create(APlan.DeclEndLine))
    .AddPair('text', APlan.DeclarationText));
  var Impl := TJSONArray.Create;
  for var K := 0 to High(APlan.ImplBlocks) do
    Impl.Add(TJSONObject.Create
      .AddPair('fromLine', TJSONNumber.Create(APlan.ImplStartLines[K]))
      .AddPair('toLine', TJSONNumber.Create(APlan.ImplEndLines[K])));
  Result.AddPair('implementations', Impl);
  var Cons := TJSONArray.Create;
  for var S in APlan.Consumers do Cons.Add(S);
  Result.AddPair('unitsThatGetTheNewUnitInUses', Cons);
  var Drop := TJSONArray.Create;
  for var S in APlan.SourceUsesToRemove do Drop.Add(S);
  Result.AddPair('unitsThatLoseTheSourceUnit', Drop);
  var Ed := TJSONArray.Create;
  for var E in APlan.Edits do Ed.Add(ExtractFileName(E.FilePath) + ': ' + E.Description);
  Result.AddPair('edits', Ed);
end;

// apply=false (the default since 2026-09-24) runs the plan and every
// refusal check, deletes the empty target file again and answers with the
// ranges that WOULD move.
function ToolMoveToNewUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, NewUnit, Err, Token: string;
  L1, C1: Integer;
  Plan: TMovePlan;
  DoApply: Boolean;
begin
  F := ExpandFileName(AArgs.GetValue<string>('file', ''));
  L1 := AArgs.GetValue<Integer>('line', 0);
  C1 := AArgs.GetValue<Integer>('column', 0);
  NewUnit := Trim(AArgs.GetValue<string>('new_unit', ''));
  if (AArgs.GetValue<string>('file', '') = '') or (L1 < 1) or (C1 < 1) or (NewUnit = '') then
    Exit(McpErr('arguments "file", "line", "column" (1-based) and "new_unit" are required'));
  if SameText(ExtractFileExt(NewUnit), '.pas') then NewUnit := ChangeFileExt(NewUnit, '');
  var NewFile := NewUnit;
  if ExtractFilePath(NewFile) = '' then NewFile := ExtractFilePath(F) + NewFile;
  NewFile := NewFile + '.pas';
  DoApply := ArgBool(AArgs, 'apply');
  Token := ArgStr(AArgs, 'token');
  var Ok := False;
  var Ident := '';
  var Msg := '';
  Plan := Default(TMovePlan);
  if not McpRunOnMain(
    procedure
    var
      C: string;
      Col0: Integer;
    begin
      if not McpReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      Ident := IdentifierAtPos(C.Replace(#13#10, #10).Split([#10]), L1 - 1, C1 - 1, Col0);
      if Ident = '' then
      begin
        Msg := 'there is no identifier at that position';
        Exit;
      end;
      if DoApply and (Token <> '') and not CheckPreviewToken(Token, 'move_to_new_unit',
        function(AFile: string): string
        begin
          if not McpReadContent(AFile, Result) then Result := '';
        end, Msg) then Exit;
      Ok := TLspMoveToUnit.ExecuteToNewUnit(Ident, F, NewFile, not DoApply, Plan, Msg);
      if Ok and not DoApply then
        Token := NewPreviewToken('move_to_new_unit', [F], [C]);
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('moved', Ident);
  Res.AddPair('new_unit', NewFile);
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('moves', MovePlanToJson(Plan));
  if Msg <> '' then Res.AddPair('note', Msg);
  if DoApply then
    Res.AddPair('saved', 'The new unit was created on disk and added to the project; ' +
      'its content and the edits of the other units are in the IDE buffers (not ' +
      'saved) - units that are not open in the IDE were changed on disk.')
  else
  begin
    Res.AddPair('token', Token);
    Res.AddPair('note2', 'Nothing was written and no file was created. Call again ' +
      'with apply=true (and this token) to move.');
  end;
  Result := McpOk(Res);
end;


// ---------------------------------------------------------------------------
//  Statement refactorings: extract variable / wrap in try..finally
// ---------------------------------------------------------------------------
//
// Both planners are pure and were written for the editor entry points; the
// tools below feed them the same way but take the selection as arguments
// (user request 2026-09-29: every refactoring reachable through the bridge,
// with the preview of 1.12.0).

// One place for "a planner produced new content -> preview or write".
function ContentResult(const AFile, AOld, ANewContent, ATool: string;
  ADoApply: Boolean; const AToken: string; AExtra: TJSONObject): string;
var
  SL: TStringList;
  NewContent, Problem, Token: string;
begin
  Token := AToken;
  NewContent := ANewContent;
  SL := TStringList.Create;
  try
    if ADoApply then
    begin
      if (Token <> '') and not CheckPreviewToken(Token, ATool,
        function(AF: string): string
        begin
          if not McpReadContent(AF, Result) then Result := '';
        end, Problem) then
      begin
        AExtra.Free;
        Exit(McpErr(Problem));
      end;
      SL.Text := NewContent;
      if not ApplyLinesMinimal(AFile, SL, AOld) then
      begin
        AExtra.Free;
        Exit(McpErr('the change could not be written'));
      end;
    end
    else
      Token := NewPreviewToken(ATool, [AFile], [AOld]);
  finally
    SL.Free;
  end;
  var Res := AExtra;
  if Res = nil then Res := TJSONObject.Create;
  Res.AddPair('file', AFile);
  Res.AddPair('applied', TJSONBool.Create(ADoApply));
  var Total := 0;
  var Ch := DiffToChanges(AFile, AOld, NewContent, 40, Total);
  Res.AddPair('changes', ChangesToJson(Ch));
  if Total > Length(Ch) then
    Res.AddPair('changesTruncated', TJSONNumber.Create(Total));
  if ADoApply then
    Res.AddPair('note', 'Changed in the IDE buffer when the file is open (not ' +
      'saved), otherwise on disk.')
  else
  begin
    Res.AddPair('token', Token);
    Res.AddPair('note', 'Nothing was written. Call again with apply=true (and ' +
      'this token) to make the change.');
  end;
  Result := McpOk(Res);
end;

// A planner that works on LINES: joined the way SplitContentLines splits,
// so a trailing empty element stays the file's final line break.
function LinesResult(const AFile, AOld: string; const ANewLines: TArray<string>;
  const ATool: string; ADoApply: Boolean; const AToken: string;
  AExtra: TJSONObject): string;
var
  LB: string;
begin
  if Pos(#13#10, AOld) > 0 then LB := #13#10
  else if Pos(#10, AOld) > 0 then LB := #10
  else LB := sLineBreak;
  Result := ContentResult(AFile, AOld, string.Join(LB, ANewLines), ATool,
    ADoApply, AToken, AExtra);
end;

function ToolExtractVariable(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Expr, Name, Err, Msg: string;
  L1, Col1: Integer;
  DoApply: Boolean;
  Plan: TExtractVarPlan;
  Planned: Boolean;
begin
  F := ArgStr(AArgs, 'file');
  L1 := ArgInt(AArgs, 'line');
  Expr := ArgStr(AArgs, 'expression');
  Col1 := ArgInt(AArgs, 'column');
  Name := Trim(ArgStr(AArgs, 'name'));
  var EndCol1 := ArgInt(AArgs, 'end_column');
  if (F = '') or (L1 < 1) then
    Exit(McpErr('arguments "file" and "line" (1-based) are required'));
  if (Expr = '') and ((Col1 < 1) or (EndCol1 <= Col1)) then
    Exit(McpErr('pass "expression" (the text to extract) or "column" + ' +
      '"end_column" (1-based, end exclusive)'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Msg := '';
  Planned := False;
  var Content := '';
  var Why := '';
  if not McpRunOnMain(
    procedure
    begin
      if not McpReadContent(F, Content) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      var Lines := Content.Replace(#13#10, #10).Split([#10]);
      if L1 > Length(Lines) then
      begin
        Msg := Format('line %d is beyond the file (%d lines)', [L1, Length(Lines)]);
        Exit;
      end;
      var LineText := Lines[L1 - 1];
      var Start0: Integer;
      var End0: Integer;
      if Expr <> '' then
      begin
        // locate the expression on the line - at "column" when given,
        // else its first occurrence
        var P := 0;
        if Col1 >= 1 then P := Pos(Expr, LineText, Col1);
        if P = 0 then P := Pos(Expr, LineText);
        if P = 0 then
        begin
          Msg := 'the expression is not on line ' + IntToStr(L1);
          Exit;
        end;
        Start0 := P - 1;
        End0 := Start0 + Length(Expr);
      end
      else
      begin
        Start0 := Col1 - 1;
        End0 := EndCol1 - 1;
        if End0 > Length(LineText) then
        begin
          Msg := 'end_column is beyond the line';
          Exit;
        end;
        Expr := Copy(LineText, Start0 + 1, End0 - Start0);
      end;
      if Name = '' then Name := SuggestVariableName(Expr);
      if not IsValidIdent(Name) then
      begin
        Msg := '"' + Name + '" is not a valid identifier';
        Exit;
      end;
      Planned := PlanExtractVariable(Lines, L1 - 1, Start0, End0, Name, Plan, Why);
    end, True, AStop, Err) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  if not Planned then
    Exit(McpErr('extract variable is not possible here: ' + Why));
  var Extra := TJSONObject.Create;
  Extra.AddPair('name', Name);
  Extra.AddPair('expression', Expr);
  Extra.AddPair('declaredBeforeLine', TJSONNumber.Create(Plan.StatementLine + 1));
  Extra.AddPair('declaration', Trim(Plan.DeclText));
  Result := LinesResult(F, Content, Plan.NewLines, 'extract_variable', DoApply,
    ArgStr(AArgs, 'token'), Extra);
end;

function ToolWrapTryFinally(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Cleanup, Err, Msg, Why, Content: string;
  From1, To1: Integer;
  DoApply, Planned, HadCleanup: Boolean;
  NewLines: TArray<string>;
begin
  F := ArgStr(AArgs, 'file');
  From1 := ArgInt(AArgs, 'from_line');
  To1 := ArgInt(AArgs, 'to_line', From1);
  Cleanup := Trim(ArgStr(AArgs, 'cleanup'));
  HadCleanup := Cleanup <> '';
  if (F = '') or (From1 < 1) or (To1 < From1) then
    Exit(McpErr('arguments "file" and "from_line" (1-based) are required; ' +
      '"to_line" defaults to from_line'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Msg := '';
  Planned := False;
  Content := '';
  Why := '';
  if not McpRunOnMain(
    procedure
    begin
      if not McpReadContent(F, Content) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      var Lines := Content.Replace(#13#10, #10).Split([#10]);
      if To1 > Length(Lines) then
      begin
        Msg := Format('to_line %d is beyond the file (%d lines)', [To1, Length(Lines)]);
        Exit;
      end;
      if not HadCleanup then
        // the statement before the range names what to release
        for var I := From1 - 2 downto 0 do
          if Trim(Lines[I]) <> '' then
          begin
            Cleanup := InferCleanup(Lines[I]);
            Break;
          end;
      Planned := PlanWrapTryFinally(Lines, From1 - 1, To1 - 1, Cleanup, NewLines, Why);
    end, True, AStop, Err) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  if not Planned then
    Exit(McpErr('wrap in try..finally is not possible here: ' + Why));
  var Extra := TJSONObject.Create;
  if Cleanup <> '' then
  begin
    Extra.AddPair('cleanup', Cleanup);
    if not HadCleanup then
      Extra.AddPair('cleanupFrom', 'inferred from the statement before the range - ' +
        'pass "cleanup" to override it');
  end
  else
    Extra.AddPair('cleanup', 'none - a TODO comment is inserted instead; pass ' +
      '"cleanup" with the statement that releases what the block acquires');
  Result := LinesResult(F, Content, NewLines, 'wrap_try_finally', DoApply,
    ArgStr(AArgs, 'token'), Extra);
end;

// apply=false reports the plan only (user request 2026-09-29).
function ToolMoveToUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Target, Err, Token: string;
  L1, C1: Integer;
  Plan: TMovePlan;
  DoApply: Boolean;
begin
  F := ExpandFileName(ArgStr(AArgs, 'file'));
  L1 := ArgInt(AArgs, 'line');
  C1 := ArgInt(AArgs, 'column');
  Target := ArgStr(AArgs, 'target_file');
  if (ArgStr(AArgs, 'file') = '') or (L1 < 1) or (C1 < 1) or (Target = '') then
    Exit(McpErr('arguments "file", "line", "column" (1-based) and ' +
      '"target_file" are required'));
  if ExtractFilePath(Target) = '' then Target := ExtractFilePath(F) + Target;
  if SameText(ExtractFileExt(Target), '') then Target := Target + '.pas';
  Target := ExpandFileName(Target);
  DoApply := ArgBool(AArgs, 'apply');
  Token := ArgStr(AArgs, 'token');
  var Ok := False;
  var Ident := '';
  var Msg := '';
  Plan := Default(TMovePlan);
  if not McpRunOnMain(
    procedure
    var
      C: string;
      Col0: Integer;
    begin
      if not McpReadContent(F, C) then
      begin
        Msg := 'file not found: ' + F;
        Exit;
      end;
      Ident := IdentifierAtPos(C.Replace(#13#10, #10).Split([#10]), L1 - 1, C1 - 1, Col0);
      if Ident = '' then
      begin
        Msg := 'there is no identifier at that position';
        Exit;
      end;
      if DoApply and (Token <> '') and not CheckPreviewToken(Token, 'move_to_unit',
        function(AFile: string): string
        begin
          if not McpReadContent(AFile, Result) then Result := '';
        end, Msg) then Exit;
      Ok := TLspMoveToUnit.ExecuteToExistingUnit(Ident, F, Target, not DoApply,
        Plan, Msg);
      if Ok and not DoApply then
        Token := NewPreviewToken('move_to_unit', [F], [C]);
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('moved', Ident);
  Res.AddPair('target', Target);
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('moves', MovePlanToJson(Plan));
  if Msg <> '' then Res.AddPair('note', Msg);
  if DoApply then
    Res.AddPair('saved', 'The edits are in the IDE buffers (not saved) for ' +
      'units open in the IDE, on disk for the others.')
  else
  begin
    Res.AddPair('token', Token);
    Res.AddPair('note2', 'Nothing was written. Call again with apply=true (and ' +
      'this token) to move.');
  end;
  Result := McpOk(Res);
end;


// ---------------------------------------------------------------------------
//  Uses cleanup: the dialog's verdicts as an edit
// ---------------------------------------------------------------------------
//
// analyze_uses reported the verdicts and left the caller to rebuild the
// edit out of remove_unit / add_unit calls. This runs what the dialog runs
// (user request 2026-09-29): UNUSED entries are removed, MOVABLE ones move
// to the implementation uses - each through the same pure planners the
// uses editor writes with, chained over one content so the preview is the
// whole change. Anything the analysis is unsure about stays untouched, and
// the answer says why per entry.
function ToolCleanupUses(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, C: string;
  Found, DoApply, DoRemove, DoMove, DoUnverified: Boolean;
  Cycle0: Integer;
  Entries: TArray<TUsesEntryInfo>;
  St: TDesignerState;
begin
  F := ArgStr(AArgs, 'file');
  if F = '' then Exit(McpErr('argument "file" is required'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  DoRemove := ArgBool(AArgs, 'remove_unused', True);
  DoMove := ArgBool(AArgs, 'move_to_implementation', False);
  // Issue #20: in a form unit whose form is NOT loaded, nothing can say
  // which units the designer would write back. Inside the IDE such an entry
  // simply reappears; on a command-line or CI build nothing re-adds it - at
  // best that is a compile error, at worst a form that streams a class
  // nobody registered any more. So it takes an explicit opt-in.
  DoUnverified := ArgBool(AArgs, 'include_unverified', False);
  Found := False;
  Cycle0 := TUnitIndex.Instance.ScanCycle;
  if not McpRunOnMain(
    procedure
    begin
      Found := McpReadContent(F, C);
      TUnitIndex.Instance.RefreshSourcesFromEditor;
      QueryDesigner(F, St);
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Found then Exit(McpErr('file not found: ' + F));
  // the index parses from DISK - wait for one full cycle like the dialog
  var Waited := 0;
  while (TUnitIndex.Instance.ScanCycle = Cycle0) and (Waited < 5000) do
  begin
    if WaitForSingleObject(AStop, 50) = WAIT_OBJECT_0 then Exit(McpErr('shutting down'));
    Inc(Waited, 50);
  end;
  var Snap := TUnitIndex.Instance.Snapshot;
  if (Snap = nil) or (Snap.IdentCount = 0) then
    Exit(McpErr('the identifier index is not ready yet - see get_status'));
  Entries := AnalyzeUsesWithDesigner(C, F, Snap, St,
    TPluginSettings.UsesCleanupKeepUnits);

  var Content := C;
  var Rows := TJSONArray.Create;
  var Removed := 0;
  var Moved := 0;
  var Skipped := 0;
  for var E0 in Entries do
  begin
    var E := E0;
    var Row := TJSONObject.Create;
    Row.AddPair('unit', E.UnitName);
    Row.AddPair('verdict', VerdictNames[E.Verdict]);
    if E.Reason <> '' then Row.AddPair('reason', E.Reason);
    // An unverified row acts on what the TEXT says, but only when the
    // caller has taken responsibility for it.
    if E.Verdict = uvUnverified then
      if DoUnverified then
        E.Verdict := ResolveUnverified(E)
      else
        Inc(Skipped);
    var Action := 'kept';
    var Next := '';
    if E.Verdict = uvUnverified then
      Action := 'kept: unverified - the form is not loaded, so nothing could ' +
        'confirm the IDE would not re-add it (pass include_unverified=true)'
    else if (E.Verdict = uvUnused) and DoRemove then
    begin
      if PlanRemoveUnitFromUsesText(Content, E.UnitName, Next) then
      begin
        Content := Next;
        Action := 'removed';
        Inc(Removed);
      end
      else
        Action := 'kept: the uses clause could not be edited safely';
    end
    else if (E.Verdict = uvMovable) and DoMove then
    begin
      // remove from the interface, add to the implementation - the same two
      // steps the dialog does, and both must succeed
      if PlanRemoveUnitFromUsesText(Content, E.UnitName, Next)
        and PlanAddUnitToUsesText(Next, E.UnitName, usImplementation, Next) then
      begin
        Content := Next;
        Action := 'moved to the implementation uses';
        Inc(Moved);
      end
      else
        Action := 'kept: the move could not be planned safely';
    end
    else if E.Verdict = uvUnused then
      Action := 'kept: remove_unused is off'
    else if E.Verdict = uvMovable then
      Action := 'kept: move_to_implementation is off'
    else if E.Verdict = uvIdeManaged then
      Action := 'kept: the form designer writes this entry itself'
    else if E.Verdict = uvKeptByUser then
      Action := 'kept: on your keep list';
    Row.AddPair('action', Action);
    Rows.Add(Row);
  end;

  var Extra := TJSONObject.Create;
  Extra.AddPair('entries', Rows);
  Extra.AddPair('removed', TJSONNumber.Create(Removed));
  Extra.AddPair('movedToImplementation', TJSONNumber.Create(Moved));
  Extra.AddPair('designerVerified', TJSONBool.Create(St.Verified));
  Extra.AddPair('designerState', St.StateText);
  if Skipped > 0 then
    Extra.AddPair('unverifiedKept', TJSONNumber.Create(Skipped));
  if Content = C then
  begin
    Extra.AddPair('file', F);
    Extra.AddPair('applied', TJSONBool.Create(False));
    Extra.AddPair('changes', TJSONArray.Create);
    Extra.AddPair('note', 'Nothing to clean up with these options. init_code, ' +
      'ide_managed, kept_by_user, unverified and unknown entries are kept ' +
      'deliberately; class helpers, operators and initialization side ' +
      'effects are invisible to a textual analysis.');
    Exit(McpOk(Extra));
  end;
  Result := ContentResult(F, C, Content, 'cleanup_uses', DoApply,
    ArgStr(AArgs, 'token'), Extra);
end;


// ---------------------------------------------------------------------------
//  Remove with
// ---------------------------------------------------------------------------
//
// One occurrence (file + line), a file, a list or the whole project. The
// rewrite itself is the wizard's - only the dialog is replaced by this
// answer (user request 2026-09-29). apply=false is the default; an
// occurrence the rewriter cannot handle is listed with the reason, never
// rewritten half way.
function ToolRemoveWith(AArgs: TJSONObject; AStop: THandle): string;
var
  Files: TArray<string>;
  Results: TArray<TWithRewriteResult>;
  F, Err, RunErr, Token: string;
  L1, Applied, Failed, SkippedNested: Integer;
  DoApply, Inline_, Ok: Boolean;
begin
  F := ArgStr(AArgs, 'file');
  if F <> '' then F := ExpandFileName(F);
  L1 := ArgInt(AArgs, 'line');
  Files := nil;
  if F <> '' then Files := [F];
  var Arr := AArgs.GetValue<TJSONArray>('files', nil);
  if Arr <> nil then
    for var V in Arr do Files := Files + [ExpandFileName(V.Value)];
  var Project := ArgBool(AArgs, 'project');
  if (Length(Files) = 0) and not Project then
    Exit(McpErr('pass "file" (optionally with "line" for a single ' +
      'with-statement), "files" or "project": true'));
  DoApply := ArgBool(AArgs, 'apply');
  Inline_ := ArgBool(AArgs, 'inline_vars', True);
  Token := ArgStr(AArgs, 'token');
  Ok := False;
  RunErr := '';
  Applied := 0;
  Failed := 0;
  SkippedNested := 0;
  var Contents: TArray<string> := nil;
  if not McpRunOnMain(
    procedure
    begin
      var All := Files;
      if Project and (Editor <> nil) then All := All + Editor.GetProjectSourceFiles;
      if DoApply and (Token <> '') and not CheckPreviewToken(Token, 'remove_with',
        function(AFile: string): string
        begin
          if not McpReadContent(AFile, Result) then Result := '';
        end, RunErr) then Exit;
      Ok := RunRemoveWithHeadless(All, F, L1, DoApply, Inline_, Results,
        Applied, Failed, SkippedNested, RunErr);
      if Ok and not DoApply then
      begin
        // the token pins the files the preview actually describes
        var Seen := TStringList.Create;
        try
          Seen.CaseSensitive := False;
          for var R in Results do
            if R.IsAutoRewritable and (Seen.IndexOf(R.FileName) < 0) then
              Seen.Add(R.FileName);
          var Pin: TArray<string> := nil;
          for var I := 0 to Seen.Count - 1 do
          begin
            var C: string;
            if McpReadContent(Seen[I], C) then
            begin
              Pin := Pin + [Seen[I]];
              Contents := Contents + [C];
            end;
          end;
          if Length(Pin) > 0 then Token := NewPreviewToken('remove_with', Pin, Contents);
        finally
          Seen.Free;
        end;
      end;
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(RunErr));
  var Rows := TJSONArray.Create;
  var Rewritable := 0;
  for var R in Results do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', R.FileName);
    O.AddPair('line', TJSONNumber.Create(R.Occurrence.KeywordPos.Line));
    O.AddPair('column', TJSONNumber.Create(R.Occurrence.KeywordPos.Col));
    var Targets := TJSONArray.Create;
    for var Tg in R.Occurrence.Targets do Targets.Add(Tg.Expression);
    O.AddPair('targets', Targets);
    if R.IsAutoRewritable then
    begin
      Inc(Rewritable);
      O.AddPair('rewritable', TJSONBool.Create(True));
      O.AddPair('before', R.OriginalText);
      O.AddPair('after', R.NewText);
    end
    else
    begin
      O.AddPair('rewritable', TJSONBool.Create(False));
      O.AddPair('reason', WithRewriteIssueText(R.Issues));
    end;
    Rows.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('found', TJSONNumber.Create(Length(Results)));
  Res.AddPair('rewritable', TJSONNumber.Create(Rewritable));
  Res.AddPair('occurrences', Rows);
  if DoApply then
  begin
    Res.AddPair('written', TJSONNumber.Create(Applied));
    if Failed > 0 then Res.AddPair('failed', TJSONNumber.Create(Failed));
    if SkippedNested > 0 then
      Res.AddPair('skippedEnclosing', TJSONNumber.Create(SkippedNested));
    Res.AddPair('note', 'Changed in the IDE buffers (not saved) for open units, ' +
      'on disk for the others.' + IfThen(SkippedNested > 0,
      ' An enclosing with-statement that contains another rewritten one is ' +
      'skipped - call again for it.', ''));
  end
  else
  begin
    if Token <> '' then Res.AddPair('token', Token);
    Res.AddPair('note', 'Nothing was written. "before"/"after" is the whole ' +
      'with-statement as it would be replaced; call again with apply=true ' +
      '(and this token) to rewrite the rewritable ones.');
  end;
  Result := McpOk(Res);
end;


// ---------------------------------------------------------------------------
//  Find original symbol
// ---------------------------------------------------------------------------
//
// Read-only. The chain is the menu entry's (DelphiLSP, then the qualifier's
// type, then the identifier index) - ResolveOriginalSymbol is shared, so an
// answer here is the place the menu would jump to (user request 2026-09-29).
// lsp_definition alone is NOT the same: it stops where DelphiLSP is silent,
// which is exactly the RSS-5463 overload case this resolves.
function ToolFindOriginalSymbol(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Note, Ident: string;
  L1, C1: Integer;
  Hits: TArray<TOriginalSymbolHit>;
  Ctx: TPosContext;
begin
  if not RequireFilePos(AArgs, F, L1, C1, Err) then Exit(McpErr(Err));
  if not GatherPosContext(F, L1, C1, True, AStop, Ctx, Err) then Exit(McpErr(Err));
  Ident := Ctx.Identifier;
  Note := '';
  Hits := nil;
  if not McpRunOnMain(
    procedure
    begin
      ResolveOriginalSymbol(Ctx.Client, F, L1 - 1, Ctx.IdentCol0, Ident,
        Hits, Note);
    end, True, AStop, Err, 120000) then Exit(McpErr(Err));
  var Res := TJSONObject.Create;
  Res.AddPair('identifier', Ident);
  var Arr := TJSONArray.Create;
  for var H in Hits do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', H.FilePath);
    O.AddPair('line', TJSONNumber.Create(H.Line + 1));
    O.AddPair('column', TJSONNumber.Create(H.Col + 1));
    O.AddPair('via', H.Via);
    if H.TypeName <> '' then O.AddPair('type', H.TypeName);
    Arr.Add(O);
  end;
  Res.AddPair('declarations', Arr);
  if Length(Hits) = 0 then
  begin
    if Note = '' then Note := 'no declaration found, and the identifier index ' +
      'does not know it either';
    Res.AddPair('note', Note);
  end
  else if Length(Hits) > 1 then
    Res.AddPair('note', 'several units declare this identifier - the menu entry ' +
      'hands this case to the Find-Unit dialog; find_unit lists the same ' +
      'candidates with their declarations')
  else if Note <> '' then
    Res.AddPair('note', Note);
  Result := McpOk(Res);
end;


// ---------------------------------------------------------------------------
//  Find unit references
// ---------------------------------------------------------------------------
//
// Which units use the given unit, and WHERE - every hit verified with
// DelphiLSP, plus one "(unused)" row per unit that lists it in its uses
// clause without referencing anything of it. Read-only. Shares the wizard's
// search (user request 2026-09-29).
function ToolFindUnitReferences(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Status, RunErr: string;
  Items: TUnitRefItems;
  Ok: Boolean;
  Max: Integer;
begin
  F := ArgStr(AArgs, 'file');
  if F = '' then Exit(McpErr('argument "file" is required (the unit whose ' +
    'references you want)'));
  F := ExpandFileName(F);
  Max := ArgInt(AArgs, 'max', 400);
  Ok := False;
  RunErr := '';
  Status := '';
  if not McpRunOnMain(
    procedure
    begin
      if not FileExists(F) then
      begin
        RunErr := 'file not found: ' + F;
        Exit;
      end;
      Ok := FindUnitReferencesHeadless(F, Items, Status, RunErr);
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then
  begin
    if RunErr = '' then RunErr := 'the search did not finish';
    Exit(McpErr(RunErr));
  end;
  var Arr := TJSONArray.Create;
  var Dead := 0;
  var Shown := 0;
  for var It in Items do
  begin
    if It.IsDead then Inc(Dead);
    if (Max > 0) and (Shown >= Max) then Continue;
    Inc(Shown);
    var O := TJSONObject.Create;
    O.AddPair('file', It.FilePath);
    if It.IsDead then
      O.AddPair('unused', TJSONBool.Create(True))
    else
    begin
      O.AddPair('identifier', It.Identifier);
      O.AddPair('line', TJSONNumber.Create(It.Line + 1));
      O.AddPair('column', TJSONNumber.Create(It.Col + 1));
    end;
    O.AddPair('preview', Trim(It.Preview));
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('unit', ChangeFileExt(ExtractFileName(F), ''));
  Res.AddPair('total', TJSONNumber.Create(Length(Items)));
  Res.AddPair('unusedEntries', TJSONNumber.Create(Dead));
  Res.AddPair('references', Arr);
  if Shown < Length(Items) then
    Res.AddPair('truncated', TJSONNumber.Create(Length(Items) - Shown));
  if Status <> '' then Res.AddPair('summary', Status);
  Res.AddPair('note', 'A row with "unused": true means the unit is in that ' +
    'file''s uses clause but nothing of it is referenced there. Every other ' +
    'row was verified with DelphiLSP (its definition leads into this unit).');
  Result := McpOk(Res);
end;

// ---------------------------------------------------------------------------
//  Project checks: interface GUIDs, DFM event handlers
// ---------------------------------------------------------------------------

function ToolInterfaceGuids(AArgs: TJSONObject; AStop: THandle): string;
var
  Entries: TArray<TInterfaceGuidEntry>;
  Err: string;
  Files: TArray<string>;
begin
  var OnlyProblems := ArgBool(AArgs, 'only_problems', True);
  Files := nil;
  if not McpRunOnMain(
    procedure
    begin
      if Editor <> nil then Files := Editor.GetProjectSourceFiles;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if Length(Files) = 0 then Exit(McpErr('no project loaded / no source files'));
  Entries := TInterfaceGuidChecker.Scan(Files);
  var Arr := TJSONArray.Create;
  var Dupes := 0;
  var Missing := 0;
  for var E in Entries do
  begin
    if E.IsDuplicate then Inc(Dupes);
    if not E.HasGuid then Inc(Missing);
    if OnlyProblems and not E.IsDuplicate and E.HasGuid then Continue;
    var O := TJSONObject.Create;
    O.AddPair('interface', E.InterfaceName);
    O.AddPair('file', E.FileName);
    O.AddPair('line', TJSONNumber.Create(E.Line));
    if E.HasGuid then O.AddPair('guid', E.Guid)
    else O.AddPair('guid', TJSONNull.Create);
    if E.IsDuplicate then O.AddPair('duplicate', TJSONBool.Create(True));
    if E.IsDispInterface then O.AddPair('dispinterface', TJSONBool.Create(True));
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('interfaces', TJSONNumber.Create(Length(Entries)));
  Res.AddPair('duplicateGuids', TJSONNumber.Create(Dupes));
  Res.AddPair('withoutGuid', TJSONNumber.Create(Missing));
  Res.AddPair('entries', Arr);
  Res.AddPair('note', 'A duplicate GUID makes Supports/QueryInterface return the ' +
    'WRONG object. An interface paired with a dispinterface on the same GUID is ' +
    'a type-library import and not counted. Read-only.' +
    IfThen(OnlyProblems, ' Only problems are listed - pass only_problems=false ' +
    'for every interface.', ''));
  Result := McpOk(Res);
end;

// One id per issue so apply can name them; bound to the place, not to an
// index, so it survives a re-check as long as the issue does.
function DfmIssueId(const AIssue: TDfmEventIssue): string;
begin
  Result := IntToHex(PreviewContentHash(LowerCase(AIssue.DfmFile) + '|' +
    AIssue.ComponentName + '|' + AIssue.EventName), 8);
end;

function ToolDfmEvents(AArgs: TJSONObject; AStop: THandle): string;
var
  Issues: TArray<TDfmEventIssue>;
  Files, SigFiles, WantIds: TArray<string>;
  Err, RunErr: string;
  DoApply: Boolean;
begin
  DoApply := ArgBool(AArgs, 'apply');
  WantIds := nil;
  var IdArr := AArgs.GetValue<TJSONArray>('fix_ids', nil);
  if IdArr <> nil then
    for var V in IdArr do WantIds := WantIds + [UpperCase(Trim(V.Value))];
  if DoApply and (Length(WantIds) = 0) then
    Exit(McpErr('apply needs "fix_ids" - list the issues first and pick the ' +
      'ones to fix (a generated empty handler can shadow an inherited one)'));
  Files := nil;
  SigFiles := nil;
  RunErr := '';
  if not McpRunOnMain(
    procedure
    begin
      if Editor = nil then Exit;
      Files := Editor.GetProjectSourceFiles;
      SigFiles := GatherSignatureFiles;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if Length(Files) = 0 then Exit(McpErr('no project loaded / no source files'));
  Issues := TDfmEventChecker.CheckProject(Files, nil, SigFiles);

  var Applied := TDictionary<string, string>.Create;   // id -> result
  try
    if DoApply then
      if not McpRunOnMain(
        procedure
        begin
          var Ctx := TFixContext.Create;
          try
            for var Iss in Issues do
            begin
              var Id := UpperCase(DfmIssueId(Iss));
              var Wanted := False;
              for var W in WantIds do
                if W = Id then Wanted := True;
              if not Wanted then Continue;
              var Reason := '';
              if TDfmEventChecker.ApplyFix(Iss, Reason, Ctx) then
                Applied.AddOrSetValue(Id, 'fixed')
              else
                Applied.AddOrSetValue(Id, 'NOT fixed: ' + Reason);
            end;
          finally
            Ctx.Free;
          end;
        end, False, AStop, Err, 300000) then Exit(McpErr(Err));

    var Arr := TJSONArray.Create;
    var Missing := 0;
    var Mismatch := 0;
    for var Iss in Issues do
    begin
      if Iss.Kind = eikMissingHandler then Inc(Missing) else Inc(Mismatch);
      var O := TJSONObject.Create;
      var Id := DfmIssueId(Iss);
      O.AddPair('id', Id);
      if Iss.Kind = eikMissingHandler then O.AddPair('kind', 'missing_handler')
      else O.AddPair('kind', 'signature_mismatch');
      O.AddPair('form', Iss.DfmFile);
      O.AddPair('unit', Iss.PasFile);
      O.AddPair('component', Iss.ComponentName);
      O.AddPair('componentType', Iss.ComponentType);
      O.AddPair('event', Iss.EventName);
      O.AddPair('handler', Iss.HandlerName);
      O.AddPair('formLine', TJSONNumber.Create(Iss.DfmLine));
      if Iss.PasLine > 0 then O.AddPair('unitLine', TJSONNumber.Create(Iss.PasLine));
      if Iss.Expected <> '' then O.AddPair('expected', Iss.Expected);
      if Iss.Actual <> '' then O.AddPair('actual', Iss.Actual);
      O.AddPair('fixable', TJSONBool.Create(
        (Iss.Kind = eikMissingHandler) or (Iss.ExpectedRawParams <> '')));
      var R: string;
      if Applied.TryGetValue(UpperCase(Id), R) then O.AddPair('result', R);
      Arr.Add(O);
    end;
    var Res := TJSONObject.Create;
    Res.AddPair('applied', TJSONBool.Create(DoApply));
    Res.AddPair('missingHandlers', TJSONNumber.Create(Missing));
    Res.AddPair('signatureMismatches', TJSONNumber.Create(Mismatch));
    Res.AddPair('issues', Arr);
    if DoApply then
      Res.AddPair('note', 'Fixed handlers are in the IDE buffers (not saved). ' +
        'A missing handler is generated as an EMPTY method - in an inherited ' +
        'form that shadows the ancestor''s handler, which is why nothing is ' +
        'fixed without naming its id.')
    else
      Res.AddPair('note', 'Nothing was changed. Pass apply=true with "fix_ids" ' +
        'to generate the missing handlers / correct the parameter lists.');
    Result := McpOk(Res);
  finally
    Applied.Free;
  end;
end;


// ---------------------------------------------------------------------------
//  Align method signature
// ---------------------------------------------------------------------------
//
// Every declaration and implementation of the method at the position -
// interface, class, implementation header - with the one they disagree
// about. apply=true aligns the divergent ones to the majority signature
// through the wizard's own step, in the order the rules require (a class
// declaration before its implementation).
function ToolSignatureCheck(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Method: string;
  L1, C1: Integer;
  Ctx: TPosContext;
  Entries: TSignatureEntries;
  DoApply: Boolean;
  Results: TDictionary<Integer, string>;
begin
  if not RequireFilePos(AArgs, F, L1, C1, Err) then Exit(McpErr(Err));
  if not GatherPosContext(F, L1, C1, True, AStop, Ctx, Err) then Exit(McpErr(Err));
  if Ctx.Client = nil then
    Exit(McpErr('no DelphiLSP session - the signature check resolves the ' +
      'declarations through it (see get_status)'));
  Method := Ctx.Identifier;
  DoApply := ArgBool(AArgs, 'apply');
  Entries := nil;
  if not McpRunOnMain(
    procedure
    begin
      Entries := TSignatureChecker.Collect(Ctx.Client, F, Method);
    end, True, AStop, Err, 300000) then Exit(McpErr(Err));
  if Length(Entries) = 0 then
    Exit(McpErr('no declaration of "' + Method + '" was found'));

  var Reference := TSignatureChecker.PickReference(Entries);
  var Ref: TSignatureEntry;
  var HaveRef := TSignatureChecker.ReferenceEntry(Entries, Reference, Ref);
  Results := TDictionary<Integer, string>.Create;
  try
    if DoApply and HaveRef then
      if not McpRunOnMain(
        procedure
        begin
          // class declarations first: an implementation is aligned with the
          // declaration of ITS unit, which must be right before it
          for var Pass := 0 to 1 do
            for var I := 0 to High(Entries) do
            begin
              if (Pass = 0) <> (Entries[I].Role in [srInterfaceDecl, srClassDecl]) then
                Continue;
              if Entries[I].Normalized = Reference then Continue;
              var Blocker := TSignatureChecker.AlignBlocker(Entries, I, Reference);
              if Blocker <> '' then
              begin
                Results.AddOrSetValue(I, 'not aligned: ' + Blocker);
                Continue;
              end;
              var Why := AlignSignatureEntry(Entries[I], Ref);
              if Why = '' then Results.AddOrSetValue(I, 'aligned')
              else Results.AddOrSetValue(I, 'not aligned: ' + Why);
            end;
        end, False, AStop, Err, 300000) then Exit(McpErr(Err));

    var Arr := TJSONArray.Create;
    var Diverging := 0;
    for var I := 0 to High(Entries) do
    begin
      var E := Entries[I];
      var O := TJSONObject.Create;
      O.AddPair('role', TSignatureChecker.RoleToString(E.Role));
      O.AddPair('container', E.Container);
      O.AddPair('file', E.FilePath);
      O.AddPair('line', TJSONNumber.Create(E.Line + 1));
      O.AddPair('signature', E.RawSignature);
      var Matches := E.Normalized = Reference;
      O.AddPair('matches', TJSONBool.Create(Matches));
      if not Matches then
      begin
        Inc(Diverging);
        var B := TSignatureChecker.AlignBlocker(Entries, I, Reference);
        if B <> '' then O.AddPair('blocked', B);
      end;
      var R: string;
      if Results.TryGetValue(I, R) then O.AddPair('result', R);
      Arr.Add(O);
    end;
    var Res := TJSONObject.Create;
    Res.AddPair('method', Method);
    Res.AddPair('applied', TJSONBool.Create(DoApply));
    Res.AddPair('entries', Arr);
    Res.AddPair('diverging', TJSONNumber.Create(Diverging));
    if HaveRef then
      Res.AddPair('reference', TJSONObject.Create
        .AddPair('file', Ref.FilePath)
        .AddPair('line', TJSONNumber.Create(Ref.Line + 1))
        .AddPair('signature', Ref.RawSignature));
    if Diverging = 0 then
      Res.AddPair('note', 'Every declaration and implementation agrees.')
    else if DoApply then
      Res.AddPair('note', 'Changed in the IDE buffers (not saved). An ' +
        'implementation header keeps its parameter NAMES - the body uses them.')
    else
      Res.AddPair('note', 'Nothing was changed. apply=true aligns the ' +
        'diverging entries with the reference signature (class declarations ' +
        'first, then the implementations).');
    Result := McpOk(Res);
  finally
    Results.Free;
  end;
end;


// ---------------------------------------------------------------------------
//  Extract interface
// ---------------------------------------------------------------------------
//
// apply=false answers with the interface text that WOULD be written - which
// is the actual review question here ("are these the right members?"), not a
// line diff: the change creates a unit (or extends one) and adds the
// interface to the class' ancestor list.
function ToolExtractInterface(AArgs: TJSONObject; AStop: THandle): string;
var
  F, IntfName, Target, Err, RunErr, Text: string;
  L1: Integer;
  Members: TArray<string>;
  AddExisting, DoApply, Ok: Boolean;
  Info: TExtractInterfaceInfo;
begin
  F := ArgStr(AArgs, 'file');
  L1 := ArgInt(AArgs, 'line');
  if (F = '') or (L1 < 1) then
    Exit(McpErr('arguments "file" and "line" (1-based, inside or at the class ' +
      'declaration) are required'));
  F := ExpandFileName(F);
  IntfName := Trim(ArgStr(AArgs, 'interface_name'));
  Target := Trim(ArgStr(AArgs, 'target_file'));
  AddExisting := ArgBool(AArgs, 'add_to_existing');
  DoApply := ArgBool(AArgs, 'apply');
  Members := nil;
  var MArr := AArgs.GetValue<TJSONArray>('members', nil);
  if MArr <> nil then
    for var V in MArr do Members := Members + [Trim(V.Value)];
  Ok := False;
  RunErr := '';
  if not McpRunOnMain(
    procedure
    begin
      Ok := ExtractInterfaceHeadless(F, L1, AddExisting, IntfName, Target,
        Members, not DoApply, Info, Text, RunErr);
    end, not DoApply, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(RunErr));
  var Sel := TJSONArray.Create;
  for var M in Info.Members do
    if M.Selected then Sel.Add(M.Name);
  var Res := TJSONObject.Create;
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('class', Info.ClassName);
  Res.AddPair('interface', Info.InterfaceName);
  Res.AddPair('members', Sel);
  if AddExisting then
  begin
    Res.AddPair('target', Info.ExistingFile);
    Res.AddPair('targetLine', TJSONNumber.Create(Info.ExistingDeclLine));
  end
  else
    Res.AddPair('target', Info.TargetFile);
  Res.AddPair('interfaceText', Text);
  if DoApply then
    Res.AddPair('note', 'The interface was written, the class got it in its ' +
      'ancestor list and the uses clauses were updated. Files open in the IDE ' +
      'were changed in the buffer (not saved).')
  else
    Res.AddPair('note', 'Nothing was written. "interfaceText" is what the ' +
      'interface would look like; "members" what it would carry (pass ' +
      '"members" to choose, the default is every public and published method ' +
      'and property). Call again with apply=true to write it.');
  Result := McpOk(Res);
end;


// ---------------------------------------------------------------------------
//  Extract method
// ---------------------------------------------------------------------------
//
// The block is given as whole lines (an agent has no editor selection).
// apply=false answers with the GENERATED CODE - the routine, the call that
// replaces the block and the declaration line - because that is what has to
// be judged here; the write itself is four editor operations, not a line
// diff, and the applier is the dialog's.
function ToolExtractMethod(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Name, Err, RunErr: string;
  From1, To1: Integer;
  DoApply, Ok: Boolean;
  Prev: TExtractMethodPreview;
begin
  F := ArgStr(AArgs, 'file');
  From1 := ArgInt(AArgs, 'from_line');
  To1 := ArgInt(AArgs, 'to_line', From1);
  Name := Trim(ArgStr(AArgs, 'name'));
  if (F = '') or (From1 < 1) or (To1 < From1) then
    Exit(McpErr('arguments "file" and "from_line" (1-based) are required; ' +
      '"to_line" defaults to from_line'));
  if Name = '' then Name := 'ExtractedMethod';
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Ok := False;
  RunErr := '';
  if not McpRunOnMain(
    procedure
    begin
      Ok := ExtractMethodHeadless(F, From1, To1, Name, DoApply, Prev, RunErr);
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(RunErr));
  var Res := TJSONObject.Create;
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('name', Name);
  Res.AddPair('file', F);
  Res.AddPair('fromLine', TJSONNumber.Create(From1));
  Res.AddPair('toLine', TJSONNumber.Create(To1));
  if Prev.EnclosingClass <> '' then Res.AddPair('class', Prev.EnclosingClass);
  Res.AddPair('parameters', TJSONNumber.Create(Prev.ParamCount));
  Res.AddPair('localVariables', TJSONNumber.Create(Prev.LocalCount));
  Res.AddPair('method', Prev.MethodText);
  Res.AddPair('call', Prev.CallText);
  if Prev.DeclText <> '' then Res.AddPair('declaration', Prev.DeclText);
  Res.AddPair('insertAtLine', TJSONNumber.Create(Prev.InsertLine));
  if Prev.ClassDeclLine > 0 then
    Res.AddPair('declarationAtLine', TJSONNumber.Create(Prev.ClassDeclLine));
  if DoApply then
    Res.AddPair('note', 'Changed in the IDE buffer (not saved, undoable with ' +
      'Ctrl+Z). The variables that moved into the new routine were removed ' +
      'from the old one''s var section.')
  else
    Res.AddPair('note', 'Nothing was written. "method" is the routine that ' +
      'would be inserted, "call" what replaces the block, "declaration" the ' +
      'line added to the class. Which variables become parameters and which ' +
      'become locals was resolved with DelphiLSP. Call again with apply=true ' +
      'to write it.');
  Result := McpOk(Res);
end;

// Adds IInterface support (FRefCount, QueryInterface, _AddRef, _Release and
// the NewInstance / AfterConstruction pair) to a class that does not descend
// from TInterfacedObject. apply=false returns the code it would add.
function ToolAddIInterface(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, Report, RunErr: string;
  L1: Integer;
  DoApply, Ok: Boolean;
begin
  F := ArgStr(AArgs, 'file');
  L1 := ArgInt(AArgs, 'line');
  if (F = '') or (L1 < 1) then
    Exit(McpErr('arguments "file" and "line" (1-based, inside or at the class ' +
      'declaration) are required'));
  F := ExpandFileName(F);
  DoApply := ArgBool(AArgs, 'apply');
  Ok := False;
  RunErr := '';
  Report := '';
  if not McpRunOnMain(
    procedure
    begin
      Ok := DelegateIInterfaceHeadless(F, L1, not DoApply, Report, RunErr);
    end, not DoApply, AStop, Err, 120000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(RunErr));
  var Res := TJSONObject.Create;
  Res.AddPair('applied', TJSONBool.Create(DoApply));
  Res.AddPair('file', F);
  Res.AddPair('report', Report);
  if DoApply then
    Res.AddPair('note', 'Changed in the IDE buffer when the file is open (not ' +
      'saved), otherwise on disk. Hold the instance as an interface from now ' +
      'on - it frees itself when the last reference drops.')
  else
    Res.AddPair('note', 'Nothing was written. "report" shows the declaration ' +
      'and implementation that would be added. Call again with apply=true.');
  Result := McpOk(Res);
end;

initialization
  RegisterMcpTool('find_unit', ToolFindUnit);
  RegisterMcpTool('add_iinterface', ToolAddIInterface);
  RegisterMcpTool('extract_method', ToolExtractMethod);
  RegisterMcpTool('extract_interface', ToolExtractInterface);
  RegisterMcpTool('signature_check', ToolSignatureCheck);
  RegisterMcpTool('interface_guids', ToolInterfaceGuids);
  RegisterMcpTool('dfm_events', ToolDfmEvents);
  RegisterMcpTool('find_unit_references', ToolFindUnitReferences);
  RegisterMcpTool('find_original_symbol', ToolFindOriginalSymbol);
  RegisterMcpTool('remove_with', ToolRemoveWith);
  RegisterMcpTool('cleanup_uses', ToolCleanupUses);
  RegisterMcpTool('move_to_unit', ToolMoveToUnit);
  RegisterMcpTool('extract_variable', ToolExtractVariable);
  RegisterMcpTool('wrap_try_finally', ToolWrapTryFinally);
  RegisterMcpTool('add_unit', ToolAddUnit);
  RegisterMcpTool('remove_unit', ToolRemoveUnit);
  RegisterMcpTool('analyze_uses', ToolAnalyzeUses);
  RegisterMcpTool('find_implementations', ToolFindImplementations);
  RegisterMcpTool('find_references', ToolFindReferences);
  RegisterMcpTool('uses_path', ToolUsesPath);
  RegisterMcpTool('uses_cycles', ToolUsesCycles);
  RegisterMcpTool('debug_consistency', ToolDebugConsistency);
  RegisterMcpTool('blame', ToolBlame);
  RegisterMcpTool('commit_info', ToolCommitInfo);
  RegisterMcpTool('rename_preview', ToolRenamePreview);
  RegisterMcpTool('rename_apply', ToolRenameApply);
  RegisterMcpTool('semantic_replace', ToolSemanticReplace);
  RegisterMcpTool('move_to_new_unit', ToolMoveToNewUnit);

finalization
  // before the BPL unloads - the wizard holds no IDE references of its own
  GRenameHost := nil;
  FreeAndNil(GRenameWizard);

end.
