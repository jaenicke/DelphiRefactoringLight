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
  Expert.SafeDeletePlan, Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Delphi.FileEncoding, Expert.PascalScanner;

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

function ToolAddUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, U, Sec, Err, Msg: string;
  Ok: Boolean;
begin
  F := ArgStr(AArgs, 'file');
  U := Trim(ArgStr(AArgs, 'unit'));
  Sec := ArgStr(AArgs, 'section', 'interface');
  if (F = '') or (U = '') then Exit(McpErr('arguments "file" and "unit" are required'));
  F := ExpandFileName(F);
  Msg := '';
  Ok := False;
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
      if SameText(Sec, 'implementation') then
        Ok := AddUnitToUses(F, U, usImplementation)
      else
        Ok := AddUnitToUses(F, U, usInterface);
      if not Ok then
        Msg := U + ' was not added - it is already reachable from that section, ' +
          'or the uses clause could not be edited safely (comments inside it)';
    end, False, AStop, Err) then Exit(McpErr(Err));
  if Msg <> '' then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('added', U);
  Res.AddPair('note', 'Edited in the IDE buffer when the file is open (not ' +
    'saved), otherwise on disk.');
  Result := McpOk(Res);
end;

function ToolRemoveUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, U, Err: string;
  Ok: Boolean;
begin
  F := ArgStr(AArgs, 'file');
  U := Trim(ArgStr(AArgs, 'unit'));
  if (F = '') or (U = '') then Exit(McpErr('arguments "file" and "unit" are required'));
  F := ExpandFileName(F);
  Ok := False;
  if not McpRunOnMain(
    procedure
    begin
      Ok := RemoveUnitFromUses(F, U);
    end, False, AStop, Err) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(U + ' could not be removed (not in a uses clause of ' + F + '?)'));
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('removed', U);
  Result := McpOk(Res);
end;

const
  VerdictNames: array[TUsesVerdict] of string = ('used', 'unused', 'movable',
    'unknown', 'init_code', 'ide_managed');

function ToolAnalyzeUses(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, C: string;
  Found: Boolean;
  Cycle0: Integer;
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
  var Entries := AnalyzeUses(C,
    function(const AIdent: string): TArray<string>
    begin
      Result := nil;
      for var H in Snap.Lookup(AIdent) do Result := Result + [H.UnitName];
    end,
    function(const AUnitName: string): Boolean
    begin
      Result := Snap.HasUnit(AUnitName);
    end,
    function(const AUnitName: string): Boolean
    begin
      Result := Snap.HasInitCode(AUnitName);
    end);
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
    Arr.Add(O);
  end;
  var Res := TJSONObject.Create;
  Res.AddPair('file', F);
  Res.AddPair('entries', Arr);
  Res.AddPair('note', 'unused = no identifier of the unit is used (remove_unit); ' +
    'movable = only used in the implementation (remove + add_unit with ' +
    'section implementation); init_code / ide_managed / unknown are KEPT ' +
    'deliberately. Class helpers, operators and initialization side effects ' +
    'are invisible to this textual analysis.');
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
  if not IsIncludeFile(F) then
  begin
    var StartBefore := Ctx.Client.GetFileDiagnosticsVersion(F);
    if McpSyncLspContent(Ctx.Client, F, Ctx.Content) then
      McpWaitLspAnalysed(Ctx.Client, F, StartBefore, AStop);
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
    var IncCtx := TLspIncludeContext.Create(Ctx.Client,
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
      if (L1 - 1 <= High(CL)) and LineDeclaresName(CL[L1 - 1], Ctx.Identifier) then
      begin
        DeclFile := F;
        DeclLine := L1 - 1;
        DeclCol := Ctx.IdentCol0;
      end
      else
        Exit(McpErr('DelphiLSP knows no declaration for ' + Ctx.Identifier + ' at that position'));
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
    IncCtx.AddTargetWithPartner(Targets, DeclFile, DeclLine, DeclCol);
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
    Links := CollectLinkedTargets(Graph, OwnerTypeName, Ctx.Identifier, Linked);
    var PreSkipped := 0;
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
              var Before := Ctx.Client.GetFileDiagnosticsVersion(Cd.FilePath);
              if McpSyncLspContent(Ctx.Client, Cd.FilePath, C) then
              begin
                Inc(SentCount);
                if not McpWaitLspAnalysed(Ctx.Client, Cd.FilePath, Before, AStop) then
                  Inc(TimedOut);
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
          on E: Exception do Answer := E.ClassName + ': ' + E.Message;
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

function ToolMoveToNewUnit(AArgs: TJSONObject; AStop: THandle): string;
var
  F, NewUnit, Err: string;
  L1, C1: Integer;
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
  var Ok := False;
  var Ident := '';
  var Msg := '';
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
      Ok := TLspMoveToUnit.ExecuteToNewUnit(Ident, F, NewFile, Msg);
    end, False, AStop, Err, 300000) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(Msg));
  var Res := TJSONObject.Create;
  Res.AddPair('moved', Ident);
  Res.AddPair('new_unit', NewFile);
  if Msg <> '' then Res.AddPair('note', Msg);
  Res.AddPair('saved', 'The new unit was created on disk and added to the project; ' +
    'its content and the edits of the other units are in the IDE buffers (not ' +
    'saved) - units that are not open in the IDE were changed on disk.');
  Result := McpOk(Res);
end;

initialization
  RegisterMcpTool('find_unit', ToolFindUnit);
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
