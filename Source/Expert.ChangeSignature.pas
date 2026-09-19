(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ChangeSignature;

// CHANGE METHOD SIGNATURE (issue #11, suggestion 9 by Ian Branch): add,
// remove, reorder and rename parameters, change their type, modifier or
// default value - and update every header of the method's FAMILY and every
// call site in one step.
//
//   1. DelphiLSP names the declaration of the identifier at the caret; the
//      Safe Delete planner locates declaration + implementation.
//   2. The family: interface methods the class implements and the other
//      classes implementing them, or - for an interface method - every
//      implementing class; for a virtual method the whole override chain
//      (Expert.InterfaceLinks). All of them must change together or the
//      code no longer compiles.
//   3. Every whole-word occurrence in the project scope is asked where it
//      leads (DelphiLSP GotoDefinition); only occurrences that lead to a
//      family header are calls. An occurrence without an answer (inactive
//      {$IFDEF} branch) is listed as NOT VERIFIED and left alone.
//   4. The rewrite itself is pure (Expert.SignatureEdit): each call is
//      classified (arguments / statement / expression / method reference /
//      property accessor), argument lists are re-ordered and completed,
//      renamed parameters are followed into the bodies.
//
// The analysis runs once; the dialog re-plans on every edit of the
// parameter grid (pure, no LSP) and shows the preview. Applying re-checks
// that no touched file changed since the analysis.

interface

uses
  System.SysUtils, System.Classes, Expert.SignatureEdit, Expert.SafeDelete;

type
  TChangeSigUnverified = record
    FilePath: string;
    Line, Col: Integer;   // 0-based
    Text: string;
    Note: string;
  end;

  TChangeSigAnalysis = record
    Ok: Boolean;
    Error: string;
    Identifier: string;
    Kind: string;                    // procedure / function / constructor / ...
    Container: string;               // owning class / interface, '' = routine
    ResultType: string;
    Sources: TArray<TSigSource>;
    Headers: TArray<TSigHeader>;     // [0] = the declaration
    Calls: TArray<TSigCall>;
    Unverified: TArray<TChangeSigUnverified>;
    Vetoes: TArray<string>;          // block every change
    Notes: TArray<string>;
    Candidates, FilesScanned: Integer;
    function Params: TArray<TSigParam>;
    function QualifiedName: string;
    function Summary: string;
  end;

/// <summary>Worker thread: declaration, family, verified occurrences.
///  AStop cancels (an event handle).</summary>
function AnalyzeChangeSignature(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>): TChangeSigAnalysis;

/// <summary>Pure: the edits for the new parameter list (vetoes of the
///  analysis are errors of the plan).</summary>
function PlanChangeSignature(const A: TChangeSigAnalysis;
  const ANew: TArray<TNewParam>): TSigPlan;

/// <summary>Main thread: applies APlan when none of the touched files
///  changed since the analysis.</summary>
function ApplyChangeSignature(const A: TChangeSigAnalysis; const APlan: TSigPlan;
  out AError: string): Boolean;

/// <summary>The line(s) an edit produces - for previews.</summary>
function SigEditPreview(const A: TChangeSigAnalysis; const AEdit: TSigEdit): string;

/// <summary>Editor entry point (menu "Change signature...").</summary>
procedure ChangeSignatureAtCursor;

implementation

uses
  Winapi.Windows, System.Types, System.Math, System.StrUtils, System.IOUtils, System.JSON,
  System.SyncObjs, System.Generics.Collections, System.UITypes,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Grids, Vcl.Graphics,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.AutoImport, Expert.UsesEditor,
  Expert.DiagStore, Expert.DialogHelper, Expert.IdeThemes, Expert.WorkerLatch,
  Expert.McpServer, Expert.McpLspTools, Expert.SafeDeletePlan, Expert.InterfaceLinks,
  Expert.IncludeExpansion, Lsp.Client, Lsp.Protocol, Lsp.Uri, Delphi.FileEncoding;

const
  MaxCandidates = 600;
  ContextText: array[TCallContext] of string = ('call', 'call without arguments',
    'function call in an expression', 'property accessor', 'method reference');

{ TChangeSigAnalysis }

function TChangeSigAnalysis.Params: TArray<TSigParam>;
begin
  if Length(Headers) > 0 then Result := Headers[0].Params else Result := nil;
end;

function TChangeSigAnalysis.QualifiedName: string;
begin
  Result := IfThen(Container <> '', Container + '.', '') + Identifier;
end;

function TChangeSigAnalysis.Summary: string;
begin
  if not Ok then Exit('Change signature is not possible: ' + Error);
  Result := Format('%s %s(%s)%s', [Kind, QualifiedName, FormatParamList(Params),
    IfThen(ResultType <> '', ': ' + ResultType, '')]) + sLineBreak +
    Format('%d header(s) of the family, %d call site(s) verified by DelphiLSP ' +
      '(%d occurrence(s) checked in %d file(s))',
      [Length(Headers), Length(Calls), Candidates, FilesScanned]);
  for var H in Headers do
    Result := Result + sLineBreak + Format('  %s - %s(%d)', [H.Caption,
      ExtractFileName(Sources[H.FileIndex].FilePath), H.Line + 1]);
  for var V in Vetoes do
    Result := Result + sLineBreak + 'BLOCKED: ' + V;
  if Length(Unverified) > 0 then
    Result := Result + sLineBreak + Format('%d occurrence(s) DelphiLSP could not resolve ' +
      '(e.g. in an inactive {$IFDEF} branch) - they are NOT changed, check them by hand',
      [Length(Unverified)]);
  for var N in Notes do
    Result := Result + sLineBreak + 'Note: ' + N;
end;

// ---------------------------------------------------------------------------
//  Analysis
// ---------------------------------------------------------------------------

type
  TSigContentCache = class
  private
    FMap: TDictionary<string, string>;
  public
    constructor Create(const AIn: TSafeDeleteInput);
    destructor Destroy; override;
    function Get(const AFile: string; out AContent: string): Boolean;
  end;

constructor TSigContentCache.Create(const AIn: TSafeDeleteInput);
begin
  inherited Create;
  FMap := TDictionary<string, string>.Create;
  for var I := 0 to Min(High(AIn.OpenFiles), High(AIn.OpenContents)) do
    FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.OpenFiles[I])), AIn.OpenContents[I]);
  FMap.AddOrSetValue(UpperCase(ExpandFileName(AIn.FileName)), AIn.Content);
end;

destructor TSigContentCache.Destroy;
begin
  FMap.Free;
  inherited;
end;

function TSigContentCache.Get(const AFile: string; out AContent: string): Boolean;
begin
  var K := UpperCase(ExpandFileName(AFile));
  if FMap.TryGetValue(K, AContent) then Exit(True);
  Result := False;
  AContent := '';
  if not FileExists(AFile) then Exit;
  try
    AContent := ReadDelphiFile(AFile);
    FMap.Add(K, AContent);
    Result := True;
  except
    Result := False;
  end;
end;

function Directives(const AContent: string; ALine0: Integer): string;
var
  HdrEnd: Integer;
begin
  Result := '';
  var Lines := SplitContentLines(AContent);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  if CollectHeader(Lines, ALine0, HdrEnd) = '' then Exit;
  Result := HeaderDirectives(Lines, ALine0, HdrEnd);
end;

function IsDispatched(const ADirs: string): Boolean;
begin
  Result := HasWholeWordCI(ADirs, 'virtual') or HasWholeWordCI(ADirs, 'dynamic') or
    HasWholeWordCI(ADirs, 'override') or HasWholeWordCI(ADirs, 'abstract');
end;

function AnalyzeChangeSignature(const AIn: TSafeDeleteInput; AStop: THandle;
  const AProgress: TProc<Integer, Integer, string>): TChangeSigAnalysis;
var
  Res: TChangeSigAnalysis;
  Src: TSigContentCache;
  IncCtx: TLspIncludeContext;
  Targets: TLspSymbolTargets;
  FileIdx: TDictionary<string, Integer>;
  HeaderNames: TDictionary<string, Boolean>;
  DeclFile, DeclContent: string;
  DeclLine, IdentCol: Integer;

  function Stopped: Boolean;
  begin
    Result := (AStop <> 0) and (WaitForSingleObject(AStop, 0) = WAIT_OBJECT_0);
  end;

  function InBds(const AFile: string): Boolean;
  begin
    Result := (AIn.BdsRoot <> '') and ExpandFileName(AFile).ToUpper.StartsWith(
      IncludeTrailingPathDelimiter(AIn.BdsRoot).ToUpper);
  end;

  function SourceIndex(const AFile: string): Integer;
  var
    S: TSigSource;
  begin
    var K := UpperCase(ExpandFileName(AFile));
    if FileIdx.TryGetValue(K, Result) then Exit;
    S.FilePath := ExpandFileName(AFile);
    if not Src.Get(S.FilePath, S.Content) then Exit(-1);
    Res.Sources := Res.Sources + [S];
    Result := High(Res.Sources);
    FileIdx.Add(K, Result);
  end;

  function NameKey(AFileIndex, ALine, ACol: Integer): string;
  begin
    Result := Format('%d|%d|%d', [AFileIndex, ALine, ACol]);
  end;

  procedure AddHeader(const AFile: string; ALine: Integer; AIsDecl, AWithBody: Boolean;
    const ACaption: string);
  var
    H: TSigHeader;
    Why: string;
  begin
    for var X in Res.Headers do
      if SameText(Res.Sources[X.FileIndex].FilePath, ExpandFileName(AFile)) and (X.Line = ALine) then
        Exit;
    if InBds(AFile) then
    begin
      Res.Vetoes := Res.Vetoes + [ACaption + ' is declared in the RTL / VCL (' +
        ExtractFileName(AFile) + ') - its signature cannot change'];
      Exit;
    end;
    var FI := SourceIndex(AFile);
    if FI < 0 then
    begin
      Res.Vetoes := Res.Vetoes + [ACaption + ': the file cannot be read (' + AFile + ')'];
      Exit;
    end;
    if not LocateSigHeader(Res.Sources[FI].Content, ALine, Res.Identifier, AIsDecl, H, Why,
      AWithBody) then
    begin
      Res.Vetoes := Res.Vetoes + [ACaption + ': ' + Why];
      Exit;
    end;
    H.FileIndex := FI;
    H.Caption := ACaption;
    Res.Headers := Res.Headers + [H];
    Targets.Add(Res.Sources[FI].FilePath, ALine);
    HeaderNames.AddOrSetValue(NameKey(FI, ALine,
      H.NameStart - SigOffsetOf(Res.Sources[FI].Content, ALine, 0)), True);
  end;

var
  StartLines, DeclLines: TArray<string>;
  Sym: TSafeDeleteSymbol;
  Why: string;
begin
  Res := Default(TChangeSigAnalysis);
  IncCtx := nil;
  Src := TSigContentCache.Create(AIn);
  FileIdx := TDictionary<string, Integer>.Create;
  HeaderNames := TDictionary<string, Boolean>.Create;
  try
    StartLines := SplitContentLines(AIn.Content);
    Res.Identifier := IdentifierAtPos(StartLines, AIn.Line0, AIn.Col0, IdentCol);
    if Res.Identifier = '' then
    begin
      Res.Error := 'there is no identifier at the caret';
      Exit(Res);
    end;
    if AIn.Client = nil then
    begin
      Res.Error := 'the DelphiLSP session is not running yet (it starts with the ' +
        'first opened project)';
      Exit(Res);
    end;
    if Assigned(AProgress) then AProgress(0, 0, 'Asking DelphiLSP for the declaration...');
    IncCtx := TLspIncludeContext.Create(AIn.Client,
      function(const APath: string; out AContent: string): Boolean
      begin
        Result := Src.Get(APath, AContent);
      end);
    IncCtx.RegisterFiles(AIn.ScopeFiles + [AIn.FileName]);

    // 1. the declaration
    if not IncCtx.OwnsDocument(AIn.FileName) then
    begin
      var Before := AIn.Client.GetFileDiagnosticsVersion(AIn.FileName);
      if McpSyncLspContent(AIn.Client, AIn.FileName, AIn.Content) then
        McpWaitLspAnalysed(AIn.Client, AIn.FileName, Before, AStop, 30000);
    end;
    var D := IncCtx.Definition(AIn.FileName, AIn.Line0, IdentCol);
    var Deadline := GetTickCount64 + 3000;
    while (Length(D) = 0) and (GetTickCount64 < Deadline) and not Stopped do
    begin
      Sleep(300);
      D := IncCtx.Definition(AIn.FileName, AIn.Line0, IdentCol);
    end;
    if Length(D) > 0 then
    begin
      DeclFile := ExpandFileName(TLspUri.FileUriToPath(D[0].Uri));
      DeclLine := D[0].Range.Start.Line;
    end
    else if LineDeclaresName(StartLines[AIn.Line0], Res.Identifier) then
    begin
      DeclFile := ExpandFileName(AIn.FileName);
      DeclLine := AIn.Line0;
    end
    else
    begin
      Res.Error := 'DelphiLSP knows no declaration for "' + Res.Identifier + '"';
      Exit(Res);
    end;
    if InBds(DeclFile) then
    begin
      Res.Error := Res.Identifier + ' is declared in the RTL / VCL (' +
        ExtractFileName(DeclFile) + ')';
      Exit(Res);
    end;
    if not Src.Get(DeclFile, DeclContent) then
    begin
      Res.Error := 'the declaring file cannot be read: ' + DeclFile;
      Exit(Res);
    end;
    DeclLines := SplitContentLines(DeclContent);
    if not PlanSafeDeleteSymbol(DeclLines, DeclLine, Res.Identifier, Sym, Why) then
    begin
      Res.Error := Why;
      Exit(Res);
    end;
    if not (Sym.Kind in [sdkMethod, sdkRoutine]) then
    begin
      Res.Error := Format('%s is a %s, not a method or routine', [Res.Identifier,
        SafeDeleteKindText(Sym.Kind)]);
      Exit(Res);
    end;
    Res.Container := Sym.Container;
    begin
      var IsCM: Boolean;
      var HdrEnd: Integer;
      var Kind: string;
      if IsHeaderLine(DeclLines[Sym.DeclLine], Kind, IsCM) then
        Res.Kind := IfThen(IsCM, 'class ', '') + Kind;
      var Q, P, R: string;
      var Hdr := CollectHeader(DeclLines, Sym.DeclLine, HdrEnd);
      if (Hdr <> '') and ParseHeader(Hdr, Kind, Q, P, R) then Res.ResultType := R;
    end;
    var Dirs := Directives(DeclContent, Sym.DeclLine);
    if HasWholeWordCI(Dirs, 'overload') then
      Res.Vetoes := Res.Vetoes + ['the method is OVERLOADED - a changed parameter list ' +
        'could collide with another overload or silently move calls to it; change it by hand'];
    if HasWholeWordCI(Dirs, 'message') then
      Res.Vetoes := Res.Vetoes + ['a MESSAGE handler has a fixed signature'];
    for var V in Sym.Vetoes do
      if Pos('PUBLISHED', V) > 0 then
        Res.Notes := Res.Notes + ['the method is published - code that finds it by NAME ' +
          '(RTTI, MethodAddress) is not seen by this check'];

    // 2. the family
    var Q := Res.QualifiedName;
    var WithBody := (Sym.ImplLine >= 0) and (Sym.ImplLine = Sym.DeclLine);
    AddHeader(DeclFile, Sym.DeclLine, True, WithBody, Q + IfThen(WithBody, '',
      IfThen(Sym.IsInterfaceMember, ' (interface)', ' (declaration)')));
    if (Sym.ImplLine >= 0) and (Sym.ImplLine <> Sym.DeclLine) then
      AddHeader(DeclFile, Sym.ImplLine, False, False, Q + ' (implementation)');
    if Length(Res.Headers) = 0 then
    begin
      Res.Error := string.Join('; ', Res.Vetoes);
      Exit(Res);
    end;
    if Sym.Kind = sdkMethod then
    begin
      if Assigned(AProgress) then AProgress(0, 0, 'Collecting interfaces and overrides...');
      var Graph := TTypeGraph.Create(AIn.ScopeFiles + [DeclFile],
        function(const APath: string; out AContent: string): Boolean
        begin
          Result := Src.Get(APath, AContent);
        end);
      try
        var Links: TArray<TMemberLink> := nil;
        if Sym.IsInterfaceMember then
          Links := Graph.ClassMethodsImplementing(Sym.Container, Res.Identifier)
        else
        begin
          var Intfs := Graph.InterfaceMethodsImplementedBy(Sym.Container, Res.Identifier);
          Links := Intfs;
          for var I in Intfs do
            Links := Links + Graph.ClassMethodsImplementing(I.TypeName, Res.Identifier);
          if IsDispatched(Dirs) then
            for var L in Graph.ClassHierarchyMembers(Sym.Container, Res.Identifier) do
            begin
              var C: string;
              if not Src.Get(L.FilePath, C) then Continue;
              if IsDispatched(Directives(C, L.Line)) then
                Links := Links + [L]
              else
                Res.Notes := Res.Notes + [Format('%s.%s hides the method (no override) - ' +
                  'not changed', [L.TypeName, Res.Identifier])];
            end;
        end;
        for var L in Links do
        begin
          if SameText(L.TypeName, Sym.Container) and SameText(ExpandFileName(L.FilePath), DeclFile) then
            Continue;
          var C: string;
          if Src.Get(L.FilePath, C) and HasWholeWordCI(Directives(C, L.Line), 'overload') then
            Res.Vetoes := Res.Vetoes + [Format('%s.%s is overloaded - change it by hand',
              [L.TypeName, Res.Identifier])];
          AddHeader(L.FilePath, L.Line, True, False, L.TypeName + '.' + Res.Identifier +
            IfThen(L.IsInterface, ' (interface)', ' (declaration)'));
          if L.ImplLine >= 0 then
            AddHeader(L.FilePath, L.ImplLine, False, False, L.TypeName + '.' +
              Res.Identifier + ' (implementation)');
        end;
      finally
        Graph.Free;
      end;
    end;

    // 3. every occurrence in the scope
    var Files: TArray<string> := [DeclFile];
    if Sym.LocalFirst < 0 then
    begin
      for var S in Res.Sources do
        if not SameText(S.FilePath, DeclFile) then Files := Files + [S.FilePath];
      for var F in AIn.ScopeFiles do
        Files := Files + [ExpandFileName(F)];
    end;
    var Seen := TDictionary<string, Boolean>.Create;
    var Cands: TArray<TPair<string, TPoint>> := nil;
    try
      for var F in Files do
      begin
        if Stopped then begin Res.Error := 'cancelled'; Exit(Res); end;
        if Seen.ContainsKey(UpperCase(F)) then Continue;
        Seen.Add(UpperCase(F), True);
        var C: string;
        if not Src.Get(F, C) then Continue;
        Inc(Res.FilesScanned);
        var Lines := SplitContentLines(C);
        var Hits: TArray<TPoint> := nil;
        if Sym.LocalFirst >= 0 then
          CodeWordHits(Lines, Res.Identifier, Sym.LocalFirst, Sym.LocalLast, Hits)
        else
          CodeWordHits(Lines, Res.Identifier, 0, High(Lines), Hits);
        var FI: Integer;
        if not FileIdx.TryGetValue(UpperCase(F), FI) then FI := -1;
        for var H in Hits do
        begin
          if (FI >= 0) and HeaderNames.ContainsKey(NameKey(FI, H.Y, H.X)) then Continue;
          if Length(Cands) >= MaxCandidates then
          begin
            Res.Vetoes := Res.Vetoes + [Format('more than %d occurrences of "%s" - the check ' +
              'stopped; this is too broad for an automatic change', [MaxCandidates, Res.Identifier])];
            Res.Error := Res.Vetoes[High(Res.Vetoes)];
            Exit(Res);
          end;
          Cands := Cands + [TPair<string, TPoint>.Create(F, H)];
        end;
      end;
    finally
      Seen.Free;
    end;
    Res.Candidates := Length(Cands);

    // 4. which of them lead to the family
    var OtherDecls := 0;
    var Verified := TDictionary<string, TArray<TPoint>>.Create;
    var OrigPath := TDictionary<string, string>.Create;
    var Synced := TDictionary<string, Boolean>.Create;
    try
      for var I := 0 to High(Cands) do
      begin
        if Stopped then begin Res.Error := 'cancelled'; Exit(Res); end;
        var F := Cands[I].Key;
        var P := Cands[I].Value;
        if Assigned(AProgress) then
          AProgress(I + 1, Length(Cands), Format('%s(%d)', [ExtractFileName(F), P.Y + 1]));
        var Key := UpperCase(F);
        var FirstInFile := not Synced.ContainsKey(Key);
        if FirstInFile then
        begin
          Synced.Add(Key, True);
          var C: string;
          if not IncCtx.OwnsDocument(F) and Src.Get(F, C) then
          begin
            var B := AIn.Client.GetFileDiagnosticsVersion(F);
            if McpSyncLspContent(AIn.Client, F, C) then
              McpWaitLspAnalysed(AIn.Client, F, B, AStop, 30000);
          end;
        end;
        var Answer: TArray<TLspLocation> := nil;
        try
          Answer := IncCtx.Definition(F, P.Y, P.X);
          var Dl := GetTickCount64 + UInt64(IfThen(FirstInFile, 3000, 600));
          while (Length(Answer) = 0) and (GetTickCount64 < Dl) and not Stopped do
          begin
            Sleep(300);
            Answer := IncCtx.Definition(F, P.Y, P.X);
          end;
        except
          Answer := nil;
        end;
        if Length(Answer) = 0 then
        begin
          var U: TChangeSigUnverified;
          var C: string;
          U.FilePath := F; U.Line := P.Y; U.Col := P.X;
          U.Text := '';
          var RawLine := '';
          if Src.Get(F, C) then
          begin
            var L := SplitContentLines(C);
            if P.Y <= High(L) then RawLine := L[P.Y];
            U.Text := Trim(RawLine);
          end;
          // DelphiLSP answers nothing AT a declaration. A declaration of the
          // same name that is not a family header belongs to another
          // symbol (an unrelated class or interface) - the family analysis
          // already took every interface / override that belongs to it
          if LineDeclaresName(RawLine, Res.Identifier) then
          begin
            Inc(OtherDecls);
            Continue;
          end;
          U.Note := IfThen(IsIncludeFile(F), 'DelphiLSP gives no answer inside this include file',
            'DelphiLSP gives no answer here - an inactive {$IFDEF} branch looks like this');
          Res.Unverified := Res.Unverified + [U];
        end
        else if Targets.Contains(TLspUri.FileUriToPath(Answer[0].Uri), Answer[0].Range.Start.Line) then
        begin
          var Arr: TArray<TPoint>;
          if not Verified.TryGetValue(Key, Arr) then Arr := nil;
          Verified.AddOrSetValue(Key, Arr + [P]);
          OrigPath.AddOrSetValue(Key, F);
        end;
      end;
      // 5. classify the calls, one masking pass per file
      for var Pair in Verified do
      begin
        var FI := SourceIndex(OrigPath[Pair.Key]);
        if FI < 0 then Continue;
        Res.Calls := Res.Calls + LocateSigCalls(Res.Sources[FI].Content, Pair.Value,
          Length(Res.Identifier), FI);
      end;
    finally
      Synced.Free;
      OrigPath.Free;
      Verified.Free;
    end;

    // 6. an event handler bound in a form: the event type fixes the signature
    if Sym.Kind = sdkMethod then
      for var H in Res.Headers do
      begin
        if not H.IsDecl then Continue;
        for var Ext in ['.dfm', '.fmx'] do
        begin
          var FF := ChangeFileExt(Res.Sources[H.FileIndex].FilePath, Ext);
          if not FileExists(FF) then Continue;
          var Txt: string;
          try
            Txt := ReadDelphiFile(FF);
          except
            Continue;
          end;
          for var L in FormTextMentions(Txt, Res.Identifier) do
          begin
            Res.Vetoes := Res.Vetoes + [Format('%s(%d) binds it as an event handler - the ' +
              'event type fixes its signature', [ExtractFileName(FF), L + 1])];
            Break;
          end;
        end;
      end;

    if OtherDecls > 0 then
      Res.Notes := Res.Notes + [Format('%d other declaration(s) of "%s" (unrelated types) ' +
        'are not part of the change', [OtherDecls, Res.Identifier])];
    if Sym.LocalFirst >= 0 then
      Res.Notes := Res.Notes + ['local routine - only its enclosing routine was checked']
    else
      Res.Notes := Res.Notes + [Format('checked the project scope (%d file(s)); code OUTSIDE ' +
        'it (other projects using this unit) is not seen', [Res.FilesScanned])];
    if Res.ResultType <> '' then
      Res.Notes := Res.Notes + ['the result type is not changed by this refactoring'];
    Res.Ok := True;
    Result := Res;
  finally
    IncCtx.Free;
    HeaderNames.Free;
    FileIdx.Free;
    Src.Free;
  end;
end;

function PlanChangeSignature(const A: TChangeSigAnalysis;
  const ANew: TArray<TNewParam>): TSigPlan;
begin
  if not A.Ok then
  begin
    Result := Default(TSigPlan);
    Result.Errors := [A.Error];
    Exit;
  end;
  Result := PlanSignatureEdits(A.Sources, A.Headers, A.Calls, ANew);
  Result.Errors := A.Vetoes + Result.Errors;
  for var U in A.Unverified do
    Result.Warnings := Result.Warnings + [Format('%s(%d) NOT VERIFIED, not changed: %s',
      [ExtractFileName(U.FilePath), U.Line + 1, U.Text])];
end;

function SigEditPreview(const A: TChangeSigAnalysis; const AEdit: TSigEdit): string;
begin
  var T := ApplySigEdits(A.Sources[AEdit.FileIndex].Content, [AEdit], AEdit.FileIndex);
  var S := AEdit.Start;
  while (S > 1) and (T[S - 1] <> #10) do Dec(S);
  var E := AEdit.Start + Length(AEdit.NewText);
  while (E <= Length(T)) and (T[E] <> #10) do Inc(E);
  Result := Trim(Copy(T, S, E - S));
end;

function ApplyChangeSignature(const A: TChangeSigAnalysis; const APlan: TSigPlan;
  out AError: string): Boolean;
var
  Cur: TArray<string>;
  Touched: TArray<Integer>;
begin
  Result := False;
  AError := '';
  if not APlan.Ok or (Length(APlan.Edits) = 0) then
  begin
    AError := 'nothing to apply';
    Exit;
  end;
  Touched := nil;
  for var E in APlan.Edits do
    if (Length(Touched) = 0) or (Touched[High(Touched)] <> E.FileIndex) then
      Touched := Touched + [E.FileIndex];
  // every file first: the whole change or nothing
  SetLength(Cur, Length(A.Sources));
  for var FI in Touched do
  begin
    var P := A.Sources[FI].FilePath;
    if (Editor = nil) or not Editor.ReadEditorContent(P, Cur[FI]) then
      try
        Cur[FI] := ReadDelphiFile(P);
      except
        on Ex: Exception do
        begin
          AError := ExtractFileName(P) + ': ' + Ex.Message;
          Exit;
        end;
      end;
    if DiagContentHash(Cur[FI]) <> DiagContentHash(A.Sources[FI].Content) then
    begin
      AError := ExtractFileName(P) + ' changed since the analysis - run it again';
      Exit;
    end;
  end;
  var Done := 0;
  for var FI in Touched do
  begin
    var SL := TStringList.Create;
    try
      SL.Text := ApplySigEdits(Cur[FI], APlan.Edits, FI);
      if not ApplyLinesMinimal(A.Sources[FI].FilePath, SL, Cur[FI]) then
      begin
        AError := Format('%s could not be changed (%d of %d file(s) are changed already - ' +
          'undo them in the editor)', [ExtractFileName(A.Sources[FI].FilePath), Done,
          Length(Touched)]);
        Exit;
      end;
      Inc(Done);
    finally
      SL.Free;
    end;
  end;
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Dialog
// ---------------------------------------------------------------------------

const
  ColModifier = 0;
  ColName = 1;
  ColType = 2;
  ColDefault = 3;
  ColValue = 4;
  ColWas = 5;

type
  TChangeSignatureDialog = class(TForm)
  private
    FAn: TChangeSigAnalysis;
    FRows: TArray<TNewParam>;
    FPlan: TSigPlan;
    FInfo, FPreview: TMemo;
    FGrid: TStringGrid;
    FBtnAdd, FBtnRemove, FBtnUp, FBtnDown, FBtnApply, FBtnCancel: TButton;
    FTimer: TTimer;
    FFilling: Boolean;
    procedure FillGrid;
    procedure ReadGrid;
    procedure Replan;
    procedure Changed;
    procedure DoTimer(Sender: TObject);
    procedure DoSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure DoSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure DoAdd(Sender: TObject);
    procedure DoRemove(Sender: TObject);
    procedure DoMove(Sender: TObject);
    procedure DoApply(Sender: TObject);
  public
    Applied: Boolean;
    constructor CreateDialog(AOwner: TComponent; const AAn: TChangeSigAnalysis);
  end;

constructor TChangeSignatureDialog.CreateDialog(AOwner: TComponent; const AAn: TChangeSigAnalysis);

  function Btn(AParent: TWinControl; const ACaption: string; AAlign: TAlign;
    AOnClick: TNotifyEvent): TButton;
  begin
    Result := TButton.Create(Self);
    Result.Parent := AParent;
    Result.Caption := ACaption;
    Result.Align := AAlign;
    Result.AlignWithMargins := True;
    Result.OnClick := AOnClick;
  end;

var
  Top, Side, Bottom: TPanel;
begin
  inherited CreateNew(AOwner);
  FAn := AAn;
  FRows := UnchangedSignature(AAn.Params);
  Caption := 'Change signature: ' + AAn.QualifiedName;
  Width := 940;
  Height := 640;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FInfo := TMemo.Create(Self);
  FInfo.Parent := Self;
  FInfo.Align := alTop;
  FInfo.AlignWithMargins := True;
  FInfo.Height := 120;
  FInfo.ReadOnly := True;
  FInfo.ScrollBars := ssVertical;
  FInfo.Text := AAn.Summary;

  Top := TPanel.Create(Self);
  Top.Parent := Self;
  Top.Align := alTop;
  Top.Top := 200;
  Top.Height := 190;
  Top.BevelOuter := bvNone;

  Side := TPanel.Create(Self);
  Side.Parent := Top;
  Side.Align := alRight;
  Side.Width := 100;
  Side.BevelOuter := bvNone;
  FBtnDown := Btn(Side, 'Move &down', alTop, DoMove);
  FBtnUp := Btn(Side, 'Move &up', alTop, DoMove);
  FBtnRemove := Btn(Side, '&Remove', alTop, DoRemove);
  FBtnAdd := Btn(Side, '&Add', alTop, DoAdd);

  FGrid := TStringGrid.Create(Self);
  FGrid.Parent := Top;
  FGrid.Align := alClient;
  FGrid.AlignWithMargins := True;
  FGrid.ColCount := 6;
  FGrid.FixedCols := 0;
  FGrid.FixedRows := 1;
  FGrid.RowCount := 2;
  FGrid.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing,
    goEditing, goTabs, goAlwaysShowEditor];
  FGrid.ColWidths[ColModifier] := 80;
  FGrid.ColWidths[ColName] := 130;
  FGrid.ColWidths[ColType] := 170;
  FGrid.ColWidths[ColDefault] := 110;
  FGrid.ColWidths[ColValue] := 160;
  FGrid.ColWidths[ColWas] := 160;
  FGrid.Cells[ColModifier, 0] := 'Modifier';
  FGrid.Cells[ColName, 0] := 'Name';
  FGrid.Cells[ColType, 0] := 'Type';
  FGrid.Cells[ColDefault, 0] := 'Default';
  FGrid.Cells[ColValue, 0] := 'Value for existing calls';
  FGrid.Cells[ColWas, 0] := 'Was';
  FGrid.OnSetEditText := DoSetEditText;
  FGrid.OnSelectCell := DoSelectCell;

  Bottom := TPanel.Create(Self);
  Bottom.Parent := Self;
  Bottom.Align := alBottom;
  Bottom.Height := 40;
  Bottom.BevelOuter := bvNone;
  FBtnCancel := Btn(Bottom, '&Close', alRight, nil);
  FBtnCancel.Cancel := True;
  FBtnCancel.ModalResult := mrCancel;
  FBtnApply := Btn(Bottom, 'A&pply', alRight, DoApply);

  FPreview := TMemo.Create(Self);
  FPreview.Parent := Self;
  FPreview.Align := alClient;
  FPreview.AlignWithMargins := True;
  FPreview.ReadOnly := True;
  FPreview.ScrollBars := ssBoth;
  FPreview.WordWrap := False;

  FTimer := TTimer.Create(Self);
  FTimer.Enabled := False;
  FTimer.Interval := 300;
  FTimer.OnTimer := DoTimer;

  FillGrid;
  Replan;
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
  ActiveControl := FGrid;
end;

procedure TChangeSignatureDialog.FillGrid;
begin
  FFilling := True;
  try
    FGrid.RowCount := Max(2, Length(FRows) + 1);
    for var C := 0 to FGrid.ColCount - 1 do FGrid.Cells[C, 1] := '';
    for var I := 0 to High(FRows) do
    begin
      var R := FRows[I];
      FGrid.Cells[ColModifier, I + 1] := R.Param.Modifier;
      FGrid.Cells[ColName, I + 1] := R.Param.Name;
      FGrid.Cells[ColType, I + 1] := R.Param.TypeText;
      FGrid.Cells[ColDefault, I + 1] := R.Param.DefaultText;
      if R.OldIndex >= 0 then
      begin
        var O := FAn.Params[R.OldIndex];
        FGrid.Cells[ColValue, I + 1] := '';
        FGrid.Cells[ColWas, I + 1] := Format('%d: %s%s', [R.OldIndex + 1,
          IfThen(O.Modifier <> '', O.Modifier + ' ', ''), O.Name + IfThen(O.TypeText <> '',
          ': ' + O.TypeText, '')]);
      end
      else
      begin
        FGrid.Cells[ColValue, I + 1] := R.CallValue;
        FGrid.Cells[ColWas, I + 1] := '(new)';
      end;
    end;
  finally
    FFilling := False;
  end;
end;

procedure TChangeSignatureDialog.ReadGrid;
begin
  for var I := 0 to High(FRows) do
  begin
    FRows[I].Param.Modifier := Trim(FGrid.Cells[ColModifier, I + 1]);
    FRows[I].Param.Name := Trim(FGrid.Cells[ColName, I + 1]);
    FRows[I].Param.TypeText := Trim(FGrid.Cells[ColType, I + 1]);
    FRows[I].Param.DefaultText := Trim(FGrid.Cells[ColDefault, I + 1]);
    if FRows[I].OldIndex < 0 then FRows[I].CallValue := Trim(FGrid.Cells[ColValue, I + 1]);
  end;
end;

procedure TChangeSignatureDialog.Replan;
var
  SL: TStringList;
begin
  FPlan := PlanChangeSignature(FAn, FRows);
  SL := TStringList.Create;
  try
    var NewList: TArray<TSigParam> := nil;
    for var R in FRows do NewList := NewList + [R.Param];
    SL.Add('New: ' + FAn.Kind + ' ' + FAn.QualifiedName + '(' + FormatParamList(NewList) + ')' +
      IfThen(FAn.ResultType <> '', ': ' + FAn.ResultType, ''));
    SL.Add('');
    if not FPlan.Ok then
    begin
      SL.Add('NOT POSSIBLE:');
      for var E in FPlan.Errors do SL.Add('  - ' + E);
      SL.Add('');
    end;
    for var W in FPlan.Warnings do SL.Add('Note: ' + W);
    if Length(FPlan.Warnings) > 0 then SL.Add('');
    if FPlan.Ok and (Length(FPlan.Edits) = 0) then
      SL.Add('No change.')
    else if FPlan.Ok then
    begin
      SL.Add(Format('%d edit(s):', [Length(FPlan.Edits)]));
      for var E in FPlan.Edits do
        SL.Add(Format('  %s(%d) [%s]  %s', [ExtractFileName(FAn.Sources[E.FileIndex].FilePath),
          E.Line + 1, E.What, SigEditPreview(FAn, E)]));
    end;
    FPreview.Lines.Assign(SL);
  finally
    SL.Free;
  end;
  FBtnApply.Enabled := FPlan.Ok and (Length(FPlan.Edits) > 0);
  FBtnRemove.Enabled := Length(FRows) > 0;
  FBtnUp.Enabled := Length(FRows) > 1;
  FBtnDown.Enabled := Length(FRows) > 1;
end;

procedure TChangeSignatureDialog.Changed;
begin
  FTimer.Enabled := False;
  FTimer.Enabled := True;
end;

procedure TChangeSignatureDialog.DoTimer(Sender: TObject);
begin
  FTimer.Enabled := False;
  ReadGrid;
  Replan;
end;

procedure TChangeSignatureDialog.DoSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
begin
  if not FFilling then Changed;
end;

procedure TChangeSignatureDialog.DoSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  // 'Was' is information, and existing parameters take their calls' values
  var Editable := (ARow - 1 <= High(FRows)) and (ACol <> ColWas) and
    not ((ACol = ColValue) and (FRows[ARow - 1].OldIndex >= 0));
  if Editable then
    FGrid.Options := FGrid.Options + [goEditing]
  else
    FGrid.Options := FGrid.Options - [goEditing];
end;

procedure TChangeSignatureDialog.DoAdd(Sender: TObject);
begin
  ReadGrid;
  var N := 'NewParam';
  var K := 1;
  var Taken := True;
  while Taken do
  begin
    Taken := False;
    for var R in FRows do
      if SameText(R.Param.Name, N) then Taken := True;
    if Taken then
    begin
      Inc(K);
      N := 'NewParam' + IntToStr(K);
    end;
  end;
  var R := Default(TNewParam);
  R.OldIndex := -1;
  R.Param.Name := N;
  R.Param.TypeText := 'Integer';
  R.CallValue := '0';
  FRows := FRows + [R];
  FillGrid;
  FGrid.Row := Length(FRows);
  FGrid.Col := ColName;
  Replan;
end;

procedure TChangeSignatureDialog.DoRemove(Sender: TObject);
begin
  ReadGrid;
  var I := FGrid.Row - 1;
  if (I < 0) or (I > High(FRows)) then Exit;
  Delete(FRows, I, 1);
  FillGrid;
  Replan;
end;

procedure TChangeSignatureDialog.DoMove(Sender: TObject);
begin
  ReadGrid;
  var I := FGrid.Row - 1;
  var J := I + IfThen(Sender = FBtnUp, -1, 1);
  if (I < 0) or (I > High(FRows)) or (J < 0) or (J > High(FRows)) then Exit;
  var T := FRows[I];
  FRows[I] := FRows[J];
  FRows[J] := T;
  FillGrid;
  FGrid.Row := J + 1;
  Replan;
end;

procedure TChangeSignatureDialog.DoApply(Sender: TObject);
var
  Err: string;
begin
  ReadGrid;
  Replan;
  if not FBtnApply.Enabled then Exit;
  if ApplyChangeSignature(FAn, FPlan, Err) then
  begin
    Applied := True;
    ModalResult := mrOk;
  end
  else
    ShowThemedMessage('Change signature failed: ' + Err);
end;

// ---------------------------------------------------------------------------
//  Editor entry point
// ---------------------------------------------------------------------------

type
  TChangeSigJob = class(TInterfacedObject)
  public
    Lock: TCriticalSection;
    Cur, Total: Integer;
    Text: string;
    Done: Boolean;
    Res: TChangeSigAnalysis;
    Stop: TEvent;
    constructor Create;
    destructor Destroy; override;
  end;

constructor TChangeSigJob.Create;
begin
  inherited;
  Lock := TCriticalSection.Create;
  Stop := TEvent.Create(nil, True, False, '');
end;

destructor TChangeSigJob.Destroy;
begin
  Stop.Free;
  Lock.Free;
  inherited;
end;

procedure ChangeSignatureAtCursor;
var
  Ctx: TEditorContext;
  Inp: TSafeDeleteInput;
  Err: string;
  Job: TChangeSigJob;
  JobRef: IInterface;
  Prog: TCheckProgressWindow;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or (Ctx.WordAtCursor = '') then
  begin
    ShowThemedMessage('Change signature: place the caret on the method or routine name.');
    Exit;
  end;
  Editor.SaveAllFiles;   // the scan reads closed files from disk
  if not GatherScanInput(Ctx.FileName, Ctx.Line - 1, Ctx.Column - 1, Inp, Err) then
  begin
    ShowThemedMessage('Change signature: ' + Err);
    Exit;
  end;
  if Inp.Client = nil then
  begin
    ShowThemedMessage('Change signature needs the DelphiLSP session, which is not running ' +
      'yet - open a project and try again in a moment.');
    Exit;
  end;

  Job := TChangeSigJob.Create;
  JobRef := Job;
  Prog := CreateCheckProgress('Change signature', Application.MainForm,
    'Analysing ' + Ctx.WordAtCursor + '...');
  try
    var ThreadRef: IInterface := JobRef;
    var StopHandle := Job.Stop.Handle;
    if not StartWorker(
      procedure
      var
        R: TChangeSigAnalysis;
      begin
        try
          R := AnalyzeChangeSignature(Inp, StopHandle,
            procedure(ACur, ATotal: Integer; AText: string)
            begin
              Job.Lock.Enter;
              try
                Job.Cur := ACur; Job.Total := ATotal; Job.Text := AText;
              finally
                Job.Lock.Leave;
              end;
            end);
        except
          on E: Exception do
          begin
            R := Default(TChangeSigAnalysis);
            R.Error := E.ClassName + ': ' + E.Message;
          end;
        end;
        Job.Lock.Enter;
        try
          Job.Res := R;
          Job.Done := True;
        finally
          Job.Lock.Leave;
        end;
        ThreadRef := nil;
      end) then
    begin
      ShowThemedMessage('Change signature: the plugin is shutting down.');
      Exit;
    end;

    while True do
    begin
      var C, T: Integer;
      var S: string;
      var D: Boolean;
      Job.Lock.Enter;
      try
        C := Job.Cur; T := Job.Total; S := Job.Text; D := Job.Done;
      finally
        Job.Lock.Leave;
      end;
      if D then Break;
      if not Prog.Visible then Job.Stop.SetEvent;
      Prog.Step(C, T, S);
      Sleep(40);
    end;
  finally
    Prog.Free;
  end;

  if not Job.Res.Ok then
  begin
    ShowThemedMessage(Job.Res.Summary);
    Exit;
  end;
  var Dlg := TChangeSignatureDialog.CreateDialog(Application.MainForm, Job.Res);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  MCP tool "change_signature"
// ---------------------------------------------------------------------------

function ParamToJson(const P: TSigParam): TJSONObject;
begin
  Result := TJSONObject.Create;
  if P.Modifier <> '' then Result.AddPair('modifier', P.Modifier);
  Result.AddPair('name', P.Name);
  if P.TypeText <> '' then Result.AddPair('type', P.TypeText);
  if P.DefaultText <> '' then Result.AddPair('default', P.DefaultText);
end;

function ToolChangeSignature(AArgs: TJSONObject; AStop: THandle): string;
var
  F, Err, GatherErr: string;
  L1, C1: Integer;
  Inp: TSafeDeleteInput;
  Ok: Boolean;
begin
  F := AArgs.GetValue<string>('file', '');
  L1 := AArgs.GetValue<Integer>('line', 0);
  C1 := AArgs.GetValue<Integer>('column', 0);
  var DoApply := AArgs.GetValue<Boolean>('apply', False);
  if (F = '') or (L1 < 1) or (C1 < 1) then
    Exit(McpErr('arguments "file", "line" and "column" (1-based) are required'));
  F := ExpandFileName(F);
  Ok := False;
  if not McpRunOnMain(
    procedure
    var
      E: string;
    begin
      Ok := GatherScanInput(F, L1 - 1, C1 - 1, Inp, E);
      if not Ok then GatherErr := E;
    end, True, AStop, Err) then Exit(McpErr(Err));
  if not Ok then Exit(McpErr(GatherErr));
  var An := AnalyzeChangeSignature(Inp, AStop, nil);
  var J := TJSONObject.Create;
  J.AddPair('identifier', An.Identifier);
  J.AddPair('summary', An.Summary);
  if not An.Ok then Exit(McpOk(J));
  J.AddPair('kind', An.Kind);
  if An.Container <> '' then J.AddPair('container', An.Container);
  if An.ResultType <> '' then J.AddPair('result_type', An.ResultType);
  var PA := TJSONArray.Create;
  for var P in An.Params do PA.Add(ParamToJson(P));
  J.AddPair('parameters', PA);
  var HA := TJSONArray.Create;
  for var H in An.Headers do
  begin
    var O := TJSONObject.Create;
    O.AddPair('what', H.Caption);
    O.AddPair('file', An.Sources[H.FileIndex].FilePath);
    O.AddPair('line', TJSONNumber.Create(H.Line + 1));
    HA.Add(O);
  end;
  J.AddPair('family', HA);
  var CA := TJSONArray.Create;
  for var I := 0 to Min(High(An.Calls), 99) do
  begin
    var C := An.Calls[I];
    var O := TJSONObject.Create;
    O.AddPair('file', An.Sources[C.FileIndex].FilePath);
    O.AddPair('line', TJSONNumber.Create(C.Line + 1));
    O.AddPair('usage', ContextText[C.Context]);
    O.AddPair('text', C.Text);
    CA.Add(O);
  end;
  J.AddPair('call_count', TJSONNumber.Create(Length(An.Calls)));
  J.AddPair('calls', CA);
  var UA := TJSONArray.Create;
  for var U in An.Unverified do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', U.FilePath);
    O.AddPair('line', TJSONNumber.Create(U.Line + 1));
    O.AddPair('text', U.Text);
    O.AddPair('note', U.Note);
    UA.Add(O);
  end;
  J.AddPair('unverified', UA);
  var VA := TJSONArray.Create;
  for var V in An.Vetoes do VA.Add(V);
  J.AddPair('vetoes', VA);

  // the new parameter list, if given
  var NewArr := AArgs.GetValue('params') as TJSONArray;
  if NewArr = nil then Exit(McpOk(J));
  var Rows: TArray<TNewParam> := nil;
  for var V in NewArr do
  begin
    if not (V is TJSONObject) then
    begin
      J.Free;
      Exit(McpErr('"params" must be an array of objects'));
    end;
    var O := TJSONObject(V);
    var R := Default(TNewParam);
    R.Param.Name := O.GetValue<string>('name', '');
    R.Param.TypeText := O.GetValue<string>('type', '');
    R.Param.Modifier := O.GetValue<string>('modifier', '');
    R.Param.DefaultText := O.GetValue<string>('default', '');
    R.CallValue := O.GetValue<string>('value', '');
    var From := O.GetValue<string>('from', '');
    R.OldIndex := -1;
    if From <> '' then
    begin
      for var K := 0 to High(An.Params) do
        if SameText(An.Params[K].Name, From) then R.OldIndex := K;
      if R.OldIndex < 0 then
      begin
        J.Free;
        Exit(McpErr('"from": the method has no parameter "' + From + '"'));
      end;
      // an existing parameter keeps what is not given
      if O.GetValue('type') = nil then R.Param.TypeText := An.Params[R.OldIndex].TypeText;
      if O.GetValue('modifier') = nil then R.Param.Modifier := An.Params[R.OldIndex].Modifier;
      if O.GetValue('default') = nil then R.Param.DefaultText := An.Params[R.OldIndex].DefaultText;
      if R.Param.Name = '' then R.Param.Name := An.Params[R.OldIndex].Name;
    end;
    Rows := Rows + [R];
  end;
  var Plan := PlanChangeSignature(An, Rows);
  var PJ := TJSONObject.Create;
  PJ.AddPair('ok', TJSONBool.Create(Plan.Ok));
  var EA := TJSONArray.Create;
  for var E in Plan.Errors do EA.Add(E);
  PJ.AddPair('errors', EA);
  var WA := TJSONArray.Create;
  for var W in Plan.Warnings do WA.Add(W);
  PJ.AddPair('warnings', WA);
  var DA := TJSONArray.Create;
  for var E in Plan.Edits do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', An.Sources[E.FileIndex].FilePath);
    O.AddPair('line', TJSONNumber.Create(E.Line + 1));
    O.AddPair('what', E.What);
    O.AddPair('result', SigEditPreview(An, E));
    DA.Add(O);
  end;
  PJ.AddPair('edits', DA);
  J.AddPair('plan', PJ);
  if DoApply then
  begin
    var Applied := False;
    var AErr := '';
    if Plan.Ok and (Length(Plan.Edits) > 0) then
    begin
      if not McpRunOnMain(
        procedure
        begin
          Applied := ApplyChangeSignature(An, Plan, AErr);
        end, False, AStop, Err) then
        AErr := Err;
    end
    else
      AErr := 'the plan has errors or no edits';
    J.AddPair('applied', TJSONBool.Create(Applied));
    if AErr <> '' then J.AddPair('apply_error', AErr);
  end;
  Result := McpOk(J);
end;

initialization
  RegisterMcpTool('change_signature', ToolChangeSignature);

end.
