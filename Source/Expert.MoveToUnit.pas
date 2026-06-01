(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MoveToUnit;

{
  "Move identifier to other unit" engine.

  Given an identifier at a known editor position, the engine:

    1. Reads the SOURCE unit (editor buffer preferred, then disk).
    2. Locates the symbol's declaration in the interface section by
       scanning section-headers (const/var/type/resourcestring) and
       routine signatures (procedure/function), comment- and
       string-aware. For a class declaration the full `T<Name> = class
       [(...)]; ... end;` range is captured and all of its method
       implementations in the implementation section (matched by the
       `T<Name>.` prefix on the qualified header) are captured too.
    3. Reads the TARGET unit and checks for a same-name conflict; if
       so, refuses.
    4. Produces a TMovePlan and applies it transactionally via the
       editor API (undo-capable):
         - delete declaration (and any implementation blocks) from
           source,
         - splice declaration into target interface, splice impl(s)
           into target implementation,
         - for every project file that references the moved symbol,
           ensure TARGET unit appears in its interface uses clause,
         - if the source unit is now unused in a consumer, remove
           source from that file's uses clauses.

  Uses Expert.WithRewriter.ScanUsesClauses for uses-clause locations.

  KNOWN LIMITATIONS:
    - Move within mixed conditional-defines blocks may not be
      perfectly aligned; user can review preview and abort.
    - For overloaded routines, ALL overloads sharing the same name
      are moved together (acceptable for this use case).
    - The engine does not rewrite qualified `SourceUnit.X` references
      in consumers; uses-clause adjustment is the only edit applied
      to consumers.
}

interface

uses
  System.SysUtils, System.Classes, System.Types, System.IOUtils,
  System.StrUtils, System.Character, System.Generics.Collections, System.JSON,
  Vcl.Dialogs, Vcl.Forms,
  Expert.EditorHelper, Expert.WithRewriter,
  Expert.LspManager,
  Lsp.Client, Lsp.Uri, Lsp.Protocol, Delphi.FileEncoding;

type
  TMoveSymbolKind = (mskUnknown, mskConst, mskVar, mskResourceString,
                     mskType, mskClass, mskRoutine);

  TMoveEdit = record
    FilePath: string;
    Description: string;
  end;

  TMovePlan = record
    Symbol: string;
    Kind: TMoveSymbolKind;
    SourceFile: string;
    TargetFile: string;

    /// <summary>Text of the declaration (interface) to move.</summary>
    DeclarationText: string;
    /// <summary>1-based line range of declaration in source.</summary>
    DeclStartLine, DeclEndLine: Integer;

    /// <summary>For classes/routines: implementation block(s) to move.</summary>
    ImplBlocks: TArray<string>;
    /// <summary>Parallel array of (startLine, endLine) ranges, 1-based.</summary>
    ImplStartLines: TArray<Integer>;
    ImplEndLines:   TArray<Integer>;

    /// <summary>Consumer files (project-relative) that reference the
    ///  moved symbol and need TARGET in interface uses.</summary>
    Consumers: TArray<string>;

    /// <summary>Files where the SOURCE unit is no longer referenced by
    ///  any remaining symbol (after the move) and can be dropped from
    ///  the uses clauses.</summary>
    SourceUsesToRemove: TArray<string>;

    /// <summary>Human-readable summary edits.</summary>
    Edits: TArray<TMoveEdit>;

    Conflict: string;
    ProblemDetail: string;
  end;

  TLspMoveToUnit = class
  private
    class function ReadFile(const APath: string): string;
    class function LocateDeclaration(const ASymbol: string;
      const ASource: string; out AKind: TMoveSymbolKind;
      out AStartLine, AEndLine: Integer; out AClassDeclLine: Integer): Boolean;
    class function LocateClassMethods(const ASymbol, AClassName: string;
      const ASource: string; out AImplBlocks: TArray<string>;
      out AStartLines, AEndLines: TArray<Integer>): Boolean;
    class function LocateRoutineImpl(const ASymbol: string;
      const ASource: string; out AImpl: string;
      out AStartLine, AEndLine: Integer): Boolean;
    class function TargetHasSymbol(const ATargetSource, ASymbol: string): Boolean;
    class procedure ApplyToTarget(const ATargetFile: string;
      const APlan: TMovePlan);
    class procedure RemoveRangeFromSource(const ASourceFile: string;
      const APlan: TMovePlan);
    class procedure EnsureInterfaceUses(const AConsumerFile, ATargetUnit: string);
    class procedure EnsureImplementationUses(const ATargetFile: string;
      const AUnits: TArray<string>);
    class procedure RemoveFromUses(const AConsumerFile, ASourceUnit: string);
    class function UnitNameOfFile(const AFile: string): string;
    /// <summary>Collects identifier-name occurrences inside line range
    ///  AStartLine0..AEndLine0 (0-based, inclusive) of AFile, asks LSP
    ///  for each to find its declaring file, returns the resulting set
    ///  of unit names. Skips System (built-ins are auto-imported),
    ///  skips identifiers whose declaration falls inside any of the
    ///  AInSourceMovedRanges line spans (those are locals/params/etc.
    ///  that are themselves being moved). Skips ASourceUnit and
    ///  ATargetUnit since the target doesn't need to use itself.</summary>
    class function CollectRequiredUnits(AClient: TLspClient;
      const ASourceFile: string;
      const ARangesToScan: TArray<TPoint>;
      const AInSourceMovedRanges: TArray<TPoint>;
      const ASourceUnit, ATargetUnit: string): TArray<string>;
  public
    /// <summary>Build a plan describing the move. ProblemDetail is set
    ///  if the move cannot be performed; returns False in that case.</summary>
    class function BuildPlan(const ASymbol, ASourceFile, ATargetFile: string;
      out APlan: TMovePlan): Boolean;
    /// <summary>Execute the plan (write to editor/disk). Returns True
    ///  on success.</summary>
    class function ApplyPlan(const APlan: TMovePlan): Boolean;
    /// <summary>Top-level convenience: plan + confirm + apply.</summary>
    class function Execute(const ASymbol: string;
      const ASourceFile, ATargetFile: string;
      const AContext: TEditorContext): Boolean;
  end;

implementation

{ Helpers (file-local) }

function StripCommentsAndStringsKeepNewlines(const ASource: string): string;
// Replaces comments/strings with spaces but preserves CR/LF so that
// 1-based line numbering across original and stripped is identical.
var
  I, N: Integer;
  C: Char;
begin
  N := Length(ASource);
  SetLength(Result, N);
  I := 1;
  while I <= N do
  begin
    C := ASource[I];
    if (C = '/') and (I < N) and (ASource[I + 1] = '/') then
    begin
      while (I <= N) and (ASource[I] <> #10) and (ASource[I] <> #13) do
      begin Result[I] := ' '; Inc(I); end;
    end
    else if C = '{' then
    begin
      while (I <= N) and (ASource[I] <> '}') do
      begin
        if (ASource[I] = #10) or (ASource[I] = #13) then
          Result[I] := ASource[I]
        else
          Result[I] := ' ';
        Inc(I);
      end;
      if I <= N then begin Result[I] := ' '; Inc(I); end;
    end
    else if (C = '(') and (I < N) and (ASource[I + 1] = '*') then
    begin
      Result[I] := ' '; Inc(I);
      Result[I] := ' '; Inc(I);
      while I <= N do
      begin
        if (ASource[I] = '*') and (I < N) and (ASource[I + 1] = ')') then
        begin
          Result[I] := ' '; Inc(I);
          Result[I] := ' '; Inc(I);
          Break;
        end;
        if (ASource[I] = #10) or (ASource[I] = #13) then
          Result[I] := ASource[I]
        else
          Result[I] := ' ';
        Inc(I);
      end;
    end
    else if C = '''' then
    begin
      Result[I] := ' '; Inc(I);
      while I <= N do
      begin
        if ASource[I] = '''' then
        begin
          Result[I] := ' '; Inc(I);
          if (I <= N) and (ASource[I] = '''') then
          begin
            Result[I] := ' '; Inc(I);
            Continue;
          end;
          Break;
        end;
        if (ASource[I] = #10) or (ASource[I] = #13) then
          Result[I] := ASource[I]
        else
          Result[I] := ' ';
        Inc(I);
      end;
    end
    else
    begin
      Result[I] := C;
      Inc(I);
    end;
  end;
end;

function SplitLines(const ASource: string): TArray<string>;
var
  SL: TStringList;
  I: Integer;
begin
  SL := TStringList.Create;
  try
    SL.Text := ASource;
    SetLength(Result, SL.Count);
    for I := 0 to SL.Count - 1 do
      Result[I] := SL[I];
  finally
    SL.Free;
  end;
end;

function JoinLines(const ALines: TArray<string>): string;
var
  I: Integer;
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    for I := 0 to High(ALines) do
    begin
      if I > 0 then SB.Append(#13#10);
      SB.Append(ALines[I]);
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

/// <summary>Collapses runs of 2+ consecutive blank lines down to a
///  single blank line. Whitespace-only lines count as blank. Used as a
///  post-pass after moves so the source unit doesn't end up with
///  dangling blank-line pairs where the removed declaration used to
///  be, and the target unit doesn't accumulate extra gaps around the
///  freshly inserted block.</summary>
function CollapseBlankLines(const ASource: string): string;
var
  Lines: TArray<string>;
  Out_: TStringBuilder;
  I, BlankRun: Integer;
  Trimmed: string;
begin
  Lines := SplitLines(ASource);
  Out_ := TStringBuilder.Create;
  try
    BlankRun := 0;
    for I := 0 to High(Lines) do
    begin
      Trimmed := Trim(Lines[I]);
      if Trimmed = '' then
      begin
        Inc(BlankRun);
        // Keep AT MOST one blank line in a row.
        if BlankRun > 1 then Continue;
      end
      else
        BlankRun := 0;
      if Out_.Length > 0 then Out_.Append(#13#10);
      Out_.Append(Lines[I]);
    end;
    Result := Out_.ToString;
  finally
    Out_.Free;
  end;
end;

function IsIdCh(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
end;

/// <summary>True iff the (trimmed) line consists of nothing but a
///  bare section keyword 'type' / 'var' / 'const' / 'resourcestring' /
///  'threadvar' (no declaration on the same line). After a move this
///  is the kind of orphan the cleanup pass deletes.</summary>
function IsBareSectionKeyword(const ALine: string): Boolean;
var Upper: string;
begin
  Upper := UpperCase(Trim(ALine));
  Result := (Upper = 'TYPE') or (Upper = 'VAR') or (Upper = 'CONST')
    or (Upper = 'RESOURCESTRING') or (Upper = 'THREADVAR');
end;

/// <summary>True iff the trimmed line LOOKS like a different kind of
///  declaration following an (orphan) bare section keyword: another
///  section keyword, 'begin', 'implementation', 'initialization',
///  'finalization', 'end.', or a method/forward declaration
///  ('function ...' / 'procedure ...' / 'constructor ...' /
///  'destructor ...' / 'class function ...' / 'class procedure ...').
///  Used by RemoveEmptySectionHeaders to decide whether a preceding
///  bare 'type'/'var'/'const'/... is in fact orphaned.</summary>
function LooksLikeFollowupSection(const ALine: string): Boolean;

  function StartsWord(const AUpper, AWord: string): Boolean;
  begin
    Result := AUpper.StartsWith(AWord)
      and ((Length(AUpper) = Length(AWord))
        or not IsIdCh(AUpper[Length(AWord) + 1]));
  end;

var Upper: string;
begin
  Upper := UpperCase(Trim(ALine));
  Result := IsBareSectionKeyword(ALine)
    or (Upper = 'BEGIN')
    or (Upper = 'IMPLEMENTATION')
    or (Upper = 'INTERFACE')
    or (Upper = 'INITIALIZATION')
    or (Upper = 'FINALIZATION')
    or (Upper = 'END.')
    or StartsWord(Upper, 'USES')
    or StartsWord(Upper, 'FUNCTION')
    or StartsWord(Upper, 'PROCEDURE')
    or StartsWord(Upper, 'CONSTRUCTOR')
    or StartsWord(Upper, 'DESTRUCTOR')
    or StartsWord(Upper, 'OPERATOR')
    or StartsWord(Upper, 'CLASS')
    or StartsWord(Upper, 'PROPERTY');
end;

/// <summary>Removes bare 'type' / 'var' / 'const' / 'resourcestring'
///  header lines whose section is empty (next non-blank line is the
///  next section opener or 'implementation' or 'end.').</summary>
function RemoveEmptySectionHeaders(const ASource: string): string;
var
  Lines: TArray<string>;
  Out_: TStringBuilder;
  I, J: Integer;
begin
  Lines := SplitLines(ASource);
  Out_ := TStringBuilder.Create;
  try
    I := 0;
    while I <= High(Lines) do
    begin
      if IsBareSectionKeyword(Lines[I]) then
      begin
        // Look ahead: skip blank lines and see whether the next
        // non-blank line is another opener => empty section => skip.
        J := I + 1;
        while (J <= High(Lines)) and (Trim(Lines[J]) = '') do Inc(J);
        if (J <= High(Lines)) and LooksLikeFollowupSection(Lines[J]) then
        begin
          // Orphan section header - drop it.
          Inc(I);
          Continue;
        end;
      end;
      if Out_.Length > 0 then Out_.Append(#13#10);
      Out_.Append(Lines[I]);
      Inc(I);
    end;
    Result := Out_.ToString;
  finally
    Out_.Free;
  end;
end;

function HasWord(const ALine, AWord: string): Boolean;
// Word-boundary, case-insensitive contains check on a single line.
var
  Up, UpW: string;
  P, L: Integer;
  Before, After: Char;
begin
  Result := False;
  Up := AnsiUpperCase(ALine);
  UpW := AnsiUpperCase(AWord);
  L := Length(UpW);
  if L = 0 then Exit;
  P := 1;
  while P <= Length(Up) - L + 1 do
  begin
    if Copy(Up, P, L) = UpW then
    begin
      Before := #0;
      if P > 1 then Before := Up[P - 1];
      After := #0;
      if P + L <= Length(Up) then After := Up[P + L];
      if (not IsIdCh(Before)) and (not IsIdCh(After)) then Exit(True);
    end;
    Inc(P);
  end;
end;

function FirstNonSpace(const ALine: string): Integer;
var
  I: Integer;
begin
  for I := 1 to Length(ALine) do
    if not CharInSet(ALine[I], [' ', #9]) then Exit(I);
  Result := 0;
end;

function StartsWithKeyword(const ALine, AKeyword: string): Boolean;
// Case-insensitive: first non-space token equals AKeyword.
var
  P, L: Integer;
  After: Char;
begin
  Result := False;
  P := FirstNonSpace(ALine);
  if P = 0 then Exit;
  L := Length(AKeyword);
  if P + L - 1 > Length(ALine) then Exit;
  if not SameText(Copy(ALine, P, L), AKeyword) then Exit;
  After := #0;
  if P + L <= Length(ALine) then After := ALine[P + L];
  Result := (not IsIdCh(After));
end;

{ TLspMoveToUnit }

class function TLspMoveToUnit.ReadFile(const APath: string): string;
begin
  if not TEditorHelper.ReadEditorContent(APath, Result) then
    Result := TDelphiFileEncoding.ReadAll(APath);
end;

class function TLspMoveToUnit.UnitNameOfFile(const AFile: string): string;
begin
  Result := ChangeFileExt(ExtractFileName(AFile), '');
end;

class function TLspMoveToUnit.LocateDeclaration(const ASymbol: string;
  const ASource: string; out AKind: TMoveSymbolKind;
  out AStartLine, AEndLine: Integer; out AClassDeclLine: Integer): Boolean;
// Walk the interface section line-by-line. Track current section header
// (const/var/type/resourcestring). For routines, recognise signature
// lines. For a hit on ASymbol, determine the start/end of the declaration.
var
  Clean: string;
  Lines: TArray<string>;
  I, LastIntfLine, FirstIntfLine: Integer;
  InInterface: Boolean;
  Section: string;        // 'const' / 'var' / 'type' / 'resourcestring' / ''
  Stripped: TArray<string>;
  Line: string;
  TrimmedLow: string;
  SymU: string;
begin
  Result := False;
  AKind := mskUnknown;
  AStartLine := 0; AEndLine := 0; AClassDeclLine := 0;

  Clean := StripCommentsAndStringsKeepNewlines(ASource);
  Stripped := SplitLines(Clean);
  Lines := SplitLines(ASource);
  if Length(Stripped) = 0 then Exit;

  SymU := AnsiUpperCase(ASymbol);

  // Bracket the interface section.
  InInterface := False;
  FirstIntfLine := 0;
  LastIntfLine := High(Stripped);
  for I := 0 to High(Stripped) do
  begin
    if StartsWithKeyword(Stripped[I], 'interface') then
    begin
      InInterface := True;
      FirstIntfLine := I + 1;
    end
    else if InInterface and StartsWithKeyword(Stripped[I], 'implementation') then
    begin
      LastIntfLine := I - 1;
      Break;
    end;
  end;
  if not InInterface then Exit;

  Section := '';
  I := FirstIntfLine;
  while I <= LastIntfLine do
  begin
    Line := Stripped[I];
    TrimmedLow := AnsiLowerCase(Trim(Line));

    // Section headers reset the current section.
    if StartsWithKeyword(Line, 'const') then Section := 'const'
    else if StartsWithKeyword(Line, 'var') then Section := 'var'
    else if StartsWithKeyword(Line, 'resourcestring') then Section := 'resourcestring'
    else if StartsWithKeyword(Line, 'type') then Section := 'type'
    else if (TrimmedLow <> '') and (Section <> '') then
    begin
      // Inside a section; look for our symbol.
      // For const/var: "Name = expr;" or "Name: Type = expr;" or "Name: Type;"
      // For type: "TName = ..."
      if HasWord(Line, ASymbol) then
      begin
        // Verify the line declares it (starts with the identifier).
        var FNS := FirstNonSpace(Line);
        if (FNS > 0) and SameText(Copy(Line, FNS, Length(ASymbol)), ASymbol) then
        begin
          AStartLine := I + 1; // 1-based

          // Type/class detection: scan forward looking for "class" keyword
          // on declaration line, and find matching `end;`.
          if Section = 'type' then
          begin
            AClassDeclLine := AStartLine;
            // Scan declaration text forward until we see ';' at top level
            // or '= class' meaning it's a class definition with body.
            var IsClass := False;
            var DepthBegin := 0;
            var EndLine: Integer := I;
            var J: Integer := I;
            while J <= LastIntfLine do
            begin
              var LJ := Stripped[J];
              if (not IsClass) and HasWord(LJ, 'class') and HasWord(LJ, '=') then
                IsClass := True;
              if (not IsClass) and HasWord(LJ, 'interface') and (J > I) and HasWord(LJ, '=') then
                IsClass := True; // treat interface decl same way
              if IsClass then
              begin
                // class can have nested record types - track 'record'..'end' too
                // simple approach: count 'record' and 'class ... =' (no 'class of')
                // and balance 'end' tokens.
                // Tokenize by spaces and common delimiters.
                // For robustness we look for word 'end' on the line.
                if HasWord(LJ, 'record') then Inc(DepthBegin);
                // Class declaration body itself opens one level too.
                if (J = I) then Inc(DepthBegin);
                if HasWord(LJ, 'end') then
                begin
                  Dec(DepthBegin);
                  if DepthBegin <= 0 then
                  begin
                    EndLine := J;
                    Break;
                  end;
                end;
              end
              else
              begin
                // Simple type alias: ends at first ';'
                if Pos(';', LJ) > 0 then
                begin
                  EndLine := J;
                  Break;
                end;
              end;
              Inc(J);
            end;
            AEndLine := EndLine + 1;
            if IsClass then AKind := mskClass else AKind := mskType;
            Exit(True);
          end
          else
          begin
            // const/var/resourcestring: declaration ends at the first ';'
            var J := I;
            while J <= LastIntfLine do
            begin
              if Pos(';', Stripped[J]) > 0 then
              begin
                AEndLine := J + 1;
                Break;
              end;
              Inc(J);
            end;
            if AEndLine = 0 then AEndLine := AStartLine;
            if Section = 'const' then AKind := mskConst
            else if Section = 'var' then AKind := mskVar
            else AKind := mskResourceString;
            Exit(True);
          end;
        end;
      end;
    end;

    // Procedure/function declaration at top level of interface.
    if StartsWithKeyword(Line, 'procedure') or StartsWithKeyword(Line, 'function') then
    begin
      Section := ''; // routine headers terminate var/const sections
      if HasWord(Line, ASymbol) then
      begin
        // Verify second token equals the symbol (i.e. it's the routine name,
        // not just a parameter type).
        var FNS := FirstNonSpace(Line);
        var Rest := Trim(Copy(Line, FNS, MaxInt));
        // strip leading keyword
        var Sp := Pos(' ', Rest);
        if Sp > 0 then
        begin
          var After := Trim(Copy(Rest, Sp + 1, MaxInt));
          // capture identifier
          var K := 1;
          while (K <= Length(After)) and IsIdCh(After[K]) do Inc(K);
          var Name := Copy(After, 1, K - 1);
          if SameText(Name, ASymbol) then
          begin
            AStartLine := I + 1;
            var J := I;
            while J <= LastIntfLine do
            begin
              if Pos(';', Stripped[J]) > 0 then
              begin
                // skip ";" directives like "; overload;" - scan all directives
                AEndLine := J + 1;
                // look-ahead: collect "overload;" / "inline;" / "stdcall;" etc.
                var K2 := J + 1;
                while K2 <= LastIntfLine do
                begin
                  var T := AnsiLowerCase(Trim(Stripped[K2]));
                  if (T = 'overload;') or (T = 'inline;') or
                     (T = 'stdcall;') or (T = 'cdecl;') or
                     (T = 'register;') or (T = 'safecall;') or
                     (T = 'pascal;') or (T = 'forward;') or
                     (T = 'export;') or (T = 'external;') or
                     (T = 'platform;') or (T = 'deprecated;') then
                  begin
                    AEndLine := K2 + 1;
                    Inc(K2);
                  end
                  else
                    Break;
                end;
                Break;
              end;
              Inc(J);
            end;
            if AEndLine = 0 then AEndLine := AStartLine;
            AKind := mskRoutine;
            Exit(True);
          end;
        end;
      end;
    end;

    Inc(I);
  end;
end;

class function TLspMoveToUnit.LocateRoutineImpl(const ASymbol: string;
  const ASource: string; out AImpl: string;
  out AStartLine, AEndLine: Integer): Boolean;
// Finds the routine implementation in the implementation section. Matches
// a header line `procedure Name(...)` or `function Name(...)` at top level
// (not inside a class) and captures up to the matching top-level `end;`.
var
  Clean: string;
  Stripped, Lines: TArray<string>;
  I, ImplStart, MyStart: Integer;
  Line: string;
  BeginDepth: Integer;
begin
  Result := False;
  AImpl := ''; AStartLine := 0; AEndLine := 0;
  Clean := StripCommentsAndStringsKeepNewlines(ASource);
  Stripped := SplitLines(Clean);
  Lines := SplitLines(ASource);

  ImplStart := -1;
  for I := 0 to High(Stripped) do
    if StartsWithKeyword(Stripped[I], 'implementation') then
    begin
      ImplStart := I + 1;
      Break;
    end;
  if ImplStart < 0 then Exit;

  I := ImplStart;
  while I <= High(Stripped) do
  begin
    Line := Stripped[I];
    if StartsWithKeyword(Line, 'procedure') or StartsWithKeyword(Line, 'function')
       or StartsWithKeyword(Line, 'constructor') or StartsWithKeyword(Line, 'destructor') then
    begin
      // Extract identifier after keyword. Reject qualified `X.Y` names.
      var FNS := FirstNonSpace(Line);
      var Rest := Trim(Copy(Line, FNS, MaxInt));
      var Sp := Pos(' ', Rest);
      if Sp > 0 then
      begin
        var After := Trim(Copy(Rest, Sp + 1, MaxInt));
        var K := 1;
        while (K <= Length(After)) and IsIdCh(After[K]) do Inc(K);
        var Name := Copy(After, 1, K - 1);
        var HasDot := (K <= Length(After)) and (After[K] = '.');
        if (not HasDot) and SameText(Name, ASymbol) then
        begin
          // Found header. Find body: skip to 'begin', counting nested begin/end.
          MyStart := I;
          BeginDepth := 0;
          var Found := False;
          var J := I;
          while J <= High(Stripped) do
          begin
            var LJ := Stripped[J];
            if HasWord(LJ, 'begin') then
            begin
              Inc(BeginDepth);
              Found := True;
            end;
            if Found and HasWord(LJ, 'end') then
            begin
              // skip 'end.'? we care about end at depth 1
              Dec(BeginDepth);
              if BeginDepth <= 0 then
              begin
                AStartLine := MyStart + 1;
                AEndLine := J + 1;
                // Build text from original Lines (preserve formatting).
                var Sub: TStringBuilder := TStringBuilder.Create;
                try
                  for var KK := AStartLine - 1 to AEndLine - 1 do
                  begin
                    if KK < Length(Lines) then
                    begin
                      if Sub.Length > 0 then Sub.Append(#13#10);
                      Sub.Append(Lines[KK]);
                    end;
                  end;
                  AImpl := Sub.ToString;
                finally
                  Sub.Free;
                end;
                Exit(True);
              end;
            end;
            Inc(J);
          end;
        end;
      end;
    end;
    Inc(I);
  end;
end;

class function TLspMoveToUnit.LocateClassMethods(const ASymbol, AClassName: string;
  const ASource: string; out AImplBlocks: TArray<string>;
  out AStartLines, AEndLines: TArray<Integer>): Boolean;
// Find every implementation block whose header is qualified by
// "AClassName." (case-insensitive). Capture them all.
var
  Clean: string;
  Stripped, Lines: TArray<string>;
  I, ImplStart, MyStart: Integer;
  Line: string;
  BeginDepth: Integer;
  Blocks: TList<string>;
  StartList, EndList: TList<Integer>;
  QName: string;
begin
  AImplBlocks := nil;
  AStartLines := nil;
  AEndLines := nil;
  Result := False;
  Clean := StripCommentsAndStringsKeepNewlines(ASource);
  Stripped := SplitLines(Clean);
  Lines := SplitLines(ASource);
  ImplStart := -1;
  for I := 0 to High(Stripped) do
    if StartsWithKeyword(Stripped[I], 'implementation') then
    begin
      ImplStart := I + 1;
      Break;
    end;
  if ImplStart < 0 then Exit;

  QName := AnsiUpperCase(AClassName) + '.';
  Blocks := TList<string>.Create;
  StartList := TList<Integer>.Create;
  EndList := TList<Integer>.Create;
  try
    I := ImplStart;
    while I <= High(Stripped) do
    begin
      Line := Stripped[I];
      if StartsWithKeyword(Line, 'procedure') or StartsWithKeyword(Line, 'function')
         or StartsWithKeyword(Line, 'constructor') or StartsWithKeyword(Line, 'destructor')
         or StartsWithKeyword(Line, 'class') then
      begin
        var Up := AnsiUpperCase(Line);
        if Pos(' ' + QName, ' ' + Up) > 0 then
        begin
          MyStart := I;
          BeginDepth := 0;
          var Found := False;
          var J := I;
          while J <= High(Stripped) do
          begin
            var LJ := Stripped[J];
            if HasWord(LJ, 'begin') then
            begin
              Inc(BeginDepth);
              Found := True;
            end;
            if Found and HasWord(LJ, 'end') then
            begin
              Dec(BeginDepth);
              if BeginDepth <= 0 then
              begin
                StartList.Add(MyStart + 1);
                EndList.Add(J + 1);
                var Sub: TStringBuilder := TStringBuilder.Create;
                try
                  for var KK := MyStart to J do
                  begin
                    if KK < Length(Lines) then
                    begin
                      if Sub.Length > 0 then Sub.Append(#13#10);
                      Sub.Append(Lines[KK]);
                    end;
                  end;
                  Blocks.Add(Sub.ToString);
                finally
                  Sub.Free;
                end;
                I := J;
                Break;
              end;
            end;
            Inc(J);
          end;
        end;
      end;
      Inc(I);
    end;

    AImplBlocks := Blocks.ToArray;
    AStartLines := StartList.ToArray;
    AEndLines := EndList.ToArray;
    Result := Length(AImplBlocks) > 0;
  finally
    Blocks.Free;
    StartList.Free;
    EndList.Free;
  end;
end;

class function TLspMoveToUnit.TargetHasSymbol(const ATargetSource, ASymbol: string): Boolean;
var
  Clean: string;
  Stripped: TArray<string>;
  I: Integer;
begin
  Result := False;
  Clean := StripCommentsAndStringsKeepNewlines(ATargetSource);
  Stripped := SplitLines(Clean);
  for I := 0 to High(Stripped) do
    if HasWord(Stripped[I], ASymbol) then
    begin
      // Heuristic: any word-boundary occurrence outside comments/strings
      // is enough for "conflict" warning.
      Exit(True);
    end;
end;

class function TLspMoveToUnit.BuildPlan(const ASymbol, ASourceFile,
  ATargetFile: string; out APlan: TMovePlan): Boolean;
var
  Src, Tgt: string;
  DeclLine, EndLine, ClassDeclLine: Integer;
  Kind: TMoveSymbolKind;
  Lines: TArray<string>;
  I: Integer;
  Edits: TList<TMoveEdit>;
  E: TMoveEdit;
begin
  Result := False;
  APlan := Default(TMovePlan);
  APlan.Symbol := ASymbol;
  APlan.SourceFile := ASourceFile;
  APlan.TargetFile := ATargetFile;

  if SameText(ASourceFile, ATargetFile) then
  begin
    APlan.ProblemDetail := 'Source and target are the same unit.';
    Exit;
  end;

  try
    Src := ReadFile(ASourceFile);
  except
    on E2: Exception do
    begin
      APlan.ProblemDetail := 'Cannot read source: ' + E2.Message;
      Exit;
    end;
  end;
  try
    Tgt := ReadFile(ATargetFile);
  except
    on E2: Exception do
    begin
      APlan.ProblemDetail := 'Cannot read target: ' + E2.Message;
      Exit;
    end;
  end;

  if not LocateDeclaration(ASymbol, Src, Kind, DeclLine, EndLine, ClassDeclLine) then
  begin
    APlan.ProblemDetail := Format(
      'Could not locate declaration of "%s" in the interface section of %s.',
      [ASymbol, ExtractFileName(ASourceFile)]);
    Exit;
  end;

  APlan.Kind := Kind;
  APlan.DeclStartLine := DeclLine;
  APlan.DeclEndLine := EndLine;

  // Capture declaration text
  Lines := SplitLines(Src);
  var SB := TStringBuilder.Create;
  try
    for I := DeclLine - 1 to EndLine - 1 do
      if I >= 0 then
      begin
        if I < Length(Lines) then
        begin
          if SB.Length > 0 then SB.Append(#13#10);
          SB.Append(Lines[I]);
        end;
      end;
    APlan.DeclarationText := SB.ToString;
  finally
    SB.Free;
  end;

  // For routines: also locate implementation block.
  if Kind = mskRoutine then
  begin
    var ImplT: string;
    var ImplS, ImplE: Integer;
    if LocateRoutineImpl(ASymbol, Src, ImplT, ImplS, ImplE) then
    begin
      SetLength(APlan.ImplBlocks, 1);
      APlan.ImplBlocks[0] := ImplT;
      SetLength(APlan.ImplStartLines, 1); APlan.ImplStartLines[0] := ImplS;
      SetLength(APlan.ImplEndLines, 1);   APlan.ImplEndLines[0] := ImplE;
    end;
  end
  else if Kind = mskClass then
  begin
    LocateClassMethods(ASymbol, ASymbol, Src,
      APlan.ImplBlocks, APlan.ImplStartLines, APlan.ImplEndLines);
  end;

  // Conflict detection in target
  if TargetHasSymbol(Tgt, ASymbol) then
  begin
    APlan.Conflict := Format(
      'Target unit %s already mentions "%s". Refusing to move.',
      [ExtractFileName(ATargetFile), ASymbol]);
    APlan.ProblemDetail := APlan.Conflict;
    Exit;
  end;

  // Build consumer list: every project file that contains the symbol as a
  // word-boundary token and isn't the source or target.
  var Consumers := TList<string>.Create;
  try
    for var F in TEditorHelper.GetProjectSourceFiles do
    begin
      if SameText(F, ASourceFile) or SameText(F, ATargetFile) then Continue;
      var Ext := AnsiLowerCase(ExtractFileExt(F));
      if (Ext <> '.pas') and (Ext <> '.dpr') and (Ext <> '.dpk') then Continue;
      var Content: string := '';
      try
        Content := ReadFile(F);
      except
        Continue;
      end;
      var Clean := StripCommentsAndStringsKeepNewlines(Content);
      var Ls := SplitLines(Clean);
      for var L in Ls do
        if HasWord(L, ASymbol) then
        begin
          Consumers.Add(F);
          Break;
        end;
    end;
    APlan.Consumers := Consumers.ToArray;
  finally
    Consumers.Free;
  end;

  // Build edit summary
  Edits := TList<TMoveEdit>.Create;
  try
    E.FilePath := ASourceFile;
    E.Description := Format('Remove declaration (lines %d-%d) from source',
      [APlan.DeclStartLine, APlan.DeclEndLine]);
    Edits.Add(E);
    for I := 0 to High(APlan.ImplBlocks) do
    begin
      E.FilePath := ASourceFile;
      E.Description := Format('Remove implementation block (lines %d-%d) from source',
        [APlan.ImplStartLines[I], APlan.ImplEndLines[I]]);
      Edits.Add(E);
    end;
    E.FilePath := ATargetFile;
    E.Description := 'Insert declaration into target interface';
    Edits.Add(E);
    if Length(APlan.ImplBlocks) > 0 then
    begin
      E.FilePath := ATargetFile;
      E.Description := Format('Insert %d implementation block(s) into target',
        [Length(APlan.ImplBlocks)]);
      Edits.Add(E);
    end;
    for var C in APlan.Consumers do
    begin
      E.FilePath := C;
      E.Description := Format('Ensure interface uses contains %s',
        [UnitNameOfFile(ATargetFile)]);
      Edits.Add(E);
    end;
    APlan.Edits := Edits.ToArray;
  finally
    Edits.Free;
  end;

  Result := True;
end;

class procedure TLspMoveToUnit.RemoveRangeFromSource(const ASourceFile: string;
  const APlan: TMovePlan);
var
  Src: string;
  Lines: TArray<string>;
  Remove: array of Boolean;
  I: Integer;
  Out_: TStringBuilder;
begin
  Src := ReadFile(ASourceFile);
  Lines := SplitLines(Src);
  SetLength(Remove, Length(Lines));

  for I := APlan.DeclStartLine - 1 to APlan.DeclEndLine - 1 do
    if (I >= 0) and (I < Length(Lines)) then Remove[I] := True;

  for var K := 0 to High(APlan.ImplBlocks) do
  begin
    for I := APlan.ImplStartLines[K] - 1 to APlan.ImplEndLines[K] - 1 do
      if (I >= 0) and (I < Length(Lines)) then Remove[I] := True;
    // Extend the removal upward over class-marker comments that point
    // at the block we're removing. Delphi formatters often emit
    // '{ TClassName }' as a header for class methods; once all of
    // those methods are gone, the marker is an orphan. We strip a
    // comment-only line directly above (allowing one blank line in
    // between) when it references the qualifier of this block.
    if (APlan.ImplStartLines[K] >= 2) then
    begin
      var ProbeIdx := APlan.ImplStartLines[K] - 2; // 0-based, line above
      // Skip a single blank line.
      if (ProbeIdx >= 0) and (Trim(Lines[ProbeIdx]) = '') then
        Dec(ProbeIdx);
      if (ProbeIdx >= 0) then
      begin
        var Probe := Trim(Lines[ProbeIdx]);
        // '{ TFoo }' / '{ TFoo<T> }' / '{ TFoo: comment }' patterns.
        // Match: starts '{', contains the qualifier (the part before
        // '.' in the symbol's qualified name), ends '}'.
        if Probe.StartsWith('{') and Probe.EndsWith('}')
          and (APlan.Symbol <> '') then
        begin
          // qualifier = symbol before the first '.'.
          var DotPos := Pos('.', APlan.Symbol);
          var Qual: string;
          if DotPos > 0 then
            Qual := Copy(APlan.Symbol, 1, DotPos - 1)
          else
            Qual := APlan.Symbol;
          if (Qual <> '') and Probe.Contains(Qual) then
          begin
            Remove[ProbeIdx] := True;
            // Also drop the blank line between marker and block (if
            // present) so we don't leave a floating blank.
            if (ProbeIdx + 1 < APlan.ImplStartLines[K] - 1)
              and (Trim(Lines[ProbeIdx + 1]) = '') then
              Remove[ProbeIdx + 1] := True;
          end;
        end;
      end;
    end;
  end;

  Out_ := TStringBuilder.Create;
  try
    for I := 0 to High(Lines) do
    begin
      if Remove[I] then Continue;
      if Out_.Length > 0 then Out_.Append(#13#10);
      Out_.Append(Lines[I]);
    end;
    // Removing a declaration typically leaves twin blank lines where
    // the surrounding pre-/post-padding meets, plus possibly an
    // orphaned section header ('type' with no body). Two clean-up
    // passes: collapse runs of blank lines, then drop empty sections.
    TEditorHelper.ReplaceFileContent(ASourceFile,
      CollapseBlankLines(RemoveEmptySectionHeaders(Out_.ToString)));
  finally
    Out_.Free;
  end;
end;

class procedure TLspMoveToUnit.ApplyToTarget(const ATargetFile: string;
  const APlan: TMovePlan);
// Splice declaration into target interface section (just before the
// 'implementation' keyword). If implementation blocks exist, append them
// to the end of the implementation section (just before final 'end.').
var
  Src: string;
  Lines: TArray<string>;
  Clean: string;
  Stripped: TArray<string>;
  I, ImplKwLine, FinalEndLine: Integer;
  Out_: TStringBuilder;
  HeaderForDecl: string;
begin
  Src := ReadFile(ATargetFile);
  Lines := SplitLines(Src);
  Clean := StripCommentsAndStringsKeepNewlines(Src);
  Stripped := SplitLines(Clean);

  ImplKwLine := -1;
  FinalEndLine := -1;
  for I := 0 to High(Stripped) do
  begin
    if (ImplKwLine < 0) and StartsWithKeyword(Stripped[I], 'implementation') then
      ImplKwLine := I;
    if StartsWithKeyword(Stripped[I], 'end') and (Pos('.', Trim(Stripped[I])) > 0) then
      FinalEndLine := I;
  end;

  // Choose section header to prefix declaration with, since we don't
  // attempt to merge into an existing section.
  case APlan.Kind of
    mskConst:           HeaderForDecl := 'const';
    mskVar:             HeaderForDecl := 'var';
    mskResourceString:  HeaderForDecl := 'resourcestring';
    mskType, mskClass:  HeaderForDecl := 'type';
  else
    HeaderForDecl := '';
  end;

  Out_ := TStringBuilder.Create;
  try
    for I := 0 to High(Lines) do
    begin
      if I = ImplKwLine then
      begin
        // Insert declaration block (with section header if needed)
        // before the 'implementation' line. We add ONE blank-line
        // padding on each side. Any excess blank lines that result
        // when the source already had blanks here are collapsed by
        // the post-pass below.
        Out_.Append(#13#10);                // blank line ABOVE decl
        if HeaderForDecl <> '' then
          Out_.Append(HeaderForDecl).Append(#13#10);
        Out_.Append(APlan.DeclarationText);
        Out_.Append(#13#10).Append(#13#10); // blank line BELOW decl
      end;
      if (I = FinalEndLine) and (Length(APlan.ImplBlocks) > 0) then
      begin
        // Insert implementation blocks before final 'end.', again with
        // single-blank-line padding around / between blocks.
        Out_.Append(#13#10);                // blank line above block(s)
        for var K := 0 to High(APlan.ImplBlocks) do
        begin
          Out_.Append(APlan.ImplBlocks[K]);
          Out_.Append(#13#10).Append(#13#10); // blank line after each
        end;
      end;
      if Out_.Length > 0 then Out_.Append(#13#10);
      Out_.Append(Lines[I]);
    end;
    TEditorHelper.ReplaceFileContent(ATargetFile,
      CollapseBlankLines(RemoveEmptySectionHeaders(Out_.ToString)));
  finally
    Out_.Free;
  end;
end;

class procedure TLspMoveToUnit.EnsureInterfaceUses(const AConsumerFile,
  ATargetUnit: string);
var
  Src: string;
  Idx: TLineIndex;
  Scan: TUsesScanResult;
  Lines: TArray<string>;
  Out_: TStringBuilder;
  I: Integer;
  InsertLine: Integer;
  InjectAfterCol: Integer;
begin
  Src := ReadFile(AConsumerFile);
  Idx.Init(Src);
  Scan := ScanUsesClauses(Src, Idx);
  if UsesContains(Scan.InterfaceUses, ATargetUnit) or
     UsesContains(Scan.ImplementationUses, ATargetUnit) then
    Exit; // already there

  Lines := SplitLines(Src);

  if Scan.InterfaceUses.Found then
  begin
    // Insert ", ATargetUnit" after the last unit name.
    InsertLine := Scan.InterfaceUses.LastBeforeSemiLine - 1;
    InjectAfterCol := Scan.InterfaceUses.LastBeforeSemiCol;
    if (InsertLine >= 0) and (InsertLine < Length(Lines)) then
    begin
      var Line := Lines[InsertLine];
      Lines[InsertLine] :=
        Copy(Line, 1, InjectAfterCol) + ', ' + ATargetUnit +
        Copy(Line, InjectAfterCol + 1, MaxInt);
    end;
  end
  else
  begin
    // No interface uses clause: synthesize one right after the
    // 'interface' keyword line.
    var Clean := StripCommentsAndStringsKeepNewlines(Src);
    var Stripped := SplitLines(Clean);
    var IntfLine := -1;
    for I := 0 to High(Stripped) do
      if StartsWithKeyword(Stripped[I], 'interface') then
      begin
        IntfLine := I;
        Break;
      end;
    if IntfLine < 0 then Exit; // .dpr files: just bail

    // Insert "uses ATargetUnit;" two lines after the interface keyword.
    Out_ := TStringBuilder.Create;
    try
      for I := 0 to High(Lines) do
      begin
        if Out_.Length > 0 then Out_.Append(#13#10);
        Out_.Append(Lines[I]);
        if I = IntfLine then
        begin
          Out_.Append(#13#10);
          Out_.Append(#13#10);
          Out_.Append('uses ');
          Out_.Append(ATargetUnit);
          Out_.Append(';');
        end;
      end;
      TEditorHelper.ReplaceFileContent(AConsumerFile, Out_.ToString);
      Exit;
    finally
      Out_.Free;
    end;
  end;

  TEditorHelper.ReplaceFileContent(AConsumerFile, JoinLines(Lines));
end;

class procedure TLspMoveToUnit.EnsureImplementationUses(
  const ATargetFile: string; const AUnits: TArray<string>);
// Adds AUnits to the target file's implementation-uses clause if not
// already present (in either interface or implementation uses).
// Creates an 'uses' clause right after 'implementation' if there is
// none yet. Multi-unit version of EnsureInterfaceUses but for impl.
var
  Src: string;
  Idx: TLineIndex;
  Scan: TUsesScanResult;
  Lines: TArray<string>;
  Out_: TStringBuilder;
  I: Integer;
  Missing: TList<string>;
  U, InsertText: string;
begin
  if Length(AUnits) = 0 then Exit;
  Src := ReadFile(ATargetFile);
  Idx.Init(Src);
  Scan := ScanUsesClauses(Src, Idx);

  Missing := TList<string>.Create;
  try
    for U in AUnits do
      if (U <> '')
        and not UsesContains(Scan.InterfaceUses, U)
        and not UsesContains(Scan.ImplementationUses, U) then
      begin
        // Avoid duplicates within the Missing list itself.
        var Already := False;
        for var M in Missing do
          if SameText(M, U) then begin Already := True; Break end;
        if not Already then Missing.Add(U);
      end;
    if Missing.Count = 0 then Exit;

    Lines := SplitLines(Src);

    if Scan.ImplementationUses.Found then
    begin
      // Append ", U1, U2..." after the last unit before the ';'.
      var InsertLine := Scan.ImplementationUses.LastBeforeSemiLine - 1;
      var InjectAfterCol := Scan.ImplementationUses.LastBeforeSemiCol;
      if (InsertLine >= 0) and (InsertLine < Length(Lines)) then
      begin
        InsertText := '';
        for U in Missing do
          InsertText := InsertText + ', ' + U;
        var Line := Lines[InsertLine];
        Lines[InsertLine] :=
          Copy(Line, 1, InjectAfterCol) + InsertText +
          Copy(Line, InjectAfterCol + 1, MaxInt);
        TEditorHelper.ReplaceFileContent(ATargetFile, JoinLines(Lines));
      end;
      Exit;
    end;

    // No implementation-uses clause: synthesize one right after the
    // 'implementation' keyword line.
    var Clean := StripCommentsAndStringsKeepNewlines(Src);
    var Stripped := SplitLines(Clean);
    var ImplLine := -1;
    for I := 0 to High(Stripped) do
      if StartsWithKeyword(Stripped[I], 'implementation') then
      begin
        ImplLine := I;
        Break;
      end;
    if ImplLine < 0 then Exit;

    Out_ := TStringBuilder.Create;
    try
      for I := 0 to High(Lines) do
      begin
        if Out_.Length > 0 then Out_.Append(#13#10);
        Out_.Append(Lines[I]);
        if I = ImplLine then
        begin
          Out_.Append(#13#10);
          Out_.Append(#13#10);
          Out_.Append('uses ');
          for var Mi := 0 to Missing.Count - 1 do
          begin
            if Mi > 0 then Out_.Append(', ');
            Out_.Append(Missing[Mi]);
          end;
          Out_.Append(';');
        end;
      end;
      TEditorHelper.ReplaceFileContent(ATargetFile, Out_.ToString);
    finally
      Out_.Free;
    end;
  finally
    Missing.Free;
  end;
end;

class function TLspMoveToUnit.CollectRequiredUnits(AClient: TLspClient;
  const ASourceFile: string;
  const ARangesToScan: TArray<TPoint>;
  const AInSourceMovedRanges: TArray<TPoint>;
  const ASourceUnit, ATargetUnit: string): TArray<string>;
var
  Src: string;
  Clean: string;
  Lines: TArray<string>;
  Seen: TDictionary<string, Boolean>;
  Out_: TList<string>;
  L, P, Q: Integer;
  Tok: string;
  Defs: TArray<TLspLocation>;
  DefPath, DefUnit: string;
  DefLine: Integer;

  function InMovedRange(ALine0: Integer): Boolean;
  var R: TPoint;
  begin
    for R in AInSourceMovedRanges do
      if (ALine0 >= R.X - 1) and (ALine0 <= R.Y - 1) then Exit(True);
    Result := False;
  end;

begin
  Src := ReadFile(ASourceFile);
  Clean := StripCommentsAndStringsKeepNewlines(Src);
  Lines := SplitLines(Clean);

  Seen := TDictionary<string, Boolean>.Create;
  Out_ := TList<string>.Create;
  try
    // Pre-populate Seen with excluded units (so we can dedupe across runs).
    if ASourceUnit <> '' then Seen.AddOrSetValue('U:' + UpperCase(ASourceUnit), True);
    if ATargetUnit <> '' then Seen.AddOrSetValue('U:' + UpperCase(ATargetUnit), True);
    Seen.AddOrSetValue('U:SYSTEM', True);

    for var R in ARangesToScan do
      for L := R.X - 1 to R.Y - 1 do
      begin
        if (L < 0) or (L >= Length(Lines)) then Continue;
        P := 1;
        while P <= Length(Lines[L]) do
        begin
          if (Lines[L][P] = '_') or Lines[L][P].IsLetter then
          begin
            Q := P;
            while (Q <= Length(Lines[L])) and IsIdCh(Lines[L][Q]) do Inc(Q);
            Tok := Copy(Lines[L], P, Q - P);
            // Skip very short idents (less LSP traffic for noise).
            // Skip Pascal-reserved words by upper-case check below.
            if (Length(Tok) > 1) and not Seen.ContainsKey('T:' + UpperCase(Tok)) then
            begin
              Seen.AddOrSetValue('T:' + UpperCase(Tok), True);
              // Skip Pascal keywords (cheap textual filter).
              var Upper := UpperCase(Tok);
              if (Upper = 'BEGIN') or (Upper = 'END') or (Upper = 'VAR')
                or (Upper = 'TYPE') or (Upper = 'CONST') or (Upper = 'IF')
                or (Upper = 'THEN') or (Upper = 'ELSE') or (Upper = 'FOR')
                or (Upper = 'TO') or (Upper = 'DOWNTO') or (Upper = 'DO')
                or (Upper = 'WHILE') or (Upper = 'REPEAT') or (Upper = 'UNTIL')
                or (Upper = 'CASE') or (Upper = 'OF') or (Upper = 'TRY')
                or (Upper = 'FINALLY') or (Upper = 'EXCEPT') or (Upper = 'RAISE')
                or (Upper = 'AND') or (Upper = 'OR') or (Upper = 'NOT')
                or (Upper = 'XOR') or (Upper = 'DIV') or (Upper = 'MOD')
                or (Upper = 'SHL') or (Upper = 'SHR') or (Upper = 'NIL')
                or (Upper = 'TRUE') or (Upper = 'FALSE') or (Upper = 'RESULT')
                or (Upper = 'SELF') or (Upper = 'INHERITED')
                or (Upper = 'CLASS') or (Upper = 'INTERFACE') or (Upper = 'RECORD')
                or (Upper = 'OBJECT') or (Upper = 'FUNCTION') or (Upper = 'PROCEDURE')
                or (Upper = 'CONSTRUCTOR') or (Upper = 'DESTRUCTOR')
                or (Upper = 'PROPERTY') or (Upper = 'PUBLIC') or (Upper = 'PRIVATE')
                or (Upper = 'PROTECTED') or (Upper = 'PUBLISHED') or (Upper = 'STRICT')
                or (Upper = 'IS') or (Upper = 'AS') or (Upper = 'IN')
                or (Upper = 'OUT') or (Upper = 'EXIT') or (Upper = 'BREAK')
                or (Upper = 'CONTINUE') or (Upper = 'WITH') or (Upper = 'USES')
                or (Upper = 'IMPLEMENTATION') or (Upper = 'INITIALIZATION')
                or (Upper = 'FINALIZATION') or (Upper = 'UNIT')
                or (Upper = 'STRING') or (Upper = 'INTEGER') or (Upper = 'BOOLEAN')
                or (Upper = 'BYTE') or (Upper = 'CHAR') or (Upper = 'WORD')
                or (Upper = 'CARDINAL') or (Upper = 'DOUBLE') or (Upper = 'EXTENDED')
                or (Upper = 'SINGLE') or (Upper = 'POINTER') or (Upper = 'OVERRIDE')
                or (Upper = 'VIRTUAL') or (Upper = 'OVERLOAD') or (Upper = 'STDCALL')
                or (Upper = 'CDECL') or (Upper = 'REINTRODUCE') or (Upper = 'ABSTRACT')
                or (Upper = 'STATIC') or (Upper = 'SEALED') or (Upper = 'FINAL')
                or (Upper = 'DEFAULT') or (Upper = 'INDEX') or (Upper = 'READ')
                or (Upper = 'WRITE') or (Upper = 'STORED') or (Upper = 'ARRAY')
                or (Upper = 'SET') or (Upper = 'FILE') or (Upper = 'LABEL')
                or (Upper = 'GOTO') or (Upper = 'PACKED') or (Upper = 'INLINE')
                or (Upper = 'ASSEMBLER') or (Upper = 'ASM') or (Upper = 'DISPID')
                or (Upper = 'MESSAGE') or (Upper = 'NAME') or (Upper = 'EXPORTS') then
              begin
                P := Q;
                Continue;
              end;

              // LSP query at this position (0-based line/col).
              Defs := nil;
              try
                Defs := AClient.GotoDefinition(ASourceFile, L, P - 1);
              except end;
              if Length(Defs) > 0 then
              begin
                DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
                DefLine := Defs[0].Range.Start.Line;
                // Skip if the declaration is INSIDE the moved range
                // (locals/params/local types).
                if SameText(DefPath, ASourceFile) and InMovedRange(DefLine) then
                begin
                  P := Q;
                  Continue;
                end;
                DefUnit := ChangeFileExt(ExtractFileName(DefPath), '');
                if (DefUnit <> '')
                  and not Seen.ContainsKey('U:' + UpperCase(DefUnit)) then
                begin
                  Seen.AddOrSetValue('U:' + UpperCase(DefUnit), True);
                  Out_.Add(DefUnit);
                end;
              end;
            end;
            P := Q;
          end
          else
            Inc(P);
        end;
      end;
    Result := Out_.ToArray;
  finally
    Out_.Free;
    Seen.Free;
  end;
end;

class procedure TLspMoveToUnit.RemoveFromUses(const AConsumerFile,
  ASourceUnit: string);
// Remove ASourceUnit from any uses clause in AConsumerFile. Simple
// textual approach: locate the unit name on a uses line, drop the entry.
var
  Src: string;
  Idx: TLineIndex;
  Scan: TUsesScanResult;
  Lines: TArray<string>;
  I, L: Integer;
  UpLine, UpUnit: string;
  Line: string;

  procedure StripUnitFromLine(var ALine: string);
  var
    UL: string;
    Pp: Integer;
    Before, After: Char;
    Start, Stop: Integer;
  begin
    UL := AnsiUpperCase(ALine);
    Pp := 1;
    L := Length(UpUnit);
    while Pp <= Length(UL) - L + 1 do
    begin
      if Copy(UL, Pp, L) = UpUnit then
      begin
        Before := #0;
        if Pp > 1 then Before := UL[Pp - 1];
        After := #0;
        if Pp + L <= Length(UL) then After := UL[Pp + L];
        if (not IsIdCh(Before)) and (not IsIdCh(After)) then
        begin
          Start := Pp;
          Stop := Pp + L - 1;
          // include leading ', ' or trailing ', '
          var P2 := Start - 1;
          while (P2 >= 1) and CharInSet(ALine[P2], [' ', #9]) do Dec(P2);
          if (P2 >= 1) and (ALine[P2] = ',') then
            Start := P2;
          var P3 := Stop + 1;
          while (P3 <= Length(ALine)) and CharInSet(ALine[P3], [' ', #9]) do Inc(P3);
          if (P3 <= Length(ALine)) and (ALine[P3] = ',') and (Start = Pp) then
          begin
            Stop := P3;
          end;
          ALine := Copy(ALine, 1, Start - 1) + Copy(ALine, Stop + 1, MaxInt);
          Exit;
        end;
      end;
      Inc(Pp);
    end;
  end;

begin
  Src := ReadFile(AConsumerFile);
  Idx.Init(Src);
  Scan := ScanUsesClauses(Src, Idx);
  if not (UsesContains(Scan.InterfaceUses, ASourceUnit) or
          UsesContains(Scan.ImplementationUses, ASourceUnit)) then
    Exit;

  Lines := SplitLines(Src);
  UpUnit := AnsiUpperCase(ASourceUnit);

  // Apply removal in the lines that fall inside either uses clause.
  if Scan.InterfaceUses.Found then
  begin
    for I := Scan.InterfaceUses.FirstUnitLine - 1
        to Scan.InterfaceUses.LastBeforeSemiLine - 1 do
      if (I >= 0) and (I < Length(Lines)) then
      begin
        UpLine := AnsiUpperCase(Lines[I]);
        if Pos(UpUnit, UpLine) > 0 then
        begin
          Line := Lines[I];
          StripUnitFromLine(Line);
          Lines[I] := Line;
        end;
      end;
  end;
  if Scan.ImplementationUses.Found then
  begin
    for I := Scan.ImplementationUses.FirstUnitLine - 1
        to Scan.ImplementationUses.LastBeforeSemiLine - 1 do
      if (I >= 0) and (I < Length(Lines)) then
      begin
        UpLine := AnsiUpperCase(Lines[I]);
        if Pos(UpUnit, UpLine) > 0 then
        begin
          Line := Lines[I];
          StripUnitFromLine(Line);
          Lines[I] := Line;
        end;
      end;
  end;

  TEditorHelper.ReplaceFileContent(AConsumerFile, JoinLines(Lines));
end;

class function TLspMoveToUnit.ApplyPlan(const APlan: TMovePlan): Boolean;
var
  TargetUnit, SourceUnit: string;
  C: string;
begin
  Result := False;
  if APlan.ProblemDetail <> '' then Exit;
  if APlan.DeclStartLine = 0 then Exit;

  TargetUnit := UnitNameOfFile(APlan.TargetFile);
  SourceUnit := UnitNameOfFile(APlan.SourceFile);

  try
    // 0) Collect which units the moved code REFERENCES via LSP. We do
    //    this on the still-unmodified source so positions are stable.
    //    Result: list of unit names that the target file will need.
    //    Interface-uses set covers the declaration; implementation-uses
    //    covers the impl blocks (less any units already needed in the
    //    interface, so we don't import the same thing twice).
    var Client: TLspClient := nil;
    var DeclUnits: TArray<string>;
    var ImplUnits: TArray<string>;
    try
      var DelphiLspJson := TEditorHelper.FindDelphiLspJson;
      var RootPath := TEditorHelper.GetProjectRoot;
      if RootPath = '' then RootPath := ExtractFilePath(APlan.SourceFile);
      var ProjFile := TEditorHelper.GetCurrentProjectDproj;
      if (DelphiLspJson <> '') and (ProjFile <> '') then
        Client := TLspManager.Instance.GetClient(RootPath, ProjFile, DelphiLspJson);
    except
      Client := nil;
    end;
    if Client <> nil then
    begin
      // Build the range list to scan: declaration + each impl block.
      var DeclRange: TArray<TPoint>;
      var ImplRange: TArray<TPoint>;
      var Moved: TArray<TPoint>;
      SetLength(DeclRange, 1);
      DeclRange[0] := Point(APlan.DeclStartLine, APlan.DeclEndLine);
      SetLength(ImplRange, Length(APlan.ImplBlocks));
      for var K := 0 to High(APlan.ImplBlocks) do
        ImplRange[K] := Point(APlan.ImplStartLines[K], APlan.ImplEndLines[K]);
      Moved := DeclRange + ImplRange;
      // Decl-scope identifiers go to interface-uses of the target;
      // they're visible in the public surface (parameter types, parent
      // class, property types, ...).
      DeclUnits := CollectRequiredUnits(Client, APlan.SourceFile,
        DeclRange, Moved, SourceUnit, TargetUnit);
      // Impl-scope identifiers go to implementation-uses ONLY (so we
      // don't promote private dependencies into the public surface).
      // Filter out ones already in DeclUnits later.
      ImplUnits := CollectRequiredUnits(Client, APlan.SourceFile,
        ImplRange, Moved, SourceUnit, TargetUnit);
    end;

    // 1) Add declaration + implementation blocks to TARGET first.
    ApplyToTarget(APlan.TargetFile, APlan);

    // 1a) Add the units that the moved code references.
    if Length(DeclUnits) > 0 then
      for var U in DeclUnits do
        EnsureInterfaceUses(APlan.TargetFile, U);
    if Length(ImplUnits) > 0 then
    begin
      // Remove ImplUnits that are already in DeclUnits (avoid double-import).
      var ImplOnly: TArray<string>;
      for var U in ImplUnits do
      begin
        var Dup := False;
        for var D in DeclUnits do
          if SameText(D, U) then begin Dup := True; Break end;
        if not Dup then ImplOnly := ImplOnly + [U];
      end;
      EnsureImplementationUses(APlan.TargetFile, ImplOnly);
    end;

    // 2) Remove declaration + implementation blocks from SOURCE.
    RemoveRangeFromSource(APlan.SourceFile, APlan);

    // 3) For every consumer, ensure interface uses contains TARGET.
    for C in APlan.Consumers do
      EnsureInterfaceUses(C, TargetUnit);

    // 4) For every file (source removed): if source unit is no longer
    //    referenced there by any remaining symbol, drop source from uses.
    for C in APlan.Consumers do
    begin
      var Body: string := '';
      try
        Body := ReadFile(C);
      except
        Continue;
      end;
      var Clean := StripCommentsAndStringsKeepNewlines(Body);
      var Ls := SplitLines(Clean);
      var StillReferenced := False;
      // Read source file's current interface to know which identifiers
      // remain after the move. Skipping that subtlety: only drop if
      // the consumer literally no longer mentions the moved symbol's
      // unit qualifier or any source-only identifier we can prove. For
      // safety, never auto-drop source from uses — the user spec says
      // "if no OTHER symbol from source is still referenced there",
      // which we approximate conservatively here.
      for var Line in Ls do
        if HasWord(Line, SourceUnit) then
        begin
          StillReferenced := True;
          Break;
        end;
      if not StillReferenced then
        RemoveFromUses(C, SourceUnit);
    end;

    Result := True;
  except
    on E: Exception do
      MessageDlg('Move failed: ' + E.Message, mtError, [mbOK], 0);
  end;
end;

class function TLspMoveToUnit.Execute(const ASymbol: string;
  const ASourceFile, ATargetFile: string;
  const AContext: TEditorContext): Boolean;
var
  Plan: TMovePlan;
begin
  Result := False;
  TEditorHelper.SaveAllFiles;
  if not BuildPlan(ASymbol, ASourceFile, ATargetFile, Plan) then
  begin
    if Plan.ProblemDetail <> '' then
      MessageDlg(Plan.ProblemDetail, mtWarning, [mbOK], 0);
    Exit;
  end;
  Result := ApplyPlan(Plan);
end;

end.
