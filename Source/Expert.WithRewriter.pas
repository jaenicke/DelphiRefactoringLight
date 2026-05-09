(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WithRewriter;

// Computes the textual rewrite for one TWithOccurrence using DelphiLSP.
//
// Strategy
// --------
// Per target T_i (in declared order):
//   1) LSP textDocument/definition on the target's last identifier
//      -> location of the declaring identifier (field, var, parameter).
//   2) Read the declaration line, extract the type name via a small
//      scan ("Ident : Type" or ": Type" form).
//   3) LSP textDocument/definition on the type name in step 2
//      -> location of the type declaration (e.g. "TFoo = class").
//   4) Compute the class source range via begin/try/case/asm-vs-end
//      balancing, starting from the declaration line. This gives a
//      (TypeFile, ClassStartLine, ClassEndLine) triple per target.
//
// Per body identifier (bare name not preceded by ".", "&" or "@"):
//   1) LSP textDocument/definition
//   2) Walk targets from right to left; on the first whose type file
//      matches the definition file AND whose class range contains the
//      definition line -> qualify with that target's prefix.
//   3) Pascal 'with A, B do ...' semantics: when a name is declared in
//      multiple targets, the RIGHTMOST target wins.
//   4) No matching target -> leave the identifier unchanged.
//
// Complex targets are auto-introduced as inline vars
// (Delphi 10.3+ syntax). A "complex target" is any target whose textual
// expression is not a dotted identifier path Ident(.Ident)*.
// The inline var is named __withN where N is the 1-based target index.
//
// Multi-target ("with A, B do") rewrites ARE produced automatically:
// per body identifier the rewriter picks the rightmost target whose
// class contains the definition. Complex targets (non-dotted) are
// hoisted to '__withN' inline-vars in declaration order so the
// expression is evaluated exactly once, matching the original
// 'with' semantics.

interface

uses
  System.Classes, System.SysUtils, System.Generics.Collections,
  Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Expert.WithScanner;

type
  TWithRewriteIssue = (
    /// <summary>"with A, B do …" — multi-target. Not auto-rewritten.</summary>
    wriMultipleTargets,
    /// <summary>Target type could not be resolved via LSP.</summary>
    wriTypeUnresolved,
    /// <summary>Class range could not be determined.</summary>
    wriClassRangeUnknown,
    /// <summary>Inline-var introduction would shadow an existing name
    ///  even after numeric bump (extremely unlikely).</summary>
    wriNameClash
  );

  TWithRewriteIssues = set of TWithRewriteIssue;

  /// <summary>How the rewriter decided which target a body-identifier
  ///  belongs to. Used by the dialog's debug view.</summary>
  TWithDebugMatch = (
    /// <summary>No target accepted the ref - left unchanged.</summary>
    dmNone,
    /// <summary>Matched via direct (parsed) class member.</summary>
    dmMember,
    /// <summary>Matched via LSP GotoDefinition (typically an inherited
    ///  member).</summary>
    dmLsp
  );

  TWithDebugTargetInfo = record
    Expression: string;
    TypeFile: string;
    ClassStartLine: Integer;
    ClassEndLine: Integer;
    Members: TArray<string>;
    InlineVarName: string;
    QualifyPrefix: string;
    Resolved: Boolean;
  end;

  TWithDebugRefInfo = record
    Name: string;
    Line: Integer;
    Col: Integer;
    /// <summary>True iff LSP returned at least one location.</summary>
    LspHadResult: Boolean;
    LspFile: string;
    LspLine: Integer;
    MatchSource: TWithDebugMatch;
    /// <summary>Index into Targets[] of the target that won, -1 if none.</summary>
    MatchedTargetIdx: Integer;
    AppliedPrefix: string;
  end;

  TWithRewriteDebug = record
    Targets: TArray<TWithDebugTargetInfo>;
    Refs: TArray<TWithDebugRefInfo>;
  end;

  TWithRewriteResult = record
    /// <summary>Source file the occurrence belongs to.</summary>
    FileName: string;
    /// <summary>Original occurrence (carried along for UI display).</summary>
    Occurrence: TWithOccurrence;
    /// <summary>Original full source range that will be replaced
    ///  (covers from the 'with' keyword through the body terminator).</summary>
    ReplaceRange: TWithSourceRange;
    /// <summary>Original source slice corresponding to ReplaceRange.</summary>
    OriginalText: string;
    /// <summary>Rewritten replacement text. Empty if HasUnsupportedShape
    ///  or Issues prevent automatic rewriting.</summary>
    NewText: string;
    /// <summary>Set of analysis problems. If non-empty, the dialog
    ///  should mark the entry but still show OriginalText so the user
    ///  can review the location.</summary>
    Issues: TWithRewriteIssues;
    /// <summary>Diagnostic snapshot of what LSP / Member-fallback said
    ///  about each target type and each body identifier. Used by the
    ///  dialog's debug view; never read by the rewriter itself.</summary>
    Debug: TWithRewriteDebug;
    /// <summary>True iff the rewrite is mechanically valid (no Issues
    ///  AND NewText non-empty).</summary>
    function IsAutoRewritable: Boolean;
  end;

  TWithRewriter = class
  public
    /// <summary>Computes the rewrite for one occurrence. Reads ASource
    ///  to extract the body text and to determine class ranges in
    ///  unrelated files via TFile.ReadAllLines as needed. Never raises;
    ///  on LSP error sets the appropriate issue and returns a partial
    ///  result with OriginalText filled.</summary>
    class function Rewrite(AClient: TLspClient; const AFileName: string;
      const ASource: string; const AOccurrence: TWithOccurrence): TWithRewriteResult; static;
  end;

implementation

uses
  System.IOUtils, System.Character, System.StrUtils;

{ Helpers: 1-based <-> string-index, line slicing }

type
  TLineIndex = record
  private
    FStarts: TArray<Integer>;  // 1-based string index of first char of each line; FStarts[0] => line 1
    FSrc: string;
  public
    procedure Init(const ASource: string);
    function LineCount: Integer;
    function LineText(ALine1: Integer): string;
    /// <summary>Converts (1-based line, 1-based column) to a 1-based
    ///  string index into the source.</summary>
    function PosToIndex(ALine1, ACol1: Integer): Integer;
    /// <summary>Slice from (StartLine,StartCol) through (EndLine,EndCol) inclusive.</summary>
    function Slice(AStart, AEnd: TWithSourcePos): string;
  end;

procedure TLineIndex.Init(const ASource: string);
var
  I, N: Integer;
  Starts: TList<Integer>;
begin
  FSrc := ASource;
  Starts := TList<Integer>.Create;
  try
    Starts.Add(1);
    N := Length(ASource);
    I := 1;
    while I <= N do
    begin
      if ASource[I] = #13 then
      begin
        if (I < N) and (ASource[I + 1] = #10) then
          Inc(I);
        Starts.Add(I + 1);
      end
      else if ASource[I] = #10 then
        Starts.Add(I + 1);
      Inc(I);
    end;
    FStarts := Starts.ToArray;
  finally
    Starts.Free;
  end;
end;

function TLineIndex.LineCount: Integer;
begin
  Result := Length(FStarts);
end;

function TLineIndex.LineText(ALine1: Integer): string;
var
  Start, Stop: Integer;
begin
  if (ALine1 < 1) or (ALine1 > Length(FStarts)) then Exit('');
  Start := FStarts[ALine1 - 1];
  if ALine1 < Length(FStarts) then
    Stop := FStarts[ALine1] - 1
  else
    Stop := Length(FSrc) + 1;
  // Strip trailing CR/LF
  while (Stop - 1 >= Start) and
    ((FSrc[Stop - 1] = #10) or (FSrc[Stop - 1] = #13)) do
    Dec(Stop);
  Result := Copy(FSrc, Start, Stop - Start);
end;

function TLineIndex.PosToIndex(ALine1, ACol1: Integer): Integer;
begin
  if (ALine1 < 1) or (ALine1 > Length(FStarts)) then Exit(0);
  Result := FStarts[ALine1 - 1] + ACol1 - 1;
end;

function TLineIndex.Slice(AStart, AEnd: TWithSourcePos): string;
var
  IStart, IEnd: Integer;
begin
  IStart := PosToIndex(AStart.Line, AStart.Col);
  IEnd := PosToIndex(AEnd.Line, AEnd.Col);
  if (IStart < 1) or (IEnd < IStart) then Exit('');
  Result := Copy(FSrc, IStart, IEnd - IStart + 1);
end;

{ Helpers: identifier classification }

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := (C = '_') or C.IsLetter;
end;

function IsIdentCont(C: Char): Boolean; inline;
begin
  Result := (C = '_') or C.IsLetterOrDigit;
end;

function IsPascalKeyword(const S: string): Boolean;
const
  KW: array[0..68] of string = (
    'and','array','as','asm','begin','case','class','const','constructor',
    'destructor','dispinterface','div','do','downto','else','end','except',
    'exports','file','finalization','finally','for','function','goto','if',
    'implementation','in','inherited','initialization','inline','interface',
    'is','label','library','mod','nil','not','object','of','or','out',
    'packed','procedure','program','property','raise','record','repeat',
    'resourcestring','set','shl','shr','string','then','threadvar','to',
    'try','type','unit','until','uses','var','while','with','xor','true',
    'false','self','result'
  );
var
  I: Integer;
begin
  for I := Low(KW) to High(KW) do
    if SameText(S, KW[I]) then Exit(True);
  Result := False;
end;

/// <summary>Splits a simple dotted-identifier path into its segments.
///  Returns nil for non-simple expressions.</summary>
function SplitDottedSegments(const AExpr: string): TArray<string>;
var
  Trimmed: string;
  Segs: TList<string>;
  Start, I: Integer;
begin
  Trimmed := AExpr.Trim;
  Segs := TList<string>.Create;
  try
    Start := 1;
    I := 1;
    while I <= Length(Trimmed) do
    begin
      if Trimmed[I] = '.' then
      begin
        Segs.Add(Copy(Trimmed, Start, I - Start));
        Start := I + 1;
      end;
      Inc(I);
    end;
    if Start <= Length(Trimmed) then
      Segs.Add(Copy(Trimmed, Start, Length(Trimmed) - Start + 1));
    Result := Segs.ToArray;
  finally
    Segs.Free;
  end;
end;

/// <summary>Returns the leading-whitespace prefix of the first
///  non-empty line in ABodyOnly (the indentation used for body
///  statements). Falls back to ADefaultIndent when not derivable.</summary>
function GetBodyStatementIndent(const ABodyOnly, ADefaultIndent: string): string;
var
  I, LineStart: Integer;
  WS: string;
begin
  I := 1;
  while I <= Length(ABodyOnly) do
  begin
    if (ABodyOnly[I] = #10) or (ABodyOnly[I] = #13) then
    begin
      LineStart := I + 1;
      if (ABodyOnly[I] = #13) and (LineStart <= Length(ABodyOnly))
        and (ABodyOnly[LineStart] = #10) then
        Inc(LineStart);
      WS := '';
      while (LineStart <= Length(ABodyOnly))
        and ((ABodyOnly[LineStart] = ' ') or (ABodyOnly[LineStart] = #9)) do
      begin
        WS := WS + ABodyOnly[LineStart];
        Inc(LineStart);
      end;
      if (LineStart <= Length(ABodyOnly))
        and (ABodyOnly[LineStart] <> #10) and (ABodyOnly[LineStart] <> #13) then
        Exit(WS);
      I := LineStart;
    end
    else
      Inc(I);
  end;
  Result := ADefaultIndent;
end;

/// <summary>Derives a temp-var name from a target's last segment using
///  Delphi convention: drop a leading 'F' member-prefix, then prepend
///  'L'. Returns '' for inputs that don't lead to a valid identifier.</summary>
function DeriveTempName(const ALastSegment: string): string;
var
  Base: string;
begin
  Base := ALastSegment;
  // Strip leading F-prefix when the next char is upper-case (field
  // convention TFoo.FBar -> 'Bar').
  if (Length(Base) >= 2) and (Base[1] = 'F') and Base[2].IsUpper then
    Base := Copy(Base, 2, MaxInt);
  if (Base = '') or not IsIdentStart(Base[1]) then Exit('');
  Result := 'L' + Base;
end;

/// <summary>True iff the expression is a dotted-identifier path
///  Ident(.Ident)* — i. e. safe to use directly (no side effects, no
///  parens).</summary>
function IsSimpleDottedTarget(const AExpr: string): Boolean;
var
  I: Integer;
  ExpectIdent: Boolean;
  Trimmed: string;
begin
  Trimmed := AExpr.Trim;
  if Trimmed = '' then Exit(False);
  Result := True;
  ExpectIdent := True;
  I := 1;
  while I <= Length(Trimmed) do
  begin
    if ExpectIdent then
    begin
      if not IsIdentStart(Trimmed[I]) then Exit(False);
      Inc(I);
      while (I <= Length(Trimmed)) and IsIdentCont(Trimmed[I]) do Inc(I);
      ExpectIdent := False;
    end
    else
    begin
      if Trimmed[I] <> '.' then Exit(False);
      Inc(I);
      ExpectIdent := True;
    end;
  end;
  if ExpectIdent then Exit(False); // ended with '.'
end;

{ Type-name extraction from a declaration line }

/// <summary>Given a line like
///    "FFoo: TFoo;"  /  "FFoo: TFoo = nil;"  /  "Foo, Bar: TBaz;"
///  extracts the type-name token after the ':'. Returns '' on
///  failure. Works on field, var and parameter declarations.</summary>
function ExtractTypeNameFromDeclarationLine(const ALine: string;
  out ATypeNameCol1: Integer): string;
var
  I, N: Integer;
  ColonPos: Integer;
begin
  Result := '';
  ATypeNameCol1 := 0;
  N := Length(ALine);
  ColonPos := 0;

  // Find ':' that is not part of ':=' (not really a concern in decls
  // but doesn't hurt) and not inside parentheses (parameter list start).
  var ParenDepth := 0;
  I := 1;
  while I <= N do
  begin
    case ALine[I] of
      '(': Inc(ParenDepth);
      ')': if ParenDepth > 0 then Dec(ParenDepth);
      ':':
        if (ParenDepth = 0) and ((I = N) or (ALine[I + 1] <> '=')) then
        begin
          ColonPos := I;
          Break;
        end;
    end;
    Inc(I);
  end;
  if ColonPos = 0 then Exit;

  // Skip whitespace after ':'
  I := ColonPos + 1;
  while (I <= N) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  if (I > N) or not IsIdentStart(ALine[I]) then Exit;

  ATypeNameCol1 := I;
  var Start := I;
  while (I <= N) and IsIdentCont(ALine[I]) do Inc(I);
  Result := Copy(ALine, Start, I - Start);
end;

{ Class range detection in an arbitrary file }

/// <summary>Reads AFile and returns (StartLine, EndLine) of the class
///  declaration that begins at AClassDeclLine1 (1-based).
///  We balance "begin/try/case/asm/record" openers vs. "end" closers.
///  For "TFoo = class(...)" the opener is the keyword 'class' on the
///  declaration line itself; we treat the declaration line as opening
///  one block and look for the matching end.
///  Falls back to (AClassDeclLine1, AClassDeclLine1) on error.</summary>
procedure FindClassRangeInFile(const AFile: string; AClassDeclLine1: Integer;
  out AStartLine, AEndLine: Integer);
var
  Source: string;
  Index: TLineIndex;
  I, Depth: Integer;
  Line, Upper: string;
  IsForwardDecl: Boolean;
begin
  AStartLine := AClassDeclLine1;
  AEndLine := AClassDeclLine1;

  if not TFile.Exists(AFile) then Exit;
  try
    Source := TFile.ReadAllText(AFile);
  except
    Exit;
  end;
  Index.Init(Source);
  if (AClassDeclLine1 < 1) or (AClassDeclLine1 > Index.LineCount) then Exit;

  Line := Index.LineText(AClassDeclLine1);
  Upper := UpperCase(Line);

  // Forward declaration "TFoo = class;" / "TFoo = interface;" - no body.
  // Must end with ';' (after trim). Bodyless inheritance lists like
  // "TFoo = class" (open body follows) must NOT be classified as forward.
  var Trimmed := Trim(Upper);
  IsForwardDecl :=
    (Trimmed.EndsWith('CLASS;') or Trimmed.EndsWith('INTERFACE;') or
     Trimmed.EndsWith('OBJECT;') or Trimmed.EndsWith('DISPINTERFACE;'))
    and not Upper.Contains('CLASS(');
  if IsForwardDecl then
  begin
    AEndLine := AClassDeclLine1;
    Exit;
  end;

  // The declaration line itself starts a record/class body. Some lines
  // already contain the matching 'end' (single-line declaration). Walk
  // forward token-by-token (line-by-line is good enough; we do a coarse
  // scan looking for 'end' as a whole word).
  Depth := 1;
  for I := AClassDeclLine1 + 1 to Index.LineCount do
  begin
    Line := Index.LineText(I);
    Upper := UpperCase(Line);

    // Count nested record / class / case / try / begin openers and end closers.
    // We do a token-aware scan: split on non-ident chars.
    var P := 1;
    while P <= Length(Upper) do
    begin
      if IsIdentStart(Upper[P]) then
      begin
        var Q := P;
        while (Q <= Length(Upper)) and IsIdentCont(Upper[Q]) do Inc(Q);
        var Word := Copy(Upper, P, Q - P);
        if (Word = 'RECORD') or (Word = 'CASE') or (Word = 'TRY')
          or (Word = 'BEGIN') or (Word = 'ASM') then
          Inc(Depth)
        else if (Word = 'CLASS') or (Word = 'INTERFACE')
                or (Word = 'OBJECT') or (Word = 'DISPINTERFACE') then
        begin
          // class method markers, class-helper etc — only count as opener
          // if it looks like a type definition. Conservative: don't count;
          // these are rare inside class bodies and would cause Depth to
          // never balance.
        end
        else if Word = 'END' then
        begin
          Dec(Depth);
          if Depth = 0 then
          begin
            AEndLine := I;
            Exit;
          end;
        end;
        P := Q;
      end
      else
        Inc(P);
    end;
  end;

  // Unbalanced — fall back to declaration line only.
  AEndLine := AClassDeclLine1;
end;

{ Scanning bare identifiers in a body }

type
  TBodyIdentRef = record
    Name: string;
    /// <summary>Position of the first char of the identifier
    ///  (1-based line/col, file-relative).</summary>
    Pos: TWithSourcePos;
    /// <summary>String index of the first char in the source.</summary>
    StartIdx: Integer;
    /// <summary>String index just past the last char.</summary>
    EndIdx: Integer;
  end;

/// <summary>Scans the inner body and returns all bare identifiers
///  (those NOT preceded by '.', '&' or '@'). Strings/comments are
///  skipped. Pascal keywords are excluded.
///  All positions are absolute file positions, derived from the line
///  index passed in, so the caller can map back to the original source.</summary>
function CollectBodyIdentifiers(const ASource: string;
  const AInner: TWithSourceRange; const AIndex: TLineIndex): TArray<TBodyIdentRef>;
var
  Refs: TList<TBodyIdentRef>;
  StartIdx, EndIdx, I, N: Integer;
  Line, Col: Integer;
  Ref: TBodyIdentRef;
  PrevSig: Char;
begin
  Refs := TList<TBodyIdentRef>.Create;
  try
    StartIdx := AIndex.PosToIndex(AInner.StartPos.Line, AInner.StartPos.Col);
    EndIdx := AIndex.PosToIndex(AInner.EndPos.Line, AInner.EndPos.Col);
    if (StartIdx < 1) or (EndIdx < StartIdx) then Exit(Refs.ToArray);

    Line := AInner.StartPos.Line;
    Col := AInner.StartPos.Col;
    PrevSig := #0;
    I := StartIdx;
    N := EndIdx;
    while I <= N do
    begin
      var Ch := ASource[I];

      // Track newlines for position tracking
      if Ch = #13 then
      begin
        Inc(I);
        if (I <= N) and (ASource[I] = #10) then Inc(I);
        Inc(Line);
        Col := 1;
        Continue;
      end;
      if Ch = #10 then
      begin
        Inc(I);
        Inc(Line);
        Col := 1;
        Continue;
      end;

      // Skip strings
      if Ch = '''' then
      begin
        Inc(I); Inc(Col);
        while I <= N do
        begin
          if ASource[I] = '''' then
          begin
            if (I < N) and (ASource[I + 1] = '''') then
            begin
              Inc(I, 2); Inc(Col, 2);
            end
            else
            begin
              Inc(I); Inc(Col);
              Break;
            end;
          end
          else if (ASource[I] = #10) or (ASource[I] = #13) then
            Break
          else
          begin
            Inc(I); Inc(Col);
          end;
        end;
        PrevSig := '''';
        Continue;
      end;

      // Skip char constants
      if Ch = '#' then
      begin
        Inc(I); Inc(Col);
        if (I <= N) and (ASource[I] = '$') then begin Inc(I); Inc(Col); end;
        while (I <= N) and CharInSet(ASource[I], ['0'..'9','a'..'f','A'..'F']) do
        begin
          Inc(I); Inc(Col);
        end;
        PrevSig := '#';
        Continue;
      end;

      // Line comment
      if (Ch = '/') and (I < N) and (ASource[I + 1] = '/') then
      begin
        while (I <= N) and (ASource[I] <> #10) and (ASource[I] <> #13) do
        begin
          Inc(I); Inc(Col);
        end;
        Continue;
      end;

      // Curly comment
      if Ch = '{' then
      begin
        while (I <= N) and (ASource[I] <> '}') do
        begin
          if ASource[I] = #10 then begin Inc(Line); Col := 1; Inc(I); Continue; end;
          if ASource[I] = #13 then
          begin
            Inc(I);
            if (I <= N) and (ASource[I] = #10) then Inc(I);
            Inc(Line); Col := 1;
            Continue;
          end;
          Inc(I); Inc(Col);
        end;
        if I <= N then begin Inc(I); Inc(Col); end;
        Continue;
      end;

      // Paren-star comment
      if (Ch = '(') and (I < N) and (ASource[I + 1] = '*') then
      begin
        Inc(I, 2); Inc(Col, 2);
        while I <= N do
        begin
          if (ASource[I] = '*') and (I < N) and (ASource[I + 1] = ')') then
          begin
            Inc(I, 2); Inc(Col, 2);
            Break;
          end;
          if ASource[I] = #10 then begin Inc(Line); Col := 1; Inc(I); Continue; end;
          if ASource[I] = #13 then
          begin
            Inc(I);
            if (I <= N) and (ASource[I] = #10) then Inc(I);
            Inc(Line); Col := 1;
            Continue;
          end;
          Inc(I); Inc(Col);
        end;
        Continue;
      end;

      // Identifier?
      if IsIdentStart(Ch) then
      begin
        var IdStart := I;
        var IdStartCol := Col;
        var IdStartLine := Line;
        while (I <= N) and IsIdentCont(ASource[I]) do
        begin
          Inc(I); Inc(Col);
        end;
        var Name := Copy(ASource, IdStart, I - IdStart);
        var IsMember := (PrevSig = '.') or (PrevSig = '&') or (PrevSig = '@');
        if (not IsMember) and (not IsPascalKeyword(Name)) then
        begin
          Ref.Name := Name;
          Ref.Pos.Line := IdStartLine;
          Ref.Pos.Col := IdStartCol;
          Ref.StartIdx := IdStart;
          Ref.EndIdx := I;
          Refs.Add(Ref);
        end;
        PrevSig := 'a';
        Continue;
      end;

      // Whitespace -> don't update PrevSig
      if (Ch = ' ') or (Ch = #9) then
      begin
        Inc(I); Inc(Col);
        Continue;
      end;

      // Other punctuation — record as PrevSig
      PrevSig := Ch;
      Inc(I); Inc(Col);
    end;

    Result := Refs.ToArray;
  finally
    Refs.Free;
  end;
end;

{ Per-target type info via LSP }

type
  TTargetTypeInfo = record
    Expression: string;            // verbatim from the source
    InlineVarName: string;         // empty if target is simple-dotted
    QualifyPrefix: string;         // what to prepend to body identifiers
    TypeFile: string;
    ClassStartLine: Integer;       // 1-based
    ClassEndLine: Integer;         // 1-based, inclusive
    /// <summary>Names of fields / properties / methods declared directly
    ///  in this class/record body (not inherited). Used as a fallback when
    ///  DelphiLSP fails to resolve bare body-identifiers in multi-target
    ///  with-clauses.</summary>
    Members: TArray<string>;
    Resolved: Boolean;
  end;

/// <summary>Parses the class/record body between AStartLine+1 and
///  AEndLine-1 of AFile and returns the names of directly-declared
///  fields, properties and methods. Inherited members are NOT included
///  - by design, since LSP already handles those.
///  Visibility section keywords (private/protected/public/published,
///  optionally prefixed by 'strict') and modifier keywords like 'class'
///  are skipped.</summary>
function CollectClassMembers(const AFile: string;
  AStartLine, AEndLine: Integer): TArray<string>;
var
  Source, Line, Upper, Word: string;
  Index: TLineIndex;
  I, P, Q: Integer;
  List: TList<string>;
  Pending: TStringList;
  HadFirstWord: Boolean;
  FirstWord: string;
  procedure FlushPending;
  var S: string;
  begin
    for S in Pending do
      if (S <> '') then List.Add(S);
    Pending.Clear;
  end;
begin
  Result := nil;
  if (AStartLine < 1) or (AEndLine <= AStartLine) then Exit;
  if not TFile.Exists(AFile) then Exit;
  try
    Source := TFile.ReadAllText(AFile);
  except
    Exit;
  end;
  Index.Init(Source);

  List := TList<string>.Create;
  Pending := TStringList.Create;
  try
    for I := AStartLine + 1 to AEndLine - 1 do
    begin
      if (I < 1) or (I > Index.LineCount) then Continue;
      Line := Index.LineText(I);
      Upper := UpperCase(Line);

      // Strip line comment after //
      P := Pos('//', Line);
      if P > 0 then
      begin
        Line := Copy(Line, 1, P - 1);
        Upper := Copy(Upper, 1, P - 1);
      end;

      // Detect method-like lines: starts (after optional 'class') with
      // procedure/function/constructor/destructor/operator. The name is
      // the next identifier.
      HadFirstWord := False;
      FirstWord := '';
      P := 1;
      while P <= Length(Upper) do
      begin
        if IsIdentStart(Upper[P]) then
        begin
          Q := P;
          while (Q <= Length(Upper)) and IsIdentCont(Upper[Q]) do Inc(Q);
          Word := Copy(Upper, P, Q - P);
          if not HadFirstWord then
          begin
            HadFirstWord := True;
            FirstWord := Word;
            // Skip 'strict' / 'class' modifiers and visibility keywords.
            if (Word = 'STRICT') or (Word = 'CLASS') then
            begin
              HadFirstWord := False;
              P := Q;
              Continue;
            end;
            if (Word = 'PRIVATE') or (Word = 'PROTECTED') or (Word = 'PUBLIC')
              or (Word = 'PUBLISHED') or (Word = 'TYPE') or (Word = 'CONST')
              or (Word = 'VAR') or (Word = 'THREADVAR') then
            begin
              // section header - skip whole line
              Pending.Clear;
              Break;
            end;
            if (Word = 'PROCEDURE') or (Word = 'FUNCTION')
              or (Word = 'CONSTRUCTOR') or (Word = 'DESTRUCTOR')
              or (Word = 'OPERATOR') then
            begin
              // Next ident is the method name.
              P := Q;
              while (P <= Length(Line)) and ((Line[P] = ' ') or (Line[P] = #9)) do Inc(P);
              if (P <= Length(Line)) and IsIdentStart(Line[P]) then
              begin
                Q := P;
                while (Q <= Length(Line)) and IsIdentCont(Line[Q]) do Inc(Q);
                List.Add(Copy(Line, P, Q - P));
              end;
              Break;
            end;
            if Word = 'PROPERTY' then
            begin
              // Next ident is property name.
              P := Q;
              while (P <= Length(Line)) and ((Line[P] = ' ') or (Line[P] = #9)) do Inc(P);
              if (P <= Length(Line)) and IsIdentStart(Line[P]) then
              begin
                Q := P;
                while (Q <= Length(Line)) and IsIdentCont(Line[Q]) do Inc(Q);
                List.Add(Copy(Line, P, Q - P));
              end;
              Break;
            end;
            if (Word = 'END') or (Word = 'CASE') or (Word = 'BEGIN') then
            begin
              Pending.Clear;
              Break;
            end;
            // Otherwise: it's a field name (or list). Fall through.
            Pending.Add(Copy(Line, P, Q - P));
            P := Q;
            Continue;
          end
          else
          begin
            // Subsequent words on a non-method line: only collect if
            // followed/preceded by ',' (field list). Stop on ':'.
            // We already added FirstWord; just keep collecting names
            // until we see ':'.
            Pending.Add(Copy(Line, P, Q - P));
            P := Q;
            Continue;
          end;
        end
        else if Upper[P] = ':' then
        begin
          // type follows - flush pending field names and stop.
          FlushPending;
          Break;
        end
        else if Upper[P] = '=' then
        begin
          // const-style declaration ("X = 5;") - keep first name only.
          if Pending.Count > 0 then
            List.Add(Pending[0]);
          Pending.Clear;
          Break;
        end
        else if Upper[P] = ';' then
        begin
          // statement terminator without ':' - drop, but note: forward
          // method decls already added. Just clear pending.
          Pending.Clear;
          Break;
        end;
        Inc(P);
      end;

      // End-of-line: if a field declaration spans multiple lines we
      // might still have Pending entries waiting for ':'. Conservative:
      // discard them — multi-line field-list declarations are rare.
      Pending.Clear;
    end;

    Result := List.ToArray;
  finally
    Pending.Free;
    List.Free;
  end;
end;

function ResolveTargetType(AClient: TLspClient; const AFileName: string;
  const ATarget: TWithTarget): TTargetTypeInfo;
var
  DefLocs: TArray<TLspLocation>;
  DeclFile: string;
  DeclLine: Integer;
  DeclLineText: string;
  TypeName: string;
  TypeNameCol1: Integer;
  TypeDefLocs: TArray<TLspLocation>;
  TmpIndex: TLineIndex;
  TmpSource: string;
begin
  Result := Default(TTargetTypeInfo);
  Result.Expression := ATarget.Expression;
  Result.Resolved := False;

  // Position of the LAST identifier of the (possibly dotted) target.
  // ATarget.Range.EndPos points at its last char; LSP wants the position
  // ON the identifier — using the start-of-last-token would be cleaner
  // but pointing anywhere on the identifier works for DelphiLSP.
  // For simple-dotted targets we use the EndPos as-is. For complex
  // targets (function calls) we fall back to EndPos as well; LSP is
  // robust to that.
  try
    DefLocs := AClient.GotoDefinition(AFileName,
      ATarget.Range.EndPos.Line - 1, ATarget.Range.EndPos.Col - 1);
  except
    Exit;
  end;
  if Length(DefLocs) = 0 then Exit;

  DeclFile := DefLocs[0].Uri;
  // TLspLocation.Uri is a file:// URI. Convert via Lsp.Uri if needed —
  // for now do the simple file:/// strip.
  DeclFile := TLspUri.FileUriToPath(DeclFile);

  DeclLine := DefLocs[0].Range.Start.Line + 1; // -> 1-based

  // Read the declaration line and extract the type name.
  if not TFile.Exists(DeclFile) then Exit;
  try
    TmpSource := TFile.ReadAllText(DeclFile);
  except
    Exit;
  end;
  TmpIndex.Init(TmpSource);
  DeclLineText := TmpIndex.LineText(DeclLine);
  TypeName := ExtractTypeNameFromDeclarationLine(DeclLineText, TypeNameCol1);
  if TypeName = '' then Exit;

  // Now resolve the type name itself.
  try
    TypeDefLocs := AClient.GotoDefinition(DeclFile, DeclLine - 1, TypeNameCol1 - 1);
  except
    Exit;
  end;
  if Length(TypeDefLocs) = 0 then Exit;

  Result.TypeFile := TLspUri.FileUriToPath(TypeDefLocs[0].Uri);

  FindClassRangeInFile(Result.TypeFile,
    TypeDefLocs[0].Range.Start.Line + 1,
    Result.ClassStartLine, Result.ClassEndLine);

  // Parse direct (non-inherited) member names. Used as a fallback when
  // DelphiLSP fails to resolve bare body-identifiers in a multi-target
  // with-clause (it appears LSP only checks the rightmost target's class).
  Result.Members := CollectClassMembers(Result.TypeFile,
    Result.ClassStartLine, Result.ClassEndLine);

  Result.Resolved := True;
end;

{ TWithRewriteResult }

function TWithRewriteResult.IsAutoRewritable: Boolean;
begin
  Result := (Issues = []) and (NewText <> '');
end;

{ TWithRewriter }

class function TWithRewriter.Rewrite(AClient: TLspClient; const AFileName: string;
  const ASource: string; const AOccurrence: TWithOccurrence): TWithRewriteResult;
var
  Index: TLineIndex;
  ReplaceStart, ReplaceEnd: TWithSourcePos;
  Targets: TArray<TTargetTypeInfo>;
  I: Integer;
  BodyRefs: TArray<TBodyIdentRef>;
  Ref: TBodyIdentRef;
  RefDef: TArray<TLspLocation>;
  RefDefFile: string;
  RefDefLine: Integer;
  IndentStr: string;
  LineStartIdx: Integer;
  InnerText, BodyOnly: string;
  HasInlineVar: Boolean;
  Builder: TStringBuilder;
begin
  Result := Default(TWithRewriteResult);
  Result.FileName := AFileName;
  Result.Occurrence := AOccurrence;

  Index.Init(ASource);

  // ReplaceRange covers from 'with' through the body terminator.
  ReplaceStart := AOccurrence.KeywordPos;
  ReplaceEnd := AOccurrence.BodyRange.EndPos;
  Result.ReplaceRange.StartPos := ReplaceStart;
  Result.ReplaceRange.EndPos := ReplaceEnd;
  Result.OriginalText := Index.Slice(ReplaceStart, ReplaceEnd);

  if Length(AOccurrence.Targets) = 0 then
  begin
    Include(Result.Issues, wriClassRangeUnknown);
    Exit;
  end;

  // Resolve each target's type.
  SetLength(Targets, Length(AOccurrence.Targets));
  for I := 0 to High(AOccurrence.Targets) do
  begin
    Targets[I] := ResolveTargetType(AClient, AFileName, AOccurrence.Targets[I]);
    if not Targets[I].Resolved then
    begin
      Include(Result.Issues, wriTypeUnresolved);
      // Snapshot what we have so the debug view can still show partial
      // resolution.
      SetLength(Result.Debug.Targets, Length(Targets));
      for var DI := 0 to High(Targets) do
      begin
        Result.Debug.Targets[DI].Expression     := Targets[DI].Expression;
        Result.Debug.Targets[DI].TypeFile       := Targets[DI].TypeFile;
        Result.Debug.Targets[DI].ClassStartLine := Targets[DI].ClassStartLine;
        Result.Debug.Targets[DI].ClassEndLine   := Targets[DI].ClassEndLine;
        Result.Debug.Targets[DI].Members        := Targets[DI].Members;
        Result.Debug.Targets[DI].InlineVarName  := Targets[DI].InlineVarName;
        Result.Debug.Targets[DI].QualifyPrefix  := Targets[DI].QualifyPrefix;
        Result.Debug.Targets[DI].Resolved       := Targets[DI].Resolved;
      end;
      Exit;
    end;
    // Pick the inline-var name for this target.
    //   - single-letter dotted/single segment ("a", "x")     -> no temp
    //     (already short; introducing 'La := a' is just noise)
    //   - dotted path or non-trivial single ident            -> 'L' + last
    //   - complex expression (call, brackets, ...)           -> '__withN'
    // Cross-target collisions and body-identifier collisions fall back
    // to '__withN'.
    if IsSimpleDottedTarget(Targets[I].Expression) then
    begin
      var Segs := SplitDottedSegments(Targets[I].Expression);
      var Last := '';
      if Length(Segs) > 0 then Last := Segs[High(Segs)];
      if (Length(Segs) = 1) and (Length(Last) <= 1) then
      begin
        // 'a' / 'x' — short enough to leave alone.
        Targets[I].QualifyPrefix := Targets[I].Expression + '.';
        Targets[I].InlineVarName := '';
      end
      else
      begin
        Targets[I].InlineVarName := DeriveTempName(Last);
        if Targets[I].InlineVarName = '' then
          Targets[I].InlineVarName := Format('__with%d', [I + 1]);
        Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
      end;
    end
    else
    begin
      Targets[I].InlineVarName := Format('__with%d', [I + 1]);
      Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
    end;
  end;

  // Cross-target collision check: if two targets ended up with the same
  // InlineVarName, fall back to '__withN' for the later one(s).
  for I := 0 to High(Targets) do
    if Targets[I].InlineVarName <> '' then
      for var J := 0 to I - 1 do
        if SameText(Targets[J].InlineVarName, Targets[I].InlineVarName) then
        begin
          Targets[I].InlineVarName := Format('__with%d', [I + 1]);
          Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
          Break;
        end;

  // Determine indentation of the line containing the 'with' keyword.
  LineStartIdx := Index.PosToIndex(ReplaceStart.Line, 1);
  IndentStr := '';
  var P := LineStartIdx;
  while (P <= Length(ASource))
    and ((ASource[P] = ' ') or (ASource[P] = #9)) do
  begin
    IndentStr := IndentStr + ASource[P];
    Inc(P);
  end;

  // Collect body-identifier references and resolve each via LSP.
  BodyRefs := CollectBodyIdentifiers(ASource, AOccurrence.BodyInnerRange, Index);

  // Body-identifier collision check: if a body identifier already uses
  // the name we picked for an inline-var, fall back to '__withN'.
  for I := 0 to High(Targets) do
    if Targets[I].InlineVarName <> '' then
      for var BR := 0 to High(BodyRefs) do
        if SameText(BodyRefs[BR].Name, Targets[I].InlineVarName) then
        begin
          Targets[I].InlineVarName := Format('__with%d', [I + 1]);
          Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
          Break;
        end;

  // Build a set of (StartIdx -> qualifying prefix) to apply when
  // rewriting. We iterate refs in order and on each successful match
  // record which prefix to inject. Refs we leave alone get '' as prefix.
  var DebugRefs := TList<TWithDebugRefInfo>.Create;
  var Prefixes := TDictionary<Integer, string>.Create;
  try
    for Ref in BodyRefs do
    begin
      var Matched := False;
      var DbgRef: TWithDebugRefInfo;
      DbgRef := Default(TWithDebugRefInfo);
      DbgRef.Name := Ref.Name;
      DbgRef.Line := Ref.Pos.Line;
      DbgRef.Col := Ref.Pos.Col;
      DbgRef.MatchSource := dmNone;
      DbgRef.MatchedTargetIdx := -1;

      // Step 1: rightmost-wins lookup against the directly-declared
      // members of each target. This handles names declared in the
      // class/record body itself (e.g. fields, immediate properties /
      // methods). It does NOT cover inherited members - those need LSP.
      // Doing this first avoids relying on DelphiLSP, which appears to
      // check only the rightmost target's class for bare identifiers
      // inside a multi-target with-clause.
      for I := High(Targets) downto 0 do
      begin
        var Hit := False;
        for var M := 0 to High(Targets[I].Members) do
          if SameText(Targets[I].Members[M], Ref.Name) then
          begin
            Hit := True;
            Break;
          end;
        if Hit then
        begin
          Prefixes.AddOrSetValue(Ref.StartIdx, Targets[I].QualifyPrefix);
          Matched := True;
          DbgRef.MatchSource := dmMember;
          DbgRef.MatchedTargetIdx := I;
          DbgRef.AppliedPrefix := Targets[I].QualifyPrefix;
          Break;
        end;
      end;

      // Step 2: if no direct-member match, ask LSP. This is what catches
      // INHERITED members - e.g. TButton.Caption is declared in
      // TControl, so it's not in TButton's directly-parsed Members but
      // LSP resolves it to a location inside TControl. We accept any
      // resolution that points into a target's own class range; the
      // ancestor case is handled by the next step.
      if not Matched then
      begin
        try
          RefDef := AClient.GotoDefinition(AFileName, Ref.Pos.Line - 1, Ref.Pos.Col - 1);
        except
          RefDef := nil;
        end;
        DbgRef.LspHadResult := Length(RefDef) > 0;
        if Length(RefDef) > 0 then
        begin
          RefDefFile := TLspUri.FileUriToPath(RefDef[0].Uri);
          RefDefLine := RefDef[0].Range.Start.Line + 1;
          DbgRef.LspFile := RefDefFile;
          DbgRef.LspLine := RefDefLine;

          for I := High(Targets) downto 0 do
          begin
            if SameText(RefDefFile, Targets[I].TypeFile)
              and (RefDefLine >= Targets[I].ClassStartLine)
              and (RefDefLine <= Targets[I].ClassEndLine) then
            begin
              Prefixes.AddOrSetValue(Ref.StartIdx, Targets[I].QualifyPrefix);
              Matched := True;
              DbgRef.MatchSource := dmLsp;
              DbgRef.MatchedTargetIdx := I;
              DbgRef.AppliedPrefix := Targets[I].QualifyPrefix;
              Break;
            end;
          end;
        end;
      end;

      DebugRefs.Add(DbgRef);
    end;

    // Snapshot debug info for the dialog.
    SetLength(Result.Debug.Targets, Length(Targets));
    for I := 0 to High(Targets) do
    begin
      Result.Debug.Targets[I].Expression     := Targets[I].Expression;
      Result.Debug.Targets[I].TypeFile       := Targets[I].TypeFile;
      Result.Debug.Targets[I].ClassStartLine := Targets[I].ClassStartLine;
      Result.Debug.Targets[I].ClassEndLine   := Targets[I].ClassEndLine;
      Result.Debug.Targets[I].Members        := Targets[I].Members;
      Result.Debug.Targets[I].InlineVarName  := Targets[I].InlineVarName;
      Result.Debug.Targets[I].QualifyPrefix  := Targets[I].QualifyPrefix;
      Result.Debug.Targets[I].Resolved       := Targets[I].Resolved;
    end;
    Result.Debug.Refs := DebugRefs.ToArray;

    // Rebuild body with prefixes injected.
    var InnerStartIdx := Index.PosToIndex(
      AOccurrence.BodyInnerRange.StartPos.Line, AOccurrence.BodyInnerRange.StartPos.Col);
    var InnerEndIdx := Index.PosToIndex(
      AOccurrence.BodyInnerRange.EndPos.Line, AOccurrence.BodyInnerRange.EndPos.Col);
    InnerText := Copy(ASource, InnerStartIdx, InnerEndIdx - InnerStartIdx + 1);

    // Walk InnerText, applying prefixes by absolute string index.
    Builder := TStringBuilder.Create;
    try
      var Cursor := InnerStartIdx;
      while Cursor <= InnerEndIdx do
      begin
        var Prefix: string;
        if Prefixes.TryGetValue(Cursor, Prefix) then
          Builder.Append(Prefix);
        Builder.Append(ASource[Cursor]);
        Inc(Cursor);
      end;
      BodyOnly := Builder.ToString;
    finally
      Builder.Free;
    end;
  finally
    Prefixes.Free;
    DebugRefs.Free;
  end;

  // Compose the new text: optional inline-var(s), then body content
  // without the 'with X do begin' wrapper.
  HasInlineVar := False;
  for I := 0 to High(Targets) do
    if Targets[I].InlineVarName <> '' then HasInlineVar := True;

  Builder := TStringBuilder.Create;
  try
    case AOccurrence.BodyKind of
      wbkBeginEnd:
        begin
          // 'begin' + (var-decls) + body + 'end'. Var-decls go INSIDE
          // begin/end so their scope matches the original 'with'.
          Builder.Append('begin');
          if HasInlineVar then
          begin
            var BodyIndent := GetBodyStatementIndent(BodyOnly, IndentStr + '  ');
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.AppendLine.Append(BodyIndent)
                       .Append('var ').Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';');
          end;
          Builder.Append(BodyOnly);
          Builder.Append('end');
        end;
      wbkSingle:
        begin
          if HasInlineVar then
          begin
            // Wrap single statement in begin/end so var-decls are
            // syntactically valid.
            var BodyIndent := IndentStr + '  ';
            Builder.Append('begin');
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.AppendLine.Append(BodyIndent)
                       .Append('var ').Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';');
            Builder.AppendLine.Append(BodyIndent).Append(BodyOnly.Trim);
            Builder.AppendLine.Append(IndentStr).Append('end');
          end
          else
            Builder.Append(BodyOnly.TrimLeft);
        end;
    end;

    Result.NewText := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

end.
