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
// The inline var is named after the target's last identifier ('L' +
// segment), e.g. 'LAdd' for 'lvList.Items.Add'. When the expression has
// no usable identifier or collides with another target / body
// identifier, a numeric suffix is appended (e.g. 'LAdd2'). The very
// last-resort fallback is 'LWithN' (N = 1-based target index). Names
// never start with an underscore — this matches the Delphi style guide.
//
// Multi-target ("with A, B do") rewrites ARE produced automatically:
// per body identifier the rewriter picks the rightmost target whose
// class contains the definition. Complex targets (non-dotted) are
// hoisted to derived inline-vars in declaration order so the
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
    wriNameClash,
    /// <summary>Rewrite would introduce a Delphi 10.3+ inline-var
    ///  declaration but the user opted out of inline-var emission
    ///  (compatibility with older compilers).</summary>
    wriRequiresInlineVar,
    /// <summary>The with-statement sits inside an inactive {$IFDEF}
    ///  region (DelphiLSP-Diagnostic H2655/H2656, tag=Unnecessary).
    ///  The code is dead for the current build configuration; rewriting
    ///  it would be misleading. Left untouched.</summary>
    wriInactiveRegion,
    /// <summary>DelphiLSP has not delivered any publishDiagnostics for
    ///  this file within the wait window, so we cannot tell whether the
    ///  with-statement sits inside an inactive {$IFDEF}-region. To be
    ///  safe we leave the occurrence untouched - rewriting code that
    ///  might be dead would silently produce wrong output.</summary>
    wriLspNoDiagnostics
  );

  /// <summary>Caller-controlled toggles for the rewriter.</summary>
  TWithRewriteSettings = record
    /// <summary>When False, the rewriter still computes the inline-var
    ///  form for preview purposes but adds wriRequiresInlineVar to
    ///  occurrences whose rewrite would require a 10.3+ inline-var
    ///  declaration. Auto-apply will skip those occurrences; cases that
    ///  can be qualified directly (e.g. 'with FParser do' or
    ///  'with p^ do' after the field-qualification fix) remain
    ///  auto-rewritable in either mode.</summary>
    UseInlineVars: Boolean;
    class function Defaults: TWithRewriteSettings; static;
  end;

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

  /// <summary>One entry in a target's inheritance chain: file + 1-based
  ///  line range of an ancestor class body. The rewriter uses these
  ///  ranges to match body identifiers that LSP resolves to a parent
  ///  class (inherited members).</summary>
  TAncestorClassRange = record
    TypeFile: string;
    ClassStartLine: Integer;       // 1-based
    ClassEndLine: Integer;         // 1-based, inclusive
  end;

  TWithDebugTargetInfo = record
    Expression: string;
    TypeFile: string;
    ClassStartLine: Integer;
    ClassEndLine: Integer;
    Members: TArray<string>;
    InlineVarName: string;
    QualifyPrefix: string;
    Resolved: Boolean;
    /// <summary>Diagnostic note on why resolution failed. Empty when
    ///  Resolved=True. Examples: "LSP GotoDefinition returned 0 results",
    ///  "LSP exception: ELspError: Server not responding",
    ///  "decl line has no type name".</summary>
    ResolveNote: string;
    /// <summary>Resolved ancestor class ranges (file + start..end).</summary>
    Ancestors: TArray<TAncestorClassRange>;
    /// <summary>Unit-qualifier hints from a dotted parent class
    ///  declaration (e.g. ['Grids'] for 'class(Grids.TStringGrid)').</summary>
    ParentUnitHints: TArray<string>;
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

  /// <summary>One variable declaration to be added at the top of the
  ///  enclosing method when running in classic (non-inline-var) mode.</summary>
  TClassicVarDecl = record
    Name: string;
    TypeName: string;
  end;

  /// <summary>Classic-mode (non-inline-var) rewrite artifacts. Computed
  ///  in parallel to the inline NewText so the dialog can switch
  ///  between modes without re-running the rewriter. The orchestrator
  ///  aggregates these across all selected items per file/method when
  ///  applying.</summary>
  TClassicRewrite = record
    /// <summary>True iff a classic-mode rewrite is fully derivable for
    ///  this occurrence (enclosing method found, all target types have
    ///  a resolvable unit name).</summary>
    Supported: Boolean;
    /// <summary>Replacement text for ReplaceRange. Same shape as
    ///  inline-mode NewText but with plain assignments
    ///  ("LWith1 := Expr;") instead of inline-var declarations
    ///  ("var LWith1 := Expr;").</summary>
    BodyText: string;
    /// <summary>Stable key identifying the enclosing method
    ///  ("FileName:HeaderLine"). Used by the orchestrator to merge
    ///  var-section insertions when multiple with-statements live in
    ///  the same method.</summary>
    MethodKey: string;
    HasVarSection: Boolean;
    /// <summary>1-based line of the last existing var-decl in the
    ///  method's var section (0 if none).</summary>
    VarSectionLastLine: Integer;
    MethodBodyBeginLine: Integer;
    MethodBodyBeginCol: Integer;
    /// <summary>Recommended indent for var lines.</summary>
    LocalIndent: string;
    /// <summary>Variable declarations to emit at method top.</summary>
    VarDecls: TArray<TClassicVarDecl>;
    /// <summary>Unit names (case-preserved) to add to the
    ///  implementation uses clause. Already filtered against existing
    ///  uses clauses; empty if none needed.</summary>
    AddUnits: TArray<string>;
    /// <summary>Implementation-keyword location (1-based). Used when
    ///  no implementation-uses clause exists and a new one has to be
    ///  inserted.</summary>
    ImplKeywordLine: Integer;
    ImplKeywordCol: Integer;
    IntfUsesFound: Boolean;
    IntfUsesLastLine, IntfUsesLastCol: Integer;
    ImplUsesFound: Boolean;
    ImplUsesLastLine, ImplUsesLastCol: Integer;
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
    /// <summary>Classic-mode (non-inline-var) rewrite artifacts.</summary>
    Classic: TClassicRewrite;
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
      const ASource: string; const AOccurrence: TWithOccurrence;
      const ASettings: TWithRewriteSettings): TWithRewriteResult; static;
  end;

  /// <summary>1-based line index over a string source.</summary>
  TLineIndex = record
  private
    FStarts: TArray<Integer>;
    FSrc: string;
  public
    procedure Init(const ASource: string);
    function LineCount: Integer;
    function LineText(ALine1: Integer): string;
    function PosToIndex(ALine1, ACol1: Integer): Integer;
    function Slice(AStart, AEnd: TWithSourcePos): string;
  end;

  TUsesClauseLocation = record
    Found: Boolean;
    FirstUnitLine: Integer;
    FirstUnitCol: Integer;
    LastBeforeSemiLine: Integer;
    LastBeforeSemiCol: Integer;
    Units: TArray<string>;
  end;

  TUsesScanResult = record
    InterfaceUses: TUsesClauseLocation;
    ImplementationUses: TUsesClauseLocation;
    ImplementationKeywordLine: Integer;
    ImplementationKeywordCol: Integer;
  end;

/// <summary>Scans ASource and reports the interface/implementation uses
///  clause locations. Comment/string-aware.</summary>
function ScanUsesClauses(const ASource: string;
  const AIndex: TLineIndex): TUsesScanResult;

/// <summary>True iff AUnit appears in AClause.Units (case-insensitive).</summary>
function UsesContains(const AClause: TUsesClauseLocation;
  const AUnit: string): Boolean;

implementation

uses
  System.IOUtils, System.Character, System.StrUtils;

{ Helpers: 1-based <-> string-index, line slicing }

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

/// <summary>For wbkCompound bodies (try/case/asm), the first line of
///  ABody is the opener keyword without leading whitespace, but lines
///  2..N still carry the original indentation that was relative to the
///  opener's column in the source. After replacement the opener lives
///  at the with-keyword's indent (typically less), so we left-shift
///  the tail by (AOriginalCol - ANewCol) leading spaces. Empty lines
///  are kept empty. Lines with less leading whitespace than the shift
///  amount are kept (we only strip what's actually there - never
///  consume non-whitespace).</summary>
function ShiftCompoundBody(const ABody: string;
  AOriginalCol, ANewCol: Integer): string;
var
  Shift: Integer;
  Builder: TStringBuilder;
  LineStart, I, N, Stripped: Integer;
begin
  Shift := AOriginalCol - ANewCol;
  if Shift <= 0 then Exit(ABody);
  N := Length(ABody);
  Builder := TStringBuilder.Create;
  try
    LineStart := 1;
    I := 1;
    // First line: emit as-is up to the first newline.
    while (I <= N) and not CharInSet(ABody[I], [#10, #13]) do Inc(I);
    Builder.Append(Copy(ABody, LineStart, I - LineStart));
    // Subsequent lines: strip up to Shift leading spaces, then emit.
    while I <= N do
    begin
      // Copy the line-break verbatim.
      if (ABody[I] = #13) and (I < N) and (ABody[I + 1] = #10) then
      begin
        Builder.Append(#13#10);
        Inc(I, 2);
      end
      else
      begin
        Builder.Append(ABody[I]);
        Inc(I);
      end;
      // Walk forward while either consuming up to `Shift` spaces, or
      // until we hit a non-space char or end-of-line.
      Stripped := 0;
      while (I <= N) and (Stripped < Shift) and (ABody[I] = ' ') do
      begin
        Inc(I);
        Inc(Stripped);
      end;
      // Now emit the rest of the line (including any spaces beyond
      // Shift, up to next newline).
      LineStart := I;
      while (I <= N) and not CharInSet(ABody[I], [#10, #13]) do Inc(I);
      Builder.Append(Copy(ABody, LineStart, I - LineStart));
    end;
    Result := Builder.ToString;
  finally
    Builder.Free;
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
  // convention TFoo.FBar -> 'Bar') or leading T-prefix when followed by
  // upper-case (type convention TWinControl -> 'WinControl'). The
  // upper-case condition keeps real identifiers like 'TestFoo' / 'Free'
  // intact.
  if (Length(Base) >= 2) and Base[2].IsUpper
    and ((Base[1] = 'F') or (Base[1] = 'T')) then
    Base := Copy(Base, 2, MaxInt);
  if (Base = '') or not IsIdentStart(Base[1]) then Exit('');
  Result := 'L' + Base;
end;

/// <summary>Picks a Delphi-style temp-var name when DeriveTempName failed
///  (e.g. complex expression with calls, brackets, or empty last segment).
///  Walks AExpr right-to-left looking for the last identifier-looking
///  token and returns 'L' + that token. If no usable token is found,
///  returns 'LWith' + AIndex (1-based) so we never emit an underscore-
///  prefixed name — leading underscore violates the Delphi style guide.</summary>
function FallbackTempName(const AExpr: string; AIndex: Integer): string;
var
  S: string;
  I, J: Integer;
begin
  S := AExpr.TrimRight;
  // Strip a trailing caret (pointer-deref) - irrelevant for the name.
  while (S <> '') and (S[Length(S)] = '^') do
    S := Copy(S, 1, Length(S) - 1).TrimRight;
  // Walk back to find the end of the last ident-like run.
  I := Length(S);
  while (I >= 1) and not IsIdentCont(S[I]) do Dec(I);
  if I >= 1 then
  begin
    J := I;
    while (J >= 1) and IsIdentCont(S[J]) do Dec(J);
    Inc(J);
    if (J <= I) and IsIdentStart(S[J]) then
    begin
      var Token := Copy(S, J, I - J + 1);
      var Tmp := DeriveTempName(Token);
      if Tmp <> '' then Exit(Tmp);
    end;
  end;
  Result := Format('LWith%d', [AIndex]);
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

/// <summary>True iff the expression is a dotted-identifier path optionally
///  ending in a single '^' (pointer dereference). Safe to use directly
///  as a qualifying prefix (no side effects, no calls).</summary>
function IsSimpleDottedOrCaretTarget(const AExpr: string): Boolean;
var
  Trimmed: string;
begin
  Trimmed := AExpr.Trim;
  if Trimmed = '' then Exit(False);
  if Trimmed[Length(Trimmed)] = '^' then
    Trimmed := Copy(Trimmed, 1, Length(Trimmed) - 1).TrimRight;
  Result := IsSimpleDottedTarget(Trimmed);
end;

/// <summary>True iff the expression is PROVABLY side-effect-free as a
///  qualifying prefix: a single identifier, or 'Self.Field', optionally
///  with a single trailing '^'. Anything with two or more dots is
///  rejected — intermediate segments could be property accessors with
///  side effects, and the final segment could be a parameterless
///  function (e. g. 'lvList.Items.Add'). When this returns False, the
///  caller must use a temp variable to preserve 'with' semantics.</summary>
function IsSideEffectFreeTarget(const AExpr: string): Boolean;
var
  Trimmed: string;
  Segs: TArray<string>;
begin
  Trimmed := AExpr.Trim;
  if Trimmed = '' then Exit(False);
  if Trimmed[Length(Trimmed)] = '^' then
    Trimmed := Copy(Trimmed, 1, Length(Trimmed) - 1).TrimRight;
  if not IsSimpleDottedTarget(Trimmed) then Exit(False);
  Segs := SplitDottedSegments(Trimmed);
  case Length(Segs) of
    0: Result := False;
    1: Result := True;                            // 'Foo'
    2: Result := SameText(Segs[0], 'Self');       // 'Self.Field' only
  else
    Result := False;                              // 2+ dots — not safe
  end;
end;

/// <summary>If AExpr ends in '^' (after trim), returns AExpr without
///  the trailing caret; otherwise returns AExpr unchanged.</summary>
function StripTrailingCaret(const AExpr: string): string;
begin
  Result := AExpr.TrimRight;
  if (Result <> '') and (Result[Length(Result)] = '^') then
    Result := Copy(Result, 1, Length(Result) - 1).TrimRight;
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
  // Tolerate a leading '^' for inline pointer-type declarations like
  //   "p: ^TFoo;"
  // The target expression typically dereferences the pointer ("p^"),
  // so the type we care about is the pointee. Skip the caret (and any
  // whitespace before the pointee identifier).
  if (I <= N) and (ALine[I] = '^') then
  begin
    Inc(I);
    while (I <= N) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  end;
  if (I > N) or not IsIdentStart(ALine[I]) then Exit;

  ATypeNameCol1 := I;
  var Start := I;
  while (I <= N) and IsIdentCont(ALine[I]) do Inc(I);
  Result := Copy(ALine, Start, I - Start);
end;

/// <summary>True iff (after trimming leading WS) the declaration line
///  starts with a method keyword (function/procedure/constructor/
///  destructor/operator), possibly preceded by 'class'. Used to decide
///  whether a 'with X do' target is a side-effecting call (keep temp
///  var) or a plain value (inline qualifier).</summary>
function DeclLineDeclaresMethod(const ALine: string): Boolean;
var
  I: Integer;
  Word: string;
begin
  Result := False;
  I := 1;
  while (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  // Optional 'class' modifier
  if (I <= Length(ALine)) and IsIdentStart(ALine[I]) then
  begin
    var Q := I;
    while (Q <= Length(ALine)) and IsIdentCont(ALine[Q]) do Inc(Q);
    Word := UpperCase(Copy(ALine, I, Q - I));
    if Word = 'CLASS' then
    begin
      I := Q;
      while (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
      if (I <= Length(ALine)) and IsIdentStart(ALine[I]) then
      begin
        Q := I;
        while (Q <= Length(ALine)) and IsIdentCont(ALine[Q]) do Inc(Q);
        Word := UpperCase(Copy(ALine, I, Q - I));
      end
      else
        Word := '';
    end;
    Result := (Word = 'FUNCTION') or (Word = 'PROCEDURE')
           or (Word = 'CONSTRUCTOR') or (Word = 'DESTRUCTOR')
           or (Word = 'OPERATOR');
  end;
end;

/// <summary>For a constructor/destructor declaration like
///    "constructor TFoo.Create(const X: Integer);"
///    "class constructor TFoo.Init;"
///    "destructor TFoo.Destroy; override;"
///  extracts the class qualifier (TFoo) and its 1-based column. The
///  class qualifier is the implicit return-type of the (de)constructor
///  for our purposes (the instance the constructor builds). Returns ''
///  if the line is not a (de)constructor with a dotted name.</summary>
function ExtractConstructorClassName(const ALine: string;
  out AClassCol1: Integer): string;
var
  I, N: Integer;
  ClsStart: Integer;
  Word: string;

  procedure SkipWS;
  begin
    while (I <= N) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  end;

  function ReadIdent: string;
  var Start: Integer;
  begin
    Result := '';
    SkipWS;
    if (I > N) or not IsIdentStart(ALine[I]) then Exit;
    Start := I;
    while (I <= N) and IsIdentCont(ALine[I]) do Inc(I);
    Result := Copy(ALine, Start, I - Start);
  end;

begin
  Result := '';
  AClassCol1 := 0;
  N := Length(ALine);
  I := 1;
  SkipWS;
  if I > N then Exit;

  // Optional leading 'class' modifier.
  Word := ReadIdent;
  if SameText(Word, 'class') then
    Word := ReadIdent;

  if not (SameText(Word, 'constructor') or SameText(Word, 'destructor')) then
    Exit;

  // Class name follows.
  SkipWS;
  ClsStart := I;
  Word := ReadIdent;
  if Word = '' then Exit;

  // Must be followed by '.' (qualified method name); otherwise the
  // line is a bare 'destructor;' declaration inside an interface body,
  // not useful here.
  SkipWS;
  if (I > N) or (ALine[I] <> '.') then Exit;

  AClassCol1 := ClsStart;
  Result := Word;
end;

/// <summary>Tries to parse a type-alias line and returns the RHS
///  identifier we should hop to. Handles both:
///     "PFoo = ^TFoo;"   (pointer alias)
///     "TLifo = _L_List;" (plain alias)
///  In either case we want to GotoDef the RHS identifier and follow the
///  chain. Returns '' when the line doesn't look like a simple alias
///  (e. g. a 'record', 'class', 'set of', generic, etc.) so the caller
///  stops hopping and uses the current line as the type declaration.</summary>
function ExtractPointeeNameFromAliasLine(const ALine: string;
  out APointeeCol1: Integer): string;
var
  I, N, EqPos: Integer;
begin
  Result := '';
  APointeeCol1 := 0;
  N := Length(ALine);

  // Find first '=' not part of '<=' / '>=' / ':=' / '=>'.
  EqPos := 0;
  I := 1;
  while I <= N do
  begin
    if ALine[I] = '=' then
    begin
      EqPos := I;
      Break;
    end;
    Inc(I);
  end;
  if EqPos = 0 then Exit;

  // After '=' skip WS and an optional '^' (pointer alias form).
  I := EqPos + 1;
  while (I <= N) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  if (I <= N) and (ALine[I] = '^') then
  begin
    Inc(I);
    while (I <= N) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  end;
  if (I > N) or not IsIdentStart(ALine[I]) then Exit;

  // Read the candidate identifier.
  var Start := I;
  while (I <= N) and IsIdentCont(ALine[I]) do Inc(I);
  var Candidate := Copy(ALine, Start, I - Start);

  // Reject keywords that introduce a real type body. The caller should
  // stop hopping here and use the current line as the type definition.
  var Upper := UpperCase(Candidate);
  if (Upper = 'CLASS') or (Upper = 'INTERFACE') or (Upper = 'RECORD')
    or (Upper = 'OBJECT') or (Upper = 'DISPINTERFACE') or (Upper = 'PACKED')
    or (Upper = 'SET') or (Upper = 'ARRAY') or (Upper = 'FILE')
    or (Upper = 'FUNCTION') or (Upper = 'PROCEDURE') or (Upper = 'REFERENCE')
    or (Upper = 'TYPE') then
    Exit('');

  APointeeCol1 := Start;
  Result := Candidate;
end;

/// <summary>Given a class declaration line like
///    "TFoo = class(Vcl.Controls.TWinControl)"
///    "TBar = class(TBase, ISomeIntf)"
///  extracts the LAST identifier of the parent class reference (so
///  "Vcl.Controls.TWinControl" yields "TWinControl") plus its 1-based
///  column. Returns '' when the line is not a class/interface decl with
///  a parent, or the parent is just 'TObject' / absent (open).</summary>
function ExtractParentClassFromDeclLine(const ALine: string;
  out AParentCol1: Integer): string;
var
  I, N, OpenParen, CloseParen: Integer;
begin
  Result := '';
  AParentCol1 := 0;
  N := Length(ALine);
  // Locate '(' that follows 'class' / 'interface' / 'object' /
  // 'dispinterface' on the same line. We only need to look for the
  // first '(' since these keywords introduce the parent list.
  OpenParen := 0;
  I := 1;
  while I <= N do
  begin
    if ALine[I] = '(' then begin OpenParen := I; Break end;
    Inc(I);
  end;
  if OpenParen = 0 then Exit;
  // Find matching ')' (no nesting in a normal parent list)
  CloseParen := 0;
  I := OpenParen + 1;
  while I <= N do
  begin
    if ALine[I] = ')' then begin CloseParen := I; Break end;
    Inc(I);
  end;
  if CloseParen = 0 then CloseParen := N + 1;

  // Take FIRST entry from the parent list - that's the actual ancestor;
  // anything after a comma is an interface.
  var Inner := Copy(ALine, OpenParen + 1, CloseParen - OpenParen - 1);
  var CommaPos := Pos(',', Inner);
  if CommaPos > 0 then
    Inner := Copy(Inner, 1, CommaPos - 1);
  Inner := Inner.Trim;

  // Inner now looks like 'Vcl.Controls.TWinControl' or 'TBase'.
  // Walk from the end to find the last identifier and its position
  // relative to the original line.
  var LastIdentStart: Integer;
  var P := Length(Inner);
  while (P >= 1) and not IsIdentCont(Inner[P]) do Dec(P);
  if P < 1 then Exit;
  var EndP := P;
  while (P >= 1) and IsIdentCont(Inner[P]) do Dec(P);
  Inc(P);
  if (P > EndP) or not IsIdentStart(Inner[P]) then Exit;
  Result := Copy(Inner, P, EndP - P + 1);
  // 'TObject' as a literal ancestor is the universal root - walking
  // higher is not useful (and would just loop on System).
  if SameText(Result, 'TObject') then Exit('');
  // Compute 1-based column in the original line.
  LastIdentStart := OpenParen + 1; // start of Inner in the original
  // Skip the same leading whitespace we trimmed off in Inner.
  while (LastIdentStart <= N)
    and CharInSet(ALine[LastIdentStart], [' ', #9]) do
    Inc(LastIdentStart);
  // P is the 1-based position WITHIN the trimmed Inner. We need its
  // offset within the ORIGINAL Inner (untrimmed). Easier: re-scan the
  // original ALine from LastIdentStart, applying the same comma cut.
  var ScanStart := LastIdentStart;
  var ScanEnd := CloseParen - 1;
  if CommaPos > 0 then
    ScanEnd := LastIdentStart + CommaPos - 2;
  // Walk back from ScanEnd to find the last ident token in the
  // original line.
  var QQ := ScanEnd;
  while (QQ >= ScanStart) and not IsIdentCont(ALine[QQ]) do Dec(QQ);
  if QQ < ScanStart then Exit;
  var EndQ := QQ;
  while (QQ >= ScanStart) and IsIdentCont(ALine[QQ]) do Dec(QQ);
  Inc(QQ);
  AParentCol1 := QQ;
  Result := Copy(ALine, QQ, EndQ - QQ + 1);
  if SameText(Result, 'TObject') then
  begin
    Result := '';
    AParentCol1 := 0;
  end;
end;

/// <summary>Scans the class/record body lines AStartLine..AEndLine of
///  AFile, looking for a declaration whose member-name matches
///  AMemberName. Returns the LAST identifier of its declared type
///  (e.g. 'TCanvas' for "property Canvas: Vcl.Graphics.TCanvas read FCanvas;")
///  plus the 1-based line/col of that identifier. Used by the property-
///  type walker to navigate dotted target expressions like
///  'LListBox.Canvas' when DelphiLSP can't resolve the property
///  directly.</summary>
function FindMemberTypeInClassBody(const AFile: string;
  AStartLine, AEndLine: Integer; const AMemberName: string;
  out ATypeLine: Integer; out ATypeCol1: Integer): string;
var
  Source: string;
  Index: TLineIndex;
  L: Integer;
begin
  Result := '';
  ATypeLine := 0;
  ATypeCol1 := 0;
  if not TFile.Exists(AFile) then Exit;
  try Source := TFile.ReadAllText(AFile); except Exit end;
  Index.Init(Source);
  for L := AStartLine to AEndLine do
  begin
    if (L < 1) or (L > Index.LineCount) then Break;
    var LineText := Index.LineText(L);
    var N := Length(LineText);
    var P := 1;
    while (P <= N) and CharInSet(LineText[P], [' ', #9]) do Inc(P);
    if P > N then Continue;

    // Skip optional class-/strict-/method-keyword modifiers, then
    // optionally 'property' / 'function' / 'procedure' / 'constructor'
    // / 'destructor'. We loop because Delphi allows compound modifiers
    // like 'class function' or 'class property'.
    while P <= N do
    begin
      if not IsIdentStart(LineText[P]) then Break;
      var QQ := P;
      while (QQ <= N) and IsIdentCont(LineText[QQ]) do Inc(QQ);
      var Word := UpperCase(Copy(LineText, P, QQ - P));
      if (Word = 'CLASS') or (Word = 'STRICT') or (Word = 'PROPERTY')
        or (Word = 'FUNCTION') or (Word = 'PROCEDURE')
        or (Word = 'CONSTRUCTOR') or (Word = 'DESTRUCTOR') then
      begin
        P := QQ;
        while (P <= N) and CharInSet(LineText[P], [' ', #9]) do Inc(P);
        Continue;
      end;
      Break;
    end;

    // Read the member name.
    if (P > N) or not IsIdentStart(LineText[P]) then Continue;
    var NameStart := P;
    while (P <= N) and IsIdentCont(LineText[P]) do Inc(P);
    var Name := Copy(LineText, NameStart, P - NameStart);
    if not SameText(Name, AMemberName) then Continue;

    // Optional comma-list (other names sharing the same type). Walk
    // past until we hit ':'.
    while (P <= N) and (LineText[P] <> ':') and (LineText[P] <> '=') do Inc(P);
    if (P > N) or (LineText[P] <> ':') then Continue;
    Inc(P);  // skip ':'
    while (P <= N) and CharInSet(LineText[P], [' ', #9]) do Inc(P);

    // Skip optional '^' (pointer-to-T)
    if (P <= N) and (LineText[P] = '^') then
    begin
      Inc(P);
      while (P <= N) and CharInSet(LineText[P], [' ', #9]) do Inc(P);
    end;

    if (P > N) or not IsIdentStart(LineText[P]) then Continue;
    var TypeStart := P;
    while (P <= N) and IsIdentCont(LineText[P]) do Inc(P);

    // Follow dotted type names: take the LAST identifier as the type
    // proper (so 'Vcl.Graphics.TCanvas' yields 'TCanvas').
    var LastSegStart := TypeStart;
    while (P <= N) and (LineText[P] = '.')
      and (P + 1 <= N) and IsIdentStart(LineText[P + 1]) do
    begin
      Inc(P);
      LastSegStart := P;
      while (P <= N) and IsIdentCont(LineText[P]) do Inc(P);
    end;
    Result := Copy(LineText, LastSegStart, P - LastSegStart);
    ATypeLine := L;
    ATypeCol1 := LastSegStart;
    Exit;
  end;
end;

{ Unit name + uses-clause helpers }

/// <summary>Reads the first 'unit Foo;' / 'unit Foo.Bar;' header from
///  ASource and returns the unit identifier ('Foo' / 'Foo.Bar').
///  Comment/whitespace tolerant. Returns '' for non-unit files
///  (program, library, package) or on parse failure.</summary>
function GetUnitNameFromSource(const ASource: string): string;
var
  I, N: Integer;

  function SkipWS: Boolean;
  begin
    while I <= N do
    begin
      if CharInSet(ASource[I], [' ', #9, #10, #13]) then
      begin
        Inc(I);
        Continue;
      end;
      if (ASource[I] = '/') and (I < N) and (ASource[I + 1] = '/') then
      begin
        while (I <= N) and (ASource[I] <> #10) and (ASource[I] <> #13) do Inc(I);
        Continue;
      end;
      if ASource[I] = '{' then
      begin
        while (I <= N) and (ASource[I] <> '}') do Inc(I);
        if I <= N then Inc(I);
        Continue;
      end;
      if (ASource[I] = '(') and (I < N) and (ASource[I + 1] = '*') then
      begin
        Inc(I, 2);
        while I <= N do
        begin
          if (ASource[I] = '*') and (I < N) and (ASource[I + 1] = ')') then
          begin
            Inc(I, 2);
            Break;
          end;
          Inc(I);
        end;
        Continue;
      end;
      Break;
    end;
    Result := I <= N;
  end;

  function ReadIdent: string;
  var
    Start: Integer;
  begin
    Result := '';
    if (I > N) or not IsIdentStart(ASource[I]) then Exit;
    Start := I;
    while (I <= N) and IsIdentCont(ASource[I]) do Inc(I);
    Result := Copy(ASource, Start, I - Start);
  end;

var
  Tok, Full: string;
begin
  Result := '';
  I := 1;
  N := Length(ASource);
  if not SkipWS then Exit;
  Tok := ReadIdent;
  if not SameText(Tok, 'unit') then Exit;
  if not SkipWS then Exit;
  Full := ReadIdent;
  if Full = '' then Exit;
  // Allow dotted unit names: Foo.Bar.Baz
  while True do
  begin
    var Save := I;
    if not SkipWS then Break;
    if (I > N) or (ASource[I] <> '.') then
    begin
      I := Save;
      Break;
    end;
    Inc(I);
    if not SkipWS then Exit('');
    var Next := ReadIdent;
    if Next = '' then Exit('');
    Full := Full + '.' + Next;
  end;
  Result := Full;
end;

function ScanUsesClauses(const ASource: string;
  const AIndex: TLineIndex): TUsesScanResult;
type
  TSect = (secOuter, secInterface, secImpl);
var
  I, N, Line, Col: Integer;
  Sect: TSect;
  Tok: string;
  TokLine, TokCol: Integer;

  procedure AdvancePastTriviaTo(out OutLine, OutCol: Integer);
  begin
    while I <= N do
    begin
      var Ch := ASource[I];
      if Ch = #13 then
      begin
        Inc(I);
        if (I <= N) and (ASource[I] = #10) then Inc(I);
        Inc(Line); Col := 1;
        Continue;
      end;
      if Ch = #10 then begin Inc(I); Inc(Line); Col := 1; Continue; end;
      if CharInSet(Ch, [' ', #9]) then begin Inc(I); Inc(Col); Continue; end;
      if (Ch = '/') and (I < N) and (ASource[I + 1] = '/') then
      begin
        while (I <= N) and (ASource[I] <> #10) and (ASource[I] <> #13) do
        begin Inc(I); Inc(Col); end;
        Continue;
      end;
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
      if (Ch = '(') and (I < N) and (ASource[I + 1] = '*') then
      begin
        Inc(I, 2); Inc(Col, 2);
        while I <= N do
        begin
          if (ASource[I] = '*') and (I < N) and (ASource[I + 1] = ')') then
          begin Inc(I, 2); Inc(Col, 2); Break; end;
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
      Break;
    end;
    OutLine := Line; OutCol := Col;
  end;

  function ReadToken(out ATokLine, ATokCol: Integer): string;
  begin
    AdvancePastTriviaTo(ATokLine, ATokCol);
    Result := '';
    if I > N then Exit;
    if IsIdentStart(ASource[I]) then
    begin
      var Start := I;
      while (I <= N) and IsIdentCont(ASource[I]) do
      begin
        Inc(I); Inc(Col);
      end;
      Result := Copy(ASource, Start, I - Start);
    end
    else
    begin
      Result := ASource[I];
      Inc(I); Inc(Col);
    end;
  end;

  procedure ParseUsesClauseInto(var ALoc: TUsesClauseLocation);
  var
    Units: TList<string>;
    UnitName: string;
    UL, UC: Integer;
    LastIdentLine, LastIdentCol: Integer;
    SawIdent: Boolean;
  begin
    // Already consumed the 'uses' keyword.
    Units := TList<string>.Create;
    SawIdent := False;
    LastIdentLine := 0;
    LastIdentCol := 0;
    try
      while True do
      begin
        Tok := ReadToken(UL, UC);
        if Tok = '' then Break;
        if (Length(Tok) = 1) and (Tok[1] = ';') then Break;
        if (Length(Tok) = 1) and (Tok[1] = ',') then Continue;
        if (Length(Tok) = 1) and (Tok[1] = '.') then Continue;
        if not IsIdentStart(Tok[1]) then Break;

        // Compose dotted unit name
        UnitName := Tok;
        if not ALoc.Found then
        begin
          ALoc.Found := True;
          ALoc.FirstUnitLine := UL;
          ALoc.FirstUnitCol := UC;
        end;
        // Position of last char of this token (line/col approx)
        LastIdentLine := UL;
        LastIdentCol := UC + Length(Tok) - 1;
        SawIdent := True;

        // Lookahead for '.IDENT' continuations
        while True do
        begin
          var SaveI := I; var SaveLine := Line; var SaveCol := Col;
          var DL, DC: Integer;
          var Peek := ReadToken(DL, DC);
          if (Length(Peek) = 1) and (Peek[1] = '.') then
          begin
            var Next := ReadToken(DL, DC);
            if (Next <> '') and IsIdentStart(Next[1]) then
            begin
              UnitName := UnitName + '.' + Next;
              LastIdentLine := DL;
              LastIdentCol := DC + Length(Next) - 1;
              Continue;
            end
            else
            begin
              // Roll back the unexpected token
              I := SaveI; Line := SaveLine; Col := SaveCol;
              Break;
            end;
          end
          else if SameText(Peek, 'in') then
          begin
            // "Foo in 'Foo.pas'" — skip the string literal
            // (the next ReadToken returns the apostrophe; we then have
            // to read the contents — simpler: skip everything until
            // the next ',' or ';' at top level).
            while I <= N do
            begin
              if ASource[I] = ';' then Break;
              if ASource[I] = ',' then Break;
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
            Break;
          end
          else
          begin
            I := SaveI; Line := SaveLine; Col := SaveCol;
            Break;
          end;
        end;

        Units.Add(UnitName);
      end;
      ALoc.Units := Units.ToArray;
      if SawIdent then
      begin
        ALoc.LastBeforeSemiLine := LastIdentLine;
        ALoc.LastBeforeSemiCol := LastIdentCol;
      end;
    finally
      Units.Free;
    end;
  end;

begin
  Result := Default(TUsesScanResult);
  if AIndex.LineCount = 0 then Exit;
  I := 1;
  N := Length(ASource);
  Line := 1;
  Col := 1;
  Sect := secOuter;

  while I <= N do
  begin
    Tok := ReadToken(TokLine, TokCol);
    if Tok = '' then Break;
    if not IsIdentStart(Tok[1]) then Continue;

    if SameText(Tok, 'interface') and (Sect = secOuter) then
    begin
      Sect := secInterface;
    end
    else if SameText(Tok, 'implementation') and (Sect <> secImpl) then
    begin
      Sect := secImpl;
      Result.ImplementationKeywordLine := TokLine;
      Result.ImplementationKeywordCol := TokCol;
    end
    else if SameText(Tok, 'uses') then
    begin
      case Sect of
        secInterface: ParseUsesClauseInto(Result.InterfaceUses);
        secImpl:      ParseUsesClauseInto(Result.ImplementationUses);
      else
        // 'uses' outside known sections — ignore.
      end;
    end
    else if SameText(Tok, 'initialization') or SameText(Tok, 'finalization')
         or SameText(Tok, 'end') then
    begin
      // Past the area where uses can still appear.
      if Sect = secImpl then Break;
    end;
  end;
end;

function UsesContains(const AClause: TUsesClauseLocation; const AUnit: string): Boolean;
var
  I: Integer;
begin
  for I := 0 to High(AClause.Units) do
    if SameText(AClause.Units[I], AUnit) then Exit(True);
  Result := False;
end;

{ Enclosing-method detection for classic-mode var insertion }

type
  TMethodInfo = record
    HeaderLine: Integer;      // 1-based line of 'procedure'/'function'/...
    BodyBeginLine: Integer;   // 1-based line of the method's own 'begin'
    BodyBeginCol: Integer;    // 1-based col of that 'begin'
    BodyEndLine: Integer;     // 1-based line of the matching 'end'
    HasVarSection: Boolean;   // 'var' keyword found between header and begin
    VarSectionLastLine: Integer;  // 1-based line of last var-decl, or 0
    LocalIndent: string;      // indent recommended for var entries
  end;

/// <summary>Walk ASource once, identifying all top-level
///  procedure/function/constructor/destructor/operator declarations
///  that have an implementation body in this file. Returns them in
///  the order they appear. Comment/string-aware. Nested procedures
///  are also captured.
///  The detection treats any of these keywords at the beginning of a
///  logical token as a candidate method start; it then looks for the
///  matching 'begin'..'end' block. Forward declarations (header
///  followed directly by ';' or a directive list and another ';',
///  without 'begin') are skipped.</summary>
function FindAllMethodsInSource(const ASource: string;
  const AIndex: TLineIndex): TArray<TMethodInfo>;
type
  TPState = (psOuter, psInHeader, psInVar, psInBody);
var
  Methods: TList<TMethodInfo>;
  Stack: TList<TMethodInfo>;
  StateStack: TList<TPState>;
  BodyDepthStack: TList<Integer>;  // begin/try/case/asm depth in body

  procedure PushNew(AHeaderLine: Integer);
  var
    M: TMethodInfo;
  begin
    M := Default(TMethodInfo);
    M.HeaderLine := AHeaderLine;
    Stack.Add(M);
    StateStack.Add(psInHeader);
    BodyDepthStack.Add(0);
  end;

  procedure FinalizeTop(AEndLine: Integer);
  var
    M: TMethodInfo;
  begin
    if Stack.Count = 0 then Exit;
    M := Stack[Stack.Count - 1];
    M.BodyEndLine := AEndLine;
    if M.BodyBeginLine > 0 then
      Methods.Add(M);
    Stack.Delete(Stack.Count - 1);
    StateStack.Delete(StateStack.Count - 1);
    BodyDepthStack.Delete(BodyDepthStack.Count - 1);
  end;

  function TopState: TPState;
  begin
    if StateStack.Count = 0 then Exit(psOuter);
    Result := StateStack[StateStack.Count - 1];
  end;

  procedure SetTopState(AState: TPState);
  begin
    if StateStack.Count = 0 then Exit;
    StateStack[StateStack.Count - 1] := AState;
  end;

  procedure SetTopMethod(const AMI: TMethodInfo);
  begin
    if Stack.Count = 0 then Exit;
    Stack[Stack.Count - 1] := AMI;
  end;

  function TopMethod: TMethodInfo;
  begin
    if Stack.Count = 0 then Exit(Default(TMethodInfo));
    Result := Stack[Stack.Count - 1];
  end;

  function TopDepth: Integer;
  begin
    if BodyDepthStack.Count = 0 then Exit(0);
    Result := BodyDepthStack[BodyDepthStack.Count - 1];
  end;

  procedure SetTopDepth(AValue: Integer);
  begin
    if BodyDepthStack.Count = 0 then Exit;
    BodyDepthStack[BodyDepthStack.Count - 1] := AValue;
  end;

var
  I, N, Line, Col: Integer;
  TokenStart, TokenStartLine, TokenStartCol: Integer;
  WordUC: string;
  MI: TMethodInfo;
  IsMethodKW: Boolean;
begin
  Methods := TList<TMethodInfo>.Create;
  Stack := TList<TMethodInfo>.Create;
  StateStack := TList<TPState>.Create;
  BodyDepthStack := TList<Integer>.Create;
  try
    I := 1;
    N := Length(ASource);
    Line := 1;
    Col := 1;
    while I <= N do
    begin
      var Ch := ASource[I];

      // Linebreaks
      if Ch = #13 then
      begin
        Inc(I);
        if (I <= N) and (ASource[I] = #10) then Inc(I);
        Inc(Line); Col := 1;
        Continue;
      end;
      if Ch = #10 then
      begin
        Inc(I); Inc(Line); Col := 1;
        Continue;
      end;

      // String literal
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
        Continue;
      end;

      // Char constant
      if Ch = '#' then
      begin
        Inc(I); Inc(Col);
        if (I <= N) and (ASource[I] = '$') then begin Inc(I); Inc(Col); end;
        while (I <= N) and CharInSet(ASource[I], ['0'..'9','a'..'f','A'..'F']) do
        begin
          Inc(I); Inc(Col);
        end;
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
        TokenStart := I;
        TokenStartLine := Line;
        TokenStartCol := Col;
        while (I <= N) and IsIdentCont(ASource[I]) do
        begin
          Inc(I); Inc(Col);
        end;
        WordUC := UpperCase(Copy(ASource, TokenStart, I - TokenStart));

        // Dispatch by current state.
        case TopState of
          psOuter, psInBody:
            begin
              IsMethodKW :=
                (WordUC = 'PROCEDURE') or (WordUC = 'FUNCTION') or
                (WordUC = 'CONSTRUCTOR') or (WordUC = 'DESTRUCTOR') or
                (WordUC = 'OPERATOR');
              if (TopState = psOuter) and IsMethodKW then
                PushNew(TokenStartLine)
              else if TopState = psInBody then
              begin
                if (WordUC = 'BEGIN') or (WordUC = 'TRY')
                  or (WordUC = 'CASE') or (WordUC = 'ASM') then
                  SetTopDepth(TopDepth + 1)
                else if WordUC = 'END' then
                begin
                  SetTopDepth(TopDepth - 1);
                  if TopDepth = 0 then
                    FinalizeTop(TokenStartLine);
                end
                else if IsMethodKW then
                  // Nested method declaration. Should appear in psInVar
                  // (between outer header and outer begin), not psInBody,
                  // but Delphi syntactically allows nothing useful here.
                  // Conservative: ignore.
                  ;
              end;
            end;

          psInHeader:
            begin
              if WordUC = 'VAR' then
              begin
                SetTopState(psInVar);
              end
              else if WordUC = 'BEGIN' then
              begin
                MI := TopMethod;
                MI.BodyBeginLine := TokenStartLine;
                MI.BodyBeginCol := TokenStartCol;
                SetTopMethod(MI);
                SetTopState(psInBody);
                SetTopDepth(1);
              end
              else if (WordUC = 'FORWARD') or (WordUC = 'EXTERNAL')
                   or (WordUC = 'OVERLOAD') or (WordUC = 'OVERRIDE')
                   or (WordUC = 'VIRTUAL') or (WordUC = 'DYNAMIC')
                   or (WordUC = 'ABSTRACT') or (WordUC = 'REINTRODUCE')
                   or (WordUC = 'CDECL') or (WordUC = 'STDCALL')
                   or (WordUC = 'REGISTER') or (WordUC = 'PASCAL')
                   or (WordUC = 'INLINE') or (WordUC = 'ASSEMBLER')
                   or (WordUC = 'STATIC') or (WordUC = 'FINAL')
                   or (WordUC = 'MESSAGE') or (WordUC = 'DISPID') then
                // Directive — stay in header.
                ;
            end;

          psInVar:
            begin
              if WordUC = 'BEGIN' then
              begin
                MI := TopMethod;
                MI.BodyBeginLine := TokenStartLine;
                MI.BodyBeginCol := TokenStartCol;
                SetTopMethod(MI);
                SetTopState(psInBody);
                SetTopDepth(1);
              end
              else if (WordUC = 'PROCEDURE') or (WordUC = 'FUNCTION')
                or (WordUC = 'CONSTRUCTOR') or (WordUC = 'DESTRUCTOR')
                or (WordUC = 'OPERATOR') then
              begin
                // Nested method inside this method's var section
                PushNew(TokenStartLine);
              end
              else if (WordUC = 'CONST') or (WordUC = 'TYPE')
                or (WordUC = 'LABEL') or (WordUC = 'RESOURCESTRING') then
                // Stays in pseudo-var area; we still treat it as 'in_var' so
                // 'begin' continues to land us in body.
                ;
              // Otherwise: it's a var declaration line. Capture this line
              // as the last seen var-decl. We approximate: ANY identifier
              // we see here updates HasVarSection + VarSectionLastLine.
              if not ((WordUC = 'BEGIN') or (WordUC = 'PROCEDURE') or
                      (WordUC = 'FUNCTION') or (WordUC = 'CONSTRUCTOR') or
                      (WordUC = 'DESTRUCTOR') or (WordUC = 'OPERATOR') or
                      (WordUC = 'TYPE') or (WordUC = 'CONST') or
                      (WordUC = 'LABEL') or (WordUC = 'RESOURCESTRING') or
                      (WordUC = 'VAR')) then
              begin
                MI := TopMethod;
                MI.HasVarSection := True;
                MI.VarSectionLastLine := TokenStartLine;
                // Track indent of first var-decl seen.
                if MI.LocalIndent = '' then
                  MI.LocalIndent := StringOfChar(' ', TokenStartCol - 1);
                SetTopMethod(MI);
              end;
            end;
        end;
        Continue;
      end;

      // Whitespace
      Inc(I); Inc(Col);
    end;
    Result := Methods.ToArray;
  finally
    BodyDepthStack.Free;
    StateStack.Free;
    Stack.Free;
    Methods.Free;
  end;
end;

/// <summary>Of all methods produced by FindAllMethodsInSource, returns
///  the INNERMOST one whose body range contains ATargetLine. Returns
///  False if none matches.</summary>
function FindEnclosingMethod(const AMethods: TArray<TMethodInfo>;
  ATargetLine: Integer; out AInfo: TMethodInfo): Boolean;
var
  I: Integer;
  Best: TMethodInfo;
  HaveBest: Boolean;
begin
  HaveBest := False;
  Best := Default(TMethodInfo);
  for I := 0 to High(AMethods) do
  begin
    if (AMethods[I].BodyBeginLine > 0)
      and (AMethods[I].BodyBeginLine <= ATargetLine)
      and (AMethods[I].BodyEndLine >= ATargetLine) then
    begin
      // Innermost = method with the largest BodyBeginLine satisfying it
      if (not HaveBest) or (AMethods[I].BodyBeginLine > Best.BodyBeginLine) then
      begin
        Best := AMethods[I];
        HaveBest := True;
      end;
    end;
  end;
  if HaveBest then
  begin
    AInfo := Best;
    Result := True;
  end
  else
    Result := False;
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
    /// <summary>Ancestor class ranges (parents up the inheritance chain).
    ///  Direct class is NOT included — that's TypeFile/ClassStartLine/
    ///  ClassEndLine. Used to recognise INHERITED members in body refs
    ///  (e.g. a body ref RowCount that LSP resolves to Vcl.Grids.pas line
    ///  665 — that's TStringGrid's range, which is an ancestor of a local
    ///  TStringGrid descendant).</summary>
    Ancestors: TArray<TAncestorClassRange>;
    /// <summary>Names of fields / properties / methods declared directly
    ///  in this class/record body (not inherited). Used as a fallback when
    ///  DelphiLSP fails to resolve bare body-identifiers in multi-target
    ///  with-clauses.</summary>
    Members: TArray<string>;
    Resolved: Boolean;
    /// <summary>True iff the target's last segment resolves to a
    ///  function/procedure/constructor/destructor — i.e. each evaluation
    ///  is a separate call with potential side effects. The rewriter
    ///  must introduce a temp var in that case to preserve single-eval
    ///  semantics of 'with'.</summary>
    IsMethod: Boolean;
    /// <summary>Verbatim type name as written in the declaration of the
    ///  target's last identifier (e.g. 'TStringGrid', 'PFoo',
    ///  'TArray&lt;Integer&gt;'). Empty when the rewriter could not
    ///  read the declaration line. Used for classic-mode var emission
    ///  ('LFoo: TStringGrid;').</summary>
    TypeName: string;
    /// <summary>File the target's last identifier was DECLARED in
    ///  (from the first GotoDefinition hop). Used as a fallback for
    ///  body-ref matching when TypeFile could not be resolved: members
    ///  of the target's type live in the same unit as the method/
    ///  property the target invokes, so any body-ref whose LSP result
    ///  lands in this file is very likely a sibling member.</summary>
    DeclFile: string;
    /// <summary>Diagnostic note populated when Resolved=False, describing
    ///  the reason resolution failed. Surfaced to the dialog debug view.</summary>
    ResolveNote: string;
    /// <summary>Unit qualifier from the declared parent class, e.g.
    ///  'Grids' for 'TStringGrid = class(Grids.TStringGrid)'. Empty when
    ///  there is no parent or the parent is unqualified. Used by Step 3
    ///  matching as a heuristic for inherited members when the explicit
    ///  ancestor walk failed (DelphiLSP often resolves the LAST segment
    ///  of a dotted parent to the LOCAL declaration of the same name,
    ///  not the unit-qualified one).</summary>
    ParentUnitHints: TArray<string>;
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

/// <summary>Resolves a TYPE NAME (e. g. 'TCanvas') at a known position
///  in a file to its class declaration's file + 1-based line range.
///  Follows pointer / simple aliases like 'PCanvas = ^TCanvas;' and
///  'tLiFo = _L_List;'. Used both directly (by ResolveTargetType after
///  Stage 1) and by the property-chain walker for dotted targets.</summary>
function ResolveTypeNameToClassBody(AClient: TLspClient;
  const AHomeFile: string; AHomeLine, AHomeCol1: Integer;
  out AClassFile: string; out AClassStart, AClassEnd: Integer): Boolean;
var
  TypeDefLocs: TArray<TLspLocation>;
begin
  Result := False;
  AClassFile := '';
  AClassStart := 0;
  AClassEnd := 0;

  try
    TypeDefLocs := AClient.GotoDefinition(AHomeFile, AHomeLine - 1, AHomeCol1 - 1);
  except
    Exit;
  end;
  if Length(TypeDefLocs) = 0 then Exit;

  // RefreshDocument the decl file so position queries on it work.
  try
    AClient.RefreshDocument(TLspUri.FileUriToPath(TypeDefLocs[0].Uri));
  except end;

  var AliasFile := TLspUri.FileUriToPath(TypeDefLocs[0].Uri);
  var AliasLine := TypeDefLocs[0].Range.Start.Line + 1;

  // Walk pointer / simple aliases (same logic as ResolveTargetType,
  // condensed). Up to 4 hops.
  for var Hop := 1 to 4 do
  begin
    var AliasSrc: string;
    if not TFile.Exists(AliasFile) then Break;
    try AliasSrc := TFile.ReadAllText(AliasFile); except Break end;
    var AliasIdx: TLineIndex;
    AliasIdx.Init(AliasSrc);
    if (AliasLine < 1) or (AliasLine > AliasIdx.LineCount) then Break;
    var PointeeCol1: Integer;
    var Pointee := ExtractPointeeNameFromAliasLine(
      AliasIdx.LineText(AliasLine), PointeeCol1);
    if Pointee = '' then Break;
    var HopLocs: TArray<TLspLocation>;
    try
      HopLocs := AClient.GotoDefinition(AliasFile, AliasLine - 1, PointeeCol1 - 1);
    except HopLocs := nil end;
    var Reused := False;
    if Length(HopLocs) > 0 then
    begin
      var NewFile := TLspUri.FileUriToPath(HopLocs[0].Uri);
      var NewLine := HopLocs[0].Range.Start.Line + 1;
      if SameText(NewFile, AliasFile) and (NewLine = AliasLine) then
        Reused := True
      else
      begin
        AliasFile := NewFile;
        AliasLine := NewLine;
      end;
    end;
    if (Length(HopLocs) = 0) or Reused then
    begin
      // Text fallback inside the same file.
      var Found := False;
      for var L := 1 to AliasIdx.LineCount do
      begin
        var Stripped := TrimLeft(AliasIdx.LineText(L));
        if Stripped.StartsWith(Pointee, True) then
        begin
          var After := Copy(Stripped, Length(Pointee) + 1, MaxInt);
          if (After = '') or not IsIdentCont(After[1]) then
          begin
            var AfterT := TrimLeft(After);
            if (AfterT <> '') and (AfterT[1] = '=') then
            begin
              AliasLine := L;
              Found := True;
              Break;
            end;
          end;
        end;
      end;
      if not Found then Break;
    end;
  end;

  FindClassRangeInFile(AliasFile, AliasLine, AClassStart, AClassEnd);
  if (AClassStart < 1) or (AClassEnd < AClassStart) then Exit;
  AClassFile := AliasFile;
  Result := True;
end;

/// <summary>For a dotted target like 'LListBox.Canvas' that DelphiLSP
///  couldn't resolve directly, walks the chain segment-by-segment:
///    1) GotoDef on first segment → var decl → extract base type
///    2) Find base type's class body
///    3) For each subsequent segment, find that member in the class
///       body, extract its declared type, walk to that type's class.
///  Returns Resolved=True only when the WHOLE chain resolved.</summary>
function ResolvePropertyChain(AClient: TLspClient;
  const AFileName: string; const ATarget: TWithTarget): TTargetTypeInfo;
var
  Expr: string;
  Segs: TArray<string>;
  CurFile: string;
  CurStart, CurEnd: Integer;
begin
  Result := Default(TTargetTypeInfo);
  Result.Expression := ATarget.Expression;
  Result.Resolved := False;

  Expr := ATarget.Expression.TrimRight;
  while (Expr <> '') and (Expr[Length(Expr)] = '^') do
    Expr := Copy(Expr, 1, Length(Expr) - 1).TrimRight;
  Segs := Expr.Split(['.']);
  for var SI := 0 to High(Segs) do Segs[SI] := Segs[SI].Trim;
  if Length(Segs) < 2 then Exit;
  for var SI := 0 to High(Segs) do
    if (Segs[SI] = '') or not IsIdentStart(Segs[SI][1]) then Exit;

  // Step 1: GotoDef on first segment at its known position.
  var DefLocs: TArray<TLspLocation>;
  try
    DefLocs := AClient.GotoDefinition(AFileName,
      ATarget.Range.StartPos.Line - 1, ATarget.Range.StartPos.Col - 1);
  except
    Exit;
  end;
  if Length(DefLocs) = 0 then Exit;

  var DeclFile := TLspUri.FileUriToPath(DefLocs[0].Uri);
  var DeclLine := DefLocs[0].Range.Start.Line + 1;
  if not TFile.Exists(DeclFile) then Exit;
  var DeclSrc: string;
  try DeclSrc := TFile.ReadAllText(DeclFile); except Exit end;
  var DeclIdx: TLineIndex;
  DeclIdx.Init(DeclSrc);
  if (DeclLine < 1) or (DeclLine > DeclIdx.LineCount) then Exit;
  var DeclLineText := DeclIdx.LineText(DeclLine);
  var TypeCol1: Integer;
  var TypeName := ExtractTypeNameFromDeclarationLine(DeclLineText, TypeCol1);
  if TypeName = '' then Exit;

  // Step 2: resolve that type name to its class body.
  if not ResolveTypeNameToClassBody(AClient, DeclFile,
    DeclLine, TypeCol1, CurFile, CurStart, CurEnd) then Exit;

  // Step 3: walk each subsequent segment.
  for var I := 1 to High(Segs) do
  begin
    // Find the member in the current class body (and any direct
    // ancestor we can reach textually).
    var MemTypeLine, MemTypeCol1: Integer;
    var MemType := FindMemberTypeInClassBody(CurFile, CurStart, CurEnd,
      Segs[I], MemTypeLine, MemTypeCol1);
    if MemType = '' then
    begin
      // Try walking up parents textually: read the class decl line,
      // extract parent, resolve parent's class body, search there.
      // Up to 6 hops.
      var SearchFile := CurFile;
      var SearchStart := CurStart;
      var Found := False;
      for var AH := 1 to 6 do
      begin
        var Src: string;
        if not TFile.Exists(SearchFile) then Break;
        try Src := TFile.ReadAllText(SearchFile); except Break end;
        var Idx: TLineIndex;
        Idx.Init(Src);
        if (SearchStart < 1) or (SearchStart > Idx.LineCount) then Break;
        var ParentCol1: Integer;
        var ParentName := ExtractParentClassFromDeclLine(
          Idx.LineText(SearchStart), ParentCol1);
        if (ParentName = '') or (ParentCol1 = 0) then Break;
        var ParentFile: string;
        var ParentStart, ParentEnd: Integer;
        if not ResolveTypeNameToClassBody(AClient, SearchFile,
          SearchStart, ParentCol1, ParentFile, ParentStart, ParentEnd) then Break;
        MemType := FindMemberTypeInClassBody(ParentFile,
          ParentStart, ParentEnd, Segs[I], MemTypeLine, MemTypeCol1);
        if MemType <> '' then
        begin
          // Found in this ancestor. Use that ancestor's coordinates
          // for the next type-name resolution.
          SearchFile := ParentFile;
          Found := True;
          Break;
        end;
        SearchFile := ParentFile;
        SearchStart := ParentStart;
      end;
      if not Found then Exit;
      // Replace CurFile with where we found the member so the
      // subsequent GotoDef hits an open document.
      CurFile := SearchFile;
    end;

    // Found the member with type MemType on (CurFile, MemTypeLine, MemTypeCol1).
    // Resolve that type name to its class body.
    if not ResolveTypeNameToClassBody(AClient, CurFile,
      MemTypeLine, MemTypeCol1, CurFile, CurStart, CurEnd) then Exit;
  end;

  // Build the result.
  Result.TypeFile := CurFile;
  Result.ClassStartLine := CurStart;
  Result.ClassEndLine := CurEnd;
  Result.TypeName := ''; // not tracked through the walk
  Result.DeclFile := DeclFile;
  Result.IsMethod := False;
  Result.Members := CollectClassMembers(CurFile, CurStart, CurEnd);

  // Ancestors of the final class (so inherited members of TCanvas etc.
  // also match).
  begin
    var WalkFile := CurFile;
    var WalkStart := CurStart;
    for var Step := 1 to 6 do
    begin
      var Src: string;
      if not TFile.Exists(WalkFile) then Break;
      try Src := TFile.ReadAllText(WalkFile); except Break end;
      var Idx: TLineIndex;
      Idx.Init(Src);
      if (WalkStart < 1) or (WalkStart > Idx.LineCount) then Break;
      var ParentCol1: Integer;
      var ParentName := ExtractParentClassFromDeclLine(
        Idx.LineText(WalkStart), ParentCol1);
      if (ParentName = '') or (ParentCol1 = 0) then Break;
      var ParentFile: string;
      var ParentStart, ParentEnd: Integer;
      if not ResolveTypeNameToClassBody(AClient, WalkFile,
        WalkStart, ParentCol1, ParentFile, ParentStart, ParentEnd) then Break;
      var Anc: TAncestorClassRange;
      Anc.TypeFile := ParentFile;
      Anc.ClassStartLine := ParentStart;
      Anc.ClassEndLine := ParentEnd;
      Result.Ancestors := Result.Ancestors + [Anc];
      WalkFile := ParentFile;
      WalkStart := ParentStart;
    end;
  end;

  Result.Resolved := True;
  Result.ResolveNote := 'resolved via property-chain walker (multi-segment dotted target)';
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
  LspLine, LspCol: Integer;       // 0-based LSP position for GotoDef
  EndsWithCaret: Boolean;
begin
  Result := Default(TTargetTypeInfo);
  Result.Expression := ATarget.Expression;
  Result.Resolved := False;

  // LSP position of the LAST identifier of the (possibly dotted) target.
  // The scanner sets EndPos AFTER the last character of the target. Many
  // expression shapes don't end in an identifier:
  //   p^                  -> ends in '^'
  //   Foo()               -> ends in ')'
  //   Foo(a, b)           -> ends in ')'
  //   Items[i]            -> ends in ']'
  //   (Foo as TBar)       -> ends in ')'
  // We need to back up onto the last identifier character before any
  // trailing group of pointer-derefs / parens / brackets / whitespace.
  // Strings inside () are walked over verbatim - we only care that they
  // contribute one balanced opener/closer.
  LspLine := ATarget.Range.EndPos.Line - 1;
  LspCol  := ATarget.Range.EndPos.Col - 1;
  begin
    var ExprT := ATarget.Expression.TrimRight;
    var I := Length(ExprT);
    // Strip trailing whitespace + balanced (...) / [...] groups + ^ /
    // dots. Stops at the FIRST identifier-character we encounter.
    while I >= 1 do
    begin
      var Ch := ExprT[I];
      if (Ch = ' ') or (Ch = #9) then begin Dec(I); Continue end;
      if Ch = '^' then begin Dec(I); Continue end;
      if Ch = '.' then begin Dec(I); Continue end;
      if (Ch = ')') or (Ch = ']') then
      begin
        // Walk back over a balanced group, ignoring string literals.
        var Closer := Ch;
        var Opener: Char;
        if Closer = ')' then Opener := '(' else Opener := '[';
        var Depth := 1;
        Dec(I);
        while (I >= 1) and (Depth > 0) do
        begin
          if ExprT[I] = '''' then
          begin
            Dec(I);
            while (I >= 1) and (ExprT[I] <> '''') do Dec(I);
            if I >= 1 then Dec(I);
            Continue;
          end;
          if ExprT[I] = Closer then Inc(Depth)
          else if ExprT[I] = Opener then Dec(Depth);
          Dec(I);
        end;
        Continue;
      end;
      // Anything else (incl. identifier chars) -> stop. I now points at
      // the LAST character of the identifier we want to query.
      Break;
    end;
    // Translate the 1-based-within-expr position I back to a source
    // (line, col) by walking the expression. The expression is verbatim
    // source from StartPos, so we just count newlines/columns.
    if I >= 1 then
    begin
      var SrcLine := ATarget.Range.StartPos.Line; // 1-based
      var SrcCol  := ATarget.Range.StartPos.Col;  // 1-based
      for var K := 1 to I - 1 do
      begin
        if ExprT[K] = #10 then
        begin
          Inc(SrcLine);
          SrcCol := 1;
        end
        else if ExprT[K] = #13 then
        begin
          // CR or CRLF - treat as newline; skip following LF.
          Inc(SrcLine);
          SrcCol := 1;
          // (next iteration will skip LF if present)
        end
        else
          Inc(SrcCol);
      end;
      LspLine := SrcLine - 1;        // -> 0-based
      LspCol  := SrcCol - 1;         // -> 0-based, points at I
    end;
  end;
  EndsWithCaret := False; // already handled by the walker above
  if EndsWithCaret then ;

  // Retry-Schleife: DelphiLSP wirft beim ersten Aufruf nach Cold-Start
  // gelegentlich noch 'Server not responding', auch wenn die Probe in
  // EnsureProjectIndexed schon durch war. Bis zu 5 Versuche mit jeweils
  // 1 s Pause - das deckt den Warmlauf-Tail ab, ohne den happy path
  // ernsthaft zu verzoegern.
  var LastError: string := '';
  for var Attempt := 1 to 5 do
  begin
    LastError := '';
    try
      DefLocs := AClient.GotoDefinition(AFileName, LspLine, LspCol);
    except
      on E: Exception do
      begin
        LastError := Format('%s: %s', [E.ClassName, E.Message]);
        SetLength(DefLocs, 0);
      end;
    end;
    if Length(DefLocs) > 0 then Break;
    if Attempt < 5 then Sleep(1000);
  end;
  if Length(DefLocs) = 0 then
  begin
    // Fallback 1: Property-Chain-Walker fuer mehrteilige Targets wie
    // 'LListBox.Canvas'. DelphiLSP weigert sich oft, das LETZTE Segment
    // (die Property) aufzuloesen, kann aber das erste Segment (die
    // Variable) plus die Klasse danach. Der Walker geht die Kette
    // textuell durch.
    var WalkerExpr := ATarget.Expression.TrimRight;
    while (WalkerExpr <> '') and (WalkerExpr[Length(WalkerExpr)] = '^') do
      WalkerExpr := Copy(WalkerExpr, 1, Length(WalkerExpr) - 1).TrimRight;
    var WalkerNote: string := '';
    if Pos('.', WalkerExpr) > 0 then
    begin
      var Walked := ResolvePropertyChain(AClient, AFileName, ATarget);
      if Walked.Resolved and (Walked.TypeFile <> '')
        and (Walked.ClassStartLine > 0) then
      begin
        Result := Walked;
        Exit;
      end;
      WalkerNote := Format(
        ' Property-chain walker also failed (Resolved=%s, TypeFile=%s, '
        + 'ClassStart=%d). ',
        [BoolToStr(Walked.Resolved, True), ExtractFileName(Walked.TypeFile),
         Walked.ClassStartLine]);
    end;

    // Keine Aufloesung moeglich: weder Stage 1 noch der Property-Chain-
    // Walker fanden was. Das passiert v.a. fuer Code in inaktiven
    // {$IFDEF}-Bloecken - der Compiler liest ihn nicht, also indexiert
    // LSP ihn nicht. In dem Fall waeren auch alle Body-Refs nicht
    // verlaesslich aufzuloesen, eine partielle Umschreibung
    // (Temp-Var ohne Body-Qualifizierung) waere irrefuehrend.
    // Daher: Resolved=False, wriTypeUnresolved, with-Statement
    // bleibt komplett unveraendert.
    Result.Resolved := False;
    if LastError <> '' then
      Result.ResolveNote := Format(
        'LSP unable to resolve target: GotoDefinition raised after 5 '
        + 'retries: %s.%s Likely cause: code is in an inactive '
        + '{$IFDEF}-block or LSP has not indexed this region. '
        + 'Statement left unchanged.',
        [LastError, WalkerNote])
    else
      Result.ResolveNote := Format(
        'LSP returned 0 locations for target @ L%d:C%d after 5 retries.%s '
        + 'Likely cause: code is in an inactive {$IFDEF}-block or LSP '
        + 'has not indexed this region. Statement left unchanged.',
        [LspLine, LspCol, WalkerNote]);
    Exit;
  end;

  DeclFile := DefLocs[0].Uri;
  // TLspLocation.Uri is a file:// URI. Convert via Lsp.Uri if needed —
  // for now do the simple file:/// strip.
  DeclFile := TLspUri.FileUriToPath(DeclFile);
  Result.DeclFile := DeclFile;

  DeclLine := DefLocs[0].Range.Start.Line + 1; // -> 1-based

  // Read the declaration line and extract the type name.
  if not TFile.Exists(DeclFile) then
  begin
    // Sonderfall "Precompiled-Package-Typ": LSP kennt den Typ und
    // verweist auf eine .pas (z.B. DockForm.pas in DesignIde.bpl),
    // die nur als .dcu ausgeliefert wird - die Quelle ist nicht da.
    // Damit koennen wir Members nicht parsen, aber die Temp-Var-
    // Entscheidung treffen wir rein textuell und die deckt bereits
    // Bug #1 (Mehrfach-Evaluation) ab. Body-Refs bleiben dann
    // unqualifiziert - aber der Code ist KEINE Dead-Code-Region
    // (LSP konnte den Typ ja aufloesen, also ist die {$IFDEF}-Region
    // hier aktiv).
    Result.Resolved := True;
    Result.ResolveNote := Format(
      'partial: decl file does not exist on disk (%s) - likely a type '
      + 'from a precompiled package (e.g. DesignIde). Temp var will be '
      + 'emitted, body-ref qualification skipped (needs manual review).',
      [ExtractFileName(DeclFile)]);
    // Heuristik fuer "X as TType": TType direkt aus dem Ausdruck holen,
    // damit DeriveTempName ein lesbares Temp-Var ableiten kann.
    var ExprT := ATarget.Expression.TrimRight;
    var Lower := AnsiLowerCase(ExprT);
    var AsPos := Pos(' as ', Lower);
    if AsPos > 0 then
    begin
      var AfterAs := Trim(Copy(ExprT, AsPos + 4, MaxInt));
      var EndP := 1;
      while (EndP <= Length(AfterAs)) and IsIdentCont(AfterAs[EndP]) do
        Inc(EndP);
      if EndP > 1 then
        Result.TypeName := Copy(AfterAs, 1, EndP - 1);
    end;
    Exit;
  end;
  try
    TmpSource := TFile.ReadAllText(DeclFile);
  except
    on E: Exception do
    begin
      Result.ResolveNote := Format(
        'cannot read decl file %s: %s: %s',
        [ExtractFileName(DeclFile), E.ClassName, E.Message]);
      Exit;
    end;
  end;
  TmpIndex.Init(TmpSource);
  DeclLineText := TmpIndex.LineText(DeclLine);

  // Methodendetektion: GotoDef hat uns auf die Deklarationszeile
  // des Bezeichners gebracht. Wenn die Zeile mit
  // 'function'/'procedure'/'class function'/... beginnt, ist das
  // Target ein Methodenaufruf — wir MÜSSEN dann eine Temp-Variable
  // erzeugen, weil sonst bei direkter Inline-Qualifizierung der Call
  // mehrfach ausgewertet würde.
  Result.IsMethod := DeclLineDeclaresMethod(DeclLineText);

  TypeName := ExtractTypeNameFromDeclarationLine(DeclLineText, TypeNameCol1);

  // Sonderfall Konstruktor/Destruktor: hat keinen ':'-Rueckgabetyp.
  // Bei 'with TFoo.Create(...) do' landet LSP auf
  //   'constructor TFoo.Create(const FileName: string);'
  // Der "Typ" ist hier implizit die Klasse, in der der Konstruktor
  // deklariert ist (TFoo). Den finden wir am Class-Qualifier des
  // Methodennamens. Greift bevor wir auf den 'class = class(...)'-
  // Sonderfall fallen.
  if TypeName = '' then
  begin
    TypeName := ExtractConstructorClassName(DeclLineText, TypeNameCol1);
    if TypeName <> '' then
      Result.IsMethod := True;
  end;

  // Sonderfall: die Decl-Zeile IST selbst eine Typ-Deklaration
  //   'Ident = class(Parent)' / 'interface(Base)' / 'record' / 'object'
  // Das passiert wenn das Target ein Typ-Cast war ('Comp as TWinControl')
  // - GotoDef springt dann direkt auf die Typ-Deklaration, nicht auf
  // eine Variable. ExtractTypeNameFromDeclarationLine findet hier kein
  // ':' nach dem Variablennamen und gibt '' zurueck. Wir behandeln das
  // explizit: die LHS des '=' IST der gesuchte Typ, DeclFile/DeclLine
  // ist seine Lokation - Stage 2 entfaellt komplett.
  if TypeName = '' then
  begin
    var Trim1 := TrimLeft(DeclLineText);
    var EqPos := Pos('=', Trim1);
    if EqPos > 0 then
    begin
      var After := TrimLeft(Copy(Trim1, EqPos + 1, MaxInt));
      var AfterUpper := UpperCase(After);
      if AfterUpper.StartsWith('CLASS')
        or AfterUpper.StartsWith('INTERFACE')
        or AfterUpper.StartsWith('RECORD')
        or AfterUpper.StartsWith('OBJECT')
        or AfterUpper.StartsWith('DISPINTERFACE') then
      begin
        // LHS-Identifier extrahieren (vor dem ersten Whitespace/'=').
        var LhsEnd := 1;
        while (LhsEnd <= Length(Trim1)) and IsIdentCont(Trim1[LhsEnd]) do
          Inc(LhsEnd);
        if LhsEnd > 1 then
        begin
          TypeName := Copy(Trim1, 1, LhsEnd - 1);
          Result.TypeName := TypeName;
          Result.TypeFile := DeclFile;
          // Klassen-Range direkt aus der gefundenen Zeile bestimmen.
          FindClassRangeInFile(DeclFile, DeclLine,
            Result.ClassStartLine, Result.ClassEndLine);
          Result.Members := CollectClassMembers(Result.TypeFile,
            Result.ClassStartLine, Result.ClassEndLine);
          Result.Resolved := True;
          Result.ResolveNote := Format(
            'decl line is the type declaration itself ("%s"); skipped stage 2.',
            [Trim(DeclLineText)]);
          Exit;
        end;
      end;
    end;
    Result.ResolveNote := Format(
      'decl line %d in %s has no type name; line text: %s',
      [DeclLine, ExtractFileName(DeclFile), Trim(DeclLineText)]);
    Exit;
  end;
  Result.TypeName := TypeName;

  // Sonderfall Methoden-Target: das LSP-Ziel war eine Funktion, deren
  // Rueckgabetyp wir wollen. ExtractTypeNameFromDeclarationLine liefert
  // bei 'function Foo: TBar;' bereits TBar (es springt zum ersten ':'
  // ausserhalb der Parameterliste). Gut.

  // Now resolve the type name itself.
  // Position-based queries (textDocument/definition with a {line,col})
  // only work reliably for documents the LSP has opened via didOpen.
  // The first GotoDefinition succeeded because Unit11.pas was didOpen'd
  // up-front. The decl file we landed on (e.g. Vcl.ComCtrls.pas) was
  // NOT didOpen'd - DelphiLSP knew of its existence from its index, but
  // refuses position-based queries on it. Force a RefreshDocument here
  // so the second resolution stage has the file in the workspace.
  try
    AClient.RefreshDocument(DeclFile);
  except
    // Non-fatal: if didOpen fails the GotoDefinition below will simply
    // return 0 locations and the diagnostic note covers the failure.
  end;
  try
    TypeDefLocs := AClient.GotoDefinition(DeclFile, DeclLine - 1, TypeNameCol1 - 1);
  except
    on E: Exception do
    begin
      Result.ResolveNote := Format(
        'LSP GotoDefinition for type "%s" in %s @ L%d:C%d raised %s: %s',
        [TypeName, ExtractFileName(DeclFile), DeclLine - 1, TypeNameCol1 - 1,
         E.ClassName, E.Message]);
      Exit;
    end;
  end;
  if Length(TypeDefLocs) = 0 then
  begin
    // Degrade gracefully: we still have IsMethod, TypeName and the first
    // decl file. The rewriter only needs IsMethod+TypeName to pick the
    // temp-var path (which is what fixes side-effect duplication). Member
    // qualification will fall back to LSP body-ref resolution when
    // TypeFile/ClassRange are unknown.
    Result.Resolved := True;
    Result.ResolveNote := Format(
      'partial: type "%s" decl-line known, but LSP could not resolve the '
      + 'type itself in %s @ L%d:C%d (decl line text: "%s") - falling back '
      + 'to body-ref-only qualification',
      [TypeName, ExtractFileName(DeclFile), DeclLine - 1, TypeNameCol1 - 1,
       Trim(DeclLineText)]);
    Exit;
  end;

  Result.TypeFile := TLspUri.FileUriToPath(TypeDefLocs[0].Uri);

  // Pointer-Alias folgen: 'PFoo = ^TFoo;' / einfacher Alias 'A = B;' —
  // der Class-Range-Finder erkennt das nicht als Klassenkörper. Hier
  // explizit dereferenzieren (maximal ein paar Hops, um Endlosschleifen
  // über fehlerhafte Quellen auszuschliessen).
  begin
    var AliasFile := Result.TypeFile;
    var AliasLine := TypeDefLocs[0].Range.Start.Line + 1;
    var Hops := 0;
    while Hops < 4 do
    begin
      var AliasSrc: string;
      if not TFile.Exists(AliasFile) then Break;
      try
        AliasSrc := TFile.ReadAllText(AliasFile);
      except
        Break;
      end;
      var AliasIdx: TLineIndex;
      AliasIdx.Init(AliasSrc);
      var AliasLineText := AliasIdx.LineText(AliasLine);
      var PointeeCol1: Integer;
      var Pointee := ExtractPointeeNameFromAliasLine(AliasLineText, PointeeCol1);
      if Pointee = '' then Break;

      var Hop: TArray<TLspLocation>;
      try
        Hop := AClient.GotoDefinition(AliasFile, AliasLine - 1, PointeeCol1 - 1);
      except
        Hop := nil;
      end;

      // Fallback wenn LSP nichts findet ODER wieder die gleiche Stelle
      // zurueckgibt (DelphiLSP-Quirk: GotoDef auf den Pointee in einem
      // Pointer-Alias landet manchmal auf dem Alias selber statt auf der
      // tatsaechlichen Typ-Deklaration). Wir suchen den Pointee dann
      // textuell in derselben Datei: '<Pointee> = ...'.
      var Reused := False;
      if Length(Hop) > 0 then
      begin
        var NewFile := TLspUri.FileUriToPath(Hop[0].Uri);
        var NewLine := Hop[0].Range.Start.Line + 1;
        if SameText(NewFile, AliasFile) and (NewLine = AliasLine) then
          Reused := True
        else
        begin
          AliasFile := NewFile;
          AliasLine := NewLine;
        end;
      end;
      if (Length(Hop) = 0) or Reused then
      begin
        // Text-Fallback: scan derselben Datei nach 'Pointee = ' am
        // Zeilenanfang (nach optionalem Whitespace). Klappt fuer normale
        // type-Block-Layouts. Wenn nichts gefunden: aufhoeren.
        var Found := False;
        for var ScanLine := 1 to AliasIdx.LineCount do
        begin
          var LT := AliasIdx.LineText(ScanLine);
          var Stripped := TrimLeft(LT);
          if Stripped.StartsWith(Pointee, True) then
          begin
            // word-boundary: nach Pointee muss Nicht-Ident-Char kommen.
            var After := Copy(Stripped, Length(Pointee) + 1, MaxInt);
            if (After = '') or not IsIdentCont(After[1]) then
            begin
              // '=' nach optionalem Whitespace?
              var AfterTrim := TrimLeft(After);
              if (AfterTrim <> '') and (AfterTrim[1] = '=') then
              begin
                AliasLine := ScanLine;
                Found := True;
                Break;
              end;
            end;
          end;
        end;
        if not Found then Break;
      end;
      Inc(Hops);
    end;
    Result.TypeFile := AliasFile;
    FindClassRangeInFile(AliasFile, AliasLine,
      Result.ClassStartLine, Result.ClassEndLine);
  end;

  // Parse direct (non-inherited) member names. Used as a fallback when
  // DelphiLSP fails to resolve bare body-identifiers in a multi-target
  // with-clause (it appears LSP only checks the rightmost target's class).
  Result.Members := CollectClassMembers(Result.TypeFile,
    Result.ClassStartLine, Result.ClassEndLine);

  // Vererbungskette aufbauen: bei einer 'class(Parent)'-Deklaration
  // hopsen wir GotoDef-weise hoch und sammeln die Class-Ranges aller
  // Vorfahren. Body-Refs, die LSP auf eine geerbte Property (z.B.
  // TStringGrid.RowCount in Vcl.Grids.pas) aufloest, matchen sonst
  // nicht, weil der direkte Klassen-Range nur den lokalen Descendant
  // abdeckt. Maximal 6 Hops gegen Endlosschleifen.
  if (Result.TypeFile <> '')
    and (Result.ClassStartLine > 0) then
  begin
    var CurFile := Result.TypeFile;
    var CurLine := Result.ClassStartLine;
    for var Step := 1 to 6 do
    begin
      var CurSrc: string;
      if not TFile.Exists(CurFile) then Break;
      try CurSrc := TFile.ReadAllText(CurFile); except Break end;
      var CurIdx: TLineIndex;
      CurIdx.Init(CurSrc);
      if (CurLine < 1) or (CurLine > CurIdx.LineCount) then Break;
      var ParentCol1: Integer;
      var ParentName := ExtractParentClassFromDeclLine(
        CurIdx.LineText(CurLine), ParentCol1);
      if (ParentName = '') or (ParentCol1 = 0) then Break;

      // Wenn der Parent in der Quelle dotted geschrieben war
      // ('Grids.TStringGrid'), die Unit-Segmente als Hint merken -
      // GotoDef auf die letzte Komponente landet wegen Name-Shadowing
      // oft im LOKALEN Descendant, der explizite Ancestor-Walk bringt
      // dann nichts mehr. Step 3 im Body-Ref-Matching nutzt diese Hints
      // als File-Pattern.
      var ParentLine := CurIdx.LineText(CurLine);
      var OpenP := Pos('(', ParentLine);
      if OpenP > 0 then
      begin
        var ClosP := Pos(')', ParentLine, OpenP + 1);
        if ClosP = 0 then ClosP := Length(ParentLine) + 1;
        var Inner := Copy(ParentLine, OpenP + 1, ClosP - OpenP - 1);
        var CommaP := Pos(',', Inner);
        if CommaP > 0 then Inner := Copy(Inner, 1, CommaP - 1);
        Inner := Inner.Trim;
        if Inner.Contains('.') then
        begin
          var Segs := Inner.Split(['.']);
          // alle Segmente AUSSER dem letzten (Klassennamen) sind
          // Unit-Qualifier - die merken wir uns. Duplikate vermeiden.
          for var SI := 0 to Length(Segs) - 2 do
          begin
            var Seg := Segs[SI].Trim;
            var Exists := False;
            for var EH in Result.ParentUnitHints do
              if SameText(EH, Seg) then begin Exists := True; Break end;
            if (Seg <> '') and not Exists then
              Result.ParentUnitHints := Result.ParentUnitHints + [Seg];
          end;
        end;
      end;

      var ParentDef: TArray<TLspLocation>;
      try
        ParentDef := AClient.GotoDefinition(CurFile,
          CurLine - 1, ParentCol1 - 1);
      except
        Break;
      end;
      if Length(ParentDef) = 0 then Break;

      var NewFile := TLspUri.FileUriToPath(ParentDef[0].Uri);
      var NewLine := ParentDef[0].Range.Start.Line + 1;
      if SameText(NewFile, CurFile) and (NewLine = CurLine) then Break;

      var AncStart, AncEnd: Integer;
      FindClassRangeInFile(NewFile, NewLine, AncStart, AncEnd);
      if (AncStart < 1) or (AncEnd < AncStart) then Break;

      var Anc: TAncestorClassRange;
      Anc.TypeFile := NewFile;
      Anc.ClassStartLine := AncStart;
      Anc.ClassEndLine := AncEnd;
      Result.Ancestors := Result.Ancestors + [Anc];

      CurFile := NewFile;
      CurLine := AncStart;
    end;
  end;

  Result.Resolved := True;
end;

{ TWithRewriteSettings }

class function TWithRewriteSettings.Defaults: TWithRewriteSettings;
begin
  Result.UseInlineVars := True;
end;

{ TWithRewriteResult }

function TWithRewriteResult.IsAutoRewritable: Boolean;
begin
  Result := (Issues = []) and (NewText <> '');
end;

{ TWithRewriter }

class function TWithRewriter.Rewrite(AClient: TLspClient; const AFileName: string;
  const ASource: string; const AOccurrence: TWithOccurrence;
  const ASettings: TWithRewriteSettings): TWithRewriteResult;
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
        Result.Debug.Targets[DI].ResolveNote    := Targets[DI].ResolveNote;
        Result.Debug.Targets[DI].Ancestors      := Targets[DI].Ancestors;
        Result.Debug.Targets[DI].ParentUnitHints := Targets[DI].ParentUnitHints;
      end;
      Exit;
    end;
    // Pick the inline-var name for this target.
    //
    //   - Target is a side-effecting call (IsMethod) → temp var
    //     'LWithN' or 'L' + name (so the call runs only once, matching
    //     'with' semantics).
    //   - Target is a simple dotted path (e.g. 'FParser') OR ends with
    //     '^' (e.g. 'p^', 'Self.FNode^') and is otherwise safe →
    //     NO temp var, qualify with the expression directly.
    //   - Anything else (parens, calls inside, brackets, …) →
    //     temp var 'LWithN'.
    //
    // Cross-target collisions and body-identifier collisions fall back
    // to 'LWithN'.
    if Targets[I].IsMethod then
    begin
      // Even simple-dotted paths get a temp when the last segment
      // resolves to a method — direct inline qualification would call
      // it multiple times.
      // When we know the TypeName from LSP resolution, prefer it as
      // the naming base - it describes what the variable *holds*, which
      // is more useful than the method name. Especially important for
      // constructor targets ('TFoo.Create' -> 'LFoo' instead of 'LCreate'),
      // and avoids accidental hits on identifier-looking content inside
      // string literals of the call arguments
      // ('TRegIniFile.Create(''\Software\GExperts'')' would otherwise
      // produce 'LGExperts' from the string).
      if Targets[I].TypeName <> '' then
        Targets[I].InlineVarName := DeriveTempName(Targets[I].TypeName)
      else if IsSimpleDottedTarget(Targets[I].Expression) then
      begin
        var Segs := SplitDottedSegments(Targets[I].Expression);
        var Last := '';
        if Length(Segs) > 0 then Last := Segs[High(Segs)];
        Targets[I].InlineVarName := DeriveTempName(Last);
      end;
      if Targets[I].InlineVarName = '' then
        Targets[I].InlineVarName := FallbackTempName(Targets[I].Expression, I + 1);
      Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
    end
    else if IsSideEffectFreeTarget(Targets[I].Expression) then
    begin
      // No temp var: qualify directly. Works for 'FParser', 'Self.FFoo',
      // 'p^', 'Self.FNode^'. Members are accessed via 'expr.Member';
      // for caret form this gives 'p^.Info' which is the desired shape.
      // Chains with 2+ dots (e. g. 'A.B.C') are intentionally NOT inlined
      // here — intermediate property accessors / parameterless functions
      // could have side effects, so they fall through to the temp-var
      // branch below to preserve single-evaluation 'with' semantics.
      Targets[I].InlineVarName := '';
      Targets[I].QualifyPrefix := Targets[I].Expression.TrimRight + '.';
    end
    else
    begin
      Targets[I].InlineVarName := FallbackTempName(Targets[I].Expression, I + 1);
      Targets[I].QualifyPrefix := Targets[I].InlineVarName + '.';
    end;
  end;

  // Cross-target collision check: if two targets ended up with the same
  // InlineVarName, append the 1-based index of the later one so we get
  // 'LFoo' / 'LFoo2' instead of duplicates. No underscores - those are
  // discouraged by the Delphi style guide.
  for I := 0 to High(Targets) do
    if Targets[I].InlineVarName <> '' then
      for var J := 0 to I - 1 do
        if SameText(Targets[J].InlineVarName, Targets[I].InlineVarName) then
        begin
          Targets[I].InlineVarName :=
            Targets[I].InlineVarName + IntToStr(I + 1);
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
  // the name we picked for an inline-var, append the 1-based target
  // index. Bumping rather than re-deriving guarantees we move away from
  // the colliding name; FallbackTempName would just regenerate the same
  // 'L'+last-segment.
  for I := 0 to High(Targets) do
    if Targets[I].InlineVarName <> '' then
      for var BR := 0 to High(BodyRefs) do
        if SameText(BodyRefs[BR].Name, Targets[I].InlineVarName) then
        begin
          Targets[I].InlineVarName :=
            Targets[I].InlineVarName + IntToStr(I + 1);
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
            // Direct class range
            var InRange :=
                  SameText(RefDefFile, Targets[I].TypeFile)
              and (RefDefLine >= Targets[I].ClassStartLine)
              and (RefDefLine <= Targets[I].ClassEndLine);
            // ... or any ancestor's class range (inherited members)
            if not InRange then
              for var A := 0 to High(Targets[I].Ancestors) do
                if SameText(RefDefFile, Targets[I].Ancestors[A].TypeFile)
                  and (RefDefLine >= Targets[I].Ancestors[A].ClassStartLine)
                  and (RefDefLine <= Targets[I].Ancestors[A].ClassEndLine) then
                begin
                  InRange := True;
                  Break;
                end;
            if InRange then
            begin
              Prefixes.AddOrSetValue(Ref.StartIdx, Targets[I].QualifyPrefix);
              Matched := True;
              DbgRef.MatchSource := dmLsp;
              DbgRef.MatchedTargetIdx := I;
              DbgRef.AppliedPrefix := Targets[I].QualifyPrefix;
              Break;
            end;
          end;

          // Fallback: when the second LSP resolve step (TypeName -> type
          // definition) failed, Targets[I].TypeFile is empty and the
          // class-range is unknown. But we still know which file the
          // target's method/property was declared in (DeclFile from the
          // first hop). Members of the target's type almost always live
          // in the same .pas file as the method we resolved (Delphi
          // convention: one class, one unit). Use file-equality with
          // DeclFile as a soft match when the strict range check above
          // could not run.
          if not Matched then
          begin
            for I := High(Targets) downto 0 do
            begin
              if (Targets[I].TypeFile = '')
                and (Targets[I].DeclFile <> '')
                and SameText(RefDefFile, Targets[I].DeclFile) then
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

          // Letzter Fallback: ParentUnitHints. Bei einer lokalen Klasse
          // wie 'TStringGrid = class(Grids.TStringGrid)' scheitert der
          // explizite Ancestor-Walk haeufig (LSP loest 'TStringGrid' im
          // Parent-Ausdruck auf den lokalen Descendant). Wir haben dann
          // wenigstens 'Grids' als Hint gespeichert. Wenn ein Body-Ref
          // LSP-Result in einer Datei landet, deren Basename eines der
          // Hint-Segmente enthaelt ('Vcl.Grids.pas' -> Token 'Grids'
          // matcht 'Grids'), akzeptieren wir das als geerbten Member.
          if not Matched then
          begin
            var RefBase := ChangeFileExt(ExtractFileName(RefDefFile), '');
            var RefSegs := RefBase.Split(['.']);
            for I := High(Targets) downto 0 do
            begin
              if Length(Targets[I].ParentUnitHints) = 0 then Continue;
              var HintHit := False;
              for var H in Targets[I].ParentUnitHints do
                for var RS in RefSegs do
                  if SameText(H, RS) then begin HintHit := True; Break end;
              if HintHit then
              begin
                Prefixes.AddOrSetValue(Ref.StartIdx, Targets[I].QualifyPrefix);
                DbgRef.MatchSource := dmLsp;
                DbgRef.MatchedTargetIdx := I;
                DbgRef.AppliedPrefix := Targets[I].QualifyPrefix;
                Break;
              end;
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
      Result.Debug.Targets[I].ResolveNote    := Targets[I].ResolveNote;
      Result.Debug.Targets[I].Ancestors      := Targets[I].Ancestors;
      Result.Debug.Targets[I].ParentUnitHints := Targets[I].ParentUnitHints;
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
      wbkCompound:
        begin
          // try..end / case..end / asm..end body. We do NOT wrap in an
          // extra begin..end - the compound block has its own scope
          // shape. Var-decls go IN FRONT OF the block; since we already
          // need inline-vars at statement level (Delphi 10.3+), this is
          // syntactically fine inside any procedure body.
          //
          // Indentation: BodyOnly[1] is the opener keyword ('try' etc.)
          // with NO leading whitespace (we sliced from there). But lines
          // 2..N still carry their ORIGINAL indent, computed relative to
          // the opener's old column. After replacement the opener lives
          // at IndentStr (the 'with' keyword's indent), which is
          // typically less than the original opener column. We shift the
          // tail of the body left by that difference so the block stays
          // self-consistent.
          if HasInlineVar then
          begin
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.Append('var ').Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';')
                       .AppendLine.Append(IndentStr);
          end;
          Builder.Append(ShiftCompoundBody(
            BodyOnly.TrimLeft,
            AOccurrence.BodyInnerRange.StartPos.Col - 1,
            Length(IndentStr)));
        end;
    end;

    Result.NewText := Builder.ToString;
  finally
    Builder.Free;
  end;

  // === Classic-mode (non-inline-var) artifacts =============================
  //
  // Compute these whenever the inline form uses any inline-var, so the
  // dialog can switch modes without re-running the rewriter. Failure
  // here leaves Result.Classic.Supported = False; the dialog will then
  // add wriRequiresInlineVar when the user opts out of inline vars.
  Result.Classic.Supported := False;
  if HasInlineVar then
  begin
    // 1) Body text: same shape as NewText but with plain 'Name := Expr;'
    //    in place of 'var Name := Expr;'.
    Builder := TStringBuilder.Create;
    try
      case AOccurrence.BodyKind of
        wbkBeginEnd:
          begin
            Builder.Append('begin');
            var BodyIndent := GetBodyStatementIndent(BodyOnly, IndentStr + '  ');
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.AppendLine.Append(BodyIndent)
                       .Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';');
            Builder.Append(BodyOnly);
            Builder.Append('end');
          end;
        wbkSingle:
          begin
            var BodyIndent := IndentStr + '  ';
            Builder.Append('begin');
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.AppendLine.Append(BodyIndent)
                       .Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';');
            Builder.AppendLine.Append(BodyIndent).Append(BodyOnly.Trim);
            Builder.AppendLine.Append(IndentStr).Append('end');
          end;
        wbkCompound:
          begin
            // Classic mode for try/case/asm: prepend assignments at the
            // same indent level, then emit the block (left-shifted by
            // the original-vs-new opener-indent difference).
            for I := 0 to High(Targets) do
              if Targets[I].InlineVarName <> '' then
                Builder.Append(Targets[I].InlineVarName)
                       .Append(' := ').Append(Targets[I].Expression).Append(';')
                       .AppendLine.Append(IndentStr);
            Builder.Append(ShiftCompoundBody(
              BodyOnly.TrimLeft,
              AOccurrence.BodyInnerRange.StartPos.Col - 1,
              Length(IndentStr)));
          end;
      end;
      Result.Classic.BodyText := Builder.ToString;
    finally
      Builder.Free;
    end;

    // 2) Find the enclosing method.
    var Methods := FindAllMethodsInSource(ASource, Index);
    var MI: TMethodInfo;
    var HaveMethod := FindEnclosingMethod(Methods, AOccurrence.KeywordPos.Line, MI);
    if HaveMethod then
    begin
      Result.Classic.MethodKey :=
        AFileName + ':' + IntToStr(MI.HeaderLine);
      Result.Classic.HasVarSection := MI.HasVarSection;
      Result.Classic.VarSectionLastLine := MI.VarSectionLastLine;
      Result.Classic.MethodBodyBeginLine := MI.BodyBeginLine;
      Result.Classic.MethodBodyBeginCol := MI.BodyBeginCol;
      if MI.LocalIndent <> '' then
        Result.Classic.LocalIndent := MI.LocalIndent
      else
        Result.Classic.LocalIndent := '  ';

      // 3) Per-target var-decls.
      var VarDecls := TList<TClassicVarDecl>.Create;
      try
        for I := 0 to High(Targets) do
          if Targets[I].InlineVarName <> '' then
          begin
            var VD: TClassicVarDecl;
            VD.Name := Targets[I].InlineVarName;
            VD.TypeName := Targets[I].TypeName;
            VarDecls.Add(VD);
          end;
        Result.Classic.VarDecls := VarDecls.ToArray;
      finally
        VarDecls.Free;
      end;

      // 4) Uses-clause: scan + determine missing units.
      var UsesScan := ScanUsesClauses(ASource, Index);
      Result.Classic.IntfUsesFound := UsesScan.InterfaceUses.Found;
      Result.Classic.IntfUsesLastLine := UsesScan.InterfaceUses.LastBeforeSemiLine;
      Result.Classic.IntfUsesLastCol := UsesScan.InterfaceUses.LastBeforeSemiCol;
      Result.Classic.ImplUsesFound := UsesScan.ImplementationUses.Found;
      Result.Classic.ImplUsesLastLine := UsesScan.ImplementationUses.LastBeforeSemiLine;
      Result.Classic.ImplUsesLastCol := UsesScan.ImplementationUses.LastBeforeSemiCol;
      Result.Classic.ImplKeywordLine := UsesScan.ImplementationKeywordLine;
      Result.Classic.ImplKeywordCol := UsesScan.ImplementationKeywordCol;

      var AddUnits := TList<string>.Create;
      try
        var AllResolved := True;
        for I := 0 to High(Targets) do
          if Targets[I].InlineVarName <> '' then
          begin
            // Determine the unit name of this type's source file.
            if (Targets[I].TypeFile = '') or (not TFile.Exists(Targets[I].TypeFile)) then
            begin
              AllResolved := False;
              Break;
            end;
            var TypeSrc: string;
            try
              TypeSrc := TFile.ReadAllText(Targets[I].TypeFile);
            except
              AllResolved := False;
              Break;
            end;
            var UN := GetUnitNameFromSource(TypeSrc);
            if UN = '' then
            begin
              AllResolved := False;
              Break;
            end;
            // Already in either uses clause? Skip.
            if UsesContains(UsesScan.InterfaceUses, UN) then Continue;
            if UsesContains(UsesScan.ImplementationUses, UN) then Continue;
            // Already added in this pass?
            var Dup := False;
            for var K := 0 to AddUnits.Count - 1 do
              if SameText(AddUnits[K], UN) then begin Dup := True; Break; end;
            if not Dup then AddUnits.Add(UN);
          end;
        if AllResolved then
        begin
          Result.Classic.AddUnits := AddUnits.ToArray;
          Result.Classic.Supported := True;
        end;
      finally
        AddUnits.Free;
      end;
    end;
  end;

  // Wenn das Rewrite eine 10.3+ Inline-Variable bräuchte, der Aufrufer
  // aber nicht im Inline-Modus arbeiten will, kennzeichnen wir das hier.
  // Wenn allerdings Classic.Supported = True ist, lässt sich das auch ohne
  // Inline-Var erledigen — dann KEIN Issue setzen.
  if HasInlineVar and (not ASettings.UseInlineVars) and (not Result.Classic.Supported) then
    Include(Result.Issues, wriRequiresInlineVar);
end;

end.
