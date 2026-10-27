(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SafeDeletePlan;

// SAFE DELETE, the pure half (idea: Ian Branch, issue #11): WHAT would be
// deleted for the symbol declared at a given line, and which properties of
// the declaration forbid deleting it at all. The usage check (text scan +
// DelphiLSP verification, form files) lives in Expert.SafeDelete.
//
// Supported: methods (class / record / interface), free routines,
// fields, variables, constants, properties and single-line types.
// Refused (a "veto" with the reason): overloaded, virtual / dynamic /
// abstract / override / message methods (dispatch can reach them without
// a textual call), published members (streaming / RTTI), multi-line
// declarations of data, and types with a body.

interface

uses
  System.SysUtils, System.Types;

type
  TSafeDeleteKind = (sdkUnknown, sdkMethod, sdkRoutine, sdkField, sdkProperty,
    sdkVariable, sdkConstant, sdkType);

  /// <summary>Whole lines FirstLine..LastLine (0-based, inclusive) are
  ///  removed - or, with HasReplacement, replaced by the ONE line
  ///  Replacement (a name dropped from "A, B, C: Integer;").</summary>
  TSafeDeleteEdit = record
    FirstLine, LastLine: Integer;
    HasReplacement: Boolean;
    Replacement: string;
  end;

  TSafeDeleteSymbol = record
    Name: string;
    Kind: TSafeDeleteKind;
    Container: string;        // owning class / record / interface ('' = unit level)
    ContainerLine: Integer;   // its header line, -1 = none
    IsInterfaceMember: Boolean;
    Parents: TArray<string>;  // the container's parent list (interface veto)
    DeclLine: Integer;        // the declaration
    ImplLine: Integer;        // implementation header, -1 = none
    LocalFirst, LocalLast: Integer;  // a routine-local symbol: its routine; else -1
    Edits: TArray<TSafeDeleteEdit>;
    Vetoes: TArray<string>;
  end;

/// <summary>Plans the deletion of AName declared on ADeclLine0 of ALines
///  (the declaring unit). The line may also be the IMPLEMENTATION header of
///  a method or routine - the declaration is then located. False with
///  AWhy when the line does not declare AName or the shape is not
///  supported; vetoes do NOT make it fail (ASym.Vetoes lists them).</summary>
function PlanSafeDeleteSymbol(const ALines: TArray<string>; ADeclLine0: Integer;
  const AName: string; out ASym: TSafeDeleteSymbol; out AWhy: string): Boolean;

/// <summary>ALine declares AName (routine header, property, "A, B: T",
///  "X = ..."). Used when DelphiLSP gives no definition at the caret.</summary>
function LineDeclaresName(const ALine, AName: string): Boolean;

/// <summary>DelphiLSP's definition for a caret that sits ON a declaration
///  names ANOTHER symbol: the caret line declares AName itself, and the
///  answer lies in another file. Measured (forum 2026-09-22): for the field
///  "ABC: Integer;" of a class in a unit that uses Winapi.Windows, DelphiLSP
///  answers Winapi.Windows.pas:20496 - the global type ABC - and rename
///  refused "declared in the RAD Studio installation". A declaration and
///  its implementation live in ONE unit, so a cross-file answer cannot be
///  its partner. Not foreign: an override and a typeless property
///  redeclaration ("property Caption;") - those really belong to the
///  ancestor.</summary>
function DeclarationAnswerIsForeign(const ACaretLine, AName, ACaretFile,
  AAnswerFile: string): Boolean;

/// <summary>True when a scan has NO anchor for the symbol: DelphiLSP did not
///  answer the declaration query AND the caret line does not declare the
///  name, so the caret is a USE whose declaration is unknown.
///  WHY IT MATTERS (forum 2026-09-30, with both logs): on a cold session the
///  declaration query came back empty for a use of GlobalConfig.Formulare.BTB,
///  the caret was taken as the declaration, and every one of the 341
///  candidates that DelphiLSP later resolved CORRECTLY to the real
///  declaration was dropped as "leads to another symbol" - 29 hits instead of
///  327, two of them wrong. The same run two minutes later, session warm, was
///  right. An unknown anchor must SUPPRESS the filtering, never guess an
///  anchor.</summary>
function DeclarationAnchorUnknown(AHasDefinition: Boolean;
  const ACaretLine, AName: string): Boolean;

/// <summary>The line range of the routine HEADER that ALINE0 belongs to, when
///  that header declares ANAME. AFirst is the line the header starts on - the
///  one that carries the name - and ALast the line its parameter list closes
///  on. False when no such header stands within a few lines above ALINE0.
///  WHY (forum 2026-09-30, second log): DelphiLSP answered the declaration of
///  "BTB" with ROM_Utils.pas:15589:1 while the routine starts at 15584 - a
///  CONTINUATION line of a wrapped parameter list, where the name does not
///  occur at all. The partner query is asked at the name's column on that
///  line, finds nothing, so the symbol's position set keeps that single line -
///  and every candidate the server resolves to the header's FIRST line then
///  counts as "another symbol". In that run it silently dropped the
///  declaration, the implementation and 7 calls, out of the WARM run: the one
///  that looked right.</summary>
function DeclarationHeaderSpan(const ALines: TArray<string>; ALine0: Integer;
  const AName: string; out AFirst, ALast: Integer): Boolean;

/// <summary>The position the candidates' own answers agree on, and how many
///  of them do; '' / -1 when they do not agree clearly enough (fewer than two
///  answers, or no position reaching AMinShare of them).
///  WHY: while the declaration query stays unanswered the scan has no anchor
///  and marks EVERY hit unverified - although its own answers already name the
///  declaration. In the reported cold run 326 of 341 candidates resolved to
///  ROM_Utils.pas:15589, the very position the warm run took two minutes later
///  as the declaration. That is evidence, not a guess - the same "the most
///  frequent target IS the symbol" rule semantic replace already uses.</summary>
function DominantAnswer(const AFiles: TArray<string>; const ALines: TArray<Integer>;
  out AFile: string; out ALine: Integer; AMinShare: Double = 0.6): Integer;

type
  /// <summary>What to do with an occurrence whose DelphiLSP answer names a
  ///  position that is NOT one of the symbol's.</summary>
  TForeignAnswerVerdict = (
    favDrop,        // it belongs to another symbol - remove the row
    favKeepMarked); // say where the answer led and let the user judge

/// <summary>THE one rule for "the server answered, and it points elsewhere",
///  whichever pass asked. A CLEAR answer is evidence: with an anchor to
///  compare against and a session that aborted nothing, such a row belongs to
///  another symbol and is dropped.
///  It is kept and MARKED when there is no anchor (nothing to compare against)
///  or when the session aborted a request (issue #13 - then the answer comes
///  from exactly the degraded state we do not trust, and removing a real
///  reference is the worse error).
///  WHY THIS IS ONE FUNCTION: the third cold-session report (forum
///  2026-09-30) had the SAME finding judged twice in one run - the
///  derived-anchor pass KEPT three rows resolving to UGlobalRomConfig.pas:1586
///  while the second attempt DROPPED two more of exactly that shape, so the
///  cold run listed 339 rows where the warm one lists 336. A verdict must not
///  depend on which pass happened to see the answer.</summary>
function ForeignAnswerVerdict(AHaveAnchor: Boolean;
  AAbortedRequests: Integer): TForeignAnswerVerdict;

/// <summary>ALines after AEdits (applied bottom-up; overlapping edits are
///  not expected - the planner never produces them).</summary>
function ApplySafeDeleteEdits(const ALines: TArray<string>;
  const AEdits: TArray<TSafeDeleteEdit>): TArray<string>;

/// <summary>0-based lines of a TEXT form file (.dfm/.fmx) that mention
///  AName as a whole word - component names, event handler bindings,
///  references. Deliberately text-level: a form binds by NAME.</summary>
function FormTextMentions(const AText, AName: string): TArray<Integer>;

/// <summary>Words after the signature's ';' on the header's last line
///  AHdrEnd (+ the following lines that only hold directives):
///  ' virtual; overload;' - test with a whole-word search.</summary>
function HeaderDirectives(const ALines: TArray<string>; AStart, AHdrEnd: Integer): string;

/// <summary>Does the member AMEMBER of type ATYPE, declared in ACONTENT,
///  PROVE that it is a different symbol than the one being renamed?
///  Evidence only, and only two kinds of it: 'reintroduce' (the member
///  deliberately HIDES the inherited one, so it is a symbol of its own)
///  and an 'overload' whose parameter count differs from ADECLPARAMS
///  (-1 = unknown, then only reintroduce counts). Anything unclear
///  answers False and stays part of the rename - one occurrence too many
///  is visible in the preview, while skipping a real one does not
///  compile (audit #40, M21b).</summary>
function MemberIsOtherSymbol(const AContent, AType, AMember: string;
  ADeclParams: Integer; out AReason: string): Boolean;

/// <summary>The number of parameters of a routine declaration or header
///  line (its parameter list may wrap over the following lines), -1 when
///  the line has no parameter list at all.</summary>
function DeclaredParamCount(const ALines: TArray<string>; ALine0: Integer): Integer;

/// <summary>Human-readable kind ("method", "field", ...).</summary>
function SafeDeleteKindText(AKind: TSafeDeleteKind): string;

/// <summary>The routine header at ALine0 declares AName as one of its
///  PARAMETERS (its own name is not one). DelphiLSP answers nothing at a
///  parameter in a declaration, so a rename started there had no anchor
///  at all and was refused (issue #49).</summary>
function HeaderDeclaresParameter(const ALines: TArray<string>;
  ALine0: Integer; const AName: string): Boolean;

type
  /// <summary>What could be established about the two headers a parameter
  ///  is declared in. Delphi requires the two names to be identical, so
  ///  this is not a matter of degree: either both places are known, or the
  ///  routine really has one header, or the rename must not happen.</summary>
  TParamPairing = (
    ppUnknown,        // could not be established - refuse
    ppPaired,         // both headers are known
    ppSingleHeader);  // the routine has exactly one header

  /// <summary>Everything the rename of a PARAMETER needs. POSITIONS are the
  ///  places at which the name is really declared (0-based, X = column,
  ///  Y = line), sorted; SPANS are the header spans they live in (Y = first
  ///  line, X = last). The spans are what makes a line-based target safe: on
  ///  a header line only the positions are the parameter - the routine's own
  ///  name, a type or a default value spelled like it are not (issue
  ///  #50).</summary>
  TParamRenamePlan = record
    Pairing: TParamPairing;
    Positions: TArray<TPoint>;
    Spans: TArray<TPoint>;
    Reason: string;
  end;

/// <summary>The plan for renaming the parameter ANAME of the routine whose
///  header ALINE0 belongs to. Parsed on the MASKED text over the header's
///  own span, so a comment in the parameter list, a header longer than
///  twelve lines and "const [ref]" are read like any other header; a shape
///  that cannot be read at all (an operator header, a list that is never
///  closed) answers ppUnknown WITH the reason, because renaming one of two
///  names Delphi requires to be identical is a broken unit, not a smaller
///  change (issue #50).</summary>
function PlanParameterRename(const ALines: TArray<string>; ALine0: Integer;
  const AName: string): TParamRenamePlan;

/// <summary>The line in the routine's OTHER header on which the SAME
///  parameter stands - the projection of PlanParameterRename that older
///  callers and its own tests use. A parameter is declared TWICE - in the routine's
///  declaration and in its implementation header - and Delphi requires the
///  two names to be identical: MEASURED with dcc32 37.0 for a class method
///  and for a free routine, they differ and the unit does not compile
///  (E2037 "Declaration of 'X' differs from previous declaration", plus an
///  undeclared identifier for every use in the body). DelphiLSP answers
///  nothing at either of the two, so nothing a rename verifies against can
///  pair them and the second one was left behind (issue #49).
///  ALine0 is one of the two header lines and AName one of its parameters.
///  The answer is 0-based and need not be the partner's FIRST line - a
///  wrapped parameter list puts the name further down.
///  False when there is no such partner, which includes an overloaded
///  routine whose declarations carry another name at that place, and a free
///  routine whose only other header is a 'forward' one in the
///  implementation section (not searched - missing it leaves the occurrence
///  listed as unverified, which is the safe direction).</summary>
function FindParameterPartnerLine(const ALines: TArray<string>;
  ALine0: Integer; const AName: string; out APartnerLine0: Integer): Boolean;

implementation

uses
  System.Classes, System.StrUtils, System.Math, System.Generics.Collections,
  Expert.AutoImport, Expert.UnitIndex, Expert.PascalScanner,
  Expert.InterfaceLinks, Expert.SignatureEdit;

// ---- a parameter is declared TWICE, and only the names count ------------

type
  // One routine header, read over its OWN span on the MASKED text: a
  // comment inside the parameter list is then not code, a header of twenty
  // lines is still one header, and the positions are the real ones.
  THeaderSpan = record
    First, Last: Integer;          // 0-based lines; Last carries the ';'
    EndCol: Integer;               // 1-based column of that ';'
    HasList: Boolean;
    Qualified: string;             // 'Foo' or 'TFoo.Foo'
    Params: string;                // the parameter list as text
    Names: TArray<TPoint>;         // where the parameter NAMES stand
    NameTexts: TArray<string>;     // ...and what they are, same order
  end;

const
  // A header longer than this is pathological; the answer is then "could
  // not be read", which refuses a rename instead of half-doing it.
  MaxHeaderSpanLines = 60;

function IsParamModifierWord(const AWord: string): Boolean;
begin
  Result := SameText(AWord, 'const') or SameText(AWord, 'var')
    or SameText(AWord, 'out') or SameText(AWord, 'constref');
end;

// ALine0 must BE a header line of AMASKED. Walks its span and collects the
// positions at which parameter NAMES are declared: at bracket depth 1, in
// the part of a group before its ':'. So never the routine's own name (it
// stands before the '('), never a type, never a default value, and never
// the parameters of a procedural parameter's own list (depth 2).
function ReadHeaderSpanAt(const AMasked: TArray<string>; ALine0: Integer;
  out ASpan: THeaderSpan): Boolean;
var
  Kind, Joined, Q, P, Rt: string;
  IsCM: Boolean;
  Depth, Angle, I, J, K: Integer;
  Started, InNames: Boolean;
begin
  Result := False;
  ASpan := Default(THeaderSpan);
  if (ALine0 < 0) or (ALine0 > High(AMasked)) then Exit;
  if not IsHeaderLine(Trim(AMasked[ALine0]), Kind, IsCM) then Exit;
  ASpan.First := ALine0;
  ASpan.Last := -1;
  Depth := 0;
  Angle := 0;
  Started := False;
  InNames := False;
  for I := ALine0 to Min(ALine0 + MaxHeaderSpanLines, High(AMasked)) do
  begin
    var S := AMasked[I];
    J := 1;
    while (J <= Length(S)) and (ASpan.Last < 0) do
    begin
      var C := S[J];
      if IsIdentStart(C) then
      begin
        K := J;
        while (K <= Length(S)) and IsIdentChar(S[K]) do Inc(K);
        if InNames and (Depth = 1) and (Angle = 0)
          and not IsParamModifierWord(Copy(S, J, K - J)) then
        begin
          ASpan.Names := ASpan.Names + [Point(J - 1, I)];
          ASpan.NameTexts := ASpan.NameTexts + [Copy(S, J, K - J)];
        end;
        J := K;
        Continue;
      end;
      case C of
        '(', '[':
          begin
            Inc(Depth);
            if (C = '(') and (Depth = 1) and not Started then
            begin
              Started := True;
              ASpan.HasList := True;
              InNames := True;
            end;
          end;
        ')', ']':
          begin
            if Depth > 0 then Dec(Depth);
            if Depth = 0 then InNames := False;
          end;
        '<': Inc(Angle);
        '>': if Angle > 0 then Dec(Angle);
        ';':
          if Depth = 0 then
          begin
            ASpan.Last := I;
            ASpan.EndCol := J;
          end
          else if (Depth = 1) and (Angle = 0) then InNames := True;
        ',': if (Depth = 1) and (Angle = 0) then InNames := True;
        ':': if (Depth = 1) and (Angle = 0) then InNames := False;
      end;
      Inc(J);
    end;
    if ASpan.Last >= 0 then Break;
  end;
  if ASpan.Last < 0 then Exit;        // never closed within the cap
  // ...and the join stops AT it: the first ';' of the text is the one
  // INSIDE the parameter list as soon as there is more than one group.
  Joined := '';
  for I := ASpan.First to ASpan.Last - 1 do
    Joined := Joined + ' ' + AMasked[I];
  Joined := Joined + ' ' + Copy(AMasked[ASpan.Last], 1, ASpan.EndCol);
  if not ParseHeader(Trim(Joined), Kind, Q, P, Rt) then Exit;
  ASpan.Qualified := Q;
  ASpan.Params := P;
  Result := True;
end;

// The line the routine header covering ALine0 STARTS on - ALine0 itself
// when it is one. A wrapped parameter list is answered by DelphiLSP with
// whatever line its range starts on, and that need not be the first (1.15.4
// measured a declaration reported at a continuation line).
function HeaderLineOfMasked(const AMasked: TArray<string>; ALine0: Integer): Integer;
var
  Kind: string;
  IsCM: Boolean;
  Span: THeaderSpan;
begin
  Result := -1;
  if (ALine0 < 0) or (ALine0 > High(AMasked)) then Exit;
  for var I := ALine0 downto Max(0, ALine0 - MaxHeaderSpanLines) do
  begin
    if not IsHeaderLine(Trim(AMasked[I]), Kind, IsCM) then Continue;
    // the nearest header above either still covers ALine0 or ALine0 belongs
    // to something else entirely
    if ReadHeaderSpanAt(AMasked, I, Span) and (Span.Last >= ALine0) then
      Result := I;
    Exit;
  end;
end;

function HeaderLineOf(const ALines: TArray<string>; ALine0: Integer): Integer;
begin
  Result := HeaderLineOfMasked(MaskCommentsAndStrings(ALines), ALine0);
end;

// The parameter names of the routine header at ALine0, in order, and its
// (possibly qualified) name. "A, B: Integer" is two names, which is what
// makes the INDEX the identity of a parameter.
function RoutineParameterNamesOf(const ALines: TArray<string>; ALine0: Integer;
  out AQualified: string): TArray<string>;
var
  Masked: TArray<string>;
  Span: THeaderSpan;
begin
  Result := nil;
  AQualified := '';
  Masked := MaskCommentsAndStrings(ALines);
  ALine0 := HeaderLineOfMasked(Masked, ALine0);
  if ALine0 < 0 then Exit;
  if not ReadHeaderSpanAt(Masked, ALine0, Span) then Exit;
  AQualified := Span.Qualified;
  Result := Span.NameTexts;
end;

function HeaderDeclaresParameter(const ALines: TArray<string>;
  ALine0: Integer; const AName: string): Boolean;
var
  Q: string;
begin
  Result := False;
  if AName = '' then Exit;
  for var N in RoutineParameterNamesOf(ALines, ALine0, Q) do
    if SameText(N, AName) then Exit(True);
end;

// Same names, modifiers and types, in the same order - DEFAULTS ignored,
// because they live in the declaration only. That is what tells two
// overloads apart which share a parameter name at the same index, and it
// is why a pair that does not match is no pair at all.
function ParamListsMatch(const A, B: string): Boolean;
var
  PA, PB: TArray<TSigParam>;
begin
  PA := ParseParamList(A);
  PB := ParseParamList(B);
  Result := Length(PA) = Length(PB);
  if not Result then Exit;
  for var I := 0 to High(PA) do
    if not (SameText(Trim(PA[I].Name), Trim(PB[I].Name))
        and SameText(Trim(PA[I].Modifier), Trim(PB[I].Modifier))
        and SameText(Trim(PA[I].TypeText), Trim(PB[I].TypeText))) then
      Exit(False);
end;

// The type body a MEMBER DECLARATION at ALine0 belongs to: the nearest
// opener above it, with its KIND. Deliberately not EnclosingContainerName /
// ClassOpenerName, which accept 'class' and 'record' only - an INTERFACE
// method has no implementation ANYWHERE, and that is not the same as
// failing to find one (issue #50). A member's own type is always the
// nearest opener above it: anything between would have to be nested, and
// then the nested one IS the answer.
function NearestTypeOpenerAbove(const ALines: TArray<string>; ALine0: Integer;
  out AName, AKeyword: string): Boolean;
begin
  Result := False;
  AName := '';
  AKeyword := '';
  for var I := Min(ALine0, High(ALines)) downto 0 do
  begin
    var T := Trim(StripLineComment(ALines[I]));
    var Eq := Pos('=', T);
    if Eq <= 1 then Continue;
    var N := Trim(Copy(T, 1, Eq - 1));
    if (N = '') or not IsIdentStart(N[1]) then Continue;
    var Rest := TrimLeft(Copy(T, Eq + 1, MaxInt));
    if StartsText('packed ', Rest) then Rest := TrimLeft(Copy(Rest, 8, MaxInt));
    var K := 1;
    while (K <= Length(Rest)) and IsIdentChar(Rest[K]) do Inc(K);
    var Kw := LowerCase(Copy(Rest, 1, K - 1));
    if (Kw <> 'class') and (Kw <> 'interface') and (Kw <> 'dispinterface')
      and (Kw <> 'record') and (Kw <> 'object') then Continue;
    if Rest.EndsWith(';') then Continue;    // a forward declaration or an alias
    var Lt := Pos('<', N);
    if Lt > 0 then N := Copy(N, 1, Lt - 1);
    AName := N;
    AKeyword := Kw;
    Exit(True);
  end;
end;

// Two routine names are the same when they are the same apart from their
// generic arguments: an implementation header writes "TList<T>.Add" where
// the declaration says "Add" inside "TList<T>".
function SameRoutineName(const A, B: string): Boolean;

  function Strip(const S: string): string;
  begin
    Result := '';
    var Angle := 0;
    for var I := 1 to Length(S) do
      case S[I] of
        '<': Inc(Angle);
        '>': if Angle > 0 then Dec(Angle);
      else
        if Angle = 0 then Result := Result + S[I];
      end;
  end;

begin
  Result := SameText(Strip(A), Strip(B));
end;

// Every header at column 1 named ANAME: a top-level routine's interface
// declaration, a 'forward' one and its implementation all qualify.
function TopLevelHeadersNamed(const AMasked: TArray<string>;
  const AName: string): TArray<Integer>;
var
  Span: THeaderSpan;
begin
  Result := nil;
  for var I := 0 to High(AMasked) do
  begin
    var Raw := AMasked[I];
    if (Raw = '') or CharInSet(Raw[1], [' ', #9]) then Continue;
    if not ReadHeaderSpanAt(AMasked, I, Span) then Continue;
    if SameRoutineName(Span.Qualified, AName) then Result := Result + [I];
  end;
end;

// The headers named ANAME inside the ONE top-level block ALine0 sits in -
// that is where a nested routine and its own 'forward' declaration live,
// and nowhere else. A nested routine of the same name in ANOTHER routine
// is a different symbol and must never be paired with this one.
function NestedHeadersNamed(const AMasked: TArray<string>; ALine0: Integer;
  const AName: string): TArray<Integer>;
var
  Span: THeaderSpan;
  Outer: Integer;
begin
  Result := nil;
  Outer := -1;
  for var I := ALine0 downto 0 do
    if (AMasked[I] <> '') and not CharInSet(AMasked[I][1], [' ', #9]) then
    begin
      Outer := I;
      Break;
    end;
  if Outer < 0 then Exit;
  for var I := Outer + 1 to High(AMasked) do
  begin
    var Raw := AMasked[I];
    if (Raw <> '') and not CharInSet(Raw[1], [' ', #9]) then Break;
    if not ReadHeaderSpanAt(AMasked, I, Span) then Continue;
    if SameRoutineName(Span.Qualified, AName) then Result := Result + [I];
  end;
end;

function PlanParameterRename(const ALines: TArray<string>; ALine0: Integer;
  const AName: string): TParamRenamePlan;
var
  Masked: TArray<string>;
  Span, Other: THeaderSpan;
  Cands, Matches: TArray<Integer>;
  Plan: TParamRenamePlan;

  procedure Take(const ASpan: THeaderSpan);
  begin
    for var I := 0 to High(ASpan.NameTexts) do
      if SameText(ASpan.NameTexts[I], AName) then
        Plan.Positions := Plan.Positions + [ASpan.Names[I]];
    Plan.Spans := Plan.Spans + [Point(ASpan.Last, ASpan.First)];
  end;

begin
  Plan := Default(TParamRenamePlan);
  Plan.Pairing := ppUnknown;
  Result := Plan;
  if (AName = '') or (Length(ALines) = 0) then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  var H := HeaderLineOfMasked(Masked, ALine0);
  if H < 0 then
  begin
    Plan.Reason := 'that line does not belong to a routine header';
    Result := Plan;
    Exit;
  end;
  if not ReadHeaderSpanAt(Masked, H, Span) then
  begin
    Plan.Reason := 'this routine header cannot be read (an operator header, '
      + 'or a parameter list that is never closed), so the other place the '
      + 'parameter is declared in cannot be found either';
    Result := Plan;
    Exit;
  end;
  Take(Span);
  if Length(Plan.Positions) = 0 then
  begin
    Plan.Reason := Format('"%s" is not a parameter of %s', [AName, Span.Qualified]);
    Plan.Spans := nil;
    Result := Plan;
    Exit;
  end;

  // WHICH of the two headers is this one? Not by section - a class may be
  // declared in the implementation section, and a nested type anywhere
  // (issue #50) - but by what encloses it.
  Cands := nil;
  var Dot := LastDelimiter('.', Span.Qualified);
  if Dot > 0 then
  begin
    // a method's IMPLEMENTATION header -> its declaration in the class
    var Owner := OwnerTypeOfImplHeader(ALines[H]);
    var Member := Copy(Span.Qualified, Dot + 1, MaxInt);
    var Lt := Pos('<', Member);
    if Lt > 0 then Member := Copy(Member, 1, Lt - 1);
    if (Owner <> '') and (Member <> '') then
      Cands := FindMemberDeclarationLines(string.Join(#10, ALines), Owner, Member);
  end
  else if not ((ALines[H] <> '') and CharInSet(ALines[H][1], [' ', #9])) then
    // a TOP-LEVEL routine: its interface declaration, a 'forward' one and
    // its implementation all carry the bare name at column 1
    Cands := TopLevelHeadersNamed(Masked, Span.Qualified)
  else
  begin
    // INDENTED and unqualified: either nested in a routine body or a member
    // declaration. What the nearest unindented line above is decides, and
    // nothing else - the section does not (a class may be declared in the
    // implementation section, a type may be nested).
    var Outer := -1;
    for var I := H - 1 downto 0 do
      if (ALines[I] <> '') and not CharInSet(ALines[I][1], [' ', #9]) then
      begin
        Outer := I;
        Break;
      end;
    var OKind: string;
    var OIsCM: Boolean;
    if (Outer >= 0) and IsHeaderLine(Trim(Masked[Outer]), OKind, OIsCM) then
      // nested in that routine's body: its own 'forward' declaration is the
      // only thing that may carry the name, and nothing outside the block
      Cands := NestedHeadersNamed(Masked, H, Span.Qualified)
    else
    begin
      var TName, TKw: string;
      if not NearestTypeOpenerAbove(ALines, H, TName, TKw) then
      begin
        Plan.Reason := 'the type this declaration belongs to could not be '
          + 'identified, so the implementation header that carries the same '
          + 'parameter is unknown';
        Result := Plan;
        Exit;
      end;
      var Dirs := HeaderDirectives(ALines, Span.First, Span.Last);
      if (TKw = 'interface') or (TKw = 'dispinterface')
        or HasWholeWordCI(Dirs, 'abstract') or HasWholeWordCI(Dirs, 'external') then
      begin
        // declared once BY DESIGN - there is no implementation anywhere
        Plan.Pairing := ppSingleHeader;
        Result := Plan;
        Exit;
      end;
      // EVERY implementation header of that qualified name is a candidate:
      // with overloads the first one need not be ours (issue #50)
      Cands := TopLevelHeadersNamed(Masked, TName + '.' + Span.Qualified);
      if Length(Cands) = 0 then
      begin
        Plan.Reason := Format('no implementation of %s.%s was found in this '
          + 'unit, so the second place the parameter is declared in is unknown',
          [TName, Span.Qualified]);
        Result := Plan;
        Exit;
      end;
    end;
  end;

  // our own header is not a candidate for being the other one
  var Rest: TArray<Integer> := nil;
  for var P0 in Cands do
    if P0 <> Span.First then Rest := Rest + [P0];
  Cands := Rest;

  Matches := nil;
  for var P0 in Cands do
  begin
    if not ReadHeaderSpanAt(Masked, P0, Other) then Continue;
    if not ParamListsMatch(Span.Params, Other.Params) then Continue;
    Matches := Matches + [P0];
  end;

  if Length(Matches) = 1 then
  begin
    ReadHeaderSpanAt(Masked, Matches[0], Other);
    Take(Other);
    Plan.Pairing := ppPaired;
    // sorted, so a caller can read the declaration and the body off the
    // ends (two or three entries - no comparer needed for that)
    for var I := 1 to High(Plan.Positions) do
      for var J := High(Plan.Positions) downto I do
        if (Plan.Positions[J].Y < Plan.Positions[J - 1].Y)
          or ((Plan.Positions[J].Y = Plan.Positions[J - 1].Y)
              and (Plan.Positions[J].X < Plan.Positions[J - 1].X)) then
        begin
          var Tmp := Plan.Positions[J];
          Plan.Positions[J] := Plan.Positions[J - 1];
          Plan.Positions[J - 1] := Tmp;
        end;
  end
  else if Length(Matches) > 1 then
    Plan.Reason := Format('%d headers of %s fit this one, so which of them '
      + 'declares the same parameter cannot be decided here',
      [Length(Matches), Span.Qualified])
  else if Length(Cands) = 0 then
    // nothing else carries that name: a nested routine, or a routine that
    // is only defined in the implementation section
    Plan.Pairing := ppSingleHeader
  else
    Plan.Reason := Format('the other header of %s does not carry the same '
      + 'parameter list, so the two cannot be paired', [Span.Qualified]);
  Result := Plan;
end;

function FindParameterPartnerLine(const ALines: TArray<string>;
  ALine0: Integer; const AName: string; out APartnerLine0: Integer): Boolean;
var
  Plan: TParamRenamePlan;
begin
  APartnerLine0 := -1;
  Plan := PlanParameterRename(ALines, ALine0, AName);
  Result := False;
  if (Plan.Pairing <> ppPaired) or (Length(Plan.Spans) = 0) then Exit;
  // the anchor's own span is the first one - the partner is the position
  // outside it
  for var Pt in Plan.Positions do
    if (Pt.Y < Plan.Spans[0].Y) or (Pt.Y > Plan.Spans[0].X) then
    begin
      APartnerLine0 := Pt.Y;
      Exit(True);
    end;
end;

function DeclaredParamCount(const ALines: TArray<string>; ALine0: Integer): Integer;
var
  Text_: string;
begin
  Result := -1;
  if (ALine0 < 0) or (ALine0 > High(ALines)) then Exit;
  Text_ := JoinOpenParenLines(ALines, ALine0);
  var Op := Pos('(', Text_);
  if Op = 0 then Exit;
  var Depth := 0;
  var Cl := 0;
  for var I := Op to Length(Text_) do
  begin
    if Text_[I] = '(' then Inc(Depth)
    else if Text_[I] = ')' then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        Cl := I;
        Break;
      end;
    end;
  end;
  if Cl = 0 then Exit;
  Result := Length(ParseParamList(Copy(Text_, Op + 1, Cl - Op - 1)));
end;

function MemberIsOtherSymbol(const AContent, AType, AMember: string;
  ADeclParams: Integer; out AReason: string): Boolean;
var
  Lines: TArray<string>;
begin
  Result := False;
  AReason := '';
  if (AType = '') or (AMember = '') then Exit;
  var DeclLines := FindMemberDeclarationLines(AContent, AType, AMember);
  // No declaration found, or several: the type has its own overload set
  // and which one an implementation belongs to cannot be decided here.
  if Length(DeclLines) <> 1 then Exit;
  Lines := SplitContentLines(AContent);
  var L := DeclLines[0];
  if (L < 0) or (L > High(Lines)) then Exit;
  var Dirs := LowerCase(HeaderDirectives(Lines, L, L));
  if HasWholeWordCI(Dirs, 'reintroduce') then
  begin
    AReason := AType + '.' + AMember + ' is declared "reintroduce" - it ' +
      'hides the inherited member and is a symbol of its own';
    Exit(True);
  end;
  if (ADeclParams >= 0) and HasWholeWordCI(Dirs, 'overload') then
  begin
    var N := DeclaredParamCount(Lines, L);
    if (N >= 0) and (N <> ADeclParams) then
    begin
      AReason := Format('%s.%s is an overload with %d parameter(s) while ' +
        'the renamed declaration has %d', [AType, AMember, N, ADeclParams]);
      Exit(True);
    end;
  end;
end;

function SafeDeleteKindText(AKind: TSafeDeleteKind): string;
begin
  case AKind of
    sdkMethod: Result := 'method';
    sdkRoutine: Result := 'routine';
    sdkField: Result := 'field';
    sdkProperty: Result := 'property';
    sdkVariable: Result := 'variable';
    sdkConstant: Result := 'constant';
    sdkType: Result := 'type';
  else
    Result := 'symbol';
  end;
end;

// The line's CODE: comments ({ }, (* *), //) and string contents blanked
// (MaskCommentsAndStrings keeps the length), trimmed. Analysis only - edits
// always use the raw line.
function Code(const S: string): string;
begin
  Result := Trim(MaskCommentsAndStrings([S])[0]);
end;

function FirstWordU(const S: string): string;
var
  T: string;
  I: Integer;
begin
  T := TrimLeft(S);
  I := 1;
  while (I <= Length(T)) and IsIdentChar(T[I]) do Inc(I);
  Result := UpperCase(Copy(T, 1, I - 1));
end;

// Header line of the class / record / interface / object whose BODY
// contains ALine0 (-1 = none), with its name.
function EnclosingTypeHeader(const ALines: TArray<string>; ALine0: Integer;
  out AName: string; out AIsInterface: Boolean): Integer;
var
  Depth: Integer;
  T, U, Rest: string;
  P: Integer;
begin
  Result := -1;
  AName := '';
  AIsInterface := False;
  Depth := 0;
  for var L := ALine0 - 1 downto 0 do
  begin
    T := Code(ALines[L]);
    U := UpperCase(T);
    if (U = 'END;') or (U = 'END') then
    begin
      Inc(Depth);
      Continue;
    end;
    if (U = 'IMPLEMENTATION') or (U = 'INTERFACE') then Exit;
    P := Pos('=', T);
    if P > 1 then
    begin
      Rest := UpperCase(Trim(Copy(T, P + 1, MaxInt)));
      if Rest.StartsWith('PACKED ') then Rest := TrimLeft(Copy(Rest, 8, MaxInt));
      var Opens := (Rest.StartsWith('CLASS') and not Rest.StartsWith('CLASS OF'))
        or Rest.StartsWith('RECORD') or Rest.StartsWith('OBJECT')
        or Rest.StartsWith('INTERFACE') or Rest.StartsWith('DISPINTERFACE');
      if Opens and not Rest.EndsWith(';') then   // not a forward / alias
      begin
        if Depth = 0 then
        begin
          AName := Trim(Copy(T, 1, P - 1));
          var LT := Pos('<', AName);
          if LT > 0 then AName := Trim(Copy(AName, 1, LT - 1));
          AIsInterface := Rest.StartsWith('INTERFACE') or Rest.StartsWith('DISPINTERFACE');
          Exit(L);
        end;
        Dec(Depth);
      end;
    end;
  end;
end;

// Parent list of a type header line "TFoo = class(TBase, IA, IB)".
function HeaderParents(const ALine: string): TArray<string>;
var
  A, B: Integer;
begin
  Result := nil;
  A := Pos('(', ALine);
  B := Pos(')', ALine);
  if (A = 0) or (B <= A) then Exit;
  for var S in Copy(ALine, A + 1, B - A - 1).Split([',']) do
    if Trim(S) <> '' then Result := Result + [Trim(S)];
end;

// Implementation header of AQualified ("TFoo.Bar" or a plain name for a
// top-level routine), -1 = none. Nested routines (indented) never match a
// plain name.
function FindImplHeader(const ALines: TArray<string>; const AQualified: string): Integer;
var
  Kind, Hdr, Q, P, R: string;
  IsCM: Boolean;
  HdrEnd, Impl: Integer;
  Plain: Boolean;
begin
  Result := -1;
  Impl := ImplementationLineOf(ALines);
  if Impl = MaxInt then Exit;
  Plain := Pos('.', AQualified) = 0;
  for var L := Impl + 1 to High(ALines) do
  begin
    if not IsHeaderLine(StripLineComment(ALines[L]), Kind, IsCM) then Continue;
    if Plain and (ALines[L] <> TrimLeft(ALines[L])) then Continue;   // nested
    Hdr := CollectHeader(ALines, L, HdrEnd);
    if (Hdr = '') or not ParseHeader(Hdr, Kind, Q, P, R) then Continue;
    // "TOuter.TInner.Bar" also matches a container given as "TInner"
    var LT := Pos('<', Q);
    if LT > 0 then   // TList<T>.Add -> TList.Add
      Q := Copy(Q, 1, LT - 1) + Copy(Q, Pos('>', Q) + 1, MaxInt);
    if SameText(Q, AQualified) or (not Plain and Q.ToUpper.EndsWith('.' + AQualified.ToUpper)) then
      Exit(L);
  end;
end;

// Words after the signature's ';' on the header's last line (+ the lines
// that only hold directives): 'virtual', 'override', ...
function HeaderDirectives(const ALines: TArray<string>; AStart, AHdrEnd: Integer): string;
var
  S: string;
  Depth, P: Integer;
begin
  S := StripLineComment(ALines[AHdrEnd]);
  Depth := 0;
  P := 0;
  var From := 1;
  if AHdrEnd <> AStart then From := 1;
  for var I := From to Length(S) do
    case S[I] of
      '(', '[': Inc(Depth);
      ')', ']': Dec(Depth);
      ';': if Depth = 0 then begin P := I; Break; end;
    end;
  Result := ' ' + Copy(S, P + 1, MaxInt);
  // directive-only follow-up lines ("      override;")
  for var L := AHdrEnd + 1 to High(ALines) do
  begin
    var T := Code(ALines[L]);
    var W := FirstWordU(T);
    if (W = 'VIRTUAL') or (W = 'OVERRIDE') or (W = 'OVERLOAD') or (W = 'ABSTRACT') or
       (W = 'DYNAMIC') or (W = 'REINTRODUCE') or (W = 'STDCALL') or (W = 'CDECL') or
       (W = 'INLINE') or (W = 'MESSAGE') or (W = 'STATIC') or (W = 'FINAL') then
      Result := Result + ' ' + T
    else
      Break;
  end;
end;

// Visibility keyword in force at ALine0 inside the type body starting at
// AHeader ('' = the default section at the top).
function VisibilityAt(const ALines: TArray<string>; AHeader, ALine0: Integer): string;
var
  Depth: Integer;
begin
  Result := '';
  Depth := 0;
  for var L := AHeader + 1 to ALine0 - 1 do
  begin
    var T := Code(ALines[L]);
    var U := UpperCase(T);
    if (ClassOpenerName(T) <> '') or ((Pos('= RECORD', U) > 0) and not U.EndsWith(';')) then
      Inc(Depth)
    else if (U = 'END;') or (U = 'END') then
      Dec(Depth)
    else if Depth = 0 then
    begin
      if U.StartsWith('STRICT ') then U := TrimLeft(Copy(U, 8, MaxInt));
      if (U = 'PRIVATE') or (U = 'PROTECTED') or (U = 'PUBLIC') or (U = 'PUBLISHED') or
         (U = 'AUTOMATED') then
        Result := LowerCase(U);
    end;
  end;
end;

// The section keyword ('var', 'const', 'type', ...) that governs a unit-
// or routine-level declaration on ALine0 ('' = unknown).
function SectionAt(const ALines: TArray<string>; ALine0: Integer): string;
begin
  Result := '';
  for var L := ALine0 downto 0 do
  begin
    var W := FirstWordU(Code(ALines[L]));
    if (W = 'VAR') or (W = 'CONST') or (W = 'TYPE') or (W = 'THREADVAR') or
       (W = 'RESOURCESTRING') then
      Exit(LowerCase(W));
    if (W = 'BEGIN') or (W = 'IMPLEMENTATION') or (W = 'INTERFACE') or
       (W = 'PROCEDURE') or (W = 'FUNCTION') or (W = 'CONSTRUCTOR') or
       (W = 'DESTRUCTOR') then
      Exit;
  end;
end;

procedure AddEdit(var ASym: TSafeDeleteSymbol; AFirst, ALast: Integer;
  const ALines: TArray<string>);
var
  E: TSafeDeleteEdit;
begin
  // XML doc comments directly above belong to the declaration - one rule,
  // in the unit the quick fix can reach too (forum #225)
  AFirst := DocCommentStart(ALines, AFirst);
  // no double blank line left behind
  if (AFirst > 0) and (ALast < High(ALines)) and (Trim(ALines[AFirst - 1]) = '')
    and (Trim(ALines[ALast + 1]) = '') then
    Inc(ALast);
  E := Default(TSafeDeleteEdit);
  E.FirstLine := AFirst;
  E.LastLine := ALast;
  ASym.Edits := ASym.Edits + [E];
end;

procedure Veto(var ASym: TSafeDeleteSymbol; const AWhy: string);
begin
  ASym.Vetoes := ASym.Vetoes + [AWhy];
end;

// Names of a data declaration line: "A, B: Integer;" / "X = 5;" /
// "var X: T;" / "class var F: T;". AColon receives the position of the
// ':' or '=' separating names and the rest, ANamesStart where the names
// begin.
function SplitDeclNames(const ALine: string; out ANamesStart, ASep: Integer;
  out AIsEquals: Boolean): TArray<string>;
var
  T, U: string;
  Lead, I, Depth: Integer;
begin
  Result := nil;
  ASep := 0;
  AIsEquals := False;
  T := StripLineComment(ALine);
  Lead := 1;
  while (Lead <= Length(T)) and CharInSet(T[Lead], [' ', #9]) do Inc(Lead);
  U := UpperCase(Copy(T, Lead, MaxInt));
  for var KW in ['CLASS VAR ', 'VAR ', 'CONST ', 'THREADVAR ', 'TYPE ', 'RESOURCESTRING '] do
    if U.StartsWith(KW) then
    begin
      Inc(Lead, Length(KW));
      while (Lead <= Length(T)) and CharInSet(T[Lead], [' ', #9]) do Inc(Lead);
      Break;
    end;
  ANamesStart := Lead;
  Depth := 0;
  for I := Lead to Length(T) do
  begin
    case T[I] of
      '(', '[', '<': Inc(Depth);
      ')', ']', '>': Dec(Depth);
      ':': if (Depth = 0) and ((I = Length(T)) or (T[I + 1] <> '=')) then begin ASep := I; Break; end;
      '=': if Depth = 0 then begin ASep := I; AIsEquals := True; Break; end;
    end;
  end;
  if ASep = 0 then Exit;
  for var S in Copy(T, Lead, ASep - Lead).Split([',']) do
  begin
    var N := Trim(S);
    var LT := Pos('<', N);
    if LT > 0 then N := Trim(Copy(N, 1, LT - 1));
    if (N = '') or not IsValidIdent(N) then Exit(nil);
    Result := Result + [N];
  end;
end;

function DeclarationAnchorUnknown(AHasDefinition: Boolean;
  const ACaretLine, AName: string): Boolean;
begin
  Result := (not AHasDefinition) and not LineDeclaresName(ACaretLine, AName);
end;

function DeclarationHeaderSpan(const ALines: TArray<string>; ALine0: Integer;
  const AName: string; out AFirst, ALast: Integer): Boolean;
const
  MaxUp = 40;      // a parameter list longer than this is not a wrapped header
  MaxDown = 40;
var
  I, Depth: Integer;
  Masked: TArray<string>;
begin
  Result := False;
  AFirst := ALine0;
  ALast := ALine0;
  if (ALine0 < 0) or (ALine0 > High(ALines)) or (AName = '') then Exit;
  // comments and strings masked: a name inside one is no declaration
  Masked := MaskCommentsAndStrings(ALines);
  // UP to the line that carries the name and declares it
  I := ALine0;
  while (I >= 0) and (ALine0 - I <= MaxUp) do
  begin
    if HasWholeWordCI(Masked[I], AName) and LineDeclaresName(Masked[I], AName) then
    begin
      AFirst := I;
      Result := True;
      Break;
    end;
    Dec(I);
  end;
  if not Result then Exit;
  // DOWN while the parameter list is still open - that is the header's extent
  Depth := 0;
  ALast := AFirst;
  for I := AFirst to Min(High(Masked), AFirst + MaxDown) do
  begin
    for var C in Masked[I] do
      if C = '(' then Inc(Depth)
      else if C = ')' then Dec(Depth);
    ALast := I;
    if (Depth <= 0) and (I > AFirst) then Break;
    if (Depth <= 0) and (Pos(';', Masked[I]) > 0) then Break;
  end;
  if ALast < ALine0 then ALast := ALine0;   // the answer itself always belongs
end;

function ForeignAnswerVerdict(AHaveAnchor: Boolean;
  AAbortedRequests: Integer): TForeignAnswerVerdict;
begin
  if AHaveAnchor and (AAbortedRequests <= 0) then
    Result := favDrop
  else
    Result := favKeepMarked;
end;

function DominantAnswer(const AFiles: TArray<string>; const ALines: TArray<Integer>;
  out AFile: string; out ALine: Integer; AMinShare: Double): Integer;
var
  Counts: TDictionary<string, Integer>;
  Total, Best: Integer;
  BestKey: string;
begin
  Result := 0;
  AFile := '';
  ALine := -1;
  Total := Min(Length(AFiles), Length(ALines));
  if Total < 2 then Exit;
  Counts := TDictionary<string, Integer>.Create;
  try
    Best := 0;
    BestKey := '';
    for var I := 0 to Total - 1 do
    begin
      if AFiles[I] = '' then Continue;
      var K := UpperCase(AFiles[I]) + '|' + IntToStr(ALines[I]);
      var N := 0;
      Counts.TryGetValue(K, N);
      Inc(N);
      Counts.AddOrSetValue(K, N);
      if N > Best then
      begin
        Best := N;
        BestKey := K;
        AFile := AFiles[I];
        ALine := ALines[I];
      end;
    end;
    // a clear majority of the ANSWERED ones - a handful of scattered answers
    // must not be promoted to "the declaration"
    var Answered := 0;
    for var I := 0 to Total - 1 do
      if AFiles[I] <> '' then Inc(Answered);
    if (Best >= 2) and (Answered > 0) and (Best / Answered >= AMinShare) then
      Result := Best
    else
    begin
      AFile := '';
      ALine := -1;
    end;
  finally
    Counts.Free;
  end;
end;

function DeclarationAnswerIsForeign(const ACaretLine, AName, ACaretFile,
  AAnswerFile: string): Boolean;
begin
  Result := False;
  if (AAnswerFile = '') or (ACaretFile = '') then Exit;
  if SameText(ExpandFileName(AAnswerFile), ExpandFileName(ACaretFile)) then Exit;
  if not LineDeclaresName(ACaretLine, AName) then Exit;
  var Code := LowerCase(Trim(StripLineComment(ACaretLine)));
  if HasWholeWordCI(Code, 'override') then Exit;
  if StartsText('property', Code) and (Pos(':', Code) = 0) then Exit;
  Result := True;
end;

function LineDeclaresName(const ALine, AName: string): Boolean;
var
  Kind, Hdr, Q, P, R: string;
  IsCM, IsEq: Boolean;
  NS, Sep, HdrEnd: Integer;
begin
  Result := False;
  if IsHeaderLine(StripLineComment(ALine), Kind, IsCM) then
  begin
    Hdr := CollectHeader([ALine], 0, HdrEnd);
    if Hdr = '' then Hdr := StripLineComment(ALine) + ';';
    if not ParseHeader(Hdr, Kind, Q, P, R) then Exit;
    var Dot := LastDelimiter('.', Q);
    Exit(SameText(Copy(Q, Dot + 1, MaxInt), AName));
  end;
  var T := Code(ALine);
  if T.ToUpper.StartsWith('CLASS PROPERTY ') then T := Trim(Copy(T, 7, MaxInt));
  if T.ToUpper.StartsWith('PROPERTY ') then
  begin
    var I := 10;
    while (I <= Length(T)) and CharInSet(T[I], [' ', #9]) do Inc(I);
    var S := I;
    while (I <= Length(T)) and IsIdentChar(T[I]) do Inc(I);
    Exit(SameText(Copy(T, S, I - S), AName));
  end;
  for var N in SplitDeclNames(ALine, NS, Sep, IsEq) do
    if SameText(N, AName) then Exit(True);
end;

function PlanSafeDeleteSymbol(const ALines: TArray<string>; ADeclLine0: Integer;
  const AName: string; out ASym: TSafeDeleteSymbol; out AWhy: string): Boolean;
var
  Content, Kind, Hdr, Qual, Params, Ret, T, U, Dirs, CName: string;
  IsCM, IsIntf: Boolean;
  Impl, HdrEnd, DeclLine, F, L: Integer;
begin
  Result := False;
  ASym := Default(TSafeDeleteSymbol);
  ASym.Name := AName;
  ASym.ImplLine := -1;
  ASym.ContainerLine := -1;
  ASym.LocalFirst := -1;
  ASym.LocalLast := -1;
  AWhy := '';
  if (ADeclLine0 < 0) or (ADeclLine0 > High(ALines)) then
  begin
    AWhy := 'the declaration line is outside the unit';
    Exit;
  end;
  Content := string.Join(sLineBreak, ALines);
  Impl := ImplementationLineOf(ALines);
  DeclLine := ADeclLine0;
  T := Code(ALines[DeclLine]);

  // ---- routines and methods ------------------------------------------------
  if IsHeaderLine(T, Kind, IsCM) then
  begin
    Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
    if (Hdr = '') or not ParseHeader(Hdr, Kind, Qual, Params, Ret) then
    begin
      AWhy := 'the routine header could not be read';
      Exit;
    end;
    var Dot := LastDelimiter('.', Qual);
    if not SameText(Copy(Qual, Dot + 1, MaxInt), AName) then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    if Dot > 0 then
    begin
      // an IMPLEMENTATION header "TFoo.Bar": the declaration is in the type
      var Owner := Copy(Qual, 1, Dot - 1);
      var OD := LastDelimiter('.', Owner);
      Owner := Copy(Owner, OD + 1, MaxInt);
      var LT := Pos('<', Owner);
      if LT > 0 then Owner := Copy(Owner, 1, LT - 1);
      DeclLine := FindMemberDeclarationLine(Content, Owner, AName);
      if DeclLine < 0 then
      begin
        AWhy := Format('the declaration of %s.%s was not found', [Owner, AName]);
        Exit;
      end;
      Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
      if (Hdr = '') or not IsHeaderLine(Code(ALines[DeclLine]), Kind, IsCM) then
      begin
        AWhy := 'the declaration of the method could not be read';
        Exit;
      end;
    end
    else if (DeclLine > Impl) and (ALines[DeclLine] = TrimLeft(ALines[DeclLine])) then
    begin
      // an unqualified top-level header in the implementation: either the
      // body of a routine declared in the interface, or impl-only
      for var I := 0 to Min(Impl, High(ALines)) - 1 do
        if IsHeaderLine(Code(ALines[I]), Kind, IsCM) and (ALines[I] = TrimLeft(ALines[I]))
          and LineDeclaresName(ALines[I], AName) then
        begin
          DeclLine := I;
          Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
          Break;
        end;
    end;
    ASym.DeclLine := DeclLine;
    Dirs := HeaderDirectives(ALines, DeclLine, HdrEnd);
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    if ASym.ContainerLine >= 0 then
    begin
      ASym.Kind := sdkMethod;
      ASym.Container := CName;
      ASym.IsInterfaceMember := IsIntf;
      ASym.Parents := HeaderParents(Code(ALines[ASym.ContainerLine]));
      if not IsIntf then
        ASym.ImplLine := FindImplHeader(ALines, CName + '.' + AName);
    end
    else
    begin
      ASym.Kind := sdkRoutine;
      if DeclLine > Impl then
        ASym.ImplLine := DeclLine          // implementation-only (or nested)
      else
        ASym.ImplLine := FindImplHeader(ALines, AName);
      // a routine nested in another one is local to it
      if (ASym.ImplLine >= 0) and (ALines[ASym.ImplLine] <> TrimLeft(ALines[ASym.ImplLine])) then
      begin
        var OF_, OL_: Integer;
        if FindEnclosingRoutineRange(Content, ASym.ImplLine - 1, OF_, OL_) then
        begin
          ASym.LocalFirst := OF_;
          ASym.LocalLast := OL_;
        end;
      end;
    end;
    // what may reach a method WITHOUT a textual call
    for var W in ['overload', 'virtual', 'dynamic', 'abstract', 'override', 'message'] do
      if HasWholeWordCI(Dirs, W) then
        Veto(ASym, Format('the %s is declared "%s" - %s', [SafeDeleteKindText(ASym.Kind), W,
          IfThen(W = 'overload', 'another overload could silently take over its calls',
          'it can be reached through dispatch without a textual call')]));
    if HasWholeWordCI(Dirs, 'external') then ASym.ImplLine := -1;
    // edits: the declaration (when separate) and the implementation
    if ASym.ImplLine <> DeclLine then
      AddEdit(ASym, DeclLine, HdrEnd, ALines);
    if ASym.ImplLine >= 0 then
    begin
      if not FindEnclosingRoutineRange(Content, ASym.ImplLine, F, L) or (F <> ASym.ImplLine) then
      begin
        AWhy := 'the implementation block could not be delimited';
        Exit;
      end;
      AddEdit(ASym, F, L, ALines);
    end;
    // (declared but never implemented: deleting the declaration is enough)
  end
  // ---- properties ------------------------------------------------------------
  else if T.ToUpper.StartsWith('PROPERTY ') or T.ToUpper.StartsWith('CLASS PROPERTY ') then
  begin
    if not LineDeclaresName(ALines[DeclLine], AName) then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    ASym.Kind := sdkProperty;
    ASym.DeclLine := DeclLine;
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    ASym.Container := CName;
    ASym.IsInterfaceMember := IsIntf;
    if not T.EndsWith(';') then
      Veto(ASym, 'the property declaration spans several lines - not supported');
    AddEdit(ASym, DeclLine, DeclLine, ALines);
  end
  // ---- data: fields, variables, constants, single-line types -------------------
  else
  begin
    var NS, Sep: Integer;
    var IsEq: Boolean;
    var Names := SplitDeclNames(ALines[DeclLine], NS, Sep, IsEq);
    var Idx := -1;
    for var I := 0 to High(Names) do
      if SameText(Names[I], AName) then Idx := I;
    if Idx < 0 then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    ASym.DeclLine := DeclLine;
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    ASym.Container := CName;
    U := UpperCase(Trim(Copy(StripLineComment(ALines[DeclLine]), Sep + 1, MaxInt)));
    if ASym.ContainerLine >= 0 then
      ASym.Kind := sdkField
    else
    begin
      var Sec := SectionAt(ALines, DeclLine);
      if (Sec = 'type') or (IsEq and (Sec = '') and (U.StartsWith('CLASS') or U.StartsWith('RECORD'))) then
        ASym.Kind := sdkType
      else if (Sec = 'const') or (Sec = 'resourcestring') or IsEq then
        ASym.Kind := sdkConstant
      else
        ASym.Kind := sdkVariable;
      var RF, RL: Integer;
      if FindEnclosingRoutineRange(Content, DeclLine, RF, RL) then
      begin
        ASym.LocalFirst := RF;
        ASym.LocalLast := RL;
      end;
    end;
    if (ASym.Kind = sdkType) and ((ClassOpenerName(Code(ALines[DeclLine])) <> '')
      or U.StartsWith('CLASS') or U.StartsWith('RECORD') or U.StartsWith('INTERFACE')
      or U.StartsWith('PACKED')) and not Code(ALines[DeclLine]).EndsWith(';') then
    begin
      Veto(ASym, 'types with a body are not supported yet');
    end
    else if not Code(ALines[DeclLine]).EndsWith(';') then
      Veto(ASym, 'the declaration spans several lines - not supported');
    if Length(Names) > 1 then
    begin
      // drop the name from the list, keep everything else as written
      var S := ALines[DeclLine];
      var Kept: TArray<string> := nil;
      for var I := 0 to High(Names) do
        if I <> Idx then Kept := Kept + [Names[I]];
      var E := Default(TSafeDeleteEdit);
      E.FirstLine := DeclLine;
      E.LastLine := DeclLine;
      E.HasReplacement := True;
      E.Replacement := Copy(S, 1, NS - 1) + string.Join(', ', Kept) +
        IfThen(IsEq, ' ', '') + Copy(S, Sep, MaxInt);
      ASym.Edits := ASym.Edits + [E];
    end
    else
    begin
      // the whole line - and a section keyword left without declarations
      F := DeclLine;
      L := DeclLine;
      var W := FirstWordU(Code(ALines[DeclLine]));
      var KeywordOnLine := (W = 'VAR') or (W = 'CONST') or (W = 'TYPE') or (W = 'THREADVAR');
      if not KeywordOnLine then
      begin
        var Prev := DeclLine - 1;
        while (Prev >= 0) and (Code(ALines[Prev]) = '') do Dec(Prev);
        var Next := DeclLine + 1;
        while (Next <= High(ALines)) and (Code(ALines[Next]) = '') do Inc(Next);
        var PW := UpperCase(Code(IfThen(Prev >= 0, ALines[Prev], '')));
        var NW := '';
        if Next <= High(ALines) then NW := FirstWordU(Code(ALines[Next]));
        if ((PW = 'VAR') or (PW = 'CONST') or (PW = 'TYPE') or (PW = 'THREADVAR') or
            (PW = 'CLASS VAR') or (PW = 'RESOURCESTRING'))
          and ((NW = 'BEGIN') or (NW = 'VAR') or (NW = 'CONST') or (NW = 'TYPE') or
            (NW = 'THREADVAR') or (NW = 'RESOURCESTRING') or (NW = 'LABEL') or
            (NW = 'PROCEDURE') or (NW = 'FUNCTION') or (NW = 'CONSTRUCTOR') or
            (NW = 'DESTRUCTOR') or (NW = 'CLASS') or (NW = 'IMPLEMENTATION') or
            (NW = 'INITIALIZATION') or (NW = 'FINALIZATION') or (NW = 'END') or
            (NW = 'PRIVATE') or (NW = 'PROTECTED') or (NW = 'PUBLIC') or
            (NW = 'PUBLISHED') or (NW = 'STRICT') or (NW = 'PROPERTY') or (NW = '')) then
          F := Prev;
      end;
      AddEdit(ASym, F, L, ALines);
    end;
  end;

  // ---- container checks -------------------------------------------------------
  if (ASym.ContainerLine >= 0) and not ASym.IsInterfaceMember then
  begin
    ASym.Parents := HeaderParents(Code(ALines[ASym.ContainerLine]));
    if VisibilityAt(ALines, ASym.ContainerLine, ASym.DeclLine) = 'published' then
      Veto(ASym, 'the member is PUBLISHED - it can be used by name (form streaming, RTTI)');
  end;
  Result := True;
end;

function ApplySafeDeleteEdits(const ALines: TArray<string>;
  const AEdits: TArray<TSafeDeleteEdit>): TArray<string>;
var
  Sorted: TArray<TSafeDeleteEdit>;
  Res: TArray<string>;
begin
  Sorted := Copy(AEdits);
  // bottom-up
  for var I := 0 to High(Sorted) - 1 do
    for var J := I + 1 to High(Sorted) do
      if Sorted[J].FirstLine > Sorted[I].FirstLine then
      begin
        var Tmp := Sorted[I];
        Sorted[I] := Sorted[J];
        Sorted[J] := Tmp;
      end;
  Res := Copy(ALines);
  for var E in Sorted do
  begin
    if (E.FirstLine < 0) or (E.LastLine > High(Res)) or (E.LastLine < E.FirstLine) then Continue;
    var Tail := Copy(Res, E.LastLine + 1, MaxInt);
    Res := Copy(Res, 0, E.FirstLine);
    if E.HasReplacement then Res := Res + [E.Replacement];
    Res := Res + Tail;
  end;
  Result := Res;
end;

function FormTextMentions(const AText, AName: string): TArray<Integer>;
var
  Lines: TArray<string>;
  U, W: string;
  P, E: Integer;
begin
  Result := nil;
  if AName = '' then Exit;
  W := UpperCase(AName);
  Lines := AText.Replace(#13#10, #10).Split([#10]);
  for var L := 0 to High(Lines) do
  begin
    U := UpperCase(Lines[L]);
    P := Pos(W, U);
    while P > 0 do
    begin
      E := P + Length(W);
      if ((P = 1) or not IsIdentChar(U[P - 1])) and ((E > Length(U)) or not IsIdentChar(U[E])) then
      begin
        Result := Result + [L];
        Break;
      end;
      P := Pos(W, U, P + 1);
    end;
  end;
end;

end.
