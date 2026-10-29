(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ConstParams;

// The pure half of the project check "Missing const (string parameters)"
// (idea: TurboMagic in the forum, thread #226 ff.). A string parameter
// without const makes every call raise and lower a reference count; const
// passes the caller's reference as it is.
//
// EVERY RULE BELOW WAS MEASURED with dcc32 37.0 before it was written, and
// two of them are not in anybody's condition list:
// * a method that IMPLEMENTS AN INTERFACE member carries no directive at
//   all, so a family rule reading virtual/override cannot see it - and with
//   const it stops satisfying the interface: E2291 "Missing implementation
//   of interface method IFoo.Bar";
// * `procedure(S: string)` and `procedure(const S: string)` are
//   INCOMPATIBLE types (E2009 "Parameter lists differ"), so a routine used
//   as a method pointer, and above all an anonymous method whose procedural
//   type dictates its signature, must keep its parameter as it is. In this
//   repository's own, const-disciplined sources 42 of 44 const-less string
//   parameters are exactly that case.
// Not a condition, also measured: the CALLING CONVENTION does not matter
// (register / cdecl / stdcall / safecall / pascal all compile with a const
// string), and a property accessor takes const (`property H: string write
// SetH` accepts `SetH(const Value: string)`).
//
// AND WHY THIS IS A LIST, never a sweep that rewrites a project: a const
// string parameter ALIASES the caller's variable without raising its
// reference count, so if the body - or anything it calls - releases that
// variable, the parameter dangles. Measured on twelve lines: the same
// program prints its string with a value parameter and dies with
// EInvalidPointer with const. No compiler warning, and not decidable from
// the text. So this check PROPOSES and names its reasons; the decision
// belongs to whoever reads the row.

interface

uses
  System.SysUtils, System.Types;

type
  /// <summary>Why a const-less string parameter is NOT proposed. cprNone is
  ///  a candidate.</summary>
  TConstParamRefusal = (
    cprNone,
    /// <summary>The body assigns it, passes it on as var/out or uses it as
    ///  a loop variable - const would not compile (E2064).</summary>
    cprBodyWrites,
    /// <summary>The body takes its address.</summary>
    cprBodyAddress,
    /// <summary>virtual / override / dynamic / message / abstract: the
    ///  whole family would have to change together.</summary>
    cprFamily,
    /// <summary>The owner type implements an interface that declares this
    ///  member. Nothing in the header says so, and const breaks the
    ///  implementation (measured: E2291).</summary>
    cprInterfaceBound,
    /// <summary>A published method: a .dfm binds it as an event handler by
    ///  NAME, the compiler never sees that assignment, and the signature
    ///  has to match at run time. The one case that fails silently.</summary>
    cprPublished,
    /// <summary>The routine's own name is used as a value in this unit, so
    ///  it is (or may be) a method pointer - measured E2009.</summary>
    cprMethodReference,
    /// <summary>The routine has ONE header: an interface or dispinterface
    ///  member, an abstract method, an external declaration. Changing it
    ///  reaches implementers or an ABI this check cannot see.</summary>
    cprSingleHeader,
    /// <summary>The two headers a parameter is declared in could not be
    ///  established, so only one of them would be changed (E2037).</summary>
    cprNotPaired,
    /// <summary>The routine's body could not be delimited, so "the body
    ///  does not write it" was never answered.</summary>
    cprBodyUnknown,
    /// <summary>The two headers group the parameters differently, so one
    ///  const would not cover the same names in both.</summary>
    cprGroupDiffers,
    /// <summary>One modifier serves a whole group ("A, B: string"), and one
    ///  member of this group is refused - so the group is.</summary>
    cprGroupMixed,
    /// <summary>The header could not be read as a parameter list at
    ///  all.</summary>
    cprHeaderUnreadable);

  /// <summary>One group of const-less string parameters of one routine.
  ///  A GROUP is the unit of work, not a parameter: one modifier serves
  ///  "A, B: string", so const is written once, in front of the group's
  ///  first name, and the group is all-or-nothing.
  ///  POSITIONS are every place a 'const ' goes - the group's first name in
  ///  the declaration AND in the implementation header, because Delphi
  ///  requires the two lists to agree and a half-changed routine does not
  ///  compile (E2037).</summary>
  TConstParamFinding = record
    FileName: string;
    /// <summary>'TFoo.Bar' or 'Bar'.</summary>
    Routine: string;
    /// <summary>The group's names, in order.</summary>
    ParamNames: TArray<string>;
    /// <summary>'string', 'UnicodeString', ...</summary>
    ParamType: string;
    /// <summary>Where the row points: the group's first name in the header
    ///  this was found through (0-based).</summary>
    Line, Col: Integer;
    /// <summary>Every place a 'const ' goes (0-based, X = column,
    ///  Y = line), sorted so a writer can apply them from the end.</summary>
    Positions: TArray<TPoint>;
    Refusal: TConstParamRefusal;
    /// <summary>One sentence: what was found and what it means. Always
    ///  filled, for a candidate as well.</summary>
    Reason: string;
    /// <summary>Stable id of the PLACE, so an answer can name a row
    ///  whatever the sort order is.</summary>
    Id: string;
  end;

  /// <summary>Answers whether AMEMBER of AOWNERTYPE is bound by something
  ///  this file cannot see: an interface the type implements, or an
  ///  ancestor's method. The pure collector cannot know it - the type graph
  ///  of the whole project does - so it comes in as a lookup, like the
  ///  designer answer of the uses cleanup does. nil = not asked, and then
  ///  the check says so in the row's reason instead of implying it
  ///  looked.</summary>
  TConstBoundLookup = reference to function(const AOwnerType, AMember: string;
    out AWhy: string): Boolean;

/// <summary>FNV-1a over the lower-cased key - the same shape the other
///  project checks' ids have. It names the PLACE (file, routine, first
///  parameter of the group), never a list position.</summary>
function ConstParamIssueId(const AFile, ARoutine, AParam: string): string;

/// <summary>Every group of const-less string parameters in ALINES, with its
///  verdict. AINCLUDEREFUSED adds the refused ones - the dialog counts them
///  and the MCP answer can name them, because "it was not proposed" and "it
///  was never looked at" must not read the same.
///  ALINES are the RAW lines; everything is decided on the masked text.</summary>
function CollectConstParamCandidates(const AFile: string;
  const ALines: TArray<string>; const ABound: TConstBoundLookup;
  AIncludeRefused: Boolean = False): TArray<TConstParamFinding>;

/// <summary>ALINES with 'const ' written in at every position of AFINDING.
///  False with a reason when the text no longer looks the way the finding
///  describes - the check may have run minutes ago, and a stale position
///  must never be written to.</summary>
function PlanConstParamEdit(const ALines: TArray<string>;
  const AFinding: TConstParamFinding; out ANew: TArray<string>;
  out AWhy: string): Boolean;

/// <summary>One short phrase per refusal, for a column or a count.</summary>
function ConstParamRefusalText(ARefusal: TConstParamRefusal): string;

/// <summary>The string types this check proposes const for. Deliberately
///  the WRITTEN names only: an alias (TCaption, TFileName) cannot be
///  resolved without the identifier index, and a check that guessed would
///  propose const for something that is not a string at all.</summary>
function IsStringTypeName(const AType: string): Boolean;

implementation

uses
  System.StrUtils, System.Math, System.Generics.Collections,
  Expert.PascalScanner, Expert.SafeDeletePlan, Expert.SignatureEdit,
  Expert.SelectionValidator, Expert.AutoImport;

function ConstParamIssueId(const AFile, ARoutine, AParam: string): string;
var
  H: UInt32;
begin
  H := 2166136261;
  for var C in LowerCase(AFile + '|' + ARoutine + '|' + AParam) do
  begin
    H := H xor UInt32(Ord(C));
    H := H * 16777619;
  end;
  Result := IntToHex(H, 8);
end;

function ConstParamRefusalText(ARefusal: TConstParamRefusal): string;
begin
  case ARefusal of
    cprNone: Result := 'can take const';
    cprBodyWrites: Result := 'the body writes it';
    cprBodyAddress: Result := 'the body takes its address';
    cprFamily: Result := 'part of a virtual family';
    cprInterfaceBound: Result := 'bound by an interface';
    cprPublished: Result := 'published - a form may bind it';
    cprMethodReference: Result := 'used as a method pointer';
    cprSingleHeader: Result := 'declared once (interface / abstract / external)';
    cprNotPaired: Result := 'the two headers could not be paired';
    cprBodyUnknown: Result := 'the body could not be delimited';
    cprGroupDiffers: Result := 'the headers group the parameters differently';
    cprGroupMixed: Result := 'another parameter of its group is refused';
    cprHeaderUnreadable: Result := 'the header could not be read';
  else
    Result := '';
  end;
end;

function IsStringTypeName(const AType: string): Boolean;
begin
  var T := Trim(AType);
  Result := SameText(T, 'string') or SameText(T, 'UnicodeString')
    or SameText(T, 'AnsiString') or SameText(T, 'WideString');
end;

// ---------------------------------------------------------------------------

type
  // One group of a header: the indices into the span's Names it covers.
  TParamGroup = record
    First, Last: Integer;
  end;

// The groups of a header, derived from the MASKED text around the names a
// span reports: a name whose preceding code character is '(' or ';' opens a
// group, every following name belongs to it until the next one does.
function GroupsOfSpan(const AMasked: TArray<string>;
  const ASpan: THeaderSpan): TArray<TParamGroup>;

  // The last code character BEFORE the name, ACROSS LINE BREAKS. A wrapped
  // parameter list carries the group's ';' at the end of the previous line,
  // so a walk that stops at column 1 reads the first parameter of every
  // continuation line as a continuation of the group above it - measured
  // live on "function ReadCachedLines(AFile: string;" / "  ACache: ...",
  // where it merged the two and the group was then no candidate at all.
  function CodeCharBefore(ALine, ACol: Integer): Char;
  begin
    Result := #0;
    while ALine >= ASpan.First do
    begin
      var S := '';
      if (ALine >= 0) and (ALine <= High(AMasked)) then S := AMasked[ALine];
      var J := ACol;
      if J > Length(S) then J := Length(S);
      while J >= 1 do
      begin
        if S[J] > ' ' then Exit(S[J]);
        Dec(J);
      end;
      Dec(ALine);
      ACol := MaxInt;
    end;
  end;

begin
  Result := nil;
  for var I := 0 to High(ASpan.Names) do
  begin
    var P := ASpan.Names[I];
    // '(' opens the list, ';' ends a group, ',' continues one.
    var Opens := CharInSet(CodeCharBefore(P.Y, P.X), ['(', ';']);
    if Opens or (Length(Result) = 0) then
    begin
      var G: TParamGroup;
      G.First := I;
      G.Last := I;
      Result := Result + [G];
    end
    else
      Result[High(Result)].Last := I;
  end;
end;

// The visibility section ALINE0 sits in, '' when it is not in a type body.
function VisibilityAt(const AMasked: TArray<string>; ALine0: Integer): string;
begin
  Result := '';
  for var I := ALine0 - 1 downto Max(0, ALine0 - 600) do
  begin
    var T := LowerCase(Trim(AMasked[I]));
    while T.EndsWith(';') do T := Copy(T, 1, Length(T) - 1);
    if (T = 'published') or (T = 'public') or (T = 'private')
      or (T = 'protected') or (T = 'strict private')
      or (T = 'strict protected') then Exit(T);
    if (T = 'implementation') or (T = 'interface') then Exit('');
  end;
end;

function BodyWritesName(const ALines: TArray<string>; AFirst, ALast: Integer;
  const AName: string): Boolean;
var
  Block: TArray<string>;
begin
  Block := nil;
  for var I := Max(0, AFirst) to Min(ALast, High(ALines)) do
    Block := Block + [ALines[I]];
  for var W in WrittenIdentifiersIn(Block) do
    if SameText(W, AName) then Exit(True);
  Result := False;
end;

function BodyTakesAddress(const AMasked: TArray<string>; AFirst, ALast: Integer;
  const AName: string): Boolean;
begin
  for var I := Max(0, AFirst) to Min(ALast, High(AMasked)) do
  begin
    var S := AMasked[I];
    var P := 1;
    repeat
      P := Pos('@', S, P);
      if P = 0 then Break;
      var Q := P + 1;
      while (Q <= Length(S)) and (S[Q] = ' ') do Inc(Q);
      var R := Q;
      while (R <= Length(S)) and IsIdentChar(S[R]) do Inc(R);
      if SameText(Copy(S, Q, R - Q), AName) then Exit(True);
      Inc(P);
    until False;
  end;
  Result := False;
end;

// Is the ROUTINE itself used as a value anywhere in this unit? ccReference
// is "a call OR a method reference, the text cannot tell" - and for this
// question that uncertainty has to count against const, because the two
// signatures are incompatible types (E2009). Deliberately this unit only:
// a cross-unit method pointer is a compile error the user sees at once,
// while proposing it would be a row that cannot be applied.
function UsedAsValueInUnit(const AMasked: TArray<string>;
  const ASimpleName: string; const ASpans: TArray<TParamSpan>): Boolean;
begin
  for var L := 0 to High(AMasked) do
  begin
    // the routine's own headers are not uses of it
    var InHeader := False;
    for var Sp in ASpans do
      if (L >= Sp.First) and (L <= Sp.Last) then InHeader := True;
    if InHeader then Continue;
    var S := AMasked[L];
    var P := 1;
    while P <= Length(S) do
    begin
      if IsIdentStart(S[P]) and ((P = 1) or not IsIdentChar(S[P - 1])) then
      begin
        var R := P;
        while (R <= Length(S)) and IsIdentChar(S[R]) do Inc(R);
        if SameText(Copy(S, P, R - P), ASimpleName) then
        begin
          var O, C: Integer;
          if CallContextAt(S, P, R - 1, O, C) = ccReference then Exit(True);
        end;
        P := R;
      end
      else
        Inc(P);
    end;
  end;
  Result := False;
end;

function CollectConstParamCandidates(const AFile: string;
  const ALines: TArray<string>; const ABound: TConstBoundLookup;
  AIncludeRefused: Boolean): TArray<TConstParamFinding>;
var
  Masked: TArray<string>;
  Seen: TList<string>;

  // 'const ' goes in front of the group's FIRST name - in every header the
  // parameter is declared in, which is what keeps the two lists equal.
  function GroupPositions(const APlan: TParamRenamePlan; const AFirstName: string;
    out APos: TArray<TPoint>; out AWhy: string): Boolean;
  var
    Span: THeaderSpan;
  begin
    APos := nil;
    AWhy := '';
    for var Sp in APlan.Spans do
    begin
      if not ReadHeaderSpanFrom(Masked, Sp.First, Sp.FirstCol + 1, Span) then
      begin
        AWhy := Format('the header at line %d could not be read again', [Sp.First + 1]);
        Exit(False);
      end;
      var Groups := GroupsOfSpan(Masked, Span);
      var Found := False;
      for var G in Groups do
        for var I := G.First to G.Last do
          if SameText(Span.NameTexts[I], AFirstName) then
          begin
            if I <> G.First then
            begin
              AWhy := Format('%s does not open its group in the header at ' +
                'line %d, so one const would not cover the same names in ' +
                'both', [AFirstName, Sp.First + 1]);
              Exit(False);
            end;
            APos := APos + [Span.Names[G.First]];
            Found := True;
          end;
      if not Found then
      begin
        AWhy := Format('%s was not found in the header at line %d',
          [AFirstName, Sp.First + 1]);
        Exit(False);
      end;
    end;
    Result := Length(APos) > 0;
    if not Result and (AWhy = '') then AWhy := 'no place for the const was found';
  end;

  procedure Add(const ARoutine: string; const ANames: TArray<string>;
    const AType: string; ALine, ACol: Integer; const APos: TArray<TPoint>;
    ARefusal: TConstParamRefusal; const AReason: string);
  var
    F: TConstParamFinding;
  begin
    if (ARefusal <> cprNone) and not AIncludeRefused then Exit;
    F := Default(TConstParamFinding);
    F.FileName := AFile;
    F.Routine := ARoutine;
    F.ParamNames := ANames;
    F.ParamType := AType;
    F.Line := ALine;
    F.Col := ACol;
    F.Positions := APos;
    F.Refusal := ARefusal;
    F.Reason := AReason;
    F.Id := ConstParamIssueId(AFile, ARoutine, ANames[0]);
    if Seen.Contains(F.Id) then Exit;   // the same group is found once per header
    Seen.Add(F.Id);
    Result := Result + [F];
  end;

var
  Span: THeaderSpan;
begin
  Result := nil;
  if Length(ALines) = 0 then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  Seen := TList<string>.Create;
  try
    for var L := 0 to High(Masked) do
      for var Start in DeclarationStartsOnLine(Masked[L], True) do
      begin
        if not ReadHeaderSpanFrom(Masked, L, Start, Span) then Continue;
        if not Span.HasList then Continue;
        var Params := ParseParamList(Span.Params);
        if Length(Params) <> Length(Span.Names) then Continue;
        var Groups := GroupsOfSpan(Masked, Span);
        for var G in Groups do
        begin
          // a candidate group: every member a string, none with a modifier
          var IsCand := True;
          for var I := G.First to G.Last do
            if (Trim(Params[I].Modifier) <> '')
              or not IsStringTypeName(Params[I].TypeText) then IsCand := False;
          if not IsCand then Continue;

          var Names: TArray<string> := nil;
          for var I := G.First to G.Last do Names := Names + [Span.NameTexts[I]];
          var Line := Span.Names[G.First].Y;
          var Col := Span.Names[G.First].X;
          var Typ := Trim(Params[G.First].TypeText);

          // the plan is the authority on the two headers and the body
          var Plan := PlanParameterRename(ALines, L, Names[0], Col);

          // Read EVERY header of the routine ONCE, before any rule: the
          // same group is found once per header, so the name this is
          // reported under has to be the same whichever header it came
          // through - the QUALIFIED one, which only the implementation
          // carries. And a directive stands on the DECLARATION, which is
          // the other span exactly then.
          var Routine := Span.Qualified;
          var Owner := '';
          var DeclLine := -1;
          var Dirs := '';
          for var Sp in Plan.Spans do
          begin
            var S2: THeaderSpan;
            if not ReadHeaderSpanFrom(Masked, Sp.First, Sp.FirstCol + 1, S2) then
              Continue;
            Dirs := Dirs + ' ' + LowerCase(HeaderDirectives(ALines, Sp.First, Sp.Last));
            if Pos('.', S2.Qualified) > 0 then
            begin
              Routine := S2.Qualified;
              Owner := Copy(S2.Qualified, 1, Pos('.', S2.Qualified) - 1);
            end
            else
              DeclLine := Sp.First;
          end;
          if Routine = '' then Routine := '(unnamed)';
          if DeclLine < 0 then DeclLine := Plan.Spans[0].First;
          if Plan.Pairing = ppUnknown then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprNotPaired, Plan.Reason);
            Continue;
          end;
          if Plan.Pairing = ppSingleHeader then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprSingleHeader,
              'this routine is declared once - an interface or dispinterface ' +
              'member, an abstract method or an external declaration - so ' +
              'const would reach implementers or an ABI this check cannot see');
            Continue;
          end;

          var Fam := '';
          for var D in ['virtual', 'override', 'abstract', 'dynamic', 'message'] do
            if Pos(D, Dirs) > 0 then Fam := D;
          if Fam <> '' then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprFamily,
              Format('declared %s, so every member of the family would have ' +
              'to change together', [Fam]));
            Continue;
          end;

          var Member := Routine;
          if Pos('.', Member) > 0 then
            Member := Copy(Member, Pos('.', Member) + 1, MaxInt);

          if Owner <> '' then
          begin
            if VisibilityAt(Masked, DeclLine) = 'published' then
            begin
              Add(Routine, Names, Typ, Line, Col, nil, cprPublished,
                'a published method: a .dfm binds an event handler by NAME, ' +
                'the compiler never sees that assignment, and the signature ' +
                'has to match at run time');
              Continue;
            end;
            if Assigned(ABound) then
            begin
              var Why := '';
              if ABound(Owner, Member, Why) then
              begin
                Add(Routine, Names, Typ, Line, Col, nil, cprInterfaceBound, Why);
                Continue;
              end;
            end;
          end;

          if not Plan.ScopeKnown then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprBodyUnknown,
              'the routine''s body could not be delimited, so "the body does ' +
              'not write it" was never answered');
            Continue;
          end;

          var Written := '';
          for var N in Names do
            if BodyWritesName(ALines, Plan.BodyFirst, Plan.BodyLast, N) then
              Written := N;
          if Written <> '' then
          begin
            // one modifier serves the group, so a written member refuses the
            // whole group rather than itself
            var R1 := cprBodyWrites;
            if Length(Names) > 1 then R1 := cprGroupMixed;
            Add(Routine, Names, Typ, Line, Col, nil, R1,
              Format('the body writes %s, so const would not compile', [Written]));
            Continue;
          end;
          var Addr := '';
          for var N in Names do
            if BodyTakesAddress(Masked, Plan.BodyFirst, Plan.BodyLast, N) then
              Addr := N;
          if Addr <> '' then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprBodyAddress,
              Format('the body takes the address of %s', [Addr]));
            Continue;
          end;

          if UsedAsValueInUnit(Masked, Member, Plan.Spans) then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprMethodReference,
              Format('%s is used as a value in this unit, so it is or may be ' +
              'a method pointer - and a const parameter makes the two ' +
              'signatures incompatible types', [Member]));
            Continue;
          end;

          var Pos2: TArray<TPoint>;
          var Why2 := '';
          if not GroupPositions(Plan, Names[0], Pos2, Why2) then
          begin
            Add(Routine, Names, Typ, Line, Col, nil, cprGroupDiffers, Why2);
            Continue;
          end;

          var Note := '';
          if not Assigned(ABound) and (Owner <> '') then
            Note := ' (no type graph was asked, so an interface this type ' +
              'implements was not checked)';
          Add(Routine, Names, Typ, Line, Col, Pos2, cprNone,
            Format('%s: %s is passed by value; nothing writes it and its two ' +
            'headers are known%s', [Routine, string.Join(', ', Names), Note]));
        end;
      end;
  finally
    Seen.Free;
  end;
end;

function PlanConstParamEdit(const ALines: TArray<string>;
  const AFinding: TConstParamFinding; out ANew: TArray<string>;
  out AWhy: string): Boolean;
var
  Order: TArray<TPoint>;
begin
  ANew := nil;
  AWhy := '';
  Result := False;
  if Length(AFinding.Positions) = 0 then
  begin
    AWhy := 'the finding carries no place for the const';
    Exit;
  end;
  // Every position must still carry the name - the check may have run
  // minutes ago, and a stale position must never be written to.
  for var P in AFinding.Positions do
  begin
    if (P.Y < 0) or (P.Y > High(ALines)) then
    begin
      AWhy := Format('line %d is gone', [P.Y + 1]);
      Exit;
    end;
    var S := ALines[P.Y];
    if not SameText(Copy(S, P.X + 1, Length(AFinding.ParamNames[0])),
      AFinding.ParamNames[0]) then
    begin
      AWhy := Format('line %d no longer carries %s at column %d - run the ' +
        'check again', [P.Y + 1, AFinding.ParamNames[0], P.X + 1]);
      Exit;
    end;
  end;
  // bottom-up, so no position shifts under a later insert
  Order := Copy(AFinding.Positions);
  for var I := 0 to High(Order) do
    for var J := I + 1 to High(Order) do
      if (Order[J].Y > Order[I].Y)
        or ((Order[J].Y = Order[I].Y) and (Order[J].X > Order[I].X)) then
      begin
        var T := Order[I];
        Order[I] := Order[J];
        Order[J] := T;
      end;
  ANew := Copy(ALines);
  for var P in Order do
    ANew[P.Y] := Copy(ANew[P.Y], 1, P.X) + 'const ' +
      Copy(ANew[P.Y], P.X + 1, MaxInt);
  Result := True;
end;

end.
