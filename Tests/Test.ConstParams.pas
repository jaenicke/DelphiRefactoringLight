(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Test.ConstParams;

// The project check "Missing const (string parameters)" - TurboMagic's idea
// from the forum. Every case below is a shape that was MEASURED with
// dcc32 37.0 before the rule was written, and the two interesting ones are
// the shapes nobody's condition list had: a method bound by an INTERFACE
// (its header carries no directive at all, and const makes it stop
// satisfying the interface - E2291) and a routine used as a METHOD POINTER
// (const makes the two signatures incompatible types - E2009).

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TConstParamTests = class
  public
    /// <summary>The shape the check exists for: a method whose body only
    ///  reads the parameter. One const per header, and the two of them
    ///  together are what keeps the unit compiling (E2037 otherwise).</summary>
    [Test] procedure AReadOnlyStringParameterCanTakeConst;
    /// <summary>Every refusal, each with its own routine in one unit, so a
    ///  rule that stops firing is named rather than quietly accepted.</summary>
    [Test] procedure EveryRefusalFiresForItsOwnReason;
    /// <summary>One modifier serves a GROUP ("A, B: string"), so const is
    ///  written once in front of the group - and a group with one written
    ///  member is refused as a whole.</summary>
    [Test] procedure AGroupIsAllOrNothing;
    /// <summary>The ids name the place, and a stale position is refused
    ///  instead of written to.</summary>
    [Test] procedure TheIdNamesThePlaceAndAStalePositionIsRefused;
    /// <summary>What the MCP tool reads, the schema must declare: an
    ///  argument it does not list is dropped before the handler sees it and
    ///  nobody is told (1.16.11).</summary>
    [Test] procedure TheConstToolDeclaresWhatItTakes;
    /// <summary>A WRAPPED parameter list: the group's ';' stands at the end
    ///  of the previous LINE, so the continuation line opens a new group.
    ///  Measured live - every fixture here had a one-line list, and the
    ///  check answered 0 findings AND 0 refusals for 113 real files.</summary>
    [Test] procedure AWrappedParameterListStillHasItsGroups;
  end;

implementation

uses
  System.SysUtils, System.Types, System.Classes,
  Expert.PascalScanner, Expert.ConstParams, Mcp.Protocol;

function Lines(const AText: string): TArray<string>;
begin
  Result := SplitEditorLines(AText);
end;

function FindingOf(const AFindings: TArray<TConstParamFinding>;
  const ARoutine: string): TConstParamFinding;
begin
  Result := Default(TConstParamFinding);
  for var F in AFindings do
    if SameText(F.Routine, ARoutine) then Exit(F);
end;

// ---------------------------------------------------------------------------

procedure TConstParamTests.AReadOnlyStringParameterCanTakeConst;
const
  Src =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TFoo = class'#13#10 +
    '  private'#13#10 +
    '    FH: string;'#13#10 +
    '    procedure Keep(AFile: string);'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'procedure TFoo.Keep(AFile: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AnsiUpperCase(AFile);'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
var
  New: TArray<string>;
  Why: string;
begin
  var L := Lines(Src);
  var Found := CollectConstParamCandidates('U.pas', L, nil);
  Assert.AreEqual(1, Integer(Length(Found)), 'one candidate');
  var F := Found[0];
  Assert.AreEqual('TFoo.Keep', F.Routine, 'the routine is named');
  Assert.AreEqual('AFile', F.ParamNames[0]);
  Assert.AreEqual(cprNone, F.Refusal, F.Reason);
  Assert.AreEqual(2, Integer(Length(F.Positions)),
    'BOTH headers get the const - one of them alone is E2037');

  Assert.IsTrue(PlanConstParamEdit(L, F, New, Why), Why);
  var Text := string.Join(#10, New);
  Assert.IsTrue(Text.Contains('procedure Keep(const AFile: string);'),
    'the declaration: ' + Text);
  Assert.IsTrue(Text.Contains('procedure TFoo.Keep(const AFile: string);'),
    'the implementation header: ' + Text);
  Assert.IsFalse(Text.Contains('const const'),
    'and the same parameter is found once, not once per header');
end;

procedure TConstParamTests.EveryRefusalFiresForItsOwnReason;
const
  Src =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TFoo = class'#13#10 +
    '  private'#13#10 +
    '    FH: string;'#13#10 +
    '    procedure Writes(AText: string);'#13#10 +
    '    procedure Virt(AText: string); virtual;'#13#10 +
    '    procedure Bound(AText: string);'#13#10 +
    '    procedure Pointed(AText: string);'#13#10 +
    '    procedure Takes(AText: string);'#13#10 +
    '  published'#13#10 +
    '    procedure Handler(AText: string);'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'procedure TFoo.Writes(AText: string);'#13#10 +
    'begin'#13#10 +
    '  AText := AText + ''x'';'#13#10 +
    '  FH := AText;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Virt(AText: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AText;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Bound(AText: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AText;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Pointed(AText: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AText;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Takes(AText: string);'#13#10 +
    'begin'#13#10 +
    '  FH := PString(@AText)^;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Handler(AText: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AText;'#13#10 +
    'end;'#13#10 +
    'procedure Install(AFoo: TFoo);'#13#10 +
    'var'#13#10 +
    '  P: TProc<string>;'#13#10 +
    'begin'#13#10 +
    '  P := AFoo.Pointed;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
var
  Asked: TStringList;
begin
  Asked := TStringList.Create;
  try
    var Bound: TConstBoundLookup :=
      function(const AOwner, AMember: string; out AWhy: string): Boolean
      begin
        Asked.Add(AOwner + '.' + AMember);
        Result := SameText(AMember, 'Bound');
        if Result then
          AWhy := 'IThing declares Bound, and const would stop TFoo from ' +
            'satisfying it (E2291)';
      end;

    var Found := CollectConstParamCandidates('U.pas', Lines(Src), Bound, True);

    Assert.AreEqual(cprBodyWrites, FindingOf(Found, 'TFoo.Writes').Refusal,
      'the body assigns it');
    Assert.AreEqual(cprFamily, FindingOf(Found, 'TFoo.Virt').Refusal,
      'virtual - the family would have to move together');
    Assert.AreEqual(cprInterfaceBound, FindingOf(Found, 'TFoo.Bound').Refusal,
      'the lookup says an interface declares it');
    Assert.AreEqual(cprPublished, FindingOf(Found, 'TFoo.Handler').Refusal,
      'published - a .dfm may bind it, which the compiler never sees');
    Assert.AreEqual(cprMethodReference, FindingOf(Found, 'TFoo.Pointed').Refusal,
      'assigned as a value, so the signatures must stay compatible');
    Assert.AreEqual(cprBodyAddress, FindingOf(Found, 'TFoo.Takes').Refusal,
      'the body takes its address');

    Assert.IsTrue(Asked.Count >= 1, 'the type graph was asked at all');
    // and nothing of this unit is proposed
    var Candidates := 0;
    for var F in Found do
      if F.Refusal = cprNone then Inc(Candidates);
    Assert.AreEqual(0, Candidates, 'none of these six may be proposed');
  finally
    Asked.Free;
  end;
end;

procedure TConstParamTests.AGroupIsAllOrNothing;
const
  Src =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TFoo = class'#13#10 +
    '  private'#13#10 +
    '    FH: string;'#13#10 +
    '    procedure Clean(A, B: string);'#13#10 +
    '    procedure Dirty(A, B: string);'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'procedure TFoo.Clean(A, B: string);'#13#10 +
    'begin'#13#10 +
    '  FH := A + B;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Dirty(A, B: string);'#13#10 +
    'begin'#13#10 +
    '  B := A;'#13#10 +
    '  FH := B;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
var
  New: TArray<string>;
  Why: string;
begin
  var L := Lines(Src);
  var Found := CollectConstParamCandidates('U.pas', L, nil, True);

  var C := FindingOf(Found, 'TFoo.Clean');
  Assert.AreEqual(cprNone, C.Refusal, C.Reason);
  Assert.AreEqual(2, Integer(Length(C.ParamNames)), 'the group carries both names');
  Assert.AreEqual(2, Integer(Length(C.Positions)),
    'ONE const per header, in front of the group - not one per name');
  Assert.IsTrue(PlanConstParamEdit(L, C, New, Why), Why);
  var Text := string.Join(#10, New);
  Assert.IsTrue(Text.Contains('procedure Clean(const A, B: string);'),
    'one modifier for the group: ' + Text);

  Assert.AreEqual(cprGroupMixed, FindingOf(Found, 'TFoo.Dirty').Refusal,
    'one written member refuses the whole group - a modifier cannot be ' +
    'given to half of it');
end;

procedure TConstParamTests.TheIdNamesThePlaceAndAStalePositionIsRefused;
const
  Src =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TFoo = class'#13#10 +
    '  private'#13#10 +
    '    FH: string;'#13#10 +
    '    procedure One(A: string);'#13#10 +
    '    procedure Two(A: string);'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'procedure TFoo.One(A: string);'#13#10 +
    'begin'#13#10 +
    '  FH := A;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Two(A: string);'#13#10 +
    'begin'#13#10 +
    '  FH := A;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
var
  New: TArray<string>;
  Why: string;
begin
  var L := Lines(Src);
  var Found := CollectConstParamCandidates('U.pas', L, nil);
  Assert.AreEqual(2, Integer(Length(Found)), 'both routines are candidates');
  Assert.AreNotEqual(Found[0].Id, Found[1].Id,
    'two groups of the same parameter NAME must not share one id');
  Assert.AreEqual(ConstParamIssueId('u.PAS', 'tfoo.one', 'a'),
    ConstParamIssueId('U.pas', 'TFoo.One', 'A'), 'the id ignores case');

  // the check may have run minutes ago
  var Stale := Found[0];
  Stale.Positions := [Point(0, 1)];
  Assert.IsFalse(PlanConstParamEdit(L, Stale, New, Why),
    'a position that no longer carries the name is refused');
  Assert.IsTrue(Why.Contains('no longer carries'), Why);
end;

procedure TConstParamTests.TheConstToolDeclaresWhatItTakes;
var
  Args: TArray<string>;

  function Declares(const AName: string): Boolean;
  begin
    Result := False;
    for var A in Args do
      if SameText(A, AName) then Exit(True);
  end;

begin
  Args := KnownToolArguments('const_params');
  Assert.IsTrue(Length(Args) > 0, 'the tool is in the shipped definitions');
  Assert.IsTrue(Declares('apply'), 'apply is declared');
  Assert.IsTrue(Declares('fix_ids'), 'fix_ids is declared');
  Assert.IsTrue(Declares('include_refused'), 'include_refused is declared');
  Assert.IsTrue(Declares('max'), 'max is declared');
  Assert.IsFalse(Declares('file'), 'and nothing it does not read');
end;

procedure TConstParamTests.AWrappedParameterListStillHasItsGroups;
const
  // The real header of TLspFindUnitReferencesWizard.ReadCachedLines, which
  // is the candidate the live run lost: AFile ends its group with the ';'
  // at the end of its line, and ACache - no string - starts a new one.
  Src =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'uses System.Generics.Collections;'#13#10 +
    'type'#13#10 +
    '  TFoo = class'#13#10 +
    '  private'#13#10 +
    '    FH: string;'#13#10 +
    '    function ReadCachedLines(AFile: string;'#13#10 +
    '      ACache: TDictionary<string, Integer>): Integer;'#13#10 +
    '    procedure Pair(AFirst,'#13#10 +
    '      ASecond: string);'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'function TFoo.ReadCachedLines(AFile: string;'#13#10 +
    '  ACache: TDictionary<string, Integer>): Integer;'#13#10 +
    'begin'#13#10 +
    '  FH := AFile;'#13#10 +
    '  Result := ACache.Count;'#13#10 +
    'end;'#13#10 +
    'procedure TFoo.Pair(AFirst,'#13#10 +
    '  ASecond: string);'#13#10 +
    'begin'#13#10 +
    '  FH := AFirst + ASecond;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
begin
  var L := Lines(Src);
  var Found := CollectConstParamCandidates('U.pas', L, nil, True);

  var Read := FindingOf(Found, 'TFoo.ReadCachedLines');
  Assert.AreEqual('AFile', string.Join(',', Read.ParamNames),
    'AFile ends its group at the ";" of its own line - ACache, which is no ' +
    'string, must NOT be swallowed into that group (the live defect)');
  Assert.AreEqual(cprNone, Read.Refusal, Read.Reason);
  Assert.AreEqual(2, Integer(Length(Read.Positions)),
    'both headers of the wrapped routine carry the position');

  // ...and the other direction still holds: a group really does continue
  // across the line break when the previous line ends in a comma.
  var Pair := FindingOf(Found, 'TFoo.Pair');
  Assert.AreEqual('AFirst,ASecond', string.Join(',', Pair.ParamNames),
    'one modifier serves "AFirst," / "ASecond: string" - it is ONE group ' +
    'although it is written on two lines');
end;

initialization
  TDUnitX.RegisterTestFixture(TConstParamTests);

end.
