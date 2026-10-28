unit Test.ParameterRename;

{ Issue #50, Ian Branch: his audit of the parameter pairing that 2.6.1 added
  for issue #49. Every case here is one where the rename of a parameter must
  either reach BOTH headers or refuse - Delphi requires the two names to be
  identical (E2037, plus an undeclared identifier for every use in the body),
  so renaming one of them is not a smaller change, it is a broken unit. }

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TParameterRenameAuditTests = class
  public
    /// <summary>#50.1: two overloads that carry the same parameter name at
    ///  the same index must not pair crosswise - and a default value, which
    ///  lives in the declaration only, must not stop them from pairing.</summary>
    [Test] procedure OverloadsPairByTheWholeList;
    /// <summary>#50.2/#50.3: where the pairing cannot be established the
    ///  answer is a REFUSAL, and a routine that legitimately has one header
    ///  says so instead.</summary>
    [Test] procedure WhatCannotBePairedIsRefused;
    /// <summary>#50.3: the header shapes the old parse could not read - a
    ///  class declared in the implementation section, a comment in the
    ///  parameter list, a header longer than twelve lines, "const [ref]".</summary>
    [Test] procedure TheHeaderShapesThatUsedToFail;
    /// <summary>#50.4: a target is a LINE, so the routine's own name, a type
    ///  or a default value spelled like the parameter was renamed with it.
    ///  The positions say which places really declare the parameter.</summary>
    [Test] procedure OnlyTheParameterItselfIsAPosition;
  end;

  [TestFixture]
  TParameterRenameShapeTests = class
  public
    /// <summary>#54.A/F/G: the owner of a member declaration is the
    ///  NESTING. A nested interface, record or class between the class
    ///  header and the declaration ended the search at ITSELF - for a
    ///  nested interface that even answered "this routine has ONE header",
    ///  which the wizard ACCEPTS, so the declaration was renamed and the
    ///  implementation header was not: E2037 plus an undeclared identifier
    ///  for every use in the body, measured with dcc32 37.0.</summary>
    [Test] procedure AMemberAfterANestedTypeIsPaired;
    /// <summary>#54.B: a parameter spelled like its own routine. The
    ///  rename's targets are line-keyed, so a CALL to the routine resolves
    ///  to a header line and was renamed with the parameter (E2003). The
    ///  parameter's scope is its routine's body, and that is what tells the
    ///  two apart.</summary>
    [Test] procedure AParameterOfItsOwnNameLeavesTheCallersAlone;
    /// <summary>#54.C: a header span is not a LINE. With the body written
    ///  after the ';' on the header's own line, the body's use of the
    ///  parameter sat "on the header" and was silently dropped -
    ///  E2003.</summary>
    [Test] procedure ABodyOnTheHeaderLineIsNotTheHeader;
    /// <summary>#54.D: Delphi allows TWO headers on one line, and then only
    ///  the caret's column says which one the rename means. Without it the
    ///  first header was read - the wrong parameter renamed in both of ITS
    ///  headers, which compiles.</summary>
    [Test] procedure TheCaretColumnPicksTheHeader;
    /// <summary>#54.E/H: a 'class operator' is a routine header like any
    ///  other, and an implementation header WITHOUT a parameter list is
    ///  legal Delphi that fits any declaration of that name - both were
    ///  refused.</summary>
    [Test] procedure AnOperatorAndAnOmittedListArePaired;
    /// <summary>What the old-vs-new sweep over the RAD Studio sources
    ///  found in the walk #54.A needed: a type declared in TWO {$IF}
    ///  branches has two openers and ONE 'end;', so it answered
    ///  'TControl.TControl' for every member of the biggest classes in the
    ///  VCL - which CreateImplStub was WRITING as a stub's qualifier - and
    ///  the extra opener also left the depth one too low, which gave a
    ///  sibling type further down a false outer level and a top-level
    ///  routine a container it has not got.</summary>
    [Test] procedure ATypeDeclaredInTwoBranchesIsOneType;
    /// <summary>The cross-unit half of #54.B, which 2.9.5 still renamed
    ///  (Ian Branch, measured live before the fix:
    ///  "[5] XuB.pas:14:3 => XuA.pas:9 -> MATCH" and the preview answered
    ///  "ATitle('x');"). A parameter cannot be used outside its routine, so
    ///  once the body is delimited and holds no {$I}nclude, no other file
    ///  can hold a use of it - and with an include nothing outside this
    ///  file is decided at all.</summary>
    [Test] procedure ACallFromAnotherUnitIsNotTheParameterEither;
  end;

implementation

uses
  System.SysUtils, System.Types, System.Classes, System.TypInfo,
  Expert.SafeDeletePlan, Expert.AutoImport, Expert.SignatureEdit,
  Expert.PascalScanner;

function L(const AText: string): TArray<string>;
begin
  Result := SplitContentLines(AText);
end;

procedure TParameterRenameAuditTests.OverloadsPairByTheWholeList;
const
  Src =
    'unit U;'#13#10 +                                              // 0
    'interface'#13#10 +                                            // 1
    'type'#13#10 +                                                 // 2
    '  TFoo = class'#13#10 +                                       // 3
    '    procedure Init(const Title: string); overload;'#13#10 +    // 4
    '    procedure Init(const Title: string; X: Integer = 0); overload;'#13#10 + // 5
    '  end;'#13#10 +                                               // 6
    'implementation'#13#10 +                                       // 7
    'procedure TFoo.Init(const Title: string);'#13#10 +            // 8
    'begin'#13#10 +                                                // 9
    'end;'#13#10 +                                                 // 10
    'procedure TFoo.Init(const Title: string; X: Integer);'#13#10 + // 11
    'begin'#13#10 +                                                // 12
    'end;'#13#10 +                                                 // 13
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
begin
  // the ONE-parameter declaration belongs to the one-parameter body
  Plan := PlanParameterRename(L(Src), 4, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'declaration 4 pairs with a body');
  Assert.AreEqual(8, Plan.Positions[High(Plan.Positions)].Y,
    'and it is the body at line 8, not the other overload''s');

  // ...and the TWO-parameter one with the two-parameter body, although its
  // own list carries a default the body does not repeat
  Plan := PlanParameterRename(L(Src), 5, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the default value must not stop the pairing');
  Assert.AreEqual(11, Plan.Positions[High(Plan.Positions)].Y,
    'the body with two parameters (the reported crosswise pairing)');

  // and from the body back to its own declaration
  Plan := PlanParameterRename(L(Src), 11, 'Title');
  Assert.AreEqual(5, Plan.Positions[0].Y, 'body 11 -> declaration 5');

  // THE POLARITY: a pair that does not match is no pair. Somebody renamed
  // one of the two names by hand, so the unit does not compile - and a
  // rename that now changes the other one by itself is a guess about which
  // of the two the user meant.
  var Half := string(Src).Replace('procedure TFoo.Init(const Title: string);',
    'procedure TFoo.Init(const Caption: string);');
  Plan := PlanParameterRename(L(Half), 4, 'Title');
  Assert.AreEqual('ppUnknown', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a declaration whose body carries another list is not paired');
  Assert.AreNotEqual('', Plan.Reason, 'and it says why');
end;

procedure TParameterRenameAuditTests.WhatCannotBePairedIsRefused;
const
  // a NESTED routine named like the interface routine (#50.2)
  Nested =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Run(const Title: string);'#13#10 +            // 2
    'implementation'#13#10 +                                 // 3
    'procedure Run(const Title: string);'#13#10 +            // 4
    'begin'#13#10 +                                          // 5
    'end;'#13#10 +                                           // 6
    'procedure Other;'#13#10 +                               // 7
    '  procedure Run(const Title: string);'#13#10 +           // 8
    '  begin'#13#10 +                                        // 9
    '  end;'#13#10 +                                          // 10
    'begin'#13#10 +                                          // 11
    'end;'#13#10 +                                           // 12
    'end.'#13#10;
  // an interface method and an abstract one have ONE header each
  Single_ =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  IFoo = interface'#13#10 +                             // 3
    '    procedure Go(const Title: string);'#13#10 +          // 4
    '  end;'#13#10 +                                         // 5
    '  TBase = class'#13#10 +                                // 6
    '    procedure Go(const Title: string); virtual; abstract;'#13#10 + // 7
    '  end;'#13#10 +                                         // 8
    'implementation'#13#10 +                                 // 9
    'end.'#13#10;
  // a shape nothing here can read: a parameter list that is never closed
  // (#50.3). It used to be an operator header - 2.9.4 reads those (issue
  // #54.E), and an assertion that keeps passing for another reason is no
  // assertion, so the shape was replaced rather than the expectation.
  Oper =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TRec = record'#13#10 +                                // 3
    '    procedure Go(const Title: string'#13#10 +           // 4
    '  end;'#13#10 +                                         // 5
    'implementation'#13#10 +                                 // 6
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
begin
  // the nested routine is declared ONCE - it must never pair with the
  // interface declaration of the unrelated routine of the same name
  Plan := PlanParameterRename(L(Nested), 8, 'Title');
  Assert.AreEqual('ppSingleHeader', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a nested routine has no second header');
  Assert.AreEqual(1, Integer(Length(Plan.Positions)), 'so exactly one position');
  Assert.AreEqual(8, Plan.Positions[0].Y, 'its own header');

  // the top-level routine of the same name still pairs correctly
  Plan := PlanParameterRename(L(Nested), 2, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the interface declaration pairs with the body at 4');
  Assert.AreEqual(4, Plan.Positions[High(Plan.Positions)].Y, 'the body at column 1');

  // an interface method has no implementation ANYWHERE - that is not a
  // failure to find one
  Plan := PlanParameterRename(L(Single_), 4, 'Title');
  Assert.AreEqual('ppSingleHeader', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'an interface method is declared once');
  Plan := PlanParameterRename(L(Single_), 7, 'Title');
  Assert.AreEqual('ppSingleHeader', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'and so is an abstract one');

  // ...while a header this does not understand is a REFUSAL, not a rename
  // of half the places
  Plan := PlanParameterRename(L(Oper), 4, 'Title');
  Assert.AreEqual('ppUnknown', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a header that cannot be read is a refusal, not a rename of half the '
    + 'places');
  Assert.AreNotEqual('', Plan.Reason, 'and it says why');
end;

{ TParameterRenameShapeTests }

procedure TParameterRenameShapeTests.AMemberAfterANestedTypeIsPaired;
const
  // #54.A, his fixture verbatim
  NestedIntf =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TOuter = class'#13#10 +                               // 3
    '  public type'#13#10 +                                  // 4
    '    IInner = interface'#13#10 +                         // 5
    '      procedure X;'#13#10 +                             // 6
    '    end;'#13#10 +                                       // 7
    '  public'#13#10 +                                       // 8
    '    procedure Go(A: Integer);'#13#10 +                  // 9
    '  end;'#13#10 +                                         // 10
    'implementation'#13#10 +                                 // 11
    'procedure TOuter.Go(A: Integer);'#13#10 +               // 12
    'begin'#13#10 +                                          // 13
    '  if A > 0 then ;'#13#10 +                              // 14
    'end;'#13#10 +                                           // 15
    'end.'#13#10;
  // #54.F - a nested RECORD, which was refused
  NestedRec =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TOuter = class'#13#10 +
    '  public type'#13#10 +
    '    TInner = record'#13#10 +
    '      V: Integer;'#13#10 +
    '    end;'#13#10 +
    '  public'#13#10 +
    '    procedure Go(A: Integer);'#13#10 +                  // 9
    '  end;'#13#10 +
    'implementation'#13#10 +
    'procedure TOuter.Go(A: Integer);'#13#10 +               // 12
    'begin'#13#10 +
    '  if A > 0 then ;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
  // #54.G - a method OF a nested class, and a SECOND outer class with a
  // nested type of the same name: the qualifier names the whole path, so
  // the two must not be mixed up
  Nested2 =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TOuterA = class'#13#10 +                              // 3
    '  public type'#13#10 +                                  // 4
    '    TInner = class'#13#10 +                             // 5
    '      procedure Go(A: Integer);'#13#10 +                // 6
    '    end;'#13#10 +                                       // 7
    '  end;'#13#10 +                                         // 8
    '  TOuterB = class'#13#10 +                              // 9
    '  public type'#13#10 +                                  // 10
    '    TInner = class'#13#10 +                             // 11
    '      procedure Go(A: Integer);'#13#10 +                // 12
    '    end;'#13#10 +                                       // 13
    '  end;'#13#10 +                                         // 14
    'implementation'#13#10 +                                 // 15
    'procedure TOuterA.TInner.Go(A: Integer);'#13#10 +       // 16
    'begin'#13#10 +                                          // 17
    'end;'#13#10 +                                           // 18
    'procedure TOuterB.TInner.Go(A: Integer);'#13#10 +       // 19
    'begin'#13#10 +                                          // 20
    'end;'#13#10 +                                           // 21
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
begin
  Plan := PlanParameterRename(L(NestedIntf), 9, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the nested interface is nested, not the owner - the reported case');
  Assert.AreEqual(2, Integer(Length(Plan.Positions)), 'both headers');
  Assert.AreEqual(9, Plan.Positions[0].Y, 'the declaration');
  Assert.AreEqual(12, Plan.Positions[1].Y, 'and the implementation header');
  // ...and the same answer from the other side
  Plan := PlanParameterRename(L(NestedIntf), 12, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'started on the implementation header');

  Plan := PlanParameterRename(L(NestedRec), 9, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a nested record is no owner either');

  Plan := PlanParameterRename(L(Nested2), 12, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a method OF a nested class pairs with its own implementation');
  Assert.AreEqual(19, Plan.Positions[High(Plan.Positions)].Y,
    'TOuterB.TInner.Go - not TOuterA''s, which carries the same names');
  // and from the implementation side the qualifier decides, not the
  // innermost name: 'TOuterA.TInner.Go' may only pair with TOuterA's
  Plan := PlanParameterRename(L(Nested2), 16, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'and back again');
  Assert.AreEqual(6, Plan.Positions[0].Y, 'TOuterA''s declaration');
  // ...and the SECOND one must find its own, which the innermost type NAME
  // cannot say: pairing it with TOuterA's renames a foreign class, and
  // refusing that candidate leaves "one header", which renames half of the
  // two names Delphi requires to be equal.
  Plan := PlanParameterRename(L(Nested2), 19, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'TOuterB.TInner.Go has a declaration of its own');
  Assert.AreEqual(12, Plan.Positions[0].Y,
    'TOuterB''s - not TOuterA''s, whose nested type carries the same name');
end;

procedure TParameterRenameShapeTests.AParameterOfItsOwnNameLeavesTheCallersAlone;
const
  // #54.B, his fixture verbatim
  SelfNamed =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Title(const Title: string);'#13#10 +          // 2
    'procedure Caller;'#13#10 +                              // 3
    'implementation'#13#10 +                                 // 4
    'procedure Title(const Title: string);'#13#10 +          // 5
    'begin'#13#10 +                                          // 6
    '  if Title = '''' then ;'#13#10 +                       // 7
    'end;'#13#10 +                                           // 8
    'procedure Caller;'#13#10 +                              // 9
    'begin'#13#10 +                                          // 10
    '  Title(''x'');'#13#10 +                                // 11
    'end;'#13#10 +                                           // 12
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Why: string;
begin
  Plan := PlanParameterRename(L(SelfNamed), 2, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'both headers, as before - the defect was downstream');
  Assert.IsTrue(Plan.ScopeKnown, 'and the routine''s body is delimited');
  Assert.AreEqual(5, Plan.BodyFirst, 'from its implementation header');
  Assert.AreEqual(8, Plan.BodyLast, 'to its end');

  // THE REPORTED DAMAGE: the call resolves to the implementation header's
  // LINE, which is one of the rename's targets, so it was renamed with the
  // parameter - E2003, measured
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 11, 2, Why))),
    'the CALL Title(''x'') is not the parameter');
  Assert.AreNotEqual('', Why, 'and the drop says why');
  // what must still be renamed
  Assert.AreEqual('poParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 7, 5, Why))),
    'the use INSIDE the body is the parameter');
  Assert.AreEqual('poParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 2, 22, Why))),
    'and so is the declaration itself');
  // ...while the routine's own name on that very line is not
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 2, 10, Why))),
    'the routine''s own name on the header line is not the parameter');
end;

procedure TParameterRenameShapeTests.ABodyOnTheHeaderLineIsNotTheHeader;
const
  // #54.C, his fixture verbatim
  OneLine =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Go(A: Integer);'#13#10 +                      // 2
    'implementation'#13#10 +                                 // 3
    'procedure Go(A: Integer); begin if A > 0 then ; end;'#13#10 + // 4
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Why: string;
begin
  Plan := PlanParameterRename(L(OneLine), 2, 'A');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'both headers');
  Assert.AreEqual(2, Integer(Length(Plan.Spans)), 'and both spans');
  // the ';' at column 24 ends the signature - everything after it is the
  // BODY and must not be read as "on the header"
  Assert.AreEqual(24, Plan.Spans[1].EndCol, 'the span ends at the '';''');
  Assert.AreEqual('poParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 4, 13, Why))),
    'the parameter itself');
  // THE REPORTED DAMAGE: this was "on the header but not a position" and
  // was dropped, so the body kept the old name - E2003
  Assert.AreNotEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 4, 35, Why))),
    'the body''s use is NOT on the header - it must not be dropped');
end;

procedure TParameterRenameShapeTests.TheCaretColumnPicksTheHeader;
const
  // #54.D, his fixture verbatim
  Two =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TFoo = class'#13#10 +                                 // 3
    '    procedure SetA(A: Integer); procedure SetB(A: Integer);'#13#10 + // 4
    '  end;'#13#10 +                                         // 5
    'implementation'#13#10 +                                 // 6
    'procedure TFoo.SetA(A: Integer);'#13#10 +               // 7
    'begin'#13#10 +                                          // 8
    '  if A > 0 then ;'#13#10 +                              // 9
    'end;'#13#10 +                                           // 10
    'procedure TFoo.SetB(A: Integer);'#13#10 +               // 11
    'begin'#13#10 +                                          // 12
    '  if A > 0 then ;'#13#10 +                              // 13
    'end;'#13#10 +                                           // 14
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Why: string;
begin
  // the caret on SetB's parameter (column 47)
  Plan := PlanParameterRename(L(Two), 4, 'A', 47);
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'SetB''s parameter');
  Assert.AreEqual(47, Plan.Positions[0].X, 'the one under the caret');
  Assert.AreEqual(11, Plan.Positions[1].Y, 'and SetB''s implementation header');
  // ...and SetA's parameter on the SAME line is not it. Without this the
  // rename changed three places of the wrong routine (issue #54.D).
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 4, 19, Why))),
    'SetA''s parameter is another symbol');

  // the caret on SetA's parameter picks the OTHER header
  Plan := PlanParameterRename(L(Two), 4, 'A', 19);
  Assert.AreEqual(19, Plan.Positions[0].X, 'SetA''s parameter');
  Assert.AreEqual(7, Plan.Positions[1].Y, 'and SetA''s implementation header');

  // and without a column the line is ambiguous - refuse, do not guess
  Plan := PlanParameterRename(L(Two), 4, 'A');
  Assert.AreEqual('ppUnknown', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'two headers on one line and no column: refuse');
  Assert.IsTrue(Plan.Reason.Contains('2 routine headers'),
    'and the reason says how many: ' + Plan.Reason);
end;

procedure TParameterRenameShapeTests.AnOperatorAndAnOmittedListArePaired;
const
  // #54.E - his fixture, with the implementation it needs to pair
  Oper =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TRec = record'#13#10 +                                // 3
    '    class operator Add(const Title: TRec; B: TRec): TRec;'#13#10 + // 4
    '  end;'#13#10 +                                         // 5
    'implementation'#13#10 +                                 // 6
    'class operator TRec.Add(const Title: TRec; B: TRec): TRec;'#13#10 + // 7
    'begin'#13#10 +                                          // 8
    '  Result := Title;'#13#10 +                             // 9
    'end;'#13#10 +                                           // 10
    'end.'#13#10;
  // #54.H - the implementation header leaves the list out, which is legal
  NoList =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TFoo = class'#13#10 +                                 // 3
    '    procedure Go(A: Integer);'#13#10 +                  // 4
    '  end;'#13#10 +                                         // 5
    'implementation'#13#10 +                                 // 6
    'procedure TFoo.Go;'#13#10 +                             // 7
    'begin'#13#10 +                                          // 8
    '  if A > 0 then ;'#13#10 +                              // 9
    'end;'#13#10 +                                           // 10
    'end.'#13#10;
  // ...but a parameterless OVERLOAD is not the other header of this one
  Overl =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TFoo = class'#13#10 +                                 // 3
    '    procedure Go(A: Integer); overload;'#13#10 +         // 4
    '    procedure Go; overload;'#13#10 +                     // 5
    '  end;'#13#10 +                                         // 6
    'implementation'#13#10 +                                 // 7
    'procedure TFoo.Go;'#13#10 +                             // 8
    'begin'#13#10 +                                          // 9
    'end;'#13#10 +                                           // 10
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Why: string;
begin
  Plan := PlanParameterRename(L(Oper), 4, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'an operator header is read like any other');
  Assert.AreEqual(2, Integer(Length(Plan.Positions)), 'both headers');
  Assert.AreEqual(7, Plan.Positions[1].Y, 'the implementation at column 1');

  Plan := PlanParameterRename(L(NoList), 4, 'A');
  Assert.AreEqual('ppSingleHeader', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the parameter is declared ONCE - the other header carries no list');
  Assert.AreEqual(1, Integer(Length(Plan.Positions)), 'so one position');
  Assert.AreEqual(2, Integer(Length(Plan.Spans)),
    'but both headers are spans - nothing on the listless one is the parameter');
  Assert.IsTrue(Plan.ScopeKnown, 'and its body is where the parameter is used');
  Assert.AreEqual('poParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 9, 5, Why))),
    'the body''s use is the parameter');
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 7, 15, Why))),
    'the routine''s own name on the listless header is not');

  // Delphi forbids omitting the list for an overload, so the parameterless
  // sibling must not be taken for the other header of this one
  Plan := PlanParameterRename(L(Overl), 4, 'A');
  Assert.AreEqual('ppSingleHeader', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the body of the OTHER overload is not this one''s header');
  Assert.AreEqual(1, Integer(Length(Plan.Positions)), 'one position');
  Assert.AreEqual(4, Plan.Positions[0].Y, 'its own declaration');
end;

procedure TParameterRenameShapeTests.ATypeDeclaredInTwoBranchesIsOneType;
const
  // Vcl.Controls' TControl / TWinControl and Vcl.ActnMan's
  // TActionDragBaseClass have exactly this shape
  Src =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '{$IF Defined(CLR)}'#13#10 +                             // 3
    '  TBase = class(TA)'#13#10 +                            // 4
    '{$ELSE}'#13#10 +                                        // 5
    '  TBase = class(TB)'#13#10 +                            // 6
    '{$ENDIF}'#13#10 +                                       // 7
    '    procedure Go(A: Integer);'#13#10 +                   // 8
    '  end;'#13#10 +                                         // 9
    '  TOther = class'#13#10 +                               // 10
    '    procedure Also(B: Integer);'#13#10 +                 // 11
    '  end;'#13#10 +                                         // 12
    'procedure Top(C: Integer);'#13#10 +                      // 13
    'implementation'#13#10 +                                 // 14
    'end.'#13#10;
var
  Kw: string;
begin
  var Lines := L(Src);
  // the member of the type that is declared twice
  Assert.AreEqual('TBase', EnclosingContainerName(Lines, 8),
    'two branches are ONE type, not a type nested in itself');
  // the SIBLING below it - the extra opener used to make TBase its outer
  // level
  Assert.AreEqual('TOther', EnclosingContainerName(Lines, 11),
    'the next type is a sibling, not a nested one');
  // ...and a top-level routine has no container at all
  Assert.AreEqual('', EnclosingContainerName(Lines, 13),
    'a routine declared beside the types is in none of them');
  // the path answers the same, with the keyword the filter above needs
  Assert.AreEqual('TBase', EnclosingTypePath(MaskCommentsAndStrings(Lines), 8, Kw),
    'the path says the same');
  Assert.AreEqual('class', Kw, 'and what kind of body it is');
end;

procedure TParameterRenameAuditTests.TheHeaderShapesThatUsedToFail;
const
  // a class declared in the IMPLEMENTATION section: the branch must be
  // chosen by "inside a type", not by the section (#50.3)
  Local =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'implementation'#13#10 +                                 // 2
    'type'#13#10 +                                           // 3
    '  TLocal = class'#13#10 +                               // 4
    '    procedure Go(const Title: string);'#13#10 +          // 5
    '  end;'#13#10 +                                         // 6
    'procedure TLocal.Go(const Title: string);'#13#10 +      // 7
    'begin'#13#10 +                                          // 8
    'end;'#13#10 +                                           // 9
    'end.'#13#10;
  // a { } comment inside the parameter list, with a bracket in it
  Comment =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Go(const Title: string { the name ( of it });'#13#10 + // 2
    'implementation'#13#10 +                                 // 3
    'procedure Go(const Title: string);'#13#10 +             // 4
    'begin'#13#10 +                                          // 5
    'end;'#13#10 +                                           // 6
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Long: string;
  I: Integer;
begin
  Plan := PlanParameterRename(L(Local), 5, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a class in the implementation section pairs with its own body');
  Assert.AreEqual(7, Plan.Positions[High(Plan.Positions)].Y, 'the body at line 7');

  Plan := PlanParameterRename(L(Comment), 2, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a comment in the list is not code: ' + Plan.Reason);

  // a header longer than the twelve lines CollectHeader reads
  Long := 'unit U;'#13#10'interface'#13#10'procedure Go('#13#10;
  for I := 1 to 14 do
    Long := Long + '  const P' + IntToStr(I) + ': string;'#13#10;
  Long := Long + '  const Title: string);'#13#10 +
    'implementation'#13#10'procedure Go('#13#10;
  for I := 1 to 14 do
    Long := Long + '  const P' + IntToStr(I) + ': string;'#13#10;
  Long := Long + '  const Title: string);'#13#10'begin'#13#10'end;'#13#10'end.'#13#10;
  Plan := PlanParameterRename(L(Long), 2, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'a header of 16 lines is still a header: ' + Plan.Reason);
  Assert.AreEqual(17, Plan.Positions[0].Y, 'the name stands on the 15th line of it');

  // "const [ref] A: T" - the attribute comes AFTER the modifier
  var Ref := ParseParamList('const [ref] Title: TRec; var Other: Integer');
  Assert.AreEqual(2, Integer(Length(Ref)), 'two parameters');
  Assert.AreEqual('Title', Ref[0].Name, 'the attribute is not part of the name');
  Assert.AreEqual('TRec', Ref[0].TypeText, 'and the type is read');
end;

procedure TParameterRenameAuditTests.OnlyTheParameterItselfIsAPosition;
const
  // the routine is named like its parameter, a type is spelled like it, and
  // a default value names a constant of that name (#50.4)
  Src =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Title(const Title: string);'#13#10 +          // 2
    'procedure Other(const A: Title);'#13#10 +               // 3
    'procedure Third(const X: Integer = Title);'#13#10 +     // 4
    'implementation'#13#10 +                                 // 5
    'procedure Title(const Title: string);'#13#10 +          // 6
    'begin'#13#10 +                                          // 7
    'end;'#13#10 +                                           // 8
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
begin
  Plan := PlanParameterRename(L(Src), 2, 'Title');
  Assert.AreEqual('ppPaired', GetEnumName(TypeInfo(TParamPairing), Ord(Plan.Pairing)),
    'the parameter pairs with the body''s own header');
  Assert.AreEqual(2, Integer(Length(Plan.Positions)), 'exactly two places declare it');
  // "procedure Title(const Title: string);" - the parameter stands at 23,
  // the routine's own name at 10, and that one must not be renamed with it
  Assert.AreEqual(22, Plan.Positions[0].X, 'the name INSIDE the list, 0-based');
  Assert.AreEqual(2, Plan.Positions[0].Y, 'on the declaration line');
  Assert.AreEqual(22, Plan.Positions[1].X, 'and the same place in the body''s header');

  // a TYPE spelled like the parameter is no parameter of that name
  Assert.IsFalse(HeaderDeclaresParameter(L(Src), 3, 'Title'),
    'a type named like it does not declare it');
  // ...and neither does a default value naming a constant
  Assert.IsFalse(HeaderDeclaresParameter(L(Src), 4, 'Title'),
    'a default value naming it does not declare it');
end;

procedure TParameterRenameShapeTests.ACallFromAnotherUnitIsNotTheParameterEither;
const
  // his two-unit fixture's declaring half, verbatim
  OwnUnit =
    'unit A;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Title(const Title: string);'#13#10 +          // 2
    'implementation'#13#10 +                                 // 3
    'procedure Title(const Title: string);'#13#10 +          // 4
    'begin'#13#10 +                                          // 5
    '  if Title = '''' then ;'#13#10 +                       // 6
    'end;'#13#10 +                                           // 7
    'end.'#13#10;
  // ...and the one case in which the body's text is NOT wholly here
  WithInclude =
    'unit A;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Title(const Title: string);'#13#10 +          // 2
    'implementation'#13#10 +                                 // 3
    'procedure Title(const Title: string);'#13#10 +          // 4
    'begin'#13#10 +                                          // 5
    '  {$I TitleBody.inc}'#13#10 +                           // 6
    'end;'#13#10 +                                           // 7
    'end.'#13#10;
  // {$I+} is the IO-check switch and no include at all - the directive the
  // rule must NOT fall for
  WithSwitch =
    'unit A;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'procedure Title(const Title: string);'#13#10 +          // 2
    'implementation'#13#10 +                                 // 3
    'procedure Title(const Title: string);'#13#10 +          // 4
    'begin'#13#10 +                                          // 5
    '  {$I+} if Title = '''' then ;'#13#10 +                 // 6
    'end;'#13#10 +                                           // 7
    'end.'#13#10;
var
  Plan: TParamRenamePlan;
  Why: string;
begin
  Plan := PlanParameterRename(L(OwnUnit), 2, 'Title');
  Assert.IsTrue(Plan.ScopeKnown, 'the body is delimited');
  Assert.IsTrue(Plan.BodySelfContained,
    'and its text is wholly in this unit - no include in it');

  // THE REPORTED DAMAGE: unit B's "Title('x')" resolves to a header LINE of
  // THIS unit, so it matched; the rule had no opinion outside this file and
  // the call was renamed.
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, False, 13, 2, Why))),
    'a call from ANOTHER unit is not the parameter');
  Assert.IsTrue(Why.Contains('another file'), 'and the drop says why: ' + Why);
  // the line number there means nothing - any position in another file is
  // answered the same way, including one that would be inside the body range
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, False, 6, 5, Why))),
    'a line INSIDE this unit''s body range is not the parameter elsewhere');
  // ...while this file is judged exactly as before
  Assert.AreEqual('poParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 6, 5, Why))),
    'the body''s own use is still the parameter');

  // An {$I} in the body: the scope reaches into a file this plan never read,
  // so nothing outside this one is decided - the occurrence is KEPT.
  Plan := PlanParameterRename(L(WithInclude), 2, 'Title');
  Assert.IsTrue(Plan.ScopeKnown, 'the body is still delimited');
  Assert.IsFalse(Plan.BodySelfContained,
    'but its text is not wholly here');
  Assert.AreEqual('poUndecided',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, False, 13, 2, Why))),
    'so another file is not decided, and nothing there is dropped');
  Assert.AreEqual('poNotTheParameter',
    GetEnumName(TypeInfo(TParamOccurrence),
      Ord(ClassifyParamOccurrence(Plan, True, 2, 10, Why))),
    'while THIS file is judged as before - the routine''s own name');

  // and the switch that begins like an include is none
  Plan := PlanParameterRename(L(WithSwitch), 2, 'Title');
  Assert.IsTrue(Plan.BodySelfContained,
    '{$I+} is the IO-check switch, not an include');
end;

initialization
  TDUnitX.RegisterTestFixture(TParameterRenameAuditTests);
  TDUnitX.RegisterTestFixture(TParameterRenameShapeTests);

end.
