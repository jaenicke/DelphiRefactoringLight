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

implementation

uses
  System.SysUtils, System.Types, System.Classes, System.TypInfo,
  Expert.SafeDeletePlan, Expert.AutoImport, Expert.SignatureEdit;

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
  // a shape nothing here can read: an operator header (#50.3)
  Oper =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TRec = record'#13#10 +                                // 3
    '    class operator Add(const Title: TRec; B: TRec): TRec;'#13#10 + // 4
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
    'an operator header is not read here - refuse instead of guessing');
  Assert.AreNotEqual('', Plan.Reason, 'and it says why');
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

initialization
  TDUnitX.RegisterTestFixture(TParameterRenameAuditTests);

end.
