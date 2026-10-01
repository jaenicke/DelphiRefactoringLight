(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  "Move to unit": which implementation lines are taken along with the
///  symbol, and where they land in the target. The pure halves only
///  (LocateMoveImplementation, SpliceMoveIntoTarget).
/// </summary>
unit Test.MoveToUnitBody;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TMoveToUnitBodyTests = class
  public
    /// <summary>Audit C4: the body ran to the first 'end' at depth 0, so the
    ///  'end' of a try, case or asm block (or of a nested routine) cut it.</summary>
    [Test] procedure RoutineBody_RunsToItsOwnEnd;
    /// <summary>Audit C4, the same counter in the class-method locator.</summary>
    [Test] procedure ClassMethodBody_RunsToItsOwnEnd;
    /// <summary>Audit H28: "TList2&lt;T&gt;.Add" was never matched.</summary>
    [Test] procedure GenericClassMethods_AreFound;
    /// <summary>Audit H26: the blocks were inserted before the final 'end.',
    ///  i.e. inside an initialization / finalization section.</summary>
    [Test] procedure Splice_GoesBeforeInitialization;
    /// <summary>An asm LABEL named @@end is not the block's end.</summary>
    [Test] procedure AsmLabelNamedEnd_IsNotTheBlocksEnd;
  end;

implementation

uses
  System.SysUtils, Expert.MoveToUnit;

{ TMoveToUnitBodyTests }

procedure TMoveToUnitBodyTests.RoutineBody_RunsToItsOwnEnd;
const
  Src =
    'unit U;'#13#10 +                      // 1
    'interface'#13#10 +                    // 2
    'procedure Run;'#13#10 +               // 3
    'procedure Outer;'#13#10 +             // 4
    'function Fast: Integer;'#13#10 +      // 5
    'implementation'#13#10 +               // 6
    'procedure Run;'#13#10 +               // 7
    'var'#13#10 +                          // 8
    '  I: Integer;'#13#10 +                // 9
    'begin'#13#10 +                        // 10
    '  try'#13#10 +                        // 11
    '    case I of'#13#10 +                // 12
    '      0: I := 1;'#13#10 +             // 13
    '    end;'#13#10 +                     // 14
    '  finally'#13#10 +                    // 15
    '    I := 0;'#13#10 +                  // 16
    '  end;'#13#10 +                       // 17
    '  I := 2;'#13#10 +                    // 18
    'end;'#13#10 +                         // 19
    'procedure Outer;'#13#10 +             // 20
    'type'#13#10 +                         // 21
    '  TRec = record'#13#10 +              // 22
    '    case Integer of'#13#10 +          // 23
    '      0: (A: Integer);'#13#10 +       // 24
    '  end;'#13#10 +                       // 25
    '  procedure Inner;'#13#10 +           // 26
    '  begin'#13#10 +                      // 27
    '  end;'#13#10 +                       // 28
    'begin'#13#10 +                        // 29
    '  Inner;'#13#10 +                     // 30
    'end;'#13#10 +                         // 31
    'function Fast: Integer;'#13#10 +      // 32
    'asm'#13#10 +                          // 33
    '  MOV EAX, 1'#13#10 +                 // 34
    'end;'#13#10 +                         // 35
    'procedure Other;'#13#10 +             // 36
    'begin'#13#10 +                        // 37
    'end;'#13#10 +                         // 38
    'end.';
var
  S, E: TArray<Integer>;
begin
  Assert.IsTrue(LocateMoveImplementation('Run', Src, S, E), 'found: Run');
  Assert.AreEqual(7, S[0]);
  Assert.AreEqual(19, E[0], 'try/case end their own blocks, not the routine');
  Assert.IsTrue(LocateMoveImplementation('Outer', Src, S, E), 'found: Outer');
  Assert.AreEqual(20, S[0]);
  Assert.AreEqual(31, E[0], 'local record and nested routine do not end the routine');
  Assert.IsTrue(LocateMoveImplementation('Fast', Src, S, E), 'found: Fast');
  Assert.AreEqual(32, S[0]);
  Assert.AreEqual(35, E[0], 'an asm body ends at its own end');
end;

procedure TMoveToUnitBodyTests.ClassMethodBody_RunsToItsOwnEnd;
const
  Src =
    'unit U;'#13#10 +                      // 1
    'interface'#13#10 +                    // 2
    'type'#13#10 +                         // 3
    '  TFoo = class'#13#10 +               // 4
    '    procedure X;'#13#10 +             // 5
    '    procedure Y;'#13#10 +             // 6
    '  end;'#13#10 +                       // 7
    'implementation'#13#10 +               // 8
    'procedure TFoo.X;'#13#10 +            // 9
    'begin'#13#10 +                        // 10
    '  try'#13#10 +                        // 11
    '  finally'#13#10 +                    // 12
    '  end;'#13#10 +                       // 13
    '  Y;'#13#10 +                         // 14
    'end;'#13#10 +                         // 15
    'procedure TFoo.Y;'#13#10 +            // 16
    'begin'#13#10 +                        // 17
    'end;'#13#10 +                         // 18
    'end.';
var
  S, E: TArray<Integer>;
begin
  Assert.IsTrue(LocateMoveImplementation('TFoo', Src, S, E), 'found: TFoo');
  Assert.AreEqual(2, Integer(Length(S)), 'both methods');
  Assert.AreEqual(9, S[0]);
  Assert.AreEqual(15, E[0], 'the try block does not end TFoo.X');
  Assert.AreEqual(16, S[1]);
  Assert.AreEqual(18, E[1]);
end;

procedure TMoveToUnitBodyTests.GenericClassMethods_AreFound;
const
  Src =
    'unit U;'#13#10 +                      // 1
    'interface'#13#10 +                    // 2
    'type'#13#10 +                         // 3
    '  TList2<T> = class'#13#10 +          // 4
    '    procedure Add(const A: T);'#13#10 + // 5
    '    function Count: Integer;'#13#10 + // 6
    '  end;'#13#10 +                       // 7
    'implementation'#13#10 +               // 8
    'procedure TList2<T>.Add(const A: T);'#13#10 + // 9
    'begin'#13#10 +                        // 10
    'end;'#13#10 +                         // 11
    'function TList2<T>.Count: Integer;'#13#10 + // 12
    'begin'#13#10 +                        // 13
    '  Result := 0;'#13#10 +               // 14
    'end;'#13#10 +                         // 15
    'end.';
var
  S, E: TArray<Integer>;
begin
  Assert.IsTrue(LocateMoveImplementation('TList2', Src, S, E),
    'the methods of a generic class are found');
  Assert.AreEqual(2, Integer(Length(S)), 'both methods');
  Assert.AreEqual(9, S[0]);
  Assert.AreEqual(11, E[0]);
  Assert.AreEqual(12, S[1]);
  Assert.AreEqual(15, E[1]);
end;

procedure TMoveToUnitBodyTests.AsmLabelNamedEnd_IsNotTheBlocksEnd;
const
  Src =
    'unit U;'#13#10 +                      // 1
    'interface'#13#10 +                    // 2
    'function Fast: Integer;'#13#10 +      // 3
    'implementation'#13#10 +               // 4
    'function Fast: Integer;'#13#10 +      // 5
    'asm'#13#10 +                          // 6
    '  CMP EAX, 0'#13#10 +                 // 7
    '  JZ  @@end'#13#10 +                  // 8
    '  MOV EAX, 1'#13#10 +                 // 9
    '@@end:'#13#10 +                       // 10
    'end;'#13#10 +                         // 11
    'procedure Other;'#13#10 +             // 12
    'begin'#13#10 +                        // 13
    'end;'#13#10 +                         // 14
    'end.';
var
  S, E: TArray<Integer>;
begin
  Assert.IsTrue(LocateMoveImplementation('Fast', Src, S, E), 'found: Fast');
  Assert.AreEqual(5, S[0]);
  Assert.AreEqual(11, E[0], 'a label @@end is not the asm block''s end');
end;

procedure TMoveToUnitBodyTests.Splice_GoesBeforeInitialization;
const
  Moved = 'procedure Moved;'#13#10'begin'#13#10'end;';
  procedure Check(const ATarget, ASection: string);
  var
    Plan: TMovePlan;
    R: string;
  begin
    Plan := Default(TMovePlan);
    Plan.Kind := mskRoutine;
    Plan.DeclarationText := 'procedure Moved;';
    Plan.ImplBlocks := [Moved];
    R := SpliceMoveIntoTarget(ATarget, Plan);
    Assert.IsTrue(Pos(Moved, R) > 0, ASection + ': the body is inserted');
    Assert.IsTrue(Pos(Moved, R) < Pos(ASection, R),
      ASection + ': the body lands before the section, not inside it');
    Assert.IsTrue(Pos('procedure Moved;', R) < Pos('implementation', R),
      ASection + ': the declaration goes into the interface');
  end;
begin
  Check('unit T;'#13#10'interface'#13#10'implementation'#13#10 +
    'initialization'#13#10'  Exit;'#13#10'end.', 'initialization');
  Check('unit T;'#13#10'interface'#13#10'implementation'#13#10 +
    'finalization'#13#10'  Exit;'#13#10'end.', 'finalization');
  // no section: still before the final end.
  Check('unit T;'#13#10'interface'#13#10'implementation'#13#10'end.', 'end.');
end;

initialization
  TDUnitX.RegisterTestFixture(TMoveToUnitBodyTests);

end.
