(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Extract method (audit #22): the token positions, the Result promotion,
///  the var removal in the enclosing routine, and the selection validator
///  on complete try..end / case..end blocks. The pure halves only.
/// </summary>
unit Test.ExtractMethod;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TExtractMethodTests = class
  public
    [Test] procedure Tokens_FirstLineIsOffsetByTheStartColumn;
    [Test] procedure ResultPromotion_OnlyWhenTheOldValueIsNotRead;
    [Test] procedure VarRemoval_TouchesOnlyTheEnclosingRoutine;
    [Test] procedure VarRemoval_EmptiedSectionLosesItsKeyword;
    [Test] procedure Validator_CompleteTryBlockIsAccepted;
    [Test] procedure Validator_CompleteCaseAndAsmBlocksAreAccepted;
    [Test] procedure Validator_StrayEndIsStillRefused;
    [Test] procedure Validator_ElseAfterAClosedCaseIsAnOrphan;
  end;

implementation

uses
  System.SysUtils, System.Generics.Collections, Expert.ExtractMethod,
  Expert.SelectionValidator;

const
  // a unit variable X and three routines that each declare their own X;
  // Run (line 11) is the enclosing routine
  VarUnit: array[0..20] of string = (
    'unit U;',                // 1
    'interface',              // 2
    'implementation',         // 3
    'var',                    // 4
    '  X: Integer;',          // 5  unit variable
    'procedure Other;',       // 6
    'var',                    // 7
    '  X, Y: Integer;',       // 8  another routine's
    'begin',                  // 9
    'end;',                   // 10
    'procedure Run;',         // 11
    'var',                    // 12
    '  X: Integer;',          // 13
    '  Z, X2: string;',       // 14
    'begin',                  // 15
    '  X := 1;',              // 16
    'end;',                   // 17
    'procedure After;',       // 18
    'var',                    // 19
    '  X: Integer;',          // 20
    'begin end;');            // 21

function Lines(const AItems: array of string): TArray<string>;
begin
  SetLength(Result, Length(AItems));
  for var I := 0 to High(AItems) do Result[I] := AItems[I];
end;

// no file context: only the block's own structure is checked
function Check(const ABlock: string): TValidationResult;
begin
  Result := TSelectionValidator.Validate(ABlock, nil, 10, 12, 1, '');
end;

procedure TExtractMethodTests.Tokens_FirstLineIsOffsetByTheStartColumn;
begin
  // the selection starts at column 5 of line 10
  var T := ExtractBlockTokens('Total := Extra;'#13#10'  Foo(Total, Count);', 10, 5);
  Assert.AreEqual<Integer>(4, Length(T), 'Total, Extra, Foo, Count');
  Assert.AreEqual('Total', T[0].Ident, False);
  Assert.AreEqual(9, T[0].Line0);
  Assert.AreEqual(4, T[0].Col0, 'column 5 of the file line, 0-based');
  Assert.AreEqual('Extra', T[1].Ident, False);
  Assert.AreEqual(13, T[1].Col0);
  Assert.AreEqual('Foo', T[2].Ident, False);
  Assert.AreEqual(10, T[2].Line0);
  Assert.AreEqual(2, T[2].Col0, 'later lines start at column 1');
end;

procedure TExtractMethodTests.ResultPromotion_OnlyWhenTheOldValueIsNotRead;
begin
  Assert.IsFalse(CanReturnViaResult('Total := Total + Extra;', 'Total'),
    'the reported case: the old value is read');
  Assert.IsTrue(CanReturnViaResult('Total := Extra * 2;'#10'Show(Total);', 'Total'),
    'written first, then read: Result is defined');
  Assert.IsFalse(CanReturnViaResult('if C then'#10'  Total := 1;', 'Total'),
    'a conditional write leaves Result undefined');
  Assert.IsFalse(CanReturnViaResult('if C then Total := 1;', 'Total'),
    'the same on one line');
  Assert.IsFalse(CanReturnViaResult('Show(Total);'#10'Total := 1;', 'Total'),
    'read before it is written');
  Assert.IsTrue(CanReturnViaResult('// Total is set here'#10'Total := 1;', 'Total'),
    'a comment is no read');
end;

procedure TExtractMethodTests.VarRemoval_TouchesOnlyTheEnclosingRoutine;
begin
  var P := PlanLocalVarRemoval(Lines(VarUnit), 11, ['X']);
  Assert.AreEqual<Integer>(1, Length(P), 'exactly one line changes');
  Assert.AreEqual(13, P[0].Key, 'the line in Run');
  Assert.IsTrue(P[0].Value = #0, 'deleted');

  P := PlanLocalVarRemoval(Lines(VarUnit), 11, ['Z']);
  Assert.AreEqual<Integer>(1, Length(P));
  Assert.AreEqual(14, P[0].Key);
  Assert.AreEqual('  X2: string;', P[0].Value, False, 'the other name on the line stays');
end;

procedure TExtractMethodTests.VarRemoval_EmptiedSectionLosesItsKeyword;
begin
  var P := PlanLocalVarRemoval(Lines(VarUnit), 11, ['X', 'Z', 'X2']);
  Assert.AreEqual<Integer>(3, Length(P), 'the keyword and both declarations');
  Assert.AreEqual(12, P[0].Key, 'the var keyword goes');
  Assert.IsTrue(P[0].Value = #0, 'the keyword line is deleted');
  Assert.AreEqual(13, P[1].Key);
  Assert.AreEqual(14, P[2].Key);

  // "var X: Integer;" on the keyword line with a survivor after it: the
  // keyword stays for the survivor
  P := PlanLocalVarRemoval(Lines(['procedure Run;', 'var X: Integer;', '  Y: Integer;',
    'begin', 'end;']), 1, ['X']);
  Assert.AreEqual<Integer>(1, Length(P));
  Assert.AreEqual(2, P[0].Key);
  Assert.AreEqual('var', P[0].Value, False, 'the keyword stays for Y');
end;

procedure TExtractMethodTests.Validator_CompleteTryBlockIsAccepted;
begin
  var R := Check('try'#13#10'  A;'#13#10'finally'#13#10'  B;'#13#10'end;');
  Assert.IsFalse(R.HasErrors, R.FormatIssues);
  R := Check('begin'#13#10'  try A; except on E: Exception do B; end;'#13#10'end;');
  Assert.IsFalse(R.HasErrors, R.FormatIssues);
end;

procedure TExtractMethodTests.Validator_CompleteCaseAndAsmBlocksAreAccepted;
begin
  var R := Check('case X of'#13#10'  1: A;'#13#10'else'#13#10'  B;'#13#10'end;');
  Assert.IsFalse(R.HasErrors, R.FormatIssues);
  R := Check('asm'#13#10'  MOV EAX, 1'#13#10'end;');
  Assert.IsFalse(R.HasErrors, R.FormatIssues);
end;

procedure TExtractMethodTests.Validator_StrayEndIsStillRefused;
begin
  var R := Check('  A;'#13#10'end;');
  Assert.IsTrue(R.HasErrors, 'a stray end must still be refused');
  Assert.IsTrue(R.FormatIssues.Contains('"end" without matching'), R.FormatIssues);
end;

procedure TExtractMethodTests.Validator_ElseAfterAClosedCaseIsAnOrphan;
begin
  var R := Check('case X of'#13#10'  1: A;'#13#10'end;'#13#10'else'#13#10'  B;');
  Assert.IsTrue(R.FormatIssues.Contains('"else" without matching'), R.FormatIssues);
end;

initialization
  TDUnitX.RegisterTestFixture(TExtractMethodTests);

end.
