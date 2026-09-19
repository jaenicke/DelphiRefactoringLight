(*
 * Written by Ian Branch for his fork gllDelphiRefactoringLight and
 * contributed together with his code audit (GitHub issue #10); adapted
 * to this code base by Sebastian Jänicke.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Regression tests for TWithScanner - the pure, LSP-free half of the
///  with-refactoring pipeline.
///
///  The scanner deliberately resolves nothing; it only finds with-statements
///  and delimits their targets and bodies. That makes it fully testable
///  without a running DelphiLSP, and it is the layer everything else builds
///  on: if the scanner mis-delimits a body, every downstream decision about
///  which identifiers belong to which scope is made on bad input.
///
///  The nested-with cases matter most. Expert.WithRefactorWizard refuses to
///  rewrite an outer with whose body contains another with, and it decides
///  that by testing the inner KeywordPos against the outer BodyInnerRange.
///  Those tests pin that predicate.
/// </summary>
unit Test.WithScanner;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TWithScannerTests = class
  private
    /// <summary>Asserts a 1-based position equals the expected line/column.</summary>
    procedure CheckPos(const AWhat: string; ALine, ACol, AExpLine, AExpCol: Integer);
  public
    [Test] procedure ScanSource_NoWith_ReturnsEmpty;
    [Test] procedure ScanSource_SimpleBeginEnd_FindsOneWithCorrectTarget;
    [Test] procedure ScanSource_IsCaseInsensitive;
    [Test] procedure ScanSource_SingleStatementBody_IsWbkSingle;
    [Test] procedure ScanSource_CompoundTryBody_IsWbkCompound;
    [Test] procedure ScanSource_MultipleTargets_AreSplitOnTopLevelCommas;
    [Test] procedure ScanSource_CommasInsideCallParens_DoNotSplitTargets;
    [Test] procedure ScanSource_DottedTargetExpression_IsPreservedVerbatim;
    [Test] procedure ScanSource_WithInStringLiteral_IsIgnored;
    [Test] procedure ScanSource_WithInLineComment_IsIgnored;
    [Test] procedure ScanSource_WithInBraceComment_IsIgnored;
    [Test] procedure ScanSource_WithInParenStarComment_IsIgnored;
    [Test] procedure ScanSource_MalformedWithNoDo_IsSkippedWithoutRaising;
    [Test] procedure ScanSource_NestedWith_FindsBoth;
    [Test] procedure ScanSource_NestedWith_OuterBodyContainsInnerKeyword;
    [Test] procedure ScanSource_SiblingWiths_DoNotContainEachOther;
    [Test] procedure ScanSource_TwoSiblingBeginEndWiths_FindsBoth;
    [Test] procedure ScanSource_TwoSiblingSingleStatementWiths_FindsBoth;
  end;

implementation

uses
  System.SysUtils,
  Expert.WithScanner;

const
  /// <summary>Line break used when composing test sources. The scanner
  ///  counts lines on LF, so CRLF and LF both work; CRLF is used to match
  ///  what the IDE actually hands it.</summary>
  NL = #13#10;

/// <summary>Builds a source string from an array of lines.</summary>
function Src(const ALines: array of string): string;
var
  I: Integer;
begin
  Result := '';
  for I := Low(ALines) to High(ALines) do
    Result := Result + ALines[I] + NL;
end;

/// <summary>True when APos lies inside ARange (inclusive, 1-based). This is
///  the same containment predicate the wizard uses to detect a nested with,
///  reproduced here so the tests pin the behaviour rather than the code.</summary>
function PosInRange(const APos: TWithSourcePos; const ARange: TWithSourceRange): Boolean;
begin
  Result :=
    ((APos.Line > ARange.StartPos.Line)
      or ((APos.Line = ARange.StartPos.Line) and (APos.Col >= ARange.StartPos.Col)))
    and
    ((APos.Line < ARange.EndPos.Line)
      or ((APos.Line = ARange.EndPos.Line) and (APos.Col <= ARange.EndPos.Col)));
end;

{ TWithScannerTests }

procedure TWithScannerTests.CheckPos(const AWhat: string; ALine, ACol, AExpLine, AExpCol: Integer);
begin
  Assert.AreEqual(AExpLine, ALine, AWhat + ': line');
  Assert.AreEqual(AExpCol, ACol, AWhat + ': col');
end;

procedure TWithScannerTests.ScanSource_NoWith_ReturnsEmpty;
begin
  Assert.AreEqual<Integer>(0, Length(TWithScanner.ScanSource(Src([
    'procedure Foo;',
    'begin',
    '  Bar := 1;',
    'end;']))));
end;

procedure TWithScannerTests.ScanSource_SimpleBeginEnd_FindsOneWithCorrectTarget;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'procedure Foo;',        // 1
    'begin',                 // 2
    '  with FBar do',        // 3
    '    begin',             // 4
    '    Baz := 1;',         // 5
    '    end;',              // 6
    'end;']));               // 7

  Assert.AreEqual<Integer>(1, Length(Occs), 'occurrence count');
  Assert.AreEqual<Integer>(1, Length(Occs[0].Targets), 'target count');
  Assert.AreEqual('FBar', Occs[0].Targets[0].Expression, 'target expression');
  CheckPos('keyword', Occs[0].KeywordPos.Line, Occs[0].KeywordPos.Col, 3, 3);
  Assert.IsTrue(Occs[0].BodyKind = wbkBeginEnd, 'body kind should be begin..end');
end;

procedure TWithScannerTests.ScanSource_IsCaseInsensitive;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'WITH FBar DO',
    '  Baz := 1;']));
  Assert.AreEqual<Integer>(1, Length(Occs), 'uppercase WITH/DO should still be found');
  Assert.AreEqual('FBar', Occs[0].Targets[0].Expression);
end;

procedure TWithScannerTests.ScanSource_SingleStatementBody_IsWbkSingle;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'with FBar do',
    '  Baz := 1;']));
  Assert.AreEqual<Integer>(1, Length(Occs));
  Assert.IsTrue(Occs[0].BodyKind = wbkSingle, 'single statement body');
end;

procedure TWithScannerTests.ScanSource_CompoundTryBody_IsWbkCompound;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'with FBar do',
    '  try',
    '    Baz := 1;',
    '  finally',
    '    Quux;',
    '  end;']));
  Assert.AreEqual<Integer>(1, Length(Occs));
  Assert.IsTrue(Occs[0].BodyKind = wbkCompound, 'try..end is a compound body');
end;

procedure TWithScannerTests.ScanSource_MultipleTargets_AreSplitOnTopLevelCommas;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'with FBar, FBaz do',
    '  Quux := 1;']));
  Assert.AreEqual<Integer>(1, Length(Occs), 'one with-statement');
  Assert.AreEqual<Integer>(2, Length(Occs[0].Targets), 'two targets');
  Assert.AreEqual('FBar', Occs[0].Targets[0].Expression);
  Assert.AreEqual('FBaz', Occs[0].Targets[1].Expression);
end;

procedure TWithScannerTests.ScanSource_CommasInsideCallParens_DoNotSplitTargets;
var
  Occs: TArray<TWithOccurrence>;
begin
  // A comma inside the argument list is NOT a target separator. Getting this
  // wrong would turn one target into two and produce nonsense prefixes.
  Occs := TWithScanner.ScanSource(Src([
    'with GetThing(1, 2) do',
    '  Quux := 1;']));
  Assert.AreEqual<Integer>(1, Length(Occs));
  Assert.AreEqual<Integer>(1, Length(Occs[0].Targets), 'call arguments must not split the target');
  Assert.AreEqual('GetThing(1, 2)', Occs[0].Targets[0].Expression);
end;

procedure TWithScannerTests.ScanSource_DottedTargetExpression_IsPreservedVerbatim;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'with MainForm.SystemDatabaseQuery do',
    '  Open;']));
  Assert.AreEqual<Integer>(1, Length(Occs));
  Assert.AreEqual('MainForm.SystemDatabaseQuery', Occs[0].Targets[0].Expression);
end;

procedure TWithScannerTests.ScanSource_WithInStringLiteral_IsIgnored;
begin
  Assert.AreEqual<Integer>(0, Length(TWithScanner.ScanSource(Src([
    'begin',
    '  S := ''with FBar do'';',
    'end;']))), 'a with inside a string literal is not code');
end;

procedure TWithScannerTests.ScanSource_WithInLineComment_IsIgnored;
begin
  Assert.AreEqual<Integer>(0, Length(TWithScanner.ScanSource(Src([
    'begin',
    '  // with FBar do Baz := 1;',
    'end;']))), 'a with inside a // comment is not code');
end;

procedure TWithScannerTests.ScanSource_WithInBraceComment_IsIgnored;
begin
  Assert.AreEqual<Integer>(0, Length(TWithScanner.ScanSource(Src([
    'begin',
    '  { with FBar do Baz := 1; }',
    'end;']))), 'a with inside a { } comment is not code');
end;

procedure TWithScannerTests.ScanSource_WithInParenStarComment_IsIgnored;
begin
  Assert.AreEqual<Integer>(0, Length(TWithScanner.ScanSource(Src([
    'begin',
    '  (* with FBar do Baz := 1; *)',
    'end;']))), 'a with inside a (* *) comment is not code');
end;

procedure TWithScannerTests.ScanSource_MalformedWithNoDo_IsSkippedWithoutRaising;
var
  Occs: TArray<TWithOccurrence>;
begin
  // Documented contract: malformed with-statements are skipped silently and
  // the scanner never raises.
  Occs := TWithScanner.ScanSource(Src([
    'begin',
    '  with FBar',
    'end;']));
  Assert.AreEqual<Integer>(0, Length(Occs), 'a with without do yields no occurrence');
end;

procedure TWithScannerTests.ScanSource_NestedWith_FindsBoth;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(Src([
    'with FOuter do',        // 1
    '  begin',               // 2
    '  with FInner do',      // 3
    '    begin',             // 4
    '    Baz := 1;',         // 5
    '    end;',              // 6
    '  end;']));             // 7

  Assert.AreEqual<Integer>(2, Length(Occs), 'both the outer and the inner with are reported');
end;

procedure TWithScannerTests.ScanSource_NestedWith_OuterBodyContainsInnerKeyword;
var
  Occs: TArray<TWithOccurrence>;
  I, OuterIdx, InnerIdx: Integer;
begin
  // This is the predicate Expert.WithRefactorWizard relies on to refuse an
  // outer with that contains a nested one. If the scanner's BodyInnerRange
  // ever stops covering the inner keyword, that refusal silently stops
  // firing and the outer rewrite starts stealing the inner's identifiers.
  Occs := TWithScanner.ScanSource(Src([
    'with FOuter do',        // 1
    '  begin',               // 2
    '  with FInner do',      // 3
    '    begin',             // 4
    '    Baz := 1;',         // 5
    '    end;',              // 6
    '  end;']));             // 7

  Assert.AreEqual<Integer>(2, Length(Occs));

  OuterIdx := -1;
  InnerIdx := -1;
  for I := 0 to High(Occs) do
    if Occs[I].Targets[0].Expression = 'FOuter' then
      OuterIdx := I
    else if Occs[I].Targets[0].Expression = 'FInner' then
      InnerIdx := I;

  Assert.IsTrue(OuterIdx >= 0, 'outer occurrence not found');
  Assert.IsTrue(InnerIdx >= 0, 'inner occurrence not found');

  Assert.IsTrue(
    PosInRange(Occs[InnerIdx].KeywordPos, Occs[OuterIdx].BodyInnerRange),
    'the inner with keyword must lie inside the outer body range');

  Assert.IsFalse(
    PosInRange(Occs[OuterIdx].KeywordPos, Occs[InnerIdx].BodyInnerRange),
    'containment must not be symmetric');
end;

procedure TWithScannerTests.ScanSource_SiblingWiths_DoNotContainEachOther;
var
  Occs: TArray<TWithOccurrence>;
begin
  // Guards the other direction: two independent with-statements must not be
  // seen as nested, or the wizard would refuse perfectly rewritable blocks.
  Occs := TWithScanner.ScanSource(Src([
    'begin',                 // 1
    '  with FOne do',        // 2
    '    Baz := 1;',         // 3
    '  with FTwo do',        // 4
    '    Quux := 2;',        // 5
    'end;']));               // 6

  Assert.AreEqual<Integer>(2, Length(Occs), 'two sibling with-statements');
  Assert.IsFalse(PosInRange(Occs[1].KeywordPos, Occs[0].BodyInnerRange),
    'sibling must not be seen as nested');
  Assert.IsFalse(PosInRange(Occs[0].KeywordPos, Occs[1].BodyInnerRange),
    'sibling must not be seen as nested (reverse)');
end;

procedure TWithScannerTests.ScanSource_TwoSiblingBeginEndWiths_FindsBoth;
var
  Occs: TArray<TWithOccurrence>;
begin
  // Same shape as the sibling test but with begin..end bodies, to separate
  // "sibling handling is broken" from "single-statement body handling is
  // broken".
  Occs := TWithScanner.ScanSource(Src([
    'begin',
    '  with FOne do',
    '    begin',
    '    Baz := 1;',
    '    end;',
    '  with FTwo do',
    '    begin',
    '    Quux := 2;',
    '    end;',
    'end;']));
  Assert.AreEqual<Integer>(2, Length(Occs), 'two sibling begin..end with-statements');
end;

procedure TWithScannerTests.ScanSource_TwoSiblingSingleStatementWiths_FindsBoth;
var
  Occs: TArray<TWithOccurrence>;
begin
  // Two single-statement withs with no enclosing begin, to isolate whether
  // the enclosing block matters.
  Occs := TWithScanner.ScanSource(Src([
    'with FOne do',
    '  Baz := 1;',
    'with FTwo do',
    '  Quux := 2;']));
  Assert.AreEqual<Integer>(2, Length(Occs), 'two sibling single-statement with-statements');
end;

initialization
  TDUnitX.RegisterTestFixture(TWithScannerTests);

end.
