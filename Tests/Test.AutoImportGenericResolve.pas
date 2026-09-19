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
///  Regression tests for the E2003 identifier extraction and generic
///  candidate filtering that the 2026-08-30 upstream merge added to
///  Expert.AutoImport - the resolution half of the generics fix. (The index
///  half is pinned in Test.UnitIndexGenerics.)
///
///  The defect these guard against is subtle and silent. For a generic use
///  site the compiler reports 'TList&lt;&gt;' but points its diagnostic RANGE at
///  (or into) the type ARGUMENTS, so reading the token at the range yields
///  'Integer' - or a type parameter - and the quick fix resolves a completely
///  different identifier. Nothing errors; the user is simply offered the
///  wrong unit, or none.
///
///  All three routines are pure, and upstream exposes them in the interface
///  for exactly this. The unit needs no live index: TUnitIndex is only
///  reached by ResolveQuickFixes, which is deliberately NOT tested here -
///  it exits early without a snapshot, and a snapshot means indexing the
///  machine's whole library path, which is an integration test, not this.
/// </summary>
unit Test.AutoImportGenericResolve;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TAutoImportGenericResolveTests = class
  public
    // ---- E2003IdentFromDiag ----
    [Test] procedure GenericUse_TakesTheQuotedName_NotTheRangeToken;
    [Test] procedure GenericUse_AnchorsOnTheRealTokenNotTheRangeStart;
    [Test] procedure MultiParamGeneric_StripsTheWholeArityMarker;
    [Test] procedure NonGenericUse_UsesTheRangeTokenAndItsPosition;
    [Test] procedure MessageWithoutQuotes_FallsBackToTheRangeToken;
    [Test] procedure QuotedTextThatIsNotAnIdentifier_IsRejected;
    [Test] procedure NameNotPresentOnTheLine_FallsBackToTheRangeStart;
    [Test] procedure NameNowhereInTheBuffer_IsTreatedAsStale;
    [Test] procedure LineOutOfRange_ReturnsEmpty;
    [Test] procedure AnchorsOnAWholeWord_NotInsideALongerIdentifier;

    // ---- IsGenericUseAt ----
    [Test] procedure AngleBracketAfterToken_IsAGenericUse;
    [Test] procedure SemicolonAfterToken_IsNotAGenericUse;
    [Test] procedure WhitespaceBeforeAngleBracket_IsStillAGenericUse;
    [Test] procedure TokenAtEndOfLine_IsNotAGenericUse;
    [Test] procedure KnownLimitation_LessThanOperatorReadsAsGenericUse;

    // ---- FilterHitsByGenericUse ----
    [Test] procedure GenericUse_KeepsOnlyGenericCandidates;
    [Test] procedure BareUse_KeepsOnlyNonGenericCandidates;
    [Test] procedure NoCandidateOfTheRightForm_KeepsThemAll;
    [Test] procedure EmptyCandidateList_StaysEmpty;
  end;

implementation

uses
  System.SysUtils,
  Lsp.Protocol, Expert.UnitIndex, Expert.AutoImport;

// Record builders: plain implementation-section functions, so the fixture's
// published surface stays exactly the list of tests.
function MakeDiag(ALine, ACharStart, ACharEnd: Integer;
  const AName: string): TLspErrorDiag;
begin
  Result := Default(TLspErrorDiag);
  Result.Code := 'E2003';
  Result.Severity := 1;
  // The real message is localized around the quoted name; only the text
  // between the first pair of single quotes is ever read.
  Result.Message := 'E2003 Undeclared identifier: ' + QuotedStr(AName);
  Result.Range.Start.Line := ALine;
  Result.Range.Start.Character := ACharStart;
  Result.Range.End_.Line := ALine;
  Result.Range.End_.Character := ACharEnd;
end;

function MakeHit(const AUnitName: string; AIsGeneric: Boolean): TFindUnitHit;
begin
  Result := Default(TFindUnitHit);
  Result.Identifier := 'TList';
  Result.UnitName := AUnitName;
  Result.Path := 'C:\src\' + AUnitName + '.pas';
  Result.IsGeneric := AIsGeneric;
end;

{ E2003IdentFromDiag }

procedure TAutoImportGenericResolveTests.GenericUse_TakesTheQuotedName_NotTheRangeToken;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  //                 0         1         2
  //                 0123456789012345678901
  Lines := ['    FIds: TList<Integer>;'];
  // The range sits on the type ARGUMENT, which is the whole problem.
  Assert.AreEqual('Integer', Copy(Lines[0], 17, 7),
    'precondition: the range really does cover the type argument');
  Assert.AreEqual('TList',
    E2003IdentFromDiag(Lines, MakeDiag(0, 16, 23, 'TList<>'), Col0, Len),
    'the quoted name wins over the token at the range');
end;

procedure TAutoImportGenericResolveTests.GenericUse_AnchorsOnTheRealTokenNotTheRangeStart;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  Lines := ['    FIds: TList<Integer>;'];
  E2003IdentFromDiag(Lines, MakeDiag(0, 16, 23, 'TList<>'), Col0, Len);
  // Col0/Len drive the editor's dotted underline and the rename
  // replacement - a right name at the wrong offset marks the wrong text.
  Assert.AreEqual<Integer>(10, Col0, 'anchored at the real "TList"');
  Assert.AreEqual<Integer>(5, Len, 'and only as long as the bare name');
  Assert.AreEqual('TList', Copy(Lines[0], Col0 + 1, Len),
    'the anchor actually covers the identifier');
end;

procedure TAutoImportGenericResolveTests.MultiParamGeneric_StripsTheWholeArityMarker;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  Lines := ['  FMap: TDictionary<string, Integer>;'];
  Assert.AreEqual('TDictionary',
    E2003IdentFromDiag(Lines, MakeDiag(0, 20, 26, 'TDictionary<,>'), Col0, Len),
    'everything from "<" on is arity, not part of the name');
end;

procedure TAutoImportGenericResolveTests.NonGenericUse_UsesTheRangeTokenAndItsPosition;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  //             0123456789012345678
  Lines := ['  FLegacy: TList;'];
  Assert.AreEqual('TList',
    E2003IdentFromDiag(Lines, MakeDiag(0, 11, 16, 'TList'), Col0, Len),
    'message and range agree');
  Assert.AreEqual<Integer>(11, Col0, 'and the range position is used directly');
  Assert.AreEqual<Integer>(5, Len, 'with the token length');
end;

procedure TAutoImportGenericResolveTests.MessageWithoutQuotes_FallsBackToTheRangeToken;
var
  Lines: TArray<string>;
  D: TLspErrorDiag;
  Col0, Len: Integer;
begin
  // A localization that does not quote the name must not break resolution.
  Lines := ['  FLegacy: TList;'];
  D := MakeDiag(0, 11, 16, 'TList');
  D.Message := 'Undeklarierter Bezeichner TList';   // no quotes at all
  Assert.AreEqual('TList', E2003IdentFromDiag(Lines, D, Col0, Len),
    'falls back to the token at the diagnostic range');
  Assert.AreEqual<Integer>(11, Col0, 'anchored at the range');
end;

procedure TAutoImportGenericResolveTests.QuotedTextThatIsNotAnIdentifier_IsRejected;
var
  Lines: TArray<string>;
  D: TLspErrorDiag;
  Col0, Len: Integer;
begin
  // Some messages quote a FILE, not a symbol. Taking that as the identifier
  // would send the index looking for "Unit1.pas".
  Lines := ['  FLegacy: TList;'];
  D := MakeDiag(0, 11, 16, 'TList');
  D.Message := 'E2003 something about ''Unit1.pas'' here';
  Assert.AreEqual('TList', E2003IdentFromDiag(Lines, D, Col0, Len),
    'a quoted non-identifier is discarded and the range token used instead');
end;

procedure TAutoImportGenericResolveTests.NameNotPresentOnTheLine_FallsBackToTheRangeStart;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
  Ident: string;
begin
  // Message and range disagree AND the reported name is nowhere on the
  // DIAGNOSED line - a multi-line construct, so the declaration carrying it
  // sits on another line of the same buffer. The name is still right; only
  // the anchor is a guess, and it must be the range start rather than
  // garbage.
  Lines := ['    FIds := Create;', '  FIds: TList<Integer>;'];
  Ident := E2003IdentFromDiag(Lines, MakeDiag(0, 12, 18, 'TList<>'), Col0, Len);
  Assert.AreEqual('TList', Ident, 'the reported name is still returned');
  Assert.AreEqual<Integer>(12, Col0, 'anchor falls back to the range start');
  Assert.AreEqual<Integer>(5, Len, 'with the reported name''s length');
end;

procedure TAutoImportGenericResolveTests.NameNowhereInTheBuffer_IsTreatedAsStale;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  // The OTHER half of the same rule, added upstream in 2c2075a: when the
  // reported identifier occurs nowhere in the WHOLE buffer, the diagnostic
  // is stale - Error Insight lags one edit behind, so deleting the offending
  // line re-publishes the old E2003 against the new content, where it would
  // otherwise stick for good ("the fix stays although I removed the line").
  // Such a diagnostic must resolve to nothing rather than to a rough anchor.
  Lines := ['    FIds := Create;'];
  Assert.AreEqual('',
    E2003IdentFromDiag(Lines, MakeDiag(0, 12, 18, 'TList<>'), Col0, Len),
    'an identifier absent from the entire buffer is a stale diagnostic');
  Assert.AreEqual<Integer>(0, Col0, 'and carries no anchor');
  Assert.AreEqual<Integer>(0, Len, 'nor a length');
end;

procedure TAutoImportGenericResolveTests.LineOutOfRange_ReturnsEmpty;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  Lines := ['  FLegacy: TList;'];
  Assert.AreEqual('',
    E2003IdentFromDiag(Lines, MakeDiag(9, 0, 5, 'TList'), Col0, Len),
    'a diagnostic past the end of the buffer resolves to nothing');
  Assert.AreEqual('',
    E2003IdentFromDiag(Lines, MakeDiag(-1, 0, 5, 'TList'), Col0, Len),
    'and so does a negative line');
end;

procedure TAutoImportGenericResolveTests.AnchorsOnAWholeWord_NotInsideALongerIdentifier;
var
  Lines: TArray<string>;
  Col0, Len: Integer;
begin
  // "MyTListHelper" contains "TList". Anchoring inside it would underline
  // the middle of an unrelated identifier and, on a rename fix, corrupt it.
  //             0         1         2         3
  //             0123456789012345678901234567890123456789
  Lines := ['  MyTListHelper.Wrap(FIds: TList<Integer>);'];
  E2003IdentFromDiag(Lines, MakeDiag(0, 33, 40, 'TList<>'), Col0, Len);
  Assert.AreEqual('TList', Copy(Lines[0], Col0 + 1, Len),
    'the anchor covers exactly "TList"');
  Assert.AreEqual<Integer>(27, Col0,
    'and it is the standalone occurrence, not the one inside MyTListHelper');
end;

{ IsGenericUseAt }

procedure TAutoImportGenericResolveTests.AngleBracketAfterToken_IsAGenericUse;
begin
  Assert.IsTrue(IsGenericUseAt('    FIds: TList<Integer>;', 10, 5),
    '"TList<" is a generic use site');
end;

procedure TAutoImportGenericResolveTests.SemicolonAfterToken_IsNotAGenericUse;
begin
  Assert.IsFalse(IsGenericUseAt('  FLegacy: TList;', 11, 5),
    '"TList;" is not');
end;

procedure TAutoImportGenericResolveTests.WhitespaceBeforeAngleBracket_IsStillAGenericUse;
begin
  Assert.IsTrue(IsGenericUseAt('    FIds: TList <Integer>;', 10, 5),
    'spaces between the name and "<" do not change what is meant');
  Assert.IsTrue(IsGenericUseAt('    FIds: TList'#9'<Integer>;', 10, 5),
    'nor does a tab');
end;

procedure TAutoImportGenericResolveTests.TokenAtEndOfLine_IsNotAGenericUse;
begin
  // Nothing follows, so there is nothing to read - and reading past the end
  // must not raise.
  Assert.IsFalse(IsGenericUseAt('  TList', 2, 5),
    'a token at end of line is not a generic use');
end;

procedure TAutoImportGenericResolveTests.KnownLimitation_LessThanOperatorReadsAsGenericUse;
begin
  // DOCUMENTED BEHAVIOUR, NOT A DESIRED PROPERTY. The check is purely
  // lexical: the next non-blank character after the token. A comparison
  // therefore reads as a generic use. It is harmless in practice because the
  // caller only reaches it for an E2003 identifier, and FilterHitsByGenericUse
  // keeps every candidate when none matches the assumed form. Pinned so that
  // changing it is a visible decision rather than an accident.
  Assert.IsTrue(IsGenericUseAt('  if Count < Limit then', 5, 5),
    'a "<" operator is indistinguishable from a type-argument list here');
end;

{ FilterHitsByGenericUse }

procedure TAutoImportGenericResolveTests.GenericUse_KeepsOnlyGenericCandidates;
var
  Hits, Kept: TArray<TFindUnitHit>;
begin
  // The real pair: System.Classes.TList versus
  // System.Generics.Collections.TList<T>.
  Hits := [MakeHit('System.Classes', False),
           MakeHit('System.Generics.Collections', True)];
  Kept := FilterHitsByGenericUse(Hits, True);
  Assert.AreEqual<Integer>(1, Length(Kept), 'one candidate survives');
  Assert.AreEqual('System.Generics.Collections', Kept[0].UnitName,
    'a TList<Integer> site is never offered the non-generic TList');
end;

procedure TAutoImportGenericResolveTests.BareUse_KeepsOnlyNonGenericCandidates;
var
  Hits, Kept: TArray<TFindUnitHit>;
begin
  Hits := [MakeHit('System.Classes', False),
           MakeHit('System.Generics.Collections', True)];
  Kept := FilterHitsByGenericUse(Hits, False);
  Assert.AreEqual<Integer>(1, Length(Kept), 'one candidate survives');
  Assert.AreEqual('System.Classes', Kept[0].UnitName,
    'and the filter works in the other direction too');
end;

procedure TAutoImportGenericResolveTests.NoCandidateOfTheRightForm_KeepsThemAll;
var
  Hits, Kept: TArray<TFindUnitHit>;
begin
  // Deliberate: offering nothing is worse than offering an imperfect guess,
  // because the index's genericity flag is heuristic.
  Hits := [MakeHit('System.Classes', False), MakeHit('EBaseClasses', False)];
  Kept := FilterHitsByGenericUse(Hits, True);
  Assert.AreEqual<Integer>(2, Length(Kept),
    'with no matching form at all, every candidate is kept');
end;

procedure TAutoImportGenericResolveTests.EmptyCandidateList_StaysEmpty;
var
  Kept: TArray<TFindUnitHit>;
begin
  Kept := FilterHitsByGenericUse(nil, True);
  Assert.AreEqual<Integer>(0, Length(Kept), 'nothing in, nothing out');
end;

initialization
  TDUnitX.RegisterTestFixture(TAutoImportGenericResolveTests);

end.
