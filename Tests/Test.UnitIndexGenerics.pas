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
///  Regression tests for the GENERIC-declaration handling in
///  Expert.UnitIndex.ParseUnit - the half of the 2026-08-30 upstream merge
///  that decides whether a candidate unit can satisfy a "TList&lt;Integer&gt;"
///  use site at all.
///
///  ParseUnit is the right seam: it is a pure function over a file, needs
///  neither DelphiLSP nor the IDE (the Editor global is only touched by the
///  project-scope gatherers), and upstream exposes it in the interface
///  precisely so the console suite can reach it.
///
///  Two behaviours are pinned here, and both are load-bearing:
///
///  1. The trailing '&lt;' MARKER on a generic declaration. It is what becomes
///     TFindUnitHit.IsGeneric when the snapshot map is built, and therefore
///     what stops the quick fix offering System.Classes' non-generic TList
///     for a "TList&lt;Integer&gt;" use site. Drop the marker and the filter
///     silently degrades to offering everything.
///
///  2. StripAngleSpans. A generic constraint ("Store&lt;T: record&gt;") used to
///     count as a nested RECORD opener, so its phantom END swallowed every
///     declaration after the class - the unit was indexed, but half its
///     contents went missing. That failure is invisible from the outside:
///     the identifier is simply never offered.
/// </summary>
unit Test.UnitIndexGenerics;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TUnitIndexGenericsTests = class
  private
    FDir: string;
    /// <summary>Writes ALines as a .pas into the fixture's temp directory
    ///  and returns its full path. The lines land inside a type section.</summary>
    function WriteUnit(const AUnitName: string;
      const ALines: array of string): string;
    /// <summary>Parses AFile and returns its exported identifiers.</summary>
    function Parse(const AFile: string): TArray<string>;
    /// <summary>Exact (case-sensitive) membership - the generic marker is
    ///  part of the string, so SameText would hide a missing '&lt;'.</summary>
    function Has(const AIdents: TArray<string>; const AIdent: string): Boolean;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure GenericClass_IsMarked;
    [Test] procedure NonGenericClass_IsNotMarked;
    [Test] procedure BothFormsInOneUnit_AreIndexedSeparately;
    [Test] procedure GenericRecord_IsMarked;
    [Test] procedure MultiParamGeneric_IsMarkedUnderItsBaseName;
    [Test] procedure GenericRoutine_IsMarked;
    [Test] procedure ForwardDeclaration_DoesNotLoseTheRealDeclaration;
    [Test] procedure GenericConstraint_DoesNotSwallowLaterDeclarations;
    [Test] procedure GenericConstraint_DoesNotLeakClassMembersToTopLevel;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils,
  Expert.UnitIndex;

{ TUnitIndexGenericsTests }

procedure TUnitIndexGenericsTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath,
    'RLGenerics_' + TGUID.NewGuid.ToString.Replace('{', '').Replace('}', ''));
  TDirectory.CreateDirectory(FDir);
end;

procedure TUnitIndexGenericsTests.TearDown;
begin
  if (FDir <> '') and TDirectory.Exists(FDir) then
    TDirectory.Delete(FDir, True);
end;

function TUnitIndexGenericsTests.WriteUnit(const AUnitName: string;
  const ALines: array of string): string;
var
  SL: TStringList;
  L: string;
begin
  Result := TPath.Combine(FDir, AUnitName + '.pas');
  SL := TStringList.Create;
  try
    SL.Add('unit ' + AUnitName + ';');
    SL.Add('');
    SL.Add('interface');
    SL.Add('');
    SL.Add('type');
    for L in ALines do
      SL.Add(L);
    SL.Add('');
    SL.Add('implementation');
    SL.Add('');
    SL.Add('end.');
    SL.WriteBOM := True;
    SL.SaveToFile(Result, TEncoding.UTF8);
  finally
    SL.Free;
  end;
end;

function TUnitIndexGenericsTests.Parse(const AFile: string): TArray<string>;
var
  UnitName: string;
  HasInit: Boolean;
begin
  Result := ParseUnit(AFile, UnitName, HasInit);
end;

function TUnitIndexGenericsTests.Has(const AIdents: TArray<string>;
  const AIdent: string): Boolean;
var
  S: string;
begin
  for S in AIdents do
    if S = AIdent then Exit(True);
  Result := False;
end;

procedure TUnitIndexGenericsTests.GenericClass_IsMarked;
var
  Idents: TArray<string>;
begin
  Idents := Parse(WriteUnit('GenClass', [
    '  TBox<T> = class(TObject)',
    '  public',
    '    procedure Put(const AItem: T);',
    '  end;']));
  // The KEY is the bare name (that is what the user types and what the
  // compiler names in its E2003); the '<' rides along as the flag.
  Assert.IsTrue(Has(Idents, 'TBox<'),
    'generic class must be indexed as "TBox<" (base name + generic marker)');
  Assert.IsFalse(Has(Idents, 'TBox'),
    'the unmarked name must not also be present - it would make the generic '
    + 'declaration look like a non-generic candidate');
end;

procedure TUnitIndexGenericsTests.NonGenericClass_IsNotMarked;
var
  Idents: TArray<string>;
begin
  Idents := Parse(WriteUnit('PlainClass', [
    '  TBox = class(TObject)',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'TBox'), 'non-generic class indexed under its name');
  Assert.IsFalse(Has(Idents, 'TBox<'), 'non-generic class must carry no marker');
end;

procedure TUnitIndexGenericsTests.BothFormsInOneUnit_AreIndexedSeparately;
var
  Idents: TArray<string>;
begin
  // The shape that matters in the RTL: System.Classes.TList (non-generic)
  // versus System.Generics.Collections.TList<T>. Here both live in one unit
  // so a single parse has to keep them apart.
  Idents := Parse(WriteUnit('BothForms', [
    '  TBag = class(TObject)',
    '  end;',
    '',
    '  TBag<T> = class(TObject)',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'TBag'), 'the non-generic declaration survives');
  Assert.IsTrue(Has(Idents, 'TBag<'), 'the generic declaration survives');
end;

procedure TUnitIndexGenericsTests.GenericRecord_IsMarked;
var
  Idents: TArray<string>;
begin
  Idents := Parse(WriteUnit('GenRecord', [
    '  TPair<K, V> = record',
    '    Key: K;',
    '    Value: V;',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'TPair<'), 'generic record marked like a generic class');
end;

procedure TUnitIndexGenericsTests.MultiParamGeneric_IsMarkedUnderItsBaseName;
var
  Idents: TArray<string>;
begin
  // The compiler reports these as 'TDictionary<,>'; the index must key the
  // bare name so the reported name and the indexed key meet.
  Idents := Parse(WriteUnit('MultiParam', [
    '  TMap<TKey, TValue> = class(TObject)',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'TMap<'),
    'multi-parameter generic keyed under its base name');
end;

procedure TUnitIndexGenericsTests.GenericRoutine_IsMarked;
var
  Idents: TArray<string>;
begin
  Idents := Parse(WriteUnit('GenRoutine', [
    '  TMarker = record',
    '    Value: Integer;',
    '  end;',
    '',
    'function Wrap<T>(const AValue: T): TMarker;',
    'procedure Plain(const AValue: Integer);']));
  Assert.IsTrue(Has(Idents, 'Wrap<'), 'generic routine carries the marker');
  Assert.IsTrue(Has(Idents, 'Plain'), 'non-generic routine does not');
  Assert.IsFalse(Has(Idents, 'Plain<'), 'and must not gain one');
end;

procedure TUnitIndexGenericsTests.ForwardDeclaration_DoesNotLoseTheRealDeclaration;
var
  Idents: TArray<string>;
begin
  // System.Classes' actual shape: "TList = class;" at line 257, the real
  // declaration at 270. ParseUnit reports the name twice; that is benign
  // (DedupeByUnitName collapses it downstream), but the real declaration
  // must still be there - and must not be mistaken for a generic one.
  Idents := Parse(WriteUnit('FwdDecl', [
    '  TNode = class;',
    '',
    '  TNode = class(TObject)',
    '  public',
    '    Next: TNode;',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'TNode'), 'the declaration is indexed');
  Assert.IsFalse(Has(Idents, 'TNode<'), 'a forward declaration is not generic');
end;

procedure TUnitIndexGenericsTests.GenericConstraint_DoesNotSwallowLaterDeclarations;
var
  Idents: TArray<string>;
begin
  // Without StripAngleSpans the RECORD inside the constraint counted as a
  // nested record opener, so the class's own END did not close it and every
  // later declaration was parsed as if still inside the class body.
  Idents := Parse(WriteUnit('Constraint', [
    '  THolder = class(TObject)',
    '  public',
    '    procedure Store<T: record>(const AItem: T);',
    '  end;',
    '',
    '  TAfterTheClass = class(TObject)',
    '  end;',
    '',
    '  TAlsoAfter = record',
    '    Value: Integer;',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'THolder'), 'the class itself is indexed');
  Assert.IsTrue(Has(Idents, 'TAfterTheClass'),
    'a constraint must not swallow the declarations that follow the class');
  Assert.IsTrue(Has(Idents, 'TAlsoAfter'), 'nor the ones after that');
end;

procedure TUnitIndexGenericsTests.GenericConstraint_DoesNotLeakClassMembersToTopLevel;
var
  Idents: TArray<string>;
begin
  // The mirror-image failure of the one above: if the constraint threw the
  // end-counting off the other way, the class body would be treated as top
  // level and its methods would be indexed as exported identifiers.
  Idents := Parse(WriteUnit('NoLeak', [
    '  THolder = class(TObject)',
    '  public',
    '    procedure Store<T: class>(const AItem: T);',
    '    procedure MemberOnly(const A: Integer);',
    '  end;']));
  Assert.IsTrue(Has(Idents, 'THolder'), 'the class itself is indexed');
  Assert.IsFalse(Has(Idents, 'MemberOnly'),
    'a class method is not an exported identifier of the unit');
  Assert.IsFalse(Has(Idents, 'Store<'), 'nor is a generic class method');
end;

initialization
  TDUnitX.RegisterTestFixture(TUnitIndexGenericsTests);

end.
