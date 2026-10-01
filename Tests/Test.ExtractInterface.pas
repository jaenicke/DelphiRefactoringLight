(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Extract interface / Add to existing interface / Add IInterface support:
///  the text routines of Expert.ExtractInterface - the class header
///  matcher, the project interface scanner, what is spliced into an
///  existing interface, finding a declaration again after an edit, and
///  the IInterface directives.
/// </summary>
unit Test.ExtractInterface;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TExtractInterfaceTests = class
  public
    /// <summary>"TMyClass = class": the word "class" inside the NAME made
    ///  the header matcher skip the line, and the class above it was
    ///  parsed instead.</summary>
    [Test] procedure ClassNameContainingClass_IsTheClassParsed;
    /// <summary>The scanner lists interfaces only: not a class with
    ///  TInterfacedObject in its ancestor list, and not a forward
    ///  "IFoo = interface;", which ran to the next type's end.</summary>
    [Test] procedure ScanFile_OnlyRealInterfaceDeclarations;
    /// <summary>With every selected member already in the interface there
    ///  is nothing to splice - a whole "IFoo = interface ... end;" was
    ///  spliced into the existing interface.</summary>
    [Test] procedure Splice_NothingLeft_IsEmpty;
    /// <summary>After the uses clause grew by three lines, the interface is
    ///  found again by name - walking on from the old line ended at the
    ///  previous interface's end.</summary>
    [Test] procedure FindDecl_AfterTheFileGrewAboveIt;
    /// <summary>The class is found again from its old line after an
    ///  interface above it in the same unit was extended.</summary>
    [Test] procedure RelocateClass_AfterTheFileGrewAboveIt;
    /// <summary>On a TComponent-style base only QueryInterface is virtual;
    ///  "override" on _AddRef/_Release does not compile (E2170).</summary>
    [Test] procedure IInterfaceDirective_OnlyQueryInterfaceOnComponents;
  end;

implementation

uses
  System.SysUtils, System.IOUtils, Expert.ExtractInterface;

procedure TExtractInterfaceTests.ClassNameContainingClass_IsTheClassParsed;
var
  Info: TExtractInterfaceInfo;
begin
  Assert.IsTrue(TExtractInterfaceEngine.ParseClassAtLine([
    'unit U;',                  // 1
    'interface',                // 2
    'type',                     // 3
    '  TOther = class',         // 4
    '    procedure A;',         // 5
    '  end;',                   // 6
    '  TMyClass = class',       // 7
    '    procedure Run;',       // 8
    '  end;',                   // 9
    'implementation',           // 10
    'end.'], '', 8, Info));
  Assert.AreEqual('TMyClass', Info.ClassName, False);
  Assert.AreEqual(7, Info.ClassDeclLine);
end;

procedure TExtractInterfaceTests.ScanFile_OnlyRealInterfaceDeclarations;
var
  FileName, Found: string;
begin
  FileName := TPath.GetTempFileName;
  TFile.WriteAllText(FileName, string.Join(sLineBreak, [
    'unit U;',                                    // 1
    'interface',                                  // 2
    'type',                                       // 3
    '  TImpl = class(TInterfacedObject, IBar)',   // 4
    '    procedure X;',                           // 5
    '  end;',                                     // 6
    '  IFwd = interface;',                        // 7
    '  IBar = interface',                         // 8
    '    procedure X;',                           // 9
    '  end;',                                     // 10
    '  IFwd = interface',                         // 11
    '    procedure Y;',                           // 12
    '  end;',                                     // 13
    'implementation',                             // 14
    'end.']));
  try
    Found := '';
    for var Loc in TProjectInterfaceScanner.ScanFile(FileName) do
      Found := Found + Format('%s %d-%d|', [Loc.InterfaceName, Loc.DeclLine, Loc.EndLine]);
    Assert.AreEqual('IBar 8-10|IFwd 11-13|', Found, False);
  finally
    TFile.Delete(FileName);
  end;
end;

procedure TExtractInterfaceTests.Splice_NothingLeft_IsEmpty;
var
  Info: TExtractInterfaceInfo;
begin
  Info := Default(TExtractInterfaceInfo);
  Info.InterfaceName := 'IFoo';
  Info.Guid := '{00000000-0000-0000-0000-000000000001}';
  SetLength(Info.Members, 1);
  Info.Members[0].Name := 'Run';
  Info.Members[0].Kind := mkMethod;
  Info.Members[0].Signature := 'procedure Run';
  Info.Members[0].Selected := False;
  Assert.AreEqual('', string.Join('|', TExtractInterfaceEngine.InterfaceSpliceLines(Info)),
    False, 'nothing selected: nothing to splice');
  Info.Members[0].Selected := True;
  Assert.AreEqual('    procedure Run;',
    string.Join('|', TExtractInterfaceEngine.InterfaceSpliceLines(Info)), False);
end;

procedure TExtractInterfaceTests.FindDecl_AfterTheFileGrewAboveIt;
var
  Loc: TInterfaceDeclLocation;
begin
  // IFoo was at line 7 before three uses lines were added above it
  Assert.IsTrue(TProjectInterfaceScanner.FindDecl([
    'unit U;',                  // 1
    'interface',                // 2
    '',                         // 3 (added)
    'uses',                     // 4 (added)
    '  A;',                     // 5 (added)
    'type',                     // 6
    '  IOther = interface',     // 7
    '    procedure P;',         // 8
    '  end;',                   // 9
    '  IFoo = interface',       // 10
    '    procedure Q;',         // 11
    '  end;',                   // 12
    'implementation',           // 13
    'end.'], 'IFoo', 7, Loc));
  Assert.AreEqual(10, Loc.DeclLine);
  Assert.AreEqual(12, Loc.EndLine, 'the end of IFoo, not of IOther');
end;

procedure TExtractInterfaceTests.RelocateClass_AfterTheFileGrewAboveIt;
var
  DeclLine, EndLine: Integer;
begin
  // TFoo was at line 5 before a member was spliced into IFoo above it
  Assert.IsTrue(TExtractInterfaceEngine.RelocateClass([
    'type',                     // 1
    '  IFoo = interface',       // 2
    '    procedure A;',         // 3
    '    procedure B;',         // 4 (spliced in)
    '  end;',                   // 5
    '  TFoo = class',           // 6
    '    procedure A;',         // 7
    '  end;'], 'TFoo', 5, DeclLine, EndLine));
  Assert.AreEqual(6, DeclLine);
  Assert.AreEqual(8, EndLine);
end;

procedure TExtractInterfaceTests.IInterfaceDirective_OnlyQueryInterfaceOnComponents;
begin
  Assert.AreEqual('override; ',
    TExtractInterfaceEngine.IInterfaceDirective('TComponent', 'QueryInterface'), False);
  Assert.AreEqual('', TExtractInterfaceEngine.IInterfaceDirective('TComponent', '_AddRef'), False,
    '_AddRef is static in TComponent');
  Assert.AreEqual('', TExtractInterfaceEngine.IInterfaceDirective('TComponent', '_Release'), False,
    '_Release is static in TComponent');
  Assert.AreEqual('', TExtractInterfaceEngine.IInterfaceDirective('TObject', 'QueryInterface'), False);
end;

initialization
  TDUnitX.RegisterTestFixture(TExtractInterfaceTests);

end.
