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

    /// <summary>Issue #44, the reported class verbatim. Two defects in one
    ///  shape: a WRAPPED parameter list was cut at the ';' inside it, so
    ///  "procedure Test3( const AParam1: string;" was the whole method and
    ///  the continuation line became a FIELD member; and a '&'-ESCAPED
    ///  property name was parsed as EMPTY, which generated
    ///  "function Get: String" and "property : String read Get".</summary>
    [Test] procedure WrappedParametersAndEscapedNames_Issue44;
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

procedure TExtractInterfaceTests.WrappedParametersAndEscapedNames_Issue44;
var
  Info: TExtractInterfaceInfo;
  Names, Sigs: string;
  I: Integer;
begin
  // The reporter's class, verbatim.
  Assert.IsTrue(TExtractInterfaceEngine.ParseClassAtLine([
    'unit U;',                                      // 1
    'interface',                                    // 2
    'type',                                         // 3
    '  TMyObject = class(TObject)',                 // 4
    '  private',                                    // 5
    '    FString: String;',                         // 6
    '  protected',                                  // 7
    '  public',                                     // 8
    '    procedure &Integer;',                      // 9
    '    procedure &Test;',                         // 10
    '    procedure Test2( const AParam: Boolean);', // 11
    '    procedure Test3( const AParam1: string;',  // 12
    '                     const AParam2: Boolean);',// 13
    '    property &String: String read FString;',   // 14
    '  end;',                                       // 15
    'implementation',                               // 16
    'end.'], '', 4, Info));

  for I := 0 to High(Info.Members) do
  begin
    Names := Names + '|' + Info.Members[I].Name;
    Sigs := Sigs + '|' + Info.Members[I].Signature;
    // What the reporter ticked: the public members, not the private field.
    Info.Members[I].Selected := Info.Members[I].Visibility = mvPublic;
  end;

  // ONE entry for Test3, and its continuation line is not a member of its
  // own - it used to arrive as a "field" called "const AParam2".
  Assert.AreEqual('|FString|&Integer|&Test|Test2|Test3|&String', Names, False,
    'six members, and the escaped names are names');
  Assert.IsTrue(Sigs.Contains(
    '|procedure Test3( const AParam1: string; const AParam2: Boolean)'),
    'the wrapped parameter list stays with its header: ' + Sigs);

  Info.InterfaceName := 'IMyObject';
  Info.Guid := '{23E240A6-23A0-4DC8-883E-3BA021063342}';

  // The reporter's expected preview, verbatim.
  Assert.AreEqual(
    '  IMyObject = interface' + sLineBreak +
    '    [''{23E240A6-23A0-4DC8-883E-3BA021063342}'']' + sLineBreak +
    '    procedure &Integer;' + sLineBreak +
    '    procedure &Test;' + sLineBreak +
    '    procedure Test2( const AParam: Boolean);' + sLineBreak +
    '    procedure Test3( const AParam1: string; const AParam2: Boolean);' + sLineBreak +
    '    function GetString: String;' + sLineBreak +
    '    property &String: String read GetString;' + sLineBreak +
    '  end;',
    TExtractInterfaceEngine.BuildInterfaceText(Info), False,
    'the escaped property keeps its & and its getter drops it');
end;

initialization
  TDUnitX.RegisterTestFixture(TExtractInterfaceTests);

end.
