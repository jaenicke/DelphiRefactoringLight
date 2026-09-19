(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Version 1.3.0 features: interface <-> class links for find references,
///  the include expander for debugging, and the property converter. The
///  pure halves only.
/// </summary>
unit Test.UserIdeas13;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TUserIdeas13Tests = class
  public
    [Test] procedure InterfaceLinks_BothDirections;
    [Test] procedure IncludeExpander_MarkedAndNested;
    [Test] procedure PropertyConverter_FieldsToAccessors;
    [Test] procedure PropertyConverter_TrivialAccessorsToFields;
  end;

implementation

uses
  System.SysUtils, System.IOUtils, Expert.InterfaceLinks, Expert.IncludeExpansion,
  Expert.PropertyConvert;

procedure TUserIdeas13Tests.InterfaceLinks_BothDirections;
begin
  var Dir := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_links');
  TDirectory.CreateDirectory(Dir);
  var F := TPath.Combine(Dir, 'UL.pas');
  TFile.WriteAllText(F, string.Join(#13#10, [
    'unit UL;', 'interface', 'type',
    '  IBase = interface', '    procedure Bar;', '  end;',
    '  IFoo = interface(IBase)', '  end;',
    '  TFoo = class(TInterfacedObject, IFoo)', '    procedure Bar;', '  end;',
    '  TPlain = class', '    procedure Bar;', '  end;',
    'implementation',
    'procedure TFoo.Bar; begin end;', 'procedure TPlain.Bar; begin end;', 'end.']));
  var G := TTypeGraph.Create([F], nil, False);
  try
    var Up := G.InterfaceMethodsImplementedBy('TFoo', 'Bar');
    Assert.AreEqual<Integer>(1, Length(Up), 'IBase.Bar through IFoo');
    Assert.AreEqual('IBase', Up[0].TypeName);
    Assert.AreEqual(4, Up[0].Line);
    var Down := G.ClassMethodsImplementing('IBase', 'Bar');
    Assert.AreEqual<Integer>(1, Length(Down), 'TPlain does not implement IBase');
    Assert.AreEqual('TFoo', Down[0].TypeName);
    Assert.AreEqual(15, Down[0].ImplLine);
    var T := TLinkedTargets.Create;
    try
      CollectLinkedTargets(G, 'TFoo', 'Bar', T);
      Assert.AreEqual('declared in interface IBase', T.DeclLabel(F, 4));
      Assert.AreEqual('call via interface IBase', T.CallLabel(F, 4));
    finally
      T.Free;
    end;
  finally
    G.Free;
    TDirectory.Delete(Dir, True);
  end;
end;

procedure TUserIdeas13Tests.IncludeExpander_MarkedAndNested;
var
  Count: Integer;
begin
  var Dir := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_marked');
  TDirectory.CreateDirectory(Dir);
  try
    TFile.WriteAllText(TPath.Combine(Dir, 'b.inc'), 'B; {$I c.inc}');
    TFile.WriteAllText(TPath.Combine(Dir, 'c.inc'), 'C; // tail');
    var R := ExpandIncludesMarked('  X; {$I b.inc} Y;'#13#10'end.', Dir, nil, Count);
    Assert.AreEqual(2, Count);
    Assert.AreEqual('  X; '#13#10 +
      '// >>> include begin: {$I b.inc}'#13#10'B; '#13#10 +
      '// >>> include begin: {$I c.inc}'#13#10'C; // tail'#13#10'// <<< include end: c.inc'#13#10 +
      '// <<< include end: b.inc'#13#10' Y;'#13#10'end.', R,
      'the code after the directive must not end up in the include''s // comment');
    Assert.AreEqual('{$I nope.inc}', ExpandIncludesMarked('{$I nope.inc}', Dir, nil, Count));
    Assert.AreEqual(0, Count);
  finally
    TDirectory.Delete(Dir, True);
  end;
end;

procedure TUserIdeas13Tests.PropertyConverter_FieldsToAccessors;
begin
  var Src := TArray<string>.Create('unit UB;', 'interface', 'type', '  TBar = class',
    '    FX: Integer;', '  public', '    property X: Integer read FX write FX;',
    '    property Items[I: Integer]: Integer read FX;', '  end;', '',
    'implementation', '', 'end.');
  var P := PlanPropertyConversion(Src, 6, 7, pcToAccessors, True, False, nil);
  Assert.AreEqual<Integer>(2, Length(P.Items));
  Assert.IsTrue(P.Items[0].Ok);
  Assert.AreEqual('    property X: Integer read GetX write FX;', P.Items[0].After,
    'only the getter was asked for');
  Assert.IsFalse(P.Items[1].Ok, 'array property');
  var Out_ := string.Join('|', P.NewLines);
  Assert.Contains(Out_, '    FX: Integer;|  private|    function GetX: Integer;|  public',
    'a new private section BEFORE the first visibility keyword');
  Assert.Contains(Out_, '|function TBar.GetX: Integer;|begin|  Result := FX;|end;||end.');
end;

procedure TUserIdeas13Tests.PropertyConverter_TrivialAccessorsToFields;
begin
  var Src := TArray<string>.Create('unit UT;', 'interface', 'type', '  TFoo = class',
    '  private', '    FCount: Integer;', '    function GetCount: Integer;',
    '    procedure SetCount(const V: Integer);', '  public',
    '    property Count: Integer read GetCount write SetCount;', '  end;', '',
    'implementation', '',
    'function TFoo.GetCount: Integer;', 'begin', '  Exit(FCount);', 'end;', '',
    'procedure TFoo.SetCount(const V: Integer);', 'begin', '  FCount := V;', 'end;', '',
    'end.');
  var P := PlanPropertyConversion(Src, 9, 9, pcToFields, True, True, nil);
  Assert.AreEqual(1, P.OkCount);
  var Out_ := string.Join('|', P.NewLines);
  Assert.Contains(Out_, 'property Count: Integer read FCount write FCount;');
  Assert.DoesNotContain(Out_, 'GetCount');
  Assert.DoesNotContain(Out_, 'SetCount');
  Assert.Contains(Out_, 'implementation||end.');
end;

initialization
  TDUnitX.RegisterTestFixture(TUserIdeas13Tests);

end.
