(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Where PlanEventHandler (Expert.EventGen) puts the declaration of a
///  generated event handler, driven by in-memory unit text.
/// </summary>
unit Test.EventGenPlan;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TEventGenPlanTests = class
  public
    /// <summary>The declaration goes at the END of the private section,
    ///  after the fields: straight after "private" it came ahead of them,
    ///  which does not compile (E2169).</summary>
    [Test] procedure Declaration_GoesAfterThePrivateFields;
    /// <summary>The first "private" in the text belongs to a nested class;
    ///  the handler goes into the outer class's own private section.</summary>
    [Test] procedure NestedClassPrivate_IsNotTheClasssPrivate;
  end;

implementation

uses
  System.SysUtils, Expert.EventGen;

function ClickInfo: TProcTypeInfo;
begin
  Result := Default(TProcTypeInfo);
  Result.Kind := pkMethod;
  Result.Params := 'Sender: TObject';
end;

procedure TEventGenPlanTests.Declaration_GoesAfterThePrivateFields;
var
  Plan: TGenPlan;
begin
  Plan := PlanEventHandler([
    'unit U;',                          // 0
    'interface',                        // 1
    'type',                             // 2
    '  TForm1 = class(TForm)',          // 3
    '  private',                        // 4
    '    FCount: Integer;',             // 5
    '    FName: string;',               // 6
    '  public',                         // 7
    '    procedure Test;',              // 8
    '  end;',                           // 9
    'implementation',                   // 10
    'procedure TForm1.Test;',           // 11
    'begin',                            // 12
    '  Button1.OnClick := ',            // 13
    'end;',                             // 14
    '',                                 // 15
    'end.'], 13, 'Button1Click', ClickInfo);
  Assert.IsTrue(Plan.Ok, Plan.Reason);
  Assert.AreEqual(7, Plan.DeclLine0, 'before "public", after FCount and FName');
  Assert.IsTrue(Plan.DeclText.Contains('procedure Button1Click(Sender: TObject);'), Plan.DeclText);
end;

procedure TEventGenPlanTests.NestedClassPrivate_IsNotTheClasssPrivate;
var
  Plan: TGenPlan;
begin
  Plan := PlanEventHandler([
    'unit U;',                          // 0
    'interface',                        // 1
    'type',                             // 2
    '  TForm1 = class(TForm)',          // 3
    '  public type',                    // 4
    '    TInner = class',               // 5
    '    private',                      // 6
    '      FZ: Integer;',               // 7
    '    end;',                         // 8
    '  private',                        // 9
    '    FCount: Integer;',             // 10
    '  public',                         // 11
    '    procedure Test;',              // 12
    '  end;',                           // 13
    'implementation',                   // 14
    'procedure TForm1.Test;',           // 15
    'begin',                            // 16
    '  Button1.OnClick := ',            // 17
    'end;',                             // 18
    '',                                 // 19
    'end.'], 17, 'Button1Click', ClickInfo);
  Assert.IsTrue(Plan.Ok, Plan.Reason);
  Assert.AreEqual(11, Plan.DeclLine0, 'TForm1''s own private section, after FCount');
end;

initialization
  TDUnitX.RegisterTestFixture(TEventGenPlanTests);

end.
