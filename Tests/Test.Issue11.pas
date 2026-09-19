(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Tests for the features built from Ian Branch's ideas in issue #11:
///  aligning a method signature, "Remove with" without conditionals, the
///  rename conflict check, the E2291 quick fix, extract variable and wrap
///  in try..finally. The pure planners only - the editor side is thin.
/// </summary>
unit Test.Issue11;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TIssue11Tests = class
  public
    [Test] procedure AlignSignature_ReplacesSignatureKeepsDirectives;
    [Test] procedure AlignSignature_RefusesCommentInHeader;
    [Test] procedure RemoveWith_ConditionalsDetected;
    [Test] procedure RenameConflict_CountsCodeOnly;
    [Test] procedure E2291_MessageAndDeclaration;
    [Test] procedure E2291_InsertionIntoPublicSection;
    [Test] procedure ExtractVariable_PlainCase;
    [Test] procedure ExtractVariable_RefusesUnsafeSpots;
    [Test] procedure WrapTryFinally_InfersCleanup;
    [Test] procedure WrapTryFinally_WrapsAndRefusesPartialBlocks;
  end;

implementation

uses
  System.SysUtils, Expert.SignatureCheck, Expert.WithScanner, Expert.UnitIndex,
  Expert.AutoImport, Expert.StatementRefactor;

procedure TIssue11Tests.AlignSignature_ReplacesSignatureKeepsDirectives;
var
  R: TArray<string>;
begin
  Assert.AreEqual('procedure Bar(const A: Integer)',
    TSignatureChecker.SignatureForContainer('procedure TFoo.Bar(const A: Integer);', ''));
  Assert.AreEqual('function TFoo.Baz(X: Integer): string',
    TSignatureChecker.SignatureForContainer('function Baz(X: Integer): string', 'TFoo'));
  Assert.IsTrue(TSignatureChecker.ReplaceSignature(TArray<string>.Create(
    '  TFoo = class', '    class function Bar(A: Integer): Boolean; virtual; abstract;', '  end;'),
    1, 'function Bar(A: Int64): Boolean', R));
  Assert.AreEqual('    class function Bar(A: Int64): Boolean; virtual; abstract;', R[1]);
  Assert.IsTrue(TSignatureChecker.ReplaceSignature(TArray<string>.Create(
    '    procedure Bar(A: Integer;', '      B: string); override;', '  end;'),
    0, 'procedure Bar(A: Integer; const B: string)', R));
  Assert.AreEqual<Integer>(2, Length(R), 'a wrapped header becomes one line');
  Assert.AreEqual('    procedure Bar(A: Integer; const B: string); override;', R[0]);
end;

procedure TIssue11Tests.AlignSignature_RefusesCommentInHeader;
var
  R: TArray<string>;
begin
  Assert.IsFalse(TSignatureChecker.ReplaceSignature(TArray<string>.Create(
    '    procedure Bar(A: Integer; { doc }', '      B: string);'), 0, 'procedure Bar', R),
    'a comment inside the replaced span would be lost');
end;

procedure TIssue11Tests.RemoveWith_ConditionalsDetected;
begin
  Assert.IsFalse(TWithScanner.SourceHasConditionals(
    '{$R *.dfm} {$I+} {$I-} {$INLINE ON} begin with A do B := 1; end.'));
  for var S in TArray<string>.Create('{$IFDEF X}', '{$ifndef X}', '{$IF Defined(X)}',
    '{$IFOPT R+}', '{$ELSE}', '(*$IFDEF X*)', '{$I Some.inc}', '{$INCLUDE Some.inc}') do
    Assert.IsTrue(TWithScanner.SourceHasConditionals('begin ' + S + ' end'), S);
end;

procedure TIssue11Tests.RenameConflict_CountsCodeOnly;
var
  W: TArray<Integer>;
begin
  W := CodeWordLines(TArray<string>.Create(
    'var Count: Integer;', '// Count in a comment', 'S := ''Count'';',
    'X := FCount + Counter;', 'Self.Count := 1;'), 'count');
  Assert.AreEqual<Integer>(2, Length(W));
  Assert.AreEqual<Integer>(0, W[0]);
  Assert.AreEqual<Integer>(4, W[1]);
end;

procedure TIssue11Tests.E2291_MessageAndDeclaration;
var
  Intf, Meth, Decl: string;
begin
  Assert.IsTrue(E2291MethodFromMessage('E2291 Missing implementation of interface method IFoo.Bar', Intf, Meth));
  Assert.AreEqual('IFoo', Intf);
  Assert.AreEqual('Bar', Meth);
  Assert.IsTrue(E2291MethodFromMessage('E2291 Fehlende Implementierung ''Sys.IFoo.Bar''.', Intf, Meth));
  Assert.AreEqual('Sys.IFoo', Intf);
  Assert.IsTrue(InterfaceMethodDecl(TArray<string>.Create(
    '    function Get(const I: Integer): string; stdcall;'), 0, Decl));
  Assert.AreEqual('function Get(const I: Integer): string; stdcall;', Decl,
    'the calling convention is part of the interface contract');
  Assert.IsFalse(InterfaceMethodDecl(TArray<string>.Create(
    '    procedure Put(X: Integer); overload;'), 0, Decl), 'overloads are ambiguous');
end;

procedure TIssue11Tests.E2291_InsertionIntoPublicSection;
var
  Ins, First: Integer;
  Text: string;
begin
  Assert.IsTrue(PlanInterfaceMethodInsertion(TArray<string>.Create(
    '  TFoo = class(TInterfacedObject, IFoo)', '  private', '    FX: Integer;',
    '  public', '    procedure Qux;', '  end;'),
    0, TArray<string>.Create('procedure Bar;'), Ins, Text, First));
  Assert.AreEqual<Integer>(5, Ins, 'end of the public section');
  Assert.AreEqual('    procedure Bar;' + sLineBreak, Text);
  Assert.IsTrue(PlanInterfaceMethodInsertion(TArray<string>.Create(
    '  TFoo = class(TInterfacedObject, IFoo)', '  private', '    FX: Integer;', '  end;'),
    0, TArray<string>.Create('procedure Bar;'), Ins, Text, First));
  Assert.AreEqual<Integer>(3, Ins);
  Assert.AreEqual<Integer>(4, First);
  Assert.AreEqual('  public' + sLineBreak + '    procedure Bar;' + sLineBreak, Text,
    'no public section yet - one is added');
end;

function ExtractSource: TArray<string>;
begin
  Result := TArray<string>.Create(
    'procedure P;',                                   // 0
    'begin',                                          // 1
    '  Total := Foo.Bar.Count + 1;',                  // 2
    '  if Assigned(X) and (X.Foo > 0) then',          // 3
    '    Y := 1;',                                    // 4
    '  if A then',                                    // 5
    '    Z := Calc(3);',                              // 6
    '  while Q.Next do Step;',                        // 7
    '  if B then W := Sum(4);',                       // 8
    '  S := ''Foo.Bar.Count'';',                      // 9
    'end;');                                          // 10
end;

procedure TIssue11Tests.ExtractVariable_PlainCase;
var
  Plan: TExtractVarPlan;
  Why: string;
begin
  Assert.AreEqual('LCount', SuggestVariableName('Foo.Bar.Count'));
  Assert.AreEqual('LName', SuggestVariableName('GetName(X, Y)'));
  Assert.AreEqual('LCaption', SuggestVariableName('Items[I].Caption'));
  Assert.IsTrue(PlanExtractVariable(ExtractSource, 2, 11, 24, 'LCount', Plan, Why), Why);
  Assert.AreEqual('  var LCount := Foo.Bar.Count;', Plan.NewLines[2]);
  Assert.AreEqual('  Total := LCount + 1;', Plan.NewLines[3]);
end;

procedure TIssue11Tests.ExtractVariable_RefusesUnsafeSpots;
var
  Plan: TExtractVarPlan;
  Why: string;
begin
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 3, 22, 27, 'LFoo', Plan, Why),
    'after a short-circuit "and" - hoisting would dereference nil');
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 6, 9, 16, 'LCalc', Plan, Why),
    'sole statement of a then-branch');
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 7, 8, 14, 'LNext', Plan, Why),
    'loop condition - evaluated per iteration');
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 8, 17, 23, 'LSum', Plan, Why),
    'branch on the same line');
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 9, 8, 21, 'LText', Plan, Why),
    'inside a string');
  Assert.IsFalse(PlanExtractVariable(ExtractSource, 2, 11, 24, 'Total', Plan, Why),
    'name already used in the routine');
end;

procedure TIssue11Tests.WrapTryFinally_InfersCleanup;
begin
  Assert.AreEqual('L.Free', InferCleanup('  L := TStringList.Create;'));
  Assert.AreEqual('Obj.Free', InferCleanup('  var Obj: TFoo := TFoo.Create(Self);'));
  Assert.AreEqual('Memo1.Lines.EndUpdate', InferCleanup('  Memo1.Lines.BeginUpdate;'));
  Assert.AreEqual('FLock.Leave', InferCleanup('  FLock.Enter;'));
  Assert.AreEqual('TMonitor.Exit(FList)', InferCleanup('  TMonitor.Enter(FList);'));
  Assert.AreEqual('', InferCleanup('  X := 5;'));
end;

procedure TIssue11Tests.WrapTryFinally_WrapsAndRefusesPartialBlocks;
var
  Src, R: TArray<string>;
  Why: string;
begin
  Src := TArray<string>.Create('  L := TStringList.Create;', '  L.Add(''a'');',
    '  if X then', '  begin', '    Y;', '  end;', '  Z;');
  Assert.IsTrue(PlanWrapTryFinally(Src, 1, 5, 'L.Free', R, Why), Why);
  Assert.AreEqual('  L := TStringList.Create;|  try|    L.Add(''a'');|    if X then|' +
    '    begin|      Y;|    end;|  finally|    L.Free;|  end;|  Z;', string.Join('|', R));
  Assert.IsFalse(PlanWrapTryFinally(Src, 1, 4, '', R, Why), 'begin without its end');
end;

initialization
  TDUnitX.RegisterTestFixture(TIssue11Tests);

end.
