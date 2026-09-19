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
    [Test] procedure SafeDelete_MethodDeclarationAndBody;
    [Test] procedure SafeDelete_VetoesDispatchAndPublished;
    [Test] procedure SafeDelete_DataDeclarations;
    [Test] procedure Scanner_TokensAndPositions;
    [Test] procedure Scanner_MultiLineStringIsOneToken;
    [Test] procedure Scanner_UnicodeIdentifiersAndEscapes;
    [Test] procedure Scanner_StripAndMask;
  end;

implementation

uses
  System.SysUtils, Expert.SignatureCheck, Expert.WithScanner, Expert.UnitIndex,
  Expert.AutoImport, Expert.StatementRefactor, Expert.SafeDeletePlan,
  Expert.PascalScanner;

const
  SafeDemo: array[0..28] of string = (
    'unit Demo;',                                        // 0
    'interface',                                         // 1
    'type',                                              // 2
    '  TFoo = class',                                    // 3
    '  private',                                         // 4
    '    FA, FB: Integer;',                              // 5
    '    procedure Unused;',                             // 6
    '    procedure Virt; virtual;',                      // 7
    '  published',                                       // 8
    '    procedure Pub;',                                // 9
    '  end;',                                            // 10
    '',                                                  // 11
    'const',                                             // 12
    '  CLone = 5;',                                      // 13
    '',                                                  // 14
    'implementation',                                    // 15
    '',                                                  // 16
    'procedure TFoo.Unused;',                            // 17
    'begin',                                             // 18
    'end;',                                              // 19
    '',                                                  // 20
    'procedure TFoo.Virt;',                              // 21
    'begin',                                             // 22
    'end;',                                              // 23
    '',                                                  // 24
    'procedure TFoo.Pub;',                               // 25
    'begin',                                             // 26
    'end;',                                              // 27
    'end.');                                             // 28

function SafeDemoLines: TArray<string>;
begin
  SetLength(Result, Length(SafeDemo));
  for var I := 0 to High(SafeDemo) do Result[I] := SafeDemo[I];
end;

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

procedure TIssue11Tests.SafeDelete_MethodDeclarationAndBody;
var
  Sym: TSafeDeleteSymbol;
  Why: string;
begin
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 6, 'Unused', Sym, Why), Why);
  Assert.AreEqual<Integer>(Ord(sdkMethod), Ord(Sym.Kind));
  Assert.AreEqual('TFoo', Sym.Container);
  Assert.AreEqual(17, Sym.ImplLine);
  Assert.AreEqual<Integer>(0, Length(Sym.Vetoes));
  Assert.AreEqual<Integer>(2, Length(Sym.Edits));
  var Res := ApplySafeDeleteEdits(SafeDemoLines, Sym.Edits);
  Assert.IsFalse(string.Join('|', Res).Contains('Unused'));
  Assert.AreEqual<Integer>(Length(SafeDemo) - 5, Length(Res));
  // the implementation header leads to the same plan
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 17, 'Unused', Sym, Why), Why);
  Assert.AreEqual(6, Sym.DeclLine);
  // a line that does not declare the name
  Assert.IsFalse(PlanSafeDeleteSymbol(SafeDemoLines, 18, 'Unused', Sym, Why));
end;

procedure TIssue11Tests.SafeDelete_VetoesDispatchAndPublished;
var
  Sym: TSafeDeleteSymbol;
  Why: string;
begin
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 7, 'Virt', Sym, Why), Why);
  Assert.AreEqual<Integer>(1, Length(Sym.Vetoes), 'virtual');
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 9, 'Pub', Sym, Why), Why);
  Assert.AreEqual<Integer>(1, Length(Sym.Vetoes), 'published');
end;

procedure TIssue11Tests.SafeDelete_DataDeclarations;
var
  Sym: TSafeDeleteSymbol;
  Why: string;
begin
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 5, 'FB', Sym, Why), Why);
  Assert.AreEqual<Integer>(Ord(sdkField), Ord(Sym.Kind));
  Assert.IsTrue(Sym.Edits[0].HasReplacement);
  Assert.AreEqual('    FA: Integer;', Sym.Edits[0].Replacement);
  Assert.IsTrue(PlanSafeDeleteSymbol(SafeDemoLines, 13, 'CLone', Sym, Why), Why);
  Assert.AreEqual<Integer>(Ord(sdkConstant), Ord(Sym.Kind));
  Assert.AreEqual(12, Sym.Edits[0].FirstLine, 'the lone "const" goes with it');
  var Hits := FormTextMentions('object B: TButton'#13#10'  OnClick = UnusedClick', 'UnusedClick');
  Assert.AreEqual<Integer>(1, Length(Hits));
  Assert.AreEqual(1, Hits[0]);
end;

function ScanAll(const AText: string; AComments: Boolean = False): TArray<TPasToken>;
var
  S: TPascalScanner;
  T: TPasToken;
begin
  Result := nil;
  S := TPascalScanner.Create(AText, AComments, AComments);
  try
    while S.Next(T) do Result := Result + [T];
  finally
    S.Free;
  end;
end;

procedure TIssue11Tests.Scanner_TokensAndPositions;
begin
  var T := ScanAll('X := A[1..5];'#13#10'  Y := $FF + 1.5e3; // c');
  Assert.AreEqual<Integer>(15, Length(T));
  Assert.IsTrue(T[1].IsSymbol(':='));
  Assert.AreEqual('1', T[4].Text, '"1..5" is a range, not a float');
  Assert.IsTrue(T[5].IsSymbol('..'));
  Assert.AreEqual(1, T[9].Line);
  Assert.AreEqual(2, T[9].Col);
  Assert.AreEqual('$FF', T[11].Text);
  Assert.AreEqual<Integer>(Ord(ptNumber), Ord(T[13].Kind));
  var C := ScanAll('{$IFDEF X}a (* b *) // c', True);
  Assert.AreEqual<Integer>(Ord(ptDirective), Ord(C[0].Kind));
  Assert.AreEqual<Integer>(Ord(ptComment), Ord(C[2].Kind));
  Assert.AreEqual<Integer>(Ord(ptComment), Ord(C[3].Kind));
end;

procedure TIssue11Tests.Scanner_MultiLineStringIsOneToken;
begin
  var T := ScanAll('S := '''''''#13#10'  begin end'#13#10'  '''''';'#13#10'X');
  Assert.AreEqual<Integer>(5, Length(T), 'begin/end inside the literal are no tokens');
  Assert.AreEqual<Integer>(Ord(ptString), Ord(T[2].Kind));
  Assert.IsTrue(T[3].IsSymbol(';'));
  Assert.AreEqual(2, T[3].Line);
  Assert.AreEqual(3, MultiLineStringOpener('x := ''''''', 6));
  Assert.AreEqual(0, MultiLineStringOpener('''''''x', 1));
end;

procedure TIssue11Tests.Scanner_UnicodeIdentifiersAndEscapes;
var
  S: TPascalScanner;
  Tok, All: string;
begin
  Assert.IsTrue(IsIdentifier('Größe'));
  Assert.IsTrue(IsIdentifier('_x1'));
  Assert.IsFalse(IsIdentifier('1x'));
  Assert.IsFalse(IsIdentifier('a.b'));
  var T := ScanAll('&begin Größe');
  Assert.AreEqual<Integer>(2, Length(T));
  Assert.IsFalse(T[0].IsWord('begin'), 'an escaped identifier is not the keyword');
  Assert.AreEqual('Größe', T[1].Text);
  // the selection validator's stream
  All := '';
  S := TPascalScanner.Create('if a(x, ''s'') then c[1] := 2; (.x.)');
  try
    while S.NextToken(Tok) do All := All + Tok + ' ';
  finally
    S.Free;
  end;
  Assert.AreEqual('IF A ( X ) THEN C [ ] ; [ X ] ', All);
end;

procedure TIssue11Tests.Scanner_StripAndMask;
begin
  Assert.AreEqual('  Url = ''http://x'';', StripLineComment('  Url = ''http://x''; // note'));
  Assert.AreEqual('  X := 1;', StripLineComment('  X := 1;'));
  var M := MaskCommentsAndStrings(TArray<string>.Create(
    'A := ''x''; { c', 'still } B', 'S := ''''''', '  Name', '  ''''''; C // d'));
  Assert.AreEqual('A :=    ;    ', M[0]);
  Assert.AreEqual('        B', M[1]);
  Assert.AreEqual('      ', M[3], 'content of a multi-line string');
  Assert.AreEqual('     ; C     ', M[4]);
end;

initialization
  TDUnitX.RegisterTestFixture(TIssue11Tests);

end.
