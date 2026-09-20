(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Change method signature (issue #11, suggestion 9 by Ian Branch): the
///  pure half - parameter / argument lists, how an occurrence is used, the
///  plan for a family of headers plus its calls, and the override chain.
/// </summary>
unit Test.ChangeSignature;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TChangeSignatureTests = class
  public
    [Test] procedure Arguments_ReorderAndDefaults;
    [Test] procedure CallContext_Classifies;
    [Test] procedure Plan_HeadersBodyAndCalls;
    [Test] procedure Plan_BlocksWhatCannotFollow;
    [Test] procedure Plan_InheritedCallPassesNewParameter;
    [Test] procedure TypeGraph_OverrideChain;
    [Test] procedure MemberDeclaration_AfterNestedClasses;
    [Test] procedure ReferenceKinds_Classified;
    [Test] procedure BatchFixes_And_SemanticVerdicts;
    [Test] procedure MemberResolution_WithoutLsp;
    [Test] procedure MemberResolution_PrivatePublicOverloadAcrossUnits;
    [Test] procedure MemberResolution_UnqualifiedOverloadByArgumentType;
    [Test] procedure MemberResolution_SelfQualifiedWrappedCall;
  end;

implementation

uses
  System.SysUtils, System.Types, System.IOUtils, Expert.SignatureEdit, Expert.InterfaceLinks,
  Expert.UnitIndex, Expert.ReferenceKind, Expert.PascalScanner, Expert.SemanticReplace,
  Expert.AutoImport;

const
  Src: array[0..18] of string = (
    'unit U1;',
    'interface',
    'type',
    '  TFoo = class',
    '    procedure Bar(A: Integer; B: string = ''b'');',
    '    function Get: Integer;',
    '  end;',
    'implementation',
    'procedure TFoo.Bar(A: Integer; B: string);',
    'begin',
    '  Writeln(A, B, Self.A);',
    '  if A > 0 then Bar(A - 1);',
    'end;',
    'function TFoo.Get: Integer;',
    'begin',
    '  Bar(1, ''x'');',
    '  Result := 0;',
    'end;',
    'end.');

function Content: string;
begin
  Result := string.Join(#13#10, Src);
end;

procedure Setup(out AHeaders: TArray<TSigHeader>; out ACalls: TArray<TSigCall>;
  out ASources: TArray<TSigSource>);
var
  HD, HI: TSigHeader;
  Why: string;
begin
  Assert.IsTrue(LocateSigHeader(Content, 4, 'Bar', True, HD, Why), Why);
  Assert.IsTrue(LocateSigHeader(Content, 8, 'Bar', False, HI, Why), Why);
  AHeaders := [HD, HI];
  ACalls := LocateSigCalls(Content, [Point(Pos('Bar(', Src[11]) - 1, 11),
    Point(Pos('Bar(', Src[15]) - 1, 15)], 3, 0);
  SetLength(ASources, 1);
  ASources[0].FilePath := 'U1.pas';
  ASources[0].Content := Content;
end;

procedure TChangeSignatureTests.Arguments_ReorderAndDefaults;
var
  T, Note, Err: string;
begin
  var Old := ParseParamList('A: Integer; B: string = ''b''; C: Boolean = True');
  Assert.AreEqual(3, Integer(Length(Old)));
  var New_: TArray<TNewParam>;
  SetLength(New_, 3);
  New_[0].Param := Old[2]; New_[0].Param.DefaultText := ''; New_[0].OldIndex := 2;
  New_[1].Param := Old[0]; New_[1].OldIndex := 0;
  New_[2].Param := ParseParamList('D: Integer')[0]; New_[2].OldIndex := -1; New_[2].CallValue := '0';
  Assert.AreEqual('', ValidateSignatureChange(Old, New_));
  // C's default the call relied on is written out now
  Assert.IsTrue(RewriteArguments(['X'], Old, New_, T, Note, Err), Err);
  Assert.AreEqual('True, X, 0', T);
  var S := 'Foo(TDictionary<string, Integer>.Create, a < b, ''x,y'')';
  var Args := SplitArguments(S, 4, MatchingBracket(S, 4));
  Assert.AreEqual(3, Integer(Length(Args)));
  Assert.AreEqual('TDictionary<string, Integer>.Create', Args[0].Text);
end;

procedure TChangeSignatureTests.CallContext_Classifies;

  function Ctx(const AText: string): TCallContext;
  var
    O, C: Integer;
  begin
    var P := Pos('Foo', AText);
    Result := CallContextAt(AText, P, P + 2, O, C);
  end;

begin
  Assert.IsTrue(Ctx('  Foo(1);') = ccArgs);
  Assert.IsTrue(Ctx('  Obj.Foo;') = ccStatement);
  Assert.IsTrue(Ctx('  inherited Foo;') = ccStatement);
  Assert.IsTrue(Ctx('  if Foo then Exit;') = ccExpression);
  Assert.IsTrue(Ctx('  Btn.OnClick := Foo;') = ccReference);
  Assert.IsTrue(Ctx('  P := @Foo;') = ccReference);
  Assert.IsTrue(Ctx('  property X: Integer read Foo;') = ccAccessor);
end;

procedure TChangeSignatureTests.Plan_HeadersBodyAndCalls;
var
  H: TArray<TSigHeader>;
  C: TArray<TSigCall>;
  S: TArray<TSigSource>;
begin
  Setup(H, C, S);
  var New_: TArray<TNewParam>;
  SetLength(New_, 3);
  New_[0].Param := H[0].Params[0]; New_[0].Param.Name := 'Count'; New_[0].OldIndex := 0;
  New_[1].Param := ParseParamList('C: Boolean')[0]; New_[1].OldIndex := -1; New_[1].CallValue := 'True';
  New_[2].Param := H[0].Params[1]; New_[2].OldIndex := 1;
  var Plan := PlanSignatureEdits(S, H, C, New_);
  Assert.IsTrue(Plan.Ok, string.Join('|', Plan.Errors));
  var Out_ := ApplySigEdits(Content, Plan.Edits, 0).Split([#10]);
  Assert.AreEqual('    procedure Bar(Count: Integer; C: Boolean; B: string = ''b'');', Out_[4]);
  Assert.AreEqual('procedure TFoo.Bar(Count: Integer; C: Boolean; B: string);', Out_[8]);
  Assert.AreEqual('  Writeln(Count, B, Self.A);', Out_[10]);
  Assert.AreEqual('  if Count > 0 then Bar(Count - 1, True);', Out_[11]);
  Assert.AreEqual('  Bar(1, True, ''x'');', Out_[15]);
end;

procedure TChangeSignatureTests.Plan_BlocksWhatCannotFollow;
var
  H: TArray<TSigHeader>;
  C: TArray<TSigCall>;
  S: TArray<TSigSource>;
begin
  Setup(H, C, S);
  // removing A - the body still uses it
  var New_: TArray<TNewParam>;
  SetLength(New_, 1);
  New_[0].Param := H[0].Params[1]; New_[0].OldIndex := 1;
  var Plan := PlanSignatureEdits(S, H, C, New_);
  Assert.IsFalse(Plan.Ok);
  Assert.IsTrue(Pos('still used', string.Join('|', Plan.Errors)) > 0);
  // a method reference blocks a real change, not a rename
  var Ref := C[1];
  Ref.Context := ccReference;
  New_ := UnchangedSignature(H[0].Params);
  New_[0].Param.TypeText := 'Int64';
  Plan := PlanSignatureEdits(S, H, [C[0], Ref], New_);
  Assert.IsTrue(Pos('method reference', string.Join('|', Plan.Errors)) > 0);
  New_ := UnchangedSignature(H[0].Params);
  New_[1].Param.Name := 'Text';
  Plan := PlanSignatureEdits(S, H, [C[0], Ref], New_);
  Assert.IsTrue(Plan.Ok, string.Join('|', Plan.Errors));
  // a comment inside a parameter list is left to the user
  var HX: TSigHeader;
  var Why: string;
  Assert.IsFalse(LocateSigHeader('procedure Foo(A: Integer { n });', 0, 'Foo', True, HX, Why));
end;

procedure TChangeSignatureTests.Plan_InheritedCallPassesNewParameter;
begin
    // "inherited Go(...)" in an override hands the new parameter on
    var ISrc := string.Join(#13#10, ['unit U2;', 'interface', 'type', '  TBase = class',
      '    function Go(Count: Integer): Integer; virtual;', '  end;', '  TChild = class(TBase)',
      '    function Go(Count: Integer): Integer; override;', '  end;', 'implementation',
      'function TBase.Go(Count: Integer): Integer;', 'begin', '  Result := Count;', 'end;',
      'function TChild.Go(Count: Integer): Integer;', 'begin',
      '  Result := inherited Go(Count) + 1;', 'end;', 'end.']);
    var IH: TArray<TSigHeader>;
    SetLength(IH, 4);
    var IWhy: string;
    var IOk := LocateSigHeader(ISrc, 4, 'Go', True, IH[0], IWhy) and
      LocateSigHeader(ISrc, 7, 'Go', True, IH[1], IWhy) and
      LocateSigHeader(ISrc, 10, 'Go', False, IH[2], IWhy) and
      LocateSigHeader(ISrc, 14, 'Go', False, IH[3], IWhy);
    var ILines := ISrc.Replace(#13#10, #10).Split([#10]);
    var ICalls := LocateSigCalls(ISrc, [Point(Pos('Go(', ILines[16]) - 1, 16)], 2, 0);
    var ISources: TArray<TSigSource>;
    SetLength(ISources, 1);
    ISources[0].FilePath := 'U2.pas';
    ISources[0].Content := ISrc;
    var INew: TArray<TNewParam>;
    SetLength(INew, 2);
    INew[0].Param := IH[0].Params[0]; INew[0].Param.Name := 'N'; INew[0].OldIndex := 0;
    INew[1].Param := ParseParamList('Scale: Integer = 1')[0]; INew[1].OldIndex := -1;
    var IPlan := PlanSignatureEdits(ISources, IH, ICalls, INew);
    var IOut := ApplySigEdits(ISrc, IPlan.Edits, 0).Split([#10]);
  Assert.IsTrue(IOk, IWhy);
  Assert.IsTrue(IPlan.Ok, string.Join('|', IPlan.Errors));
  Assert.AreEqual('  Result := inherited Go(N, Scale) + 1;', IOut[16]);
  Assert.AreEqual('  Result := N;', IOut[12]);
  Assert.AreEqual('    function Go(N: Integer; Scale: Integer = 1): Integer; override;', IOut[7]);
end;

procedure TChangeSignatureTests.TypeGraph_OverrideChain;
begin
  var Dir := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_chain');
  TDirectory.CreateDirectory(Dir);
  var F := TPath.Combine(Dir, 'UC.pas');
  TFile.WriteAllText(F, string.Join(#13#10, [
    'unit UC;',
    'interface',
    'type',
    '  TBase = class',
    '    procedure Run(A: Integer); virtual;',
    '  end;',
    '  TMid = class(TBase)',
    '    procedure Run(A: Integer); override;',
    '  end;',
    '  TLeaf = class(TMid)',
    '    procedure Run(A: Integer); override;',
    '  end;',
    '  TOther = class',
    '    procedure Run(A: Integer); virtual;',
    '  end;',
    'implementation',
    'procedure TBase.Run(A: Integer); begin end;',
    'procedure TMid.Run(A: Integer); begin end;',
    'procedure TLeaf.Run(A: Integer); begin end;',
    'procedure TOther.Run(A: Integer); begin end;',
    'end.']));
  try
    var G := TTypeGraph.Create([F], nil, False);
    try
      // from the middle: up to the introducing class, down to every override
      var L := G.ClassHierarchyMembers('TMid', 'Run');
      var Names := '';
      for var X in L do Names := Names + X.TypeName + ';';
      Assert.AreEqual(3, Integer(Length(L)), Names);
      Assert.IsTrue(Pos('TBase;', Names) > 0, Names);
      Assert.IsTrue(Pos('TLeaf;', Names) > 0, Names);
      Assert.IsTrue(Pos('TOther', Names) = 0, Names);
      for var X in L do
        Assert.IsTrue(X.ImplLine > 0, X.TypeName + ' has no implementation line');
    finally
      G.Free;
    end;
  finally
    TDirectory.Delete(Dir, True);
  end;
end;

procedure TChangeSignatureTests.MemberDeclaration_AfterNestedClasses;
begin
  // members after NESTED classes (change signature on TUnitIndex.Search)
  var NSrc := string.Join(#13#10, ['unit N;', 'interface', 'type', '  TOuter = class',
    '  private type', '    TInner = class', '      procedure Run;', '    end;',
    '    TWorker = class(TThread)', '    protected', '      procedure Execute; override;',
    '    end;', '    TFwd = class;', '    TMeta = class of TObject;', '  public',
    '    class procedure Make;', '    function Search(const S: string): Integer;', '  end;',
    'implementation', 'end.']);
  Assert.AreEqual(16, FindMemberDeclarationLine(NSrc, 'TOuter', 'Search'));
  Assert.AreEqual(15, FindMemberDeclarationLine(NSrc, 'TOuter', 'Make'));
  Assert.AreEqual(-1, FindMemberDeclarationLine(NSrc, 'TOuter', 'Execute'));
  Assert.AreEqual(-1, FindMemberDeclarationLine(NSrc, 'TOuter', 'Run'));
  Assert.AreEqual(6, FindMemberDeclarationLine(NSrc, 'TInner', 'Run'));
end;

procedure TChangeSignatureTests.ReferenceKinds_Classified;
begin
  var KSrc := string.Join(#13#10, [
    'unit K;',                                            // 0
    'interface',                                          // 1
    'uses',                                               // 2
    '  System.SysUtils, KUnit;',                          // 3
    'type',                                               // 4
    '  TFoo = class(TBase)',                              // 5
    '    FCount: Integer;',                               // 6
    '    procedure Run(A: Integer);',                     // 7
    '    function Get: Integer;',                         // 8
    '    property Count: Integer read FCount write FCount;', // 9
    '  end;',                                             // 10
    'const',                                              // 11
    '  cRed = 1;',                                        // 12
    'implementation',                                     // 13
    'procedure TFoo.Run(A: Integer);',                    // 14
    'var',                                                // 15
    '  X: TFoo;',                                         // 16
    'begin',                                              // 17
    '  FCount := A;',                                     // 18
    '  X := Self; X.Run(1);',                             // 19
    '  inherited Run(A);',                                // 20
    '  if Get > 0 then Inc(FCount);',                     // 21
    '  case A of',                                        // 22
    '    cRed: Run(3);',                                  // 23
    '  end;',                                             // 24
    '  P := @Run; OnDone := Run; Y := Get;',              // 25
    '  Run;',                                             // 26
    'end;',                                               // 27
    'end.']);                                             // 28
  var KLines := KSrc.Replace(#13#10, #10).Split([#10]);
  // the kind of the N-th (1-based) whole-word occurrence of AWord on ALine
  var KindAt := function(ALine: Integer; const AWord: string; ANth: Integer;
    ASym: TRefSymbolKind): string
    begin
      Result := '?';
      var S := KLines[ALine];
      var Hit := 0;
      for var P := 1 to Length(S) - Length(AWord) + 1 do
        if (Copy(S, P, Length(AWord)) = AWord) and ((P = 1) or not IsIdentChar(S[P - 1])) and
           ((P + Length(AWord) > Length(S)) or not IsIdentChar(S[P + Length(AWord)])) then
        begin
          Inc(Hit);
          if Hit = ANth then
          begin
            var R := ClassifyReferences(KSrc, [Point(P - 1, ALine)], Length(AWord), ASym);
            Exit(RefKindText(R[0]));
          end;
        end;
    end;
  var KExpect: TArray<TArray<string>> := [
    // line, word, nth, symbol kind, expected
    ['6', 'FCount', '1', 'D', 'Declaration'],
    ['9', 'FCount', '1', 'D', 'Property accessor'],
    ['9', 'FCount', '2', 'D', 'Property accessor'],
    ['18', 'FCount', '1', 'D', 'Write'],
    ['21', 'FCount', '1', 'D', 'Read'],
    ['7', 'Run', '1', 'P', 'Declaration'],
    ['14', 'Run', '1', 'P', 'Implementation'],
    ['19', 'Run', '1', 'P', 'Call'],
    ['20', 'Run', '1', 'P', 'Inherited call'],
    ['23', 'Run', '1', 'P', 'Call'],
    ['25', 'Run', '1', 'P', 'Method reference'],
    ['25', 'Run', '2', 'P', 'Method reference'],
    ['26', 'Run', '1', 'P', 'Call'],
    ['8', 'Get', '1', 'F', 'Declaration'],
    ['21', 'Get', '1', 'F', 'Call'],
    ['25', 'Get', '1', 'F', 'Call'],
    ['12', 'cRed', '1', 'D', 'Declaration'],
    ['23', 'cRed', '1', 'D', 'Read'],
    ['5', 'TFoo', '1', 'T', 'Declaration'],
    ['14', 'TFoo', '1', 'T', 'Type use'],
    ['16', 'TFoo', '1', 'T', 'Type use'],
    ['5', 'TBase', '1', 'T', 'Type use'],
    ['3', 'KUnit', '1', 'U', 'Uses clause']];
  var KFails := '';
  for var X in KExpect do
  begin
    var Sym := rsUnknown;
    case X[3][1] of
      'D': Sym := rsData;
      'P': Sym := rsProcedure;
      'F': Sym := rsFunction;
      'T': Sym := rsType;
    end;
    var Got := KindAt(StrToInt(X[0]), X[1], StrToInt(X[2]), Sym);
    if Got <> X[4] then
      KFails := KFails + Format(' %s@%s#%s=%s(want %s)', [X[1], X[0], X[2], Got, X[4]]);
  end;
  // the symbol's own kind from its declaration line
  if SymbolKindFromDeclLine('    procedure Run(A: Integer);', 'Run') <> rsProcedure then KFails := KFails + ' sk-proc';
  if SymbolKindFromDeclLine('    class function Get: Integer;', 'Get') <> rsFunction then KFails := KFails + ' sk-func';
  if SymbolKindFromDeclLine('    FA, FCount: Integer;', 'FCount') <> rsData then KFails := KFails + ' sk-field';
  if SymbolKindFromDeclLine('  TFoo = class(TBase)', 'TFoo') <> rsType then KFails := KFails + ' sk-class';
  if SymbolKindFromDeclLine('  cRed = 1;', 'cRed') <> rsData then KFails := KFails + ' sk-const';
  if SymbolKindFromDeclLine('  TMyInt = Integer;', 'TMyInt') <> rsType then KFails := KFails + ' sk-alias';
  if SymbolKindFromDeclLine('  TKind = (kA, kB);', 'TKind') <> rsType then KFails := KFails + ' sk-enum';
  if SymbolKindFromDeclLine('    property Count: Integer read FCount;', 'Count') <> rsData then KFails := KFails + ' sk-prop';
  if SymbolKindFromDeclLine('procedure TFoo.Run(A: Integer);', 'Run') <> rsProcedure then KFails := KFails + ' sk-impl';
  Assert.AreEqual('', KFails, 'reference kinds:' + KFails);
end;


procedure TChangeSignatureTests.BatchFixes_And_SemanticVerdicts;
begin
  var BFails := '';
  // ---- semantic replace: the verdicts and leaving matches out
  var SRules: TArray<TSemanticReplaceRule>;
  SetLength(SRules, 2);
  SRules[0].Find := 'A.Go'; SRules[0].Replace := 'B.Go';
  SRules[1].Find := 'X.Run'; SRules[1].Replace := 'Y.Run'; SRules[1].DeclaredIn := 'Vcl.Forms';
  var STargets: TArray<TMatchTarget>;
  SetLength(STargets, 6);
  STargets[0].RuleIdx := 0; STargets[0].TargetFile := 'C:\p\x.pas'; STargets[0].TargetLine := 10;
  STargets[1].RuleIdx := 0; STargets[1].TargetFile := 'C:\p\x.pas'; STargets[1].TargetLine := 10;
  STargets[2].RuleIdx := 0; STargets[2].TargetFile := 'C:\p\y.pas'; STargets[2].TargetLine := 3;
  STargets[3].RuleIdx := 0; STargets[3].TargetFile := ''; STargets[3].TargetLine := -1;
  STargets[4].RuleIdx := 1; STargets[4].TargetFile := 'C:\bds\Vcl.Forms.pas'; STargets[4].TargetLine := 5;
  STargets[5].RuleIdx := 1; STargets[5].TargetFile := 'C:\p\Other.pas'; STargets[5].TargetLine := 1;
  var SDom: TArray<string>;
  var SV := TSemanticReplaceEngine.VerifyVerdicts(SRules, STargets, SDom);
  if (SV[0] <> mvVerified) or (SV[1] <> mvVerified) or (SV[2] <> mvOtherSymbol) or
     (SV[3] <> mvNoAnswer) or (SV[4] <> mvVerified) or (SV[5] <> mvWrongUnit) then
    BFails := BFails + ' sr-verdicts';
  if SDom[0] <> 'x.pas:11' then BFails := BFails + ' sr-dominant(' + SDom[0] + ')';
  var SSrc := 'begin A.Go; A.Go; end';
  var SMatches := TSemanticReplaceEngine.FindAllMatches(SSrc, [SRules[0]]);
  var SStats: TSemanticReplaceStats;
  if (Length(SMatches) <> 2) or (TSemanticReplaceEngine.ApplyToText(SSrc, [SRules[0]],
     [SMatches[1].Offset], SStats) <> 'begin B.Go; A.Go; end') or (SStats.Occurrences <> 1) then
    BFails := BFails + ' sr-skip';
  if (Length(SMatches) = 2) and (TSemanticReplaceEngine.VerifyOffset(SRules[0], SMatches[0]) <>
     SMatches[0].Offset + 2) then
    BFails := BFails + ' sr-verify-offset';
  // ---- batch quick fixes: order, relocation, a real batch
  var BFixes: TArray<TQuickFix>;
  SetLength(BFixes, 3);
  BFixes[0].Kind := qfAddUnit; BFixes[0].Line := 20;
  BFixes[1].Kind := qfRemoveVar; BFixes[1].Line := 5;
  BFixes[2].Kind := qfRemoveVar; BFixes[2].Line := 12;
  var BOrd := OrderFixesForBatch(BFixes);
  if (BOrd[0].Line <> 12) or (BOrd[1].Line <> 5) or (BOrd[2].Kind <> qfAddUnit) then
    BFails := BFails + ' batch-order';
  if RelocateFixLine(['a', 'b', 'x', 'c'], 'x', 1, 1) <> 2 then BFails := BFails + ' reloc-delta';
  if RelocateFixLine(['a', 'b', 'x', 'c'], 'x', 3, 0) <> 2 then BFails := BFails + ' reloc-near';
  if RelocateFixLine(['a', 'b', 'c'], 'x', 1, 0) <> -1 then BFails := BFails + ' reloc-gone';
  // (a real batch on a file runs in the console suite - it needs the
  // editor stand-in that writes files; this project has none)
  Assert.AreEqual('', BFails, BFails);
end;

procedure TChangeSignatureTests.MemberResolution_WithoutLsp;
// Delphi 13.1 answers NOTHING for a member whose class declares it as a
// private/public overload pair (RSS-5463) - the use site then has to be
// resolved from the sources: qualifier -> its declared type -> member.
const
  Src =
    'unit MuDemo;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TBase = class'#13#10 +
    '  public'#13#10 +
    '    procedure Shared;'#13#10 +
    '  end;'#13#10 +
    ''#13#10 +
    '  TMyClassA = class(TBase)'#13#10 +
    '  private'#13#10 +
    '    procedure Init; overload;'#13#10 +
    '  public'#13#10 +
    '    procedure Init(const ABoolean: Boolean); overload;'#13#10 +
    '  end;'#13#10 +
    ''#13#10 +
    '  TMyRec1 = record'#13#10 +
    '  public'#13#10 +
    '    procedure Init;'#13#10 +
    '  end;'#13#10 +
    ''#13#10 +
    'implementation'#13#10 +
    ''#13#10 +
    'procedure TMyClassB.Test;'#13#10 +
    'var'#13#10 +
    '  lMyClassA: TMyClassA;'#13#10 +
    'begin'#13#10 +
    '  lMyClassA := TMyClassA.Create;'#13#10 +
    '  lMyClassA.Init(True);'#13#10 +
    '  lMyClassA.Shared;'#13#10 +
    'end;'#13#10 +
    ''#13#10 +
    'end.';
var
  Link: TMemberLink;
  Ambiguous: Boolean;
begin
  var Lines := Src.Replace(#13#10, #10).Split([#10]);
  var UseInit := -1;
  var UseShared := -1;
  for var I := 0 to High(Lines) do
  begin
    if Pos('lMyClassA.Init', Lines[I]) > 0 then UseInit := I;
    if Pos('lMyClassA.Shared', Lines[I]) > 0 then UseShared := I;
  end;

  Assert.AreEqual('lMyClassA',
    QualifierBefore(Lines[UseInit], Pos('Init', Lines[UseInit]) - 1), 'qualifier');
  Assert.AreEqual('TMyClassA', DeclaredTypeOfIdentifier(Src, UseInit, 'lMyClassA'),
    'declared type of the qualifier');

  var FileName := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_member.pas');
  TFile.WriteAllText(FileName, Src, TEncoding.UTF8);
  var Graph := TTypeGraph.Create([FileName], nil, False);
  try
    // inherited member: found in the ANCESTOR, unambiguous
    Assert.IsTrue(Graph.FindMember('TMyClassA', 'Shared', Link, Ambiguous), 'Shared found');
    Assert.AreEqual('TBase', Link.TypeName, 'Shared declared in TBase');
    Assert.IsFalse(Ambiguous, 'Shared is not overloaded');
    // the overload pair: a position, but not THE declaration
    Assert.IsTrue(Graph.FindMember('TMyClassA', 'Init', Link, Ambiguous), 'Init found');
    Assert.IsTrue(Ambiguous, 'Init is overloaded');
    // a RECORD is a type with members too (the forum example of post #150
    // is records) - without them such a use site stays unresolvable
    Assert.IsTrue(Graph.FindMember('TMyRec1', 'Init', Link, Ambiguous), 'record member');
    Assert.AreEqual('TMyRec1', Link.TypeName, 'record type');
    Assert.IsFalse(Ambiguous, 'the record declares Init once');

    Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileName, Src, UseShared,
      Pos('Shared', Lines[UseShared]) - 1, 'Shared', Link)), 'use site resolved');
    // an overloaded call is decided by its ARGUMENT COUNT: "Init(True)"
    // can only reach the declaration that takes one parameter
    Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileName, Src, UseInit,
      Pos('Init', Lines[UseInit]) - 1, 'Init', Link)), 'overloaded use site');
    Assert.IsTrue(Link.Text.Contains('ABoolean'), 'the one-parameter overload');

    // the valuable answer: NOT our symbol, so the scans can drop it
    Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileName, Src, UseShared,
      Pos('Shared', Lines[UseShared]) - 1, 'Shared',
      function(AFile: string; ALine: Integer): Boolean begin Result := False; end,
      function(ATypeName: string): Boolean begin Result := False; end, Link)), 'other symbol');
    Assert.AreEqual(Ord(uuOurs), Ord(ClassifyUnansweredUse(Graph, FileName, Src, UseShared,
      Pos('Shared', Lines[UseShared]) - 1, 'Shared',
      function(AFile: string; ALine: Integer): Boolean begin Result := ALine = 5; end,
      function(ATypeName: string): Boolean begin Result := True; end, Link)), 'our symbol');
  finally
    Graph.Free;
    TFile.Delete(FileName);
  end;
end;

procedure TChangeSignatureTests.MemberResolution_PrivatePublicOverloadAcrossUnits;
// The reported constellation (forum #156, Delphi 13.1 bug RSS-5463): the
// class declares Init as a private/public OVERLOAD PAIR in one unit and is
// used from ANOTHER one. DelphiLSP then answers nothing at all there - no
// definition, no completion - so everything below has to hold without it.
const
  UnitA =
    'unit MuClassA;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TMyClassA = class(TObject)'#13#10 +
    '  private'#13#10 +
    '    procedure Init; overload;'#13#10 +
    '  public'#13#10 +
    '    procedure Init(const ABoolean: Boolean); overload;'#13#10 +
    '    procedure Free2;'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'end.';
  UnitB =
    'unit MuClassB;'#13#10 +
    'interface'#13#10 +
    'uses MuClassA;'#13#10 +
    'type'#13#10 +
    '  TMyClassB = class'#13#10 +
    '  public'#13#10 +
    '    procedure Test;'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    ''#13#10 +
    'procedure TMyClassB.Test;'#13#10 +
    'var'#13#10 +
    '  lMyClassA: TMyClassA;'#13#10 +
    'begin'#13#10 +
    '  lMyClassA := TMyClassA.Create;'#13#10 +
    '  try'#13#10 +
    '    lMyClassA.Init(True);'#13#10 +
    '    lMyClassA.Free2;'#13#10 +
    '  finally'#13#10 +
    '    lMyClassA.Free;'#13#10 +
    '  end;'#13#10 +
    'end;'#13#10 +
    ''#13#10 +
    'end.';
var
  Link: TMemberLink;
  Ambiguous: Boolean;
begin
  var Dir := TPath.Combine(TPath.GetTempPath, 'rl_member_units');
  TDirectory.CreateDirectory(Dir);
  var FileA := TPath.Combine(Dir, 'MuClassA.pas');
  var FileB := TPath.Combine(Dir, 'MuClassB.pas');
  TFile.WriteAllText(FileA, UnitA, TEncoding.UTF8);
  TFile.WriteAllText(FileB, UnitB, TEncoding.UTF8);
  try
    var Lines := UnitB.Replace(#13#10, #10).Split([#10]);
    var UseInit := -1;
    var UseFree2 := -1;
    for var I := 0 to High(Lines) do
    begin
      if Pos('lMyClassA.Init', Lines[I]) > 0 then UseInit := I;
      if Pos('lMyClassA.Free2', Lines[I]) > 0 then UseFree2 := I;
    end;
    // the variable is declared in the OTHER unit's type - the type name is
    // all the use site itself gives us
    Assert.AreEqual('TMyClassA', DeclaredTypeOfIdentifier(UnitB, UseInit, 'lMyClassA'),
      'type of the qualifier');

    var Graph := TTypeGraph.Create([FileA, FileB], nil, False);
    try
      // the overload pair: found, but not pinnable to ONE declaration
      Assert.IsTrue(Graph.FindMember('TMyClassA', 'Init', Link, Ambiguous), 'Init found');
      Assert.IsTrue(Ambiguous, 'the private/public pair is overloaded');
      Assert.AreEqual(FileA.ToUpper, Link.FilePath.ToUpper, 'declared in the other unit');
      // one argument, one declaration that takes one - decidable
      Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileB, UnitB, UseInit,
        Pos('Init', Lines[UseInit]) - 1, 'Init', Link)), 'use site resolved by argument count');
      Assert.IsTrue(Link.Text.Contains('ABoolean'), 'the public overload');

      // resolved, but to a declaration that is not the searched one:
      // another symbol, not an unverified row the user has to judge
      Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileB, UnitB, UseInit,
        Pos('Init', Lines[UseInit]) - 1, 'Init',
        function(AFile: string; ALine: Integer): Boolean begin Result := False; end,
        function(ATypeName: string): Boolean
        begin Result := SameText(ATypeName, 'TMyClassA'); end, Link)), 'not the searched overload');
      // ... and when it IS the searched declaration, it counts
      Assert.AreEqual(Ord(uuOurs), Ord(ClassifyUnansweredUse(Graph, FileB, UnitB, UseInit,
        Pos('Init', Lines[UseInit]) - 1, 'Init',
        function(AFile: string; ALine: Integer): Boolean
        begin Result := ALine = Link.Line; end,
        function(ATypeName: string): Boolean
        begin Result := SameText(ATypeName, 'TMyClassA'); end, Link)), 'the searched overload');
      Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileB, UnitB, UseInit,
        Pos('Init', Lines[UseInit]) - 1, 'Init',
        function(AFile: string; ALine: Integer): Boolean begin Result := False; end,
        function(ATypeName: string): Boolean begin Result := False; end, Link)),
        'Init of a class we are not renaming');

      // a NON-overloaded member of the same class resolves exactly, so the
      // scans can use it (this is what rename needs to stay complete)
      Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileB, UnitB, UseFree2,
        Pos('Free2', Lines[UseFree2]) - 1, 'Free2', Link)), 'Free2 resolved');
      Assert.AreEqual(8, Link.Line, 'declaration line of Free2 in MuClassA');
    finally
      Graph.Free;
    end;
  finally
    TFile.Delete(FileA);
    TFile.Delete(FileB);
  end;
end;

procedure TChangeSignatureTests.MemberResolution_UnqualifiedOverloadByArgumentType;
// The tester's record (forum #156, 2026-09-20): nested classes in a strict
// private section, two overloads with the SAME number of parameters, and an
// UNQUALIFIED call inside one of them. Neither the argument count nor a
// qualifier helps - only the TYPECAST in the first argument does.
const
  Src =
    'unit RecDemo;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    '  TMyRecord = record'#13#10 +
    '  strict private'#13#10 +
    '   type'#13#10 +
    '    TMyListA = class'#13#10 +
    '    end;'#13#10 +
    '    TMyListB = class'#13#10 +
    '    end;'#13#10 +
    '  strict private'#13#10 +
    '   class procedure Init(AListe: TMyListA; ASpur: Integer); overload; static;'#13#10 +
    '  public'#13#10 +
    '   class procedure Init(AListe: TMyListB; ASpur: Integer); overload; static;'#13#10 +
    '  end;'#13#10 +
    ''#13#10 +
    'implementation'#13#10 +
    ''#13#10 +
    'class procedure TMyRecord.Init(AListe: TMyListB; ASpur: Integer);'#13#10 +
    'begin'#13#10 +
    '  Init(TMyListA(AListe), ASpur);'#13#10 +
    'end;'#13#10 +
    ''#13#10 +
    'class procedure TMyRecord.Init(AListe: TMyListA; ASpur: Integer);'#13#10 +
    'begin'#13#10 +
    'end;'#13#10 +
    ''#13#10 +
    'end.';
var
  Link: TMemberLink;
begin
  var Lines := Src.Replace(#13#10, #10).Split([#10]);
  var CallLine := -1;
  var DeclA := -1;
  var DeclB := -1;
  for var I := 0 to High(Lines) do
  begin
    if Pos('Init(TMyListA(AListe)', Lines[I]) > 0 then CallLine := I;
    if Pos('Init(AListe: TMyListA', Lines[I]) > 0 then
      if DeclA < 0 then DeclA := I;
    if Pos('Init(AListe: TMyListB', Lines[I]) > 0 then
      if DeclB < 0 then DeclB := I;
  end;

  var FileName := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_record.pas');
  TFile.WriteAllText(FileName, Src, TEncoding.UTF8);
  var Graph := TTypeGraph.Create([FileName], nil, False);
  try
    // both overloads are declarations of the same name
    // Integer(): on Win64 Length answers NativeInt, and AreEqual cannot
    // infer its generic argument from two different integer types
    Assert.AreEqual(2, Integer(Length(Graph.FindMembers('TMyRecord', 'Init'))), 'two overloads');
    // the unqualified call belongs to the enclosing type, and the typecast
    // says which overload it reaches
    Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileName, Src,
      CallLine, Pos('Init', Lines[CallLine]) - 1, 'Init', Link)), 'call resolved');
    Assert.AreEqual(DeclA, Link.Line, 'the TMyListA overload');
    // searching the OTHER overload must not list that call
    Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileName, Src,
      CallLine, Pos('Init', Lines[CallLine]) - 1, 'Init',
      function(AFile: string; ALine: Integer): Boolean begin Result := ALine = DeclB; end,
      function(ATypeName: string): Boolean
      begin Result := SameText(ATypeName, 'TMyRecord'); end, Link)), 'not the B overload');
  finally
    Graph.Free;
    TFile.Delete(FileName);
  end;
end;

procedure TChangeSignatureTests.MemberResolution_SelfQualifiedWrappedCall;
// Forum 2026-09-20: searching TTestRecord.Init (another unit) listed the
// call "Self.Init( nil, ..." inside TMyRecord.Test as UNVERIFIED. Three
// things had to come together: "Self" is no declared identifier, the CALL
// wraps over six lines and so does the DECLARATION it reaches - so neither
// the type nor the argument count could be determined.
const
  SrcB =
    'unit UnitB;'#13#10 +
    'interface'#13#10 +
    'type'#13#10 +
    ' TMyRecord = record'#13#10 +
    ' private'#13#10 +
    '  type'#13#10 +
    '   TMyInnerRecord = record'#13#10 +
    '   public'#13#10 +
    '    procedure Init( AInteger1,'#13#10 +
    '                    AInteger2,'#13#10 +
    '                    AInteger3: Integer;'#13#10 +
    '                    AString: String);'#13#10 +
    '   end;'#13#10 +
    ' strict private'#13#10 +
    '  procedure Init; overload;'#13#10 +
    '  procedure Init( AObject: TObject;'#13#10 +
    '                  ABoolean1,'#13#10 +
    '                  ABoolean2: Boolean;'#13#10 +
    '                  AInteger1,'#13#10 +
    '                  AInteger2: Integer;'#13#10 +
    '                  AString: String); overload;'#13#10 +
    ' public'#13#10 +
    '   procedure Test;'#13#10 +
    ' end;'#13#10 +
    'implementation'#13#10 +
    'procedure TMyRecord.Test;'#13#10 +
    'begin'#13#10 +
    '  Self.Init( nil,'#13#10 +
    '             True,'#13#10 +
    '             True,'#13#10 +
    '             1,'#13#10 +
    '             2,'#13#10 +
    '             '''');'#13#10 +
    'end;'#13#10 +
    'end.';
var
  Link: TMemberLink;
begin
  var Lines := SrcB.Replace(#13#10, #10).Split([#10]);
  var CallLine := -1;
  var DeclSix := -1;
  var DeclNone := -1;
  for var I := 0 to High(Lines) do
  begin
    if Pos('Self.Init(', Lines[I]) > 0 then CallLine := I;
    if Pos('procedure Init( AObject', Lines[I]) > 0 then DeclSix := I;
    if Pos('procedure Init; overload', Lines[I]) > 0 then DeclNone := I;
  end;
  var FileB := TPath.Combine(TPath.GetTempPath, 'rl_dunitx_self_b.pas');
  TFile.WriteAllText(FileB, SrcB, TEncoding.UTF8);
  var Graph := TTypeGraph.Create([FileB], nil, False);
  try
    var Col := Pos('Init(', Lines[CallLine]) - 1;
    // Self. resolves to the enclosing type, and the six arguments pick the
    // six-parameter overload although both wrap over several lines
    Assert.AreEqual(Ord(murResolved), Ord(ResolveMemberUse(Graph, FileB, SrcB,
      CallLine, Col, 'Init', Link)), 'Self.Init resolved');
    Assert.AreEqual(DeclSix, Link.Line, 'the six-parameter overload');
    // the parameterless overload is not reached by that call
    Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileB, SrcB,
      CallLine, Col, 'Init',
      function(AFile: string; ALine: Integer): Boolean begin Result := ALine = DeclNone; end,
      function(ATypeName: string): Boolean
      begin Result := SameText(ATypeName, 'TMyRecord'); end, Link)), 'not the empty overload');
    // and the same-named member of a record in ANOTHER unit is not it either
    Assert.AreEqual(Ord(uuOtherSymbol), Ord(ClassifyUnansweredUse(Graph, FileB, SrcB,
      CallLine, Col, 'Init',
      function(AFile: string; ALine: Integer): Boolean begin Result := False; end,
      function(ATypeName: string): Boolean
      begin Result := SameText(ATypeName, 'TTestRecord'); end, Link)), 'not TTestRecord.Init');
    // the wrapped join keeps the first line's columns and runs to the end
    var Joined := JoinOpenParenLines(Lines, CallLine, 40);
    Assert.IsTrue(Joined.StartsWith('  Self.Init( nil,'), 'first line kept');
    Assert.IsTrue(Joined.EndsWith(');'), 'continued to the closing bracket');
  finally
    Graph.Free;
    TFile.Delete(FileB);
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TChangeSignatureTests);

end.
