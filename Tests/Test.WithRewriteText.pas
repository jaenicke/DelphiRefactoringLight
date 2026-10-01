(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Pins the TEXT the remove-with rewriter produces, for shapes that came out
///  compiling but bound differently or ran differently.
///
///  Test.WithRewriter proves the rewriter REFUSES when it does not know. This
///  unit proves that, when it does know, what it writes is right. It needs a
///  language server that actually resolves: TResolvingLspClient answers
///  GotoDefinition from a name table over a small declarations unit written to
///  a temporary folder, so the whole resolve chain (target, type, class range,
///  members) runs headless against real files on disk.
/// </summary>
unit Test.WithRewriteText;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TWithRewriteTextTests = class
  private
    FDir: string;
  public
    [Setup] procedure SetUp;
    [TearDown] procedure TearDown;

    /// <summary>A begin..end inside a single-statement body must not end the
    ///  body at its first ';', or Y goes out bare and binds elsewhere.</summary>
    [Test] procedure SingleBody_IfThenBeginEnd_QualifiesEveryStatement;
    /// <summary>The else of an if INSIDE the body belongs to the body.</summary>
    [Test] procedure SingleBody_IfThenElse_QualifiesTheElseBranch;
    /// <summary>The ';' between repeat and until is not the body's end.</summary>
    [Test] procedure SingleBody_RepeatUntil_QualifiesTheWholeLoop;
    /// <summary>A while loop's begin..end body belongs to the with body.</summary>
    [Test] procedure SingleBody_WhileBeginEnd_QualifiesTheWholeLoop;
    /// <summary>A with nested inside a single-statement body's begin..end is
    ///  found, so the nested-with guard can see it.</summary>
    [Test] procedure SingleBody_NestedWithInBeginEnd_IsFound;
    /// <summary>A with straight after a case label is found at top level.</summary>
    [Test] procedure CaseLabel_WithAtTopLevel_IsFound;
    /// <summary>A with after a case label inside another with's body is found,
    ///  so the outer rewrite cannot claim the inner one's members unseen.</summary>
    [Test] procedure CaseLabel_WithInsideWithBody_IsFound;
    /// <summary>A compound body with a temp in a single-statement slot (after
    ///  then) must be wrapped in begin..end, or the try runs unconditionally.</summary>
    [Test] procedure CompoundBodyAfterThen_IsWrappedInBeginEnd;
    /// <summary>The classic (non-inline) form of the same shape is wrapped too.</summary>
    [Test] procedure CompoundBodyAfterThen_ClassicIsWrappedInBeginEnd;
    /// <summary>Counter-case: in a statement list the declaration and the block
    ///  may stand side by side, so no extra begin..end is added there.</summary>
    [Test] procedure CompoundBodyInStatementList_IsNotWrapped;
    /// <summary>A rewrite that needs no temp is applicable in classic mode as-is.</summary>
    [Test] procedure NoTemp_ClassicIsSupportedWithTheSameText;
  end;

implementation

uses
  System.SysUtils, System.Character, System.IOUtils, System.Generics.Collections,
  Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Expert.WithScanner, Expert.WithRewriter;

const
  NL = #13#10;

type
  /// <summary>A TLspClient whose GotoDefinition answers from a name table: the
  ///  identifier under the requested position is looked up by name.</summary>
  TResolvingLspClient = class(TLspClient)
  private
    FNames: TDictionary<string, TLspLocation>;
  public
    constructor Create; reintroduce;
    destructor Destroy; override;
    /// <summary>Maps AName (case-insensitive) to a 0-based position in AFile.</summary>
    procedure Map(const AName, AFile: string; ALine0, ACol0: Integer);
    /// <summary>Answers from the name table; never contacts a server.</summary>
    function GotoDefinition(const AFilePath: string; ALine, ACol: Integer): TArray<TLspLocation>; override;
    /// <summary>No-op; the fake has no document state.</summary>
    procedure RefreshDocument(const AFilePath: string); override;
  end;

{ TResolvingLspClient }

constructor TResolvingLspClient.Create;
begin
  inherited Create('');
  FNames := TDictionary<string, TLspLocation>.Create;
end;

destructor TResolvingLspClient.Destroy;
begin
  FNames.Free;
  inherited;
end;

procedure TResolvingLspClient.Map(const AName, AFile: string; ALine0, ACol0: Integer);
var
  Loc: TLspLocation;
begin
  Loc := Default(TLspLocation);
  Loc.Uri := TLspUri.PathToFileUri(AFile);
  Loc.Range.Start.Line := ALine0;
  Loc.Range.Start.Character := ACol0;
  Loc.Range.End_ := Loc.Range.Start;
  FNames.AddOrSetValue(UpperCase(AName), Loc);
end;

function TResolvingLspClient.GotoDefinition(const AFilePath: string;
  ALine, ACol: Integer): TArray<TLspLocation>;
var
  Lines: TArray<string>;
  S: string;
  B, E: Integer;
  Loc: TLspLocation;
begin
  Result := nil;
  if not TFile.Exists(AFilePath) then Exit;
  Lines := TFile.ReadAllLines(AFilePath);
  if (ALine < 0) or (ALine > High(Lines)) then Exit;
  S := Lines[ALine];
  B := ACol + 1;
  if (B < 1) or (B > Length(S)) then Exit;
  if not (S[B].IsLetterOrDigit or (S[B] = '_')) then Exit;
  while (B > 1) and (S[B - 1].IsLetterOrDigit or (S[B - 1] = '_')) do Dec(B);
  E := ACol + 1;
  while (E < Length(S)) and (S[E + 1].IsLetterOrDigit or (S[E + 1] = '_')) do Inc(E);
  if FNames.TryGetValue(UpperCase(Copy(S, B, E - B + 1)), Loc) then
    Result := [Loc];
end;

procedure TResolvingLspClient.RefreshDocument(const AFilePath: string);
begin
  // deliberately nothing
end;

{ Fixture helpers }

/// <summary>The declarations unit every fixture resolves against.</summary>
function TypesUnitLines: TArray<string>;
begin
  Result := [
    'unit Types1;',                        // 0
    '',                                    // 1
    'interface',                           // 2
    '',                                    // 3
    'type',                                // 4
    '  TFoo = class',                      // 5
    '  public',                            // 6
    '    F: Boolean;',                     // 7
    '    X: Integer;',                     // 8
    '    Y: Integer;',                     // 9
    '    procedure Free;',                 // 10
    '  end;',                              // 11
    '',                                    // 12
    'var',                                 // 13
    '  B: TFoo;',                          // 14
    '  B2: TFoo;',                         // 15
    '',                                    // 16
    'function MakeFoo: TFoo;',             // 17
    '',                                    // 18
    'implementation',                      // 19
    '',                                    // 20
    'end.'];                               // 21
end;

/// <summary>Builds the resolving client over the declarations unit in ADir.</summary>
function MakeClient(const ADir: string): TResolvingLspClient;
var
  TypesFile: string;
  Lines: TArray<string>;

  procedure MapAt(const AName: string; ALine0: Integer);
  begin
    Result.Map(AName, TypesFile, ALine0, Pos(AName, Lines[ALine0]) - 1);
  end;

begin
  TypesFile := TPath.Combine(ADir, 'Types1.pas');
  Lines := TypesUnitLines;
  Result := TResolvingLspClient.Create;
  MapAt('TFoo', 5);
  MapAt('F', 7);
  MapAt('X', 8);
  MapAt('Y', 9);
  MapAt('Free', 10);
  MapAt('B', 14);
  MapAt('B2', 15);
  MapAt('MakeFoo', 17);
end;

/// <summary>A unit whose single routine is AHeader + begin + ABody + end.</summary>
function UnitSource(const AHeader: string; const ABody: array of string): string;
var
  S: string;
begin
  Result := 'unit Unit1;' + NL + NL + 'interface' + NL + NL + 'uses' + NL + '  Types1;' + NL + NL +
    'implementation' + NL + NL + AHeader + NL + 'begin' + NL;
  for S in ABody do
    Result := Result + S + NL;
  Result := Result + 'end;' + NL + NL + 'end.' + NL;
end;

/// <summary>Writes ASource as Unit1.pas and rewrites its first with-statement
///  (in source order).</summary>
function RewriteFirst(const ADir, ASource: string): TWithRewriteResult;
var
  Occs: TArray<TWithOccurrence>;
  Client: TResolvingLspClient;
  FileName: string;
  First: Integer;
begin
  FileName := TPath.Combine(ADir, 'Unit1.pas');
  TFile.WriteAllText(FileName, ASource);
  Occs := TWithScanner.ScanSource(ASource);
  Assert.IsTrue(Length(Occs) > 0, 'fixture should contain a with-statement');
  // ScanSource lists a nested with before its outer one.
  First := 0;
  for var I := 1 to High(Occs) do
    if (Occs[I].KeywordPos.Line < Occs[First].KeywordPos.Line)
      or ((Occs[I].KeywordPos.Line = Occs[First].KeywordPos.Line)
        and (Occs[I].KeywordPos.Col < Occs[First].KeywordPos.Col)) then
      First := I;
  Client := MakeClient(ADir);
  try
    Result := TWithRewriter.Rewrite(Client, FileName, ASource, Occs[First],
      TWithRewriteSettings.Defaults);
  finally
    Client.Free;
  end;
end;

/// <summary>Rewrites a one-line body in a plain routine.</summary>
function RewriteLine(const ADir, ALine: string): TWithRewriteResult;
begin
  Result := RewriteFirst(ADir, UnitSource('procedure Go;', [ALine]));
end;

/// <summary>The compound-body fixture: 'with MakeFoo do try .. finally .. end',
///  standing after 'if C then' (AAfterThen) or in the routine's statement list.</summary>
function CompoundSource(AAfterThen: Boolean): string;
begin
  if AAfterThen then
    Result := UnitSource('procedure Go(C: Boolean);', [
      '  if C then',
      '    with MakeFoo do',
      '    try',
      '      X := 1;',
      '    finally',
      '      Free;',
      '    end;'])
  else
    Result := UnitSource('procedure Go(C: Boolean);', [
      '  with MakeFoo do',
      '  try',
      '    X := 1;',
      '  finally',
      '    Free;',
      '  end;']);
end;

{ TWithRewriteTextTests }

procedure TWithRewriteTextTests.SetUp;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'RLWithRewriteText_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
  TFile.WriteAllLines(TPath.Combine(FDir, 'Types1.pas'), TypesUnitLines);
end;

procedure TWithRewriteTextTests.TearDown;
begin
  if TDirectory.Exists(FDir) then
    TDirectory.Delete(FDir, True);
end;

procedure TWithRewriteTextTests.SingleBody_IfThenBeginEnd_QualifiesEveryStatement;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do if F then begin X := 1; Y := 2; end;');
  Assert.IsTrue(R.IsAutoRewritable, 'a fully resolved with must be rewritable');
  Assert.AreEqual('if B.F then begin B.X := 1; B.Y := 2; end;', R.NewText, False);
end;

procedure TWithRewriteTextTests.SingleBody_IfThenElse_QualifiesTheElseBranch;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do if F then X := 1 else Y := 2;');
  Assert.AreEqual('if B.F then B.X := 1 else B.Y := 2;', R.NewText, False);
end;

procedure TWithRewriteTextTests.SingleBody_RepeatUntil_QualifiesTheWholeLoop;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do repeat X := 1; Y := 2 until F;');
  Assert.AreEqual('repeat B.X := 1; B.Y := 2 until B.F;', R.NewText, False);
end;

procedure TWithRewriteTextTests.SingleBody_WhileBeginEnd_QualifiesTheWholeLoop;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do while F do begin X := 1; Y := 2 end;');
  Assert.AreEqual('while B.F do begin B.X := 1; B.Y := 2 end;', R.NewText, False);
end;

procedure TWithRewriteTextTests.SingleBody_NestedWithInBeginEnd_IsFound;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(
    'begin' + NL + '  with B do if F then begin with B2 do Y := 1; X := 2 end;' + NL + 'end;' + NL);
  Assert.AreEqual<Integer>(2, Length(Occs), 'the inner with must be found as well');
end;

procedure TWithRewriteTextTests.CaseLabel_WithAtTopLevel_IsFound;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(
    'begin' + NL + '  case N of 1: with B do X := 1; end;' + NL + 'end;' + NL);
  Assert.AreEqual<Integer>(1, Length(Occs), 'a with after a case label is a with');
end;

procedure TWithRewriteTextTests.CaseLabel_WithInsideWithBody_IsFound;
var
  Occs: TArray<TWithOccurrence>;
begin
  Occs := TWithScanner.ScanSource(
    'begin' + NL + '  with B do case N of 1: with B2 do Y := 1; end;' + NL + 'end;' + NL);
  Assert.AreEqual<Integer>(2, Length(Occs), 'the with after the case label must be found');
end;

procedure TWithRewriteTextTests.CompoundBodyAfterThen_IsWrappedInBeginEnd;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, CompoundSource(True));
  Assert.AreEqual(
    'begin' + NL +
    '      var LFoo := MakeFoo;' + NL +
    '      try' + NL +
    '        LFoo.X := 1;' + NL +
    '      finally' + NL +
    '        LFoo.Free;' + NL +
    '      end' + NL +
    '    end', R.NewText, False);
end;

procedure TWithRewriteTextTests.CompoundBodyAfterThen_ClassicIsWrappedInBeginEnd;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, CompoundSource(True));
  Assert.IsTrue(R.Classic.Supported, 'the classic form is derivable here');
  Assert.AreEqual(
    'begin' + NL +
    '      LFoo := MakeFoo;' + NL +
    '      try' + NL +
    '        LFoo.X := 1;' + NL +
    '      finally' + NL +
    '        LFoo.Free;' + NL +
    '      end' + NL +
    '    end', R.Classic.BodyText, False);
end;

procedure TWithRewriteTextTests.CompoundBodyInStatementList_IsNotWrapped;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, CompoundSource(False));
  Assert.AreEqual(
    'var LFoo := MakeFoo;' + NL +
    '  try' + NL +
    '    LFoo.X := 1;' + NL +
    '  finally' + NL +
    '    LFoo.Free;' + NL +
    '  end', R.NewText, False);
end;

procedure TWithRewriteTextTests.NoTemp_ClassicIsSupportedWithTheSameText;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do X := 1;');
  Assert.IsTrue(R.Classic.Supported, 'nothing about this rewrite needs an inline variable');
  Assert.AreEqual(R.NewText, R.Classic.BodyText, False);
  Assert.AreEqual<Integer>(0, Length(R.Classic.VarDecls), 'and no declaration is added');
end;

initialization
  TDUnitX.RegisterTestFixture(TWithRewriteTextTests);

end.
