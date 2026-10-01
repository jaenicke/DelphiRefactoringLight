(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 * Test cases contributed by Ian Branch (code audit, issue #22).
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Audit repro tests for Remove with (issue #22). Each test names the audit
///  item it reproduces. RED means the item is still open; a test can be
///  deleted when it goes green.
///
///  TResolvingLspClient answers GotoDefinition from a name table over a small
///  declarations unit written to a temporary folder, so the rewriter's whole
///  resolve chain (target, type, class range, members) runs headless against
///  real files.
/// </summary>
unit Test.AuditReproRemoveWith;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TAuditReproRemoveWithTests = class
  private
    FDir: string;
  public
    [Setup] procedure SetUp;
    [TearDown] procedure TearDown;
    /// <summary>H11: the temp for a record-typed target is a COPY, so the
    ///  write through the with is lost.</summary>
    [Test] procedure H11_RecordFieldTarget_TempIsNotACopy;
    /// <summary>H13: the classic temp of a with inside an anonymous method
    ///  is declared in the OUTER method and shared by every call.</summary>
    [Test] procedure H13_WithInAnonymousMethod_ClassicTempIsNotInOuterMethod;
    /// <summary>M40a: '@Field' in a with body is not qualified.</summary>
    [Test] procedure M40a_AddressOfMember_IsQualified;
    /// <summary>M40a: '&amp;Name' in a with body is not qualified.</summary>
    [Test] procedure M40a_EscapedIdentifier_IsQualified;
    /// <summary>M40b: a 'var' parameter is taken for the method's var
    ///  section.</summary>
    [Test] procedure M40b_VarParameter_IsNotAVarSection;
    /// <summary>M40c: the temp name is not checked against the method's
    ///  own locals.</summary>
    [Test] procedure M40c_TempName_AvoidsExistingLocal;
  end;

implementation

uses
  System.SysUtils, System.Character, System.IOUtils, System.Generics.Collections,
  Lsp.Client, Lsp.Protocol, Lsp.Uri,
  Expert.WithScanner, Expert.WithRewriter;

const
  NL = #13#10;

type
  /// <summary>A TLspClient whose GotoDefinition answers from a name table:
  ///  the identifier under the requested position is looked up by name.</summary>
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
  inherited Create('');  // no exe path - Start is never called
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
  StartIdx, EndIdx: Integer;
  Loc: TLspLocation;
begin
  Result := nil;
  if not TFile.Exists(AFilePath) then Exit;
  Lines := TFile.ReadAllLines(AFilePath);
  if (ALine < 0) or (ALine > High(Lines)) then Exit;
  S := Lines[ALine];
  StartIdx := ACol + 1;
  if (StartIdx < 1) or (StartIdx > Length(S)) then Exit;
  if not (S[StartIdx].IsLetterOrDigit or (S[StartIdx] = '_')) then Exit;
  while (StartIdx > 1) and (S[StartIdx - 1].IsLetterOrDigit or (S[StartIdx - 1] = '_')) do
    Dec(StartIdx);
  EndIdx := ACol + 1;
  while (EndIdx < Length(S)) and (S[EndIdx + 1].IsLetterOrDigit or (S[EndIdx + 1] = '_')) do
    Inc(EndIdx);
  if FNames.TryGetValue(UpperCase(Copy(S, StartIdx, EndIdx - StartIdx + 1)), Loc) then
    Result := [Loc];
end;

procedure TResolvingLspClient.RefreshDocument(const AFilePath: string);
begin
  // deliberately nothing
end;

{ Fixture helpers }

/// <summary>The declarations unit every fixture resolves against; the
///  trailing comments give each 0-based line.</summary>
function TypesUnitLines: TArray<string>;
begin
  Result := [
    'unit Types1;',            // 0
    '',                        // 1
    'interface',               // 2
    '',                        // 3
    'type',                    // 4
    '  TFoo = class',          // 5
    '  public',                // 6
    '    F: Boolean;',         // 7
    '    X: Integer;',         // 8
    '    Y: Integer;',         // 9
    '    &Type: Integer;',     // 10
    '    procedure Free;',     // 11
    '  end;',                  // 12
    '',                        // 13
    '  TRec = record',         // 14
    '    Count: Integer;',     // 15
    '  end;',                  // 16
    '',                        // 17
    '  PRec = ^TRec;',         // 18
    '',                        // 19
    '  TOuter = record',       // 20
    '    Inner: TRec;',        // 21
    '  end;',                  // 22
    '',                        // 23
    'var',                     // 24
    '  B: TFoo;',              // 25
    '  B2: TFoo;',             // 26
    '  R: TOuter;',            // 27
    '',                        // 28
    'function MakeFoo: TFoo;', // 29
    'function GetRecPtr: PRec;', // 30
    '',                        // 31
    'implementation',          // 32
    '',                        // 33
    'end.'];                   // 34
end;

/// <summary>The resolving client over Types1.pas in ADir; the caller frees it.</summary>
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
  MapAt('Type', 10);
  MapAt('Free', 11);
  MapAt('TRec', 14);
  MapAt('Count', 15);
  MapAt('PRec', 18);
  MapAt('TOuter', 20);
  MapAt('Inner', 21);
  MapAt('B', 25);
  MapAt('B2', 26);
  MapAt('R', 27);
  MapAt('MakeFoo', 29);
  MapAt('GetRecPtr', 30);
end;

/// <summary>Unit1 with a single routine: AHeader + ADecls + begin + ABody + end.
///  The routine's header is line 10 and, without ADecls, its begin is line 11.</summary>
function UnitSource(const AHeader: string; const ADecls, ABody: array of string): string;
begin
  Result := 'unit Unit1;' + NL + NL + 'interface' + NL + NL + 'uses' + NL + '  Types1;' + NL + NL +
    'implementation' + NL + NL + AHeader + NL;
  for var S in ADecls do
    Result := Result + S + NL;
  Result := Result + 'begin' + NL;
  for var S in ABody do
    Result := Result + S + NL;
  Result := Result + 'end;' + NL + NL + 'end.' + NL;
end;

/// <summary>Writes ASource as Unit1.pas and rewrites its first with-statement
///  (in source order) with the default settings.</summary>
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

/// <summary>Rewrites a one-line body in 'procedure Go(C: Boolean; N: Integer;
///  P: Pointer)' with a local S.</summary>
function RewriteLine(const ADir, ALine: string): TWithRewriteResult;
begin
  Result := RewriteFirst(ADir, UnitSource('procedure Go(C: Boolean; N: Integer; P: Pointer);',
    ['var', '  S: string;'], [ALine]));
end;

{ TAuditReproRemoveWithTests }

procedure TAuditReproRemoveWithTests.SetUp;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'RLAuditReproRemoveWith_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
  TFile.WriteAllLines(TPath.Combine(FDir, 'Types1.pas'), TypesUnitLines);
end;

procedure TAuditReproRemoveWithTests.TearDown;
begin
  if TDirectory.Exists(FDir) then
    TDirectory.Delete(FDir, True);
end;

procedure TAuditReproRemoveWithTests.H11_RecordFieldTarget_TempIsNotACopy;
var
  R: TWithRewriteResult;
begin
  // 'with R.Inner do Count := 0;' writes into the global R.
  // 'var LRec := R.Inner; LRec.Count := 0;' writes into a copy and leaves R unchanged.
  R := RewriteLine(FDir, '  with R.Inner do Count := 0;');
  Assert.IsFalse(R.NewText.Contains(':= R.Inner;'),
    'H11: the temp holds a COPY of the record, so the write is lost: ' + R.NewText);
end;

procedure TAuditReproRemoveWithTests.H13_WithInAnonymousMethod_ClassicTempIsNotInOuterMethod;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, UnitSource('procedure Go(C: Boolean);', [], [
    '  Run(procedure',
    '    begin',
    '      with MakeFoo do X := 1;',
    '    end);']));
  // The outer routine's begin is line 11. A classic temp declared there is
  // captured by the anonymous method and shared by every invocation.
  Assert.IsFalse(R.Classic.Supported and (R.Classic.MethodBodyBeginLine = 11),
    'H13: the classic temp goes into the OUTER method''s var section (method begin line ' +
    IntToStr(R.Classic.MethodBodyBeginLine) + ')');
end;

procedure TAuditReproRemoveWithTests.M40a_AddressOfMember_IsQualified;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do P := @X;');
  Assert.AreEqual('P := @B.X;', R.NewText, False,
    'M40a: @Field is taken for member access and left unqualified');
end;

procedure TAuditReproRemoveWithTests.M40a_EscapedIdentifier_IsQualified;
var
  R: TWithRewriteResult;
begin
  R := RewriteLine(FDir, '  with B do &Type := 1;');
  Assert.AreEqual('B.&Type := 1;', R.NewText, False,
    'M40a: &Name is taken for member access and left unqualified');
end;

procedure TAuditReproRemoveWithTests.M40b_VarParameter_IsNotAVarSection;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, UnitSource('procedure Go(var A: Integer);', [], [
    '  with MakeFoo do X := A;']));
  Assert.IsTrue(R.Classic.Supported, 'precondition: the classic form is derivable here');
  Assert.IsFalse(R.Classic.HasVarSection,
    'M40b: the var PARAMETER is taken for a var section (VarSectionLastLine ' +
    IntToStr(R.Classic.VarSectionLastLine) + '), so the temp is inserted without a var keyword');
end;

procedure TAuditReproRemoveWithTests.M40c_TempName_AvoidsExistingLocal;
var
  R: TWithRewriteResult;
begin
  R := RewriteFirst(FDir, UnitSource('procedure Go(C: Boolean);', ['var', '  LFoo: Integer;'], [
    '  with MakeFoo do X := 1;',
    '  LFoo := 2;']));
  Assert.IsTrue(R.Classic.Supported, 'precondition: the classic form is derivable here');
  Assert.IsTrue(Length(R.Classic.VarDecls) > 0, 'precondition: the rewrite declares a temp');
  for var D in R.Classic.VarDecls do
    Assert.IsFalse(SameText(D.Name, 'LFoo'),
      'M40c: the temp reuses the name of the existing local LFoo (classic: duplicate declaration ''' +
      D.Name + ': ' + D.TypeName + ''')');
end;

initialization
  TDUnitX.RegisterTestFixture(TAuditReproRemoveWithTests);

end.
