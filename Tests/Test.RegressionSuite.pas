(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
/// <summary>
///  Regression tests ported from the project's console suite (the one used
///  during development, which also covers IDE-near code) - the pure,
///  IDE-free cases that document fixed bugs: the issue #10 audit findings,
///  forum reports and tester rounds. Each test names what went wrong.
/// </summary>
unit Test.RegressionSuite;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TFileEncodingRegressionTests = class
  private
    FDir: string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;
    [Test] procedure AnsiTarget_UnrepresentableChar_IsUpgradedNotLost;
    [Test] procedure ReadLines_AllLineBreakKinds;
    [Test] procedure InvalidUtf8_IsReadAsAnsi;
  end;

  [TestFixture]
  TUsesClauseRegressionTests = class
  public
    [Test] procedure CommentedOutUnit_IsNotInTheClause;
    [Test] procedure UnitInImplementationClause_IsNotReachableFromInterface;
  end;

  [TestFixture]
  TQuickFixSafetyTests = class
  public
    [Test] procedure CanTakeSemicolon_RefusesContinuations;
    [Test] procedure CanTakeSemicolon_AcceptsStatements;
    [Test] procedure SoleBranchStatement_IsRecognised;
    [Test] procedure OrdinaryStatement_IsNoSoleBranch;
    [Test] procedure ConstructorStub_ChainsToAncestor;
    [Test] procedure H2219Name_FromGenericAndQualifiedMessages;
  end;

  [TestFixture]
  TWithScannerRegressionTests = class
  public
    [Test] procedure SiblingWithsInsideBlock_AreBothFound;
    [Test] procedure NestedThenSibling_AllThreeFound;
    [Test] procedure MultiLineStringContent_IsNotScanned;
  end;

  [TestFixture]
  TSourceScanRegressionTests = class
  public
    [Test] procedure CommentMask_CarriesBlockCommentsAcrossLines;
    [Test] procedure ForwardDeclaration_LosesAgainstRealOne;
    [Test] procedure EnclosingRoutine_WithAnonymousMethods;
  end;

  /// <summary>The version is shown by the IDE (splash, About box, status
  ///  window) from Expert.Version, by Windows from the .dproj version info,
  ///  and at the top of the README - all three must name the SAME version,
  ///  or nobody can tell whether the installed build is the current one.</summary>
  [TestFixture]
  TVersionTests = class
  public
    [Test] procedure ReadmeDprojAndPluginVersionAgree;
  end;

  [TestFixture]
  TMiscRegressionTests = class
  public
    [Test] procedure LspUri_EncodedDriveColon_IsLocal;
    [Test] procedure LspUri_OldFiveSlashUncForm;
    [Test] procedure LspUri_DrivePathWithUmlautRoundTrips;
    [Test] procedure BackupPath_MirrorsTheFullPath;
    [Test] procedure VcsOutput_MixedUtf8AndAnsiLines;
    [Test] procedure WorkerLatch_WaitsAndRefusesAfterShutdown;
  end;

  /// <summary>The MCP pipe is how Claude Code reaches this IDE. On a 64-bit
  ///  IDE it was never created: the token buffer of GetTokenInformation was
  ///  a byte array (alignment 1), and on an odd stack address the API
  ///  answers ERROR_NOACCESS (998). These tests run on BOTH platforms, so
  ///  the 64-bit build is the one that matters here.</summary>
  [TestFixture]
  TMcpPipeRegressionTests = class
  public
    [Test] procedure SecurityDescriptor_IsBuiltOnThisPlatform;
    [Test] procedure Start_ReportsThePipeItReallyCreated;
  end;

  /// <summary>Issue #13: while DelphiLSP loads a big project (12-30 s) its
  ///  controller aborts every request after 10 s. The only honest signal
  ///  that it is busy is its own "$/progress" notification - which we
  ///  ignored, so the first search after opening a project read the
  ///  failures as results ("1 of 341 candidate(s) verified").</summary>
  [TestFixture]
  TLspProgressTests = class
  public
    [Test] procedure ProgressNotification_BeginReportEnd;
    [Test] procedure ProgressNotification_RejectsWhatIsNotProgress;
  end;

  /// <summary>A registry value of the WRONG TYPE made TRegistry.ReadBool
  ///  raise out of TPluginSettings.Load, and the IDE then refused to load
  ///  the whole BPL ("Ungueltiger Datentyp fuer 'VerifySession'"). Settings
  ///  are a convenience - they may never decide whether the plugin loads.</summary>
  [TestFixture]
  TSettingsRobustnessTests = class
  public
    [Test] procedure WrongValueType_FallsBackToTheDefault;
    [Test] procedure RenameChoices_AreRememberedAndOutOfRangeIsRefused;
  end;

  /// <summary>DelphiLSP answers a "go to definition" for a static class
  ///  method with the IMPLEMENTATION header and a range starting at column
  ///  0. The partner query (implementation -> declaration) was asked right
  ///  there - on the keyword "class" - and answered nothing, so the
  ///  declaration was missing from the target set and find references
  ///  dropped every correctly resolved call (PsyPrax, 2026-09-21: 1 of 9
  ///  verified). The query must be asked at the NAME.</summary>
  [TestFixture]
  TPartnerQueryTests = class
  public
    [Test] procedure NameColumn_ImplementationHeaderPrefersTheMember;
    [Test] procedure NameColumn_DeclarationCommentAndMisses;
  end;

  /// <summary>A declaration whose implementation exists with ANOTHER
  ///  signature (the user added a parameter to the declaration, 2026-09-21)
  ///  got "create empty implementation" - a second body - and the class
  ///  header got "implement the interface method" - a second declaration.
  ///  Both only add errors. The fixes are to make declaration and
  ///  implementation agree, in either direction; an OVERLOADED declaration
  ///  keeps the stub (a new overload needs its own body).</summary>
  [TestFixture]
  TAlignSignatureFixTests = class
  public
    [Test] procedure ChangedDeclaration_OffersBothAlignDirections;
    [Test] procedure AlignDeclaration_KeepsDirectivesAndTrailingDefaults;
  end;

  /// <summary>The uses-graph SCC pass used to recurse once per unit along
  ///  the chain, so a deep enough graph overflowed the stack - and a stack
  ///  overflow inside a design-time BPL takes the IDE with it rather than
  ///  failing the analysis. Measured on the recursive version through this
  ///  same seam: 6,000 units passed, 7,000 raised EStackOverflow.</summary>
  [TestFixture]
  TUsesGraphDepthTests = class
  public
    [Test] procedure DeepCycle_DoesNotOverflowTheStack;
    [Test] procedure Components_AreStillCorrectOnASmallGraph;
    [Test] procedure DeepEnumerateCycles_DoesNotOverflowTheStack;
    [Test] procedure DeepEnumerateCyclesThrough_DoesNotOverflowTheStack;
    [Test] procedure DeepEdgeLevers_DoesNotOverflowTheStack;
  end;

  /// <summary>Uses cleanup called a unit UNUSED although a method of its
  ///  class helper was called ("Button1.Dummy" with TButtonHelper = class
  ///  helper for TButton in Unit1). The using unit never names the helper,
  ///  only the member - which the index did not know, because it skips
  ///  class bodies. Helper members are indexed with HelperMemberPrefix
  ///  now.</summary>
  [TestFixture]
  THelperUsageTests = class
  public
    [Test] procedure Parser_IndexesHelperMembersWithThePrefix;
    [Test] procedure Snapshot_KeepsHelperMembersOutOfTheNormalLookup;
    [Test] procedure Cleanup_CountsAMemberCallAsUseOfTheHelperUnit;
  end;

  /// <summary>Issue #16 (Ian Branch): the unit index cache wrote its
  ///  counts with SizeOf of an INFERRED type - 8 bytes on Win64 - and read
  ///  4, so the 64-bit IDE read four bytes of a unit name as an identifier
  ///  count and zero-filled a 12-16 GB array (SetLength succeeds on Win64,
  ///  so the "corrupt -> rebuild" handler never ran). A cache file must
  ///  round-trip, and a malformed one must be REJECTED without allocating
  ///  more than it holds.</summary>
  [TestFixture]
  TUnitIndexCacheTests = class
  public
    [Test] procedure RoundTrip_KeepsUnitsAndIdentifiers;
    [Test] procedure Win64ShapedCounts_AreRejected;
    [Test] procedure AbsurdIdentifierCount_IsRejectedWithoutAllocating;
    [Test] procedure AbsurdStringLength_IsRejected;
    [Test] procedure TruncatedAndTrailing_AreRejected;
  end;

  /// <summary>Forum: "Move to unit" moved only the FIRST line of a
  ///  routine declaration whose parameter list wraps - the declaration was
  ///  cut at the first ';', which sits inside the parameter list.</summary>
  [TestFixture]
  TMoveDeclarationRangeTests = class
  public
    [Test] procedure WrappedParameterList_IsMovedWhole;
    [Test] procedure ConstantsAndTypesWithInnerSemicolons;
  end;

  /// <summary>Two rename reports (forum 2026-09-22): "TestX" renamed back
  ///  to "Test" showed the right preview but changed nothing, and a field
  ///  named "ABC" could not be renamed at all once the unit used
  ///  Winapi.Windows ("declared in the RAD Studio installation").</summary>
  [TestFixture]
  TRenameGuardTests = class
  public
    [Test] procedure EditGuard_MatchesWholeWordsOnly;
    [Test] procedure ForeignAnswerAtADeclaration_IsRecognised;
  end;

implementation

uses
  System.SysUtils, System.IOUtils, System.Classes, System.SyncObjs,
  Delphi.FileEncoding, Expert.UsesEditor, Expert.AutoImport, Expert.UnitIndex,
  Expert.WithScanner, Lsp.Uri, Rename.WorkspaceEdit, Expert.VcsBlame,
  Expert.WorkerLatch, Expert.Version, Expert.PascalScanner, System.RegularExpressions,
  Winapi.Windows, Mcp.PipeServer, System.JSON, Lsp.Protocol,
  System.Win.Registry, Expert.PluginSettings, Expert.UsesGraph,
  Expert.UsesCleanup, Expert.MoveToUnit, Expert.SafeDeletePlan;

const
  NL = sLineBreak;

{ TFileEncodingRegressionTests }

procedure TFileEncodingRegressionTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath,
    'RefLightRegression_' + FormatDateTime('yyyymmddhhnnsszzz', Now));
  TDirectory.CreateDirectory(FDir);
end;

procedure TFileEncodingRegressionTests.TearDown;
begin
  if (FDir <> '') and TDirectory.Exists(FDir) then
    try
      TDirectory.Delete(FDir, True);
    except
    end;
end;

procedure TFileEncodingRegressionTests.AnsiTarget_UnrepresentableChar_IsUpgradedNotLost;
var
  F, Content: string;
begin
  // A file detected as ANSI receives a character outside its code page
  // (a rename to a CJK name, an arrow in a string). Writing through the
  // ANSI encoding would substitute '?' - irrecoverably.
  F := TPath.Combine(FDir, 'ansi.pas');
  Content := 'X := ''' + #$2192 + ' ' + #$4E2D + ''';';
  TDelphiFileEncoding.WriteAll(F, Content, TEncoding.Default);
  Assert.AreEqual(Content, TDelphiFileEncoding.ReadAll(F));
end;

procedure TFileEncodingRegressionTests.ReadLines_AllLineBreakKinds;
var
  F: string;
  L: TArray<string>;
begin
  F := TPath.Combine(FDir, 'lines.pas');
  TFile.WriteAllText(F, 'a'#13#10'b'#10'c'#13'd'#13#10, TEncoding.UTF8);
  L := TDelphiFileEncoding.ReadLines(F);
  Assert.AreEqual<Integer>(4, Length(L), 'CRLF, LF and CR each end a line; no empty last line');
  Assert.AreEqual('a', L[0]);
  Assert.AreEqual('b', L[1]);
  Assert.AreEqual('c', L[2]);
  Assert.AreEqual('d', L[3]);
end;

procedure TFileEncodingRegressionTests.InvalidUtf8_IsReadAsAnsi;
var
  F: string;
begin
  F := TPath.Combine(FDir, 'cp1252.pas');
  TFile.WriteAllBytes(F, TBytes.Create($47, $72, $E4, $73, $73, $65));   // "Gr?sse" in CP1252
  Assert.AreEqual('Gr' + #$E4 + 'sse', TDelphiFileEncoding.ReadAll(F));
end;

{ TUsesClauseRegressionTests }

procedure TUsesClauseRegressionTests.CommentedOutUnit_IsNotInTheClause;
begin
  // issue #10 (A8): a unit inside a { } comment was taken for a listed one,
  // so a legitimate "add unit" was refused
  Assert.IsFalse(UnitInUsesText('unit U;' + NL + 'uses A {, B};' + NL, 'B'));
  Assert.IsTrue(UnitInUsesText('unit U;' + NL + 'uses A {, B}, B;' + NL, 'B'));
  Assert.IsFalse(UnitInUsesText('unit U;' + NL + 'uses' + NL + '  A,' + NL +
    '  { B; removed }' + NL + '  C;' + NL, 'B'));
end;

procedure TUsesClauseRegressionTests.UnitInImplementationClause_IsNotReachableFromInterface;
var
  Src: string;
begin
  Src := 'unit U;' + NL + 'interface' + NL + 'uses A;' + NL +
    'implementation' + NL + 'uses X, Target, Y;' + NL + 'end.';
  Assert.IsFalse(UnitInUsesSection(Src, 'Target', usInterface));
  Assert.IsTrue(UnitInUsesSection(Src, 'Target', usImplementation));
  Assert.IsTrue(UnitInUsesSection(Src, 'A', usInterface));
end;

{ TQuickFixSafetyTests }

procedure TQuickFixSafetyTests.CanTakeSemicolon_RefusesContinuations;
begin
  for var S in TArray<string>.Create(
    '  with FList', '  case K', '  Result := Foo(A,', '  X := A shl',
    '  if A and', '  I := I mod', '  strict private', '  implementation') do
    Assert.IsFalse(CanTakeSemicolon(S), '"' + S + '" cannot take a semicolon');
end;

procedure TQuickFixSafetyTests.CanTakeSemicolon_AcceptsStatements;
begin
  for var S in TArray<string>.Create(
    '  S := ''abc''', '  if A then Foo', '  for I := 0 to 9 do Bar(I)',
    '  while A do Step', '  Foo(1, 2)', '  X := 1 // note', '  raise') do
    Assert.IsTrue(CanTakeSemicolon(S), '"' + S + '" is a statement');
end;

procedure TQuickFixSafetyTests.SoleBranchStatement_IsRecognised;
begin
  // issue #10 (B9): deleting the dead assignment after "if ... then" made
  // the FOLLOWING statement conditional
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('if not FCached then', '  FRetries := 0;'), 1));
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('  if A then // note', '    X := 1;'), 1));
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('  if A then { multi', '  line }', '    X := 1;'), 2));
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('case K of', '  1:', '    X := 1;'), 2));
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('  else', '    X := 1;'), 1));
  Assert.IsTrue(IsSoleBranchStatement(TArray<string>.Create('while A do', 'X := 1;'), 1));
end;

procedure TQuickFixSafetyTests.OrdinaryStatement_IsNoSoleBranch;
begin
  Assert.IsFalse(IsSoleBranchStatement(TArray<string>.Create('begin', '  X := 1;'), 1));
  Assert.IsFalse(IsSoleBranchStatement(TArray<string>.Create('  Foo;', '', '  X := 1;'), 2));
  Assert.IsFalse(IsSoleBranchStatement(TArray<string>.Create('  X := 1;'), 0));
end;

procedure TQuickFixSafetyTests.ConstructorStub_ChainsToAncestor;
begin
  // issue #10 (B8): an empty constructor body compiles and fails far away
  Assert.AreEqual('  inherited;', StubBodyLine('constructor', False,
    ' constructor Create(AOwner: TComponent); override;', '  TBar = class(TComponent)'));
  Assert.AreEqual('  inherited Create;', StubBodyLine('constructor', False,
    ' constructor Create(A: Integer);', '  TFoo = class'));
  Assert.AreEqual('  inherited Create;', StubBodyLine('constructor', False,
    ' constructor Create(A: Integer);', '  TFoo = class(TObject)'));
  Assert.StartsWith('  // TODO', StubBodyLine('constructor', False,
    ' constructor CreateX(A: Integer);', '  TBar = class(TComponent, IFoo)'));
  Assert.AreEqual('  inherited;', StubBodyLine('destructor', False,
    ' destructor Destroy; override;', '  TFoo = class'));
  Assert.AreEqual('', StubBodyLine('constructor', False,
    ' constructor Create(A: Integer);', '  TRec = record'), 'records have no inheritance');
  Assert.AreEqual('', StubBodyLine('constructor', True,
    ' class constructor Create;', '  TFoo = class'), 'class constructors chain nothing');
  Assert.AreEqual('', StubBodyLine('procedure', False, ' procedure Foo;', '  TFoo = class'));
end;

procedure TQuickFixSafetyTests.H2219Name_FromGenericAndQualifiedMessages;
begin
  // tester: "2 fixable, declined empty, 1 fix" - the provider bailed out
  // silently on the quoted generic / qualified names
  Assert.AreEqual('Test', H2219NameFromMessage('H2219 Private symbol ''Test<T>'' declared but never used'));
  Assert.AreEqual('Test', H2219NameFromMessage('H2219 Private symbol ''TForm4.Test'' declared but never used'));
  Assert.AreEqual('FCount', H2219NameFromMessage('H2219 Privates Symbol ''FCount'' deklariert, aber nie verwendet'));
end;

{ TWithScannerRegressionTests }

procedure TWithScannerRegressionTests.SiblingWithsInsideBlock_AreBothFound;
begin
  Assert.AreEqual<Integer>(2, Length(TWithScanner.ScanSource('begin' + NL +
    '  with FOne do Baz := 1;' + NL + '  with FTwo do Quux := 2;' + NL + 'end;')));
end;

procedure TWithScannerRegressionTests.NestedThenSibling_AllThreeFound;
begin
  Assert.AreEqual<Integer>(3, Length(TWithScanner.ScanSource('begin' + NL +
    '  with A do' + NL + '    with B do X := 1;' + NL + '  with C do Y := 2;' + NL + 'end;')));
end;

procedure TWithScannerRegressionTests.MultiLineStringContent_IsNotScanned;
begin
  // Delphi 12+ ''' literals: their CONTENT was scanned as code
  Assert.AreEqual<Integer>(1, Length(TWithScanner.ScanSource(
    'S := ''''''' + NL + '  with X do Y;' + NL + '  '''''';' + NL + 'with A do B := 1;')));
end;

{ TSourceScanRegressionTests }

procedure TSourceScanRegressionTests.CommentMask_CarriesBlockCommentsAcrossLines;
var
  L, M: TArray<string>;

  function IsCode(ALine: Integer; const AWord: string; ANth: Integer): Boolean;
  var
    P: Integer;
  begin
    P := 0;
    for var K := 1 to ANth do
      P := Pos(AWord, L[ALine], P + 1);
    Result := (P > 0) and (M[ALine][P] = L[ALine][P]);
  end;

begin
  // forum report: find references / rename listed words inside the lines
  // of a MULTI-LINE { } comment
  L := TArray<string>.Create(
    '{',                              // 0
    '  Hier wird oefter mal init aufgerufen',   // 1
    '}',                              // 2
    '{ init } Init;',                 // 3
    '(* init',                        // 4
    '   init *) Init; // init',       // 5
    'S := ''{ init'' + Init;');       // 6
  M := MaskCommentsAndStrings(L);
  Assert.IsFalse(IsCode(1, 'init', 1), 'line inside the comment');
  Assert.IsFalse(IsCode(3, 'init', 1));
  Assert.IsTrue(IsCode(3, 'Init', 1));
  Assert.IsFalse(IsCode(5, 'init', 1));
  Assert.IsTrue(IsCode(5, 'Init', 1));
  Assert.IsFalse(IsCode(5, 'init', 2), '// comment');
  Assert.IsFalse(IsCode(6, 'init', 1), 'inside a string');
  Assert.IsTrue(IsCode(6, 'Init', 1), 'after the string');
  for var K := 0 to High(L) do
    Assert.AreEqual<Integer>(Length(L[K]), Length(M[K]), 'line lengths are kept');
end;

procedure TSourceScanRegressionTests.ForwardDeclaration_LosesAgainstRealOne;
begin
  // DevExpress declares nearly every class forward first - "find original
  // symbol" landed on the bodyless line
  Assert.AreEqual<Integer>(5, FindDeclarationLine(
    'unit cxButtons;' + NL + 'interface' + NL + 'type' + NL +
    '  TcxButton = class;' + NL + '' + NL +
    '  TcxButton = class(TcxCustomButton)' + NL + '  end;' + NL +
    'implementation' + NL + 'end.', 'TcxButton'));
  Assert.AreEqual<Integer>(3, FindDeclarationLine(
    'unit U;' + NL + 'interface' + NL + 'type' + NL +
    '  TOnlyFwd = class;' + NL + 'implementation' + NL + 'end.', 'TOnlyFwd'),
    'with nothing else the forward line is still the answer');
end;

procedure TSourceScanRegressionTests.EnclosingRoutine_WithAnonymousMethods;
var
  Src: string;
  F, L: Integer;
begin
  // tester: no handler generated once the method contained an anonymous
  // method - "end);" was never counted as the end of a block
  Src :=
    'implementation' + NL +                                       // 0
    'procedure TForm1.FormCreate(Sender: TObject);' + NL +        // 1
    'begin' + NL +                                                // 2
    '  OnShow := ' + NL +                                         // 3
    '  TThread.Queue(nil,' + NL +                                 // 4
    '    procedure' + NL +                                        // 5
    '    begin' + NL +                                            // 6
    '      X := 1;' + NL +                                        // 7
    '    end);' + NL +                                            // 8
    '  Foo(procedure(Sender: TObject)' + NL +                     // 9
    '    begin' + NL +                                            // 10
    '    end, 1);' + NL +                                         // 11
    '  OnClick := ' + NL +                                        // 12
    'end;' + NL +                                                 // 13
    'end.';
  for var Caret in TArray<Integer>.Create(3, 7, 12) do
  begin
    Assert.IsTrue(FindEnclosingRoutineRange(Src, Caret, F, L), 'caret ' + IntToStr(Caret));
    Assert.AreEqual<Integer>(1, F);
    Assert.AreEqual<Integer>(13, L);
  end;
end;

{ TVersionTests }

procedure TVersionTests.ReadmeDprojAndPluginVersionAgree;
var
  Root, Readme, Dproj: string;
  M: TMatch;
  Parts: TArray<string>;
begin
  // the repository root: walk up from the test exe (Tests\<platform>\<config>)
  Root := ExtractFilePath(ParamStr(0));
  while (Root <> '') and not TFile.Exists(TPath.Combine(Root, 'README.md')) do
  begin
    var Up := ExtractFilePath(ExcludeTrailingPathDelimiter(Root));
    if Up = Root then Root := '' else Root := Up;
  end;
  Assert.IsTrue(Root <> '', 'repository root (README.md) not found above the test exe');
  Readme := TFile.ReadAllText(TPath.Combine(Root, 'README.md'));
  Dproj := TFile.ReadAllText(TPath.Combine(Root, 'Packages\DelphiRefactoringLight.dproj'));

  M := TRegEx.Match(Readme, '\*\*Version ([0-9.]+)\*\*');
  Assert.IsTrue(M.Success, 'README.md has no "**Version x.y.z**" line');
  Assert.AreEqual(PluginVersion, M.Groups[1].Value, 'README version vs Expert.Version');

  Parts := PluginVersion.Split(['.']);
  while Length(Parts) < 3 do Parts := Parts + ['0'];
  Assert.IsTrue(Dproj.Contains('<VerInfo_MajorVer>' + Parts[0] + '</VerInfo_MajorVer>'), 'dproj MajorVer');
  Assert.IsTrue(Dproj.Contains('<VerInfo_MinorVer>' + Parts[1] + '</VerInfo_MinorVer>'), 'dproj MinorVer');
  Assert.IsTrue(Dproj.Contains('<VerInfo_Release>' + Parts[2] + '</VerInfo_Release>'), 'dproj Release');
  Assert.IsTrue(Dproj.Contains('FileVersion=' + Parts[0] + '.' + Parts[1] + '.' + Parts[2] + '.0;'),
    'dproj FileVersion key');
  Assert.IsTrue(Dproj.Contains('ProductVersion=' + PluginVersion + ';'), 'dproj ProductVersion key');
end;

{ TMiscRegressionTests }

procedure TMiscRegressionTests.LspUri_EncodedDriveColon_IsLocal;
begin
  // DelphiLSP's own form: the drive colon percent-encoded. A drive test
  // before decoding turned every local path into "\\c:\..."
  Assert.AreEqual('c:\Temp\X.pas', TLspUri.FileUriToPath('file:///c%3A/Temp/X.pas'));
  Assert.AreEqual('C:\A B\Größe.pas', TLspUri.FileUriToPath('file:///C%3A/A%20B/Gr%C3%B6%C3%9Fe.pas'));
  // no drive and no share: a (relative) file name, never "\\Unit.pas"
  Assert.AreEqual('Unit.pas', TLspUri.FileUriToPath('file:///Unit.pas'));
  Assert.AreEqual('\\srv\share\x.pas', TLspUri.FileUriToPath('file:////srv/share/x.pas'));
end;

procedure TMiscRegressionTests.LspUri_OldFiveSlashUncForm;
begin
  // what this code base itself used to emit for a UNC path
  Assert.AreEqual('\\server\share\X.pas', TLspUri.FileUriToPath('file://///server/share/X.pas'));
end;

procedure TMiscRegressionTests.LspUri_DrivePathWithUmlautRoundTrips;
const
  P = 'C:\A B\Größe.pas';
begin
  Assert.AreEqual(P, TLspUri.FileUriToPath(TLspUri.PathToFileUri(P)));
end;

procedure TMiscRegressionTests.BackupPath_MirrorsTheFullPath;
begin
  // issue #10 (A10): Client\Utils.pas and Server\Utils.pas shared one backup
  Assert.AreEqual('D:\bk\C\src\Client\Utils.pas', MirroredBackupPath('D:\bk', 'C:\src\Client\Utils.pas'));
  Assert.AreEqual('D:\bk\C\src\Server\Utils.pas', MirroredBackupPath('D:\bk', 'C:\src\Server\Utils.pas'));
  Assert.AreEqual('D:\bk\UNC\srv\share\x.pas', MirroredBackupPath('D:\bk', '\\srv\share\x.pas'));
end;

procedure TMiscRegressionTests.VcsOutput_MixedUtf8AndAnsiLines;
var
  All: TBytes;
begin
  // a diff reproduces the BYTES of the files - a cp1252 source inside UTF-8
  // git output raised EEncodingError
  All := TEncoding.UTF8.GetBytes('Author: J' + #$00E4 + 'nicke' + #10) +
    TEncoding.ANSI.GetBytes('+  Gr' + #$00F6 + #$00DF + 'e := 1;' + #10) +
    TEncoding.UTF8.GetBytes('@@ ' + #$00FC + ' @@');
  Assert.AreEqual('Author: J' + #$00E4 + 'nicke' + #10 + '+  Gr' + #$00F6 + #$00DF +
    'e := 1;' + #10 + '@@ ' + #$00FC + ' @@', DecodeVcsOutput(All, Length(All)));
end;

procedure TMiscRegressionTests.WorkerLatch_WaitsAndRefusesAfterShutdown;
var
  Done: Integer;
begin
  // issue #10 (A7): workers must be drained before the package unloads
  ResetWorkerLatch;
  try
    Done := 0;
    Assert.IsTrue(StartWorker(
      procedure
      begin
        Sleep(200);
        TInterlocked.Increment(Done);
      end));
    Assert.IsTrue(ShutdownWorkersAndWait(5000), 'the running worker finishes');
    Assert.AreEqual<Integer>(1, Done);
    Assert.IsFalse(StartWorker(procedure begin end), 'no new worker after the shutdown began');
    Assert.IsTrue(WorkersShuttingDown);
  finally
    ResetWorkerLatch;
  end;
end;

{ TMcpPipeRegressionTests }

procedure TMcpPipeRegressionTests.SecurityDescriptor_IsBuiltOnThisPlatform;
var
  Err: DWORD;
begin
  // The step the 64-bit IDE failed at. 998 = ERROR_NOACCESS, which for
  // GetTokenInformation means "your buffer is not aligned".
  Assert.IsTrue(McpPipeSecurityProbe(Err),
    Format('the pipe security descriptor must be buildable, error %d: %s',
      [Err, SysErrorMessage(Err)]));
end;

procedure TMcpPipeRegressionTests.Start_ReportsThePipeItReallyCreated;
var
  Srv: TMcpPipeServer;
  Name: string;
begin
  // Start used to return True whatever the listener did, so an IDE without
  // a pipe looked like a running server.
  Name := '\\.\pipe\RefLightTest-' + UIntToStr(GetCurrentProcessId) + '-' +
    FormatDateTime('hhnnsszzz', Now);
  Srv := TMcpPipeServer.Create(Name,
    function(const ARequest: string; AStop: THandle): string
    begin
      Result := '{"ok":true}';
    end);
  try
    Assert.IsTrue(Srv.Start, 'Start reports the created pipe: ' + Srv.LastError);
    Assert.IsTrue(Srv.Listening, 'an instance is waiting');
    Assert.IsTrue(WaitNamedPipe(PChar(Name), 1000), 'the pipe is reachable');
    Srv.Stop;
    Assert.IsFalse(Srv.Listening, 'after Stop nothing listens');
  finally
    Srv.Free;
  end;
end;

{ TLspProgressTests }

procedure TLspProgressTests.ProgressNotification_BeginReportEnd;

  function Parse(const AJson: string; out AToken, AKind, AText: string): Boolean;
  begin
    var O := TJSONObject.ParseJSONValue(AJson) as TJSONObject;
    try
      Result := ParseProgressNotification(O, AToken, AKind, AText);
    finally
      O.Free;
    end;
  end;

var
  Token, Kind, Text: string;
begin
  // what DelphiLSP sends while it loads a project
  Assert.IsTrue(Parse('{"token":"1","value":{"kind":"begin","title":"Loading project",' +
    '"message":"Rezepturen.dpr"}}', Token, Kind, Text));
  Assert.AreEqual('1', Token);
  Assert.AreEqual('begin', Kind);
  Assert.AreEqual('Loading project Rezepturen.dpr', Text);
  // a NUMERIC token identifies the same task
  Assert.IsTrue(Parse('{"token":7,"value":{"kind":"report","message":"units"}}',
    Token, Kind, Text));
  Assert.AreEqual('7', Token);
  Assert.AreEqual('report', Kind);
  Assert.IsTrue(Parse('{"token":7,"value":{"kind":"end"}}', Token, Kind, Text));
  Assert.AreEqual('end', Kind);
  Assert.AreEqual('', Text, 'an end carries no text');
end;

procedure TLspProgressTests.ProgressNotification_RejectsWhatIsNotProgress;
var
  Token, Kind, Text: string;

  function Parse(const AJson: string): Boolean;
  begin
    var O := TJSONObject.ParseJSONValue(AJson) as TJSONObject;
    try
      Result := ParseProgressNotification(O, Token, Kind, Text);
    finally
      O.Free;
    end;
  end;

begin
  Assert.IsFalse(Parse('{"value":{"kind":"begin"}}'), 'no token');
  Assert.IsFalse(Parse('{"token":"1"}'), 'no value');
  Assert.IsFalse(Parse('{"token":"1","value":{"kind":"whatever"}}'), 'unknown kind');
  Assert.IsFalse(ParseProgressNotification(nil, Token, Kind, Text), 'nil params');
end;

{ TSettingsRobustnessTests }

procedure TSettingsRobustnessTests.WrongValueType_FallsBackToTheDefault;
var
  Reg: TRegistry;
  Key: string;
begin
  // write a STRING where the plugin writes a DWORD - exactly what broke
  // the package load
  Key := TPluginSettings.RegistryKey;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Assert.IsTrue(Reg.OpenKey(Key, True), 'settings key');
    var Had := Reg.ValueExists('VerifySession');
    var Old := 0;
    if Had and (Reg.GetDataType('VerifySession') = rdInteger) then
      Old := Reg.ReadInteger('VerifySession');
    try
      Reg.WriteString('VerifySession', 'certainly not a boolean');
      Reg.CloseKey;
      // Load must survive it and keep the default
      try
        TPluginSettings.Load;
      except
        on E: Exception do
          Assert.Fail('Load raised on a bad value: ' + E.ClassName + ': ' + E.Message);
      end;
      Assert.IsTrue(TPluginSettings.VerifySession, 'the default survives');
    finally
      // put the user's own value back
      if Reg.OpenKey(Key, True) then
      begin
        Reg.DeleteValue('VerifySession');
        if Had then Reg.WriteInteger('VerifySession', Old);
        Reg.CloseKey;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TSettingsRobustnessTests.RenameChoices_AreRememberedAndOutOfRangeIsRefused;
var
  Reg: TRegistry;
  Key: string;
  HadScope, HadBackup: Boolean;
  OldScope: Integer;
  OldBackup: Boolean;
begin
  // Issue: several renames in one unit meant switching the scope and
  // unticking "Create backup" every time. The dialog now stores both;
  // this pins the round trip and the range check (a combo index from a
  // future version with more scopes must not select nothing).
  Key := TPluginSettings.RegistryKey;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Assert.IsTrue(Reg.OpenKey(Key, True), 'settings key');
    HadScope := Reg.ValueExists('RenameScope') and (Reg.GetDataType('RenameScope') = rdInteger);
    HadBackup := Reg.ValueExists('RenameBackup') and (Reg.GetDataType('RenameBackup') = rdInteger);
    OldScope := 0; OldBackup := True;
    if HadScope then OldScope := Reg.ReadInteger('RenameScope');
    if HadBackup then OldBackup := Reg.ReadBool('RenameBackup');
    Reg.CloseKey;
    try
      TPluginSettings.Load;
      TPluginSettings.RenameScope := 1;       // "In current unit"
      TPluginSettings.RenameBackup := False;
      TPluginSettings.Save;
      TPluginSettings.RenameScope := 0;
      TPluginSettings.RenameBackup := True;
      TPluginSettings.Load;
      Assert.AreEqual(1, TPluginSettings.RenameScope, 'scope survives a reload');
      Assert.IsFalse(TPluginSettings.RenameBackup, 'backup switch survives a reload');

      Assert.IsTrue(Reg.OpenKey(Key, True));
      Reg.WriteInteger('RenameScope', 7);
      Reg.CloseKey;
      TPluginSettings.Load;
      Assert.AreEqual(0, TPluginSettings.RenameScope, 'out of range -> whole project');
    finally
      if Reg.OpenKey(Key, True) then
      begin
        Reg.DeleteValue('RenameScope');
        Reg.DeleteValue('RenameBackup');
        if HadScope then Reg.WriteInteger('RenameScope', OldScope);
        if HadBackup then Reg.WriteBool('RenameBackup', OldBackup);
        Reg.CloseKey;
      end;
      TPluginSettings.Load;
    end;
  finally
    Reg.Free;
  end;
end;

{ THelperUsageTests }

const
  HelperUnit1 =
    'unit Unit1;'#13#10 +
    'interface'#13#10 +
    'uses'#13#10 +
    '   Vcl.StdCtrls;'#13#10 +
    'type'#13#10 +
    '  TButtonHelper = class helper for TButton'#13#10 +
    '  public'#13#10 +
    '    procedure Dummy;'#13#10 +
    '    class function Make(const A: string): TButton; static;'#13#10 +
    '    property Hint2: string read GetHint2;'#13#10 +
    '  private type'#13#10 +
    '    TInner = class'#13#10 +
    '      procedure NotAHelperMember;'#13#10 +
    '    end;'#13#10 +
    '  end;'#13#10 +
    '  TPlain = class'#13#10 +
    '    procedure PlainMember;'#13#10 +
    '  end;'#13#10 +
    '  TIntHelper = record helper for Integer'#13#10 +
    '    function Twice: Integer;'#13#10 +
    '  end;'#13#10 +
    'procedure AfterAll;'#13#10 +
    'implementation'#13#10 +
    'procedure TButtonHelper.Dummy;'#13#10 +
    'begin'#13#10 +
    'end;'#13#10 +
    'end.';

procedure THelperUsageTests.Parser_IndexesHelperMembersWithThePrefix;
var
  F, UnitName: string;
  HasInit: Boolean;
  Ids: TArray<string>;
begin
  F := TPath.Combine(TPath.GetTempPath, 'RlHelperUnit1.pas');
  TFile.WriteAllText(F, HelperUnit1);
  try
    Ids := ParseUnit(F, UnitName, HasInit);
  finally
    TFile.Delete(F);
  end;
  var Joined := '|' + string.Join('|', Ids) + '|';
  Assert.Contains(Joined, '|TButtonHelper|');
  Assert.Contains(Joined, '|.Dummy|', 'helper method');
  Assert.Contains(Joined, '|.Make|', 'class function of the helper');
  Assert.Contains(Joined, '|.Hint2|', 'helper property');
  Assert.Contains(Joined, '|.Twice|', 'record helper for a simple type');
  Assert.DoesNotContain(Joined, '|Dummy|', 'never as a top-level identifier');
  Assert.DoesNotContain(Joined, 'NotAHelperMember', 'member of a NESTED type');
  Assert.DoesNotContain(Joined, 'PlainMember', 'members of ordinary classes stay out');
  Assert.Contains(Joined, '|TPlain|');
  Assert.Contains(Joined, '|AfterAll|', 'the parser is back on track after the helper');
end;

procedure THelperUsageTests.Snapshot_KeepsHelperMembersOutOfTheNormalLookup;
var
  Src: TUnitSource;
  Snap: IUnitSnapshot;
begin
  Src.UnitName := 'Unit1';
  Src.Path := 'C:\x\Unit1.pas';
  Src.Idents := TArray<string>.Create('TButtonHelper', HelperMemberPrefix + 'Dummy');
  Src.HasInit := False;
  Snap := BuildUnitSnapshot([Src]);
  Assert.AreEqual<Integer>(0, Length(Snap.Lookup('Dummy')), 'no add-unit for a bare Dummy');
  Assert.AreEqual<Integer>(1, Length(Snap.Lookup(HelperMemberPrefix + 'dummy')));
  Assert.AreEqual<Integer>(0, Length(Snap.Search('Dumm', 50)), 'not in the Find Unit search');
  Assert.AreEqual<Integer>(0, Length(Snap.FuzzyIdentifiers('Dummx', 2, 10)), 'no "did you mean"');
  // and the layered view the IDE actually uses
  Snap := ComposeUnitSnapshots(nil, Snap);
  Assert.AreEqual<Integer>(1, Length(Snap.Lookup(HelperMemberPrefix + 'Dummy')));
  Assert.AreEqual<Integer>(0, Length(Snap.Search('Dumm', 50)));
end;

procedure THelperUsageTests.Cleanup_CountsAMemberCallAsUseOfTheHelperUnit;
const
  Unit2 =
    'unit Unit2;'#13#10 +
    'interface'#13#10 +
    'uses'#13#10 +
    '  Vcl.StdCtrls, Unit1, UnitOther;'#13#10 +
    'procedure Run(Button: TButton);'#13#10 +
    'implementation'#13#10 +
    'procedure Run(Button: TButton);'#13#10 +
    'begin'#13#10 +
    '  Button.Dummy;'#13#10 +
    '  Dummy2;'#13#10 +
    'end;'#13#10 +
    'end.';
var
  Infos: TArray<TUsesEntryInfo>;
begin
  Infos := AnalyzeUses(Unit2,
    function(const AIdent: string): TArray<string>
    begin
      Result := nil;
      if SameText(AIdent, 'TButton') then Result := ['Vcl.StdCtrls']
      else if SameText(AIdent, HelperMemberPrefix + 'Dummy') then Result := ['Unit1']
      else if SameText(AIdent, HelperMemberPrefix + 'Dummy2') then Result := ['UnitOther'];
    end,
    function(const AUnit: string): Boolean begin Result := True; end,
    function(const AUnit: string): Boolean begin Result := False; end);
  var Found := 0;
  for var E in Infos do
  begin
    if SameText(E.UnitName, 'Unit1') then
    begin
      Inc(Found);
      Assert.IsTrue(E.Verdict <> uvUnused, 'Unit1 is used through Button.Dummy');
    end;
    if SameText(E.UnitName, 'UnitOther') then
    begin
      Inc(Found);
      // an UNQUALIFIED Dummy2 is no member access - no helper involved
      Assert.IsTrue(E.Verdict = uvUnused, 'a bare call does not count as a helper use');
    end;
  end;
  Assert.AreEqual(2, Found);
end;

{ TUnitIndexCacheTests }

function CacheTempFile(const AName: string): string;
begin
  Result := TPath.Combine(TPath.GetTempPath, 'RlIdxTest_' + AName + '.idx');
end;

function SampleUnits: TArray<TUnitSource>;
begin
  SetLength(Result, 2);
  Result[0].UnitName := 'Unit1';
  Result[0].Path := 'C:\x\Unit1.pas';
  Result[0].Idents := TArray<string>.Create('TFoo', 'Bar', '.Dummy', 'TList<');
  Result[0].HasInit := True;
  Result[1].UnitName := 'Ümlaut.Unit';
  Result[1].Path := 'C:\x\Ümlaut.Unit.pas';
  Result[1].Idents := nil;
  Result[1].HasInit := False;
end;

// The magic block (length + text) of the CURRENT format, taken from a file
// the production writer produced - the test must not know the constant.
function MagicPrefix: TBytes;
var
  F: string;
  B: TBytes;
  L: Integer;
begin
  F := CacheTempFile('magic');
  WriteUnitIndexCacheFile(F, nil);
  try
    B := TFile.ReadAllBytes(F);
  finally
    TFile.Delete(F);
  end;
  Move(B[0], L, 4);
  Result := Copy(B, 0, 4 + L);
end;

procedure AppendInt(var B: TBytes; V: Integer);
begin
  var P := Length(B);
  SetLength(B, P + 4);
  Move(V, B[P], 4);
end;

procedure AppendInt64(var B: TBytes; V: Int64);
begin
  var P := Length(B);
  SetLength(B, P + 8);
  Move(V, B[P], 8);
end;

procedure AppendStr(var B: TBytes; const S: string);
begin
  var U := TEncoding.UTF8.GetBytes(S);
  AppendInt(B, Length(U));
  var P := Length(B);
  SetLength(B, P + Length(U));
  if Length(U) > 0 then Move(U[0], B[P], Length(U));
end;

// One entry's fixed part in the CURRENT layout, up to (not including) the
// include count.
procedure AppendEntryHead(var B: TBytes; const APath, AUnit: string);
begin
  AppendInt(B, $494E4455);        // entry sentinel
  AppendStr(B, APath);
  AppendStr(B, AUnit);
  AppendInt64(B, 0);              // MTime
  AppendInt64(B, 0);              // Size
  var P := Length(B);
  SetLength(B, P + 1); B[P] := 0; // HasInit
  AppendInt64(B, 0);              // IncStamp
end;

procedure TUnitIndexCacheTests.RoundTrip_KeepsUnitsAndIdentifiers;
var
  F: string;
  R: TArray<TUnitSource>;
begin
  F := CacheTempFile('roundtrip');
  WriteUnitIndexCacheFile(F, SampleUnits);
  try
    Assert.IsFalse(TFile.Exists(F + '.tmp'), 'the temporary file is renamed into place');
    R := ReadUnitIndexCacheFile(F);
  finally
    TFile.Delete(F);
  end;
  Assert.AreEqual<Integer>(2, Length(R));
  for var S in R do
    if SameText(S.UnitName, 'Unit1') then
    begin
      Assert.AreEqual('TFoo|Bar|.Dummy|TList<', string.Join('|', S.Idents));
      Assert.IsTrue(S.HasInit);
    end
    else
    begin
      Assert.AreEqual('Ümlaut.Unit', S.UnitName, 'UTF-8 names survive');
      Assert.AreEqual<Integer>(0, Length(S.Idents));
    end;
end;

procedure TUnitIndexCacheTests.Win64ShapedCounts_AreRejected;
var
  B: TBytes;
  F: string;
begin
  // exactly what the old writer produced on Win64: every count 8 bytes wide
  B := MagicPrefix;
  AppendInt64(B, 1);                              // unit count, 8 bytes
  AppendEntryHead(B, 'C:\x\A.pas', 'A');
  AppendInt64(B, 0);                              // include count, 8 bytes
  AppendInt64(B, 1);                              // identifier count, 8 bytes
  AppendStr(B, 'TFoo');
  AppendInt(B, $444E4549);
  F := CacheTempFile('win64');
  TFile.WriteAllBytes(F, B);
  try
    Assert.AreEqual<Integer>(0, Length(ReadUnitIndexCacheFile(F)),
      'out of step from the first entry - rebuild, never half-read');
  finally
    TFile.Delete(F);
  end;
end;

procedure TUnitIndexCacheTests.AbsurdIdentifierCount_IsRejectedWithoutAllocating;
var
  B: TBytes;
  F: string;
begin
  // the value measured in the report: four bytes of ASCII as a count
  B := MagicPrefix;
  AppendInt(B, 1);
  AppendEntryHead(B, 'C:\x\A.pas', 'A');
  AppendInt(B, 0);                                // no includes
  AppendInt(B, 1635069299);                       // identifier count
  AppendStr(B, 'TFoo');
  F := CacheTempFile('absurd');
  TFile.WriteAllBytes(F, B);
  try
    var T0 := GetTickCount64;
    var R := ReadUnitIndexCacheFile(F);
    // before the fix this line allocated ~13 GB on Win64 (and raised
    // EOutOfMemory on Win32); now the count is checked against the bytes
    // that are left before anything is allocated
    Assert.AreEqual<Integer>(0, Length(R));
    Assert.IsTrue(GetTickCount64 - T0 < 2000, 'rejected at once');
  finally
    TFile.Delete(F);
  end;
end;

procedure TUnitIndexCacheTests.AbsurdStringLength_IsRejected;
var
  B: TBytes;
  F: string;
begin
  B := MagicPrefix;
  AppendInt(B, 1);
  AppendInt(B, $494E4455);
  AppendInt(B, 2000000000);                       // path "length"
  F := CacheTempFile('strlen');
  TFile.WriteAllBytes(F, B);
  try
    Assert.AreEqual<Integer>(0, Length(ReadUnitIndexCacheFile(F)));
  finally
    TFile.Delete(F);
  end;
end;

procedure TUnitIndexCacheTests.TruncatedAndTrailing_AreRejected;
var
  F: string;
  B: TBytes;
begin
  F := CacheTempFile('cut');
  WriteUnitIndexCacheFile(F, SampleUnits);
  try
    B := TFile.ReadAllBytes(F);
    // an IDE killed while writing - the old in-place writer left exactly this
    TFile.WriteAllBytes(F, Copy(B, 0, Length(B) - 7));
    Assert.AreEqual<Integer>(0, Length(ReadUnitIndexCacheFile(F)), 'truncated');
    // bytes behind the end sentinel mean the reader is out of step
    TFile.WriteAllBytes(F, B + [1, 2, 3, 4]);
    Assert.AreEqual<Integer>(0, Length(ReadUnitIndexCacheFile(F)), 'trailing bytes');
    // and the untouched file still loads
    TFile.WriteAllBytes(F, B);
    Assert.AreEqual<Integer>(2, Length(ReadUnitIndexCacheFile(F)), 'intact');
  finally
    TFile.Delete(F);
  end;
end;

{ TMoveDeclarationRangeTests }

procedure TMoveDeclarationRangeTests.WrappedParameterList_IsMovedWhole;
const
  // the forum example verbatim
  Src =
    'unit UnitA;'#13#10 +                        // 1
    ''#13#10 +                                   // 2
    'interface'#13#10 +                          // 3
    ''#13#10 +                                   // 4
    ' procedure Test( AParam1: Integer;'#13#10 + // 5
    '                 AParam2: Integer);'#13#10 +// 6
    ''#13#10 +                                   // 7
    'implementation'#13#10 +                     // 8
    ''#13#10 +
    'procedure Test( AParam1,'#13#10 +
    '                AParam2: Integer);'#13#10 +
    'begin'#13#10 +
    ''#13#10 +
    'end;'#13#10 +
    'end.';
var
  S, E: Integer;
begin
  Assert.IsTrue(LocateMoveDeclaration('Test', Src, S, E), 'found: Test');
  Assert.AreEqual(5, S);
  Assert.AreEqual(6, E, 'the second parameter line belongs to the declaration');
  // a directive on its own line after a wrapped header still goes along
  Assert.IsTrue(LocateMoveDeclaration('Run',
    'unit U;'#13#10'interface'#13#10 +
    'procedure Run(A: Integer;'#13#10 +          // 3
    '  B: string);'#13#10 +                      // 4
    '  overload;'#13#10 +                        // 5
    'procedure Other;'#13#10 +
    'implementation'#13#10'end.', S, E));
  Assert.AreEqual(3, S);
  Assert.AreEqual(5, E);
end;

procedure TMoveDeclarationRangeTests.ConstantsAndTypesWithInnerSemicolons;
const
  Src =
    'unit U;'#13#10 +                                        // 1
    'interface'#13#10 +                                      // 2
    'type'#13#10 +                                           // 3
    '  TProc2 = procedure(A: Integer;'#13#10 +               // 4
    '    B: Integer);'#13#10 +                               // 5
    '  TFoo ='#13#10 +                                       // 6
    '    class(TObject)'#13#10 +                             // 7
    '    procedure X;'#13#10 +                               // 8
    '  end;'#13#10 +                                         // 9
    '  TRec = packed record'#13#10 +                         // 10
    '    A: Integer;'#13#10 +                                // 11
    '    B: string;'#13#10 +                                 // 12
    '  end;'#13#10 +                                         // 13
    '  TRef = class of TFoo;'#13#10 +                        // 14
    '  TFwd = class;'#13#10 +                                // 15
    '  TOuter = class(TObject)'#13#10 +                      // 16
    '  public type'#13#10 +                                  // 17
    '    TInner = class'#13#10 +                             // 18
    '      procedure Y;'#13#10 +                             // 19
    '    end;'#13#10 +                                       // 20
    '    TLater = class;'#13#10 +                            // 21
    '  public'#13#10 +                                       // 22
    '    procedure Z;'#13#10 +                               // 23
    '  end;'#13#10 +                                         // 24
    '  TFwd = class(TFoo)'#13#10 +                           // 25
    '  end;'#13#10 +                                         // 26
    '  TOnlyFwd = class;'#13#10 +                            // 27
    'const'#13#10 +                                          // 28
    '  R: TPoint = (X: 1;'#13#10 +                           // 29
    '    Y: 2);'#13#10 +                                     // 30
    '  S = ''a;b'';'#13#10 +                                 // 31
    'implementation'#13#10 +
    'end.';
var
  S, E: Integer;
begin
  Assert.IsTrue(LocateMoveDeclaration('TProc2', Src, S, E), 'found: TProc2');
  Assert.AreEqual(4, S); Assert.AreEqual(5, E, 'procedural type with a wrapped list');
  Assert.IsTrue(LocateMoveDeclaration('TFoo', Src, S, E), 'found: TFoo');
  Assert.AreEqual(6, S); Assert.AreEqual(9, E, '"TFoo =" with the class on the next line');
  Assert.IsTrue(LocateMoveDeclaration('TRec', Src, S, E), 'found: TRec');
  Assert.AreEqual(10, S); Assert.AreEqual(13, E, 'a record runs to its end, not to its first field');
  Assert.IsTrue(LocateMoveDeclaration('TRef', Src, S, E), 'found: TRef');
  Assert.AreEqual(14, S); Assert.AreEqual(14, E, '"class of" has no body');
  Assert.IsTrue(LocateMoveDeclaration('TFwd', Src, S, E), 'found: TFwd');
  Assert.AreEqual(25, S); Assert.AreEqual(26, E, 'the real declaration wins over the forward one');
  Assert.IsTrue(LocateMoveDeclaration('TOnlyFwd', Src, S, E), 'found: TOnlyFwd');
  Assert.AreEqual(27, S); Assert.AreEqual(27, E, 'a forward declaration has no body');
  Assert.IsTrue(LocateMoveDeclaration('TOuter', Src, S, E), 'found: TOuter');
  Assert.AreEqual(16, S); Assert.AreEqual(24, E, 'nested class opens a level, nested forward does not');
  Assert.IsTrue(LocateMoveDeclaration('R', Src, S, E), 'found: R');
  Assert.AreEqual(29, S); Assert.AreEqual(30, E, 'record constant');
  Assert.IsTrue(LocateMoveDeclaration('S', Src, S, E), 'found: S');
  Assert.AreEqual(31, S); Assert.AreEqual(31, E, 'a ";" inside a string does not count');
end;

{ TRenameGuardTests }

procedure TRenameGuardTests.EditGuard_MatchesWholeWordsOnly;
var
  B: TBytes;
begin
  B := TEncoding.UTF8.GetBytes('    TestX: Integer;'#13#10'  Größe: Test;');
  // the report: the position holds "TestX" - "Test" is NOT there
  Assert.IsFalse(Utf8BufferHoldsAt(B, Length(B), 4, 'Test', True),
    '"Test" is only the start of "TestX"');
  Assert.IsTrue(Utf8BufferHoldsAt(B, Length(B), 4, 'TestX', True));
  Assert.IsTrue(Utf8BufferHoldsAt(B, Length(B), 4, 'testx', True), 'case-insensitive');
  // behind a non-ASCII identifier the byte offsets still fit ("Größe" is 7 bytes)
  var P := Length(TEncoding.UTF8.GetBytes('    TestX: Integer;'#13#10'  Größe: '));
  Assert.IsTrue(Utf8BufferHoldsAt(B, Length(B), P, 'Test', True), 'followed by ";"');
  Assert.IsFalse(Utf8BufferHoldsAt(B, Length(B), P, 'Tes', True), 'a prefix is no match');
  // a multi-line old text (with rewriter) - line breaks compared without #13
  Assert.IsTrue(Utf8BufferHoldsAt(B, Length(B), 4, 'TestX: Integer;'#13#10, True));
  Assert.IsFalse(Utf8BufferHoldsAt(B, Length(B), Length(B) - 2, 'Test;', True), 'beyond the end');
end;

procedure TRenameGuardTests.ForeignAnswerAtADeclaration_IsRecognised;
const
  Unit1 = 'C:\x\Unit1.pas';
  Win = 'C:\RAD\source\rtl\win\Winapi.Windows.pas';
begin
  // measured: DelphiLSP answers the FIELD declaration with Windows' type ABC
  Assert.IsTrue(DeclarationAnswerIsForeign('    ABC: Integer;', 'ABC', Unit1, Win));
  Assert.IsTrue(DeclarationAnswerIsForeign('    property ABC: Integer read FABC;', 'ABC', Unit1, Win));
  Assert.IsTrue(DeclarationAnswerIsForeign('    procedure ABC;', 'ABC', Unit1, Win));
  // the same file is the partner (declaration <-> implementation)
  Assert.IsFalse(DeclarationAnswerIsForeign('    procedure ABC;', 'ABC', Unit1, Unit1));
  // a use is no declaration - its answer stands
  Assert.IsFalse(DeclarationAnswerIsForeign('  C.ABC := 1;', 'ABC', Unit1, Win));
  // these really belong to the ancestor
  Assert.IsFalse(DeclarationAnswerIsForeign('    procedure Paint; override;', 'Paint', Unit1, Win));
  Assert.IsFalse(DeclarationAnswerIsForeign('    property Caption;', 'Caption', Unit1, Win));
end;

{ TPartnerQueryTests }

procedure TPartnerQueryTests.NameColumn_ImplementationHeaderPrefersTheMember;
const
  // the reported line, verbatim
  Impl = 'class function TGemTiFunctions.IsConnectorUnreachable(const AResultCode: Integer; ' +
    'const AErrorMessage: string): Boolean;';
begin
  // column 31 (0-based) = where DelphiLSP answers; column 0 is "class"
  Assert.AreEqual(31, NameColumnOnLine(Impl, 'IsConnectorUnreachable', 0));
  // case does not matter in Pascal
  Assert.AreEqual(31, NameColumnOnLine(Impl, 'isconnectorunreachable', 0));
  // a name that is also the TYPE part: the member after the dot wins
  // ("function Foo." = 13 characters)
  Assert.AreEqual(13, NameColumnOnLine('function Foo.Foo: Integer;', 'Foo', 0));
  // spaces around the dot are legal Delphi
  Assert.AreEqual(19, NameColumnOnLine('procedure TFoo  .  Bar;', 'Bar', 0));
end;

procedure TPartnerQueryTests.NameColumn_DeclarationCommentAndMisses;
begin
  // a declaration: the only occurrence
  Assert.AreEqual(19, NameColumnOnLine(
    '    class function IsConnectorUnreachable(const A: Integer): Boolean; static;',
    'IsConnectorUnreachable'));
  // a longer name containing it is no hit (whole word only)
  Assert.AreEqual(-1, NameColumnOnLine(
    'class function IsConnectorUnreachableError(const A: Integer): Boolean;',
    'IsConnectorUnreachable'));
  // a mention in a trailing comment is no hit either
  Assert.AreEqual(10, NameColumnOnLine('procedure Run; // calls Run again', 'Run'));
  Assert.AreEqual(-1, NameColumnOnLine('x := 1; // Run', 'Run'));
  // without a dot the hint decides between several occurrences
  // ("  A := B(A, A);" - the A's stand at 2, 9 and 12)
  Assert.AreEqual(12, NameColumnOnLine('  A := B(A, A);', 'A', 12));
  Assert.AreEqual(9, NameColumnOnLine('  A := B(A, A);', 'A', 9));
  // a hint between occurrences falls back to the first one
  Assert.AreEqual(2, NameColumnOnLine('  A := B(A, A);', 'A', 10));
end;

{ TAlignSignatureFixTests }

function AlignDemoLines: TArray<string>;
begin
  // the shape of the report: the class implements an interface, and the
  // DECLARATION got an extra first parameter
  Result := TArray<string>.Create(
    'unit KbDemo;',                                                          // 0
    'interface',                                                             // 1
    'type',                                                                  // 2
    '  IKb = interface',                                                     // 3
    '    procedure BindKeyboard(const BindingServices: IInterface);',        // 4
    '  end;',                                                                // 5
    '  TKb = class(TInterfacedObject, IKb)',                                 // 6
    '  public',                                                              // 7
    '    procedure BindKeyboard(const A: Integer; const BindingServices: IInterface);', // 8
    '    procedure Other(X: Integer); overload;',                           // 9
    '    procedure Other; overload;',                                        // 10
    '  end;',                                                                // 11
    'implementation',                                                        // 12
    'procedure TKb.BindKeyboard(const BindingServices: IInterface);',       // 13
    'begin',                                                                 // 14
    'end;',                                                                  // 15
    'procedure TKb.Other;',                                                  // 16
    'begin',                                                                 // 17
    'end;',                                                                  // 18
    'end.');                                                                 // 19
end;

function Diag(const ACode, AMsg: string; ALine, ACol, ALen: Integer): TLspErrorDiag;
begin
  Result := Default(TLspErrorDiag);
  Result.Code := ACode;
  Result.Message := AMsg;
  Result.Severity := 1;
  Result.Range.Start.Line := ALine;
  Result.Range.Start.Character := ACol;
  Result.Range.End_.Line := ALine;
  Result.Range.End_.Character := ACol + ALen;
end;

procedure TAlignSignatureFixTests.ChangedDeclaration_OffersBothAlignDirections;
var
  Lines: TArray<string>;
  Fixes: TArray<TQuickFix>;
begin
  Lines := AlignDemoLines;
  Assert.AreEqual(13, FindImplementationOfDecl(Lines, 8), 'found by NAME despite the other signature');
  Assert.AreEqual(16, FindImplementationOfDecl(Lines, 10), 'the parameterless Other');
  Fixes := ResolveQuickFixes(string.Join(#13#10, Lines), [
    Diag('E2291', 'E2291 Missing implementation of interface method KbDemo.IKb.BindKeyboard', 6, 2, 3),
    Diag('E2065', 'E2065 Unsatisfied forward or external declaration: ''TKb.BindKeyboard''', 8, 14, 12),
    Diag('E2065', 'E2065 Unsatisfied forward or external declaration: ''TKb.Other''', 9, 14, 5),
    Diag('E2037', 'E2037 Declaration of ''BindKeyboard'' differs from previous declaration', 13, 14, 12)]);
  var AtDecl := ''; var AtClass := ''; var AtOverload := ''; var AtImpl := '';
  for var F in Fixes do
    case F.Line of
      8: AtDecl := AtDecl + QuickFixKindText(F.Kind) + '|';
      6: AtClass := AtClass + QuickFixKindText(F.Kind) + '|';
      9: AtOverload := AtOverload + QuickFixKindText(F.Kind) + '|';
      13: AtImpl := AtImpl + QuickFixKindText(F.Kind) + '|';
    end;
  Assert.AreEqual('Align implementation header|Align declaration with implementation|', AtDecl,
    'at the edited declaration: both directions, and NO empty second body');
  Assert.AreEqual('', AtClass, 'no second declaration from the interface');
  Assert.AreEqual('Create implementation stub|', AtOverload,
    'a new OVERLOAD really needs its own body');
  Assert.AreEqual('Align implementation header|', AtImpl, 'the E2037 offer is unchanged');
  for var F in Fixes do
    if (F.Line = 8) and (F.Kind = qfAlignHeader) then
      Assert.AreEqual(13, F.AuxLine, 'anchored at the declaration, pointing at the body');
end;

procedure TAlignSignatureFixTests.AlignDeclaration_KeepsDirectivesAndTrailingDefaults;
var
  NewLines: TArray<string>;
begin
  Assert.IsTrue(PlanDeclToImplAlign(AlignDemoLines, 8, 13, NewLines));
  Assert.AreEqual('    procedure BindKeyboard(const BindingServices: IInterface);', NewLines[8],
    'indentation kept, parameters from the implementation');
  // 'class', directives and a surviving default value stay; the implementation
  // added a parameter WITHOUT default behind B, so B loses its default
  // (Delphi allows defaults only on a trailing run)
  var L := TArray<string>.Create(
    'unit U;', 'interface', 'type', '  TFoo = class',
    '    class function Run(A: Integer; B: string = ''x''): Boolean; static;',
    '    procedure Go(A: Integer = 1);',
    '    procedure Go3(A: Integer; B: string = ''x'');', '  end;', 'implementation',
    'class function TFoo.Run(A: Integer; B: string; C: Boolean): Boolean;',
    'begin', 'end;', 'procedure TFoo.Go(A: Integer; B: Integer);', 'begin', 'end;',
    'procedure TFoo.Go3(Z: Byte; A: Integer; B: string);', 'begin', 'end;', 'end.');
  Assert.AreEqual(9, FindImplementationOfDecl(L, 4));
  Assert.IsTrue(PlanDeclToImplAlign(L, 4, 9, NewLines));
  Assert.AreEqual('    class function Run(A: Integer; B: string; C: Boolean): Boolean; static;',
    NewLines[4], 'class + static kept; B loses its default because C follows without one');
  Assert.IsTrue(PlanDeclToImplAlign(L, 5, 12, NewLines));
  Assert.AreEqual('    procedure Go(A: Integer; B: Integer);', NewLines[5],
    'a default in front of a new plain parameter would not compile');
  Assert.IsTrue(PlanDeclToImplAlign(L, 6, 15, NewLines));
  Assert.AreEqual('    procedure Go3(Z: Byte; A: Integer; B: string = ''x'');', NewLines[6],
    'a surviving trailing default is KEPT - defaults live in the declaration');
  // an implementation header WITHOUT parameter list fits any declaration:
  // aligning to it would throw the declaration's parameters away
  var L2 := TArray<string>.Create('unit U;', 'interface', 'type', '  TFoo = class',
    '    procedure Put(A: Integer);', '  end;', 'implementation',
    'procedure TFoo.Put;', 'begin', 'end;', 'end.');
  Assert.IsFalse(PlanDeclToImplAlign(L2, 4, 7, NewLines), 'never strip the parameters');
end;

{ TUsesGraphDepthTests }

// A ring of AUnits units - U0 uses U1 ... U(n-1) uses U0 - built entirely
// through Analyze's own AReader seam, so the test touches no disk and no
// IDE. One ring is ONE strongly connected component, which is the worst
// case for the SCC pass: it must reach the whole chain before it can close
// a single component.
function RingResult(AUnits: Integer): TUsesCycleResult;
var
  Files: TArray<string>;
begin
  SetLength(Files, AUnits);
  for var I := 0 to AUnits - 1 do
    Files[I] := 'C:\ring\U' + IntToStr(I) + '.pas';
  Result := TUsesGraphAnalyzer.Analyze(Files, nil,
    function(APath: string): string
    var
      Idx: Integer;
    begin
      Idx := StrToIntDef(ChangeFileExt(ExtractFileName(APath), '').Substring(1), -1);
      if Idx < 0 then Exit('');
      Result := 'unit U' + IntToStr(Idx) + ';'#13#10 +
                'interface'#13#10 +
                'uses U' + IntToStr((Idx + 1) mod AUnits) + ';'#13#10 +
                'implementation'#13#10 +
                'end.';
    end);
end;

procedure TUsesGraphDepthTests.DeepCycle_DoesNotOverflowTheStack;
var
  Res: TUsesCycleResult;
begin
  // 20,000 is roughly three times the depth at which the recursive version
  // died. That margin IS the test, so do not lower it casually: anything at
  // or under ~6,000 passes either way and proves nothing.
  Res := RingResult(20000);
  try
    Assert.AreEqual<Integer>(20000, Length(Res.UnitNames), 'every unit is in the graph');
    Assert.AreEqual<Integer>(20000, Length(Res.Edges), 'a ring of N units has N cycle edges');
    Assert.AreEqual<Integer>(1, Length(Res.GroupInfos), 'the ring is a single component');
    Assert.AreEqual<Integer>(20000, Res.GroupInfos[0].UnitCount, 'and it spans every unit');
  finally
    Res.Free;
  end;
end;

procedure TUsesGraphDepthTests.Components_AreStillCorrectOnASmallGraph;
var
  Res: TUsesCycleResult;
begin
  // The companion to the depth test: the rewrite must still compute the
  // SAME partition. A 5-unit ring is small enough to check by hand - one
  // group, five units, five edges, girth five.
  Res := RingResult(5);
  try
    Assert.AreEqual<Integer>(1, Length(Res.GroupInfos), 'one group');
    Assert.AreEqual<Integer>(5, Res.GroupInfos[0].UnitCount, 'five units in it');
    Assert.AreEqual<Integer>(5, Res.GroupInfos[0].ShortestCycle, 'the girth of a 5-ring is 5');
    Assert.AreEqual<Integer>(5, Length(Res.Edges), 'five cycle edges');
  finally
    Res.Free;
  end;
end;

procedure TUsesGraphDepthTests.DeepEnumerateCycles_DoesNotOverflowTheStack;
var
  Res: TUsesCycleResult;
  Trunc: Boolean;
begin
  // The cycle enumerator recurses once per unit ON THE CURRENT PATH, and on
  // a ring the path is the whole ring. Its ceiling was LOWER than the SCC
  // pass's - it overflowed at 6,000 where StrongConnect reached 7,000 -
  // because the frame carries more locals. 12,000 is past both.
  Res := RingResult(12000);
  try
    var Cycles := Res.EnumerateCycles(10, Trunc, 0);
    Assert.AreEqual<Integer>(1, Length(Cycles), 'a ring has exactly one simple cycle');
    Assert.AreEqual<Integer>(12000, Length(Cycles[0].Units), 'and it spans every unit');
    Assert.IsFalse(Trunc, 'well under the count cap, so nothing was truncated');
  finally
    Res.Free;
  end;
end;

procedure TUsesGraphDepthTests.DeepEnumerateCyclesThrough_DoesNotOverflowTheStack;
var
  Res: TUsesCycleResult;
  Trunc: Boolean;
begin
  // Same walk, rooted at one unit - the Path tab of the results dialog.
  Res := RingResult(12000);
  try
    var Cycles := Res.EnumerateCyclesThrough('U0', 10, Trunc, 0);
    Assert.AreEqual<Integer>(1, Length(Cycles), 'one cycle through U0');
    Assert.AreEqual<Integer>(12000, Length(Cycles[0].Units), 'spanning every unit');
  finally
    Res.Free;
  end;
end;

procedure TUsesGraphDepthTests.DeepEdgeLevers_DoesNotOverflowTheStack;
var
  Res: TUsesCycleResult;
begin
  // EdgeLevers runs CountCycleNodes once per unique dependency, so this is
  // N+1 full SCC passes - the slowest test here by far, and the reason the
  // ring is 11,000 rather than larger. The recursive form of that pass
  // reached 8,000 and died at 10,000.
  Res := RingResult(11000);
  try
    var Levers := Res.EdgeLevers;
    Assert.AreEqual<Integer>(11000, Length(Levers), 'one lever per edge of the ring');
    // Removing any single edge of a ring breaks the whole cycle, so every
    // unit leaves the cyclic set - the same answer for every lever.
    Assert.AreEqual<Integer>(11000, Levers[0].UnitsFreed, 'cutting one edge frees the whole ring');
  finally
    Res.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TFileEncodingRegressionTests);
  TDUnitX.RegisterTestFixture(TUsesClauseRegressionTests);
  TDUnitX.RegisterTestFixture(TQuickFixSafetyTests);
  TDUnitX.RegisterTestFixture(TWithScannerRegressionTests);
  TDUnitX.RegisterTestFixture(TSourceScanRegressionTests);
  TDUnitX.RegisterTestFixture(TMiscRegressionTests);
  TDUnitX.RegisterTestFixture(TVersionTests);
  TDUnitX.RegisterTestFixture(TMcpPipeRegressionTests);
  TDUnitX.RegisterTestFixture(TLspProgressTests);
  TDUnitX.RegisterTestFixture(TSettingsRobustnessTests);
  TDUnitX.RegisterTestFixture(TPartnerQueryTests);
  TDUnitX.RegisterTestFixture(TAlignSignatureFixTests);
  TDUnitX.RegisterTestFixture(TUsesGraphDepthTests);
  TDUnitX.RegisterTestFixture(THelperUsageTests);
  TDUnitX.RegisterTestFixture(TUnitIndexCacheTests);
  TDUnitX.RegisterTestFixture(TMoveDeclarationRangeTests);
  TDUnitX.RegisterTestFixture(TRenameGuardTests);

end.
