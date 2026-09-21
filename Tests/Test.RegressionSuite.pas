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

implementation

uses
  System.SysUtils, System.IOUtils, System.Classes, System.SyncObjs,
  Delphi.FileEncoding, Expert.UsesEditor, Expert.AutoImport, Expert.UnitIndex,
  Expert.WithScanner, Lsp.Uri, Rename.WorkspaceEdit, Expert.VcsBlame,
  Expert.WorkerLatch, Expert.Version, Expert.PascalScanner, System.RegularExpressions,
  Winapi.Windows, Mcp.PipeServer, System.JSON, Lsp.Protocol;

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

end.
