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
  // The designer-managed uses fixture names TUsesEntryInfo / TUsesVerdict in
  // its own helpers, so those two live here rather than only below.
  DUnitX.TestFramework, Expert.UsesCleanup, Expert.EditorHelperIntf;

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
    /// <summary>Audit #38, L1g: the worker COUNT drops in the closure's
    ///  finally - before the closure epilogue releases what it captured and
    ///  before the thread itself has ended. A shutdown that watches only
    ///  the count therefore lets the BPL unload while plugin code is still
    ///  running on that thread; the Sleep(50) it used instead of a join was
    ///  a timing allowance, not a guarantee.</summary>
    [Test] procedure WorkerLatch_WaitsForTheThreadNotJustTheCount;
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
    /// <summary>An install that could not replace RefactoringLightMcp.exe left
    ///  a bridge nine days old talking to a current plugin, and NOTHING could
    ///  see it: the exe reported its version to its MCP client only. Now every
    ///  request carries it.</summary>
    [Test] procedure BridgeVersion_IsStampedAndReadBack;
    [Test] procedure BridgeVersion_MismatchIsNamedNotGuessed;
    [Test] procedure BridgeVersion_ReachesTheServerOverARealPipe;
    /// <summary>Audit #22, H4: Stop waited 5 s for the handlers and then
    ///  returned as if all was well, so the owner freed the server (and the
    ///  package unloaded) under a handler that ignores the stop event. Stop
    ///  now says so, and Destroy waits for the last handler.</summary>
    [Test] procedure Stop_ReportsAHandlerThatOutlivesTheDeadline;
    /// <summary>Audit #22, H5: handler threads were bare anonymous threads,
    ///  invisible to ShutdownWorkersAndWait - the unload never waited for a
    ///  handler that outlived Stop. They are started through the latch now.
    ///  </summary>
    [Test] procedure PipeHandler_IsCountedAsAWorker;
    /// <summary>Audit #36, M43: the bridge gave add_iinterface 45 s while
    ///  its own handler allows 120 s - so it reported "no answer from the
    ///  IDE" while the IDE was still writing. The list it was missing from
    ///  is hand-maintained, which is how it got there (stage 2 added six
    ///  tools and five got an entry), so the classification is TOTAL now
    ///  and this test is what refuses an unclassified one.</summary>
    [Test] procedure EveryToolGetsADeliberateWait;
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

  /// <summary>Found by running the check on the user's real project
  ///  (2026-09-30): it listed "result :" and "TAbortERezeptResult" as
  ///  interfaces without a GUID. Both lines merely CONTAIN the text
  ///  "= interface" - an assignment "result := InterfaceArrayFind(...)" and
  ///  an alias "TFoo = Interfaces.GTIDLL.TFoo;".</summary>
  [TestFixture]
  TInterfaceDeclLineTests = class
  public
    [Test] procedure RealDeclarations_AreRecognised;
    [Test] procedure AssignmentsAndAliases_AreNot;
  end;

  /// <summary>Found on the user's real project (2026-09-30): the signature
  ///  check reported TGemTiFunctions.IsConnectorUnreachable as DIVERGING from
  ///  a declaration it matches exactly. Normalize wanted to strip the
  ///  "TClass." qualifier so a declaration and its implementation compare
  ///  equal, but it copied from the dot onwards - taking the KEYWORD with it,
  ///  so every implementation differed from its own declaration.</summary>
  [TestFixture]
  TSignatureQualifierTests = class
  public
    [Test] procedure ImplementationAndDeclaration_NormalizeEqual;
  end;

  /// <summary>A forum report (2026-09-30) came with both logs: the first
  ///  Find-References run of a session listed 29 of 327 references, two of
  ///  them wrong; the second run, two minutes later, was right. The logs
  ///  name the cause - the declaration query answered nothing on the cold
  ///  session, the CARET was taken as the declaration, and every candidate
  ///  DelphiLSP later resolved correctly to the real declaration was dropped
  ///  as "leads to another symbol". An unanswered declaration on a USE line
  ///  is NO anchor.</summary>
  [TestFixture]
  TDeclarationAnchorTests = class
  public
    [Test] procedure NoAnswerOnAUseLine_IsNoAnchor;
    [Test] procedure NoAnswerOnADeclaration_IsTheAnchor;
    /// <summary>Second round of the same report: the answer pointed at a
    ///  CONTINUATION line of a wrapped header, where the name does not occur -
    ///  so the partner query found nothing and the warm run silently dropped
    ///  the declaration, the implementation and 7 calls.</summary>
    [Test] procedure AWrappedHeaderIsOnePosition;
    /// <summary>...and the cold run marked all 341 hits unverified although
    ///  326 of its own answers named the declaration.</summary>
    [Test] procedure TheAnswersThemselvesNameTheDeclaration;
    /// <summary>THIRD round of the same report: "beim ersten Durchlauf habe
    ///  ich 3 UNVERIFIED Treffer die im zweiten Durchlauf nicht mehr da sind".
    ///  The cold run listed 339 rows, the warm one 336 - and the three extra
    ///  rows were the SAME finding the second attempt had already dropped two
    ///  of, in the very same run. One rule for both passes.</summary>
    [Test] procedure AForeignAnswerIsJudgedTheSameByBothPasses;
  end;

  /// <summary>A forum report (2026-09-30): two event handlers sat on the
  ///  SAME line as the field declared before them
  ///  ("b_Cancel: TButton;procedure FormCreate(Sender: TObject);"), and the
  ///  DFM check reported both as MISSING - it only ever looked at the first
  ///  token of a line. Worse, its own "already declared?" guard looked the
  ///  same way, so applying the fix would have added a SECOND declaration
  ///  of each.</summary>
  [TestFixture]
  TDfmGluedDeclarationTests = class
  public
    [Test] procedure DeclarationStarts_AfterASemicolonOnTheSameLine;
    [Test] procedure DeclarationStarts_IgnoreCommentsStringsAndParameters;
  end;

  /// <summary>The quick fixes as TEXT (user request 2026-09-24: the MCP
  ///  tools preview a fix before it is applied). PlanQuickFixText is what
  ///  ApplyQuickFix writes, so a preview can never describe something else
  ///  than what happens.</summary>
  [TestFixture]
  TQuickFixPreviewTests = class
  private
    function Source: string;
  public
    [Test] procedure RemoveVar_KeepsTheSiblingsAndRefusesAStaleAnchor;
    [Test] procedure RemoveLastVar_TakesTheVarKeywordLineAlong;
    [Test] procedure DeclareVar_GoesBeforeBeginAndAddsItsUnit;
    [Test] procedure GeneratingKinds_HaveNoTextPreview;
  end;

  /// <summary>Issue #20: the uses cleanup offered units for removal that the
  ///  IDE's FORM DESIGNER writes back on the next save (cxGraphics,
  ///  cxControls, ... with DevExpress; Data.DB for a plain TDBGrid, because
  ///  its selection editor asks for it through
  ///  ISelectionEditor.RequiresUnits). None of them appears as an identifier
  ///  in the .pas, so no textual analysis can see them - the designer of the
  ///  loaded form is the only exact source, and a hand-written list of four
  ///  VCL units can never cover third-party or in-house packages.
  ///  AnalyzeUses stays pure: what the designer answered comes in as a
  ///  lookup, so these tests script it.</summary>
  [TestFixture]
  TDesignerManagedUsesTests = class
  private
    /// <summary>A form unit that uses a unit it never NAMES - exactly the
    ///  shape of the report.</summary>
    function FormUnit: string;
    function Analyze(const AContent: string;
      const ADesignerUnits: TArray<string>; const AKeepList: string = '';
      AFormUnverified: Boolean = False): TArray<TUsesEntryInfo>;
    function VerdictOf(const AEntries: TArray<TUsesEntryInfo>;
      const AUnit: string): TUsesVerdict;
    function ReasonOf(const AEntries: TArray<TUsesEntryInfo>;
      const AUnit: string): string;
  public
    [Test] procedure ADesignerRequiredUnitIsKeptAndNamesTheComponent;
    [Test] procedure ADesignerRequiredUnitIsNotMovedDownEither;
    [Test] procedure UnitScopeNamesMatchBothWays;
    [Test] procedure TheKeepListTakesMasks;
    [Test] procedure AFormWhoseDesignerCannotBeAskedIsUnverified;
    [Test] procedure AnIncompleteAnswerStillProtectsWhatItDidReport;
    [Test] procedure OptingIntoAnUnverifiedRowFollowsTheText;
  end;

  /// <summary>A fork maintainer's report (2026-10-01), reproduced live on
  ///  this repository's own Expert.OptionsFrame.pas: the clause names came
  ///  out of the RAW text with only // comments cut, so a directive was
  ///  GLUED onto its neighbour ("Expert.PluginSettings {$IFNDEF
  ///  STANDALONE_BUILD}"), a ';' inside a { } comment ENDED the clause and
  ///  hid every unit behind it, and an "in '<path>'" tail stayed part of the
  ///  name. The names are read from the MASKED text now - and because the
  ///  LSP session reports which lines are inactive, a conditional clause can
  ///  be judged per entry instead of not at all.</summary>
  [TestFixture]
  TUsesClauseParsingTests = class
  private
    function Analyze(const AClause: string;
      const AInactive: TArray<Integer> = nil): TArray<TUsesEntryInfo>;
    function Names(const AEntries: TArray<TUsesEntryInfo>): string;
    function VerdictOf(const AEntries: TArray<TUsesEntryInfo>;
      const AUnit: string): TUsesVerdict;
    function ReasonOf(const AEntries: TArray<TUsesEntryInfo>;
      const AUnit: string): string;
  public
    [Test] procedure ADirectiveIsNotPartOfTheUnitName;
    [Test] procedure ASemicolonInACommentDoesNotEndTheClause;
    [Test] procedure AnInPathTailIsNotPartOfTheName;
    [Test] procedure APlainClauseIsUnchangedAndStillJudged;
    [Test] procedure AnUnknownVerdictSaysWhy;
    [Test] procedure AnInactiveBranchIsNeverJudgedButItsNeighbourIs;
  end;

  /// <summary>Two guards the plugin keeps on ITSELF, both earned.
  ///  (1) ToolsAPI is main thread only, and breaking that rule does not
  ///  fail - it corrupts a buffer now and then. A fork audit found three MCP
  ///  tools applying their edit from the pipe handler thread, which no test
  ///  could have caught, so every write path reports itself now.
  ///  (2) analyze_uses answered 'Expert.PluginSettings {$IFNDEF
  ///  STANDALONE_BUILD}' as a unit NAME for two days. That is structurally
  ///  impossible, and a tool that notices it says so instead of passing it
  ///  to a caller who has no way to doubt it.
  ///  (3) an argument name that is not in a tool's schema is never read, so
  ///  the call ran with a default and the answer looked like a tool that
  ///  did something else than it was asked (found by making that mistake:
  ///  buffer_read with "from_line" answered the whole file).</summary>
  [TestFixture]
  TSelfProtectionTests = class
  public
    [Test] procedure AWriteFromAWorkerThreadIsRecordedAndCanBeRefused;
    [Test] procedure AnImpossibleUnitNameIsReportedAsADefect;
    [Test] procedure AnArgumentTheToolDoesNotKnowIsReported;
    /// <summary>Audit #36, H34: a preview token recorded only the tool
    ///  name and the buffer hashes, so the token from a preview of
    ///  "add_unit Foo" was accepted by "add_unit apply=true Bar" - the
    ///  answer then described Bar while the user had reviewed Foo.</summary>
    [Test] procedure APreviewTokenBelongsToItsArguments;
  end;

  /// <summary>Audit issue #41, the three High findings of "Remove with".
  ///  All three produce code that COMPILES and behaves differently, which
  ///  is why each is refused rather than guessed at.</summary>
  [TestFixture]
  TRemoveWithSafetyTests = class
  public
    /// <summary>H11: a temp of a record / object type is a COPY, so a
    ///  write through the with is lost.</summary>
    [Test] procedure AValueTypeIsRecognisedInItsDeclaration;
    /// <summary>H13: the classic temp of a with inside an anonymous method
    ///  would land in the OUTER method's var section - one variable shared
    ///  by every invocation.</summary>
    [Test] procedure AWithInsideAnAnonymousMethodIsRecognised;
    /// <summary>H15: an insertion carries no old text, so the apply has to
    ///  check the OTHER edits of the same file before writing any.</summary>
    [Test] procedure AnEditIsVerifiedAgainstTheCurrentText;
  end;

implementation

uses
  System.SysUtils, System.IOUtils, System.Classes, System.SyncObjs,
  Delphi.FileEncoding, Expert.UsesEditor, Expert.AutoImport, Expert.UnitIndex,
  Expert.DfmEventCheck, Expert.SignatureCheck, Expert.InterfaceGuidCheck,
  Expert.WithScanner, Lsp.Uri, Rename.WorkspaceEdit, Expert.VcsBlame,
  Expert.WorkerLatch, Expert.Version, Expert.PascalScanner, System.RegularExpressions,
  Winapi.Windows, Mcp.PipeServer, Mcp.Protocol, Mcp.Bridge, System.JSON, Lsp.Protocol,
  System.Win.Registry, Expert.PluginSettings, Expert.UsesGraph,
  Expert.MoveToUnit, Expert.SafeDeletePlan, Expert.McpTools, Expert.WithRewriter,
  System.StrUtils;

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

type
  // Its destructor is the "plugin code still running after the counter
  // dropped": the worker's closure captures it, so it is released by the
  // closure EPILOGUE - after the latch has already been left. Shape taken
  // from Ian Branch's own repro for this item (audit #38, L1g).
  TClosureTail = class(TInterfacedObject)
  public
    destructor Destroy; override;
  end;

var
  GTailRan: Integer = 0;

destructor TClosureTail.Destroy;
begin
  Sleep(300);
  AtomicExchange(GTailRan, 1);
  inherited;
end;

// THE POINT OF THESE TWO HELPERS, and my own first version got it wrong:
// a variable captured by a closure lives in a frame object that the
// ENCLOSING METHOD holds as well. Written inside the test method, clearing
// the local would destroy the tail on the MAIN thread and the test would
// pass whatever the shutdown does. Built here, only the returned TProc
// holds the frame - and StartTailWorker keeps the test method from holding
// the TProc itself.
function MakeTailProc(AStarted: TEvent): TProc;
var
  Tail: IInterface;
begin
  Tail := TClosureTail.Create;
  Result :=
    procedure
    begin
      if Tail <> nil then
        AStarted.SetEvent;
    end;
end;

function StartTailWorker(AStarted: TEvent): Boolean;
begin
  Result := StartWorker(MakeTailProc(AStarted));
end;

procedure TMiscRegressionTests.WorkerLatch_WaitsForTheThreadNotJustTheCount;
var
  Started: TEvent;
begin
  ResetWorkerLatch;
  AtomicExchange(GTailRan, 0);
  Started := TEvent.Create(nil, True, False, '');
  try
    Assert.IsTrue(StartTailWorker(Started), 'the worker starts');
    Assert.IsTrue(Started.WaitFor(5000) = wrSignaled, 'the worker runs');
    Assert.IsTrue(ShutdownWorkersAndWait(10000),
      'the shutdown reports that every worker is gone');
    Assert.AreEqual<Integer>(1, AtomicCmpExchange(GTailRan, 0, 0),
      'and it may only report that once the thread really ended - the ' +
      'captured object''s destructor is plugin code on that thread');
  finally
    // let the tail finish before the next test, whatever the verdict
    for var I := 1 to 100 do
      if AtomicCmpExchange(GTailRan, 0, 0) = 1 then Break else Sleep(50);
    Started.Free;
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

procedure TMcpPipeRegressionTests.BridgeVersion_IsStampedAndReadBack;
begin
  var R := StampBridgeVersion('{"method":"call","tool":"x"}', '1.15.2');
  Assert.AreEqual('{"bridge":"1.15.2","method":"call","tool":"x"}', R,
    'the stamp goes behind the brace, the rest is untouched');
  Assert.AreEqual('1.15.2', BridgeVersionOfRequest(R));
  // it must stay valid JSON for the parser the server really uses
  var O := TJSONObject.ParseJSONValue(R) as TJSONObject;
  Assert.IsNotNull(O, 'a stamped request still parses');
  try
    Assert.AreEqual('1.15.2', O.GetValue<string>('bridge'));
    Assert.AreEqual('call', O.GetValue<string>('method'));
  finally
    O.Free;
  end;
  // an object without members must not end up as '{"bridge":"x",}'
  Assert.AreEqual('{"bridge":"1.15.2"}', StampBridgeVersion('{}', '1.15.2'));
  // a bridge from before this change sends nothing - not an error, just unknown
  Assert.AreEqual('', BridgeVersionOfRequest('{"method":"context"}'));
  Assert.AreEqual('{"method":"context"}',
    StampBridgeVersion('{"method":"context"}', ''), 'no version, no stamp');
end;

procedure TMcpPipeRegressionTests.BridgeVersion_MismatchIsNamedNotGuessed;
begin
  Assert.AreEqual('', BridgeVersionProblem('1.15.2', '1.15.2'),
    'equal versions have nothing to report');
  Assert.AreEqual('', BridgeVersionProblem('1.15.2', ''),
    'without a plugin version there is no verdict to give');
  var S := BridgeVersionProblem('1.2.0', '1.15.2');
  Assert.IsTrue(S.Contains('1.2.0') and S.Contains('1.15.2'),
    'BOTH numbers must be in the message, that is the whole point: ' + S);
  Assert.IsTrue(S.Contains('install.cmd'), 'and what to do about it: ' + S);
  Assert.IsTrue(BridgeVersionProblem('', '1.15.2').Contains('1.15.2'),
    'a bridge too old to report its version is still reported');
end;

procedure TMcpPipeRegressionTests.BridgeVersion_ReachesTheServerOverARealPipe;
var
  Srv: TMcpPipeServer;
  Seen, Probe, Resp, Err, Mine: string;
begin
  // Serve THIS process's own pipe and send through the PRODUCTION transport,
  // because the point is not that a stamped request can be read - it is that
  // the one funnel every request goes through really stamps. A future request
  // site that bypasses it fails here.
  // A real Claude Code bridge polls every Refactoring Light pipe it finds, so
  // this must not assume it is the only client: our own request is recognised
  // by a probe value, and only our own PID is looked up.
  Probe := 'p' + FormatDateTime('hhnnsszzz', Now);
  Srv := TMcpPipeServer.Create(McpPipeName(GetCurrentProcessId),
    function(const ARequest: string; AStop: THandle): string
    begin
      if ARequest.Contains(Probe) then Seen := ARequest;
      Result := '{"ok":true}';
    end);
  try
    Assert.IsTrue(Srv.Start, 'the test pipe is there: ' + Srv.LastError);
    var T: IMcpTransport := TPipeTransport.Create;
    Assert.IsTrue(T.Request(GetCurrentProcessId,
      '{"method":"context","probe":"' + Probe + '"}', 5000, Resp, Err),
      'the request went through: ' + Err);
    Assert.AreEqual(BridgeVersion, BridgeVersionOfRequest(Seen),
      'the production transport stamps its version into every request');
    for var C in Srv.RecentClients(60000) do
      if C.Pid = GetCurrentProcessId then Mine := C.Version;
    Assert.AreEqual(BridgeVersion, Mine,
      'and the IDE side keeps it per client, which is what the status row reads');
    Srv.Stop;
  finally
    Srv.Free;
  end;
end;

// Sends one request carrying AProbe to this process's pipe on a thread of
// its own (the handler under test blocks, so the caller must not).
function StartProbeClient(const AProbe: string): TThread;
begin
  Result := TThread.CreateAnonymousThread(
    procedure
    var
      Resp, Err: string;
    begin
      McpPipeRequest(GetCurrentProcessId,
        '{"method":"context","probe":"' + AProbe + '"}', 10000, Resp, Err);
    end);
  Result.FreeOnTerminate := False;
  Result.Start;
end;

procedure TMcpPipeRegressionTests.Stop_ReportsAHandlerThatOutlivesTheDeadline;
var
  Srv: TMcpPipeServer;
  Entered, Release: TEvent;
  Client: TThread;
  Probe: string;
begin
  // A real Claude Code bridge may poll this pipe too - only OUR request
  // (recognised by the probe) blocks, every other one is answered at once.
  Probe := 'h4-' + FormatDateTime('hhnnsszzz', Now);
  Entered := TEvent.Create(nil, True, False, '');
  Release := TEvent.Create(nil, True, False, '');
  try
    Srv := TMcpPipeServer.Create(McpPipeName(GetCurrentProcessId),
      function(const ARequest: string; AStop: THandle): string
      begin
        if ARequest.Contains(Probe) then
        begin
          Entered.SetEvent;
          // ignores AStop - like an LSP wait or a long lsp_request
          Release.WaitFor(10000);
        end;
        Result := '{"ok":true}';
      end);
    try
      Assert.IsTrue(Srv.Start, 'the test pipe is there: ' + Srv.LastError);
      Client := StartProbeClient(Probe);
      try
        Assert.IsTrue(Entered.WaitFor(5000) = wrSignaled, 'the handler runs');
        Assert.IsFalse(Srv.Stop(100),
          'a handler still runs after the deadline - Stop must say so');
        Assert.IsTrue(Srv.ActiveHandlers >= 1, 'and it is still counted');
        Release.SetEvent;
        Assert.IsTrue(Srv.Stop(5000), 'once it has left, the stop is clean');
        Assert.AreEqual(0, Srv.ActiveHandlers);
      finally
        Release.SetEvent;
        Client.WaitFor;
        Client.Free;
      end;
    finally
      Srv.Free;
    end;
  finally
    Release.Free;
    Entered.Free;
  end;
end;

procedure TMcpPipeRegressionTests.PipeHandler_IsCountedAsAWorker;
var
  Srv: TMcpPipeServer;
  Entered, Release: TEvent;
  Client: TThread;
  Probe: string;
  Before, During: Integer;
begin
  Probe := 'h5-' + FormatDateTime('hhnnsszzz', Now);
  ResetWorkerLatch;
  Before := ActiveWorkerCount;
  Entered := TEvent.Create(nil, True, False, '');
  Release := TEvent.Create(nil, True, False, '');
  try
    Srv := TMcpPipeServer.Create(McpPipeName(GetCurrentProcessId),
      function(const ARequest: string; AStop: THandle): string
      begin
        if ARequest.Contains(Probe) then
        begin
          Entered.SetEvent;
          Release.WaitFor(10000);
        end;
        Result := '{"ok":true}';
      end);
    try
      Assert.IsTrue(Srv.Start, 'the test pipe is there: ' + Srv.LastError);
      Client := StartProbeClient(Probe);
      try
        Assert.IsTrue(Entered.WaitFor(5000) = wrSignaled, 'the handler runs');
        During := ActiveWorkerCount;
      finally
        Release.SetEvent;
        Client.WaitFor;
        Client.Free;
      end;
      Srv.Stop;
    finally
      Srv.Free;
    end;
  finally
    Release.Free;
    Entered.Free;
  end;
  Assert.IsTrue(During > Before, Format('a running pipe handler must be ' +
    'counted by the worker latch (workers before: %d, while it ran: %d)',
    [Before, During]));
end;

procedure TMcpPipeRegressionTests.EveryToolGetsADeliberateWait;
var
  Missing: TArray<string>;
begin
  // THE point of the test: no tool of the shipped list may fall through.
  // A new tool fails here with its own name until someone decides whether
  // it answers in seconds or can run for minutes.
  // ... and it must not pass because the list came back empty: count the
  // tools first, or a broken McpToolDefinitions would make this vacuous.
  var Arr := McpToolDefinitions;
  try
    Assert.IsTrue((Arr <> nil) and (Arr.Count > 40),
      'the shipped tool list is what this test walks');
  finally
    Arr.Free;
  end;
  Missing := UnclassifiedWaitTools;
  Assert.AreEqual(0, Integer(Length(Missing)),
    'every tool needs a deliberate wait; unclassified: ' +
    string.Join(', ', Missing));

  // the reported one, and the neighbour that was close enough to 45 s to
  // be luck rather than design
  Assert.AreEqual(Ord(twLong), Ord(ToolWaitClass('add_iinterface')),
    'its handler allows 120 s');
  Assert.AreEqual(Ord(twLong), Ord(ToolWaitClass('get_quick_fixes')));
  // and the bound really is above what the IDE allows itself
  Assert.IsTrue(ToolTimeout('add_iinterface') > 120000,
    'the bridge must not be the one that gives up first');
  Assert.IsTrue(ToolTimeout('rename_apply') >= 300000);

  // a read stays short - a wedged IDE must not block a buffer_read for
  // minutes
  Assert.AreEqual(Ord(twQuick), Ord(ToolWaitClass('buffer_read')));
  Assert.AreEqual(Ord(twQuick), Ord(ToolWaitClass('lsp_hover')));
  Assert.AreEqual(Ord(twQuick), Ord(ToolWaitClass('ide_instances')));
  Assert.IsTrue(ToolTimeout('buffer_read') < 60000);

  // an unknown name is unclassified (that is what the first assertion can
  // fail on) but at RUNTIME it still gets the safe bound
  Assert.AreEqual(Ord(twUnclassified), Ord(ToolWaitClass('no_such_tool')));
  Assert.IsTrue(ToolTimeout('no_such_tool') >= 300000,
    'too short is the dangerous direction, so unclassified waits long');
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

{ TInterfaceDeclLineTests }

procedure TInterfaceDeclLineTests.RealDeclarations_AreRecognised;
var
  Name: string;
  Disp: Boolean;
begin
  Assert.IsTrue(IsInterfaceDeclLine('  IFoo = interface', Name, Disp));
  Assert.AreEqual('IFoo', Name);
  Assert.IsFalse(Disp);
  Assert.IsTrue(IsInterfaceDeclLine('  IFoo = interface(IBase)', Name, Disp));
  Assert.AreEqual('IFoo', Name);
  Assert.IsTrue(IsInterfaceDeclLine('  IFoo = dispinterface', Name, Disp));
  Assert.IsTrue(Disp, 'a dispinterface is one');
  // generics, as the real project has them
  Assert.IsTrue(IsInterfaceDeclLine('  IList<T: IObject> = interface', Name, Disp));
  Assert.AreEqual('IList<T: IObject>', Name);
  // a FORWARD declaration carries no GUID - the real one elsewhere does
  Assert.IsFalse(IsInterfaceDeclLine('  IFoo = interface;', Name, Disp));
end;

procedure TInterfaceDeclLineTests.AssignmentsAndAliases_AreNot;
var
  Name: string;
  Disp: Boolean;
begin
  // the two lines the real project produced, verbatim
  Assert.IsFalse(IsInterfaceDeclLine(
    '  result := InterfaceArrayFind(aInterfaceArray,aItem);', Name, Disp),
    'an assignment is no declaration');
  Assert.IsFalse(IsInterfaceDeclLine(
    '  TAbortERezeptResult = Interfaces.GTIDLL.TAbortERezeptResult;', Name, Disp),
    'an alias to another unit''s type is no declaration');
  // near misses
  Assert.IsFalse(IsInterfaceDeclLine('  X := Interfaces.Foo;', Name, Disp));
  Assert.IsFalse(IsInterfaceDeclLine('  if A <= InterfaceCount then', Name, Disp));
  Assert.IsFalse(IsInterfaceDeclLine('  TFoo = InterfaceHelper;', Name, Disp));
  // ... and the keyword still wins when it really is one
  Assert.IsTrue(IsInterfaceDeclLine('  IFoo=interface', Name, Disp), 'no spaces');
end;

{ TSignatureQualifierTests }

procedure TSignatureQualifierTests.ImplementationAndDeclaration_NormalizeEqual;
begin
  // the reported pair, verbatim
  Assert.AreEqual(
    TSignatureChecker.Normalize(
      'function IsConnectorUnreachable(const AResultCode: Integer; ' +
      'const AErrorMessage: string): Boolean;'),
    TSignatureChecker.Normalize(
      'function TGemTiFunctions.IsConnectorUnreachable(const AResultCode: ' +
      'Integer; const AErrorMessage: string): Boolean;'),
    'the qualifier must not take the keyword with it');
  // the keyword SURVIVES - a procedure and a function of the same name and
  // parameters must still differ
  Assert.AreNotEqual(
    TSignatureChecker.Normalize('procedure TFoo.Bar(A: Integer);'),
    TSignatureChecker.Normalize('function TFoo.Bar(A: Integer): Boolean;'));
  Assert.IsTrue(TSignatureChecker.Normalize(
    'function TFoo.Bar(A: Integer): Boolean;').StartsWith('function '),
    'the keyword is kept');
  // a constructor, and a generic type qualifier
  Assert.AreEqual(
    TSignatureChecker.Normalize('constructor Create(AOwner: TComponent);'),
    TSignatureChecker.Normalize('constructor TFoo.Create(AOwner: TComponent);'));
  Assert.AreEqual(
    TSignatureChecker.Normalize('procedure Add(const AItem: T);'),
    TSignatureChecker.Normalize('procedure TList<T>.Add(const AItem: T);'));
  // an UNqualified implementation header (a free routine) is unchanged
  Assert.AreEqual(
    TSignatureChecker.Normalize('procedure DoIt(A: Integer);'),
    TSignatureChecker.Normalize('procedure DoIt(A: Integer);'));
end;

{ TDeclarationAnchorTests }

procedure TDeclarationAnchorTests.AWrappedHeaderIsOnePosition;
var
  First, Last: Integer;
begin
  // the reported shape: the routine starts on the first line, DelphiLSP
  // answered with the LAST one (column 1), where "BTB" does not occur
  var Lines: TArray<string> := [
    'implementation',                                     // 0
    '',                                                   // 1
    'procedure BTB( const xEintrag: String;',             // 2
    '               const xTyp: TBtbTyp;',                // 3
    '               const xRezeptur: String;',            // 4
    '               const xKunde: Integer;',              // 5
    '               const xBemerkung: String );',         // 6
    'begin',                                              // 7
    '  DoSomething;',                                     // 8
    'end;'];                                              // 9
  Assert.IsTrue(DeclarationHeaderSpan(Lines, 6, 'BTB', First, Last),
    'the continuation line belongs to the header above it');
  Assert.AreEqual(2, First, 'the header starts where the NAME stands');
  Assert.AreEqual(6, Last, 'and ends where its parameter list closes');
  // asked at the header line itself the answer must not move
  Assert.IsTrue(DeclarationHeaderSpan(Lines, 2, 'BTB', First, Last));
  Assert.AreEqual(2, First);
  Assert.AreEqual(6, Last);
  // a line inside the BODY is no header of BTB - the walk up must not run
  // into the header and turn a call into the declaration
  Assert.IsFalse(DeclarationHeaderSpan(Lines, 8, 'DoSomething', First, Last),
    'a call is not a header');
  // and a one-line header stays one line
  var One: TArray<string> := ['procedure BTB(const x: String);'];
  Assert.IsTrue(DeclarationHeaderSpan(One, 0, 'BTB', First, Last));
  Assert.AreEqual(0, First);
  Assert.AreEqual(0, Last);
end;

procedure TDeclarationAnchorTests.TheAnswersThemselvesNameTheDeclaration;
var
  F: string;
  L: Integer;
begin
  // the reported run: 326 of 341 candidates answered ROM_Utils.pas:15589,
  // 9 answered 15584 (the same routine) and 5 another symbol
  var Files: TArray<string> := nil;
  var Lines: TArray<Integer> := nil;
  for var I := 1 to 326 do
  begin
    Files := Files + ['D:\Sources\ROM_Utils.pas'];
    Lines := Lines + [15588];
  end;
  for var I := 1 to 9 do
  begin
    Files := Files + ['D:\Sources\ROM_Utils.pas'];
    Lines := Lines + [15583];
  end;
  for var I := 1 to 5 do
  begin
    Files := Files + ['D:\Sources\UGlobalRomConfig.pas'];
    Lines := Lines + [1585];
  end;
  Assert.AreEqual(326, DominantAnswer(Files, Lines, F, L), 'the agreed position');
  Assert.AreEqual('D:\Sources\ROM_Utils.pas', F);
  Assert.AreEqual(15588, L);

  // scattered answers are NOT a declaration - three files, no majority
  var S: TArray<string> := ['a.pas', 'b.pas', 'c.pas', 'd.pas'];
  var SL: TArray<Integer> := [1, 2, 3, 4];
  Assert.AreEqual(0, DominantAnswer(S, SL, F, L), 'no agreement, no anchor');
  Assert.AreEqual('', F);
  // a single answer is not evidence either
  Assert.AreEqual(0, DominantAnswer(['a.pas'], [1], F, L));
  // unanswered candidates ('' file) do not count against the majority
  Assert.AreEqual(3, DominantAnswer(['x.pas', '', 'x.pas', '', 'x.pas'],
    [7, 0, 7, 0, 7], F, L), 'three agreeing answers among five candidates');
  Assert.AreEqual(7, L);
end;

procedure TDeclarationAnchorTests.NoAnswerOnAUseLine_IsNoAnchor;
begin
  // the reported lines, verbatim
  Assert.IsTrue(DeclarationAnchorUnknown(False,
    '  lCds.FieldByName(''WERT'').AsString := GlobalConfig.Formulare.BTB;', 'BTB'),
    'a use of a nested record field');
  Assert.IsTrue(DeclarationAnchorUnknown(False,
    '  BTB(Format(_(''%s: Kunde %s gespeichert''), A, B));', 'BTB'), 'a call');
  // an answer always wins - there IS an anchor then
  Assert.IsFalse(DeclarationAnchorUnknown(True,
    '  lCds.FieldByName(''WERT'').AsString := GlobalConfig.Formulare.BTB;', 'BTB'));
  // a declaration line of ANOTHER name is no anchor for ours
  Assert.IsTrue(DeclarationAnchorUnknown(False,
    '    procedure SetStatus(const AText: string);', 'BTB'));
end;

procedure TDeclarationAnchorTests.NoAnswerOnADeclaration_IsTheAnchor;
begin
  // DelphiLSP answers null AT a declaration - the caret IS the symbol there,
  // which is what makes renaming an interface method from its declaration
  // work (IRenameHost.SetStatus); this must not regress
  Assert.IsFalse(DeclarationAnchorUnknown(False,
    '    procedure SetStatus(const AText: string);', 'SetStatus'));
  Assert.IsFalse(DeclarationAnchorUnknown(False, '    BTB: string;', 'BTB'));
  Assert.IsFalse(DeclarationAnchorUnknown(False,
    '    property BTB: string read FBTB write SetBTB;', 'BTB'));
end;

{ TDfmGluedDeclarationTests }

procedure TDfmGluedDeclarationTests.DeclarationStarts_AfterASemicolonOnTheSameLine;
const
  // the reported line, verbatim
  Glued = '    b_Cancel: TButton;procedure FormCreate(Sender: TObject); ' +
    'procedure ControlListBeforeDrawItem(AIndex: Integer; ACanvas: TCanvas;';
var
  Starts: TArray<Integer>;
begin
  Starts := DeclarationStartsOnLine(Glued);
  Assert.AreEqual(2, Integer(Length(Starts)), 'both handlers are declarations');
  Assert.AreEqual('procedure', Copy(Glued, Starts[0], 9));
  Assert.AreEqual('procedure', Copy(Glued, Starts[1], 9));
  // a field alone is none
  Assert.AreEqual(0, Integer(Length(DeclarationStartsOnLine('    b_Cancel: TButton;'))));
  // an ordinary declaration: one start, at the first non-blank character
  Starts := DeclarationStartsOnLine('    procedure FormCreate(Sender: TObject);');
  Assert.AreEqual(1, Integer(Length(Starts)));
  Assert.AreEqual(5, Starts[0]);
  // a continuation line of a wrapped header is no start
  Assert.AreEqual(0, Integer(Length(DeclarationStartsOnLine(
    '      ARect: TRect; AState: TOwnerDrawState);'))));
end;

procedure TDfmGluedDeclarationTests.DeclarationStarts_IgnoreCommentsStringsAndParameters;
begin
  // a PROCEDURAL PARAMETER is inside the list - not a second declaration
  Assert.AreEqual(1, Integer(Length(DeclarationStartsOnLine(
    '    procedure Run(P: procedure; A: Integer);'))), 'procedural parameter');
  // 'class procedure' answers AT the procedure, so the name parsing
  // downstream is unchanged
  var CLine := '    class procedure Init; static;';
  var CS := DeclarationStartsOnLine(CLine);
  Assert.AreEqual(1, Integer(Length(CS)));
  Assert.AreEqual('procedure', Copy(CLine, CS[0], 9), 'the class prefix is skipped');
  // comments and strings open nothing
  Assert.AreEqual(0, Integer(Length(DeclarationStartsOnLine('    // procedure Foo;'))));
  Assert.AreEqual(0, Integer(Length(DeclarationStartsOnLine('    (* procedure Foo; *)'))));
  Assert.AreEqual(0, Integer(Length(DeclarationStartsOnLine('    C := ''x;procedure Foo;'';'))));
  Assert.AreEqual(1, Integer(Length(DeclarationStartsOnLine(
    '    { procedure Foo; } procedure Bar;'))), 'only the real one');
end;

{ TQuickFixPreviewTests }

function TQuickFixPreviewTests.Source: string;
begin
  Result :=
    'unit QPrev;'#13#10 +
    ''#13#10 +
    'interface'#13#10 +
    ''#13#10 +
    'uses'#13#10 +
    '  System.SysUtils;'#13#10 +
    ''#13#10 +
    'implementation'#13#10 +
    ''#13#10 +
    'procedure Foo;'#13#10 +
    'var'#13#10 +
    '  A, X, B: Integer;'#13#10 +
    'begin'#13#10 +
    '  A := 1;'#13#10 +
    'end;'#13#10 +
    ''#13#10 +
    'end.'#13#10;
end;

procedure TQuickFixPreviewTests.RemoveVar_KeepsTheSiblingsAndRefusesAStaleAnchor;
var
  F: TQuickFix;
  NewText: string;
begin
  F := Default(TQuickFix);
  F.Kind := qfRemoveVar;
  F.Line := 11;                       // '  A, X, B: Integer;'
  F.Identifier := 'X';
  Assert.IsTrue(PlanQuickFixText(Source, F, 0, NewText));
  Assert.IsTrue(Pos('  A, B: Integer;', NewText) > 0, 'the siblings stay');
  Assert.AreEqual(0, Pos(' X', NewText), 'X is gone');
  // nothing else moved: exactly one line differs from the original
  var Differs := 0;
  var Old := Source.Split([#13#10]);
  var New := NewText.Split([#13#10]);
  Assert.AreEqual(Length(Old), Length(New), 'same line count');
  for var I := 0 to High(Old) do
    if Old[I] <> New[I] then Inc(Differs);
  Assert.AreEqual(1, Differs);
  // an anchor that no longer holds the name is stale - and writes nothing
  F.Identifier := 'NotThere';
  Assert.IsFalse(PlanQuickFixText(Source, F, 0, NewText));
end;

procedure TQuickFixPreviewTests.RemoveLastVar_TakesTheVarKeywordLineAlong;
var
  F: TQuickFix;
  NewText: string;
begin
  var Src := StringReplace(Source, '  A, X, B: Integer;', '  X: Integer;', [rfReplaceAll]);
  F := Default(TQuickFix);
  F.Kind := qfRemoveVar;
  F.Line := 11;
  F.Identifier := 'X';
  Assert.IsTrue(PlanQuickFixText(Src, F, 0, NewText));
  Assert.AreEqual(0, Pos('X: Integer', NewText), 'the declaration is gone');
  Assert.AreEqual(0, Pos(#13#10'var'#13#10, NewText), 'the empty var section too');
  Assert.IsTrue(Pos('procedure Foo;'#13#10'begin', NewText) > 0);
end;

procedure TQuickFixPreviewTests.DeclareVar_GoesBeforeBeginAndAddsItsUnit;
var
  F: TQuickFix;
  NewText: string;
begin
  var Src := StringReplace(Source, '  A := 1;', '  Q := TStringList.Create;',
    [rfReplaceAll]);
  F := Default(TQuickFix);
  F.Kind := qfDeclareVar;
  F.Line := 13;
  F.Col := 2;
  F.TokenLen := 1;
  F.Identifier := 'Q';
  F.NewText := 'TStringList';
  F.FollowUpUnit := 'System.Classes';
  F.Section := usInterface;
  Assert.IsTrue(PlanQuickFixText(Src, F, 0, NewText));
  Assert.IsTrue(Pos('  Q: TStringList;'#13#10'begin', NewText) > 0, 'declared before begin');
  Assert.IsTrue(Pos('System.Classes', NewText) > 0, 'the unit comes with it');
end;

procedure TQuickFixPreviewTests.GeneratingKinds_HaveNoTextPreview;
var
  F: TQuickFix;
  NewText: string;
begin
  // these build code out of a header elsewhere in the file - their
  // appliers stay, and the preview says so instead of guessing
  for var K in [qfImplStub, qfClassStub, qfImplIntfMethod, qfAlignHeader] do
  begin
    F := Default(TQuickFix);
    F.Kind := K;
    F.Line := 9;
    Assert.IsFalse(PlanQuickFixText(Source, F, 0, NewText), QuickFixKindText(K));
  end;
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
  // The big rings take minutes and the runner prints nothing while a test
  // runs, which looks like a hang (reported with measurements in audit #38:
  // 118 s in one test, 132 s of a 150 s run). Say what is happening - the
  // depth itself must NOT be lowered, see the comment in the deep test.
  if AUnits >= 5000 then
  begin
    Writeln(Format('    [building a %d-unit ring - this takes a while]',
      [AUnits]));
    Flush(Output);
  end;
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

{ TDesignerManagedUsesTests }

function TDesignerManagedUsesTests.FormUnit: string;
begin
  // The DevExpress units of the report are not mentioned anywhere in the
  // source - the DESIGNER puts them there. Vcl.Forms is used (TForm),
  // Vcl.Dialogs only in the implementation.
  Result :=
    'unit Unit1;'#13#10 +
    'interface'#13#10 +
    'uses'#13#10 +
    '  Vcl.Forms, Vcl.Controls, cxGraphics, cxControls, Data.DB;'#13#10 +
    'type'#13#10 +
    '  TForm1 = class(TForm)'#13#10 +
    '  end;'#13#10 +
    'implementation'#13#10 +
    'uses'#13#10 +
    '  Vcl.Dialogs;'#13#10 +
    'procedure Go;'#13#10 +
    'begin'#13#10 +
    '  ShowMessage(''hi'');'#13#10 +
    'end;'#13#10 +
    'end.';
end;

function TDesignerManagedUsesTests.Analyze(const AContent: string;
  const ADesignerUnits: TArray<string>; const AKeepList: string;
  AFormUnverified: Boolean): TArray<TUsesEntryInfo>;
var
  Required: TArray<TDesignerRequiredUnit>;
  R: TDesignerRequiredUnit;
begin
  Required := nil;
  for var U in ADesignerUnits do
  begin
    R.UnitName := U;
    R.Reason := 'cxGrid1: TcxGrid';
    Required := Required + [R];
  end;
  Result := AnalyzeUses(AContent,
    function(const AIdent: string): TArray<string>
    begin
      // only what the .pas really names
      Result := nil;
      if SameText(AIdent, 'TForm') then Result := ['Vcl.Forms']
      else if SameText(AIdent, 'TControl') then Result := ['Vcl.Controls']
      else if SameText(AIdent, 'ShowMessage') then Result := ['Vcl.Dialogs'];
    end,
    function(const AUnit: string): Boolean begin Result := True; end,
    function(const AUnit: string): Boolean begin Result := False; end,
    DesignerRequiredLookup(Required),
    function(const AUnit: string): Boolean
    begin
      Result := MatchesKeepList(AUnit, AKeepList);
    end,
    AFormUnverified);
end;

function TDesignerManagedUsesTests.VerdictOf(
  const AEntries: TArray<TUsesEntryInfo>; const AUnit: string): TUsesVerdict;
begin
  for var E in AEntries do
    if SameText(E.UnitName, AUnit) then Exit(E.Verdict);
  Assert.Fail('no entry for ' + AUnit);
  Result := uvUnknown;
end;

function TDesignerManagedUsesTests.ReasonOf(
  const AEntries: TArray<TUsesEntryInfo>; const AUnit: string): string;
begin
  Result := '';
  for var E in AEntries do
    if SameText(E.UnitName, AUnit) then Exit(E.Reason);
end;

procedure TDesignerManagedUsesTests.ADesignerRequiredUnitIsKeptAndNamesTheComponent;
var
  E: TArray<TUsesEntryInfo>;
begin
  // The reported case: the designer needs cxGraphics, the source never
  // names it.
  E := Analyze(FormUnit, ['cxGraphics']);
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, VerdictOf(E, 'cxGraphics'));
  // and the row has to SAY which component does it, or the user cannot
  // judge the answer
  Assert.AreEqual('cxGrid1: TcxGrid', ReasonOf(E, 'cxGraphics'));
  // cxControls is NOT in the designer answer, so it stays a candidate
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'cxControls'));
  // REGRESSION: without a designer lookup the old behaviour is unchanged -
  // that is what every non-form unit keeps doing.
  E := Analyze(FormUnit, []);
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'cxGraphics'));
  Assert.AreEqual<TUsesVerdict>(uvUsed, VerdictOf(E, 'Vcl.Forms'));
end;

procedure TDesignerManagedUsesTests.ADesignerRequiredUnitIsNotMovedDownEither;
const
  Src =
    'unit Unit1;'#13#10 +
    'interface'#13#10 +
    'uses'#13#10 +
    '  Data.DB;'#13#10 +
    'implementation'#13#10 +
    'procedure Go(D: TDataSet);'#13#10 +
    'begin'#13#10 +
    'end;'#13#10 +
    'end.';
var
  E: TArray<TUsesEntryInfo>;
begin
  // Data.DB is only used in the implementation, so the analysis alone would
  // offer to move it down - but the designer always writes into the
  // INTERFACE uses, so moving it starts the same tug-of-war as removing it.
  E := AnalyzeUses(Src,
    function(const AIdent: string): TArray<string>
    begin
      Result := nil;
      if SameText(AIdent, 'TDataSet') then Result := ['Data.DB'];
    end,
    function(const AUnit: string): Boolean begin Result := True; end,
    function(const AUnit: string): Boolean begin Result := False; end,
    DesignerRequiredLookup([]));
  Assert.AreEqual<TUsesVerdict>(uvMovable, VerdictOf(E, 'Data.DB'),
    'without the designer it is movable');

  var Req: TDesignerRequiredUnit;
  Req.UnitName := 'Data.DB';
  Req.Reason := 'DBGrid1: TDBGrid';
  E := AnalyzeUses(Src,
    function(const AIdent: string): TArray<string>
    begin
      Result := nil;
      if SameText(AIdent, 'TDataSet') then Result := ['Data.DB'];
    end,
    function(const AUnit: string): Boolean begin Result := True; end,
    function(const AUnit: string): Boolean begin Result := False; end,
    DesignerRequiredLookup([Req]));
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, VerdictOf(E, 'Data.DB'),
    'the designer writes it into the INTERFACE uses - do not move it');
end;

procedure TDesignerManagedUsesTests.UnitScopeNamesMatchBothWays;
var
  E: TArray<TUsesEntryInfo>;
begin
  // The designer reports the RTTI name ('Vcl.Forms'), older code writes the
  // short one ('Forms'). Matched both ways, so it errs towards KEEPING.
  Assert.IsTrue(SameUnitIgnoringScope('Forms', 'Vcl.Forms'));
  Assert.IsTrue(SameUnitIgnoringScope('Vcl.Forms', 'Forms'));
  Assert.IsTrue(SameUnitIgnoringScope('vcl.forms', 'Vcl.Forms'), 'case');
  Assert.IsFalse(SameUnitIgnoringScope('Vcl.Forms', 'Vcl.FormsX'));
  Assert.IsFalse(SameUnitIgnoringScope('Forms', 'MyForms'),
    'a name that merely ENDS with it is another unit');
  // and through the analysis
  E := Analyze(StringReplace(FormUnit, 'Vcl.Controls, cxGraphics',
    'Controls, cxGraphics', []), ['Vcl.Controls']);
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, VerdictOf(E, 'Controls'));
end;

procedure TDesignerManagedUsesTests.TheKeepListTakesMasks;
var
  E: TArray<TUsesEntryInfo>;
begin
  // What the designer cannot report (a unit some other IDE expert writes on
  // save) is what this list is for.
  Assert.IsTrue(MatchesKeepList('dxSkinsCore', 'dxSkin*;MyCompany.*'));
  Assert.IsTrue(MatchesKeepList('MyCompany.Utils', 'dxSkin*;MyCompany.*'));
  Assert.IsFalse(MatchesKeepList('MyCompanyX', 'dxSkin*;MyCompany.*'),
    'MyCompany.* must not swallow MyCompanyX');
  Assert.IsFalse(MatchesKeepList('cxGraphics', ''), 'an empty list keeps nothing');
  Assert.IsTrue(MatchesKeepList('CXGRAPHICS', 'cxgraphics'), 'case-insensitive');
  Assert.IsTrue(MatchesKeepList('dxSkinsCore', ' dxSkin* ; x '), 'blanks around a mask');
  E := Analyze(FormUnit, [], 'cx*');
  Assert.AreEqual<TUsesVerdict>(uvKeptByUser, VerdictOf(E, 'cxControls'));
  // the designer answer WINS over the keep list: its reason is the useful one
  E := Analyze(FormUnit, ['cxGraphics'], 'cx*');
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, VerdictOf(E, 'cxGraphics'));
  Assert.AreEqual<TUsesVerdict>(uvKeptByUser, VerdictOf(E, 'cxControls'));
end;

procedure TDesignerManagedUsesTests.AFormWhoseDesignerCannotBeAskedIsUnverified;
var
  E: TArray<TUsesEntryInfo>;
begin
  // Outside the IDE (MCP on a closed form, standalone, or a form that failed
  // to load) nothing can say what the designer would re-add - and nothing
  // re-adds it on a CI build either, so a removal there is a compile error
  // at best and a form that streams an unregistered class at worst.
  E := Analyze(FormUnit, [], '', True);
  Assert.AreEqual<TUsesVerdict>(uvUnverified, VerdictOf(E, 'cxGraphics'));
  Assert.AreEqual<TUsesVerdict>(uvUnverified, VerdictOf(E, 'Data.DB'));
  Assert.AreEqual<TUsesVerdict>(uvUsed, VerdictOf(E, 'Vcl.Forms'),
    'a unit that IS used stays used - this is not about usage');
  // the same content WITHOUT a form file keeps the old verdict
  E := Analyze(FormUnit, [], '', False);
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'cxGraphics'));
end;

procedure TDesignerManagedUsesTests.AnIncompleteAnswerStillProtectsWhatItDidReport;
var
  E: TArray<TUsesEntryInfo>;
begin
  // A selection editor raised while we asked: what it DID report is still
  // true, the rest is unknown. So the reported unit stays protected and
  // everything else becomes unverified - never the other way round.
  E := Analyze(FormUnit, ['cxGraphics'], '', True);
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, VerdictOf(E, 'cxGraphics'));
  Assert.AreEqual<TUsesVerdict>(uvUnverified, VerdictOf(E, 'cxControls'));
end;

procedure TDesignerManagedUsesTests.OptingIntoAnUnverifiedRowFollowsTheText;
var
  Entry: TUsesEntryInfo;
begin
  // ONE rule for the dialog's Apply and the MCP tool's include_unverified.
  Entry := Default(TUsesEntryInfo);
  Entry.Verdict := uvUnverified;
  Entry.UsageCount := 0;
  Assert.AreEqual<TUsesVerdict>(uvUnused, ResolveUnverified(Entry));
  Entry.UsageCount := 3;
  Assert.AreEqual<TUsesVerdict>(uvMovable, ResolveUnverified(Entry));
  // every other verdict is passed through untouched
  Entry.Verdict := uvIdeManaged;
  Assert.AreEqual<TUsesVerdict>(uvIdeManaged, ResolveUnverified(Entry));
  Entry.Verdict := uvUsed;
  Assert.AreEqual<TUsesVerdict>(uvUsed, ResolveUnverified(Entry));
end;

procedure TDeclarationAnchorTests.AForeignAnswerIsJudgedTheSameByBothPasses;
begin
  // The reported cold run: 0 aborted requests, an anchor derived from 300
  // agreeing answers - and three rows answering UGlobalRomConfig.pas:1586
  // (the record field GlobalConfig.Formulare.BTB), which the warm run does
  // not list at all. A clear answer naming another declaration is EVIDENCE,
  // so it is dropped - exactly what the second attempt did with two more
  // rows of that shape in the same run.
  Assert.AreEqual<TForeignAnswerVerdict>(favDrop, ForeignAnswerVerdict(True, 0),
    'an anchor and a healthy session: the answer decides');
  // A DERIVED anchor is still an anchor - that is what made the cold and the
  // warm run disagree by exactly three rows.
  Assert.AreEqual<TForeignAnswerVerdict>(favDrop, ForeignAnswerVerdict(True, 0),
    'the derived anchor counts too');
  // Issue #13: an aborted request means the answers come from a degraded
  // session, so nothing is removed on their word.
  Assert.AreEqual<TForeignAnswerVerdict>(favKeepMarked, ForeignAnswerVerdict(True, 1),
    'one aborted request is enough to stop dropping');
  Assert.AreEqual<TForeignAnswerVerdict>(favKeepMarked, ForeignAnswerVerdict(True, 28));
  // No anchor at all: there is nothing to measure "elsewhere" against, and
  // dropping here is what turned 327 references into 29 (first report).
  Assert.AreEqual<TForeignAnswerVerdict>(favKeepMarked, ForeignAnswerVerdict(False, 0),
    'without an anchor nothing may be dropped');
  Assert.AreEqual<TForeignAnswerVerdict>(favKeepMarked, ForeignAnswerVerdict(False, 5));
end;

{ TUsesClauseParsingTests }

function TUsesClauseParsingTests.Analyze(const AClause: string;
  const AInactive: TArray<Integer>): TArray<TUsesEntryInfo>;
var
  Inact: TLineInactive;
begin
  // A whole unit around the clause, so the line numbers are real: the
  // clause starts on line 4 (0-based).
  var Src := 'unit U;'#13#10 + 'interface'#13#10 + ''#13#10 + AClause +
    #13#10 + 'implementation'#13#10 + 'end.';
  Inact := nil;
  if System.Length(AInactive) > 0 then
    Inact :=
      function(ALine: Integer): Boolean
      begin
        Result := False;
        for var L in AInactive do
          if L = ALine then Exit(True);
      end;
  Result := AnalyzeUses(Src,
    function(const AIdent: string): TArray<string> begin Result := nil; end,
    function(const AUnit: string): Boolean begin Result := True; end,
    function(const AUnit: string): Boolean begin Result := False; end,
    nil, nil, False, Inact);
end;

function TUsesClauseParsingTests.Names(
  const AEntries: TArray<TUsesEntryInfo>): string;
begin
  Result := '';
  for var E in AEntries do
  begin
    if Result <> '' then Result := Result + ',';
    Result := Result + E.UnitName;
  end;
end;

function TUsesClauseParsingTests.VerdictOf(
  const AEntries: TArray<TUsesEntryInfo>; const AUnit: string): TUsesVerdict;
begin
  for var E in AEntries do
    if SameText(E.UnitName, AUnit) then Exit(E.Verdict);
  Assert.Fail('no entry for ' + AUnit);
  Result := uvUsed;
end;

function TUsesClauseParsingTests.ReasonOf(
  const AEntries: TArray<TUsesEntryInfo>; const AUnit: string): string;
begin
  Result := '';
  for var E in AEntries do
    if SameText(E.UnitName, AUnit) then Exit(E.Reason);
end;

procedure TUsesClauseParsingTests.ADirectiveIsNotPartOfTheUnitName;
begin
  // The reported shape, and the one this repository's own options frame has.
  Assert.AreEqual('A,B,C,D', Names(Analyze(
    'uses'#13#10 +
    '  A, B, C {$IFNDEF STANDALONE_BUILD},'#13#10 +
    '  D{$ENDIF};')),
    'the directive must not stick to C or to D');
  // ... and a clause whose branches span several lines, with a two-line
  // brace comment in the middle
  Assert.AreEqual('A,B,C,D', Names(Analyze(
    'uses'#13#10 +
    '  A,'#13#10 +
    '  {$IFDEF DEBUG}'#13#10 +
    '  B,'#13#10 +
    '  {$ELSE}'#13#10 +
    '  C,'#13#10 +
    '  {$ENDIF}'#13#10 +
    '  { a comment'#13#10 +
    '    over two lines }'#13#10 +
    '  D;')));
end;

procedure TUsesClauseParsingTests.ASemicolonInACommentDoesNotEndTheClause;
begin
  // Measured live before the fix: analyze_uses answered ONE entry named
  // "System.SysUtils { was: System.Math" and B and C were missing entirely.
  Assert.AreEqual('A,B,C', Names(Analyze(
    'uses A { was: X; }, B (* old; *), C;')),
    'a '';'' inside a comment is not the terminator');
end;

procedure TUsesClauseParsingTests.AnInPathTailIsNotPartOfTheName;
var
  E: TArray<TUsesEntryInfo>;
begin
  E := Analyze('uses A in ''..\src\A.pas'', B in ''B.pas'';');
  Assert.AreEqual('A,B', Names(E), 'the path is not part of the name');
  // ... but such an entry belongs to a PROJECT file, and removing it takes
  // the unit out of the project - so it is reported, never offered.
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'A'));
  Assert.Contains(ReasonOf(E, 'A'), 'project file');
end;

procedure TUsesClauseParsingTests.APlainClauseIsUnchangedAndStillJudged;
var
  E: TArray<TUsesEntryInfo>;
begin
  // REGRESSION: an ordinary clause must behave exactly as before, and a
  // trailing // comment must not block the analysis either.
  E := Analyze('uses A, // the one we need'#13#10 + '  B;');
  Assert.AreEqual('A,B', Names(E));
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'A'),
    'nothing uses it in this fixture, and a // tail does not stop the analysis');
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'B'));
end;

procedure TUsesClauseParsingTests.AnUnknownVerdictSaysWhy;
var
  E: TArray<TUsesEntryInfo>;
begin
  // The row used to read "not analysable" with no reason given, leaving
  // three different causes indistinguishable.
  E := Analyze('uses A, B {$IFDEF X}, C{$ENDIF};');
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'A'));
  Assert.Contains(ReasonOf(E, 'A'), 'directives');
  // no indexed source is the other cause, and it must name itself
  E := AnalyzeUses('unit U;'#13#10 + 'interface'#13#10 + 'uses A;'#13#10 +
    'implementation'#13#10 + 'end.',
    function(const AIdent: string): TArray<string> begin Result := nil; end,
    function(const AUnit: string): Boolean begin Result := False; end,
    function(const AUnit: string): Boolean begin Result := False; end);
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'A'));
  Assert.Contains(ReasonOf(E, 'A'), 'no indexed source');
end;

procedure TUsesClauseParsingTests.AnInactiveBranchIsNeverJudgedButItsNeighbourIs;
var
  E: TArray<TUsesEntryInfo>;
begin
  // The LSP session reports which lines the compiler does not see
  // (TLspClient.IsLineInactive, already used by remove-with). With that, a
  // conditional clause no longer has to stay unanalysed as a whole:
  // 0-based lines of the whole fixture unit:
  //   3: uses
  //   4:   A, {$IFDEF X}
  //   5:   B,            <- inactive in this configuration
  //   6:   {$ELSE}
  //   7:   C,            <- active
  //   8:   {$ENDIF} D;
  var Clause :=
    'uses'#13#10 +
    '  A, {$IFDEF X}'#13#10 +
    '  B,'#13#10 +
    '  {$ELSE}'#13#10 +
    '  C,'#13#10 +
    '  {$ENDIF} D;';
  E := Analyze(Clause, [5]);
  Assert.AreEqual('A,B,C,D', Names(E), 'every entry is still reported');
  // B is not compiled here, so our usage analysis has no evidence about it
  // in the configuration where it IS compiled - never judged.
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'B'));
  Assert.Contains(ReasonOf(E, 'B'), 'inactive in this configuration');
  // the active ones are judged normally, which is the whole point
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'A'));
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'C'));
  Assert.AreEqual<TUsesVerdict>(uvUnused, VerdictOf(E, 'D'));
  // WITHOUT that information nothing in a conditional clause is judged -
  // the behaviour before this round, and what the standalone keeps doing.
  E := Analyze(Clause);
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'A'));
  Assert.AreEqual<TUsesVerdict>(uvUnknown, VerdictOf(E, 'D'));
  Assert.Contains(ReasonOf(E, 'A'), 'directives');
end;

{ TSelfProtectionTests }

procedure TSelfProtectionTests.AWriteFromAWorkerThreadIsRecordedAndCanBeRefused;
var
  Before: Integer;
  Raised, RanOffMain: Boolean;
  T: TThread;
begin
  // the test itself runs on the main thread, so nothing is recorded there
  Before := MainThreadViolations;
  Assert.IsTrue(OnMainThread, 'the suite runs on the main thread');
  RequireMainThread('probe_from_main');
  Assert.AreEqual(Before, MainThreadViolations,
    'a main-thread call is not a violation');

  // ... and a worker thread is one, by name
  T := TThread.CreateAnonymousThread(
    procedure
    begin
      RequireMainThread('probe_from_worker');
    end);
  T.FreeOnTerminate := False;
  try
    T.Start;
    T.WaitFor;
  finally
    T.Free;
  end;
  Assert.AreEqual(Before + 1, MainThreadViolations, 'the violation is counted');
  Assert.AreEqual('probe_from_worker', LastOffMainThreadCall,
    'and the call that did it is named');

  // With the strict switch it is fatal instead of merely counted. It is OFF
  // by default on purpose: the three known offenders would turn into hard
  // errors before they are fixed.
  Raised := False;
  RanOffMain := False;
  StrictMainThread := True;
  try
    T := TThread.CreateAnonymousThread(
      procedure
      begin
        try
          RequireMainThread('probe_strict');
          RanOffMain := True;      // must NOT be reached
        except
          on E: EOffMainThread do Raised := True;
        end;
      end);
    T.FreeOnTerminate := False;
    try
      T.Start;
      T.WaitFor;
    finally
      T.Free;
    end;
  finally
    StrictMainThread := False;
  end;
  Assert.IsTrue(Raised, 'strict mode raises EOffMainThread');
  Assert.IsFalse(RanOffMain, 'and the write never happens');
end;

procedure TSelfProtectionTests.AnImpossibleUnitNameIsReportedAsADefect;
begin
  // the two names analyze_uses really answered, verbatim
  Assert.IsTrue(ImplausibleUnitName('Expert.PluginSettings {$IFNDEF STANDALONE_BUILD}'));
  Assert.IsTrue(ImplausibleUnitName('Expert.BlameGutter{$ENDIF}'));
  Assert.IsTrue(ImplausibleUnitName('System.SysUtils { was: System.Math'));
  Assert.IsTrue(ImplausibleUnitName('A in ''..\src\A.pas'''));
  Assert.IsTrue(ImplausibleUnitName(''), 'an empty name is not a unit either');
  Assert.IsTrue(ImplausibleUnitName('Vcl.'), 'a trailing dot has no segment');
  Assert.IsTrue(ImplausibleUnitName('.Vcl.Forms'));
  Assert.IsTrue(ImplausibleUnitName('2Fast'), 'a name cannot start with a digit');
  // ... and everything a real unit name looks like
  Assert.IsFalse(ImplausibleUnitName('Vcl.Forms'));
  Assert.IsFalse(ImplausibleUnitName('Winapi.Windows'));
  Assert.IsFalse(ImplausibleUnitName('cxGraphics'));
  Assert.IsFalse(ImplausibleUnitName('U2'));
  Assert.IsFalse(ImplausibleUnitName('_Private.Unit_2'));
  Assert.IsFalse(ImplausibleUnitName('mormot.core.interfaces'));

  // the sentence a tool puts into its answer
  Assert.AreEqual('', UnitNameSelfCheck(['Vcl.Forms', 'System.Classes']),
    'a clean answer carries no selfCheck at all');
  var S := UnitNameSelfCheck(['Vcl.Forms', 'Expert.BlameGutter{$ENDIF}']);
  Assert.IsTrue(S <> '', 'one bad name is enough');
  Assert.Contains(S, 'Expert.BlameGutter{$ENDIF}', 'it quotes the name');
  Assert.Contains(S, 'report', 'and says it is a defect in the plugin');
end;

procedure TSelfProtectionTests.APreviewTokenBelongsToItsArguments;

  function Args(const APairs: array of string): TJSONObject;
  begin
    Result := TJSONObject.Create;
    var I := 0;
    while I < Length(APairs) - 1 do
    begin
      Result.AddPair(APairs[I], APairs[I + 1]);
      Inc(I, 2);
    end;
  end;

var
  Problem: string;
begin
  var Reader: TFunc<string, string> :=
    function(AFile: string): string
    begin
      Result := 'unchanged content';   // the buffer never moves in this test
    end;

  // the fingerprint itself: order does not matter, apply/token/instance do
  // not belong to it
  var A1 := Args(['file', 'U.pas', 'unit', 'Foo']);
  var A2 := Args(['unit', 'Foo', 'file', 'U.pas']);
  var A3 := Args(['file', 'U.pas', 'unit', 'Foo', 'apply', 'true',
    'token', 'x', 'instance', '4711']);
  var A4 := Args(['file', 'U.pas', 'unit', 'Bar']);
  try
    Assert.AreEqual(PreviewArgsFingerprint(A1), PreviewArgsFingerprint(A2),
      'the order the client sends the pairs in must not matter');
    Assert.AreEqual(PreviewArgsFingerprint(A1), PreviewArgsFingerprint(A3),
      'apply / token / instance say nothing about WHAT is changed');
    Assert.AreNotEqual(PreviewArgsFingerprint(A1), PreviewArgsFingerprint(A4),
      'another unit is another change');

    // the token of a preview with A1 must not apply A4
    var Token := NewPreviewToken('add_unit', ['U.pas'], ['unchanged content'], A1);
    Assert.IsTrue(CheckPreviewToken(Token, 'add_unit', Reader, A1, Problem),
      'the arguments it was previewed with: ' + Problem);
    Assert.IsFalse(CheckPreviewToken(Token, 'add_unit', Reader, A4, Problem),
      'OTHER arguments must be refused');
    Assert.Contains(Problem, 'arguments', 'and the reason says why: ' + Problem);
    // the old call shape (no arguments) still works
    Assert.IsTrue(CheckPreviewToken(Token, 'add_unit', Reader, Problem));
  finally
    A1.Free; A2.Free; A3.Free; A4.Free;
  end;
end;

procedure TSelfProtectionTests.AnArgumentTheToolDoesNotKnowIsReported;
var
  Known: TArray<string>;
begin
  // the schema is the contract, and it is read from the SHIPPED list
  Known := KnownToolArguments('buffer_read');
  Assert.IsTrue(Length(Known) > 0, 'buffer_read declares arguments');
  Assert.IsTrue(MatchText('start_line', Known), 'start_line is one of them');
  Assert.IsTrue(MatchText('instance', Known),
    'and instance is accepted everywhere - the bridge may add it');

  // the mistake that found this: buffer_read takes start_line / end_line
  var Unknown := UnknownToolArguments('buffer_read',
    ['file', 'from_line', 'to_line']);
  Assert.AreEqual(2, Integer(Length(Unknown)), 'both wrong names are named');
  Assert.AreEqual('from_line', Unknown[0]);
  var Note := UnknownArgumentNote('buffer_read', ['file', 'from_line']);
  Assert.Contains(Note, 'from_line', 'the note quotes what was dropped');
  Assert.Contains(Note, 'start_line', 'and says what the tool does take');

  // a correct call must stay silent
  Assert.AreEqual(0, Integer(Length(UnknownToolArguments('buffer_read',
    ['file', 'start_line', 'end_line', 'instance']))));
  Assert.AreEqual('', UnknownArgumentNote('buffer_read',
    ['file', 'start_line']));
  // case is irrelevant - JSON keys are not Pascal, but the schema is ours
  Assert.AreEqual(0, Integer(Length(UnknownToolArguments('buffer_read',
    ['FILE', 'Start_Line']))));

  // the bridge's own tools are in the same list, so they are checked too
  Assert.AreEqual(1, Integer(Length(UnknownToolArguments('ide_instances',
    ['whatever']))), 'ide_instances is described here as well');
  // a tool the list does NOT describe (an IDE newer than this unit) has
  // nothing to measure against, so it reports nothing rather than
  // everything
  Assert.AreEqual(0, Integer(Length(UnknownToolArguments('no_such_tool',
    ['x', 'y']))));
  Assert.AreEqual('', UnknownArgumentNote('no_such_tool', ['x']));
end;

{ TRemoveWithSafetyTests }

procedure TRemoveWithSafetyTests.AValueTypeIsRecognisedInItsDeclaration;
begin
  // the reported shape: "with R.Inner do Count := 0" with Inner a record
  Assert.IsTrue(DeclaredTypeIsValueType('  TInner = record', ''),
    'a record is a value type');
  Assert.IsTrue(DeclaredTypeIsValueType('  TInner = packed record', ''),
    'a packed record too');
  Assert.IsTrue(DeclaredTypeIsValueType('  TOld = object', ''),
    'an old-style object too');
  Assert.IsTrue(DeclaredTypeIsValueType('  TInner =', '    record'),
    'the keyword may stand on the next line');
  // ... and everything that is a REFERENCE must stay rewritable
  Assert.IsFalse(DeclaredTypeIsValueType('  TFoo = class(TObject)', ''));
  Assert.IsFalse(DeclaredTypeIsValueType('  TFoo = class', '  private'));
  Assert.IsFalse(DeclaredTypeIsValueType('  IFoo = interface', ''));
  Assert.IsFalse(DeclaredTypeIsValueType('  IFoo = dispinterface', ''));
  // 'procedure of object' must not be read as 'object'
  Assert.IsFalse(DeclaredTypeIsValueType('  TEvent = procedure of object;', ''),
    'the FIRST word after the = decides');
  Assert.IsFalse(DeclaredTypeIsValueType('  TFoo = class helper for TBar', ''));
  Assert.IsFalse(DeclaredTypeIsValueType('no equals sign here', ''));
end;

procedure TRemoveWithSafetyTests.AWithInsideAnAnonymousMethodIsRecognised;
const
  // 1 unit U; 2 interface 3 implementation
  Src =
    'unit U;'#13#10 +                               // 1
    'interface'#13#10 +                             // 2
    'implementation'#13#10 +                        // 3
    'procedure Outer;'#13#10 +                      // 4
    'begin'#13#10 +                                 // 5
    '  with FFoo do'#13#10 +                        // 6  <- NOT anonymous
    '    Bar;'#13#10 +                              // 7
    '  TParallel.For(0, 9, procedure(I: Integer)'#13#10 +   // 8
    '    begin'#13#10 +                             // 9
    '      with TFoo.Create do'#13#10 +             // 10 <- anonymous
    '      try'#13#10 +                             // 11
    '        Run;'#13#10 +                          // 12
    '      finally'#13#10 +                         // 13
    '        Free;'#13#10 +                         // 14
    '      end;'#13#10 +                            // 15
    '    end);'#13#10 +                             // 16
    '  with FBaz do'#13#10 +                        // 17 <- after it: NOT
    '    Qux;'#13#10 +                              // 18
    'end;'#13#10 +                                  // 19
    'end.';
var
  Lines: TArray<string>;
begin
  Lines := Src.Replace(#13#10, #10).Split([#10]);
  Assert.IsFalse(InsideAnonymousMethod(Lines, 5, 6),
    'a with directly in the method body is not inside an anonymous one');
  Assert.IsTrue(InsideAnonymousMethod(Lines, 5, 10),
    'the with inside the anonymous method must be recognised');
  Assert.IsFalse(InsideAnonymousMethod(Lines, 5, 17),
    'after the anonymous method ends, we are back in the outer body');
end;

procedure TRemoveWithSafetyTests.AnEditIsVerifiedAgainstTheCurrentText;
const
  Content = 'unit U;'#13#10'begin'#13#10'  with A do B;'#13#10'end.';
begin
  // line 3 starts with two blanks, so the 'with' is at column 3
  Assert.IsTrue(TextMatchesAt(Content, 3, 3, 'with A do B;'),
    'the text is still where the edit expects it');
  Assert.IsFalse(TextMatchesAt(Content, 3, 3, 'with X do B;'),
    'a changed line must not pass');
  Assert.IsFalse(TextMatchesAt(Content, 2, 3, 'with A do B;'),
    'the same text one line up is not a match');
  Assert.IsFalse(TextMatchesAt(Content, 99, 1, 'anything'),
    'a line beyond the end is no match');
  // an old text spanning lines: the line breaks need not agree
  Assert.IsTrue(TextMatchesAt(Content, 2, 1, 'begin'#10'  with A do B;'),
    'CR is ignored on both sides');
  // an insertion has nothing to verify - which is exactly why the caller
  // must check the other edits of the file
  Assert.IsTrue(TextMatchesAt(Content, 1, 1, ''),
    'an empty old text always matches');
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
  TDUnitX.RegisterTestFixture(TQuickFixPreviewTests);
  TDUnitX.RegisterTestFixture(TDfmGluedDeclarationTests);
  TDUnitX.RegisterTestFixture(TDeclarationAnchorTests);
  TDUnitX.RegisterTestFixture(TSignatureQualifierTests);
  TDUnitX.RegisterTestFixture(TInterfaceDeclLineTests);
  TDUnitX.RegisterTestFixture(TDesignerManagedUsesTests);
  TDUnitX.RegisterTestFixture(TUsesClauseParsingTests);
  TDUnitX.RegisterTestFixture(TSelfProtectionTests);
  TDUnitX.RegisterTestFixture(TRemoveWithSafetyTests);

end.
