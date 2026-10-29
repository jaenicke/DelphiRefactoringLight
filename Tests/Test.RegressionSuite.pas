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
  // System.SysUtils for TBytes - the line-ending fixture names it in its
  // own helper declaration.
  // Expert.StatementRefactor for TDeclareRequest / TDeclarePlace, which the
  // constant fixture names in its own request builder.
  DUnitX.TestFramework, System.SysUtils, Expert.UsesCleanup,
  Expert.StatementRefactor,
  Expert.EditorHelperIntf;

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
    [Test] procedure VcsOutput_MixedUtf8AndAnsiLastLine;
    [Test] procedure VcsOutput_Utf8RespectsByteCount;
    [Test] procedure WorkerLatch_WaitsAndRefusesAfterShutdown;
    /// <summary>Audit #38, L1g: the worker COUNT drops in the closure's
    ///  finally - before the closure epilogue releases what it captured and
    ///  before the thread itself has ended. A shutdown that watches only
    ///  the count therefore lets the BPL unload while plugin code is still
    ///  running on that thread; the Sleep(50) it used instead of a join was
    ///  a timing allowance, not a guarantee.</summary>
    [Test] procedure WorkerLatch_WaitsForTheThreadNotJustTheCount;
    /// <summary>...and the two subsystems the latch does NOT start - the
    ///  live quick-fix checker and the blame reader - note themselves, so
    ///  the unload can wait for their threads too instead of only for
    ///  their counters. What it still cannot finish is RECORDED, because
    ///  a finalization has no other channel than the marker file.</summary>
    [Test] procedure WorkerLatch_JoinsAWorkerItDidNotStart;
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
    /// <summary>Found live on 2026-10-04: the age probe asked the PROCESS
    ///  for its image, and after an install that image is the file the
    ///  install renamed aside - older than the process, of course. So the
    ///  row said "the installed exe is still the old one, run install.cmd"
    ///  in exactly the case where the install had worked and only a new
    ///  session was missing; the user followed it twice.</summary>
    [Test] procedure AnAsideExeNamesTheInstallThatReplacedIt;
    /// <summary>User, 2026-10-17: "Kannst du der MCP Bridge nicht eine
    ///  Neustartfunktion verpassen, wenn die neue Exe da liegt, aber noch
    ///  die alte umbenannte Version geladen ist? inklusive loeschen der
    ///  alten Version." The decision and the cleanup rule; the hand-over
    ///  itself needs two processes and real pipes and is measured by
    ///  scratchpad relaunch\probe.ps1.</summary>
    [Test] procedure AReplacedBridgeHandsOverAndTidiesUp;
    [Test] procedure BridgeVersion_ReachesTheServerOverARealPipe;
    /// <summary>Audit #22, H4: Stop waited 5 s for the handlers and then
    ///  returned as if all was well, so the owner freed the server (and the
    ///  package unloaded) under a handler that ignores the stop event. Stop
    ///  now says so, and Destroy waits for the last handler.</summary>
    [Test] procedure Stop_ReportsAHandlerThatOutlivesTheDeadline;
    /// <summary>Issue #22, the PR #31 follow-up: a handler in a wait that
    ///  ignores the stop event made the IDE give up at the first deadline,
    ///  leak the server and pin the package - although ending what it waits
    ///  for (DelphiLSP) would have let it leave a moment later. The wake
    ///  action runs exactly once, only when needed, and the stop is clean
    ///  after it.</summary>
    [Test] procedure Stop_WakesAStragglerAndThenStopsCleanly;
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
    /// <summary>A reviewing model with nothing but the tool list said it
    ///  could use almost everything - and then: the shared apply text
    ///  promises "plus a token" while only half the writing tools hand one
    ///  out, so for the other half it could not tell whether an apply is
    ///  protected against changes at all. It cannot be told from the
    ///  outside, so each tool has to say it.</summary>
    [Test] procedure EveryWritingToolSaysWhatProtectsItsApply;
    /// <summary>Same review: six parameters had no description at all, so
    ///  their meaning and their default could only be found by trying. The
    ///  sweep found 46 of them - including the apply of convert_properties,
    ///  which is why this rule is a test and not a round of fixes.</summary>
    [Test] procedure EveryToolParameterIsDescribed;
    /// <summary>The "guide" tool answers what this server can do - the one
    ///  text a fresh session can read before it knows the toolbox, and the
    ///  only one that costs nothing until it is asked for. A guide that
    ///  misses a tool, or names one that is gone, is worse than none, so
    ///  the two lists are held against each other in all three directions:
    ///  every tool in exactly one section, every section name a real tool,
    ///  and every name present in the prose that section shows.</summary>
    [Test] procedure TheGuideAndTheToolListAgree;
    /// <summary>The initialize instructions are the ONLY text every
    ///  session with this server registered pays for, whether it calls a
    ///  tool or not (measured 2026-10-19: 946 characters, ~235 tokens,
    ///  against ~680 for the 61 names and ~13 800 for all schemas). They
    ///  were cut to that size in 1.11.1 and must not grow back into a
    ///  catalogue - that is what the guide tool is for, which costs
    ///  nothing until it is called.</summary>
    [Test] procedure TheAlwaysLoadedInstructionsStayShort;
    /// <summary>Reported by a reviewing session on 2.0.6: an empty lsp_*
    ///  answer always blamed a precompiled .dcu as the "usual cause" - and
    ///  in the case it had, hover ANSWERED and only the definition was
    ///  unresolved, so the hint sent it looking in the wrong place and
    ///  never named find_original_symbol, which does answer there. Plus
    ///  2.0.6's own regression: the tool's sentence and the generic text
    ///  both opened with "DelphiLSP returned nothing".</summary>
    [Test] procedure AnEmptyAnswerNamesTheRightCause;
    /// <summary>Opening or closing a project throws away unsaved work when
    ///  nobody is asked - and the IDE's own question cannot be answered
    ///  through a pipe, it only blocks the call. Both tools therefore have
    ///  to OFFER the decision and say that they refuse without it, and the
    ///  file list has to be a reviewable change like every other write.
    ///  </summary>
    [Test] procedure AProjectActionSaysWhatHappensToUnsavedWork;
    /// <summary>Measured 2026-10-20 through the session log: 103 .dcu in
    ///  the directory the .delphilsp.json names made DelphiLSP load 60
    ///  project units FROM those, and for their symbols it answers a
    ///  definition without a file - so go to definition, find references
    ///  and rename come back empty with nothing saying why. They were
    ///  OLDER than the sources and were used anyway, so only an empty
    ///  directory helps. This pins the three pure steps: find the
    ///  directory, find what shadows a source, move it aside.</summary>
    [Test] procedure TheLspDcuShadowIsFoundAndMovable;
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
    /// <summary>User request 2026-10-20: a session cannot click the
    ///  options page, so the settings are readable and writable by NAME -
    ///  which only works while the descriptor, the getters, the setters
    ///  and Save agree about every one of them. Two hand-kept lists are
    ///  what made the legacy migration drop five values (audit #37,
    ///  L4e).</summary>
    [Test] procedure EverySettingIsDescribedAndWiredUp;
    /// <summary>...and "for this IDE instance only" has to mean it: a
    ///  Save triggered elsewhere (the rename dialog stores its scope, the
    ///  uses cleanup its keep list) must not make a temporary change
    ///  permanent behind the user's back.</summary>
    [Test] procedure ASessionOnlyValueIsNeverPersisted;
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
    /// <summary>Issue #22, his follow-up: signature help walked back over
    ///  an ASCII-only character set, so a call name with an umlaut
    ///  answered its own TAIL - "Groesse(" with the real umlaut asked
    ///  about "e". The shared, Unicode-aware reader is one function now,
    ///  and the declaration end it needs next door is too.</summary>
    [Test] procedure IdentifierBefore_IsUnicodeAware;
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
    /// <summary>Issue #49: renaming a parameter from inside the body
    ///  changed the implementation header and left the declaration in the
    ///  class alone - and Delphi requires the two names to be identical
    ///  (measured with dcc32 37.0: E2037 plus an undeclared identifier for
    ///  every use in the body), so the result did not compile.</summary>
    [Test] procedure AParameterIsDeclaredTwiceAndBothMustFollow;
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
    /// <summary>Found the same way (2026-10-03): "[ '{...}']" - a blank
    ///  after the bracket - compiles, and the compiler takes it as the
    ///  interface's GUID (assigning the interface to a TGUID works, E2232
    ///  otherwise). The scan looked for the exact text "['{" and listed
    ///  the interface as "(no GUID)".</summary>
    [Test] procedure BlanksInsideTheGuidBrackets_AreStillAGuid;
    /// <summary>The fix the check never had (user request, 2026-10-04): a
    ///  declaration gets a FRESH GUID. What it must never do is pick which
    ///  of several declarations on one GUID keeps it - the caller names the
    ///  ones to change - and what it must never write is a GUID that is not
    ///  one, or a line the check did not report.</summary>
    [Test] procedure AFreshGuidGoesWhereTheCheckFoundTheOldOne;
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
    /// <summary>Forum #225: "Remove unused private ..." took the
    ///  declaration and the body and left the XMLDOC above them behind,
    ///  documenting nothing.</summary>
    [Test] procedure RemovePrivate_TakesTheDocCommentAlong;
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
    /// <summary>Audit #36, H36: a quick-fix id carries the GENERATION of
    ///  the listing it came from. With an unchanged buffer but changed
    ///  diagnostics the indices shift, and an id of the first listing
    ///  used to select a different fix in the second one.</summary>
    [Test] procedure AFixIdNamesItsOwnListing;
    /// <summary>Found live on 2026-10-04: extract_method on a file the IDE
    ///  did not have open left it loaded and MODIFIED with no tab - an edit
    ///  nobody can see and Ctrl+Z cannot reach. Only the quick-fix tools
    ///  cleaned up after themselves; the rule is shared now.</summary>
    [Test] procedure AModuleTheCallOpenedIsNotLeftBehind;
    /// <summary>Measured live (2.0.4): interface_guids answered "The
    ///  declarations you named carry a fresh GUID now" after an apply whose
    ///  only id matched nothing at all - a static note asserting success,
    ///  which is the lie 1.25.1 found in the line-ending tool. The sentence
    ///  is DERIVED from the outcomes now, in one shared place.</summary>
    [Test] procedure AnApplyNoteSaysWhatReallyHappened;
    /// <summary>Reported by a reviewing model (2026-10-19): an empty
    ///  lsp_hover answer carried "note" twice, because AddPair APPENDS.
    ///  A duplicate key is valid JSON of which a parser keeps one pair -
    ///  its choice - so one of the two sentences was lost at random.</summary>
    [Test] procedure AnAnswerCarriesNoKeyTwice;
    /// <summary>The second half of the 2026-10-26 IDE hang: a closure
    ///  handed to TThread.Queue runs inside CheckSynchronize, and when it
    ///  RAISES, CheckSynchronize calls ApplicationHandleException
    ///  (System.Classes.pas:16374) - in the VCL a modal
    ///  TApplication.MessageBox, inside whichever IDE transaction pumped
    ///  the queue, and one that does not even raise ModalLevel. So the
    ///  exception may not escape; QueueSafe is the one place that holds
    ///  that rule, and it counts what it swallows.</summary>
    [Test] procedure AQueuedClosureNeverLetsAnExceptionEscape;
  end;

  /// <summary>A forum screenshot (2026-10-02) showed the Extract Method
  ///  dialog saying nothing but "No .delphilsp.json found." - a dead end:
  ///  the file is generated by the IDE, and only while a Code Insight
  ///  option is on. Nine places reported this in five wordings, three of
  ///  them with an option path that was missing its "Delphi" level.</summary>
  [TestFixture]
  TLspConfigHintTests = class
  public
    [Test] procedure TheHintNamesTheFileAndTheSwitch;
    [Test] procedure TheOptionPathIsSpelledInOnePlaceOnly;
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

  /// <summary>Audit issue #40, the three items of its remainder that are
  ///  decided by pure code: the UNIT rename must only touch places where a
  ///  unit NAME can stand (M13), the semantic-replace preview must show
  ///  the occurrence the apply changes (L5a), and a descendant's
  ///  reintroduced or differently-overloaded member is not the symbol
  ///  being renamed (M21b).</summary>
  /// <summary>Audit issue #37, the two remaining items that pure code
  ///  decides: a second file of the same unit name is left out of the
  ///  uses graph and must be NAMED (L3k), and the library path has to be
  ///  expanded per PLATFORM instead of always for Win32 (L3c).</summary>
  /// <summary>Audit issue #39, the four items of its remainder that pure
  ///  code decides: extract method must see every WRITE in the block
  ///  (L7f) and refuse a name that is no identifier (L7e), the MCP
  ///  extract_variable must match its expression as CODE (L7d), and an
  ///  LF-only buffer is more than one line (L7p).</summary>
  [TestFixture]
  TExtractAndCompletionTests = class
  public
    [Test] procedure EveryWriteInTheBlockIsFound;
    [Test] procedure AMethodNameMustBeAnIdentifier;
    [Test] procedure TheExpressionIsMatchedAsCode;
    [Test] procedure LinesAreSplitOnEveryLineBreakStyle;
    /// <summary>Forum #224: a local "Handles: array[0..0] of THandle"
    ///  arrived in the extracted method as "Handles: array;" - the type
    ///  was read as ONE identifier after the colon. And what a local may
    ///  be written as, a parameter may not.</summary>
    [Test] procedure AWrittenOutTypeIsTheWholeType;
    /// <summary>Forum #226/#227: the extracted method was given a
    ///  parameter for "AWorkSheet.Cell" and none for LineStr, and it
    ///  declared a local "AsFloat: string" that nothing uses. Every
    ///  identifier of the block was asked about, the MEMBER after a dot
    ///  included - and a member belongs to a type, never to the block's
    ///  surroundings.</summary>
    [Test] procedure AMemberIsNoFreeIdentifierOfTheBlock;
    /// <summary>Forum 2026-10-03: the generated CALL must stand where the
    ///  extracted code stood. The selection cannot say that - a drag that
    ///  starts at the first statement carries no leading whitespace, and
    ///  the call went in at column 1.</summary>
    [Test] procedure TheCallKeepsTheCodesOwnIndentation;
    /// <summary>Forum 2026-10-04: a block that assigns the enclosing
    ///  FUNCTION's Result was extracted into a PROCEDURE with a local
    ///  variable named Result - of the right type, so it compiles, and the
    ///  function's own result is never assigned.</summary>
    [Test] procedure ResultIsTheFunctionsReturnValueNotAVariable;
    /// <summary>Measured live on 2026-10-04 with dcc32: extracting a block
    ///  whose local is the utterly ordinary "S" produced the parameter "AS"
    ///  - 'as' is an OPERATOR, so the unit did not compile (E2029 four
    ///  times). A generated name must never be a reserved word.</summary>
    [Test] procedure AGeneratedParameterNameIsNoReservedWord;
  end;

  /// <summary>Since L3c resolves $(DXVCL), the DFM check reads DevExpress's
  ///  own event types - and compared their parameter types as text. cxTL's
  ///  OnGetNodeImageIndex says "var AIndex: TcxImageIndex", the handler the
  ///  IDE wrote says TImageIndex, and cxGraphics declares "TcxImageIndex =
  ///  System.UITypes.TImageIndex;": the same type, yet listed as a
  ///  signature mismatch and pre-ticked for the auto-fix (user report on
  ///  1.16.32).</summary>
  [TestFixture]
  TDfmAliasOnlyTests = class
  private
    FDir: string;
    function WriteFile(const AName, AText: string): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;
    [Test] procedure AnotherNameForTheSameTypeIsANoteNotAMismatch;
  end;

  /// <summary>"Edit methods" (user, 2026-10-03): the pure planner behind
    ///  the dialog - which members may travel alone, what a move writes into
    ///  both units, what it refuses, and what it NAMES instead of doing.
    ///  </summary>
  [TestFixture]
  TMethodEditTests = class
  private
    function SourceUnit: TArray<string>;
    function TargetUnit: TArray<string>;
  public
    [Test] procedure TheListSaysWhatCanTravelAlone;
    [Test] procedure AMethodMovesWithItsBodyAndIsRequalified;
    [Test] procedure AMemberThatCannotTravelAloneIsRefused;
    [Test] procedure ModifiersAreEditedOnTheDeclaration;
    /// <summary>Both uses clauses have to follow a move, and the one
    ///  direction that can close a circle (a moved DECLARATION naming a type
    ///  of the old unit) is named instead of risked.</summary>
    [Test] procedure TheUsesClausesFollowAndTheCircleIsNamed;
    /// <summary>The ordinary case: the target class lives in the SAME unit.
    ///  Two independent rewrites of one file would lose half the work, so the
    ///  removal and the insertion happen in one text - and the insertion
    ///  points have to be located again in the already shortened one.</summary>
    [Test] procedure AClassInTheSameUnitIsOneRewrite;
    /// <summary>A signature is changed at BOTH headers or at neither - a
    ///  declaration whose implementation was not rewritten is a unit that
    ///  does not compile. The modifiers reach every selected member.</summary>
    [Test] procedure ASignatureChangeReachesBothHeaders;
    /// <summary>The order the request is carried out in: the signature is
    ///  rewritten BEFORE the move, so a member travels with the list it is
    ///  supposed to have - and a request that asks for nothing says so
    ///  instead of writing an unchanged file.</summary>
    [Test] procedure TheRequestIsCarriedOutInOneOrder;
    /// <summary>The MCP tool's own contract: it takes the arguments the
    ///  dialog has, it offers apply, and it needs no preview token because
    ///  every call re-runs its own analysis.</summary>
    [Test] procedure TheToolDeclaresWhatItTakes;
    /// <summary>User, 2026-10-03: "Wenn ich drei Methoden markiere und den
    ///  Dialog aufrufe, sind diese nicht markiert." A SELECTION names the
    ///  members the dialog ticks - the declarations it covers, and the
    ///  bodies, because either is what a user marks.</summary>
    [Test] procedure ASelectionNamesEveryMemberItCovers;
    /// <summary>User, 2026-10-03, after moving five methods into another
    ///  unit: "Im Ziel sind ploetzlich zwei Zeilen zwischen den Methoden."
    ///  A body's delete range already carries the blank line that followed
    ///  it, so the separator was added on top of one that was there.</summary>
    [Test] procedure MovedBodiesAreSeparatedByExactlyOneBlankLine;
    /// <summary>Same report, second half: BOTH units had grown a blank line
    ///  at the END. SplitContentLines keeps the final break as a trailing
    ///  empty element, so the lines of a planned file must be joined and
    ///  assigned to TStringList.Text - adding them one by one lets the
    ///  list's own break turn that element into a new line.</summary>
    [Test] procedure AnUnchangedPlanIsByteIdentical;
    /// <summary>User, 2026-10-04: after the apply the dialog closes and its
    ///  occurrence list goes with it - "man muesste danach noch einen
    ///  nicht-modalen Dialog haben, der die Referenzen auf die alte Position
    ///  zum Durchklicken anzeigt". What that window may show is this rule:
    ///  every occurrence EXCEPT the edit's own result, with an honest total.
    /// </summary>
    [Test] procedure WhatStillNeedsHandWorkAfterTheEdit;
    /// <summary>User, 2026-10-04: he moved five methods whose headers name
    ///  IOTAKeyContext / TKeyBindingResult into a unit without ToolsAPI in
    ///  its uses. The move said nothing - the compiler did. A unit is only
    ///  named when the SOURCE's own uses clause has it: the moved code
    ///  compiled there, so that is where its types come from.</summary>
    [Test] procedure TheTargetGetsTheUnitsTheMovedCodeNeeds;
    /// <summary>User, 2026-10-04: "Wenn ich die erste bearbeite und dabei
    ///  eine Zeile loesche, passen die Fundstellen danach nicht mehr." A row
    ///  remembers the TEXT it stood on, so its line is found again.</summary>
    [Test] procedure AFindingFollowsItsLineThroughAnEdit;
    /// <summary>A preview must not report a line as changed that did not
    ///  change. Found live: the first edit_method call answered 45 changed
    ///  lines for a move that touches eight, with pairs like
    ///  "implementation" -> "implementation" in them - the shared diff
    ///  pairs the shifted region line by line and emitted every pair.</summary>
    [Test] procedure APreviewReportsOnlyRealChanges;
  end;

  /// <summary>The line-ending check the IDE runs when a file is opened
  ///  (user request, 2026-10-15: "Delphi kommt nicht damit klar, wenn nicht
  ///  das volle Windows Zeilenende verwendet wird, insbesondere wenn
  ///  verschiedene in der Datei vorkommen"). The classification and the
  ///  three gates that decide whether a file may be REWRITTEN are pure;
  ///  the warning dialog and the notifier are not, and the note in
  ///  CLAUDE.md says so.</summary>
  [TestFixture]
  TLineEndingCheckTests = class
  private
    FDir: string;
    function WriteRaw(const AName: string; const ABytes: TBytes): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;
    [Test] procedure MixedAndLfOnlyAreNamedAsTheProblemTheyAre;
    /// <summary>The gates: only source text, never a binary form file -
    ///  the fix rewrites BYTES, so a wrong answer here destroys a file.</summary>
    [Test] procedure NothingButSourceTextMayBeRewritten;
    /// <summary>The fix the dialog offers, end to end: every break becomes
    ///  CRLF, the encoding and the content survive, and a second run is a
    ///  no-op.</summary>
    [Test] procedure TheFixProducesCrlfAndChangesNothingElse;
    /// <summary>"Abbrechen scheint das gleiche wie Leave it zu machen. Es
    ///  sollte aber nur nix tun, aber beim naechsten Mal nachfragen, denn
    ///  wenn man das nicht will, kann man das ja auswaehlen" (user,
    ///  2026-10-15). Cancel decides NOTHING - the file was remembered
    ///  before the dialog even opened, so a cancelled question silenced it
    ///  for the whole session.</summary>
    [Test] procedure CancelDecidesNothingAndLeaveItDoes;
    /// <summary>The project check (user request, 2026-10-15: "einen
    ///  Project check analog zum dfm Check"): which files it reports, which
    ///  it must never report, and that every finding carries a stable id -
    ///  an apply names ids, never list positions.</summary>
    [Test] procedure TheProjectCheckReportsOnlyWhatItMayRewrite;
    /// <summary>...and the tool declares what it reads. An argument the
    ///  schema does not list is dropped silently (1.16.11).</summary>
    [Test] procedure TheLineEndingToolDeclaresWhatItTakes;
    /// <summary>Found live on 2026-10-15: one id named, three OTHER files
    ///  listed - the row that was acted on had been cut away by the row
    ///  budget, so what happened to it was invisible.</summary>
    [Test] procedure WhatWasActedOnIsNeverCutFromTheAnswer;
  end;

  /// <summary>Forum idea 2026-10-15 ("Extract Constant" for a number,
  ///  "Extract Resource string" for a string) in the shape the user chose:
  ///  ONE function that takes the selection as the value and is told the
  ///  type and the place.</summary>
  [TestFixture]
  TDeclareConstantTests = class
  private
    function Unit1: TArray<string>;
    function ColOf(const ALine, AText: string): Integer;
    function Occurrences(const AText, AWhat: string): Integer;
    function ReqFor(const ALines: TArray<string>; ALine0, AStart, AEnd: Integer;
      const AName, AType: string; APlace: TDeclarePlace;
      ASection: TDeclareSection; AReplaceAll: Boolean): TDeclareRequest;
  public
    [Test] procedure OnlyALiteralCanBecomeAConstant;
    [Test] procedure TheNameIsOfferedFromWhatTheValueMeans;
    [Test] procedure ALocalConstantLandsInItsOwnSection;
    [Test] procedure AResourceStringIsAUnitLevelStringOnly;
    [Test] procedure AUnitConstantJoinsTheSectionThatIsThere;
    [Test] procedure WhatWouldNotCompileIsRefusedWithAReason;
    /// <summary>Request 2026-10-16: the value is the user's, not the
    ///  selection's - and without a selection the whole thing is defined
    ///  from scratch, where nothing in the code may change.</summary>
    [Test] procedure TheValueIsEditableAndMayComeFromNowhere;
    /// <summary>...and the tool declares every argument its handler reads:
    ///  one the schema does not list is dropped before the handler sees it,
    ///  and the caller is told nothing (1.16.11).</summary>
    [Test] procedure TheConstantToolDeclaresWhatItTakes;
    /// <summary>Screenshot 2026-10-16: nested routines before the
    ///  outer begin hid the routine the caret is really in.</summary>
    [Test] procedure ANestedRoutineDoesNotHideTheOneAroundIt;
    [Test] procedure AClassMemberIsNoNestedRoutine;
  end;

  /// <summary>Forum idea 2026-10-16: convert a project's sources from one
  ///  encoding to another ("man wuerde angeben von wo nach wo"). The
  ///  measurement that shaped it is in TSourceEncoding - a BOM-less UTF-8
  ///  unit is read by the compiler as ANSI.</summary>
  [TestFixture]
  TFileEncodingCheckTests = class
  private
    FDir: string;
    function WriteRawEnc(const AName: string; const ABytes: TBytes): string;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;
    [Test] procedure WhatAFileISAndWhyItMatters;
    [Test] procedure OnlyWhatWasAskedForAndMayBeRewritten;
    [Test] procedure BothChecksShareOneRowRule;
    [Test] procedure TheEncodingToolDeclaresWhatItTakes;
    /// <summary>Found by running the check against a real project
    ///  (1.31.3): it offered to convert the shipped Indy IdHTTP.pas,
    ///  because the scope takes every unit OPEN in the editor and the
    ///  BDS-root guard only ever covered the uses closure. A check that
    ///  rewrites BYTES must never reach into the installation.</summary>
    [Test] procedure NoCheckRewritesTheInstallation;
  end;

  /// <summary>Forum report 2026-10-15: Delphi had been installed by an admin
  ///  account, so rsvars.bat and the registry name THAT account's Documents
  ///  folder as the common directory - and every build of a user without
  ///  those rights died with MSB3191 on a path it cannot create.</summary>
  [TestFixture]
  TInstallOutputRootTests = class
  public
    [Test] procedure ABuildNeverWritesWhereItMayNot;
    /// <summary>Forum #223: moving the three OUTPUTS was not enough -
    ///  CodeGear.Delphi.Targets creates directories from properties they
    ///  do not reach, so the build still died with MSB3191 on folders it
    ///  never writes to. The ROOT has to move with them.</summary>
    [Test] procedure TheBuildIsToldWhereItMayCreateDirectories;
    /// <summary>What a config may name as the built BPL. {$LIBSUFFIX AUTO}
    ///  appends the PRODUCT suffix, and only Delphi 13 has it equal to its
    ///  registry version ("37.0" -> 370): Delphi 11 registers as 22.0 with
    ///  280, Delphi 12 as 23.0 with 290. So the suffix is read from the
    ///  installation, or installing into either of them would register a
    ///  file name that does not exist.</summary>
    [Test] procedure TheLibSuffixComesFromTheInstallation;
    /// <summary>Installing over a file somebody is RUNNING. Windows
    ///  refuses to overwrite such a file but allows renaming it, and an
    ///  installer that does not know this reports success without
    ///  replacing anything - which is how a stale MCP bridge survived four
    ///  versions (1.22.1).</summary>
    [Test] procedure AFileInUseIsMovedAsideAndReplaced;
    /// <summary>Issue #52: 2.6.9 gated only the artefact copy on the
    ///  build, so after a failed build the packages and experts were still
    ///  registered and the post events still verified the previous build's
    ///  file. This pins the policy InstallEntry gates those steps on:
    ///  nothing after a failed build - and everything after a good one, or
    ///  an entry without a build would stop installing.</summary>
    [Test] procedure NothingFollowsAFailedBuild;
    /// <summary>...and the other half of that question, found while
    ///  reviewing #52: the build can succeed while its artefact does not
    ///  reach the target (a missing source file, a target nobody may
    ///  write). The post event then answers for the file that was already
    ///  there and the install reports it as verified - how a nine-day-old
    ///  MCP bridge survived (1.15.2). A failed copy withdraws the post
    ///  events and NOTHING else: a package is registered where the BUILD
    ///  writes its BPL.</summary>
    [Test] procedure APostEventNeverVerifiesAFileTheInstallDidNotWrite;
  end;

  [TestFixture]
  /// <summary>A unit that no path can reach because it is shipped COMPILED
  ///  inside a package - DesignIntf and the rest of source\ToolsAPI
  ///  (user report, 2026-10-28: "wird darauf hingewiesen, dass die Unit
  ///  nicht im Pfad ist. In dem Fall reicht aber designide in requires,
  ///  etwas anderes ist gar nicht moeglich, da nichts weiter mitgeliefert
  ///  wird").</summary>
  TUnitFromPackageTests = class
  public
    /// <summary>Which package ships a unit, and the measurement that
    ///  bounds the claim: all 42 units of source\ToolsAPI resolve with
    ///  -LUdesignide alone, while the 67 of SOURCE\PROPERTY EDITORS are in
    ///  no .dcp at all - for those the search path really is the way, so
    ///  the rule must not answer for them.</summary>
    [Test] procedure OnlyWhereThePackageIsMeasuredIsItNamed;
    /// <summary>...and whether the project already requires it, which is
    ///  what decides between "nothing to warn about" and "the requires
    ///  clause is the way in". Read off the MASKED text, and 'requires'
    ///  opens a clause only as the first word of its line.</summary>
    [Test] procedure TheRequiresClauseIsRead;
  end;

  [TestFixture]
  TCheckScopeTests = class
  public
    [Test] procedure ADuplicateUnitNameIsReportedAsSkipped;
    [Test] procedure ThePathVariablesKnowTheirPlatform;
    /// <summary>The follow-up of L3c, found by looking at the status window
    ///  of a running IDE: the IDE's base registry key IS the version key,
    ///  so stripping a level off it unconditionally landed on
    ///  'Software\Embarcadero\BDS' - a key that opens and holds nothing.
    ///  The LIBRARY scope of the identifier index was empty from 1.16.24
    ///  to 1.16.30 (measured: 0 units), so every RTL / VCL / third-party
    ///  identifier was unknown.</summary>
    [Test] procedure TheIdeVersionKeyIsFoundByEvidence;
    /// <summary>The other follow-up of L3c: the DFM check's signature index
    ///  gets the shared library walk now, and its private copy had left the
    ///  Studio's source tree out. With FMX indexed ahead of the VCL, the
    ///  first-wins tables checked VCL handlers against FMX event types -
    ///  measured with 1.16.32 on a VCL/DevExpress form whose handlers the
    ///  compiler accepts: eight signature mismatches, every one of them
    ///  ticked for the auto-fix.</summary>
    [Test] procedure TheDfmSignatureScopeLeavesTheStudioSourcesOut;
  end;

  [TestFixture]
  TUnitRenameScopeTests = class
  public
    [Test] procedure OnlyUnitPlacesCountForAUnitRename;
    /// <summary>The same rule over this repository's OWN sources: every
    ///  entry of every real uses clause must be recognised. A filter that
    ///  drops one would make a unit rename leave that reference
    ///  behind - and a fixture cannot cover the shapes a hundred real
    ///  units have.</summary>
    [Test] procedure NoRealUsesClauseEntryIsMissed;
    /// <summary>The other side of the same risk: MemberIsOtherSymbol
    ///  DROPS a rename target, so a false positive silently leaves a
    ///  member unrenamed. Swept over the repository's own overrides,
    ///  where the answer must always be "this IS the symbol".</summary>
    [Test] procedure NoRealOverrideIsTakenForAnotherSymbol;
    [Test] procedure ThePreviewChangesTheOccurrenceTheApplyChanges;
    [Test] procedure AReintroducedMemberIsAnotherSymbol;
  end;

implementation

uses
  System.IOUtils, System.Classes, System.SyncObjs,
  Delphi.FileEncoding, Expert.UsesEditor, Expert.AutoImport, Expert.UnitIndex,
  Expert.DfmEventCheck, Expert.SignatureCheck, Expert.InterfaceGuidCheck,
  Expert.WithScanner, Lsp.Uri, Rename.WorkspaceEdit, Expert.VcsBlame,
  Expert.WorkerLatch, Expert.Version, Expert.PascalScanner, System.RegularExpressions,
  Winapi.Windows, Mcp.PipeServer, Mcp.Protocol, Mcp.Bridge, System.JSON, Lsp.Protocol,
  System.Win.Registry, Expert.PluginSettings, Expert.UsesGraph,
  Expert.MoveToUnit, Expert.SafeDeletePlan, Expert.McpTools, Expert.WithRewriter,
  Expert.SemanticReplace, Expert.SelectionValidator, Expert.ExtractMethod,
  Expert.SignatureEdit, Expert.MethodEdit, System.StrUtils,
  Expert.DebugConsistency, DIH.Types, DIH.Placeholders, DIH.FileOps, DIH.Builder,
  // Generics.Collections: TJSONObject.GetPair is inline and the
  // doubled-note test calls it (H2443 without this entry).
  System.DateUtils, System.Generics.Collections;

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
  All, AnsiLine: TBytes;
begin
  // A diff reproduces source BYTES. Do not encode German text through ANSI:
  // on a cp1251 Windows installation that loses characters before the test.
  // These bytes are invalid UTF-8; their ANSI meaning depends on Windows.
  AnsiLine := TBytes.Create($2B, $20, $20, $47, $72, $F6, $DF, $65,
    $20, $3A, $3D, $20, $31, $3B, $0A);
  All := TEncoding.UTF8.GetBytes('Author: J' + #$00E4 + 'nicke' + #10) +
    AnsiLine +
    TEncoding.UTF8.GetBytes('@@ ' + #$00FC + ' @@');
  Assert.AreEqual('Author: J' + #$00E4 + 'nicke' + #10 +
    TEncoding.ANSI.GetString(AnsiLine) + '@@ ' + #$00FC + ' @@',
    DecodeVcsOutput(All, Length(All)));
end;

procedure TMiscRegressionTests.VcsOutput_MixedUtf8AndAnsiLastLine;
var
  All, AnsiLine: TBytes;
  Header: string;
begin
  Header := 'Author: ' + #$0410 + #$043B + #$0435 + #$043A + #$0441 + #13#10;
  AnsiLine := TBytes.Create($2B, $20, $E0, $FF); // no final newline
  All := TEncoding.UTF8.GetBytes(Header) + AnsiLine;
  Assert.AreEqual(Header + TEncoding.ANSI.GetString(AnsiLine),
    DecodeVcsOutput(All, Length(All)));
end;

procedure TMiscRegressionTests.VcsOutput_Utf8RespectsByteCount;
var
  All: TBytes;
  Text: string;
  Count: Integer;
begin
  Text := #$041F + #$0440 + #$0438 + #$0432 + #$0435 + #$0442 +
    ' ' + #$D83D + #$DE80 + #13#10;
  All := TEncoding.UTF8.GetBytes(Text);
  Count := Length(All);
  All := All + TBytes.Create($FF, $FE); // unread buffer capacity is not output
  Assert.AreEqual(Text, DecodeVcsOutput(All, Count));
  Assert.AreEqual('', DecodeVcsOutput(All, 0));
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

var
  GSelfNotedRan: Integer = 0;

// A worker the LATCH DID NOT START - which is what the live quick-fix
// checker and the blame reader really are: they raise a counter of their
// own, so ShutdownWorkersAndWait never sees them and returns at once. Audit
// #38's L1g was fixed for the latch's own workers and left those two with a
// counter, and a counter is not a join: it drops in the closure's finally,
// i.e. before the closure epilogue (which releases everything it captured,
// with code from this package) and the thread's own exit path have run.
function StartSelfNotedThread(AStarted: TEvent): TThread;
begin
  Result := TThread.CreateAnonymousThread(
    procedure
    begin
      NoteCurrentWorkerThread;
      AStarted.SetEvent;
      Sleep(400);
      AtomicExchange(GSelfNotedRan, 1);
    end);
  Result.FreeOnTerminate := True;
  Result.Start;
end;

procedure TMiscRegressionTests.WorkerLatch_JoinsAWorkerItDidNotStart;
var
  Started: TEvent;
begin
  ResetWorkerLatch;
  AtomicExchange(GSelfNotedRan, 0);
  Started := TEvent.Create(nil, True, False, '');
  try
    StartSelfNotedThread(Started);
    Assert.IsTrue(Started.WaitFor(5000) = wrSignaled, 'the worker runs');
    Assert.IsTrue(JoinNotedWorkers(10000),
      'the join reports that the noted thread is gone');
    Assert.AreEqual<Integer>(1, AtomicCmpExchange(GSelfNotedRan, 0, 0),
      'and it may only report that once the thread really ended - a counter ' +
      'would have let the unload through while this was still running');

    // AND WHAT IT COULD NOT FINISH IS RECORDED, because the only channel
    // out of a finalization is the marker file the next load reads.
    NoteShutdownProblem('a worker did not end');
    NoteShutdownProblem('a worker did not end');   // the same twice says nothing twice
    var P := ShutdownProblems;
    Assert.AreEqual<Integer>(1, Length(P), 'recorded once');
    Assert.AreEqual('a worker did not end', P[0]);
    ResetWorkerLatch;
    Assert.AreEqual<Integer>(0, Length(ShutdownProblems),
      'and a reset means "as if nothing had happened"');
  finally
    for var I := 1 to 100 do
      if AtomicCmpExchange(GSelfNotedRan, 0, 0) = 1 then Break else Sleep(50);
    Started.Free;
    ResetWorkerLatch;
  end;
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

procedure TMcpPipeRegressionTests.AReplacedBridgeHandsOverAndTidiesUp;
const
  Aside = 'C:\x\mcp\RefactoringLightMcp.exe.208982825.old';
  Canon = 'C:\x\mcp\RefactoringLightMcp.exe';
var
  Plan: TBridgeRelaunch;
  Everything, Nothing: TFunc<string, Boolean>;
begin
  Everything := function(APath: string): Boolean begin Result := True; end;
  Nothing := function(APath: string): Boolean begin Result := False; end;

  // The normal case - this IS the installed exe, nothing to do.
  Plan := PlanBridgeRelaunch(Canon, Everything, False);
  Assert.IsFalse(Plan.Wanted, 'the installed exe never hands over');
  Assert.AreEqual('', Plan.NewExe, False);
  Assert.IsTrue(Plan.Reason.Contains('installed exe'), Plan.Reason);

  // An install moved us aside and wrote the new one: hand over.
  Plan := PlanBridgeRelaunch(Aside, Everything, False);
  Assert.IsTrue(Plan.Wanted, 'an aside image hands over: ' + Plan.Reason);
  Assert.AreEqual(Canon, Plan.NewExe, False);

  // ...but only when the successor is really there. A packaged host
  // redirects %LOCALAPPDATA%, so the aside copy can sit in a view the
  // install never wrote to (1.22.3) - then there is nothing to hand to.
  Plan := PlanBridgeRelaunch(Aside, Nothing, False);
  Assert.IsFalse(Plan.Wanted, 'no successor, no hand-over');
  Assert.AreEqual('', Plan.NewExe, False, 'and no path is invented');
  Assert.IsTrue(Plan.Reason.Contains('not there'), Plan.Reason);

  // --no-relaunch is the explicit kill switch and still refuses.
  Plan := PlanBridgeRelaunch(Aside, Everything, True);
  Assert.IsFalse(Plan.Wanted, 'suppressed: serve from here');
  Assert.IsTrue(Plan.Reason.Contains('no-relaunch'), Plan.Reason);

  // ...BUT A SUCCESSOR MAY HAND OVER AGAIN, which is what 1.33.0 forbade by
  // passing --no-relaunch to every one of them: the NEXT install renames
  // that process aside too, and then it served stale for the rest of the
  // session (measured in a live IDE, 2026-10-28 - bridge 2.8.2 against
  // plugin 2.9.4, running from an aside copy). A loop needs none of that
  // ban: a successor runs from the INSTALLED name, which the first case
  // above refuses.
  Plan := PlanBridgeRelaunch(Aside, Everything, False, 1);
  Assert.IsTrue(Plan.Wanted, 'the second hand-over of a chain: ' + Plan.Reason);
  Assert.AreEqual(Canon, Plan.NewExe, False);
  // ...up to a backstop, so a pathological chain cannot grow forever.
  Plan := PlanBridgeRelaunch(Aside, Everything, False, MaxRelaunchHops);
  Assert.IsFalse(Plan.Wanted, 'at the cap it serves from here');
  Assert.IsTrue(Plan.Reason.Contains('already handed the pipes on'), Plan.Reason);

  // The cleanup: every aside copy but the one we run from. That one is
  // in use - Windows refuses it, and the NEXT bridge removes it.
  var Gone := AsideBridgeExesToDelete(
    [Canon, Aside, 'C:\x\mcp\RefactoringLightMcp.exe.7.old',
     'C:\x\mcp\notes.txt', 'C:\x\mcp\RefactoringLightMcp.exe.old'], Aside);
  Assert.AreEqual(1, Integer(Length(Gone)), 'one deletable copy');
  Assert.AreEqual('C:\x\mcp\RefactoringLightMcp.exe.7.old', Gone[0], False);
end;

procedure TMcpPipeRegressionTests.AnAsideExeNamesTheInstallThatReplacedIt;
var
  Dir, Exe, Aside: string;
begin
  // The shape buildmcp.cmd writes: "<exe>.<number>.old" beside the exe.
  Assert.AreEqual('C:\x\mcp\RefactoringLightMcp.exe',
    CanonicalBridgeExePath('C:\x\mcp\RefactoringLightMcp.exe.1052527944.old'),
    'the aside name maps back to the name it was installed under');
  Assert.AreEqual('C:\x\mcp\RefactoringLightMcp.exe',
    CanonicalBridgeExePath('C:\x\mcp\RefactoringLightMcp.exe.7.old'),
    'any number');

  // Everything else is NOT ours to interpret - a guess here would send the
  // user the wrong way, which is the whole point of this round.
  Assert.AreEqual('', CanonicalBridgeExePath('C:\x\mcp\RefactoringLightMcp.exe'),
    'the plain exe is no aside file');
  Assert.AreEqual('', CanonicalBridgeExePath('C:\x\mcp\RefactoringLightMcp.exe.old'),
    'the OLD fixed name carries no number - 1.22.1 replaced it');
  Assert.AreEqual('', CanonicalBridgeExePath('C:\x\mcp\RefactoringLightMcp.exe.keep.old'),
    'not a number');
  Assert.AreEqual('', CanonicalBridgeExePath('C:\x\mcp\something.7.old'),
    'the name it maps back to must be an exe');
  Assert.AreEqual('', CanonicalBridgeExePath(''), 'nothing to read');

  // And the verdict itself, on real files: a process whose image was moved
  // aside while the installed name exists again IS running a replaced exe.
  Dir := TPath.Combine(TPath.GetTempPath, 'rl_aside_' + IntToStr(GetCurrentProcessId));
  TDirectory.CreateDirectory(Dir);
  try
    Exe := TPath.Combine(Dir, 'RefactoringLightMcp.exe');
    Aside := Exe + '.4711.old';
    TFile.WriteAllText(Aside, 'old');
    Assert.AreEqual(Exe, CanonicalBridgeExePath(Aside), 'the name maps back');
    Assert.IsFalse(TFile.Exists(Exe),
      'nothing installed yet, so the install did not finish - no advice to start a session');
    TFile.WriteAllText(Exe, 'new');
    Assert.IsTrue(TFile.Exists(CanonicalBridgeExePath(Aside)),
      'the installed name is back: only the running session is old');
  finally
    TDirectory.Delete(Dir, True);
  end;
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

  // User, 2026-10-04: "version mismatch trotz install.cmd und Claude Code
  // Neustart". Measured: the exe on disk was the NEW one and only his running
  // bridge was old - so the message must not send him to install.cmd again.
  // The two cases need OPPOSITE actions, which is why the age decides.
  var Replaced := BridgeVersionProblem('1.16.29', '1.22.0', beaReplaced);
  Assert.IsTrue(Replaced.Contains('1.16.29') and Replaced.Contains('1.22.0'),
    'both numbers, whatever the age: ' + Replaced);
  Assert.IsTrue(Replaced.Contains('new Claude Code session'),
    'the exe is already right, so a new session is the whole fix: ' + Replaced);
  // THE SESSION COMES FIRST, which is the whole point of 1.22.3: the exe on
  // disk is already right, so sending the user to the install was what made
  // him run it twice. It may still be NAMED afterwards, and since 2.9.5 it
  // is - the live state of 2026-10-28 was a bridge replaced at 10:54 (so
  // "replaced") whose on-disk version was 2.9.0 against a plugin of 2.9.4,
  // because the IDE's own "Install" renews the BPL alone. A new session
  // alone could not have fixed that, and saying nothing about it would send
  // the next report in a circle of its own.
  Assert.IsTrue(Pos('new Claude Code session', Replaced) <
                Pos('install.cmd', Replaced + 'install.cmd'),
    'the new session is the FIRST thing it says, the install at most the ' +
    'second: ' + Replaced);
  // AND IT SAYS FROM WHEN THE BRIDGE DOES IT BY ITSELF, because until 2.9.5
  // it did so ONCE per session: the successor of a hand-over was started
  // with --no-relaunch, so a SECOND install left it serving stale (live,
  // 2026-10-28). The sentence promised otherwise.
  Assert.IsTrue(Replaced.Contains(FirstRepeatedHandOver),
    'the version from which a bridge hands over as often as it is '
    + 'replaced: ' + Replaced);

  var AsOld := BridgeVersionProblem('1.16.29', '1.22.0', beaAsOld);
  Assert.IsTrue(AsOld.Contains('install.cmd'),
    'an exe as old as its process was never replaced: ' + AsOld);
  // beaUnknown must behave like "not replaced": naming the install is the
  // safe direction when Windows would not say.
  Assert.AreEqual(AsOld, BridgeVersionProblem('1.16.29', '1.22.0', beaUnknown),
    'an unknown age reads as "the install did not replace it"');
  Assert.AreEqual(AsOld, BridgeVersionProblem('1.16.29', '1.22.0'),
    'and that is also the default, so every old caller keeps its wording');
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

procedure TMcpPipeRegressionTests.Stop_WakesAStragglerAndThenStopsCleanly;
var
  Srv: TMcpPipeServer;
  Entered, Release: TEvent;
  Client: TThread;
  Probe: string;
  Wakes: Integer;
begin
  Probe := 'wake-' + FormatDateTime('hhnnsszzz', Now);
  Wakes := 0;
  Entered := TEvent.Create(nil, True, False, '');
  Release := TEvent.Create(nil, True, False, '');
  try
    Srv := TMcpPipeServer.Create(McpPipeName(GetCurrentProcessId),
      function(const ARequest: string; AStop: THandle): string
      begin
        if ARequest.Contains(Probe) then
        begin
          Entered.SetEvent;
          // ignores AStop - only the wake action ends this wait, as ending
          // DelphiLSP ends an LSP one
          Release.WaitFor(10000);
        end;
        Result := '{"ok":true}';
      end);
    try
      Assert.IsTrue(Srv.Start, 'the test pipe is there: ' + Srv.LastError);
      Client := StartProbeClient(Probe);
      try
        Assert.IsTrue(Entered.WaitFor(5000) = wrSignaled, 'the handler runs');
        Assert.IsTrue(Srv.Stop(100,
          procedure
          begin
            Inc(Wakes);
            Release.SetEvent;
          end, 5000),
          'the woken handler leaves within the grace period - a clean stop');
        Assert.AreEqual(1, Wakes, 'the wake action ran once');
        Assert.AreEqual(0, Srv.ActiveHandlers);
        Assert.IsTrue(Srv.Stop(100,
          procedure
          begin
            Inc(Wakes);
          end),
          'nothing running: a clean stop');
        Assert.AreEqual(1, Wakes, 'and no wake when no handler needs one');
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

procedure TMcpPipeRegressionTests.EveryToolParameterIsDescribed;
var
  Props: Integer;
  Bad: string;

  procedure Note(const AWhat: string);
  begin
    if Bad <> '' then Bad := Bad + '; ';
    Bad := Bad + AWhat;
  end;

begin
  Props := 0;
  Bad := '';
  var Arr := McpToolDefinitions;
  try
    Assert.IsTrue((Arr <> nil) and (Arr.Count > 40), 'the shipped list');
    for var V in Arr do
    begin
      var Tool := V as TJSONObject;
      var Name := Tool.GetValue<string>('name', '');
      var Schema := Tool.GetValue<TJSONObject>('inputSchema', nil);
      if Schema = nil then Continue;
      var PropsObj := Schema.GetValue<TJSONObject>('properties', nil);
      if PropsObj = nil then Continue;
      for var Pair in PropsObj do
      begin
        Inc(Props);
        var Body := Pair.JsonValue as TJSONObject;
        var Desc := Body.GetValue<string>('description', '');
        if Trim(Desc) = '' then
        begin
          Note(Name + '.' + Pair.JsonString.Value + ' has no description');
          Continue;
        end;
        // "1-based column" does not say HOW it counts - the question a
        // reviewer asked, and the answer belongs where it is asked.
        if SameText(Pair.JsonString.Value, 'column') or
           SameText(Pair.JsonString.Value, 'end_column') then
          if not Desc.ToLower.Contains('tab') then
            Note(Name + '.' + Pair.JsonString.Value +
              ' does not say how it counts a tab');
      end;
    end;
  finally
    Arr.Free;
  end;

  Assert.AreEqual('', Bad, Bad);
  Assert.IsTrue(Props > 200, 'every parameter of every tool was looked at');
end;

procedure TMcpPipeRegressionTests.TheLspDcuShadowIsFoundAndMovable;
var
  Dir, Cfg, Err, Target: string;
begin
  // 1. the option, in the shapes a real dccOptions line has.
  Assert.AreEqual('.\DCU\Win64\Debug',
    DcuOutputFromDccOptions('-$O- --no-config -NU.\DCU\Win64\Debug -NSSystem'),
    'the -NU of our own configuration, verbatim');
  Assert.AreEqual('C:\out dir\dcu',
    DcuOutputFromDccOptions('-I"c:\x" -NU"C:\out dir\dcu" -R"c:\y"'),
    'a quoted path with a space');
  Assert.AreEqual('', DcuOutputFromDccOptions('-$O- -NSSystem -NB"c:\dcp"'),
    'no -NU at all');
  Assert.AreEqual('', DcuOutputFromDccOptions(''), 'nothing');

  Dir := TPath.Combine(TPath.GetTempPath, 'rl-dcu-' + IntToStr(GetCurrentThreadId));
  try
    TDirectory.CreateDirectory(Dir);
    // 2. a config file as the IDE writes it: the options are ONE JSON
    //    string, so the backslashes are escaped in it.
    Cfg := TPath.Combine(Dir, 'Demo.delphilsp.json');
    TFile.WriteAllText(Cfg,
      '{"settings":{"project":"file:///x","dccOptions":"-$O- ' +
      '-NU.\\\\DCU\\\\Win64\\\\Debug -NSSystem","projectFiles":[]}}');
    Assert.AreEqual(TPath.Combine(Dir, 'DCU\Win64\Debug'),
      LspDcuDirOfConfig(Cfg, Dir),
      'the directory, resolved against the project like the compiler does');

    // 3. WHAT THE DIRECTORY SAYS ABOUT ITSELF. MEASURED on a live session
    //    (2026-10-20, eight states, one variable each): a unit the LSP
    //    loads from a .dcu DOES answer with its source position, so the
    //    finding is not "there are .dcu files here". What broke go to
    //    definition was a directory that does not MATCH the sources: a
    //    unit missing from it, or a .dcu older than its own source.
    var DcuDir := TPath.Combine(Dir, 'DCU\Win64\Debug');
    TDirectory.CreateDirectory(DcuDir);
    var Far := TPath.Combine(Dir, 'Elsewhere');
    TDirectory.CreateDirectory(Far);
    for var N in ['Unit1.pas', 'Unit2.pas', 'Demo.dpr', 'Main.dfm'] do
      TFile.WriteAllText(TPath.Combine(Dir, N), 'unit ' + N + ';');
    TFile.WriteAllText(TPath.Combine(Far, 'Hidden.pas'), 'unit Hidden;');
    var Sources: TArray<string> := [TPath.Combine(Dir, 'Unit1.pas'), TPath.Combine(Dir, 'Unit2.pas'),
                    TPath.Combine(Dir, 'Demo.dpr'), TPath.Combine(Dir, 'Main.dfm'),
                    TPath.Combine(Far, 'Hidden.pas')];
    // Unit1 + Hidden built, Unit2 not - and a .dcu of a unit the project
    // does not have must not count at all.
    for var N in ['Unit1.dcu', 'Hidden.dcu', 'Vcl.Forms.dcu'] do
      TFile.WriteAllText(TPath.Combine(DcuDir, N), 'x');
    var St := InspectLspDcuDir(DcuDir, Sources, [Dir]);
    // The .dfm is no unit - and neither is the project FILE: measured over
    // three projects of this repository, no .dcu is ever named after the
    // .dpk or the .dpr (that file produces the .bpl or the .exe). Counting
    // it reported the directory as incomplete by one, for a file that can
    // never satisfy it, while the status row said it matches (live, 2.5.0).
    Assert.AreEqual(3, St.UnitCount,
      'three units: Unit1, Unit2, Hidden - not Main.dfm and not Demo.dpr');
    Assert.AreEqual(2, St.Present, 'Unit1.dcu and Hidden.dcu');
    Assert.AreEqual(1, Integer(Length(St.Missing)),
      'only Unit2 has none: ' + string.Join(', ', St.Missing));
    // Its OWN spelling, not the lower-cased lookup key: the live 2.5.0
    // answer read "First missing: delphirefactoringlight".
    Assert.IsTrue(St.Missing[0] = 'Unit2',
      'and it is Unit2, never the project file, spelled as the source ' +
      'is: ' + St.Missing[0]);
    Assert.AreEqual(0, Integer(Length(St.Stale)),
      'and none of them is older than its source');

    // The same for a PACKAGE, which is what the live finding was about.
    TFile.WriteAllText(TPath.Combine(Dir, 'Pack.dpk'), 'package Pack;');
    var StP := InspectLspDcuDir(DcuDir,
      [TPath.Combine(Dir, 'Unit1.pas'), TPath.Combine(Dir, 'Pack.dpk')], [Dir]);
    Assert.AreEqual(1, StP.UnitCount, 'the .dpk is no unit either');
    Assert.AreEqual(0, Integer(Length(StP.Missing)),
      'so a fully built package reports nothing missing: ' +
      string.Join(', ', StP.Missing));

    // A .dcu older than its source is the one per-file finding left - and
    // whether moving it aside HELPS depends on the compiler being able to
    // reach the source instead.
    TFile.SetLastWriteTime(TPath.Combine(DcuDir, 'Unit1.dcu'),
      IncHour(Now, -5));
    TFile.SetLastWriteTime(TPath.Combine(DcuDir, 'Hidden.dcu'),
      IncHour(Now, -5));
    St := InspectLspDcuDir(DcuDir, Sources, [Dir]);
    Assert.AreEqual(2, Integer(Length(St.Stale)),
      'both are older than their source now');
    var Repl := '';
    for var I := 0 to High(St.Stale) do
      if St.StaleReplaceable[I] then
        Repl := Repl + ExtractFileName(St.Stale[I]) + ' ';
    Assert.IsTrue(Repl = 'Unit1.dcu ',
      'only Unit1''s source lies in a directory the LSP searches, so only ' +
      'there is removing the .dcu a cure; Hidden.dcu is the ONLY way to ' +
      'that unit and removing it would cost the answer (measured)');

    // And the count a complete directory is compared against comes from
    // the configuration itself.
    var Cfg3 := TPath.Combine(Dir, 'Count.delphilsp.json');
    TFile.WriteAllText(Cfg3,
      '{"settings":{"dccOptions":"-$O-","projectFiles":[' +
      '{"name":"A","file":"file:///a"},{"name":"B","file":"file:///b"}]}}');
    Assert.AreEqual(2, LspProjectUnitCount(Cfg3), 'two units listed');
    Assert.AreEqual(0, LspProjectUnitCount(TPath.Combine(Dir, 'none.json')),
      'and no file is no claim');

    // ...and the search path is read as ONE option whose value continues
    // past a closing quote - reading only to that quote is the mistake
    // that said "the source directory is on no search path" when it was.
    var Cfg2 := TPath.Combine(Dir, 'Paths.delphilsp.json');
    TFile.WriteAllText(Cfg2,
      '{"settings":{"dccOptions":"-$O- -U\"c:\\\\rtl\";..\\\\Source;\"c:\\\\more\" ' +
      '-NSSystem"}}');
    var Dirs := LspUnitSearchDirs(Cfg2, Dir);
    Assert.IsTrue(Length(Dirs) >= 4,
      'project dir + three -U entries: ' + string.Join(' | ', Dirs));
    var HasSource := False;
    // '..\Source' is resolved against the project directory, so it is
    // the PARENT's Source - that is what the compiler does with it.
    var Want := ExcludeTrailingPathDelimiter(ExpandFileName(
      TPath.Combine(Dir, '..\Source')));
    for var D in Dirs do
      if SameText(D, Want) then HasSource := True;
    Assert.IsTrue(HasSource,
      'the relative entry BETWEEN two quoted ones is found as ' + Want +
      ': ' + string.Join(' | ', Dirs));

    // 4. moved aside, not deleted - and the finding is gone afterwards.
    var Stale1 := TPath.Combine(DcuDir, 'Unit1.dcu');
    var Aside := DcuAsideFolder(DcuDir);
    Assert.IsTrue(MoveDcuAside(Stale1, Aside, Target, Err), Err);
    Assert.IsFalse(TFile.Exists(Stale1),
      'the compiler no longer finds it there');
    Assert.IsTrue(TFile.Exists(Target), 'but it still exists: ' + Target);
    St := InspectLspDcuDir(DcuDir, [TPath.Combine(Dir, 'Unit1.pas')], [Dir]);
    Assert.AreEqual(0, Integer(Length(St.Stale)),
      'and no stale .dcu is reported any more');
    Assert.IsFalse(MoveDcuAside(Stale1, Aside, Target, Err),
      'moving it twice is refused');
    Assert.IsTrue(Err <> '', 'with a reason');
  finally
    try
      if TDirectory.Exists(Dir) then TDirectory.Delete(Dir, True);
    except
    end;
  end;
end;

procedure TMcpPipeRegressionTests.AProjectActionSaysWhatHappensToUnsavedWork;

  function Tool(AArr: TJSONArray; const AName: string): TJSONObject;
  begin
    Result := nil;
    for var V in AArr do
      if SameText((V as TJSONObject).GetValue<string>('name', ''), AName) then
        Exit(V as TJSONObject);
  end;

  function Props(AArr: TJSONArray; const AName: string): TJSONObject;
  begin
    Result := nil;
    var T := Tool(AArr, AName);
    if T = nil then Exit;
    var S := T.GetValue<TJSONObject>('inputSchema', nil);
    if S <> nil then Result := S.GetValue<TJSONObject>('properties', nil);
  end;

begin
  var Arr := McpToolDefinitions;
  try
    for var N in ['project_open', 'project_close'] do
    begin
      var T := Tool(Arr, N);
      Assert.IsNotNull(T, N + ' is a tool of this server');
      var P := Props(Arr, N);
      Assert.IsNotNull(P, N + ' has a schema');
      Assert.IsNotNull(P.GetValue<TJSONObject>('save_modified', nil),
        N + ' must OFFER what happens to unsaved work');
      var D := T.GetValue<string>('description', '');
      Assert.IsTrue(D.ToLower.Contains('refuse'),
        N + ' must say that it refuses while something is unsaved: ' + D);
      Assert.IsTrue(D.ToLower.Contains('unsaved'),
        N + ' must name unsaved work: ' + D);
    end;

    // The project file is not a buffer: no Ctrl+Z, so it is a reviewable
    // change like every other write - and saving it is a choice.
    var PF := Props(Arr, 'project_files');
    Assert.IsNotNull(PF, 'project_files has a schema');
    for var A in ['add', 'remove', 'apply', 'save'] do
      Assert.IsNotNull(PF.GetValue<TJSONObject>(A, nil),
        'project_files declares "' + A + '"');
    var FD := Tool(Arr, 'project_files').GetValue<string>('description', '');
    Assert.IsTrue(FD.Contains('Ctrl+Z'),
      'and says why it needs apply at all: ' + FD);
  finally
    Arr.Free;
  end;
end;

procedure TMcpPipeRegressionTests.AnEmptyAnswerNamesTheRightCause;
const
  Mine = 'DelphiLSP returned nothing for this position.';
  F = 'C:\p\Expert.LspManager.pas';

  function Occurrences(const AText, AWhat: string): Integer;
  begin
    Result := 0;
    var P := Pos(AWhat, AText);
    while P > 0 do
    begin
      Inc(Result);
      P := Pos(AWhat, AText, P + 1);
    end;
  end;

begin
  // The reported case: the unit IS analysed (hover answers), only this
  // position is unresolved.
  var Pos_ := LspEmptyAnswerNote(Mine, '', F, -1, True, True, True);
  Assert.AreEqual(1, Occurrences(Pos_, 'DelphiLSP returned nothing'),
    'the sentence is said ONCE, not by the prefix AND the generic text: ' +
    Pos_);
  Assert.IsTrue(Pos_.StartsWith(Mine), 'the tool speaks first: ' + Pos_);
  // MEASURED 2026-10-20 (session log): with hover answering all along, the
  // cause was the DECLARING unit coming from a stale .dcu. So an analysed
  // unit must name BOTH causes and what tells them apart - the earlier
  // version of this test demanded the opposite and was wrong.
  Assert.IsTrue(Pos_.Contains('DECLARING unit'),
    'the other cause is named: ' + Pos_);
  Assert.IsTrue(Pos_.Contains('file:///(-1:0)'),
    'with the log line that tells the two apart: ' + Pos_);
  Assert.IsTrue(Pos_.Contains('find_original_symbol'),
    'and the way out is named: ' + Pos_);
  Assert.IsTrue(Pos_.Contains('RSS-5463'),
    'with the known case of it: ' + Pos_);

  // The other half: nothing of the unit was ever analysed. THERE the .dcu
  // is the usual cause, and that is where the recipe belongs.
  var Silent := LspEmptyAnswerNote('', '', F, -1, False, True, True);
  Assert.IsTrue(Silent.StartsWith('DelphiLSP returned nothing'),
    'without a prefix the note says it itself: ' + Silent);
  Assert.IsTrue(Silent.Contains('.dcu'), 'the usual cause: ' + Silent);
  // How to see it: the SETTING (which the settings tool can now flip) and
  // the log file it writes - not the environment variable, which needs the
  // IDE to be restarted with it.
  Assert.IsTrue(Silent.Contains('LspLogging') and
    Silent.Contains('RefactoringLight*.log'),
    'and how to see it: ' + Silent);
  Assert.IsFalse(Silent.Contains('RSS-5463'),
    'but not the overload story, which is about a unit that IS analysed: ' +
    Silent);

  // #50.5: four tools passed their DESCRIPTION where the empty-answer
  // sentence belongs, so the one thing the note exists to say was gone and
  // the paragraph about resolving a symbol followed a sentence about
  // ranges. Both halves are their own parameter now.
  const Desc = 'raw LSP DocumentSymbol[] - ranges are 0-based.';
  var Syms := LspEmptyAnswerNote('DelphiLSP returned no symbols for this unit.',
    Desc, F, -1, True, False, False);
  Assert.IsTrue(Syms.StartsWith('DelphiLSP returned no symbols'),
    'what an empty answer MEANS comes first: ' + Syms);
  Assert.IsTrue(Syms.Contains(Desc),
    'and the description is kept beside it, not instead of it: ' + Syms);
  Assert.IsFalse(Syms.Contains('RSS-5463'),
    'document symbols resolve no symbol - that lead would be wrong here: ' + Syms);
  Assert.IsFalse(Syms.Contains('find_original_symbol'), 'nor that one');
  Assert.IsTrue(Syms.Contains('server''s own answer'),
    'it says what an analysed unit means instead: ' + Syms);

  // ...while a tool that DOES resolve a position keeps the diagnosis - and
  // lsp_implementation still does not point at the index, because having
  // no implementation is its normal answer.
  var Impl := LspEmptyAnswerNote('DelphiLSP returned no implementation for '
    + 'this position - for a method that neither overrides nor implements '
    + 'one that is the normal answer.', 'x', F, -1, True, False, True);
  Assert.IsTrue(Impl.StartsWith('DelphiLSP returned no implementation'),
    'the normal case is stated first: ' + Impl);
  Assert.IsTrue(Impl.Contains('RSS-5463'),
    'the diagnosis stays where a position IS resolved: ' + Impl);
  Assert.IsFalse(Impl.Contains('find_original_symbol'),
    'but it is not the answer to "there is no implementation": ' + Impl);

  // A file we have just sent is simply not ready yet - that comes first,
  // before any cause at all.
  var Fresh := LspEmptyAnswerNote(Mine, '', F, 2500, False, False, True);
  Assert.IsTrue(Fresh.Contains('2 s ago'), 'how long ago, in seconds: ' + Fresh);
  Assert.IsTrue(Fresh.Contains('Expert.LspManager.pas'),
    'and which file: ' + Fresh);
  Assert.IsFalse(Fresh.Contains('find_original_symbol'),
    'a tool that is not about declarations does not point there: ' + Fresh);
  Assert.AreEqual(1, Occurrences(Fresh, 'DelphiLSP returned nothing'),
    'still once: ' + Fresh);
end;

procedure TMcpPipeRegressionTests.TheAlwaysLoadedInstructionsStayShort;
var
  Named: Integer;
  Names: string;
begin
  var Instr := McpServerInstructions;
  Assert.IsTrue(Length(Instr) > 400,
    'the instructions must still say something: ' + IntToStr(Length(Instr)));
  Assert.IsTrue(Length(Instr) <= 1100,
    Format('every session pays for this text - %d characters is too much, ' +
      'put it in the guide instead', [Length(Instr)]));

  // It may name the few tools a session cannot do without (the three that
  // exist with no IDE, and the buffer rule it states), but it is NOT the
  // catalogue: a session gets the names anyway, and the guide has room.
  Named := 0;
  Names := '';
  var Arr := McpToolDefinitions;
  try
    for var V in Arr do
    begin
      var N := (V as TJSONObject).GetValue<string>('name', '');
      if (N <> '') and Instr.Contains(N) then
      begin
        Inc(Named);
        Names := Names + ' ' + N;
      end;
    end;
  finally
    Arr.Free;
  end;
  Assert.IsTrue(Named <= 12,
    Format('the instructions name %d tools (%s) - that is a list, and a ' +
      'list belongs in the guide', [Named, Names.Trim]));
  Assert.IsTrue(Instr.Contains('guide'),
    'but the tool that explains the rest has to be findable: ' + Instr);
end;

procedure TMcpPipeRegressionTests.TheGuideAndTheToolListAgree;
var
  Bad: string;
  Counted: Integer;

  procedure Note(const AWhat: string);
  begin
    // ALL of them at once: one name at a time hides how stale it is.
    if Bad <> '' then Bad := Bad + '; ';
    Bad := Bad + AWhat;
  end;

begin
  Bad := '';
  Counted := 0;
  var Secs := McpGuideSections;
  Assert.IsTrue(Length(Secs) > 5,
    'the guide must have its sections: ' + IntToStr(Length(Secs)));

  var Text := McpGuideText;
  Assert.IsTrue(Text.Contains('guide is this text'),
    'the rendered guide is what the tool answers');

  var Arr := McpToolDefinitions;
  try
    Assert.IsTrue((Arr <> nil) and (Arr.Count > 40),
      'the shipped tool list is what this test walks');

    // 1. every tool is in EXACTLY one section, and 3. the prose that
    //    section shows really names it.
    for var V in Arr do
    begin
      var Name := (V as TJSONObject).GetValue<string>('name', '');
      if Name = '' then Continue;
      Inc(Counted);
      var Where := 0;
      for var S in Secs do
        for var T in S.Tools do
          if SameText(T, Name) then
          begin
            Inc(Where);
            if not S.Text.Contains(Name) then
              Note(Format('%s is listed under "%s" but that text does not ' +
                'name it', [Name, S.Title]));
          end;
      if Where = 0 then
        Note(Name + ' is in no section of the guide')
      else if Where > 1 then
        Note(Format('%s is in %d sections', [Name, Where]));
    end;

    // 2. and no section names a tool that does not exist any more.
    for var S in Secs do
      for var T in S.Tools do
      begin
        var Found := False;
        for var V in Arr do
          if SameText((V as TJSONObject).GetValue<string>('name', ''), T) then
            Found := True;
        if not Found then
          Note(Format('"%s" names %s, which is not a tool of this server',
            [S.Title, T]));
      end;
  finally
    Arr.Free;
  end;

  Assert.IsTrue(Counted > 40,
    Format('the sweep must see the tools, counted %d', [Counted]));
  Assert.AreEqual('', Bad, Bad);
end;

procedure TMcpPipeRegressionTests.EveryWritingToolSaysWhatProtectsItsApply;
var
  Token, Fresh, Ids, Guard: Integer;
  Bad: string;

  procedure Note(const AWhat: string);
  begin
    // ALL of them: fixing one at a time hides how many there are.
    if Bad <> '' then Bad := Bad + '; ';
    Bad := Bad + AWhat;
  end;

begin
  Token := 0;
  Fresh := 0;
  Ids := 0;
  Guard := 0;
  Bad := '';
  var Arr := McpToolDefinitions;
  try
    Assert.IsTrue((Arr <> nil) and (Arr.Count > 40),
      'the shipped tool list is what this test walks');
    for var V in Arr do
    begin
      var Tool := V as TJSONObject;
      var Name := Tool.GetValue<string>('name', '');
      var Schema := Tool.GetValue<TJSONObject>('inputSchema', nil);
      if Schema = nil then Continue;
      var Props := Schema.GetValue<TJSONObject>('properties', nil);
      if Props = nil then Continue;
      var ApplyObj := Props.GetValue<TJSONObject>('apply', nil);
      if ApplyObj = nil then Continue;          // a read-only tool

      var Desc := ApplyObj.GetValue<string>('description', '');
      var HasToken := Props.GetValue<TJSONObject>('token', nil) <> nil;
      if HasToken then
      begin
        Inc(Token);
        if not Desc.ToLower.Contains('token') then
          Note(Name + ': it hands out a token and does not say so');
      end
      else
      begin
        // The promise must not be made where it cannot be kept, and the
        // caller has to learn what DOES protect the apply instead.
        if Desc.ToLower.Contains('plus a token') then
          Note(Name + ': promises a token it has no parameter for');
        if not Desc.ToLower.Contains('no token') then
          Note(Name + ': has no token and never says what protects an apply');
        // The three shapes an apply can be protected by, and a tool has
        // to name the one it uses: it looks again (re-runs its analysis),
        // it works from the ids of a scan, or it compares the file with
        // what the analysis saw and refuses a file that moved on.
        if Desc.ToLower.Contains('re-runs') then Inc(Fresh)
        else if Desc.ToLower.Contains('re-scans') then Inc(Ids)
        else if Desc.ToLower.Contains('compares') then Inc(Guard)
        else
          Note(Name + ': says "no token" without naming what it does instead');
      end;
    end;
  finally
    Arr.Free;
  end;

  Assert.AreEqual('', Bad, Bad);
  // ...and the test must not pass because it found nothing to look at.
  Assert.IsTrue(Token >= 9, 'the tools that pin a preview with a token');
  Assert.IsTrue(Fresh >= 5, 'the tools that re-run their own analysis');
  Assert.IsTrue(Ids >= 4, 'the tools that act on the ids of a scan');
  Assert.IsTrue(Guard >= 1, 'and the one that compares before it writes');
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

procedure TSettingsRobustnessTests.EverySettingIsDescribedAndWiredUp;
var
  Bad: string;
  Err, Text, Back: string;
begin
  var Infos := TPluginSettings.SettingInfos;
  Assert.IsTrue(Length(Infos) >= 15,
    'the descriptor must hold the settings: ' + IntToStr(Length(Infos)));
  TPluginSettings.Load;
  Bad := '';
  for var Info in Infos do
  begin
    // Readable by name...
    if not TPluginSettings.GetValueText(Info.Name, Back) then
    begin
      Bad := Bad + Info.Name + ' cannot be READ by name; ';
      Continue;
    end;
    // ...and writable, with its OWN value, which must round-trip. That is
    // what catches a setting added to the descriptor and nowhere else.
    if not TPluginSettings.SetValueText(Info.Name, Back, Err) then
      Bad := Bad + Info.Name + ': ' + Err + '; '
    else
    begin
      TPluginSettings.GetValueText(Info.Name, Text);
      if Text <> Back then
        Bad := Bad + Format('%s: wrote "%s", reads "%s"; ',
          [Info.Name, Back, Text]);
    end;
    if Info.DefaultText = '' then
      if Info.Kind <> pskString then
        Bad := Bad + Info.Name + ' has no default; ';
    if Info.Summary = '' then
      Bad := Bad + Info.Name + ' has no summary; ';
  end;
  Assert.AreEqual('', Bad, Bad);

  // The three refusals a caller can run into, each with its reason.
  Assert.IsFalse(TPluginSettings.SetValueText('NoSuchThing', '1', Err),
    'an unknown name is refused');
  Assert.IsTrue(Err.Contains('NoSuchThing'), 'and named: ' + Err);
  Assert.IsFalse(TPluginSettings.SetValueText('LiveBlame', 'maybe', Err),
    'a switch takes true or false');
  Assert.IsTrue(Err.Contains('switch'), 'and says so: ' + Err);
  Assert.IsFalse(TPluginSettings.SetValueText('BlameInfo', '9', Err),
    'a number outside its range is refused');
  Assert.IsTrue(Err.Contains('between'), 'with the range: ' + Err);
end;

procedure TSettingsRobustnessTests.ASessionOnlyValueIsNeverPersisted;
var
  Reg: TRegistry;
  Key, Err, Persisted: string;
  Had: Boolean;
  Old: Boolean;
begin
  Key := TPluginSettings.RegistryKey;
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Assert.IsTrue(Reg.OpenKey(Key, True), 'settings key');
    Had := Reg.ValueExists('BlameUseTortoise') and
      (Reg.GetDataType('BlameUseTortoise') = rdInteger);
    Old := True;
    if Had then Old := Reg.ReadBool('BlameUseTortoise');
    Reg.CloseKey;
    try
      // A known state in the registry, then the opposite for THIS session.
      TPluginSettings.Load;
      Assert.IsTrue(TPluginSettings.SetValueText('BlameUseTortoise', 'true',
        Err), Err);
      TPluginSettings.Save;

      Assert.IsTrue(TPluginSettings.SetValueText('BlameUseTortoise', 'false',
        Err), Err);
      TPluginSettings.MarkSessionOnly('BlameUseTortoise', True);
      Assert.IsFalse(TPluginSettings.BlameUseTortoise,
        'the change takes effect in this session at once');
      Assert.IsTrue(TPluginSettings.IsSessionOnly('BlameUseTortoise',
        Persisted), 'and it is marked as session-only');
      Assert.AreEqual('True', Persisted,
        'with the value the registry still holds');

      // THE POINT: a Save from anywhere else must not persist it.
      TPluginSettings.Save;
      // NOT OpenKeyReadOnly: it sets TRegistry.FAccess to KEY_READ and
      // never puts it back (read in System.Win.Registry), so every later
      // OpenKey on this object is read-only and the restore below would
      // raise "Fehler beim Setzen der Daten".
      Assert.IsTrue(Reg.OpenKey(Key, False), 'read the key back');
      try
        Assert.IsTrue(Reg.ReadBool('BlameUseTortoise'),
          'the registry keeps what it had - a session-only value is never ' +
          'written, whoever saves');
      finally
        Reg.CloseKey;
      end;

      // Clearing the mark makes the current value ordinary again.
      TPluginSettings.MarkSessionOnly('BlameUseTortoise', False);
      Assert.IsFalse(TPluginSettings.IsSessionOnly('BlameUseTortoise',
        Persisted), 'the mark is gone');
      TPluginSettings.Save;
      Assert.IsTrue(Reg.OpenKey(Key, False));
      try
        Assert.IsFalse(Reg.ReadBool('BlameUseTortoise'),
          'and now the value is stored');
      finally
        Reg.CloseKey;
      end;
    finally
      // the user's own value, whatever this test did
      TPluginSettings.MarkSessionOnly('BlameUseTortoise', False);
      if Reg.OpenKey(Key, True) then
      try
        if Had then Reg.WriteBool('BlameUseTortoise', Old)
        else if Reg.ValueExists('BlameUseTortoise') then
          Reg.DeleteValue('BlameUseTortoise');
      finally
        Reg.CloseKey;
      end;
      TPluginSettings.Load;
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

procedure TInterfaceDeclLineTests.BlanksInsideTheGuidBrackets_AreStillAGuid;
var
  F: string;
  E: TArray<TInterfaceGuidEntry>;
begin
  F := TPath.Combine(TPath.GetTempPath, 'RLGuidBlanks_' + TGUID.NewGuid.ToString + '.pas');
  TFile.WriteAllText(F,
    'unit UBlanks;' + NL + 'interface' + NL + 'type' + NL +
    '  ILeft = interface' + NL +
    '    [ ''{11111111-2222-3333-4444-000000000001}'']' + NL +
    '  end;' + NL +
    '  IBoth = interface' + NL +
    '    [' + #9 + '''{11111111-2222-3333-4444-000000000002}'' ]' + NL +
    '  end;' + NL +
    '  IInline = interface(IInterface) [ ''{11111111-2222-3333-4444-000000000001}'' ]' + NL +
    '  end;' + NL +
    '  INotAGuid = interface' + NL +
    '    procedure Foo(const A: array of string); // [ ''{11111111-2222-3333-4444-000000000003}'' ]' + NL +
    '  end;' + NL +
    'implementation' + NL + 'end.' + NL, TEncoding.UTF8);
  try
    E := TInterfaceGuidChecker.Scan([F]);
  finally
    TFile.Delete(F);
  end;
  Assert.AreEqual(4, Integer(Length(E)), 'four declarations');
  Assert.AreEqual('{11111111-2222-3333-4444-000000000001}', E[0].Guid,
    'a blank after the bracket');
  Assert.AreEqual('{11111111-2222-3333-4444-000000000002}', E[1].Guid,
    'a tab after the bracket, a blank before the closing one');
  Assert.AreEqual('{11111111-2222-3333-4444-000000000001}', E[2].Guid,
    'blanks on the declaration line');
  Assert.IsTrue(E[0].IsDuplicate and E[2].IsDuplicate,
    'the two with one GUID are duplicates whatever the blanks');
  Assert.IsFalse(E[3].HasGuid, 'a GUID in a comment is still none: ' + E[3].Guid);
end;

procedure TInterfaceDeclLineTests.AFreshGuidGoesWhereTheCheckFoundTheOldOne;
const
  G1 = '{11111111-2222-3333-4444-000000000001}';
  NewG = '{AABBCCDD-1122-3344-5566-778899AABBCC}';
var
  Lines: TArray<string>;
  Plan: TGuidEditPlan;
begin
  // Only the text of a GUID may ever reach a source file.
  Assert.IsTrue(IsGuidText(G1), 'the shape Delphi accepts');
  Assert.IsTrue(IsGuidText('{aabbccdd-1122-3344-5566-778899aabbcc}'), 'lower case');
  Assert.IsFalse(IsGuidText(''), 'nothing');
  Assert.IsFalse(IsGuidText('11111111-2222-3333-4444-000000000001'), 'no braces');
  Assert.IsFalse(IsGuidText('{11111111-2222-3333-4444-00000000000G}'), 'no hex');
  Assert.IsFalse(IsGuidText('{1111111-22222-3333-4444-000000000001}'), 'group widths');

  // 1) The GUID sits on a line of its OWN below the declaration, which is
  //    where the check's three-line look-ahead finds it - so that is the
  //    line the fix must rewrite, and nothing else may move.
  Lines := SplitContentLines(
    'unit UFix;'#13#10 + 'interface'#13#10 + 'type'#13#10 +
    '  ICopy = interface'#13#10 +
    '    [''' + G1 + ''']'#13#10 +
    '    procedure Foo;'#13#10 +
    '  end;'#13#10 + 'implementation'#13#10 + 'end.'#13#10);
  Plan := PlanInterfaceGuidEdit(Lines, 3, NewG);
  Assert.IsTrue(Plan.Ok, 'the declaration is at line 4: ' + Plan.Problem);
  Assert.AreEqual(4, Plan.Line, 'the GUID line, not the declaration');
  Assert.IsFalse(Plan.Inserted, 'it had one - it is replaced');
  Assert.AreEqual(G1, Plan.OldGuid, 'what it replaced is reported');
  Assert.AreEqual('    [''' + NewG + ''']', Plan.Lines[4],
    'the brackets, quotes and indentation stay as the author wrote them');
  Assert.AreEqual(Lines[3], Plan.Lines[3], 'the declaration is untouched');
  Assert.AreEqual(Integer(Length(Lines)), Integer(Length(Plan.Lines)),
    'replacing adds no line');

  // 2) No GUID: a line is INSERTED below the declaration - which is also
  //    right for a parent list, where appending to the declaration line
  //    would have to know where the brackets end.
  Lines := SplitContentLines(
    'unit UFix;'#13#10 + 'interface'#13#10 + 'type'#13#10 +
    '  INone = interface(IInterface)'#13#10 +
    '    procedure Foo;'#13#10 +
    '  end;'#13#10 + 'implementation'#13#10 + 'end.'#13#10);
  Plan := PlanInterfaceGuidEdit(Lines, 3, NewG);
  Assert.IsTrue(Plan.Ok, Plan.Problem);
  Assert.IsTrue(Plan.Inserted, 'there was none');
  Assert.AreEqual('', Plan.OldGuid, 'nothing was replaced');
  Assert.AreEqual(4, Plan.Line);
  Assert.AreEqual('    [''' + NewG + ''']', Plan.Lines[4],
    'indented like the member below it');
  Assert.AreEqual('  INone = interface(IInterface)', Plan.Lines[3],
    'the parent list is untouched');
  Assert.AreEqual('    procedure Foo;', Plan.Lines[5], 'and nothing is lost');

  // 3) The shapes a human has to decide: nothing is written and the reason
  //    says which one it is.
  Lines := SplitContentLines(
    'unit UFix;'#13#10 + 'interface'#13#10 + 'type'#13#10 +
    '  IOneLine = interface end;'#13#10 +
    '  IReal = interface'#13#10 + '  end;'#13#10 +
    'implementation'#13#10 + 'end.'#13#10);
  Plan := PlanInterfaceGuidEdit(Lines, 3, NewG);
  Assert.IsFalse(Plan.Ok, 'opened and closed on one line');
  Assert.IsTrue(Plan.Problem.Contains('one line'), Plan.Problem);
  Plan := PlanInterfaceGuidEdit(Lines, 2, NewG);
  Assert.IsFalse(Plan.Ok, '"type" declares no interface');
  Assert.IsTrue(Plan.Problem.Contains('does not declare'), Plan.Problem);
  Plan := PlanInterfaceGuidEdit(Lines, 4, 'not-a-guid');
  Assert.IsFalse(Plan.Ok, 'and a GUID that is none is never written');

  // 4) By NAME in the current text, because any edit since the check moves
  //    the line - and TWO answers mean a conditional declaration, where the
  //    caller must refuse instead of picking one.
  Lines := SplitContentLines(
    'unit UFix;'#13#10 + 'interface'#13#10 + 'type'#13#10 +
    '  IPlain = interface'#13#10 + '  end;'#13#10 +
    '{$IFDEF A}'#13#10 +
    '  ITwice = interface'#13#10 + '  end;'#13#10 +
    '{$ELSE}'#13#10 +
    '  ITwice = interface'#13#10 + '  end;'#13#10 +
    '{$ENDIF}'#13#10 + 'implementation'#13#10 + 'end.'#13#10);
  Assert.AreEqual(1, Integer(Length(InterfaceDeclLines(Lines, 'IPlain'))),
    'one declaration');
  Assert.AreEqual(3, InterfaceDeclLines(Lines, 'IPlain')[0], 'at its line');
  Assert.AreEqual(2, Integer(Length(InterfaceDeclLines(Lines, 'ITwice'))),
    'a conditional declaration is TWO - the caller refuses');
  Assert.AreEqual(0, Integer(Length(InterfaceDeclLines(Lines, 'INothing'))),
    'and a name that is not declared is none');

  // The tool reads "apply" and "fix_ids", so its SCHEMA has to list them:
  // an argument a tool does not declare is dropped silently (1.16.11), and
  // "apply without fix_ids" would then look like "apply everything".
  var Known := KnownToolArguments('interface_guids');
  Assert.IsTrue(Length(Known) > 0, 'interface_guids is in the shipped list');
  for var N in ['only_problems', 'fix_ids', 'apply', 'instance'] do
    Assert.IsTrue(MatchText(N, Known), N + ' is an argument of interface_guids');
  Assert.IsFalse(MatchText('token', Known),
    'the analysis re-runs inside the apply, so there is no revision to pin');
  var Def := '';
  var Arr := McpToolDefinitions;
  try
    for var E in Arr do
      if (E is TJSONObject) and
         (TJSONObject(E).GetValue<string>('name', '') = 'interface_guids') then
        Def := TJSONObject(E).GetValue<string>('description', '');
  finally
    Arr.Free;
  end;
  Assert.IsFalse(Def.Contains('Read-only'),
    'it writes now - a description that still says read-only is a lie');
  Assert.IsTrue(Def.Contains('leave out'),
    'and it has to say which declaration keeps the GUID');
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

procedure TQuickFixPreviewTests.RemovePrivate_TakesTheDocCommentAlong;
const
  Src =
    'unit U;'#13#10 +                                        // 0
    'interface'#13#10 +                                      // 1
    'type'#13#10 +                                           // 2
    '  TFoo = class'#13#10 +                                 // 3
    '  private'#13#10 +                                      // 4
    '    /// <summary>Kept: it documents Other.</summary>'#13#10 +  // 5
    '    procedure Other;'#13#10 +                           // 6
    '    /// <summary>What Bar does.</summary>'#13#10 +      // 7
    '    /// <remarks>Over two lines.</remarks>'#13#10 +     // 8
    '    procedure Bar;'#13#10 +                             // 9
    '  end;'#13#10 +                                         // 10
    'implementation'#13#10 +                                 // 11
    '/// <summary>The body of Bar.</summary>'#13#10 +        // 12
    'procedure TFoo.Bar;'#13#10 +                            // 13
    'begin'#13#10 +                                          // 14
    'end;'#13#10 +                                           // 15
    'procedure TFoo.Other;'#13#10 +                          // 16
    'begin'#13#10 +                                          // 17
    'end;'#13#10 +                                           // 18
    'end.'#13#10;
var
  Info: TPrivateMember;
begin
  Assert.IsTrue(FindRemovablePrivate(Src, 9, 'Bar', Info), 'Bar is removable');
  var After := RemovePrivateMemberText(Src, Info);

  Assert.IsFalse(After.Contains('What Bar does'),
    'the doc of the declaration goes with it');
  Assert.IsFalse(After.Contains('Over two lines'),
    'the whole block, not just its last line');
  Assert.IsFalse(After.Contains('The body of Bar'),
    'and the doc above the implementation too');
  Assert.IsFalse(After.Contains('procedure Bar;'), 'the declaration itself');
  Assert.IsFalse(After.Contains('procedure TFoo.Bar;'), 'and its body');

  // ...and nothing that belongs to the member next door
  Assert.IsTrue(After.Contains('it documents Other'),
    'the doc of the OTHER member stays');
  Assert.IsTrue(After.Contains('procedure Other;'), 'and so does it');
  Assert.IsTrue(After.Contains('procedure TFoo.Other;'), 'with its body');

  // MEASURED LIVE (2026-10-26): the removal left the unit ONE LINE
  // LONGER - the join went through TStringList.Text, which appends a break
  // of its own, so the trailing empty element SplitContentLines keeps for
  // the file's final break became a blank line at the end. The region
  // between the two neighbours was always right; the END was not.
  Assert.IsTrue(After.EndsWith('end.'#13#10),
    'the unit still ends exactly as it did: ' + After.Substring(After.Length - 12).Replace(#13, '<CR>').Replace(#10, '<LF>'));
  Assert.IsFalse(After.EndsWith('end.'#13#10#13#10),
    'and NOT with a blank line the removal left behind');
  Assert.AreEqual(Integer(Length(SplitContentLines(Src)) - 7),
    Integer(Length(SplitContentLines(After))),
    'exactly the 7 removed lines (2 doc + declaration, 1 doc + 3 body)');

  // ...and the file's own line break survives: this quick fix must not
  // turn an LF-only unit into a CRLF one (the whole-file fallback of
  // ApplyLinesMinimal writes exactly this text).
  var Lf := Src.Replace(#13#10, #10);
  Assert.IsTrue(FindRemovablePrivate(Lf, 9, 'Bar', Info), 'the same member, LF only');
  var AfterLf := RemovePrivateMemberText(Lf, Info);
  Assert.AreEqual(0, Integer(Pos(#13, AfterLf)), 'no CR was invented');
  Assert.IsTrue(AfterLf.EndsWith('end.'#10), 'and the final LF is still one');

  // the rule on its own, including "there is none"
  var L := SplitContentLines(Src);
  Assert.AreEqual(7, DocCommentStart(L, 9), 'two doc lines above the declaration');
  Assert.AreEqual(16, DocCommentStart(L, 16), 'a member without one is itself');
  Assert.AreEqual(0, DocCommentStart(L, 0), 'the first line has nothing above it');
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

procedure TRenameGuardTests.AParameterIsDeclaredTwiceAndBothMustFollow;
const
  // the reporter's shape, plus a wrapped list and a free routine
  Src =
    'unit U;'#10 +                                          // 0
    'interface'#10 +                                        // 1
    'type'#10 +                                             // 2
    '  TForm1 = class(TObject)'#10 +                        // 3
    '  public'#10 +                                         // 4
    '    procedure SetTitle(const Title: string);'#10 +     // 5
    '    function Wrapped(const A: Integer;'#10 +           // 6
    '      const Title: string): Boolean;'#10 +             // 7
    '  end;'#10 +                                           // 8
    'procedure Free1(const Title: string);'#10 +            // 9
    'implementation'#10 +                                   // 10
    'procedure TForm1.SetTitle(const Title: string);'#10 +  // 11
    'begin'#10 +                                            // 12
    'end;'#10 +                                             // 13
    'function TForm1.Wrapped(const A: Integer;'#10 +        // 14
    '  const Title: string): Boolean;'#10 +                 // 15
    'begin'#10 +                                            // 16
    'end;'#10 +                                             // 17
    'procedure Free1(const Title: string);'#10 +            // 18
    'begin'#10 +                                            // 19
    'end;'#10 +                                             // 20
    'end.';                                                 // 21
  // Overloads: the partner is the declaration that carries that parameter
  // AT THAT PLACE. Both of these declare A and B - at swapped places - so
  // "the other declaration mentions the name somewhere" is not enough.
  Ovl =
    'unit V;'#10 +                                                     // 0
    'interface'#10 +                                                   // 1
    'type'#10 +                                                        // 2
    '  TC = class'#10 +                                                // 3
    '    procedure Init(const B: string; const A: Integer); overload;'#10 + // 4
    '    procedure Init(const A: Integer; const B: string); overload;'#10 + // 5
    '  end;'#10 +                                                      // 6
    'implementation'#10 +                                              // 7
    'procedure TC.Init(const B: string; const A: Integer);'#10 +       // 8
    'begin'#10 +                                                       // 9
    'end;'#10 +                                                        // 10
    'procedure TC.Init(const A: Integer; const B: string);'#10 +       // 11
    'begin'#10 +                                                       // 12
    'end;'#10 +                                                        // 13
    'end.';                                                            // 14
var
  P: Integer;
begin
  var Lines := string(Src).Split([#10]);

  // THE REPORTED CASE: from the implementation header to the declaration
  Assert.IsTrue(FindParameterPartnerLine(Lines, 11, 'Title', P),
    'the class declaration carries the same parameter');
  Assert.AreEqual(5, P, 'and that is the line the rename also has to change');
  // ...and the same rename started on the declaration
  Assert.IsTrue(FindParameterPartnerLine(Lines, 5, 'Title', P), 'the other way');
  Assert.AreEqual(11, P);

  // a WRAPPED list: the partner's name is not on its first line
  Assert.IsTrue(FindParameterPartnerLine(Lines, 14, 'Title', P), 'wrapped');
  Assert.AreEqual(7, P, 'the continuation line, not the header line');
  Assert.IsTrue(FindParameterPartnerLine(Lines, 6, 'Title', P), 'and back');
  Assert.AreEqual(15, P);
  // ...and the anchor itself may BE a continuation line: that is what
  // DelphiLSP answers for a wrapped header (1.15.4)
  Assert.IsTrue(FindParameterPartnerLine(Lines, 7, 'Title', P),
    'a continuation line belongs to its header');
  Assert.AreEqual(15, P);
  Assert.IsTrue(FindParameterPartnerLine(Lines, 15, 'Title', P), 'the same there');
  Assert.AreEqual(7, P);

  // a FREE routine declared in the interface section - same rule
  Assert.IsTrue(FindParameterPartnerLine(Lines, 18, 'Title', P), 'free routine');
  Assert.AreEqual(9, P);
  Assert.IsTrue(FindParameterPartnerLine(Lines, 9, 'Title', P), 'and back');
  Assert.AreEqual(18, P);

  // what is NOT a parameter of that header
  Assert.IsFalse(FindParameterPartnerLine(Lines, 11, 'SetTitle', P),
    'the routine''s own name is renamed through its own partner query');
  Assert.IsFalse(FindParameterPartnerLine(Lines, 12, 'Title', P),
    '"begin" is no header');
  Assert.IsFalse(FindParameterPartnerLine(Lines, 11, 'Caption', P),
    'a name the header does not declare');

  // OVERLOADS: only the declaration that carries that parameter at that
  // place - the other one is a different routine
  var OL := string(Ovl).Split([#10]);
  Assert.IsTrue(FindParameterPartnerLine(OL, 11, 'B', P), 'B of the second overload');
  Assert.AreEqual(5, P,
    'the declaration that has B SECOND - line 4 has a B too, at the front');
  Assert.IsTrue(FindParameterPartnerLine(OL, 8, 'B', P), 'B of the first overload');
  Assert.AreEqual(4, P, 'and that one pairs the other way round');
  Assert.IsTrue(FindParameterPartnerLine(OL, 11, 'A', P), 'A of the second');
  Assert.AreEqual(5, P);

  // and the predicate the rename's anchor decision asks
  Assert.IsTrue(HeaderDeclaresParameter(Lines, 5, 'Title'),
    'a declaration declares its parameters - DelphiLSP answers nothing there');
  Assert.IsFalse(HeaderDeclaresParameter(Lines, 5, 'SetTitle'), 'its own name');
  Assert.IsFalse(HeaderDeclaresParameter(Lines, 4, 'Title'), '"public" is no header');
end;

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

procedure TPartnerQueryTests.IdentifierBefore_IsUnicodeAware;
var
  Col: Integer;
  Line: string;
begin
  // His two red-first cases. The '(' is what signature help asks about, so
  // APos is its position and the identifier ends right before it.
  Line := '  Größe(';
  Assert.AreEqual('Größe', IdentifierBefore(Line, Pos('(', Line), Col), False,
    'the walk back must not stop at the umlaut');
  Assert.AreEqual(3, Col);
  Line := '  Ärger(';
  Assert.AreEqual('Ärger', IdentifierBefore(Line, Pos('(', Line), Col), False,
    'nor at a leading one');
  // Plain ASCII, blanks before the bracket, and the cases with no name
  Line := '  Foo2 (';
  Assert.AreEqual('Foo2', IdentifierBefore(Line, Pos('(', Line), Col), False);
  Assert.AreEqual(3, Col);
  Assert.AreEqual('', IdentifierBefore('  (', 3, Col), False, 'nothing before it');
  Assert.AreEqual(0, Col);
  Assert.AreEqual('', IdentifierBefore('', 1, Col), False, 'an empty line');

  // The declaration end that extract interface and move to unit share: the
  // ';' inside a parameter list does not end the header (issue #44).
  var Depth := 0;
  Assert.IsFalse(LineEndsDeclaration('procedure Test3( const AParam1: string;', Depth),
    'the '';'' inside the open parameter list is not the end');
  Assert.AreEqual(1, Depth, 'and the depth is carried to the next line');
  Assert.IsTrue(LineEndsDeclaration('const AParam2: Boolean);', Depth),
    'the one after the closing bracket is');
  Depth := 0;
  Assert.IsTrue(LineEndsDeclaration('procedure Test2( const A: Boolean);', Depth),
    'a one-line header ends on its own line');

  Assert.AreEqual('String', UnescapedIdentifier('&String'), False);
  Assert.AreEqual('String', UnescapedIdentifier('String'), False, 'idempotent');
  Assert.AreEqual('', UnescapedIdentifier(''), False);
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
  Raised, RanOffMain, WasStrict: Boolean;
  T: TThread;
begin
  // 2.0 turned the counter into an exception, so the DEFAULT is part of
  // the contract now - a revert has to be a deliberate edit, never a
  // drift. The evidence for the flip is in the declaration's own doc.
  WasStrict := StrictMainThread;
  Assert.IsTrue(WasStrict, 'StrictMainThread is ON by default since 2.0');
  try

  // Both halves set the mode they are about; whatever the default is,
  // this test puts it back (it used to end on a hard False, which from
  // 2.0 on would have switched the guard off for the rest of the run).
  StrictMainThread := False;

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

  // With the strict switch it is fatal instead of merely counted - which
  // is what the plugin ships with since 2.0.
  Raised := False;
  RanOffMain := False;
  StrictMainThread := True;
  begin
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
  end;
  Assert.IsTrue(Raised, 'strict mode raises EOffMainThread');
  Assert.IsFalse(RanOffMain, 'and the write never happens');

  finally
    StrictMainThread := WasStrict;
  end;
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


procedure TSelfProtectionTests.AModuleTheCallOpenedIsNotLeftBehind;
begin
  // The call opened it and nobody can see it: that is ours to settle.
  Assert.AreEqual(Integer(Ord(msvSaveAndClose)),
    Integer(Ord(ModuleSweepVerdict(False, False, False, True))),
    'a module the call opened and wrote goes to disk');
  Assert.AreEqual(Integer(Ord(msvDiscard)),
    Integer(Ord(ModuleSweepVerdict(False, False, False, False))),
    'a failed call leaves nothing behind');

  // Ownership beats everything: the IDE had it, so rename's "changed but
  // not saved, undoable" must stay exactly that.
  Assert.AreEqual(Integer(Ord(msvLeave)),
    Integer(Ord(ModuleSweepVerdict(True, False, False, True))),
    'a module the IDE already had is never touched');
  // buffer_open promised a headless buffer that SURVIVES the call.
  Assert.AreEqual(Integer(Ord(msvLeave)),
    Integer(Ord(ModuleSweepVerdict(False, True, False, True))),
    'a bridge-opened headless buffer stays open');
  // And a tab the user can see is the user's, whatever we think we opened.
  Assert.AreEqual(Integer(Ord(msvLeave)),
    Integer(Ord(ModuleSweepVerdict(False, False, True, True))),
    'a module with a view belongs to the user');
  Assert.AreEqual(Integer(Ord(msvLeave)),
    Integer(Ord(ModuleSweepVerdict(True, True, True, False))),
    'ownership wins over a failed call too');
end;

procedure TSelfProtectionTests.AFixIdNamesItsOwnListing;
var
  H: Cardinal;
  Gen, Idx: Integer;
begin
  // The id says which listing it came from.
  Assert.AreEqual('0000002A-7-3', MakeFixId($2A, 7, 3), 'the id format');
  Assert.IsTrue(ParseFixId('0000002A-7-3', H, Gen, Idx), 'and it parses back');
  Assert.AreEqual(Cardinal($2A), H, 'the buffer state');
  Assert.AreEqual(7, Gen, 'the generation');
  Assert.AreEqual(3, Idx, 'the index');

  // The OLDER two-part form still parses and means "the newest listing of
  // this buffer state" - a client that kept an id must not break.
  Assert.IsTrue(ParseFixId('0000002A-3', H, Gen, Idx), 'the old form');
  Assert.AreEqual(Cardinal($2A), H);
  Assert.AreEqual(-1, Gen, 'no generation = the newest listing');
  Assert.AreEqual(3, Idx);

  // Nonsense is refused rather than read as something.
  Assert.IsFalse(ParseFixId('', H, Gen, Idx), 'empty');
  Assert.IsFalse(ParseFixId('nothex00-1-2', H, Gen, Idx), 'no hash');
  Assert.IsFalse(ParseFixId('0000002A-1-x', H, Gen, Idx), 'no index');
  Assert.IsFalse(ParseFixId('0000002A--1', H, Gen, Idx), 'a negative generation');
end;

procedure TSelfProtectionTests.AnApplyNoteSaysWhatReallyHappened;
const
  Thing = 'declaration(s)';
  Done = 'carry a fresh GUID now';
  Where = 'in the IDE buffer when the file was open (not saved)';
begin
  // The same rule for what an answer OFFERS. Measured live (2.5.1):
  // with a complete DCU output the check found no such finding at all
  // and still ended with "A "LSP reads a DCU" finding can be fixed
  // from here" - and since 2.5.0 only a stale .dcu whose source the
  // compiler can reach instead is movable, so the sentence was wrong
  // in both directions.
  Assert.AreEqual('', DebugFixOfferNote(0),
    'nothing fixable means the answer offers nothing');
  Assert.AreEqual('', DebugFixOfferNote(-1), 'and a bad count too');
  var Offer := DebugFixOfferNote(3);
  Assert.IsTrue(Offer.StartsWith('3 finding(s)'),
    'and with three of them it says how many: ' + Offer);
  Assert.IsTrue(Offer.Contains('older than its source'),
    'naming WHICH finding has that cure: ' + Offer);
  Assert.IsTrue(Offer.Contains('is a build'),
    'and that the others have another one: ' + Offer);

  // The measured case: apply=true with one id that matches no finding.
  // Nothing happened, so the note may not say the opposite.
  var Nothing := AppliedByIdsNote(0, 0, 1, Thing, Done, Where);
  Assert.IsTrue(Nothing.StartsWith('NOTHING was changed'),
    'an apply that matched nothing says so first: ' + Nothing);
  Assert.IsFalse(Nothing.Contains(Done),
    'and it does not claim the declarations were done: ' + Nothing);
  Assert.IsTrue(Nothing.Contains('unknownIds'),
    'while the id that matched nothing is pointed at: ' + Nothing);

  // The ordinary case: two done, and the counts are the scan BEFORE.
  var Two := AppliedByIdsNote(2, 0, 0, Thing, Done, Where);
  Assert.IsTrue(Two.StartsWith('2 declaration(s) carry a fresh GUID now'),
    'what happened, counted: ' + Two);
  Assert.IsTrue(Two.Contains('BEFORE'),
    'and that the rows describe the state before it: ' + Two);
  Assert.IsFalse(Two.Contains('could NOT'), 'nothing failed: ' + Two);

  // A partial apply names both halves, and an apply in which every named
  // entry failed is not success with a footnote.
  var Mixed := AppliedByIdsNote(1, 2, 0, Thing, Done, Where);
  Assert.IsTrue(Mixed.StartsWith('1 declaration(s)'), 'the one that worked');
  Assert.IsTrue(Mixed.Contains('2 could NOT be done'),
    'and the two that did not: ' + Mixed);

  var AllFailed := AppliedByIdsNote(0, 3, 0, Thing, Done, Where);
  Assert.IsTrue(AllFailed.StartsWith('NOTHING was changed'),
    'three named, three failed: ' + AllFailed);
  Assert.IsTrue(AllFailed.Contains('3 declaration(s)'),
    'and it says how many were named: ' + AllFailed);
end;

procedure TSelfProtectionTests.AnAnswerCarriesNoKeyTwice;
var
  O: TJSONObject;
  Notes: Integer;
begin
  O := TJSONObject.Create;
  try
    // The shape of the report: two code paths both have something to say.
    SetJsonNote(O, 'DelphiLSP returned nothing for this position.');
    SetJsonNote(O, 'Unit1.pas was (re)sent 2 s ago - retry shortly.');

    Notes := 0;
    for var I := 0 to O.Count - 1 do
      if O.Pairs[I].JsonString.Value = 'note' then Inc(Notes);
    Assert.AreEqual(1, Notes,
      'ONE "note" key, whatever how many places write it: ' + O.ToJSON);
    Assert.AreEqual('Unit1.pas was (re)sent 2 s ago - retry shortly.',
      O.GetValue<string>('note'), 'and it is the one that was set last');

    // The first write still has to work on an object that has no note yet.
    var P := TJSONObject.Create;
    try
      SetJsonNote(P, 'only one');
      Assert.AreEqual('only one', P.GetValue<string>('note'), 'the first note');
    finally
      P.Free;
    end;
  finally
    O.Free;
  end;
end;


{ TLspConfigHintTests }

procedure TSelfProtectionTests.AQueuedClosureNeverLetsAnExceptionEscape;
var
  Ran: Integer;
begin
  // The counter is global, so measure the DIFFERENCE - another fixture may
  // have queued something.
  var Before := QueuedErrorCount;
  Ran := 0;

  QueueSafe(
    procedure
    begin
      Inc(Ran);
      raise Exception.Create('a queued closure that raises');
    end);
  ForceQueueSafe(
    procedure
    begin
      Inc(Ran);
    end);

  // This is the main thread, so CheckSynchronize is what runs them - the
  // same call the IDE makes from inside its transactions.
  for var I := 1 to 20 do
    if not CheckSynchronize(0) then Break;

  Assert.AreEqual(2, Ran, 'both closures ran');
  Assert.AreEqual(Before + 1, QueuedErrorCount,
    'the one that raised is counted, and nothing escaped to the caller');
  Assert.IsTrue(LastQueuedError.Contains('a queued closure that raises'),
    'and it is NAMED - swallowing is only defensible with a channel: ' +
    LastQueuedError);
end;

procedure TLspConfigHintTests.TheHintNamesTheFileAndTheSwitch;
begin
  var Path := LspConfigOptionPath;
  // Every level of the path, in order - the "Delphi" level is the one the
  // three wizards that spelled it themselves left out.
  var Last := 0;
  for var Part in ['Tools', 'Options', 'Editor', 'Language', 'Delphi',
    'Code Insight', 'Generate LSP Config'] do
  begin
    var P := Pos(Part, Path);
    Assert.IsTrue(P > Last,
      Format('"%s" must follow the level before it in: %s', [Part, Path]));
    Last := P;
  end;

  var Short := LspConfigMissingHint;
  Assert.IsTrue(Pos('.delphilsp.json', Short) > 0, 'it names the file');
  Assert.IsTrue(Pos(Path, Short) > 0, 'and the option path');
  Assert.AreEqual(0, Integer(Pos(#13, Short) + Pos(#10, Short)),
    'the short form goes into a status bar - one line');

  var Long := LspConfigMissingHintLong;
  Assert.IsTrue(Pos(Short, Long) > 0, 'the long form starts from the short one');
  Assert.IsTrue(Pos('LSP-Konfiguration generieren', Long) > 0,
    'and names the option as a German IDE spells it');
  Assert.IsTrue(Pos('.dproj', Long) > 0, 'and where the file appears');
end;

procedure TLspConfigHintTests.TheOptionPathIsSpelledInOnePlaceOnly;
var
  Dir: string;
  Spelt: TArray<string>;
begin
  // The point of the shared text: a second copy is free to drift, and the
  // three that existed had drifted (no "Delphi" level). "Generate LSP
  // Config" is the option's own name, so it cannot appear by accident.
  Dir := ExtractFilePath(ParamStr(0));
  for var Up := 1 to 6 do
  begin
    Dir := ExtractFilePath(ExcludeTrailingPathDelimiter(Dir));
    if TDirectory.Exists(TPath.Combine(Dir, 'Source')) then Break;
  end;
  Dir := TPath.Combine(Dir, 'Source');
  Assert.IsTrue(TDirectory.Exists(Dir),
    'the repository Source folder must be reachable from ' + ParamStr(0));
  var Checked := 0;
  for var Mask in ['*.pas', '*.inc'] do
    for var F in TDirectory.GetFiles(Dir, Mask) do
    begin
      var Content := '';
      try
        Content := ReadDelphiFile(F);
      except
        Continue;
      end;
      Inc(Checked);
      if Pos('Generate LSP Config', Content) > 0 then
        Spelt := Spelt + [ExtractFileName(F)];
    end;
  Assert.IsTrue(Checked > 50,
    Format('the sweep must really read the units (%d)', [Checked]));
  Assert.AreEqual('Expert.EditorHelperIntf.pas', string.Join(', ', Spelt),
    'exactly one unit may spell the option out');
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




{ TExtractAndCompletionTests }

procedure TExtractAndCompletionTests.EveryWriteInTheBlockIsFound;
var
  Written: TArray<string>;

  function Has(const AName: string): Boolean;
  begin
    Result := False;
    for var S in Written do
      if SameText(S, AName) then Exit(True);
  end;

begin
  // Only "^identifier :=" counted, so these three made N / I / S a CONST
  // parameter and the extracted method did not compile.
  Written := WrittenIdentifiersIn([
    '  Inc(N);',
    '  for I := 0 to 9 do',
    '    ReadLn(S);',
    '  Total := Total + 1;',
    '  Other.Field := 2;',       // a member, not a local
    '  if Flag then Count := 1;' // an assignment that does not start the line
  ]);
  Assert.IsTrue(Has('N'), 'Inc(N) writes N');
  Assert.IsTrue(Has('I'), 'the for-loop variable is written');
  Assert.IsTrue(Has('S'), 'ReadLn(S) writes S');
  Assert.IsTrue(Has('Total'), 'a plain assignment');
  Assert.IsTrue(Has('Count'), 'an assignment after "then"');
  Assert.IsFalse(Has('Flag'), 'a condition is only read');

  // Nothing inside a comment or a string counts.
  Written := WrittenIdentifiersIn([
    '  // Hidden := 1;',
    '  S := ''Quoted := 2'';',
    '  { Blocked := 3 }'
  ]);
  Assert.IsTrue(Has('S'), 'the real assignment');
  Assert.IsFalse(Has('Hidden'), 'a // comment');
  Assert.IsFalse(Has('Quoted'), 'a string literal');
  Assert.IsFalse(Has('Blocked'), 'a block comment');
end;

procedure TExtractAndCompletionTests.AGeneratedParameterNameIsNoReservedWord;
begin
  // The reported case: a string variable called S.
  Assert.AreEqual('AS1', ExtractedParameterName('S'),
    'A + S is "as", the operator');
  // The same shape with the other words that start with A.
  Assert.AreEqual('And1', ExtractedParameterName('nd'));
  Assert.AreEqual('Array1', ExtractedParameterName('rray'));
  Assert.AreEqual('Asm1', ExtractedParameterName('sm'));
  // A variable ALREADY named AS takes the keep-branch, so the check has to
  // sit after both branches - not inside the prefixing one.
  Assert.AreEqual('AS1', ExtractedParameterName('AS'),
    'already prefixed, and still a reserved word');
  // Everything else keeps the plain convention.
  Assert.AreEqual('ACount', ExtractedParameterName('Count'));
  Assert.AreEqual('AList', ExtractedParameterName('List'));
  Assert.AreEqual('AValue', ExtractedParameterName('AValue'),
    'an A-prefixed name is kept');
  Assert.AreEqual('Aindex', ExtractedParameterName('index'),
    'lower case after the A is not the keep-branch');
end;

procedure TExtractAndCompletionTests.AMethodNameMustBeAnIdentifier;
begin
  Assert.AreEqual('', MethodNameProblem('DoSomething'), 'a plain name');
  Assert.AreEqual('', MethodNameProblem('_Do2'), 'digits and underscore');
  Assert.IsTrue(MethodNameProblem('') <> '', 'empty');
  Assert.IsTrue(ContainsText(MethodNameProblem('2Foo'), 'identifier'),
    '2Foo is no identifier');
  Assert.IsTrue(MethodNameProblem('My Method') <> '', 'a blank inside');
  Assert.IsTrue(ContainsText(MethodNameProblem('begin'), 'reserved'),
    'begin is a reserved word');
  Assert.IsTrue(ContainsText(MethodNameProblem('BEGIN'), 'reserved'),
    '... whatever its case');
end;

procedure TExtractAndCompletionTests.TheExpressionIsMatchedAsCode;
const
  LineA = '  X := ''Count: '' + IntToStr(Count);';
  LineB = '  Y := MaxIdx + Idx;';
begin
  // The string literal must not win over the real occurrence.
  var P := CodeOccurrenceOf(LineA, 'Count', 0);
  Assert.IsTrue(P > 0, 'there IS a code occurrence');
  Assert.AreEqual('Count);', Copy(LineA, P, 7), 'and it is the argument');
  // "Idx" must not match inside "MaxIdx".
  P := CodeOccurrenceOf(LineB, 'Idx', 0);
  Assert.IsTrue(P > 0, 'Idx occurs on its own too');
  Assert.AreEqual(' Idx;', Copy(LineB, P - 1, 5), 'the standalone one');
  // Nothing to find at all.
  Assert.AreEqual(0, CodeOccurrenceOf('  // Count', 'Count', 0),
    'a comment is no occurrence');
  Assert.AreEqual(0, CodeOccurrenceOf('  S := ''Count'';', 'Count', 0),
    'a string is no occurrence');
  // An expression that is not an identifier keeps working.
  Assert.IsTrue(CodeOccurrenceOf('  X := A + B;', 'A + B', 0) > 0,
    'a composed expression');
end;

procedure TExtractAndCompletionTests.AMemberIsNoFreeIdentifierOfTheBlock;

  function NamesOf(const ABlock: string): string;
  begin
    Result := '|';
    for var T in ExtractBlockTokens(ABlock, 10, 1) do
      Result := Result + T.Ident + '|';
  end;

begin
  // the reporter's line, verbatim
  var Names := NamesOf(
    'AWorkSheet.Cell[''A'' + LineStr].AsString := RemoveAcceleratorKey(FFieldName.AsString);'#10 +
    'AWorkSheet.Cell[''B'' + LineStr].AsFloat := Total;');

  // what the block really needs from around it
  Assert.IsTrue(Names.Contains('|AWorkSheet|'), 'the sheet: ' + Names);
  Assert.IsTrue(Names.Contains('|LineStr|'),
    'the local the reported method was NOT given: ' + Names);
  Assert.IsTrue(Names.Contains('|RemoveAcceleratorKey|'), 'a routine it calls');
  Assert.IsTrue(Names.Contains('|FFieldName|'), 'a field it reads');
  Assert.IsTrue(Names.Contains('|Total|'), 'and the one on the right');

  // ...and what belongs to a type rather than to the block
  Assert.IsFalse(Names.Contains('|Cell|'),
    'the indexed property became a parameter: ' + Names);
  Assert.IsFalse(Names.Contains('|AsString|'), 'a property of the cell');
  Assert.IsFalse(Names.Contains('|AsFloat|'),
    'the reported "var AsFloat: string;" that nothing uses');

  // A name that is a MEMBER here and a free identifier three characters on
  // must still be found - and at the FREE occurrence, or the LSP is asked
  // about the member again.
  var Toks := ExtractBlockTokens('Foo.Value := Value;', 10, 1);
  var Col := -1;
  for var T in Toks do
    if SameText(T.Ident, 'Value') then Col := T.Col0;
  Assert.AreEqual(13, Col, 'the second Value, not the member at column 5');

  // A RANGE is no member access: "cA..cZ" are two free identifiers.
  Names := NamesOf('if C in [cA..cZ] then Hit := True;');
  Assert.IsTrue(Names.Contains('|cA|') and Names.Contains('|cZ|'),
    'a range is not a member access: ' + Names);

  // A qualified unit reference keeps its first part only - that one is
  // classified as a unit and dropped, the rest never costs a request.
  Names := NamesOf('System.SysUtils.Trim(S);');
  Assert.IsTrue(Names.Contains('|System|'), 'the first part: ' + Names);
  Assert.IsFalse(Names.Contains('|SysUtils|') or Names.Contains('|Trim|'),
    'what follows a dot is no free identifier: ' + Names);
end;

procedure TExtractAndCompletionTests.AWrittenOutTypeIsTheWholeType;
begin
  // The hovers DelphiLSP 37.0 really sends, measured 2026-10-21 - the
  // reported one with its markdown link and code fence around it.
  Assert.AreEqual('array [0..0] of UInt64',
    TypeTextFromHover('[ExDemo.pas:18](file:///C%3A/x/ExDemo.pas#18)'#13#10#13#10 +
      '```'#13#10'var Handles: array [0..0] of UInt64'#13#10'```'#13#10),
    'the reported case - it used to arrive as "array"');
  Assert.AreEqual('array of string', TypeTextFromHover('var Dyn: array of string'));
  Assert.AreEqual('set of Byte', TypeTextFromHover('var Bytes: set of Byte'));
  // the colon INSIDE the parameter list is not the one that separates
  Assert.AreEqual('procedure(X: Integer) of object',
    TypeTextFromHover('var Cb: procedure(X: Integer) of object'),
    'a procedural type carries a colon of its own');
  Assert.AreEqual('Integer', TypeTextFromHover('var Plain: Integer'),
    'the ordinary case is unchanged');
  Assert.AreEqual('TDictionary<string, Integer>',
    TypeTextFromHover('var D: TDictionary<string, Integer>'), 'generics too');
  Assert.AreEqual('Integer', TypeTextFromHover('const Max: Integer = 5'),
    'a default value belongs to the declaration, not to the type');
  Assert.AreEqual('', TypeTextFromHover('property TButton.OnClick'),
    'no colon, no type');

  // MEASURED with dcc32 37.0: only a NAMED type works as a parameter type.
  Assert.IsTrue(TypeIsNamed('Integer'));
  Assert.IsTrue(TypeIsNamed('System.UInt64'), 'dotted');
  Assert.IsTrue(TypeIsNamed('TDictionary<string, Integer>'), 'generic');
  Assert.IsFalse(TypeIsNamed('array [0..0] of UInt64'),
    'E2029 in a parameter list - and that is the reported variable');
  Assert.IsFalse(TypeIsNamed('array of string'),
    'it compiles there, but as an OPEN array - SetLength on it is E2008');
  Assert.IsFalse(TypeIsNamed('set of Byte'));
  Assert.IsFalse(TypeIsNamed('string[20]'));
  Assert.IsFalse(TypeIsNamed('^Integer'));
  Assert.IsFalse(TypeIsNamed('procedure(X: Integer) of object'));
  Assert.IsFalse(TypeIsNamed('record X: Integer; end'));
  Assert.IsFalse(TypeIsNamed(''), 'nothing is no name');

  // Issue #51: a name is a name by the LANGUAGE's rule, not by an ASCII
  // set. All three of these compile as a parameter type (measured,
  // dcc32 37.0 - scratchpad issue51/probe.py), so reading them as
  // "written out in place" refused an extraction for nothing.
  Assert.IsTrue(TypeIsNamed('TK' + #$00FC + 'ndigung'),
    'a non-ASCII letter is a letter');
  Assert.IsTrue(TypeIsNamed('Vertr' + #$00E4 + 'ge.T' + #$00C4 + 'nderung'),
    'in every dotted part');
  Assert.IsTrue(TypeIsNamed('&Type'), 'an escaped reserved word is a name');
  Assert.IsTrue(TypeIsNamed('Unit1.&Type'), 'qualified as well');
  Assert.IsTrue(TypeIsNamed('TList<TK' + #$00FC + 'ndigung>'),
    'and as a generic argument');
  // ...and what must stay refused: a name is not a sentence.
  Assert.IsFalse(TypeIsNamed('&'), 'an escape with nothing behind it');
  Assert.IsFalse(TypeIsNamed('A.'), 'a trailing dot is no part');
  Assert.IsFalse(TypeIsNamed('array of TK' + #$00FC + 'ndigung'),
    'the umlaut does not make an open array a name');
  Assert.IsFalse(TypeIsNamed('1Foo'), 'a digit cannot start one');

  // And the rule the type alone cannot decide (issue #51): a parameter
  // FORWARDED from the enclosing header already is what its declaration
  // says, so repeating it verbatim compiles - measured for const and var.
  // A local dynamic array promoted to a parameter is still refused: there
  // the declaration would change what the code means (E2008 on SetLength).
  Assert.IsTrue(ParamTypeRefusesExtraction('array of Integer', False),
    'a local that would BECOME an open array');
  Assert.IsFalse(ParamTypeRefusesExtraction('array of Integer', True),
    'but a forwarded open array is one already - the reported case');
  Assert.IsFalse(ParamTypeRefusesExtraction('Integer', False),
    'a named type never refuses');
  Assert.IsFalse(ParamTypeRefusesExtraction('', False),
    'and an unknown type is not a refusal either');
end;

procedure TExtractAndCompletionTests.LinesAreSplitOnEveryLineBreakStyle;
begin
  Assert.AreEqual(3, Integer(Length(SplitEditorLines('a'#10'b'#10'c'))), 'LF only');
  Assert.AreEqual(3, Integer(Length(SplitEditorLines('a'#13#10'b'#13#10'c'))), 'CRLF');
  Assert.AreEqual(3, Integer(Length(SplitEditorLines('a'#13'b'#13'c'))), 'CR only');
  var L := SplitEditorLines('unit U;'#10'interface'#10);
  Assert.AreEqual('interface', L[1], 'and the lines themselves are right');
end;


procedure TUnitRenameScopeTests.NoRealUsesClauseEntryIsMissed;
var
  Dir, Content: string;
  Files: TArray<string>;
  Checked, Missed: Integer;
  MissedWhere: string;
begin
  // The repository's own Source folder, found from the test exe (…\Tests\Win32\Debug).
  // Walk up from the test exe until the repository's Source folder is
  // there. NOT an Assert.Pass when it is not: a sweep that silently skips
  // proves nothing, and that is exactly what the first version did (the
  // negative check failed only the fixture, not this test).
  Dir := ExtractFilePath(ParamStr(0));
  for var Up := 1 to 6 do
  begin
    Dir := ExtractFilePath(ExcludeTrailingPathDelimiter(Dir));
    if TDirectory.Exists(TPath.Combine(Dir, 'Source')) then Break;
  end;
  Dir := TPath.Combine(Dir, 'Source');
  Assert.IsTrue(TDirectory.Exists(Dir),
    'the repository Source folder must be reachable from ' + ParamStr(0));
  Files := TDirectory.GetFiles(Dir, '*.pas');
  Assert.IsTrue(Length(Files) > 50, 'the sweep must see the real units');
  Checked := 0;
  Missed := 0;
  MissedWhere := '';
  for var F in Files do
  begin
    try
      Content := ReadDelphiFile(F);
    except
      Continue;
    end;
    var Lines := SplitEditorLines(Content);
    var Masked := MaskCommentsAndStrings(Lines);
    var Regions := CollectUnitRefRegions(Masked);
    for var E in TUsesGraphAnalyzer.ParseUsesEntries(Content) do
    begin
      if (E.Line < 1) or (E.Line > Length(Lines)) then Continue;
      // the LAST segment of a dotted name is what a rename would look for
      var Name := E.UnitName;
      var Dot := LastDelimiter('.', Name);
      if Dot > 0 then Name := Copy(Name, Dot + 1, MaxInt);
      var C := Pos(Name, Lines[E.Line - 1]);
      if C = 0 then Continue;          // wrapped differently - not this test's point
      Inc(Checked);
      if not UnitNameOccurrenceIsUnitReference(Masked, Regions, E.Line - 1, C,
        Length(Name)) then
      begin
        Inc(Missed);
        if MissedWhere = '' then
          MissedWhere := Format('%s:%d %s', [ExtractFileName(F), E.Line, E.UnitName]);
      end;
    end;
  end;
  Assert.IsTrue(Checked > 300,
    Format('the sweep must really look at the clauses (checked %d)', [Checked]));
  Assert.AreEqual(0, Missed, Format('%d uses entry/entries not recognised, ' +
    'first: %s', [Missed, MissedWhere]));
end;


procedure TUnitRenameScopeTests.NoRealOverrideIsTakenForAnotherSymbol;
var
  Dir, Content, Why: string;
  Files: TArray<string>;
  Checked, Wrong: Integer;
  WrongWhere: string;
begin
  Dir := ExtractFilePath(ParamStr(0));
  for var Up := 1 to 6 do
  begin
    Dir := ExtractFilePath(ExcludeTrailingPathDelimiter(Dir));
    if TDirectory.Exists(TPath.Combine(Dir, 'Source')) then Break;
  end;
  Dir := TPath.Combine(Dir, 'Source');
  Assert.IsTrue(TDirectory.Exists(Dir),
    'the repository Source folder must be reachable from ' + ParamStr(0));
  Files := TDirectory.GetFiles(Dir, '*.pas');
  Checked := 0;
  Wrong := 0;
  WrongWhere := '';
  for var F in Files do
  begin
    try
      Content := ReadDelphiFile(F);
    except
      Continue;
    end;
    var Lines := SplitEditorLines(Content);
    for var L := 0 to High(Lines) do
    begin
      var T := Trim(Lines[L]);
      // "procedure Foo(...); override;" inside a class body, with its
      // owning type read from the lines above.
      if not HasWholeWordCI(T, 'override') then Continue;
      var Kind: string;
      var IsClassMethod: Boolean;
      if not IsHeaderLine(T, Kind, IsClassMethod) then Continue;
      var Owner := EnclosingContainerName(Lines, L);
      if Owner = '' then Continue;
      // the member's name: the token after the keyword
      var Rest := Trim(Copy(T, Pos(' ', T) + 1, MaxInt));
      if IsClassMethod then Rest := Trim(Copy(Rest, Pos(' ', Rest) + 1, MaxInt));
      var NameEnd := 1;
      while (NameEnd <= Length(Rest)) and IsIdentChar(Rest[NameEnd]) do Inc(NameEnd);
      var Member := Copy(Rest, 1, NameEnd - 1);
      if Member = '' then Continue;
      Inc(Checked);
      // An override IS the symbol - whatever the parameter count, which is
      // why the rule only ever acts on reintroduce or an overload.
      if MemberIsOtherSymbol(Content, Owner, Member, -1, Why)
        or MemberIsOtherSymbol(Content, Owner, Member, 1, Why) then
      begin
        Inc(Wrong);
        if WrongWhere = '' then
          WrongWhere := Format('%s:%d %s.%s - %s',
            [ExtractFileName(F), L + 1, Owner, Member, Why]);
      end;
    end;
  end;
  // 63 override declarations exist in Source today (counted with grep);
  // the bound only has to prove the sweep really looked.
  Assert.IsTrue(Checked >= 50,
    Format('the sweep must really see overrides (checked %d)', [Checked]));
  Assert.AreEqual(0, Wrong, Format('%d override(s) taken for another ' +
    'symbol, first: %s', [Wrong, WrongWhere]));
end;

{ TDfmAliasOnlyTests }

procedure TDfmAliasOnlyTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'RLDfmAlias_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
end;

procedure TDfmAliasOnlyTests.TearDown;
begin
  if (FDir <> '') and TDirectory.Exists(FDir) then
    try
      TDirectory.Delete(FDir, True);
    except
    end;
end;

function TDfmAliasOnlyTests.WriteFile(const AName, AText: string): string;
begin
  Result := TPath.Combine(FDir, AName);
  TFile.WriteAllText(Result, AText, TEncoding.UTF8);
end;

procedure TDfmAliasOnlyTests.AnotherNameForTheSameTypeIsANoteNotAMismatch;
var
  Lib, Tree, Pas: string;
  Issues: TArray<TDfmEventIssue>;
begin
  // the shape of cxGraphics / cxTL: a unit-qualified alias, and a plain one
  Lib := WriteFile('ULibGraphics.pas',
    'unit ULibGraphics;' + NL + 'interface' + NL + 'uses System.UITypes;' + NL +
    'type' + NL +
    '  TLibImageIndex = System.UITypes.TImageIndex;' + NL +
    '  TLibCount = Integer;' + NL +
    'implementation' + NL + 'end.' + NL);
  Tree := WriteFile('ULibTree.pas',
    'unit ULibTree;' + NL + 'interface' + NL +
    'uses System.Classes, ULibGraphics;' + NL +
    'type' + NL +
    '  TLibGetImageIndexEvent = procedure(Sender: TObject;' + NL +
    '    var AIndex: TLibImageIndex) of object;' + NL +
    '  TLibCountEvent = procedure(Sender: TObject; ACount: TLibCount) of object;' + NL +
    '  TLibTree = class(TComponent)' + NL +
    '  private' + NL +
    '    FOnGetImageIndex: TLibGetImageIndexEvent;' + NL +
    '    FOnCount: TLibCountEvent;' + NL +
    '  published' + NL +
    '    property OnGetImageIndex: TLibGetImageIndexEvent read FOnGetImageIndex write FOnGetImageIndex;' + NL +
    '    property OnCount: TLibCountEvent read FOnCount write FOnCount;' + NL +
    '  end;' + NL +
    'implementation' + NL + 'end.' + NL);
  Pas := WriteFile('UAliasForm.pas',
    'unit UAliasForm;' + NL + 'interface' + NL +
    'uses System.Classes, System.UITypes, Vcl.Forms, ULibTree;' + NL +
    'type' + NL +
    '  TForm1 = class(TForm)' + NL +
    '    Tree1: TLibTree;' + NL +
    '    Tree2: TLibTree;' + NL +
    '    procedure Tree1GetImageIndex(Sender: TObject; var AIndex: TImageIndex);' + NL +
    '    procedure Tree1Count(Sender: TObject; ACount: Integer);' + NL +
    '    procedure Tree2GetImageIndex(Sender: TObject; var AIndex: Word);' + NL +
    '  end;' + NL +
    'implementation' + NL + 'end.' + NL);
  WriteFile('UAliasForm.dfm',
    'object Form1: TForm1' + NL +
    '  object Tree1: TLibTree' + NL +
    '    OnGetImageIndex = Tree1GetImageIndex' + NL +
    '    OnCount = Tree1Count' + NL +
    '  end' + NL +
    '  object Tree2: TLibTree' + NL +
    '    OnGetImageIndex = Tree2GetImageIndex' + NL +
    '  end' + NL +
    'end' + NL);

  Issues := TDfmEventChecker.CheckProject([Pas], nil, [Lib, Tree]);
  Assert.AreEqual(3, Integer(Length(Issues)),
    'two handlers that only name their types differently, one that is wrong');
  for var Iss in Issues do
  begin
    Assert.AreEqual(Ord(eikSignatureMismatch), Ord(Iss.Kind), Iss.HandlerName);
    if SameText(Iss.HandlerName, 'Tree1GetImageIndex') then
    begin
      Assert.IsTrue(Iss.AliasOnly,
        'TLibImageIndex = System.UITypes.TImageIndex is TImageIndex');
      Assert.AreEqual('TLibImageIndex = TImageIndex', Iss.AliasNote);
      Assert.IsTrue(Iss.ExpectedRawParams <> '',
        'the names can still be aligned when the user ticks the row');
    end
    else if SameText(Iss.HandlerName, 'Tree1Count') then
    begin
      Assert.IsTrue(Iss.AliasOnly, 'TLibCount = Integer is Integer');
      Assert.AreEqual('TLibCount = Integer', Iss.AliasNote);
    end
    else
    begin
      Assert.AreEqual('Tree2GetImageIndex', Iss.HandlerName);
      Assert.IsFalse(Iss.AliasOnly,
        'Word is no other name for TImageIndex - that one IS a mismatch');
      Assert.AreEqual('', Iss.AliasNote);
    end;
  end;
end;

procedure TExtractAndCompletionTests.TheCallKeepsTheCodesOwnIndentation;
var
  Block: TArray<string>;
begin
  Block := ['    N := 0;', '    Inc(N);'];
  // The two ways a user selects the same block must give the same call.
  Assert.AreEqual('    ', CallIndentFor(Block, '    N := 0;'#10'    Inc(N);'#10),
    'dragged from the left margin');
  Assert.AreEqual('    ', CallIndentFor(Block, 'N := 0;'#10'    Inc(N);'#10),
    'dragged from the first statement - the reported case');

  // A tab stays a tab: re-indenting someone''s file is not our business.
  Assert.AreEqual(#9#9, CallIndentFor([#9#9'DoIt;'], 'DoIt;'#10),
    'tabs are kept verbatim');

  // A blank line carries no indentation, so the first REAL line decides.
  Assert.AreEqual('  ', CallIndentFor(['', '  DoIt;'], 'DoIt;'#10),
    'a blank first line is skipped');

  // Without file lines (the buffer could not be read) the old rule stands.
  Assert.AreEqual('  ', CallIndentFor(nil, '  N := 0;'#10'    Inc(N);'#10),
    'fallback: the smallest indentation of the selected text');
  Assert.AreEqual('', CallIndentFor(nil, ''),
    'nothing to go by at all');
end;

{ TMethodEditTests }

function TMethodEditTests.SourceUnit: TArray<string>;
begin
  Result := [
    'unit Old;',                          // 0
    '',
    'interface',
    '',
    'type',
    '  TOld = class',                     // 5
    '  private',
    '    FColumns: Integer;',
    '    procedure Recalc;',
    '  public',
    '    constructor Create;',            // 10
    '    function FormatRow(const ARow: string): string;',
    '    procedure Paint; virtual;',
    '  end;',
    '',
    'implementation',                     // 15
    '',
    'constructor TOld.Create;',
    'begin',
    '  FColumns := 1;',
    'end;',                               // 20
    '',
    'procedure TOld.Recalc;',
    'begin',
    'end;',
    '',                                   // 25
    'function TOld.FormatRow(const ARow: string): string;',
    'begin',
    '  Recalc;',
    '  Result := ARow + IntToStr(FColumns);',
    'end;',                               // 30
    '',
    'procedure TOld.Paint;',
    'begin',
    'end;',
    '',
    'end.'];
end;

function TMethodEditTests.TargetUnit: TArray<string>;
begin
  Result := [
    'unit New;',
    '',
    'interface',
    '',
    'type',
    '  TNew = class',
    '  private',
    '    FKept: Integer;',
    '  public',
    '    procedure Go;',
    '  end;',
    '',
    'implementation',
    '',
    'procedure TNew.Go;',
    'begin',
    'end;',
    '',
    'end.'];
end;

procedure TMethodEditTests.TheListSaysWhatCanTravelAlone;
var
  Members: TArray<TClassMemberInfo>;
  Names: string;
  Paint: TClassMemberInfo;
begin
  Members := ClassMembersOf(SourceUnit, 'TOld');
  Names := '';
  Paint := Default(TClassMemberInfo);
  for var M in Members do
  begin
    if Names <> '' then Names := Names + ', ';
    Names := Names + M.Name;
    if SameText(M.Name, 'Paint') then Paint := M;
  end;
  // Declaration order, methods only - fields and properties are not rows.
  Assert.AreEqual('Recalc, Create, FormatRow, Paint', Names,
    'the methods of TOld, in the order they are declared');
  Assert.AreEqual('Paint', Paint.Name, 'the virtual one is in the list');
  Assert.IsFalse(Paint.Movable, 'but it cannot travel alone');
  Assert.IsTrue(ContainsText(Paint.Why, 'virtual'),
    'and the row says why: ' + Paint.Why);
  for var M in Members do
    if SameText(M.Name, 'FormatRow') then
      Assert.IsTrue(M.Movable, 'a plain method can be moved');
end;

procedure TMethodEditTests.AMethodMovesWithItsBodyAndIsRequalified;
var
  Plan: TMethodMovePlan;
  Src, Tgt: string;
  Own: string;
begin
  Plan := PlanMethodMove(SourceUnit, TargetUnit, 'TOld', ['FormatRow'],
    'TNew', 'private');
  Assert.IsTrue(Plan.Ok, 'the move is planned: ' + Plan.Error);
  Src := string.Join(NL, Plan.SourceLines);
  Tgt := string.Join(NL, Plan.TargetLines);

  // Gone from the old class, declaration AND body.
  Assert.IsFalse(ContainsText(Src, 'function FormatRow'),
    'the declaration left TOld');
  Assert.IsFalse(ContainsText(Src, 'TOld.FormatRow'),
    'and so did the body');
  Assert.IsTrue(ContainsText(Src, 'procedure TOld.Recalc;'),
    'everything else stays untouched');

  // Arrived in the new one, under the section that was asked for.
  Assert.IsTrue(ContainsText(Tgt, 'function FormatRow(const ARow: string): string;'),
    'the declaration arrived');
  Assert.IsTrue(Pos('FormatRow', Tgt) > Pos('private', Tgt),
    'in the private section');
  Assert.IsTrue(Pos('FormatRow', Tgt) < Pos('public', Tgt),
    'and not below the public one');
  Assert.IsTrue(ContainsText(Tgt, 'function TNew.FormatRow(const ARow: string): string;'),
    'the implementation header is requalified');
  Assert.IsTrue(ContainsText(Tgt, 'Result := ARow + IntToStr(FColumns);'),
    'the body came along verbatim');
  Assert.IsTrue(Pos('TNew.FormatRow', Tgt) > Pos('implementation', Tgt),
    'and sits in the implementation section');
  // One blank line separates it, not two - the first live apply left a
  // double blank behind every moved body.
  Assert.IsFalse(ContainsText(Tgt, NL + NL + NL),
    'no double blank line where the body landed');

  // And it SAYS what it did not do.
  Own := '';
  for var Iss in Plan.Issues do
    if Iss.Kind = meiOwnMember then Own := Iss.Text;
  Assert.IsTrue(ContainsText(Own, 'FColumns'), 'the field stays behind: ' + Own);
  Assert.IsTrue(ContainsText(Own, 'Recalc'), 'and so does the method it calls');
end;

procedure TMethodEditTests.AMemberThatCannotTravelAloneIsRefused;
var
  Plan: TMethodMovePlan;
  Veto: string;
begin
  Plan := PlanMethodMove(SourceUnit, TargetUnit, 'TOld', ['Paint'],
    'TNew', 'public');
  Assert.IsFalse(Plan.Ok, 'a virtual method is not moved');
  Assert.AreEqual(0, Integer(Length(Plan.Moved)), 'nothing moved');
  Veto := '';
  for var Iss in Plan.Issues do
    if Iss.Kind = meiVeto then Veto := Iss.Text;
  Assert.IsTrue(ContainsText(Veto, 'virtual'), 'and the reason is named: ' + Veto);
  // An unknown target class is an error, not a silent no-op.
  Plan := PlanMethodMove(SourceUnit, TargetUnit, 'TOld', ['FormatRow'],
    'TNowhere', 'private');
  Assert.IsTrue(ContainsText(Plan.Error, 'TNowhere'), Plan.Error);
end;

procedure TMethodEditTests.ModifiersAreEditedOnTheDeclaration;
begin
  // A ';' inside the parameter list is not the end of the signature.
  Assert.AreEqual('    procedure Go(A: Integer; B: string); virtual;',
    ApplyModifiersToDecl('    procedure Go(A: Integer; B: string);',
      ['virtual'], []),
    'added, indentation kept');
  Assert.AreEqual('    procedure Go; virtual;',
    ApplyModifiersToDecl('    procedure Go; virtual;', ['virtual'], []),
    'never twice');
  Assert.AreEqual('    procedure Go; virtual;',
    ApplyModifiersToDecl('    procedure Go; overload; virtual;', [], ['overload']),
    'removed, the others keep their order');
  Assert.AreEqual('    class procedure Go; static; inline;',
    ApplyModifiersToDecl('    class procedure Go; static;', ['inline'], []),
    'a class method keeps its prefix');
  Assert.AreEqual('  FCount: Integer',
    ApplyModifiersToDecl('  FCount: Integer', ['virtual'], []),
    'a line without a ";" is left alone');
end;

procedure TMethodEditTests.TheUsesClausesFollowAndTheCircleIsNamed;
var
  Src, Tgt: TArray<string>;
  Plan: TMethodMovePlan;
  T, Note: string;
begin
  // The declaration names a type of the SOURCE unit, and the source keeps
  // calling what moves - both uses clauses have to react.
  Src := [
    'unit Old;',
    '',
    'interface',
    '',
    'uses System.SysUtils;',
    '',
    'type',
    '  TRowData = record',
    '    Value: string;',
    '  end;',
    '',
    '  TOld = class',
    '  public',
    '    function FormatRow(const ARow: TRowData): string;',
    '    procedure Dump;',
    '  end;',
    '',
    'implementation',
    '',
    'function TOld.FormatRow(const ARow: TRowData): string;',
    'begin',
    '  Result := ARow.Value;',
    'end;',
    '',
    'procedure TOld.Dump;',
    'begin',
    '  Writeln(FormatRow(Default(TRowData)));',
    'end;',
    '',
    'end.'];
  Tgt := [
    'unit New;',
    '',
    'interface',
    '',
    'uses System.Classes;',
    '',
    'type',
    '  TNew = class',
    '  public',
    '    procedure Go;',
    '  end;',
    '',
    'implementation',
    '',
    'procedure TNew.Go;',
    'begin',
    'end;',
    '',
    'end.'];

  Plan := PlanMethodMove(Src, Tgt, 'TOld', ['FormatRow'], 'TNew', 'public');
  Assert.IsTrue(Plan.Ok, 'planned: ' + Plan.Error);

  // The target needs the source - and in the INTERFACE, because the moved
  // DECLARATION names TRowData.
  Assert.AreEqual('Old', Plan.TargetUsesAdded, 'the target gained the source unit');
  Assert.IsTrue(Plan.TargetUsesInInterface, 'in its interface uses');
  T := string.Join(NL, Plan.TargetLines);
  Assert.IsTrue(Pos('Old', T) < Pos('implementation', T),
    'the entry stands above the implementation keyword');

  // ... and exactly that direction can close a circle, so it is said.
  Note := '';
  for var Iss in Plan.Issues do
    if Iss.Kind = meiNote then Note := Iss.Text;
  Assert.IsTrue(ContainsText(Note, 'circle'), 'the risk is named: ' + Note);

  // The source still calls FormatRow, so it needs the target unit - in the
  // implementation, where no circle can form.
  Assert.AreEqual('New', Plan.SourceUsesAdded, 'the source gained the target unit');
  Assert.IsTrue(Pos('New', string.Join(NL, Plan.SourceLines)) >
    Pos('implementation', string.Join(NL, Plan.SourceLines)),
    'and that one went into the implementation uses');
end;

procedure TMethodEditTests.AClassInTheSameUnitIsOneRewrite;
var
  Lines: TArray<string>;
  Plan: TMethodMovePlan;
  Txt: string;

  // How often does a line with that text exist? Counting the result is the
  // only way to see a doubled insertion - a second copy still "contains" it.
  function Occurrences(const AText: string): Integer;
  begin
    Result := 0;
    for var L in Plan.TargetLines do
      if ContainsText(L, AText) then Inc(Result);
  end;

begin
  Lines := [
    'unit Two;',
    '',
    'interface',
    '',
    'type',
    '  TFirst = class',
    '  public',
    '    function Describe: string;',
    '    procedure Keep;',
    '  end;',
    '',
    '  TSecond = class',
    '  private',
    '    FTag: Integer;',
    '  end;',
    '',
    'implementation',
    '',
    'function TFirst.Describe: string;',
    'begin',
    '  Result := ''first'';',
    'end;',
    '',
    'procedure TFirst.Keep;',
    'begin',
    'end;',
    '',
    'end.'];

  Plan := PlanMethodMove(Lines, Lines, 'TFirst', ['Describe'], 'TSecond',
    'private');
  Assert.IsTrue(Plan.Ok, 'planned: ' + Plan.Error);

  // ONE result, not two halves: both arrays are the same finished text.
  Assert.AreEqual(string.Join(NL, Plan.SourceLines),
    string.Join(NL, Plan.TargetLines), 'one file, one result');
  Txt := string.Join(NL, Plan.TargetLines);

  // The removal happened ONCE, and the insertion found the shifted places.
  Assert.AreEqual(1, Occurrences('function Describe: string;'),
    'the declaration exists exactly once');
  Assert.AreEqual(1, Occurrences('Result := ''first'';'),
    'and so does the body');
  Assert.IsTrue(ContainsText(Txt, 'function TSecond.Describe: string;'),
    'requalified to the new owner');
  Assert.IsFalse(ContainsText(Txt, 'TFirst.Describe'),
    'the old header is gone');
  Assert.IsTrue(ContainsText(Txt, 'procedure TFirst.Keep;'),
    'the sibling is untouched');
  Assert.IsTrue(Pos('Describe', Txt) > Pos('TSecond', Txt),
    'the declaration moved into the second class');

  // No uses clause can be involved in a move inside one unit.
  Assert.AreEqual('', Plan.SourceUsesAdded, 'no uses entry for the source');
  Assert.AreEqual('', Plan.TargetUsesAdded, 'none for the target either');

  // And the same class is no target at all.
  Plan := PlanMethodMove(Lines, Lines, 'TFirst', ['Describe'], 'TFirst',
    'private');
  Assert.IsFalse(Plan.Ok, 'moving into its own class is refused');
  Assert.IsTrue(ContainsText(Plan.Error, 'same'), Plan.Error);
end;

procedure TMethodEditTests.ASignatureChangeReachesBothHeaders;
var
  Lines, Res: TArray<string>;
  Err, Txt: string;
  Params: TArray<TSigParam>;
begin
  Lines := SourceUnit;

  // The parameter list of FormatRow gains a second parameter.
  SetLength(Params, 2);
  Params[0].Modifier := 'const';
  Params[0].Name := 'ARow';
  Params[0].TypeText := 'string';
  Params[1].Name := 'AWidth';
  Params[1].TypeText := 'Integer';
  Params[1].DefaultText := '0';
  Res := ApplySignatureInClass(Lines, 'TOld', 'FormatRow', Params, 'string', Err);
  Assert.AreEqual('', Err, 'rewritten without a complaint');
  Txt := string.Join(NL, Res);
  Assert.IsTrue(ContainsText(Txt,
    'function FormatRow(const ARow: string; AWidth: Integer = 0): string;'),
    'the declaration carries the new list - with the default value');
  Assert.IsTrue(ContainsText(Txt,
    'function TOld.FormatRow(const ARow: string; AWidth: Integer): string;'),
    'and so does the implementation header - WITHOUT the default');
  Assert.IsTrue(ContainsText(Txt, '  Result := ARow + IntToStr(FColumns);'),
    'the body is untouched');

  // A member without an implementation says so instead of changing one half.
  Res := ApplySignatureInClass(Lines, 'TOld', 'Nothing', Params, 'string', Err);
  Assert.IsTrue(ContainsText(Err, 'not found'), Err);
  Assert.AreEqual(string.Join(NL, Lines), string.Join(NL, Res),
    'and nothing was written');

  // The modifiers reach EVERY selected member, and only those.
  Res := ApplyModifiersInClass(Lines, 'TOld', ['Recalc', 'FormatRow'],
    ['virtual'], []);
  Txt := string.Join(NL, Res);
  Assert.IsTrue(ContainsText(Txt, 'procedure Recalc; virtual;'), 'the first one');
  Assert.IsTrue(ContainsText(Txt,
    'function FormatRow(const ARow: string): string; virtual;'), 'the second one');
  Assert.IsTrue(ContainsText(Txt, 'constructor Create;' + NL),
    'the one nobody selected keeps its line');

  // The target classes a move can offer: top level only.
  var Names := string.Join(', ', ClassNamesOf(SourceUnit));
  Assert.AreEqual('TOld', Names, 'the only class of the unit');
end;

procedure TMethodEditTests.TheRequestIsCarriedOutInOneOrder;
var
  Req: TMethodEditRequest;
  Res: TMethodEditResult;
  Params: TArray<TSigParam>;
  Txt: string;
begin
  // Nothing asked for is not "write the file unchanged".
  Req := Default(TMethodEditRequest);
  Req.Members := ['FormatRow'];
  Res := PlanMethodEdit('Old.pas', string.Join(NL, SourceUnit), 'TOld', Req, '');
  Assert.IsFalse(Res.Ok, 'a request without an action is refused');
  Assert.IsTrue(ContainsText(Res.Error, 'nothing is changed'), Res.Error);

  // No member selected is its own answer.
  Req.Members := nil;
  Req.AddModifiers := ['inline'];
  Res := PlanMethodEdit('Old.pas', string.Join(NL, SourceUnit), 'TOld', Req, '');
  Assert.IsFalse(Res.Ok, 'and so is an empty selection');

  // Signature AND move AND modifier in one request: the member arrives in
  // the new class with the NEW list and the added directive.
  SetLength(Params, 1);
  Params[0].Modifier := 'const';
  Params[0].Name := 'ARow';
  Params[0].TypeText := 'TRowText';
  Req := Default(TMethodEditRequest);
  Req.Members := ['FormatRow'];
  Req.TargetFile := 'New.pas';
  Req.TargetClass := 'TNew';
  Req.Section := 'public';
  Req.AddModifiers := ['inline'];
  Req.SigMember := 'FormatRow';
  Req.SigParams := Params;
  Req.SigResultType := 'string';
  Res := PlanMethodEdit('Old.pas', string.Join(NL, SourceUnit), 'TOld', Req,
    string.Join(NL, TargetUnit));
  Assert.IsTrue(Res.Ok, 'planned: ' + Res.Error);
  Txt := string.Join(NL, Res.TargetLines);
  Assert.IsTrue(ContainsText(Txt,
    'function FormatRow(const ARow: TRowText): string; inline;'),
    'the new list and the new directive are both there: ' + Txt);
  Assert.IsTrue(ContainsText(Txt,
    'function TNew.FormatRow(const ARow: TRowText): string;'),
    'and the implementation header carries the new list too');
  Assert.IsFalse(ContainsText(string.Join(NL, Res.SourceLines), 'FormatRow'),
    'nothing of it stayed behind');

  // The signature change is the one thing that reaches further than this
  // dialog, so it has to SAY so.
  var Said := False;
  for var I in Res.Issues do
    if ContainsText(I.Text, 'Change signature') then Said := True;
  Assert.IsTrue(Said, 'the call sites are named as the user''s job');
end;

procedure TMethodEditTests.ASelectionNamesEveryMemberItCovers;
var
  L: TArray<string>;
begin
  L := SourceUnit;

  // Three DECLARATIONS marked (lines 10..12) - the reported case.
  Assert.AreEqual('Create,FormatRow,Paint',
    string.Join(',', MembersInLineRange(L, ClassMembersOf(L, 'TOld'), 'TOld', 10, 12)),
    'three marked declarations are three members, in declaration order');

  // The two ends may arrive in either order (a selection dragged upwards).
  Assert.AreEqual('Create,FormatRow,Paint',
    string.Join(',', MembersInLineRange(L, ClassMembersOf(L, 'TOld'), 'TOld', 12, 10)),
    'the ends may arrive the other way round');

  // BODIES marked instead (Recalc 22..24, FormatRow 26..30): the same
  // question, so the same answer - and FormatRow is named ONCE although the
  // range covers five of its lines.
  Assert.AreEqual('Recalc,FormatRow',
    string.Join(',', MembersInLineRange(L, ClassMembersOf(L, 'TOld'), 'TOld', 22, 30)),
    'marked implementations name their members, each one once');

  // One line inside one body is that member, nothing else.
  Assert.AreEqual('FormatRow',
    string.Join(',', MembersInLineRange(L, ClassMembersOf(L, 'TOld'), 'TOld', 28, 28)),
    'a caret-sized range inside a body');

  // A range that covers no member at all answers nothing - the wizard then
  // falls back to the caret instead of opening an empty dialog.
  Assert.AreEqual('',
    string.Join(',', MembersInLineRange(L, ClassMembersOf(L, 'TOld'), 'TOld', 2, 4)),
    'interface / type lines name no member');

  // A SIBLING CLASS's body in the same unit must not count, however much
  // its member is called like ours - that is what the owner-type test of the
  // implementation header is for.
  var S: TArray<string> := [
    'unit Two;',                        // 0
    'interface',
    'type',
    '  TOne = class',
    '    procedure Run;',               // 4
    '  end;',
    '  TTwo = class',
    '    procedure Run;',               // 7
    '  end;',
    'implementation',
    'procedure TOne.Run;',              // 10
    'begin',
    'end;',
    'procedure TTwo.Run;',              // 13
    'begin',
    'end;',
    'end.'];
  Assert.AreEqual('Run',
    string.Join(',', MembersInLineRange(S, ClassMembersOf(S, 'TOne'), 'TOne', 10, 12)),
    'TOne.Run''s own body');
  Assert.AreEqual('',
    string.Join(',', MembersInLineRange(S, ClassMembersOf(S, 'TOne'), 'TOne', 13, 15)),
    'the same-named member of the SIBLING class is not ours');
end;

procedure TMethodEditTests.MovedBodiesAreSeparatedByExactlyOneBlankLine;
var
  Plan: TMethodMovePlan;

  // The blank lines between ALINES[AFrom] and the next non-blank line.
  function BlanksAfter(const ALines: TArray<string>; AFrom: Integer): Integer;
  begin
    Result := 0;
    var I := AFrom + 1;
    while (I <= High(ALines)) and (Trim(ALines[I]) = '') do
    begin
      Inc(Result);
      Inc(I);
    end;
  end;

  function LineOf(const ALines: TArray<string>; const AText: string): Integer;
  begin
    for var I := 0 to High(ALines) do
      if ContainsText(ALines[I], AText) then Exit(I);
    Result := -1;
  end;

begin
  // TWO members at once - the reported case needs more than one body.
  Plan := PlanMethodMove(SourceUnit, TargetUnit, 'TOld', ['Recalc', 'FormatRow'],
    'TNew', 'private');
  Assert.IsTrue(Plan.Ok, 'the move is planned: ' + Plan.Error);

  var Tgt := Plan.TargetLines;
  var First := LineOf(Tgt, 'TNew.Recalc');
  var Second := LineOf(Tgt, 'TNew.FormatRow');
  Assert.IsTrue(First >= 0, 'the first body arrived');
  Assert.IsTrue(Second >= 0, 'the second body arrived');

  // Between the END of the first body and the second header: exactly one.
  var EndOfFirst := -1;
  for var I := First to Second do
    if SameText(Trim(Tgt[I]), 'end;') then EndOfFirst := I;
  Assert.IsTrue(EndOfFirst >= 0, 'the first body ends before the second');
  Assert.AreEqual(1, BlanksAfter(Tgt, EndOfFirst),
    'exactly ONE blank line between two moved bodies');

  // And one under the last body, so the final "end." does not sit on it.
  var Final := LineOf(Tgt, 'end.');
  Assert.IsTrue(Final > Second, 'the unit still ends after the bodies');
  var EndOfLast := -1;
  for var I := Second to Final do
    if SameText(Trim(Tgt[I]), 'end;') then EndOfLast := I;
  Assert.AreEqual(1, BlanksAfter(Tgt, EndOfLast),
    'one blank line between the last moved body and what follows');

  // No double blank anywhere in either file - the shape the user saw.
  Assert.IsFalse(ContainsText(string.Join(NL, Tgt), NL + NL + NL),
    'no double blank line in the target');
  Assert.IsFalse(ContainsText(string.Join(NL, Plan.SourceLines), NL + NL + NL),
    'and none left behind in the source');
end;

procedure TMethodEditTests.AnUnchangedPlanIsByteIdentical;
var
  SL: TStringList;

  procedure RoundTrip(const AContent, AWhat: string);
  begin
    var L := SplitContentLines(AContent);
    Assert.AreEqual(AContent, JoinPlannedLines(L, AContent),
      'split and join are inverse: ' + AWhat);
    // ...and the way the apply hands those lines to ApplyLinesMinimal must
    // not add a line of its own.
    SL.Text := JoinPlannedLines(L, AContent);
    var ViaAdd := TStringList.Create;
    try
      for var S in L do ViaAdd.Add(S);
      Assert.AreEqual(SL.Count + 1, ViaAdd.Count,
        'adding the elements one by one is what produced the extra line: ' + AWhat);
    finally
      ViaAdd.Free;
    end;
  end;

begin
  SL := TStringList.Create;
  try
    RoundTrip('unit U;' + #13#10 + 'implementation' + #13#10 + 'end.' + #13#10,
      'CRLF with a final break');
    RoundTrip('unit U;'#10'implementation'#10'end.'#10, 'LF only');
    // No final break: there is no trailing element, so nothing can grow -
    // the join must still return the text unchanged.
    var NoBreak := 'unit U;' + #13#10 + 'end.';
    Assert.AreEqual(NoBreak,
      JoinPlannedLines(SplitContentLines(NoBreak), NoBreak),
      'a file without a final break keeps it that way');

    // The real plan, in the shape the apply has it: the lines come from
    // SplitContentLines, so they carry the file's final break as their
    // trailing empty element - and moving a member must not change that.
    var SrcText := string.Join(#13#10, SourceUnit) + #13#10;
    var TgtText := string.Join(#13#10, TargetUnit) + #13#10;
    var Plan := PlanMethodMove(SplitContentLines(SrcText),
      SplitContentLines(TgtText), 'TOld', ['FormatRow'], 'TNew', 'private');
    Assert.IsTrue(Plan.Ok, 'the move is planned: ' + Plan.Error);
    var Written := JoinPlannedLines(Plan.SourceLines, SrcText);
    Assert.IsTrue(Written.EndsWith('end.' + #13#10),
      'the source still ends with exactly one break after "end."');
    Assert.IsFalse(Written.EndsWith('end.' + #13#10 + #13#10),
      'and not with a blank line the move added');
    var TgtWritten := JoinPlannedLines(Plan.TargetLines, TgtText);
    Assert.IsTrue(TgtWritten.EndsWith('end.' + #13#10),
      'and so does the target');
    Assert.IsFalse(TgtWritten.EndsWith('end.' + #13#10 + #13#10),
      'the target grew no blank line either - both files did');
  finally
    SL.Free;
  end;
end;

procedure TMethodEditTests.WhatStillNeedsHandWorkAfterTheEdit;
var
  Plan: TMethodMovePlan;
  Refs: TArray<TPostEditRef>;
  Total: Integer;
  Capped: Boolean;

  function KindOf(const AFile: string; ALine: Integer): string;
  begin
    Result := '';
    for var R in Refs do
      if SameText(R.FilePath, AFile) and (R.Line = ALine) then Exit(R.Kind);
  end;

begin
  Plan := PlanMethodMove(SourceUnit, TargetUnit, 'TOld', ['FormatRow'],
    'TNew', 'private');
  Assert.IsTrue(Plan.Ok, 'the move is planned: ' + Plan.Error);
  var TgtText := string.Join(NL, Plan.TargetLines);

  // A third unit that CALLS the moved member - the rows that matter.
  var CallerText := string.Join(NL, [
    'unit Caller;',                               // 0
    'interface',
    'implementation',
    'uses New;',
    'procedure Go(O: TNew);',                     // 4
    'begin',
    '  WriteLn(O.FormatRow(''a''));',             // 6
    '  WriteLn(O.FormatRow(''b''));',             // 7
    'end;',
    'end.']);

  Refs := CollectPostEditRefs(['C:\p\New.pas', 'C:\p\Caller.pas'],
    [TgtText, CallerText], ['FormatRow'],
    'C:\p\New.pas', 'TNew', TgtText, 0, 0, Total, Capped);

  // The new declaration and the new body are the edit's own result.
  for var R in Refs do
    Assert.IsFalse(SameText(R.FilePath, 'C:\p\New.pas'),
      'the target''s own declaration and body are not rows, but ' +
      R.Kind + ' was listed');

  Assert.AreEqual(2, Total, 'the two calls in the third unit');
  Assert.AreEqual(2, Integer(Length(Refs)), 'and both are listed');
  Assert.IsFalse(Capped, 'nothing was capped');
  Assert.AreEqual('Call', KindOf('C:\p\Caller.pas', 6), 'a call is a call');
  Assert.AreEqual('FormatRow', Refs[0].Member, 'the row names its member');
  Assert.AreEqual('WriteLn(O.FormatRow(''a''));', Refs[0].Text,
    'and carries the line to look at');

  // Sorted by file and line - the order someone works through them in.
  for var I := 1 to High(Refs) do
    Assert.IsTrue((CompareText(Refs[I - 1].FilePath, Refs[I].FilePath) < 0) or
      ((SameText(Refs[I - 1].FilePath, Refs[I].FilePath)) and
       (Refs[I - 1].Line <= Refs[I].Line)), 'rows are sorted');

  // THE BUDGET, and the honest total beside it: a member called Create
  // occurs 1685 times in this repository, which is no window content - but
  // "2 found, 1 listed" must not read as "1 found".
  Refs := CollectPostEditRefs(['C:\p\New.pas', 'C:\p\Caller.pas'],
    [TgtText, CallerText], ['FormatRow'],
    'C:\p\New.pas', 'TNew', TgtText, 1, 0, Total, Capped);
  Assert.AreEqual(1, Integer(Length(Refs)), 'one row per member at most');
  Assert.AreEqual(2, Total, 'while the total still says what there is');
  Assert.IsTrue(Capped, 'and it says that it capped');
end;

procedure TMethodEditTests.TheTargetGetsTheUnitsTheMovedCodeNeeds;
var
  Src, Tgt: TArray<string>;
  Plan: TMethodMovePlan;

  // How many LINES of the target name that text - a second uses entry
  // would show up here.
  function Lines_(const AText: string): Integer;
  begin
    Result := 0;
    for var L in Plan.TargetLines do
      if ContainsText(L, AText) then Inc(Result);
  end;

  // The identifier index, as this test knows it.
  function Lookup: TMethodEditUnitLookup;
  begin
    Result :=
      function(const AIdent: string): TArray<string>
      begin
        Result := nil;
        if SameText(AIdent, 'IOTAKeyContext') then Result := ['ToolsAPI'];
        if SameText(AIdent, 'TKeyBindingResult') then Result := ['ToolsAPI'];
        if SameText(AIdent, 'TStringList') then Result := ['System.Classes'];
        // The real risk, and the reason the SOURCE's uses clause decides:
        // the index knows a Count somewhere, and the body's own field is
        // called Count. A unit nobody here uses must never be added.
        if SameText(AIdent, 'Count') then Result := ['Winapi.Windows'];
        // Named in a COMMENT only - the masking must not even ask.
        if SameText(AIdent, 'TcxGrid') then Result := ['cxGrid'];
      end;
  end;

begin
  Src := [
    'unit Old;',
    'interface',
    'uses ToolsAPI, System.Classes;',        // both reachable HERE
    'type',
    '  TOld = class',
    '  public',
    '    procedure KeyProc(const Context: IOTAKeyContext;',
    '      var BindingResult: TKeyBindingResult);',
    '  end;',
    'implementation',
    'procedure TOld.KeyProc(const Context: IOTAKeyContext;',
    '  var BindingResult: TKeyBindingResult);',
    'begin',
    '  FList := TStringList.Create;',
    '  Count := 0;',                         // a field of TOld, not a type
    '  // TcxGrid is only named in a COMMENT',
    'end;',
    'end.'];
  Tgt := [
    'unit New;',
    'interface',
    'uses System.Classes;',                  // has Classes, NOT ToolsAPI
    'type',
    '  TNew = class',
    '  private',
    '  end;',
    'implementation',
    'end.'];

  // Lookup() with the parentheses: without them Delphi passes the nested
  // FUNCTION itself, not what it returns.
  Plan := PlanMethodMove(Src, Tgt, 'TOld', ['KeyProc'], 'TNew', 'private',
    Lookup());
  Assert.IsTrue(Plan.Ok, 'the move is planned: ' + Plan.Error);
  var TgtText := string.Join(NL, Plan.TargetLines);

  // The unit the DECLARATION needs goes into the interface uses.
  Assert.IsTrue(ContainsText(TgtText, 'ToolsAPI'),
    'the target gets ToolsAPI - the moved header names its types');
  var UsesLine := '';
  for var L in Plan.TargetLines do
    if StartsText('uses', TrimLeft(L)) and (UsesLine = '') then UsesLine := L;
  Assert.IsTrue(ContainsText(UsesLine, 'ToolsAPI'),
    'and it goes into the INTERFACE clause: ' + UsesLine);

  // What the target already has is not added again, and a unit NOBODY uses
  // is never invented from a name in a comment.
  Assert.AreEqual(1, Lines_('System.Classes'),
    'System.Classes was already there - not added a second time');
  var AllUses := '';
  for var L in Plan.TargetLines do
    if StartsText('uses', TrimLeft(L)) then AllUses := AllUses + L + NL;
  Assert.IsFalse(ContainsText(AllUses, 'Winapi.Windows'),
    'a unit the SOURCE does not use is never added: ' + AllUses);
  Assert.IsFalse(ContainsText(AllUses, 'cxGrid'),
    'a name in a comment is not looked up: ' + AllUses);

  // And the plan SAYS it, so the preview shows what happened.
  var Said := False;
  for var I in Plan.Issues do
    if ContainsText(I.Text, 'ToolsAPI') then Said := True;
  Assert.IsTrue(Said, 'the plan names the unit it added');

  // Without the index: nothing is invented, and nothing breaks.
  Plan := PlanMethodMove(Src, Tgt, 'TOld', ['KeyProc'], 'TNew', 'private', nil);
  Assert.IsTrue(Plan.Ok, 'the move still works without a lookup');
  var NoUses := '';
  for var L in Plan.TargetLines do
    if StartsText('uses', TrimLeft(L)) then NoUses := NoUses + L + NL;
  Assert.IsFalse(ContainsText(NoUses, 'ToolsAPI'),
    'and says nothing it cannot know: ' + NoUses);
end;

procedure TMethodEditTests.AFindingFollowsItsLineThroughAnEdit;
begin
  var Lines: TArray<string> := [
    'procedure Go;',
    'begin',
    '  DoFirst;',
    '  DoSecond;',
    '  DoThird;',
    'end;'];

  // Nothing changed: the row stays exactly where it was.
  var R := RelocateLine(Lines, 'DoSecond;', 3);
  Assert.AreEqual(3, R.Line, 'unchanged file, unchanged line');
  Assert.IsFalse(R.Moved or R.Stale, 'and nothing to report');

  // The reported case: editing the FIRST finding deletes a line above.
  var Shorter: TArray<string> := [
    'procedure Go;',
    'begin',
    '  DoSecond;',
    '  DoThird;',
    'end;'];
  R := RelocateLine(Shorter, 'DoSecond;', 3);
  Assert.AreEqual(2, R.Line, 'the line moved up by one');
  Assert.IsTrue(R.Moved, 'and the window says so');
  Assert.IsFalse(R.Stale, 'it was found, so it is not stale');

  // Inserting far above moves it by more than any fixed window.
  var Grown: TArray<string> := [];
  for var I := 1 to 400 do Grown := Grown + ['  Filler;'];
  Grown := Grown + Lines;
  R := RelocateLine(Grown, 'DoSecond;', 3);
  Assert.AreEqual(403, R.Line, 'found 400 lines further down');
  Assert.IsTrue(R.Moved, 'moved');

  // The row's own line was rewritten: nothing can be said, so the row keeps
  // its old position and is MARKED - never moved to a guess.
  var Edited: TArray<string> := [
    'procedure Go;',
    'begin',
    '  DoFirst;',
    '  DoSecondRenamed;',
    '  DoThird;',
    'end;'];
  R := RelocateLine(Edited, 'DoSecond;', 3);
  Assert.IsTrue(R.Stale, 'the text is gone - that is reported');
  Assert.AreEqual(3, R.Line, 'and the old line is kept, not a guess');

  // Indentation is not part of the anchor: the rows carry a TRIMMED
  // preview, and re-indenting a block must not lose every row.
  var Indented: TArray<string> := [
    'procedure Go;',
    'begin',
    '    DoFirst;',
    '    DoSecond;',
    '    DoThird;',
    'end;'];
  R := RelocateLine(Indented, 'DoSecond;', 3);
  Assert.IsFalse(R.Stale, 'a re-indented line is the same line');
  Assert.AreEqual(3, R.Line, 'and it is still line 3');

  // Two identical lines: the NEAREST one wins - an edit moves code, it
  // rarely duplicates it.
  var Twice: TArray<string> := [
    '  DoSecond;',
    'procedure Go;',
    'begin',
    '  DoFirst;',
    '  DoSecond;',
    'end;'];
  R := RelocateLine(Twice, 'DoSecond;', 3);
  Assert.AreEqual(4, R.Line, 'the nearest occurrence, not the first');

  // Nothing remembered -> the old behaviour, and no empty anchor may ever
  // match a blank line.
  R := RelocateLine(Lines, '', 3);
  Assert.AreEqual(3, R.Line, 'no anchor, no change');
  Assert.IsFalse(R.Moved or R.Stale, 'and nothing claimed');
  R := RelocateLine(Lines, '   ', 99);
  Assert.AreEqual(99, R.Line, 'a blank anchor is no anchor');
  Assert.IsFalse(R.Stale, 'and it is not reported as lost');
end;

procedure TMethodEditTests.TheToolDeclaresWhatItTakes;
var
  Known: TArray<string>;
  Def: string;
begin
  Known := KnownToolArguments('edit_method');
  Assert.IsTrue(Length(Known) > 0, 'edit_method is in the shipped tool list');
  for var N in ['file', 'line', 'members', 'target_class', 'target_unit',
    'section', 'add_modifiers', 'remove_modifiers', 'apply', 'instance'] do
    Assert.IsTrue(MatchText(N, Known), N + ' is an argument of edit_method');
  // No token: every call re-runs its own analysis, so there is no window
  // between a preview and an apply that a token would have to cover.
  Assert.IsFalse(MatchText('token', Known),
    'edit_method pins no buffer revision, so it must not pretend to');

  // The definition's text is what the model reads - it has to name the
  // limit, not only the feature.
  var Arr := McpToolDefinitions;
  try
    for var E in Arr do
      if (E is TJSONObject) and
         (TJSONObject(E).GetValue<string>('name', '') = 'edit_method') then
        Def := TJSONObject(E).GetValue<string>('description', '');
  finally
    Arr.Free;
  end;
  Assert.IsTrue(Def <> '', 'the tool has a description');
  Assert.Contains(Def, 'change_signature',
    'it points at the verified path for a parameter list');
  Assert.Contains(Def, 'not rewritten',
    'and says that the call sites are not touched');
end;

procedure TMethodEditTests.APreviewReportsOnlyRealChanges;
var
  Old, New: string;
  Total: Integer;
  Changes: TArray<TPreviewChange>;
begin
  // A member MOVES down: everything between the two places shifts by one, so
  // a line-by-line pairing sees a "change" on every line of it.
  Old := 'unit U;' + NL + 'A' + NL + 'MOVED' + NL + 'B' + NL + 'C' + NL +
    'D' + NL + 'end.';
  New := 'unit U;' + NL + 'A' + NL + 'B' + NL + 'C' + NL + 'D' + NL +
    'MOVED' + NL + 'end.';

  Changes := DiffToChanges('U.pas', Old, New, 200, Total);
  for var C in Changes do
    Assert.AreNotEqual(C.Before, C.After,
      'a line that did not change is not a change: "' + C.Before + '"');
  Assert.AreEqual(2, Total,
    'the move is one deletion and one insertion, not the whole span');
  Assert.AreEqual(2, Integer(Length(Changes)));

  // With BLANK lines around the moved block the LCS has a choice, and it
  // used to pair a deleted blank with an inserted one - a "change" from
  // nothing to nothing, which the live preview really showed.
  Old := 'unit U;' + NL + 'A' + NL + '' + NL + 'MOVED' + NL + '' + NL +
    'B' + NL + 'end.';
  New := 'unit U;' + NL + 'A' + NL + '' + NL + 'B' + NL + '' + NL +
    'MOVED' + NL + 'end.';
  Changes := DiffToChanges('U.pas', Old, New, 200, Total);
  for var C in Changes do
    Assert.AreNotEqual(C.Before, C.After, 'nothing to nothing is no change');
  Assert.AreEqual(2, Total, 'still one deletion and one insertion');

  // ... and a real edit is still reported, with its line.
  Old := 'unit U;' + NL + 'A' + NL + 'MOVED' + NL + 'B' + NL + 'C' + NL +
    'D' + NL + 'end.';
  New := StringReplace(Old, 'C', 'C2', [rfReplaceAll]);
  Changes := DiffToChanges('U.pas', Old, New, 200, Total);
  Assert.AreEqual(1, Total, 'one line really changed');
  Assert.AreEqual('C', Changes[0].Before);
  Assert.AreEqual('C2', Changes[0].After);
  Assert.AreEqual(5, Changes[0].Line, '1-based, the line that changed');
end;

procedure TExtractAndCompletionTests.ResultIsTheFunctionsReturnValueNotAVariable;
var
  Lines: TArray<string>;
begin
  // WHICH routine the block sits in decides what Result means.
  Lines := [
    'function TFoo.Describe(const ARow: TRow): string;',          // 0
    'begin',
    'end;',
    '',
    'procedure TFoo.Paint;',                                      // 4
    'begin',
    'end;',
    '',
    'constructor TFoo.Create;',                                   // 8
    'begin',
    'end;',
    '',
    'function TFoo.Items(const AKey: string;',                    // 12
    '  AFlag: Boolean): TList<Integer>;',
    'begin',
    'end;'];
  Assert.AreEqual('string', EnclosingResultType(Lines, 0),
    'the function''s own result type');
  Assert.AreEqual('', EnclosingResultType(Lines, 4),
    'a procedure has no Result - there a local named Result is legal');
  Assert.AreEqual('', EnclosingResultType(Lines, 8),
    'a constructor''s Result is the instance and is never handed on');
  Assert.AreEqual('TList<Integer>', EnclosingResultType(Lines, 12),
    'a wrapped header, read to its own end');
  Assert.AreEqual('', EnclosingResultType(Lines, 1), 'not a header at all');

  // Is Result mentioned as CODE?
  Assert.IsTrue(BlockMentionsResult('  Result := ARow.Value;'));
  Assert.IsTrue(BlockMentionsResult('  if X then Result := ''a'';'));
  Assert.IsFalse(BlockMentionsResult('  Query.Result := 1;'),
    'a member called Result belongs to something else');
  Assert.IsFalse(BlockMentionsResult('  // Result is set below'),
    'a comment is no code');
  Assert.IsFalse(BlockMentionsResult('  Log(''Result'');'),
    'and neither is a string');
  Assert.IsFalse(BlockMentionsResult('  X := 1;'));

  // The reported shape: a plain assignment CAN become the function's result.
  Assert.IsTrue(CanReturnViaResult(
    '  Result := ARow.Value + IntToStr(FColumns);', 'Result'),
    'the forum case: the block assigns Result, so the extracted routine ' +
    'is a function of that type');

  // What must be refused instead of carried over wrongly.
  Assert.IsFalse(CanReturnViaResult('  Result := Result + ''x'';', 'Result'),
    'it reads the value the function already has');
  Assert.IsFalse(CanReturnViaResult(
    '  if AFlag then' + NL + '    Result := ''a'';', 'Result'),
    'a conditional assignment leaves the old value in place on one path');
  Assert.IsFalse(CanReturnViaResult(
    '  Writeln(Result);' + NL + '  Result := ''b'';', 'Result'),
    'the first mention reads it');
end;

{ TLineEndingCheckTests }

procedure TLineEndingCheckTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'rl_eol_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
end;

procedure TLineEndingCheckTests.TearDown;
begin
  try
    if TDirectory.Exists(FDir) then TDirectory.Delete(FDir, True);
  except
  end;
end;

function TLineEndingCheckTests.WriteRaw(const AName: string;
  const ABytes: TBytes): string;
begin
  Result := TPath.Combine(FDir, AName);
  TFile.WriteAllBytes(Result, ABytes);
end;

procedure TLineEndingCheckTests.MixedAndLfOnlyAreNamedAsTheProblemTheyAre;
var
  Info: TLineEndingInfo;

  function Classify(const AText: string): TLineEndingInfo;
  begin
    Result := ClassifyLineEndings(TEncoding.UTF8.GetBytes(AText));
  end;

begin
  // The normal case must stay quiet - this check runs on every file the
  // IDE opens.
  Info := Classify('unit A;' + #13#10 + 'end.' + #13#10);
  Assert.IsTrue(Info.Kind = lekCrLf, 'all CRLF');
  Assert.IsFalse(Info.NeedsFix, 'nothing to fix');
  Assert.AreEqual('', LineEndingProblemText(Info), 'and nothing to say');

  // A one-line file has no line break at all and is not a problem either.
  Info := Classify('unit A;');
  Assert.IsTrue(Info.Kind = lekNone, 'no break');
  Assert.IsFalse(Info.NeedsFix, 'no break is no defect');

  // The git-checkout case.
  Info := Classify('unit A;' + #10 + 'end.' + #10);
  Assert.IsTrue(Info.Kind = lekLf, 'LF only');
  Assert.IsTrue(Info.NeedsFix, 'LF only must be offered');
  Assert.IsTrue(LineEndingProblemText(Info).Contains('LF only'),
    'the text names LF: ' + LineEndingProblemText(Info));

  // CR only.
  Info := Classify('unit A;' + #13 + 'end.' + #13);
  Assert.IsTrue(Info.Kind = lekCr, 'CR only');
  Assert.IsTrue(Info.NeedsFix, 'CR only must be offered');

  // THE REPORTED CASE: more than one sort in ONE file, which is what the
  // IDE handles worst. The counts have to appear, because they are what
  // tells the user it is the half-converted file they feared.
  Info := Classify('a' + #13#10 + 'b' + #10 + 'c' + #13 + 'd' + #13#10);
  Assert.IsTrue(Info.Kind = lekMixed, 'mixed');
  Assert.AreEqual(2, Info.CrLf, 'two CRLF');
  Assert.AreEqual(1, Info.LfOnly, 'one LF');
  Assert.AreEqual(1, Info.CrOnly, 'one CR');
  var Text := LineEndingProblemText(Info);
  Assert.IsTrue(Text.Contains('MIXED'), 'said plainly: ' + Text);
  Assert.IsTrue(Text.Contains('2 CRLF') and Text.Contains('1 LF') and
    Text.Contains('1 CR'), 'with the counts: ' + Text);
end;

procedure TLineEndingCheckTests.NothingButSourceTextMayBeRewritten;
var
  Dfm: TBytes;
begin
  Assert.IsTrue(LineEndingCheckCandidate('C:\x\U.pas'), '.pas');
  Assert.IsTrue(LineEndingCheckCandidate('C:\x\U.INC'), '.inc, any case');
  Assert.IsTrue(LineEndingCheckCandidate('C:\x\P.dpr'), '.dpr');
  Assert.IsTrue(LineEndingCheckCandidate('C:\x\P.dpk'), '.dpk');
  // A FORM FILE IS THE DANGEROUS ONE: a binary .dfm is bytes, and turning
  // an $0A inside it into $0D$0A would destroy the form.
  Assert.IsFalse(LineEndingCheckCandidate('C:\x\U.dfm'), '.dfm never');
  Assert.IsFalse(LineEndingCheckCandidate('C:\x\U.fmx'), '.fmx never');
  // The IDE rewrites these itself.
  Assert.IsFalse(LineEndingCheckCandidate('C:\x\P.dproj'), '.dproj never');
  Assert.IsFalse(LineEndingCheckCandidate('C:\x\readme.txt'), '.txt never');

  // The second gate, for a form saved under a source name: $FF 'TPF0'.
  Dfm := [$FF, Ord('T'), Ord('P'), Ord('F'), Ord('0'), 1, 2, 10, 3];
  Assert.IsTrue(LooksBinaryContent(Dfm), 'a binary form is binary');
  Assert.IsFalse(LooksBinaryContent(TEncoding.UTF8.GetBytes('unit A;'#13#10)),
    'plain source is not');
  // UTF-16 text is full of NUL bytes and must NOT read as binary - both
  // the counter and the converter handle it by code unit.
  Assert.IsFalse(LooksBinaryContent(TEncoding.Unicode.GetPreamble +
    TEncoding.Unicode.GetBytes('unit A;'#13#10)), 'UTF-16 text is text');
  Assert.IsTrue(LooksBinaryContent([Ord('M'), Ord('Z'), 0, 1, 2]),
    'a NUL without a UTF-16 BOM is binary');
end;

procedure TLineEndingCheckTests.TheFixProducesCrlfAndChangesNothingElse;
var
  Path: string;
  Before, After: TBytes;
  Info: TLineEndingInfo;
begin
  // UTF-8 WITH BOM and mixed endings - our own sources' encoding, and the
  // shape a half-converted checkout has.
  Before := TEncoding.UTF8.GetPreamble +
    TEncoding.UTF8.GetBytes('unit A;' + #13#10 + '// Grösse' + #10 +
      'implementation' + #13 + 'end.' + #10);
  Path := WriteRaw('U.pas', Before);

  Assert.IsTrue(ConvertFileToCrLf(Path), 'it had something to fix');
  After := TFile.ReadAllBytes(Path);
  Info := ClassifyLineEndings(After);
  Assert.IsTrue(Info.Kind = lekCrLf, 'all CRLF afterwards');
  Assert.AreEqual(4, Info.CrLf, 'all four breaks');

  // The ENCODING and the content must survive - this rewrites bytes.
  Assert.IsTrue((After[0] = $EF) and (After[1] = $BB) and (After[2] = $BF),
    'the BOM is still there');
  var Txt := TEncoding.UTF8.GetString(After, 3, Length(After) - 3);
  Assert.IsTrue(Txt.Contains('Grösse'), 'the umlaut survived: ' + Txt);
  Assert.AreEqual('unit A;|// Grösse|implementation|end.|',
    Txt.Replace(#13#10, '|'), 'nothing but the breaks changed');

  // Idempotent: nothing left to do means False, not another rewrite.
  Assert.IsFalse(ConvertFileToCrLf(Path), 'a second run changes nothing');
  // '=' on two dynamic arrays compares REFERENCES, so it is always False
  // for two separate reads - my first version of this assertion failed for
  // that reason alone. Compare the bytes.
  var Again := TFile.ReadAllBytes(Path);
  Assert.AreEqual(Length(After), Length(Again), 'same length');
  Assert.IsTrue((Length(Again) = 0) or
    CompareMem(@After[0], @Again[0], Length(After)), 'and writes nothing');
end;

procedure TLineEndingCheckTests.CancelDecidesNothingAndLeaveItDoes;
var
  E: TLineEndingEffect;
begin
  // THE REPORTED CASE: Cancel changes nothing AND settles nothing.
  E := LineEndingChoiceEffect(lecCancel);
  Assert.IsFalse(E.Correct, 'cancel corrects nothing');
  Assert.IsFalse(E.AlwaysThisSession, 'cancel stores nothing');
  Assert.IsFalse(E.AlwaysSaved, 'cancel saves nothing');
  Assert.IsFalse(E.DoNotAskAgain,
    'cancel must ask again next time - that is the whole difference to Leave it');

  // ... and "Leave it" is the answer that does silence the file.
  E := LineEndingChoiceEffect(lecLeave);
  Assert.IsFalse(E.Correct, 'leave it corrects nothing');
  Assert.IsTrue(E.DoNotAskAgain, 'leave it is a decision');
  Assert.IsFalse(E.AlwaysThisSession, 'and says nothing about other files');

  // The three correcting answers, and only the last one is persisted.
  E := LineEndingChoiceEffect(lecCorrectOnce);
  Assert.IsTrue(E.Correct, 'correct once corrects');
  Assert.IsFalse(E.AlwaysThisSession, 'and asks again for the next file');
  E := LineEndingChoiceEffect(lecCorrectSession);
  Assert.IsTrue(E.Correct and E.AlwaysThisSession, 'correct + this session');
  Assert.IsFalse(E.AlwaysSaved, 'but not beyond it');
  E := LineEndingChoiceEffect(lecCorrectAlways);
  Assert.IsTrue(E.Correct and E.AlwaysThisSession and E.AlwaysSaved,
    'correct + always, saved');

  // The dialog's result mapped to an answer: not confirmed is cancel
  // whatever the radio buttons say (Escape and the close box land here
  // too), and an index we never offered is no answer either.
  Assert.IsTrue(LineEndingChoiceOfIndex(0, False) = lecCancel,
    'not confirmed on option 1 is still cancel');
  Assert.IsTrue(LineEndingChoiceOfIndex(3, False) = lecCancel,
    'not confirmed on "leave it" is still cancel');
  Assert.IsTrue(LineEndingChoiceOfIndex(0, True) = lecCorrectOnce, 'option 1');
  Assert.IsTrue(LineEndingChoiceOfIndex(1, True) = lecCorrectSession, 'option 2');
  Assert.IsTrue(LineEndingChoiceOfIndex(2, True) = lecCorrectAlways, 'option 3');
  Assert.IsTrue(LineEndingChoiceOfIndex(3, True) = lecLeave, 'option 4');
  Assert.IsTrue(LineEndingChoiceOfIndex(9, True) = lecCancel,
    'an answer we did not offer changes nothing');
end;

procedure TLineEndingCheckTests.TheProjectCheckReportsOnlyWhatItMayRewrite;
var
  Issues: TArray<TLineEndingIssue>;
  Files: TArray<string>;
  Reader: TLineEndingReader;

  function Add(const AName: string; const ABytes: TBytes): string;
  begin
    Result := WriteRaw(AName, ABytes);
    Files := Files + [Result];
  end;

  function Utf8(const AText: string): TBytes;
  begin
    Result := TEncoding.UTF8.GetBytes(AText);
  end;

  function Found(const AFile: string): Boolean;
  begin
    Result := False;
    for var Issue in Issues do
      if SameText(Issue.FileName, AFile) then Exit(True);
  end;

begin
  Files := nil;
  var Clean := Add('Clean.pas', Utf8('unit Clean;'#13#10'end.'#13#10));
  var Lf := Add('Lf.pas', Utf8('unit Lf;'#10'end.'#10));
  var Mixed := Add('Mixed.pas', Utf8('a'#13#10'b'#10'c'#13#10));
  var Inc_ := Add('Part.inc', Utf8('const X = 1;'#10));
  // A TEXT .dfm with LF endings - still none of our business, because the
  // gate cannot tell it from a binary one by name alone.
  var Dfm := Add('Form.dfm', Utf8('object Form1: TForm1'#10'end'#10));
  // A binary form that somebody saved under a source name: the second gate.
  var Bin := Add('Binary.pas', TBytes.Create($FF, Ord('T'), Ord('P'), Ord('F'),
    Ord('0'), 1, 10, 2, 10));
  var Gone := TPath.Combine(FDir, 'NoSuchFile.pas');
  Files := Files + [Gone];
  // The same file twice - a project list plus its include files can name
  // one file twice, and one row per file is what the dialog shows.
  Files := Files + [Lf];

  Reader :=
    function(const AFile: string; out ABytes: TBytes): Boolean
    begin
      ABytes := nil;
      if not TFile.Exists(AFile) then Exit(False);
      ABytes := TFile.ReadAllBytes(AFile);
      Result := True;
    end;
  Issues := CollectLineEndingIssues(Files, Reader);

  Assert.IsTrue(Found(Lf), 'an LF-only unit is a finding');
  Assert.IsTrue(Found(Mixed), 'a mixed unit is a finding');
  Assert.IsTrue(Found(Inc_), 'an include file counts too');
  Assert.IsFalse(Found(Clean), 'a CRLF unit is not');
  Assert.IsFalse(Found(Dfm), 'a form file is NEVER rewritten');
  Assert.IsFalse(Found(Bin), 'and neither is anything binary');
  Assert.IsFalse(Found(Gone), 'a file that is not there is no finding');
  Assert.AreEqual(3, Integer(Length(Issues)),
    'three findings, and the duplicate is listed once');

  // The id names the PLACE, so it survives a re-scan that reorders the
  // list - that is what an apply addresses.
  Assert.AreEqual(LineEndingIssueId(Lf), LineEndingIssueId(Lf),
    'the same file, the same id');
  Assert.AreEqual(LineEndingIssueId(Lf), LineEndingIssueId(UpperCase(Lf)),
    'and the case of the path does not change it');
  Assert.AreNotEqual(LineEndingIssueId(Lf), LineEndingIssueId(Mixed),
    'two files, two ids');
  for var Issue in Issues do
    Assert.AreEqual(LineEndingIssueId(Issue.FileName), Issue.Id,
      'the issue carries its own id');
end;

procedure TLineEndingCheckTests.TheLineEndingToolDeclaresWhatItTakes;
var
  Args: TArray<string>;

  function Declares(const AName: string): Boolean;
  begin
    Result := False;
    for var A in Args do
      if SameText(A, AName) then Exit(True);
  end;

begin
  Args := KnownToolArguments('line_endings');
  Assert.IsTrue(Length(Args) > 0, 'the tool is in the shipped definitions');
  // The handler reads these two, so the schema must name them - an
  // argument the schema does not list is dropped before the handler ever
  // sees it, and the caller is told nothing (1.16.11).
  Assert.IsTrue(Declares('apply'), 'apply is declared');
  Assert.IsTrue(Declares('fix_ids'), 'fix_ids is declared');
  Assert.IsTrue(Declares('max'), 'max is declared');
  Assert.IsFalse(Declares('files'), 'and nothing it does not read');
end;

procedure TLineEndingCheckTests.WhatWasActedOnIsNeverCutFromTheAnswer;
var
  Issues: TArray<TLineEndingIssue>;

  function Issue(const AFile: string): TLineEndingIssue;
  begin
    Result := Default(TLineEndingIssue);
    Result.FileName := AFile;
    Result.Info := ClassifyLineEndings(TEncoding.UTF8.GetBytes('a'#10'b'#10));
    Result.Id := LineEndingIssueId(AFile);
  end;

  function Ids(const ARows: TArray<Integer>): string;
  begin
    Result := '';
    for var R in ARows do
      Result := Result + Issues[R].FileName + ' ';
  end;

begin
  Issues := [Issue('A.pas'), Issue('B.pas'), Issue('C.pas'), Issue('D.pas')];

  // The reported shape: one id named, a budget of three. The acted row is
  // the LAST finding, so the old "first N" rule dropped exactly it.
  var Rows := SelectLineEndingRows(Issues, [Issues[3].Id], 3);
  Assert.IsTrue(Ids(Rows).Contains('D.pas'),
    'the file that was corrected must be in the answer');
  Assert.AreEqual('A.pas B.pas C.pas D.pas ', Ids(Rows),
    'and it does not cost one of the other rows, nor jump to the front');

  // A budget of zero is how a caller asks for counts only - but what it
  // just corrected is not a row it can do without.
  Rows := SelectLineEndingRows(Issues, [Issues[1].Id], 0);
  Assert.AreEqual('B.pas ', Ids(Rows), 'only the acted one, and that one');
  Assert.AreEqual('', Ids(SelectLineEndingRows(Issues, nil, 0)),
    'nothing acted on, no budget - no rows');
  Assert.AreEqual('A.pas B.pas ', Ids(SelectLineEndingRows(Issues, nil, 2)),
    'and a plain listing is still the first N');
  Assert.AreEqual('A.pas B.pas C.pas D.pas ',
    Ids(SelectLineEndingRows(Issues, nil, 99)),
    'a budget nobody reaches lists everything');
  // The id comes back from a client that may upper-case it.
  Assert.AreEqual('C.pas ',
    Ids(SelectLineEndingRows(Issues, [UpperCase(Issues[2].Id)], 0)),
    'the case of the id does not decide whether the row survives');

  // An id of an older scan: nothing was done for it, and the answer has to
  // say so - the note alone would read as "your file is CRLF now".
  var Unknown := UnknownLineEndingIds(Issues, [Issues[0].Id, 'DEADBEEF', '']);
  Assert.AreEqual(1, Integer(Length(Unknown)), 'one id named nothing');
  Assert.AreEqual('DEADBEEF', Unknown[0], 'and it is named');
  Assert.AreEqual(0, Integer(Length(UnknownLineEndingIds(Issues,
    [LowerCase(Issues[0].Id)]))), 'a known id in another case is known');
end;

{ TDeclareConstantTests }

// The fixture every case works on: a number used twice in one routine, a
// string literal, and a magic number in a second routine.
function TDeclareConstantTests.Unit1: TArray<string>;
begin
  Result := [
    'unit U1;',                                    //  0
    '',                                            //  1
    'interface',                                   //  2
    '',                                            //  3
    'uses',                                        //  4
    '  System.SysUtils;',                          //  5
    '',                                            //  6
    'type',                                        //  7
    '  TWorker = class',                           //  8
    '  private',                                   //  9
    '    FCount: Integer;',                        // 10
    '  public',                                    // 11
    '    procedure Run;',                          // 12
    '  end;',                                      // 13
    '',                                            // 14
    'implementation',                              // 15
    '',                                            // 16
    'procedure TWorker.Run;',                      // 17
    'var',                                         // 18
    '  I: Integer;',                               // 19
    'begin',                                       // 20
    '  Timeout := 5000;',                          // 21
    '  if I > 5000 then',                           // 22
    '    ShowMessage(''Datei nicht gefunden'');',  // 23
    '  // 5000 in a comment stays',                // 24
    'end;',                                        // 25
    '',                                            // 26
    'end.',                                        // 27
    ''];                                           // 28
end;

function TDeclareConstantTests.Occurrences(const AText, AWhat: string): Integer;
var
  P: Integer;
begin
  Result := 0;
  P := Pos(AWhat, AText);
  while P > 0 do
  begin
    Inc(Result);
    P := Pos(AWhat, AText, P + Length(AWhat));
  end;
end;

// Goes through DeclareRequestFor, so what the editor will hand the dialog
// is what the tests exercise.
function TDeclareConstantTests.ReqFor(const ALines: TArray<string>;
  ALine0, AStart, AEnd: Integer; const AName, AType: string;
  APlace: TDeclarePlace; ASection: TDeclareSection;
  AReplaceAll: Boolean): TDeclareRequest;
begin
  Result := DeclareRequestFor(ALines, ALine0, AStart, AEnd);
  Result.Name := AName;
  Result.TypeName := AType;
  Result.Place := APlace;
  Result.Section := ASection;
  Result.ReplaceAll := AReplaceAll;
end;

function TDeclareConstantTests.ColOf(const ALine, AText: string): Integer;
begin
  Result := Pos(AText, ALine) - 1;
  Assert.IsTrue(Result >= 0, 'the fixture must contain ' + AText);
end;

procedure TDeclareConstantTests.OnlyALiteralCanBecomeAConstant;
begin
  Assert.AreEqual(Integer(Ord(dvNumber)), Integer(Ord(ClassifyDeclareValue('5000'))));
  Assert.AreEqual(Integer(Ord(dvNumber)), Integer(Ord(ClassifyDeclareValue('$FF'))));
  Assert.AreEqual(Integer(Ord(dvNumber)), Integer(Ord(ClassifyDeclareValue('2 * 1024'))),
    'a constant EXPRESSION of numbers is a number');
  Assert.AreEqual(Integer(Ord(dvString)), Integer(Ord(ClassifyDeclareValue('''abc'''))));
  Assert.AreEqual(Integer(Ord(dvString)), Integer(Ord(ClassifyDeclareValue('''a'' + ''b'''))));
  Assert.AreEqual(Integer(Ord(dvString)), Integer(Ord(ClassifyDeclareValue('#13#10'))));
  // An identifier would have to be in scope where the DECLARATION goes,
  // which is not the same place as the use site.
  Assert.AreEqual(Integer(Ord(dvOther)), Integer(Ord(ClassifyDeclareValue('2 * Factor'))),
    'an identifier is not a literal');
  Assert.AreEqual(Integer(Ord(dvOther)), Integer(Ord(ClassifyDeclareValue('1, 2'))),
    'two values are not one value');
  Assert.AreEqual(Integer(Ord(dvOther)), Integer(Ord(ClassifyDeclareValue('''a'' + 1'))),
    'a string and a number together are neither');
  Assert.AreEqual(Integer(Ord(dvNone)), Integer(Ord(ClassifyDeclareValue('   '))));
end;

procedure TDeclareConstantTests.TheNameIsOfferedFromWhatTheValueMeans;
begin
  // A string says what it means, so it names itself.
  Assert.AreEqual('SDateiNichtGefunden',
    SuggestConstantName('''Datei nicht gefunden''', '', dvString));
  Assert.AreEqual('SText', SuggestConstantName('''...''', '', dvString),
    'a literal with no words at all still gives a usable name');
  // A number says nothing, so it is named after what it is assigned to.
  Assert.AreEqual('CTimeout', SuggestConstantName('5000', '  Timeout := ', dvNumber));
  Assert.AreEqual('CTimeout', SuggestConstantName('5000', '  SetTimeout(', dvNumber),
    'Set/Get is not part of the meaning');
  Assert.AreEqual('CLimit', SuggestConstantName('5000', '  if Limit > ', dvNumber),
    'a comparison names what is compared');
  Assert.AreEqual('CValue', SuggestConstantName('5000', '  X := Y + ', dvNumber), 'without an identifier to the left there is nothing to go on');
end;

procedure TDeclareConstantTests.ALocalConstantLandsInItsOwnSection;
var
  Lines: TArray<string>;
  Plan: TDeclarePlan;
  Why: string;
  Col: Integer;
begin
  Lines := Unit1;
  Col := ColOf(Lines[21], '5000');
  Assert.IsTrue(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'CTimeout', '',
    dplLocal, dsConst, False), Plan, Why), Why);
  var Text := string.Join(#10, Plan.NewLines);
  // The routine already has a VAR section, so the const needs its own - a
  // const cannot be added under a var.
  Assert.IsTrue(Text.Contains('const' + #10 + '  CTimeout = 5000;'),
    'a const section of its own: ' + Plan.DeclText);
  Assert.IsTrue(Text.Contains('  Timeout := CTimeout;'), 'the value is replaced');
  Assert.IsTrue(Text.Contains('  if I > 5000 then'),
    'and ONLY the selected occurrence, because replace-all was off');
  Assert.AreEqual(1, Plan.Replaced);
  // ...and with replace-all the second use goes too, but the one in the
  // comment never does.
  Assert.IsTrue(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'CTimeout', '',
    dplLocal, dsConst, True), Plan, Why), Why);
  Text := string.Join(#10, Plan.NewLines);
  Assert.AreEqual(2, Plan.Replaced, 'both uses in the routine');
  Assert.IsTrue(Text.Contains('  if I > CTimeout then'), 'the second use too');
  Assert.IsTrue(Text.Contains('// 5000 in a comment stays'),
    'a comment is not code');
end;

procedure TDeclareConstantTests.AResourceStringIsAUnitLevelStringOnly;
var
  Lines: TArray<string>;
  Plan: TDeclarePlan;
  Why: string;
  Col: Integer;
begin
  Lines := Unit1;
  Col := ColOf(Lines[23], '''Datei');
  // What Delphi allows: a string, at unit level, untyped.
  Assert.IsTrue(PlanDeclare(Lines, ReqFor(Lines, 23, Col, Col + Length('''Datei nicht gefunden'''), 'SNotFound', '',
    dplImplementation, dsResourceString, False), Plan, Why), Why);
  var Text := string.Join(#10, Plan.NewLines);
  Assert.IsTrue(Text.Contains('resourcestring' + #10 +
    '  SNotFound = ''Datei nicht gefunden'';'), 'the section is created: ' + Plan.DeclText);
  Assert.IsTrue(Text.Contains('ShowMessage(SNotFound);'), 'and used at the site');
  // ...and what it does not.
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 23, Col, Col + Length('''Datei nicht gefunden'''), 'SNotFound', '',
    dplLocal, dsResourceString, False), Plan, Why));
  Assert.IsTrue(Why.Contains('unit level'), Why);
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 23, Col, Col + Length('''Datei nicht gefunden'''), 'SNotFound', 'string',
    dplImplementation, dsResourceString, False), Plan, Why));
  Assert.IsTrue(Why.Contains('no type'), Why);
  Col := ColOf(Lines[21], '5000');
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'SNotFound', '',
    dplImplementation, dsResourceString, False), Plan, Why));
  Assert.IsTrue(Why.Contains('string literal'), Why);
end;

procedure TDeclareConstantTests.WhatWouldNotCompileIsRefusedWithAReason;
var
  Lines: TArray<string>;
  Plan: TDeclarePlan;
  Why: string;
  Col: Integer;
begin
  Lines := Unit1;
  Col := ColOf(Lines[21], '5000');

  // A name that is not a name.
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'begin', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('reserved word'), Why);
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, '2Foo', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('identifier'), Why);

  // A name that is taken: at unit level the whole unit counts, locally the
  // routine - and "I" is a local of exactly this routine.
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'I', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('already occurs'), Why);
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 4, 'FCount', '',
    dplInterface, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('already occurs'), Why);

  // The number in the COMMENT is not code.
  var CCol := ColOf(Lines[24], '5000');
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 24, CCol, CCol + 4, 'CX', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('comment'), Why);

  // Half a literal is not a literal.
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col + 2, 'CX', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('exactly one'), Why);

  // Nothing selected and nothing typed: since the value is the user's own
  // field, that is the question to ask - not "select something".
  Assert.IsFalse(PlanDeclare(Lines, ReqFor(Lines, 21, Col, Col, 'CX', '',
    dplLocal, dsConst, False), Plan, Why));
  Assert.IsTrue(Why.Contains('enter a value'), Why);
end;

procedure TDeclareConstantTests.AUnitConstantJoinsTheSectionThatIsThere;
var
  Lines: TArray<string>;
  Plan: TDeclarePlan;
  Why: string;
  Col: Integer;
begin
  // A unit that already has an implementation const section: the new
  // declaration joins it instead of opening a second one.
  Lines := [
    'unit U2;',                       // 0
    'interface',                      // 1
    'implementation',                 // 2
    '',                               // 3
    'const',                          // 4
    '  CFirst = 1;',                  // 5
    '',                               // 6
    'procedure Go;',                  // 7
    'begin',                          // 8
    '  Sleep(250);',                  // 9
    'end;',                           // 10
    '',                               // 11
    'end.',                           // 12
    ''];
  Col := ColOf(Lines[9], '250');
  Assert.IsTrue(PlanDeclare(Lines, ReqFor(Lines, 9, Col, Col + 3, 'CDelay', '',
    dplImplementation, dsConst, False), Plan, Why), Why);
  var Text := string.Join(#10, Plan.NewLines);
  Assert.AreEqual(1, Occurrences(Text, 'const'), 'no second const section');
  Assert.IsTrue(Text.Contains('  CFirst = 1;' + #10 + '  CDelay = 250;'),
    'it joins the one that is there: ' + Text);
  Assert.IsTrue(Text.Contains('  Sleep(CDelay);'));
  // A typed one is written as a typed constant.
  Assert.IsTrue(PlanDeclare(Lines, ReqFor(Lines, 9, Col, Col + 3, 'CDelay', 'Cardinal',
    dplImplementation, dsConst, False), Plan, Why), Why);
  Assert.IsTrue(string.Join(#10, Plan.NewLines).Contains('  CDelay: Cardinal = 250;'),
    Plan.DeclText);
end;

procedure TDeclareConstantTests.TheConstantToolDeclaresWhatItTakes;
var
  Args: TArray<string>;

  function Declares(const AName: string): Boolean;
  begin
    Result := False;
    for var A in Args do
      if SameText(A, AName) then Exit(True);
  end;

begin
  Args := KnownToolArguments('declare_constant');
  Assert.IsTrue(Length(Args) > 0, 'the tool is in the shipped definitions');
  // Exactly what ToolDeclareConstant reads.
  for var Needed in ['file', 'line', 'literal', 'column',
    'end_column', 'value', 'name', 'type', 'section', 'place',
    'replace_all', 'apply', 'token'] do
    Assert.IsTrue(Declares(Needed), Needed + ' is declared');
  Assert.IsFalse(Declares('expression'),
    'and nothing it does not read - that is extract_variable''s name for it');
end;

procedure TDeclareConstantTests.AClassMemberIsNoNestedRoutine;
var
  Lines: TArray<string>;
  First, Last: Integer;
begin
  // Found by the old-vs-new sweep over real sources (UGTIKim.pas:982)
  // while fixing the nested-routine report: a class declared in the
  // IMPLEMENTATION section has method declarations that look exactly like
  // a nested routine''s header. Reading them as nested made the walk run
  // through the whole implementation section, so a blank line between two
  // routines answered with a 2000-line "routine". What separates the two
  // is that a nested routine is INDENTED deeper than its own header.
  Lines := [
    'unit U;',                                   //  0
    'interface',                                 //  1
    'implementation',                            //  2
    '',                                          //  3
    'type',                                      //  4
    '  TEvents = class',                         //  5
    '    procedure Handle(A: TObject);',         //  6  <- a MEMBER
    '  end;',                                    //  7
    '',                                          //  8
    'procedure TEvents.Handle(A: TObject);',     //  9
    'begin',                                     // 10
    '  DoIt;',                                   // 11
    'end;',                                      // 12
    '',                                          // 13
    'procedure Later;',                          // 14
    'begin',                                     // 15
    '  DoMore;',                                 // 16
    'end;',                                      // 17
    '',                                          // 18
    'end.',                                      // 19
    ''];

  Assert.IsFalse(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 6,
    First, Last), 'a class member declaration is no routine body');
  Assert.IsFalse(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 13,
    First, Last), 'and the blank line between two routines is in none');
  Assert.IsFalse(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 8,
    First, Last), 'nor the one below the class');
  // The real routines are still found, and each one only to its own end.
  Assert.IsTrue(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 11,
    First, Last));
  Assert.AreEqual(9, First);
  Assert.AreEqual(12, Last, 'its OWN end, not the next routine''s');
  Assert.IsTrue(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 16,
    First, Last));
  Assert.AreEqual(14, First);
end;

procedure TDeclareConstantTests.ANestedRoutineDoesNotHideTheOneAroundIt;
var
  Lines: TArray<string>;
  First, Last: Integer;
  Plan: TDeclarePlan;
  Req: TDeclareRequest;
  Why: string;
  Col: Integer;
begin
  // The reported shape (screenshot 2026-10-16, KimSendMail): a routine
  // whose declaration part holds nested routines, and the caret far below
  // them in the OUTER body. "the selection is not inside a routine" was
  // the answer, and the user named the cause.
  Lines := [
    'unit U;',                                   //  0
    'interface',                                 //  1
    'implementation',                            //  2
    '',                                          //  3
    'procedure KimSendMail;',                    //  4
    '',                                          //  5
    '  procedure Inner1;',                       //  6
    '  begin',                                   //  7
    '    DoA;',                                  //  8
    '  end;',                                    //  9
    '',                                          // 10
    '  function Inner2: Integer;',               // 11
    '  begin',                                   // 12
    '    Result := 1;',                          // 13
    '  end;',                                    // 14
    '',                                          // 15
    'var',                                       // 16
    '  I: Integer;',                             // 17
    'begin',                                     // 18
    '  LastErrorMessage := ''no recipient'';',   // 19
    '  Done;',                                   // 20
    'end;',                                      // 21
    '',                                          // 22
    'procedure After;',                          // 23
    'begin',                                     // 24
    '  X := 1;',                                 // 25
    'end;',                                      // 26
    '',                                          // 27
    'end.',                                      // 28
    ''];

  // The routine around line 19 is KimSendMail, not Inner2.
  Assert.IsTrue(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 19,
    First, Last), 'line 19 IS inside a routine');
  Assert.AreEqual(4, First, 'the OUTER header, not the last nested one');
  Assert.AreEqual(21, Last, 'and its own end');

  // Inside a nested routine the answer is still that nested routine.
  Assert.IsTrue(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 13,
    First, Last));
  Assert.AreEqual(11, First, 'a line inside Inner2 belongs to Inner2');
  Assert.AreEqual(14, Last);

  // A line between two routines belongs to neither.
  Assert.IsFalse(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 22,
    First, Last), 'the blank line between two routines is in no routine');

  // The routine AFTER one with nested routines is still found.
  Assert.IsTrue(FindEnclosingRoutineRange(string.Join(#13#10, Lines), 25,
    First, Last));
  Assert.AreEqual(23, First);

  // ...and that is what the dialog needs: a local const in KimSendMail.
  Col := ColOf(Lines[19], '''no recipient''');
  Req := DeclareRequestFor(Lines, 19, Col, Col + Length('''no recipient'''));
  Req.Name := 'SNoRecipient';
  Req.Place := dplLocal;
  Assert.IsTrue(PlanDeclare(Lines, Req, Plan, Why), Why);
  var Text := string.Join(#10, Plan.NewLines);
  // Its own const section, after the nested routines and right before the
  // outer 'begin' - which is where Delphi allows it and where the nested
  // routines cannot see it (they must not; the use site is in the body).
  Assert.IsTrue(Text.Contains('const' + #10 + '  SNoRecipient = ''no recipient'';'),
    'its own const section in the OUTER routine: ' + Plan.DeclText);
  Assert.IsTrue(Text.IndexOf('const' + #10 + '  SNoRecipient') >
    Text.IndexOf('function Inner2'), 'and it stands AFTER the nested ones');
  Assert.IsTrue(Text.Contains('LastErrorMessage := SNoRecipient;'),
    'and the use site');
  Assert.IsTrue(Text.Contains('  function Inner2: Integer;'),
    'the nested routines are untouched');
end;

procedure TDeclareConstantTests.TheValueIsEditableAndMayComeFromNowhere;
var
  Lines: TArray<string>;
  Req: TDeclareRequest;
  Plan: TDeclarePlan;
  Why: string;
  Col: Integer;
begin
  Lines := Unit1;

  // The selection OFFERS the value and the name - that is all it does.
  Col := ColOf(Lines[21], '5000');
  Req := DeclareRequestFor(Lines, 21, Col, Col + 4);
  Assert.IsTrue(Req.HasSelection);
  Assert.AreEqual('5000', Req.Value, 'the value is offered from the selection');
  Assert.AreEqual('CTimeout', Req.Name, 'and so is the name');

  // ...and the user may change it. The DECLARATION gets the new value, the
  // code gets the name where the selection was.
  Req.Value := '$1388';
  Req.Name := 'CTimeout';
  Req.Place := dplLocal;
  Assert.IsTrue(PlanDeclare(Lines, Req, Plan, Why), Why);
  var Text := string.Join(#10, Plan.NewLines);
  Assert.IsTrue(Text.Contains('  CTimeout = $1388;'),
    'the edited value is what is declared: ' + Plan.DeclText);
  Assert.IsTrue(Text.Contains('  Timeout := CTimeout;'), 'and the use site');
  Assert.AreEqual(1, Plan.Replaced);

  // An edited value still has to BE a value.
  Req.Value := 'Factor * 2';
  Assert.IsFalse(PlanDeclare(Lines, Req, Plan, Why));
  Assert.IsTrue(Why.Contains('must be a literal'), Why);

  // FROM SCRATCH: no selection. The caret still says which routine dplLocal
  // means, but nothing in the code is touched.
  Req := DeclareRequestFor(Lines, 21, 0, 0);
  Assert.IsFalse(Req.HasSelection, 'nothing selected');
  Assert.AreEqual('CValue', Req.Name, 'a name to start from');
  Req.Name := 'CRetries';
  Req.Value := '3';
  Req.Place := dplLocal;
  Assert.IsTrue(PlanDeclare(Lines, Req, Plan, Why), Why);
  Text := string.Join(#10, Plan.NewLines);
  Assert.IsTrue(Text.Contains('  CRetries = 3;'), Plan.DeclText);
  Assert.AreEqual(0, Plan.Replaced, 'nothing was replaced');
  Assert.IsTrue(Plan.Note.Contains('nothing in the code'), Plan.Note);
  Assert.IsTrue(Text.Contains('  Timeout := 5000;'),
    'the line the caret was on is untouched');
  // Replace-all cannot mean anything without a selection, so it is ignored
  // rather than guessing which literal was meant.
  Req.ReplaceAll := True;
  Assert.IsTrue(PlanDeclare(Lines, Req, Plan, Why), Why);
  Assert.AreEqual(0, Plan.Replaced, 'and it stays ignored');
  // From scratch at unit level needs no routine at all.
  Req := DeclareRequestFor(Lines, 0, 0, 0);
  Req.Name := 'SCaption';
  Req.Value := '''Titel''';
  Req.Place := dplInterface;
  Req.Section := dsResourceString;
  Assert.IsTrue(PlanDeclare(Lines, Req, Plan, Why), Why);
  Assert.IsTrue(string.Join(#10, Plan.NewLines).Contains(
    'resourcestring' + #10 + '  SCaption = ''Titel'';'), Plan.DeclText);
  // ...and an empty value is the one thing it cannot do.
  Req.Value := '  ';
  Assert.IsFalse(PlanDeclare(Lines, Req, Plan, Why));
  Assert.IsTrue(Why.Contains('enter a value'), Why);
end;

{ TFileEncodingCheckTests }

function TFileEncodingCheckTests.WriteRawEnc(const AName: string;
  const ABytes: TBytes): string;
begin
  Result := TPath.Combine(FDir, AName);
  TFile.WriteAllBytes(Result, ABytes);
end;

procedure TFileEncodingCheckTests.Setup;
begin
  FDir := TPath.Combine(TPath.GetTempPath, 'rl_enc_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(FDir);
end;

procedure TFileEncodingCheckTests.TearDown;
begin
  try
    if TDirectory.Exists(FDir) then TDirectory.Delete(FDir, True);
  except
  end;
end;

procedure TFileEncodingCheckTests.WhatAFileISAndWhyItMatters;
const
  OE = #$00F6;   // LATIN SMALL LETTER O WITH DIAERESIS
var
  Utf8: TBytes;
begin
  // The measured case (2026-10-16, dcc32 37.0): the SAME literal compiled
  // from a BOM-less UTF-8 file gave 71 114 195 182 115 115 101 - two
  // characters where one belongs - while a BOM and an ANSI file both gave
  // 71 114 246 115 115 101. So these two classes are not the same question.
  Utf8 := TEncoding.UTF8.GetBytes('Gr' + OE + 'sse');
  Assert.AreEqual(Integer(Ord(secUtf8NoBom)),
    Integer(Ord(ClassifySourceEncoding(Utf8))),
    'valid UTF-8 with high bytes and no BOM - the trap');
  Assert.AreEqual(Integer(Ord(secUtf8Bom)),
    Integer(Ord(ClassifySourceEncoding(TBytes.Create($EF, $BB, $BF) + Utf8))));
  Assert.AreEqual(Integer(Ord(secAnsi)),
    Integer(Ord(ClassifySourceEncoding(TBytes.Create(71, 114, $F6, 115)))),
    'a lone high byte is not UTF-8');
  Assert.AreEqual(Integer(Ord(secUtf16)),
    Integer(Ord(ClassifySourceEncoding(TBytes.Create($FF, $FE, 71, 0)))));
  // ASCII is its OWN answer: calling it "UTF-8 without BOM" would make
  // every plain unit a finding, and converting it changes no byte.
  Assert.AreEqual(Integer(Ord(secAscii)),
    Integer(Ord(ClassifySourceEncoding(TEncoding.ASCII.GetBytes('unit U;')))));
  Assert.AreEqual(Integer(Ord(secAscii)),
    Integer(Ord(ClassifySourceEncoding(nil))), 'nothing is ASCII, not ANSI');
  Assert.IsTrue(SourceEncodingNote(secUtf8NoBom).Contains('ANSI'),
    'the note says what the compiler does: ' + SourceEncodingNote(secUtf8NoBom));
end;

procedure TFileEncodingCheckTests.OnlyWhatWasAskedForAndMayBeRewritten;
const
  OE = #$00F6;
var
  Issues: TArray<TEncodingIssue>;
  Files: TArray<string>;
  Reader: TEncodingReader;

  function Add(const AName: string; const ABytes: TBytes): string;
  begin
    Result := WriteRawEnc(AName, ABytes);
    Files := Files + [Result];
  end;

  function Found(const AFile: string): Boolean;
  begin
    Result := False;
    for var Issue in Issues do
      if SameText(Issue.FileName, AFile) then Exit(True);
  end;

begin
  Files := nil;
  var U8 := TEncoding.UTF8.GetBytes('unit U; // Gr' + OE + 'sse');
  var NoBom := Add('NoBom.pas', U8);
  var WithBom := Add('WithBom.pas', TBytes.Create($EF, $BB, $BF) + U8);
  var Ansi := Add('Ansi.pas', TBytes.Create(117, 110, 105, 116, 32, $F6, 59));
  var Plain := Add('Plain.pas', TEncoding.ASCII.GetBytes('unit Plain;'));
  // A form file is NEVER rewritten - measured 2026-10-16: 1368 of 1381 are
  // pure ASCII anyway, so excluding them costs nothing.
  var Dfm := Add('Form.dfm', TBytes.Create(111, 98, 106, $F6, 99, 116));
  var Bin := Add('Binary.pas', TBytes.Create($FF, Ord('T'), Ord('P'), Ord('F'),
    Ord('0'), 1, $F6, 2));
  var Gone := TPath.Combine(FDir, 'NoSuchFile.pas');
  Files := Files + [Gone, NoBom];   // and the same file twice

  Reader :=
    function(const AFile: string; out ABytes: TBytes): Boolean
    begin
      ABytes := nil;
      if not TFile.Exists(AFile) then Exit(False);
      ABytes := TFile.ReadAllBytes(AFile);
      Result := True;
    end;

  // The default pair: what the compiler gets wrong, plus plain ANSI.
  Issues := CollectEncodingIssues(Files, [secUtf8NoBom, secAnsi], secUtf8Bom,
    Reader);
  Assert.IsTrue(Found(NoBom), 'BOM-less UTF-8 is a finding');
  Assert.IsTrue(Found(Ansi), 'and so is ANSI');
  Assert.IsFalse(Found(WithBom), 'a file already in the target is not');
  Assert.IsFalse(Found(Plain), 'and pure ASCII was not asked for');
  Assert.IsFalse(Found(Dfm), 'a form file is NEVER rewritten');
  Assert.IsFalse(Found(Bin), 'and neither is anything binary');
  Assert.IsFalse(Found(Gone), 'a file that is not there is no finding');
  Assert.AreEqual(2, Integer(Length(Issues)),
    'two findings, and the duplicate is listed once');

  // "von wo nach wo" is the whole point: ask for ANSI only and the
  // BOM-less file stays untouched.
  Issues := CollectEncodingIssues(Files, [secAnsi], secUtf8Bom, Reader);
  Assert.AreEqual(1, Integer(Length(Issues)));
  Assert.IsTrue(Found(Ansi));
  // ...and ASCII is reachable, but only when it is named.
  Issues := CollectEncodingIssues(Files, [secAscii], secUtf8Bom, Reader);
  Assert.IsTrue(Found(Plain), 'asked for, so offered');

  // The ids name the place.
  Assert.AreEqual(EncodingIssueId(Plain), EncodingIssueId(UpperCase(Plain)),
    'the case of the path does not change the id');
  Assert.AreNotEqual(EncodingIssueId(Plain), EncodingIssueId(Ansi));
end;

procedure TFileEncodingCheckTests.BothChecksShareOneRowRule;
var
  Ids: TArray<string>;

  function Picked(const ARows: TArray<Integer>): string;
  begin
    Result := '';
    for var R in ARows do Result := Result + Ids[R] + ' ';
  end;

begin
  Ids := ['A', 'B', 'C', 'D'];
  // The acted row is the LAST one, so a "first N" rule would drop exactly
  // it - which is what the line-ending tool did live (1.25.1).
  Assert.AreEqual('A B C D ', Picked(SelectReportRows(Ids, ['D'], 3)),
    'what was acted on survives the budget, and costs no other row');
  Assert.AreEqual('B ', Picked(SelectReportRows(Ids, ['b'], 0)),
    'budget 0 means counts plus what you just changed, case-insensitively');
  Assert.AreEqual('', Picked(SelectReportRows(Ids, nil, 0)));
  Assert.AreEqual('A B ', Picked(SelectReportRows(Ids, nil, 2)),
    'and a plain listing is still the first N');
  var Unknown := UnknownReportIds(Ids, ['A', 'ZZZZ', '']);
  Assert.AreEqual(1, Integer(Length(Unknown)));
  Assert.AreEqual('ZZZZ', Unknown[0], 'an id nothing carries is named');
end;

procedure TFileEncodingCheckTests.NoCheckRewritesTheInstallation;
var
  Root, Own, Shipped: string;
  Reader: TEncodingReader;
  LineReader: TLineEndingReader;
begin
  // The two files of the reported shape: one of the project, one under
  // the RAD Studio root the user had open in the editor.
  Root := TPath.Combine(FDir, 'Studio');
  TDirectory.CreateDirectory(TPath.Combine(Root, 'source'));
  // A non-ASCII character, or the bytes are pure ASCII and the file is
  // not an ANSI finding at all - the first run of this test said so.
  Own := WriteRawEnc('Own.pas',
    TEncoding.ASCII.GetBytes('unit Own; // Gr') + TBytes.Create($F6) +
    TEncoding.ASCII.GetBytes('sse'#13#10));
  Shipped := TPath.Combine(Root, 'source\IdHTTP.pas');
  TFile.WriteAllBytes(Shipped,
    TEncoding.ASCII.GetBytes('unit IdHTTP; // Gr') + TBytes.Create($F6) +
    TEncoding.ASCII.GetBytes('sse'#13#10));

  Reader :=
    function(const AFile: string; out ABytes: TBytes): Boolean
    begin
      Result := TFile.Exists(AFile);
      if Result then ABytes := TFile.ReadAllBytes(AFile);
    end;

  // Without a root both are findings - that is the state that shipped.
  var All := CollectEncodingIssues([Own, Shipped], [secAnsi], secUtf8Bom,
    Reader);
  Assert.AreEqual(2, Integer(Length(All)),
    'both are ANSI, so both qualify as far as the encoding goes');

  // With it, the installation's file is gone and the project's stays.
  var Mine := CollectEncodingIssues([Own, Shipped], [secAnsi], secUtf8Bom,
    Reader, nil, Root);
  Assert.AreEqual(1, Integer(Length(Mine)),
    'the file under the RAD Studio root is NOT a finding');
  Assert.AreEqual(LowerCase(Own), LowerCase(Mine[0].FileName),
    'and the project''s own file still is');

  // The same gate, same rule, for the line-ending check.
  LineReader :=
    function(const AFile: string; out ABytes: TBytes): Boolean
    begin
      Result := Reader(AFile, ABytes);
    end;
  var Ends := CollectLineEndingIssues(
    [WriteRawEnc('Lf.pas', TEncoding.UTF8.GetBytes('unit Lf;'#10)),
     TPath.Combine(Root, 'source\Lf2.pas')], LineReader, nil, Root);
  TFile.WriteAllBytes(TPath.Combine(Root, 'source\Lf2.pas'),
    TEncoding.UTF8.GetBytes('unit Lf2;'#10));
  Ends := CollectLineEndingIssues(
    [TPath.Combine(FDir, 'Lf.pas'), TPath.Combine(Root, 'source\Lf2.pas')],
    LineReader, nil, Root);
  Assert.AreEqual(1, Integer(Length(Ends)),
    'the line-ending check applies the same rule');
  Assert.IsFalse(Ends[0].FileName.Contains('Studio'),
    'and it is the project''s file that is left');

  // A path that only LOOKS like the root is none of it - "C:\Studio2"
  // must not be swallowed by "C:\Studio".
  Assert.IsFalse(UnderInstallRoot(Root + '2\X.pas', Root),
    'the root is a FOLDER, not a prefix of the name');
  Assert.IsFalse(UnderInstallRoot(Own, ''),
    'no root given = the gate is off');
end;

procedure TFileEncodingCheckTests.TheEncodingToolDeclaresWhatItTakes;
var
  Args: TArray<string>;

  function Declares(const AName: string): Boolean;
  begin
    Result := False;
    for var A in Args do
      if SameText(A, AName) then Exit(True);
  end;

begin
  Args := KnownToolArguments('file_encoding');
  Assert.IsTrue(Length(Args) > 0, 'the tool is in the shipped definitions');
  for var Needed in ['from', 'to', 'fix_ids', 'max', 'apply'] do
    Assert.IsTrue(Declares(Needed), Needed + ' is declared');
  Assert.IsFalse(Declares('files'), 'and nothing it does not read');
end;

{ TInstallOutputRootTests }


procedure TInstallOutputRootTests.AFileInUseIsMovedAsideAndReplaced;
var
  Dir, Target, Source, Aside, Stamp, DayBefore, DayAfter: string;
  H: THandle;
  Failed: Boolean;
  P: Integer;
begin
  Dir := TPath.Combine(TPath.GetTempPath, 'dih_inuse_' + TGUID.NewGuid.ToString);
  TDirectory.CreateDirectory(Dir);
  try
    Target := TPath.Combine(Dir, 'tool.exe');
    Source := TPath.Combine(Dir, 'new.bin');
    TFile.WriteAllText(Target, 'OLD');
    TFile.WriteAllText(Source, 'NEW');

    // Exactly the shape a RUNNING image has: readers and renames are
    // allowed, writers are not. FILE_SHARE_DELETE is what makes the move
    // possible while the file is in use - a test that locks it outright
    // would measure something else.
    H := CreateFile(PChar(Target), GENERIC_READ,
      FILE_SHARE_READ or FILE_SHARE_DELETE, nil, OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL, 0);
    Assert.IsTrue(H <> INVALID_HANDLE_VALUE, 'the probe must hold the file');
    try
      // The plain copy has to fail, or the test proves nothing.
      Failed := False;
      try
        TFile.Copy(Source, Target, True);
      except
        Failed := True;
      end;
      Assert.IsTrue(Failed, 'a file in use cannot simply be overwritten');

      DayBefore := FormatDateTime('yyyymmdd', Now);
      Assert.IsTrue(ReplaceFilePossiblyInUse(Source, Target, Aside),
        'but it can be replaced');
      DayAfter := FormatDateTime('yyyymmdd', Now);
      Assert.AreEqual('NEW', TFile.ReadAllText(Target), False,
        'the new bytes are in place');
      Assert.IsTrue(Aside <> '', 'and the old file was moved aside');
      Assert.IsTrue(TFile.Exists(Aside), 'under a name that is still there');
      Assert.AreEqual('OLD', TFile.ReadAllText(Aside), False,
        'carrying what was installed before');
      // The aside name must be UNIQUE: a fixed ".old" is held by whatever
      // was started from it last time, and then the move has nowhere to go.
      Assert.AreNotEqual(Target + '.old', Aside, False,
        'a fixed aside name is what blocked the next install');

      // ...and unique ACROSS PROCESSES, which is the half that was wrong:
      // Random without Randomize starts every process at the same first
      // value, so every install walked the same collisions. A timestamp
      // does not. 17 digits (yyyymmddhhnnsszzz) plus the try counter - a
      // 6-digit random number fails both of these.
      P := Aside.LastDelimiter('.');
      Stamp := Aside.Substring(0, P);           // drop ".old"
      Stamp := Stamp.Substring(Stamp.LastDelimiter('.') + 1);
      Assert.IsTrue(Stamp.Length >= 17,
        'the aside name carries a timestamp, not a 6-digit random: ' + Aside);
      Assert.IsTrue(Stamp.StartsWith(DayBefore) or Stamp.StartsWith(DayAfter),
        'and it is TODAY''s timestamp: ' + Aside);

      // The shape is a CONTRACT, not a taste. Since 1.30.1 the installer
      // writes the MCP bridge's aside copy, and this is the function that
      // reads it back - it tells "the install replaced the exe under this
      // process" (start a new session) from "the install never finished"
      // (run install.cmd). The two ends have to agree.
      Assert.AreEqual(Target, CanonicalBridgeExePath(Aside), False,
        'the consumer must recognise the name the installer writes');
    finally
      CloseHandle(H);
    end;

    // With nothing holding it, the target is simply overwritten and no
    // copy is left behind.
    TFile.WriteAllText(Source, 'NEWER');
    Assert.IsTrue(ReplaceFilePossiblyInUse(Source, Target, Aside), 'plain copy');
    Assert.AreEqual('', Aside, False, 'nothing to move aside');
    Assert.AreEqual('NEWER', TFile.ReadAllText(Target), False, 'replaced');
  finally
    TDirectory.Delete(Dir, True);
  end;
end;

procedure TInstallOutputRootTests.NothingFollowsAFailedBuild;
begin
  Assert.IsTrue(AfterBuildSteps(False) = [],
    'after a failed build no step may run - not the copy, not a ' +
    'registration, not a post event');
  Assert.IsTrue(AfterBuildSteps(True) =
    [abCopyArtefacts, abRegisterPackages, abRegisterExperts, abPostEvents],
    'after a good build (or none) every step runs');
end;

procedure TInstallOutputRootTests.APostEventNeverVerifiesAFileTheInstallDidNotWrite;
begin
  Assert.IsTrue(WithdrawnByCopy(True) = [abPostEvents],
    'a failed artefact copy withdraws the post events - they would answer ' +
    'for the file the target already had');
  Assert.IsTrue(WithdrawnByCopy(False) = [],
    'a copy that arrived withdraws nothing');
  // The registrations are deliberately NOT withdrawn, and that is the half
  // a blanket "nothing follows a failed copy" would get wrong: the BPL the
  // IDE loads is written by the BUILD (DCC_BplOutput), not by the file
  // copy, so it is there and registering it is right.
  Assert.IsTrue(AfterBuildSteps(True) - WithdrawnByCopy(True) =
    [abCopyArtefacts, abRegisterPackages, abRegisterExperts],
    'a failed copy leaves the registrations alone');
  Assert.IsTrue(AfterBuildSteps(False) - WithdrawnByCopy(True) = [],
    'and after a failed build there is nothing left to withdraw');
end;

procedure TInstallOutputRootTests.TheBuildIsToldWhereItMayCreateDirectories;
begin
  // The ordinary machine: the common directory is writable, the resolver
  // leaves the root alone, and so does the command line.
  var Cmd := MSBuildCommandLine('C:\p\X.dproj', 'Win32', 'Release',
    'C:\o\dcu', 'C:\o\bpl', 'C:\o\dcp', 'C:\o\hpp', '', '');
  Assert.IsTrue(Cmd.Contains('/tv:4.0'), 'the toolset is still named');
  Assert.IsTrue(Cmd.Contains('/p:DCC_BplOutput="C:\o\bpl"'), 'the outputs');
  Assert.IsFalse(Cmd.ToUpper.Contains('BDSCOMMONDIR'),
    'nothing is moved while Delphi''s own common directory can be written');

  // The forum machine: Delphi was installed by another account, the
  // resolver moved the output root - and the build has to be told, or the
  // targets file dies creating a Dcp and an hpp folder under the old one.
  Cmd := MSBuildCommandLine('C:\p\X.dproj', 'Win64', 'Debug',
    'C:\o\dcu', 'C:\o\bpl', 'C:\o\dcp', 'C:\o\hpp',
    'C:\Users\me\AppData\Local\Embarcadero\Studio\37.0', '/v:d');
  Assert.IsTrue(
    Cmd.Contains('/p:BDSCOMMONDIR="C:\Users\me\AppData\Local\Embarcadero\Studio\37.0"'),
    'the root the resolver chose: ' + Cmd);
  Assert.IsTrue(Cmd.Contains('/p:DCC_BplOutput="C:\o\bpl"'),
    'and the outputs still say where the artefacts really go');
  Assert.IsTrue(Cmd.EndsWith('/v:d'), 'extra parameters stay at the end');
end;

procedure TInstallOutputRootTests.TheLibSuffixComesFromTheInstallation;

  function FakeIde(const AVer, ARtl: string): string;
  begin
    Result := TPath.Combine(TPath.GetTempPath,
      'rl-libsuffix-' + AVer + '-' + ARtl);
    TDirectory.CreateDirectory(TPath.Combine(Result, 'bin'));
    if ARtl <> '' then
      TFile.WriteAllText(TPath.Combine(Result, 'bin\rtl' + ARtl + '.bpl'), 'x');
  end;

var
  Dirs: TArray<string>;
begin
  try
    // Delphi 11 and 12: the installation says 280 / 290 while the registry
    // branch says 22.0 / 23.0 - the version is exactly what must NOT decide.
    Dirs := Dirs + [FakeIde('22.0', '280')];
    Assert.AreEqual('280', ProductLibSuffix(Dirs[0], '22.0'),
      'Delphi 11 registers as 22.0 and its packages carry 280');
    Dirs := Dirs + [FakeIde('23.0', '290')];
    Assert.AreEqual('290', ProductLibSuffix(Dirs[1], '23.0'),
      'Delphi 12 registers as 23.0 and its packages carry 290');

    // No installation to read: the version without its dot, which is what
    // this answered before and what Delphi 13 needs - so nothing moves.
    Dirs := Dirs + [FakeIde('37.0', '')];
    Assert.AreEqual('370', ProductLibSuffix(Dirs[2], '37.0'),
      'the fallback is the old rule');
    Assert.AreEqual('220', ProductLibSuffix('', '22.0'),
      'and it is also what an unknown root answers');

    // A name that is not rtl<digits>.bpl decides nothing.
    Dirs := Dirs + [FakeIde('x', 'proxy')];
    Assert.AreEqual('370', ProductLibSuffix(Dirs[3], '37.0'),
      'rtlproxy.bpl is no product suffix');

    // And the REAL installation this runs on: Delphi 13 ships rtl370.bpl,
    // so the measured answer and the fallback agree here. That agreement
    // is the whole reason a D13 install cannot change behaviour.
    var Root := FindBdsRoot;
    if Root <> '' then
      Assert.AreEqual(StringReplace(ExtractFileName(ExcludeTrailingPathDelimiter(Root)),
        '.', '', [rfReplaceAll]), ProductLibSuffix(Root, '37.0'),
        'the installed IDE: its own rtl*.bpl and its version agree');
  finally
    for var D in Dirs do
      if TDirectory.Exists(D) then TDirectory.Delete(D, True);
  end;
end;

procedure TInstallOutputRootTests.ABuildNeverWritesWhereItMayNot;
const
  Common = 'C:\Users\ADMHUMMMA\Documents\Embarcadero\Studio\37.0';
  Local  = 'C:\Users\normaluser\AppData\Local';
  Public_ = 'C:\Users\Public\Documents\Embarcadero\Studio\37.0';
var
  Root, Note: string;
  Asked: TArray<string>;

  function ProbeAllowing(const AAllowed: string): TDIHWriteProbe;
  begin
    Result :=
      function(const ADir: string): Boolean
      begin
        Asked := Asked + [ADir];
        Result := (AAllowed <> '') and ADir.StartsWith(AAllowed, True);
      end;
  end;

begin
  // The ordinary machine: the common directory is writable, so nothing moves
  // and nothing is said. This is the case every existing user is in, and it
  // must stay byte for byte what it was.
  Asked := nil;
  DecideBuildOutputRoot(Public_, '37.0', Local, ProbeAllowing('C:'), Root, Note);
  Assert.AreEqual(Public_, Root, 'a writable common directory is used as it is');
  Assert.AreEqual('', Note, 'and there is nothing to report');
  Assert.AreEqual(1, Integer(Length(Asked)),
    'the fallback is not even probed when it is not needed');

  // The reported machine: Delphi installed by ADMHUMMMA, this account cannot
  // write there.
  Asked := nil;
  DecideBuildOutputRoot(Common, '37.0', Local, ProbeAllowing(Local), Root, Note);
  Assert.AreEqual(Local + '\Embarcadero\Studio\37.0', Root,
    'the build goes to a directory this account owns');
  Assert.IsTrue(Note.Contains(Common) and Note.Contains(Root),
    'and the log names BOTH paths - the next report must not have to guess: ' + Note);
  Assert.IsTrue(Note.Contains('registered'),
    'it also says the package is registered from there, which is what makes ' +
    'it work at all');

  // Neither writable: that build is going to fail, and saying why beats the
  // MSB3191 it would fail with. The root stays the IDE's value, so nothing
  // is registered from a directory that does not exist.
  Asked := nil;
  DecideBuildOutputRoot(Common, '37.0', Local, ProbeAllowing(''), Root, Note);
  Assert.AreEqual(Common, Root, 'no invented directory when nothing is writable');
  Assert.IsTrue(Note.StartsWith('WARNING'), 'and it is a warning: ' + Note);
  Assert.AreEqual(2, Integer(Length(Asked)), 'both were tried');

  // A probe we cannot run decides nothing - the caller keeps the IDE's value.
  DecideBuildOutputRoot(Common, '37.0', Local, nil, Root, Note);
  Assert.AreEqual(Common, Root, 'without a probe, nothing is substituted');
  Assert.AreEqual('', Note, 'and nothing is claimed');
end;

{ TCheckScopeTests }

procedure TCheckScopeTests.ADuplicateUnitNameIsReportedAsSkipped;
var
  R: TUsesCycleResult;
begin
  // Two files named Foo.pas in different folders: only one can be the
  // graph's node, and the other one's uses edges are simply not in the
  // answer. That used to happen without a word.
  R := TUsesGraphAnalyzer.Analyze(
    ['C:\a\Foo.pas', 'C:\b\Foo.pas', 'C:\a\Bar.pas'], nil,
    function(AFile: string): string
    begin
      if SameText(AFile, 'C:\a\Foo.pas') then
        Result := 'unit Foo;' + NL + 'interface' + NL + 'uses Bar;' + NL +
          'implementation' + NL + 'end.'
      else if SameText(AFile, 'C:\b\Foo.pas') then
        Result := 'unit Foo;' + NL + 'interface' + NL + 'implementation' + NL + 'end.'
      else
        Result := 'unit Bar;' + NL + 'interface' + NL + 'uses Foo;' + NL +
          'implementation' + NL + 'end.';
    end);
  try
    Assert.AreEqual(1, Integer(Length(R.SkippedFiles)),
      'the second Foo.pas must be reported as skipped');
    Assert.AreEqual('C:\b\Foo.pas', R.SkippedFiles[0]);
    // ... and the analysis itself still works on the first one.
    Assert.IsTrue(Length(R.Edges) > 0, 'Foo <-> Bar is a cycle');
  finally
    R.Free;
  end;

  // Nothing to report when every unit name is unique.
  R := TUsesGraphAnalyzer.Analyze(['C:\a\Foo.pas'], nil,
    function(AFile: string): string
    begin
      Result := 'unit Foo;' + NL + 'interface' + NL + 'implementation' + NL + 'end.';
    end);
  try
    Assert.AreEqual(0, Integer(Length(R.SkippedFiles)), 'no duplicates');
  finally
    R.Free;
  end;
end;

procedure TCheckScopeTests.ThePathVariablesKnowTheirPlatform;
const
  Root = 'C:\Studio\37.0';
begin
  // $(Platform) was hard-coded to Win32, so a Win64 library path
  // resolved to directories that do not exist - and with two Delphi
  // versions installed the check read the wrong one's paths.
  Assert.AreEqual('C:\lib\Win32\Release',
    ExpandIdeVars('C:\lib\$(Platform)\$(Config)', Root, nil),
    'Win32 stays the default');
  Assert.AreEqual('C:\lib\Win64\Release',
    ExpandIdeVars('C:\lib\$(Platform)\$(Config)', Root, nil, 'Win64'),
    'and Win64 is honoured when asked for');
  Assert.AreEqual(Root + '\lib\Win64',
    ExpandIdeVars('$(BDSLIB)\$(Platform)', Root, nil, 'Win64'),
    'together with the other built-ins');
  // The IDE's own version key must not carry our settings branch.
  var Key := IdeVersionRegistryKey;
  Assert.IsFalse(ContainsText(Key, 'DelphiRefactoringLight'),
    'the version key is the IDE''s, not ours: ' + Key);
  // ... and it must be a VERSION key: the Library / Browsing paths live
  // below 'BDS\37.0', not below 'BDS'. An empty answer is allowed (another
  // machine, no RAD Studio in HKCU), a wrong one is not.
  if Key <> '' then
  begin
    var Last := Copy(Key, Key.LastIndexOf('\') + 2, MaxInt);
    var Num: Double;
    Assert.IsTrue(TryStrToFloat(Last, Num, TFormatSettings.Invariant),
      'the last segment must be the IDE version: ' + Key);
  end;
end;

procedure TCheckScopeTests.TheIdeVersionKeyIsFoundByEvidence;
var
  NoProbe: TFunc<string, Boolean>;
begin
  // EVIDENCE = the key has a RootDir value; only below that one do the
  // Library / Browsing paths exist.
  var HasRootDir: TFunc<string, Boolean> :=
    function(AKey: string): Boolean
    begin
      Result := SameText(AKey, 'Software\Embarcadero\BDS\37.0');
    end;
  Assert.AreEqual('Software\Embarcadero\BDS\37.0',
    TrimToIdeVersionKey('Software\Embarcadero\BDS\37.0', HasRootDir),
    'the base key already IS the version key and must be left alone');
  Assert.AreEqual('Software\Embarcadero\BDS\37.0',
    TrimToIdeVersionKey('Software\Embarcadero\BDS\37.0\DelphiRefactoringLight',
      HasRootDir), 'a caller''s own branch below it is cut back');
  Assert.AreEqual('Software\Embarcadero\BDS\37.0',
    TrimToIdeVersionKey('Software\Embarcadero\BDS\37.0\', HasRootDir),
    'a trailing backslash is not a level');
  // The hive root opens fine and holds nothing - answering it was the bug.
  Assert.AreEqual('', TrimToIdeVersionKey('Software\Embarcadero\BDS', HasRootDir),
    'the hive root is no version key');
  Assert.AreEqual('', TrimToIdeVersionKey('', HasRootDir), 'nothing to trim');
  NoProbe := nil;
  Assert.AreEqual('', TrimToIdeVersionKey('Software\Embarcadero\BDS\37.0', NoProbe),
    'without a probe there is no evidence, so no key');
end;

procedure TCheckScopeTests.TheDfmSignatureScopeLeavesTheStudioSourcesOut;
const
  Root = 'C:\Studio\37.0\';   // RootDir as the registry holds it
var
  Dirs: TArray<string>;
begin
  Dirs := SignatureLibraryDirs([
    'C:\Studio\37.0\\source\fmx',        // what ExpandIdeVars makes of $(BDS)\source\fmx
    'C:\Studio\37.0\SOURCE\VCL',
    'C:\Studio\37.0\lib\Win64\release',
    'C:\Studio\37.0',
    'C:\DevExpress\VCL\ExpressEditors Library\Sources',
    'C:\Studio\37.0x\Sources',           // next to the root, not below it
    'C:\Users\Public\Documents\Embarcadero\Studio\37.0\Dcp'], Root);
  Assert.AreEqual(3, Integer(Length(Dirs)),
    'only the three outside the Studio tree: ' + string.Join('; ', Dirs));
  Assert.AreEqual('C:\DevExpress\VCL\ExpressEditors Library\Sources', Dirs[0],
    'a third-party directory is kept');
  Assert.AreEqual('C:\Studio\37.0x\Sources', Dirs[1], 'a sibling of the root is kept');
  Assert.AreEqual('C:\Users\Public\Documents\Embarcadero\Studio\37.0\Dcp', Dirs[2],
    'the public documents are not the Studio tree');
  // the same root without its trailing backslash
  Assert.AreEqual(3, Integer(Length(SignatureLibraryDirs(
    ['C:\Studio\37.0\source\rtl\common', 'C:\a', 'C:\b', 'C:\c'], 'C:\Studio\37.0'))),
    'RootDir without a trailing backslash');
  Assert.AreEqual(2, Integer(Length(SignatureLibraryDirs(['C:\a', 'C:\b'], ''))),
    'no root, nothing to leave out');
end;

{ TUnitRenameScopeTests }

procedure TUnitRenameScopeTests.OnlyUnitPlacesCountForAUnitRename;
var
  Lines, Masked: TArray<string>;
  Regions: TArray<TUnitRefRegion>;

  function IsUnitRef(ALine0: Integer): Boolean;
  var
    C: Integer;
  begin
    // the first occurrence of 'Settings' as a whole word on that line
    C := Pos('Settings', Lines[ALine0]);
    Assert.IsTrue(C > 0, 'the fixture line must mention Settings: ' +
      Lines[ALine0]);
    Result := UnitNameOccurrenceIsUnitReference(Masked, Regions, ALine0, C,
      Length('Settings'));
  end;

begin
  // The reported shape: a unit named Settings, and in another unit a
  // variable of the same name. A unit rename is a whole-word text replace
  // with no LSP check, so without a filter it renamed the variable too.
  Lines := [
    'unit Consumer;',                       // 0
    '',                                     // 1
    'interface',                            // 2
    '',                                     // 3
    'uses',                                 // 4
    '  System.Classes, Settings;',          // 5
    '',                                     // 6
    'implementation',                       // 7
    '',                                     // 8
    'var',                                  // 9
    '  Settings: TObject;',                 // 10
    '',                                     // 11
    'procedure Go;',                        // 12
    'begin',                                // 13
    '  Settings := TObject.Create;',        // 14
    '  Settings.Free;',                     // 15
    '  Settings.DoIt(Other.Settings);',     // 16
    '  // Settings is the unit',            // 17
    'end;',                                 // 18
    '',                                     // 19
    'end.'];                                // 20
  Masked := MaskCommentsAndStrings(Lines);
  Regions := CollectUnitRefRegions(Masked);

  Assert.IsTrue(IsUnitRef(5), 'the uses-clause entry IS the unit');
  Assert.IsFalse(IsUnitRef(10), 'a variable declaration is not the unit');
  Assert.IsFalse(IsUnitRef(14), 'an assignment to a variable is not the unit');
  // 'Settings.Free' cannot be told from a qualified unit reference by text,
  // and a qualified unit reference that is NOT renamed does not compile -
  // so this one counts, deliberately.
  Assert.IsTrue(IsUnitRef(15), 'a qualifier counts (it may be the unit)');
  // ... but the MEMBER half of 'Other.Settings' never does. The first
  // occurrence on line 16 is the qualifier, so ask for the second one.
  var P2 := Pos('Other.Settings', Lines[16]) + Length('Other.');
  Assert.IsFalse(UnitNameOccurrenceIsUnitReference(Masked, Regions, 16, P2,
    Length('Settings')), 'Other.Settings is a member, not the unit');

  // And the header of the unit itself, in the renamed unit's own file.
  Lines := ['unit Settings;', '', 'interface', '', 'implementation', '', 'end.'];
  Masked := MaskCommentsAndStrings(Lines);
  Regions := CollectUnitRefRegions(Masked);
  Assert.IsTrue(UnitNameOccurrenceIsUnitReference(Masked, Regions, 0,
    Pos('Settings', Lines[0]), Length('Settings')),
    'the unit header names the unit');

  // A .dpk: 'contains' / 'requires' open a clause too, but 'Contains' as
  // a method call must not ("if List.Contains(Settings) then").
  Lines := [
    'package Demo;',                        // 0
    'requires',                             // 1
    '  rtl, Settings;',                     // 2
    'contains',                             // 3
    '  Settings in ''Settings.pas'';',      // 4
    'end.'];                                // 5
  Masked := MaskCommentsAndStrings(Lines);
  Regions := CollectUnitRefRegions(Masked);
  Assert.IsTrue(UnitNameOccurrenceIsUnitReference(Masked, Regions, 2,
    Pos('Settings', Lines[2]), 8), 'a requires entry IS the unit');
  Assert.IsTrue(UnitNameOccurrenceIsUnitReference(Masked, Regions, 4,
    Pos('Settings', Lines[4]), 8), 'a contains entry IS the unit');

  Lines := [
    'unit U;',                              // 0
    'implementation',                       // 1
    'procedure P;',                         // 2
    'begin',                                // 3
    '  if List.Contains(Settings) then Settings := nil;',   // 4
    'end;',                                 // 5
    'end.'];                                // 6
  Masked := MaskCommentsAndStrings(Lines);
  Regions := CollectUnitRefRegions(Masked);
  Assert.IsFalse(UnitNameOccurrenceIsUnitReference(Masked, Regions, 4,
    Pos('Settings', Lines[4]), 8),
    'a method named Contains does not open a uses clause');
end;

procedure TUnitRenameScopeTests.ThePreviewChangesTheOccurrenceTheApplyChanges;
begin
  // The preview used StringReplace, which rewrites the FIRST occurrence on
  // the line - on "Foo('Foo', Foo);" it showed the wrong one changed.
  var Line := '  Foo(''Foo'', Foo);';
  var Col := Pos('Foo', Line);                       // the call itself
  Assert.AreEqual('  Bar(''Foo'', Foo);',
    PreviewReplacedLine(Line, 'Foo', 'Bar', Col), 'the call');
  var Last := Pos('Foo);', Line);                    // the argument
  Assert.IsTrue(Last > Col, 'the fixture has a later occurrence');
  Assert.AreEqual('  Foo(''Foo'', Bar);',
    PreviewReplacedLine(Line, 'Foo', 'Bar', Last), 'the argument');
  // case-insensitive match: the line's own spelling is replaced
  Assert.AreEqual('  Bar := 1;',
    PreviewReplacedLine('  foo := 1;', 'Foo', 'Bar', 3), 'lower case');
  // nothing to splice: the line is handed back unchanged rather than
  // guessing a position
  Assert.AreEqual('  Other := 1;',
    PreviewReplacedLine('  Other := 1;', 'Foo', 'Bar', 3), 'no match there');
  Assert.AreEqual('  Foo;', PreviewReplacedLine('  Foo;', 'Foo', 'Bar', 99),
    'a column past the line');
end;

procedure TUnitRenameScopeTests.AReintroducedMemberIsAnotherSymbol;
const
  Src =
    'unit U;' + NL +
    'interface' + NL +
    'type' + NL +
    '  TBase = class' + NL +
    '    procedure Init(A: Integer); virtual;' + NL +
    '  end;' + NL +
    '  TChild = class(TBase)' + NL +
    '    procedure Init(A: Integer); reintroduce;' + NL +
    '  end;' + NL +
    '  TOther = class(TBase)' + NL +
    '    procedure Init(A: Integer); override;' + NL +
    '  end;' + NL +
    '  TWide = class(TBase)' + NL +
    '    procedure Init(A: Integer; B: string); overload;' + NL +
    '  end;' + NL +
    'implementation' + NL +
    'end.';
var
  Why: string;
begin
  // The renamed declaration is TBase.Init(A: Integer) - one parameter.
  var Lines := SplitContentLines(Src);
  Assert.AreEqual(1, DeclaredParamCount(Lines, 4),
    'the renamed declaration has one parameter');

  Assert.IsTrue(MemberIsOtherSymbol(Src, 'TChild', 'Init', 1, Why),
    'a reintroduced member hides the base one');
  Assert.Contains(Why, 'reintroduce');

  Why := '';
  Assert.IsFalse(MemberIsOtherSymbol(Src, 'TOther', 'Init', 1, Why),
    'an override IS the symbol');
  Assert.AreEqual('', Why);

  Assert.IsTrue(MemberIsOtherSymbol(Src, 'TWide', 'Init', 1, Why),
    'an overload with another parameter count is another symbol');
  Assert.Contains(Why, '2 parameter');

  // Without a known parameter count only reintroduce counts - guessing
  // would drop a real occurrence, which does not compile.
  Assert.IsFalse(MemberIsOtherSymbol(Src, 'TWide', 'Init', -1, Why),
    'an unknown signature decides nothing');
  Assert.IsFalse(MemberIsOtherSymbol(Src, 'TNoSuchType', 'Init', 1, Why),
    'an unknown type decides nothing');
end;

{ TUnitFromPackageTests }

procedure TUnitFromPackageTests.OnlyWhereThePackageIsMeasuredIsItNamed;
const
  Root = 'C:\Program Files (x86)\Embarcadero\Studio\37.0';
begin
  // the reported case
  Assert.AreEqual('designide', PackageShippingUnit(
    Root + '\source\ToolsAPI\DesignIntf.pas', Root),
    'DesignIntf is shipped inside designide.dcp - measured with dcc32');
  // the folder, however it is spelled, and whatever the unit
  Assert.AreEqual('designide', PackageShippingUnit(
    Root + '\SOURCE\TOOLSAPI\DesignEditors.pas', Root),
    'the same folder in another spelling');
  Assert.AreEqual('designide', PackageShippingUnit(
    Root + '\source\ToolsAPI\ToolsAPI.pas', Root),
    'and every other unit of it - 42 of 42 resolve with -LUdesignide');

  // ...and NOT where the measurement says otherwise: those units are in no
  // .dcp, so adding their folder to the search path is the only way and
  // must keep being offered
  Assert.AreEqual('', PackageShippingUnit(
    Root + '\SOURCE\PROPERTY EDITORS\ActnRes.pas', Root),
    'the property editors are shipped as SOURCE, not inside a package');
  // a unit of the user's own code is never one of these
  Assert.AreEqual('', PackageShippingUnit('C:\Work\MyApp\Unit1.pas', Root),
    'a project unit is no installation unit');
  // and without a root there is no claim to make
  Assert.AreEqual('', PackageShippingUnit(
    Root + '\source\ToolsAPI\DesignIntf.pas', ''),
    'no installation root, no answer');
end;

procedure TUnitFromPackageTests.TheRequiresClauseIsRead;
const
  // his own Package1.dpk, verbatim
  Dpk =
    'package Package1;'#13#10 +
    #13#10 +
    '{$R *.res}'#13#10 +
    '{$IMPLICITBUILD ON}'#13#10 +
    #13#10 +
    'requires'#13#10 +
    '  designide,'#13#10 +
    '  rtl;'#13#10 +
    #13#10 +
    'contains'#13#10 +
    '  DockFrame in ''DockFrame.pas'' {MyDockFrame: TFrame};'#13#10 +
    #13#10 +
    'end.'#13#10;
  // a commented-out entry is not in the clause
  Commented =
    'package P;'#13#10 +
    'requires rtl, { designide, } vcl;'#13#10 +
    'end.'#13#10;
  // 'requires' is reserved in a package only, so neither a call of that
  // name nor a name that BEGINS with it may open a clause. Only the SECOND
  // of the two is a rule a fixture can isolate: after a method call's
  // "Requires" comes a '(' or an operator, and nothing is collected
  // whatever the rule does, while "RequiresDesignide := True" would hand
  // out "Designide" as a required package. The first-word half stays
  // because ClauseOpener states the same rule for every clause - it is
  // defence, and its mutation cannot go red (see neg.py).
  NotAClause =
    'unit U;'#13#10 +
    'interface'#13#10 +
    'implementation'#13#10 +
    'procedure Go;'#13#10 +
    'begin'#13#10 +
    '  if List.Requires(designide) then Exit;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
  // ...and a NAME that begins with it. Its own fixture, because the clause
  // scan stops at the first ';' it reaches: in one text the line above
  // would end the scan before this one is read, and the test would pass
  // for that reason instead of for the rule.
  NameStartsWithRequires =
    'unit U;'#13#10 +
    'implementation'#13#10 +
    'procedure Go;'#13#10 +
    'begin'#13#10 +
    '  RequiresDesignide := True;'#13#10 +
    'end;'#13#10 +
    'end.'#13#10;
var
  Names: TArray<string>;
begin
  Names := PackageRequiresNames(Dpk);
  Assert.AreEqual(2, Integer(Length(Names)), 'two packages are required');
  Assert.AreEqual('designide', Names[0], 'and this is the one that matters');
  Assert.AreEqual('rtl', Names[1], 'the clause ends at its own semicolon - '
    + 'nothing of "contains" is in it');

  Names := PackageRequiresNames(Commented);
  Assert.AreEqual(2, Integer(Length(Names)), 'rtl and vcl');
  for var N in Names do
    Assert.AreNotEqual('designide', N,
      'a commented-out entry does not make it required');

  Assert.AreEqual(0, Integer(Length(PackageRequiresNames(NotAClause))),
    'List.Requires( is a call, not a clause');
  Assert.AreEqual(0, Integer(Length(PackageRequiresNames(NameStartsWithRequires))),
    'and RequiresDesignide is a name, not the clause plus a package');
  Assert.AreEqual(0, Integer(Length(PackageRequiresNames(''))),
    'and an empty text has no clause either');
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
  TDUnitX.RegisterTestFixture(TMethodEditTests);
  TDUnitX.RegisterTestFixture(TLspConfigHintTests);
  TDUnitX.RegisterTestFixture(TRemoveWithSafetyTests);
  TDUnitX.RegisterTestFixture(TUnitRenameScopeTests);
  TDUnitX.RegisterTestFixture(TDfmAliasOnlyTests);
  TDUnitX.RegisterTestFixture(TLineEndingCheckTests);
  TDUnitX.RegisterTestFixture(TInstallOutputRootTests);
  TDUnitX.RegisterTestFixture(TDeclareConstantTests);
  TDUnitX.RegisterTestFixture(TFileEncodingCheckTests);
  TDUnitX.RegisterTestFixture(TUnitFromPackageTests);
  TDUnitX.RegisterTestFixture(TCheckScopeTests);
  TDUnitX.RegisterTestFixture(TExtractAndCompletionTests);

end.
