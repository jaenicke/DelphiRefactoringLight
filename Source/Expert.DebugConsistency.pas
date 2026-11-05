(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DebugConsistency;

// "The debugger does not stop at my breakpoint" / "the debug info does not
// match the code that runs" (colleagues of the user, 2026-09). Most causes
// leave traces on disk or in the project options that can be found without
// a debug session:
//
//  * a unit source exists TWICE - the editor shows one copy, the compiler
//    finds the other one first on the search path;
//  * a stray precompiled .dcu of a project unit lies in another directory of
//    the search path, or a unit is only available as .dcu (no source the
//    compiler can see) while its source is open in the editor;
//  * line endings that are not CRLF (LF only / mixed, typically after a git
//    checkout) - the classic reason for breakpoints on the wrong line;
//  * debug information / local symbols switched off or optimization on,
//    in the active configuration or by a directive in the source;
//  * the executable is older than the sources (the last build failed), its
//    symbol files (.rsm/.tds/.map) do not belong to it, or the debugger
//    starts a different host application.
//
// This unit is PURE (files + an input record, no ToolsAPI) so the console
// suite can drive it; the IDE side gathers the input.

interface

uses
  System.SysUtils, System.Generics.Collections, Delphi.FileEncoding;

type
  TDebugIssueSeverity = (dsInfo, dsWarning, dsProblem);
  TDebugIssueKind = (dkDuplicateSource, dkStrayDcu, dkDcuOnly, dkLineEndings,
    dkDebugOption, dkSourceDirective, dkExecutable, dkSymbolFile,
    dkHostApplication, dkDuplicateBinary, dkLspDcu);

  TDebugIssue = record
    Kind: TDebugIssueKind;
    Severity: TDebugIssueSeverity;
    FileName: string;   // the file the issue is about
    Line: Integer;      // 1-based, 0 = not line specific
    Reason: string;
    Hint: string;
    /// <summary>dkLineEndings only: the file can be converted to CRLF.</summary>
    Fixable: Boolean;
  end;

  TDebugCheckInput = record
    ProjectFile: string;               // .dproj
    ProjectDir: string;
    /// <summary>Project sources: .dpr/.dpk and all units of the project.</summary>
    ProjectSources: TArray<string>;
    /// <summary>Source files open in the editor (any, project or not).</summary>
    OpenFiles: TArray<string>;
    /// <summary>Directories the COMPILER searches for units, in its order:
    ///  project directory, project search path, IDE library path.</summary>
    UnitSearchDirs: TArray<string>;
    /// <summary>Directories below which duplicates are NOT reported (the
    ///  shipped RTL/VCL - it has platform copies by design).</summary>
    IgnoreDirs: TArray<string>;
    DcuOutputDir: string;              // '' = next to each source
    /// <summary>The project''s .delphilsp.json, '' when there is
    ///  none. Its DCU output is what DelphiLSP reads INSTEAD of the
    ///  sources when .dcu files lie there.</summary>
    LspConfigFile: string;
    TargetFile: string;                // exe / dll / bpl the build produces
    HostApplication: string;           // run parameters, '' = the target
    ConfigName: string;
    /// <summary>Option values of the active configuration as the project
    ///  stores them ('' = unknown / not available).</summary>
    DebugInformation: string;
    LocalDebugSymbols: string;
    Optimize: string;
    /// <summary>Directories Windows searches for DLLs/packages (PATH etc.).</summary>
    BinarySearchDirs: TArray<string>;
  end;

  TDebugProgress = reference to function(ACurrent, ATotal: Integer;
    const AText: string): Boolean;   // False = cancel

/// <summary>Runs all checks. AProgress may be nil.</summary>
function RunDebugConsistencyCheck(const AInput: TDebugCheckInput;
  const AProgress: TDebugProgress): TArray<TDebugIssue>;

/// <summary>The DCU output directory a .delphilsp.json asks the compiler
///  for: the -NU option of its "dccOptions", unquoted and relative to the
///  project as the compiler reads it. '' when there is none.</summary>
function DcuOutputFromDccOptions(const AOptions: string): string;

/// <summary>The DCU output directory of the LSP configuration, resolved
///  against AProjectDir. '' when the file cannot be read or names
///  none.</summary>
function LspDcuDirOfConfig(const AConfigFile, AProjectDir: string): string;

/// <summary>How many units the LSP configuration lists for this project
///  (its projectFiles) - the number a DCU output directory that matches the
///  project has one .dcu for. 0 when the file cannot be read.</summary>
function LspProjectUnitCount(const AConfigFile: string): Integer;

/// <summary>The directories the LSP configuration tells the compiler to
///  SEARCH for units (-U of its dccOptions), resolved against the project.
///  A unit whose source lies in none of them can only be reached through
///  the project file list or a .dcu.</summary>
function LspUnitSearchDirs(const AConfigFile, AProjectDir: string): TArray<string>;

type
  /// <summary>What the DCU output directory of the LSP configuration holds,
  ///  measured against the project's sources.</summary>
  TLspDcuDirState = record
    /// <summary>Project units that compile to a .dcu at all.</summary>
    UnitCount: Integer;
    /// <summary>...and have one in this directory.</summary>
    Present: Integer;
    /// <summary>Unit names that do not - empty when the directory matches
    ///  the project.</summary>
    Missing: TArray<string>;
    /// <summary>.dcu files here that are OLDER than their own source, with
    ///  the source path in the same order.</summary>
    Stale: TArray<string>;
    StaleSource: TArray<string>;
    /// <summary>...and for each of those, whether the compiler could use
    ///  the source instead (its directory is one the LSP searches). Only
    ///  then does moving the .dcu aside help; otherwise the .dcu is the
    ///  only way to that unit and removing it costs the answer.</summary>
    StaleReplaceable: TArray<Boolean>;
  end;

/// <summary>What that directory holds, measured against ASources.
///  MEASURED 2026-10-20, eight states on a live session with one variable
///  changed per step (probe: go to definition on a symbol of another unit
///  of the same project):
///  * a unit the LSP loads from a .dcu DOES come back with its source
///    position - with debug information and without it, from the output
///    directory and from a search path directory, as a .dcu and packed in
///    the installed package's .dcp, and even when that one file is four
///    days old. Six of the eight states answered correctly.
///  * what failed was an INCOMPLETE directory: 107 units from one build
///    plus the queried unit's .dcu missing, so it had to come from another
///    build - then the answer was empty and the server logged "Kibitz
///    definition found in file:///(-1:0)"; with every provider gone it
///    could not compile at all ("kkError"). Putting that one file back made
///    it answer again.
///  So the finding is NOT "there are .dcu files here" (they are what makes
///  it work) but "this directory does not match your sources": units
///  missing from it, or a .dcu older than its own source - and the cure is
///  a build, not deleting things. Moving a .dcu aside only helps where the
///  compiler can reach the source instead (a .dpr project whose units are
///  on its search path - the documented case of another session, whose log
///  showed 60 units loaded from .dcu and removing them fixed it).</summary>
function InspectLspDcuDir(const ADir: string;
  const ASources, AReachableDirs: TArray<string>): TLspDcuDirState;

/// <summary>Moves one .dcu into a subdirectory of its own directory, so
///  the compiler no longer finds it and the file is still there. The
///  subdirectory is created when needed; the name is kept, so a second run
///  into the same folder would overwrite - it is one folder per run.
///  </summary>
function MoveDcuAside(const AFile, AAsideDir: string;
  out ATarget, AError: string): Boolean;

/// <summary>The folder name one "move aside" run uses, stamped so two runs
///  never mix.</summary>
function DcuAsideFolder(const ADcuDir: string): string;

/// <summary>What file a compiler message means, as a path the
///  diagnostics store can be asked with. A message carries either the full
///  path - then that IS the answer - or only the base name, and then only
///  an UNAMBIGUOUS match among the project's own files is one: two
///  Utils.pas in two folders answer '' rather than a guess, because the
///  store is keyed by path and a wrong key puts a warning on the wrong
///  file. ACANDIDATES are the project's source files (plus whatever is
///  open). No disk access: a path the compiler named is taken as given,
///  and a file that cannot be read drops out when its content is read.</summary>
function ResolveCompilerMessageFile(const AName: string;
  const ACandidates: TArray<string>): string;

/// <summary>An id for one finding: its kind and its file, so an apply can
///  name a finding instead of an index (the rule every check here
///  follows).</summary>
function DebugIssueId(const AIssue: TDebugIssue): string;

/// <summary>Counts line breaks in raw file bytes (UTF-16 BOMs handled).</summary>
procedure CountLineEndings(const ABytes: TBytes; out ACrLf, ALfOnly, ACrOnly: Integer);

/// <summary>Rewrites every line break as CRLF, keeping the encoding (works
///  on the bytes; UTF-16 by code unit). False when nothing had to change
///  or the file could not be written.</summary>
function ConvertFileToCrLf(const AFile: string): Boolean;

type
  /// <summary>What a file's line breaks look like as a whole.</summary>
  TLineEndingKind = (
    lekNone,     // no line break at all (a one-line file)
    lekCrLf,     // all CRLF - what Delphi expects
    lekLf,       // all LF (a git checkout without autocrlf)
    lekCr,       // all CR (classic Mac, or a half-converted file)
    lekMixed);   // more than one of them - the case the IDE handles worst

  TLineEndingInfo = record
    CrLf: Integer;
    LfOnly: Integer;
    CrOnly: Integer;
    Kind: TLineEndingKind;
    /// <summary>True for everything but lekNone and lekCrLf.</summary>
    NeedsFix: Boolean;
  end;

/// <summary>Classifies the line breaks of raw file bytes. Same counting as
///  CountLineEndings (UTF-16 BOMs handled), plus the verdict the warning
///  and the fix decision are made from.</summary>
function ClassifyLineEndings(const ABytes: TBytes): TLineEndingInfo;

/// <summary>One sentence naming what is wrong, with the counts - "mixed
///  line endings (12 CRLF, 3 LF)". '' when nothing needs fixing.</summary>
function LineEndingProblemText(const AInfo: TLineEndingInfo): string;

/// <summary>A file whose line endings we may check AND rewrite: Delphi
///  source text only. Deliberately NOT .dfm/.fmx - a BINARY form file is
///  bytes, and turning an $0A inside it into $0D$0A would destroy it;
///  and not .dproj/.groupproj, which the IDE rewrites itself.</summary>
function LineEndingCheckCandidate(const AFileName: string): Boolean;

/// <summary>THE THIRD GATE of the two project checks, and the one a live
///  run found missing (1.31.3): a file under the RAD Studio installation
///  is not ours to rewrite. The scope collects every unit OPEN in the
///  editor, and Ctrl+click into an Indy or RTL source puts it there - the
///  encoding check then offered to convert a shipped Indy unit. The
///  line-ending WARNING has had this rule since 1.24.0; this is the same
///  rule, in one place. AInstallRoot = '' switches it off.</summary>
function UnderInstallRoot(const AFileName, AInstallRoot: string): Boolean;

/// <summary>Content that must never be rewritten byte-wise: a binary form
///  file ($FF 'TPF0'), or a NUL byte in a file that is not UTF-16. The
///  extension gate above is the first guard; this is the second, because a
///  wrong answer here corrupts a file.</summary>
function LooksBinaryContent(const ABytes: TBytes): Boolean;

type
  /// <summary>What the user answered to the line-ending warning. CANCEL is
  ///  its own answer and NOT the same as "Leave it": it decides nothing, so
  ///  the next open asks again (user, 2026-10-15 - "wenn man das nicht
  ///  will, kann man das ja auswaehlen").</summary>
  TLineEndingChoice = (
    lecCorrectOnce,     // correct these files, ask again next time
    lecCorrectSession,  // ... and stop asking until the IDE closes
    lecCorrectAlways,   // ... and save that in the options
    lecLeave,           // change nothing, and do not report these again
    lecCancel);         // change nothing, and decide nothing

  TLineEndingEffect = record
    Correct: Boolean;            // rewrite the files now
    AlwaysThisSession: Boolean;  // no more questions in this IDE session
    AlwaysSaved: Boolean;        // ... and in the next one
    DoNotAskAgain: Boolean;      // remember THESE files for this session
  end;

/// <summary>Everything one answer means, in one place: what the dialog
///  does, what it stores, and whether the files are remembered.</summary>
function LineEndingChoiceEffect(AChoice: TLineEndingChoice): TLineEndingEffect;

/// <summary>The dialog's result as an answer. AConfirmed = False (Cancel,
///  Escape, the close box) is lecCancel whatever the radio buttons say, and
///  so is an index the dialog cannot have produced.</summary>
function LineEndingChoiceOfIndex(AIndex: Integer;
  AConfirmed: Boolean): TLineEndingChoice;

type
  /// <summary>One finding of the project-wide line-ending check.</summary>
  TLineEndingIssue = record
    FileName: string;
    Info: TLineEndingInfo;
    /// <summary>Stable over a re-scan, because it names the PLACE and not
    ///  a position in the list - the same reason the DFM check and the
    ///  GUID check carry one. An apply names ids, never indexes.</summary>
    Id: string;
  end;

  /// <summary>Reads a file's raw bytes. Returns False when it cannot be
  ///  read at all (then the file is simply not a finding).</summary>
  TLineEndingReader = reference to function(const AFile: string;
    out ABytes: TBytes): Boolean;

  TLineEndingProgress = reference to procedure(ACurrent, ATotal: Integer;
    const AFile: string);

/// <summary>FNV-1a over the file's path, lower case - the same shape the
///  other checks' ids have.</summary>
function LineEndingIssueId(const AFileName: string): string;

/// <summary>Every file of AFiles that Delphi would stumble over: not pure
///  CRLF, and allowed to be rewritten at all (LineEndingCheckCandidate +
///  LooksBinaryContent). Duplicates in AFiles are reported once. AProgress
///  may be nil.</summary>
function CollectLineEndingIssues(const AFiles: TArray<string>;
  const AReader: TLineEndingReader;
  const AProgress: TLineEndingProgress = nil;
  const AInstallRoot: string = ''): TArray<TLineEndingIssue>;

type
  /// <summary>One file an encoding conversion would touch.</summary>
  TEncodingIssue = record
    FileName: string;
    Found: TSourceEncoding;
    /// <summary>Stable over a re-scan - it names the PLACE, like the
    ///  line-ending, DFM and GUID ids.</summary>
    Id: string;
  end;

  TEncodingReader = reference to function(const AFile: string;
    out ABytes: TBytes): Boolean;

/// <summary>FNV-1a over the lower-cased path.</summary>
function EncodingIssueId(const AFileName: string): string;

/// <summary>Every file of AFiles whose encoding is in AFROM and which may
///  be rewritten at all - the SAME two gates the line-ending check uses
///  (LineEndingCheckCandidate, LooksBinaryContent), so a form file or
///  anything binary is never a finding. A file that already IS ATo is
///  skipped even when its class is in AFrom, and a duplicate in AFiles is
///  reported once. AProgress may be nil.
///  WHY FORM FILES STAY OUT, measured 2026-10-16 over 1381 of them: 1368
///  are pure ASCII, i.e. a conversion would change nothing, and the rest
///  the IDE re-streams itself on the next save - so excluding them costs
///  nothing and removes the one case that could corrupt a binary
///  form.</summary>
function CollectEncodingIssues(const AFiles: TArray<string>;
  AFrom: TSourceEncodings; ATo: TSourceEncoding;
  const AReader: TEncodingReader;
  const AProgress: TLineEndingProgress = nil;
  const AInstallRoot: string = ''): TArray<TEncodingIssue>;

/// <summary>The rows an answer must carry, by id: every finding that was
///  ACTED on, whatever AMax says, plus as many of the rest as fit. Shared
///  by both project checks - a budget that drops the row the call asked
///  about is the defect 1.25.1 found live.</summary>
function SelectReportRows(const AIds, AActedIds: TArray<string>;
  AMax: Integer): TArray<Integer>;

/// <summary>Ids the caller named that no finding carries. Shared.</summary>
function UnknownReportIds(const AIds, ANamed: TArray<string>): TArray<string>;

/// <summary>What the answer of the debug-consistency check may say about
///  fixing, derived from the findings it really has. AFixable is how many
///  of them carry a cure this tool can apply. MEASURED LIVE (2.5.1): with
///  a complete DCU output the check finds nothing of that kind and the
///  answer still ended with "a finding can be fixed from here" - and since
///  2.5.0 only a STALE .dcu whose source the compiler can reach instead is
///  movable at all, so promising it for every finding of that kind is the
///  other half of the same wrong sentence. '' = say nothing about a
///  fix.</summary>
function DebugFixOfferNote(AFixable: Integer): string;

/// <summary>What an apply by ids really DID, as one sentence - derived,
///  never asserted. Measured live (2.0.4): interface_guids answered "the
///  declarations you named carry a fresh GUID now" after an apply in which
///  nothing matched at all, which is the lie 1.25.1 found in the
///  line-ending tool. ADone / AFailed are what the per-entry outcomes say,
///  AUnknown how many ids matched no finding. AThing is what is counted
///  ("declaration(s)"), ADoneText what happened to them ("carry a fresh
///  GUID now"), AWhere where the change lives.</summary>
function AppliedByIdsNote(ADone, AFailed, AUnknown: Integer;
  const AThing, ADoneText, AWhere: string): string;

/// <summary>Which findings an answer must carry, as indexes into AIssues:
///  every one that was ACTED on, whatever AMax says, plus as many of the
///  rest as fit. A row that was acted on IS the answer to the call - a
///  budget that cuts it away leaves the caller unable to see what
///  happened, which is the one thing it asked about. AMax &lt;= 0 means
///  "only the acted ones".</summary>
function SelectLineEndingRows(const AIssues: TArray<TLineEndingIssue>;
  const AActedIds: TArray<string>; AMax: Integer): TArray<Integer>;

/// <summary>Ids the caller named that no finding carries: the file is
///  already CRLF, is not a file we may rewrite, or the ids come from an
///  older scan. NOTHING was done for them, and staying silent about that
///  reads as success.</summary>
function UnknownLineEndingIds(const AIssues: TArray<TLineEndingIssue>;
  const AIds: TArray<string>): TArray<string>;

type
  TSourceDirective = record
    Line: Integer;       // 1-based
    Text: string;        // e.g. '{$D-}'
    What: string;        // 'debug information off' ...
  end;

/// <summary>Compiler directives in AContent that remove or degrade debug
///  information: {$D-}/{$DEBUGINFO OFF}, {$L-}/{$LOCALSYMBOLS OFF},
///  {$O+}/{$OPTIMIZATION ON}, also combined forms like {$D-,L-}. Comments
///  and strings are skipped.</summary>
function FindDebugDirectives(const AContent: string): TArray<TSourceDirective>;

/// <summary>Include files referenced by {$I name} / {$INCLUDE name} in
///  AContent (names only, as written).</summary>
function FindIncludeNames(const AContent: string): TArray<string>;

/// <summary>'true'/'1'/'2' = on, 'false'/'0' = off, '' = unknown.</summary>
function OptionState(const AValue: string): Integer;   // 1 on, 0 off, -1 unknown

implementation

uses
  System.Classes, System.IOUtils, System.StrUtils, System.DateUtils,
  System.Math;

function OptionState(const AValue: string): Integer;
var
  V: string;
begin
  V := LowerCase(Trim(AValue));
  if V = '' then Exit(-1);
  if (V = 'false') or (V = '0') then Exit(0);
  if (V = 'true') or (V = '1') or (V = '2') then Exit(1);
  Result := -1;
end;

// ---------------------------------------------------------------------------
//  line endings
// ---------------------------------------------------------------------------

function DcuOutputFromDccOptions(const AOptions: string): string;
var
  I: Integer;
begin
  Result := '';
  I := 1;
  // The options are ONE command line: -NU<dir>, the directory possibly
  // quoted, possibly relative. Everything else is skipped by hand rather
  // than split, because a quoted path may contain spaces.
  while I < Length(AOptions) - 2 do
  begin
    if (AOptions[I] = '-') and (AOptions[I + 1] = 'N') and
       (AOptions[I + 2] = 'U') then
    begin
      var J := I + 3;
      while (J <= Length(AOptions)) and (AOptions[J] = ' ') do Inc(J);
      if (J <= Length(AOptions)) and (AOptions[J] = '"') then
      begin
        Inc(J);
        var K := J;
        while (K <= Length(AOptions)) and (AOptions[K] <> '"') do Inc(K);
        Exit(Copy(AOptions, J, K - J));
      end;
      var K := J;
      while (K <= Length(AOptions)) and (AOptions[K] > ' ') do Inc(K);
      Exit(Copy(AOptions, J, K - J));
    end;
    Inc(I);
  end;
end;

// The dccOptions value is a JSON string: every backslash and every quote
// inside it is escaped. Replacing only the backslashes leaves a backslash
// in front of each quote, and a quoted path then parses as an entry of its
// own (measured 2026-10-20: "c:\rtl" came out as C:\c:\rtl).
function JsonUnescape(const AText: string): string;
var
  I: Integer;
begin
  Result := '';
  I := 1;
  while I <= Length(AText) do
  begin
    if (AText[I] = '\') and (I < Length(AText)) then
    begin
      case AText[I + 1] of
        '\', '"', '/': begin Result := Result + AText[I + 1]; Inc(I, 2); Continue; end;
        'n': begin Result := Result + #10; Inc(I, 2); Continue; end;
        'r': begin Result := Result + #13; Inc(I, 2); Continue; end;
        't': begin Result := Result + #9; Inc(I, 2); Continue; end;
      end;
    end;
    Result := Result + AText[I];
    Inc(I);
  end;
end;

function LspDcuDirOfConfig(const AConfigFile, AProjectDir: string): string;
var
  Text, Opts: string;
begin
  Result := '';
  if (AConfigFile = '') or not TFile.Exists(AConfigFile) then Exit;
  try
    Text := TDelphiFileEncoding.ReadAll(AConfigFile);
  except
    Exit;
  end;
  // Deliberately NOT a JSON parse: the file is written by the IDE, the one
  // value we need is a command line inside one string, and a parser would
  // bring its own failure modes for a diagnostic that must never throw.
  var P := Pos('"dccOptions"', Text);
  if P = 0 then Exit;
  P := PosEx('"', Text, P + Length('"dccOptions"') + 1);
  if P = 0 then Exit;
  var Q := PosEx('"', Text, P + 1);
  while (Q > 1) and (Text[Q - 1] = '\') do Q := PosEx('"', Text, Q + 1);
  if Q = 0 then Exit;
  Opts := JsonUnescape(Copy(Text, P + 1, Q - P - 1));
  Result := DcuOutputFromDccOptions(Opts);
  if Result = '' then Exit;
  if TPath.IsRelativePath(Result) and (AProjectDir <> '') then
    Result := TPath.Combine(AProjectDir, Result);
  Result := ExcludeTrailingPathDelimiter(ExpandFileName(Result));
end;

function LspProjectUnitCount(const AConfigFile: string): Integer;
var
  Text: string;
begin
  Result := 0;
  if (AConfigFile = '') or not TFile.Exists(AConfigFile) then Exit;
  try
    Text := TDelphiFileEncoding.ReadAll(AConfigFile);
  except
    Exit;
  end;
  var P := Pos('"projectFiles"', Text);
  if P = 0 then Exit;
  // One entry per unit: {"name": "...", "file": "file:///..."}. Counting
  // the name keys is enough and needs no JSON parser in a 1 s tick.
  var Q := PosEx('"name"', Text, P);
  while Q > 0 do
  begin
    Inc(Result);
    Q := PosEx('"name"', Text, Q + 1);
  end;
end;

function LspUnitSearchDirs(const AConfigFile, AProjectDir: string): TArray<string>;
var
  Text, Opts, Raw: string;
begin
  Result := nil;
  if AProjectDir <> '' then
    Result := [ExcludeTrailingPathDelimiter(ExpandFileName(AProjectDir))];
  if (AConfigFile = '') or not TFile.Exists(AConfigFile) then Exit;
  try
    Text := TDelphiFileEncoding.ReadAll(AConfigFile);
  except
    Exit;
  end;
  var P := Pos('"dccOptions"', Text);
  if P = 0 then Exit;
  P := PosEx('"', Text, P + Length('"dccOptions"') + 1);
  if P = 0 then Exit;
  var Q := PosEx('"', Text, P + 1);
  while (Q > 1) and (Text[Q - 1] = '\') do Q := PosEx('"', Text, Q + 1);
  if Q = 0 then Exit;
  Opts := JsonUnescape(Copy(Text, P + 1, Q - P - 1));

  // -U is ONE option whose value continues past a closing quote:
  //   -U"c:\rtl";..\Source;"c:\more"
  // Reading only up to that quote is the mistake that cost an hour on
  // 2026-10-20 ("the source directory is on no search path" - it was).
  var I := Pos('-U', Opts);
  if I = 0 then Exit;
  Inc(I, 2);
  Raw := '';
  while I <= Length(Opts) do
  begin
    if Opts[I] = '"' then
    begin
      Inc(I);
      while (I <= Length(Opts)) and (Opts[I] <> '"') do
      begin
        Raw := Raw + Opts[I];
        Inc(I);
      end;
      Inc(I);   // the closing quote
      Continue;
    end;
    if Opts[I] = ' ' then Break;   // the option ends at an unquoted blank
    Raw := Raw + Opts[I];
    Inc(I);
  end;

  for var Entry in Raw.Split([';']) do
  begin
    var D := Trim(Entry);
    if D = '' then Continue;
    if TPath.IsRelativePath(D) and (AProjectDir <> '') then
      D := TPath.Combine(AProjectDir, D);
    try
      D := ExcludeTrailingPathDelimiter(ExpandFileName(D));
    except
      Continue;
    end;
    var Known := False;
    for var E in Result do
      if SameText(E, D) then Known := True;
    if not Known then Result := Result + [D];
  end;
end;

function InspectLspDcuDir(const ADir: string;
  const ASources, AReachableDirs: TArray<string>): TLspDcuDirState;
var
  Want: TDictionary<string, string>;   // unit name -> its source path
begin
  Result := Default(TLspDcuDirState);
  if ADir = '' then Exit;
  Want := TDictionary<string, string>.Create;
  try
    for var S in ASources do
    begin
      // ONLY a unit. MEASURED 2026-10-20 over three projects of this
      // repository (a package and two programs, four output
      // directories, 110 / 132 .dcu each): the project FILE itself -
      // the .dpk or .dpr - never produces a .dcu of its own name; it
      // produces the .bpl or the .exe. Counting it made the directory
      // look incomplete by exactly one, for a file that can never
      // satisfy it, and the status row (which counts the
      // configuration's own projectFiles) said the opposite.
      if SameText(ExtractFileExt(S), '.pas') then
        Want.AddOrSetValue(
          LowerCase(ChangeFileExt(ExtractFileName(S), '')), S);
    end;
    Result.UnitCount := Want.Count;
    if Want.Count = 0 then Exit;
    for var Pair in Want do
    begin
      // From the SOURCE's own spelling: the dictionary key is lower
      // case for the lookup, and a row naming "unit1.dcu" for
      // Unit1.pas reads like a file nobody has.
      var F := TPath.Combine(ADir,
        ChangeFileExt(ExtractFileName(Pair.Value), '.dcu'));
      if not TFile.Exists(F) then
      begin
        // The unit's OWN spelling, for the same reason the .dcu path
        // above takes it: "First missing: delphirefactoringlight" is
        // what the live answer read (2.5.0), and nobody has a unit of
        // that name.
        Result.Missing := Result.Missing +
          [ChangeFileExt(ExtractFileName(Pair.Value), '')];
        Continue;
      end;
      Inc(Result.Present);
      // The .dcu describes the source it was compiled from. Older than
      // that source means its positions belong to code that has changed.
      var Older: Boolean;
      try
        Older := TFile.Exists(Pair.Value) and
          (TFile.GetLastWriteTime(F) < TFile.GetLastWriteTime(Pair.Value));
      except
        Older := False;
      end;
      if not Older then Continue;
      // Can the compiler use the source instead? Only then is removing
      // this file a cure rather than the loss of the only way to the unit.
      var SrcDir := '';
      try
        SrcDir := ExcludeTrailingPathDelimiter(
          ExpandFileName(ExtractFileDir(Pair.Value)));
      except
        SrcDir := '';
      end;
      var Reachable := False;
      for var D in AReachableDirs do
        if (D <> '') and SameText(D, SrcDir) then Reachable := True;
      Result.Stale := Result.Stale + [F];
      Result.StaleSource := Result.StaleSource + [Pair.Value];
      Result.StaleReplaceable := Result.StaleReplaceable + [Reachable];
    end;
  finally
    Want.Free;
  end;
end;

function DcuAsideFolder(const ADcuDir: string): string;
begin
  if ADcuDir = '' then Exit('');
  Result := TPath.Combine(ADcuDir,
    'aside-' + FormatDateTime('yyyymmdd-hhnnss', Now));
end;

function MoveDcuAside(const AFile, AAsideDir: string;
  out ATarget, AError: string): Boolean;
begin
  ATarget := '';
  AError := '';
  Result := False;
  if not TFile.Exists(AFile) then
  begin
    AError := 'the file is not there any more';
    Exit;
  end;
  if AAsideDir = '' then
  begin
    AError := 'no target directory';
    Exit;
  end;
  ATarget := TPath.Combine(AAsideDir, ExtractFileName(AFile));
  try
    if not TDirectory.Exists(AAsideDir) then TDirectory.CreateDirectory(AAsideDir);
    // MOVED, never deleted: a .dcu is cheap to rebuild but the user said
    // nothing about deleting anything, and the installer moves the bridge
    // aside for the same reason.
    TFile.Move(AFile, ATarget);
    Result := True;
  except
    on E: Exception do
    begin
      AError := E.Message;
      ATarget := '';
    end;
  end;
end;

function ResolveCompilerMessageFile(const AName: string;
  const ACandidates: TArray<string>): string;
var
  Base, F, Full: string;
begin
  Result := '';
  if AName = '' then Exit;
  if ExtractFilePath(AName) <> '' then Exit(ExpandFileName(AName));
  Base := AName;
  for F in ACandidates do
    if SameText(ExtractFileName(F), Base) then
    begin
      // DISTINCT paths, not occurrences: the candidates are the project's
      // sources PLUS the open files, so the active unit is legitimately
      // in there twice - counting occurrences would have skipped exactly
      // the file that matters.
      Full := ExpandFileName(F);
      if (Result <> '') and not SameText(Result, Full) then
        Exit('');   // two Utils.pas: no answer is the answer
      Result := Full;
    end;
end;

function DebugIssueId(const AIssue: TDebugIssue): string;
var
  H: Cardinal;
begin
  // FNV-1a over kind + file: the PLACE, so an id from an older scan finds
  // nothing rather than something else.
  H := 2166136261;
  for var C in IntToStr(Ord(AIssue.Kind)) + '|' + LowerCase(AIssue.FileName) do
  begin
    H := H xor Cardinal(Ord(C));
    H := H * 16777619;
  end;
  Result := IntToHex(H, 8);
end;

procedure CountLineEndings(const ABytes: TBytes; out ACrLf, ALfOnly, ACrOnly: Integer);
var
  I, N, Step, Start: Integer;
  BigEndian: Boolean;

  function Unit_(AIdx: Integer): Word;
  begin
    if Step = 1 then Exit(ABytes[AIdx]);
    if BigEndian then
      Result := (ABytes[AIdx] shl 8) or ABytes[AIdx + 1]
    else
      Result := ABytes[AIdx] or (ABytes[AIdx + 1] shl 8);
  end;

begin
  ACrLf := 0;
  ALfOnly := 0;
  ACrOnly := 0;
  N := Length(ABytes);
  Step := 1;
  Start := 0;
  BigEndian := False;
  if (N >= 2) and (ABytes[0] = $FF) and (ABytes[1] = $FE) then
  begin
    Step := 2;
    Start := 2;
  end
  else if (N >= 2) and (ABytes[0] = $FE) and (ABytes[1] = $FF) then
  begin
    Step := 2;
    Start := 2;
    BigEndian := True;
  end;
  I := Start;
  while I + Step - 1 < N do
  begin
    case Unit_(I) of
      13:
        if (I + 2 * Step - 1 < N) and (Unit_(I + Step) = 10) then
        begin
          Inc(ACrLf);
          Inc(I, Step);
        end
        else
          Inc(ACrOnly);
      10:
        Inc(ALfOnly);
    end;
    Inc(I, Step);
  end;
end;

function ConvertFileToCrLf(const AFile: string): Boolean;
var
  Src, Dst: TBytes;
  N, I, J, Step, Start: Integer;
  BigEndian: Boolean;
  CrLf, LfOnly, CrOnly: Integer;

  function U(AIdx: Integer): Word;
  begin
    if Step = 1 then Exit(Src[AIdx]);
    if BigEndian then
      Result := (Src[AIdx] shl 8) or Src[AIdx + 1]
    else
      Result := Src[AIdx] or (Src[AIdx + 1] shl 8);
  end;

  procedure Put(AValue: Word);
  begin
    if Step = 1 then
    begin
      Dst[J] := Byte(AValue);
      Inc(J);
    end
    else if BigEndian then
    begin
      Dst[J] := Byte(AValue shr 8);
      Dst[J + 1] := Byte(AValue);
      Inc(J, 2);
    end
    else
    begin
      Dst[J] := Byte(AValue);
      Dst[J + 1] := Byte(AValue shr 8);
      Inc(J, 2);
    end;
  end;

begin
  Result := False;
  try
    Src := TFile.ReadAllBytes(AFile);
  except
    Exit;
  end;
  CountLineEndings(Src, CrLf, LfOnly, CrOnly);
  if LfOnly + CrOnly = 0 then Exit;
  N := Length(Src);
  Step := 1;
  Start := 0;
  BigEndian := False;
  if (N >= 2) and (Src[0] = $FF) and (Src[1] = $FE) then
  begin
    Step := 2; Start := 2;
  end
  else if (N >= 2) and (Src[0] = $FE) and (Src[1] = $FF) then
  begin
    Step := 2; Start := 2; BigEndian := True;
  end;
  SetLength(Dst, N + (LfOnly + CrOnly) * Step);
  J := 0;
  for I := 0 to Start - 1 do
  begin
    Dst[J] := Src[I];
    Inc(J);
  end;
  I := Start;
  while I + Step - 1 < N do
  begin
    case U(I) of
      13:
        begin
          Put(13);
          Put(10);
          if (I + 2 * Step - 1 < N) and (U(I + Step) = 10) then Inc(I, Step);
        end;
      10:
        begin
          Put(13);
          Put(10);
        end;
    else
      Put(U(I));
    end;
    Inc(I, Step);
  end;
  // an odd trailing byte of a broken UTF-16 file stays as it was
  while I < N do
  begin
    Dst[J] := Src[I];
    Inc(J);
    Inc(I);
  end;
  SetLength(Dst, J);
  try
    TFile.WriteAllBytes(AFile, Dst);
    Result := True;
  except
    Result := False;
  end;
end;

function ClassifyLineEndings(const ABytes: TBytes): TLineEndingInfo;
var
  Sorts: Integer;
begin
  Result := Default(TLineEndingInfo);
  CountLineEndings(ABytes, Result.CrLf, Result.LfOnly, Result.CrOnly);
  Sorts := 0;
  if Result.CrLf > 0 then Inc(Sorts);
  if Result.LfOnly > 0 then Inc(Sorts);
  if Result.CrOnly > 0 then Inc(Sorts);
  if Sorts = 0 then
    Result.Kind := lekNone
  else if Sorts > 1 then
    Result.Kind := lekMixed
  else if Result.CrLf > 0 then
    Result.Kind := lekCrLf
  else if Result.LfOnly > 0 then
    Result.Kind := lekLf
  else
    Result.Kind := lekCr;
  Result.NeedsFix := not (Result.Kind in [lekNone, lekCrLf]);
end;

function LineEndingProblemText(const AInfo: TLineEndingInfo): string;

  function Counts: string;
  begin
    Result := '';
    if AInfo.CrLf > 0 then
      Result := Format('%d CRLF', [AInfo.CrLf]);
    if AInfo.LfOnly > 0 then
    begin
      if Result <> '' then Result := Result + ', ';
      Result := Result + Format('%d LF', [AInfo.LfOnly]);
    end;
    if AInfo.CrOnly > 0 then
    begin
      if Result <> '' then Result := Result + ', ';
      Result := Result + Format('%d CR', [AInfo.CrOnly]);
    end;
  end;

begin
  case AInfo.Kind of
    lekMixed:
      Result := Format('MIXED line endings (%s)', [Counts]);
    lekLf:
      Result := Format('LF only, no CR (%d line(s))', [AInfo.LfOnly]);
    lekCr:
      Result := Format('CR only, no LF (%d line(s))', [AInfo.CrOnly]);
  else
    Result := '';
  end;
end;

function LineEndingCheckCandidate(const AFileName: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(AFileName));
  Result := (Ext = '.pas') or (Ext = '.inc') or (Ext = '.dpr') or
            (Ext = '.dpk') or (Ext = '.inc');
end;

function UnderInstallRoot(const AFileName, AInstallRoot: string): Boolean;
begin
  Result := (AInstallRoot <> '') and
    LowerCase(TPath.GetFullPath(AFileName)).StartsWith(
      LowerCase(IncludeTrailingPathDelimiter(TPath.GetFullPath(AInstallRoot))));
end;

function LooksBinaryContent(const ABytes: TBytes): Boolean;
var
  N, I, Limit: Integer;
  Utf16: Boolean;
begin
  N := Length(ABytes);
  // A binary .dfm/.fmx starts with $FF followed by 'TPF0'. The extension
  // gate should keep those out; this catches one saved under another name.
  if (N >= 5) and (ABytes[0] = $FF) and (ABytes[1] = Ord('T')) and
     (ABytes[2] = Ord('P')) and (ABytes[3] = Ord('F')) and
     (ABytes[4] = Ord('0')) then
    Exit(True);
  // UTF-16 text is full of NUL bytes, so the NUL test only applies when no
  // UTF-16 BOM says otherwise - CountLineEndings and ConvertFileToCrLf
  // both read those files by code unit.
  Utf16 := (N >= 2) and
           (((ABytes[0] = $FF) and (ABytes[1] = $FE)) or
            ((ABytes[0] = $FE) and (ABytes[1] = $FF)));
  if Utf16 then Exit(False);
  Limit := N;
  if Limit > 4096 then Limit := 4096;
  for I := 0 to Limit - 1 do
    if ABytes[I] = 0 then Exit(True);
  Result := False;
end;

function LineEndingChoiceEffect(AChoice: TLineEndingChoice): TLineEndingEffect;
begin
  Result := Default(TLineEndingEffect);
  case AChoice of
    lecCorrectOnce:
      Result.Correct := True;
    lecCorrectSession:
      begin
        Result.Correct := True;
        Result.AlwaysThisSession := True;
      end;
    lecCorrectAlways:
      begin
        Result.Correct := True;
        Result.AlwaysThisSession := True;
        Result.AlwaysSaved := True;
      end;
    lecLeave:
      // The only answer that silences these files: the user said so.
      Result.DoNotAskAgain := True;
    lecCancel:
      // Nothing. Not even "do not ask again" - that is what Leave it is
      // for, and a cancelled dialog must cost the user nothing.
      ;
  end;
end;

function LineEndingChoiceOfIndex(AIndex: Integer;
  AConfirmed: Boolean): TLineEndingChoice;
begin
  if not AConfirmed then Exit(lecCancel);
  case AIndex of
    0: Result := lecCorrectOnce;
    1: Result := lecCorrectSession;
    2: Result := lecCorrectAlways;
    3: Result := lecLeave;
  else
    Result := lecCancel;   // an answer we did not offer changes nothing
  end;
end;

function LineEndingIssueId(const AFileName: string): string;
var
  H: Cardinal;
  B: Byte;
  Key: string;
begin
  Key := LowerCase(AFileName);
  H := 2166136261;
  for B in TEncoding.UTF8.GetBytes(Key) do
    H := (H xor B) * 16777619;
  Result := IntToHex(H, 8);
end;

function CollectLineEndingIssues(const AFiles: TArray<string>;
  const AReader: TLineEndingReader;
  const AProgress: TLineEndingProgress;
  const AInstallRoot: string): TArray<TLineEndingIssue>;
var
  Seen: TDictionary<string, Boolean>;
  Bytes: TBytes;
  Issue: TLineEndingIssue;
  I: Integer;
  Key: string;
begin
  Result := nil;
  if not Assigned(AReader) then Exit;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for I := 0 to High(AFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(I + 1, Length(AFiles), AFiles[I]);
      Key := LowerCase(AFiles[I]);
      if Seen.ContainsKey(Key) then Continue;
      Seen.Add(Key, True);
      // The same two gates the warning at file open uses: the fix rewrites
      // BYTES, so a form file or anything binary is none of our business.
      if not LineEndingCheckCandidate(AFiles[I]) then Continue;
      if UnderInstallRoot(AFiles[I], AInstallRoot) then Continue;
      if not AReader(AFiles[I], Bytes) then Continue;
      if LooksBinaryContent(Bytes) then Continue;
      Issue := Default(TLineEndingIssue);
      Issue.Info := ClassifyLineEndings(Bytes);
      if not Issue.Info.NeedsFix then Continue;
      Issue.FileName := AFiles[I];
      Issue.Id := LineEndingIssueId(AFiles[I]);
      Result := Result + [Issue];
    end;
  finally
    Seen.Free;
  end;
end;

function SelectReportRows(const AIds, AActedIds: TArray<string>;
  AMax: Integer): TArray<Integer>;
var
  I, Left: Integer;

  function Acted(const AId: string): Boolean;
  var
    S: string;
  begin
    for S in AActedIds do
      if SameText(S, AId) then Exit(True);
    Result := False;
  end;

begin
  Result := nil;
  Left := AMax;
  // The order of the findings is kept - an acted row does not jump to the
  // front, it simply cannot be dropped.
  for I := 0 to High(AIds) do
    if Acted(AIds[I]) then
      Result := Result + [I]
    else if Left > 0 then
    begin
      Result := Result + [I];
      Dec(Left);
    end;
end;

function UnknownReportIds(const AIds, ANamed: TArray<string>): TArray<string>;
var
  Name, Id: string;
  Known: Boolean;
begin
  Result := nil;
  for Name in ANamed do
  begin
    if Trim(Name) = '' then Continue;
    Known := False;
    for Id in AIds do
      if SameText(Id, Name) then
      begin
        Known := True;
        Break;
      end;
    if not Known then Result := Result + [Name];
  end;
end;

function DebugFixOfferNote(AFixable: Integer): string;
begin
  if AFixable <= 0 then Exit('');
  Result := Format('%d finding(s) of this scan carry a cure this tool can ' +
    'apply: a .dcu that is older than its source AND whose source the ' +
    'compiler could use instead - apply=true with such an id moves that ' +
    '.dcu into a subdirectory, so DelphiLSP reads the source again ' +
    '(lsp_restart afterwards). Every other finding names its own cure, ' +
    'which is a build.', [AFixable]);
end;

function AppliedByIdsNote(ADone, AFailed, AUnknown: Integer;
  const AThing, ADoneText, AWhere: string): string;
begin
  if ADone > 0 then
  begin
    Result := Format('%d %s %s, %s.', [ADone, AThing, ADoneText, AWhere]);
    if AFailed > 0 then
      Result := Result + Format(' %d could NOT be done - "result" says why ' +
        'per entry.', [AFailed]);
    Result := Result + ' The counts above are from the scan BEFORE the ' +
      'change - list again for the state now.';
  end
  else if AFailed > 0 then
    Result := Format('NOTHING was changed: all %d %s this call named failed ' +
      '- "result" says why per entry.', [AFailed, AThing])
  else
    Result := 'NOTHING was changed - no finding of this scan was acted on.';

  if AUnknown > 0 then
    Result := Result + Format(' %d of the ids you named matched no finding ' +
      'of this scan ("unknownIds"), so nothing was done for them: already ' +
      'in order, not something this check fixes, or an id from an older ' +
      'scan - list again.', [AUnknown]);
end;

// Both checks answer the same two questions about their own findings, so
// they ask ONE implementation - only the record differs.
function IdsOf(const AIssues: TArray<TLineEndingIssue>): TArray<string>; overload;
var
  I: Integer;
begin
  SetLength(Result, Length(AIssues));
  for I := 0 to High(AIssues) do Result[I] := AIssues[I].Id;
end;

function IdsOf(const AIssues: TArray<TEncodingIssue>): TArray<string>; overload;
var
  I: Integer;
begin
  SetLength(Result, Length(AIssues));
  for I := 0 to High(AIssues) do Result[I] := AIssues[I].Id;
end;

function SelectLineEndingRows(const AIssues: TArray<TLineEndingIssue>;
  const AActedIds: TArray<string>; AMax: Integer): TArray<Integer>;
begin
  Result := SelectReportRows(IdsOf(AIssues), AActedIds, AMax);
end;

function UnknownLineEndingIds(const AIssues: TArray<TLineEndingIssue>;
  const AIds: TArray<string>): TArray<string>;
begin
  Result := UnknownReportIds(IdsOf(AIssues), AIds);
end;

function EncodingIssueId(const AFileName: string): string;
var
  H: Cardinal;
  B: Byte;
begin
  H := 2166136261;
  for B in TEncoding.UTF8.GetBytes(LowerCase(AFileName)) do
    H := (H xor B) * 16777619;
  Result := IntToHex(H, 8);
end;

function CollectEncodingIssues(const AFiles: TArray<string>;
  AFrom: TSourceEncodings; ATo: TSourceEncoding;
  const AReader: TEncodingReader;
  const AProgress: TLineEndingProgress;
  const AInstallRoot: string): TArray<TEncodingIssue>;
var
  Seen: TDictionary<string, Boolean>;
  Bytes: TBytes;
  Issue: TEncodingIssue;
  I: Integer;
  Key: string;
  Found: TSourceEncoding;
begin
  Result := nil;
  if not Assigned(AReader) then Exit;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for I := 0 to High(AFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(I + 1, Length(AFiles), AFiles[I]);
      Key := LowerCase(AFiles[I]);
      if Seen.ContainsKey(Key) then Continue;
      Seen.Add(Key, True);
      // The same two gates the line-ending check uses: this rewrites
      // BYTES, so a form file or anything binary is none of our business.
      if not LineEndingCheckCandidate(AFiles[I]) then Continue;
      if UnderInstallRoot(AFiles[I], AInstallRoot) then Continue;
      if not AReader(AFiles[I], Bytes) then Continue;
      if LooksBinaryContent(Bytes) then Continue;
      Found := ClassifySourceEncoding(Bytes);
      if Found = ATo then Continue;   // already there
      if not (Found in AFrom) then Continue;
      Issue := Default(TEncodingIssue);
      Issue.FileName := AFiles[I];
      Issue.Found := Found;
      Issue.Id := EncodingIssueId(AFiles[I]);
      Result := Result + [Issue];
    end;
  finally
    Seen.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  directives
// ---------------------------------------------------------------------------

// Walks AContent and hands every "{$...}" / "(*$...*)" directive body
// (without the braces) with its 1-based line to AProc. Comments ({ } (* *)
// //) and string literals are skipped.
procedure ForEachDirective(const AContent: string;
  const AProc: TProc<Integer, string, string>);
var
  I, N, Line, StartLine: Integer;
  C: Char;
begin
  N := Length(AContent);
  I := 1;
  Line := 1;
  while I <= N do
  begin
    C := AContent[I];
    if C = #10 then
    begin
      Inc(Line);
      Inc(I);
      Continue;
    end;
    if C = '''' then
    begin
      Inc(I);
      while (I <= N) and not CharInSet(AContent[I], ['''', #10]) do Inc(I);
      Inc(I);
      Continue;
    end;
    if (C = '/') and (I < N) and (AContent[I + 1] = '/') then
    begin
      while (I <= N) and (AContent[I] <> #10) do Inc(I);
      Continue;
    end;
    if (C = '{') or ((C = '(') and (I < N) and (AContent[I + 1] = '*')) then
    begin
      var Paren := C = '(';
      var Open := I;
      StartLine := Line;
      if Paren then Inc(I, 2) else Inc(I);
      var BodyStart := I;
      while I <= N do
      begin
        if AContent[I] = #10 then Inc(Line);
        if (not Paren) and (AContent[I] = '}') then Break;
        if Paren and (AContent[I] = '*') and (I < N) and (AContent[I + 1] = ')') then Break;
        Inc(I);
      end;
      if (BodyStart <= N) and (AContent[BodyStart] = '$') then
      begin
        var Body := Copy(AContent, BodyStart + 1, I - BodyStart - 1);
        var Full: string;
        if Paren then
          Full := Copy(AContent, Open, I - Open + 2)
        else
          Full := Copy(AContent, Open, I - Open + 1);
        AProc(StartLine, Body, Full);
      end;
      if Paren then Inc(I, 2) else Inc(I);
      Continue;
    end;
    Inc(I);
  end;
end;

function FindDebugDirectives(const AContent: string): TArray<TSourceDirective>;
var
  Found: TList<TSourceDirective>;
begin
  Found := TList<TSourceDirective>.Create;
  try
    ForEachDirective(AContent,
      procedure(ALine: Integer; ABody, AFull: string)
      var
        D: TSourceDirective;
        U, Name, Value: string;
        SpaceP: Integer;
      begin
        U := UpperCase(Trim(ABody));
        // long form: "DEBUGINFO OFF", "LOCALSYMBOLS OFF", "OPTIMIZATION ON"
        SpaceP := Pos(' ', U);
        if SpaceP > 0 then
        begin
          Name := Copy(U, 1, SpaceP - 1);
          Value := Trim(Copy(U, SpaceP + 1, MaxInt));
          D.Line := ALine;
          D.Text := AFull;
          D.What := '';
          if (Name = 'DEBUGINFO') and StartsStr('OFF', Value) then
            D.What := 'debug information off'
          else if (Name = 'LOCALSYMBOLS') and StartsStr('OFF', Value) then
            D.What := 'local symbols off'
          else if (Name = 'OPTIMIZATION') and StartsStr('ON', Value) then
            D.What := 'optimization on';
          if D.What <> '' then Found.Add(D);
          Exit;
        end;
        // short, possibly combined form: "D-", "D-,L-", "O+,W-"
        for var Part in U.Split([',']) do
        begin
          var P := Trim(Part);
          if Length(P) <> 2 then Continue;
          D.Line := ALine;
          D.Text := AFull;
          D.What := '';
          if P = 'D-' then D.What := 'debug information off'
          else if P = 'L-' then D.What := 'local symbols off'
          else if P = 'O+' then D.What := 'optimization on';
          if D.What <> '' then Found.Add(D);
        end;
      end);
    Result := Found.ToArray;
  finally
    Found.Free;
  end;
end;

function FindIncludeNames(const AContent: string): TArray<string>;
var
  Found: TList<string>;
begin
  Found := TList<string>.Create;
  try
    ForEachDirective(AContent,
      procedure(ALine: Integer; ABody, AFull: string)
      var
        B, U, Name: string;
      begin
        B := Trim(ABody);
        U := UpperCase(B);
        if StartsStr('INCLUDE ', U) then
          Name := Trim(Copy(B, 9, MaxInt))
        else if StartsStr('I ', U) then
          Name := Trim(Copy(B, 3, MaxInt))
        else
          Exit;
        // "{$I+}" / "{$I-}" is I/O checking, not an include
        if (Name = '') or (Name = '+') or (Name = '-') then Exit;
        if (Length(Name) >= 2) and (Name[1] = '''') and (Name[Length(Name)] = '''') then
          Name := Copy(Name, 2, Length(Name) - 2);
        Found.Add(Name);
      end);
    Result := Found.ToArray;
  finally
    Found.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  the check
// ---------------------------------------------------------------------------

function NormDir(const ADir: string): string;
begin
  Result := ExcludeTrailingPathDelimiter(TPath.GetFullPath(ADir));
end;

function IsBelow(const AFile: string; const ADirs: TArray<string>): Boolean;
begin
  for var D in ADirs do
    if (D <> '') and StartsText(IncludeTrailingPathDelimiter(D), AFile) then
      Exit(True);
  Result := False;
end;

function FileTime(const AFile: string): TDateTime;
begin
  try
    Result := TFile.GetLastWriteTime(AFile);
  except
    Result := 0;
  end;
end;

function TimeText(const AFile: string): string;
begin
  Result := FormatDateTime('yyyy-mm-dd hh:nn', FileTime(AFile));
end;

function RunDebugConsistencyCheck(const AInput: TDebugCheckInput;
  const AProgress: TDebugProgress): TArray<TDebugIssue>;
var
  Issues: TList<TDebugIssue>;
  SearchDirs, IgnoreDirs: TArray<string>;
  Sources: TDictionary<string, string>;   // UPPER unit name -> project source path
  Checked: TDictionary<string, Boolean>;  // files already line-checked
  Step, Total: Integer;
  Cancelled: Boolean;

  procedure Add(AKind: TDebugIssueKind; ASev: TDebugIssueSeverity;
    const AFile: string; ALine: Integer; const AReason, AHint: string;
    AFixable: Boolean = False);
  var
    I: TDebugIssue;
  begin
    I.Kind := AKind;
    I.Severity := ASev;
    I.FileName := AFile;
    I.Line := ALine;
    I.Reason := AReason;
    I.Hint := AHint;
    I.Fixable := AFixable;
    Issues.Add(I);
  end;

  function Tick(const AText: string): Boolean;
  begin
    Inc(Step);
    if Assigned(AProgress) and not AProgress(Step, Total, AText) then
      Cancelled := True;
    Result := not Cancelled;
  end;

  // first directory of the compiler's search order holding AFileName
  function FirstOnSearchPath(const AFileName: string): string;
  begin
    for var D in SearchDirs do
      if TFile.Exists(TPath.Combine(D, AFileName)) then
        Exit(TPath.Combine(D, AFileName));
    Result := '';
  end;

  procedure CheckLineEndings(const AFile: string; AOpen: Boolean);
  var
    Bytes: TBytes;
    CrLf, Lf, Cr: Integer;
  begin
    if Checked.ContainsKey(UpperCase(AFile)) then Exit;
    Checked.Add(UpperCase(AFile), True);
    if not TFile.Exists(AFile) then Exit;
    try
      Bytes := TFile.ReadAllBytes(AFile);
    except
      Exit;
    end;
    CountLineEndings(Bytes, CrLf, Lf, Cr);
    if Lf + Cr = 0 then Exit;
    var Hint := 'the debugger maps breakpoints by CRLF lines - convert the file to CRLF ' +
      '(and check git''s core.autocrlf / .gitattributes, or it comes back)';
    if AOpen then
      Hint := Hint + '. The file is open in the editor: close it before converting';
    if CrLf = 0 then
      Add(dkLineEndings, dsProblem, AFile, 0,
        Format('line endings are LF only (%d lines)', [Lf + Cr]), Hint, not AOpen)
    else
      Add(dkLineEndings, dsProblem, AFile, 0,
        Format('mixed line endings: %d CRLF, %d LF only, %d CR only', [CrLf, Lf, Cr]),
        Hint, not AOpen);
  end;

  function IsOpen(const AFile: string): Boolean;
  begin
    for var F in AInput.OpenFiles do
      if SameText(F, AFile) then Exit(True);
    Result := False;
  end;

  procedure CheckDirectivesAndIncludes(const AFile: string);
  var
    Content: string;
  begin
    try
      Content := TFile.ReadAllText(AFile);
    except
      Exit;
    end;
    for var D in FindDebugDirectives(Content) do
      Add(dkSourceDirective, dsWarning, AFile, D.Line,
        Format('%s turns %s', [D.Text, D.What]),
        'breakpoints in code compiled this way do not bind or jump around ' +
        '- remove the directive or limit it to release builds ({$IFDEF DEBUG})');
    for var Name in FindIncludeNames(Content) do
    begin
      var Inc_ := '';
      var Candidates: TArray<string> := [TPath.Combine(ExtractFilePath(AFile), Name)];
      if ExtractFileExt(Name) = '' then
        Candidates := Candidates + [TPath.Combine(ExtractFilePath(AFile), Name + '.inc')];
      for var C in Candidates do
        if TFile.Exists(C) then
        begin
          Inc_ := C;
          Break;
        end;
      if Inc_ = '' then
      begin
        Inc_ := FirstOnSearchPath(Name);
        if (Inc_ = '') and (ExtractFileExt(Name) = '') then
          Inc_ := FirstOnSearchPath(Name + '.inc');
      end;
      if (Inc_ <> '') and not IsBelow(Inc_, IgnoreDirs) then
        CheckLineEndings(TPath.GetFullPath(Inc_), IsOpen(Inc_));
    end;
  end;

begin
  Issues := TList<TDebugIssue>.Create;
  Sources := TDictionary<string, string>.Create;
  Checked := TDictionary<string, Boolean>.Create;
  try
    Step := 0;
    Cancelled := False;
    Total := Length(AInput.ProjectSources) * 2 + Length(AInput.OpenFiles) + 4;

    SearchDirs := nil;
    for var D in AInput.UnitSearchDirs do
      if (Trim(D) <> '') and TDirectory.Exists(D) then
        SearchDirs := SearchDirs + [NormDir(D)];
    IgnoreDirs := nil;
    for var D in AInput.IgnoreDirs do
      if Trim(D) <> '' then IgnoreDirs := IgnoreDirs + [NormDir(D)];

    for var S in AInput.ProjectSources do
      if SameText(ExtractFileExt(S), '.pas') then
        Sources.AddOrSetValue(UpperCase(ChangeFileExt(ExtractFileName(S), '')), S);

    // ---- 1. duplicate sources and stray DCUs of project units ---------------
    for var S in AInput.ProjectSources do
    begin
      if not Tick('search path: ' + ExtractFileName(S)) then Break;
      if not SameText(ExtractFileExt(S), '.pas') then Continue;
      var UnitFile := ExtractFileName(S);
      var DcuFile := ChangeFileExt(UnitFile, '.dcu');
      for var D in SearchDirs do
      begin
        if IsBelow(D + PathDelim, IgnoreDirs) then Continue;
        var Other := TPath.Combine(D, UnitFile);
        if TFile.Exists(Other) and not SameText(TPath.GetFullPath(Other), TPath.GetFullPath(S)) then
          Add(dkDuplicateSource, dsWarning, Other, 0,
            Format('second copy of project unit %s (project uses %s)', [UnitFile, S]),
            'opening this copy (e.g. via "find declaration") shows code that is NOT ' +
            'compiled - breakpoints set there never bind. Delete or rename one copy');
        var Dcu := TPath.Combine(D, DcuFile);
        var OwnDcuDir := AInput.DcuOutputDir;
        if OwnDcuDir = '' then OwnDcuDir := ExtractFilePath(S);
        if TFile.Exists(Dcu) and not SameText(NormDir(D), NormDir(OwnDcuDir)) then
          Add(dkStrayDcu, dsWarning, Dcu, 0,
            Format('precompiled %s from %s outside the project''s DCU output directory',
              [DcuFile, TimeText(Dcu)]),
            'a unit compiled elsewhere can be linked instead of the project''s own ' +
            '(different code, different or no debug info) - delete the stray .dcu');
      end;
    end;

    // ---- 2. open files the compiler does not take from where they are ------
    for var F in AInput.OpenFiles do
    begin
      if Cancelled or not Tick('open file: ' + ExtractFileName(F)) then Break;
      if not SameText(ExtractFileExt(F), '.pas') then Continue;
      if IsBelow(F, IgnoreDirs) then Continue;
      var UnitKey := UpperCase(ChangeFileExt(ExtractFileName(F), ''));
      var ProjectCopy: string;
      if Sources.TryGetValue(UnitKey, ProjectCopy) then
      begin
        if not SameText(TPath.GetFullPath(ProjectCopy), TPath.GetFullPath(F)) then
          Add(dkDuplicateSource, dsProblem, F, 0,
            Format('this copy is open, but the project compiles %s', [ProjectCopy]),
            'breakpoints in this editor tab never bind - close it and open the ' +
            'project''s copy');
        Continue;
      end;
      // not a project unit: which copy does the compiler find?
      var Compiled := FirstOnSearchPath(ExtractFileName(F));
      if Compiled <> '' then
      begin
        if not SameText(TPath.GetFullPath(Compiled), TPath.GetFullPath(F)) then
          Add(dkDuplicateSource, dsProblem, F, 0,
            Format('the compiler takes %s from the search path, not this copy', [Compiled]),
            'breakpoints in this editor tab never bind - open the copy the compiler uses, ' +
            'or fix the search path order');
      end
      else
      begin
        var Dcu := FirstOnSearchPath(ChangeFileExt(ExtractFileName(F), '.dcu'));
        if Dcu <> '' then
          Add(dkDcuOnly, dsWarning, F, 0,
            Format('the compiler does not see this source - it links the precompiled %s (%s)',
              [Dcu, TimeText(Dcu)]),
            'breakpoints bind only if that .dcu was built from exactly this source WITH ' +
            'debug information. To debug it, add the source directory to the project ' +
            'search path (or rebuild the library in debug mode)');
      end;
    end;

    // ---- 3. line endings, directives, include files -------------------------
    for var S in AInput.ProjectSources do
    begin
      if Cancelled or not Tick('line endings: ' + ExtractFileName(S)) then Break;
      CheckLineEndings(TPath.GetFullPath(S), IsOpen(S));
      CheckDirectivesAndIncludes(S);
    end;
    for var F in AInput.OpenFiles do
      if not IsBelow(F, IgnoreDirs) and (SameText(ExtractFileExt(F), '.pas')
        or SameText(ExtractFileExt(F), '.inc')) then
        CheckLineEndings(TPath.GetFullPath(F), True);

    // ---- 4. options of the active configuration -----------------------------
    if not Cancelled and Tick('project options') then
    begin
      var Cfg := AInput.ConfigName;
      if Cfg = '' then Cfg := 'active configuration';
      if OptionState(AInput.DebugInformation) = 0 then
        Add(dkDebugOption, dsProblem, AInput.ProjectFile, 0,
          Format('"%s": debug information is switched off', [Cfg]),
          'Project > Options > Building > Delphi Compiler > Compiling > Debug information');
      if OptionState(AInput.LocalDebugSymbols) = 0 then
        Add(dkDebugOption, dsWarning, AInput.ProjectFile, 0,
          Format('"%s": local symbols are switched off', [Cfg]),
          'local variables cannot be inspected - Project > Options > Compiling > Local symbols');
      if OptionState(AInput.Optimize) = 1 then
        Add(dkDebugOption, dsWarning, AInput.ProjectFile, 0,
          Format('"%s": optimization is switched on', [Cfg]),
          'optimized code merges and drops lines - breakpoints jump or are ignored, ' +
          'variables show "inaccessible". Switch it off for debugging');
    end;

    // ---- 5. target and symbol files -----------------------------------------
    if not Cancelled and Tick('target') and (AInput.TargetFile <> '') then
    begin
      if not TFile.Exists(AInput.TargetFile) then
        Add(dkExecutable, dsInfo, AInput.TargetFile, 0, 'the target does not exist yet',
          'build the project')
      else
      begin
        var TargetTime := FileTime(AInput.TargetFile);
        var Newest := '';
        var NewestTime: TDateTime := 0;
        for var S in AInput.ProjectSources do
        begin
          var T := FileTime(S);
          if T > NewestTime then
          begin
            NewestTime := T;
            Newest := S;
          end;
        end;
        if (Newest <> '') and (NewestTime > IncSecond(TargetTime, 2)) then
          Add(dkExecutable, dsProblem, AInput.TargetFile, 0,
            Format('built %s, but %s was changed later (%s)',
              [TimeText(AInput.TargetFile), ExtractFileName(Newest), TimeText(Newest)]),
            'the running code is older than the source - did the last build fail? ' +
            'Build (not just run) and check the Messages window');
        for var Ext in ['.rsm', '.tds', '.map'] do
        begin
          var Sym := ChangeFileExt(AInput.TargetFile, Ext);
          if TFile.Exists(Sym) and (Abs(SecondsBetween(FileTime(Sym), TargetTime)) > 300) then
            Add(dkSymbolFile, dsWarning, Sym, 0,
              Format('symbol file from %s does not belong to the target from %s',
                [TimeText(Sym), TimeText(AInput.TargetFile)]),
              'stale symbol files make the debugger show wrong lines/variables - ' +
              'delete it and rebuild');
        end;
      end;

      if (AInput.HostApplication <> '')
        and not SameText(TPath.GetFullPath(AInput.HostApplication), TPath.GetFullPath(AInput.TargetFile)) then
      begin
        var IsExe := SameText(ExtractFileExt(AInput.TargetFile), '.exe');
        if IsExe then
          Add(dkHostApplication, dsProblem, AInput.HostApplication, 0,
            Format('Run > Parameters starts %s instead of the built %s',
              [AInput.HostApplication, ExtractFileName(AInput.TargetFile)]),
            'breakpoints in this project bind only if that program loads the code ' +
            'built here - clear the host application for an exe project')
        else
        begin
          Add(dkHostApplication, dsInfo, AInput.HostApplication, 0,
            Format('the debugger starts the host %s, which must load %s',
              [AInput.HostApplication, AInput.TargetFile]),
            'make sure the host loads THIS file and not a copy from its own directory or the PATH');
          if not TFile.Exists(AInput.HostApplication) then
            Add(dkHostApplication, dsProblem, AInput.HostApplication, 0,
              'the host application does not exist', 'Run > Parameters > Host application');
        end;
      end;
    end;

    // ---- 6. copies of a DLL / package target where Windows may load them ----
    if not Cancelled and Tick('binaries') and (AInput.TargetFile <> '')
      and not SameText(ExtractFileExt(AInput.TargetFile), '.exe') then
    begin
      var Name := ExtractFileName(AInput.TargetFile);
      var Dirs := AInput.BinarySearchDirs;
      if AInput.HostApplication <> '' then
        Dirs := [ExtractFilePath(AInput.HostApplication)] + Dirs;
      var Seen := TDictionary<string, Boolean>.Create;
      try
        for var D in Dirs do
        begin
          if Trim(D) = '' then Continue;
          var Cand := TPath.Combine(D, Name);
          var Key := UpperCase(TPath.GetFullPath(Cand));
          if Seen.ContainsKey(Key) then Continue;
          Seen.Add(Key, True);
          if TFile.Exists(Cand) and not SameText(TPath.GetFullPath(Cand), TPath.GetFullPath(AInput.TargetFile)) then
            Add(dkDuplicateBinary, dsWarning, Cand, 0,
              Format('another %s (%s) where Windows looks for it', [Name, TimeText(Cand)]),
              'if the host loads this copy, your breakpoints never bind - delete it or ' +
              'make the host load the built file');
        end;
      finally
        Seen.Free;
      end;
    end;

    // ---- 7. what the LSP reads instead of the sources -----------------------
    // MEASURED 2026-10-20: 103 .dcu in the directory the .delphilsp.json
    // names made DelphiLSP load 60 project units FROM THE DCU, and for
    // their symbols it then reports a definition without a file
    // ("definition found in file:///(-1:0)") - so go to definition, find
    // references and rename come back empty for exactly those units. The
    // DCUs were OLDER than the sources and were used anyway, which is why
    // keeping them current is no answer: the directory has to be empty.
    if not Cancelled and Tick('LSP configuration') and
       (AInput.LspConfigFile <> '') then
    begin
      var LspDcu := LspDcuDirOfConfig(AInput.LspConfigFile, AInput.ProjectDir);
      if LspDcu <> '' then
      begin
        var DcuState := InspectLspDcuDir(LspDcu, AInput.ProjectSources,
          LspUnitSearchDirs(AInput.LspConfigFile, AInput.ProjectDir));
        // A PARTIALLY built output directory is the state that was measured
        // to break go to definition: the units that are there come from one
        // build, the missing one has to come from another (the installed
        // package''s .dcp, a copy on the library path) - and then the
        // compiler resolves it without a source position. A directory with
        // NO .dcu at all is not reported: everything then comes from one
        // place, which is consistent.
        if (DcuState.Present > 0) and (Length(DcuState.Missing) > 0) then
          Add(dkLspDcu, dsWarning, LspDcu, 0,
            Format('%d of %d project units have no .dcu in the DCU output ' +
              'of this configuration, %d do',
              [Length(DcuState.Missing), DcuState.UnitCount,
               DcuState.Present]),
            'the LSP searches this directory the way the compiler does. ' +
            'MEASURED: with a COMPLETE directory go to definition across ' +
            'units answers with the source position - and with one unit ' +
            'missing from it the same query answered nothing, because that ' +
            'unit then came from another build (the server logged "Kibitz ' +
            'definition found in file:///(-1:0)"); putting the file back ' +
            'made it answer again. So the cure is to BUILD the project for ' +
            'this configuration, not to remove anything. First missing: ' +
            string.Join(', ', Copy(DcuState.Missing, 0,
              Min(5, Length(DcuState.Missing)))));
        for var I := 0 to High(DcuState.Stale) do
          Add(dkLspDcu, dsWarning, DcuState.Stale[I], 0,
            Format('this .dcu (%s) is OLDER than its source (%s)',
              [TimeText(DcuState.Stale[I]),
               TimeText(DcuState.StaleSource[I])]),
            IfThen(DcuState.StaleReplaceable[I],
              'the compiler can reach the source as well (its directory is ' +
              'on the LSP''s search path), so this .dcu can shadow it: the ' +
              'answers would describe an older version of the file. ' +
              'Build the project, or move the .dcu aside (this check can) ' +
              'and restart the session with lsp_restart',
              'the source is in no directory the LSP searches, so this ' +
              '.dcu is the ONLY way to that unit - moving it aside does ' +
              'not help, it costs go to definition for the unit entirely ' +
              '(measured). Build the project for this configuration so ' +
              'the .dcu matches the source again'),
            DcuState.StaleReplaceable[I]);
      end;
    end;

    Result := Issues.ToArray;
  finally
    Checked.Free;
    Sources.Free;
    Issues.Free;
  end;
end;

end.
