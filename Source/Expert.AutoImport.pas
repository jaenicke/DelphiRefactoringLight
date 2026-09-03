(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.AutoImport;

// "Add unit for a missing symbol" (auto-import), driven by DelphiLSP's own
// diagnostics: an undeclared identifier is reported as code E2003, from which
// we take the identifier, look it up in the background unit index, and add the
// declaring unit to the file's uses clause.
//
// Two entry points:
//   AddUnitForIdentifierAtCursor - resolves the identifier under the caret,
//     showing a small caret-anchored chooser popup (VS-like) when there is
//     more than one candidate unit or section choice.
//   ResolveMissingUnits          - a batch dialog listing every unresolved
//     identifier in the file with its suggested unit + target section.
//
// Live indicator ("lightbulb"): StartAutoImportLive watches the active
// buffer. Diagnostics come from TWO COMPLEMENTARY sources that merge:
//   * IDE: Expert.StructureErrors pushes the IDE's own Error Insight
//     results (Structure view notifier) via LiveReportErrorDiags - fast,
//     but ERRORS only.
//   * Both hosts: a low-frequency poller runs our own LSP diagnostics in
//     the background once per idle buffer state (all ToolsAPI access on
//     the main thread; the worker only waits on LSP). This is the only
//     source of HINT fixes (H2443/H2164); its result is the superset and
//     replaces the structure result for the same buffer state.
// When the caret sits on a line with a missing identifier whose unit is
// known, a small non-focus-stealing button appears just below the caret;
// clicking it opens the chooser popup. The hint is bound to the error line:
// it hides when the caret moves elsewhere, the line scrolls out of view or
// the buffer changes.

interface

uses
  System.SysUtils, Lsp.Protocol, Expert.UsesEditor, Expert.UnitIndex;

type
  TQuickFixKind = (qfAddUnit, qfRenameIdent, qfFixUsesName, qfRemoveUses,
    qfAlignHeader, qfRemoveVar, qfInsertSemi, qfInitVar, qfRemoveAssign,
    qfAddReintroduce, qfImplStub, qfClassStub, qfRemoveToken,
    qfRemovePrivate);

  /// <summary>One concrete, applicable fix action derived from a compiler
  ///  diagnostic. See ResolveQuickFixes for the providers.</summary>
  TQuickFix = record
    Kind: TQuickFixKind;
    Line: Integer;              // 0-based line the fix anchors to
    Caption: string;            // short hint-label text
    // token replacement / marker anchor:
    Col: Integer;               // 0-based start column of the token
    TokenLen: Integer;
    NewText: string;            // replacement / generated text (kind-specific)
    // uses handling:
    Identifier: string;         // affected identifier (kind-specific)
    UnitNames: TArray<string>;  // qfAddUnit: candidate units (deduped)
    Section: TUsesSection;      // target section for uses additions
    FollowUpUnit: string;       // qfRenameIdent: unit to add after renaming
    OldUnit: string;            // qfRemoveUses: unit to remove
    AuxLine: Integer;           // qfInitVar: 0-based line of the routine's 'begin'
  end;

/// <summary>Turns compiler diagnostics into concrete quick fixes. Pure
///  function of (content, diagnostics) plus the immutable unit-index
///  snapshot - safe on any thread. Providers: E2003 (add unit / "did you
///  mean" rename), F2613/F2063 (fix or remove a uses entry), E2037 (align
///  the implementation header with the declaration), H2443 (add the unit
///  the hint names to uses), H2164 (remove an unused variable),
///  E2029/E2066 (insert a missing ';'), W1036 (initialize a variable with
///  Default at the routine's begin), H2077 (remove a dead assignment),
///  W1010 (add reintroduce), E2065 (generate an empty implementation for
///  an unsatisfied forward - method, routine, or class forward
///  declaration stub).</summary>
function ResolveQuickFixes(const AContent: string;
  const ADiags: TArray<TLspErrorDiag>): TArray<TQuickFix>;

/// <summary>Identifier of an E2003 diagnostic. Prefers the QUOTED name from
///  the (localized) message - for generic types the compiler reports
///  'TList<>' while the range points at/into the type ARGUMENTS, so the
///  token at the range would be 'Integer' (or a type parameter). Falls back
///  to the token at the range position. '' when nothing usable; ACol0/ALen
///  anchor the identifier on its line (markers, rename replacement).</summary>
function E2003IdentFromDiag(const ALines: TArray<string>;
  const D: TLspErrorDiag; out ACol0, ALen: Integer): string;

/// <summary>True when the identifier at ACol0/ALen (0-based col) is
///  followed by '<' - a generic instantiation ("TList<Integer>").</summary>
function IsGenericUseAt(const ALine: string; ACol0, ALen: Integer): Boolean;

/// <summary>True when AIdent is visibly DECLARED somewhere in AContent -
///  as a parameter, local/global variable or field ('X: Typ', or
///  'var/const/out X'). An E2003 "undeclared identifier" for such an
///  identifier is an Error-Insight false positive (frequent around
///  anonymous-method parameters) or a stale diagnostic: neither
///  "add unit" nor "did you mean" makes sense then.</summary>
function IdentDeclaredInFile(const AContent, AIdent: string): Boolean;

/// <summary>Drops candidate units whose declaration FORM cannot satisfy
///  the use form: "TList<...>" is never System.Classes' non-generic TList,
///  a bare "TList" never the generic one. Only filters when at least one
///  matching-form candidate exists - otherwise all hits are kept.</summary>
function FilterHitsByGenericUse(const AHits: TArray<TFindUnitHit>;
  AGenericUse: Boolean): TArray<TFindUnitHit>;

/// <summary>Executes one quick fix. AUnitChoice picks the candidate unit
///  for qfAddUnit (index into UnitNames).</summary>
function ApplyQuickFix(const AFile: string; const AFix: TQuickFix;
  AUnitChoice: Integer): Boolean;

/// <summary>E2037 fix: rewrites the implementation header at/around ALine0
///  (0-based) so parameters and return type match the declaration, keeping
///  the implementation's parameter names. False when the declaration is
///  not found or ambiguous (overloads).</summary>
function AlignImplHeaderToDecl(const AFile: string; ALine0: Integer): Boolean;

procedure AddUnitForIdentifierAtCursor;
procedure ResolveMissingUnits;

procedure StartAutoImportLive;
procedure StopAutoImportLive;

/// <summary>True when the live checker holds FRESH results (matching the
///  current buffer content) for AFile; ACount is then the number of missing
///  identifiers with at least one known unit. Used by menus to enable /
///  annotate their quick-fix entry. False = no live data (caller should
///  keep the entry enabled and let the on-demand path decide).</summary>
function LiveFreshInfo(const AFile: string; out ACount: Integer): Boolean;

/// <summary>Tags for the diagnostics sources. WHICH source produced the
///  current result decides what an empty payload means: from the SAME
///  source it is a retraction ("re-evaluated, the error is gone") and the
///  fixes go away at once; from a DIFFERENT one it is just silence and
///  the existing fixes stay.</summary>
const
  SrcLsp = 'lsp';
  SrcStructure = 'structure';
  SrcCompiler = 'compiler';

/// <summary>Feeds the live checker with error diagnostics obtained from an
///  EXTERNAL source - in the IDE that is the Structure view, which mirrors
///  Delphi's own Error Insight (see Expert.StructureErrors). AContent must
///  be the buffer content the diagnostics refer to. These fast, errors-only
///  results are published immediately; the complementary LSP pass (which
///  adds the hint fixes) still runs for the same buffer state and replaces
///  them as the superset. Main thread only.</summary>
procedure LiveReportErrorDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>; const ASource: string = SrcStructure);

/// <summary>Opens the quick-fix chooser popup for the fixes on the CARET
///  line, taken from the FRESH live results of AFile. False when there is
///  no fresh fix there. Used by the Structure-view double-click
///  integration (the IDE navigates to the error first).</summary>
function LiveShowFixAtCaret(const AFile: string): Boolean;

/// <summary>True while the live checker is still resolving (or has a
///  resolve queued) - a False from LiveShowFixAtCaret then means "not
///  ready yet", not "nothing to fix"; the caller may retry shortly.</summary>
function LiveResolveBusy: Boolean;

/// <summary>The FRESH quick fixes anchored to the 0-based line ALine0 of
///  AFile - read-only access for the editor line markers
///  (Expert.QuickFixMarkers). Empty when the live results are stale or
///  belong to another file. Main thread only.</summary>
function LiveFixesForLine(const AFile: string; ALine0: Integer): TArray<TQuickFix>;

/// <summary>ALL quick fixes the live check resolved for AFile - only when
///  the results are FRESH for the active buffer state (False otherwise).
///  Main thread only.</summary>
function LiveAllFixes(const AFile: string; out AFixes: TArray<TQuickFix>): Boolean;

/// <summary>True when AUnit is already usable from ASection of AContent
///  (an interface-clause entry serves both sections, an implementation
///  entry only the implementation). The guard that tells a STALE
///  add-unit diagnostic - our LSP can lag one edit behind - from a real
///  one.</summary>
function UnitReachableFrom(const AContent, AUnit: string;
  ASection: TUsesSection): Boolean;

/// <summary>Drops every published live result. Called when the active
///  PROJECT changes: the results are keyed by (file, content hash), so
///  an untouched buffer would otherwise keep showing fixes resolved for
///  the old project context (stale diagnostics survive a reopen).
///  State only - the next poll tick hides the hint.</summary>
procedure LiveResetResults;

/// <summary>Snapshot of the live checker for the status window: the
///  watched buffer, whether an analysis/resolution is in flight, how many
///  fixes the last completed analysis produced and which source answered
///  (our LSP session = incl. hints, or the Structure view = errors only).
///  Reads UI-thread state only - call it from a timer tick.</summary>
procedure LiveStatusInfo(out AFile: string; out AAnalysing, AResolving,
  AFromLsp, AFresh: Boolean; out AFixCount: Integer);

type
  /// <summary>Record describing one private member that H2219 flagged and
  ///  that we could remove: where it is declared, where it is implemented
  ///  and whether that implementation actually does anything.</summary>
  TPrivateMember = record
    Name: string;
    TypeName: string;
    DeclLine: Integer;        // 0-based, the declaration inside the class
    DeclText: string;         // exact text - stale check at apply time
    ImplFirst: Integer;       // 0-based, header line ( -1 = none: a field )
    ImplLast: Integer;        // 0-based, its final "end;"
    BodyEmpty: Boolean;       // no statements between begin and end
    BodyLines: Integer;       // statements the user would lose
  end;

/// <summary>Locates the private member AName of the type enclosing
///  ADeclLine plus its implementation. False when it cannot be removed
///  safely - a virtual/override/abstract/message method, an overload, a
///  name used elsewhere in the unit, or an implementation that cannot be
///  delimited. Pure and testable; the caller decides about the body.</summary>
function FindRemovablePrivate(const AContent: string; ADeclLine0: Integer;
  const AName: string; out AInfo: TPrivateMember): Boolean;

/// <summary>Content with AInfo's declaration AND implementation removed.</summary>
function RemovePrivateMemberText(const AContent: string;
  const AInfo: TPrivateMember): string;

type
  /// <summary>Asked before a non-empty body is deleted. Unassigned means
  ///  "no UI available" - the fix then REFUSES rather than deleting code
  ///  silently (that is also what keeps the console tests non-modal).</summary>
  TRemoveConfirmFunc = reference to function(const AInfo: TPrivateMember): Boolean;

var
  /// <summary>Set by the IDE / standalone startup to a themed dialog.</summary>
  RemovePrivateConfirm: TRemoveConfirmFunc = nil;

/// <summary>Why the LAST ResolveQuickFixes call did not turn a
///  diagnostic into a fix ('' when it did). Thread-local, so the worker
///  that ran the resolution reads its OWN reason right after the call.
///  "1 diag, 1 fixable -> 0 fix" is otherwise a dead end.</summary>
function LastResolveNote: string;

/// <summary>Bare member name out of an H2219 message. The compiler
///  quotes what it knows - "Test", "Test&lt;T&gt;" for a generic method, or the
///  qualified "TForm4.Test" - so the type prefix and the generic
///  parameter list are cut off. '' when nothing usable remains.</summary>
function H2219NameFromMessage(const AMessage: string): string;

/// <summary>False when a ';' can never be missing at the END of ALine:
///  the line closes with a block opener (begin/then/else/do/of/try/...),
///  an operator or a separator. A missing semicolon always follows a
///  STATEMENT. Guards against a stale diagnostic - after the offending
///  line is deleted the source still reports the old position for a
///  while, and the fix would then offer to insert ';' after 'begin'.</summary>
function CanTakeSemicolon(const ALine: string): Boolean;

/// <summary>True for diagnostic codes a quick-fix provider handles.
///  Everything else can never produce a fix (E2250 & co).</summary>
function IsHandledDiagCode(const ACode: string): Boolean;

/// <summary>Pipeline counters of the last completed resolution for the
///  status window: how many diagnostics the resolver SAW, how many of
///  them carry a code we have a provider for, and where they came from.
///  Turns a bare "0 fixes" into an answer to WHY.</summary>
procedure LiveDiagStats(out ASeen, AHandled: Integer; out ACodes: string);

/// <summary>Why the last resolution turned a FIXABLE diagnostic into no
///  fix at all ('' when it did produce one). See LastResolveNote.</summary>
function LiveDeclineNote: string;

/// <summary>One line per diagnostics source describing its LAST pass -
///  how many diagnostics it saw, how many fixes came out and whether the
///  result was published or dropped. This is what tells "the source never
///  ran" apart from "it ran and its result was discarded".</summary>
procedure LiveSourceStats(out AStructure, ALsp: string);

/// <summary>Class name of the popup that currently keeps the hint
///  hidden ('' when none) - the status window names it, so a wrong
///  suppression can be identified instead of guessed at.</summary>
function LiveHintBlocker: string;

/// <summary>Menu entry point: modal overview of EVERY quick fix in the
///  active unit (line-sorted, sortable columns). Selecting one jumps to
///  its line and opens the regular fix popup there. Waits briefly for a
///  fresh live analysis when none is available yet.</summary>
procedure ShowAllQuickFixes;

var
  /// <summary>Set by Expert.QuickFixMarkers: invoked from the poll TICK
  ///  (safe WM_TIMER context) whenever the published live results change,
  ///  so the markers can invalidate the editor. Never called from
  ///  notifier/queue contexts.</summary>
  GLiveRepaintHook: TProc = nil;

/// <summary>Called after a REAL compile (IOTAIDENotifier50.AfterCompile,
///  not code-insight): re-arms ONE LSP analysis of the active buffer even
///  though the Structure view already answered it. The compiler's full
///  message set - including hints like H2443, which never reach the
///  Structure view and whose window has no read API - is only obtainable
///  through the plugin's own LSP session. State-only (notifier context);
///  the poll tick performs the analysis.</summary>
procedure LiveRefreshAfterCompile;

implementation

uses
  System.Classes, System.Types, System.UITypes,
  System.Generics.Collections, System.Generics.Defaults,
  System.IOUtils, System.Math, System.StrUtils,
  System.Hash, System.JSON, System.SyncObjs,
  Winapi.Windows,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Graphics, Vcl.Dialogs,
  Lsp.Client, Lsp.Uri, Expert.LspManager,
  Expert.EditorHelperIntf, Expert.UnitAvailability,
  Expert.DfmEventCheck,   // MergeParamNames (shared with the DFM auto-fix)
  Expert.DialogHelper, Expert.IdeThemes, Expert.ListViewSort;

type
  TMissingIdent = record
    Identifier: string;         // bare name (compared to the word at cursor)
    GenericUse: Boolean;        // use site is "Name<...>" - display as "Name<>"
    Line: Integer;        // 0-based (from the diagnostic range)
    Section: TUsesSection;// smart default (interface if used above impl)
    Units: TArray<TFindUnitHit>;
  end;

// ---------------------------------------------------------------------------
//  Engine
// ---------------------------------------------------------------------------

function SplitContentLines(const AContent: string): TArray<string>;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

// 0-based line index of the 'implementation' keyword, or MaxInt if none.
function ImplementationLineOf(const ALines: TArray<string>): Integer;
var I: Integer;
begin
  for I := 0 to High(ALines) do
    if SameText(Trim(ALines[I]), 'implementation') then Exit(I);
  Result := MaxInt;
end;

// The same unit name is often found under several paths (3rd-party libs
// typically ship duplicate source dirs). For a uses clause only the NAME
// matters, so collapse the hits to one per unit name (first hit wins -
// Lookup lists the project scope first).
function DedupeByUnitName(const AHits: TArray<TFindUnitHit>): TArray<TFindUnitHit>;
var
  Seen: TDictionary<string, Boolean>;
  Res: TList<TFindUnitHit>;
begin
  if Length(AHits) <= 1 then Exit(AHits);
  Seen := TDictionary<string, Boolean>.Create;
  Res := TList<TFindUnitHit>.Create;
  try
    for var H in AHits do
      if not Seen.ContainsKey(UpperCase(H.UnitName)) then
      begin
        Seen.Add(UpperCase(H.UnitName), True);
        Res.Add(H);
      end;
    Result := Res.ToArray;
  finally
    Res.Free;
    Seen.Free;
  end;
end;

// Used to drop STALE add-unit diagnostics: our LSP (and Error Insight)
// can lag one edit behind, so right after the user added the unit the
// very same E2003 is reported again - without this guard the fix stuck
// to the now-correct buffer state forever.
function UnitReachableFrom(const AContent, AUnit: string;
  ASection: TUsesSection): Boolean;
begin
  Result := UnitInUsesSection(AContent, AUnit, usInterface)
    or ((ASection = usImplementation)
        and UnitInUsesSection(AContent, AUnit, usImplementation));
end;

function ReadCurrentContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := (Editor <> nil) and Editor.ReadEditorContent(AFile, AContent);
  if not Result then
  begin
    if not TFile.Exists(AFile) then Exit(False);
    try AContent := TFile.ReadAllText(AFile); Result := True; except Result := False; end;
  end;
end;

// ---- shared Pascal-header helpers (E2037 align + E2065 stub) --------------

function IsHeaderLine(const S: string; out AKind: string;
  out AIsClassMethod: Boolean): Boolean;
var
  T: string;
begin
  Result := False;
  AKind := '';
  T := Trim(S);
  AIsClassMethod := StartsText('class ', T);
  if AIsClassMethod then T := Trim(Copy(T, 7, MaxInt));
  for var KW in ['procedure', 'function', 'constructor', 'destructor'] do
    if StartsText(KW + ' ', T) then
    begin
      AKind := KW;
      Exit(True);
    end;
end;

// Joins lines from AStart until the first ';' outside parentheses;
// AEndLine returns the last joined line. '' when no terminator found.
function CollectHeader(const ALines: TArray<string>; AStart: Integer;
  out AEndLine: Integer): string;
var
  I, J, Depth: Integer;
  L: string;
begin
  Result := '';
  Depth := 0;
  for I := AStart to Min(AStart + 11, High(ALines)) do
  begin
    L := StripLineComment(ALines[I]);
    for J := 1 to Length(L) do
    begin
      case L[J] of
        '(', '[': Inc(Depth);
        ')', ']': if Depth > 0 then Dec(Depth);
        ';': if Depth = 0 then
          begin
            AEndLine := I;
            Exit(Trim(Result + ' ' + Copy(L, 1, J)));
          end;
      end;
    end;
    Result := Result + ' ' + L;
  end;
  Result := '';
end;

// Splits a collected header into name / params / return type.
function ParseHeader(const AHeader, AKind: string;
  out AQualified, AParams, ARetType: string): Boolean;
var
  T: string;
  P, Depth, I, OpenP, CloseP, ColonP: Integer;
begin
  Result := False;
  T := Trim(AHeader);
  if StartsText('class ', T) then T := Trim(Copy(T, 7, MaxInt));
  if not StartsText(AKind + ' ', T) then Exit;
  T := Trim(Copy(T, Length(AKind) + 2, MaxInt));
  // Name runs until '(' / ':' / ';'.
  P := 1;
  while (P <= Length(T)) and not CharInSet(T[P], ['(', ':', ';']) do Inc(P);
  AQualified := Trim(Copy(T, 1, P - 1));
  if AQualified = '' then Exit;
  AParams := '';
  ARetType := '';
  OpenP := 0; CloseP := 0; Depth := 0;
  for I := 1 to Length(T) do
    case T[I] of
      '(': begin if Depth = 0 then OpenP := I; Inc(Depth); end;
      ')': begin Dec(Depth); if Depth = 0 then begin CloseP := I; Break; end; end;
    end;
  if (OpenP > 0) and (CloseP > OpenP) then
    AParams := Trim(Copy(T, OpenP + 1, CloseP - OpenP - 1));
  // Return type: ':' after the params (or after the name) up to ';'.
  ColonP := 0; Depth := 0;
  for I := Max(1, CloseP + 1) to Length(T) do
    case T[I] of
      '(', '[': Inc(Depth);
      ')', ']': if Depth > 0 then Dec(Depth);
      ':': if Depth = 0 then begin ColonP := I; Break; end;
      ';': if Depth = 0 then Break;
    end;
  if ColonP > 0 then
  begin
    I := ColonP + 1;
    while (I <= Length(T)) and (T[I] <> ';') do Inc(I);
    ARetType := Trim(Copy(T, ColonP + 1, I - ColonP - 1));
  end;
  Result := True;
end;


// Turns the E2003 diagnostics for a buffer into missing-identifier records
// with candidate units. Pure function of (content, diagnostics) - safe on
// ANY thread (TUnitIndex.Lookup reads an immutable snapshot).
function ExtractMissing(const AContent: string;
  const ADiags: TArray<TLspErrorDiag>): TArray<TMissingIdent>;
var
  Lines: TArray<string>;
  Res: TList<TMissingIdent>;
  Seen: TDictionary<string, Boolean>;
  ImplLine: Integer;
begin
  Result := nil;
  Lines := SplitContentLines(AContent);
  ImplLine := ImplementationLineOf(Lines);

  Res := TList<TMissingIdent>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for var D in ADiags do
    begin
      if not SameText(D.Code, 'E2003') then Continue;   // undeclared identifier
      if (D.Range.Start.Line < 0) or (D.Range.Start.Line > High(Lines)) then Continue;
      // Message-first extraction (generics: message says 'TList<>', the
      // range points at/into the type arguments).
      var ECol0, ELen: Integer;
      var Ident := E2003IdentFromDiag(Lines, D, ECol0, ELen);
      if Ident = '' then Continue;
      if IdentDeclaredInFile(AContent, Ident) then Continue;   // EI false positive
      if Seen.ContainsKey(UpperCase(Ident)) then Continue;
      Seen.Add(UpperCase(Ident), True);

      var EGenericUse := IsGenericUseAt(Lines[D.Range.Start.Line], ECol0, ELen);
      var Hits := DedupeByUnitName(FilterHitsByGenericUse(
        TUnitIndex.Instance.Lookup(Ident), EGenericUse));
      if Length(Hits) = 0 then Continue;   // nothing we can add - skip

      var M: TMissingIdent;
      M.Identifier := Ident;
      M.GenericUse := EGenericUse;
      M.Line := D.Range.Start.Line;
      if M.Line < ImplLine then M.Section := usInterface else M.Section := usImplementation;
      // Stale diagnostic? Every candidate already reachable -> nothing to add.
      var Usable: TArray<TFindUnitHit> := nil;
      for var H in Hits do
        if not UnitReachableFrom(AContent, H.UnitName, M.Section) then
          Usable := Usable + [H];
      if Length(Usable) = 0 then Continue;
      Hits := Usable;
      M.Units := Hits;
      Res.Add(M);
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
    Seen.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Quick-fix model + resolver
// ---------------------------------------------------------------------------
//
// One diagnostic can yield several concrete FIX ACTIONS. Providers:
//   E2003 undeclared identifier  -> add declaring unit to uses (index hit),
//                                   or "did you mean" rename (fuzzy index)
//   F2613 unit not found         -> fix the unit name in uses (fuzzy) or
//                                   remove the entry
//   E2037 declaration differs    -> align the implementation header with
//                                   the declaration (DFM-fix style merge)
//   H2443 inline not expanded    -> add the unit named in the hint to uses
// ResolveQuickFixes is a pure function of (content, diagnostics) plus the
// immutable index snapshot - safe on any thread. (TQuickFix and the public
// entry points are declared in the interface section.)

// Token at the 1-based column ACol1 of ALine; identifier chars, optionally
// including '.' (for dotted unit names). 0-based start col out.
function TokenAt(const ALine: string; ACol1: Integer; AAllowDots: Boolean;
  out AStartCol0, ALen: Integer): Boolean;

  function IsTokCh(C: Char): Boolean;
  begin
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_'])
      or (AAllowDots and (C = '.'));
  end;

var
  P, StartP, EndP: Integer;
begin
  Result := False;
  P := ACol1;
  if (P > Length(ALine)) or ((P >= 1) and not IsTokCh(ALine[P])) then
    P := ACol1 - 1;
  if (P < 1) or (P > Length(ALine)) or not IsTokCh(ALine[P]) then Exit;
  StartP := P;
  while (StartP > 1) and IsTokCh(ALine[StartP - 1]) do Dec(StartP);
  EndP := P;
  while (EndP < Length(ALine)) and IsTokCh(ALine[EndP + 1]) do Inc(EndP);
  // Trim stray dots at the edges and reject number literals.
  while (StartP <= EndP) and (ALine[StartP] = '.') do Inc(StartP);
  while (EndP >= StartP) and (ALine[EndP] = '.') do Dec(EndP);
  if (EndP < StartP) or CharInSet(ALine[StartP], ['0'..'9']) then Exit;
  AStartCol0 := StartP - 1;
  ALen := EndP - StartP + 1;
  Result := True;
end;

// First '...'-quoted token of a (localized) diagnostic message - for H2443
// that is the inline function's name (the marker anchor).
function FirstQuoted(const S: string): string;
var
  I, StartQ: Integer;
begin
  Result := '';
  StartQ := 0;
  for I := 1 to Length(S) do
    if S[I] = '''' then
    begin
      if StartQ = 0 then
        StartQ := I
      else
      begin
        Result := Copy(S, StartQ + 1, I - StartQ - 1);
        Exit;
      end;
    end;
end;

// Last '...'-quoted token of a (localized) diagnostic message - for H2443
// that is the unit name ("... because unit 'System.Math' is not ...").
function LastQuoted(const S: string): string;
var
  I, EndQ: Integer;
begin
  Result := '';
  EndQ := 0;
  for I := Length(S) downto 1 do
    if S[I] = '''' then
    begin
      if EndQ = 0 then
        EndQ := I
      else
      begin
        Result := Copy(S, I + 1, EndQ - I - 1);
        Exit;
      end;
    end;
end;

function LooksLikeUnitName(const S: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  if (S = '') or CharInSet(S[1], ['0'..'9', '.']) then Exit;
  for I := 1 to Length(S) do
    if not CharInSet(S[I], ['A'..'Z', 'a'..'z', '0'..'9', '_', '.']) then Exit;
  Result := S[Length(S)] <> '.';
end;

function IsBareIdent(const S: string): Boolean;
var
  I: Integer;
begin
  Result := (S <> '') and not CharInSet(S[1], ['0'..'9']);
  if Result then
    for I := 1 to Length(S) do
      if not CharInSet(S[I], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
        Exit(False);
end;

function IdentDeclaredInFile(const AContent, AIdent: string): Boolean;
var
  Lines: TArray<string>;
  Raw, L, U, UIdent: string;
  P, AfterIdx, K: Integer;
begin
  Result := False;
  if (AContent = '') or (AIdent = '') then Exit;
  UIdent := UpperCase(AIdent);
  Lines := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  for Raw in Lines do
  begin
    L := Trim(Raw);
    if (L = '') or L.StartsWith('//') then Continue;
    P := Pos('//', L);
    if P > 0 then L := TrimRight(Copy(L, 1, P - 1));
    U := UpperCase(L);
    P := Pos(UIdent, U);
    while P > 0 do
    begin
      AfterIdx := P + Length(UIdent);
      // whole word?
      if ((P = 1) or not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_']))
        and ((AfterIdx > Length(U)) or not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_'])) then
      begin
        // 'X : Typ' (not 'X :=') - parameter, var, field.
        K := AfterIdx;
        while (K <= Length(U)) and CharInSet(U[K], [' ', #9]) do Inc(K);
        if (K <= Length(U)) and (U[K] = ':')
          and ((K = Length(U)) or (U[K + 1] <> '=')) then
          Exit(True);
        // 'var X' / 'const X' / 'out X' immediately before (covers
        // 'for var X in ...' and untyped inline vars).
        K := P - 1;
        while (K >= 1) and CharInSet(U[K], [' ', #9]) do Dec(K);
        for var W in ['VAR', 'CONST', 'OUT'] do
          if (K >= Length(W)) and (Copy(U, K - Length(W) + 1, Length(W)) = W)
            and ((K - Length(W) = 0) or not CharInSet(U[K - Length(W)], ['A'..'Z', '0'..'9', '_'])) then
            Exit(True);
      end;
      P := PosEx(UIdent, U, P + 1);
    end;
  end;
end;

function IsGenericUseAt(const ALine: string; ACol0, ALen: Integer): Boolean;
var
  P: Integer;
begin
  P := ACol0 + ALen + 1;   // 1-based char after the token
  while (P <= Length(ALine)) and CharInSet(ALine[P], [' ', #9]) do
    Inc(P);
  Result := (P <= Length(ALine)) and (ALine[P] = '<');
end;

function FilterHitsByGenericUse(const AHits: TArray<TFindUnitHit>;
  AGenericUse: Boolean): TArray<TFindUnitHit>;
var
  HasMatchingForm: Boolean;
  H: TFindUnitHit;
begin
  HasMatchingForm := False;
  for H in AHits do
    if H.IsGeneric = AGenericUse then
    begin
      HasMatchingForm := True;
      Break;
    end;
  if not HasMatchingForm then Exit(AHits);
  Result := nil;
  for H in AHits do
    if H.IsGeneric = AGenericUse then
      Result := Result + [H];
end;

function E2003IdentFromDiag(const ALines: TArray<string>;
  const D: TLspErrorDiag; out ACol0, ALen: Integer): string;
var
  LnTxt, TokTxt: string;
  P, LtP, TCol0, TLen: Integer;
begin
  Result := '';
  ACol0 := 0; ALen := 0;
  if (D.Range.Start.Line < 0) or (D.Range.Start.Line > High(ALines)) then Exit;
  LnTxt := ALines[D.Range.Start.Line];

  // Quoted name from the message; everything from '<' on is the generic
  // arity marker ("TList<>", "TDictionary<,>").
  Result := FirstQuoted(D.Message);
  LtP := Pos('<', Result);
  if LtP > 0 then Result := Copy(Result, 1, LtP - 1);
  if not IsBareIdent(Result) then Result := '';

  TokTxt := '';
  if TokenAt(LnTxt, D.Range.Start.Character + 1, False, TCol0, TLen) then
    TokTxt := Copy(LnTxt, TCol0 + 1, TLen);

  if Result = '' then
  begin
    // No usable quote - the range token is all we have.
    Result := TokTxt;
    ACol0 := TCol0;
    ALen := TLen;
    Exit;
  end;
  if SameText(TokTxt, Result) then
  begin
    ACol0 := TCol0;
    ALen := TLen;
    Exit;
  end;
  // Range and message disagree (generics): anchor at the first whole-word
  // occurrence of the reported name on the line.
  ALen := Length(Result);
  ACol0 := D.Range.Start.Character;   // last resort for the markers
  P := 1;
  while P <= Length(LnTxt) - ALen + 1 do
  begin
    P := PosEx(UpperCase(Result), UpperCase(LnTxt), P);
    if P = 0 then Break;
    if ((P = 1) or not CharInSet(LnTxt[P - 1], ['A'..'Z', 'a'..'z', '0'..'9', '_']))
       and ((P + ALen > Length(LnTxt))
         or not CharInSet(LnTxt[P + ALen], ['A'..'Z', 'a'..'z', '0'..'9', '_'])) then
    begin
      ACol0 := P - 1;
      Exit;
    end;
    Inc(P);
  end;
  // The reported name is NOT on the diagnosed line. DelphiLSP / Error
  // Insight lag one edit behind: when the user deletes the offending
  // line, the old E2003 gets re-published for the NEW buffer - keyed to
  // the new hash it then sticks for good ("the fix stays although I
  // removed the line"). A diagnostic whose identifier no longer occurs
  // ANYWHERE in the content cannot be acted on - treat it as stale.
  for var LI := 0 to High(ALines) do
  begin
    P := 1;
    while P <= Length(ALines[LI]) - ALen + 1 do
    begin
      P := PosEx(UpperCase(Result), UpperCase(ALines[LI]), P);
      if P = 0 then Break;
      if ((P = 1) or not CharInSet(ALines[LI][P - 1], ['A'..'Z', 'a'..'z', '0'..'9', '_']))
         and ((P + ALen > Length(ALines[LI]))
           or not CharInSet(ALines[LI][P + ALen], ['A'..'Z', 'a'..'z', '0'..'9', '_'])) then
        Exit;   // still used somewhere - keep the fix (anchor stays rough)
      Inc(P);
    end;
  end;
  Result := '';
  ACol0 := 0;
  ALen := 0;
end;

threadvar
  GResolveNote: string;

function LastResolveNote: string;
begin
  Result := GResolveNote;
end;

// Records why a diagnostic produced no fix (status window).
procedure Decline(const ACode, AIdent, AWhy: string);
var
  S: string;
begin
  S := ACode;
  if AIdent <> '' then S := S + ' "' + AIdent + '"';
  S := S + ': ' + AWhy;
  if GResolveNote = '' then GResolveNote := S
  else if Length(GResolveNote) < 300 then GResolveNote := GResolveNote + '; ' + S;
end;

function ResolveQuickFixes(const AContent: string;
  const ADiags: TArray<TLspErrorDiag>): TArray<TQuickFix>;
var
  Lines: TArray<string>;
  ImplLine: Integer;
  Res: TList<TQuickFix>;
  Seen: TDictionary<string, Boolean>;
  FuzzyCache: TDictionary<string, TArray<TFindUnitHit>>;
  Snap: IUnitSnapshot;

  function SectionFor(ALine0: Integer): TUsesSection;
  begin
    if ALine0 < ImplLine then Result := usInterface else Result := usImplementation;
  end;

  function DiagToken(const D: TLspErrorDiag; AAllowDots: Boolean;
    out ACol0, ALen: Integer; out AText: string): Boolean;
  var
    LnTxt: string;
  begin
    Result := False;
    if (D.Range.Start.Line < 0) or (D.Range.Start.Line > High(Lines)) then Exit;
    LnTxt := Lines[D.Range.Start.Line];
    Result := TokenAt(LnTxt, D.Range.Start.Character + 1, AAllowDots, ACol0, ALen);
    if Result then
      AText := Copy(LnTxt, ACol0 + 1, ALen);
  end;

  procedure AddE2003(const D: TLspErrorDiag);
  var
    Col0, Len, MaxDist: Integer;
    Ident: string;
    Hits, Cands: TArray<TFindUnitHit>;
    F: TQuickFix;
  begin
    // Message-first extraction: for "TList<Integer>" the compiler reports
    // 'TList<>' but the range points at/into the type arguments, so the
    // token there would be 'Integer' (or a type parameter T).
    Ident := E2003IdentFromDiag(Lines, D, Col0, Len);
    if Ident = '' then
    begin
      Decline('E2003', '', 'identifier not found in the buffer (stale diagnostic)');
      Exit;
    end;
    // Error-Insight false positive / lag: the identifier IS declared in
    // this very file (typical: an anonymous-method parameter) - neither
    // add-unit nor did-you-mean applies.
    if IdentDeclaredInFile(AContent, Ident) then
    begin
      Decline('E2003', Ident, 'declared in this file - Error Insight false positive');
      Exit;
    end;
    // Semantic dedup: both diagnostic sources report one E2003 per
    // OCCURRENCE, so the same identifier twice on a line would otherwise
    // produce identical fixes. One fix per (identifier, line).
    if Seen.ContainsKey('I|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident)) then Exit;
    Seen.Add('I|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident), True);

    if Snap = nil then
    begin
      Decline('E2003', Ident, 'identifier index not ready yet');
      Exit;
    end;
    // "TList<Integer>" must not offer System.Classes' non-generic TList
    // (and a bare "TList" not the generic one). Filter BEFORE the unit
    // dedupe so per-declaration genericity flags are still intact.
    var GenericUse := (D.Range.Start.Line >= 0)
      and (D.Range.Start.Line <= High(Lines))
      and IsGenericUseAt(Lines[D.Range.Start.Line], Col0, Len);
    Hits := DedupeByUnitName(FilterHitsByGenericUse(Snap.Lookup(Ident), GenericUse));
    var Sect := SectionFor(D.Range.Start.Line);
    if Length(Hits) > 0 then
    begin
      // STALE-DIAGNOSTIC GUARD (same as H2443): when every candidate unit
      // is already reachable from this section, the identifier resolves
      // and the E2003 is a leftover from a previous buffer state - offer
      // nothing (otherwise the fix stuck for good on an untouched buffer).
      var Usable: TArray<TFindUnitHit> := nil;
      for var H in Hits do
        if not UnitReachableFrom(AContent, H.UnitName, Sect) then
          Usable := Usable + [H];
      if Length(Usable) = 0 then
      begin
        Decline('E2003', Ident,
          'every candidate unit is already reachable (stale diagnostic)');
        Exit;
      end;
      Hits := Usable;
      // Display form: "TList<>" makes clear the GENERIC variant is meant.
      var Disp := Ident;
      if GenericUse then Disp := Disp + '<>';
      F := Default(TQuickFix);
      F.Kind := qfAddUnit;
      F.Line := D.Range.Start.Line;
      F.Col := Col0;           // for the editor line markers
      F.TokenLen := Len;
      F.Identifier := Disp;
      F.Caption := Format('Add unit for "%s"', [Disp]);
      F.Section := Sect;
      for var H in Hits do
        F.UnitNames := F.UnitNames + [H.UnitName];
      Res.Add(F);
      Exit;
    end;

    // Unknown everywhere -> "did you mean" via fuzzy index scan.
    if Length(Ident) < 4 then
    begin
      Decline('E2003', Ident, 'no unit declares it, too short for a "did you mean"');
      Exit;
    end;
    if not FuzzyCache.TryGetValue(UpperCase(Ident), Cands) then
    begin
      if Length(Ident) <= 5 then MaxDist := 1 else MaxDist := 2;
      Cands := Snap.FuzzyIdentifiers(Ident, MaxDist, 3);
      FuzzyCache.Add(UpperCase(Ident), Cands);
    end;
    if Length(Cands) = 0 then
      Decline('E2003', Ident,
        'no unit declares it and no similar identifier exists - nothing to offer');
    for var C in Cands do
    begin
      F := Default(TQuickFix);
      F.Kind := qfRenameIdent;
      F.Line := D.Range.Start.Line;
      F.Col := Col0;
      F.TokenLen := Len;
      F.NewText := C.Identifier;
      F.Caption := Format('Did you mean "%s"?', [C.Identifier]);
      F.Section := SectionFor(F.Line);
      // If no declaring unit of the corrected identifier is reachable yet,
      // plan to add the best one right after the rename.
      var CandHits := DedupeByUnitName(Snap.Lookup(C.Identifier));
      var Reachable := False;
      for var H in CandHits do
        if UnitInUsesText(AContent, H.UnitName) then
        begin
          Reachable := True;
          Break;
        end;
      if (not Reachable) and (Length(CandHits) > 0) then
        F.FollowUpUnit := CandHits[0].UnitName;
      Res.Add(F);
    end;
  end;

  procedure AddF2613(const D: TLspErrorDiag);
  var
    Col0, Len: Integer;
    UnitName, LnTxt, Entry: string;
    F: TQuickFix;
  begin
    if not DiagToken(D, True, Col0, Len, UnitName) then Exit;
    if Seen.ContainsKey('U|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(UnitName)) then Exit;
    Seen.Add('U|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(UnitName), True);

    // The clause entry may be syntactically BROKEN ('PP.;' - trailing dot):
    // TokenAt trims edge dots, but removal/replacement must cover the RAW
    // entry text up to the ','/';' delimiter, or the leftover dot keeps
    // the clause broken (and the remove action cannot even find it).
    LnTxt := Lines[D.Range.Start.Line];
    Entry := UnitName;
    var EndP := Col0 + Len;   // 0-based char AFTER the trimmed token
    while (EndP < Length(LnTxt))
      and not CharInSet(LnTxt[EndP + 1], [',', ';', ' ', #9, '/', '{']) do
      Inc(EndP);
    if EndP > Col0 + Len then
    begin
      Entry := Copy(LnTxt, Col0 + 1, EndP - Col0);
      Len := EndP - Col0;
    end;

    if Snap <> nil then
      for var Cand in Snap.FuzzyUnitNames(UnitName, 2, 3) do
      begin
        // A candidate that is ALREADY listed in a uses clause would end up
        // twice after the rename (a compile error) - the right action for
        // that case is the remove fix below.
        if UnitInUsesText(AContent, Cand) then Continue;
        F := Default(TQuickFix);
        F.Kind := qfFixUsesName;
        F.Line := D.Range.Start.Line;
        F.Col := Col0;
        F.TokenLen := Len;   // replaces the FULL broken entry
        F.NewText := Cand;
        F.Caption := Format('Unit "%s"?', [Cand]);
        Res.Add(F);
      end;
    F := Default(TQuickFix);
    F.Kind := qfRemoveUses;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.OldUnit := Entry;      // raw entry text, incl. a trailing dot
    F.Caption := Format('Remove "%s" from uses', [Entry]);
    Res.Add(F);
  end;

  procedure AddE2037(const D: TLspErrorDiag);
  var
    Col0, Len: Integer;
    Tok: string;
    F: TQuickFix;
  begin
    F := Default(TQuickFix);
    F.Kind := qfAlignHeader;
    F.Line := D.Range.Start.Line;
    F.Caption := 'Align with declaration';
    // Marker anchor: the method name at the diagnostic position (falls
    // back to the painter's first-word marker when unavailable).
    if DiagToken(D, False, Col0, Len, Tok) then
    begin
      F.Col := Col0;
      F.TokenLen := Len;
    end;
    Res.Add(F);
  end;

  procedure AddH2443(const D: TLspErrorDiag);
  var
    UnitName, FuncName, Tok: string;
    Col0, Len, P: Integer;
    F: TQuickFix;
  begin
    UnitName := LastQuoted(D.Message);
    if not LooksLikeUnitName(UnitName) then Exit;
    if UnitInUsesText(AContent, UnitName) then Exit;   // stale diagnostic
    // One add-unit fix per unit, no matter how many call sites hint it.
    if Seen.ContainsKey('H|' + UpperCase(UnitName)) then Exit;
    Seen.Add('H|' + UpperCase(UnitName), True);

    // Marker anchor: the inline function's call site. Prefer the token at
    // the diagnostic position; else locate the first-quoted name (the
    // function) in the line - without this the editor marker would sit
    // under the line's indentation instead of under the call.
    Col0 := 0;
    Len := 0;
    if not DiagToken(D, False, Col0, Len, Tok) then
    begin
      Col0 := 0;
      Len := 0;
      FuncName := FirstQuoted(D.Message);
      if (FuncName <> '') and (D.Range.Start.Line >= 0)
        and (D.Range.Start.Line <= High(Lines)) then
      begin
        P := Pos(UpperCase(FuncName), UpperCase(Lines[D.Range.Start.Line]));
        if P > 0 then
        begin
          Col0 := P - 1;
          Len := Length(FuncName);
        end;
      end;
    end;
    F := Default(TQuickFix);
    F.Kind := qfAddUnit;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := UnitName;
    F.UnitNames := [UnitName];
    F.Section := SectionFor(D.Range.Start.Line);
    F.Caption := Format('Add %s to uses', [UnitName]);
    Res.Add(F);
  end;

  procedure AddH2164(const D: TLspErrorDiag);
  var
    Col0, Len, ColonP, I: Integer;
    Ident, LnTxt: string;
    F: TQuickFix;
  begin
    if not DiagToken(D, False, Col0, Len, Ident) then Exit;
    if Seen.ContainsKey('V|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident)) then Exit;
    Seen.Add('V|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident), True);
    // Only offer the removal when the line plausibly IS a var declaration
    // we know how to rewrite: a ':' at top level before any ';', and no
    // ':=' (an initialized inline var may carry a side effect - skipped).
    LnTxt := Lines[D.Range.Start.Line];
    if Pos(':=', LnTxt) > 0 then Exit;
    ColonP := 0;
    for I := 1 to Length(LnTxt) do
      case LnTxt[I] of
        ':': begin ColonP := I; Break; end;
        ';': Break;
      end;
    if (ColonP = 0) or (Col0 + Len >= ColonP) then Exit;   // ident must precede ':'
    F := Default(TQuickFix);
    F.Kind := qfRemoveVar;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := Ident;
    F.Caption := Format('Remove unused variable "%s"', [Ident]);
    Res.Add(F);
  end;

  // E2029 "';' expected but X found" / E2066 "Missing operator or
  // semicolon": insert the ';' at the end of the statement it terminates.
  procedure AddInsertSemi(const D: TLspErrorDiag);
  var
    L0, C0, InsLine, InsCol, I: Integer;
    S, Prefix: string;
    F: TQuickFix;
  begin
    L0 := D.Range.Start.Line;
    C0 := D.Range.Start.Character;
    if (L0 < 0) or (L0 > High(Lines)) then Exit;

    // Text BEFORE the unexpected token on the same line? Then the ';'
    // belongs right after it; otherwise after the previous non-blank line.
    InsLine := L0;
    S := StripLineComment(Lines[L0]);
    Prefix := Copy(S, 1, Min(Max(C0, 0), Length(S)));
    if Trim(Prefix) = '' then
    begin
      I := L0 - 1;
      while (I >= 0) and (Trim(StripLineComment(Lines[I])) = '') do Dec(I);
      if I < 0 then
      begin
        Decline(D.Code, '', 'nothing before the reported position');
        Exit;
      end;
      InsLine := I;
      S := StripLineComment(Lines[I]);
      Prefix := S;
    end;
    InsCol := Length(Prefix);
    while (InsCol >= 1) and (Prefix[InsCol] = ' ') do Dec(InsCol);
    if InsCol < 1 then
    begin
      Decline(D.Code, '', 'insertion line is empty');
      Exit;
    end;
    if CharInSet(Prefix[InsCol], [';', ',']) then
    begin
      Decline(D.Code, '', 'the line already ends with ; or , (stale)');
      Exit;   // stale / odd spot
    end;
    // STALE-DIAGNOSTIC GUARD (tester: the fix survived deleting the line
    // it referred to). Error Insight / the compiler keep reporting the
    // old position for a moment; the line we would append to has shifted
    // and is now something like 'begin'. A ';' never follows a block
    // opener - so this cannot be a real missing semicolon.
    if not CanTakeSemicolon(Copy(Prefix, 1, InsCol)) then
    begin
      Decline(D.Code, '',
        'a ";" cannot follow "' + Trim(Copy(Prefix, 1, InsCol)) + '" (stale)');
      Exit;
    end;

    if Seen.ContainsKey('S|' + IntToStr(InsLine)) then Exit;
    Seen.Add('S|' + IntToStr(InsLine), True);

    F := Default(TQuickFix);
    F.Kind := qfInsertSemi;
    // Anchor (hint/marker/caret matching) at the DIAGNOSED line - that is
    // where the IDE points the user; the actual insertion position lives
    // in AuxLine/Col (usually the previous non-blank line).
    F.Line := D.Range.Start.Line;
    F.AuxLine := InsLine;
    F.Col := InsCol - 1;   // 0-based col of the last statement char
    F.TokenLen := 0;       // marker: first-word fallback on the diag line
    F.NewText := ';';
    F.Caption := 'Insert missing ";"';
    Res.Add(F);
  end;

  // E2029 "Declaration/Statement expected but identifier 'k' found": a
  // stray token typed outside any block - remove it. Language-independent
  // trigger: exactly ONE quoted part in the message (the ';'-expected
  // variant quotes TWO and is handled by AddInsertSemi), and that part is
  // a bare identifier verifiably present on the diagnosed line.
  procedure AddRemoveStray(const D: TLspErrorDiag);
  var
    L0, C0, QCnt, I, P, AfterIdx, FoundCol, Occurrences: Integer;
    Tok, S, U, UTok: string;
    F: TQuickFix;
  begin
    QCnt := 0;
    for I := 1 to Length(D.Message) do
      if D.Message[I] = '''' then Inc(QCnt);
    if QCnt <> 2 then Exit;
    Tok := FirstQuoted(D.Message);
    if not IsBareIdent(Tok) then Exit;

    L0 := D.Range.Start.Line;
    C0 := D.Range.Start.Character;
    if (L0 < 0) or (L0 > High(Lines)) then Exit;
    S := StripLineComment(Lines[L0]);
    U := UpperCase(S);
    UTok := UpperCase(Tok);

    // Whole-word occurrences on the line; the diag column picks one,
    // otherwise it must be unambiguous.
    FoundCol := -1;
    Occurrences := 0;
    P := Pos(UTok, U);
    while P > 0 do
    begin
      AfterIdx := P + Length(UTok);
      if ((P = 1) or not CharInSet(U[P - 1], ['A'..'Z', '0'..'9', '_']))
        and ((AfterIdx > Length(U)) or not CharInSet(U[AfterIdx], ['A'..'Z', '0'..'9', '_'])) then
      begin
        Inc(Occurrences);
        if (C0 >= P - 1) and (C0 < AfterIdx - 1) then FoundCol := P - 1
        else if FoundCol < 0 then FoundCol := P - 1;
      end;
      P := PosEx(UTok, U, P + 1);
    end;
    if Occurrences = 0 then Exit;
    if (Occurrences > 1)
      and not ((C0 >= 0) and (FoundCol >= 0) and (C0 >= FoundCol)
               and (C0 < FoundCol + Length(Tok))) then Exit;   // ambiguous

    if Seen.ContainsKey('RT|' + IntToStr(L0) + '|' + UTok) then Exit;
    Seen.Add('RT|' + IntToStr(L0) + '|' + UTok, True);

    F := Default(TQuickFix);
    F.Kind := qfRemoveToken;
    F.Line := L0;
    F.Col := FoundCol;
    F.TokenLen := Length(Tok);
    F.Identifier := Tok;
    // Apply-time validation: the applier requires the line's EXACT text.
    F.NewText := Lines[L0];
    F.Caption := Format('Remove stray "%s"', [Tok]);
    Res.Add(F);
  end;

  // H2219 "Private symbol 'X' declared but never used": remove it -
  // declaration AND implementation, or nothing at all.
  procedure AddH2219(const D: TLspErrorDiag);
  var
    Ident: string;
    Info: TPrivateMember;
    F: TQuickFix;
    Col0: Integer;
  begin
    Ident := H2219NameFromMessage(D.Message);
    if Ident = '' then
    begin
      Decline('H2219', FirstQuoted(D.Message),
        'no usable member name in the message');
      Exit;
    end;
    if not FindRemovablePrivate(AContent, D.Range.Start.Line, Ident, Info) then
    begin
      Decline('H2219', Ident,
        'not safely removable (contract directive, overload, still ' +
        'referenced, or the implementation cannot be delimited)');
      Exit;
    end;
    if Seen.ContainsKey('P|' + IntToStr(Info.DeclLine)) then Exit;
    Seen.Add('P|' + IntToStr(Info.DeclLine), True);

    F := Default(TQuickFix);
    F.Kind := qfRemovePrivate;
    F.Line := Info.DeclLine;
    F.Identifier := Ident;
    F.OldUnit := Info.TypeName;
    F.NewText := Info.DeclText;
    Col0 := Pos(UpperCase(Ident), UpperCase(Info.DeclText)) - 1;
    if Col0 < 0 then Col0 := 0;
    F.Col := Col0;
    F.TokenLen := Length(Ident);
    if Info.ImplFirst < 0 then
      F.Caption := Format('Remove unused private "%s"', [Ident])
    else if Info.BodyEmpty then
      F.Caption := Format('Remove unused private "%s" (declaration + empty implementation)',
        [Ident])
    else
      F.Caption := Format('Remove unused private "%s" (declaration + implementation, %d statement(s))',
        [Ident, Info.BodyLines]);
    Res.Add(F);
  end;

  // W1036 "Variable 'X' might not have been initialized": initialize it
  // with Default(<type>) right after the enclosing routine's 'begin'.
  procedure AddW1036(const D: TLspErrorDiag);
  var
    Col0, Len, HdrLine, BeginLine, I, ColonP: Integer;
    Ident, Kind, T, TypeText, Indent: string;
    IsCM, FoundName: Boolean;
    F: TQuickFix;
  begin
    if not DiagToken(D, False, Col0, Len, Ident) then Exit;
    if Seen.ContainsKey('W36|' + UpperCase(Ident) + '|' + IntToStr(D.Range.Start.Line)) then Exit;
    Seen.Add('W36|' + UpperCase(Ident) + '|' + IntToStr(D.Range.Start.Line), True);

    // Enclosing routine header: walking upward, every bare 'end;' we pass
    // closes a NESTED routine (or a try/case block) - skip one header per
    // counted terminator, otherwise a use below a nested routine would
    // anchor to the nested header and the fix would edit the wrong body.
    // Over-counting from try/case 'end;' lines only degrades to "no fix".
    HdrLine := -1;
    var EndCount := 0;
    for I := D.Range.Start.Line downto Max(0, D.Range.Start.Line - 400) do
    begin
      T := Trim(StripLineComment(Lines[I]));
      if SameText(T, 'end;') then
        Inc(EndCount)
      else if IsHeaderLine(Lines[I], Kind, IsCM) then
      begin
        if EndCount > 0 then
          Dec(EndCount)
        else
        begin
          HdrLine := I;
          Break;
        end;
      end;
    end;
    if HdrLine < 0 then Exit;

    BeginLine := -1;
    TypeText := '';
    for I := HdrLine + 1 to Min(High(Lines), HdrLine + 200) do
    begin
      T := Trim(StripLineComment(Lines[I]));
      // A NESTED routine header before the 'begin'? Its var block and its
      // 'begin' would be mistaken for the outer routine's - refuse the fix
      // for routines with nested subroutines rather than guessing.
      var K2: string;
      var CM2: Boolean;
      if IsHeaderLine(T, K2, CM2) then Exit;
      if SameText(T, 'begin') then
      begin
        BeginLine := I;
        Break;
      end;
      if Pos(':=', T) > 0 then Continue;   // inline var / statement - not a decl
      ColonP := Pos(':', T);
      if ColonP > 1 then
      begin
        FoundName := False;
        for var N in Copy(T, 1, ColonP - 1).Split([',']) do
          if SameText(Trim(N), Ident) then
          begin
            FoundName := True;
            Break;
          end;
        if FoundName then
        begin
          TypeText := Trim(Copy(T, ColonP + 1, MaxInt));
          if TypeText.EndsWith(';') then
            TypeText := Trim(Copy(TypeText, 1, Length(TypeText) - 1));
        end;
      end;
    end;
    if (BeginLine < 0) or (TypeText = '') then Exit;
    // Default(T) requires a plain (dotted) TYPE IDENTIFIER - reject inline
    // type constructions ('set of X', '^Integer', 'array of Y', 'record',
    // parameterized shapes) outright.
    if not CharInSet(TypeText[1], ['A'..'Z', 'a'..'z', '_']) then Exit;
    for I := 1 to Length(TypeText) do
      if not CharInSet(TypeText[I], ['A'..'Z', 'a'..'z', '0'..'9', '_', '.']) then Exit;

    Indent := Copy(Lines[BeginLine], 1,
      Length(Lines[BeginLine]) - Length(TrimLeft(Lines[BeginLine]))) + '  ';

    F := Default(TQuickFix);
    F.Kind := qfInitVar;
    F.Line := D.Range.Start.Line;   // hint/marker anchor: the flagged use
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := Ident;
    F.AuxLine := BeginLine;
    F.NewText := Indent + Ident + ' := Default(' + TypeText + ');';
    F.Caption := Format('Initialize "%s" at begin', [Ident]);
    Res.Add(F);
  end;

  // H2077 "Value assigned to 'X' never used": remove the dead assignment -
  // but ONLY when the whole statement sits on one line and the right-hand
  // side is a SAFE LITERAL. A blacklist ('no parentheses') is not enough
  // in Delphi: parameterless functions, constructors and property getters
  // are invoked WITHOUT parens ('X := Random;', 'L := TStringList.Create;'),
  // so any identifier RHS may carry a side effect. Whitelist instead:
  // numeric/string/char literals, nil, True/False, [].
  procedure AddH2077(const D: TLspErrorDiag);

    function IsSafeLiteral(const S: string): Boolean;
    var
      I: Integer;
      V: string;
    begin
      Result := False;
      V := Trim(S);
      if V = '' then Exit;
      if SameText(V, 'nil') or SameText(V, 'True') or SameText(V, 'False')
        or (V = '[]') or (V = '''''') then Exit(True);
      // String literal: fully quoted, no embedded expression.
      if (Length(V) >= 2) and (V[1] = '''') and (V[Length(V)] = '''') then
        Exit(True);
      // Char/numeric literal: #13, $FF, 3.14, -42.
      I := 1;
      if CharInSet(V[1], ['-', '+', '#', '$']) then I := 2;
      if I > Length(V) then Exit;
      for var K := I to Length(V) do
        if not CharInSet(V[K], ['0'..'9', '.', 'a'..'f', 'A'..'F']) then Exit;
      // Hex chars only allowed for #/$ prefixed literals.
      if not CharInSet(V[1], ['#', '$']) then
        for var K := I to Length(V) do
          if not CharInSet(V[K], ['0'..'9', '.']) then Exit;
      Result := True;
    end;

  var
    Col0, Len, P: Integer;
    Ident, T, RHS: string;
    F: TQuickFix;
  begin
    if not DiagToken(D, False, Col0, Len, Ident) then Exit;
    if Seen.ContainsKey('H77|' + IntToStr(D.Range.Start.Line)) then Exit;
    Seen.Add('H77|' + IntToStr(D.Range.Start.Line), True);

    T := Trim(StripLineComment(Lines[D.Range.Start.Line]));
    P := Pos(':=', T);
    if P <= 1 then Exit;
    if not SameText(Trim(Copy(T, 1, P - 1)), Ident) then Exit;   // plain "X := ..." only
    RHS := Trim(Copy(T, P + 2, MaxInt));
    if (RHS = '') or not RHS.EndsWith(';') then Exit;
    RHS := Trim(Copy(RHS, 1, Length(RHS) - 1));
    if not IsSafeLiteral(RHS) then Exit;

    F := Default(TQuickFix);
    F.Kind := qfRemoveAssign;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := Ident;
    F.NewText := T;   // exact statement - the applier refuses on any drift
    F.Caption := Format('Remove dead assignment to "%s"', [Ident]);
    Res.Add(F);
  end;

  // W1010 "Method hides virtual method of base type": add 'reintroduce;'.
  procedure AddW1010(const D: TLspErrorDiag);
  var
    Col0, Len: Integer;
    Ident, Kind, S: string;
    IsCM: Boolean;
    F: TQuickFix;
  begin
    if not DiagToken(D, False, Col0, Len, Ident) then Exit;
    if Seen.ContainsKey('W10|' + IntToStr(D.Range.Start.Line)) then Exit;
    Seen.Add('W10|' + IntToStr(D.Range.Start.Line), True);

    S := StripLineComment(Lines[D.Range.Start.Line]);
    if not IsHeaderLine(S, Kind, IsCM) then Exit;
    if Pos(';', S) = 0 then Exit;                       // multi-line decl - skip
    if Pos('REINTRODUCE', UpperCase(S)) > 0 then Exit;  // already there

    F := Default(TQuickFix);
    F.Kind := qfAddReintroduce;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := Ident;
    F.Caption := 'Add reintroduce';
    Res.Add(F);
  end;

  // E2065 "Unsatisfied forward or external declaration": generate an empty
  // implementation - or, for a forward CLASS declaration ("TNode = class;")
  // that never got its full declaration, a class stub.
  procedure AddE2065(const D: TLspErrorDiag);
  var
    Col0, Len: Integer;
    Ident, T, Rest, Kind: string;
    IsCM: Boolean;
    F: TQuickFix;
  begin
    if not DiagToken(D, False, Col0, Len, Ident) then Exit;
    if Seen.ContainsKey('E65|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident)) then Exit;
    Seen.Add('E65|' + IntToStr(D.Range.Start.Line) + '|' + UpperCase(Ident), True);

    T := Trim(StripLineComment(Lines[D.Range.Start.Line]));

    // Forward class declaration: "<Ident> = class;"
    if StartsText(Ident, T) then
    begin
      Rest := Trim(Copy(T, Length(Ident) + 1, MaxInt));
      if StartsText('=', Rest) and SameText(Trim(Copy(Rest, 2, MaxInt)), 'class;') then
      begin
        F := Default(TQuickFix);
        F.Kind := qfClassStub;
        F.Line := D.Range.Start.Line;
        F.Col := Col0;
        F.TokenLen := Len;
        F.Identifier := Ident;
        F.Caption := Format('Declare class "%s" (stub)', [Ident]);
        Res.Add(F);
        Exit;
      end;
    end;

    // Method / routine declaration without implementation.
    if not IsHeaderLine(T, Kind, IsCM) then Exit;
    if Pos('EXTERNAL', UpperCase(T)) > 0 then Exit;   // cannot generate that

    F := Default(TQuickFix);
    F.Kind := qfImplStub;
    F.Line := D.Range.Start.Line;
    F.Col := Col0;
    F.TokenLen := Len;
    F.Identifier := Ident;
    F.Caption := Format('Create empty implementation of "%s"', [Ident]);
    Res.Add(F);
  end;

var
  Key: string;
begin
  Result := nil;
  Lines := SplitContentLines(AContent);
  ImplLine := ImplementationLineOf(Lines);
  Snap := TUnitIndex.Instance.Snapshot;

  GResolveNote := '';
  Res := TList<TQuickFix>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  FuzzyCache := TDictionary<string, TArray<TFindUnitHit>>.Create;
  try
    for var D in ADiags do
    begin
      Key := Format('%s|%d|%d', [UpperCase(D.Code), D.Range.Start.Line,
        D.Range.Start.Character]);
      if Seen.ContainsKey(Key) then Continue;
      Seen.Add(Key, True);

      // F2613 "Unit not found" and F2063 "Could not compile used unit":
      // the IDE's Error Insight reports a broken/misspelled uses entry as
      // F2063 (observed empirically), the batch compiler as F2613 - both
      // anchor at the uses entry and get the same repair actions.
      if SameText(D.Code, 'E2003') then AddE2003(D)
      else if SameText(D.Code, 'F2613') or SameText(D.Code, 'F2063') then AddF2613(D)
      else if SameText(D.Code, 'E2037') then AddE2037(D)
      else if SameText(D.Code, 'H2443') then AddH2443(D)
      else if SameText(D.Code, 'H2164') then AddH2164(D)
      else if SameText(D.Code, 'H2219') then AddH2219(D)
      // E2029 only when the EXPECTED token is ';' (first quoted part of the
      // localized message); E2066 is always the missing-semicolon case.
      else if (SameText(D.Code, 'E2029') and (FirstQuoted(D.Message) = ';'))
        or SameText(D.Code, 'E2066') then AddInsertSemi(D)
      else if SameText(D.Code, 'E2029') then AddRemoveStray(D)
      else if SameText(D.Code, 'W1036') then AddW1036(D)
      else if SameText(D.Code, 'H2077') then AddH2077(D)
      else if SameText(D.Code, 'W1010') then AddW1010(D)
      else if SameText(D.Code, 'E2065') then AddE2065(D);
    end;
    Result := Res.ToArray;
  finally
    FuzzyCache.Free;
    Seen.Free;
    Res.Free;
  end;
end;

var
  // True while an on-demand gather is running; the live poller skips its
  // tick then, so two didClose/didOpen sequences never interleave.
  GOnDemandBusy: Boolean = False;

// Runs DelphiLSP over AFile, collects the undeclared-identifier (E2003)
// diagnostics, resolves each identifier via the unit index and returns the
// ones that have at least one candidate unit. AStatus reports progress.
// MAIN THREAD ONLY (talks to the editor and pumps messages).
function GatherMissing(const AFile: string; AStatus: TProc<string>): TArray<TMissingIdent>;
var
  Root, Proj, Json, Content: string;
  Client: TLspClient;
  Diags: TArray<TLspErrorDiag>;
  Before, I: Integer;
begin
  Result := nil;
  if (Editor = nil) or (AFile = '') then Exit;
  Json := Editor.FindDelphiLspJson;
  if Json = '' then
  begin
    if Assigned(AStatus) then AStatus('DelphiLSP not configured for this project.');
    Exit;
  end;
  Root := Editor.GetProjectRoot;
  if Root = '' then Root := ExtractFilePath(AFile);
  Proj := Editor.GetCurrentProjectDproj;

  GOnDemandBusy := True;
  try
    // Keep the index warm and the file synced to disk (so LSP sees the buffer).
    TUnitIndex.Instance.RefreshSourcesFromEditor;
    Editor.SaveFile(AFile);

    if Assigned(AStatus) then AStatus('Starting LSP...');
    Client := TLspManager.Instance.GetClient(Root, Proj, Json);
    if Client = nil then Exit;

    // Force a FRESH analysis (RefreshDocument sends a didChange; documentSymbol
    // triggers parsing) and wait for the resulting diagnostics push.
    Before := Client.GetDiagnosticsCount;
    Client.RefreshDocument(AFile);
    try Client.GetDocumentSymbols(AFile, 15000).Free; except end;
    if Assigned(AStatus) then AStatus('Analysing (waiting for diagnostics)...');
    for I := 1 to 40 do
    begin
      if Client.GetDiagnosticsCount > Before then Break;
      Sleep(150);
      Application.ProcessMessages;
    end;
    Sleep(300);   // let this file's specific push settle
    Diags := Client.GetErrorDiagnostics(AFile);
  finally
    GOnDemandBusy := False;
  end;

  if not ReadCurrentContent(AFile, Content) then Exit;
  Result := ExtractMissing(Content, Diags);
end;

function SectionName(ASection: TUsesSection): string;
begin
  if ASection = usInterface then Result := 'interface' else Result := 'implementation';
end;

function ApplyOne(const AFile, AUnit: string; ASection: TUsesSection): Boolean;
begin
  // Units the index only found via a BROWSING path are invisible to the
  // compiler - offer to make them project-visible first.
  if not EnsureUnitAvailable(AUnit) then Exit(False);
  Result := AddUnitToUses(AFile, AUnit, ASection);
end;

// ---------------------------------------------------------------------------
//  E2037 fix: align an implementation header with its declaration
// ---------------------------------------------------------------------------

// Rewrites the implementation header at/around ALine0 so that parameter
// modifiers/types and the return type match the method's declaration, while
// KEEPING the implementation's parameter names (MergeParamNames - the same
// merge the DFM auto-fix uses). Multi-line headers are collapsed into one
// rewritten line. Overloads with ambiguous declarations are refused.
function AlignImplHeaderToDecl(const AFile: string; ALine0: Integer): Boolean;
var
  Content: string;
  Lines: TArray<string>;
  HdrStart, HdrEnd, I, P, DeclLine, DeclEnd, ClassLine, SearchFrom, SearchTo: Integer;
  Kind, DeclKind, Header, Qualified, ImplParams, ImplRet: string;
  ClassName, MethodName, DeclHeader, DeclQual, DeclParams, DeclRet: string;
  IsClassMeth, DeclIsClassMeth, B: Boolean;
  Matches: Integer;
  NewHeader, Indent: string;
begin
  Result := False;
  if not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;

  // 1. Header start at/above the diagnostic line.
  HdrStart := -1;
  for I := ALine0 downto Max(0, ALine0 - 4) do
    if IsHeaderLine(Lines[I], Kind, IsClassMeth) then
    begin
      HdrStart := I;
      Break;
    end;
  if HdrStart < 0 then Exit;

  Header := CollectHeader(Lines, HdrStart, HdrEnd);
  if Header = '' then Exit;
  if not ParseHeader(Header, Kind, Qualified, ImplParams, ImplRet) then Exit;

  P := Qualified.LastDelimiter('.');   // 0-based, -1 if none
  if P >= 0 then
  begin
    ClassName := Copy(Qualified, 1, P);
    MethodName := Copy(Qualified, P + 2, MaxInt);
  end
  else
  begin
    ClassName := '';
    MethodName := Qualified;
  end;
  if MethodName = '' then Exit;

  // 2. Find the declaration: inside the class body, or (plain routines)
  //    in the interface section.
  SearchFrom := 0;
  SearchTo := High(Lines);
  if ClassName <> '' then
  begin
    ClassLine := -1;
    for I := 0 to High(Lines) do
    begin
      var T := Trim(StripLineComment(Lines[I]));
      if StartsText(ClassName, T) then
      begin
        var Rest := Trim(Copy(T, Length(ClassName) + 1, MaxInt));
        if StartsText('=', Rest) and (Pos('CLASS', UpperCase(Rest)) > 0) then
        begin
          // Skip non-body declarations that also match the pattern:
          // forward ('TNode = class;') and metaclass ('TFoo = class of X;').
          var After := Trim(Copy(Rest, 2, MaxInt));   // text after '='
          if SameText(After, 'class;') or StartsText('class of ', After) then
            Continue;
          ClassLine := I;
          Break;
        end;
      end;
    end;
    if ClassLine < 0 then Exit;
    SearchFrom := ClassLine + 1;
    SearchTo := High(Lines);
    // The class body ends at ITS 'end;' - track nested type declarations
    // ('TOpts = record ... end;', nested classes) so their terminators do
    // not cut the search region short.
    var Depth := 1;
    for I := ClassLine + 1 to High(Lines) do
    begin
      var T := Trim(StripLineComment(Lines[I]));
      var U := UpperCase(T);
      if ((Pos('= RECORD', U) > 0) or (Pos('=RECORD', U) > 0)
          or (Pos('= CLASS', U) > 0) or (Pos('=CLASS', U) > 0))
        and not U.EndsWith(';') then
        Inc(Depth)
      else if SameText(T, 'end;') then
      begin
        Dec(Depth);
        if Depth = 0 then
        begin
          SearchTo := I;
          Break;
        end;
      end;
    end;
  end
  else
    SearchTo := ImplementationLineOf(Lines) - 1;

  DeclLine := -1;
  Matches := 0;
  for I := SearchFrom to Min(SearchTo, High(Lines)) do
  begin
    if I = HdrStart then Continue;
    if not IsHeaderLine(Lines[I], DeclKind, B) then Continue;
    var T := Trim(StripLineComment(Lines[I]));
    if StartsText('class ', T) then T := Trim(Copy(T, 7, MaxInt));
    T := Trim(Copy(T, Length(DeclKind) + 2, MaxInt));
    if not StartsText(MethodName, T) then Continue;
    var After := Copy(T, Length(MethodName) + 1, MaxInt);
    if (After <> '') and not CharInSet(After[1], ['(', ':', ';', ' ']) then Continue;
    Inc(Matches);
    if DeclLine < 0 then DeclLine := I;
  end;
  if (DeclLine < 0) or (Matches > 1) then Exit;   // none, or ambiguous overloads

  DeclHeader := CollectHeader(Lines, DeclLine, DeclEnd);
  if DeclHeader = '' then Exit;
  if not IsHeaderLine(Lines[DeclLine], DeclKind, DeclIsClassMeth) then Exit;
  if not ParseHeader(DeclHeader, DeclKind, DeclQual, DeclParams, DeclRet) then Exit;

  // 3. Compose the new implementation header: declaration's modifiers /
  //    types / return type, implementation's parameter names.
  Indent := Copy(Lines[HdrStart], 1,
    Length(Lines[HdrStart]) - Length(TrimLeft(Lines[HdrStart])));
  NewHeader := Indent;
  if DeclIsClassMeth then NewHeader := NewHeader + 'class ';
  NewHeader := NewHeader + DeclKind + ' ' + Qualified;
  if DeclParams <> '' then
    NewHeader := NewHeader + '(' + MergeParamNames(ImplParams, DeclParams) + ')';
  if DeclRet <> '' then
    NewHeader := NewHeader + ': ' + DeclRet;
  NewHeader := NewHeader + ';';

  // 4. Apply: first line replaced, continuation lines removed.
  if not Editor.ReplaceLineAt(AFile, HdrStart + 1, NewHeader) then Exit;
  for I := HdrEnd downto HdrStart + 1 do
    Editor.DeleteLineAt(AFile, I + 1);
  Result := True;
end;

// Replaces the CHARACTERS [ACol0, ACol0+ATokenLen) of the 0-based line
// ALine0 with ANewText, via a whole-line rewrite. Deliberately NOT
// Editor.ReplaceSelection: the IDE implementation of that computes BYTE
// offsets in the UTF-8 buffer, so a non-ASCII character earlier on the
// line (a German comment, say) would shift the replacement. ReplaceLineAt
// is line-based and encoding-safe in both hosts.
function ReplaceTokenInLine(const AFile: string; ALine0, ACol0,
  ATokenLen: Integer; const ANewText: string): Boolean;
var
  Content, L: string;
  Lines: TArray<string>;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  L := Lines[ALine0];
  // Stale-buffer guard: the token must still fit where the fix expects it.
  if ACol0 + ATokenLen > Length(L) then Exit;
  Result := Editor.ReplaceLineAt(AFile, ALine0 + 1,
    Copy(L, 1, ACol0) + ANewText + Copy(L, ACol0 + ATokenLen + 1, MaxInt));
end;

// H2164 fix: removes the variable AIdent from the declaration line ALine0
// ("A, X, B: Integer;" -> "A, B: Integer;"; a lone "X: Integer;" removes
// the line, and a var block left empty loses its 'var' keyword line too).
function RemoveVarFromDecl(const AFile, AIdent: string; ALine0: Integer): Boolean;

  function LooksLikeDeclLine(const S: string): Boolean;
  var
    T: string;
    P: Integer;
  begin
    // "name[, name]*: type;" - enough to recognise a sibling declaration.
    T := Trim(StripLineComment(S));
    Result := False;
    if (T = '') or (Pos(':=', T) > 0) then Exit;
    P := Pos(':', T);
    if P <= 1 then Exit;
    Result := CharInSet(T[1], ['A'..'Z', 'a'..'z', '_']);
  end;

var
  Content, L, Stripped, NamesPart, Rest, Rebuilt: string;
  Lines: TArray<string>;
  ColonP, PrevIdx, NextIdx: Integer;
  Kept: TArray<string>;
  HadVarPrefix, Found: Boolean;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  L := Lines[ALine0];
  Stripped := StripLineComment(L);
  if Pos(':=', Stripped) > 0 then Exit;   // initialized inline var - refuse

  ColonP := Pos(':', Stripped);
  if ColonP <= 1 then Exit;
  NamesPart := Trim(Copy(Stripped, 1, ColonP - 1));
  Rest := Copy(L, ColonP, MaxInt);        // ': type;' + any trailing comment

  HadVarPrefix := StartsText('var ', NamesPart);
  if HadVarPrefix then
    NamesPart := Trim(Copy(NamesPart, 5, MaxInt));

  Found := False;
  Kept := nil;
  for var N in NamesPart.Split([',']) do
  begin
    var Nm := Trim(N);
    if Nm = '' then Continue;
    if SameText(Nm, AIdent) then
      Found := True
    else
      Kept := Kept + [Nm];
  end;
  if not Found then Exit;

  if Length(Kept) > 0 then
  begin
    // Rebuild the line without the removed name.
    Rebuilt := Copy(L, 1, Length(L) - Length(TrimLeft(L)));   // indentation
    if HadVarPrefix then Rebuilt := Rebuilt + 'var ';
    Rebuilt := Rebuilt + string.Join(', ', Kept) + Rest;
    Exit(Editor.ReplaceLineAt(AFile, ALine0 + 1, Rebuilt));
  end;

  // The only name on the line -> remove the whole line ...
  if not Editor.DeleteLineAt(AFile, ALine0 + 1) then Exit;
  Result := True;
  if HadVarPrefix then Exit;   // inline 'var X: T;' - nothing more to do

  // ... and drop a now-empty 'var' keyword line above it: previous
  // non-empty line must be exactly 'var', and the line that moved into
  // the deleted slot must not be a sibling declaration.
  PrevIdx := ALine0 - 1;
  while (PrevIdx >= 0) and (Trim(Lines[PrevIdx]) = '') do Dec(PrevIdx);
  if (PrevIdx < 0) or not SameText(Trim(StripLineComment(Lines[PrevIdx])), 'var') then Exit;
  NextIdx := ALine0 + 1;   // in the ORIGINAL lines: the line after the deleted one
  while (NextIdx <= High(Lines)) and (Trim(Lines[NextIdx]) = '') do Inc(NextIdx);
  if (NextIdx <= High(Lines)) and LooksLikeDeclLine(Lines[NextIdx]) then Exit;
  Editor.DeleteLineAt(AFile, PrevIdx + 1);
end;

// "<Name> = class..." / "<Name> = record..." opener that starts a type BODY
// (forward declarations, metaclasses and one-line aliases excluded).
// Handles 'class abstract'/'class sealed'/'class helper' and generic names
// ('TFoo<T>' - the generic signature stays part of the returned name, which
// is exactly what the qualified implementation header needs).
// Returns the type name, or ''.
function ClassOpenerName(const ATrimmed: string): string;
var
  EqP: Integer;
  Name, Rest: string;
begin
  Result := '';
  EqP := Pos('=', ATrimmed);
  if EqP <= 1 then Exit;
  Name := Trim(Copy(ATrimmed, 1, EqP - 1));
  if (Name = '') or not CharInSet(Name[1], ['A'..'Z', 'a'..'z', '_']) then Exit;
  for var I := 1 to Length(Name) do
    if not CharInSet(Name[I], ['A'..'Z', 'a'..'z', '0'..'9', '_', '<', '>', ',', ' ', ':']) then Exit;
  Rest := Trim(Copy(ATrimmed, EqP + 1, MaxInt));
  if not (StartsText('class', Rest) or StartsText('record', Rest)) then Exit;
  var After := '';
  if StartsText('class', Rest) then After := Trim(Copy(Rest, 6, MaxInt))
  else After := Trim(Copy(Rest, 7, MaxInt));
  if (After <> '') and CharInSet(After[1], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
  begin
    if StartsText('of ', After) then Exit;   // metaclass '= class of X'
    if not (StartsText('helper', After) or StartsText('abstract', After)
      or StartsText('sealed', After)) then Exit;
  end;
  if Rest.EndsWith(';') then Exit;   // forward decl or one-line alias
  Result := Name;
end;

// The (possibly nested) class/record whose declaration body encloses
// ALine0, FULLY QUALIFIED ('TOuter.TInner'); '' = none. Walks upward
// counting 'end;' terminators against type-body openers, then keeps
// walking to pick up enclosing containers of the container.
function EnclosingContainerName(const ALines: TArray<string>; ALine0: Integer): string;
var
  I, Depth: Integer;
  T, Name: string;
begin
  Result := '';
  Depth := 0;
  for I := ALine0 - 1 downto 0 do
  begin
    T := Trim(StripLineComment(ALines[I]));
    if SameText(T, 'end;') then
      Inc(Depth)
    else
    begin
      Name := ClassOpenerName(T);
      if Name <> '' then
      begin
        if Depth > 0 then
          Dec(Depth)
        else
        begin
          // Found the innermost container; prepend any outer ones.
          if Result = '' then
            Result := Name
          else
            Result := Name + '.' + Result;
          // continue the walk for a possible enclosing type
        end;
      end
      else if SameText(T, 'implementation') or SameText(T, 'interface') then
        Exit;
    end;
  end;
end;

// The 0-based line an implementation block should be inserted BEFORE:
// the 'initialization'/'finalization' keyword when the unit has one
// (inserting before the final 'end.' would land INSIDE that statement
// block), else the final 'end.'. -1 when the file has no implementation
// section at all (a .dpr program - its 'end.' closes the main block).
function ImplInsertLine(const ALines: TArray<string>): Integer;
var
  I, ImplLine: Integer;
  T: string;
begin
  Result := -1;
  ImplLine := ImplementationLineOf(ALines);
  if ImplLine = MaxInt then Exit;   // program file - refuse
  for I := ImplLine + 1 to High(ALines) do
  begin
    T := LowerCase(Trim(StripLineComment(ALines[I])));
    if (T = 'initialization') or (T = 'finalization') then
      Exit(I);
  end;
  for I := High(ALines) downto 0 do
    if SameText(Trim(ALines[I]), 'end.') then
      Exit(I);
end;

// E2065 fix (methods/routines): appends an empty implementation block for
// the declaration at ALine0 at the end of the implementation section.
function CreateImplStub(const AFile: string; ALine0: Integer): Boolean;
var
  Content, Kind, Hdr, Qualified, BareName, Params, Ret, ContainerName, Text: string;
  Lines: TArray<string>;
  HdrEnd, InsLine, I: Integer;
  IsCM, DummyCM: Boolean;
  ChkKind, ChkQual, ChkParams, ChkRet: string;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  if not IsHeaderLine(Lines[ALine0], Kind, IsCM) then Exit;
  Hdr := CollectHeader(Lines, ALine0, HdrEnd);
  if Hdr = '' then Exit;
  if not ParseHeader(Hdr, Kind, Qualified, Params, Ret) then Exit;
  if Pos('.', Qualified) > 0 then Exit;   // already an implementation header
  BareName := Qualified;

  ContainerName := EnclosingContainerName(Lines, ALine0);
  // An INDENTED declaration with no detected container means we failed to
  // recognise the enclosing type (exotic opener) - generating an
  // unqualified free routine would just add junk. Refuse instead.
  if (ContainerName = '') and (Length(Lines[ALine0]) > 0)
    and (Lines[ALine0][1] = ' ') then Exit;
  if ContainerName <> '' then
    Qualified := ContainerName + '.' + Qualified;

  InsLine := ImplInsertLine(Lines);
  if InsLine < 0 then Exit;   // also guards the loop below (impl exists)

  // Idempotence: refuse when a matching implementation header already
  // exists (stale diagnostics / applying the same fix twice).
  for I := ImplementationLineOf(Lines) + 1 to High(Lines) do
    if IsHeaderLine(Lines[I], ChkKind, DummyCM) then
    begin
      Hdr := CollectHeader(Lines, I, HdrEnd);
      if (Hdr <> '') and ParseHeader(Hdr, ChkKind, ChkQual, ChkParams, ChkRet)
        and (SameText(ChkQual, Qualified)
          or SameText(ChkQual, ContainerName + '.' + BareName)) then Exit;
    end;

  Text := '';
  if IsCM then Text := 'class ';
  Text := Text + Kind + ' ' + Qualified;
  if Params <> '' then Text := Text + '(' + Params + ')';
  if Ret <> '' then Text := Text + ': ' + Ret;
  Text := Text + ';' + sLineBreak + 'begin' + sLineBreak + sLineBreak
    + 'end;' + sLineBreak + sLineBreak;

  Result := Editor.InsertTextAtLineStart(AFile, InsLine + 1, Text);
end;

// E2065 fix (forward class): inserts a minimal class declaration right
// after the forward declaration line.
function CreateClassStub(const AFile, AClassName: string; ALine0: Integer): Boolean;
var
  Content, T, Indent, Text: string;
  Lines: TArray<string>;
  I: Integer;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 + 1 > High(Lines)) then Exit;
  T := Trim(StripLineComment(Lines[ALine0]));
  if not StartsText(AClassName, T) then Exit;   // stale buffer
  // Idempotence: refuse when a full (non-forward) declaration of the
  // class already exists anywhere in the file.
  for I := 0 to High(Lines) do
    if (I <> ALine0)
      and SameText(ClassOpenerName(Trim(StripLineComment(Lines[I]))), AClassName) then
      Exit;
  Indent := Copy(Lines[ALine0], 1,
    Length(Lines[ALine0]) - Length(TrimLeft(Lines[ALine0])));
  Text := sLineBreak + Indent + AClassName + ' = class(TObject)' + sLineBreak
    + Indent + 'end;' + sLineBreak;
  Result := Editor.InsertTextAtLineStart(AFile, ALine0 + 2, Text);
end;

// W1036 fix: inserts the Default() initialization right after 'begin'.
function InsertInitAtBegin(const AFile, AStmtLine: string;
  ABeginLine0: Integer): Boolean;
var
  Content: string;
  Lines: TArray<string>;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ABeginLine0 < 0) or (ABeginLine0 >= High(Lines)) then Exit;
  if not SameText(Trim(StripLineComment(Lines[ABeginLine0])), 'begin') then Exit;
  Result := Editor.InsertTextAtLineStart(AFile, ABeginLine0 + 2,
    AStmtLine + sLineBreak);
end;

// E2029/E2066 fix: inserts ';' right after the 0-based character index
// ACol0 of line ALine0 - with the resolver's semantic guards REPEATED
// against the current buffer, so a stale fix can never punch a ';' into
// an arbitrary position.
function InsertSemicolonAfter(const AFile: string; ALine0, ACol0: Integer): Boolean;
var
  Content, L: string;
  Lines: TArray<string>;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  L := Lines[ALine0];
  if (ACol0 < 0) or (ACol0 + 1 > Length(L)) then Exit;
  // The marked character must still be a statement-end character (not
  // whitespace / already terminated) and nothing may already follow it.
  if CharInSet(L[ACol0 + 1], [' ', #9, ';', ',']) then Exit;
  if (ACol0 + 2 <= Length(L)) and (L[ACol0 + 2] = ';') then Exit;
  Result := Editor.ReplaceLineAt(AFile, ALine0 + 1,
    Copy(L, 1, ACol0 + 1) + ';' + Copy(L, ACol0 + 2, MaxInt));
end;

// H2077 fix: deletes the dead assignment line - only when the line still
// contains the EXACT statement the fix was resolved for (AExpected), so a
// stale/shifted buffer can never lose a live assignment.
// E2029 stray-token fix: removes the token (plus the whitespace run
// before it) from its line; a line left empty is deleted entirely.
// AExpected = the line's text at resolve time - anything else means the
// buffer changed and the fix is stale.
function RemoveStrayToken(const AFile, AExpected, AToken: string;
  ALine0, ACol0: Integer): Boolean;
var
  Content, L, NewL: string;
  Lines: TArray<string>;
  StartP, EndP: Integer;
begin
  Result := False;
  if (Editor = nil) or (AToken = '') or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  L := Lines[ALine0];
  if L <> AExpected then Exit;                            // buffer changed
  StartP := ACol0 + 1;                                    // 1-based
  if (StartP < 1) or (StartP + Length(AToken) - 1 > Length(L)) then Exit;
  if not SameText(Copy(L, StartP, Length(AToken)), AToken) then Exit;
  EndP := StartP + Length(AToken) - 1;
  // Swallow the whitespace run before the token.
  while (StartP > 1) and CharInSet(L[StartP - 1], [' ', #9]) do Dec(StartP);
  NewL := Copy(L, 1, StartP - 1) + Copy(L, EndP + 1, MaxInt);
  if Trim(NewL) = '' then
    Result := Editor.DeleteLineAt(AFile, ALine0 + 1)
  else
    Result := Editor.ReplaceLineAt(AFile, ALine0 + 1, NewL);
end;

function RemoveDeadAssignment(const AFile, AExpected: string; ALine0: Integer): Boolean;
var
  Content: string;
  Lines: TArray<string>;
begin
  Result := False;
  if (Editor = nil) or (AExpected = '') or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  if Trim(StripLineComment(Lines[ALine0])) <> AExpected then Exit;
  Result := Editor.DeleteLineAt(AFile, ALine0 + 1);
end;

// W1010 fix: appends ' reintroduce;' after the declaration's first
// top-level ';' (reintroduce must be the first directive).
function AddReintroduceAt(const AFile: string; ALine0: Integer): Boolean;
var
  Content, L, S: string;
  Lines: TArray<string>;
  I, Depth, SemiP: Integer;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  Lines := SplitContentLines(Content);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  L := Lines[ALine0];
  S := StripLineComment(L);
  if Pos('REINTRODUCE', UpperCase(S)) > 0 then Exit;
  Depth := 0;
  SemiP := 0;
  for I := 1 to Length(S) do
    case S[I] of
      '(', '[': Inc(Depth);
      ')', ']': if Depth > 0 then Dec(Depth);
      ';': if Depth = 0 then
        begin
          SemiP := I;
          Break;
        end;
    end;
  if SemiP = 0 then Exit;
  Result := Editor.ReplaceLineAt(AFile, ALine0 + 1,
    Copy(L, 1, SemiP) + ' reintroduce;' + Copy(L, SemiP + 1, MaxInt));
end;

// ---------------------------------------------------------------------------
//  H2219 "private symbol declared but never used"
// ---------------------------------------------------------------------------
//
// Unlike H2164 (an unused LOCAL variable, which exists at exactly one
// place), a private member lives at TWO: the declaration inside the class
// and the implementation further down. Removing only one of them breaks
// the build, so the fix always removes both - and refuses whenever it
// cannot prove that this is safe.

function FindRemovablePrivate(const AContent: string; ADeclLine0: Integer;
  const AName: string; out AInfo: TPrivateMember): Boolean;
var
  Lines: TArray<string>;
  I, Depth, NameCount, ImplIdx: Integer;
  L, U, UName, Rest: string;

  function Clean(const S: string): string;
  var
    P: Integer;
  begin
    Result := Trim(S);
    P := Pos('//', Result);
    if P > 0 then Result := TrimRight(Copy(Result, 1, P - 1));
  end;

  // AUpper starts with the keyword AWord as a whole word.
  function StartsKw(const AUpper, AWord: string): Boolean;
  begin
    Result := AUpper.StartsWith(AWord)
      and ((Length(AUpper) = Length(AWord))
           or not CharInSet(AUpper[Length(AWord) + 1], ['A'..'Z', '0'..'9', '_']));
  end;

  // Whole-word occurrence of the member name.
  function HasName(const S: string): Boolean;
  var
    P, After: Integer;
    Up: string;
  begin
    Result := False;
    Up := UpperCase(S);
    P := Pos(UName, Up);
    while P > 0 do
    begin
      After := P + Length(UName);
      if ((P = 1) or not CharInSet(Up[P - 1], ['A'..'Z', '0'..'9', '_']))
        and ((After > Length(Up)) or not CharInSet(Up[After], ['A'..'Z', '0'..'9', '_'])) then
        Exit(True);
      P := PosEx(UName, Up, P + 1);
    end;
  end;

begin
  Result := False;
  AInfo := Default(TPrivateMember);
  if (AContent = '') or not IsBareIdent(AName) then Exit;
  Lines := SplitContentLines(AContent);
  if (ADeclLine0 < 0) or (ADeclLine0 > High(Lines)) then Exit;
  UName := UpperCase(AName);

  L := Clean(Lines[ADeclLine0]);
  U := UpperCase(L);
  if not HasName(L) then Exit;              // diagnostic no longer fits

  // Directives that mean the member is part of a contract - never remove.
  for var Bad in ['VIRTUAL', 'OVERRIDE', 'ABSTRACT', 'DYNAMIC', 'MESSAGE',
                  'OVERLOAD', 'DEFAULT', 'STORED', 'READ', 'WRITE'] do
    if Pos(Bad, U) > 0 then Exit;

  AInfo.Name := AName;
  AInfo.DeclLine := ADeclLine0;
  AInfo.DeclText := Lines[ADeclLine0];
  AInfo.TypeName := EnclosingContainerName(Lines, ADeclLine0);
  if AInfo.TypeName = '' then Exit;         // not inside a type body

  // The name must appear EXACTLY twice in the unit: the declaration and
  // the implementation header (or once for a field). Anything else means
  // it is referenced somewhere and H2219 is not what we think it is.
  NameCount := 0;
  for I := 0 to High(Lines) do
    if HasName(Clean(Lines[I])) then Inc(NameCount);
  if NameCount > 2 then Exit;

  // Implementation header: "procedure TType.Name" (also class/function/
  // constructor, and a generic "TType<T>.Name").
  ImplIdx := -1;
  for I := ADeclLine0 + 1 to High(Lines) do
  begin
    L := Clean(Lines[I]);
    if L = '' then Continue;
    U := UpperCase(L);
    if not (StartsKw(U, 'PROCEDURE') or StartsKw(U, 'FUNCTION')
      or StartsKw(U, 'CLASS PROCEDURE') or StartsKw(U, 'CLASS FUNCTION')
      or StartsKw(U, 'CONSTRUCTOR') or StartsKw(U, 'DESTRUCTOR')) then
      Continue;
    if HasName(L) and (Pos('.', L) > 0) then
    begin
      ImplIdx := I;
      Break;
    end;
  end;

  if ImplIdx < 0 then
  begin
    // A FIELD (or a const): declaration only, nothing else to remove.
    AInfo.ImplFirst := -1;
    AInfo.ImplLast := -1;
    AInfo.BodyEmpty := True;
    Exit(True);
  end;

  AInfo.ImplFirst := ImplIdx;

  // Walk to the routine's final "end;" - begin/case/try open a level,
  // "end" closes one. A local routine would nest, which the counter
  // handles as well.
  Depth := 0;
  AInfo.BodyLines := 0;
  for I := ImplIdx to High(Lines) do
  begin
    L := Clean(Lines[I]);
    if L = '' then Continue;
    U := UpperCase(L);
    if StartsKw(U, 'BEGIN') or StartsKw(U, 'TRY')
      or StartsKw(U, 'CASE') then
      Inc(Depth)
    else if (I > ImplIdx) and (Depth > 0) and (U = 'END;') then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        AInfo.ImplLast := I;
        AInfo.BodyEmpty := AInfo.BodyLines = 0;
        Exit(True);
      end;
    end
    else if Depth > 0 then
    begin
      // A statement inside the body - that is what the user would lose.
      if not (StartsKw(U, 'END') or StartsKw(U, 'ELSE')
        or StartsKw(U, 'EXCEPT') or StartsKw(U, 'FINALLY')
        or L.StartsWith('{') or L.StartsWith('(*')) then
        Inc(AInfo.BodyLines);
    end
    else if (I > ImplIdx) and (StartsKw(U, 'PROCEDURE')
      or StartsKw(U, 'FUNCTION')) then
      Break;   // ran into the next routine without finding our end
  end;
  // No delimitable body -> refuse rather than guess.
  Result := False;
end;

function RemovePrivateMemberText(const AContent: string;
  const AInfo: TPrivateMember): string;
var
  Lines: TArray<string>;
  SL: TStringList;
  I: Integer;
begin
  Result := AContent;
  Lines := SplitContentLines(AContent);
  if (AInfo.DeclLine < 0) or (AInfo.DeclLine > High(Lines)) then Exit;
  SL := TStringList.Create;
  try
    for I := 0 to High(Lines) do
    begin
      if I = AInfo.DeclLine then Continue;
      if (AInfo.ImplFirst >= 0) and (I >= AInfo.ImplFirst)
        and (I <= AInfo.ImplLast) then Continue;
      // A blank line directly after the removed implementation would
      // otherwise pile up.
      if (AInfo.ImplLast >= 0) and (I = AInfo.ImplLast + 1)
        and (Trim(Lines[I]) = '') then Continue;
      SL.Add(Lines[I]);
    end;
    Result := SL.Text;
  finally
    SL.Free;
  end;
end;

// Applier: re-validates against the CURRENT buffer, asks before deleting a
// non-empty body, then writes both removals in one minimal edit.
function RemoveUnusedPrivate(const AFile: string;
  const AInfo: TPrivateMember): Boolean;
var
  Content: string;
  Fresh: TPrivateMember;
  SL: TStringList;
begin
  Result := False;
  if (Editor = nil) or not ReadCurrentContent(AFile, Content) then Exit;
  // Re-locate in the CURRENT content: the buffer may have shifted since
  // the fix was resolved.
  if not FindRemovablePrivate(Content, AInfo.DeclLine, AInfo.Name, Fresh) then Exit;
  if Fresh.DeclText <> AInfo.DeclText then Exit;   // stale

  if not Fresh.BodyEmpty then
  begin
    // EXPLICIT warning with a way out - deleting a method that still does
    // something must never happen behind the user's back.
    if not Assigned(RemovePrivateConfirm) then Exit;
    if not RemovePrivateConfirm(Fresh) then Exit;
  end;

  SL := TStringList.Create;
  try
    SL.Text := RemovePrivateMemberText(Content, Fresh);
    Result := ApplyLinesMinimal(AFile, SL, Content);
  finally
    SL.Free;
  end;
end;

// Executes one quick fix. AUnitChoice picks the candidate unit for
// qfAddUnit (index into UnitNames).
function ApplyQuickFix(const AFile: string; const AFix: TQuickFix;
  AUnitChoice: Integer): Boolean;
begin
  Result := False;
  case AFix.Kind of
    qfAddUnit:
      begin
        if (AUnitChoice < 0) or (AUnitChoice > High(AFix.UnitNames)) then
          AUnitChoice := 0;
        if AUnitChoice > High(AFix.UnitNames) then Exit;
        Result := AddUnitToUses(AFile, AFix.UnitNames[AUnitChoice], AFix.Section);
      end;
    qfRenameIdent:
      begin
        Result := ReplaceTokenInLine(AFile, AFix.Line, AFix.Col,
          AFix.TokenLen, AFix.NewText);
        if Result and (AFix.FollowUpUnit <> '') then
          AddUnitToUses(AFile, AFix.FollowUpUnit, AFix.Section);
      end;
    qfFixUsesName:
      Result := ReplaceTokenInLine(AFile, AFix.Line, AFix.Col,
        AFix.TokenLen, AFix.NewText);
    qfRemoveUses:
      Result := RemoveUnitFromUses(AFile, AFix.OldUnit);
    qfAlignHeader:
      Result := AlignImplHeaderToDecl(AFile, AFix.Line);
    qfRemoveVar:
      Result := RemoveVarFromDecl(AFile, AFix.Identifier, AFix.Line);
    qfInsertSemi:
      Result := InsertSemicolonAfter(AFile, AFix.AuxLine, AFix.Col);
    qfInitVar:
      Result := InsertInitAtBegin(AFile, AFix.NewText, AFix.AuxLine);
    qfRemoveAssign:
      Result := RemoveDeadAssignment(AFile, AFix.NewText, AFix.Line);
    qfAddReintroduce:
      Result := AddReintroduceAt(AFile, AFix.Line);
    qfImplStub:
      Result := CreateImplStub(AFile, AFix.Line);
    qfClassStub:
      Result := CreateClassStub(AFile, AFix.Identifier, AFix.Line);
    qfRemoveToken:
      Result := RemoveStrayToken(AFile, AFix.NewText, AFix.Identifier,
        AFix.Line, AFix.Col);
    qfRemovePrivate:
      begin
        var Info := Default(TPrivateMember);
        Info.Name := AFix.Identifier;
        Info.TypeName := AFix.OldUnit;
        Info.DeclLine := AFix.Line;
        Info.DeclText := AFix.NewText;
        Result := RemoveUnusedPrivate(AFile, Info);
      end;
  end;
end;

// ---------------------------------------------------------------------------
//  Caret-anchored quick-fix chooser popup (VS-like)
// ---------------------------------------------------------------------------
//
// Lists the concrete fix ACTIONS for one source line: a qfAddUnit fix
// expands into one row per candidate unit; every other fix is one row.
// Double-click / Enter applies the selected action, Escape / clicking
// elsewhere cancels. The Section button switches the target uses section
// of the add-unit rows.

type
  TQuickFixAction = record
    Caption: string;
    FixIdx: Integer;      // index into FFixes
    UnitChoice: Integer;  // candidate-unit index for qfAddUnit, else -1
  end;

  TQuickFixPopup = class(TForm)
  private
    FFile: string;
    FFixes: TArray<TQuickFix>;
    FActions: TArray<TQuickFixAction>;
    FSection: TUsesSection;   // current target of the add-unit rows
    FHasUsesRows: Boolean;
    FLbl: TLabel;
    FList: TListBox;
    FSecBtn: TButton;
    FApplyBtn: TButton;
    /// <summary>Arms the close-on-deactivate handler AFTER the popup has
    ///  settled on screen: showing a THEMED form makes the IDE attach its
    ///  style hook, which can recreate/deactivate the window right after
    ///  the first Show - an immediately-armed OnDeactivate then Closed
    ///  (caFree!) the popup before the user ever saw it.</summary>
    FArmDeactivate: TTimer;
    procedure DoArmDeactivate(Sender: TObject);
    procedure RebuildActions;
    procedure DoApply(Sender: TObject);
    procedure DoToggleSection(Sender: TObject);
    procedure DoDeactivate(Sender: TObject);
    procedure DoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DoPopupClose(Sender: TObject; var Action: TCloseAction);
  public
    constructor CreatePopup(const AFile: string; const AFixes: TArray<TQuickFix>);
    procedure ShowAt(const APt: TPoint);
  end;

constructor TQuickFixPopup.CreatePopup(const AFile: string;
  const AFixes: TArray<TQuickFix>);
var
  OnlyIdent: string;
  OnlyAdds: Boolean;
begin
  inherited CreateNew(nil);
  FFile := AFile;
  FFixes := AFixes;
  BorderStyle := bsNone;
  FormStyle := fsStayOnTop;
  Color := GetThemedColor(clWindow);
  Width := 380;
  KeyPreview := True;
  // OnDeactivate is armed DEFERRED in ShowAt - see FArmDeactivate.
  OnKeyDown := DoKeyDown;
  OnClose := DoPopupClose;

  FArmDeactivate := TTimer.Create(Self);
  FArmDeactivate.Enabled := False;
  FArmDeactivate.Interval := 250;
  FArmDeactivate.OnTimer := DoArmDeactivate;

  // Section default + row kinds.
  FSection := usImplementation;
  FHasUsesRows := False;
  OnlyAdds := True;
  OnlyIdent := '';
  for var F in FFixes do
    if F.Kind = qfAddUnit then
    begin
      if not FHasUsesRows then FSection := F.Section;
      FHasUsesRows := True;
      if OnlyIdent = '' then OnlyIdent := F.Identifier;
    end
    else
      OnlyAdds := False;

  // Explicit Top values BEFORE Align: with several alTop controls the VCL
  // stacks them by their current position, so without this the last-created
  // control (the button panel) would end up on top.
  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Top := 0;
  FLbl.Align := alTop;
  FLbl.AlignWithMargins := True;
  if OnlyAdds and (OnlyIdent <> '') then
    FLbl.Caption := Format('Add unit for "%s":', [OnlyIdent])
  else
    FLbl.Caption := 'Quick fixes:';

  FList := TListBox.Create(Self);
  FList.Parent := Self;
  FList.Top := 50;
  FList.Align := alTop;
  FList.Height := 90;
  FList.AlignWithMargins := True;
  FList.OnDblClick := DoApply;   // double-click = apply straight away

  var Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Top := 200;
  Panel.Align := alTop;
  Panel.BevelOuter := bvNone;
  Panel.Height := 34;

  FSecBtn := TButton.Create(Self);
  FSecBtn.Parent := Panel;
  FSecBtn.Align := alLeft;
  FSecBtn.Width := 150;
  FSecBtn.AlignWithMargins := True;
  FSecBtn.OnClick := DoToggleSection;
  FSecBtn.Visible := FHasUsesRows;

  FApplyBtn := TButton.Create(Self);
  FApplyBtn.Parent := Panel;
  FApplyBtn.Align := alRight;
  FApplyBtn.Width := 80;
  FApplyBtn.AlignWithMargins := True;
  FApplyBtn.Caption := 'Apply';
  FApplyBtn.Default := True;
  FApplyBtn.OnClick := DoApply;

  RebuildActions;
  ClientHeight := FLbl.Height + FList.Height + Panel.Height + 20;
  EnableThemes(Self);
end;

procedure TQuickFixPopup.RebuildActions;
var
  Sel, I, U: Integer;
  A: TQuickFixAction;
begin
  Sel := FList.ItemIndex;
  FActions := nil;
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for I := 0 to High(FFixes) do
      case FFixes[I].Kind of
        qfAddUnit:
          for U := 0 to High(FFixes[I].UnitNames) do
          begin
            A.Caption := Format('Add %s to %s uses',
              [FFixes[I].UnitNames[U], SectionName(FSection)]);
            A.FixIdx := I;
            A.UnitChoice := U;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfRenameIdent:
          begin
            A.Caption := Format('Change to "%s"', [FFixes[I].NewText]);
            if FFixes[I].FollowUpUnit <> '' then
              A.Caption := A.Caption + Format('  (+ uses %s)', [FFixes[I].FollowUpUnit]);
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfFixUsesName:
          begin
            A.Caption := Format('Change unit name to "%s"', [FFixes[I].NewText]);
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfRemoveUses:
          begin
            A.Caption := Format('Remove "%s" from uses', [FFixes[I].OldUnit]);
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfAlignHeader:
          begin
            A.Caption := 'Align implementation header with declaration';
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfRemoveVar:
          begin
            A.Caption := Format('Remove unused variable "%s"', [FFixes[I].Identifier]);
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
        qfInsertSemi, qfInitVar, qfRemoveAssign, qfAddReintroduce,
        qfImplStub, qfClassStub, qfRemoveToken, qfRemovePrivate:
          begin
            // These carry an action-ready caption from the resolver.
            A.Caption := FFixes[I].Caption;
            A.FixIdx := I;
            A.UnitChoice := -1;
            FActions := FActions + [A];
            FList.Items.Add(A.Caption);
          end;
      end;
  finally
    FList.Items.EndUpdate;
  end;
  if FList.Items.Count > 0 then
  begin
    if (Sel >= 0) and (Sel < FList.Items.Count) then
      FList.ItemIndex := Sel
    else
      FList.ItemIndex := 0;
  end;
  FSecBtn.Caption := 'Section: ' + SectionName(FSection);
end;

procedure TQuickFixPopup.DoToggleSection(Sender: TObject);
begin
  if FSection = usInterface then FSection := usImplementation
  else FSection := usInterface;
  for var I := 0 to High(FFixes) do
    if FFixes[I].Kind = qfAddUnit then
      FFixes[I].Section := FSection;
  RebuildActions;
end;

procedure TQuickFixPopup.DoApply(Sender: TObject);
var
  A: TQuickFixAction;
  Ok: Boolean;
begin
  if (FList.ItemIndex < 0) or (FList.ItemIndex > High(FActions)) then Exit;
  A := FActions[FList.ItemIndex];
  // Browsing-path-only units: offer to make them project-visible first
  // (the availability dialog deactivates us - detach the close handler).
  if (FFixes[A.FixIdx].Kind = qfAddUnit)
    and (A.UnitChoice >= 0)
    and (A.UnitChoice <= High(FFixes[A.FixIdx].UnitNames)) then
  begin
    OnDeactivate := nil;
    if not EnsureUnitAvailable(FFixes[A.FixIdx].UnitNames[A.UnitChoice]) then
    begin
      Close;   // user cancelled - no message
      Exit;
    end;
  end;
  Ok := ApplyQuickFix(FFile, FFixes[A.FixIdx], A.UnitChoice);
  if not Ok then
  begin
    // ORDER MATTERS: the message must come BEFORE Close. Close triggers
    // caFree (deferred CM_RELEASE); a modal ShowMessage AFTER it pumps
    // messages, the release executes, and the VCL then returns into the
    // mouse handlers of the FREED listbox (observed AV on double-click).
    // Also detach OnDeactivate first - the message box deactivates us and
    // the handler would Close (and thereby free) us mid-pump.
    OnDeactivate := nil;
    ShowThemedMessage('The fix could not be applied (the target may already be ' +
      'in place, or the affected code could not be rewritten safely - ' +
      'e.g. IFDEFs inside a uses clause, or an ambiguous overload).');
  end;
  Close;
end;

procedure TQuickFixPopup.DoDeactivate(Sender: TObject);
begin
  Close;
end;

procedure TQuickFixPopup.DoArmDeactivate(Sender: TObject);
begin
  FArmDeactivate.Enabled := False;
  OnDeactivate := DoDeactivate;
  // The user may have clicked elsewhere during the unarmed window - the
  // Deactivate event for that is gone, so check the actual focus state.
  if GetActiveWindow <> Handle then
    Close;
end;

procedure TQuickFixPopup.DoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then Close;
end;

procedure TQuickFixPopup.DoPopupClose(Sender: TObject; var Action: TCloseAction);
begin
  // Non-modal, owner-less popup: Close alone would only hide it and leak
  // one instance per use. caFree releases it safely (deferred CM_RELEASE).
  Action := caFree;
end;

procedure TQuickFixPopup.ShowAt(const APt: TPoint);
begin
  // Keep the popup on the visible work area.
  var R := Screen.WorkAreaRect;
  Left := Min(APt.X, R.Right - Width - 4);
  Top := Min(APt.Y, R.Bottom - Height - 4);
  Show;
  FArmDeactivate.Enabled := True;   // arm close-on-deactivate deferred
end;

// True when the focused window is a code editor we know how to anchor to:
// the IDE's editor class is 'TEditControl', the standalone's is a TMemo.
// Everything else (Structure pane, Object Inspector, ...) has no usable
// text caret - GetCaretPos would return garbage for those.
function FocusedEditorWindow(out AWnd: HWND): Boolean;
var
  Cls: array[0..63] of Char;
  ClsName: string;
begin
  Result := False;
  AWnd := GetFocus;
  if AWnd = 0 then Exit;
  if GetClassName(AWnd, Cls, Length(Cls)) = 0 then Exit;
  ClsName := string(Cls);
  Result := SameText(ClsName, 'TEditControl') or SameText(ClsName, 'TMemo');
end;

// Caret position of the focused EDITOR, in screen coordinates. False when
// the focus is not on an editor or the caret lies outside the visible
// client area (e.g. its line has been scrolled out of view).
function CaretScreenPos(out APt: TPoint): Boolean;
var
  H: HWND;
  P: TPoint;
  R: TRect;
begin
  Result := False;
  if not FocusedEditorWindow(H) then Exit;
  if not GetCaretPos(P) then Exit;
  if not GetClientRect(H, R) then Exit;
  if not PtInRect(R, P) then Exit;
  ClientToScreen(H, P);
  Inc(P.Y, 18);   // just below the caret line
  APt := P;
  Result := True;
end;

// Anchor for a hint of AHeight pixels ABOVE the caret line. Code
// completion (the IDE's and ours) drops DOWN from the caret, so anything
// placed there covers its first entry - above the line nothing else
// appears. False when there is no room (caret on the first visible
// line); the caller then falls back to the position below.
function CaretAbovePos(AHeight: Integer; out APt: TPoint): Boolean;
var
  H: HWND;
  P: TPoint;
  R, WndR: TRect;
  Top: Integer;
begin
  Result := False;
  if not FocusedEditorWindow(H) then Exit;
  if not GetCaretPos(P) then Exit;
  if not GetClientRect(H, R) then Exit;
  if not PtInRect(R, P) then Exit;
  ClientToScreen(H, P);            // P.Y = top of the caret line
  Top := P.Y - AHeight - 2;
  // Stay inside the editor window - otherwise the hint would float over
  // the tab bar or the toolbar.
  if not GetWindowRect(H, WndR) then Exit;
  if (Top < WndR.Top) or (Top < Screen.WorkAreaRect.Top) then Exit;
  APt := TPoint.Create(P.X, Top);
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Batch dialog
// ---------------------------------------------------------------------------

type
  TAutoImportDialog = class(TForm)
  private
    FFile: string;
    FItems: TArray<TMissingIdent>;
    FSel: TArray<Integer>;      // chosen unit index per row
    FSecSel: TArray<TUsesSection>;
    FList: TListView;
    FCombo: TComboBox;
    FRadio: TRadioGroup;
    FBtnApply: TButton;
    FBtnClose: TButton;
    FLbl: TLabel;
    procedure Fill;
    procedure DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure DoComboChange(Sender: TObject);
    procedure DoRadioClick(Sender: TObject);
    procedure DoApply(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
  public
    constructor CreateDialog(AOwner: TComponent; const AFile: string;
      const AItems: TArray<TMissingIdent>);
  end;

constructor TAutoImportDialog.CreateDialog(AOwner: TComponent; const AFile: string;
  const AItems: TArray<TMissingIdent>);
var
  Col: TListColumn;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  FFile := AFile;
  FItems := AItems;
  SetLength(FSel, Length(FItems));
  SetLength(FSecSel, Length(FItems));
  for var I := 0 to High(FItems) do
  begin
    FSel[I] := 0;
    FSecSel[I] := FItems[I].Section;
  end;

  Caption := 'Resolve missing units';
  Width := 720;
  Height := 460;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Align := alTop;
  FLbl.AlignWithMargins := True;
  FLbl.Caption := Format('%d unresolved identifier(s) with a known unit. ' +
    'Tick the ones to add, adjust unit/section below.', [Length(FItems)]);

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.Checkboxes := True;
  FList.OnSelectItem := DoSelect;
  Col := FList.Columns.Add; Col.Caption := 'Identifier'; Col.Width := 200;
  Col := FList.Columns.Add; Col.Caption := 'Unit';       Col.Width := 220;
  Col := FList.Columns.Add; Col.Caption := 'Section';    Col.Width := 120;
  Col := FList.Columns.Add; Col.Caption := 'Options';    Col.Width := 90;

  // Bottom editing panel.
  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 96;
  Panel.BevelOuter := bvNone;

  FCombo := TComboBox.Create(Self);
  FCombo.Parent := Panel;
  FCombo.Style := csDropDownList;
  FCombo.Left := 8; FCombo.Top := 8; FCombo.Width := 300;
  FCombo.OnChange := DoComboChange;

  FRadio := TRadioGroup.Create(Self);
  FRadio.Parent := Panel;
  FRadio.Left := 320; FRadio.Top := 4; FRadio.Width := 220; FRadio.Height := 52;
  FRadio.Caption := 'Target section';
  FRadio.Columns := 2;
  FRadio.Items.Add('interface');
  FRadio.Items.Add('implementation');
  FRadio.OnClick := DoRadioClick;

  FBtnApply := TButton.Create(Self);
  FBtnApply.Parent := Panel;
  FBtnApply.Caption := '&Apply checked';
  FBtnApply.Left := 8; FBtnApply.Top := 60; FBtnApply.Width := 130;
  FBtnApply.OnClick := DoApply;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Left := 150; FBtnClose.Top := 60; FBtnClose.Width := 90;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  Fill;
  EnableListViewSorting(FList);
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TAutoImportDialog.Fill;
var Item: TListItem;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for var I := 0 to High(FItems) do
    begin
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(I));   // survives column sorting
      if FItems[I].GenericUse then
        Item.Caption := FItems[I].Identifier + '<>'
      else
        Item.Caption := FItems[I].Identifier;
      Item.Checked := True;
      Item.SubItems.Add(FItems[I].Units[FSel[I]].UnitName);
      Item.SubItems.Add(SectionName(FSecSel[I]));
      if Length(FItems[I].Units) > 1 then
        Item.SubItems.Add(Format('%d units', [Length(FItems[I].Units)]))
      else
        Item.SubItems.Add('');
    end;
  finally
    FList.Items.EndUpdate;
  end;
end;

procedure TAutoImportDialog.DoSelect(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  if not Selected or (Item = nil) then Exit;
  var Idx := NativeInt(Item.Data);
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  FCombo.Items.BeginUpdate;
  try
    FCombo.Items.Clear;
    for var H in FItems[Idx].Units do FCombo.Items.Add(H.UnitName + '   (' + H.Path + ')');
  finally
    FCombo.Items.EndUpdate;
  end;
  FCombo.ItemIndex := FSel[Idx];
  FRadio.ItemIndex := Ord(FSecSel[Idx]);
end;

procedure TAutoImportDialog.DoComboChange(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := NativeInt(FList.Selected.Data);
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if (FCombo.ItemIndex >= 0) and (FCombo.ItemIndex <= High(FItems[Idx].Units)) then
  begin
    FSel[Idx] := FCombo.ItemIndex;
    FList.Selected.SubItems[0] := FItems[Idx].Units[FSel[Idx]].UnitName;
  end;
end;

procedure TAutoImportDialog.DoRadioClick(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := NativeInt(FList.Selected.Data);
  if (Idx < 0) or (Idx > High(FItems)) then Exit;
  if FRadio.ItemIndex >= 0 then
  begin
    FSecSel[Idx] := TUsesSection(FRadio.ItemIndex);
    FList.Selected.SubItems[1] := SectionName(FSecSel[Idx]);
  end;
end;

procedure TAutoImportDialog.DoApply(Sender: TObject);
var Added, Failed: Integer;
begin
  Added := 0; Failed := 0;
  for var I := 0 to FList.Items.Count - 1 do
    if FList.Items[I].Checked then
    begin
      var Idx := NativeInt(FList.Items[I].Data);
      if (Idx < 0) or (Idx > High(FItems)) then Continue;
      if ApplyOne(FFile, FItems[Idx].Units[FSel[Idx]].UnitName, FSecSel[Idx]) then Inc(Added)
      else Inc(Failed);
    end;
  ShowThemedMessage(Format('%d unit(s) added.%s', [Added,
    IfThen(Failed > 0, Format(#13#10'%d already present / not written.', [Failed]), '')]));
  Close;
end;

procedure TAutoImportDialog.DoCloseClick(Sender: TObject);
begin
  Close;
end;

// ---------------------------------------------------------------------------
//  Live quick-fix indicator ("lightbulb")
// ---------------------------------------------------------------------------
//
// A 600 ms UI timer watches the active editor buffer (via the CHEAP
// GetActiveFileName + ReadEditorContent - never GetCurrentContext, which
// would move the caret). When the content has been stable for ~1.5 s and
// differs from the last analysed state, the poller
//   1. on the MAIN thread: captures the content, sends RefreshDocument and
//      an async documentSymbol request (both are quick pipe writes),
//   2. on a WORKER thread: waits for the documentSymbol response and the
//      diagnostics push, then extracts the missing identifiers (pure
//      function + lock-free index snapshot - no editor access),
//   3. back on the main thread: stores the results and shows/hides the
//      hint button next to the caret.
// The hint window is WS_EX_NOACTIVATE, so clicking it never steals the
// editor focus mid-typing.

type
  TAutoImportHint = class(TForm)
  private
    FLbl: TLabel;
    FOnFix: TProc;
    FBorderColor: TColor;
    procedure DoClick(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    /// <summary>Draws the 1 px frame. The label leaves that pixel free
    ///  via Padding, so the hint reads as a panel ON the editor instead
    ///  of blending into it (tester: "no border, almost the same
    ///  colour").</summary>
    procedure Paint; override;
  public
    constructor CreateHint;
    procedure SetInfo(const AText: string);
    procedure ShowNoActivateAt(const APt: TPoint);
    /// <summary>Counterpart to ShowNoActivateAt. The window is shown via
    ///  plain ShowWindow (bypassing VCL's Visible), so TForm.Hide would be
    ///  a NO-OP - it must be hidden the same WinAPI way.</summary>
    procedure HideHint;
    property OnFix: TProc read FOnFix write FOnFix;
  end;

constructor TAutoImportHint.CreateHint;
begin
  inherited CreateNew(nil);
  BorderStyle := bsNone;
  var BackCol, TextCol: TColor;
  ComputeHintColors(GetThemedColor(clInfoBk), GetThemedColor(clInfoText),
    BackCol, FBorderColor, TextCol);
  Color := BackCol;
  Height := 24;
  Width := 160;
  Cursor := crHandPoint;
  OnClick := DoClick;
  // One pixel all round stays unpainted by the label - that is the frame.
  Padding.SetBounds(1, 1, 1, 1);

  FLbl := TLabel.Create(Self);
  FLbl.Parent := Self;
  FLbl.Align := alClient;
  FLbl.Alignment := taCenter;
  FLbl.Layout := tlCenter;
  FLbl.Transparent := True;
  FLbl.Font.Color := TextCol;
  FLbl.Cursor := crHandPoint;
  FLbl.OnClick := DoClick;
end;

procedure TAutoImportHint.Paint;
begin
  inherited;
  Canvas.Brush.Style := bsClear;
  Canvas.Pen.Color := FBorderColor;
  Canvas.Pen.Width := 1;
  Canvas.Rectangle(0, 0, ClientWidth, ClientHeight);
end;

procedure TAutoImportHint.CreateParams(var Params: TCreateParams);
begin
  inherited;
  // Never activate: the user keeps typing in the editor while we appear.
  Params.ExStyle := Params.ExStyle or WS_EX_NOACTIVATE or WS_EX_TOPMOST
    or WS_EX_TOOLWINDOW;
  Params.Style := (Params.Style and not WS_CHILD) or WS_POPUP;
end;

procedure TAutoImportHint.SetInfo(const AText: string);
begin
  FLbl.Caption := #$1F4A1' ' + AText;
  Canvas.Font := FLbl.Font;
  Width := Max(120, Canvas.TextWidth(FLbl.Caption) + 26);
  Invalidate;   // the frame must follow the new width
end;

procedure TAutoImportHint.ShowNoActivateAt(const APt: TPoint);
var
  R: TRect;
begin
  R := Screen.WorkAreaRect;
  SetWindowPos(Handle, HWND_TOPMOST,
    Min(APt.X, R.Right - Width - 4), Min(APt.Y, R.Bottom - Height - 4),
    Width, Height, SWP_NOACTIVATE);
  if not IsWindowVisible(Handle) then
    ShowWindow(Handle, SW_SHOWNOACTIVATE);
end;

procedure TAutoImportHint.HideHint;
begin
  if HandleAllocated and IsWindowVisible(Handle) then
    ShowWindow(Handle, SW_HIDE);
end;

procedure TAutoImportHint.DoClick(Sender: TObject);
begin
  if Assigned(FOnFix) then FOnFix();
end;

type
  TAutoImportLive = class
  private
    FTimer: TTimer;
    FHint: TAutoImportHint;
    // Current buffer state (UI thread only).
    FFile: string;
    FHash: Integer;
    FDirty: Boolean;
    FDirtyTick: Cardinal;
    FDirtyIdleMs: Cardinal;   // idle threshold for the pending state
    FSwitched: Boolean;       // the active buffer changed on this tick
    FTickNo: Integer;         // fast ticks since start (heavy work every N)
    FHintDirty: Boolean;      // results changed - refresh the hint at once
    // Popup that currently keeps the hint hidden (see UpdateHint).
    FBlockWnd: HWND;
    FBlockSince: Cardinal;
    FBlockCls: string;
    FAnalysing: Boolean;
    // Last resolved quick fixes (UI thread only). TWO sources feed them:
    // the Structure view (fast, ERRORS only) and our own LSP session
    // (slower, errors + HINTS like H2443/H2164). The LSP result is the
    // superset and REPLACES a structure result for the same buffer state;
    // FResFromLsp guards the other direction - a late structure answer
    // must never wipe the hint fixes again.
    FResFile: string;
    FResHash: Integer;
    FResults: TArray<TQuickFix>;
    FResFromLsp: Boolean;
    // WHICH source produced FResults (SrcLsp / SrcStructure / SrcCompiler).
    // An empty payload from THAT source is a retraction and clears the
    // fixes at once; from any other source it is just silence.
    FResSource: string;
    FResDiagSig: Integer;   // fingerprint of the payload behind FResults
    // Pipeline counters of the last resolution (status window only).
    FDiagSeen: Integer;
    FDiagHandled: Integer;
    FDiagCodes: string;
    FDiagNote: string;   // why a fixable diagnostic produced no fix
    FDiagSig: Integer;   // fingerprint of the payload the LSP pass resolved
    // Per-source outcome of the last pass (status window).
    FSrcStructure: string;
    FSrcLsp: string;
    // External payload waiting for background resolution (fuzzy index
    // scans are too slow for the notifier's main-thread context).
    FPendFile: string;
    FPendContent: string;
    FPendDiags: TArray<TLspErrorDiag>;
    FPendHash: Integer;
    FPendSource: string;
    FPendSig: Integer;
    FHasPending: Boolean;
    FResolving: Boolean;
    // One-shot: run the LSP analysis on the next tick even though the
    // Structure view answered this buffer state (set after a real
    // compile, to pick up hint diagnostics like H2443).
    FCompileRefreshPending: Boolean;
    // Last results revision the editor markers were repainted for.
    FLastPaintFile: string;
    FLastPaintHash: Integer;
    procedure ApplyExternalDiags(const AFile, AContent: string;
      const ADiags: TArray<TLspErrorDiag>; const ASource: string);
    procedure StartPendingResolve;
    procedure DoTick(Sender: TObject);
    procedure StartAnalysis(const AFile, AContent: string; AHash: Integer);
    procedure AnalysisDone(const AFile: string; AHash: Integer;
      const AResults: TArray<TQuickFix>);
    procedure UpdateHint;
    procedure RunFix;
    function HasFresh(const AFile: string): Boolean;
    /// <summary>All quick fixes anchored to the caret's LINE - the hint is
    ///  bound to the error location (VS-style), not shown for arbitrary
    ///  caret positions in the file.</summary>
    function FindFixesAtCaretLine(out AFixes: TArray<TQuickFix>): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
  end;

var
  GLive: TAutoImportLive = nil;
  // Number of live resolver/analysis worker threads still running, and the
  // shutdown latch. StopAutoImportLive waits for the count to reach zero
  // and drains the sync queue BEFORE the package can unload - otherwise a
  // worker (or its queued TThread.Queue closure) would execute unmapped
  // BPL code, or race TUnitIndex's finalization.
  GLiveWorkers: Integer = 0;
  GLiveShutdown: Boolean = False;

const
  // The tick does TWO things of very different cost: reading the buffer
  // and hashing it (heavy - a 4000-line unit on every tick would be
  // wasteful), and updating the hint from already published results plus
  // the caret position (cheap). They get different cadences: the hint
  // follows the caret and a freshly published result within one FAST
  // tick, the buffer is only re-read every HeavyEvery-th one.
  LiveFastTickMs = 120;
  LiveHeavyEvery = 4;   // -> ~480 ms for the buffer read / hashing
  LiveTickMs  = 400;   // poll interval (also the max hint-update latency,
                       // since ONLY the tick may touch the hint window)
  LiveIdleMs  = 1500;  // buffer must be stable this long before analysing
  // Switching to another buffer is not typing - the content is already
  // stable, so the file the user is LOOKING AT gets analysed almost at
  // once instead of waiting out the typing pause.
  LiveSwitchIdleMs = 200;
  // How long one and the same window may suppress the hint before we
  // conclude it is not a transient completion list.
  HintBlockMaxMs = 4000;

function H2219NameFromMessage(const AMessage: string): string;
var
  P: Integer;
begin
  Result := FirstQuoted(AMessage);
  // Generic method: "Test<T>" - the arity marker is not part of the name.
  P := Pos('<', Result);
  if P > 0 then Result := Copy(Result, 1, P - 1);
  // Qualified: "TForm4.Test" - keep the member.
  P := LastDelimiter('.', Result);
  if P > 0 then Result := Copy(Result, P + 1, MaxInt);
  Result := Trim(Result);
  if not IsBareIdent(Result) then Result := '';
end;

function CanTakeSemicolon(const ALine: string): Boolean;
const
  // Keywords a statement never ends with - a ';' behind them is invalid.
  Openers: array[0..17] of string = (
    'BEGIN', 'THEN', 'ELSE', 'DO', 'OF', 'TRY', 'REPEAT', 'EXCEPT',
    'FINALLY', 'VAR', 'CONST', 'TYPE', 'TO', 'DOWNTO', 'IN', 'AND',
    'OR', 'NOT');
var
  S, LastWord: string;
  I: Integer;
begin
  Result := False;
  S := TrimRight(StripLineComment(ALine));
  if S = '' then Exit;
  // Trailing operator / separator: the statement continues on the next
  // line, so nothing is missing HERE.
  if CharInSet(S[Length(S)],
    ['+', '-', '*', '/', '=', '<', '>', ',', '(', '[', ':', '.', '@', '^']) then
    Exit;
  // Last word.
  I := Length(S);
  while (I >= 1) and CharInSet(S[I], ['A'..'Z', 'a'..'z', '0'..'9', '_']) do Dec(I);
  LastWord := UpperCase(Copy(S, I + 1, MaxInt));
  for var K in Openers do
    if LastWord = K then Exit;
  Result := True;
end;

// Fingerprint of a diagnostic payload (codes + positions). Two payloads
// with the same fingerprint would resolve to the same fixes, so re-running
// the resolver for them is pure waste - but a payload that CHANGED must be
// resolved even when the buffer did not, because the source just
// re-evaluated it (typical: Error Insight catches up after a fix was
// applied and no longer reports the error).
function DiagSignature(const ADiags: TArray<TLspErrorDiag>): Integer;
var
  S: string;
begin
  S := '';
  for var D in ADiags do
    S := S + D.Code + '|' + IntToStr(D.Range.Start.Line) + ':' +
         IntToStr(D.Range.Start.Character) + ';';
  Result := THashBobJenkins.GetHashValue(S);
end;

function IsHandledDiagCode(const ACode: string): Boolean;
const
  Handled: array[0..12] of string = (
    'E2003', 'F2613', 'F2063', 'E2037', 'H2443', 'H2164',
    'E2029', 'E2066', 'W1036', 'H2077', 'W1010', 'E2065', 'H2219');
begin
  Result := False;
  for var C in Handled do
    if SameText(C, ACode) then Exit(True);
end;

// Fills the counters from a diagnostic payload (worker context - plain
// state, no UI).
procedure NoteDiagStats(const ADiags: TArray<TLspErrorDiag>;
  out ASeen, AHandled: Integer; out ACodes: string);
var
  Seen: TStringList;
begin
  ASeen := Length(ADiags);
  AHandled := 0;
  ACodes := '';
  Seen := TStringList.Create;
  try
    Seen.Duplicates := dupIgnore;
    Seen.Sorted := True;
    for var D in ADiags do
    begin
      if IsHandledDiagCode(D.Code) then Inc(AHandled);
      if D.Code <> '' then Seen.Add(D.Code);
    end;
    for var I := 0 to Seen.Count - 1 do
    begin
      if I >= 8 then
      begin
        ACodes := ACodes + ', ...';
        Break;
      end;
      if ACodes <> '' then ACodes := ACodes + ', ';
      ACodes := ACodes + Seen[I];
    end;
  finally
    Seen.Free;
  end;
end;

constructor TAutoImportLive.Create;
begin
  inherited;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := LiveFastTickMs;
  FTimer.OnTimer := DoTick;
  FTimer.Enabled := True;
end;

destructor TAutoImportLive.Destroy;
begin
  FreeAndNil(FTimer);
  FreeAndNil(FHint);
  inherited;
end;

function TAutoImportLive.HasFresh(const AFile: string): Boolean;
begin
  Result := (AFile <> '') and SameText(FResFile, AFile)
    and SameText(FResFile, FFile) and (FResHash = FHash);
end;

function TAutoImportLive.FindFixesAtCaretLine(out AFixes: TArray<TQuickFix>): Boolean;
var
  Line, Col: Integer;
begin
  Result := False;
  AFixes := nil;
  if not HasFresh(FFile) then Exit;
  if (Editor = nil) or not Editor.GetCaretLineCol(Line, Col) then Exit;
  for var R in FResults do
    if R.Line = Line - 1 then   // R.Line is 0-based
      AFixes := AFixes + [R];
  Result := Length(AFixes) > 0;
end;

procedure TAutoImportLive.DoTick(Sender: TObject);
var
  F, Content: string;
  H: Integer;
  Tick: Cardinal;
begin
  if (Editor = nil) or GOnDemandBusy then Exit;
  // No live churn while any modal dialog (incl. our own wizards) is open.
  if Application.ModalLevel > 0 then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  F := Editor.GetActiveFileName;
  if (F = '') or not SameText(ExtractFileExt(F), '.pas') then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  Inc(FTickNo);
  // CHEAP PATH: nothing to re-read - just keep the hint in sync with the
  // caret and with results that arrived since the last tick. Skipping the
  // heavy part here is what makes a published fix appear within ~120 ms
  // instead of waiting out the poll interval.
  if (FTickNo mod LiveHeavyEvery <> 0) and SameText(F, FFile)
    and not FHintDirty then
  begin
    UpdateHint;
    Exit;
  end;
  FHintDirty := False;

  if not SameText(F, FFile) then
  begin
    FSwitched := True;
    FFile := F;
    FHash := 0;
    FDirty := False;   // hash below re-arms
    if FHint <> nil then FHint.HideHint;
  end;

  if not Editor.ReadEditorContent(F, Content) then
  begin
    if FHint <> nil then FHint.HideHint;
    Exit;
  end;

  H := THashBobJenkins.GetHashValue(Content);
  Tick := GetTickCount;
  if H <> FHash then
  begin
    FHash := H;
    FDirty := True;
    FDirtyTick := Tick;
    // Typing needs the full pause; a buffer SWITCH does not.
    if FSwitched then FDirtyIdleMs := LiveSwitchIdleMs
    else FDirtyIdleMs := LiveIdleMs;
    if FHint <> nil then FHint.HideHint;   // results are stale now
  end;
  FSwitched := False;

  if FCompileRefreshPending and (not FAnalysing)
    and (TLspManager.Instance.PeekClient <> nil) then
  begin
    // Post-compile refresh: immediate LSP round (no idle wait) - the
    // compile just produced the full diagnostic picture.
    FCompileRefreshPending := False;
    StartAnalysis(F, Content, H);
  end
  else if FDirty and (not FAnalysing) and (Tick - FDirtyTick >= FDirtyIdleMs) then
    // One LSP pass per buffer state, ALWAYS - even when the Structure
    // view answered it: that source carries only errors, the hint fixes
    // (H2443/H2164) exist solely in our own session's diagnostics. The
    // LSP result replaces the structure result as the superset.
    StartAnalysis(F, Content, H);

  UpdateHint;

  // Editor line markers: repaint once per published results revision.
  // Deliberately from THIS tick (WM_TIMER) - never from the notifier or
  // queue contexts that update the results (state-only rule).
  if ((FResFile <> FLastPaintFile) or (FResHash <> FLastPaintHash))
    and Assigned(GLiveRepaintHook) then
  begin
    FLastPaintFile := FResFile;
    FLastPaintHash := FResHash;
    try
      GLiveRepaintHook();
    except
    end;
  end;
end;

procedure TAutoImportLive.ApplyExternalDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>; const ASource: string);
var
  NewHash, Sig: Integer;
begin
  if AFile = '' then Exit;
  // STATE ONLY - no window operations. This runs inside the IDE's
  // Structure-view notification, which the IDE dispatches from
  // CheckSynchronize during its LSP refresh. Showing/hiding a window
  // there triggers a synchronous activation cascade into
  // TEditWindow.ActivateModule -> ParserThread.CancelAndLock and can
  // DEADLOCK the IDE against its parser thread. The poll tick (a plain
  // WM_TIMER, safe context) picks the results up and updates the hint.
  //
  // The RESOLUTION itself (index lookups + fuzzy scans) is too heavy for
  // this context as well - it runs on a worker thread; the latest payload
  // wins when notifications arrive faster than the resolver finishes.
  NewHash := THashBobJenkins.GetHashValue(AContent);

  // Already resolved for exactly this buffer state? The Structure pane
  // rebuilds repeatedly for the same content (every IDE LSP refresh) -
  // re-running the resolver (incl. the O(IdentCount) fuzzy scan) for a
  // byte-identical buffer would just burn a core. This also protects an
  // LSP-published result (with hint fixes) from being re-queued.
  //
  // ONE EXCEPTION, and it cost the tester every fix (status window:
  // "Structure source: 16 diag" while "from Structure: 0 diag -> 0 fix"):
  // when the published result is EMPTY and the incoming payload actually
  // carries diagnostics, this state can still be improved. Our own LSP
  // pass often finishes FIRST with nothing at all (it can be alive and
  // silent for a file), claims the state with an empty result - and the
  // Structure payload with the real errors was then dropped right here,
  // before it was ever resolved. Nothing re-analyses an unchanged buffer,
  // so that was final.
  // ... but ONLY when the payload is also unchanged. A source that
  // re-evaluates the SAME buffer and now reports FEWER diagnostics (the
  // user applied a fix, Error Insight caught up) must get through - with
  // the old test its result was dropped here and the obsolete fix stayed
  // on screen for good (tester: "added the unit, the autofix stayed").
  Sig := DiagSignature(ADiags);
  if SameText(AFile, FResFile) and (NewHash = FResHash)
    and (Sig = FResDiagSig)
    and not ((Length(FResults) = 0) and (Length(ADiags) > 0)) then
  begin
    FFile := AFile;
    FHash := NewHash;
    FHasPending := False;
    Exit;
  end;

  // A NEW buffer state: schedule the complementary LSP pass too - the
  // Structure view only carries ERRORS, the hint fixes (H2443/H2164)
  // exist only in our own LSP session's diagnostics. Without this, a
  // hint fix would only ever (re)appear after a compile.
  if (NewHash <> FHash) or not SameText(AFile, FFile) then
  begin
    FDirty := True;
    FDirtyTick := GetTickCount;
    FDirtyIdleMs := LiveIdleMs;
  end;
  FFile := AFile;
  FHash := NewHash;

  FPendFile := AFile;
  FPendContent := AContent;
  FPendDiags := ADiags;
  FPendHash := FHash;
  FPendSource := ASource;
  FPendSig := Sig;
  FHasPending := True;
  if not FResolving then
    StartPendingResolve;
end;

procedure TAutoImportLive.StartPendingResolve;
var
  RFile, RContent: string;
  RDiags: TArray<TLspErrorDiag>;
  RHash, RSig: Integer;
  RSource: string;
begin
  if not FHasPending or GLiveShutdown then Exit;
  FHasPending := False;
  FResolving := True;
  RFile := FPendFile;
  RContent := FPendContent;
  RDiags := FPendDiags;
  RHash := FPendHash;
  RSource := FPendSource;
  RSig := FPendSig;

  // Exception-safe start: if thread creation fails, FResolving must not
  // wedge True (that would silence the whole live feature until restart).
  TInterlocked.Increment(GLiveWorkers);
  try
    TThread.CreateAnonymousThread(
      procedure
      var
        R: TArray<TQuickFix>;
        SSeen, SHandled: Integer;
        SCodes, SNote: string;
      begin
        try
          SSeen := 0; SHandled := 0; SCodes := ''; SNote := '';
          try
            NoteDiagStats(RDiags, SSeen, SHandled, SCodes);
            R := ResolveQuickFixes(RContent, RDiags);
            SNote := LastResolveNote;
          except
            R := nil;
          end;
          TThread.Queue(nil,
            procedure
            begin
              if GLive = nil then Exit;
              GLive.FResolving := False;
              GLive.FDiagSeen := SSeen;
              GLive.FDiagHandled := SHandled;
              GLive.FDiagCodes := SCodes;
              GLive.FDiagNote := SNote;
              // Accept only when the buffer has not changed since - and
              // normally never downgrade an LSP result (errors + hints)
              // for the same state to this errors-only structure result.
              //
              // BUT an EMPTY LSP result must not block us (tester: the
              // Structure view listed E2003/E2066 while the live checker
              // said "0 fix(es)"). Our LSP session can be alive and still
              // answer nothing for a file; its empty result would then
              // mask the errors the IDE visibly shows, forever - the
              // buffer state is unchanged, so nothing ever re-analyses.
              var LspHasResult := GLive.FResFromLsp
                and SameText(RFile, GLive.FResFile)
                and (RHash = GLive.FResHash)
                and (Length(GLive.FResults) > 0);
              // An empty result from ANOTHER source is just silence and
              // must not wipe what we have (a Structure rebuild fires
              // again while its tree is repopulating; our LSP can be mute
              // for a file). From the SAME source it is a RETRACTION -
              // that source re-evaluated this exact buffer state and the
              // error is gone, so the fixes go immediately.
              if (Length(R) = 0) and (Length(GLive.FResults) > 0)
                and SameText(RFile, GLive.FResFile)
                and (RHash = GLive.FResHash)
                and not SameText(GLive.FResSource, RSource) then
                GLive.FSrcStructure := Format(
                  '%d diag -> 0 fix, ignored: empty, keeping %d fix(es) from %s',
                  [SSeen, Length(GLive.FResults), GLive.FResSource])
              else if not (SameText(RFile, GLive.FFile) and (RHash = GLive.FHash)) then
                GLive.FSrcStructure := Format(
                  '%d diag -> %d fix, DROPPED: buffer changed while resolving',
                  [SSeen, Length(R)])
              else if LspHasResult then
                GLive.FSrcStructure := Format(
                  '%d diag -> %d fix, DROPPED: LSP result wins for this state',
                  [SSeen, Length(R)])
              else
              begin
                GLive.FSrcStructure := Format('%d diag -> %d fix, published (%s)',
                  [SSeen, Length(R), RSource]);
                if SNote <> '' then
                  GLive.FSrcStructure := GLive.FSrcStructure +
                    ' | declined: ' + SNote;
                GLive.FResFile := RFile;
                GLive.FResHash := RHash;
                GLive.FResults := R;
                GLive.FResFromLsp := False;
                GLive.FResSource := RSource;
                GLive.FResDiagSig := RSig;
                GLive.FHintDirty := True;
              end;
              // A newer payload may have arrived while we were resolving.
              GLive.StartPendingResolve;
            end);
        finally
          TInterlocked.Decrement(GLiveWorkers);
        end;
      end).Start;
  except
    TInterlocked.Decrement(GLiveWorkers);
    FResolving := False;
  end;
end;

procedure TAutoImportLive.StartAnalysis(const AFile, AContent: string; AHash: Integer);
var
  Client: TLspClient;
  Params, TextDocObj: TJSONObject;
  Before, BeforeVer, ReqId: Integer;
begin
  // NEVER cold-start the LSP from the background poller - only piggyback
  // on a client the prewarmer / a wizard has already brought up.
  Client := TLspManager.Instance.PeekClient;
  if Client = nil then
  begin
    FDirtyTick := GetTickCount;   // retry after another idle period
    FDirtyIdleMs := LiveIdleMs;
    Exit;
  end;

  Before := Client.GetDiagnosticsCount;
  // Snapshot of THIS file's push counter - the worker waits for it to
  // change, so a push for any other file can no longer end the wait.
  BeforeVer := Client.GetFileDiagnosticsVersion(AFile);
  try
    // Push the CURRENT buffer (RefreshDocument reads the live editor
    // content) and fire an async documentSymbol to force the analysis.
    Client.RefreshDocument(AFile);
    TextDocObj := TJSONObject.Create;
    TextDocObj.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFile)));
    Params := TJSONObject.Create;
    Params.AddPair('textDocument', TextDocObj);
    ReqId := Client.SendRequestAsync('textDocument/documentSymbol', Params);
  except
    Exit;   // client just died etc. - try again on a later tick
  end;

  FAnalysing := True;
  FDirty := False;

  TInterlocked.Increment(GLiveWorkers);
  try
    TThread.CreateAnonymousThread(
      procedure
      var
        R: TArray<TQuickFix>;
        I, DSeen, DHandled, DSig: Integer;
        DCodes, DNote: string;
      begin
        try
          R := nil;
          DSeen := 0; DHandled := 0; DCodes := ''; DNote := ''; DSig := 0;
          try
            try Client.WaitForResponse(ReqId, 20000).Free; except end;
            // Wait for the push that belongs to THIS file. Waiting on the
            // session-wide counter meant any push for any other file
            // ended the wait, and we then read a stale (usually empty)
            // result for the file we actually care about - the busier the
            // session, the more often. Poll fast at first: the answer for
            // the active buffer usually comes within a few hundred ms.
            for I := 1 to 60 do
            begin
              if GLiveShutdown then Break;
              if Client.GetFileDiagnosticsVersion(AFile) > BeforeVer then Break;
              if I <= 20 then Sleep(50) else Sleep(150);
            end;
            if not GLiveShutdown then
            begin
              var Diags := Client.GetErrorDiagnostics(AFile);
              NoteDiagStats(Diags, DSeen, DHandled, DCodes);
              DSig := DiagSignature(Diags);
              R := ResolveQuickFixes(AContent, Diags);
              DNote := LastResolveNote;   // same thread - our own reason
            end;
          except
            R := nil;
          end;
          TThread.Queue(nil,
            procedure
            begin
              if GLive <> nil then
              begin
                GLive.FDiagSeen := DSeen;
                GLive.FDiagHandled := DHandled;
                GLive.FDiagCodes := DCodes;
                GLive.FDiagNote := DNote;
                GLive.FDiagSig := DSig;
                GLive.AnalysisDone(AFile, AHash, R);
              end;
            end);
        finally
          TInterlocked.Decrement(GLiveWorkers);
        end;
      end).Start;
  except
    TInterlocked.Decrement(GLiveWorkers);
    FAnalysing := False;
  end;
end;

procedure TAutoImportLive.AnalysisDone(const AFile: string; AHash: Integer;
  const AResults: TArray<TQuickFix>);
begin
  // STATE ONLY (arrives via TThread.Queue, i.e. inside CheckSynchronize -
  // see ApplyExternalDiags for why no window operation may happen here).
  // The next poll tick shows/hides the hint.
  FAnalysing := False;
  // Only accept results that still match the current buffer. LSP results
  // are normally the SUPERSET (errors + hints) and replace whatever the
  // Structure view published for this state.
  FSrcLsp := Format('%d diag -> %d fix', [FDiagSeen, Length(AResults)]);
  if not (SameText(AFile, FFile) and (AHash = FHash)) then
    FSrcLsp := FSrcLsp + ', DROPPED: buffer changed while analysing';
  if SameText(AFile, FFile) and (AHash = FHash) then
  begin
    // EXCEPTION (tester, status window showed "0 diagnostic push(es)"
    // while the Structure view listed E2003/E2066): our LSP session can
    // be alive yet never deliver a single diagnostic - on a large
    // project whose configuration it cannot fully resolve. Its EMPTY
    // result would then permanently mask the structure result, and the
    // user gets no fixes at all although the IDE shows the errors. An
    // empty LSP answer never overrides a non-empty one for the SAME
    // buffer state (same content hash = the errors are demonstrably
    // still there).
    // Same rule as for the external sources: an empty LSP answer only
    // retracts what the LSP itself published.
    if (Length(AResults) = 0) and (Length(FResults) > 0)
      and SameText(FResFile, AFile) and (FResHash = AHash)
      and not SameText(FResSource, SrcLsp) then
    begin
      FSrcLsp := FSrcLsp + ', ignored: empty, keeping ' + FResSource;
      Exit;
    end;
    FSrcLsp := FSrcLsp + ', published';
    if FDiagNote <> '' then FSrcLsp := FSrcLsp + ' | declined: ' + FDiagNote;
    FResFile := AFile;
    FResHash := AHash;
    FResults := AResults;
    FResFromLsp := True;
    FResSource := SrcLsp;
    FResDiagSig := FDiagSig;
    FHintDirty := True;
  end;
end;

// --- yielding to other popups ----------------------------------------------
//
// The hint sits AT THE CARET and is a WS_EX_NOACTIVATE window, so when the
// code completion list opens at the same spot it ends up on top of it and
// covers the FIRST entry (tester report).
//
// The test is deliberately NOT a window-class name: the IDE's completion
// list, our own completion popup and whatever a third-party plugin shows
// would each need their own name, and those names change between IDE
// versions. What they all have in common is being a VISIBLE, top-level
// WS_POPUP window of this thread that overlaps the caret area - that is
// what we look for. Our own windows are excluded by name (VCL registers
// the window class under the form's class name).

type
  PPopupProbe = ^TPopupProbe;
  TPopupProbe = record
    Exclude: HWND;
    Area: TRect;
    Found: Boolean;
    Wnd: HWND;
    Cls: string;
  end;

function PopupProbeProc(AWnd: HWND; AParam: LPARAM): BOOL; stdcall;
const
  // ONLY the hint itself. Our own completion popup and the quick-fix
  // popup must NOT be excluded - the tester reported the hint covering
  // the first entry of OUR completion list too, so we yield to those
  // exactly like we yield to the IDE's.
  OurClasses: array[0..0] of string = ('TAutoImportHint');
var
  P: PPopupProbe;
  R, Tmp: TRect;
  Buf: array[0..63] of Char;
  Nm: string;
begin
  Result := True;   // keep enumerating
  P := PPopupProbe(AParam);
  if (AWnd = P.Exclude) or not IsWindowVisible(AWnd) then Exit;
  // Only POPUPs: the IDE main window and docked panes are not, and they
  // would otherwise match everything.
  if (GetWindowLong(AWnd, GWL_STYLE) and WS_POPUP) = 0 then Exit;
  if not GetWindowRect(AWnd, R) then Exit;
  if (R.Right <= R.Left) or (R.Bottom <= R.Top) then Exit;
  if not IntersectRect(Tmp, R, P.Area) then Exit;
  if GetClassName(AWnd, Buf, Length(Buf)) > 0 then
  begin
    Nm := Buf;
    for var C in OurClasses do
      if SameText(Nm, C) then Exit;
  end;
  P.Found := True;
  P.Wnd := AWnd;
  GetClassName(AWnd, Buf, Length(Buf));
  P.Cls := Buf;
  Result := False;   // found one - stop
end;

// Returns the blocking window (0 = none) plus its class, so the status
// window can name it.
function ForeignPopupAt(const AArea: TRect; AExclude: HWND;
  out ACls: string): HWND;
var
  Probe: TPopupProbe;
begin
  Result := 0;
  ACls := '';
  Probe.Exclude := AExclude;
  Probe.Area := AArea;
  Probe.Found := False;
  Probe.Wnd := 0;
  Probe.Cls := '';
  try
    EnumThreadWindows(GetCurrentThreadId, @PopupProbeProc, LPARAM(@Probe));
  except
    Probe.Found := False;
  end;
  if Probe.Found then
  begin
    Result := Probe.Wnd;
    ACls := Probe.Cls;
  end;
end;

procedure TAutoImportLive.UpdateHint;
var
  Pt, Target: TPoint;
  Fixes: TArray<TQuickFix>;
  Text: string;
  HintWnd, Blocker: HWND;
  BlockCls: string;
  Area: TRect;
begin
  // Bound to the ERROR LINE: shown only while the caret is on a line with
  // at least one quick fix (and that line is scrolled into view -
  // CaretScreenPos rejects carets outside the visible client area).
  if FindFixesAtCaretLine(Fixes) and CaretScreenPos(Pt) then
  begin
    if FHint = nil then
    begin
      FHint := TAutoImportHint.CreateHint;
      FHint.OnFix := RunFix;
    end;
    if Length(Fixes) = 1 then
      Text := Fixes[0].Caption
    else
      Text := Format('%d quick fixes', [Length(Fixes)]);
    FHint.SetInfo(Text);   // sizes the window - needed for the anchor

    // PRIMARY position: ABOVE the caret line. Code completion (the IDE's
    // and ours) drops DOWN from the caret, so up there nothing else
    // appears - which is the robust answer to "the hint covers the first
    // entry of the list", far better than trying to detect that list.
    // Only when there is no room above (caret on the first visible line)
    // do we fall back to the old spot below the caret.
    if not CaretAbovePos(FHint.Height, Target) then
      Target := Pt;

    // Safety net for the fallback case (and for a list the IDE flips
    // upwards near the bottom of the screen): if something else already
    // occupies our target rect, stay hidden.
    HintWnd := 0;
    if FHint.HandleAllocated then HintWnd := FHint.Handle;
    Area := TRect.Create(Target.X, Target.Y,
      Target.X + FHint.Width, Target.Y + FHint.Height);
    Blocker := ForeignPopupAt(Area, HintWnd, BlockCls);
    if Blocker <> 0 then
    begin
      // SELF-HEALING: a completion list is TRANSIENT. A window that keeps
      // matching for seconds is something else (a floating tool window,
      // or a list the IDE only hides logically) and must not disable the
      // hint for good - after HintBlockMaxMs the same window is ignored.
      if Blocker <> FBlockWnd then
      begin
        FBlockWnd := Blocker;
        FBlockSince := GetTickCount;
        FBlockCls := BlockCls;
      end;
      if GetTickCount - FBlockSince < HintBlockMaxMs then
      begin
        FHint.HideHint;
        Exit;   // the next tick brings it back once the list is gone
      end;
      FBlockCls := BlockCls + ' (ignored: persistent)';
    end
    else
    begin
      FBlockWnd := 0;
      FBlockCls := '';
    end;

    FHint.ShowNoActivateAt(Target);
  end
  else if FHint <> nil then
    FHint.HideHint;
end;

procedure TAutoImportLive.RunFix;
var
  Pt: TPoint;
  Fixes: TArray<TQuickFix>;
begin
  if not FindFixesAtCaretLine(Fixes) then Exit;
  if FHint <> nil then FHint.HideHint;
  var Popup := TQuickFixPopup.CreatePopup(FFile, Fixes);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
end;

procedure StartAutoImportLive;
begin
  GLiveShutdown := False;
  if GLive = nil then
    GLive := TAutoImportLive.Create;
end;

procedure StopAutoImportLive;
var
  Waited: Integer;
begin
  // Latch first so no NEW workers start, then wait for the running ones
  // and drain their queued TThread.Queue closures - both live inside this
  // package's code, which is about to be unmapped on uninstall. While
  // waiting, keep the sync queue moving (the closures are state-only).
  GLiveShutdown := True;
  Waited := 0;
  while (TInterlocked.CompareExchange(GLiveWorkers, 0, 0) <> 0)
    and (Waited < 15000) do
  begin
    CheckSynchronize(10);
    Inc(Waited, 10);
  end;
  CheckSynchronize(0);   // drain closures queued by the last worker
  FreeAndNil(GLive);
end;

function LiveFreshInfo(const AFile: string; out ACount: Integer): Boolean;
var
  Seen: TDictionary<string, Boolean>;
begin
  // Counts only DISTINCT add-unit identifiers: the consumer is the
  // "Add unit for identifier at cursor" menu entry, whose action handles
  // exactly that case - rename/uses-repair/header fixes would inflate the
  // count with work that action cannot perform.
  ACount := 0;
  Result := (GLive <> nil) and GLive.HasFresh(AFile);
  if not Result then Exit;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for var F in GLive.FResults do
      if (F.Kind = qfAddUnit) and not Seen.ContainsKey(UpperCase(F.Identifier)) then
      begin
        Seen.Add(UpperCase(F.Identifier), True);
        Inc(ACount);
      end;
  finally
    Seen.Free;
  end;
end;

procedure LiveReportErrorDiags(const AFile, AContent: string;
  const ADiags: TArray<TLspErrorDiag>; const ASource: string);
begin
  if GLive <> nil then
    GLive.ApplyExternalDiags(AFile, AContent, ADiags, ASource);
end;

function LiveShowFixAtCaret(const AFile: string): Boolean;
var
  Pt: TPoint;
  Fixes: TArray<TQuickFix>;
begin
  Result := False;
  if (GLive = nil) or not SameText(GLive.FFile, AFile) then Exit;
  if not GLive.FindFixesAtCaretLine(Fixes) then Exit;
  var Popup := TQuickFixPopup.CreatePopup(AFile, Fixes);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
  Result := True;
end;

function LiveHintBlocker: string;
begin
  Result := '';
  if GLive <> nil then Result := GLive.FBlockCls;
end;

procedure LiveSourceStats(out AStructure, ALsp: string);
begin
  AStructure := '';
  ALsp := '';
  if GLive = nil then Exit;
  AStructure := GLive.FSrcStructure;
  ALsp := GLive.FSrcLsp;
end;

procedure LiveDiagStats(out ASeen, AHandled: Integer; out ACodes: string);
begin
  ASeen := 0;
  AHandled := 0;
  ACodes := '';
  if GLive = nil then Exit;
  ASeen := GLive.FDiagSeen;
  AHandled := GLive.FDiagHandled;
  ACodes := GLive.FDiagCodes;
end;

function LiveDeclineNote: string;
begin
  Result := '';
  if GLive <> nil then Result := GLive.FDiagNote;
end;

function LiveResolveBusy: Boolean;
begin
  Result := (GLive <> nil)
    and (GLive.FResolving or GLive.FHasPending or GLive.FAnalysing);
end;

procedure LiveStatusInfo(out AFile: string; out AAnalysing, AResolving,
  AFromLsp, AFresh: Boolean; out AFixCount: Integer);
begin
  AFile := '';
  AAnalysing := False;
  AResolving := False;
  AFromLsp := False;
  AFresh := False;
  AFixCount := 0;
  if GLive = nil then Exit;
  AFile := GLive.FFile;
  AAnalysing := GLive.FAnalysing;
  AResolving := GLive.FResolving;
  AFromLsp := GLive.FResFromLsp;
  // "Fresh" = the published results belong to the buffer state we are
  // currently watching (same file AND same content hash).
  AFresh := (GLive.FResFile <> '') and SameText(GLive.FResFile, GLive.FFile)
    and (GLive.FResHash = GLive.FHash);
  AFixCount := Length(GLive.FResults);
end;

procedure LiveRefreshAfterCompile;
begin
  if GLive <> nil then
    GLive.FCompileRefreshPending := True;   // state only - tick acts on it
end;

function LiveFixesForLine(const AFile: string; ALine0: Integer): TArray<TQuickFix>;
begin
  Result := nil;
  if (GLive = nil) or not GLive.HasFresh(AFile) then Exit;
  for var F in GLive.FResults do
    if F.Line = ALine0 then
      Result := Result + [F];
end;

function LiveAllFixes(const AFile: string; out AFixes: TArray<TQuickFix>): Boolean;
begin
  AFixes := nil;
  Result := (GLive <> nil) and GLive.HasFresh(AFile);
  if Result then
    AFixes := Copy(GLive.FResults);
end;

procedure LiveResetResults;
begin
  if GLive = nil then Exit;
  // State only - no window operation (this can arrive from an IDE
  // notifier). The next poll tick re-arms the analysis for the active
  // buffer and hides the hint while there are no results.
  GLive.FResults := nil;
  GLive.FResSource := '';
  GLive.FResDiagSig := 0;
  GLive.FResFile := '';
  GLive.FResHash := 0;
  GLive.FResFromLsp := False;
  GLive.FHasPending := False;
  GLive.FFile := '';
  GLive.FHash := 0;
  GLive.FDirty := False;
end;

// ---------------------------------------------------------------------------
//  "Show all quick fixes" overview dialog
// ---------------------------------------------------------------------------

type
  TQuickFixListDialog = class(TForm)
  private
    FList: TListView;
    FFixes: TArray<TQuickFix>;
    FBtnGoto: TButton;
    FBtnClose: TButton;
    procedure DoDblClick(Sender: TObject);
    procedure DoGotoClick(Sender: TObject);
    procedure DoCloseClick(Sender: TObject);
  public
    HasChosen: Boolean;
    ChosenFix: TQuickFix;
    constructor CreateDialog(AOwner: TComponent;
      const AFixes: TArray<TQuickFix>);
  end;

constructor TQuickFixListDialog.CreateDialog(AOwner: TComponent;
  const AFixes: TArray<TQuickFix>);
var
  Col: TListColumn;
  Item: TListItem;
  Panel: TPanel;
begin
  inherited CreateNew(AOwner);
  Caption := 'Quick fixes in this unit';
  Width := 760;
  Height := 440;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;
  HasChosen := False;

  // Line-sorted copy; rows map back through Item.Data (sortable list).
  FFixes := Copy(AFixes);
  TArray.Sort<TQuickFix>(FFixes, TComparer<TQuickFix>.Construct(
    function(const L, R: TQuickFix): Integer
    begin
      Result := L.Line - R.Line;
      if Result = 0 then Result := L.Col - R.Col;
    end));

  FList := TListView.Create(Self);
  FList.Parent := Self;
  FList.Align := alClient;
  FList.AlignWithMargins := True;
  FList.ViewStyle := vsReport;
  FList.ReadOnly := True;
  FList.RowSelect := True;
  FList.OnDblClick := DoDblClick;
  Col := FList.Columns.Add; Col.Caption := 'Line';    Col.Width := 60;
    Col.Alignment := taRightJustify;
  Col := FList.Columns.Add; Col.Caption := 'Fix';     Col.Width := 320;
  Col := FList.Columns.Add; Col.Caption := 'Details'; Col.Width := 330;

  Panel := TPanel.Create(Self);
  Panel.Parent := Self;
  Panel.Align := alBottom;
  Panel.Height := 40;
  Panel.BevelOuter := bvNone;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := Panel;
  FBtnClose.Caption := '&Close';
  FBtnClose.Align := alRight;
  FBtnClose.AlignWithMargins := True;
  FBtnClose.Cancel := True;
  FBtnClose.OnClick := DoCloseClick;

  FBtnGoto := TButton.Create(Self);
  FBtnGoto.Parent := Panel;
  FBtnGoto.Caption := '&Go to && fix';
  FBtnGoto.Align := alRight;
  FBtnGoto.Width := 110;
  FBtnGoto.AlignWithMargins := True;
  FBtnGoto.Default := True;
  FBtnGoto.OnClick := DoGotoClick;

  FList.Items.BeginUpdate;
  try
    for var I := 0 to High(FFixes) do
    begin
      Item := FList.Items.Add;
      Item.Data := Pointer(NativeInt(I));
      Item.Caption := IntToStr(FFixes[I].Line + 1);
      Item.SubItems.Add(FFixes[I].Caption);
      case FFixes[I].Kind of
        qfAddUnit:
          Item.SubItems.Add('uses: ' + string.Join(', ', FFixes[I].UnitNames));
        qfRenameIdent, qfFixUsesName:
          Item.SubItems.Add('-> ' + FFixes[I].NewText);
        qfRemoveUses:
          Item.SubItems.Add('remove ' + FFixes[I].OldUnit);
      else
        Item.SubItems.Add('');
      end;
    end;
    if FList.Items.Count > 0 then
      FList.Items[0].Selected := True;
  finally
    FList.Items.EndUpdate;
  end;

  EnableListViewSorting(FList);
  EnableThemes(Self);
  PrepareDialog(Self, AOwner);
end;

procedure TQuickFixListDialog.DoGotoClick(Sender: TObject);
begin
  if FList.Selected = nil then Exit;
  var Idx := NativeInt(FList.Selected.Data);
  if (Idx < 0) or (Idx > High(FFixes)) then Exit;
  ChosenFix := FFixes[Idx];
  HasChosen := True;
  ModalResult := mrOk;
end;

procedure TQuickFixListDialog.DoDblClick(Sender: TObject);
begin
  DoGotoClick(Sender);
end;

procedure TQuickFixListDialog.DoCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure ShowAllQuickFixes;
var
  Ctx: TEditorContext;
  Fixes: TArray<TQuickFix>;
  Waited: Integer;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if (Ctx.FileName = '') or not SameText(ExtractFileExt(Ctx.FileName), '.pas') then
  begin
    ShowThemedMessage('Open a .pas unit first.');
    Exit;
  end;

  if not LiveAllFixes(Ctx.FileName, Fixes) then
  begin
    // No fresh analysis for the current buffer state yet: re-arm one (the
    // same state-only signal the post-compile refresh uses - the poll
    // tick acts on it) and wait briefly for the result.
    LiveRefreshAfterCompile;
    Waited := 0;
    Screen.Cursor := crHourGlass;
    try
      while (Waited < 8000) and not LiveAllFixes(Ctx.FileName, Fixes) do
      begin
        Sleep(100);
        Inc(Waited, 100);
        Application.ProcessMessages;   // live tick + queued results run here
      end;
    finally
      Screen.Cursor := crDefault;
    end;
    if not LiveAllFixes(Ctx.FileName, Fixes) then
    begin
      ShowThemedMessage('The live analysis produced no result for this unit yet - ' +
        'try again in a moment.');
      Exit;
    end;
  end;

  if Length(Fixes) = 0 then
  begin
    ShowThemedMessage('No quick fixes found in this unit.');
    Exit;
  end;

  var Dlg := TQuickFixListDialog.CreateDialog(Application.MainForm, Fixes);
  try
    Dlg.ShowModal;
    if Dlg.HasChosen then
    begin
      // Jump to the fix and open the regular chooser popup there - it
      // offers the unit choice / target section exactly like the caret
      // lightbulb does.
      Editor.GotoLocation(Ctx.FileName, Dlg.ChosenFix.Line, Dlg.ChosenFix.Col,
        Dlg.ChosenFix.TokenLen);
      Application.ProcessMessages;   // let the editor scroll + place the caret
      LiveShowFixAtCaret(Ctx.FileName);
    end;
  finally
    Dlg.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Public entry points
// ---------------------------------------------------------------------------

procedure WithWaitCursor(AProc: TProc);
begin
  Screen.Cursor := crHourGlass;
  try AProc; finally Screen.Cursor := crDefault; end;
end;

procedure AddUnitForIdentifierAtCursor;
var
  Ctx: TEditorContext;
  Missing: TArray<TMissingIdent>;
  Chosen: TMissingIdent;
  Found: Boolean;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if Ctx.FileName = '' then Exit;

  Missing := nil;
  WithWaitCursor(procedure begin Missing := GatherMissing(Ctx.FileName, nil); end);

  // Prefer the missing identifier at the caret; else the word under it.
  Found := False;
  for var M in Missing do
    if (M.Line + 1 = Ctx.Line) and SameText(M.Identifier, Ctx.WordAtCursor) then
    begin Chosen := M; Found := True; Break; end;
  if not Found then
    for var M in Missing do
      if SameText(M.Identifier, Ctx.WordAtCursor) then
      begin Chosen := M; Found := True; Break; end;

  if not Found then
  begin
    // Not flagged by LSP - offer it anyway if the index knows the word.
    var Hits := DedupeByUnitName(TUnitIndex.Instance.Lookup(Ctx.WordAtCursor));
    if Length(Hits) = 0 then
    begin
      ShowThemedMessage(Format('No unit found for "%s" (and it is not reported as ' +
        'an undeclared identifier).', [Ctx.WordAtCursor]));
      Exit;
    end;
    Chosen.Identifier := Ctx.WordAtCursor;
    Chosen.Line := Ctx.Line - 1;
    Chosen.Units := Hits;
    Chosen.Section := usInterface;
    // No diagnostic line at hand: mark generic when only generic
    // declarations exist (the word can only be used as "Name<...>").
    Chosen.GenericUse := True;
    for var H in Hits do
      if not H.IsGeneric then
      begin
        Chosen.GenericUse := False;
        Break;
      end;
  end;

  // One obvious choice -> apply straight away; else show the chooser popup.
  if Length(Chosen.Units) = 1 then
  begin
    if not ApplyOne(Ctx.FileName, Chosen.Units[0].UnitName, Chosen.Section) then
      ShowThemedMessage(Format('"%s" is already reachable in uses, or the clause ' +
        'could not be rewritten (e.g. IFDEFs inside it).',
        [Chosen.Units[0].UnitName]));
    Exit;
  end;

  // Several candidates -> the quick-fix chooser popup at the caret.
  var Disp := Chosen.Identifier;
  if Chosen.GenericUse then Disp := Disp + '<>';
  var F := Default(TQuickFix);
  F.Kind := qfAddUnit;
  F.Line := Chosen.Line;
  F.Identifier := Disp;
  F.Section := Chosen.Section;
  F.Caption := Format('Add unit for "%s"', [Disp]);
  for var H in Chosen.Units do
    F.UnitNames := F.UnitNames + [H.UnitName];

  var Pt: TPoint;
  var Popup := TQuickFixPopup.CreatePopup(Ctx.FileName, [F]);
  if CaretScreenPos(Pt) then Popup.ShowAt(Pt)
  else Popup.ShowAt(Mouse.CursorPos);
end;

procedure ResolveMissingUnits;
var
  Ctx: TEditorContext;
  Missing: TArray<TMissingIdent>;
begin
  if Editor = nil then Exit;
  Ctx := Editor.GetCurrentContext;
  if Ctx.FileName = '' then
  begin ShowThemedMessage('No active editor file.'); Exit; end;

  Missing := nil;
  WithWaitCursor(procedure begin Missing := GatherMissing(Ctx.FileName, nil); end);
  if Length(Missing) = 0 then
  begin
    ShowThemedMessage('No unresolved identifiers with a known unit were found.');
    Exit;
  end;
  var Dlg := TAutoImportDialog.CreateDialog(Application.MainForm, Ctx.FileName, Missing);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
