(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitIndex;

// Background-maintained index of public identifiers -> declaring unit(s),
// over every .pas reachable from the project's search paths (project
// sources + IDE global library/browsing paths + .dproj DCC_UnitSearchPath).
//
// Design
//  - A single worker thread parses files and (re)builds the index; the
//    main thread only ever hands it the list of source roots (gathered
//    via ToolsAPI, which must run on the main thread) and reads a
//    published, immutable snapshot under a lock.
//  - Only INTERFACE-section top-level declarations are indexed (types,
//    routines, consts, vars, enum members) - exactly the identifiers a
//    "Find Unit" needs. Class/record bodies are skipped.
//  - The parsed per-file data is cached to disk (keyed by path+mtime+size)
//    so a restart is fast; the worker re-parses only changed files and
//    refreshes periodically so the index stays live.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.SyncObjs;

type
  TFindUnitHit = record
    Identifier: string;   // the matched identifier (original case)
    UnitName: string;     // e.g. 'Vcl.Controls'
    Path: string;         // full path of the declaring .pas
    IsGeneric: Boolean;   // declared with type parameters ("TList<T>")
  end;

  /// <summary>Immutable, reference-counted view of the index. Lookups run
  ///  against it WITHOUT any lock (it is never mutated after creation and
  ///  stays alive while referenced), so a caller can hold one and scan it
  ///  on a background thread without blocking the UI or the worker.</summary>
  IUnitSnapshot = interface
    ['{2B7A9C31-6E4D-4F2A-9C1B-7D5E0A3F8B44}']
    function Lookup(const AIdentifier: string): TArray<TFindUnitHit>;
    function Search(const ASub: string; AMax: Integer): TArray<TFindUnitHit>;
    /// <summary>"Did you mean" candidates: known identifiers within
    ///  AMaxDist edit operations (incl. adjacent transposition) of AIdent,
    ///  best first, at most AMax DISTINCT identifiers (each with its first
    ///  declaring unit - use Lookup for the full unit list). Case is
    ///  ignored; an exact (case-insensitive) match is never returned.
    ///  O(IdentCount) - run on a background thread.</summary>
    function FuzzyIdentifiers(const AIdent: string; AMaxDist, AMax: Integer): TArray<TFindUnitHit>;
    /// <summary>Unit names within AMaxDist edit operations of AName
    ///  (e.g. 'Windwos' -> 'Winapi.Windows' is found via its last segment
    ///  as well as the full name), best first, at most AMax.</summary>
    function FuzzyUnitNames(const AName: string; AMaxDist, AMax: Integer): TArray<string>;
    /// <summary>True when a unit of exactly this (dotted) name is indexed.</summary>
    function HasUnit(const AUnitName: string): Boolean;
    /// <summary>Full .pas path of the indexed unit AUnitName (project
    ///  scope shadows global). False when the unit is not indexed.</summary>
    function TryGetUnitPath(const AUnitName: string; out APath: string): Boolean;
    /// <summary>True when the indexed unit has an initialization or
    ///  finalization section (load-time side effects - the uses cleanup
    ///  must never flag such a unit as removable). False for unknown
    ///  units.</summary>
    function HasInitCode(const AUnitName: string): Boolean;
    function UnitCount: Integer;
    function IdentCount: Integer;
  end;

  TUnitIndex = class
  private type
    TIndexedUnit = class
      Path: string;
      UnitName: string;
      MTime: TDateTime;
      Size: Int64;
      Idents: TArray<string>;
      HasInit: Boolean;   // unit has an initialization/finalization section
    end;
    TIndexWorker = class(TThread)
    private
      FOwner: TUnitIndex;
    protected
      procedure Execute; override;
    public
      constructor Create(AOwner: TUnitIndex);
    end;
  private
    FLock: TCriticalSection;
    FWake: TEvent;
    FWorker: TIndexWorker;
    // Sources (guarded by FLock) - set by the main thread. The GLOBAL set
    // (IDE library/browsing paths, i.e. RTL/VCL/3rd-party) is shared across
    // projects and changes almost never; the PROJECT set (project sources +
    // .dproj DCC_UnitSearchPath) changes while you edit. They are scanned
    // and cached separately so switching projects does not re-index the
    // huge, unchanging global tree.
    FGlobalDirs: TArray<string>;
    FProjectDirs: TArray<string>;
    FProjectFiles: TArray<string>;
    FGlobalCachePath: string;
    FProjectCachePath: string;
    // Published snapshot (the reference is swapped under FLock; the object
    // itself is immutable, so readers scan it lock-free).
    FSnapshot: IUnitSnapshot;
    FReady: Boolean;
    FStatus: string;
    // Worker-only state (no lock needed - single worker).
    FGlobalByPath: TObjectDictionary<string, TIndexedUnit>;
    FProjectByPath: TObjectDictionary<string, TIndexedUnit>;
    FGlobalKey: string;    // cache path the global dict was last loaded for
    FProjectKey: string;   // cache path the project dict was last loaded for
    FScanCycle: Integer;   // completed worker cycles (interlocked)
    class var FInstance: TUnitIndex;
    function CacheFileFor(const AKey: string): string;
    function GetSnapshot: IUnitSnapshot;
    procedure EnsureWorker;
    procedure PublishSnapshot;
    procedure LoadCache(const APath: string; ADict: TObjectDictionary<string, TIndexedUnit>);
    procedure SaveCache(const APath: string; ADict: TObjectDictionary<string, TIndexedUnit>);
    procedure CurrentSources(out AGlobalDirs, AProjectDirs, AProjectFiles: TArray<string>;
      out AGlobalCache, AProjectCache: string);
    procedure WorkerLoop;
    function RefreshScope(const ADirs, AFiles: TArray<string>;
      ADict: TObjectDictionary<string, TIndexedUnit>; const AScopeName: string): Boolean;
    procedure SetStatus(const S: string);
  public
    constructor Create;
    destructor Destroy; override;
    class function Instance: TUnitIndex;
    class procedure ReleaseInstance;

    /// <summary>MAIN THREAD: start building the GLOBAL scope (IDE library /
    ///  browsing paths) in the background immediately - no project needed.
    ///  Safe to call at plugin startup.</summary>
    procedure StartGlobalIndex;

    /// <summary>MAIN THREAD: (re)gather the source roots from the active
    ///  project and hand them to the worker; starts the worker on first
    ///  call. Cheap to call repeatedly (e.g. whenever the dialog opens).</summary>
    procedure RefreshSourcesFromEditor;

    function Ready: Boolean;
    function StatusLine: string;

    /// <summary>Number of completed worker scan cycles. Callers that just
    ///  requested a refresh (RefreshSourcesFromEditor) can wait for this
    ///  to advance to know the snapshot reflects the current sources.</summary>
    function ScanCycle: Integer;

    /// <summary>Take a reference to the current immutable snapshot (may be
    ///  nil before the first build). Scan it on any thread without a lock.</summary>
    function Snapshot: IUnitSnapshot;

    /// <summary>Exact (case-insensitive) lookup of AIdentifier.</summary>
    function Lookup(const AIdentifier: string): TArray<TFindUnitHit>;
    /// <summary>Substring search over identifiers (case-insensitive),
    ///  returning up to AMax (identifier, unit) hits.</summary>
    function Search(const ASub: string; AMax: Integer): TArray<TFindUnitHit>;
  end;

/// <summary>Parses the interface section of a .pas into its exported
///  identifiers; AHasInit reports an initialization/finalization section
///  (load-time side effects). Exposed for the console test suite.</summary>
function ParseUnit(const AFile: string; out AUnitName: string;
  out AHasInit: Boolean): TArray<string>;

/// <summary>0-based line of AIdent's DECLARATION inside AContent, or -1.
///  Prefers a real declaration form (type/const alias, routine header,
///  property, typed field/var) over a plain occurrence, and the
///  interface section over the implementation. Lets callers jump to a
///  symbol when the LSP has no definition for it (index fallback).</summary>
function FindDeclarationLine(const AContent, AIdent: string): Integer;

/// <summary>0-based line range of the routine that CONTAINS ALine0 -
///  from its header (procedure/function/constructor/destructor, including
///  a qualified or generic one) to its final "end;". False when the line
///  is not inside a routine body. Used by the rename scope "in current
///  method". Nested routines resolve to the INNERMOST one containing the
///  line.</summary>
function FindEnclosingRoutineRange(const AContent: string; ALine0: Integer;
  out AFirst, ALast: Integer): Boolean;

/// <summary>0-based line of AMemberName's declaration INSIDE the body of
///  type ATypeName, or -1. Class members (class procedure / class
///  function / class property, and their instance counterparts) are NOT
///  in the identifier index - the parser skips class bodies - so a
///  qualified use like "TMyClass.MyClassProc" needs this second step:
///  locate the type, then walk its body. Nested records are tracked, so
///  a member of an inner record does not end the search early.</summary>
function FindMemberDeclarationLine(const AContent, ATypeName,
  AMemberName: string): Integer;

/// <summary>Expands IDE path variables - $(BDS), $(BDSCOMMONDIR), and
///  above all USER-DEFINED ones like $(DXVCL) (Tools > Options >
///  Environment Variables, stored under the IDE registry key). Values may
///  reference further variables, so expansion repeats (bounded).
///  AVars supplies the user-defined values; pass nil to read them from
///  the registry. Anything still unresolved is left in place - callers
///  skip such directories.</summary>
function ExpandIdeVars(const S, ABdsRoot: string; AVars: TStrings): string;

/// <summary>Path variables that could NOT be resolved during the last
///  library-path scan, e.g. "$(DXVCL)" - directories behind them are not
///  indexed. Shown in the status window (this silently cost DevExpress
///  users their whole source index).</summary>
function UnresolvedPathVars: string;

/// <summary>The COMPILE-relevant directories: the IDE Library "Search
///  Path" (browsing paths deliberately EXCLUDED - the compiler never
///  looks there), the active project's DCC_UnitSearchPath entries, and
///  the project directory itself. Used to detect units the index found
///  via browsing paths only, which the compiler cannot reach.</summary>
function GatherCompileSearchDirs: TArray<string>;

/// <summary>All directories the INDEX scans for the global scope: the
///  IDE Library path PLUS the Browsing path (browsing sources are what
///  let navigation work without putting them on the compiler's path).</summary>
function GatherGlobalLibraryDirs(const ABdsRoot: string): TArray<string>;

implementation

uses
  System.IOUtils, System.StrUtils, System.Win.Registry, System.Types,
  System.Generics.Defaults, System.Math, Winapi.Windows,
  Expert.EditorHelperIntf, Delphi.FileEncoding;

const
  MaxFiles      = 40000;      // hard cap (32-bit process guard)
  MaxFileBytes  = 4 * 1024 * 1024;
  RefreshIntervalMs = 30000;  // background re-scan period
  // !!! BUMP THIS ON EVERY CHANGE TO ParseUnit !!!
  // The on-disk cache is incremental (path + mtime + size), and RTL /
  // third-party sources NEVER change - so a unit that an older, buggier
  // parser got wrong keeps its broken identifier list FOREVER unless the
  // magic invalidates the cache. That is not theory: the 08-30/09-01
  // parser fixes (variant-record 'case', multi-line enums, generics,
  // wrapped parameter lists) shipped WITHOUT a bump, and users kept
  // seeing "TcxButton / TStringGrid unknown" - which made uses-cleanup
  // offer units in use for REMOVAL, and left add-unit and "find original
  // symbol" blind for those types.
  IndexParserVersion = '08';
  //  08: parser fixes of 08-30/09-01 (case/enums/generics/parameters)
  //  07: IDE path variables ($(DXVCL)) expand
  CacheMagic    = 'RLUIDX' + IndexParserVersion;
  // Snapshot map entries are unit indexes with this flag bit set for
  // GENERIC declarations (MaxFiles stays far below the bit).
  GenericBit    = $40000000;

// ---------------------------------------------------------------------------
//  Lexical helpers
// ---------------------------------------------------------------------------

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '_']);
end;

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
end;

function IsIdent(const S: string): Boolean;
var I: Integer;
begin
  Result := (S <> '') and IsIdentStart(S[1]);
  if Result then
    for I := 2 to Length(S) do
      if not IsIdentChar(S[I]) then Exit(False);
end;

// Removes comments/directives and blanks string literals from one line,
// carrying the block-comment state (0 none, 1 '{', 2 '(*') across lines.
function CleanLine(const L: string; var ABlock: Integer): string;
var
  I, N: Integer;
  InStr: Boolean;
  SB: TStringBuilder;
begin
  N := Length(L);
  InStr := False;
  SB := TStringBuilder.Create(N);
  try
    I := 1;
    while I <= N do
    begin
      if ABlock = 1 then
      begin
        if L[I] = '}' then ABlock := 0;
        Inc(I);
        Continue;
      end;
      if ABlock = 2 then
      begin
        if (L[I] = '*') and (I < N) and (L[I + 1] = ')') then begin ABlock := 0; Inc(I, 2); Continue; end;
        Inc(I);
        Continue;
      end;
      if InStr then
      begin
        if L[I] = '''' then InStr := False;
        SB.Append(' ');
        Inc(I);
        Continue;
      end;
      case L[I] of
        '{': begin ABlock := 1; Inc(I); end;
        '(':
          if (I < N) and (L[I + 1] = '*') then begin ABlock := 2; Inc(I, 2); end
          else begin SB.Append('('); Inc(I); end;
        '/':
          if (I < N) and (L[I + 1] = '/') then Break   // line comment
          else begin SB.Append('/'); Inc(I); end;
        '''': begin InStr := True; SB.Append(' '); Inc(I); end;
      else
        SB.Append(L[I]); Inc(I);
      end;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function CountWord(const AUpper, AWord: string): Integer;
var P, AfterIdx: Integer;
begin
  Result := 0;
  P := Pos(AWord, AUpper);
  while P > 0 do
  begin
    var OkBefore := (P = 1) or not IsIdentChar(AUpper[P - 1]);
    AfterIdx := P + Length(AWord);
    var OkAfter := (AfterIdx > Length(AUpper)) or not IsIdentChar(AUpper[AfterIdx]);
    if OkBefore and OkAfter then Inc(Result);
    P := PosEx(AWord, AUpper, P + 1);
  end;
end;

// Blanks balanced '<...>' spans (generic parameter lists / constraints) so
// their content never feeds the keyword counting. Unbalanced '<' is left
// untouched (comparison operators, wrapped declarations).
function StripAngleSpans(const S: string): string;
var
  I, J, Depth, SpanStart: Integer;
begin
  Result := S;
  Depth := 0;
  SpanStart := 0;
  for I := 1 to Length(Result) do
    case Result[I] of
      '<':
        begin
          if Depth = 0 then SpanStart := I;
          Inc(Depth);
        end;
      '>':
        if Depth > 0 then
        begin
          Dec(Depth);
          if Depth = 0 then
            for J := SpanStart to I do
              Result[J] := ' ';
        end;
    end;
end;

function StartsWithWord(const AUpperTrimmed, AWord: string): Boolean;
begin
  Result := AUpperTrimmed.StartsWith(AWord) and
    ((Length(AUpperTrimmed) = Length(AWord)) or
     not IsIdentChar(AUpperTrimmed[Length(AWord) + 1]));
end;

// Adds every valid identifier in the comma-separated name list preceding a
// ':' or '=' (e.g. "A, B, C").
procedure AddNameList(const ANames: string; AList: TList<string>);
var Nm: string;
begin
  for Nm in ANames.Split([',']) do
  begin
    var N := Trim(Nm);
    if IsIdent(N) then AList.Add(N);
  end;
end;

// ---------------------------------------------------------------------------
//  The interface-section parser
// ---------------------------------------------------------------------------

// CHANGING ANYTHING IN HERE MEANS BUMPING IndexParserVersion - otherwise
// existing caches keep the results of the old parser for every file whose
// timestamp does not change (i.e. all of RTL/VCL and every library).
function ParseUnit(const AFile: string; out AUnitName: string;
  out AHasInit: Boolean): TArray<string>;
type
  TSect = (secNone, secType, secConst, secVar);
var
  Raw: string;
  Lines: TArray<string>;
  Idents: TList<string>;
  I, Block, PendingEnds, ImplIdx, InParams: Integer;
  Started, InEnum: Boolean;
  Sect: TSect;
  Code, Up, Tr, TrU: string;

  // '(' minus ')' in an already comment/string-cleaned line.
  function ParenBalance(const S: string): Integer;
  begin
    Result := 0;
    for var K := 1 to Length(S) do
      if S[K] = '(' then Inc(Result)
      else if S[K] = ')' then Dec(Result);
  end;

  procedure AddEnumMembers(const S: string);
  begin
    for var M in S.Split([',']) do
    begin
      var MM := Trim(M);
      var Sp := Pos(' ', MM);        // "a = 1"
      var Eq := Pos('=', MM);
      if Eq > 0 then MM := Trim(Copy(MM, 1, Eq - 1))
      else if Sp > 0 then MM := Trim(Copy(MM, 1, Sp - 1));
      if IsIdent(MM) then Idents.Add(MM);
    end;
  end;

begin
  AUnitName := '';
  AHasInit := False;
  Result := nil;
  try
    Raw := TDelphiFileEncoding.ReadAll(AFile);
  except
    Exit;
  end;
  Lines := Raw.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);

  Idents := TList<string>.Create;
  try
    Block := 0; PendingEnds := 0; Started := False; Sect := secNone;
    InEnum := False; ImplIdx := -1; InParams := 0;
    for I := 0 to High(Lines) do
    begin
      Code := CleanLine(Lines[I], Block);
      Tr := Trim(Code);
      if Tr = '' then Continue;
      TrU := UpperCase(Tr);

      if AUnitName = '' then
      begin
        if StartsWithWord(TrU, 'UNIT') then
        begin
          var Rest := Trim(Copy(Tr, 6, MaxInt));
          var P := Pos(';', Rest);
          if P > 0 then Rest := Copy(Rest, 1, P - 1);
          AUnitName := Trim(Rest);
        end;
        // fall through - 'unit' line has no identifiers to index
      end;

      if not Started then
      begin
        if TrU = 'INTERFACE' then Started := True;
        Continue;
      end;

      // End of the interface section.
      if TrU = 'IMPLEMENTATION' then
      begin
        ImplIdx := I;
        Break;
      end;

      // Inside a class/record/interface body: only track nesting.
      // NOTE: 'case' is deliberately NOT counted - a variant-record part
      // ("record case Tag of ...") shares the record's single 'end', so
      // counting it left PendingEnds stuck and swallowed everything after
      // the first variant record (Winapi.Windows died at LARGE_INTEGER).
      // Balanced '<...>' spans are blanked first: a generic method
      // constraint ("Get<T: record>(...)") must not count as a nested
      // record opener - that phantom END swallowed every declaration
      // after the class.
      if PendingEnds > 0 then
      begin
        Up := StripAngleSpans(TrU);
        PendingEnds := PendingEnds + CountWord(Up, 'RECORD')
                       - CountWord(Up, 'END');
        if PendingEnds < 0 then PendingEnds := 0;
        Continue;
      end;

      // Enum members continuing from a previous line ("TTyp = (A," ...).
      if InEnum then
      begin
        var CP := Pos(')', Tr);
        if CP > 0 then
        begin
          AddEnumMembers(Copy(Tr, 1, CP - 1));
          InEnum := False;
        end
        else
          AddEnumMembers(Tr);
        Continue;
      end;

      // Parameter list (or const-array value) continuing from a previous
      // line. Its identifiers are PARAMETER NAMES, not declarations - a
      // wrapped header 'procedure Foo(const AValue: X;' / '  const
      // AParams: Y);' made the next line flip Sect to secConst and
      // indexed 'AParams' ("Add unit for AParams" on a parameter!).
      if InParams > 0 then
      begin
        InParams := InParams + ParenBalance(Tr);
        if InParams < 0 then InParams := 0;
        Continue;
      end;

      // Section switches.
      if StartsWithWord(TrU, 'TYPE') then
      begin
        Sect := secType;
        Tr := Trim(Copy(Tr, 5, MaxInt)); TrU := UpperCase(Tr);
        if Tr = '' then Continue;
      end
      else if StartsWithWord(TrU, 'CONST') then
      begin
        Sect := secConst;
        Tr := Trim(Copy(Tr, 6, MaxInt)); TrU := UpperCase(Tr);
        if Tr = '' then Continue;
      end
      else if StartsWithWord(TrU, 'RESOURCESTRING') then
      begin
        Sect := secConst;
        Tr := Trim(Copy(Tr, 15, MaxInt)); TrU := UpperCase(Tr);
        if Tr = '' then Continue;
      end
      else if StartsWithWord(TrU, 'VAR') then
      begin
        Sect := secVar;
        Tr := Trim(Copy(Tr, 4, MaxInt)); TrU := UpperCase(Tr);
        if Tr = '' then Continue;
      end
      else if StartsWithWord(TrU, 'THREADVAR') then
      begin
        Sect := secVar;
        Tr := Trim(Copy(Tr, 10, MaxInt)); TrU := UpperCase(Tr);
        if Tr = '' then Continue;
      end;

      // Routines are indexed regardless of the current section.
      if StartsWithWord(TrU, 'PROCEDURE') or StartsWithWord(TrU, 'FUNCTION') then
      begin
        var Rest := Trim(Copy(Tr, Pos(' ', Tr) + 1, MaxInt));
        var CutP := Length(Rest) + 1;
        for var K := 1 to Length(Rest) do
          if CharInSet(Rest[K], ['(', ':', ';', ' ', #9, '<']) then begin CutP := K; Break; end;
        var Nm := Trim(Copy(Rest, 1, CutP - 1));
        if IsIdent(Nm) then
        begin
          // Generic routine: keep the marker (see the type branch below).
          if (CutP <= Length(Rest)) and (Rest[CutP] = '<') then
            Idents.Add(Nm + '<')
          else
            Idents.Add(Nm);
        end;
        // Wrapped parameter list: skip until the parens balance again.
        InParams := ParenBalance(Tr);
        if InParams < 0 then InParams := 0;
        Continue;
      end;

      case Sect of
        secType:
          begin
            var EqP := Pos('=', Tr);
            if EqP > 1 then
            begin
              var Name := Trim(Copy(Tr, 1, EqP - 1));
              // GENERIC type declarations: "TList<T> = class ..." - index
              // the base name (the user types "TList<Integer>", the E2003
              // token / dialog search is 'TList'). Without this the name
              // failed IsIdent, the type was never indexed AND the class
              // body below was scanned as top level (its methods leaked
              // into the index). Generic declarations carry a trailing '<'
              // MARKER in the raw ident list (stripped when the snapshot
              // map is built): quick fixes must not offer System.Classes'
              // non-generic TList for a "TList<Integer>" use site.
              var IsGenericDecl := False;
              var LtP := Pos('<', Name);
              if (LtP > 1) and Name.EndsWith('>') then
              begin
                Name := Trim(Copy(Name, 1, LtP - 1));
                IsGenericDecl := True;
              end;
              // Guard against "A = B" continuation lines with commas etc.
              if IsIdent(Name) then
              begin
                if IsGenericDecl then
                  Idents.Add(Name + '<')
                else
                  Idents.Add(Name);
                var RhsU := UpperCase(Trim(Copy(Tr, EqP + 1, MaxInt)));
                var Rhs := Trim(Copy(Tr, EqP + 1, MaxInt));
                // Enum: "= ( a, b, c )" - possibly spanning MANY lines
                // (one member per line); the continuation is collected by
                // the InEnum handler above.
                if (Rhs <> '') and (Rhs[1] = '(') then
                begin
                  var CP := Pos(')', Rhs);
                  if CP > 0 then
                    AddEnumMembers(Copy(Rhs, 2, CP - 2))
                  else
                  begin
                    AddEnumMembers(Copy(Rhs, 2, MaxInt));
                    InEnum := True;
                  end;
                end
                // Body opener: class / record / object / interface.
                else
                begin
                  var Core := RhsU;
                  if StartsWithWord(Core, 'PACKED') then Core := Trim(Copy(Core, 8, MaxInt));
                  var IsClass := StartsWithWord(Core, 'CLASS') and not Core.StartsWith('CLASS OF');
                  var IsObj  := StartsWithWord(Core, 'OBJECT');
                  var IsRec  := StartsWithWord(Core, 'RECORD');
                  var IsIntf := StartsWithWord(Core, 'INTERFACE') or StartsWithWord(Core, 'DISPINTERFACE');
                  if IsClass or IsObj or IsRec or IsIntf then
                  begin
                    // A trailing ';' with no 'end' is a forward / short decl
                    // (e.g. "TFoo = class;", "TFoo = class(TBar);").
                    var Fwd := RhsU.EndsWith(';') and (CountWord(RhsU, 'END') = 0);
                    if not Fwd then
                    begin
                      // Ends expected to close the body = same-line record
                      // openers plus the primary non-record body, minus any
                      // 'end' already on this line. ('case' variant parts
                      // have no own 'end' - see the note above.)
                      var Opens := CountWord(RhsU, 'RECORD');
                      if IsClass or IsObj or IsIntf then Inc(Opens);
                      PendingEnds := Opens - CountWord(RhsU, 'END');
                      if PendingEnds < 0 then PendingEnds := 0;
                    end;
                  end;
                end;
              end;
            end;
          end;
        secConst:
          begin
            var Cut := Length(Tr) + 1;
            for var K := 1 to Length(Tr) do
              if CharInSet(Tr[K], ['=', ':']) then begin Cut := K; Break; end;
            var Nm := Trim(Copy(Tr, 1, Cut - 1));
            if IsIdent(Nm) then Idents.Add(Nm);
          end;
        secVar:
          begin
            var Cp := Pos(':', Tr);
            if Cp > 0 then AddNameList(Copy(Tr, 1, Cp - 1), Idents);
          end;
      end;

      // Any section line that leaves a '(' open continues on the next
      // line with NON-declarations: 'TNotify = procedure(const ASender:'
      // (parameter names) or 'C: TArr = (' (const values). Enums keep
      // their own richer state (InEnum collects the members).
      if not InEnum then
      begin
        var Bal := ParenBalance(Tr);
        if Bal > 0 then InParams := Bal;
      end;
    end;

    // Side-effect detection: does the unit run code at load/unload time?
    // (Feeds the uses-cleanup - such units must never be flagged unused.)
    if ImplIdx >= 0 then
      for I := ImplIdx + 1 to High(Lines) do
      begin
        Code := UpperCase(Trim(CleanLine(Lines[I], Block)));
        if (Code = 'INITIALIZATION') or (Code = 'FINALIZATION') then
        begin
          AHasInit := True;
          Break;
        end;
      end;

    Result := Idents.ToArray;
  finally
    Idents.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Source-root gathering (MAIN THREAD - uses ToolsAPI / registry)
// ---------------------------------------------------------------------------

function FindBdsRoot: string;

  function TryHive(ARootKey: HKEY): string;
  var
    Reg: TRegistry;
    Versions: TStringList;
    V, Best: string;
    BestNum, Num: Double;
    FS: TFormatSettings;
  begin
    Result := ''; Best := ''; BestNum := -1;
    FS := TFormatSettings.Invariant;
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := ARootKey;
      if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS') then
      begin
        Versions := TStringList.Create;
        try
          Reg.GetKeyNames(Versions);
          for V in Versions do
            if TryStrToFloat(V, Num, FS) and (Num > BestNum) then begin BestNum := Num; Best := V; end;
        finally
          Versions.Free;
        end;
        Reg.CloseKey;
      end;
      if (Best <> '') and Reg.OpenKeyReadOnly('Software\Embarcadero\BDS\' + Best) then
      begin
        if Reg.ValueExists('RootDir') then Result := Reg.ReadString('RootDir');
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
  end;

begin
  Result := TryHive(HKEY_CURRENT_USER);
  if Result = '' then Result := TryHive(HKEY_LOCAL_MACHINE);
end;

var
  GUnresolvedVars: string;
  GUnresolvedLock: TObject = nil;

function UnresolvedPathVars: string;
begin
  Result := GUnresolvedVars;
end;

procedure NoteUnresolved(const AVar: string);
const
  // MSBuild PROJECT properties, not IDE path variables: a dproj entry
  // like "myDir;$(DCC_UnitSearchPath)" simply appends the inherited
  // value. Nothing is missing - do not alarm the user about them.
  Benign: array[0..5] of string = ('$(DCC_', '$(ProjectDir)', '$(ProjectPath)',
    '$(OutputDir)', '$(Configuration)', '$(MSBuild');
begin
  for var B in Benign do
    if StartsText(B, AVar) then Exit;
  if Pos(AVar, GUnresolvedVars) > 0 then Exit;
  if GUnresolvedVars = '' then GUnresolvedVars := AVar
  else GUnresolvedVars := GUnresolvedVars + ', ' + AVar;
end;

// Reads the IDE's user-defined environment variables (Tools > Options >
// Environment Variables) - this is where $(DXVCL) & friends live.
procedure ReadIdeEnvVars(ADest: TStrings);
var
  Reg: TRegistry;
  Names: TStringList;
  VerKey: string;
begin
  if ADest = nil then Exit;
  Reg := TRegistry.Create(KEY_READ);
  Names := TStringList.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    VerKey := '';
    if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS') then
    begin
      Reg.GetKeyNames(Names);
      var BestNum: Double := -1;
      var FS := TFormatSettings.Invariant;
      for var V in Names do
      begin
        var Nn: Double;
        if TryStrToFloat(V, Nn, FS) and (Nn > BestNum) then
        begin
          BestNum := Nn;
          VerKey := V;
        end;
      end;
      Reg.CloseKey;
    end;
    if VerKey = '' then Exit;
    Names.Clear;
    if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS\' + VerKey +
      '\Environment Variables') then
    begin
      Reg.GetValueNames(Names);
      for var N in Names do
        try
          ADest.Values[N] := Reg.ReadString(N);
        except
        end;
      Reg.CloseKey;
    end;
  finally
    Names.Free;
    Reg.Free;
  end;
end;

function ExpandIdeVars(const S, ABdsRoot: string; AVars: TStrings): string;
const
  MaxDepth = 8;
var
  Owned: TStringList;
  Vars: TStrings;
  Depth, P, Q: Integer;
  Name, Value: string;
begin
  Result := S;
  if Pos('$(', Result) = 0 then Exit;

  Owned := nil;
  try
    Vars := AVars;
    if Vars = nil then
    begin
      Owned := TStringList.Create;
      ReadIdeEnvVars(Owned);
      Vars := Owned;
    end;

    // A variable value can itself contain variables - repeat, bounded.
    for Depth := 1 to MaxDepth do
    begin
      P := Pos('$(', Result);
      if P = 0 then Break;
      Q := P;
      while (Q <= Length(Result)) and (Result[Q] <> ')') do Inc(Q);
      if Q > Length(Result) then Break;              // unterminated
      Name := Copy(Result, P + 2, Q - P - 2);
      if Name = '' then Break;

      Value := '';
      if SameText(Name, 'BDS') then Value := ABdsRoot
      else if SameText(Name, 'BDSLIB') then
        Value := IncludeTrailingPathDelimiter(ABdsRoot) + 'lib'
      else if SameText(Name, 'BDSINCLUDE') then
        Value := IncludeTrailingPathDelimiter(ABdsRoot) + 'include'
      else if SameText(Name, 'Platform') then Value := 'Win32'
      else if SameText(Name, 'Config') then Value := 'Release'
      else
      begin
        // User-defined IDE variable first, then the process environment
        // (the IDE exports several of its own, e.g. BDSCOMMONDIR).
        Value := Vars.Values[Name];
        if Value = '' then Value := GetEnvironmentVariable(Name);
      end;

      if Value = '' then
      begin
        // Unresolvable - report it and stop (the caller drops the dir).
        NoteUnresolved('$(' + Name + ')');
        Break;
      end;
      Result := Copy(Result, 1, P - 1) + Value + Copy(Result, Q + 1, MaxInt);
    end;
  finally
    Owned.Free;
  end;
end;

function GatherIdeLibraryDirsEx(const ABdsRoot: string;
  AIncludeBrowsing: Boolean): TArray<string>;

  function ReadPath(const ASubKey, AValue: string): string;
  var Reg: TRegistry;
  begin
    Result := '';
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly(ASubKey) and Reg.ValueExists(AValue) then Result := Reg.ReadString(AValue);
      Reg.CloseKey;
    finally
      Reg.Free;
    end;
  end;

var
  Raw, Dir, VerKey: string;
  Dirs: TList<string>;
  EnvVars: TStringList;
begin
  Result := nil;
  if ABdsRoot = '' then Exit;
  GUnresolvedVars := '';
  Dirs := TList<string>.Create;
  EnvVars := TStringList.Create;
  try
    // Read the user-defined IDE variables ONCE for the whole scan.
    ReadIdeEnvVars(EnvVars);
    var Reg := TRegistry.Create(KEY_READ);
    VerKey := '';
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly('Software\Embarcadero\BDS') then
      begin
        var Vers := TStringList.Create;
        try
          Reg.GetKeyNames(Vers);
          var BestNum: Double := -1; var FS := TFormatSettings.Invariant;
          for var V in Vers do
          begin
            var Nn: Double;
            if TryStrToFloat(V, Nn, FS) and (Nn > BestNum) then begin BestNum := Nn; VerKey := V; end;
          end;
        finally
          Vers.Free;
        end;
        Reg.CloseKey;
      end;
    finally
      Reg.Free;
    end;
    if VerKey = '' then Exit;
    var Raws: TArray<string> :=
      [ReadPath('Software\Embarcadero\BDS\' + VerKey + '\Library\Win32', 'Search Path')];
    if AIncludeBrowsing then
      Raws := Raws +
        [ReadPath('Software\Embarcadero\BDS\' + VerKey + '\Library\Win32', 'Browsing Path')];
    for Raw in Raws do
      for Dir in Raw.Split([';'], TStringSplitOptions.ExcludeEmpty) do
      begin
        var D := ExpandIdeVars(Trim(Dir), ABdsRoot, EnvVars);
        if Pos('$(', D) > 0 then Continue;   // still unresolved - see
                                             // UnresolvedPathVars
        D := ExcludeTrailingPathDelimiter(D);
        if TDirectory.Exists(D) and not Dirs.Contains(D) then Dirs.Add(D);
      end;
    Result := Dirs.ToArray;
  finally
    EnvVars.Free;
    Dirs.Free;
  end;
end;

// Library + browsing dirs merged - the INDEXING scope (we want to index
// everything the user can see, not just what compiles).
function GatherGlobalLibraryDirs(const ABdsRoot: string): TArray<string>;
begin
  Result := GatherIdeLibraryDirsEx(ABdsRoot, True);
end;

function GatherDprojSearchDirs: TArray<string>;
var
  Dproj, Content, Inner, DprojDir, Ent, D: string;
  Dirs: TList<string>;
  EnvVars: TStringList;
  P1, P2: Integer;
begin
  Result := nil;
  if Editor = nil then Exit;
  Dproj := Editor.GetCurrentProjectDproj;
  if (Dproj = '') or not TFile.Exists(Dproj) then Exit;
  DprojDir := ExtractFilePath(Dproj);
  try Content := TFile.ReadAllText(Dproj); except Exit; end;
  Dirs := TList<string>.Create;
  EnvVars := TStringList.Create;
  try
    ReadIdeEnvVars(EnvVars);
    P1 := Pos('<DCC_UnitSearchPath>', Content);
    while P1 > 0 do
    begin
      P2 := PosEx('</DCC_UnitSearchPath>', Content, P1);
      if P2 = 0 then Break;
      Inner := Copy(Content, P1 + Length('<DCC_UnitSearchPath>'), P2 - P1 - Length('<DCC_UnitSearchPath>'));
      for Ent in Inner.Split([';'], TStringSplitOptions.ExcludeEmpty) do
      begin
        var E := Trim(Ent);
        if E = '' then Continue;
        // Project search paths carry the same IDE variables as the
        // library path ($(DXVCL) and friends) - expand before use.
        if Pos('$(', E) > 0 then
        begin
          E := ExpandIdeVars(E, FindBdsRoot, EnvVars);
          if Pos('$(', E) > 0 then Continue;   // still unresolved
        end;
        if TPath.IsRelativePath(E) then D := TPath.GetFullPath(TPath.Combine(DprojDir, E)) else D := E;
        D := ExcludeTrailingPathDelimiter(D);
        if TDirectory.Exists(D) and not Dirs.Contains(D) then Dirs.Add(D);
      end;
      P1 := PosEx('<DCC_UnitSearchPath>', Content, P2);
    end;
    Result := Dirs.ToArray;
  finally
    EnvVars.Free;
    Dirs.Free;
  end;
end;

function FindDeclarationLine(const AContent, AIdent: string): Integer;
var
  Lines: TArray<string>;
  ImplLine, I, P, Score, BestScore, BestLine: Integer;
  Raw, L, U, UIdent, Rest: string;

  // Whole-word position of AIdent in AUpperLine (1-based), else 0.
  function WordPos(const AUpperLine: string): Integer;
  var
    Q, AfterIdx: Integer;
  begin
    Result := 0;
    Q := Pos(UIdent, AUpperLine);
    while Q > 0 do
    begin
      AfterIdx := Q + Length(UIdent);
      if ((Q = 1) or not IsIdentChar(AUpperLine[Q - 1]))
        and ((AfterIdx > Length(AUpperLine)) or not IsIdentChar(AUpperLine[AfterIdx])) then
        Exit(Q);
      Q := PosEx(UIdent, AUpperLine, Q + 1);
    end;
  end;

begin
  Result := -1;
  if (AContent = '') or not IsIdent(AIdent) then Exit;
  Lines := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  UIdent := UpperCase(AIdent);

  ImplLine := MaxInt;
  for I := 0 to High(Lines) do
    if SameText(Trim(Lines[I]), 'implementation') then
    begin
      ImplLine := I;
      Break;
    end;

  BestScore := 0;
  BestLine := -1;
  for I := 0 to High(Lines) do
  begin
    Raw := Lines[I];
    // Cheap comment guard: full-line comments and a trailing '//'.
    L := Trim(Raw);
    if L.StartsWith('//') or L.StartsWith('{') or L.StartsWith('(*') then Continue;
    P := Pos('//', L);
    if P > 0 then L := TrimRight(Copy(L, 1, P - 1));
    if L = '' then Continue;

    U := UpperCase(L);
    P := WordPos(U);
    if P = 0 then Continue;

    Score := 1;   // plain occurrence
    Rest := TrimLeft(Copy(L, P + Length(AIdent), MaxInt));
    if StartsWithWord(U, 'PROPERTY ') then
    begin
      if WordPos(UpperCase(Copy(L, Length('property ') + 1, MaxInt))) = 1 then
        Score := 3;
    end
    else if StartsWithWord(U, 'FUNCTION') or StartsWithWord(U, 'PROCEDURE')
         or StartsWithWord(U, 'CONSTRUCTOR') or StartsWithWord(U, 'DESTRUCTOR')
         or StartsWithWord(U, 'CLASS FUNCTION') or StartsWithWord(U, 'CLASS PROCEDURE') then
    begin
      // 'function Foo(' / 'procedure TBar.Foo;' - the name must follow
      // the keyword (or the qualifying class name).
      if P > 1 then Score := 3;
    end
    else if P = 1 then
    begin
      // Line STARTS with the identifier: 'TFoo = class', 'Bar: Integer;'
      if Rest.StartsWith('=') then Score := 3
      else if Rest.StartsWith(':') then Score := 2
      else if Rest.StartsWith('<') then Score := 3;   // generic type decl
    end;

    // Interface-section declarations win over implementation ones.
    if I < ImplLine then Inc(Score, 4);

    if Score > BestScore then
    begin
      BestScore := Score;
      BestLine := I;
    end;
  end;
  Result := BestLine;
end;

function FindEnclosingRoutineRange(const AContent: string; ALine0: Integer;
  out AFirst, ALast: Integer): Boolean;
var
  Lines: TArray<string>;
  I, Depth, HdrLine: Integer;
  L, U: string;

  function Clean(const S: string): string;
  var
    P: Integer;
  begin
    Result := Trim(S);
    if Result.StartsWith('//') then Exit('');
    P := Pos('//', Result);
    if P > 0 then Result := TrimRight(Copy(Result, 1, P - 1));
  end;

  function IsHeader(const AUpper: string): Boolean;
  begin
    Result := StartsWithWord(AUpper, 'PROCEDURE')
      or StartsWithWord(AUpper, 'FUNCTION')
      or StartsWithWord(AUpper, 'CONSTRUCTOR')
      or StartsWithWord(AUpper, 'DESTRUCTOR')
      or StartsWithWord(AUpper, 'CLASS PROCEDURE')
      or StartsWithWord(AUpper, 'CLASS FUNCTION')
      or StartsWithWord(AUpper, 'OPERATOR');
  end;

begin
  Result := False;
  AFirst := -1;
  ALast := -1;
  if AContent = '' then Exit;
  Lines := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;

  // Nearest header at or above the line. A FORWARD declaration (header
  // ending in ';' with no body) is skipped by the body walk below.
  HdrLine := -1;
  for I := ALine0 downto 0 do
  begin
    L := Clean(Lines[I]);
    if L = '' then Continue;
    if IsHeader(UpperCase(L)) then
    begin
      HdrLine := I;
      Break;
    end;
  end;
  if HdrLine < 0 then Exit;

  // Walk from the header to its final "end;".
  Depth := 0;
  for I := HdrLine to High(Lines) do
  begin
    L := Clean(Lines[I]);
    if L = '' then Continue;
    U := UpperCase(L);
    if StartsWithWord(U, 'BEGIN') or StartsWithWord(U, 'TRY')
      or StartsWithWord(U, 'CASE') then
      Inc(Depth)
    else if (Depth > 0) and ((U = 'END;') or (U = 'END')) then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        // The line must actually lie inside what we walked.
        if (ALine0 >= HdrLine) and (ALine0 <= I) then
        begin
          AFirst := HdrLine;
          ALast := I;
          Result := True;
        end;
        Exit;
      end;
    end
    else if (Depth = 0) and (I > HdrLine) and IsHeader(U) then
      Exit;   // next routine started - the first one had no body
  end;
end;

function FindMemberDeclarationLine(const AContent, ATypeName,
  AMemberName: string): Integer;
var
  Lines: TArray<string>;
  I, Depth, Best, Score, BestScore: Integer;
  L, U, UMember, Rest: string;
  InBody: Boolean;

  // Whole-word position of AWord in AUpper (1-based), else 0.
  function WordPos(const AUpper, AWord: string): Integer;
  var
    Q, AfterIdx: Integer;
  begin
    Result := 0;
    Q := Pos(AWord, AUpper);
    while Q > 0 do
    begin
      AfterIdx := Q + Length(AWord);
      if ((Q = 1) or not IsIdentChar(AUpper[Q - 1]))
        and ((AfterIdx > Length(AUpper)) or not IsIdentChar(AUpper[AfterIdx])) then
        Exit(Q);
      Q := PosEx(AWord, AUpper, Q + 1);
    end;
  end;

  // Leading identifier of S ('' when it does not start with one).
  function FirstIdent(const S: string): string;
  var
    Q: Integer;
  begin
    Q := 1;
    while (Q <= Length(S)) and IsIdentChar(S[Q]) do Inc(Q);
    Result := Copy(S, 1, Q - 1);
  end;

  // Strips a line comment and trims - the cheap cleanup used below.
  function Clean(const S: string): string;
  var
    P: Integer;
  begin
    Result := Trim(S);
    if Result.StartsWith('//') then Exit('');
    P := Pos('//', Result);
    if P > 0 then Result := TrimRight(Copy(Result, 1, P - 1));
  end;

begin
  Result := -1;
  if (AContent = '') or not IsIdent(ATypeName) or not IsIdent(AMemberName) then Exit;
  Lines := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  UMember := UpperCase(AMemberName);

  InBody := False;
  Depth := 0;
  Best := -1;
  BestScore := 0;
  for I := 0 to High(Lines) do
  begin
    L := Clean(Lines[I]);
    if L = '' then Continue;
    U := UpperCase(L);

    if not InBody then
    begin
      // "TMyClass = class(...)" / "= record" / "= interface" - but NOT a
      // forward declaration ("TMyClass = class;"), which has no body.
      // A nested declaration carries the section keyword on the same
      // line ("type TInner = record").
      if StartsWithWord(U, 'TYPE') then
      begin
        L := TrimLeft(Copy(L, 5, MaxInt));
        U := UpperCase(L);
      end;
      if WordPos(U, UpperCase(ATypeName)) <> 1 then Continue;
      Rest := TrimLeft(Copy(L, Length(ATypeName) + 1, MaxInt));
      // generic declaration: skip the parameter list
      if Rest.StartsWith('<') and (Pos('>', Rest) > 0) then
        Rest := TrimLeft(Copy(Rest, Pos('>', Rest) + 1, MaxInt));
      if not Rest.StartsWith('=') then Continue;
      Rest := UpperCase(TrimLeft(Copy(Rest, 2, MaxInt)));
      if Rest.StartsWith('PACKED ') then Rest := TrimLeft(Copy(Rest, 8, MaxInt));
      if not (StartsWithWord(Rest, 'CLASS') or StartsWithWord(Rest, 'RECORD')
        or StartsWithWord(Rest, 'OBJECT') or StartsWithWord(Rest, 'INTERFACE')
        or StartsWithWord(Rest, 'DISPINTERFACE')) then Continue;
      if Rest.EndsWith(';') and (CountWord(Rest, 'END') = 0) then Continue;  // forward
      InBody := True;
      Depth := 1;
      Continue;
    end;

    // Inside the body: track nesting. Only RECORD opens a further level
    // here - counting CLASS would misfire on "class procedure".
    Depth := Depth + CountWord(StripAngleSpans(U), 'RECORD')
             - CountWord(StripAngleSpans(U), 'END');
    if Depth <= 0 then Break;

    // Member declaration forms. The declared NAME is the first
    // identifier after the keyword - searching the whole line would
    // match an accessor mention ("property Total ... read FSize").
    Score := 0;
    Rest := U;
    if StartsWithWord(Rest, 'CLASS') then
      Rest := TrimLeft(Copy(Rest, 6, MaxInt));   // 'class procedure/var/...'
    if StartsWithWord(Rest, 'PROCEDURE') or StartsWithWord(Rest, 'FUNCTION')
      or StartsWithWord(Rest, 'PROPERTY') or StartsWithWord(Rest, 'CONSTRUCTOR')
      or StartsWithWord(Rest, 'DESTRUCTOR') then
    begin
      Rest := TrimLeft(Copy(Rest, Pos(' ', Rest) + 1, MaxInt));
      if SameText(FirstIdent(Rest), AMemberName) then Score := 3;
    end
    else
    begin
      if StartsWithWord(Rest, 'VAR') then
        Rest := TrimLeft(Copy(Rest, 4, MaxInt));      // 'class var FCount:'
      if SameText(FirstIdent(Rest), AMemberName) then
      begin
        // A field/class var declaration is "Name[, Name2] : Type".
        Rest := TrimLeft(Copy(Rest, Length(AMemberName) + 1, MaxInt));
        if Rest.StartsWith(':') or Rest.StartsWith(',') then Score := 2;
      end;
    end;

    if Score > BestScore then
    begin
      BestScore := Score;
      Best := I;
      if Score = 3 then Break;   // exact declaration form - done
    end;
  end;
  Result := Best;
end;

function GatherCompileSearchDirs: TArray<string>;
var
  Dirs: TList<string>;

  procedure Add(const AList: TArray<string>);
  begin
    for var D in AList do
      if not Dirs.Contains(D) then Dirs.Add(D);
  end;

begin
  Dirs := TList<string>.Create;
  try
    Add(GatherIdeLibraryDirsEx(FindBdsRoot, False));   // library, NO browsing
    Add(GatherDprojSearchDirs);
    if Editor <> nil then
    begin
      var Dproj := Editor.GetCurrentProjectDproj;
      if Dproj <> '' then
      begin
        var D := ExcludeTrailingPathDelimiter(ExtractFilePath(Dproj));
        if (D <> '') and not Dirs.Contains(D) then Dirs.Add(D);
      end;
    end;
    Result := Dirs.ToArray;
  finally
    Dirs.Free;
  end;
end;

function HashKey(const S: string): string;
var I: Integer; H: Cardinal;
begin
  H := 2166136261;
  for I := 1 to Length(S) do begin H := H xor Ord(S[I]); H := H * 16777619; end;
  Result := IntToHex(H, 8);
end;

// ---------------------------------------------------------------------------
//  TUnitSnapshot - immutable, lock-free readable view
// ---------------------------------------------------------------------------

type
  TUnitSnapshot = class(TInterfacedObject, IUnitSnapshot)
  private
    FUnitName: TArray<string>;
    FUnitPath: TArray<string>;
    FUnitHasInit: TArray<Boolean>;
    FKeysUpper: TArray<string>;    // sorted UPPER identifiers
    FDisplay: TArray<string>;      // parallel original-case
    FMap: TDictionary<string, TArray<Integer>>;
  public
    constructor Create(const AUnitName, AUnitPath, AKeysUpper, ADisplay: TArray<string>;
      const AUnitHasInit: TArray<Boolean>;
      AMap: TDictionary<string, TArray<Integer>>);
    destructor Destroy; override;
    function Lookup(const AIdentifier: string): TArray<TFindUnitHit>;
    function Search(const ASub: string; AMax: Integer): TArray<TFindUnitHit>;
    function FuzzyIdentifiers(const AIdent: string; AMaxDist, AMax: Integer): TArray<TFindUnitHit>;
    function FuzzyUnitNames(const AName: string; AMaxDist, AMax: Integer): TArray<string>;
    function HasUnit(const AUnitName: string): Boolean;
    function TryGetUnitPath(const AUnitName: string; out APath: string): Boolean;
    function HasInitCode(const AUnitName: string): Boolean;
    function UnitCount: Integer;
    function IdentCount: Integer;
  end;

// Optimal-string-alignment edit distance (Levenshtein + adjacent
// transposition, the classic typo operations) with a cutoff: returns
// AMax + 1 as soon as the distance must exceed AMax. Inputs are expected
// to be same-cased by the caller.
function BoundedEditDistance(const A, B: string; AMax: Integer): Integer;
var
  LenA, LenB, I, J, Cost, RowMin: Integer;
  Prev2, Prev, Curr, Tmp: TArray<Integer>;
begin
  LenA := Length(A);
  LenB := Length(B);
  if Abs(LenA - LenB) > AMax then Exit(AMax + 1);
  if LenA = 0 then Exit(LenB);
  if LenB = 0 then Exit(LenA);

  SetLength(Prev2, LenB + 1);
  SetLength(Prev, LenB + 1);
  SetLength(Curr, LenB + 1);
  for J := 0 to LenB do Prev[J] := J;

  for I := 1 to LenA do
  begin
    Curr[0] := I;
    RowMin := Curr[0];
    for J := 1 to LenB do
    begin
      if A[I] = B[J] then Cost := 0 else Cost := 1;
      Curr[J] := Prev[J] + 1;                                   // deletion
      if Prev[J - 1] + Cost < Curr[J] then Curr[J] := Prev[J - 1] + Cost; // subst
      if Curr[J - 1] + 1 < Curr[J] then Curr[J] := Curr[J - 1] + 1;      // insertion
      if (I > 1) and (J > 1) and (A[I] = B[J - 1]) and (A[I - 1] = B[J])
         and (Prev2[J - 2] + 1 < Curr[J]) then
        Curr[J] := Prev2[J - 2] + 1;                            // transposition
      if Curr[J] < RowMin then RowMin := Curr[J];
    end;
    if RowMin > AMax then Exit(AMax + 1);   // cannot get better - bail out
    Tmp := Prev2; Prev2 := Prev; Prev := Curr; Curr := Tmp;
  end;
  Result := Prev[LenB];
  if Result > AMax then Result := AMax + 1;
end;

constructor TUnitSnapshot.Create(const AUnitName, AUnitPath, AKeysUpper,
  ADisplay: TArray<string>; const AUnitHasInit: TArray<Boolean>;
  AMap: TDictionary<string, TArray<Integer>>);
begin
  inherited Create;
  FUnitName := AUnitName;
  FUnitPath := AUnitPath;
  FUnitHasInit := AUnitHasInit;
  FKeysUpper := AKeysUpper;
  FDisplay := ADisplay;
  FMap := AMap;   // takes ownership
end;

destructor TUnitSnapshot.Destroy;
begin
  FMap.Free;
  inherited;
end;

function TUnitSnapshot.HasUnit(const AUnitName: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  if AUnitName = '' then Exit;
  for I := 0 to High(FUnitName) do
    if SameText(FUnitName[I], AUnitName) then
      Exit(True);
end;

function TUnitSnapshot.TryGetUnitPath(const AUnitName: string;
  out APath: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  APath := '';
  if AUnitName = '' then Exit;
  for I := 0 to High(FUnitName) do
    if SameText(FUnitName[I], AUnitName) then
    begin
      APath := FUnitPath[I];
      Exit(True);
    end;
end;

function TUnitSnapshot.HasInitCode(const AUnitName: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  if AUnitName = '' then Exit;
  for I := 0 to High(FUnitName) do
    if SameText(FUnitName[I], AUnitName) then
      Exit((I <= High(FUnitHasInit)) and FUnitHasInit[I]);
end;

function TUnitSnapshot.UnitCount: Integer;
begin
  Result := Length(FUnitName);
end;

function TUnitSnapshot.IdentCount: Integer;
begin
  Result := Length(FKeysUpper);
end;

function TUnitSnapshot.Lookup(const AIdentifier: string): TArray<TFindUnitHit>;
var
  Ids: TArray<Integer>;
  Res: TList<TFindUnitHit>;
begin
  Result := nil;
  if (AIdentifier = '') or not FMap.TryGetValue(UpperCase(AIdentifier), Ids) then Exit;
  Res := TList<TFindUnitHit>.Create;
  try
    for var Id in Ids do
    begin
      var UnitIdx := Id and not GenericBit;
      if (UnitIdx >= 0) and (UnitIdx < Length(FUnitName)) then
      begin
        var H: TFindUnitHit;
        H.Identifier := AIdentifier;
        H.UnitName := FUnitName[UnitIdx];
        H.Path := FUnitPath[UnitIdx];
        H.IsGeneric := (Id and GenericBit) <> 0;
        Res.Add(H);
      end;
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

function TUnitSnapshot.FuzzyIdentifiers(const AIdent: string;
  AMaxDist, AMax: Integer): TArray<TFindUnitHit>;
type
  TCand = record Dist, KeyIdx: Integer; end;
var
  U: string;
  L, I, D: Integer;
  Cands: TList<TCand>;
begin
  Result := nil;
  U := UpperCase(Trim(AIdent));
  L := Length(U);
  if (L < 3) or (AMax <= 0) then Exit;

  Cands := TList<TCand>.Create;
  try
    for I := 0 to High(FKeysUpper) do
    begin
      if Abs(Length(FKeysUpper[I]) - L) > AMaxDist then Continue;
      D := BoundedEditDistance(U, FKeysUpper[I], AMaxDist);
      if (D = 0) or (D > AMaxDist) then Continue;   // 0 = same identifier
      var C: TCand;
      C.Dist := D;
      C.KeyIdx := I;
      Cands.Add(C);
    end;
    Cands.Sort(TComparer<TCand>.Construct(
      function(const A, B: TCand): Integer
      begin
        Result := A.Dist - B.Dist;
        if Result = 0 then
          Result := CompareStr(FKeysUpper[A.KeyIdx], FKeysUpper[B.KeyIdx]);
      end));

    for I := 0 to Cands.Count - 1 do
    begin
      if Length(Result) >= AMax then Break;
      var Ids := FMap[FKeysUpper[Cands[I].KeyIdx]];
      if Length(Ids) = 0 then Continue;
      var UnitIdx := Ids[0] and not GenericBit;
      if UnitIdx >= Length(FUnitName) then Continue;
      var H: TFindUnitHit;
      H.Identifier := FDisplay[Cands[I].KeyIdx];
      H.UnitName := FUnitName[UnitIdx];
      H.Path := FUnitPath[UnitIdx];
      H.IsGeneric := (Ids[0] and GenericBit) <> 0;
      Result := Result + [H];
    end;
  finally
    Cands.Free;
  end;
end;

function TUnitSnapshot.FuzzyUnitNames(const AName: string;
  AMaxDist, AMax: Integer): TArray<string>;
type
  TCand = record Dist: Integer; Name: string; end;
var
  U, KU, LastSeg: string;
  I, D, P: Integer;
  Cands: TList<TCand>;
  Seen: TDictionary<string, Boolean>;
begin
  Result := nil;
  U := UpperCase(Trim(AName));
  if (Length(U) < 3) or (AMax <= 0) then Exit;

  Cands := TList<TCand>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for I := 0 to High(FUnitName) do
    begin
      KU := UpperCase(FUnitName[I]);
      if Seen.ContainsKey(KU) then Continue;
      Seen.Add(KU, True);
      if KU = U then Continue;   // exact (case-insensitive) - not a fix
      // Match against the full dotted name AND against the last segment
      // ('Windwos' should find 'Winapi.Windows').
      D := BoundedEditDistance(U, KU, AMaxDist);
      if D > AMaxDist then
      begin
        P := KU.LastDelimiter('.') + 1;   // 0-based helper -> 1-based char
        if P > 1 then
        begin
          LastSeg := Copy(KU, P + 1, MaxInt);
          if LastSeg <> U then
            D := BoundedEditDistance(U, LastSeg, AMaxDist)
          else
            D := 1;   // exact last segment: near-certain candidate
        end;
      end;
      if D > AMaxDist then Continue;
      var C: TCand;
      C.Dist := D;
      C.Name := FUnitName[I];
      Cands.Add(C);
    end;
    Cands.Sort(TComparer<TCand>.Construct(
      function(const A, B: TCand): Integer
      begin
        Result := A.Dist - B.Dist;
        if Result = 0 then Result := CompareText(A.Name, B.Name);
      end));
    for I := 0 to Cands.Count - 1 do
    begin
      if Length(Result) >= AMax then Break;
      Result := Result + [Cands[I].Name];
    end;
  finally
    Seen.Free;
    Cands.Free;
  end;
end;

function TUnitSnapshot.Search(const ASub: string; AMax: Integer): TArray<TFindUnitHit>;
var
  Needle: string;
  Res: TList<TFindUnitHit>;
  I: Integer;
begin
  Result := nil;
  if Length(Trim(ASub)) < 2 then Exit;
  Needle := UpperCase(Trim(ASub));
  Res := TList<TFindUnitHit>.Create;
  try
    for I := 0 to High(FKeysUpper) do
    begin
      if Pos(Needle, FKeysUpper[I]) = 0 then Continue;
      for var Id in FMap[FKeysUpper[I]] do
      begin
        var UnitIdx := Id and not GenericBit;
        if (UnitIdx < 0) or (UnitIdx >= Length(FUnitName)) then Continue;
        var H: TFindUnitHit;
        H.Identifier := FDisplay[I];
        H.UnitName := FUnitName[UnitIdx];
        H.Path := FUnitPath[UnitIdx];
        H.IsGeneric := (Id and GenericBit) <> 0;
        Res.Add(H);
        if Res.Count >= AMax then Break;
      end;
      if Res.Count >= AMax then Break;
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  TUnitIndex
// ---------------------------------------------------------------------------

constructor TUnitIndex.Create;
begin
  inherited Create;
  FLock := TCriticalSection.Create;
  FWake := TEvent.Create(nil, False, False, '');
  FGlobalByPath := TObjectDictionary<string, TIndexedUnit>.Create([doOwnsValues]);
  FProjectByPath := TObjectDictionary<string, TIndexedUnit>.Create([doOwnsValues]);
  FStatus := 'Index not started.';
end;

destructor TUnitIndex.Destroy;
begin
  if FWorker <> nil then
  begin
    FWorker.Terminate;
    FWake.SetEvent;
    FWorker.WaitFor;
    FWorker.Free;
  end;
  FGlobalByPath.Free;
  FProjectByPath.Free;
  FSnapshot := nil;
  FWake.Free;
  FLock.Free;
  inherited;
end;

function TUnitIndex.CacheFileFor(const AKey: string): string;
begin
  Result := TPath.Combine(
    TPath.Combine(TPath.Combine(TPath.GetHomePath, 'DelphiRefactoringLight'), 'unitindex'),
    HashKey(AKey) + '.idx');
end;

procedure TUnitIndex.EnsureWorker;
begin
  if FWorker = nil then
    FWorker := TIndexWorker.Create(Self);
end;

function TUnitIndex.GetSnapshot: IUnitSnapshot;
begin
  FLock.Enter;
  try Result := FSnapshot; finally FLock.Leave; end;
end;

function TUnitIndex.Snapshot: IUnitSnapshot;
begin
  Result := GetSnapshot;
end;

procedure TUnitIndex.StartGlobalIndex;
var
  GDirs: TStringList;
  GlobKey: string;
begin
  GDirs := TStringList.Create;
  try
    GDirs.Duplicates := dupIgnore;
    GDirs.Sorted := True;
    for var D in GatherGlobalLibraryDirs(FindBdsRoot) do GDirs.Add(D);
    GlobKey := 'global-' + HashKey(UpperCase(string.Join(';', GDirs.ToStringArray)));
    FLock.Enter;
    try
      FGlobalDirs := GDirs.ToStringArray;
      FGlobalCachePath := CacheFileFor(GlobKey);
      if FStatus = 'Index not started.' then FStatus := 'Building global index...';
    finally
      FLock.Leave;
    end;
  finally
    GDirs.Free;
  end;
  EnsureWorker;
  FWake.SetEvent;
end;

class function TUnitIndex.Instance: TUnitIndex;
begin
  if FInstance = nil then FInstance := TUnitIndex.Create;
  Result := FInstance;
end;

class procedure TUnitIndex.ReleaseInstance;
begin
  FreeAndNil(FInstance);
end;

procedure TUnitIndex.SetStatus(const S: string);
begin
  FLock.Enter;
  try FStatus := S; finally FLock.Leave; end;
end;

function TUnitIndex.Ready: Boolean;
begin
  FLock.Enter;
  try Result := FReady; finally FLock.Leave; end;
end;

function TUnitIndex.StatusLine: string;
begin
  FLock.Enter;
  try Result := FStatus; finally FLock.Leave; end;
end;

procedure TUnitIndex.RefreshSourcesFromEditor;
var
  GDirs, PDirs: TStringList;
  Files: TArray<string>;
  ProjKey, GlobKey: string;
begin
  if Editor = nil then Exit;
  GDirs := TStringList.Create;
  PDirs := TStringList.Create;
  try
    GDirs.Duplicates := dupIgnore; GDirs.Sorted := True;
    PDirs.Duplicates := dupIgnore; PDirs.Sorted := True;

    // Global scope: IDE library/browsing paths (project-independent).
    for var D in GatherGlobalLibraryDirs(FindBdsRoot) do GDirs.Add(D);
    // Project scope: the project's own search path + .dproj DCC search dirs.
    for var D in Editor.GetProjectSearchPaths.Split([';'], TStringSplitOptions.ExcludeEmpty) do
      if TDirectory.Exists(Trim(D)) then PDirs.Add(ExcludeTrailingPathDelimiter(Trim(D)));
    for var D in GatherDprojSearchDirs do PDirs.Add(D);
    Files := Editor.GetProjectSourceFiles;

    ProjKey := UpperCase(Editor.GetCurrentProjectDproj);
    if ProjKey = '' then ProjKey := 'default';
    // The global cache is keyed by the SET of global dirs, so every project
    // that sees the same library paths shares one cached global index.
    GlobKey := 'global-' + HashKey(UpperCase(string.Join(';', GDirs.ToStringArray)));

    FLock.Enter;
    try
      FGlobalDirs := GDirs.ToStringArray;
      FProjectDirs := PDirs.ToStringArray;
      FProjectFiles := Files;
      FGlobalCachePath := CacheFileFor(GlobKey);
      FProjectCachePath := CacheFileFor(ProjKey);
    finally
      FLock.Leave;
    end;
  finally
    GDirs.Free;
    PDirs.Free;
  end;

  EnsureWorker;
  FWake.SetEvent;   // trigger a (re)scan
end;

procedure TUnitIndex.CurrentSources(out AGlobalDirs, AProjectDirs, AProjectFiles: TArray<string>;
  out AGlobalCache, AProjectCache: string);
begin
  FLock.Enter;
  try
    AGlobalDirs := Copy(FGlobalDirs);
    AProjectDirs := Copy(FProjectDirs);
    AProjectFiles := Copy(FProjectFiles);
    AGlobalCache := FGlobalCachePath;
    AProjectCache := FProjectCachePath;
  finally
    FLock.Leave;
  end;
end;

// ---- snapshot publication -------------------------------------------------

procedure TUnitIndex.PublishSnapshot;
var
  Map: TDictionary<string, TList<Integer>>;
  Names, Paths, Display: TArray<string>;
  KeysUpper: TArray<string>;
  Units: TList<TIndexedUnit>;
  Seen: TDictionary<string, Boolean>;
  I, J: Integer;
begin
  // Merge the two scopes into one flat unit list (project entries win on a
  // path collision, so a project copy of a unit shadows a library copy).
  Units := TList<TIndexedUnit>.Create;
  Seen := TDictionary<string, Boolean>.Create;
  Map := TDictionary<string, TList<Integer>>.Create;
  var DispD := TDictionary<string, string>.Create;
  try
    for var U in FProjectByPath.Values do
      if not Seen.ContainsKey(UpperCase(U.Path)) then
      begin Seen.Add(UpperCase(U.Path), True); Units.Add(U); end;
    for var U in FGlobalByPath.Values do
      if not Seen.ContainsKey(UpperCase(U.Path)) then
      begin Seen.Add(UpperCase(U.Path), True); Units.Add(U); end;

    SetLength(Names, Units.Count);
    SetLength(Paths, Units.Count);
    var InitFlags: TArray<Boolean>;
    SetLength(InitFlags, Units.Count);
    for I := 0 to Units.Count - 1 do
    begin
      Names[I] := Units[I].UnitName;
      Paths[I] := Units[I].Path;
      InitFlags[I] := Units[I].HasInit;
      for J := 0 to High(Units[I].Idents) do
      begin
        var Id := Units[I].Idents[J];
        // Trailing '<' = generic-declaration marker from ParseUnit; the
        // key is the bare name, genericity travels as a flag bit on the
        // unit index (unit counts stay far below the bit).
        var Entry := I;
        if (Id <> '') and (Id[Length(Id)] = '<') then
        begin
          Id := Copy(Id, 1, Length(Id) - 1);
          Entry := I or GenericBit;
        end;
        if Id = '' then Continue;
        var Key := UpperCase(Id);
        var Lst: TList<Integer>;
        if not Map.TryGetValue(Key, Lst) then
        begin Lst := TList<Integer>.Create; Map.Add(Key, Lst); DispD.Add(Key, Id); end;
        if (Lst.Count = 0) or (Lst.Last <> Entry) then Lst.Add(Entry);
      end;
    end;

    // Flatten to arrays + a sorted key list for search.
    var Flat := TDictionary<string, TArray<Integer>>.Create;
    KeysUpper := Map.Keys.ToArray;
    TArray.Sort<string>(KeysUpper);
    SetLength(Display, Length(KeysUpper));
    for I := 0 to High(KeysUpper) do
    begin
      Flat.Add(KeysUpper[I], Map[KeysUpper[I]].ToArray);
      Display[I] := DispD[KeysUpper[I]];
    end;

    var Snap: IUnitSnapshot := TUnitSnapshot.Create(Names, Paths, KeysUpper,
      Display, InitFlags, Flat);
    FLock.Enter;
    try
      FSnapshot := Snap;   // atomic reference swap; readers scan lock-free
      FReady := True;
    finally
      FLock.Leave;
    end;
  finally
    for var L in Map.Values do L.Free;
    Map.Free;
    DispD.Free;
    Units.Free;
    Seen.Free;
  end;
end;

// ---- lookups (ANY THREAD - scan the immutable snapshot lock-free) ---------

function TUnitIndex.Lookup(const AIdentifier: string): TArray<TFindUnitHit>;
var
  Snap: IUnitSnapshot;
begin
  Snap := GetSnapshot;
  if Snap <> nil then Result := Snap.Lookup(AIdentifier) else Result := nil;
end;

function TUnitIndex.Search(const ASub: string; AMax: Integer): TArray<TFindUnitHit>;
var
  Snap: IUnitSnapshot;
begin
  Snap := GetSnapshot;
  if Snap <> nil then Result := Snap.Search(ASub, AMax) else Result := nil;
end;

// ---- persistence (WORKER THREAD) ------------------------------------------

procedure TUnitIndex.SaveCache(const APath: string;
  ADict: TObjectDictionary<string, TIndexedUnit>);
var
  FS: TFileStream;

  procedure WStr(const S: string);
  var B: TBytes; L: Integer;
  begin
    B := TEncoding.UTF8.GetBytes(S);
    L := Length(B);
    FS.WriteBuffer(L, SizeOf(L));
    if L > 0 then FS.WriteBuffer(B[0], L);
  end;

begin
  if APath = '' then Exit;
  try
    TDirectory.CreateDirectory(ExtractFilePath(APath));
    FS := TFileStream.Create(APath, fmCreate);
    try
      WStr(CacheMagic);
      var Cnt := ADict.Count;
      FS.WriteBuffer(Cnt, SizeOf(Cnt));
      for var U in ADict.Values do
      begin
        WStr(U.Path); WStr(U.UnitName);
        FS.WriteBuffer(U.MTime, SizeOf(U.MTime));
        FS.WriteBuffer(U.Size, SizeOf(U.Size));
        FS.WriteBuffer(U.HasInit, SizeOf(U.HasInit));
        var IC := Length(U.Idents);
        FS.WriteBuffer(IC, SizeOf(IC));
        for var Id in U.Idents do WStr(Id);
      end;
    finally
      FS.Free;
    end;
  except
    // cache is an optimization - ignore failures
  end;
end;

procedure TUnitIndex.LoadCache(const APath: string;
  ADict: TObjectDictionary<string, TIndexedUnit>);
var
  FS: TFileStream;

  function RStr: string;
  var B: TBytes; L: Integer;
  begin
    FS.ReadBuffer(L, SizeOf(L));
    if L <= 0 then Exit('');
    SetLength(B, L);
    FS.ReadBuffer(B[0], L);
    Result := TEncoding.UTF8.GetString(B);
  end;

begin
  if (APath = '') or not TFile.Exists(APath) then Exit;
  try
    FS := TFileStream.Create(APath, fmOpenRead or fmShareDenyWrite);
    try
      if RStr <> CacheMagic then Exit;
      var Cnt: Integer;
      FS.ReadBuffer(Cnt, SizeOf(Cnt));
      for var I := 0 to Cnt - 1 do
      begin
        var U := TIndexedUnit.Create;
        U.Path := RStr; U.UnitName := RStr;
        FS.ReadBuffer(U.MTime, SizeOf(U.MTime));
        FS.ReadBuffer(U.Size, SizeOf(U.Size));
        FS.ReadBuffer(U.HasInit, SizeOf(U.HasInit));
        var IC: Integer;
        FS.ReadBuffer(IC, SizeOf(IC));
        SetLength(U.Idents, IC);
        for var J := 0 to IC - 1 do U.Idents[J] := RStr;
        ADict.AddOrSetValue(UpperCase(U.Path), U);
      end;
    finally
      FS.Free;
    end;
  except
    ADict.Clear;   // corrupt cache -> rebuild from scratch
  end;
end;

// ---- the build / refresh pass (WORKER THREAD) -----------------------------

function TUnitIndex.RefreshScope(const ADirs, AFiles: TArray<string>;
  ADict: TObjectDictionary<string, TIndexedUnit>; const AScopeName: string): Boolean;
var
  Target: TDictionary<string, Boolean>;   // UPPER path -> present
  Processed, Total: Integer;
begin
  Result := False;
  Target := TDictionary<string, Boolean>.Create;
  try
    var All := TList<string>.Create;
    try
      for var F in AFiles do
        if SameText(ExtractFileExt(F), '.pas') and TFile.Exists(F) then All.Add(F);
      for var D in ADirs do
      begin
        if All.Count >= MaxFiles then Break;
        if not TDirectory.Exists(D) then Continue;
        try
          for var F in TDirectory.GetFiles(D, '*.pas') do
          begin
            All.Add(F);
            if All.Count >= MaxFiles then Break;
          end;
        except
        end;
      end;

      Total := All.Count; Processed := 0;
      for var F in All do
      begin
        if (FWorker <> nil) and FWorker.Terminated then Exit;
        Inc(Processed);
        if Processed mod 200 = 0 then
          SetStatus(Format('Indexing %s: %d / %d ...', [AScopeName, Processed, Total]));
        var Key := UpperCase(F);
        Target.AddOrSetValue(Key, True);
        var Info: TIndexedUnit;
        var Size: Int64;
        var MT: TDateTime;
        try
          Size := TFile.GetSize(F);
          MT := TFile.GetLastWriteTime(F);
        except
          Continue;
        end;
        if Size > MaxFileBytes then Continue;
        if ADict.TryGetValue(Key, Info) and (Info.Size = Size) and
           (Abs(Info.MTime - MT) < 1 / 86400) then
          Continue;   // unchanged
        var UName: string;
        var HasInit: Boolean;
        var Idents := ParseUnit(F, UName, HasInit);
        if UName = '' then UName := ChangeFileExt(ExtractFileName(F), '');
        var NU := TIndexedUnit.Create;
        NU.Path := F; NU.UnitName := UName; NU.Size := Size; NU.MTime := MT;
        NU.Idents := Idents;
        NU.HasInit := HasInit;
        ADict.AddOrSetValue(Key, NU);
        Result := True;
      end;

      // Drop files that disappeared from this scope.
      var Dead := TList<string>.Create;
      try
        for var K in ADict.Keys do
          if not Target.ContainsKey(K) then Dead.Add(K);
        for var K in Dead do begin ADict.Remove(K); Result := True; end;
      finally
        Dead.Free;
      end;
    finally
      All.Free;
    end;
  finally
    Target.Free;
  end;
end;

procedure TUnitIndex.WorkerLoop;
const
  GlobalRescanCycles = 20;   // ~10 min at a 30 s interval
var
  GDirs, PDirs, PFiles: TArray<string>;
  GCache, PCache: string;
  Cycle: Integer;
begin
  // Initial load from both caches - serve them immediately.
  CurrentSources(GDirs, PDirs, PFiles, GCache, PCache);
  LoadCache(GCache, FGlobalByPath);  FGlobalKey := GCache;
  LoadCache(PCache, FProjectByPath); FProjectKey := PCache;
  if (FGlobalByPath.Count > 0) or (FProjectByPath.Count > 0) then
    PublishSnapshot;

  Cycle := 0;
  while (FWorker <> nil) and not FWorker.Terminated do
  begin
    try
      CurrentSources(GDirs, PDirs, PFiles, GCache, PCache);
      var Changed := False;

      // A changed cache key means a different library set / project - reload
      // that scope's cache before refreshing it.
      if GCache <> FGlobalKey then
      begin
        FGlobalByPath.Clear; LoadCache(GCache, FGlobalByPath); FGlobalKey := GCache; Changed := True;
      end;
      if PCache <> FProjectKey then
      begin
        FProjectByPath.Clear; LoadCache(PCache, FProjectByPath); FProjectKey := PCache; Changed := True;
      end;

      // Global scope: scan only on the first pass and rarely thereafter -
      // the RTL/VCL/library tree barely changes and is shared across projects.
      if (Cycle = 0) or (Cycle mod GlobalRescanCycles = 0) then
        if RefreshScope(GDirs, nil, FGlobalByPath, 'library') then
        begin Changed := True; SaveCache(GCache, FGlobalByPath); end;

      // Project scope: refresh every cycle (this is what changes while editing).
      if RefreshScope(PDirs, PFiles, FProjectByPath, 'project') then
      begin Changed := True; SaveCache(PCache, FProjectByPath); end;

      if Changed or not FReady then PublishSnapshot;
      var Snap := GetSnapshot;
      if Snap <> nil then
        SetStatus(Format('Index ready: %d units, %d identifiers.',
          [Snap.UnitCount, Snap.IdentCount]));
    except
      // never let a parse error kill the worker
    end;
    TInterlocked.Increment(FScanCycle);
    Inc(Cycle);
    if (FWorker <> nil) and not FWorker.Terminated then
      FWake.WaitFor(RefreshIntervalMs);
  end;
end;

function TUnitIndex.ScanCycle: Integer;
begin
  Result := TInterlocked.CompareExchange(FScanCycle, 0, 0);
end;

{ TUnitIndex.TIndexWorker }

constructor TUnitIndex.TIndexWorker.Create(AOwner: TUnitIndex);
begin
  FOwner := AOwner;
  FreeOnTerminate := False;
  inherited Create(False);
  Priority := tpLower;
end;

procedure TUnitIndex.TIndexWorker.Execute;
begin
  FOwner.WorkerLoop;
end;

initialization

finalization
  TUnitIndex.ReleaseInstance;

end.
