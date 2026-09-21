(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IncludeExpansion;

// {$I file} / {$INCLUDE file} for the project-wide scans (rename, find
// references, safe delete).
//
// An include file is no unit: DelphiLSP cannot analyse it on its own, so a
// query at a position INSIDE it answers nothing - measured (scratchpad
// lspprobe\ProbeInc.dpr): an .inc sent as its own document stays silent,
// while the same position answers correctly once the INCLUDING unit is sent
// with the include expanded in place, exactly as the compiler sees it (the
// user's idea). Definitions POINTING INTO an .inc work natively.
//
// So:
//  * ExpandUnitIncludes builds that expanded text and records a SEGMENT MAP
//    (which run of which file landed where) - positions translate in both
//    directions, nested includes included.
//  * TLspIncludeContext sends an including unit expanded ONLY when a query
//    inside one of its include files is needed, answers pointing into the
//    expanded document are mapped back to the real file, and Restore sends
//    the original text again (the live checker and the diagnostics must see
//    the unit as it is).
//
// The directive rules are the index parser's (Expert.UnitIndex.
// ExpandIncludeDirectives is built on the same core): directives inside
// comments and strings are ignored, {$I+}/{$I-} is I/O checking, a name
// without extension tries .inc and .pas, nested includes resolve against
// their own folder, depth-limited; an unreadable file stays as the directive.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Lsp.Client,
  Lsp.Protocol;

type
  /// <summary>Content of a file: the editor buffer when open, else disk.
  ///  False when it cannot be read.</summary>
  TIncludeReader = reference to function(const APath: string;
    out AContent: string): Boolean;

  TIncludeSegment = record
    FileName: string;   // full path of the source of this run
    SrcOffset: Integer; // 0-based offset in that file's text
    DstOffset: Integer; // 0-based offset in the expanded text
    Length: Integer;
  end;

  TExpandedUnit = class
  private
    FFileName: string;
    FOriginal: string;
    FText: string;
    FSegments: TArray<TIncludeSegment>;
    FIncludes: TArray<string>;
    FSources: TDictionary<string, string>;         // UPPER path -> text
    FLineStarts: TDictionary<string, TArray<Integer>>;
    function LineStartsOf(const AKey, AText: string): TArray<Integer>;
    function OffsetOf(const AKey, AText: string; ALine, ACol: Integer): Integer;
    procedure PosOf(const AKey, AText: string; AOffset: Integer; out ALine, ACol: Integer);
  public
    constructor Create(const AFileName, AOriginal: string);
    destructor Destroy; override;
    /// <summary>A position in AFile (the unit itself or one of its
    ///  includes) as a position in Text. False when that file is not part
    ///  of this expansion or the position lies in a directive.</summary>
    function ToExpanded(const AFile: string; ALine, ACol: Integer;
      out AExpLine, AExpCol: Integer): Boolean;
    /// <summary>A position in Text back in its real file.</summary>
    function FromExpanded(AExpLine, AExpCol: Integer; out AFile: string;
      out ALine, ACol: Integer): Boolean;
    function Contains(const AFile: string): Boolean;
    property FileName: string read FFileName;
    property Original: string read FOriginal;
    property Text: string read FText;
    property Includes: TArray<string> read FIncludes;
    property Segments: TArray<TIncludeSegment> read FSegments;
  end;

/// <summary>AText of AFile with every resolvable include expanded, plus the
///  segment map. nil when AText includes nothing (the common case - the
///  caller then uses the file as it is). AReader supplies the include
///  files' content (nil = disk).</summary>
function ExpandUnitIncludes(const AFile, AText: string;
  const AReader: TIncludeReader): TExpandedUnit;

/// <summary>The expanded text only (index parser). AIncludes (may be nil)
///  receives every file that was read, nested ones too.</summary>
function ExpandIncludeText(const AText, ABaseDir: string; AIncludes: TStrings;
  const AReader: TIncludeReader = nil): string;

/// <summary>Include expander for DEBUGGING (user request): every resolvable
///  include of AText replaced IN PLACE by its content, framed by marker
///  lines that keep the original directive -
///    // >>> include begin: {$I foo.inc}
///    ...content...
///    // <<< include end: foo.inc
///  Line comments, because the directive itself contains braces. The
///  markers start on a line of their own; a directive sharing its line with
///  code gets that code on the lines around the block. Nested includes are
///  expanded (and marked) too. ACount = number of includes expanded; 0 means
///  the result equals AText. Revert with the version control system.</summary>
function ExpandIncludesMarked(const AText, ABaseDir: string;
  const AReader: TIncludeReader; out ACount: Integer): string;

const
  IncludeMarkerBegin = '// >>> include begin: ';
  IncludeMarkerEnd = '// <<< include end: ';

/// <summary>Full paths of the files AText of AFile includes (nested too),
///  without duplicates.</summary>
function CollectIncludeFiles(const AFile, AText: string;
  const AReader: TIncludeReader): TArray<string>;

/// <summary>The directive body ("I foo.inc", "INCLUDE 'x'") -> file name, or
///  '' when it is no include ({$I+}, {$IFDEF}, ...).</summary>
function IncludeNameOfDirective(const ABody: string): string;

/// <summary>AName (as written in the directive) resolved against ABaseDir;
///  a name without extension tries .inc and .pas. '' when not found.</summary>
function ResolveIncludeFile(const AName, ABaseDir: string): string;

/// <summary>True for a file the scans treat as an include (not a unit,
///  program or package).</summary>
function IsIncludeFile(const AFile: string): Boolean;

type
  /// <summary>Lets DelphiLSP answer at positions INSIDE include files for the
  ///  duration of one scan. Not thread-safe; one per scan.</summary>
  TLspIncludeContext = class
  private
    FClient: TLspClient;
    FReader: TIncludeReader;
    FUnits: TObjectList<TExpandedUnit>;
    FOwnerOf: TDictionary<string, TExpandedUnit>;   // UPPER include -> unit
    FActive: TList<TExpandedUnit>;
    FWaitMs: Integer;
    FActivations: Integer;
    FNotes: TStringList;
    function UnitOf(const AFile: string): TExpandedUnit;
    function Activate(AUnit: TExpandedUnit): Boolean;
    function IsActive(AUnit: TExpandedUnit): Boolean;
    function MapBack(const ALocs: TArray<TLspLocation>): TArray<TLspLocation>;
  public
    /// <summary>AReader: buffer-or-disk content (called on the creating
    ///  thread's schedule - pass a reader that is safe there).</summary>
    constructor Create(AClient: TLspClient; const AReader: TIncludeReader;
      AWaitMs: Integer = 30000);
    destructor Destroy; override;
    /// <summary>Remembers which units of AFiles include what. Reads each
    ///  unit that contains an include directive; nothing is sent yet.</summary>
    procedure RegisterFiles(const AFiles: TArray<string>);
    /// <summary>AFile is an include file of a registered unit, or a unit that
    ///  is currently sent EXPANDED: the caller must not send its own
    ///  version of it (RefreshDocument / didOpen) - it would undo the
    ///  expansion or send an include as a unit.</summary>
    function OwnsDocument(const AFile: string): Boolean;
    /// <summary>An include file this context can answer for.</summary>
    function CanResolve(const AFile: string): Boolean;
    /// <summary>GotoDefinition that also works inside include files; answers
    ///  that point into an expanded document come back in real files.</summary>
    function Definition(const AFile: string; ALine, ACol: Integer): TArray<TLspLocation>;
    /// <summary>Like TLspSymbolTargets.AddWithPartner, via Definition.
    ///  AName: the symbol's name - the partner is then asked AT the name on
    ///  that line, not at ACol. DelphiLSP reports an implementation header
    ///  as a range starting at column 0, and a query there lands on the
    ///  keyword and answers nothing: the declaration went missing from the
    ///  target set and the source pre-check threw away every correctly
    ///  resolved call (PsyPrax report, 2026-09-21).</summary>
    procedure AddTargetWithPartner(var ATargets: TLspSymbolTargets;
      const AFile: string; ALine, ACol: Integer; const AName: string = '');
    /// <summary>Sends every expanded unit's ORIGINAL text again.</summary>
    procedure Restore;
    property Activations: Integer read FActivations;
    /// <summary>What happened ("Expert.AutoImport.pas sent expanded (1
    ///  include), analysed in 640 ms") - for detail logs.</summary>
    function NotesText: string;
  end;

implementation

uses
  System.IOUtils, System.StrUtils, System.Math, Winapi.Windows, Lsp.Uri,
  Delphi.FileEncoding, Expert.PascalScanner;

const
  MaxIncludeDepth = 8;

function IsIncludeFile(const AFile: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(AFile));
  Result := (Ext <> '.pas') and (Ext <> '.dpr') and (Ext <> '.dpk');
end;

function ResolveIncludeFile(const AName, ABaseDir: string): string;
var
  N: string;
begin
  Result := '';
  N := Trim(AName);
  if (Length(N) >= 2) and (N[1] = '''') and (N[Length(N)] = '''') then
    N := Copy(N, 2, Length(N) - 2);
  N := Trim(N).Replace('/', '\');
  if N = '' then Exit;
  try
    if not TPath.IsPathRooted(N) then
      N := TPath.Combine(ABaseDir, N);
    if TFile.Exists(N) then Exit(TPath.GetFullPath(N));
    if ExtractFileExt(N) = '' then
      for var Ext in ['.inc', '.pas'] do
        if TFile.Exists(N + Ext) then Exit(TPath.GetFullPath(N + Ext));
  except
    Result := '';   // invalid path characters in the directive
  end;
end;

function IncludeNameOfDirective(const ABody: string): string;
var
  U: string;
begin
  Result := '';
  U := UpperCase(ABody);
  // "{$I+}" / "{$I-}" / "{$IFDEF" never match: a blank must follow
  if StartsStr('INCLUDE', U) and (Length(U) > 7) and CharInSet(U[8], [' ', #9]) then
    Result := Trim(Copy(ABody, 9, MaxInt))
  else if StartsStr('I', U) and (Length(U) > 1) and CharInSet(U[2], [' ', #9]) then
    Result := Trim(Copy(ABody, 3, MaxInt));
end;

function DiskReader(const APath: string; out AContent: string): Boolean;
begin
  try
    AContent := TDelphiFileEncoding.ReadAll(APath);
    Result := True;
  except
    AContent := '';
    Result := False;
  end;
end;

type
  // The one expansion core. Output = verbatim runs of the source texts +
  // the expanded include contents, each followed by a synthetic #10 (what
  // follows the directive starts on a line of its own - a trailing '//'
  // comment in the include must not swallow it).
  TExpander = class
  private
    FSB: TStringBuilder;
    FReader: TIncludeReader;
    FSegments: TList<TIncludeSegment>;
    FIncludes: TStrings;
    FSources: TDictionary<string, string>;
    FMarked: Boolean;
    FLineBreak: string;
    FCount: Integer;
    procedure Run(const AFile, AText: string; ALo, AHi: Integer);
    procedure EnsureLineStart;
  public
    constructor Create(const AReader: TIncludeReader; AIncludes: TStrings;
      ASources: TDictionary<string, string>);
    destructor Destroy; override;
    procedure Expand(const AFile, AText, ABaseDir: string; ADepth: Integer);
  end;

constructor TExpander.Create(const AReader: TIncludeReader; AIncludes: TStrings;
  ASources: TDictionary<string, string>);
begin
  inherited Create;
  FSB := TStringBuilder.Create;
  FReader := AReader;
  FSegments := TList<TIncludeSegment>.Create;
  FIncludes := AIncludes;
  FSources := ASources;
end;

destructor TExpander.Destroy;
begin
  FSegments.Free;
  FSB.Free;
  inherited;
end;

// Marked mode: the next text starts on a line of its own.
procedure TExpander.EnsureLineStart;
begin
  if (FSB.Length > 0) and not CharInSet(FSB.Chars[FSB.Length - 1], [#10, #13]) then
    FSB.Append(FLineBreak);
end;

// Appends AText[ALo..AHi) (1-based, exclusive end) verbatim and records it.
procedure TExpander.Run(const AFile, AText: string; ALo, AHi: Integer);
var
  Seg: TIncludeSegment;
begin
  if AHi <= ALo then Exit;
  Seg.FileName := AFile;
  Seg.SrcOffset := ALo - 1;
  Seg.DstOffset := FSB.Length;
  Seg.Length := AHi - ALo;
  FSegments.Add(Seg);
  FSB.Append(AText, ALo - 1, AHi - ALo);
end;

procedure TExpander.Expand(const AFile, AText, ABaseDir: string; ADepth: Integer);
var
  I, N, RunStart, Start, Stop, CloseLen: Integer;
  Body, U, Name, Path, Sub: string;
begin
  if (FSources <> nil) and (AFile <> '') then
    FSources.AddOrSetValue(UpperCase(AFile), AText);
  N := Length(AText);
  U := UpperCase(AText);
  // cheap exit - almost no unit has an include
  if (ADepth > MaxIncludeDepth) or ((Pos('{$I', U) = 0) and (Pos('(*$I', U) = 0)) then
  begin
    Run(AFile, AText, 1, N + 1);
    Exit;
  end;
  RunStart := 1;
  I := 1;
  while I <= N do
  begin
    var C := AText[I];
    if C = '''' then
    begin
      // string literal ('' inside = two literals in a row, same result)
      Inc(I);
      while (I <= N) and (AText[I] <> '''') and (AText[I] <> #10) and (AText[I] <> #13) do
        Inc(I);
      if (I <= N) and (AText[I] = '''') then Inc(I);
      Continue;
    end;
    if (C = '/') and (I < N) and (AText[I + 1] = '/') then
    begin
      while (I <= N) and (AText[I] <> #10) and (AText[I] <> #13) do Inc(I);
      Continue;
    end;
    if (C = '{') or ((C = '(') and (I < N) and (AText[I + 1] = '*')) then
    begin
      Start := I;
      Body := '';
      if C = '{' then
      begin
        Stop := PosEx('}', AText, I + 1);
        CloseLen := 1;
        if Stop = 0 then Stop := N + 1;
        if (I < N) and (AText[I + 1] = '$') then
          Body := Copy(AText, I + 2, Stop - I - 2);
      end
      else
      begin
        Stop := PosEx('*)', AText, I + 2);
        CloseLen := 2;
        if Stop = 0 then Stop := N + 1;
        if (I + 2 <= N) and (AText[I + 2] = '$') then
          Body := Copy(AText, I + 3, Stop - I - 3);
      end;
      I := Min(Stop + CloseLen, N + 1);
      Name := '';
      if Body <> '' then
        Name := IncludeNameOfDirective(Body);
      if Name <> '' then
      begin
        Path := ResolveIncludeFile(Name, ABaseDir);
        if (Path <> '') and not FReader(Path, Sub) then
          Path := '';
        if Path <> '' then
        begin
          if FIncludes <> nil then FIncludes.Add(Path);
          Inc(FCount);
          var RunEnd := Start;
          if FMarked then
          begin
            // only indentation before the directive on its line: drop it
            // instead of leaving a line of blanks behind
            var LS := Start;
            while (LS > RunStart) and CharInSet(AText[LS - 1], [' ', #9]) do Dec(LS);
            if (LS = 1) or ((LS > RunStart) and CharInSet(AText[LS - 1], [#13, #10])) then
              RunEnd := LS;
          end;
          Run(AFile, AText, RunStart, RunEnd);
          if FMarked then
          begin
            EnsureLineStart;
            FSB.Append(IncludeMarkerBegin + Copy(AText, Start, I - Start) + FLineBreak);
            Expand(Path, Sub, ExtractFileDir(Path), ADepth + 1);
            EnsureLineStart;
            FSB.Append(IncludeMarkerEnd + ExtractFileName(Path) + FLineBreak);
            // blanks and ONE line break right after the directive are
            // already covered by the marker's own line end
            var K := I;
            while (K <= N) and CharInSet(AText[K], [' ', #9]) do Inc(K);
            if (K <= N) and (AText[K] = #13) then Inc(K);
            if (K <= N) and (AText[K] = #10) then Inc(K);
            if (K > N) or (K > I) and CharInSet(AText[K - 1], [#13, #10]) then
              I := K;
          end
          else
          begin
            Expand(Path, Sub, ExtractFileDir(Path), ADepth + 1);
            FSB.Append(#10);   // synthetic, belongs to no file
          end;
          RunStart := I;
        end;
      end;
      Continue;
    end;
    Inc(I);
  end;
  Run(AFile, AText, RunStart, N + 1);
end;

function ExpandIncludeText(const AText, ABaseDir: string; AIncludes: TStrings;
  const AReader: TIncludeReader): string;
var
  E: TExpander;
  R: TIncludeReader;
begin
  // cheap exit without any allocation
  if (Pos('{$I', UpperCase(AText)) = 0) and (Pos('(*$I', UpperCase(AText)) = 0) then
    Exit(AText);
  R := AReader;
  if not Assigned(R) then R := DiskReader;
  E := TExpander.Create(R, AIncludes, nil);
  try
    E.Expand('', AText, ABaseDir, 0);
    Result := E.FSB.ToString;
  finally
    E.Free;
  end;
end;

function ExpandIncludesMarked(const AText, ABaseDir: string;
  const AReader: TIncludeReader; out ACount: Integer): string;
var
  E: TExpander;
  R: TIncludeReader;
begin
  ACount := 0;
  if (Pos('{$I', UpperCase(AText)) = 0) and (Pos('(*$I', UpperCase(AText)) = 0) then
    Exit(AText);
  R := AReader;
  if not Assigned(R) then R := DiskReader;
  E := TExpander.Create(R, nil, nil);
  try
    E.FMarked := True;
    // the file's own line break style
    if Pos(#13#10, AText) > 0 then E.FLineBreak := #13#10
    else if Pos(#10, AText) > 0 then E.FLineBreak := #10
    else E.FLineBreak := sLineBreak;
    E.Expand('', AText, ABaseDir, 0);
    ACount := E.FCount;
    if ACount = 0 then Exit(AText);
    Result := E.FSB.ToString;
    // an include with other line endings than the unit would leave the file
    // MIXED - which the debugger (and the debug consistency check) dislikes.
    // A unit with ONE style gets the whole result in that style; a unit that
    // is mixed already is left as it is.
    var CrLf := 0;
    var Lf := 0;
    for var K := 1 to Length(AText) do
      if AText[K] = #10 then
        if (K > 1) and (AText[K - 1] = #13) then Inc(CrLf) else Inc(Lf);
    if (CrLf > 0) and (Lf = 0) then
      Result := Result.Replace(#13#10, #10).Replace(#13, #10).Replace(#10, #13#10)
    else if (Lf > 0) and (CrLf = 0) then
      Result := Result.Replace(#13#10, #10);
  finally
    E.Free;
  end;
end;

function ExpandUnitIncludes(const AFile, AText: string;
  const AReader: TIncludeReader): TExpandedUnit;
var
  E: TExpander;
  Incs: TStringList;
  R: TIncludeReader;
begin
  Result := nil;
  if (Pos('{$I', UpperCase(AText)) = 0) and (Pos('(*$I', UpperCase(AText)) = 0) then
    Exit;
  R := AReader;
  if not Assigned(R) then R := DiskReader;
  Incs := TStringList.Create;
  var U := TExpandedUnit.Create(ExpandFileName(AFile), AText);
  try
    E := TExpander.Create(R, Incs, U.FSources);
    try
      E.Expand(U.FFileName, AText, ExtractFileDir(U.FFileName), 0);
      if Incs.Count = 0 then Exit;   // only unresolvable directives
      U.FText := E.FSB.ToString;
      U.FSegments := E.FSegments.ToArray;
      Incs.Sorted := True;
      Incs.Duplicates := dupIgnore;
      Incs.CaseSensitive := False;
      U.FIncludes := Incs.ToStringArray;
    finally
      E.Free;
    end;
    Result := U;
    U := nil;
  finally
    U.Free;
    Incs.Free;
  end;
end;

function CollectIncludeFiles(const AFile, AText: string;
  const AReader: TIncludeReader): TArray<string>;
var
  Incs: TStringList;
begin
  Result := nil;
  if (Pos('{$I', UpperCase(AText)) = 0) and (Pos('(*$I', UpperCase(AText)) = 0) then
    Exit;
  Incs := TStringList.Create;
  try
    Incs.Sorted := True;
    Incs.Duplicates := dupIgnore;
    Incs.CaseSensitive := False;
    ExpandIncludeText(AText, ExtractFileDir(ExpandFileName(AFile)), Incs, AReader);
    Result := Incs.ToStringArray;
  finally
    Incs.Free;
  end;
end;

{ TExpandedUnit }

constructor TExpandedUnit.Create(const AFileName, AOriginal: string);
begin
  inherited Create;
  FFileName := AFileName;
  FOriginal := AOriginal;
  FSources := TDictionary<string, string>.Create;
  FLineStarts := TDictionary<string, TArray<Integer>>.Create;
end;

destructor TExpandedUnit.Destroy;
begin
  FLineStarts.Free;
  FSources.Free;
  inherited;
end;

function TExpandedUnit.Contains(const AFile: string): Boolean;
begin
  Result := FSources.ContainsKey(UpperCase(ExpandFileName(AFile)));
end;

// 0-based offsets of every line start; a line ends at CRLF, LF or a lone CR.
function TExpandedUnit.LineStartsOf(const AKey, AText: string): TArray<Integer>;
var
  L: TList<Integer>;
  I, N: Integer;
begin
  if FLineStarts.TryGetValue(AKey, Result) then Exit;
  L := TList<Integer>.Create;
  try
    L.Add(0);
    N := Length(AText);
    I := 1;
    while I <= N do
    begin
      if AText[I] = #13 then
      begin
        if (I < N) and (AText[I + 1] = #10) then Inc(I);
        L.Add(I);
      end
      else if AText[I] = #10 then
        L.Add(I);
      Inc(I);
    end;
    Result := L.ToArray;
  finally
    L.Free;
  end;
  FLineStarts.Add(AKey, Result);
end;

function TExpandedUnit.OffsetOf(const AKey, AText: string; ALine, ACol: Integer): Integer;
begin
  var S := LineStartsOf(AKey, AText);
  if (ALine < 0) or (ALine > High(S)) then Exit(-1);
  Result := S[ALine] + ACol;
end;

procedure TExpandedUnit.PosOf(const AKey, AText: string; AOffset: Integer;
  out ALine, ACol: Integer);
begin
  var S := LineStartsOf(AKey, AText);
  // binary search: last line start <= AOffset
  var Lo := 0;
  var Hi := High(S);
  while Lo < Hi do
  begin
    var Mid := (Lo + Hi + 1) div 2;
    if S[Mid] <= AOffset then Lo := Mid else Hi := Mid - 1;
  end;
  ALine := Lo;
  ACol := AOffset - S[Lo];
end;

function TExpandedUnit.ToExpanded(const AFile: string; ALine, ACol: Integer;
  out AExpLine, AExpCol: Integer): Boolean;
var
  Key, Src: string;
  Off: Integer;
begin
  Result := False;
  AExpLine := -1;
  AExpCol := -1;
  Key := UpperCase(ExpandFileName(AFile));
  if not FSources.TryGetValue(Key, Src) then Exit;
  Off := OffsetOf(Key, Src, ALine, ACol);
  if Off < 0 then Exit;
  for var Seg in FSegments do
    if SameText(Seg.FileName, ExpandFileName(AFile)) and (Off >= Seg.SrcOffset) and
       (Off < Seg.SrcOffset + Seg.Length) then
    begin
      PosOf(#0'expanded', FText, Seg.DstOffset + Off - Seg.SrcOffset, AExpLine, AExpCol);
      Exit(True);
    end;
end;

function TExpandedUnit.FromExpanded(AExpLine, AExpCol: Integer; out AFile: string;
  out ALine, ACol: Integer): Boolean;
var
  Off: Integer;
begin
  Result := False;
  AFile := '';
  ALine := -1;
  ACol := -1;
  Off := OffsetOf(#0'expanded', FText, AExpLine, AExpCol);
  if Off < 0 then Exit;
  for var Seg in FSegments do
    if (Off >= Seg.DstOffset) and (Off < Seg.DstOffset + Seg.Length) then
    begin
      var Key := UpperCase(Seg.FileName);
      AFile := Seg.FileName;
      PosOf(Key, FSources[Key], Seg.SrcOffset + Off - Seg.DstOffset, ALine, ACol);
      Exit(True);
    end;
end;

{ TLspIncludeContext }

constructor TLspIncludeContext.Create(AClient: TLspClient;
  const AReader: TIncludeReader; AWaitMs: Integer);
begin
  inherited Create;
  FClient := AClient;
  FReader := AReader;
  if not Assigned(FReader) then FReader := DiskReader;
  FWaitMs := AWaitMs;
  FUnits := TObjectList<TExpandedUnit>.Create(True);
  FOwnerOf := TDictionary<string, TExpandedUnit>.Create;
  FActive := TList<TExpandedUnit>.Create;
  FNotes := TStringList.Create;
end;

destructor TLspIncludeContext.Destroy;
begin
  try
    Restore;
  except
    // never raise out of a destructor
  end;
  FNotes.Free;
  FActive.Free;
  FOwnerOf.Free;
  FUnits.Free;
  inherited;
end;

procedure TLspIncludeContext.RegisterFiles(const AFiles: TArray<string>);
var
  Content: string;
begin
  for var F in AFiles do
  begin
    if IsIncludeFile(F) then Continue;
    if UnitOf(F) <> nil then Continue;
    if not FReader(F, Content) then Continue;
    var U := ExpandUnitIncludes(F, Content, FReader);
    if U = nil then Continue;
    FUnits.Add(U);
    for var IncF in U.Includes do
      if not FOwnerOf.ContainsKey(UpperCase(IncF)) then
        FOwnerOf.Add(UpperCase(IncF), U);   // included twice: the first owner
  end;
end;

function TLspIncludeContext.UnitOf(const AFile: string): TExpandedUnit;
begin
  var Full := ExpandFileName(AFile);
  for var U in FUnits do
    if SameText(U.FileName, Full) then Exit(U);
  Result := nil;
end;

function TLspIncludeContext.IsActive(AUnit: TExpandedUnit): Boolean;
begin
  Result := FActive.Contains(AUnit);
end;

function TLspIncludeContext.OwnsDocument(const AFile: string): Boolean;
begin
  if CanResolve(AFile) then Exit(True);
  var U := UnitOf(AFile);
  Result := (U <> nil) and IsActive(U);
end;

function TLspIncludeContext.CanResolve(const AFile: string): Boolean;
begin
  Result := FOwnerOf.ContainsKey(UpperCase(ExpandFileName(AFile)));
end;

function TLspIncludeContext.Activate(AUnit: TExpandedUnit): Boolean;
begin
  if IsActive(AUnit) then Exit(True);
  var Before := FClient.GetFileDiagnosticsVersion(AUnit.FileName);
  var T0 := GetTickCount64;
  FClient.RefreshDocumentWith(AUnit.FileName, AUnit.Text);
  FActive.Add(AUnit);
  Inc(FActivations);
  // the analysis is done when the unit's diagnostics arrive (also with none)
  while (FClient.GetFileDiagnosticsVersion(AUnit.FileName) = Before) and
        (GetTickCount64 - T0 < UInt64(FWaitMs)) do
    Sleep(50);
  var Ok := FClient.GetFileDiagnosticsVersion(AUnit.FileName) <> Before;
  FNotes.Add(Format('%s sent with %d include(s) expanded, %s after %d ms',
    [ExtractFileName(AUnit.FileName), Length(AUnit.Includes),
     IfThen(Ok, 'analysed', 'NOT analysed (timeout)'), GetTickCount64 - T0]));
  Result := True;
end;

function TLspIncludeContext.MapBack(const ALocs: TArray<TLspLocation>): TArray<TLspLocation>;
begin
  Result := Copy(ALocs);
  for var I := 0 to High(Result) do
  begin
    var U := UnitOf(TLspUri.FileUriToPath(Result[I].Uri));
    if (U = nil) or not IsActive(U) then Continue;
    var F: string;
    var L, C: Integer;
    if U.FromExpanded(Result[I].Range.Start.Line, Result[I].Range.Start.Character, F, L, C) then
    begin
      var Len := Result[I].Range.End_.Character - Result[I].Range.Start.Character;
      Result[I].Uri := TLspUri.PathToFileUri(F);
      Result[I].Range.Start.Line := L;
      Result[I].Range.Start.Character := C;
      Result[I].Range.End_.Line := L;
      Result[I].Range.End_.Character := C + Max(Len, 0);
    end;
  end;
end;

function TLspIncludeContext.Definition(const AFile: string; ALine, ACol: Integer): TArray<TLspLocation>;
var
  U: TExpandedUnit;
  EL, EC: Integer;
begin
  U := nil;
  FOwnerOf.TryGetValue(UpperCase(ExpandFileName(AFile)), U);
  if U <> nil then
    Activate(U)                 // a position INSIDE an include file
  else
  begin
    U := UnitOf(AFile);
    if (U <> nil) and not IsActive(U) then U := nil;   // plain unit
  end;
  if U = nil then
    Exit(MapBack(FClient.GotoDefinition(AFile, ALine, ACol)));
  if not U.ToExpanded(AFile, ALine, ACol, EL, EC) then Exit(nil);
  Result := MapBack(FClient.GotoDefinition(U.FileName, EL, EC));
end;

procedure TLspIncludeContext.AddTargetWithPartner(var ATargets: TLspSymbolTargets;
  const AFile: string; ALine, ACol: Integer; const AName: string);
var
  Col: Integer;
  Content: string;
begin
  ATargets.Add(AFile, ALine);
  Col := ACol;
  if AName <> '' then
    try
      if FReader(AFile, Content) then
      begin
        var Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
        if (ALine >= 0) and (ALine <= High(Lines)) then
        begin
          var C := NameColumnOnLine(Lines[ALine], AName, ACol);
          if C >= 0 then Col := C;
        end;
      end;
    except
      // keep the given column
    end;
  try
    var D := Definition(AFile, ALine, Col);
    if Length(D) > 0 then
      ATargets.Add(TLspUri.FileUriToPath(D[0].Uri), D[0].Range.Start.Line);
  except
    // no partner - the position itself is still in the set
  end;
end;

procedure TLspIncludeContext.Restore;
begin
  for var U in FActive do
    try
      FClient.RefreshDocumentWith(U.FileName, U.Original);
    except
      // the session may be gone - nothing left to restore then
    end;
  FActive.Clear;
end;

function TLspIncludeContext.NotesText: string;
begin
  Result := FNotes.Text;
end;

end.
