(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * The IDEA comes from the "LiveBlame" module of RAD Studio Version
 * Insight (plus), which the user pointed at
 * (github.com/mjst-legacy/delphi-versioninsight - a fork whose source
 * files carry MPL headers, although GitHub reports no repository-level
 * licence). NO code was taken from there: everything here is written
 * from scratch against the official ToolsAPI editor painting API, so
 * this is a credit to the idea, not an attribution of copied code.
 *)
unit Expert.VcsBlame;

{
  "Who last touched this line?" - blame data for the editor gutter.

  Design rules that come from the rest of this plugin:
  * git runs on a WORKER thread; the result is handed back with
    TThread.Queue and updates STATE ONLY (the painter picks it up on its
    next paint - see the notifier/queue rule in CLAUDE.md).
  * The data is keyed by (file, mtime, size). A blame belongs to the file
    ON DISK; while the editor buffer is modified the line numbers no
    longer match, and a WRONG author is worse than none - the painter
    stops drawing until the file is saved again.
  * Everything fails soft: no git, no repository, a shallow clone - the
    feature simply stays dark and BlameStatus() says why.
}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TVcsKind = (vcsNone, vcsGit, vcsSvn);

  TBlameLine = record
    Kind: TVcsKind;
    Hash: string;          // git: 40 hex (all zeros = uncommitted)
                           // svn: revision number, empty = uncommitted
    Author: string;
    AuthorTime: TDateTime; // local time
    Summary: string;
    /// <summary>What the annotation shows: 8 hex for git, r1234 for svn.</summary>
    function ShortHash: string;
    function IsUncommitted: Boolean;
  end;

  TBlameLines = TArray<TBlameLine>;

  /// <summary>One file touched by a commit, with its own diff.</summary>
  TCommitFile = record
    Path: string;
    Action: string;      // M / A / D / R (as far as the VCS tells us)
    Diff: string;
  end;

  /// <summary>A commit in the parts a log view needs - metadata, the
  ///  message, and the changed files each with its diff.</summary>
  TCommitInfo = record
    Kind: TVcsKind;
    Revision: string;
    Author: string;
    DateStr: string;
    Subject: string;
    Body: string;
    Files: TArray<TCommitFile>;
  end;

/// <summary>Parses the output of "git blame --porcelain". Pure and
///  exported for the tests. Commit metadata appears only at a commit's
///  FIRST occurrence, so it is remembered and re-used for later groups -
///  that is the whole subtlety of the format.</summary>
function ParseBlamePorcelain(const AOutput: string): TBlameLines;

/// <summary>Parses the output of: svn blame --xml
///  An entry WITHOUT a commit child is a locally modified line - svn has
///  no revision for it yet, the counterpart of git's all-zero hash.
///  Hand-written scan on purpose: Xml.XMLDoc needs COM initialisation
///  and this runs on a worker thread. Pure and tested.</summary>
function ParseSvnBlameXml(const AXml: string): TBlameLines;

/// <summary>Revision -> first line of the commit message, taken from
///  svn log --xml. svn blame does NOT carry the message, so it is
///  fetched separately and merged in. Pure.</summary>
procedure ParseSvnLogXml(const AXml: string; AMessages: TStrings);

/// <summary>Which VCS a file belongs to: the working-copy markers
///  (.git / .svn) are searched upwards from its directory. Cached.</summary>
function DetectVcs(const AFile: string): TVcsKind;

/// <summary>"3 days ago" / "in 2 months" - compact and language-neutral
///  enough for an editor annotation. Pure.</summary>
function HumanAge(const AWhen, ANow: TDateTime): string;

/// <summary>For every line of the EDITOR BUFFER, the line it came from
///  in the file on disk (1-based), or 0 when it is new or edited. Lets
///  the blame stay visible while typing instead of pausing: unchanged
///  lines keep their author, the edited span shows as "not committed".
///  Pure and tested.</summary>
function MapBufferToDiskLines(const ADisk, ABuf: TArray<string>): TArray<Integer>;

/// <summary>Blame for AFile if it is loaded AND still matches the file on
///  disk. Never blocks: a miss just means "not (yet) available".</summary>
function BlameForFile(const AFile: string; out ALines: TBlameLines): Boolean;

/// <summary>Starts a background load for AFile unless one is running or
/// the cached data is still current. Cheap to call from a paint handler
/// or a timer tick.</summary>
procedure RequestBlame(const AFile: string);

/// <summary>Drops cached data (e.g. after a commit or a checkout).</summary>
procedure InvalidateBlame(const AFile: string);

/// <summary>One line for the status window: what the last attempt did.</summary>
function BlameStatus: string;

/// <summary>Splits a unified diff into its per-file sections (git:
///  "diff --git a/x b/x", svn: "Index: x"). Pure and tested.</summary>
function SplitDiffPerFile(const ADiff: string; AKind: TVcsKind): TArray<TCommitFile>;

/// <summary>The commit of ABlame, in parts: author, date, message and
///  the changed files with their diffs. Runs the client SYNCHRONOUSLY
///  (a user action, bounded by a timeout).</summary>
function GetCommitInfo(const AFile: string; const ABlame: TBlameLine;
  out AInfo: TCommitInfo): Boolean;

/// <summary>Everything the VCS knows about the commit of ABlame: for git
///  "git show --stat -p", for svn the log entry plus the diff of that
///  revision. Runs the client SYNCHRONOUSLY (a user action, bounded by a
///  timeout) and returns its output, or '' on failure - BlameStatus then
///  says why.</summary>
function CommitDetails(const AFile: string; const ABlame: TBlameLine): string;

/// <summary>Full path of TortoiseGitProc.exe / TortoiseProc.exe, or ''
///  when that client is not installed (read from its own registry key,
///  so no guessing at install paths).</summary>
function TortoisePath(AKind: TVcsKind): string;
function TortoiseAvailable(AKind: TVcsKind): Boolean;

/// <summary>Opens the commit in TortoiseGit/TortoiseSVN. False when the
///  client is missing or refused to start - the caller then falls back
///  to the built-in view.</summary>
function TortoiseShowLog(const AFile: string; const ABlame: TBlameLine): Boolean;

/// <summary>Opens the file's blame view in TortoiseGit/TortoiseSVN, on
///  the given line.</summary>
function TortoiseShowBlame(const AFile: string; AKind: TVcsKind;
  ALine: Integer): Boolean;

/// <summary>Waits for running workers - MUST be called before the BPL
///  unloads (same rule as the live checker's workers).</summary>
procedure ShutdownBlame;

implementation

uses
  System.IOUtils, System.DateUtils, System.SyncObjs, System.Math,
  System.StrUtils, System.Win.Registry, Winapi.Windows, Winapi.ShellAPI;

{ TBlameLine }

function TBlameLine.ShortHash: string;
begin
  if Kind = vcsSvn then
  begin
    if Hash = '' then Result := '' else Result := 'r' + Hash;
    Exit;
  end;
  if Length(Hash) >= 8 then Result := Copy(Hash, 1, 8) else Result := Hash;
end;

function TBlameLine.IsUncommitted: Boolean;
var
  C: Char;
begin
  if Kind = vcsSvn then Exit(Hash = '');
  Result := Hash <> '';
  for C in Hash do
    if C <> '0' then Exit(False);
end;

// ---------------------------------------------------------------------------
//  Parsing (pure)
// ---------------------------------------------------------------------------

type
  // Internal: what the porcelain parser remembers per commit. (The public
  // TCommitInfo above is the log view's structure - different thing.)
  TCommitMeta = record
    Author: string;
    Time: TDateTime;
    Summary: string;
  end;

function IsHexHash(const S: string): Boolean;
var
  I: Integer;
begin
  Result := Length(S) = 40;
  if not Result then Exit;
  for I := 1 to 40 do
    if not CharInSet(S[I], ['0'..'9', 'a'..'f', 'A'..'F']) then Exit(False);
end;

function UnixToLocal(const ASeconds: Int64): TDateTime;
begin
  try
    Result := TTimeZone.Local.ToLocalTime(UnixToDateTime(ASeconds, True));
  except
    Result := 0;
  end;
end;

function ParseBlamePorcelain(const AOutput: string): TBlameLines;
var
  Lines: TArray<string>;
  Commits: TDictionary<string, TCommitMeta>;
  Res: TList<TBlameLine>;
  CurHash, Key, Value, L: string;
  CurLineNo, SpacePos, I: Integer;
  Info: TCommitMeta;
  BL: TBlameLine;
begin
  Result := nil;
  if AOutput = '' then Exit;
  Lines := AOutput.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);

  Commits := TDictionary<string, TCommitMeta>.Create;
  Res := TList<TBlameLine>.Create;
  try
    CurHash := '';
    CurLineNo := 0;
    for I := 0 to High(Lines) do
    begin
      L := Lines[I];
      if L = '' then Continue;

      // The CONTENT line closes a group; it is the only one starting
      // with a tab. That is where the entry for the final line is made.
      if L[1] = #9 then
      begin
        if (CurHash = '') or (CurLineNo <= 0) then Continue;
        BL := Default(TBlameLine);
        BL.Kind := vcsGit;
        BL.Hash := CurHash;
        if Commits.TryGetValue(CurHash, Info) then
        begin
          BL.Author := Info.Author;
          BL.AuthorTime := Info.Time;
          BL.Summary := Info.Summary;
        end;
        while Res.Count < CurLineNo - 1 do
          Res.Add(Default(TBlameLine));     // gap (should not happen)
        if Res.Count = CurLineNo - 1 then
          Res.Add(BL)
        else if CurLineNo - 1 < Res.Count then
          Res[CurLineNo - 1] := BL;
        CurHash := '';
        Continue;
      end;

      SpacePos := Pos(' ', L);
      if SpacePos > 0 then
      begin
        Key := Copy(L, 1, SpacePos - 1);
        Value := Copy(L, SpacePos + 1, MaxInt);
      end
      else
      begin
        Key := L;
        Value := '';
      end;

      if IsHexHash(Key) then
      begin
        // "<hash> <orig-line> <final-line> [<count>]"
        CurHash := LowerCase(Key);
        var Parts := Value.Split([' ']);
        if Length(Parts) >= 2 then
          CurLineNo := StrToIntDef(Parts[1], 0)
        else
          CurLineNo := 0;
        if not Commits.ContainsKey(CurHash) then
          Commits.Add(CurHash, Default(TCommitMeta));
        Continue;
      end;

      // Metadata of the CURRENT commit - present only at its first
      // occurrence, which is why it has to be remembered.
      if CurHash = '' then Continue;
      if not Commits.TryGetValue(CurHash, Info) then Info := Default(TCommitMeta);
      if Key = 'author' then Info.Author := Value
      else if Key = 'author-time' then Info.Time := UnixToLocal(StrToInt64Def(Value, 0))
      else if Key = 'summary' then Info.Summary := Value
      else
        Continue;
      Commits.AddOrSetValue(CurHash, Info);
    end;

    Result := Res.ToArray;
  finally
    Res.Free;
    Commits.Free;
  end;
end;

function HumanAge(const AWhen, ANow: TDateTime): string;
var
  Mins, Hours, Days, Months, Years: Int64;
  Y1, M1, D1, Y2, M2, D2: Word;
begin
  if AWhen <= 0 then Exit('');
  Mins := MinutesBetween(ANow, AWhen);
  if AWhen > ANow then Exit('just now');
  if Mins < 1 then Exit('just now');
  if Mins < 60 then Exit(Format('%d min ago', [Mins]));
  Hours := HoursBetween(ANow, AWhen);
  if Hours < 24 then Exit(Format('%d hour(s) ago', [Hours]));
  Days := DaysBetween(ANow, AWhen);
  if Days < 31 then Exit(Format('%d day(s) ago', [Days]));
  // NOT MonthsBetween/YearsBetween: both work with AVERAGE day counts
  // (30.4375 / 365.25), so exactly two years came out as "1 year ago".
  // Calendar fields are exact.
  DecodeDate(AWhen, Y1, M1, D1);
  DecodeDate(ANow, Y2, M2, D2);
  Months := (Int64(Y2) - Y1) * 12 + (Int64(M2) - M1);
  if D2 < D1 then Dec(Months);          // the current month is not full
  if Months < 12 then Exit(Format('%d month(s) ago', [Months]));
  Years := Months div 12;
  Result := Format('%d year(s) ago', [Years]);
end;

// ---------------------------------------------------------------------------
//  Subversion
// ---------------------------------------------------------------------------

// Minimal XML helpers. The input is machine-generated by svn, so a scan
// is enough - and it avoids CoInitialize on a worker thread, which a real
// XML DOM would need.

function XmlUnescape(const S: string): string;
begin
  Result := S.Replace('&lt;', '<', [rfReplaceAll])
             .Replace('&gt;', '>', [rfReplaceAll])
             .Replace('&quot;', '"', [rfReplaceAll])
             .Replace('&apos;', '''', [rfReplaceAll])
             .Replace('&#10;', #10, [rfReplaceAll])
             .Replace('&amp;', '&', [rfReplaceAll]);   // last: not re-escaped
end;

// Value of AAttr in the tag that starts at APos, '' when absent.
function XmlAttr(const S: string; ATagStart: Integer; const AAttr: string): string;
var
  TagEnd, P, Q: Integer;
  Needle: string;
begin
  Result := '';
  TagEnd := PosEx('>', S, ATagStart);
  if TagEnd = 0 then Exit;
  Needle := AAttr + '="';
  P := PosEx(Needle, S, ATagStart);
  if (P = 0) or (P > TagEnd) then Exit;
  Inc(P, Length(Needle));
  Q := PosEx('"', S, P);
  if (Q = 0) or (Q > TagEnd) then Exit;
  Result := XmlUnescape(Copy(S, P, Q - P));
end;

// Text of the first <ATag>...</ATag> between AFrom and ATo.
function XmlText(const S: string; const ATag: string; AFrom, ATo: Integer): string;
var
  P, Q, TagEnd: Integer;
begin
  Result := '';
  P := PosEx('<' + ATag, S, AFrom);
  if (P = 0) or ((ATo > 0) and (P > ATo)) then Exit;
  TagEnd := PosEx('>', S, P);
  if TagEnd = 0 then Exit;
  if S[TagEnd - 1] = '/' then Exit;          // <tag/> - empty
  Q := PosEx('</' + ATag + '>', S, TagEnd);
  if (Q = 0) or ((ATo > 0) and (Q > ATo)) then Exit;
  Result := XmlUnescape(Copy(S, TagEnd + 1, Q - TagEnd - 1));
end;

// "2026-09-03T18:12:34.123456Z" -> local TDateTime (0 when unparsable).
function ParseSvnDate(const S: string): TDateTime;
var
  Y, M, D, H, N, Sec: Integer;
  UTC: TDateTime;
begin
  Result := 0;
  if Length(S) < 19 then Exit;
  Y := StrToIntDef(Copy(S, 1, 4), 0);
  M := StrToIntDef(Copy(S, 6, 2), 0);
  D := StrToIntDef(Copy(S, 9, 2), 0);
  H := StrToIntDef(Copy(S, 12, 2), 0);
  N := StrToIntDef(Copy(S, 15, 2), 0);
  Sec := StrToIntDef(Copy(S, 18, 2), 0);
  if (Y = 0) or (M = 0) or (D = 0) then Exit;
  try
    UTC := EncodeDate(Y, M, D) + EncodeTime(H, N, Sec, 0);
    Result := TTimeZone.Local.ToLocalTime(UTC);   // svn stamps are UTC
  except
    Result := 0;
  end;
end;

function ParseSvnBlameXml(const AXml: string): TBlameLines;
var
  Res: TList<TBlameLine>;
  P, EntryEnd, CommitPos, LineNo: Integer;
  BL: TBlameLine;
begin
  Result := nil;
  if AXml = '' then Exit;
  Res := TList<TBlameLine>.Create;
  try
    P := Pos('<entry', AXml);
    while P > 0 do
    begin
      EntryEnd := PosEx('</entry>', AXml, P);
      if EntryEnd = 0 then EntryEnd := Length(AXml);
      LineNo := StrToIntDef(XmlAttr(AXml, P, 'line-number'), 0);

      BL := Default(TBlameLine);
      BL.Kind := vcsSvn;
      CommitPos := PosEx('<commit', AXml, P);
      if (CommitPos > 0) and (CommitPos < EntryEnd) then
      begin
        // No <commit> at all = the line is locally modified.
        BL.Hash := XmlAttr(AXml, CommitPos, 'revision');
        BL.Author := XmlText(AXml, 'author', CommitPos, EntryEnd);
        BL.AuthorTime := ParseSvnDate(XmlText(AXml, 'date', CommitPos, EntryEnd));
      end;

      if LineNo > 0 then
      begin
        while Res.Count < LineNo - 1 do
          Res.Add(Default(TBlameLine));
        if Res.Count = LineNo - 1 then Res.Add(BL)
        else Res[LineNo - 1] := BL;
      end;

      P := PosEx('<entry', AXml, EntryEnd);
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

procedure ParseSvnLogXml(const AXml: string; AMessages: TStrings);
var
  P, EntryEnd: Integer;
  Rev, Msg: string;
  NL: Integer;
begin
  if (AMessages = nil) or (AXml = '') then Exit;
  P := Pos('<logentry', AXml);
  while P > 0 do
  begin
    EntryEnd := PosEx('</logentry>', AXml, P);
    if EntryEnd = 0 then EntryEnd := Length(AXml);
    Rev := XmlAttr(AXml, P, 'revision');
    Msg := XmlText(AXml, 'msg', P, EntryEnd);
    // Only the FIRST line: the annotation is one line of editor space.
    NL := Pos(#10, Msg);
    if NL > 0 then Msg := Copy(Msg, 1, NL - 1);
    Msg := Trim(Msg.Replace(#13, ' '));
    if Rev <> '' then AMessages.Values[Rev] := Msg;
    P := PosEx('<logentry', AXml, EntryEnd);
  end;
end;

// ---------------------------------------------------------------------------
//  Which VCS?
// ---------------------------------------------------------------------------

var
  // Declared here because DetectVcs (just below) already needs the lock.
  GLock: TCriticalSection = nil;
  GVcsCache: TDictionary<string, TVcsKind> = nil;
  GRootCache: TDictionary<string, string> = nil;

// The working copy's marker directory (...\.git or ...\.svn), searched
// upwards from the file. '' when there is none.
function VcsMarkerDir(const AFile: string; out AKind: TVcsKind): string;
var
  Dir, Probe, Key: string;
begin
  Result := '';
  AKind := vcsNone;
  if AFile = '' then Exit;
  Dir := ExtractFileDir(AFile);
  if Dir = '' then Exit;
  Key := LowerCase(Dir);

  GLock.Enter;
  try
    if (GVcsCache <> nil) and GVcsCache.TryGetValue(Key, AKind)
      and (GRootCache <> nil) and GRootCache.TryGetValue(Key, Result) then
      Exit;
  finally
    GLock.Leave;
  end;

  Probe := Dir;
  while Probe <> '' do
  begin
    // .git is a DIRECTORY in a normal clone and a FILE in a worktree or
    // a submodule - both count.
    if TDirectory.Exists(TPath.Combine(Probe, '.git'))
      or TFile.Exists(TPath.Combine(Probe, '.git')) then
    begin
      AKind := vcsGit;
      Result := TPath.Combine(Probe, '.git');
      Break;
    end;
    if TDirectory.Exists(TPath.Combine(Probe, '.svn')) then
    begin
      AKind := vcsSvn;
      Result := TPath.Combine(Probe, '.svn');
      Break;
    end;
    var Parent := ExtractFileDir(Probe);
    if (Parent = '') or SameText(Parent, Probe) then Break;
    Probe := Parent;
  end;

  GLock.Enter;
  try
    if GVcsCache <> nil then GVcsCache.AddOrSetValue(Key, AKind);
    if GRootCache <> nil then GRootCache.AddOrSetValue(Key, Result);
  finally
    GLock.Leave;
  end;
end;

function DetectVcs(const AFile: string): TVcsKind;
begin
  VcsMarkerDir(AFile, Result);
end;

// A fingerprint of the REPOSITORY state. A commit does not touch the
// working file at all, so keying the cache on the file alone left the
// gutter showing "not committed" for lines that had just been committed
// (tester). These few files change on every commit / checkout / merge:
//   git: index, HEAD, and the reflog (logs\HEAD)
//   svn: the working copy database
// Stat-ing three paths once a second is cheap; re-running blame blindly
// would not be.
function RepoStateStamp(const AFile: string): TDateTime;
var
  Marker: string;
  Kind: TVcsKind;

  procedure Newest(const APath: string);
  var
    S: TDateTime;
  begin
    try
      if TFile.Exists(APath) then
      begin
        S := TFile.GetLastWriteTime(APath);
        if S > Result then Result := S;
      end;
    except
      // unreadable - just does not contribute
    end;
  end;

begin
  Result := 0;
  Marker := VcsMarkerDir(AFile, Kind);
  if Marker = '' then Exit;
  case Kind of
    vcsGit:
      begin
        // In a worktree/submodule ".git" is a FILE; then the real
        // directory lies elsewhere and only its own mtime is available.
        if TDirectory.Exists(Marker) then
        begin
          Newest(TPath.Combine(Marker, 'index'));
          Newest(TPath.Combine(Marker, 'HEAD'));
          Newest(TPath.Combine(Marker, 'logs' + PathDelim + 'HEAD'));
        end
        else
          Newest(Marker);
      end;
    vcsSvn:
      Newest(TPath.Combine(Marker, 'wc.db'));
  end;
end;

// ---------------------------------------------------------------------------
//  Running the VCS client
// ---------------------------------------------------------------------------

function RunCapture(const ACmdLine, ADir: string; ATimeoutMs: Cardinal;
  out AOutput: string): Boolean;
var
  SA: TSecurityAttributes;
  SI: TStartupInfo;
  PI: TProcessInformation;
  ReadPipe, WritePipe: THandle;
  Buf: array[0..8191] of Byte;
  Read: DWORD;
  Avail: DWORD;
  Cmd, DirParam: string;
  Raw: TBytesStream;
  Deadline: Cardinal;
begin
  Result := False;
  AOutput := '';
  FillChar(SA, SizeOf(SA), 0);
  SA.nLength := SizeOf(SA);
  SA.bInheritHandle := True;
  if not CreatePipe(ReadPipe, WritePipe, @SA, 0) then Exit;
  Raw := TBytesStream.Create;
  try
    FillChar(SI, SizeOf(SI), 0);
    SI.cb := SizeOf(SI);
    SI.dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
    SI.wShowWindow := SW_HIDE;
    SI.hStdOutput := WritePipe;
    SI.hStdError := WritePipe;
    SI.hStdInput := 0;

    Cmd := ACmdLine;
    DirParam := ADir;
    UniqueString(Cmd);
    if not CreateProcess(nil, PChar(Cmd), nil, nil, True,
      CREATE_NO_WINDOW, nil, PChar(DirParam), SI, PI) then
    begin
      CloseHandle(ReadPipe);
      CloseHandle(WritePipe);
      Exit;
    end;
    CloseHandle(WritePipe);   // the child owns the writing end now
    try
      Deadline := GetTickCount + ATimeoutMs;
      repeat
        Avail := 0;
        if PeekNamedPipe(ReadPipe, nil, 0, nil, @Avail, nil) and (Avail > 0) then
        begin
          if ReadFile(ReadPipe, Buf, Min(Avail, SizeOf(Buf)), Read, nil)
            and (Read > 0) then
            Raw.Write(Buf, Read);
        end
        else if WaitForSingleObject(PI.hProcess, 20) = WAIT_OBJECT_0 then
        begin
          // drain what is left
          while PeekNamedPipe(ReadPipe, nil, 0, nil, @Avail, nil) and (Avail > 0) do
          begin
            if not (ReadFile(ReadPipe, Buf, Min(Avail, SizeOf(Buf)), Read, nil)
              and (Read > 0)) then Break;
            Raw.Write(Buf, Read);
          end;
          Break;
        end;
      until GetTickCount > Deadline;

      if GetTickCount > Deadline then
        TerminateProcess(PI.hProcess, 1);
      // git speaks UTF-8; decode ONCE over the whole output.
      AOutput := TEncoding.UTF8.GetString(Raw.Bytes, 0, Raw.Size);
      Result := AOutput <> '';
    finally
      CloseHandle(PI.hThread);
      CloseHandle(PI.hProcess);
      CloseHandle(ReadPipe);
    end;
  finally
    Raw.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Cache + async loading
// ---------------------------------------------------------------------------

type
  TBlameEntry = class
    Lines: TBlameLines;
    Stamp: TDateTime;    // file mtime the blame belongs to
    Size: Int64;
    Repo: TDateTime;     // repository state the blame belongs to
    Loading: Boolean;
    Failed: Boolean;
  end;

var
  GCache: TObjectDictionary<string, TBlameEntry> = nil;
  GStatus: string = 'not used yet';
  GWorkers: Integer = 0;
  GShutdown: Boolean = False;

function NormKey(const AFile: string): string;
begin
  try
    Result := LowerCase(ExpandFileName(AFile));
  except
    Result := LowerCase(AFile);
  end;
end;

function FileStamp(const AFile: string; out AStamp: TDateTime;
  out ASize: Int64): Boolean;
begin
  Result := False;
  AStamp := 0;
  ASize := 0;
  try
    if not TFile.Exists(AFile) then Exit;
    AStamp := TFile.GetLastWriteTime(AFile);
    ASize := TFile.GetSize(AFile);
    Result := True;
  except
    Result := False;
  end;
end;

procedure SetStatus(const S: string);
begin
  GLock.Enter;
  try GStatus := S; finally GLock.Leave; end;
end;

function BlameStatus: string;
begin
  if GLock = nil then Exit('not initialised');
  GLock.Enter;
  try Result := GStatus; finally GLock.Leave; end;
end;

function MapBufferToDiskLines(const ADisk, ABuf: TArray<string>): TArray<Integer>;
var
  N, M, Pre, Suf, I: Integer;
begin
  N := Length(ADisk);
  M := Length(ABuf);
  SetLength(Result, M);
  if M = 0 then Exit;
  if N = 0 then
  begin
    for I := 0 to M - 1 do Result[I] := 0;
    Exit;
  end;

  // Common prefix and common suffix. Deliberately NOT a full diff: a
  // wrong author is worse than none, so everything between the first and
  // the last difference counts as EDITED. For the usual case - typing in
  // one place - that is exact, and for scattered edits it errs towards
  // "you touched this", never towards someone else's name.
  Pre := 0;
  while (Pre < N) and (Pre < M) and (ADisk[Pre] = ABuf[Pre]) do
    Inc(Pre);

  Suf := 0;
  while (Suf < N - Pre) and (Suf < M - Pre)
    and (ADisk[N - 1 - Suf] = ABuf[M - 1 - Suf]) do
    Inc(Suf);

  for I := 0 to M - 1 do
    if I < Pre then
      Result[I] := I + 1                      // unchanged head
    else if I >= M - Suf then
      Result[I] := N - (M - I) + 1            // unchanged tail, shifted
    else
      Result[I] := 0;                         // inside the edited span
end;

function BlameForFile(const AFile: string; out ALines: TBlameLines): Boolean;
var
  E: TBlameEntry;
  Stamp: TDateTime;
  Size: Int64;
begin
  Result := False;
  ALines := nil;
  if (GCache = nil) or (AFile = '') then Exit;
  if not FileStamp(AFile, Stamp, Size) then Exit;
  GLock.Enter;
  try
    if not GCache.TryGetValue(NormKey(AFile), E) then Exit;
    // The blame belongs to the file ON DISK. Once it changed, the line
    // numbers no longer match - report "no data" rather than a wrong
    // author.
    if (E.Stamp <> Stamp) or (E.Size <> Size) then Exit;
    if Length(E.Lines) = 0 then Exit;
    ALines := E.Lines;
    Result := True;
  finally
    GLock.Leave;
  end;
end;

procedure InvalidateBlame(const AFile: string);
begin
  if GCache = nil then Exit;
  GLock.Enter;
  try
    if AFile = '' then GCache.Clear
    else GCache.Remove(NormKey(AFile));
  finally
    GLock.Leave;
  end;
end;

procedure RequestBlame(const AFile: string);
var
  Key: string;
  E: TBlameEntry;
  Stamp, RepoNow: TDateTime;
  Size: Int64;
begin
  if GShutdown or (GCache = nil) or (AFile = '') then Exit;
  if not FileStamp(AFile, Stamp, Size) then Exit;
  RepoNow := RepoStateStamp(AFile);
  Key := NormKey(AFile);

  GLock.Enter;
  try
    if GCache.TryGetValue(Key, E) then
    begin
      if E.Loading then Exit;
      // Current means: same FILE and same REPOSITORY state. A commit
      // changes neither the file's mtime nor its size, so without the
      // second half the view stays on the pre-commit answer for good.
      if (E.Stamp = Stamp) and (E.Size = Size) and (E.Repo = RepoNow) then
        Exit;
    end
    else
    begin
      E := TBlameEntry.Create;
      GCache.Add(Key, E);
    end;
    E.Loading := True;
    E.Stamp := Stamp;
    E.Size := Size;
    E.Repo := RepoNow;
    E.Lines := nil;
    E.Failed := False;
  finally
    GLock.Leave;
  end;

  TInterlocked.Increment(GWorkers);
  try
    TThread.CreateAnonymousThread(
      procedure
      var
        Dir, Output, LogOut, Cmd: string;
        Parsed: TBlameLines;
        Ok, Reported: Boolean;
        Kind: TVcsKind;
        Ent: TBlameEntry;
      begin
        try
          Parsed := nil;
          Ok := False;
          Reported := False;
          Dir := ExtractFileDir(AFile);
          Kind := DetectVcs(AFile);
          try
            case Kind of
              vcsGit:
                begin
                  // --line-porcelain would repeat the metadata per line;
                  // the plain form is a fraction of the output and the
                  // parser remembers commits anyway.
                  Cmd := Format('git --no-pager blame --porcelain -- "%s"',
                    [ExtractFileName(AFile)]);
                  if RunCapture(Cmd, Dir, 20000, Output) then
                  begin
                    if Output.StartsWith('fatal:') or Output.StartsWith('error:') then
                    begin
                      SetStatus(Trim(Copy(Output, 1, 120)));
                      Reported := True;
                    end
                    else
                    begin
                      Parsed := ParseBlamePorcelain(Output);
                      Ok := Length(Parsed) > 0;
                    end;
                  end
                  else
                  begin
                    SetStatus('git could not be started (is it on the PATH?)');
                    Reported := True;
                  end;
                end;

              vcsSvn:
                begin
                  // NOT --xml -v: current svn rejects that combination outright
                  // (E205000, seen by a tester). The XML form already
                  // carries revision, author and date per entry.
                  Cmd := Format('svn blame --xml "%s"', [ExtractFileName(AFile)]);
                  if RunCapture(Cmd, Dir, 30000, Output) then
                  begin
                    if Pos('<blame', Output) = 0 then
                    begin
                      SetStatus(Trim(Copy(Output, 1, 120)));
                      Reported := True;
                    end
                    else
                    begin
                      Parsed := ParseSvnBlameXml(Output);
                      Ok := Length(Parsed) > 0;
                      // svn blame carries no commit MESSAGE - one extra
                      // log call supplies them for the revisions that are
                      // still in the recent history.
                      if Ok then
                      begin
                        var Msgs := TStringList.Create;
                        try
                          if RunCapture(Format('svn log --xml -l 200 "%s"',
                            [ExtractFileName(AFile)]), Dir, 30000, LogOut) then
                          begin
                            ParseSvnLogXml(LogOut, Msgs);
                            for var K := 0 to High(Parsed) do
                              if Parsed[K].Hash <> '' then
                                Parsed[K].Summary := Msgs.Values[Parsed[K].Hash];
                          end;
                        finally
                          Msgs.Free;
                        end;
                      end;
                    end;
                  end
                  else
                  begin
                    SetStatus('svn could not be started (is it on the PATH?)');
                    Reported := True;
                  end;
                end;
            else
            begin
              SetStatus('the file is not in a git or svn working copy');
              Reported := True;
            end;
            end;

            if Ok then
              SetStatus(Format('%s: %d line(s) blamed (%s)',
                [ExtractFileName(AFile), Length(Parsed),
                 IfThen(Kind = vcsGit, 'git', 'svn')]))
            else if (Kind <> vcsNone) and not Reported then
              SetStatus('the client answered, but nothing could be parsed');
          except
            on E2: Exception do
              SetStatus('git failed: ' + E2.Message);
          end;

          if GShutdown then Exit;
          GLock.Enter;
          try
            if (GCache <> nil) and GCache.TryGetValue(Key, Ent) then
            begin
              Ent.Lines := Parsed;
              Ent.Loading := False;
              Ent.Failed := not Ok;
            end;
          finally
            GLock.Leave;
          end;
        finally
          TInterlocked.Decrement(GWorkers);
        end;
      end).Start;
  except
    TInterlocked.Decrement(GWorkers);
    GLock.Enter;
    try
      if GCache.TryGetValue(Key, E) then E.Loading := False;
    finally
      GLock.Leave;
    end;
  end;
end;

// ---------------------------------------------------------------------------
//  Commit details, structured (the Tortoise-style view needs parts, not
//  one block of text)
// ---------------------------------------------------------------------------

function SplitDiffPerFile(const ADiff: string; AKind: TVcsKind): TArray<TCommitFile>;
var
  Lines: TArray<string>;
  Res: TList<TCommitFile>;
  Cur: TCommitFile;
  Body: TStringBuilder;
  Have: Boolean;
  I, P: Integer;
  L, Path: string;

  procedure Flush;
  begin
    if not Have then Exit;
    Cur.Diff := Body.ToString;
    Res.Add(Cur);
    Body.Clear;
    Have := False;
  end;

begin
  Result := nil;
  if ADiff = '' then Exit;
  Lines := ADiff.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  Res := TList<TCommitFile>.Create;
  Body := TStringBuilder.Create;
  try
    Have := False;
    for I := 0 to High(Lines) do
    begin
      L := Lines[I];
      Path := '';

      if AKind = vcsGit then
      begin
        // "diff --git a/path b/path" - the b-side is the current name.
        if L.StartsWith('diff --git ') then
        begin
          P := Pos(' b/', L);
          if P > 0 then
            Path := Copy(L, P + 3, MaxInt)
          else
            Path := Copy(L, Length('diff --git ') + 1, MaxInt);
        end;
      end
      else
      begin
        // svn writes "Index: <path>" before each file's diff.
        if L.StartsWith('Index: ') then
          Path := Trim(Copy(L, 8, MaxInt));
      end;

      if Path <> '' then
      begin
        Flush;
        Cur := Default(TCommitFile);
        Cur.Path := Trim(Path);
        Cur.Action := 'M';
        Have := True;
      end;

      if Have then
      begin
        // The header lines carry the ACTION - worth showing in the list.
        if L.StartsWith('new file mode') then Cur.Action := 'A'
        else if L.StartsWith('deleted file mode') then Cur.Action := 'D'
        else if L.StartsWith('rename from') then Cur.Action := 'R';
        Body.AppendLine(L);
      end;
    end;
    Flush;
    Result := Res.ToArray;
  finally
    Body.Free;
    Res.Free;
  end;
end;

function GetCommitInfo(const AFile: string; const ABlame: TBlameLine;
  out AInfo: TCommitInfo): Boolean;
var
  Dir, Out1, Diff: string;
  Parts: TArray<string>;
begin
  Result := False;
  AInfo := Default(TCommitInfo);
  if (AFile = '') or (ABlame.Hash = '') then Exit;
  Dir := ExtractFileDir(AFile);
  AInfo.Kind := ABlame.Kind;
  AInfo.Revision := ABlame.ShortHash;

  case ABlame.Kind of
    vcsGit:
      begin
        // Metadata in a machine-readable shape - no locale-dependent
        // parsing of "git show"'s human header.
        if not RunCapture(Format(
          'git --no-pager show -s --format=%%an%%n%%aI%%n%%s%%n%%b %s',
          [ABlame.Hash]), Dir, 30000, Out1) then
        begin
          SetStatus('git could not be started (is it on the PATH?)');
          Exit;
        end;
        Parts := Out1.Replace(#13#10, #10).Split([#10]);
        if Length(Parts) > 0 then AInfo.Author := Parts[0];
        if Length(Parts) > 1 then AInfo.DateStr := Parts[1];
        if Length(Parts) > 2 then AInfo.Subject := Parts[2];
        for var K := 3 to High(Parts) do
          if (Trim(Parts[K]) <> '') or (AInfo.Body <> '') then
            AInfo.Body := AInfo.Body + Parts[K] + sLineBreak;

        if RunCapture(Format('git --no-pager show --format= --patch %s',
          [ABlame.Hash]), Dir, 60000, Diff) then
          AInfo.Files := SplitDiffPerFile(Diff, vcsGit);
      end;

    vcsSvn:
      begin
        if not RunCapture(Format('svn log -v -r %s --xml "%s"',
          [ABlame.Hash, ExtractFileName(AFile)]), Dir, 30000, Out1) then
        begin
          SetStatus('svn could not be started (is it on the PATH?)');
          Exit;
        end;
        var Msgs := TStringList.Create;
        try
          ParseSvnLogXml(Out1, Msgs);
          AInfo.Subject := Msgs.Values[ABlame.Hash];
        finally
          Msgs.Free;
        end;
        AInfo.Author := XmlText(Out1, 'author', 1, 0);
        AInfo.DateStr := XmlText(Out1, 'date', 1, 0);
        // The changed paths live in the same log entry.
        var P := Pos('<path', Out1);
        while P > 0 do
        begin
          var E := PosEx('</path>', Out1, P);
          if E = 0 then Break;
          var TagEnd := PosEx('>', Out1, P);
          var CF := Default(TCommitFile);
          CF.Action := XmlAttr(Out1, P, 'action');
          if CF.Action = '' then CF.Action := 'M';
          CF.Path := XmlUnescape(Trim(Copy(Out1, TagEnd + 1, E - TagEnd - 1)));
          AInfo.Files := AInfo.Files + [CF];
          P := PosEx('<path', Out1, E);
        end;

        if RunCapture(Format('svn diff -c %s', [ABlame.Hash]), Dir, 60000, Diff) then
        begin
          // Merge the diffs into the path list from the log.
          var Sections := SplitDiffPerFile(Diff, vcsSvn);
          for var S in Sections do
            for var K := 0 to High(AInfo.Files) do
              if EndsText(AInfo.Files[K].Path, S.Path)
                or EndsText(S.Path, AInfo.Files[K].Path) then
              begin
                AInfo.Files[K].Diff := S.Diff;
                Break;
              end;
          // Anything the log did not list (rare) still gets shown.
          for var S in Sections do
          begin
            var Known := False;
            for var K := 0 to High(AInfo.Files) do
              if AInfo.Files[K].Diff = S.Diff then Known := True;
            if not Known then AInfo.Files := AInfo.Files + [S];
          end;
        end;
      end;
  else
    Exit;
  end;

  Result := (AInfo.Author <> '') or (Length(AInfo.Files) > 0);
end;

function CommitDetails(const AFile: string; const ABlame: TBlameLine): string;
var
  Dir, Out1, Out2: string;
begin
  Result := '';
  if (AFile = '') or (ABlame.Hash = '') then Exit;
  Dir := ExtractFileDir(AFile);
  case ABlame.Kind of
    vcsGit:
      begin
        // --stat first: "what else was in this commit" is the question
        // the gutter column provokes, and the file list answers it before
        // the diff does.
        if not RunCapture(Format(
          'git --no-pager show --stat -p --no-color %s', [ABlame.Hash]),
          Dir, 30000, Result) then
          SetStatus('git could not be started (is it on the PATH?)');
      end;
    vcsSvn:
      begin
        if RunCapture(Format('svn log -v -r %s "%s"',
          [ABlame.Hash, ExtractFileName(AFile)]), Dir, 30000, Out1) then
        begin
          if RunCapture(Format('svn diff -c %s', [ABlame.Hash]), Dir, 60000, Out2) then
            Result := Out1 + sLineBreak + Out2
          else
            Result := Out1;
        end
        else
          SetStatus('svn could not be started (is it on the PATH?)');
      end;
  end;
end;

// ---------------------------------------------------------------------------
//  TortoiseGit / TortoiseSVN as the viewer
// ---------------------------------------------------------------------------
//  Both ship a documented automation interface (TortoiseGitProc.exe /
//  TortoiseProc.exe with /command:...), and both register the FULL PATH of
//  that exe under HKLM\SOFTWARE\Tortoise*\ProcPath - so no guessing at
//  install locations, and no dependency on the PATH.

function TortoisePath(AKind: TVcsKind): string;
var
  Reg: TRegistry;
  Key: string;
begin
  Result := '';
  case AKind of
    vcsGit: Key := 'SOFTWARE\TortoiseGit';
    vcsSvn: Key := 'SOFTWARE\TortoiseSVN';
  else
    Exit;
  end;
  Reg := TRegistry.Create(KEY_READ or KEY_WOW64_64KEY);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly(Key) then
    try
      if Reg.ValueExists('ProcPath') then
        Result := Reg.ReadString('ProcPath');
    finally
      Reg.CloseKey;
    end;
  except
    Result := '';
  end;
  if (Result <> '') and not TFile.Exists(Result) then Result := '';
end;

function TortoiseAvailable(AKind: TVcsKind): Boolean;
begin
  Result := TortoisePath(AKind) <> '';
end;

function LaunchTortoise(AKind: TVcsKind; const AParams: string): Boolean;
var
  Exe: string;
begin
  Result := False;
  Exe := TortoisePath(AKind);
  if Exe = '' then Exit;
  // No waiting: the Tortoise dialogs are the user's window from here on,
  // and the IDE has nothing to do with them afterwards.
  Result := ShellExecute(0, 'open', PChar(Exe), PChar(AParams),
    PChar(ExtractFileDir(Exe)), SW_SHOWNORMAL) > 32;
  if not Result then
    SetStatus('could not start ' + ExtractFileName(Exe));
end;

function TortoiseShowLog(const AFile: string; const ABlame: TBlameLine): Boolean;
begin
  // The switches differ between the two tools; both are documented in
  // their manuals ("Automating TortoiseGit/TortoiseSVN").
  case ABlame.Kind of
    vcsGit:
      Result := LaunchTortoise(vcsGit,
        Format('/command:log /path:"%s" /rev:%s', [AFile, ABlame.Hash]));
    vcsSvn:
      Result := LaunchTortoise(vcsSvn,
        Format('/command:log /path:"%s" /startrev:%s /endrev:%s',
          [AFile, ABlame.Hash, ABlame.Hash]));
  else
    Result := False;
  end;
end;

function TortoiseShowBlame(const AFile: string; AKind: TVcsKind;
  ALine: Integer): Boolean;
begin
  case AKind of
    vcsGit:
      Result := LaunchTortoise(vcsGit,
        Format('/command:blame /path:"%s" /line:%d', [AFile, ALine]));
    vcsSvn:
      Result := LaunchTortoise(vcsSvn,
        Format('/command:blame /path:"%s" /line:%d /startrev:1 /endrev:HEAD',
          [AFile, ALine]));
  else
    Result := False;
  end;
end;

procedure ShutdownBlame;
var
  Waited: Integer;
begin
  GShutdown := True;
  Waited := 0;
  while (GWorkers > 0) and (Waited < 5000) do
  begin
    Sleep(20);
    Inc(Waited, 20);
    CheckSynchronize(10);
  end;
  CheckSynchronize(0);
end;

initialization
  GLock := TCriticalSection.Create;
  GCache := TObjectDictionary<string, TBlameEntry>.Create([doOwnsValues]);
  GVcsCache := TDictionary<string, TVcsKind>.Create;
  GRootCache := TDictionary<string, string>.Create;

finalization
  ShutdownBlame;
  FreeAndNil(GCache);
  FreeAndNil(GVcsCache);
  FreeAndNil(GRootCache);
  FreeAndNil(GLock);

end.
