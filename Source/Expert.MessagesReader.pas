(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MessagesReader;

// READS the IDE's Messages window (IDE-only) - the compiler output of the
// last build, which is the ONLY diagnostics source that still works when
// Error Insight / the IDE's LSP have given up (verified on the tester's
// machine: Structure view empty, IOTAModuleErrors -> 0 entries, while the
// Messages window listed E2003/E2066/E2250 for the very same file).
//
// WHY THIS WAY (recon + the %TEMP%\RefactoringLight-messages.log probe):
//   * IOTAMessageServices is write-only - it can add messages, never read.
//   * IOTAModuleErrors.GetErrors is the documented read path but mirrors
//     Error Insight, so it is empty in exactly the situation we care about.
//   * coreide370.bpl EXPORTS its whole message model (unit Msglines). The
//     probe confirmed on this install:
//        @Msglines@LineBufferList            (global variable)
//        TLineBufferList.GetCount / GetItem
//        TLineBuffer.GetCount / GetLines / GetGroupName
//        TLine.GetLineText / GetFileName
//        TFileMessageLine.GetLine / GetColumn
//
// RTTI was the first choice (it would keep virtual dispatch correct), but
// the IDE packages ship WITHOUT usable method RTTI for these classes -
// measured on the tester's install: "TLineBufferList has no usable RTTI
// (GetCount)". So the calls go through the EXPORTED ADDRESSES instead,
// with RTTI still tried first in case a future IDE does expose it.
//
// Calling convention: the Borland "$qqr" suffix is register/__fastcall.
//   function GetCount: Integer            -> function(Self): Integer
//   function GetLines(I: Integer): TLine  -> function(Self; I): Pointer
//   function GetLineText: string          -> the string result is a HIDDEN
//     var parameter appended after Self:  procedure(Self; var Res: string)
// Virtual overrides: an exported address is one class's implementation, so
// the right symbol is chosen from the object's ClassName chain
// (TCompilerMsgLine overrides GetLineText).
//
// Every step fails soft: no symbol -> no messages, never an exception into
// the IDE, and the reason is reported through MessagesReaderProblem.

interface

uses
  Lsp.Protocol;

type
  /// <summary>One line of the Messages window that carries a compiler
  ///  code. Line and column are 1-based, the way the compiler writes
  ///  them; 0 means the line did not say.</summary>
  TCompilerMessage = record
    FileName: string;
    Line: Integer;
    Col: Integer;
    Code: string;
    Text: string;
    /// <summary>1 error, 2 warning, 4 hint - from the code's letter, which
    ///  is the only language-independent part of the message.</summary>
    Severity: Integer;
  end;

/// <summary>Compiler messages of the last build that belong to AFile,
///  converted to diagnostics (code, message, 0-based line/col).</summary>
function ReadCompilerDiagnosticsFor(const AFile: string): TArray<TLspErrorDiag>;

/// <summary>EVERY compiler message currently in the Messages window, in
///  the order they stand there. Empty when the model could not be reached -
///  MessagesReaderProblem then says why, and an empty result must never be
///  read as "the build had nothing to say".</summary>
function ReadCompilerMessages: TArray<TCompilerMessage>;

/// <summary>Total number of message lines currently in the Messages
///  window (0 when the model could not be reached) - status window.</summary>
function CompilerMessageCount: Integer;

/// <summary>Why the reader is not working, '' when it is.</summary>
function MessagesReaderProblem: string;

/// <summary>What the last walk SAW, level by level: how many buffers the
///  model reports, what they are called, how many lines each holds and how
///  many were really readable - plus, as a cross-check, the message groups
///  the IDE itself reports through the documented IOTAMessageServices.
///  Without this a count of 0 cannot be told from "I could not read it",
///  which is exactly what cost a diagnosis once (a W1058 stood in the
///  window while this reader answered 0 and said nothing).</summary>
function MessagesReaderTrace: string;

/// <summary>Diagnostic dump of all three read paths to
///  %TEMP%\RefactoringLight-messages.log.</summary>
procedure DumpMessagesWindow;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Rtti,
  System.Character, System.Generics.Collections,
  Winapi.Windows, Vcl.Forms, Vcl.Controls, ToolsAPI;

// The two IDEs mangle their exports COMPLETELY differently (measured on
// the 64-bit IDE, which reported "export @Msglines@LineBufferList not
// found"): the 32-bit packages use the classic Borland scheme, the 64-bit
// ones the Itanium C++ ABI. Names below were read out of the shipped
// binaries (bin\coreide370.bpl / bin64\coreide370.bpl).
{$IFDEF WIN64}
const
  CoreIdeDll = 'coreide370.bpl';
  LineBufferListExport = '_ZN8Msglines14LineBufferListE';
  SymListCount = '_ZN8Msglines15TLineBufferList8GetCountEv';
  SymListItem  = '_ZN8Msglines15TLineBufferList7GetItemEi';
  SymBufCount  = '_ZN8Msglines11TLineBuffer8GetCountEv';
  SymBufLines  = '_ZN8Msglines11TLineBuffer8GetLinesEi';
  SymLineText  = '_ZN12Msglinesintf5TLine11GetLineTextEv';
  SymMsgText   = '_ZN8Msglines16TCompilerMsgLine11GetLineTextEv';
  SymLineFile  = '_ZN12Msglinesintf5TLine11GetFileNameEv';
  SymMsgFile   = '_ZN8Msglines16TFileMessageLine11GetFileNameEv';
  SymMsgLine   = '_ZN8Msglines16TFileMessageLine7GetLineEv';
  SymMsgCol    = '_ZN8Msglines16TFileMessageLine9GetColumnEv';
  // Read out of bin64\coreide370.bpl with the recipe in CLAUDE.md, not
  // from memory - the two IDEs mangle their exports completely
  // differently.
  SymListName  = '_ZN8Msglines15TLineBufferList7GetNameEi';
  SymBufGroup  = '_ZN8Msglines11TLineBuffer12GetGroupNameEv';
  SymBufErrors = '_ZN8Msglines11TLineBuffer13GetErrorCountEv';
  SymSaveStr   = '_ZN11Msgviewform16TMessageViewForm14GetSaveStringsEPN6Sys'
               + 'tem7Classes8TStringsENS1_15DelphiInterfaceIN12Msglinesint'
               + 'f13IMessageGroupEEEi';
{$ELSE}
const
  CoreIdeDll = 'coreide370.bpl';
  LineBufferListExport = '@Msglines@LineBufferList';
  SymListCount = '@Msglines@TLineBufferList@GetCount$qqrv';
  SymListItem  = '@Msglines@TLineBufferList@GetItem$qqri';
  SymBufCount  = '@Msglines@TLineBuffer@GetCount$qqrv';
  SymBufLines  = '@Msglines@TLineBuffer@GetLines$qqri';
  SymLineText  = '@Msglinesintf@TLine@GetLineText$qqrv';
  SymMsgText   = '@Msglines@TCompilerMsgLine@GetLineText$qqrv';
  SymLineFile  = '@Msglinesintf@TLine@GetFileName$qqrv';
  SymMsgFile   = '@Msglines@TFileMessageLine@GetFileName$qqrv';
  SymMsgLine   = '@Msglines@TFileMessageLine@GetLine$qqrv';
  SymMsgCol    = '@Msglines@TFileMessageLine@GetColumn$qqrv';
  SymListName  = '@Msglines@TLineBufferList@GetName$qqri';
  SymBufGroup  = '@Msglines@TLineBuffer@GetGroupName$qqrv';
  SymBufErrors = '@Msglines@TLineBuffer@GetErrorCount$qqrv';
  SymSaveStr   = '@Msgviewform@TMessageViewForm@GetSaveStrings$qqrp23System'
               + '@Classes@TStrings53System@%DelphiInterface$26Msglinesintf'
               + '@IMessageGroup%i';
{$ENDIF}

var
  GProblem: string;
  GTrace: string;

function MessagesReaderProblem: string;
begin
  Result := GProblem;
end;

function MessagesReaderTrace: string;
begin
  Result := GTrace;
end;

// The global TLineBufferList instance, or nil.
function LineBufferList: TObject;
var
  H: HMODULE;
  P: Pointer;
begin
  Result := nil;
  H := GetModuleHandle(CoreIdeDll);
  if H = 0 then
  begin
    GProblem := CoreIdeDll + ' not loaded';
    Exit;
  end;
  P := GetProcAddress(H, LineBufferListExport);
  if P = nil then
  begin
    GProblem := 'export ' + LineBufferListExport + ' not found';
    Exit;
  end;
  // Exported VARIABLE: the address points at the object reference.
  Result := TObject(PPointer(P)^);
  if Result = nil then GProblem := 'message model not initialised';
end;

// --- tiny RTTI helpers (fail soft, never raise) -----------------------------

function CallInt(AObj: TObject; const AMethod: string; out AValue: Integer;
  const AArgs: array of TValue): Boolean;
var
  Ctx: TRttiContext;
  T: TRttiType;
  M: TRttiMethod;
begin
  Result := False;
  AValue := 0;
  if AObj = nil then Exit;
  try
    T := Ctx.GetType(AObj.ClassType);
    if T = nil then Exit;
    M := T.GetMethod(AMethod);
    if M = nil then Exit;
    AValue := M.Invoke(AObj, AArgs).AsInteger;
    Result := True;
  except
    Result := False;
  end;
end;

function CallObj(AObj: TObject; const AMethod: string; AIndex: Integer): TObject;
var
  Ctx: TRttiContext;
  T: TRttiType;
  M: TRttiMethod;
  V: TValue;
begin
  Result := nil;
  if AObj = nil then Exit;
  try
    T := Ctx.GetType(AObj.ClassType);
    if T = nil then Exit;
    M := T.GetMethod(AMethod);
    if M = nil then Exit;
    if AIndex >= 0 then V := M.Invoke(AObj, [AIndex])
    else V := M.Invoke(AObj, []);
    if V.IsObject then Result := V.AsObject;
  except
    Result := nil;
  end;
end;

function CallStr(AObj: TObject; const AMethod: string): string;
var
  Ctx: TRttiContext;
  T: TRttiType;
  M: TRttiMethod;
begin
  Result := '';
  if AObj = nil then Exit;
  try
    T := Ctx.GetType(AObj.ClassType);
    if T = nil then Exit;
    M := T.GetMethod(AMethod);
    if M = nil then Exit;
    Result := M.Invoke(AObj, []).AsString;
  except
    Result := '';
  end;
end;

// --- calls through the exported (mangled) addresses -------------------------

type
  TIntGetter = function(Self: Pointer): Integer; register;
  TItemGetter = function(Self: Pointer; AIndex: Integer): Pointer; register;
  TStrGetter = procedure(Self: Pointer; var ARes: string); register;
  // function GetName(AIndex: Integer): string - the hidden result
  // parameter is appended AFTER the declared ones, so register puts
  // Self, AIndex, Result in that order.
  TStrAtGetter = procedure(Self: Pointer; AIndex: Integer;
    var ARes: string); register;

function Sym(const AName: string): Pointer;
var
  H: HMODULE;
begin
  Result := nil;
  H := GetModuleHandle(CoreIdeDll);
  if H <> 0 then Result := GetProcAddress(H, PChar(AName));
end;

// True when AObj's class or any ancestor is named AName - we cannot
// 'is'-cast to a class that lives in an IDE-internal package.
function ClassIsOrDescends(AObj: TObject; const AName: string): Boolean;
var
  C: TClass;
begin
  Result := False;
  if AObj = nil then Exit;
  C := AObj.ClassType;
  while C <> nil do
  begin
    if SameText(C.ClassName, AName) then Exit(True);
    C := C.ClassParent;
  end;
end;

function ExpInt(AObj: TObject; const ASym: string; out AValue: Integer): Boolean;
var
  P: Pointer;
begin
  Result := False;
  AValue := 0;
  P := Sym(ASym);
  if (P = nil) or (AObj = nil) then Exit;
  try
    AValue := TIntGetter(P)(AObj);
    Result := True;
  except
    Result := False;
  end;
end;

function ExpItem(AObj: TObject; const ASym: string; AIndex: Integer): TObject;
var
  P: Pointer;
begin
  Result := nil;
  P := Sym(ASym);
  if (P = nil) or (AObj = nil) then Exit;
  try
    Result := TObject(TItemGetter(P)(AObj, AIndex));
  except
    Result := nil;
  end;
end;

function ExpStrAt(AObj: TObject; const ASym: string;
  AIndex: Integer): string;
var
  P: Pointer;
begin
  Result := '';
  P := Sym(ASym);
  if (P = nil) or (AObj = nil) then Exit;
  try
    TStrAtGetter(P)(AObj, AIndex, Result);
  except
    Result := '';
  end;
end;

function ExpStr(AObj: TObject; const ASym: string): string;
var
  P: Pointer;
begin
  Result := '';
  P := Sym(ASym);
  if (P = nil) or (AObj = nil) then Exit;
  try
    TStrGetter(P)(AObj, Result);
  except
    Result := '';
  end;
end;

// --- one accessor per fact, RTTI first, exported address as fallback -------

function GetIntOf(AObj: TObject; const AMethod, ASym: string;
  out AValue: Integer): Boolean;
begin
  Result := CallInt(AObj, AMethod, AValue, []);
  if not Result then Result := ExpInt(AObj, ASym, AValue);
end;

function GetObjOf(AObj: TObject; const AMethod, ASym: string;
  AIndex: Integer): TObject;
begin
  Result := CallObj(AObj, AMethod, AIndex);
  if Result = nil then Result := ExpItem(AObj, ASym, AIndex);
end;

function LineText(AObj: TObject): string;
begin
  Result := CallStr(AObj, 'GetLineText');
  if Result <> '' then Exit;
  // Pick the implementation that belongs to this object's class.
  if ClassIsOrDescends(AObj, 'TCompilerMsgLine') then
    Result := ExpStr(AObj, SymMsgText);
  if Result = '' then Result := ExpStr(AObj, SymLineText);
end;

function LineFile(AObj: TObject): string;
begin
  Result := CallStr(AObj, 'GetFileName');
  if Result <> '' then Exit;
  if ClassIsOrDescends(AObj, 'TFileMessageLine') then
    Result := ExpStr(AObj, SymMsgFile);
  if Result = '' then Result := ExpStr(AObj, SymLineFile);
end;

// --- message text parsing ---------------------------------------------------

// Compiler codes are language independent: a letter class (E/W/H/F) plus
// four digits, e.g. "E2003" in
//   [dcc32 Fehler] FrmPPFrame.pas(3721): E2003 Undeklarierter Bezeichner...
function ExtractCode(const AText: string): string;
var
  I, J: Integer;
  Digits: Boolean;
begin
  Result := '';
  for I := 1 to Length(AText) do
  begin
    if not CharInSet(AText[I], ['E', 'W', 'H', 'F']) then Continue;
    // must start a token
    if (I > 1) and AText[I - 1].IsLetterOrDigit then Continue;
    if I + 4 > Length(AText) then Break;
    Digits := True;
    for J := I + 1 to I + 4 do
      if not AText[J].IsDigit then
      begin
        Digits := False;
        Break;
      end;
    if not Digits then Continue;
    // and must END there ("E20031" is not a code)
    if (I + 5 <= Length(AText)) and AText[I + 5].IsDigit then Continue;
    Exit(Copy(AText, I, 5));
  end;
end;

// "...FrmPPFrame.pas(3721): E2003 ..." -> 3721. 0 when absent. Used only
// as a fallback: TFileMessageLine.GetLine is the authoritative source.
function ExtractLineNo(const AText: string): Integer;
var
  P, Q: Integer;
begin
  Result := 0;
  P := Pos('(', AText);
  while P > 0 do
  begin
    Q := P + 1;
    while (Q <= Length(AText)) and AText[Q].IsDigit do Inc(Q);
    if (Q > P + 1) and (Q <= Length(AText)) and (AText[Q] = ')') then
      Exit(StrToIntDef(Copy(AText, P + 1, Q - P - 1), 0));
    P := Pos('(', AText, P + 1);
  end;
end;

// "[dcc32 Fehler] FrmPPFrame.pas(3721): E2003 ..." -> "FrmPPFrame.pas": the
// token directly in front of the first "(<digits>)". A file name that only
// appears LATER in the text ("Other.pas(10): F2051 Unit Foo.pas was
// compiled...") is not the file the message is about. '' when absent.
function ExtractFilePart(const AText: string): string;
var
  P, Q, S: Integer;
begin
  Result := '';
  P := Pos('(', AText);
  while P > 0 do
  begin
    Q := P + 1;
    while (Q <= Length(AText)) and AText[Q].IsDigit do Inc(Q);
    if (Q > P + 1) and (Q <= Length(AText)) and (AText[Q] = ')') then
    begin
      S := P - 1;
      while (S >= 1) and not CharInSet(AText[S], [' ', #9, ']', '"', '''']) do Dec(S);
      Exit(Copy(AText, S + 1, P - S - 1));
    end;
    P := Pos('(', AText, P + 1);
  end;
end;

// LSP severity from the language-independent code letter: E/F error,
// W warning, H hint. Every message used to become an ERROR, so a clean
// build with 40 H2077 hints reported 40 errors.
function SeverityOfCode(const ACode: string): Integer;
begin
  Result := 1;
  if ACode = '' then Exit;
  case UpCase(ACode[1]) of
    'W': Result := 2;
    'H': Result := 4;
  end;
end;

// --- the walk ---------------------------------------------------------------

type
  TLineVisitor = reference to procedure(ALine: TObject);

// The name a buffer goes by in the window - the LIST knows it per index,
// the buffer itself knows its group. Either is enough to recognise the
// Build tab in the trace.
function BufferLabel(AList, ABuf: TObject; AIndex: Integer): string;
begin
  Result := ExpStrAt(AList, SymListName, AIndex);
  if Result = '' then Result := CallStr(ABuf, 'GetGroupName');
  if Result = '' then Result := ExpStr(ABuf, SymBufGroup);
  if Result = '' then Result := '#' + IntToStr(AIndex);
end;

// The DOCUMENTED cross-check (IOTAMessageServices): the groups the IDE
// itself reports. When it names a tab that the model above does not have,
// the window's lines are not where this reader looks - and that is a
// statement no walk of the model alone can make.
function IdeMessageGroupNames: string;
var
  Svc: IOTAMessageServices;
  I: Integer;
  G: IOTAMessageGroup;
begin
  Result := '';
  // ToolsAPI is main thread only, and the status tick is not the only
  // caller of the walk.
  if TThread.CurrentThread.ThreadID <> MainThreadID then Exit('not asked');
  try
    if not Supports(BorlandIDEServices, IOTAMessageServices, Svc) then
      Exit('no IOTAMessageServices');
    for I := 0 to Svc.GetMessageGroupCount - 1 do
    begin
      G := Svc.GetMessageGroup(I);
      if G = nil then Continue;
      if Result <> '' then Result := Result + ', ';
      Result := Result + G.Name;
    end;
    if Result = '' then Result := 'none';
  except
    on E: Exception do Result := E.ClassName;
  end;
end;

// Walks every message line of every tab. Returns the number visited - and
// records what it saw on the way, because a 0 that explains itself is the
// difference between "the window is empty" and "I cannot read it".
function WalkMessageLines(const AVisit: TLineVisitor): Integer;
var
  List, Buf, Ln: TObject;
  Bufs, Lines, Reported, Unreadable, NoLine, I, J: Integer;
  Detail: string;
begin
  Result := 0;
  GProblem := '';
  GTrace := '';
  List := LineBufferList;
  if List = nil then
  begin
    GTrace := 'the model itself is not reachable: ' + GProblem;
    Exit;
  end;

  if not GetIntOf(List, 'GetCount', SymListCount, Bufs) then
  begin
    GProblem := 'neither RTTI nor the export ' + SymListCount + ' works';
    GTrace := GProblem;
    Exit;
  end;

  Reported := 0;
  Unreadable := 0;
  NoLine := 0;
  Detail := '';
  for I := 0 to Bufs - 1 do
  begin
    Buf := GetObjOf(List, 'GetItem', SymListItem, I);
    if Buf = nil then
    begin
      Inc(Unreadable);
      Continue;
    end;
    if not GetIntOf(Buf, 'GetCount', SymBufCount, Lines) then
    begin
      Inc(Unreadable);
      Continue;
    end;
    Inc(Reported, Lines);
    if Detail <> '' then Detail := Detail + ', ';
    Detail := Detail + BufferLabel(List, Buf, I) + '=' + IntToStr(Lines);
    for J := 0 to Lines - 1 do
    begin
      Ln := GetObjOf(Buf, 'GetLines', SymBufLines, J);
      if Ln = nil then
      begin
        Inc(NoLine);
        Continue;
      end;
      Inc(Result);
      if Assigned(AVisit) then AVisit(Ln);
    end;
  end;

  GTrace := Format('%d buffer(s)', [Bufs]);
  if Detail <> '' then GTrace := GTrace + ' (' + Detail + ')';
  GTrace := GTrace + Format(', %d line(s) read', [Result]);
  if Unreadable > 0 then
    GTrace := GTrace + Format(', %d buffer(s) unreadable', [Unreadable]);
  if NoLine > 0 then
    GTrace := GTrace + Format(', %d line(s) unreadable', [NoLine]);
  GTrace := GTrace + ' | IDE groups: ' + IdeMessageGroupNames;

  // Name the two shapes that used to be silent. Neither is certainly a
  // defect - the window really can be empty - so the sentence says what
  // was seen and leaves the reading to whoever looks at it.
  if Unreadable > 0 then
    GProblem := Format('%d of %d buffer(s) could not be read', [Unreadable, Bufs])
  else if Bufs = 0 then
    GProblem := 'the model reports no buffer at all - either the window is '
      + 'empty or its lines do not live in Msglines.LineBufferList on this IDE'
  else if Reported = 0 then
    GProblem := Format('%d buffer(s) (%s), every one of them empty',
      [Bufs, Detail])
  else if Result < Reported then
    GProblem := Format('%d of %d line(s) could not be read',
      [Reported - Result, Reported]);
end;

function CompilerMessageCount: Integer;
begin
  try
    Result := WalkMessageLines(nil);
  except
    Result := 0;
  end;
end;

// ONE conversion for both readers. False when the line is no compiler
// message at all - that is the only thing the two callers filter on
// differently (one of them additionally by file).
function MessageOfLine(ALine: TObject; out AMsg: TCompilerMessage): Boolean;
begin
  Result := False;
  AMsg := Default(TCompilerMessage);
  AMsg.Text := LineText(ALine);
  if AMsg.Text = '' then Exit;
  AMsg.Code := ExtractCode(AMsg.Text);
  if AMsg.Code = '' then Exit;             // not a compiler message
  AMsg.Severity := SeverityOfCode(AMsg.Code);

  // TFileMessageLine knows file/line/column exactly; plain lines
  // only carry the text (then parse it). The line/column getters
  // are called ONLY on such a line: the export fallback is a raw
  // method pointer, and on a TLine / TCompilerMsgLine it would read
  // a field the object does not have (garbage line, or an AV).
  AMsg.FileName := LineFile(ALine);
  if ClassIsOrDescends(ALine, 'TFileMessageLine') then
  begin
    if not GetIntOf(ALine, 'GetLine', SymMsgLine, AMsg.Line) then AMsg.Line := 0;
    if not GetIntOf(ALine, 'GetColumn', SymMsgCol, AMsg.Col) then AMsg.Col := 0;
  end;
  if AMsg.Line <= 0 then AMsg.Line := ExtractLineNo(AMsg.Text);
  // The message's own "<file>(<line>)" prefix decides - not any file name
  // that happens to occur somewhere in the text.
  if AMsg.FileName = '' then AMsg.FileName := ExtractFilePart(AMsg.Text);
  Result := True;
end;

function ReadCompilerMessages: TArray<TCompilerMessage>;
var
  Res: TList<TCompilerMessage>;
begin
  Result := nil;
  Res := TList<TCompilerMessage>.Create;
  try
    try
      WalkMessageLines(
        procedure(ALine: TObject)
        var
          M: TCompilerMessage;
        begin
          if MessageOfLine(ALine, M) then Res.Add(M);
        end);
    except
      // never disturb the IDE
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

function ReadCompilerDiagnosticsFor(const AFile: string): TArray<TLspErrorDiag>;
var
  Res: TList<TLspErrorDiag>;
  Base: string;
begin
  Result := nil;
  if AFile = '' then Exit;
  Base := ExtractFileName(AFile);
  Res := TList<TLspErrorDiag>.Create;
  try
    try
      WalkMessageLines(
        procedure(ALine: TObject)
        var
          D: TLspErrorDiag;
          M: TCompilerMessage;
        begin
          if not MessageOfLine(ALine, M) then Exit;
          if M.Line <= 0 then Exit;

          // Belongs to the file we were asked about? A full path must be
          // THIS path (two Utils.pas in different folders must not mix);
          // only a bare name can be compared by name.
          if M.FileName = '' then Exit;
          if ExtractFilePath(M.FileName) <> '' then
          begin
            if not SameText(ExpandFileName(M.FileName), ExpandFileName(AFile)) then Exit;
          end
          else if not SameText(M.FileName, Base) then
            Exit;

          D := Default(TLspErrorDiag);
          D.Code := M.Code;
          D.Message := M.Text;
          D.Severity := M.Severity;
          D.Range.Start.Line := M.Line - 1;            // 0-based
          if M.Col > 0 then D.Range.Start.Character := M.Col - 1
          else D.Range.Start.Character := 0;
          D.Range.End_ := D.Range.Start;
          Res.Add(D);
        end);
    except
      // never disturb the IDE
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

// ---------------------------------------------------------------------------
//  Diagnostic dump (kept: it is how this reader was built, and it tells us
//  what changed if a future IDE version renames something)
// ---------------------------------------------------------------------------

const
  MaxDumpLines = 600;

var
  GLog: TStringList;

procedure Log(const S: string);
begin
  if (GLog <> nil) and (GLog.Count < MaxDumpLines) then
    GLog.Add(S);
end;

procedure DumpModuleErrors;
var
  MS: IOTAModuleServices;
  ME: IOTAModuleErrors;
  Errs: TOTAErrors;
begin
  Log('--- IOTAModuleErrors (documented API) ---');
  try
    if not Supports(BorlandIDEServices, IOTAModuleServices, MS) then Exit;
    if (MS.CurrentModule = nil) then begin Log('no current module'); Exit; end;
    Log('module: ' + MS.CurrentModule.FileName);
    if not Supports(MS.CurrentModule, IOTAModuleErrors, ME) then
    begin
      Log('IOTAModuleErrors NOT supported by the module');
      Exit;
    end;
    Errs := ME.GetErrors('');
    Log(Format('GetErrors -> %d entries', [Length(Errs)]));
    for var I := 0 to High(Errs) do
    begin
      if I >= 25 then begin Log('...'); Break; end;
      Log(Format('  [sev %d] %d:%d-%d:%d %s',
        [Errs[I].Severity, Errs[I].Start.Line, Errs[I].Start.CharIndex,
         Errs[I].Stop.Line, Errs[I].Stop.CharIndex, Errs[I].Text]));
    end;
  except
    on E: Exception do Log('EX: ' + E.ClassName + ': ' + E.Message);
  end;
end;

procedure DumpMessagesModel;
var
  N: Integer;
begin
  Log('--- Msglines model (the reader) ---');
  try
    N := 0;
    WalkMessageLines(
      procedure(ALine: TObject)
      var
        Text: string;
        LineNo: Integer;
      begin
        Inc(N);
        if N > 40 then Exit;
        Text := LineText(ALine);
        if not GetIntOf(ALine, 'GetLine', SymMsgLine, LineNo) then LineNo := -1;
        Log(Format('  %s | line=%d | file=%s | %s',
          [ALine.ClassName, LineNo, LineFile(ALine), Text]));
      end);
    Log(Format('total message lines: %d', [N]));
    Log('trace: ' + GTrace);
    if GProblem <> '' then Log('PROBLEM: ' + GProblem);
  except
    on E: Exception do Log('EX: ' + E.ClassName + ': ' + E.Message);
  end;
end;

type
  // procedure GetSaveStrings(AStrings: TStrings;
  //   const AGroup: IMessageGroup; AFlags: Integer)
  TSaveStringsProc = procedure(Self, AStrings, AGroup: Pointer;
    AFlags: Integer); register;

// The SECOND route, as a diagnostic only: the "Save Messages" text dump of
// the message view form. Worth probing because it is what the window shows
// rather than what the model holds - so if the model is empty and this is
// not, the lines live somewhere else on this IDE.
// DELIBERATELY only in this dump: it calls an internal IDE method on an
// object whose lifetime we do not control, which CLAUDE.md allows as an
// occasional diagnostic and nothing more (once a second it cost this
// plugin an access violation per tick).
procedure DumpSaveStrings;
var
  P: Pointer;
  Form: TForm;
  I: Integer;
  SL: TStringList;
begin
  Log('--- TMessageViewForm.GetSaveStrings (the second route) ---');
  try
    P := Sym(SymSaveStr);
    if P = nil then
    begin
      Log('export not found: ' + SymSaveStr);
      Exit;
    end;
    // By CLASS NAME out of Screen.Forms: no second exported global, and
    // the object proves what it is.
    Form := nil;
    for I := 0 to Screen.FormCount - 1 do
      if SameText(Screen.Forms[I].ClassName, 'TMessageViewForm') then
      begin
        Form := Screen.Forms[I];
        Break;
      end;
    if Form = nil then
    begin
      Log('no TMessageViewForm among the ' + IntToStr(Screen.FormCount)
        + ' forms on screen (is the Messages window open?)');
      Exit;
    end;
    SL := TStringList.Create;
    try
      // nil group = every group, if the method reads it that way. An
      // exception here is an answer too, and it is caught.
      TSaveStringsProc(P)(Form, SL, nil, 0);
      Log(Format('%d line(s)', [SL.Count]));
      for I := 0 to SL.Count - 1 do
      begin
        if I >= 40 then begin Log('  ... (capped)'); Break; end;
        Log('  ' + SL[I]);
      end;
    finally
      SL.Free;
    end;
  except
    on E: Exception do Log('EX: ' + E.ClassName + ': ' + E.Message);
  end;
end;

procedure DumpCoreIdeExports;
var
  Base: HMODULE;
  Dos: PImageDosHeader;
  Nt: PImageNtHeaders;
  ExpRva, ExpSize: Cardinal;
  Exp: PImageExportDirectory;
  Names: PCardinal;
  I, Hits: Integer;
  Nm: AnsiString;
  U: string;
begin
  Log('--- coreide exports (Msglines / MsgView) ---');
  try
    Base := GetModuleHandle(CoreIdeDll);
    if Base = 0 then begin Log(CoreIdeDll + ' not loaded?!'); Exit; end;
    Dos := PImageDosHeader(Base);
    Nt := PImageNtHeaders(PByte(Base) + Dos._lfanew);
    ExpRva := Nt.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress;
    ExpSize := Nt.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].Size;
    if (ExpRva = 0) or (ExpSize = 0) then begin Log('no export dir'); Exit; end;
    Exp := PImageExportDirectory(PByte(Base) + ExpRva);
    Names := PCardinal(PByte(Base) + Exp.AddressOfNames);
    Hits := 0;
    for I := 0 to Integer(Exp.NumberOfNames) - 1 do
    begin
      Nm := PAnsiChar(PByte(Base) + PCardinal(PByte(Names) + I * 4)^);
      U := UpperCase(string(Nm));
      if (Pos('MSGLINES', U) > 0) or (Pos('MSGVIEW', U) > 0) then
      begin
        Log('  ' + string(Nm));
        Inc(Hits);
        if Hits >= 250 then begin Log('  ... (capped)'); Break; end;
      end;
    end;
    Log(Format('%d matching exports of %d total', [Hits, Exp.NumberOfNames]));
  except
    on E: Exception do Log('EX: ' + E.ClassName + ': ' + E.Message);
  end;
end;

var
  GDumpDone: Boolean = False;
  // Keeps the process-wide RTTI pool alive: every Call* helper declares a
  // local TRttiContext, and without a long-lived one the pool was built
  // and torn down again for each of them.
  GRttiKeepAlive: TRttiContext;

procedure DumpMessagesWindow;
begin
  // ONCE per session: the walk covers ~29,000 exported names on the main
  // thread, and the exports of a loaded module cannot change - running it
  // after every compile only cost time.
  if GDumpDone then Exit;
  GDumpDone := True;
  try
    GLog := TStringList.Create;
    try
      Log('RefactoringLight messages probe  ' + DateTimeToStr(Now));
      DumpModuleErrors;
      DumpMessagesModel;
      DumpSaveStrings;
      DumpCoreIdeExports;
      GLog.SaveToFile(TPath.Combine(TPath.GetTempPath,
        'RefactoringLight-messages.log'));
    finally
      FreeAndNil(GLog);
    end;
  except
    // diagnostics must never disturb the IDE
  end;
end;

initialization
  GRttiKeepAlive := TRttiContext.Create;

finalization
  GRttiKeepAlive.Free;

end.
