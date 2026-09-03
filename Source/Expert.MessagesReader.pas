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

/// <summary>Compiler messages of the last build that belong to AFile,
///  converted to diagnostics (code, message, 0-based line/col).</summary>
function ReadCompilerDiagnosticsFor(const AFile: string): TArray<TLspErrorDiag>;

/// <summary>Total number of message lines currently in the Messages
///  window (0 when the model could not be reached) - status window.</summary>
function CompilerMessageCount: Integer;

/// <summary>Why the reader is not working, '' when it is.</summary>
function MessagesReaderProblem: string;

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
{$ENDIF}

var
  GProblem: string;

function MessagesReaderProblem: string;
begin
  Result := GProblem;
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

// --- the walk ---------------------------------------------------------------

type
  TLineVisitor = reference to procedure(ALine: TObject);

// Walks every message line of every tab. Returns the number visited.
function WalkMessageLines(const AVisit: TLineVisitor): Integer;
var
  List, Buf, Ln: TObject;
  Bufs, Lines, I, J: Integer;
begin
  Result := 0;
  GProblem := '';
  List := LineBufferList;
  if List = nil then Exit;

  if not GetIntOf(List, 'GetCount', SymListCount, Bufs) then
  begin
    GProblem := 'neither RTTI nor the export ' + SymListCount + ' works';
    Exit;
  end;
  for I := 0 to Bufs - 1 do
  begin
    Buf := GetObjOf(List, 'GetItem', SymListItem, I);
    if Buf = nil then Continue;
    if not GetIntOf(Buf, 'GetCount', SymBufCount, Lines) then Continue;
    for J := 0 to Lines - 1 do
    begin
      Ln := GetObjOf(Buf, 'GetLines', SymBufLines, J);
      if Ln = nil then Continue;
      Inc(Result);
      if Assigned(AVisit) then AVisit(Ln);
    end;
  end;
end;

function CompilerMessageCount: Integer;
begin
  try
    Result := WalkMessageLines(nil);
  except
    Result := 0;
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
          Text, FileName, Code: string;
          LineNo, ColNo: Integer;
        begin
          Text := LineText(ALine);
          if Text = '' then Exit;
          Code := ExtractCode(Text);
          if Code = '' then Exit;             // not a compiler message

          // TFileMessageLine knows file/line/column exactly; plain lines
          // only carry the text (then parse it).
          FileName := LineFile(ALine);
          if not GetIntOf(ALine, 'GetLine', SymMsgLine, LineNo) then LineNo := 0;
          if not GetIntOf(ALine, 'GetColumn', SymMsgCol, ColNo) then ColNo := 0;
          if LineNo <= 0 then LineNo := ExtractLineNo(Text);
          if LineNo <= 0 then Exit;

          // Belongs to the file we were asked about?
          if FileName <> '' then
          begin
            if not SameText(FileName, AFile)
              and not SameText(ExtractFileName(FileName), Base) then Exit;
          end
          else if Pos(LowerCase(Base), LowerCase(Text)) = 0 then
            Exit;

          D := Default(TLspErrorDiag);
          D.Code := Code;
          D.Message := Text;
          D.Severity := 1;
          D.Range.Start.Line := LineNo - 1;            // 0-based
          if ColNo > 0 then D.Range.Start.Character := ColNo - 1
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
    if GProblem <> '' then Log('PROBLEM: ' + GProblem);
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

procedure DumpMessagesWindow;
begin
  try
    GLog := TStringList.Create;
    try
      Log('RefactoringLight messages probe  ' + DateTimeToStr(Now));
      DumpModuleErrors;
      DumpMessagesModel;
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

end.
