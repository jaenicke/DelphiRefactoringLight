(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MessagesReader;

// DIAGNOSTIC PROBE (IDE-only, not in the standalone build): explores the
// candidate paths for READING the IDE's Messages window after a compile.
// ToolsAPI's IOTAMessageServices is write-only, but three real paths
// exist (binary recon of coreide370.bpl):
//   1. IOTAModuleErrors.GetErrors - documented, per-module Error-Insight
//      diagnostics (severity 1/2/3), LSP-era behavior unverified.
//   2. coreide370.bpl EXPORTS its whole message model: unit Msglines
//      (global LineBufferList -> TLineBuffer -> TLine.GetLineText) and
//      TMessageViewForm.GetSaveStrings(TStrings, group, idx) - the
//      "Save Messages" text dump.
//   3. RTTI / published-method tables on the TMessageViewForm instance
//      (one shared VCL instance - Screen.Forms sees the form).
// This unit only DUMPS what each path yields on the running IDE into
// %TEMP%\RefactoringLight-messages.log so the real reader can be built
// on verified ground. Called from TPrewarmIdeNotifier.AfterCompile.

interface

/// <summary>Writes the diagnostic dump. Never raises.</summary>
procedure DumpMessagesWindow;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Rtti,
  Winapi.Windows, Vcl.Forms, Vcl.Controls, ToolsAPI;

const
  MaxDumpLines = 600;

var
  GLog: TStringList;

procedure Log(const S: string);
begin
  if (GLog <> nil) and (GLog.Count < MaxDumpLines) then
    GLog.Add(S);
end;

// ---------------------------------------------------------------------------
// Path 1: IOTAModuleErrors on the current module (documented ToolsAPI).
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Path 2 groundwork: which Msglines/MsgView symbols does coreide REALLY
// export on this installation? (exact mangled names for the next step)
// ---------------------------------------------------------------------------

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
  Log('--- coreide370.bpl exports (Msglines / MsgView / SaveStrings) ---');
  try
    Base := GetModuleHandle('coreide370.bpl');
    if Base = 0 then begin Log('coreide370.bpl not loaded?!'); Exit; end;
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
      if (Pos('MSGLINES', U) > 0) or (Pos('MSGVIEW', U) > 0)
        or (Pos('SAVESTRINGS', U) > 0) or (Pos('LINEBUFFER', U) > 0) then
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

// ---------------------------------------------------------------------------
// Path 3 groundwork: the TMessageViewForm instance + what RTTI exposes.
// ---------------------------------------------------------------------------

procedure DumpMessageViewForm;
var
  F: TForm;
  Found: TForm;
  Ctx: TRttiContext;
  T: TRttiType;
  N: Integer;
begin
  Log('--- TMessageViewForm via Screen.Forms + RTTI ---');
  try
    Found := nil;
    for var I := 0 to Screen.FormCount - 1 do
    begin
      F := Screen.Forms[I];
      if F.ClassName = 'TMessageViewForm' then begin Found := F; Break; end;
    end;
    if Found = nil then begin Log('TMessageViewForm not among Screen.Forms'); Exit; end;
    Log('form found: ' + Found.Name);

    // Which controls live on it (the draw tree we expect)?
    for var I := 0 to Found.ComponentCount - 1 do
    begin
      if I >= 40 then begin Log('  ...'); Break; end;
      Log('  comp: ' + Found.Components[I].ClassName + ' "' + Found.Components[I].Name + '"');
    end;

    // Published probes (MethodAddress needs only the published table).
    for var M in ['EditSelectAllClick', 'EditCopyItemClick', 'mvSaveMessagesItemClick'] do
      Log(Format('  published %s: %p', [M, Found.MethodAddress(M)]));

    // Extended RTTI: is GetSaveStrings invokable?
    T := Ctx.GetType(Found.ClassType);
    if T = nil then begin Log('no RTTI type'); Exit; end;
    N := 0;
    for var Meth in T.GetMethods do
    begin
      if not SameText(Meth.Parent.Name, 'TMessageViewForm') then Continue;
      Inc(N);
      if N <= 60 then
        Log(Format('  rtti method: %s (%s, %d params, invokable=%s)',
          [Meth.Name, TRttiEnumerationType.GetName(Meth.Visibility),
           Length(Meth.GetParameters),
           BoolToStr(Meth.HasExtendedInfo, True)]));
    end;
    Log(Format('%d own RTTI methods', [N]));
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
      DumpCoreIdeExports;
      DumpMessageViewForm;
      GLog.SaveToFile(TPath.Combine(TPath.GetTempPath, 'RefactoringLight-messages.log'));
    finally
      FreeAndNil(GLog);
    end;
  except
    // diagnostics must never disturb the IDE
  end;
end;

end.
