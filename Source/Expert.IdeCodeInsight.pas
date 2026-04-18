(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IdeCodeInsight;

interface

uses
  System.SysUtils, System.Classes, System.SyncObjs, Winapi.Windows, Vcl.Forms, ToolsAPI;

type
  TDefinitionResult = record
    FileName: string;
    Line: Integer;
    CharIndex: Integer;
    Error: Boolean;
    ErrorMessage: string;
    IsValid: Boolean;
  end;

  /// <summary>
  ///  Uses the IDE-internal Code Insight manager (= the already running
  ///  DelphiLSP) for GotoDefinition queries. No external process needed,
  ///  immediate results.
  /// </summary>
  TIdeCodeInsight = class
  private
    class var FResult: TDefinitionResult;
    class var FEvent: TLightweightEvent;
    class procedure GotoDefCallback(Sender: TObject; AId: Integer; const AFileName: string; ALine, ACharIndex: Integer;
      AError: Boolean; const AMessage: string);
    class procedure GotoDefCallbackLegacy(Sender: TObject; AId: Integer; const AFileName: string; ALine: Integer;
      AError: Boolean; const AMessage: string);
  public
    /// <summary>Calls GotoDefinition via the IDE-internal LSP.
    ///  Returns the result synchronously (waits for the callback).</summary>
    class function GotoDefinition(const AFileName: string; ALine, ACol: Integer; ATimeoutMs: Cardinal = 10000): TDefinitionResult;

    /// <summary>Checks whether the IDE Code Insight manager is available.</summary>
    class function IsAvailable: Boolean;

    /// <summary>Checks whether Code Insight is ready (no longer indexing).</summary>
    class function IsReady: Boolean;

    /// <summary>Waits until Code Insight is ready. Returns False on timeout.</summary>
    class function WaitUntilReady(ATimeoutMs: Cardinal = 30000): Boolean;
  end;

implementation

class procedure TIdeCodeInsight.GotoDefCallback(Sender: TObject; AId: Integer;
  const AFileName: string; ALine, ACharIndex: Integer; AError: Boolean; const AMessage: string);
begin
  FResult.FileName := AFileName;
  FResult.Line := ALine;
  FResult.CharIndex := ACharIndex;
  FResult.Error := AError;
  FResult.ErrorMessage := AMessage;
  FResult.IsValid := (not AError) and (AFileName <> '');
  if FEvent <> nil then
    FEvent.SetEvent;
end;

class procedure TIdeCodeInsight.GotoDefCallbackLegacy(Sender: TObject; AId: Integer;
  const AFileName: string; ALine: Integer; AError: Boolean; const AMessage: string);
begin
  FResult.FileName := AFileName;
  FResult.Line := ALine;
  FResult.CharIndex := 0;
  FResult.Error := AError;
  FResult.ErrorMessage := AMessage;
  FResult.IsValid := (not AError) and (AFileName <> '');
  if FEvent <> nil then
    FEvent.SetEvent;
end;

class function TIdeCodeInsight.IsAvailable: Boolean;
var
  CIServices: IOTACodeInsightServices;
  CIManager: IOTACodeInsightManager;
  AsyncMgr: IOTAAsyncCodeInsightManager;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTACodeInsightServices, CIServices) then
    Exit;
  CIServices.GetCurrentCodeInsightManager(CIManager);
  if CIManager = nil then
  begin
    // No current manager — try all of them
    for var I := 0 to CIServices.CodeInsightManagerCount - 1 do
    begin
      CIManager := CIServices.CodeInsightManager[I];
      if Supports(CIManager, IOTAAsyncCodeInsightManager, AsyncMgr) then
        Exit(True);
    end;
    Exit;
  end;
  Result := Supports(CIManager, IOTAAsyncCodeInsightManager);
end;

class function TIdeCodeInsight.IsReady: Boolean;
var
  CIServices: IOTACodeInsightServices;
  CIManager: IOTACodeInsightManager;
  AsyncMgr: IOTAAsyncCodeInsightManager;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTACodeInsightServices, CIServices) then
    Exit;

  // Try the current manager
  CIServices.GetCurrentCodeInsightManager(CIManager);
  if CIManager = nil then
  begin
    for var I := 0 to CIServices.CodeInsightManagerCount - 1 do
    begin
      CIManager := CIServices.CodeInsightManager[I];
      if Supports(CIManager, IOTAAsyncCodeInsightManager, AsyncMgr) then
      begin
        Result := AsyncMgr.AsyncEnabled;
        Exit;
      end;
    end;
    Exit;
  end;

  if Supports(CIManager, IOTAAsyncCodeInsightManager, AsyncMgr) then
    Result := AsyncMgr.AsyncEnabled;
end;

class function TIdeCodeInsight.WaitUntilReady(ATimeoutMs: Cardinal): Boolean;
var
  StartTick: Cardinal;
begin
  StartTick := GetTickCount;
  while True do
  begin
    if IsReady then
      Exit(True);
    if GetTickCount - StartTick > ATimeoutMs then
      Exit(False);
    Application.ProcessMessages;
    Sleep(200);
  end;
end;

class function TIdeCodeInsight.GotoDefinition(const AFileName: string; ALine, ACol: Integer;
  ATimeoutMs: Cardinal): TDefinitionResult;
var
  CIServices: IOTACodeInsightServices;
  CIManager: IOTACodeInsightManager;
  AsyncMgr290: IOTAAsyncCodeInsightManager290;
  AsyncMgr: IOTAAsyncCodeInsightManager;
  I: Integer;
begin
  Result := Default(TDefinitionResult);
  Result.IsValid := False;

  if not Supports(BorlandIDEServices, IOTACodeInsightServices, CIServices) then
  begin
    Result.ErrorMessage := 'IOTACodeInsightServices nicht verfuegbar';
    Exit;
  end;

  // Find a suitable async manager
  AsyncMgr := nil;
  AsyncMgr290 := nil;

  CIServices.GetCurrentCodeInsightManager(CIManager);
  if CIManager <> nil then
  begin
    Supports(CIManager, IOTAAsyncCodeInsightManager290, AsyncMgr290);
    if AsyncMgr290 = nil then
      Supports(CIManager, IOTAAsyncCodeInsightManager, AsyncMgr);
  end;

  // Fallback: search all managers
  if (AsyncMgr290 = nil) and (AsyncMgr = nil) then
  begin
    for I := 0 to CIServices.CodeInsightManagerCount - 1 do
    begin
      CIManager := CIServices.CodeInsightManager[I];
      if Supports(CIManager, IOTAAsyncCodeInsightManager290, AsyncMgr290) then
        Break;
      if Supports(CIManager, IOTAAsyncCodeInsightManager, AsyncMgr) then
        Break;
    end;
  end;

  if (AsyncMgr290 = nil) and (AsyncMgr = nil) then
  begin
    Result.ErrorMessage := 'Kein AsyncCodeInsightManager gefunden';
    Exit;
  end;

  // Event + message pump: the callback arrives via the message queue,
  // so we must not block the main thread with WaitFor.
  FEvent := TLightweightEvent.Create;
  try
    FResult := Default(TDefinitionResult);

    // Start async call
    if AsyncMgr290 <> nil then
      AsyncMgr290.AsyncGotoDefinitionEx(AFileName, ALine, ACol, GotoDefCallback)
    else
      AsyncMgr.AsyncGotoDefinition(AFileName, ALine, ACol, GotoDefCallbackLegacy);

    // Run the message pump until the callback arrives or timeout
    var StartTick := GetTickCount;
    while True do
    begin
      if FEvent.WaitFor(0) = wrSignaled then
      begin
        Result := FResult;
        Break;
      end;
      if GetTickCount - StartTick > ATimeoutMs then
      begin
        Result.Error := True;
        Result.ErrorMessage := 'Timeout nach ' + IntToStr(ATimeoutMs) + 'ms';
        Break;
      end;
      Application.ProcessMessages;
      Sleep(10);
    end;
  finally
    FreeAndNil(FEvent);
  end;
end;

end.
