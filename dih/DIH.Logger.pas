(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit DIH.Logger;

interface

uses
  System.SysUtils, System.Classes, Winapi.Windows, DIH.Types;

type
  TDIHLogLevel = (llDetail, llInfo, llSuccess, llWarning, llError);

  TDIHLogger = class
  private
    FLogFile: TStreamWriter;
    FConsoleHandle: THandle;
    FOriginalAttr: Word;
    FIndent: Integer;
    FVerboseTargets: TDIHVerboseTargets;
    FLastBuildOutput: TStringList;
    procedure WriteColored(const AText: string; AColor: Word);
    procedure WriteToFile(const AText: string);
    function GetTimestamp: string;
    function GetIndentStr: string;
  public
    constructor Create(const ALogFileName: string);
    destructor Destroy; override;
    procedure Info(const AMsg: string); overload;
    procedure Info(const AFmt: string; const AArgs: array of const); overload;
    procedure Success(const AMsg: string); overload;
    procedure Success(const AFmt: string; const AArgs: array of const); overload;
    procedure Error(const AMsg: string); overload;
    procedure Error(const AFmt: string; const AArgs: array of const); overload;
    procedure Warning(const AMsg: string); overload;
    procedure Warning(const AFmt: string; const AArgs: array of const); overload;
    procedure Detail(const AMsg: string); overload;
    procedure Detail(const AFmt: string; const AArgs: array of const); overload;
    /// <summary>
    ///  Logs compiler/build output according to verbose settings.
    ///  Only writes to console if vtOutput, only to log if vtLog.
    ///  Output is always buffered for error display.
    /// </summary>
    procedure CompilerOutput(const AMsg: string);
    /// <summary>
    ///  Flushes the buffered compiler output to console and log on error.
    /// </summary>
    procedure FlushBuildOutputToConsole;
    /// <summary>Clears the buffered compiler output before a new build.</summary>
    procedure ClearBuildOutput;
    procedure IncIndent;
    procedure DecIndent;
    procedure Separator;
    property VerboseTargets: TDIHVerboseTargets read FVerboseTargets write FVerboseTargets;
  end;

implementation

const
  COLOR_DEFAULT = FOREGROUND_RED or FOREGROUND_GREEN or FOREGROUND_BLUE;
  COLOR_SUCCESS = FOREGROUND_GREEN or FOREGROUND_INTENSITY;
  COLOR_ERROR   = FOREGROUND_RED or FOREGROUND_INTENSITY;
  COLOR_WARNING = FOREGROUND_RED or FOREGROUND_GREEN or FOREGROUND_INTENSITY;
  COLOR_INFO    = FOREGROUND_RED or FOREGROUND_GREEN or FOREGROUND_BLUE or FOREGROUND_INTENSITY;
  COLOR_DETAIL  = FOREGROUND_RED or FOREGROUND_GREEN or FOREGROUND_BLUE;

{ TDIHLogger }

constructor TDIHLogger.Create(const ALogFileName: string);
var
  Info: TConsoleScreenBufferInfo;
begin
  inherited Create;
  FIndent := 0;
  FVerboseTargets := [vtOutput, vtLog];
  FLastBuildOutput := TStringList.Create;
  FConsoleHandle := GetStdHandle(STD_OUTPUT_HANDLE);
  if GetConsoleScreenBufferInfo(FConsoleHandle, Info) then
    FOriginalAttr := Info.wAttributes
  else
    FOriginalAttr := COLOR_DEFAULT;

  FLogFile := TStreamWriter.Create(ALogFileName, False, TEncoding.UTF8);
  FLogFile.AutoFlush := True;
  WriteToFile('=== Delphi Install Helper Log ===');
  WriteToFile('Started: ' + DateTimeToStr(Now));
  WriteToFile('');
end;

destructor TDIHLogger.Destroy;
begin
  if Assigned(FLogFile) then
  begin
    WriteToFile('');
    WriteToFile('Finished: ' + DateTimeToStr(Now));
    FLogFile.Free;
  end;
  FLastBuildOutput.Free;
  SetConsoleTextAttribute(FConsoleHandle, FOriginalAttr);
  inherited;
end;

function TDIHLogger.GetTimestamp: string;
begin
  Result := FormatDateTime('hh:nn:ss', Now);
end;

function TDIHLogger.GetIndentStr: string;
begin
  Result := StringOfChar(' ', FIndent * 2);
end;

procedure TDIHLogger.WriteColored(const AText: string; AColor: Word);
begin
  SetConsoleTextAttribute(FConsoleHandle, AColor);
  Writeln(AText);
  SetConsoleTextAttribute(FConsoleHandle, FOriginalAttr);
end;

procedure TDIHLogger.WriteToFile(const AText: string);
begin
  if Assigned(FLogFile) then
    FLogFile.WriteLine(AText);
end;

procedure TDIHLogger.Info(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;
  WriteColored(Line, COLOR_INFO);
  WriteToFile('[' + GetTimestamp + '] [INFO] ' + Line);
end;

procedure TDIHLogger.Info(const AFmt: string; const AArgs: array of const);
begin
  Info(Format(AFmt, AArgs));
end;

procedure TDIHLogger.Success(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;
  WriteColored(Line, COLOR_SUCCESS);
  WriteToFile('[' + GetTimestamp + '] [ OK ] ' + Line);
end;

procedure TDIHLogger.Success(const AFmt: string; const AArgs: array of const);
begin
  Success(Format(AFmt, AArgs));
end;

procedure TDIHLogger.Error(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;
  WriteColored(Line, COLOR_ERROR);
  WriteToFile('[' + GetTimestamp + '] [FAIL] ' + Line);
end;

procedure TDIHLogger.Error(const AFmt: string; const AArgs: array of const);
begin
  Error(Format(AFmt, AArgs));
end;

procedure TDIHLogger.Warning(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;
  WriteColored(Line, COLOR_WARNING);
  WriteToFile('[' + GetTimestamp + '] [WARN] ' + Line);
end;

procedure TDIHLogger.Warning(const AFmt: string; const AArgs: array of const);
begin
  Warning(Format(AFmt, AArgs));
end;

procedure TDIHLogger.Detail(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;
  WriteColored(Line, COLOR_DETAIL);
  WriteToFile('[' + GetTimestamp + '] [    ] ' + Line);
end;

procedure TDIHLogger.Detail(const AFmt: string; const AArgs: array of const);
begin
  Detail(Format(AFmt, AArgs));
end;

procedure TDIHLogger.CompilerOutput(const AMsg: string);
var
  Line: string;
begin
  Line := GetIndentStr + AMsg;

  // Always buffer for potential error display
  FLastBuildOutput.Add(Line);

  // Write to console only if vtOutput
  if vtOutput in FVerboseTargets then
    WriteColored(Line, COLOR_DETAIL);

  // Write to log only if vtLog
  if vtLog in FVerboseTargets then
    WriteToFile('[' + GetTimestamp + '] [    ] ' + Line);
end;

procedure TDIHLogger.FlushBuildOutputToConsole;
var
  Line: string;
begin
  if FLastBuildOutput.Count = 0 then
    Exit;

  // Flush to console if it was not already shown
  if not (vtOutput in FVerboseTargets) then
  begin
    WriteColored(GetIndentStr + '--- Build output ---', COLOR_WARNING);
    for Line in FLastBuildOutput do
      WriteColored(Line, COLOR_DETAIL);
    WriteColored(GetIndentStr + '--- End build output ---', COLOR_WARNING);
  end;

  // Flush to log if it was not already written
  if not (vtLog in FVerboseTargets) then
  begin
    WriteToFile('[' + GetTimestamp + '] [    ] --- Build output ---');
    for Line in FLastBuildOutput do
      WriteToFile('[' + GetTimestamp + '] [    ] ' + Line);
    WriteToFile('[' + GetTimestamp + '] [    ] --- End build output ---');
  end;
end;

procedure TDIHLogger.ClearBuildOutput;
begin
  FLastBuildOutput.Clear;
end;

procedure TDIHLogger.IncIndent;
begin
  Inc(FIndent);
end;

procedure TDIHLogger.DecIndent;
begin
  if FIndent > 0 then
    Dec(FIndent);
end;

procedure TDIHLogger.Separator;
const
  Sep = '----------------------------------------------------------------------';
begin
  WriteColored(Sep, COLOR_DETAIL);
  WriteToFile(Sep);
end;

end.
