(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ResourceMonitor;

// Process resource readings for the status window and a trend log.
//
// Written after a tester report of "Zu wenig Arbeitsspeicher" and
// "EPNGOutMemory in vclimg370.bpl", followed by a frozen IDE that
// eventually terminated. EPNGOutMemory is raised in exactly one situation:
// CreateCompatibleDC or CreateDIBSection returned 0. That happens when the
// process is out of GDI objects (Windows allows 10,000 per process) OR
// when a 32-bit process cannot map the bitmap's memory any more because
// its address space is exhausted or fragmented. Both causes look the same
// to the user, and both build up over hours - so the only useful next
// report is one with NUMBERS. This unit provides them:
//
//  * Sample: GDI / USER objects (current + peak), private bytes, and for a
//    32-bit process the largest FREE address-space block - the value that
//    decides whether the next big allocation still succeeds.
//  * GdiGuard: a one-line scope guard for our own timer ticks and paint
//    handlers. It books the GDI objects a call LEFT BEHIND per subsystem,
//    so a leak in this plugin shows up as a steadily growing balance with
//    a name on it (and a flat balance clears us).
//  * A background sampler that appends a line to
//    %TEMP%\RefactoringLight-resources.log every 10 minutes and whenever a
//    threshold is crossed.

interface

type
  TResourceSample = record
    GdiObjects, GdiPeak: Integer;
    UserObjects, UserPeak: Integer;
    PrivateMB: Integer;
    /// <summary>Largest free virtual address block in MB; -1 when not
    ///  measured (64-bit process - the address space is not the limit).</summary>
    LargestFreeMB: Integer;
    FreeTotalMB: Integer;
    Is32Bit: Boolean;
  end;

  TGdiScope = (gsLiveTick, gsBlameTick, gsBlamePaint, gsMarkerPaint,
    gsStatusTick, gsStructure, gsMenuUpdate);

/// <summary>Current readings. The address-space walk is only done for a
///  32-bit process and costs a few ms - do not call it per paint.</summary>
function SampleResources(AWithAddressSpace: Boolean = True): TResourceSample;

/// <summary>Scope guard: keep the result in a local variable for the
///  duration of the routine ("var G := GdiGuard(gsLiveTick);"). When it
///  goes out of scope - on every exit path - the GDI objects created and
///  not released in between are added to the scope's balance.
///  ASampleEvery samples only every n-th call (paint handlers run per
///  line; one reading per call pair would be measurable there).</summary>
function GdiGuard(AScope: TGdiScope; ASampleEvery: Integer = 1): IInterface;

/// <summary>"live +0, blame tick +0, ..." - net GDI objects per scope
///  since the plugin was loaded.</summary>
function GdiBalanceText: string;

/// <summary>One human-readable line of a sample (also used in the log).</summary>
function ResourceSampleText(const S: TResourceSample): string;

procedure StartResourceMonitor;
procedure StopResourceMonitor;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.SyncObjs,
  Winapi.Windows, Winapi.PsAPI, Vcl.ExtCtrls;

const
  ScopeNames: array[TGdiScope] of string = ('live checker', 'blame tick',
    'blame paint', 'marker paint', 'status window', 'structure view',
    'menu update');
  // Windows' default per-process quota is 10,000 for both object kinds.
  WarnHandles = 8000;
  // A 32-bit process whose largest free block drops below this can no
  // longer map a large bitmap or grow a big array.
  WarnLargestFreeMB = 64;
  LogEveryTicks = 20;          // x 30 s = every 10 minutes
  MaxLogBytes = 512 * 1024;

var
  GBalance: array[TGdiScope] of Integer;
  GCalls: array[TGdiScope] of Integer;
  GMonitor: TTimer = nil;
  GMonitorTicks: Integer = 0;
  GWarned: set of (wGdi, wUser, wAddress);

function PrivateBytesMB: Integer;
var
  C: PROCESS_MEMORY_COUNTERS_EX;
begin
  FillChar(C, SizeOf(C), 0);
  C.cb := SizeOf(C);
  if GetProcessMemoryInfo(GetCurrentProcess, @C, SizeOf(C)) then
    Result := Integer(C.PrivateUsage div (1024 * 1024))
  else
    Result := -1;
end;

procedure WalkAddressSpace(out ALargestMB, AFreeMB: Integer);
var
  SI: TSystemInfo;
  P: NativeUInt;
  MBI: TMemoryBasicInformation;
  Largest, Total: NativeUInt;
begin
  GetSystemInfo(SI);
  P := NativeUInt(SI.lpMinimumApplicationAddress);
  Largest := 0;
  Total := 0;
  while P < NativeUInt(SI.lpMaximumApplicationAddress) do
  begin
    if VirtualQuery(Pointer(P), MBI, SizeOf(MBI)) = 0 then Break;
    if MBI.State = MEM_FREE then
    begin
      Inc(Total, MBI.RegionSize);
      if MBI.RegionSize > Largest then Largest := MBI.RegionSize;
    end;
    if MBI.RegionSize = 0 then Break;
    P := NativeUInt(MBI.BaseAddress) + MBI.RegionSize;
  end;
  ALargestMB := Integer(Largest div (1024 * 1024));
  AFreeMB := Integer(Total div (1024 * 1024));
end;

function SampleResources(AWithAddressSpace: Boolean): TResourceSample;
const
  GR_GDIOBJECTS_PEAK = 2;
  GR_USEROBJECTS_PEAK = 4;
var
  H: THandle;
begin
  H := GetCurrentProcess;
  Result.GdiObjects := GetGuiResources(H, GR_GDIOBJECTS);
  Result.GdiPeak := GetGuiResources(H, GR_GDIOBJECTS_PEAK);
  Result.UserObjects := GetGuiResources(H, GR_USEROBJECTS);
  Result.UserPeak := GetGuiResources(H, GR_USEROBJECTS_PEAK);
  Result.PrivateMB := PrivateBytesMB;
  Result.Is32Bit := SizeOf(Pointer) = 4;
  Result.LargestFreeMB := -1;
  Result.FreeTotalMB := -1;
  if Result.Is32Bit and AWithAddressSpace then
    WalkAddressSpace(Result.LargestFreeMB, Result.FreeTotalMB);
end;

function ResourceSampleText(const S: TResourceSample): string;
begin
  Result := Format('GDI %d (peak %d), USER %d (peak %d), private %d MB',
    [S.GdiObjects, S.GdiPeak, S.UserObjects, S.UserPeak, S.PrivateMB]);
  if S.LargestFreeMB >= 0 then
    Result := Result + Format(', address space free %d MB, largest block %d MB',
      [S.FreeTotalMB, S.LargestFreeMB])
  else if not S.Is32Bit then
    Result := Result + ' (64-bit)';
end;

// ---------------------------------------------------------------------------
//  Scope guard
// ---------------------------------------------------------------------------

type
  TGdiGuard = class(TInterfacedObject)
  private
    FScope: TGdiScope;
    FStart: Integer;
  public
    constructor Create(AScope: TGdiScope);
    destructor Destroy; override;
  end;

constructor TGdiGuard.Create(AScope: TGdiScope);
begin
  inherited Create;
  FScope := AScope;
  FStart := GetGuiResources(GetCurrentProcess, GR_GDIOBJECTS);
end;

destructor TGdiGuard.Destroy;
begin
  // Paint handlers and ticks run on the main thread only, but the
  // interlocked add costs nothing and keeps the counters honest.
  TInterlocked.Add(GBalance[FScope],
    Integer(GetGuiResources(GetCurrentProcess, GR_GDIOBJECTS)) - FStart);
  inherited;
end;

function GdiGuard(AScope: TGdiScope; ASampleEvery: Integer): IInterface;
begin
  Result := nil;
  if (ASampleEvery > 1)
    and (TInterlocked.Increment(GCalls[AScope]) mod ASampleEvery <> 0) then
    Exit;
  Result := TGdiGuard.Create(AScope);
end;

function GdiBalanceText: string;
var
  S: TGdiScope;
begin
  Result := '';
  for S := Low(TGdiScope) to High(TGdiScope) do
  begin
    if Result <> '' then Result := Result + ', ';
    // Delphi's Format has no '+' flag - sign it by hand.
    if GBalance[S] > 0 then
      Result := Result + Format('%s +%d', [ScopeNames[S], GBalance[S]])
    else
      Result := Result + Format('%s %d', [ScopeNames[S], GBalance[S]]);
  end;
end;

// ---------------------------------------------------------------------------
//  Background sampler + trend log
// ---------------------------------------------------------------------------

type
  TMonitorSink = class
    procedure Tick(Sender: TObject);
  end;

var
  GSink: TMonitorSink = nil;

function LogPath: string;
begin
  Result := TPath.Combine(TPath.GetTempPath, 'RefactoringLight-resources.log');
end;

procedure AppendLog(const ALine: string);
begin
  try
    TFile.AppendAllText(LogPath,
      FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + '  ' + ALine + sLineBreak,
      TEncoding.UTF8);
  except
    // a diagnostics log must never disturb the IDE
  end;
end;

procedure TMonitorSink.Tick(Sender: TObject);
var
  S: TResourceSample;
  Why: string;
begin
  try
    Inc(GMonitorTicks);
    S := SampleResources(True);
    Why := '';
    if (S.GdiObjects >= WarnHandles) and not (wGdi in GWarned) then
    begin
      Include(GWarned, wGdi);
      Why := 'GDI objects near the 10,000 limit';
    end;
    if (S.UserObjects >= WarnHandles) and not (wUser in GWarned) then
    begin
      Include(GWarned, wUser);
      Why := 'USER objects near the 10,000 limit';
    end;
    if (S.LargestFreeMB >= 0) and (S.LargestFreeMB < WarnLargestFreeMB)
      and not (wAddress in GWarned) then
    begin
      Include(GWarned, wAddress);
      Why := 'address space fragmented - large allocations will fail';
    end;

    if Why <> '' then
      AppendLog('WARNING ' + Why + ': ' + ResourceSampleText(S) +
        ' | plugin GDI balance: ' + GdiBalanceText)
    else if GMonitorTicks mod LogEveryTicks = 0 then
      AppendLog(ResourceSampleText(S) + ' | plugin GDI balance: ' +
        GdiBalanceText);
  except
  end;
end;

procedure StartResourceMonitor;
begin
  if GMonitor <> nil then Exit;
  try
    // Keep the log bounded: a new session starts a fresh file once the
    // old one has grown large.
    if TFile.Exists(LogPath) and (TFile.GetSize(LogPath) > MaxLogBytes) then
      TFile.Delete(LogPath);
  except
  end;
  AppendLog('--- session start: ' + ResourceSampleText(SampleResources(True)));
  GSink := TMonitorSink.Create;
  GMonitor := TTimer.Create(nil);
  GMonitor.Interval := 30000;
  GMonitor.OnTimer := GSink.Tick;
  GMonitor.Enabled := True;
end;

procedure StopResourceMonitor;
begin
  if GMonitor = nil then Exit;
  GMonitor.Enabled := False;
  FreeAndNil(GMonitor);
  FreeAndNil(GSink);
end;

initialization

finalization
  StopResourceMonitor;

end.
