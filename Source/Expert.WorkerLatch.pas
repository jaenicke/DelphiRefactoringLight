(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WorkerLatch;

// One shutdown latch for the plugin's background threads (idea and shape
// from Ian Branch's audit, issue #10). A worker that is still running when
// the package is unloaded (Component > Install Packages, IDE shutdown)
// executes unmapped code - and so does a TThread.Queue closure it left
// behind. So:
//   * StartWorker refuses to start anything once the shutdown has begun,
//   * every worker is counted while it runs,
//   * ShutdownWorkersAndWait (first thing in the finalization, before any
//     owner object is freed) waits for the count to reach zero and keeps
//     the synchronize queue moving, so queued closures run while the code
//     is still mapped.
// The live checker (Expert.AutoImport) and the blame reader
// (Expert.VcsBlame) keep their own latches with the same rules.

interface

uses
  System.SysUtils;

/// <summary>Starts AProc on a new anonymous thread and counts it as a
///  worker until it returns. False (nothing started) once the shutdown has
///  begun - callers must then undo any "in flight" state they set.</summary>
function StartWorker(const AProc: TProc): Boolean;

/// <summary>Low-level pair for threads created elsewhere: WorkerEnter is
///  False after the shutdown began (then do not start), otherwise every
///  WorkerEnter needs exactly one WorkerLeave.</summary>
function WorkerEnter: Boolean;
procedure WorkerLeave;

/// <summary>True once the shutdown has begun - long loops in workers
///  should check it and give up.</summary>
function WorkersShuttingDown: Boolean;

/// <summary>Number of running workers.</summary>
function ActiveWorkerCount: Integer;

/// <summary>Latches (no new workers) and waits up to ATimeoutMs for the
///  running ones, pumping CheckSynchronize on the main thread. True when
///  every worker has left. Call it before freeing anything a worker may
///  still touch.</summary>
function ShutdownWorkersAndWait(ATimeoutMs: Cardinal = 35000): Boolean;

/// <summary>Test support only: re-opens the latch.</summary>
procedure ResetWorkerLatch;

implementation

uses
  System.Classes, Winapi.Windows;

var
  GCount: Integer = 0;
  GShutdown: Integer = 0;

function WorkerEnter: Boolean;
begin
  if AtomicCmpExchange(GShutdown, 0, 0) <> 0 then Exit(False);
  AtomicIncrement(GCount);
  // The shutdown may have begun between the check and the increment -
  // back out then, the waiter may already have seen zero.
  if AtomicCmpExchange(GShutdown, 0, 0) <> 0 then
  begin
    AtomicDecrement(GCount);
    Exit(False);
  end;
  Result := True;
end;

procedure WorkerLeave;
begin
  AtomicDecrement(GCount);
end;

function WorkersShuttingDown: Boolean;
begin
  Result := AtomicCmpExchange(GShutdown, 0, 0) <> 0;
end;

function ActiveWorkerCount: Integer;
begin
  Result := AtomicCmpExchange(GCount, 0, 0);
end;

function StartWorker(const AProc: TProc): Boolean;
var
  T: TThread;
begin
  if not WorkerEnter then Exit(False);
  try
    T := TThread.CreateAnonymousThread(
      procedure
      begin
        try
          AProc();
        finally
          WorkerLeave;
        end;
      end);
    T.Start;
  except
    WorkerLeave;
    raise;
  end;
  Result := True;
end;

function ShutdownWorkersAndWait(ATimeoutMs: Cardinal): Boolean;
var
  Start: UInt64;
  OnMain: Boolean;
begin
  AtomicExchange(GShutdown, 1);
  OnMain := GetCurrentThreadId = MainThreadID;
  Start := GetTickCount64;
  while (ActiveWorkerCount > 0) and (GetTickCount64 - Start < ATimeoutMs) do
    if OnMain then
      CheckSynchronize(10)
    else
      Sleep(10);
  Result := ActiveWorkerCount = 0;
  if Result then
    // The last worker has decremented the counter but may still be on its
    // way out of the closure - a moment for the epilogue.
    Sleep(50);
  if OnMain then
    for var I := 1 to 100 do
      if not CheckSynchronize(0) then Break;   // drain queued closures
end;

procedure ResetWorkerLatch;
begin
  AtomicExchange(GShutdown, 0);
end;

end.
