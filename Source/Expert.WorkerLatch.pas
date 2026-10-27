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

/// <summary>Takes the CALLING thread into the join list, for a worker this
///  latch did not start. The live checker (Expert.AutoImport) and the blame
///  reader (Expert.VcsBlame) keep their own counters because they have
///  their own place in the shutdown order - and a COUNTER IS NOT A JOIN: it
///  drops in the worker closure's own `finally`, so it reaches zero while
///  the closure epilogue (which releases everything the closure captured,
///  with code that lives in THIS package) and the thread's own exit path
///  are still to run. That is audit #38's L1g for the two subsystems that
///  were not converted with it.
///  Safe as the closure's FIRST statement, and only there: both raise their
///  counter BEFORE Start, so no waiter can pass before this line ran.</summary>
procedure NoteCurrentWorkerThread;

/// <summary>Waits until every noted worker thread has really exited,
///  pumping the synchronize queue while on the main thread. False when the
///  timeout was reached - then plugin code may still be running, which the
///  caller reports through NoteShutdownProblem rather than ignoring.
///  ShutdownWorkersAndWait does this for its own workers; a subsystem that
///  waits for a counter of its own calls this after that wait.</summary>
function JoinNotedWorkers(ATimeoutMs: Cardinal = 5000): Boolean;

/// <summary>What the UNLOAD could not finish, collected in memory while the
///  package is being unloaded and handed to Expert.RestartHint at the end
///  of the finalization, which writes it into this process's marker file.
///  The next load reads it from there - which is the only channel there is,
///  because the instance that knows is the one that is going away.
///  This unit holds it because it is the one place every subsystem of all
///  three projects can reach; the marker file itself stays owned by
///  Expert.RestartHint.</summary>
procedure NoteShutdownProblem(const AText: string);
function ShutdownProblems: TArray<string>;

/// <summary>Latches (no new workers) and waits up to ATimeoutMs for the
///  running ones, pumping CheckSynchronize on the main thread. True when
///  every worker has left. Call it before freeing anything a worker may
///  still touch.</summary>
function ShutdownWorkersAndWait(ATimeoutMs: Cardinal = 35000): Boolean;

/// <summary>TThread.Queue / ForceQueue, plus the second rule a queued
///  closure has to keep: it must not let an exception ESCAPE.
///
///  A closure handed to TThread.Queue runs inside
///  System.Classes.CheckSynchronize, and when it raises, CheckSynchronize
///  itself calls ApplicationHandleException (System.Classes.pas:16374, the
///  `SyncProc.Queued` branch) - which in the VCL ends in
///  TApplication.MessageBox, a MODAL Win32 box. The IDE pumps
///  CheckSynchronize from inside its own transactions (a package install,
///  a project load), so that box stops a transaction mid-flight and the
///  IDE hangs. It is the same wedge a modal window in the closure itself
///  produces - see Expert.RestartHint, which cost a user an IDE session on
///  2026-10-26 - and it bypasses Application.ModalLevel entirely, because
///  TApplication.MessageBox never calls ModalStarted.
///
///  So the exception is swallowed here. Swallowing is only defensible with
///  a channel: it is counted and the last one is kept
///  (QueuedErrorCount / LastQueuedError), and the status window reports
///  them.</summary>
procedure QueueSafe(const AProc: TProc);
procedure ForceQueueSafe(const AProc: TProc);

/// <summary>How many queued closures raised, and what the last one said -
///  the channel that makes swallowing defensible.</summary>
function QueuedErrorCount: Integer;
function LastQueuedError: string;

/// <summary>Test support only: re-opens the latch.</summary>
procedure ResetWorkerLatch;

implementation

uses
  System.Classes, Winapi.Windows;

var
  GCount: Integer = 0;
  GShutdown: Integer = 0;
  // Audit #38, L1g: the COUNTER drops in the closure's finally, i.e. before
  // the closure epilogue (which releases everything AProc captured) and the
  // thread's own exit path have run. A shutdown that only watches the count
  // therefore lets the BPL unload while plugin code is still executing on
  // that thread; the old Sleep(50) was a timing allowance, not a join. So
  // the thread HANDLES are kept (duplicated - TThread closes its own) and
  // waited for as well.
  GHandleLock: TObject = nil;
  GHandles: TArray<THandle>;
  // What the unload could not finish. GHandleLock guards it too - one lock
  // is enough for both and nothing here is on a hot path.
  GProblems: TArray<string>;
  // What the swallowed exceptions of QueueSafe amount to. GHandleLock
  // guards the string as well - one lock is enough for both, and nothing
  // here is on a hot path.
  GQueueErrors: Integer = 0;
  GLastQueueError: string = '';

// ASource is DUPLICATED, never kept: TThread closes its own handle in its
// destructor, and an anonymous thread frees itself the moment it ends - so
// waiting on TThread.Handle would be a race against that free. A PSEUDO
// handle (GetCurrentThread) duplicates into a real one, which is what lets
// a worker note itself.
procedure AddNotedHandle(ASource: THandle);
var
  Dup: THandle;
begin
  if (GHandleLock = nil) or (ASource = 0) then Exit;
  Dup := 0;
  if not DuplicateHandle(GetCurrentProcess, ASource,
       GetCurrentProcess, @Dup, SYNCHRONIZE, False, 0) or (Dup = 0) then
    Exit;
  TMonitor.Enter(GHandleLock);
  try
    // prune the ones that have exited, so a long IDE session does not
    // collect a handle per worker
    var Keep: TArray<THandle> := nil;
    for var Old in GHandles do
      if WaitForSingleObject(Old, 0) = WAIT_TIMEOUT then
        Keep := Keep + [Old]
      else
        CloseHandle(Old);
    GHandles := Keep + [Dup];
  finally
    TMonitor.Exit(GHandleLock);
  end;
end;

procedure NoteWorkerThread(AThread: TThread);
begin
  if AThread <> nil then
    AddNotedHandle(AThread.Handle);
end;

procedure NoteCurrentWorkerThread;
begin
  AddNotedHandle(GetCurrentThread);
end;

procedure NoteShutdownProblem(const AText: string);
begin
  if (GHandleLock = nil) or (AText = '') then Exit;
  TMonitor.Enter(GHandleLock);
  try
    // the same problem twice says nothing twice
    for var S in GProblems do
      if S = AText then Exit;
    GProblems := GProblems + [AText];
  finally
    TMonitor.Exit(GHandleLock);
  end;
end;

function ShutdownProblems: TArray<string>;
begin
  Result := nil;
  if GHandleLock = nil then Exit;
  TMonitor.Enter(GHandleLock);
  try
    Result := Copy(GProblems);
  finally
    TMonitor.Exit(GHandleLock);
  end;
end;

/// <summary>Waits for the threads still running, up to the deadline.
///  Handles that are signalled (or that we stop waiting for) are closed.
///  </summary>
function JoinWorkerThreads(ADeadline: UInt64; AOnMain: Boolean): Boolean;
var
  Pending: TArray<THandle>;
begin
  Result := True;
  if GHandleLock = nil then Exit;
  TMonitor.Enter(GHandleLock);
  try
    Pending := GHandles;
    GHandles := nil;
  finally
    TMonitor.Exit(GHandleLock);
  end;
  for var H in Pending do
  begin
    while WaitForSingleObject(H, 10) = WAIT_TIMEOUT do
    begin
      if GetTickCount64 >= ADeadline then
      begin
        Result := False;
        Break;
      end;
      // a worker may be waiting on a queued closure - keep the queue moving
      if AOnMain then CheckSynchronize(0);
    end;
    CloseHandle(H);
  end;
end;

function JoinNotedWorkers(ATimeoutMs: Cardinal): Boolean;
begin
  Result := JoinWorkerThreads(GetTickCount64 + ATimeoutMs,
    GetCurrentThreadId = MainThreadID);
end;

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
    // BEFORE Start, so a shutdown that sees the count already raised can
    // never miss this thread's handle.
    NoteWorkerThread(T);
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
  // Every worker that was counted has registered its thread handle before
  // it ran AProc, so once the count is zero the handle list is complete -
  // and waiting on it is what makes "no plugin code is running any more"
  // true, instead of hoping 50 ms was enough (audit #38, L1g).
  if not JoinWorkerThreads(Start + ATimeoutMs, OnMain) then
    Result := False;
  if OnMain then
    for var I := 1 to 100 do
      if not CheckSynchronize(0) then Break;   // drain queued closures
end;

// The closure every QueueSafe / ForceQueueSafe call really hands to the
// RTL: whatever AProc does, nothing leaves it.
// TThreadProcedure, not TProc: the two are structurally identical and
// still distinct types, so TThread.Queue takes no TProc (E2250).
function Guarded(const AProc: TProc): TThreadProcedure;
begin
  Result :=
    procedure
    begin
      if not Assigned(AProc) then Exit;
      try
        AProc();
      except
        on E: Exception do
        begin
          AtomicIncrement(GQueueErrors);
          if GHandleLock <> nil then
          begin
            TMonitor.Enter(GHandleLock);
            try
              GLastQueueError := E.ClassName + ': ' + E.Message;
            finally
              TMonitor.Exit(GHandleLock);
            end;
          end;
        end;
      end;
    end;
end;

procedure QueueSafe(const AProc: TProc);
begin
  TThread.Queue(nil, Guarded(AProc));
end;

procedure ForceQueueSafe(const AProc: TProc);
begin
  TThread.ForceQueue(nil, Guarded(AProc));
end;

function QueuedErrorCount: Integer;
begin
  Result := GQueueErrors;
end;

function LastQueuedError: string;
begin
  Result := '';
  if GHandleLock = nil then Exit;
  TMonitor.Enter(GHandleLock);
  try
    Result := GLastQueueError;
  finally
    TMonitor.Exit(GHandleLock);
  end;
end;

procedure ResetWorkerLatch;
begin
  AtomicExchange(GShutdown, 0);
  AtomicExchange(GQueueErrors, 0);
  if GHandleLock <> nil then
  begin
    TMonitor.Enter(GHandleLock);
    try
      GLastQueueError := '';
      // A reset means "as if nothing had happened" - a problem recorded by
      // a shutdown that was taken back is no longer a problem.
      GProblems := nil;
    finally
      TMonitor.Exit(GHandleLock);
    end;
  end;
end;

initialization
  GHandleLock := TObject.Create;

finalization
  // Nothing may wait on the lock any more at this point: the shutdown has
  // run (it is the first thing in the package's finalization).
  for var H in GHandles do CloseHandle(H);
  GHandles := nil;
  GProblems := nil;
  FreeAndNil(GHandleLock);

end.
