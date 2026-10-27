(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.RestartHint;

interface

uses
  System.Types, System.SysUtils;

type
  /// <summary>What the PREVIOUS instance of this package left behind in THIS
  ///  process - read at load from the marker file, which is the only channel
  ///  there is: the instance that knows is the one that is going away.</summary>
  TReloadFacts = record
    /// <summary>The package was loaded a second time in this process, i.e.
    ///  somebody pressed "Installieren". A fresh IDE start is False.</summary>
    Reloaded: Boolean;
    /// <summary>What the previous unload could not finish
    ///  (Expert.WorkerLatch.ShutdownProblems).</summary>
    Problems: TArray<string>;
    DockWasOpen: Boolean;
    DockWasDocked: Boolean;
    DockBounds: TRect;
    /// <summary>The DeskSection of the DOCK HOST that held the window, e.g.
    ///  "DockSite5" - measured from the live form before the unload
    ///  (HostDockSite -> its form -> TDesktopForm.DeskSection). That host
    ///  belongs to the IDE and survives our unload, so this string is what
    ///  makes a docked restore possible at all: the desktop file records the
    ///  membership with the SITE ("[DockSite5] TabDockClients=MessageView,
    ///  RefactoringLightStatus"), never with our own window.</summary>
    DockHostSection: string;
    /// <summary>Files the MCP bridge had opened without an editor tab and
    ///  that were still loaded at the unload.</summary>
    Headless: TArray<string>;
  end;

  /// <summary>The post-install tick, which is what this unit really is.
  ///
  ///  It used to show one unconditional box demanding an IDE restart, with a
  ///  reason nobody had ever measured ("the IDE still holds references to the
  ///  previous package version"). The audit of 2026-10-27 found that
  ///  fourteen of the fifteen registrations Register makes are taken back by
  ///  the finalization and the fifteenth is the package wizard, which the IDE
  ///  owns and no API returns - so the box asserted a mechanism that is not
  ///  there, on every install, several times a day.
  ///  What a reload really loses is INITIALISATION: state that is set up from
  ///  IDE events which cannot fire again for a project that is already open.
  ///  So the tick now CATCHES THAT UP first (through SettleHook) and shows a
  ///  box only for what it could not restore.</summary>
  TRestartHint = class
  private
    class var FFacts: TReloadFacts;
    class var FPinned: Boolean;
    class function ProcessStamp: string; static;
    class procedure PruneDeadMarkers; static;
    class procedure Arm; static;
    class function ReadMarkerLines: TArray<string>; static;
  public
    /// <summary>Called as the last statement of Register. FILE I/O AND STATE
    ///  ONLY: it reads what the previous instance wrote, rewrites the marker
    ///  and arms the timer. It may not touch ToolsAPI and may not open a
    ///  window - Register runs inside the IDE's install transaction.</summary>
    class procedure Check; static;
    /// <summary>The file this process hands over through, named after its
    ///  pid - also the one thing to look at when a report says the hint said
    ///  the wrong thing. Public because a test must be able to start from a
    ///  known state, and because a path a diagnosis needs should not have to
    ///  be guessed from the source.</summary>
    class function GetMarkerFilePath: string; static;
    /// <summary>What Check read. Valid from Check until the unload.</summary>
    class function Facts: TReloadFacts; static;
    /// <summary>The marker half of Check, and a test seam: it reads what the
    ///  previous instance of this package wrote into THIS process, consumes
    ///  the marker and answers the facts - without arming anything. Check is
    ///  this plus Arm, and nothing but Check may arm, because a tick that
    ///  runs during a test run has no IDE to catch up and no screen to put a
    ///  box on.</summary>
    class function ReadFacts: TReloadFacts; static;
    /// <summary>Writes this instance's parting knowledge into the marker, for
    ///  the instance that loads next. Called ONCE, at the end of the
    ///  package's finalization, when every shutdown has had its say.</summary>
    class procedure RecordUnload(const AProblems, AHeadless: TArray<string>;
      ADockOpen, ADockDocked: Boolean; const ADockBounds: TRect;
      const ADockHostSection: string); static;
    /// <summary>The MCP server had to pin this module: it can never be
    ///  unmapped now, so there is no next load and no marker to hand over to -
    ///  THIS instance has to say it. Measured (scratchpad pinprobe): after a
    ///  pin the .bpl cannot be opened for writing at all (sharing violation),
    ///  so the next build fails at the link step, and loading the pinned image
    ///  again raises an access violation while no initialization section
    ///  runs.</summary>
    class procedure NotePinned; static;
    /// <summary>Returns the text for the box, '' when there is nothing to
    ///  report - and it is also what performs the catch-up, because only the
    ///  IDE-only units know how. Assigned by Expert.Registration before
    ///  Check and cleared by its finalization. Runs on the MAIN thread from a
    ///  plain WM_TIMER tick, so it may use ToolsAPI - but it must not open a
    ///  window: the box is the tick's job.</summary>
    class var SettleHook: TFunc<string>;
  end;

type
  /// <summary>What the caller should DO about the status window after a
  ///  reload. The docked case is no longer a dead end: the dock host belongs
  ///  to the IDE and survives our unload, so it can be found again by its
  ///  DeskSection and asked to take the window back.</summary>
  TReloadWindowAction = (
    /// <summary>Nothing - it was not open, or this is a fresh IDE start.</summary>
    rwaNothing,
    /// <summary>It was free-standing: create it and put it at its bounds.</summary>
    rwaFloating,
    /// <summary>It was docked and we know the host section: create it and ask
    ///  that host to show it again.</summary>
    rwaRedock,
    /// <summary>It was docked but the host is unknown (a marker written before
    ///  2.8.0, or a window whose host we could not read) - then there is
    ///  nothing to put it back into and the box has to say so.</summary>
    rwaDockedUnknownHost);

/// <summary>What to do with the status window after a reload. PURE - the
///  caller performs it, and reports back through ReloadReportText whether it
///  worked.</summary>
function ReloadWindowAction(const AFacts: TReloadFacts): TReloadWindowAction;

/// <summary>WHAT A RELOAD HAS TO TELL THE USER. PURE, and deliberately apart
///  from the doing: the catch-up itself needs ToolsAPI and a live IDE, this
///  decision does not.
///  AWindowRestored is what the caller ACHIEVED - so the window is mentioned
///  only when it could not be put back, which is the whole point of this
///  round: a sentence that sends the user to the desktop menu must not appear
///  when the window is already sitting where it was.
///  '' means there is nothing to report, which is the normal outcome.</summary>
function ReloadReportText(const AFacts: TReloadFacts;
  AWindowRestored: Boolean): string;

/// <summary>Which DOCK HOST section of a desktop file lists this window, e.g.
///  "DockSite5" for "RefactoringLightStatus". PURE text work on the lines of a
///  .dsk, so it is testable - and it is the FALLBACK for the hand-over: the
///  file reflects the last SAVED desktop, which may be older than what the
///  user last did, so a live answer always wins over this one.
///  MEASURED shape (the user's own project desktop):
///    [DockSite5]
///    HostDockSite=DockBottomCenterPanel
///    TabDockClients=MessageView,RefactoringLightStatus
///  The membership is recorded with the SITE, which is why our own section
///  ([RefactoringLightStatus], Docked=1) cannot answer where it was.</summary>
function FindDockHostSection(const ALines: TArray<string>;
  const AIdentifier: string): string;

implementation

uses
  System.Classes, System.IOUtils, System.StrUtils,
  Winapi.Windows, Vcl.Forms, Vcl.ExtCtrls, Vcl.Dialogs, System.UITypes;

const
  /// <summary>Ticks to sit out before the catch-up runs and the box may
  ///  open. THE DELAY IS ALSO A GUARD: on a fresh IDE start the project-open
  ///  notification has long arrived after five seconds, so the catch-up finds
  ///  nothing to do and stays silent; on a reload it never arrives at all.</summary>
  DeferTicks = 10;
  DeferIntervalMs = 500;

  /// <summary>Marker keys. A line the reader does not know is ignored, so an
  ///  older marker left by a previous version can never break a load.</summary>
  KeyProblem = 'problem=';
  KeyHeadless = 'headless=';
  KeyDock = 'dock=';

type
  /// <summary>Opens the hint from a plain WM_TIMER tick.
  ///
  ///  HARD RULE, and this unit is what broke it: TThread.Queue/ForceQueue
  ///  procs run inside System.Classes.CheckSynchronize, and the IDE pumps
  ///  CheckSynchronize from inside its OWN transactions - a package install
  ///  among them, which is exactly when Register calls TRestartHint.Check.
  ///  A MODAL window opened there sits in the middle of that transaction,
  ///  so the install never finishes and the IDE reports "unexpected wait
  ///  event 4" over and over (reported 2026-10-26 with ShowRestartHint in
  ///  the stack, under ClearCaches -> SetCompState -> TParseThread.Add).
  ///  CheckSynchronize runs the synchronization QUEUE and dispatches no
  ///  window messages, so a WM_TIMER cannot be delivered from it: the
  ///  timer is the channel that is safe by construction. Same remedy as
  ///  TUnitRenameWatcher.DoDeferTick, for the same reason.</summary>
  TRestartHintDefer = class(TComponent)
  private
    FTimer: TTimer;
    FTicks: Integer;
    procedure DoTick(Sender: TObject);
  public
    constructor Create; reintroduce;
    destructor Destroy; override;
    procedure Arm;
  end;

var
  GDefer: TRestartHintDefer;

{ TRestartHintDefer }

constructor TRestartHintDefer.Create;
begin
  inherited Create(nil);
  FTimer := TTimer.Create(Self);
  FTimer.Enabled := False;
  FTimer.Interval := DeferIntervalMs;
  FTimer.OnTimer := DoTick;
end;

destructor TRestartHintDefer.Destroy;
begin
  if FTimer <> nil then
    FTimer.Enabled := False;
  inherited;
end;

procedure TRestartHintDefer.Arm;
begin
  FTicks := 0;
  FTimer.Enabled := True;
end;

procedure TRestartHintDefer.DoTick(Sender: TObject);
var
  Text: string;
begin
  if Application.Terminated then
  begin
    FTimer.Enabled := False;
    Exit;
  end;
  // Someone else owns the screen - the IDE's own install progress, an error
  // box, a project load. Ours can wait: a box over a box is how a user ends
  // up answering one of them blind. The tick count does not advance either,
  // so the catch-up runs when the IDE is really free.
  //
  // TWO tests, because neither covers the other. ModalLevel is raised by
  // TCustomForm.ShowModal and the task-dialog path only, so it does NOT see
  // TApplication.MessageBox - which is exactly what the IDE's own exception
  // box is (measured in Vcl.Forms.pas: MessageBox calls DisableTaskWindows
  // and never ModalStarted). But BOTH paths disable the thread's top-level
  // windows, and so does any owner-modal native dialog, so an enabled main
  // window is the wider test of the two. MainFormHandle is 0 while the IDE
  // is still starting up - then only ModalLevel speaks.
  if Application.ModalLevel > 0 then
    Exit;
  if (Application.MainFormHandle <> 0) and
     not IsWindowEnabled(Application.MainFormHandle) then
    Exit;
  Inc(FTicks);
  if FTicks < DeferTicks then
    Exit;
  // BEFORE anything that pumps: a modal box's own message loop would deliver
  // the next tick and stack a second box on top of this one.
  FTimer.Enabled := False;

  Text := '';
  if TRestartHint.FPinned then
    // The one case that is NOT about this load: this instance is the one that
    // pinned itself, so its own finalization has already run. Nothing of the
    // IDE-only units may be touched here - the hook is cleared by then - and
    // the timer survives for the one reason it is never allowed to anywhere
    // else: a pinned image can no longer be unmapped.
    Text :=
      'Delphi Refactoring Light could not be unloaded completely - a ' +
      'background request of the MCP bridge was still running.' + sLineBreak +
      sLineBreak +
      'This IDE can no longer replace the plugin''s BPL: building or ' +
      'installing it again will fail with "could not create output file", ' +
      'and the plugin is gone from this session.' + sLineBreak +
      sLineBreak +
      'Please restart RAD Studio.'
  else if Assigned(TRestartHint.SettleHook) then
    try
      Text := TRestartHint.SettleHook();
    except
      // The catch-up may not cost the user an exception box of its own; a
      // failure here means the same as "nothing to report".
      Text := '';
    end;

  if Text <> '' then
    MessageDlg(Text, mtWarning, [mbOK], 0);
end;

function ReloadWindowAction(const AFacts: TReloadFacts): TReloadWindowAction;
begin
  Result := rwaNothing;
  if not AFacts.Reloaded then Exit;
  if not AFacts.DockWasOpen then Exit;
  if not AFacts.DockWasDocked then Exit(rwaFloating);
  if AFacts.DockHostSection <> '' then
    Result := rwaRedock
  else
    Result := rwaDockedUnknownHost;
end;

function FindDockHostSection(const ALines: TArray<string>;
  const AIdentifier: string): string;
const
  KeyClients = 'tabdockclients';
var
  Section: string;
begin
  Result := '';
  if AIdentifier = '' then Exit;
  Section := '';
  for var Raw in ALines do
  begin
    var L := Trim(Raw);
    if L = '' then Continue;
    if (L[1] = '[') and (L[Length(L)] = ']') then
    begin
      Section := Copy(L, 2, Length(L) - 2);
      Continue;
    end;
    var Eq := Pos('=', L);
    if Eq <= 1 then Continue;
    if not SameText(Trim(Copy(L, 1, Eq - 1)), KeyClients) then Continue;
    // The value is a comma list of client identifiers. Compare the ENTRIES,
    // never a substring: "RefactoringLightStatus" must not be answered for
    // "RefactoringLightStatusExtra".
    for var Entry in SplitString(Copy(L, Eq + 1, MaxInt), ',') do
      if SameText(Trim(Entry), AIdentifier) then
        Exit(Section);
  end;
end;

function ReloadReportText(const AFacts: TReloadFacts;
  AWindowRestored: Boolean): string;
var
  Lines: TArray<string>;
begin
  Result := '';
  // A fresh IDE start has nothing to catch up and nothing to report: every
  // notification this package listens to is still ahead of it.
  if not AFacts.Reloaded then Exit;

  // THE STATUS WINDOW. Both cases are DONE by the caller now - a
  // free-standing one goes back to its bounds, a docked one is handed to the
  // dock host that held it (that host is the IDE's own form and survives our
  // unload, so it can be found again by its DeskSection). Only a restore that
  // did NOT work is worth a sentence, and then it names what the user can do.
  if AFacts.DockWasOpen and not AWindowRestored then
    if AFacts.DockWasDocked then
      // NOT "open it from the menu": the caller falls back to the window's
      // old bounds, so after a failed re-dock it is usually ON SCREEN and
      // merely free-standing. One action, true in both cases.
      Lines := Lines + ['The status window was docked and could not be put ' +
        'back into its dock site. Pick your desktop again in the toolbar to ' +
        'get your layout back.']
    else
      Lines := Lines + ['The status window was open and could not be ' +
        'reopened. The menu entry opens it again.'];

  // WHAT THE PREVIOUS UNLOAD COULD NOT FINISH. The only half that really
  // asks for a restart - and it is empty unless something went wrong.
  Lines := Lines + AFacts.Problems;

  if Length(Lines) = 0 then Exit;
  Result := 'Delphi Refactoring Light was installed again into this running ' +
    'IDE. Everything it could restore by itself is back; this is what it ' +
    'could not:' + sLineBreak + sLineBreak +
    string.Join(sLineBreak + sLineBreak, Lines);
end;

{ TRestartHint }

class function TRestartHint.GetMarkerFilePath: string;
begin
  // ONE MARKER PER PROCESS (audit #40, L5j): a single shared file meant
  // IDE B overwrote IDE A's pid, so reinstalling in A found B's number and
  // showed no hint - and a stale marker whose pid Windows later reused
  // produced a hint nobody earned. The file is named after the pid and
  // holds the process's CREATION TIME, which is what tells "the same IDE
  // again" from "a new process that happens to have that pid".
  Result := TPath.Combine(TPath.GetTempPath,
    Format('DelphiRefactoringLight.%d.pid', [GetCurrentProcessId]));
end;

/// <summary>The current process's creation time as a stable string - two
///  processes can share a pid over time, never a pid plus a start
///  time.</summary>
class function TRestartHint.ProcessStamp: string;
var
  Created, Exited, Kernel, User: TFileTime;
begin
  Result := '';
  if GetProcessTimes(GetCurrentProcess, Created, Exited, Kernel, User) then
    Result := Format('%d-%d', [Created.dwHighDateTime, Created.dwLowDateTime]);
end;

/// <summary>Removes the markers of processes that are gone, so %TEMP% does
///  not collect one file per IDE start for ever.</summary>
class procedure TRestartHint.PruneDeadMarkers;
var
  Files: TArray<string>;
begin
  try
    Files := TDirectory.GetFiles(TPath.GetTempPath,
      'DelphiRefactoringLight.*.pid');
  except
    Exit;
  end;
  for var F in Files do
  begin
    var Name := TPath.GetFileNameWithoutExtension(F);   // ...Light.<pid>
    var Dot := LastDelimiter('.', Name);
    if Dot <= 0 then Continue;
    var Pid := StrToUIntDef(Copy(Name, Dot + 1, MaxInt), 0);
    if (Pid = 0) or (Pid = GetCurrentProcessId) then Continue;
    // A handle we cannot open at all, or one whose process has exited,
    // means the marker is stale.
    var H := OpenProcess(SYNCHRONIZE, False, Pid);
    var Dead := H = 0;
    if H <> 0 then
    begin
      Dead := WaitForSingleObject(H, 0) = WAIT_OBJECT_0;
      CloseHandle(H);
    end;
    if Dead then
      try TFile.Delete(F); except end;
  end;
end;

class function TRestartHint.ReadMarkerLines: TArray<string>;
begin
  Result := nil;
  try
    if FileExists(GetMarkerFilePath) then
      Result := TFile.ReadAllLines(GetMarkerFilePath);
  except
    Result := nil;
  end;
end;

class procedure TRestartHint.Arm;
begin
  if GDefer = nil then
    GDefer := TRestartHintDefer.Create;
  GDefer.Arm;
end;

class function TRestartHint.Facts: TReloadFacts;
begin
  Result := FFacts;
end;

class function TRestartHint.ReadFacts: TReloadFacts;
var
  Lines: TArray<string>;
  Stamp, L: string;
begin
  PruneDeadMarkers;
  Stamp := ProcessStamp;
  FFacts := Default(TReloadFacts);

  // A marker whose FIRST line is this process's own start stamp was written
  // by an earlier instance of this package IN THIS PROCESS - that, and only
  // that, is a reload.
  Lines := ReadMarkerLines;
  if (Length(Lines) > 0) and (Stamp <> '') and (Trim(Lines[0]) = Stamp) then
  begin
    FFacts.Reloaded := True;
    for var I := 1 to High(Lines) do
    begin
      L := Trim(Lines[I]);
      if L = '' then Continue;
      if StartsText(KeyProblem, L) then
        FFacts.Problems := FFacts.Problems + [Copy(L, Length(KeyProblem) + 1, MaxInt)]
      else if StartsText(KeyHeadless, L) then
        FFacts.Headless := FFacts.Headless + [Copy(L, Length(KeyHeadless) + 1, MaxInt)]
      else if StartsText(KeyDock, L) then
      begin
        // dock=<open>,<docked>,<left>,<top>,<right>,<bottom>[,<host section>]
        // The 7th field arrived with 2.8.0; a marker without it is simply a
        // window whose host we do not know, which is what the old behaviour
        // was - so an older marker can never be a parse error.
        var P := SplitString(Copy(L, Length(KeyDock) + 1, MaxInt), ',');
        if Length(P) >= 6 then
        begin
          FFacts.DockWasOpen := P[0] = '1';
          FFacts.DockWasDocked := P[1] = '1';
          FFacts.DockBounds := TRect.Create(StrToIntDef(P[2], 0),
            StrToIntDef(P[3], 0), StrToIntDef(P[4], 0), StrToIntDef(P[5], 0));
          if Length(P) >= 7 then
            FFacts.DockHostSection := Trim(P[6]);
        end;
      end;
    end;
  end;

  // CONSUMED: the marker goes back to the bare stamp, so the next ordinary
  // IDE start cannot inherit a previous session's findings.
  try
    TFile.WriteAllText(GetMarkerFilePath, Stamp);
  except
    // Tolerate write errors silently - the marker is a convenience, and a
    // failure here must never decide whether the plugin loads.
  end;

  Result := FFacts;
end;

class procedure TRestartHint.Check;
begin
  ReadFacts;
  // ARMED ALWAYS, not only on a reload: the tick is also what catches the
  // initialisation up, and on a fresh start that costs ten ticks that find
  // nothing to do.
  Arm;
end;

class procedure TRestartHint.RecordUnload(const AProblems, AHeadless: TArray<string>;
  ADockOpen, ADockDocked: Boolean; const ADockBounds: TRect;
  const ADockHostSection: string);
var
  SL: TStringList;
begin
  // Best effort throughout: this runs while the package is being unloaded,
  // and an exception here must not cost the rest of the finalization.
  try
    SL := TStringList.Create;
    try
      SL.Add(ProcessStamp);
      for var S in AProblems do
        if S <> '' then SL.Add(KeyProblem + S);
      for var S in AHeadless do
        if S <> '' then SL.Add(KeyHeadless + S);
      if ADockOpen then
        // A section name cannot carry a comma (it is an ini section the IDE
        // writes), so the plain list stays unambiguous.
        SL.Add(Format('%s%d,%d,%d,%d,%d,%d,%s', [KeyDock, 1, Ord(ADockDocked),
          ADockBounds.Left, ADockBounds.Top, ADockBounds.Right, ADockBounds.Bottom,
          ADockHostSection]));
      SL.SaveToFile(GetMarkerFilePath, TEncoding.UTF8);
    finally
      SL.Free;
    end;
  except
  end;
end;

class procedure TRestartHint.NotePinned;
begin
  FPinned := True;
  // Not for the box - for the NEXT report: a session that keeps failing to
  // build the package can be explained from this line, long after the fact.
  try
    TFile.AppendAllText(TPath.Combine(TPath.GetTempPath,
      'DelphiRefactoringLight.pinned.log'),
      Format('%s  pid %d  the BPL stays mapped for the life of this process'#13#10,
        [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), GetCurrentProcessId]));
  except
  end;
  Arm;
end;

initialization

finalization
  // Nothing of ours may survive the package unload (audit #40, M31): a timer
  // that is still armed would tick into unmapped code.
  // THE ONE EXCEPTION, and it is narrow on purpose: when this instance pinned
  // its own module, that image can never be unmapped - so the timer is the
  // only channel left to tell the user why the next build will fail, and it
  // is safe for exactly the reason that makes everything else unsafe.
  if not TRestartHint.FPinned then
    FreeAndNil(GDefer);

end.
