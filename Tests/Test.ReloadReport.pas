unit Test.ReloadReport;

{ What a package RELOAD - "Installieren" in a running IDE - loses, and what
  the next instance can put back by itself.

  The box this is about used to be unconditional and claimed a mechanism
  nobody had measured ("the IDE still holds references to the previous
  version"). What a reload really loses is INITIALISATION: state that is set
  up from IDE notifications which cannot fire again for a project that is
  already open. So the new rule is "catch up, then report only the rest" -
  and the two halves are deliberately apart: the catch-up needs ToolsAPI and
  a live IDE, the DECISION does not, so it is checkable here. }

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TReloadReportTests = class
  public
    /// <summary>The report says only what could NOT be restored - and the
    ///  status window is the case that separates the two halves: a
    ///  free-standing one is put back, a DOCKED one can only be said.</summary>
    [Test] procedure AReloadReportsOnlyWhatItCouldNotRestore;
    /// <summary>The marker file is the only channel from the instance that
    ///  is going away to the one that loads next, so what the unload knew
    ///  has to survive exactly one hand-over - and be consumed.</summary>
    [Test] procedure TheMarkerCarriesWhatTheUnloadKnew;
  end;

implementation

uses
  System.SysUtils, System.Types, System.IOUtils,
  Expert.RestartHint;

procedure TReloadReportTests.AReloadReportsOnlyWhatItCouldNotRestore;
var
  F: TReloadFacts;
  Reopen: Boolean;
  S: string;
begin
  // A FRESH IDE START reports nothing at all - every notification this
  // package listens to is still ahead of it. That is the common case, and
  // the old box fired here too.
  F := Default(TReloadFacts);
  F.DockWasOpen := True;         // and not even an open window changes it
  F.Problems := ['something'];
  Assert.AreEqual('', ReloadReportText(F, Reopen), 'a fresh start says nothing');
  Assert.IsFalse(Reopen, 'and puts nothing back');

  // A RELOAD WITH NOTHING LEFT OVER also says nothing: everything the
  // catch-up can do has been done by the time this runs.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  Assert.AreEqual('', ReloadReportText(F, Reopen),
    'a reload whose catch-up restored everything says nothing');
  Assert.IsFalse(Reopen);

  // A FREE-STANDING status window is PUT BACK, with no word about it - the
  // caller has its bounds and ShowStatusWindowAt reproduces them exactly.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.DockWasOpen := True;
  F.DockWasDocked := False;
  F.DockBounds := TRect.Create(100, 120, 700, 500);
  Assert.AreEqual('', ReloadReportText(F, Reopen),
    'a free-standing window is restored, so there is nothing to report');
  Assert.IsTrue(Reopen, 'and the caller is told to put it back');

  // A DOCKED one is SAID. Measured in the user's own project desktop: the
  // place is stored with the dock site, not with our window, and ToolsAPI
  // has no entry point that restores one client - opening it anyway gives
  // the floating window in the middle of the screen he reported.
  F.DockWasDocked := True;
  S := ReloadReportText(F, Reopen);
  Assert.IsFalse(Reopen,
    'a docked window must NOT be opened - that is the reported defect');
  Assert.IsTrue(S.Contains('docked'), 'and it is said instead: ' + S);
  Assert.IsTrue(S.Contains('desktop'),
    'naming the one thing that can restore it: ' + S);

  // WHAT THE UNLOAD COULD NOT FINISH is the only half that really asks for
  // a restart, and it is passed through verbatim.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.Problems := ['A blame worker did not finish in time.',
    'Another add-on chained its editor context menu onto ours.'];
  S := ReloadReportText(F, Reopen);
  for var P in F.Problems do
    Assert.IsTrue(S.Contains(P), 'every problem is named: ' + P);

  // Both halves at once - the window line AND the problems.
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  S := ReloadReportText(F, Reopen);
  Assert.IsTrue(S.Contains('docked') and S.Contains('blame worker'),
    'both halves appear: ' + S);
end;

procedure TReloadReportTests.TheMarkerCarriesWhatTheUnloadKnew;
var
  F: TReloadFacts;
begin
  // START FROM A KNOWN STATE. The marker is named after this process, so
  // the one we delete here is the test exe's own - never an IDE's.
  if TFile.Exists(TRestartHint.GetMarkerFilePath) then
    TFile.Delete(TRestartHint.GetMarkerFilePath);
  try
    // NO MARKER = this package has not been loaded in this process before,
    // i.e. a fresh start. Nothing to catch up, nothing to report.
    F := TRestartHint.ReadFacts;
    Assert.IsFalse(F.Reloaded, 'without a marker it is a fresh start');

    // What an unload knows, written for the instance that loads next.
    TRestartHint.RecordUnload(
      ['A blame worker did not finish in time.'],
      ['C:\work\Unit1.pas', 'C:\work\Unit2.pas'],
      True, False, TRect.Create(10, 20, 410, 320));

    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.Reloaded,
      'a marker with THIS process''s stamp means the package was loaded here before');
    Assert.AreEqual<Integer>(1, Length(F.Problems), 'the problem survives');
    Assert.AreEqual('A blame worker did not finish in time.', F.Problems[0]);
    Assert.AreEqual<Integer>(2, Length(F.Headless), 'both headless buffers survive');
    Assert.AreEqual('C:\work\Unit2.pas', F.Headless[1]);
    Assert.IsTrue(F.DockWasOpen, 'the window was open');
    Assert.IsFalse(F.DockWasDocked, 'and free-standing');
    Assert.AreEqual<Integer>(410, F.DockBounds.Right, 'with its own geometry');

    // CONSUMED: the findings are a one-shot. Were they not, every later load
    // in this process would report a problem that was dealt with long ago -
    // and the reload flag itself stays true, because it IS still a reload.
    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.Reloaded, 'still the same process, so still a reload');
    Assert.AreEqual<Integer>(0, Length(F.Problems),
      'but the findings were consumed by the load that read them');
    Assert.IsFalse(F.DockWasOpen, 'and so was the window state');
  finally
    if TFile.Exists(TRestartHint.GetMarkerFilePath) then
      TFile.Delete(TRestartHint.GetMarkerFilePath);
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TReloadReportTests);

end.
