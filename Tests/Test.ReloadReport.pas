unit Test.ReloadReport;

{ What a package RELOAD - "Installieren" in a running IDE - loses, and what
  the next instance can put back by itself.

  The box this is about used to be unconditional and claimed a mechanism
  nobody had measured ("the IDE still holds references to the previous
  version"). What a reload really loses is INITIALISATION: state that is set
  up from IDE notifications which cannot fire again for a project that is
  already open. So the rule is "catch up, then report only the rest" - and
  the two halves are deliberately apart: the catch-up needs ToolsAPI and a
  live IDE, the DECISION does not, so it is checkable here.

  The docked window used to be the one thing that was only SAID. It is done
  now: the dock HOST is the IDE's own form and survives our unload, so it can
  be found again by its DeskSection - which the desktop file is also keyed on
  ("[DockSite5] TabDockClients=MessageView,RefactoringLightStatus"). }

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TReloadReportTests = class
  public
    /// <summary>What the caller is told to DO with the window - and the
    ///  docked case is an action now, not a dead end.</summary>
    [Test] procedure TheWindowGoesBackWhereItWas;
    /// <summary>The report says only what could NOT be restored. A sentence
    ///  that sends the user to the desktop menu must not appear while the
    ///  window is already sitting in its place.</summary>
    [Test] procedure AReloadReportsOnlyWhatItCouldNotRestore;
    /// <summary>The desktop file is the FALLBACK for the dock host, and its
    ///  shape is the user's own: the membership is recorded with the site.</summary>
    [Test] procedure TheDesktopFileNamesTheDockHost;
    /// <summary>The marker file is the only channel from the instance that
    ///  is going away to the one that loads next, so what the unload knew
    ///  has to survive exactly one hand-over - and be consumed.</summary>
    [Test] procedure TheMarkerCarriesWhatTheUnloadKnew;
  end;

implementation

uses
  System.SysUtils, System.Types, System.IOUtils,
  Expert.RestartHint;

{ The dock site of the user's own project desktop, verbatim - this is the
  shape the fallback has to read, and the reason our own section cannot
  answer: it carries Docked=1 and geometry, never the host. }
const
  RealDesktop: array[0..19] of string = (
    '[RefactoringLightStatus]',
    'PercentageSizes=1',
    'Create=1',
    'Visible=1',
    'Docked=1',
    'Left=0',
    'Top=0',
    'Width=5981',
    'Height=3044',
    'Dockable=1',
    'StayOnTop=0',
    '',
    '[DockSite5]',
    'HostDockSite=DockBottomCenterPanel',
    'DockSiteType=1',
    'Create=1',
    'Visible=1',
    'Docked=1',
    'TabPosition=1',
    'TabDockClients=MessageView,RefactoringLightStatus');

function DeskLines(const AItems: array of string): TArray<string>;
begin
  SetLength(Result, Length(AItems));
  for var I := 0 to High(AItems) do
    Result[I] := AItems[I];
end;

procedure TReloadReportTests.TheWindowGoesBackWhereItWas;
var
  F: TReloadFacts;
begin
  // A FRESH IDE START has nothing to put back.
  F := Default(TReloadFacts);
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockHostSection := 'DockSite5';
  Assert.AreEqual<TReloadWindowAction>(rwaNothing, ReloadWindowAction(F),
    'a fresh start restores nothing');

  // A RELOAD WITH THE WINDOW CLOSED: still nothing.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  Assert.AreEqual<TReloadWindowAction>(rwaNothing, ReloadWindowAction(F),
    'it was not open');

  // FREE-STANDING: put it back at its bounds.
  F.DockWasOpen := True;
  Assert.AreEqual<TReloadWindowAction>(rwaFloating, ReloadWindowAction(F),
    'a free-standing window goes back to its bounds');

  // DOCKED AND THE HOST IS KNOWN: hand it to that host. THIS is the case the
  // whole round is about - it used to be "say it and give up".
  F.DockWasDocked := True;
  F.DockHostSection := 'DockSite5';
  Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
    'a docked window is handed back to its dock host');

  // DOCKED AND THE HOST IS NOT KNOWN - an older marker, or a host we could
  // not read. Then there is genuinely nothing to put it into.
  F.DockHostSection := '';
  Assert.AreEqual<TReloadWindowAction>(rwaDockedUnknownHost,
    ReloadWindowAction(F), 'docked, but nothing says where');
end;

procedure TReloadReportTests.AReloadReportsOnlyWhatItCouldNotRestore;
var
  F: TReloadFacts;
  S: string;
begin
  // A FRESH IDE START reports nothing at all - that is the common case, and
  // the old box fired here too.
  F := Default(TReloadFacts);
  F.DockWasOpen := True;
  F.Problems := ['something'];
  Assert.AreEqual('', ReloadReportText(F, False), 'a fresh start says nothing');

  // A RESTORED WINDOW IS NOT MENTIONED, docked or not. This is the sentence
  // the user objected to: it must not appear once the window is back.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockHostSection := 'DockSite5';
  Assert.AreEqual('', ReloadReportText(F, True),
    'the window is back in its dock - there is nothing to report');

  // ONLY A RESTORE THAT FAILED is worth a sentence, and then it names what
  // the user can do.
  S := ReloadReportText(F, False);
  Assert.IsTrue(S.Contains('docked'), 'a docked window that stayed away is said: ' + S);
  Assert.IsTrue(S.Contains('desktop'), 'and it names what can still do it: ' + S);

  // A FREE-STANDING one that could not be reopened gets its own, different
  // sentence - telling someone to re-select a desktop would be nonsense.
  F.DockWasDocked := False;
  S := ReloadReportText(F, False);
  Assert.IsTrue(S.Contains('menu'), 'the free-standing case points at the menu: ' + S);
  Assert.IsFalse(S.Contains('desktop'),
    'and does NOT send the user to the desktop menu: ' + S);

  // WHAT THE UNLOAD COULD NOT FINISH is the only half that really asks for
  // a restart, and it is passed through verbatim whatever the window did.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.Problems := ['A blame worker did not finish in time.',
    'Another add-on chained its editor context menu onto ours.'];
  S := ReloadReportText(F, True);
  for var P in F.Problems do
    Assert.IsTrue(S.Contains(P), 'every problem is named: ' + P);
end;

procedure TReloadReportTests.TheDesktopFileNamesTheDockHost;
begin
  // The user's own desktop, verbatim: the host is the SITE's section.
  Assert.AreEqual('DockSite5',
    FindDockHostSection(DeskLines(RealDesktop), 'RefactoringLightStatus'),
    'the site that lists us is the answer');

  // MessageView sits in the same list, so the entry really is matched per
  // entry and not by position.
  Assert.AreEqual('DockSite5',
    FindDockHostSection(DeskLines(RealDesktop), 'MessageView'));

  // Case does not matter - an ini is not case-sensitive.
  Assert.AreEqual('DockSite5',
    FindDockHostSection(DeskLines(RealDesktop), 'refactoringlightstatus'));

  // WHOLE ENTRIES ONLY. A substring match would answer for a window that is
  // not ours, and re-docking into a stranger's site is worse than not
  // re-docking at all.
  Assert.AreEqual('',
    FindDockHostSection(DeskLines(RealDesktop), 'RefactoringLightStatusExtra'),
    'a longer name is another window');
  Assert.AreEqual('',
    FindDockHostSection(DeskLines(RealDesktop), 'Status'),
    'and a shorter one is not us either');

  // A window that is in no dock site, and the empty question.
  Assert.AreEqual('',
    FindDockHostSection(DeskLines(RealDesktop), 'RefactoringLightMcpTools'));
  Assert.AreEqual('', FindDockHostSection(DeskLines(RealDesktop), ''));
  Assert.AreEqual('', FindDockHostSection(TArray<string>(nil), 'RefactoringLightStatus'));

  // Blanks around the entries, which an ini may carry.
  Assert.AreEqual('DockSite9', FindDockHostSection(DeskLines([
    '[DockSite9]', 'TabDockClients = MessageView , RefactoringLightStatus ']),
    'RefactoringLightStatus'), 'blanks around an entry are not part of it');
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
      True, True, TRect.Create(10, 20, 410, 320), 'DockSite5');

    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.Reloaded,
      'a marker with THIS process''s stamp means the package was loaded here before');
    Assert.AreEqual<Integer>(1, Length(F.Problems), 'the problem survives');
    Assert.AreEqual('A blame worker did not finish in time.', F.Problems[0]);
    Assert.AreEqual<Integer>(2, Length(F.Headless), 'both headless buffers survive');
    Assert.AreEqual('C:\work\Unit2.pas', F.Headless[1]);
    Assert.IsTrue(F.DockWasOpen, 'the window was open');
    Assert.IsTrue(F.DockWasDocked, 'and docked');
    Assert.AreEqual('DockSite5', F.DockHostSection,
      'and THE HOST comes through - without it a docked window cannot be put back');
    Assert.AreEqual<Integer>(410, F.DockBounds.Right, 'with its own geometry');
    Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
      'so the next instance knows to re-dock');

    // CONSUMED: the findings are a one-shot. Were they not, every later load
    // in this process would report a problem that was dealt with long ago -
    // and the reload flag itself stays true, because it IS still a reload.
    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.Reloaded, 'still the same process, so still a reload');
    Assert.AreEqual<Integer>(0, Length(F.Problems),
      'but the findings were consumed by the load that read them');
    Assert.IsFalse(F.DockWasOpen, 'and so was the window state');
    Assert.AreEqual('', F.DockHostSection, 'including the host');

    // AN OLDER MARKER - six fields, written before the host was carried -
    // must still parse. It then means "docked, host unknown", which is
    // exactly what the old behaviour was.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dock=1,1,10,20,410,320' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.DockWasDocked, 'the old six-field form still parses');
    Assert.AreEqual('', F.DockHostSection, 'it just cannot say where');
    Assert.AreEqual<TReloadWindowAction>(rwaDockedUnknownHost,
      ReloadWindowAction(F), 'and that is a case of its own, not a crash');
  finally
    if TFile.Exists(TRestartHint.GetMarkerFilePath) then
      TFile.Delete(TRestartHint.GetMarkerFilePath);
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TReloadReportTests);

end.
