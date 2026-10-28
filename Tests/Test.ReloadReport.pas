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
  now, because the dock sites belong to the IDE and survive our unload - but
  not "by its DeskSection", which is what this comment claimed until 2.9.0
  and which 2.8.2 had already refuted: a host the user created by dragging
  has no DeskSection at all, so the lookup is a name / section / site ladder.
  AND A PLACE IS NOT THE WHOLE STATE: being a tab is a RELATION to another
  client, so the host's other clients travel too ("[DockSite5]
  TabDockClients=MessageView,RefactoringLightStatus" is the same list, which
  is why the saved desktop can restore the relation as well). }

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
    /// <summary>A dock host the IDE created at RUN TIME - because the user
    ///  dragged the window to dock it - has no DeskSection, and that is the
    ///  reported defect: everything was keyed on the section, so our own row
    ///  read "docked, host unknown" and the reload had nothing to put the
    ///  window back into. Its component NAME is there either way, and the
    ///  chain carries the site that host was itself docked in.</summary>
    [Test] procedure ARunTimeDockHostIsStillAHost;
    [Test] procedure APlaceWithoutItsTabIsNotARestore;
    [Test] procedure TheFirstInstallMustNotStopLooking;
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

function DockLinkOf(const AFormName, AFormClass, ASection,
  ASiteName: string): TDockLink;
begin
  Result := Default(TDockLink);
  Result.FormName := AFormName;
  Result.FormClass := AFormClass;
  Result.Section := ASection;
  Result.SiteName := ASiteName;
end;

function PlaceOf(const ALinks: array of TDockLink): TDockPlace;
begin
  Result := Default(TDockPlace);
  SetLength(Result.Links, Length(ALinks));
  for var I := 0 to High(ALinks) do
    Result.Links[I] := ALinks[I];
end;

{ The shape of the reported dock place, innermost first - and the comment
  says per link what was MEASURED and what follows from it, because the first
  version of this fixture asserted an outer link the live IDE contradicts.

  LINK 0 is measured exactly (ide_introspect, 2026-10-28): our window
  OTADockForm_1/TOTADockForm sat in the form TabDockHostForm/TTabDockHostForm,
  which the user had created by DRAGGING the window there, and whose
  DeskSection is empty - that emptiness IS the reported defect. Its child
  panel is called DockPanel. Which of the two a client's HostDockSite really
  is, the host form or that panel, was NOT measured; the restore does not care
  (a tab host is docked into through its own form either way) and the case
  where the site name IS the form's name has its own assertion below.

  LINK 1 follows from where that host sits. What is measured is the OWNER of
  the panel: DockBottomCenterPanel is a component of EditWindow_0, a
  TEditWindow (components:EditWindow_0, index 17), and the main window has no
  child of that name at all (children:AppBuilder) - which is why 2.8.2's first
  version of this fixture, and the constant behind it, named AppBuilder and
  'Main Window' wrongly. The section of the editor window was not read, so the
  link carries none. }
function DockLinkTransient(const AForm, AClass, ASection,
  ASite: string): TDockLink;
begin
  Result := DockLinkOf(AForm, AClass, ASection, ASite);
  Result.Transient := True;
end;

function MeasuredPlace: TDockPlace;
begin
  Result := PlaceOf([
    DockLinkOf('TabDockHostForm', 'TTabDockHostForm', '', 'DockPanel'),
    DockLinkOf('EditWindow_0', 'TEditWindow', '', 'DockBottomCenterPanel')]);
end;

procedure TReloadReportTests.TheWindowGoesBackWhereItWas;
var
  F: TReloadFacts;
begin
  // A FRESH IDE START has nothing to put back.
  F := Default(TReloadFacts);
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockPlace := MeasuredPlace;
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
  F.DockPlace := MeasuredPlace;
  Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
    'a docked window is handed back to its dock host');

  // DOCKED AND THE CHAIN IS NOT KNOWN - an older marker, or a window whose
  // host we could not read. Then there is genuinely nothing to put it into.
  F.DockPlace := Default(TDockPlace);
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
  Assert.AreEqual('', ReloadReportText(F, rwoNotDone),
    'a fresh start says nothing');

  // A RESTORED WINDOW IS NOT MENTIONED, docked or not. This is the sentence
  // the user objected to: it must not appear once the window is back.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockPlace := MeasuredPlace;
  Assert.AreEqual('', ReloadReportText(F, rwoRestored),
    'the window is back in its dock - there is nothing to report');

  // ONLY A RESTORE THAT FAILED is worth a sentence, and then it names what
  // the user can do.
  S := ReloadReportText(F, rwoNotDone);
  Assert.IsTrue(S.Contains('docked'), 'a docked window that stayed away is said: ' + S);
  Assert.IsTrue(S.Contains('desktop'), 'and it names what can still do it: ' + S);

  // A FREE-STANDING one that could not be reopened gets its own, different
  // sentence - telling someone to re-select a desktop would be nonsense.
  F.DockWasDocked := False;
  S := ReloadReportText(F, rwoNotDone);
  Assert.IsTrue(S.Contains('menu'), 'the free-standing case points at the menu: ' + S);
  Assert.IsFalse(S.Contains('desktop'),
    'and does NOT send the user to the desktop menu: ' + S);

  // WHAT THE UNLOAD COULD NOT FINISH is the only half that really asks for
  // a restart, and it is passed through verbatim whatever the window did.
  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.Problems := ['A blame worker did not finish in time.',
    'Another add-on chained its editor context menu onto ours.'];
  S := ReloadReportText(F, rwoRestored);
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

  // AND THE WHOLE CHAIN THE FILE CAN GIVE, which is the half that carries the
  // FIRST install of every version: the host section it lists, and the site
  // THAT host sits in. The second link is the one that is still there when
  // the host is not, so its absence would be invisible and fatal in exactly
  // the case this fallback exists for.
  var Place := DesktopDockPlace(DeskLines(RealDesktop), 'RefactoringLightStatus');
  Assert.AreEqual<Integer>(2, Length(Place.Links), 'host and the site it sits in');
  Assert.AreEqual('DockSite5', Place.Links[0].Section, 'the host comes FIRST');
  Assert.AreEqual('DockBottomCenterPanel', Place.Links[1].SiteName,
    'and the surviving site second');
  // NO OWNER IS INVENTED for that site. 2.8.2 first wrote 'Main Window' here
  // and the live IDE refuted it: DockBottomCenterPanel is a component of
  // EditWindow_0 (TEditWindow), and the main window has no child of that
  // name - so the lookup searched the wrong form and this fallback could
  // never work.
  Assert.AreEqual('', Place.Links[1].Section,
    'the file says whose panel it is nowhere, so neither does the link');
  Assert.AreEqual('', Place.Links[1].FormName);
  Assert.IsTrue(Place.Links[1].Known,
    'and a site name alone is still a link - the lookup finds its owner');

  // A host section without a HostDockSite line is one link, not a broken two.
  Assert.AreEqual<Integer>(1, Length(DesktopDockPlace(DeskLines(
    ['[DockSite9]', 'TabDockClients=RefactoringLightStatus']),
    'RefactoringLightStatus').Links), 'no panel named, no second link');
  // And a window the file does not list is no place at all.
  Assert.AreEqual<Integer>(0, Length(DesktopDockPlace(DeskLines(RealDesktop),
    'RefactoringLightMcpTools').Links));

  // THE VALUE READER belongs to the HOST's section and to no other.
  Assert.AreEqual('DockBottomCenterPanel',
    DeskValueIn(DeskLines(RealDesktop), 'DockSite5', 'HostDockSite'));
  Assert.AreEqual('', DeskValueIn(DeskLines(RealDesktop),
    'RefactoringLightStatus', 'HostDockSite'),
    'our own section does not say where it was - that is the whole point');
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
      True, True, TRect.Create(10, 20, 410, 320), MeasuredPlace);

    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.Reloaded,
      'a marker with THIS process''s stamp means the package was loaded here before');
    Assert.AreEqual<Integer>(1, Length(F.Problems), 'the problem survives');
    Assert.AreEqual('A blame worker did not finish in time.', F.Problems[0]);
    Assert.AreEqual<Integer>(2, Length(F.Headless), 'both headless buffers survive');
    Assert.AreEqual('C:\work\Unit2.pas', F.Headless[1]);
    Assert.IsTrue(F.DockWasOpen, 'the window was open');
    Assert.IsTrue(F.DockWasDocked, 'and docked');
    Assert.AreEqual<Integer>(2, Length(F.DockPlace.Links),
      'and THE WHOLE CHAIN comes through - whether the host that held the ' +
      'window survives our unload depends on how it was made, so the link ' +
      'behind it has to travel too');
    Assert.AreEqual('TabDockHostForm', F.DockPlace.Links[0].FormName,
      'innermost first: the host that held it');
    Assert.AreEqual('TTabDockHostForm', F.DockPlace.Links[0].FormClass,
      'with its class, which rejects a name recycled by ANOTHER class - ' +
      'and no more than that: every run-time dock host is a TTabDockHostForm');
    Assert.AreEqual('DockPanel', F.DockPlace.Links[0].SiteName,
      'and the site control inside it');
    Assert.AreEqual('EditWindow_0', F.DockPlace.Links[1].FormName,
      'then the form THAT host was docked in');
    Assert.AreEqual('DockBottomCenterPanel', F.DockPlace.Links[1].SiteName,
      'and the panel it sat in - which is where the window goes when the ' +
      'host is gone');
    Assert.AreEqual<Integer>(410, F.DockBounds.Right, 'with its own geometry');
    // WHETHER THE HOST'S NAME WILL STILL MEAN IT travels too, and it has to:
    // "TabDockHostForm" is the class default every TTabDockHostForm carries,
    // so for a host that dissolves on our unload the next instance must NOT
    // look it up by that name. The flag is what says so.
    TRestartHint.RecordUnload(nil, nil, True, True, TRect.Create(0, 0, 0, 0),
      PlaceOf([DockLinkTransient('TabDockHostForm', 'TTabDockHostForm', '',
                 'DockPanel'),
               DockLinkOf('EditWindow_0', 'TEditWindow', '',
                 'DockBottomCenterPanel')]));
    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.DockPlace.Links[0].Transient,
      'the drag-created host is marked as what it is');
    Assert.IsFalse(F.DockPlace.Links[1].Transient,
      'and the window that owns the panel behind it is not');
    Assert.IsTrue(F.DockPlace.Links[0].Known,
      'a transient link is still reported - it is the LOOKUP that refuses it');

    // A dockat= LINE WITHOUT THE FIFTH FIELD is a marker of 2.8.2's own first
    // builds, and nothing before it could know this. It must read as "the
    // name was trusted", never as transient - that is the direction in which
    // a wrong default would silently stop restoring anything.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dockat=DockSite5|TTabDockHostForm|DockSite5|DockPanel' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.IsFalse(F.DockPlace.Links[0].Transient,
      'four fields mean the name was trusted, as every version before this');

    // A SECTION WITH A BLANK IN IT survives the round trip - the marker's
    // fields are separated by '|', so a space inside one is ordinary text.
    TRestartHint.RecordUnload(nil, nil, True, True, TRect.Create(0, 0, 0, 0),
      PlaceOf([DockLinkOf('AppBuilder', 'TAppBuilder', 'Main Window', 'PanelBottom')]));
    F := TRestartHint.ReadFacts;
    Assert.AreEqual('Main Window', F.DockPlace.Links[0].Section,
      'a section with a blank in it comes back whole');
    Assert.AreEqual('PanelBottom', F.DockPlace.Links[0].SiteName);
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
    Assert.AreEqual<Integer>(0, Length(F.DockPlace.Links),
      'including the chain');

    // AN OLDER MARKER - six fields, written before the host was carried -
    // must still parse. It then means "docked, host unknown", which is
    // exactly what the old behaviour was.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dock=1,1,10,20,410,320' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.IsTrue(F.DockWasDocked, 'the old six-field form still parses');
    Assert.IsFalse(F.DockPlace.Known, 'it just cannot say where');
    Assert.AreEqual<TReloadWindowAction>(rwaDockedUnknownHost,
      ReloadWindowAction(F), 'and that is a case of its own, not a crash');

    // AND THE SEVEN-FIELD FORM of 2.8.0, which is not a museum piece: an
    // install over a running older version reads a marker that older version
    // wrote, so this is the shape the FIRST load after this change meets.
    // It names a host and no site, which is still a re-dock - a tab host is
    // docked into through its own form.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dock=1,1,10,20,410,320,DockSite5' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.AreEqual<Integer>(1, Length(F.DockPlace.Links),
      'the 2.8.x form is exactly what it is: one link');
    Assert.AreEqual('DockSite5', F.DockPlace.Links[0].Section,
      'named by its section, which is all that marker carries');
    Assert.AreEqual('', F.DockPlace.Links[0].FormName,
      'and no component name is invented for it - the lookup tries the ' +
      'section as well, and for a host the desktop created the two strings ' +
      'are the same anyway');
    Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
      'and a host without a site is still enough to try');

    // AND THE SHAPE 2.8.1 REALLY WROTE, which is what the first load after
    // this change meets on the reported machine: EIGHT fields, and for a
    // drag-created host the section field is EMPTY, because that is the
    // defect 2.8.2 is about. So the marker can say only "it was docked" -
    // and with no bounds either, since 2.8.1 recorded none for a docked
    // window. THIS is why the desktop-file fallback has to carry the first
    // install of every version, and why its second link matters.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dock=1,1,0,0,0,0,,DockPanel' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.AreEqual<Integer>(1, Length(F.DockPlace.Links),
      'the site name is all it carries, and that is still a link to try');
    Assert.AreEqual('DockPanel', F.DockPlace.Links[0].SiteName);
    Assert.AreEqual('', F.DockPlace.Links[0].Section,
      'the section is empty - the host was created by a drag, which IS the ' +
      'defect 2.8.2 is about');
    Assert.AreEqual('', F.DockPlace.Links[0].FormName,
      'and a marker says nothing about which form owned that site');
    // SO IT IS TRIED, and in a real IDE it resolves to nothing: "DockPanel"
    // is the name EVERY tab host gives its panel (measured: DockSite0..4 and
    // the user's drag-created host all have one), so the lookup finds several
    // owners and refuses rather than guessing. That is what leaves the
    // desktop-file fallback its turn - and with no bounds either, because
    // 2.8.1 recorded none for a docked window, the fallback is the only
    // thing that can place this window.
    Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
      'a link worth trying is tried');
    Assert.AreEqual<Integer>(0, F.DockBounds.Width,
      'and there is not even a rectangle behind it');
    // An eight-field marker that DOES name a section still re-docks.
    TFile.WriteAllText(TRestartHint.GetMarkerFilePath,
      TFile.ReadAllText(TRestartHint.GetMarkerFilePath).Trim + sLineBreak +
      'dock=1,1,10,20,410,320,DockSite5,DockPanel' + sLineBreak);
    F := TRestartHint.ReadFacts;
    Assert.AreEqual('DockSite5', F.DockPlace.Links[0].Section);
    Assert.AreEqual('DockPanel', F.DockPlace.Links[0].SiteName,
      'the eighth field is the site, and it is read');

    // AND THE dock= LINE KEEPS ITS 7TH AND 8TH FIELD, which THIS version
    // never reads (every such link has its own dockat= line). They exist for
    // an OLDER build loaded into the same process, which skips dockat=
    // entirely - so only an assertion on what is WRITTEN can keep them: a
    // comment would be weighed against "the reader never looks at this".
    TRestartHint.RecordUnload(nil, nil, True, True, TRect.Create(1, 2, 3, 4),
      PlaceOf([DockLinkOf('DockSite5', 'TTabDockHostForm', 'DockSite5',
        'DockPanel')]));
    var Dock := '';
    for var L in TFile.ReadAllLines(TRestartHint.GetMarkerFilePath) do
      if L.StartsWith('dock=') then Dock := L;
    Assert.AreNotEqual('', Dock, 'the legacy line is still written');
    var Fields := Dock.Substring(5).Split([',']);
    Assert.AreEqual<Integer>(8, Length(Fields),
      'eight fields, as 2.8.1 wrote them');
    Assert.AreEqual('DockSite5', Fields[6],
      'the 7th is the innermost host section - what a 2.8.0 reader takes');
    Assert.AreEqual('DockPanel', Fields[7],
      'and the 8th its site, which a 2.8.1 reader takes');

    // A SEPARATOR INSIDE A NAME may not move the field behind it. Both lines
    // of the marker are separated lists, and a mis-parsed list would hand the
    // next instance a name that points at ANOTHER window; a name that lost
    // its separator is simply not found again, which is the honest failure.
    TRestartHint.RecordUnload(nil, nil, True, True, TRect.Create(0, 0, 0, 0),
      PlaceOf([DockLinkOf('Odd|Name', 'TOdd,Class', 'Odd, Section', 'DockPanel')]));
    F := TRestartHint.ReadFacts;
    Assert.AreEqual<Integer>(1, Length(F.DockPlace.Links),
      'one link, and the separators did not split it into more');
    Assert.AreEqual('DockPanel', F.DockPlace.Links[0].SiteName,
      'the LAST field is still the site name - that is what a field shift breaks');
    Assert.IsFalse(F.DockPlace.Links[0].FormName.Contains('|'),
      'and no name carries a separator into the list');
    Assert.IsFalse(F.DockPlace.Links[0].Section.Contains(','),
      'neither into the comma list');
  finally
    if TFile.Exists(TRestartHint.GetMarkerFilePath) then
      TFile.Delete(TRestartHint.GetMarkerFilePath);
  end;
end;

procedure TReloadReportTests.ARunTimeDockHostIsStillAHost;
var
  F: TReloadFacts;
begin
  // THE REPORTED SHAPE, and the whole defect in one assertion: this host has
  // no DeskSection, so everything that keyed on the section called it
  // unknown - while ide_introspect names it "TabDockHostForm" and a form's
  // component name is unique among its owner's components.
  Assert.IsTrue(MeasuredPlace.Links[0].Known,
    'a host with no DeskSection is still a host - its NAME says where it is');
  Assert.IsTrue(MeasuredPlace.Known, 'so the place is known');

  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockPlace := MeasuredPlace;
  Assert.AreEqual<TReloadWindowAction>(rwaRedock, ReloadWindowAction(F),
    'and THAT is the defect: this used to be rwaDockedUnknownHost, so the ' +
    'window came back free-standing in the middle of the screen');

  // A LINK THAT NAMES NOTHING is not a link. An empty site name is fine - a
  // tab host is docked into through its own form - but a link with neither
  // name could never be looked up again.
  // A SITE NAME ALONE *IS* A LINK, and this assertion used to say the
  // opposite - my own reasoning, refuted by the live IDE: the saved desktop
  // records a dock place as exactly that ("HostDockSite=DockBottomCenterPanel",
  // with the owner named nowhere), and the lookup answers it by finding the
  // form that OWNS a control of that name. Only a link with nothing at all
  // is nothing.
  Assert.IsTrue(DockLinkOf('', '', '', 'DockPanel').Known,
    'the saved desktop names a site and no owner, and that is findable');
  Assert.IsFalse(Default(TDockLink).Known,
    'while a link with no name at all is nothing');
  Assert.IsTrue(DockLinkOf('TabDockHostForm', 'TTabDockHostForm', '', '').Known,
    'a form name alone can - and a tab host needs no site control');

  // WHAT THE STATUS ROW SAYS, which is where this defect was read in one
  // call: the host, and then what the host is itself docked in. Both
  // outcomes of a reload are named before either of them happens.
  Assert.AreEqual('TabDockHostForm (DockPanel) in ' +
    'EditWindow_0 (DockBottomCenterPanel)', MeasuredPlace.Describe,
    'the row names the host and its fallback');
  Assert.AreEqual('DockSite5 (DockPanel)',
    PlaceOf([DockLinkOf('DockSite5', 'TTabDockHostForm', 'DockSite5',
      'DockPanel')]).Describe,
    'and for a host the desktop created it reads as it always did - the ' +
    'name and the section are the same string there');
  // AND WHEN THE HOST FORM IS ITS OWN DOCK SITE the name is already there.
  // Which of the two a client's HostDockSite really is - the host form or its
  // DockPanel - was not measured, and the row has to read as one thing in
  // both cases; the restore itself does not care, because a tab host is
  // docked into through its own form either way.
  Assert.AreEqual('TabDockHostForm',
    DockLinkOf('TabDockHostForm', 'TTabDockHostForm', '',
      'TabDockHostForm').Describe,
    'a form that is its own dock site is named ONCE, not twice');

  // THE INNERMOST LINK is the host that held the window, and it is what the
  // desktop fallback compares itself against.
  Assert.AreEqual('TabDockHostForm', MeasuredPlace.Innermost.FormName,
    'innermost first');
  Assert.AreEqual('', Default(TDockPlace).Innermost.FormName,
    'and asking an empty place is not a range error');

  // A CHAIN OF ONE, built from a section - the saved desktop file, and a
  // marker written by 2.8.x. THE FIELDS, not just Known: a section is not a
  // component name, and filling FormName from it would make the lookup's
  // first pass match any form merely NAMED 'DockSite5', bypassing the
  // TDesktopForm/DeskSection test the section pass applies.
  Assert.IsTrue(DockPlaceOfSection('DockSite5').Known,
    'the desktop file names a host and nothing else, and that is enough');
  Assert.AreEqual('DockSite5', DockPlaceOfSection('DockSite5').Links[0].Section);
  Assert.AreEqual('', DockPlaceOfSection('DockSite5').Links[0].FormName,
    'a section is not a component name');
  // ASSERT THE SHAPE, not Known: with every field empty, Known is False
  // whether the early exit is there or not, so this is the only form of the
  // assertion that can go red when the guard is removed.
  Assert.AreEqual<Integer>(0, Length(DockPlaceOfSection('').Links),
    'nothing is nothing');

  // A LINK THAT NAMES NOTHING IS SKIPPED, not printed as an empty step. No
  // production path can produce such a place today - ReadDockTarget and
  // ReadFacts both drop an unknown link - so this pins the defence rather
  // than a live case.
  Assert.AreEqual('A (s)', PlaceOf([DockLinkOf('A', 'TA', '', 's'),
    Default(TDockLink)]).Describe, 'and never a dangling " in "');
end;

function MakeClient(AOrdinal: Integer; AIsSelf: Boolean;
  const ASection, AFormName, AFormClass: string): TDockClientRef;
begin
  Result := Default(TDockClientRef);
  Result.Ordinal := AOrdinal;
  Result.IsSelf := AIsSelf;
  Result.Section := ASection;
  Result.FormName := AFormName;
  Result.FormClass := AFormClass;
end;

procedure TReloadReportTests.APlaceWithoutItsTabIsNotARestore;
var
  F: TReloadFacts;
  Place: TDockPlace;
  Link: TDockLink;
  S: string;
begin
  // THE REPORTED STATE (2026-10-28): the window came back INSIDE its old dock
  // site and was no longer a TAB beside the Messages view. The old Boolean
  // could not carry that difference, so the box said nothing at all and the
  // user had to discover the loss by looking. rwoPlaceOnly IS that state.
  Link := DockLinkOf('TabDockHostForm', 'TTabDockHostForm', '', 'DockPanel');
  Link.Clients := [
    MakeClient(0, False, 'MessageView', 'MessageViewForm', 'TMessageViewForm'),
    MakeClient(1, True, 'RefactoringLightStatus', 'OTADockForm_1', 'TOTADockForm')];
  Place := PlaceOf([Link]);

  // The list answers both derived questions, and they are DERIVED rather than
  // stored on purpose: a stored second field could disagree with the list.
  Assert.AreEqual<Integer>(1, Length(Link.Siblings),
    'the siblings are the clients that are not ours');
  Assert.AreEqual('MessageView', Link.Siblings[0].Section,
    'and a sibling is named by its DeskSection, which is the lookup key');
  Assert.AreEqual(1, Link.OwnOrdinal, 'we were the second tab');

  F := Default(TReloadFacts);
  F.Reloaded := True;
  F.DockWasOpen := True;
  F.DockWasDocked := True;
  F.DockPlace := Place;

  // RESTORED says nothing, as before - the window is where it was.
  Assert.AreEqual('', ReloadReportText(F, rwoRestored),
    'a full restore is not worth a sentence');

  // PLACE ONLY says what is missing AND names the neighbour: "it is not a tab
  // any more" is only actionable with the window it should be a tab beside.
  S := ReloadReportText(F, rwoPlaceOnly);
  Assert.IsTrue(S.Contains('MessageView'),
    'the place-only sentence names the sibling: ' + S);
  Assert.IsTrue(S.Contains('tab') or S.Contains('TAB'),
    'and says what was lost: ' + S);
  Assert.IsFalse(S.Contains('could not be put back'),
    'it is NOT the sentence a failed restore gets: ' + S);

  // AND WITHOUT A RECORDED SIBLING IT SAYS NOTHING, which 2.9.1 changed on
  // purpose: rwoPlaceOnly means "not in the host it came from", and a window
  // that was ALONE in that host has no tab to miss - so a sentence there
  // would be about something we cannot know. The restraint belongs HERE and
  // nowhere else; 2.9.0 put it into the decision whether to keep LOOKING for
  // the tab, which is what cost the first install after it (see
  // OutcomeOfDockedLink).
  // A FRESH PLACE, not Place with its list emptied: TDockPlace.Links is a
  // dynamic array, so assigning the record SHARES it and clearing it through
  // one name clears it for the other. (That is what this assertion caught on
  // its first run, in the test rather than in the code.)
  F.DockPlace := PlaceOf([DockLinkOf('TabDockHostForm', 'TTabDockHostForm',
    '', 'DockPanel')]);
  Assert.AreEqual('', ReloadReportText(F, rwoPlaceOnly),
    'with no sibling recorded there is nothing we can honestly say');

  // THE MARKER CARRIES THE LIST - without that none of it survives the
  // reload. Written interleaved behind its own dockat= line, so the reader
  // attaches it to the last link it saw and no index has to agree.
  if TFile.Exists(TRestartHint.GetMarkerFilePath) then
    TFile.Delete(TRestartHint.GetMarkerFilePath);
  try
    TRestartHint.RecordUnload(nil, nil, True, True,
      TRect.Create(0, 0, 0, 0), Place);
    F := TRestartHint.ReadFacts;
    Assert.AreEqual<Integer>(1, Length(F.DockPlace.Links), 'one link went in');
    Assert.AreEqual<Integer>(2, Length(F.DockPlace.Links[0].Clients),
      'and both of its clients came back');
    Assert.AreEqual('MessageView', F.DockPlace.Links[0].Siblings[0].Section,
      'the sibling survives the round trip by its section');
    Assert.IsTrue(F.DockPlace.Links[0].Siblings[0].Known,
      'and a section IS what makes it usable as a key');
    Assert.AreEqual(1, F.DockPlace.Links[0].OwnOrdinal,
      'our own ordinal survives too, so "were we the first tab" is answerable');
  finally
    if TFile.Exists(TRestartHint.GetMarkerFilePath) then
      TFile.Delete(TRestartHint.GetMarkerFilePath);
  end;

  // AND THE SAVED DESKTOP NAMES THE SAME GROUP, which is what carries the
  // FIRST install of this version: the marker it reads was written by a
  // version that recorded no clients at all.
  Place := DesktopDockPlace(DeskLines([
    '[DockSite5]',
    'HostDockSite=DockBottomCenterPanel',
    'TabDockClients=MessageView,RefactoringLightStatus']),
    'RefactoringLightStatus');
  Assert.IsTrue(Place.Known, 'the file still names the host');
  Assert.AreEqual<Integer>(1, Length(Place.Innermost.Siblings),
    'and it names the tab group: the one client that is not us');
  Assert.AreEqual('MessageView', Place.Innermost.Siblings[0].Section,
    'by the section the lookup needs');
end;

procedure TReloadReportTests.TheFirstInstallMustNotStopLooking;
begin
  // THE DEFECT OF 2.9.0, reported after installing it: "anders als beim
  // letzten Versuch kam das Fenster nun daneben statt als Tab" - and then
  // "beim zweiten Versuch ging es", which is what named the cause. The first
  // install reads the marker the PREVIOUS version wrote, and that carries no
  // clients; 2.9.0 answered rwoRestored for a dock into the PANEL whenever
  // no sibling had been recorded, so Registration skipped the desktop
  // fallback - the one source that does know the tab group - and the box
  // stayed silent about it.
  //
  // THE RULE IS WHERE WE LANDED, and nothing about what we happen to know.
  Assert.AreEqual(rwoRestored, OutcomeOfDockedLink(0),
    'the innermost link IS the host that held the window: place and tabs');
  Assert.AreEqual(rwoPlaceOnly, OutcomeOfDockedLink(1),
    'any link further out is a dock site that cannot express a tab - so it ' +
    'is the place alone, whatever the marker carried');
  Assert.AreEqual(rwoPlaceOnly, OutcomeOfDockedLink(5),
    'and that holds however far out it is');

  // AND THE TWO WAITS ARE SEPARATE, because they wait for different things:
  // the five ticks were measured for the MODAL BOX (2.6.11 - a modal window
  // inside the install transaction wedges the IDE), while the catch-up opens
  // no modal window and is the only part the user can SEE. Waiting out the
  // box's margin is what he reported as "mehrere Sekunden, in denen visuell
  // nichts zu sehen ist".
  Assert.AreEqual(ssWait, SettleStage(1), 'tick 1 is still too early');
  Assert.AreEqual(ssCatchUp, SettleStage(CatchUpTicks),
    'the catch-up - and with it the window - runs as soon as the IDE is usable');
  Assert.AreEqual(ssWait, SettleStage(CatchUpTicks + 1),
    'and the ticks between the two stages do nothing');
  Assert.AreEqual(ssReport, SettleStage(DeferTicks),
    'the box keeps the distance that was measured for it');
  Assert.IsTrue(CatchUpTicks < DeferTicks,
    'the window must never wait for the box');
end;

initialization
  TDUnitX.RegisterTestFixture(TReloadReportTests);

end.
