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
  /// <summary>ONE LINK of a dock chain: a dock SITE control and the form that
  ///  owns it.
  ///
  ///  THE NAME IS THE KEY, and that is measured rather than chosen. 2.8.1
  ///  identified a host by its DeskSection, which is what the desktop file
  ///  uses - and a host the IDE creates when the user DRAGS a window to dock
  ///  it has none: on 2026-10-28 the status window sat in a TTabDockHostForm
  ///  named "TabDockHostForm" (ide_introspect), beside DockSite0..DockSite4
  ///  from the saved desktop, and our own row read "docked, host unknown".
  ///  The component NAME is there in both cases, and a form's name is unique
  ///  among its owner's components, so it is what finds the same live form
  ///  again after our package was unloaded and loaded into the same
  ///  process.</summary>
  /// <summary>One OTHER client of the host that held our window - its tab
  ///  SIBLING, and the thing a place alone cannot express.
  ///
  ///  WHY A PLACE IS NOT ENOUGH (user, 2026-10-28): after the reload the
  ///  window came back docked "rechts daneben" instead of as a TAB beside the
  ///  Messages view. The chain had done exactly what it promises - the
  ///  drag-created host was gone, so the restore used the next link, the
  ///  panel that host sat in - and that is the right PLACE. Being a tab is a
  ///  RELATION to another client, and a panel cannot express it: measured,
  ///  Vcl.Controls TDockTree.InsertControl opens its non-empty case with
  ///  "// Default to right-side docking / if InsertAt in [alClient, alNone]
  ///  then InsertAt := alRight", and designide's own TPercentageDockTree only
  ///  adds a dictionary entry before tail-calling it - no tab set anywhere.
  ///  So the sibling travels too, and it is the STRONGER half of the two: a
  ///  sibling is an IDE window that SURVIVES our unload, so unlike our own
  ///  host its identity still means something afterwards.</summary>
  TDockClientRef = record
    /// <summary>TDesktopForm.DeskSection - THE key. It is bound to the class
    ///  REGISTRATION rather than to streaming, and it is what the IDE itself
    ///  writes into TabDockClients / ActiveTabID, so it is the one name that
    ///  still means this window after our package was unloaded and loaded
    ///  again. '' means "not usable as a key" (see Known).</summary>
    Section: string;
    /// <summary>The component name: report and WEAK confirmation, never the
    ///  key - it is the DFM default, which is the very reason
    ///  TDockLink.Transient exists.</summary>
    FormName: string;
    /// <summary>Confirmation, and the "is it dockable at all" test at restore
    ///  time: only a TDockableForm can be spliced into a tab host.</summary>
    FormClass: string;
    /// <summary>Index in the host's DockClients, i.e. DOCK ORDER - which is
    ///  what says whether our own window was the first tab.</summary>
    Ordinal: Integer;
    /// <summary>Our own window. The list carries it on purpose: without it
    ///  "were we the first tab" could not be answered, and a restore has to
    ///  know which entries are the OTHERS.</summary>
    IsSelf: Boolean;
    /// <summary>Usable as a lookup key - see Section.</summary>
    function Known: Boolean;
  end;

  TDockLink = record
    /// <summary>The host FORM's component name - "DockSite5",
    ///  "TabDockHostForm", "AppBuilder".</summary>
    FormName: string;
    /// <summary>Its class name. A WEAK confirmation, and this doc said
    ///  more than it can keep until 2.8.2: it only ever rejects a name that
    ///  was recycled by a form of ANOTHER class, and a run-time dock host is
    ///  always a TTabDockHostForm. What rejects a recycled name of the same
    ///  class is Transient, below.</summary>
    FormClass: string;
    /// <summary>Its DeskSection, i.e. the name the desktop file uses
    ///  ("DockSite5", "Main Window"). EMPTY for a host created at run time,
    ///  which is why it cannot be the key - but it is the readable name, and
    ///  it is what the saved desktop can answer with.</summary>
    Section: string;
    /// <summary>The dock SITE control inside that form - "DockPanel" inside
    ///  a tab host, "DockBottomCenterPanel" inside the main window. The form
    ///  alone is not enough: a window docked straight into one of the main
    ///  window's own panels has the IDE's main form as its host, and there
    ///  the panel is the only thing that says WHERE.</summary>
    SiteName: string;
    /// <summary>This host's NAME is not an identity, so it must not be looked
    ///  up by name again.
    ///
    ///  MEASURED, and it took an adversarial pass over the live IDE plus the
    ///  shipped RTL to see it: "TabDockHostForm" is the design-time DFM name
    ///  that EVERY TTabDockHostForm carries - ide_introspect returns the same
    ///  17 components for TabDockHostForm and for DockSite4 - and DockSite0..N
    ///  are the ones the desktop loader renamed. System.Classes
    ///  TReader.ReadRootComponent assigns FindUniqueName(ReadStr), whose
    ///  counter restarts at 0 and whose uniqueness test asks only for LIVE
    ///  components, so the bare name is taken again by the next form the IDE
    ///  streams once it is free. And free it becomes: a drag-created host
    ///  dissolves when our window leaves it (UMFinalUndock), which is
    ///  precisely when this name stops meaning that host.
    ///  So a link that will dissolve is kept for the report but refused by
    ///  the lookup, and the restore falls through to the next link - the site
    ///  THAT host sat in, which is where its surviving client goes anyway.
    ///  A host with three or more clients stays, and then its name does
    ///  identify it.</summary>
    Transient: Boolean;
    /// <summary>The host's clients, ours among them (TDockClientRef.IsSelf) -
    ///  filled only for a link that IS a host, which is what carries the tab
    ///  RELATION across the reload. Empty for every other link and for a
    ///  marker older than 2.9.0, where the restore behaves as it did.</summary>
    Clients: TArray<TDockClientRef>;
    /// <summary>Enough to look this host up again.</summary>
    function Known: Boolean;
    /// <summary>For the status row and the report.</summary>
    function Describe: string;
    /// <summary>The siblings - every recorded client that is not our own
    ///  window, in dock order.</summary>
    function Siblings: TArray<TDockClientRef>;
    /// <summary>Our own window's place in the tab order, -1 when the list
    ///  does not say. DERIVED, not stored: a stored second field for it could
    ///  disagree with the list it describes.</summary>
    function OwnOrdinal: Integer;
  end;

  /// <summary>WHERE A WINDOW WAS DOCKED - as the CHAIN it really is, from the
  ///  host that held it outwards to the form that is not itself docked (the
  ///  IDE's main window).
  ///
  ///  WHY A CHAIN AND NOT ONE HOST: whether the host survives our unload
  ///  depends on how the USER made it, and the restore must not care which.
  ///  MEASURED in designide370.bpl, by a research pass and then a skeptic over
  ///  the same bytes: TTabDockHostForm.DoRemoveDockClient posts $506 to itself
  ///  as soon as at most one client is left, $506 is UMFinalUndock in the
  ///  class's own message table, and that hands the dock place to the
  ///  remaining client and calls TCustomForm.Release. BUT a one-byte flag
  ///  gates it, written in exactly one place - TDesktopForm.LoadDockStream -
  ///  so a host the DESKTOP LOADER created (DockSite0..N) is exempt and stays,
  ///  while one created by a DRAG dissolves. Live corroboration: DockSite0..4
  ///  are still there, and the DockSite5 of the saved desktop is not.
  ///  So the chain: the host when it is still there, and otherwise the site
  ///  THAT host was docked in - which is also where UMFinalUndock puts the
  ///  surviving client, so it is the same place on screen either way.
  ///  So the restore walks the chain and takes the DEEPEST link that still
  ///  exists; the status row then says which one it was, which is also how
  ///  the host's lifetime gets measured instead of assumed.</summary>
  TDockPlace = record
    /// <summary>Innermost (the host that held the window) first.</summary>
    Links: TArray<TDockLink>;
    function Known: Boolean;
    /// <summary>The host that held the window, or an empty link.</summary>
    function Innermost: TDockLink;
    function Describe: string;
  end;

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
    /// <summary>WHERE IT WAS DOCKED - read from the live form before the
    ///  unload (HostDockSite -> the form that owns that site -> ITS
    ///  HostDockSite -> ...), because the desktop file cannot say: it records
    ///  the membership with the SITE ("[DockSite5] TabDockClients=
    ///  MessageView,RefactoringLightStatus"), never with our own window.
    ///  Empty when the window was free-standing, or when nothing of the chain
    ///  could be read.</summary>
    DockPlace: TDockPlace;
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
      const ADockPlace: TDockPlace); static;
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
  ///  reload. The docked case is no longer a dead end: the dock sites belong
  ///  to the IDE, so the chain the window sat in can be looked up again and
  ///  asked to take it back - the host itself when it is still there, and
  ///  otherwise the site that host was docked in.</summary>
  TReloadWindowAction = (
    /// <summary>Nothing - it was not open, or this is a fresh IDE start.</summary>
    rwaNothing,
    /// <summary>It was free-standing: create it and put it at its bounds.</summary>
    rwaFloating,
    /// <summary>It was docked and the chain it sat in is known: create it and
    ///  hand it back to the deepest link that still exists.</summary>
    rwaRedock,
    /// <summary>It was docked but the chain is unknown (a marker written
    ///  before 2.8.0, one written by 2.8.0/2.8.1 for a host whose DeskSection
    ///  was empty, or a window whose host we could not read) - then there is
    ///  nothing to put it back into and the box has to say so.</summary>
    rwaDockedUnknownHost);

  /// <summary>WHAT THE CALLER ACHIEVED, and it has three values because the
  ///  reported defect lives in the gap between the last two: the window came
  ///  back INSIDE its old dock site and the restore reported success, because
  ///  the verification asks about containment - while the user could see that
  ///  it was no longer a TAB beside the Messages view but a pane next to it.
  ///  A Boolean cannot carry that difference, so it said nothing at all, and
  ///  "the user had to discover the loss by looking" is the one thing this
  ///  project does not accept of an answer.</summary>
  TReloadWindowOutcome = (
    /// <summary>Not put back at all - there was nothing left to put it into,
    ///  or the attempt failed.</summary>
    rwoNotDone,
    /// <summary>Back in its old PLACE, but not in its old RELATION: it is a
    ///  client of the right dock site and no longer a tab beside the window
    ///  it shared a tab group with.</summary>
    rwoPlaceOnly,
    /// <summary>Back the way it was - the place, and the tab relation when
    ///  there was one to restore.</summary>
    rwoRestored);

/// <summary>What to do with the status window after a reload. PURE - the
///  caller performs it, and reports back through ReloadReportText whether it
///  worked.</summary>
function ReloadWindowAction(const AFacts: TReloadFacts): TReloadWindowAction;

const
  /// <summary>Ticks to sit out before the BOX may open. THE DELAY IS ALSO A
  ///  GUARD: on a fresh IDE start the project-open notification has long
  ///  arrived after five seconds, so the catch-up finds nothing to do and
  ///  stays silent; on a reload it never arrives at all.</summary>
  DeferTicks = 10;
  /// <summary>And ticks to sit out before the CATCH-UP runs, which is a much
  ///  shorter wait on purpose (user, 2026-10-28: "das sind aber mehrere
  ///  Sekunden, in denen visuell nichts zu sehen ist").
  ///
  ///  THE TWO HAVE DIFFERENT REASONS, and 2.7.0 put them in one hook without
  ///  noticing: the five ticks were MEASURED for the modal box, because a
  ///  modal window inside the IDE's install transaction wedges it (2.6.11) -
  ///  that is the defect that whole round was about. The catch-up opens no
  ///  modal window; what it does is ordinary window work, and the two guards
  ///  below are exactly the test for "the transaction is over": ModalLevel,
  ///  and an ENABLED main window, which every modal path of the VCL disables.
  ///  The tick count does not even start until both are green, so this is a
  ///  buffer on top of a measured state rather than a guess about one - and
  ///  the user waits for the WINDOW, never for the box.</summary>
  CatchUpTicks = 2;
  DeferIntervalMs = 500;

type
  /// <summary>What the deferred tick should do at tick number N.</summary>
  TSettleStage = (
    /// <summary>Not yet - keep ticking.</summary>
    ssWait,
    /// <summary>Run the catch-up (project, headless buffers, the window) and
    ///  remember what it reports.</summary>
    ssCatchUp,
    /// <summary>Show what the catch-up could not do.</summary>
    ssReport);

/// <summary>WHEN THE DEFERRED TICK DOES WHAT. PURE, and separated because
///  the two halves wait for different reasons - see CatchUpTicks.</summary>
function SettleStage(ATicks: Integer): TSettleStage;

/// <summary>WHAT A DOCK INTO LINK NUMBER <c>AIndex</c> OF THE CHAIN
///  ACHIEVED. Trivial, named and tested because this is exactly where 2.9.0
///  was wrong, and the mistake cost the first install after it.
///
///  THE QUESTION IS WHERE WE LANDED, and nothing else. The innermost link IS
///  the host that held the window, so landing there restores the place and
///  every tab with it; any link further out is a dock site that cannot
///  express a tab at all (measured: Vcl.Controls' TDockTree turns alClient
///  into alRight in a non-empty tree, and designide's own dock tree has no
///  tab set), so it is the place ALONE.
///  WHAT 2.9.0 DID INSTEAD was to answer rwoRestored for an outer link
///  whenever no sibling had been RECORDED - meant as restraint ("do not
///  claim a lost tab when we know of no tab"), and it conflated two
///  questions: may I SAY the relation is missing, and may I STOP LOOKING for
///  it. The second one is never answered by what the marker happens to
///  carry: the saved desktop names the tab group itself. So the first
///  install after 2.9.0 - which reads the marker the PREVIOUS version wrote,
///  with no clients in it - docked into the panel, reported a restore, and
///  with that skipped the desktop fallback that would have rebuilt the tab.
///  The user saw it land beside its neighbour and the box said nothing.
///  Restraint belongs in the REPORT, which is the only place that knows
///  whether it can name what is missing.</summary>
function OutcomeOfDockedLink(AIndex: Integer): TReloadWindowOutcome;

/// <summary>WHAT A RELOAD HAS TO TELL THE USER. PURE, and deliberately apart
///  from the doing: the catch-up itself needs ToolsAPI and a live IDE, this
///  decision does not.
///  AOutcome is what the caller ACHIEVED - so the window is mentioned only
///  when something about it is still missing, which is the whole point: a
///  sentence that sends the user to the desktop menu must not appear when the
///  window is already sitting where it was, and silence must not cover a
///  window that came back in the wrong RELATION (user, 2026-10-28).
///  '' means there is nothing to report, which is the normal outcome.</summary>
function ReloadReportText(const AFacts: TReloadFacts;
  AOutcome: TReloadWindowOutcome): string;

/// <summary>A chain of ONE link, named by its section - which is all the
///  saved desktop file can say, and all a marker written by 2.8.x carries.
///  The lookup tries the component name first and the section second, so a
///  place built this way still finds a host the desktop created (there the
///  two strings are the same, "DockSite5").</summary>
function DockPlaceOfSection(const ASection: string): TDockPlace;

/// <summary>The value of one key inside one section of an ini-shaped file, or
///  '' - the .dsk reader needs exactly this and nothing more. Exported as a
///  test seam: it decides which section a HostDockSite= line is read from,
///  and reading it from the wrong one would be invisible.</summary>
function DeskValueIn(const ALines: TArray<string>; const ASection,
  AKey: string): string;

/// <summary>THE WHOLE CHAIN a desktop file can give, for the window with this
///  identifier: the dock host that lists it, and the site THAT host is docked
///  in. PURE text work on the lines of a .dsk.
///
///  WHY THE SECOND LINK MATTERS MORE THAN IT LOOKS: the file is the fallback
///  for exactly the case where the live hand-over could not say anything - and
///  the first install of any version is such a case, because it reads the
///  marker the version it replaced wrote. For 2.8.1 that marker says "docked,
///  host unknown" and carries no bounds, so without this the window would
///  come back in the middle of the screen once more. The host section it names
///  ("DockSite5") may well be gone from the session - the IDE numbers those
///  forms as it loads a desktop, and a host created by a drag releases itself
///  when its clients leave - while the panel behind it
///  ("HostDockSite=DockBottomCenterPanel") is a control of the IDE's own
///  window chrome and stays.
///  AND THE SECOND LINK NAMES NO OWNER, which is a correction of this
///  function's own first version: it wrote Section = 'Main Window' there,
///  reasoning that a main-window panel is what such a value names. MEASURED
///  on the live IDE (2026-10-28), DockBottomCenterPanel is a component of
///  EditWindow_0, a TEditWindow, and the main window has no child of that
///  name at all - so the lookup searched the wrong form and the fallback that
///  has to carry the first install of every version could never work. The
///  file records a site name without an owner, so that is what the link
///  carries, and TDockLink.Known accepts it: the lookup finds the form that
///  OWNS a control of that name, and refuses when two forms do.
///  What the sweep over all 17 desktop files on this machine really says is
///  only which NAMES appear (DockRightPanel, DockBottomCenterPanel,
///  DockBottomPanel, DockLeftPanel, RightTabDock, LeftDockTabSet) - not whose
///  they are.</summary>
function DesktopDockPlace(const ALines: TArray<string>;
  const AIdentifier: string): TDockPlace;

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

  /// <summary>Marker keys. A line the reader does not know is ignored, so an
  ///  older marker left by a previous version can never break a load.</summary>
  KeyProblem = 'problem=';
  KeyHeadless = 'headless=';
  KeyDock = 'dock=';
  /// <summary>One line per link of the dock chain, innermost first:
  ///  dockat=&lt;form name&gt;|&lt;form class&gt;|&lt;section&gt;|&lt;site name&gt;.
  ///  Its OWN key rather than more fields on dock=, so a marker written by
  ///  2.8.x still parses as exactly what it is - one link, section only.</summary>
  /// <summary>One line per recorded client of a host link:
  ///  dockclient=&lt;ordinal&gt;|&lt;is self&gt;|&lt;section&gt;|&lt;form name&gt;|&lt;form class&gt;.
  ///  ITS OWN KEY, because a variable-length list on a fixed '|' field line
  ///  is what breaks readers - and written INTERLEAVED, directly after the
  ///  dockat= line it belongs to, so the reader attaches it to the last link
  ///  it saw and no index has to agree with anything. The prefixes are
  ///  disjoint ("dockclient=" does not start with "dock="), so an older
  ///  reader skips these lines and loses only the relation.</summary>
  KeyDockClient = 'dockclient=';
  KeyDockAt = 'dockat=';
  DockChainSep = '|';

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
    /// <summary>What the catch-up stage reported, shown by the report
    ///  stage - the two are separate ticks now (see CatchUpTicks).</summary>
    FCatchUpText: string;
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
  case SettleStage(FTicks) of
    ssWait:
      Exit;
    ssCatchUp:
      begin
        // THE CATCH-UP, as soon as the IDE is usable - the window is the one
        // thing the user SEES, and it used to wait out the box's own margin.
        // The timer keeps running: the box is a separate stage, and its text
        // is what this stage produces.
        FCatchUpText := '';
        if not TRestartHint.FPinned and Assigned(TRestartHint.SettleHook) then
          try
            FCatchUpText := TRestartHint.SettleHook();
          except
            // The catch-up may not cost the user an exception box of its own;
            // a failure here means the same as "nothing to report".
            FCatchUpText := '';
          end;
        Exit;
      end;
  end;
  // BEFORE anything that pumps: a modal box's own message loop would deliver
  // the next tick and stack a second box on top of this one.
  FTimer.Enabled := False;

  Text := FCatchUpText;
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
  else if (Text = '') and Assigned(TRestartHint.SettleHook) then
    // THE CATCH-UP DID NOT RUN - the guards were red at its own tick and
    // green at this one, which is possible when a modal dialog of the IDE
    // sat exactly across tick 2. Then this stage does both.
    try
      Text := TRestartHint.SettleHook();
    except
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
  if AFacts.DockPlace.Known then
    Result := rwaRedock
  else
    Result := rwaDockedUnknownHost;
end;

function TDockClientRef.Known: Boolean;
begin
  // THE SECTION AND NOTHING ELSE, and that is the opposite polarity from
  // TDockLink.Known on purpose. A link may be looked up by its component
  // name because a host that is still there answers for it; a SIBLING is
  // looked up in order to be DOCKED INTO, so a wrong answer would splice our
  // window into a stranger's tab group and report a restore. DeskSection is
  // bound to the class registration rather than to streaming, which is why
  // it survives where a component name does not (see TDockLink.Transient).
  Result := Section <> '';
end;

function TDockLink.Siblings: TArray<TDockClientRef>;
begin
  Result := [];
  for var C in Clients do
    if not C.IsSelf then
      Result := Result + [C];
end;

function TDockLink.OwnOrdinal: Integer;
begin
  Result := -1;
  for var C in Clients do
    if C.IsSelf then
      Exit(C.Ordinal);
end;

function TDockLink.Known: Boolean;
begin
  // ANY of the three answers, and the third one is not a convenience: the
  // SAVED DESKTOP records a dock place as a site name with NO owner at all
  // ("HostDockSite=DockBottomCenterPanel"), and MEASURED on the live IDE
  // (2026-10-28) that panel belongs to EditWindow_0, a TEditWindow - not to
  // the main window, which has no child of that name. So a link that carries
  // only a site name is exactly what the file can give, and the lookup
  // answers it by finding the form that OWNS a control of that name. Naming
  // an owner we did not read would be a guess, and it was: 2.8.2 first wrote
  // "Main Window" here, which made the desktop fallback look for that panel
  // in the wrong form and never find it.
  Result := (FormName <> '') or (Section <> '') or (SiteName <> '');
end;

function TDockLink.Describe: string;
begin
  // The SECTION is the readable name where there is one ("Main Window"), and
  // for a host the desktop file created the two are the same string anyway
  // ("DockSite5"). A host created by a drag has no section, and then its
  // component name is what there is - which is exactly the case that read
  // "host unknown" before.
  if Section <> '' then
    Result := Section
  else if FormName <> '' then
    Result := FormName
  else
    // A SITE NAME WITH NO OWNER is what the saved desktop gives, and it is
    // the whole description there is: naming an owner would invent one.
    Exit(SiteName);
  // THE HOST FORM CAN BE ITS OWN DOCK SITE, and then the name is already
  // there: TControl.Dock sets FHostDockSite to whatever it docks into, and for
  // a tab host that is the form (TBaseDockHostForm.GetDockSiteControl answers
  // Self). Printing it twice would read as two different things.
  if (SiteName <> '') and not SameText(SiteName, FormName) then
    Result := Result + ' (' + SiteName + ')';
end;

function TDockPlace.Known: Boolean;
begin
  Result := (Length(Links) > 0) and Links[0].Known;
end;

function TDockPlace.Innermost: TDockLink;
begin
  if Length(Links) > 0 then
    Result := Links[0]
  else
    Result := Default(TDockLink);
end;

function TDockPlace.Describe: string;
begin
  // "TabDockHostForm (DockPanel) in Main Window (DockBottomCenterPanel)" -
  // the host, and then what it is itself docked in, which is where the window
  // goes when the host did not survive. One row, and it names both outcomes
  // before either of them happens.
  Result := '';
  for var L in Links do
  begin
    if not L.Known then Continue;
    if Result <> '' then
      Result := Result + ' in ';
    Result := Result + L.Describe;
  end;
end;

function DeskValueIn(const ALines: TArray<string>; const ASection,
  AKey: string): string;
var
  Section: string;
begin
  Result := '';
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
    if not SameText(Section, ASection) then Continue;
    var Eq := Pos('=', L);
    if Eq <= 1 then Continue;
    if SameText(Trim(Copy(L, 1, Eq - 1)), AKey) then
      Exit(Trim(Copy(L, Eq + 1, MaxInt)));
  end;
end;

function DesktopDockPlace(const ALines: TArray<string>;
  const AIdentifier: string): TDockPlace;
var
  HostSection, Panel, Clients: string;
  Link: TDockLink;
begin
  Result := Default(TDockPlace);
  HostSection := FindDockHostSection(ALines, AIdentifier);
  if HostSection = '' then Exit;
  Result := DockPlaceOfSection(HostSection);
  // The panel that host sits in - the link that is still there when the host
  // itself is not. ITS OWNER IS NOT RECORDED and is not guessed: the file
  // says "HostDockSite=DockBottomCenterPanel" and nothing else, and measured
  // on the live IDE that panel belongs to EditWindow_0 (TEditWindow), while
  // the main window has no child of that name. 2.8.2 first wrote
  // Section = 'Main Window' here from a plausible reading of the name, which
  // made the lookup search the wrong form - the one fallback that has to
  // carry the first install of every version. A site name alone is a link the
  // lookup can resolve; a wrong owner is not.
  // AND THE TAB GROUP, which this file states OUTRIGHT - it is the very key
  // FindDockHostSection searched by: "TabDockClients=MessageView,
  // RefactoringLightStatus". Those are DeskSections, i.e. exactly the key a
  // sibling lookup needs, so the saved desktop can restore the RELATION and
  // not only the place.
  // THAT IS WHAT MAKES THE FIRST INSTALL OF THIS VERSION WORK: it reads the
  // marker the PREVIOUS version wrote, which carries no clients at all - the
  // gap that cost 2.8.1 and 2.8.2 each a round ("expect it once more, exactly
  // as predicted for its own first install").
  Clients := DeskValueIn(ALines, HostSection, 'TabDockClients');
  if Clients <> '' then
    for var Name in SplitString(Clients, ',') do
    begin
      var N := Trim(Name);
      if N = '' then Continue;
      var Ref := Default(TDockClientRef);
      Ref.Section := N;
      Ref.Ordinal := Length(Result.Links[0].Clients);
      Ref.IsSelf := SameText(N, AIdentifier);
      Result.Links[0].Clients := Result.Links[0].Clients + [Ref];
    end;

  Panel := DeskValueIn(ALines, HostSection, 'HostDockSite');
  if Panel = '' then Exit;
  Link := Default(TDockLink);
  Link.SiteName := Panel;
  Result.Links := Result.Links + [Link];
end;

function DockPlaceOfSection(const ASection: string): TDockPlace;
var
  Link: TDockLink;
begin
  Result := Default(TDockPlace);
  if ASection = '' then Exit;
  Link := Default(TDockLink);
  Link.Section := ASection;
  Result.Links := [Link];
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

function SettleStage(ATicks: Integer): TSettleStage;
begin
  if ATicks = CatchUpTicks then
    Result := ssCatchUp
  else if ATicks >= DeferTicks then
    Result := ssReport
  else
    Result := ssWait;
end;

function OutcomeOfDockedLink(AIndex: Integer): TReloadWindowOutcome;
begin
  if AIndex = 0 then
    Result := rwoRestored
  else
    Result := rwoPlaceOnly;
end;

function ReloadReportText(const AFacts: TReloadFacts;
  AOutcome: TReloadWindowOutcome): string;
var
  Lines: TArray<string>;
  Sibling: string;
begin
  Result := '';
  // A fresh IDE start has nothing to catch up and nothing to report: every
  // notification this package listens to is still ahead of it.
  if not AFacts.Reloaded then Exit;

  // THE STATUS WINDOW. Both cases are DONE by the caller now - a
  // free-standing one goes back to its bounds, a docked one into the deepest
  // link of its dock chain that still exists: the host that held it while
  // that host is there, and otherwise the site THAT host was docked in. (Not
  // "by its DeskSection", which is what this comment said until 2.8.2 and is
  // the defect that round was about: a host the user created by dragging has
  // no DeskSection at all.) Only a restore that did NOT work is worth a
  // sentence, and then it names what the user can do.
  // AND THE RELATION, which is the half 2.8.2 could not report: a window
  // that came back INSIDE its old dock site but no longer as a tab beside
  // its old neighbour was reported as a restore, because the verification
  // asks about containment. The user saw it and we said nothing.
  if (AOutcome = rwoPlaceOnly) and AFacts.DockWasOpen then
  begin
    Sibling := '';
    for var C in AFacts.DockPlace.Links do
      for var S in C.Siblings do
        if (Sibling = '') and (S.Section <> '') then
          Sibling := S.Section;
    // AND WITHOUT A NAME IT SAYS NOTHING, which is where the restraint
    // belongs: rwoPlaceOnly means "not in the host it came from", and for a
    // window that was ALONE in that host there is no tab to miss. Only a
    // recorded sibling proves there was a group, so only then is there
    // something to report - and then it can be named, which is what makes
    // the sentence actionable.
    if Sibling <> '' then
      Lines := Lines + [Format('The status window is back in its dock site, ' +
        'but no longer as a TAB beside %s - it sits next to it instead. ' +
        'Dragging it onto that tab bar once puts it back, and the IDE ' +
        'remembers it with your desktop.', [Sibling])];
  end;

  if AFacts.DockWasOpen and (AOutcome = rwoNotDone) then
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
  Stamp, L, OldSection, OldSite: string;
begin
  PruneDeadMarkers;
  Stamp := ProcessStamp;
  FFacts := Default(TReloadFacts);
  OldSection := '';
  OldSite := '';

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
      else if StartsText(KeyDockAt, L) then
      begin
        // dockat=<form name>|<form class>|<section>|<site name>|<transient>,
        // one line per link, innermost first - and they arrive in the order
        // they were written, so the chain needs no index. The fifth field
        // arrived with 2.8.2; a line without it is a link whose name was
        // trusted, which is what every version before it did.
        var Q := SplitString(Copy(L, Length(KeyDockAt) + 1, MaxInt), DockChainSep);
        var Link := Default(TDockLink);
        if Length(Q) >= 1 then Link.FormName := Trim(Q[0]);
        if Length(Q) >= 2 then Link.FormClass := Trim(Q[1]);
        if Length(Q) >= 3 then Link.Section := Trim(Q[2]);
        if Length(Q) >= 4 then Link.SiteName := Trim(Q[3]);
        if Length(Q) >= 5 then Link.Transient := Trim(Q[4]) = '-1';
        if Link.Known then
          FFacts.DockPlace.Links := FFacts.DockPlace.Links + [Link];
      end
      else if StartsText(KeyDockClient, L) then
      begin
        // dockclient=<ordinal>|<is self>|<section>|<form name>|<form class>,
        // attached to the LAST link read - the writer puts these lines
        // directly after the dockat= line they belong to, so no index has to
        // agree with anything, and a line that arrives before any link is
        // dropped rather than inventing one.
        if Length(FFacts.DockPlace.Links) > 0 then
        begin
          var Q := SplitString(Copy(L, Length(KeyDockClient) + 1, MaxInt),
            DockChainSep);
          var Ref := Default(TDockClientRef);
          Ref.Ordinal := -1;
          if Length(Q) >= 1 then Ref.Ordinal := StrToIntDef(Trim(Q[0]), -1);
          if Length(Q) >= 2 then Ref.IsSelf := Trim(Q[1]) = '-1';
          if Length(Q) >= 3 then Ref.Section := Trim(Q[2]);
          if Length(Q) >= 4 then Ref.FormName := Trim(Q[3]);
          if Length(Q) >= 5 then Ref.FormClass := Trim(Q[4]);
          if Ref.Known or (Ref.FormName <> '') then
          begin
            var Idx := High(FFacts.DockPlace.Links);
            FFacts.DockPlace.Links[Idx].Clients :=
              FFacts.DockPlace.Links[Idx].Clients + [Ref];
          end;
        end;
      end
      else if StartsText(KeyDock, L) then
      begin
        // dock=<open>,<docked>,<left>,<top>,<right>,<bottom>
        //      [,<host section>[,<site name>]]
        // The 7th field arrived with 2.8.0, the 8th with 2.8.1; a marker
        // without them is a window whose host, or whose dock site, we do not
        // know - which is what the old behaviour was, so an older marker can
        // never be a parse error. That is not a theoretical case either: an
        // install over a running older version reads exactly such a marker,
        // every time.
        var P := SplitString(Copy(L, Length(KeyDock) + 1, MaxInt), ',');
        if Length(P) >= 6 then
        begin
          FFacts.DockWasOpen := P[0] = '1';
          FFacts.DockWasDocked := P[1] = '1';
          FFacts.DockBounds := TRect.Create(StrToIntDef(P[2], 0),
            StrToIntDef(P[3], 0), StrToIntDef(P[4], 0), StrToIntDef(P[5], 0));
          if Length(P) >= 7 then
            OldSection := Trim(P[6]);
          if Length(P) >= 8 then
            OldSite := Trim(P[7]);
        end;
      end;
    end;

    // A MARKER FROM 2.8.0 / 2.8.1 carries no chain, only the section and site
    // of the one host it knew - and that is the shape the FIRST load after
    // this change meets, every time, because the install reads what the
    // version it replaced wrote. One link, found by section.
    if (Length(FFacts.DockPlace.Links) = 0) and
       ((OldSection <> '') or (OldSite <> '')) then
    begin
      var Link := Default(TDockLink);
      Link.Section := OldSection;
      Link.SiteName := OldSite;
      if Link.Known then
        FFacts.DockPlace.Links := [Link];
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

function NoSeparators(const S: string): string;
begin
  // The marker lines are separated lists, so a name that carries a separator
  // would move every field behind it - and a mis-parsed list hands the next
  // instance a name that points at ANOTHER window. A component name cannot
  // contain either character, but the host section is an ini section the IDE
  // writes, so this is measured rather than assumed: a stripped name is
  // simply not found again, which is the honest failure.
  Result := StringReplace(S, ',', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, DockChainSep, ' ', [rfReplaceAll]);
end;

class procedure TRestartHint.RecordUnload(const AProblems, AHeadless: TArray<string>;
  ADockOpen, ADockDocked: Boolean; const ADockBounds: TRect;
  const ADockPlace: TDockPlace);
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
      begin
        // THE 7TH AND 8TH FIELD ARE FOR SOMEBODY ELSE, and nothing in this
        // version reads them: every link that fills them has its own dockat=
        // line, which is what ReadFacts prefers. They are what an OLDER build
        // loaded into this same process would find - a downgrade install -
        // because it skips every dockat= line (that key does not start with
        // "dock="). Writing them costs two strings and keeps that reader
        // working; the test asserts they are still there.
        SL.Add(Format('%s%d,%d,%d,%d,%d,%d,%s,%s', [KeyDock, 1, Ord(ADockDocked),
          ADockBounds.Left, ADockBounds.Top, ADockBounds.Right, ADockBounds.Bottom,
          NoSeparators(ADockPlace.Innermost.Section),
          NoSeparators(ADockPlace.Innermost.SiteName)]));
        // AND THE WHOLE CHAIN, which is what a restore really needs: the
        // host that held the window may be gone by then - a drag-created one
        // dissolves when our window leaves it - and the next link is the
        // site THAT host was docked in.
        for var Link in ADockPlace.Links do
          if Link.Known then
          begin
            SL.Add(KeyDockAt + string.Join(DockChainSep,
              [NoSeparators(Link.FormName), NoSeparators(Link.FormClass),
               NoSeparators(Link.Section), NoSeparators(Link.SiteName),
               BoolToStr(Link.Transient, False)]));
            // AND ITS CLIENTS, right behind it - the tab RELATION, which the
            // place alone cannot carry (see TDockClientRef). Interleaved on
            // purpose: the reader attaches them to the last link it saw, so
            // the two keys cannot disagree about which host they describe.
            for var C in Link.Clients do
              if C.Known or (C.FormName <> '') then
                SL.Add(KeyDockClient + string.Join(DockChainSep,
                  [IntToStr(C.Ordinal), BoolToStr(C.IsSelf, False),
                   NoSeparators(C.Section), NoSeparators(C.FormName),
                   NoSeparators(C.FormClass)]));
          end;
      end;
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
