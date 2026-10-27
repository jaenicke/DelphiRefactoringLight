# Changelog

Every released version of Delphi Refactoring Light, newest first. Each entry
links to the full notes of that release. Versions before 1.7.0 were
development builds and were never released.

---

## 2.6.10 - 2026-10-26

### Extract Method refused two things the compiler accepts

Issue #51 by Ian Branch, both of them in the parameter-type check 2.6.3
added, and both measured with dcc32 37.0 before anything was changed.

A parameter the block uses but that comes from the **enclosing routine's
own header** was refused when its type is written out in place - the
reported case is `const A: array of Integer`. That parameter already *is*
an open array, and forwarding its declaration verbatim compiles and means
the same thing. What 2.6.3 was built for stays refused: a **local**
dynamic array promoted to a parameter really does lose `SetLength`
(E2008), which the same probe confirms.

And a type name is a name by the language's rule, not by an ASCII set:
`TKündigung`, `Verträge.TÄnderung` and the escaped `&Type` are legal type
names, compile as parameter types, and were read as "written out in place".

Both refused rather than wrote wrong code, so they cost an extraction, not
a unit.

---

## 2.6.9 - 2026-10-26

### A build that failed does not get its artefacts installed

Found while reviewing the installer helper before porting it upstream, by
three of four review dimensions independently.

Since 1.30.1 the installer copies a file the build it just ran PRODUCED -
that is how the MCP bridge reaches its install folder. The build's result
was discarded, though, so a compile error left the previous build's exe
being copied over the target and logged as "Build artefacts installed: 1",
with the entry's post event then verifying the old binary. That is exactly
the stale-binary class of bug the staged copy was added to prevent, and it
had a nine-day and a four-version precedent in this plugin's own history.

The copy now happens only when the build succeeded, and when it did not,
the log says how many artefacts were NOT installed and that the target
keeps the file it had. A copy pass in which every file failed no longer
writes a green success line either.

### The name an install moves a running file aside under

It promised to be unique and was not: it came from `Random` without
`Randomize`, so every process started at the same first value and every
install walked the same collisions - and after twenty tries the loop left
with a name that exists, where the rename then failed with a message about
a destination nobody had asked for. It is a timestamp now, and a name that
is still taken is refused with its reason.

The shape has a consumer: since 1.30.1 the installer writes the bridge's
aside copy, and `CanonicalBridgeExePath` reads it back to tell "the install
replaced the exe under this running session" from "the install never
finished" - two situations that need opposite advice. A test now holds
those two ends together.

---

## 2.6.8 - 2026-10-26

### A parameter is declared at two POSITIONS, and only one pair of them

Issue #50, an audit of the parameter pairing 2.6.1 had just shipped. Five
points, and the first four are the same mistake seen from four sides: a
parameter was looked for by NAME and on LINES, where it really stands at two
positions in two headers that have to be the same routine.

What that produced, each case measured against the compiler:

- two overloads sharing a parameter name paired crosswise, so renaming
  `Init(const Title: string)` renamed the `Title` of `Init(const Title:
  string; X: Integer)` as well - and from a declaration the FIRST
  implementation of that name was taken, whichever it was;
- a nested routine paired with an interface declaration of the same name,
  i.e. with a routine in another type entirely;
- where the pairing could not be established the rename went ahead with one
  header, which does not compile (E2037, plus an undeclared identifier for
  every use in the body). Six shapes reached that: a class declared in the
  implementation section, a nested type, a `{ }` comment inside the
  parameter list, a header longer than twelve lines, a `const [ref]`
  parameter, and a `class operator`;
- and the targets were whole LINES, so a parameter named like its own
  routine renamed the routine name too, a type or a default value spelled
  like the parameter counted as the parameter, and the owner-type scan -
  which exists to reach a method's overrides - ran with the parameter's name
  and could rename a same-named parameter of a descendant.

A parameter rename is therefore planned as a whole now: both headers are
read with their real parameter lists, they pair only when those lists agree
in names, modifiers and types, exactly one candidate has to fit, and every
renamed place is a position inside a parameter list - never a line. A
routine that is declared once by design (an interface or dispinterface
member, `abstract`, `external`, a nested routine) is recognised as such
instead of being paired with something; anything else that cannot be
established REFUSES the rename and says why.

### An empty answer lost the one thing it exists to say

The fifth point, and it is older: 2.1.1 let a tool phrase the empty answer
in its own words, and four tools then passed their DESCRIPTION there. So
`lsp_implementation`, `lsp_document_symbols`, `lsp_signature_help` and
`lsp_completion` never said that nothing came back, and the paragraph about
an unresolved symbol followed a sentence about ranges.

The two are separate now - what an empty answer MEANS for this tool, and
what the answer is - and the symbol diagnosis only goes to the tools that
resolve a position. `lsp_implementation` no longer points at
`find_original_symbol` either: having no implementation is its normal
answer.

---

## 2.6.7 - 2026-10-26

### "Remove unused private" grew the unit by a blank line

Found by running the fix of 2.6.4 on a real unit: the member and both its
documentation blocks went, the two neighbours kept exactly one blank line
between them - and the unit had one line more than before, at its very end.

The removal joined its lines through a `TStringList`, which appends a line
break of its own; the empty element that stands for the file's own final
break then became a new blank line. Every removal added one. The same join
also used the system line break, so the whole-file path would have rewritten
an LF-only unit as CRLF - exactly what the line-ending check exists to
prevent.

That rule already existed twice in this plugin (the method move has it, and
a copy sat forty lines below the defect in the same unit). It has one home
now, next to the split it inverts, and the two copies use it.

---

## 2.6.6 - 2026-10-26

### Extract Method: a member is not something the block needs

Two reports from the same method: the extracted routine was given a
parameter for `AWorkSheet.Cell` and none for the `LineStr` it really uses,
and it declared a local `AsFloat: string` that nothing reads - `AsFloat`
being a `Double` property of a cell.

Both start in the same place: the block's identifiers were collected with
every name the text holds, the MEMBER after a dot included. For
`AWorkSheet.Cell['A' + LineStr].AsString` that is `Cell` and `AsString` as
well - names that belong to a type and that the new routine can never be
given. Each one costs a DelphiLSP round trip, and wherever the answer does
not place it safely, such a name ends up as a parameter or as a local of
the new routine. A member is skipped now, as the scan that looks for uses
OUTSIDE the block has always done. A range (`cA..cZ`) is not a member
access, and a name that is a member here and a variable three characters
on is still found - at the right place.

### ...and the selection's columns are the ones in the file

`IOTAEditBlock` reports DISPLAY columns - a tab counts as its width on
screen - while everything that works with a selection uses them as
positions in the text. On a tab-indented file that shifts every position
of the selection's FIRST line, so Extract Method asked DelphiLSP about
whatever stood there instead. The conversion the completion popup already
uses is applied to a selection now; on a line without a tab nothing changes at
all.

---

## 2.6.5 - 2026-10-21

### The install on a machine where Delphi belongs to another account

Where Delphi was installed by a different user, its installer writes that
user's Documents folder into `rsvars.bat` and the registry as the common
directory - and msbuild then dies with `MSB3191` on folders the build never
writes to. 1.25.3 moved the outputs (BPL, DCP, `.hpp`) away from there,
which was not enough: `CodeGear.Delphi.Targets` also builds its
"create these directories" list from `$(DelphiHPPOutputPath)` and a `Dcp`
path of its own, which no `DCC_*` override reaches. Measured on both
projects: two `MSB3191` survived.

The ROOT moves with the outputs now - as an msbuild property, so the
environment keeps Delphi's own value and the `bds.exe` fallback still finds
the IDE's packages. For the helper that is built first this is
unconditional (it produces an exe and writes none of those folders); for
the entries of the install it happens only where this account really cannot
write the common directory, i.e. exactly where the output root moved too.
The artefacts still land where the explicit outputs say - measured.

---

## 2.6.4 - 2026-10-21

### "Remove unused private ..." takes the XMLDOC with it

Suggested in the forum: the quick fix removed the declaration and the
implementation of an unused private member and left the `///` comments
above them behind, documenting nothing. A `///` block describes what
follows it, so it goes with it - above the declaration and above the body.
Safe delete has always done this; it is one rule in one place now, and
both ask it.

---

## 2.6.3 - 2026-10-21

### Extract Method keeps a type that is written out in place

Reported in the forum: a local `Handles: array[0..0] of THandle` arrived in
the extracted method as `Handles: array;`, which compiles nowhere. The type
was read as the first identifier after the colon, so everything from the
`[` or the ` of ` on was dropped - the same for `set of Byte`,
`string[20]`, `^Integer` and `procedure(X: Integer) of object`. Both paths
that determine a type had it: the DelphiLSP hover and the fallback that
reads the declaration line when there is no hover.

A local that MOVES into the new method keeps its type verbatim now. A
variable the block SHARES with the code around it would become a
parameter - and there such a type is not allowed at all (measured with
dcc32: E2029). That case is refused with the reason and the way out
(give the type a name) instead of generating something that does not
compile. `array of T` is refused too: it is accepted in a parameter list,
but as an OPEN array, which is no longer the dynamic array it came from -
`SetLength` on it does not compile.

---

## 2.6.2 - 2026-10-21

### A rename on a parameter stops waiting for an answer that cannot come

Starting a rename on a parameter in a declaration works since 2.6.1 - after
82 seconds of nothing. Before the scan, the rename asks BOTH DelphiLSP
sessions for the symbol's declaration, to find out which of the two can
answer at all; a freshly started agent session answers everything with null
and nothing else can tell them apart. For a parameter neither one ever
answers, because a parameter is declared at that very position and the
server says nothing at a declaration - measured on a warm session that
resolves a use of the same parameter two lines further down. So that probe
could only ever end in its own timeout, and the scan then continues exactly
as it does when the probe is skipped. It is skipped now.

Only for a parameter: a caret on a method declaration IS answered when that
method overrides one, and that answer - the ancestor - is what makes the
rename reach the whole family.

---

## 2.6.1 - 2026-10-20

### Renaming a parameter changes both of its declarations (issue #49)

Reported by stepand76: put the cursor on a parameter in the method body,
rename it, and the implementation header and the body followed - the
declaration in the class did not. Measured with dcc32 37.0 before anything
was written: Delphi requires a routine's parameter names to be **identical**
in its declaration and in its implementation header, for a class method and
for a free routine alike. So the result does not compile - E2037
"Declaration of 'X' differs from previous declaration", plus an undeclared
identifier for every use in the body.

Why it happened: a parameter is declared twice, and DelphiLSP answers
nothing at either of the two. The rename verifies every occurrence against
the positions it knows the symbol by, so the second declaration answered
"no idea" and was listed as unverified and left alone - the safe direction
for an unknown occurrence, and exactly wrong for this one.

The pairing now comes from the sources, where it is decidable: the same
parameter, at the same place in the list, in the routine's other header -
a wrapped parameter list included, and for an overloaded routine only the
declaration that really carries that parameter at that place. Both halves
of the rename reach it, so it no longer matters whether the cursor is in
the body, on the implementation header or on the declaration.

And starting the rename ON the parameter in the class declaration works at
all now: DelphiLSP answers nothing there, the caret line was not recognised
as declaring anything, and the rename refused after waiting for an answer
that cannot come.

---

## 2.6.0 - 2026-10-20

### One menu entry for a move, and a dialog that settles the target

"Move to unit..." and "Move to new unit..." were two editor-menu entries
for the same refactoring, and the second one had no dialog at all: it asked
for a name in a text box and then either moved or came back with a message
box. They are one entry now, and its dialog carries the decision:

- **into an existing unit** - the project's other units, filterable, with
  the preview of the planned edits it always had;
- **into a new unit** - its name, suggested from the symbol
  (`TCustomerList` becomes `CustomerList`) and checked while it is typed: a
  name that cannot be a unit name, or one whose file is already there, says
  so instead of enabling **Move**.

What the new unit's own refusal needs - deciding whether the moved
declaration can live without the unit it came from, which would otherwise
be a circular unit reference - asks DelphiLSP, so it runs when **Move** is
pressed rather than on every keystroke, and a refusal keeps the dialog open
with its reason instead of closing it and showing a message box.

The shortcut (`Ctrl+Shift+M` by default) opens the same dialog, so both
targets are reachable from the keyboard. Nothing changed for the MCP tools:
`move_to_unit` and `move_to_new_unit` each name their target explicitly and
both stay.

### `lsp_restart` starts a session again - and says so only when it did

`lsp_restart` ends the DelphiLSP session and starts a new one, because the
`lsp_*` tools never cold-start one by themselves. The starting half went
through the path the IDE's own project-open notifier uses - and that one
debounces re-opens of the **same** project and honours the "prewarm on
project open" setting. So for the project that is already open it did
nothing at all, while the answer still reported `restarted: true`, because
the call had not raised. What was left was a plugin without a session:
every `lsp_*` tool and `find_references` refused until the IDE itself was
restarted.

The restart has its own path now: no debounce, no prewarm setting, and
`restarted` is true only when a start really began. A start that was
already running is not raced - the request is parked and runs when that one
has finished, and the answer says `queued`. Where nothing started at all,
`reason` names it: `start_again=false`, no project, no `.delphilsp.json`,
no source files, or the plugin is shutting down.

---

## 2.5.2 - 2026-10-20

### The answer stops offering a fix it does not have

With the DCU output complete the check reports no "LSP reads a DCU"
finding at all - and its answer still ended with "a finding can be fixed
from here: apply=true with its id moves the .dcu aside". That sentence was
a constant, and since 2.5.0 it is wrong in the other direction as well:
only a `.dcu` that is older than its source **and** whose source the
compiler could use instead has that cure; for every other finding of that
kind the cure is a build, which is what the rows say and what `apply`
answers when they are named.

The tail is derived from the findings now - it says how many carry that
cure, which finding has it, and that the others name their own; and with
none of them it says nothing about a fix.

---

## 2.5.1 - 2026-10-20

### What the first live run of the new DCU check reported

The 2.5.0 rules held on the project that was in the reported state - 108
"move this .dcu aside" findings became 6, and after the build the finding
asks for, 1. That last one was a defect: the check counted the project's
own `.dpk` as a unit that should have a `.dcu`, so it reported the output
directory as incomplete by exactly one, for a file that can never satisfy
it - while the status row, which counts the configuration's own unit list,
said the directory matches.

Measured over three projects of this repository (a package and two
programs, four output directories): **no `.dcu` is ever named after the
`.dpk` or the `.dpr`** - that file produces the `.bpl` or the `.exe`. Only
a `.pas` can be missing one now, which is also what the check's own source
map has always done. And the missing unit is named as its source spells it,
not as the lower-cased lookup key the answer printed.

---

## 2.5.0 - 2026-10-20

### What really decides a cross-unit "go to definition"

Eight states on a live session, one variable changed per step, the same
probe every time - go to definition on a symbol declared in another unit of
the same project:

| the project's DCU output | the declaring unit came from | answer |
|---|---|---|
| complete | that directory | the source position |
| complete, built WITHOUT debug information | that directory | the source position |
| that one `.dcu` moved to a search-path directory | there | the source position |
| ...replaced by the Release build's `.dcu` | a Release artefact | the source position |
| ...only the installed package's `.dcp` left | the `.dcp` | the source position |
| one four-day-old `.dcu` among fresh ones | that directory | the source position |
| **one unit's `.dcu` missing from it** | another build | **nothing** |
| ...and no other provider at all | nowhere | **nothing** |

So a unit the LSP loads from a `.dcu` *does* come back with its source
position - which makes two things from 2.4.x wrong, and this release
corrects them:

- **The claim.** "A unit it loads from a `.dcu` has no source position" was
  the sentence in the status window, in the check's hint and in the
  empty-answer note of every `lsp_*` tool. Six of the eight states above
  refute it. What fails is a DCU output directory that does not **match**
  the sources: the units that are there come from one build, the missing one
  has to come from another, and then the compiler resolves it without a
  position (`Kibitz definition found in file:///(-1:0)` in the session log).
  Putting the single file back made the same query answer again.
- **The cure.** The check offered to move every one of those `.dcu` files
  aside - 108 rows on its own repository - and that is how you *create* the
  broken state. It now reports only what does not match: an **incomplete**
  directory (one row, with the first missing units named and a build as the
  cure) and a `.dcu` that is **older than its own source** (one row each).
  Moving a file aside survives only where the compiler could use the source
  instead, which is the one documented case it was built for: a project
  whose units are on its search path, where a stale `.dcu` really does
  shadow them.

### Smaller things

- The status row says `N .dcu for M unit(s)` in green when the directory
  matches the project and `INCOMPLETE - N .dcu for M unit(s)` in red when it
  does not. The unit count comes from the configuration's own `projectFiles`.
- `debug_consistency` with `apply` refuses a finding whose cure is a build
  instead of trying to move a directory aside, and says which cure it has.
- A `.dcu` is reported under the spelling of its source (`Unit1.dcu`, not
  `unit1.dcu`), and two compiler hints are gone from the build.

---

## 2.4.2 - 2026-10-20

### What the DCU finding may and may not claim

Measured against this repository with the session log on: with 108 `.dcu`
in the LSP's DCU output a definition ANSWERED, and after moving all 108
aside - directory empty, no `.dcu` on any search path, the project's own
source directory even added to it - the same query answered `null`, with
the agent log saying `Kibitz definition found in file:///(-1:0)`. That line
therefore means "the compiler had no source position"; a `.dcu` is one
cause of it, not proof. One case IS documented for real (another session's
log: 60 units loaded as DCU, and removing them fixed it), so the finding
stays - as a **warning about a risk**, with what it depends on and what
proves it per case, instead of a problem that names a cause.

## 2.4.1 - 2026-10-20

### What the first live call of the new tools found

- **`lsp_restart` now restarts.** It ended the session and promised that the
  next request would start a fresh one - which is false for the `lsp_*`
  tools: they deliberately never cold-start, so they answered "the session
  is not running yet" instead. A new session is started right away now
  (`start_again=false` keeps the old behaviour, and then the answer says
  what that means).
- **The status row "DCU output of that configuration" said "none named"**
  while `debug_consistency` found 108 `.dcu` through the same config file.
  The row computed its value on every tenth tick only - which is right for
  the window, whose tick runs every second, and wrong for the snapshot
  `get_status` uses: that one ticks once per CALL, so nine answers out of
  ten came from nothing. It is looked up on the first call as well now,
  like the `.delphilsp.json` row above it.

## 2.4.0 - 2026-10-20

### When DelphiLSP reads a .dcu instead of your source

A session log finally named the cause of answers that had looked like a
plugin defect for weeks: the `.delphilsp.json` tells the compiler where to
put DCUs (`-NU`), and a `.dcu` of a project unit lying there is **read
instead of the source**. For every symbol of such a unit DelphiLSP then
reports a definition without a file - the log says `definition found in
file:///(-1:0)` - so go to definition, find references and rename come back
empty while hover still works. Measured: 103 DCUs made 60 project units
load that way, and they were OLDER than the sources and were used anyway,
so keeping them current is no answer. Only an empty directory is.

It is visible now and it can be fixed:

- The **status window** says it under DelphiLSP: the DCU output that
  configuration names, and whether it is empty ("good") or holds N `.dcu`
  the session may read instead of the sources (red). One directory listing
  every tenth tick.
- **Project checks > Debug consistency** - and the `debug_consistency`
  tool - report one finding per shadowing `.dcu`, with the source it hides
  and whether it is even older than that source. Those findings, and only
  those, can be fixed from there: `apply=true` with their ids **moves** the
  files into a dated subdirectory of their own folder (never deletes them),
  after which `lsp_restart` gets a session that reads the sources again.

Measured against this repository while writing it: 108 `.dcu` in exactly
the directory its own `-NU` names, all 108 shadowing a project source.

## 2.3.0 - 2026-10-20

### The project itself: open it, close it, add and remove files

`project_open` opens a project or a project group, `project_close` closes
the group and every open file (the IDE has no "close only the project"),
and `project_files` lists what the project carries and adds or removes
entries.

The guards are the point, not the calls. Both open and close REFUSE while
anything is unsaved and name the files, because the IDE would ask about
them - and a question nobody can answer blocks the call instead of
reaching a human; `save_modified=true` decides it up front. `project_files`
takes `apply` like every other writing tool: the project file is not an
editor buffer, so a change to it is not undoable with Ctrl+Z. Adding a
file does not create or open it, removing one does not delete it from
disk, and the `.dproj` is written unless `save=false`.

## 2.2.0 - 2026-10-20

### The settings, by name - permanently or for this IDE only

The plugin's options were reachable only through its page in Tools >
Options, which a Claude Code session cannot click. The new `settings` tool
lists all fifteen with their current value, their default, their range or
choices and what they do - and sets one by name. `scope` decides how long
it holds: `permanent` writes the registry (this IDE and every one started
later), `session` changes it for the RUNNING IDE only. There is no default
scope on purpose: a change that may outlive the IDE is not one to guess at.

"Session only" means it: the registry keeps the value it had, and a `Save`
triggered anywhere else - the rename dialog storing its scope, the uses
cleanup its keep list - writes that old value rather than the temporary
one. `Save` is driven by the same descriptor the tool lists, so a setting
can no longer be in one place and missing from the other (two hand-kept
lists are what made the legacy migration drop five values).

### `lsp_restart`

Closing and reopening a project does not restart the plugin's DelphiLSP
session, so a session that has gone wrong - typically because it loaded
`.dcu` files instead of sources - could only be fixed by restarting the
IDE. `lsp_restart` ends it (the verification session with it); the next
request that needs DelphiLSP starts a fresh one.

### An empty LSP answer, corrected

2.1.1 said that a unit the session HAS analysed cannot be a `.dcu`
problem. A session log then showed otherwise: hover answered at the very
position, and the symbol's DECLARING unit had been loaded from a stale
`.dcu`, so the compiler reported its definition without a file. Both
causes are named now, with the log line that tells them apart
(`definition found in file:///(-1:0)`).

## 2.1.1 - 2026-10-19

### An empty LSP answer names the right cause

When an `lsp_*` tool answered nothing, the note always blamed a precompiled
`.dcu` as the "usual cause" - and in the case a reviewing session actually
hit, `lsp_hover` ANSWERED at that very position and only the definition was
unresolved, so the hint pointed at the wrong place and never mentioned
`find_original_symbol`, which does answer there. The note is derived now:
when the session has analysed the unit, it says that this is about the
POSITION (the private/public overload pair of Delphi 13.1 RSS-5463 is the
known case) and names `find_original_symbol`; only for a unit it has
analysed nothing of does the `.dcu` paragraph and the log recipe appear.
Fixed with it, a regression of 2.0.6: the tool's own sentence and the
generic text both opened with "DelphiLSP returned nothing", so the note read
"DelphiLSP returned nothing for this position. DelphiLSP returned nothing.
If ...". The text is one pure rule with a test now.

## 2.1.0 - 2026-10-19

### A "guide" tool: what this server can do, in one call

A fresh Claude Code session sees the tool NAMES and the always-loaded
instructions - 946 characters, deliberately no catalogue, because every
session pays for them whether it calls a tool or not. What it could not get
anywhere was the overview: which tool for which job, and the three rules
that cost a round when you do not know them (the IDE buffer is the truth,
`apply=false` first, an id or token belongs to ONE answer).

`guide` answers exactly that as plain text, grouped by job and covering
every tool. It takes no arguments, changes nothing, and - unlike every other
tool here - it needs **no running IDE**: the bridge answers it itself, so a
session that starts before the IDE is up can read it. It costs nothing until
it is called, which is why it may be as long as it is useful (~2 500 tokens)
while the instructions stay short.

The guide cannot go stale unnoticed: a test holds it against the shipped
tool definitions in all three directions - every tool in exactly one
section, no section naming a tool that does not exist, and every name
really present in the prose that section shows. It found its first defect
immediately: the `lsp_*` section described the tools without naming them.

## 2.0.6 - 2026-10-19

### An answer carries no key twice

An empty `lsp_hover` answer had the key `"note"` TWICE: one place added the
freshness diagnosis, the next line added its own sentence. That is valid
JSON of which a parser keeps one pair - its choice - so one of the two
sentences was lost at random. The two sentences are ONE note now, and in the
`lsp_*` tools every note goes through a helper that REPLACES, so a second
one cannot become a second key. A sweep over the whole unit was the evidence
that `lsp_hover` was the only such place; the four tools whose note is
purely informational were converted with it, and they now add the reason
when their answer is empty - the hint `lsp_definition` and `lsp_hover`
already gave.

### Which unit need not be open

`find_original_symbol` named "a unit that is not open" as a case the `lsp_*`
tools cannot answer, while those tools say the unit need not be open - a
contradiction in wording, not in substance: the tools send the unit you ASK
ABOUT, and what can stay invisible to them is the unit a symbol is DECLARED
in. Both sentences name their unit now.

## 2.0.5 - 2026-10-19

### An apply says what it really did

`interface_guids` and `dfm_events` had a fixed note for `apply=true`, so an
apply whose ids matched nothing at all still answered "The declarations you
named carry a fresh GUID now" - measured live, with `unknownIds` naming the
id right next to it. Both tools derive the sentence from the outcomes now,
in one shared place: how many were done, how many failed, how many ids
matched no finding - and "NOTHING was changed" when that is the truth. They
also report `fixed` / `notFixed` as numbers, like the two project checks that
had it right.

## 2.0.4 - 2026-10-19

### A unit with no source is a finding, not a gap

Measured on the project of the report that produced the check's
`no_source_found` kind: the row said "the project lists a unit that is not
there", while the notes above it counted the same unit among those that
"could not be read" - one answer contradicting itself, and the half-answer
the kind was introduced to avoid. The two cases have their own sentence now:
a file that is there and unreadable was not checked, a unit with no file at
all is reported below as a finding (with the one place the check cannot look
named: a path only the compiler knows).

## 2.0.3 - 2026-10-19

A reviewing model was given the MCP server with nothing but the tool list and
asked whether it could use the tools from the descriptions alone. It could,
for almost everything - and the two gaps it named are fixed here. Nothing in
the IDE changes; this is about what an agent can know before it calls.

### The rules file of `semantic_replace` is described in the tool itself

The format of `<project root>\semantic-replace.json` was nowhere in the MCP
surface, so an agent could run existing rules but never write one. The tool
now carries the shape (`find`, `replace`, `uses` as an array, `declaredIn`,
the optional `localVar` block), what the match does and does not do (dotted
identifier paths, not regular expressions; case-insensitive; never inside a
comment or a string; a rule without `find` is dropped), and that the file is
plain JSON rather than a Delphi source.

### Every writing tool says what protects its `apply`

The shared `apply` text ended with "plus a token" - and only 9 of the 23
writing tools hand one out, so for the rest it was impossible to tell from
the outside whether an apply is protected against a file that moved on. Each
of them now names its own guard:

- **a token** (9 tools): the preview pins the files *and* the arguments.
- **it looks again** (8): `apply` re-runs the tool's own analysis on the
  current buffers, so a plan made for a state that is gone cannot be carried
  out.
- **the ids of a scan** (4): `apply` re-scans and acts on the `fix_ids` you
  name; an id names the PLACE, so an old one finds nothing rather than
  something else.
- **it compares** (2): `apply` reads the files again and refuses when they
  differ from what was analysed.

A contract test walks the shipped list and fails on a tool that promises a
token it has no parameter for, or that says "no token" without naming what it
does instead. That sweep found five more writing tools the report could not
see, because their `apply` text was written per tool instead of shared -
among them `convert_properties`, whose `apply` had no description at all.

### Every parameter says what it is

Six parameters were named in the review as undocumented; the sweep that
checks it found **46** of 219 - among them the `apply` of
`convert_properties`, which had no description at all. All of them carry one
now, written from the handler with the default it really uses: `find_unit`
`partial`/`max`, `rename_preview` `scope` (project, unit or method - there
is no "selected"), `include_open_units`/`include_used_units`,
`convert_properties` `direction`/`getter`/`setter`/`to_line`, `uses_cycles`
`unit`/`max`, `blame` `start_line`/`end_line`, `commit_info` `diff_file`,
`lsp_references` `include_declaration`, `lsp_completion` `prefix`/`max`,
`lsp_request` `method`/`params`/`timeout_ms`, and the rest.

### Three answers the descriptions did not give

- **How a column counts**: "1-based" does not say what happens to a tab or
  to a character outside ASCII. Every `column` and `end_column` says it now:
  characters of the line as stored, a tab counts as ONE, not as its width on
  screen.
- **Whether `lsp_*` needs an open unit**: it does not. Each of the eight
  tools says that it reads the file (editor buffer first, else disk), sends
  that text to the session and waits for its analysis - and that it cannot
  make DelphiLSP resolve a unit its own compile does not reach, where
  `find_original_symbol` is the fallback that asks the identifier index.
- **What `edit_method` moves**: the routines a class declares. A field or a
  property is never listed and never moves, so a method that uses one leaves
  it behind - which the answer already named, without the description ever
  saying so. (The record comment in the source claimed "property / field"
  too; it was wrong and is corrected.)

### An id that matches nothing is said out loud

`interface_guids` and `dfm_events` ignored a `fix_ids` entry that matched no
finding, while their note said the places you named are done - the same
defect the line-ending check had in 1.25.1, in the one answer a caller cannot
check. Both report `unknownIds` now.

### Downloads

The same three files as 2.0.1: the package for the 32-bit and the 64-bit IDE,
and the MCP bridge.

---

## 2.0.1 — 2026-10-17

Everything since 1.23.2 (versions 1.23.3 … 2.0.1, 37 commits). The major
number changes because of one behaviour change, at the end of this entry.

### Downloads

The compiled package is attached to this release, so a team can roll it out
without building anything (issue #48):

| file | for |
| --- | --- |
| `DelphiRefactoringLight370-Win32.bpl` | the 32-bit IDE (`bin\bds.exe`) |
| `DelphiRefactoringLight370-Win64.bpl` | the 64-bit IDE (`bin64\bds.exe`) |
| `RefactoringLightMcp.exe` | the MCP bridge for Claude Code (optional) |

Both packages are built Release for RAD Studio 13 (37.0). Load one through
*Component > Install Packages*, or replace the file your team already loads.
Installing from the sources with `install.cmd` stays the normal way: it also
registers the package, builds and registers the bridge, and works on machines
without a command-line compiler. The detailed `.map` files are not attached by
default — ask for them when an address from a crash log has to be resolved,
because that needs the map of exactly this build.

### New

**Declare as constant** — editor menu, `Ctrl+Alt+Shift+C`, MCP tool
`declare_constant`. Select a number or a string and declare it as a `const` or
a `resourcestring`: local, in the class, or at unit level, optionally replacing
every other occurrence in the routine or the unit. The dialog checks live — the
preview line either shows the declaration that would be written or says in red
why it cannot be, so there is no second code path that could disagree with the
apply. It also opens with nothing selected, to define a constant from scratch.
A value that names an identifier is refused: that name need not be in scope
where the declaration goes.

**Project check "Line endings"** — and a warning the moment the IDE *opens* a
file with LF-only or mixed line endings, with "correct now", "this session" or
"always". Delphi copes badly with mixed endings, and the open is the right
moment: nothing has been typed yet, so the file can simply be rewritten and
re-read. Form files, binaries and anything inside the RAD Studio installation
are never touched. Measured on a real project: 25 findings in 2000 files, one
of them genuinely mixed (1035 CRLF + 55 LF).

**Project check "File encoding"** — the default pair is the one that fixes a
real defect: a source that is valid UTF-8 but carries **no BOM** is read by the
compiler as ANSI, so every non-ASCII character silently becomes two (measured
with dcc32 37.0). Pure ASCII is never a finding — its bytes are the same in
every encoding — and form files are left alone, because the IDE writes a
non-ASCII character as `#nnnn` anyway (measured: 1368 of 1381 form files of a
real project are pure ASCII).

**Project check "Unused project units"** (issue #47) — units the project file
lists although nothing in the project names them any more. A unit counts as
used when any other unit uses it (units on the search path included) or when
the project body names one of its identifiers, which is how a form reached only
through `Application.CreateForm` is used. Two kinds: *unused*, and *unused, has
initialization* — the latter a hint, because such a unit is often in the project
exactly to register itself. Read-only: it never edits the project file. Measured
on a real project: 86 findings in 1822 units.

Four new MCP tools, 56 → **60**: `declare_constant`, `line_endings`,
`file_encoding`, `unused_units`.

### Extract interface

Four reports, all of them from real classes:

- a **wrapped parameter list** was cut at the `;` inside it, and the
  continuation line came back as a "field"; an **`&`-escaped name**
  (`property &String`) parsed as empty (#44);
- an **array property** lost its index list, and whether its accessor is a field
  or a method was decided by the name starting with `Get`/`Set` instead of by
  the class's own member list (#46);
- the interface unit no longer copies the source unit's **whole uses clause** —
  it takes only what the generated interface really needs, which is the point of
  extracting an interface in the first place;
- an **existing target unit** is extended now (the interface is added at the end
  of its type block, its uses clause topped up) instead of asking whether to
  overwrite the file.

And the command now says whether the class **can implement an interface at
all** — a `TObject` descendant provides no `IInterface`, so the extracted
interface would not compile — and offers to fix it right there: derive from
`TInterfacedObject`, or add the `IInterface` implementation to the class itself.
Neither is preselected: both change how the instance is freed.

### The MCP bridge replaces itself

Windows cannot overwrite a running program, so an install while Claude Code was
running used to leave that session on the old bridge for good. A bridge of
1.33.0 or newer notices **between two requests** that an install has moved it
aside, starts the installed exe on its own stdin/stdout and waits for it: the
session keeps talking to the same process it started, while the new code
answers. Aside copies of earlier installs are removed at the next start.

### Installing

A run of reports from machines with no command-line compiler, or where Delphi
was installed by a different account:

- the install set an environment variable named `CONFIG`, and **MSBuild reads
  every environment variable as a property**, so `$(Config)` became the path of
  the config file and the build died trying to create a directory from it;
- a `bds.exe` build that printed a **fatal error reported success**, because
  bds.exe exits 0 either way. The error file has the last word now, decided by
  the compiler code rather than the localized severity word;
- the **MCP bridge is built by the installer itself**, as one row in the same
  summary — including through the IDE on machines where msbuild is not
  available, where the bridge was simply never built before;
- the package registration is **suspended during a bds build**, so the IDE
  started for the build neither reports a missing BPL nor locks the file the
  linker is about to write;
- msbuild is pinned to toolset 4.0 (a machine whose default was 2.0 failed with
  MSB4036);
- when the common directory belongs to another account, the build output goes to
  `%LOCALAPPDATA%` instead of failing with MSB3191;
- the install window could not be closed: a dialog watcher outlived the install
  and kept the console alive. It ends by itself now, and only ever watches the
  IDE it started.

### Behaviour change (this is what makes it 2.0)

A write to an editor buffer **from a worker thread now raises** a named
exception instead of only being counted. ToolsAPI and the editor buffers are
main-thread only; breaking that rule does not fail, it corrupts a buffer now and
then — the worst way for a defect to behave. The flip had a condition written
down when the counter was introduced: one live round over the write paths in
which the counter stays at zero. That round ran before the flip.

### Fixes

- **Nested routines declared before a method's `begin` hid the method itself**
  from the enclosing-routine walk — which feeds declare-as-constant, extract
  variable, the rename scope "in current method" and the source resolution of
  find references and rename. A sweep over 344 files: 21,201 lines newly
  attributed to their routine, not one range changed.
- No check ever rewrites a file of the **RAD Studio installation** (an Indy
  source open in the editor was offered for conversion).
- The editor menu's refactorings are **two alphabetical blocks** now — "introduce
  something new" and "change what is there"; Remove with and Expand include
  files moved to the project-wide group they share a scope dialog with.
- **"Go to"** selects the identifier instead of marking from the start of the
  line (interface GUID check, DFM event check), and in Unused project units it
  points at the entry in the `.dpr` for a unit that has no file at all.
- A message about a write **says where the write went** — the editor buffer
  (undoable) or the disk. For a project-wide check most files are not open.
- Result windows: **Close really closes**, and the shared progress window paints
  its text before the scan starts instead of standing there empty.
- An MCP handler stuck in an LSP wait no longer makes the package unload leak
  and pin itself (**PR #45**).
- The option chooser measures its text instead of assuming a height — a long
  prompt was drawn behind the first radio button.

### Upgrading

Run `install.cmd` and restart RAD Studio, or load the attached `.bpl`. A running
Claude Code session picks up the new bridge by itself from 1.33.0 on; an older
one needs a new session.

Thanks to **nLeonardt95** for issues #44, #46, #47 and #48, and to
**Ian Branch** for PR #45.

---

## 1.23.2 — 2026-10-05

[Full notes](https://github.com/jaenicke/DelphiRefactoringLight/releases/tag/1.23.2)
— covers 1.11.3 … 1.23.2: 64 versions in 98 commits, and the DUnitX suite grew
from 176 to 351 tests on Win32 and Win64.

- **New feature "Edit methods"**: edit what a class declares — move members to
  another class (also in another unit), change the signature, set modifiers, for
  several members at once. The occurrences are verified with DelphiLSP *before*
  the edit, and afterwards a window lists what still needs hand work.
- **Every refactoring is reachable through the MCP bridge**, and every writing
  tool shows the change before it writes it.
- **An external code audit by Ian Branch** — 12 pull requests plus 7 area issues
  with 123 findings — implemented in full, including several findings where code
  could be lost or silently changed.
- **Find references and rename** work on a cold DelphiLSP session and on large
  projects without dropping occurrences.
- **The plugin reports its own failures**: the status window explains the LSP
  state in plain language and in colour, the main-thread rule is counted, and a
  tool answer that cannot be right is flagged as a defect.
- The interface GUID check fixes what it finds; a result row follows its line
  through your own edits; uses cleanup asks the form designer what it cannot see
  in the source.

---

## 1.11.2 — 2026-09-22

[Full notes](https://github.com/jaenicke/DelphiRefactoringLight/releases/tag/1.11.2)
— a small release with one important fix: **if you use the 64-bit IDE, please
update.**

- The 64-bit IDE could **freeze for about a minute and use 14 GB of memory**
  (#16, thanks @ijbranch). The identifier index cache was written with 8-byte
  numbers by the 64-bit build while the reader expects 4, so four bytes of a unit
  name were read as a count — and on Win64 that multi-gigabyte allocation
  *succeeds*, so the safety net never ran. Both IDEs now write identical cache
  files, every length is checked before memory is reserved, and loading is much
  faster (60 ms instead of 4 s for the complete RTL/VCL).
- **MCP bridge**: less context in every Claude Code session, the bridge carries
  the plugin's version, and the README explains `--scope user` vs `--scope local`.

---

## 1.11.0 — 2026-09-21

[Full notes](https://github.com/jaenicke/DelphiRefactoringLight/releases/tag/1.11.0)
— mostly about **find references and rename being right the first time**.

- **An aborted request no longer deletes a result** (#13, thanks @nLeonardt95):
  while a big project loads, DelphiLSP cancels requests after 10 s, and those
  cancellations were treated as "not a reference".
- Occurrences DelphiLSP does not answer for are **resolved from the sources**
  (qualifier → declared type → member), which also covers the private/public
  overload pair DelphiLSP stays silent about since Delphi 13.1 (RSS-5463).
- **Overloads are told apart** by argument count and, where the text shows them,
  argument types.
- A second, lightweight DelphiLSP session in "agent" mode does the verification
  (14.3 s → 6.5 s for 21 candidates, identical answers).
- **Closing the result window stops the search.**
- **No stack overflow on deep uses graphs** (#15, thanks @ijbranch).
- The 64-bit IDE's **MCP named pipe** works (Windows error 998, a misaligned
  buffer), and a broken registry value can no longer stop the plugin from loading.

---

## 1.7.0 — 2026-09-19 — first official release

[Full notes](https://github.com/jaenicke/DelphiRefactoringLight/releases/tag/1.7.0)
— a refactoring and navigation plugin for the Delphi IDE (RAD Studio 13, 32-bit
and 64-bit), built on DelphiLSP: wherever it matters, every occurrence is
verified by the language server instead of being matched as text.

- **Refactorings**: rename (project, unit, method or selected units, form files
  included), change signature across the whole method family, safe delete,
  extract method, extract variable, wrap in try..finally, convert properties,
  remove with, move to (new) unit, extract/extend interface, add IInterface
  support, align method signature, semantic replace, expand include files.
- **Navigation**: find references with the kind of each occurrence, find
  implementations, find unit references, find original symbol, find the unit for
  an identifier.
- **Code assistance**: live quick fixes (missing unit, "did you mean", broken
  uses entry, missing variable, interface methods, empty implementations, unused
  variables and private members, missing semicolon), code completion with
  generated event handlers and anonymous methods, and uses cleanup.
