# Refactoring Light - Standalone

A Windows-VCL executable that runs the Refactoring Light refactorings
without the Delphi IDE. The wizards, dialogs and engines are the same
codebase as the [IDE plugin](../README.md); only the editor host
differs.

## When to use the standalone

- You want to refactor a Delphi codebase from a machine where the IDE
  itself is locked, slow to start, or unstable.
- You want to script bulk refactorings (e.g. apply a Semantic Replace
  rule set across an entire project) without bringing up RAD Studio.
- You are evaluating the wizards before paying the cost of installing
  the BPL into the IDE.
- The IDE keeps crashing on a particular refactoring and you need a
  process that you can kill cleanly.

## Build

`Standalone\build.bat` drives an `msbuild` of
`RefactoringLightStandalone.dproj`. Defaults to `Debug`; pass `Release`
to optimize:

```cmd
cd C:\Beispiele\DelphiRefactoringLight\Standalone
build.bat               REM Debug
build.bat Release       REM Release
```

The build pipeline:

- `rsvars.bat` sets up `BDS`, `PATH`, etc. for the active RAD Studio
  installation (currently hard-coded to BDS 37.0).
- `msbuild` reads `RefactoringLightStandalone.dproj`. The `.dproj`
  defines `STANDALONE_BUILD` in `DCC_Define` and adds `..\Source` to
  `DCC_UnitSearchPath`, so the build needs no extra command-line flags.
- Per-config DCC defines: `DEBUG` (stack frames + range checks) vs.
  `RELEASE` (optimized, no debug info).
- Output: `Standalone\Output\RefactoringLightStandalone.exe`.
- DCU cache: `Standalone\DCU\`.

The `.dproj` opens directly in RAD Studio so you can `F9`-run with the
debugger attached during wizard development.

## What the executable provides

Top-level menu structure:

```
File
├── Open project...           (Ctrl+O)
├── Save active file          (Ctrl+S)
└── Exit

Refactor
├── Rename identifier...      (Ctrl+Alt+Shift+R)
├── Find references           (Ctrl+Alt+Shift+F)
├── Find implementations      (Ctrl+Alt+Shift+I)
├── Align method signature...
├── -----
├── Remove with
│   ├── At cursor only
│   ├── In current unit
│   ├── In selected units...
│   └── In whole project...
├── Move to unit...
├── Find unit references...
├── -----
├── Extract / extend interface
│   ├── Extract new interface from class...
│   ├── Add to existing interface...
│   └── Add IInterface support to class...
├── -----
└── Semantic replace
    ├── In current unit
    ├── In selected units...
    ├── In whole project...
    └── Edit rules...
```

Layout:

- **Left**: `TTreeView` listing every source file the `.dproj` declares
  via `DCCReference Include="..."`. Click to load into the editor.
- **Center**: `TMemo` showing the active file's content. Plain text;
  no syntax highlighting (v1 - switching to `TSynEdit` is planned).
- **Bottom**: `TStatusBar` showing the project root, the active file
  name, and the 1-based caret line/column.

After each refactor, the editor and tree refresh automatically:

- The Memo reloads from the standalone state's buffer (which the
  wizard writes through `Editor.ReplaceFileContent`), preserving the
  caret's logical line/column and the top-of-view scroll position.
- The tree re-reads the `.dproj` so a new file created by Extract
  Interface (which writes a `<DCCReference Include="..."/>` entry)
  appears immediately.
- A wait cursor and `Self.Enabled := False` mask the (still
  synchronous) refactor execution; the LSP queries do not get a
  background thread in v1.

## Wizard parity matrix

All wizards reach the editor through the `IEditorHelper` interface, so
they work against the standalone host with no per-wizard changes. Where
they behave differently from the IDE plugin, the reason is the host -
not the wizard.

| Wizard                       | Standalone | Notes                                                                                                    |
| ---------------------------- | :--------: | -------------------------------------------------------------------------------------------------------- |
| Rename identifier            |    yes     | Needs LSP. Edits go through `Editor.ApplyEditViaEditor` → disk write + buffer sync.                      |
| Find references              |    yes     | Needs LSP.                                                                                               |
| Find implementations         |    yes     | Needs LSP.                                                                                               |
| Align method signature       |    yes     | Needs LSP.                                                                                               |
| Remove with (any scope)      |    yes     | Needs LSP for type resolution.                                                                           |
| Move to unit                 |    yes     | Cross-file edits via the same path as Rename.                                                            |
| Find unit references         |    yes     | Pure text scan + LSP `documentSymbol`. Logs to `<project_root>/UnitRefsTrace_*.log`.                     |
| Extract new interface        |    yes     | New unit lives next to the source class. `Editor.AddFileToActiveProject` patches the `.dproj` XML.       |
| Add to existing interface    |    yes     | As above.                                                                                                |
| Add IInterface support       |    yes     | In-place edit of the class declaration.                                                                  |
| Semantic replace (all modes) |    yes     | Reads `semantic-replace.json` from the project root.                                                     |
| Extract Method               |    no      | Relies on the IDE's `IOTAEditBlock` to read the user's text selection. Without `TSynEdit`, no selection. |
| Code Completion              |    no      | Inline list-driven UI only meaningful inside an LSP-aware editor.                                        |

## Requirements at runtime

1. **DelphiLSP available.** The wizards that touch LSP look up
   `DelphiLsp.exe` first from the registry (`HKCU/HKLM
   Software\Embarcadero\BDS\<ver>\RootDir + bin\DelphiLsp.exe`), then
   fall back to the hard-coded path
   `C:\Program Files (x86)\Embarcadero\Studio\37.0\bin\DelphiLsp.exe`.
   No RAD Studio install on the target machine → no LSP-driven
   wizards; the standalone still runs Semantic Replace and Find Unit
   References (which is a text scan).
2. **`<project>.delphilsp.json` next to the `.dproj`.** Without it,
   DelphiLSP cannot index the project. The standalone reads the same
   file you generate for IDE use (`File > New > Other... > Delphi LSP
   Project Settings` in RAD Studio, or hand-written).
3. **The `.dproj` must list every source file as a `DCCReference`.**
   The standalone parses `.dproj/ItemGroup/DCCReference[@Include]` to
   build the source-file list; files reached purely via `uses`-clause
   search paths but not referenced in the `.dproj` are invisible to
   the tree and to project-wide refactorings.

## How it works internally

The wizards depend only on the `IEditorHelper` interface
([Source/Expert.EditorHelperIntf.pas](../Source/Expert.EditorHelperIntf.pas)).
Two implementations satisfy it:

```
+----------------------------------+
|        Refactor wizards          |  (no ToolsAPI references in their
|  (Rename, Extract Interface,     |   bodies - only Editor.X calls)
|   Semantic Replace, ...)         |
+----------------+-----------------+
                 | uses
                 v
+----------------------------------+
|     Expert.EditorHelperIntf      |
|     IEditorHelper interface      |
|     + global Editor accessor     |
+--------+--------------+----------+
         |              |
   IDE plugin     Standalone .exe
         |              |
         v              v
+---------------+ +-------------------------+
| TIDEEditor    | | TStandaloneEditor       |
| Helper        | | Helper                  |
| (ToolsAPI)    | | (FState: dproj XML +    |
|               | |  in-memory editor buf)  |
+---------------+ +-------------------------+
```

- `SetEditorImpl(...)` is called in `TMainForm.FormCreate` with a fresh
  `TStandaloneEditorHelper`. Wizards calling `Editor.GetCurrentContext`,
  `Editor.ReplaceFileContent` etc. transparently hit the standalone
  implementation.
- `TStandaloneProjectState` is the standalone's "what the editor knows"
  store: `.dproj` path, project root, `DCCReference` list, an
  `active file` pointer with cursor line/col, and a dictionary of open
  editor buffers (filename → live text).
- Idempotent `Editor.AddFileToActiveProject` rewrites the `.dproj` XML
  to append `<DCCReference Include="..."/>` entries; the standalone
  reloads `FState.SourceFiles` afterwards so the tree updates.

Some IDE-only services (`IOTAEditorServices`, `IOTAModuleServices`,
`IOTAIDEThemingServices`, the `TNotifierObject` base class) are
hidden behind `{$IFNDEF STANDALONE_BUILD}` guards in the wizards,
`Expert.EditorHelper`, `Expert.DialogHelper`, `Expert.IdeThemes` and
`Expert.LspManager`. The standalone build never references ToolsAPI;
the IDE plugin build is unchanged.

## Known limitations

- **UI blocks while a refactor runs.** The wizards run on the main
  thread. The wait cursor is the v1 workaround; threading the LSP
  queries is the v2 plan.
- **One file visible at a time.** Switching tree entries replaces the
  Memo content; there is no tab control yet.
- **Plain Memo, not TSynEdit.** No syntax highlighting; selection-based
  refactorings (Extract Method) are therefore unreachable. TSynEdit
  is the v2 candidate.
- **No "dirty" indicator.** A modified buffer is auto-flushed to disk
  on every Memo `OnChange`, so there is no concept of unsaved changes
  - which is also why `Ctrl+S` is essentially a re-flush.
- **No deep `.dproj` rewrite.** Beyond `DCCReference Include` entries,
  the standalone does not touch search paths, defines, or
  configurations.

## Troubleshooting

- **"Refactoring failed: ..."** popups bubble up from a wizard
  exception; the standalone catches them in `TMainForm.RunWizard` so
  the form stays usable. Check the message text for the root cause -
  most often a missing `.delphilsp.json` or a broken `DelphiLsp.exe`
  resolution.
- **The tree is empty after Open project.** The `.dproj` does not
  declare any `<DCCReference Include="*.pas|*.dpr|*.dpk"/>` entries.
  Add files to the project in RAD Studio (or hand-edit), save, and
  reopen.
- **Wizard finds the wrong project.** Standalone uses the `.dproj` you
  last loaded via `File > Open project...`. There is no `File > New
  group` concept; if you need to switch projects, just open another.
