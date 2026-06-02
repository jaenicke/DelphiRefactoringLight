# Delphi Refactoring Light

A design-time package for **Delphi 13** that connects to the built-in Delphi Language Server (`DelphiLSP.exe`) to provide nine refactoring features directly in the editor:

| Shortcut               | Feature                                                                                                                                              |
|------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------|
| `Ctrl+Alt+Shift+R`     | **Rename** &mdash; rename an identifier project-wide, semantically (including interface implementations)                                             |
| `Ctrl+Alt+Shift+U`     | **Find References** &mdash; list every occurrence of an identifier                                                                                   |
| `Ctrl+Alt+Shift+F`     | **Find Unit References (project-wide)** &mdash; list every file in the project that imports the current unit, with each identifier of that unit it actually uses |
| `Ctrl+Alt+Shift+I`     | **Find Implementations** &mdash; list all class implementations of an interface/virtual method                                                       |
| `Ctrl+Alt+Shift+Space` | **Code Completion** &mdash; suggestions via DelphiLSP                                                                                                |
| `Ctrl+Alt+Shift+M`     | **Extract Method** &mdash; move the selected block into a new method                                                                                 |
| `Ctrl+Alt+Shift+A`     | **Align method signature** &mdash; compare a method's class/interface declaration with its implementation and highlight mismatches                   |
| `Ctrl+Alt+Shift+W`     | **Remove with** &mdash; rewrite a `with` statement as inline-vars + qualified accesses. Scope (at cursor / current unit / selected units / project-wide) is picked from the submenu; the shortcut defaults to "at cursor only" |
| `Ctrl+Shift+M`         | **Move identifier to other unit** &mdash; move a type / class / routine / const / var to another existing unit and update consumer `uses` clauses    |

Unlike purely text-based tools, this package uses the actual LSP requests that DelphiLSP advertises in its `initialize` response: `textDocument/definition`, `textDocument/declaration`, `textDocument/implementation`, `textDocument/documentSymbol`, `textDocument/hover`, `textDocument/completion`, and `publishDiagnostics` push notifications (used to detect inactive `{$IFDEF}` regions when DelphiLSP delivers them &mdash; diagnostic code `H2655`/`H2656` with tag `Unnecessary`). DelphiLSP does **not** implement `textDocument/rename`, `textDocument/references`, `textDocument/foldingRange`, `textDocument/selectionRange` or `textDocument/documentHighlight` &mdash; for rename and find-references the package therefore runs a project-wide text search and verifies every candidate semantically via `textDocument/definition`. Identifiers that happen to share a name but belong to different symbols are cleanly distinguished.

The package starts its own DelphiLSP session in single-process mode. An earlier version used `serverType: controller` (the same mode the IDE itself uses) to maximise the diagnostic coverage, but extensive testing showed that DelphiLSP's controller-mode sub-agents (`Agent0`/`Agent1`) need parent-process / COM-bridge context only `BDS.exe` can provide &mdash; spawned by anything else they crash silently and every subsequent `textDocument/hover` returns *Internal server error*. Single-process mode resolves hover reliably; the trade-off is that inactive-region diagnostics now depend on whatever DelphiLSP volunteers without the controller's `returnDccFlags`/`returnHoverModel` hints. By default the LSP is **pre-warmed automatically when a project opens**, so the first refactoring action does not have to pay the cold-start cost; this can be turned off in *Tools &rarr; Options &rarr; Refactoring Light*. Every refactoring dialog shows the current LSP warm-up status in its title bar.

In the last three weeks I tested with big projects and used it myself in real life. A few fixes were neccessary, but now it should be a good help though it might not work in all cases. Important is, that you can always use undo, because changes are applied in a way to make this work.

**If you encounter any problems, please let me now, so I can fix it.**

## Features in Detail

### Rename (`Ctrl+Alt+Shift+R`)
- Reads the editor position and picks up the identifier under the cursor.
- Runs a text search across all project files, then verifies each candidate semantically via `textDocument/definition`.
- Extends the candidate set via `TImplementationFinder` with class method implementations &mdash; so renaming an interface method also renames the implementations in every class that implements that interface.
- Shows a preview dialog with a list view: file, line, kind (`Interface declaration`, `Class declaration`, `Implementation`, `Call`, ...), original line, and preview line. A second tab holds the full diagnostic log.
- Optional per-file backup.
- Applies the changes byte-precisely via `IOTAEditWriter` and reloads modified modules in the IDE.

### Find References (`Ctrl+Alt+Shift+U`)
- Reads the identifier under the cursor.
- Tries `textDocument/references` on the LSP server first.
- If that returns nothing (or the server does not support it), falls back to the same strategy as Rename: project-wide text search plus per-candidate verification via `textDocument/definition`.
- Shows the results in a dialog (file, line, column, line preview).
- **Double-click** or **Enter** jumps to the location.

### Find Unit References (`Ctrl+Alt+Shift+F`)
- Works on the **active unit** &mdash; no editor cursor needed.
- Step 1 (textual): scans every project source file's `uses` clauses (interface + implementation) for the current unit's name. Files with a `uses` entry but no actual symbol usage get a single "dead reference" row.
- Step 2 (LSP): queries `textDocument/documentSymbol` for the unit, builds the set of exported identifier names, then scans every using-file for matching tokens (comment- and string-aware).
- Step 3 (LSP verify): for every candidate it asks `textDocument/definition` and only keeps hits that resolve back into the original unit. Common names like `Create` / `Free` that happen to live in another unit get filtered out.
- Result list: one row per occurrence with file, line, column, identifier and line preview; dead-reference rows show "&lt;unit&gt; is listed in the uses clause but no symbols of it are used here".
- **Double-click** / **Enter** jumps to the location.

### Find Implementations (`Ctrl+Alt+Shift+I`)
- Place the cursor on an **interface method declaration** or on a **call** of an interface/virtual method.
- Uses LSP `textDocument/definition` to reach the actual method declaration (so calls like `aa.Bar` inside `TXyz.Test` resolve correctly to `ITest.Bar` rather than to the containing class).
- Determines the containing type at the declaration (e.g. `ITest`).
- Scans all project files for class method implementation lines (`procedure TClass.Method`, `function TClass.Method`, etc.) and keeps only those classes that actually implement the containing type, either directly or via inheritance.
- Results are shown in the same dialog as Find References (title "Implementations: &lt;name&gt;").
- **Double-click** or **Enter** jumps to the implementation.
- The same `TImplementationFinder` class is also used internally by the Rename feature.

### Code Completion (`Ctrl+Alt+Shift+Space`)
- Queries `textDocument/completion` at the cursor position.
- Displays suggestions in an owner-drawn popup (sort, filter, detail tooltip).
- Inserts the chosen entry into the editor cleanly (via `IOTAEditWriter`, byte-precise, no Auto-Indent interference).

### Align Method Signature (`Ctrl+Alt+Shift+A`)
- Place the cursor on a method name (in the class/interface declaration or on the implementation).
- Queries `textDocument/documentSymbol` for the current file and walks the tree to collect every match: methods inside class types, methods inside interface types, and stand-alone implementations.
- Resolves the counterpart in another unit via `textDocument/definition` and queries that file as well.
- The actual signature text is read **directly from the saved source file**, not from the LSP's `name` field &mdash; DelphiLSP serves that field from a symbol cache that lags behind editor edits, so reading the file gives the truth on disk.
- Each entry is normalized (whitespace and case folded, leading `TClass.` qualifier stripped) and compared against the majority signature.
- The interactive dialog lists every entry with role (`Class decl.`, `Interface decl.`, `Implementation`), container, file, line, match flag and the full signature. Rows that differ from the majority are tinted red.
- **Double-click** to jump to a location (dialog stays open), **Enter** to jump and close.
- v1 is diagnostic only &mdash; no automatic rewriting; review and fix the mismatches yourself.

### Remove with (`Ctrl+Alt+Shift+W`)
- The editor context menu exposes the action as a submenu with four scopes:
  - **At cursor only** &mdash; just the `with`-statement that encloses the caret. Default for the global shortcut, since it's the fastest single-edit case. If no `with` encloses the caret, falls back to the current unit.
  - **In current unit** &mdash; the active editor file.
  - **In selected units...** &mdash; opens a multi-select list of all project source files; scan only the chosen ones.
  - **In whole project...** &mdash; project-wide scan. On very large code bases (10k+ files) this can take many minutes; that's why it now requires an explicit submenu pick instead of being the default action.
- Saves any unsaved editor buffers, starts / reuses DelphiLSP and ensures the chosen file set is indexed. Per-file: triggers `RefreshDocument` (which sends `didOpen` plus a `didChange` v2 with the full content, mirroring what the IDE itself sends), actively requests `textDocument/documentSymbol` to force analysis, then waits up to 90 s (30 s per file in project-wide scope) for any inactive-region diagnostics. The dialog title bar reflects the active scope and the current LSP warm-up status.
- Scans every `*.pas` / `*.dpr` / `*.dpk` in the chosen scope for `with ... do` statements:
  - **`begin..end`** bodies &mdash; full block.
  - **Single statement** bodies &mdash; one expression terminated by `;`.
  - **Compound bodies** &mdash; `try..end` / `case..end` / `asm..end`. The block is kept verbatim and var-decls are emitted right above it (no redundant `begin..end` wrapping). Body indentation is left-shifted by the difference between the original opener column and the new with-keyword column.
- For each occurrence, hoists the with-target(s) into Pascal **inline variables** at the with-statement location (Delphi 10.3+ syntax) and rewrites every body identifier with the appropriate qualifying prefix:
  - Side-effect-free single identifiers (`FParser`, `Self.FFoo`, `p^`, `Self.FNode^`) keep the prefix directly &mdash; no temp variable.
  - Multi-segment dotted paths and any expression with parens / brackets / `as`-cast are forced into a temp var so each side-effect runs exactly once.
  - Temp-var naming uses the LSP-resolved type when known (e.g. `Form.GetBitmap()` &rarr; `LBitmap`, `TRegIniFile.Create(...)` &rarr; `LRegIniFile`); leading `F` (field convention) and `T` (type convention) are stripped. A textual fallback ("walk back to the last identifier in the expression") covers cases where LSP cannot help. Final-fallback name is `LWithN` (1-based). Names never start with `_` &mdash; the Delphi style guide prefers leading capital letters.
  - Cross-target collisions and collisions with body-identifiers append the target index to disambiguate (`LAdd` / `LAdd2`).
- Target-type resolution chain &mdash; on top of the basic `textDocument/definition` query, the engine walks several Pascal-specific shapes:
  1. **Pointer / simple aliases** (`tLiFo = _L_List;`, `_L_List = ^_L_List_Node;`) are hopped through (max 4 hops) with a text-fallback for the DelphiLSP self-reference quirk.
  2. **Constructor targets** (`with TFoo.Create(...) do`) are recognised by parsing `constructor TFoo.Create(...)` lines &mdash; the class qualifier is the implicit return type.
  3. **`as`-cast targets** (`with Comp as TWinControl do`) follow LSP to the class declaration directly.
  4. **Inheritance walk** &mdash; up to 6 ancestors via `class(Parent)` so inherited members match (e.g. `RowCount` on a local `TStringGrid` descendant resolves against `Vcl.Grids.TStringGrid`).
  5. **Parent-unit hints** &mdash; when LSP self-refs on a `class(UnitX.TBar)` parent, the unit qualifier is stored as a fallback hint: any body-ref whose LSP result lands in a file with that name segment matches.
  6. **Pre-compiled types** (e.g. `TDockableForm` from `DockForm.dcu` in `DesignIde.bpl`, no `.pas` shipped) &mdash; resolved partially via the constructor / decl-line heuristic, temp var is still emitted; body refs the engine cannot prove stay unqualified for manual review.
- Body identifiers are mapped to the right target via four strategies, in order:
  1. **Direct member parsing** of each target's class/record body (rightmost-target-wins, matching Pascal `with` semantics).
  2. **`textDocument/definition`** as fallback for inherited members (in the direct class range or any ancestor's range).
  3. **`DeclFile` soft match** for cases where the type's full source is unavailable but the decl file is known.
  4. **`ParentUnitHints`** (basename of the body-ref's LSP result vs the parent's unit qualifier).
- **Inactive `{$IFDEF}`-region detection**: when DelphiLSP pushes `publishDiagnostics` with code `H2655`/`H2656` and tag `Unnecessary` for a file, the wizard skips any `with`-statement inside such a region with status *"inactive $IFDEF region &mdash; skipped"*. If DelphiLSP delivers **no** diagnostics for a file at all (which can happen in single-process mode for files outside the active build configuration, or when the server has not finished analysing), all occurrences in that file are skipped with status *"LSP no diagnostics &mdash; skipped (dead-code unknown)"*. The wizard deliberately does **not** fall back to a text-based `{$IFDEF}` scanner &mdash; Pascal's `$IF defined(...)`, `$IFOPT`, nested `$IF`s and project-specific defines make a reliable scanner impractical, and silently rewriting code that might be dead would be worse than the honest skip.
- Review dialog with two tabs:
  - **Diff** &mdash; before / after side by side, per occurrence.
  - **Debug** &mdash; per-target type info (resolved type file, class line range, parsed direct members, chosen inline-var name, qualify-prefix, ancestor list, parent hints, resolve note) and per-body-identifier resolution (LSP result, match source, applied prefix). Useful for verifying the rewrite is sound before applying.
- **Apply selected**, **Apply all** or **Close**. Applied edits go through `IOTAEditWriter` so they are individually undoable in the IDE.
- v1 limitations: nested `with`-statements inside the body are flagged as multi-target / manual review.

### Extract Method (`Ctrl+Alt+Shift+M`)
- Validates the selection with a Pascal tokenizer (paren balance, `if`/`then`/`else`, `repeat`/`until`, `try`/`except`/`finally`, no selection crossing method boundaries, ...).
- For every identifier in the selection, asks LSP `textDocument/hover` for the symbol kind (local / parameter / field / global) and `textDocument/definition` for the declaration site.
- Determines:
  - **Parameters** of the new method (locals from the surrounding method that are read inside the block).
  - **Local variables** of the new method (declarations from the block itself).
  - **Return value** if a single trailing identifier is written-then-read.
- Parameters are prefixed with `A` in the new method signature and body (Delphi convention, e.g. `ACertFilename` instead of `CertFilename`); the call site keeps the original variable names.
- Generates the new method (class-qualified if applicable), replaces the selected block with the call, and removes the now-unused `var` entries in the original method.
- Refreshes the form and class structure in the IDE (`Module.Refresh(False)`, `FormEditor.MarkModified`).

### Move identifier to other unit (`Ctrl+Shift+M`)
- Place the cursor on a top-level symbol to move: **type / class / interface / record**, **routine** (`function`/`procedure`), **constant**, **variable**, or **resource string**. For classes, the move includes the class declaration *and* every method-implementation block in the same unit (matched by `TClass.Method` qualifier; overloads with the same name move together).
- A modal dialog lists the project's other `.pas` files; pick the target unit. Existing units only &mdash; new ones are not created.
- The engine performs the move:
  1. **Locate** the declaration range (interface section, including any preceding `type` / `var` / `const` keyword as appropriate) and the impl block range(s).
  2. **Collect required uses**: for each identifier referenced inside the moved range it asks `textDocument/definition` and records the declaring file's unit name. Identifiers that resolve to `System` (built-ins like `Length`, `IntToStr`, ...) are filtered out. Identifiers whose declaration falls **inside the moved range itself** (local vars, parameters, the symbol's own type members) are also filtered.
  3. **Insert** the declaration into the target's interface section (just before `implementation`) with one blank line of padding on each side; insert the impl block(s) just before the final `end.`. A section header (`type` / `var` / `const` / `resourcestring`) is prepended automatically if needed.
  4. **Add required `uses`** to the target: interface-side uses for identifiers referenced from the moved declaration (visible-surface types), implementation-side uses for identifiers referenced only from the impl body. Avoids double-imports.
  5. **Remove** the declaration and impl block(s) from the source. The cleanup pass strips orphan `{ TClassName }` class-marker comments above moved method blocks, drops now-empty section headers (`type`/`var`/`const`/`resourcestring` with no body below them), and collapses runs of blank lines.
  6. **Update consumer `uses`**: every other project file that already referenced the moved symbol via the old unit name gets the target unit added to its interface uses. If the source unit is no longer referenced from that consumer at all, it is removed.
- Source-unit `uses` are intentionally left intact (over-imports are harmless; the engine does not attempt to prove unused-uses safety here).
- Edits are applied via `IOTAEditWriter` and individually undoable.
- v1 limitations:
  - No new-unit creation.
  - Cross-unit qualified references (`SourceUnit.X`) in consumers are not rewritten &mdash; only the `uses` clause is adjusted.
  - The cleanup is heuristic; review the diff before committing.

## Requirements

- **Delphi 13** (BDS 37.0) &mdash; tested with the bundled `DelphiLSP.exe`. The LSP executable path is read from the registry (`HKCU\Software\Embarcadero\BDS\<version>\RootDir`) with a fallback via `IOTAServices.GetRootDirectory`, so the package works with any standard RAD Studio installation path.
- One `*.delphilsp.json` per project next to the `.dpr`/`.dpk` (see `DelphiRefactoringLight.delphilsp.json` for an example). The package hands this file to the LSP via `workspace/didChangeConfiguration`.

## Installation

### Recommended: via the Delphi Install Helper (DIH)

```cmd
install.cmd
```

The script:
1. builds the helper tool `delinst.exe` from the `dih\` sub-folder,
2. reads `DelphiRefactoringLight.xml`,
3. builds `Packages\DelphiRefactoringLight.dproj` (Win32 / Release), and
4. registers the resulting BPL in the IDE's `Known Packages` (`HKCU\Software\Embarcadero\BDS\37.0\Known Packages`, value `$(BDSCOMMONDIR)\Bpl\DelphiRefactoringLight.bpl`).

The `Known Packages` mechanism is the standard registration path for BPL-based wizards in Delphi: the IDE loads the package, calls its `Register` procedure, which Registers the main wizard via `RegisterPackageWizard` and the key bindings via `IOTAKeyboardServices`.

An **IDE restart** is required afterwards so the package is loaded cleanly. The package itself shows a restart hint dialog after installation.

Further scripts:
- `rebuild.cmd` &mdash; build only (no IDE registration).
- `uninstall.cmd` &mdash; remove the package from the IDE registry.

### Alternative: manual install inside the IDE

1. Open `Packages\DelphiRefactoringLight.dproj` in RAD Studio.
2. Platform: `Win32`, configuration: `Release`.
3. Right-click the package in the Project Manager &rarr; **Install**.
4. Restart the IDE.

> **Note:** The manual path also registers under `Known Packages`, but with whichever output directory the IDE used (usually `.\Output\...`).
> If you later switch to `install.cmd`, a second entry with a different path appears &mdash; the package is then loaded twice (visible as duplicate keyboard shortcut conflicts and access violations). Remove the stale entry manually:
>
> ```cmd
> reg delete "HKCU\Software\Embarcadero\BDS\37.0\Known Packages" /v "<old path>\DelphiRefactoringLight.bpl" /f
> ```

## Project Layout

```
DelphiRefactoringLight/
|
|-- Source/                                  # All Pascal units
|   |-- Expert.Registration.pas              # Register, wizard init
|   |-- Expert.RenameWizard.pas              # IOTAMenuWizard for Rename
|   |-- Expert.RenameDialog.pas              # Preview dialog with list view + details
|   |-- Expert.CompletionWizard.pas          # Completion trigger
|   |-- Expert.CompletionPopup.pas           # Owner-drawn popup
|   |-- Expert.ExtractMethod.pas             # Extract-method logic
|   |-- Expert.ExtractMethodDialog.pas       # Progress / preview dialog
|   |-- Expert.SelectionValidator.pas        # Pascal tokenizer + validation
|   |-- Expert.FindReferencesWizard.pas      # Find-references logic
|   |-- Expert.FindReferencesDialog.pas      # Results dialog with list view
|   |-- Expert.ImplementationFinder.pas      # Shared impl finder (rename + find-impl)
|   |-- Expert.FindImplementationsWizard.pas # Find-implementations wizard
|   |-- Expert.SignatureCheck.pas            # Signature collection / normalization
|   |-- Expert.SignatureCheckDialog.pas      # Align-signature dialog
|   |-- Expert.SignatureCheckWizard.pas      # Align-signature wizard
|   |-- Expert.WithScanner.pas               # Tokenizer for `with` occurrences
|   |-- Expert.WithRewriter.pas              # Per-occurrence inline-var + prefix rewrite
|   |-- Expert.WithRefactorDialog.pas        # Review dialog (Diff / Debug tabs)
|   |-- Expert.WithRefactorWizard.pas        # Project-wide remove-with orchestrator
|   |-- Expert.UnitReferencesWizard.pas      # Find-unit-references wizard
|   |-- Expert.UnitReferencesDialog.pas      # Find-unit-references dialog
|   |-- Expert.UnitRenameWatcher.pas         # Catches IDE unit-renames -> rename wizard
|   |-- Expert.MoveToUnit.pas                # Move-identifier engine
|   |-- Expert.MoveToUnitDialog.pas          # Target-unit picker dialog
|   |-- Expert.MoveToUnitWizard.pas          # Move-identifier wizard
|   |-- Expert.LspPrewarmer.pas              # IOTAIDENotifier: pre-warms LSP on project open
|   |-- Expert.PluginSettings.pas            # Plugin settings (registry-backed)
|   |-- Expert.OptionsFrame.pas / .dfm       # Tools > Options > Refactoring Light page
|   |-- Expert.OptionsPage.pas               # ToolsAPI options-page registration
|   |-- Expert.ContextMenu.pas               # "Refactoring Light" submenu in the editor popup
|   |-- Expert.Shortcuts.pas                 # User-configurable shortcut storage
|   |-- Expert.KeyBinding.pas                # Shortcut registration
|   |-- Expert.EditorHelper.pas              # ToolsAPI wrappers
|   |-- Expert.LspManager.pas                # LSP client singleton + project indexer
|   |-- Expert.IdeCodeInsight.pas            # IDE-internal code insight wrapper
|   |-- Expert.RestartHint.pas               # PID-based restart hint
|   |-- Lsp.Client.pas                       # LSP client (asynchronous)
|   |-- Lsp.JsonRpc.pas                      # JSON-RPC 2.0 transport
|   |-- Lsp.Protocol.pas                     # LSP data types
|   |-- Lsp.Uri.pas                          # URI <-> path
|   |-- Rename.WorkspaceEdit.pas             # WorkspaceEdit apply helper
|   `-- Delphi.FileEncoding.pas              # BOM detection
|
|-- Packages/                                # Delphi project files
|   |-- DelphiRefactoringLight.dpk           # Package source
|   |-- DelphiRefactoringLight.dproj         # Project file
|   `-- DelphiRefactoringLight.res           # Resources
|
|-- dih/                                     # Delphi Install Helper tool
|   |-- delinst.dpr / .dproj / .res
|   |-- DIH.*.pas                            # Engine, builder, registry, ...
|   |-- builddih.cmd                         # builds delinst.exe
|   `-- closedialog.ps1                      # helper script for bds.exe
|
|-- DelphiRefactoringLight.xml               # DIH configuration
|-- install.cmd                              # Install via DIH
|-- uninstall.cmd                            # Uninstall via DIH
`-- rebuild.cmd                              # Build only
```

## Architecture Notes

- **Singleton LSP client** (`TLspManager`): one `DelphiLSP.exe` instance shared across all requests. LSP executable path is resolved from the registry via `IOTAServices.GetBaseRegistryKey`, with a fallback to `GetRootDirectory` and a last-resort hard-coded path.
- **Single-process LSP mode** (not `serverType: controller`): the package sends minimal `initializationOptions` and a full set of client capabilities (hover, definition, documentSymbol, publishDiagnostics with `tagSupport`, ...). An earlier release tried to mirror the IDE's `serverType: controller` + `agentCount: 2` setup &mdash; this is what RAD Studio itself uses to get inactive-region diagnostics &mdash; but DelphiLSP's `Agent0`/`Agent1` sub-processes terminate within seconds when spawned outside `BDS.exe` (verified by replaying the IDE's exact byte sequence into a fresh `DelphiLsp_real.exe`: every `textDocument/hover` came back as `-32800 "Request removed"` or `-32603 "Internal server error"`). Single-process mode trades reduced diagnostic coverage for reliable hover/definition, which most refactorings depend on.
- **Sequential requests**: DelphiLSP does not tolerate parallel requests &mdash; the client uses `BatchSize = 1`.
- **`didChange` after `didOpen`**: `RefreshDocument` follows `didOpen` with a `didChange` notification (version 2, full document text). The IDE proxy log shows DelphiLSP only starts diagnostic analysis after the first `didChange`; without it, `documentSymbol` and `hover` time out on cold files.
- **Project-open LSP pre-warmer** (`TLspPrewarmer`): an `IOTAIDENotifier` listens for `ofnFileOpened` of `.dproj`/`.dpr` files and kicks off a background thread (priority `THREAD_PRIORITY_BELOW_NORMAL`) that calls `EnsureProjectIndexed`. By the time the user triggers their first refactoring, DelphiLSP usually has the project fully analysed. Opt out via *Tools &rarr; Options &rarr; Refactoring Light &rarr; "Pre-warm DelphiLSP ..."*.
- **Per-file diagnostic wait**: refactoring wizards that depend on diagnostics (currently the *Remove with* wizard, for inactive-region detection) explicitly block until `publishDiagnostics` for the file in question has arrived (up to 30 s timeout per file). Idempotent: files that already have diagnostics return instantly. When no diagnostics arrive within the window, affected occurrences are reported as *"LSP no diagnostics &mdash; skipped (dead-code unknown)"* rather than rewritten on guesswork.
- **Inactive `{$IFDEF}`-region tracking**: `TLspClient` parses every `publishDiagnostics` notification, filters for `source = "DelphiLSP"` + `tag = Unnecessary` (or `code = H2655`/`H2656`), and stores the resulting line-range table per file. `IsLineInactive(file, line)` is a direct lookup. `HasReceivedDiagnostics(file)` tells callers whether a `False` from `IsLineInactive` means "verified active" or "no data".
- **`IOTAEditWriter`** instead of `InsertText`: byte-precise edits without IDE auto-indent interference.
- **`Module.Refresh(False)`** instead of `True`: reloads the form module without discarding in-editor changes.
- **PID-based restart hint**: a marker file in `%TEMP%` stores the process ID; if it matches the current IDE, the package was re-installed during the running session and a restart hint is shown.

## AI Disclosure

This project was developed with the assistance of [Claude](https://claude.ai) (Anthropic). The architecture, design decisions, requirements, and quality assurance were provided by the human author. The AI assisted with code generation and documentation.

## License

This project is licensed under the [Mozilla Public License 2.0](https://mozilla.org/MPL/2.0/).
