# Delphi Refactoring Light

WARNING: **This is an experimental preview version. Please use with caution. The code completion is meant as an experiment and will most likely be removed later.**

A design-time package for **Delphi 13** that connects to the built-in Delphi Language Server (`DelphiLSP.exe`) to provide six refactoring features directly in the editor:

| Shortcut           | Feature                                                        |
|--------------------|----------------------------------------------------------------|
| `Ctrl+Alt+Shift+R`     | **Rename** &mdash; rename an identifier project-wide, semantically (including interface implementations) |
| `Ctrl+Alt+Shift+U`     | **Find References** &mdash; list every occurrence of an identifier |
| `Ctrl+Alt+Shift+I`     | **Find Implementations** &mdash; list all class implementations of an interface/virtual method |
| `Ctrl+Alt+Shift+Space` | **Code Completion** &mdash; suggestions via DelphiLSP          |
| `Ctrl+Alt+Shift+M`     | **Extract Method** &mdash; move the selected block into a new method |
| `Ctrl+Alt+Shift+A`     | **Align method signature** &mdash; compare a method's class/interface declaration with its implementation and highlight mismatches |

Unlike purely text-based tools, this package uses the actual LSP requests `textDocument/rename`, `textDocument/references`, `textDocument/implementation`, `textDocument/completion`, `textDocument/hover`, and `textDocument/definition`. Identifiers that happen to share a name but belong to different symbols are cleanly distinguished.

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
|   |-- Expert.KeyBinding.pas                # Shortcut registration
|   |-- Expert.EditorHelper.pas              # ToolsAPI wrappers
|   |-- Expert.LspManager.pas                # LSP client singleton
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
- **Sequential requests**: DelphiLSP does not tolerate parallel requests &mdash; the client uses `BatchSize = 1`.
- **`IOTAEditWriter`** instead of `InsertText`: byte-precise edits without IDE auto-indent interference.
- **`Module.Refresh(False)`** instead of `True`: reloads the form module without discarding in-editor changes.
- **PID-based restart hint**: a marker file in `%TEMP%` stores the process ID; if it matches the current IDE, the package was re-installed during the running session and a restart hint is shown.

## AI Disclosure

This project was developed with the assistance of [Claude](https://claude.ai) (Anthropic). The architecture, design decisions, requirements, and quality assurance were provided by the human author. The AI assisted with code generation and documentation.

## License

This project is licensed under the [Mozilla Public License 2.0](https://mozilla.org/MPL/2.0/).
