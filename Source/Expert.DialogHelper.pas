(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DialogHelper;

{
  Shared helpers for the plugin's dialogs:

    - PrepareDialog: ties the dialog to the IDE main window (so it stays
      on the same monitor instead of jumping to a secondary one) and
      applies the active IDE VCL style (dark / light).

    - RegisterDialogClass: registers the form class with the IDE's
      theming services so newly created instances pick up the theme on
      construction.

  The theming uses IOTAIDEThemingServices (Delphi 10.4+); if the IDE
  doesn't expose it (older versions) or theming is disabled, the
  functions degrade gracefully.

  The monitor-jumping issue is fixed by setting PopupMode := pmExplicit
  and PopupParent := <main IDE window>. Without this, a non-owned form
  shown via TForm.Show may be placed on whichever monitor the OS
  currently considers "default", which on multi-monitor setups isn't
  always the one the IDE is running on.
}

interface

uses
  System.Classes, Vcl.Forms, Vcl.Controls;

/// <summary>Registers AClass with the IDE theming service so new
///  instances pick up the active theme. Safe to call multiple times;
///  the theming service deduplicates internally.</summary>
procedure RegisterDialogClass(AClass: TCustomFormClass);

/// <summary>Wires AForm to the IDE main window (PopupParent) and
///  applies the active IDE theme. Call from the dialog's constructor
///  AFTER all child controls have been created so the theme reaches
///  them as well.</summary>
procedure PrepareDialog(AForm: TForm; AOwner: TComponent);

implementation

uses
  System.SysUtils {$IFNDEF STANDALONE_BUILD}, ToolsAPI {$ENDIF};

{$IFNDEF STANDALONE_BUILD}
function ThemingServices: IOTAIDEThemingServices;
begin
  if not Supports(BorlandIDEServices, IOTAIDEThemingServices, Result) then
    Result := nil;
end;
{$ENDIF}

procedure RegisterDialogClass(AClass: TCustomFormClass);
{$IFNDEF STANDALONE_BUILD}
var
  TS: IOTAIDEThemingServices;
{$ENDIF}
begin
{$IFNDEF STANDALONE_BUILD}
  TS := ThemingServices;
  if (TS <> nil) and TS.IDEThemingEnabled then
  try
    TS.RegisterFormClass(AClass);
  except
    // Defensive: don't let a theming bug break dialog creation.
  end;
{$ENDIF}
end;

procedure PrepareDialog(AForm: TForm; AOwner: TComponent);
var
  Anchor: TCustomForm;
  {$IFNDEF STANDALONE_BUILD}
  TS: IOTAIDEThemingServices;
  {$ENDIF}
begin
  // Anchor the dialog to the IDE main window. This both fixes the
  // multi-monitor jump and makes Windows treat the dialog as a child
  // of the IDE for task switching / focus purposes.
  if AOwner is TCustomForm then
    Anchor := TCustomForm(AOwner)
  else
    Anchor := Application.MainForm;

  if Anchor <> nil then
  begin
    AForm.PopupMode := pmExplicit;
    AForm.PopupParent := Anchor;
  end;

{$IFNDEF STANDALONE_BUILD}
  // Apply the IDE theme to the form and all its child controls.
  TS := ThemingServices;
  if (TS <> nil) and TS.IDEThemingEnabled then
  try
    TS.ApplyTheme(AForm);
  except
    // See RegisterDialogClass.
  end;
{$ENDIF}
end;

end.
