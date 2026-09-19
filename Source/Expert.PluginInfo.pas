(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.PluginInfo;

// The plugin on the IDE's splash screen and in Help > About (IDE-only).
//
// SPLASH: SplashScreenServices is the FIRST service of the IDE - available
// while packages load, before BorlandIDEServices exists. So the entry is
// added from this unit's INITIALIZATION section; Register would be too late
// on a normal start (the splash is gone by then). A package installed into
// a running IDE finds SplashScreenServices = nil and simply skips it.
//
// ABOUT BOX: needs BorlandIDEServices, so it is added from Register
// (RegisterAboutBox) and MUST be removed before the BPL unloads.
//
// The 24x24 image is drawn at runtime (no resource to maintain); the IDE
// takes the LOWER-LEFT pixel as the transparent colour, so the picture keeps
// that corner free.

interface

procedure RegisterAboutBox;
procedure UnregisterAboutBox;

implementation

uses
  Winapi.Windows, System.SysUtils, System.Types, Vcl.Graphics, ToolsAPI,
  Expert.Version;

var
  GLogo: TBitmap = nil;
  GAboutIndex: Integer = -1;

function Logo: TBitmap;
const
  Transparent = clFuchsia;
begin
  if GLogo <> nil then Exit(GLogo);
  GLogo := TBitmap.Create;
  GLogo.PixelFormat := pf24bit;
  GLogo.SetSize(24, 24);
  with GLogo.Canvas do
  begin
    Brush.Color := Transparent;
    FillRect(Rect(0, 0, 24, 24));
    // rounded tile, kept one pixel off the edges (the lower-left pixel
    // must stay transparent)
    Brush.Color := RGB(0, 102, 204);
    Pen.Color := RGB(0, 72, 150);
    RoundRect(1, 1, 23, 23, 7, 7);
    Font.Name := 'Segoe UI';
    Font.Style := [fsBold];
    Font.Height := -11;
    Font.Color := clWhite;
    Font.Quality := fqAntialiased;   // no ClearType colour fringes
    Brush.Style := bsClear;
    var S := 'RL';
    var Sz := TextExtent(S);
    TextOut((24 - Sz.cx) div 2, (24 - Sz.cy) div 2, S);
  end;
  Result := GLogo;
end;

function AboutText: string;
begin
  Result := PluginName + ' ' + PluginVersion + sLineBreak + sLineBreak +
    'Refactoring and code-insight helpers on top of DelphiLSP: rename (incl. ' +
    'form files), find references / implementations, quick fixes (missing ' +
    'units, undeclared variables, ...), uses cleanup, circular unit ' +
    'references, debug consistency check, live blame (git / svn), code ' +
    'completion with generated event handlers - and an MCP bridge for ' +
    'Claude Code.' + sLineBreak + sLineBreak +
    PluginCopyright + sLineBreak +
    'Mozilla Public License 2.0' + sLineBreak +
    PluginUrl;
end;

procedure RegisterAboutBox;
var
  Svc: IOTAAboutBoxServices;
begin
  if GAboutIndex >= 0 then Exit;
  if Supports(BorlandIDEServices, IOTAAboutBoxServices, Svc) then
    try
      GAboutIndex := Svc.AddPluginInfo(PluginName + ' ' + PluginVersion,
        AboutText, Logo.Handle, False, 'Open Source (MPL 2.0)');
    except
      GAboutIndex := -1;
    end;
end;

procedure UnregisterAboutBox;
var
  Svc: IOTAAboutBoxServices;
begin
  if GAboutIndex < 0 then Exit;
  try
    if Supports(BorlandIDEServices, IOTAAboutBoxServices, Svc) then
      Svc.RemovePluginInfo(GAboutIndex);
  except
  end;
  GAboutIndex := -1;
end;

initialization
  if SplashScreenServices <> nil then
    try
      SplashScreenServices.AddPluginBitmap(PluginName + ' ' + PluginVersion,
        Logo.Handle, False, 'Open Source');
    except
      // a splash entry is decoration - never let it stop the package
    end;

finalization
  UnregisterAboutBox;
  FreeAndNil(GLogo);

end.
