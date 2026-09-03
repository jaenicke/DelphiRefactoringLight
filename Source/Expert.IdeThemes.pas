(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IdeThemes;

interface

uses
  System.UITypes,
  Vcl.Controls,
  Vcl.Forms;

//function IdeThemesEnabled: Boolean;
//function IsDarkMode: Boolean;
function GetThemedColor(AColor: TColor): TColor;

/// <summary>Perceived brightness of AColor, 0 (black) to 255 (white).</summary>
function ColorLuminance(AColor: TColor): Integer;

/// <summary>Colours for the auto-fix hint, derived from the theme's info
///  colours. The hint sits ON the editor, so it must not blend into it:
///  the background is pushed AWAY from the theme's brightness and gets a
///  clearly visible border. Pure function of its inputs - the production
///  path passes the themed clInfoBk/clInfoText, a preview can pass
///  anything.</summary>
procedure ComputeHintColors(ABaseBack, ABaseText: TColor;
  out ABack, ABorder, AText: TColor);
procedure EnableThemes(AForm: TCustomForm);

implementation

uses
  {$IFNDEF STANDALONE_BUILD} ToolsApi, {$ENDIF}
  System.SysUtils, System.Math, Vcl.Graphics;

function ColorLuminance(AColor: TColor): Integer;
var
  RGBVal: Cardinal;
begin
  RGBVal := ColorToRGB(AColor);
  // Rec. 601 weights - good enough to tell a dark theme from a light one.
  Result := (30 * (RGBVal and $FF) + 59 * ((RGBVal shr 8) and $FF) +
             11 * ((RGBVal shr 16) and $FF)) div 100;
end;

// Moves each channel towards white (positive) or black (negative).
function ShiftColor(AColor: TColor; ADelta: Integer): TColor;
var
  RGBVal: Cardinal;
  R, G, B: Integer;
begin
  RGBVal := ColorToRGB(AColor);
  R := EnsureRange(Integer(RGBVal and $FF) + ADelta, 0, 255);
  G := EnsureRange(Integer((RGBVal shr 8) and $FF) + ADelta, 0, 255);
  B := EnsureRange(Integer((RGBVal shr 16) and $FF) + ADelta, 0, 255);
  Result := TColor(R or (G shl 8) or (B shl 16));
end;

procedure ComputeHintColors(ABaseBack, ABaseText: TColor;
  out ABack, ABorder, AText: TColor);
var
  Lum: Integer;
begin
  Lum := ColorLuminance(ABaseBack);
  if Lum < 128 then
  begin
    // Dark theme: lift the panel off the editor and use a bright border.
    ABack := ShiftColor(ABaseBack, 22);
    ABorder := ShiftColor(ABaseBack, 80);
  end
  else
  begin
    // Light theme: keep the familiar pale info colour, darken the border
    // enough to read as a frame rather than a smudge.
    ABack := ABaseBack;
    ABorder := ShiftColor(ABaseBack, -90);
  end;
  AText := ABaseText;
end;

{$IFDEF STANDALONE_BUILD}
// In standalone, theming is a no-op: the IDE theme service is not
// available and our VCL forms use their own colors.
function GetThemedColor(AColor: TColor): TColor;
begin Result := AColor; end;
procedure EnableThemes(AForm: TCustomForm);
begin end;
{$ELSE}

function IdeThemesEnabled: Boolean;
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    Result := Service.IDEThemingEnabled
  end
  else
    Result := False;
end;

function IsDarkMode: Boolean;
var
  Service: IOTAIDEThemingServices;
begin
  Result := False;
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if Service.IDEThemingEnabled then
    begin
      Result := Service.ActiveTheme.Contains('Dark', True);
    end;
  end;
end;

procedure EnableThemes(AForm: TCustomForm);
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if Service.IDEThemingEnabled then
    begin
      Service.RegisterFormClass(TCustomFormClass(AForm.ClassType));
      Service.ApplyTheme(AForm);
    end;
  end;
end;


function GetThemedColor(AColor: TColor): TColor;
var
  Service: IOTAIDEThemingServices;
begin
  if Supports(BorlandIDEServices, IOTAIDEThemingServices, Service) then
  begin
    if (Service.IDEThemingEnabled) and Assigned(Service.StyleServices) then
      Result := Service.StyleServices.GetSystemColor(AColor)
    else
      Result := AColor;
  end
  else
    Result := AColor;
end;
{$ENDIF}

end.
