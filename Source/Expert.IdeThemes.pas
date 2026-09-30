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
  System.Classes,
  System.Types,
  System.UITypes,
  Winapi.Windows,
  Winapi.Messages,
  Vcl.Controls,
  Vcl.ComCtrls,
  Vcl.Forms;

/// <summary>True while the IDE really is theming. Everything below leaves a
///  control untouched when it is False - without a theme the native look is
///  the right one.</summary>
function IdeThemesEnabled: Boolean;
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

type
  /// <summary>How a status row reads at a glance: good / waiting / broken /
  ///  nothing special (user request 2026-09-30 - the status window was a
  ///  wall of equal-looking text).</summary>
  TStatusLevel = (slNeutral, slGood, slWait, slBad);

/// <summary>A readable colour for ALEVEL on ABACKGROUND. Derived, not fixed:
///  the saturated green / amber / red of a light theme is unreadable on a
///  dark editor background, so there the hue is lightened until it carries -
///  the same approach ComputeHintColors takes for the hint window.
///  slNeutral answers ADefaultText unchanged.</summary>
function StatusLevelColor(ALevel: TStatusLevel; ABackground,
  ADefaultText: TColor): TColor;

procedure EnableThemes(AForm: TCustomForm);

type
  /// <summary>A report list view that paints its BODY and its COLUMN HEADER
  ///  in colours it is given.
  ///  Why it exists (measured on the status window, 2026-09-30): the IDE's
  ///  theming service does not reach the children of a FRAME - the list stayed
  ///  in system colours while the custom-drawn cells were themed, so what used
  ///  to be merely white became a white/dark patchwork. And the header is a
  ///  separate native control that no Color property reaches at all.
  ///  The colours are PROPERTIES, not calls into the theming service, so this
  ///  painting can be rendered and looked at outside the IDE.</summary>
  TThemedListView = class(TListView)
  private
    FHeaderColor: TColor;
    FHeaderTextColor: TColor;
    FHeaderLineColor: TColor;
    FThemed: Boolean;
    procedure PaintHeaderItem(ADC: HDC; AIndex: Integer; const ARect: TRect;
      AHot: Boolean);
    procedure WMNotify(var AMessage: TWMNotify); message WM_NOTIFY;
  public
    /// <summary>Body and header colours, the header derived from ABack so it
    ///  stays a visible but quiet strip in either theme. Assigns only what
    ///  really differs and answers whether anything changed, so a poll tick may
    ///  call it without repainting every second. Until it is called the control
    ///  behaves exactly like a plain TListView.</summary>
    function ApplyColors(ABack, AText: TColor): Boolean;
    property Themed: Boolean read FThemed;
  end;

/// <summary>Header background and separator line for a list whose body is
///  ABACK. Derived from the body rather than taken from the theme, because a
///  theme's button face can be the very same colour as its window colour - and
///  then the header would not read as a header at all. Pure.</summary>
procedure ListHeaderColors(ABack: TColor; out AHeader, ALine: TColor);

/// <summary>Gives ALV the IDE's colours; for a TThemedListView that includes
///  the column header. Nothing happens while the IDE is not theming.</summary>
function ApplyThemeToListView(ALV: TListView): Boolean;

/// <summary>The same for a whole control tree (a dockable frame with its list,
///  memo and labels). Data areas get the window colour, the frame around them
///  the button face - the split the IDE uses for its own panes.</summary>
function ApplyThemeToControls(AParent: TWinControl): Boolean;

implementation

uses
  {$IFNDEF STANDALONE_BUILD} ToolsApi, {$ENDIF}
  System.SysUtils, System.Math, Winapi.CommCtrl,
  Vcl.Graphics, Vcl.StdCtrls, Vcl.ExtCtrls;

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

function StatusLevelColor(ALevel: TStatusLevel; ABackground,
  ADefaultText: TColor): TColor;
const
  // TColor is $00BBGGRR. Tones for a LIGHT background: dark enough to read
  // on white without shouting.
  GoodLight = TColor($00107010);   // RGB(16, 112, 16)
  WaitLight = TColor($000070B0);   // RGB(176, 112, 0) - amber
  BadLight  = TColor($000000C0);   // RGB(192, 0, 0)
  // ... and for a DARK background: the same hues, lightened until they carry.
  GoodDark  = TColor($0080D880);   // RGB(128, 216, 128)
  WaitDark  = TColor($0060C8F0);   // RGB(240, 200, 96)
  BadDark   = TColor($008080FF);   // RGB(255, 128, 128)
begin
  if ALevel = slNeutral then Exit(ADefaultText);
  if ColorLuminance(ABackground) < 128 then
    case ALevel of
      slGood: Result := GoodDark;
      slWait: Result := WaitDark;
    else
      Result := BadDark;
    end
  else
    case ALevel of
      slGood: Result := GoodLight;
      slWait: Result := WaitLight;
    else
      Result := BadLight;
    end;
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

{ TThemedListView }

procedure ListHeaderColors(ABack: TColor; out AHeader, ALine: TColor);
begin
  // The header must differ from the rows to read as a header, but only just.
  if ColorLuminance(ABack) < 128 then
  begin
    AHeader := ShiftColor(ABack, 18);
    ALine := ShiftColor(ABack, 45);
  end
  else
  begin
    AHeader := ShiftColor(ABack, -14);
    ALine := ShiftColor(ABack, -55);
  end;
end;

function TThemedListView.ApplyColors(ABack, AText: TColor): Boolean;
var
  Head, Line: TColor;
begin
  ListHeaderColors(ABack, Head, Line);
  Result := (not FThemed) or (Color <> ABack) or (Font.Color <> AText) or
    (FHeaderColor <> Head);
  if not Result then Exit;
  FThemed := True;
  FHeaderColor := Head;
  FHeaderTextColor := AText;
  FHeaderLineColor := Line;
  Color := ABack;
  Font.Color := AText;
  // the native grid lines are a fixed light grey - on a dark background they
  // draw a bright cage around every cell, so they go
  GridLines := ColorLuminance(ABack) >= 128;
  if HandleAllocated then Invalidate;
end;

procedure TThemedListView.PaintHeaderItem(ADC: HDC; AIndex: Integer;
  const ARect: TRect; AHot: Boolean);
var
  Cv: TCanvas;
  R: TRect;
  Item: THDItem;
  Buf: array[0..255] of Char;
  Txt: string;
  Asc, HasArrow: Boolean;
begin
  Cv := TCanvas.Create;
  try
    Cv.Handle := ADC;
    R := ARect;
    if AHot then
      Cv.Brush.Color := ShiftColor(FHeaderColor, 12)
    else
      Cv.Brush.Color := FHeaderColor;
    Cv.Brush.Style := bsSolid;
    Cv.FillRect(R);
    Cv.Pen.Color := FHeaderLineColor;
    Cv.MoveTo(R.Right - 1, R.Top + 3);        // column separator
    Cv.LineTo(R.Right - 1, R.Bottom - 3);
    Cv.MoveTo(R.Left, R.Bottom - 1);          // baseline under the whole header
    Cv.LineTo(R.Right, R.Bottom - 1);

    // Caption and sort arrow come from the HEADER, not from Columns[]: the
    // arrow is set through HDITEM.fmt (Expert.ListViewSort), and a repaint can
    // arrive while the column list is being rebuilt.
    Txt := '';
    HasArrow := False;
    Asc := True;
    FillChar(Item, SizeOf(Item), 0);
    Item.Mask := HDI_TEXT or HDI_FORMAT;
    Item.pszText := @Buf[0];
    Item.cchTextMax := Length(Buf);
    if Header_GetItem(ListView_GetHeader(Handle), AIndex, Item) then
    begin
      Txt := Item.pszText;
      HasArrow := (Item.fmt and (HDF_SORTUP or HDF_SORTDOWN)) <> 0;
      Asc := (Item.fmt and HDF_SORTUP) <> 0;
    end;

    R.Left := R.Left + 6;
    R.Right := R.Right - 6;
    if HasArrow and (R.Right - R.Left > 16) then
    begin
      var Cx := R.Right - 5;
      var Cy := (R.Top + R.Bottom) div 2;
      Cv.Brush.Color := FHeaderTextColor;
      Cv.Pen.Color := FHeaderTextColor;
      if Asc then
        Cv.Polygon([Point(Cx - 4, Cy + 2), Point(Cx + 4, Cy + 2), Point(Cx, Cy - 3)])
      else
        Cv.Polygon([Point(Cx - 4, Cy - 2), Point(Cx + 4, Cy - 2), Point(Cx, Cy + 3)]);
      R.Right := R.Right - 14;
    end;
    Cv.Font := Font;
    Cv.Font.Color := FHeaderTextColor;
    Cv.Brush.Style := bsClear;
    DrawText(Cv.Handle, PChar(Txt), Length(Txt), R,
      DT_LEFT or DT_VCENTER or DT_SINGLELINE or DT_END_ELLIPSIS or DT_NOPREFIX);
  finally
    Cv.Handle := 0;
    Cv.Free;
  end;
end;

procedure TThemedListView.WMNotify(var AMessage: TWMNotify);
var
  Cd: PNMCustomDraw;
begin
  inherited;
  if not FThemed then Exit;
  if (AMessage.NMHdr = nil) or (AMessage.NMHdr.code <> NM_CUSTOMDRAW) then Exit;
  if not HandleAllocated then Exit;
  if AMessage.NMHdr.hwndFrom <> ListView_GetHeader(Handle) then Exit;
  Cd := PNMCustomDraw(AMessage.NMHdr);
  case Cd.dwDrawStage of
    CDDS_PREPAINT:
      AMessage.Result := CDRF_NOTIFYITEMDRAW or CDRF_NOTIFYPOSTPAINT;
    CDDS_POSTPAINT:
      begin
        // The strip BEHIND the last column is no item, so no item draw ever
        // reaches it and the default painting leaves it in the system colour -
        // a white block in the corner of a dark header (seen in the render,
        // not in the code). Filling it at PREPAINT does not help: the default
        // painting runs afterwards and paints over it. So it happens LAST.
        var HW := AMessage.NMHdr.hwndFrom;
        var HR, IR: TRect;
        // qualified: inside a TControl "GetClientRect" is the class's own
        if Winapi.Windows.GetClientRect(HW, HR) then
        begin
          var X := HR.Left;
          var Last := Header_GetItemCount(HW) - 1;
          if (Last >= 0) and Header_GetItemRect(HW, Last, @IR) then X := IR.Right;
          if X < HR.Right then
          begin
            var Cv := TCanvas.Create;
            try
              Cv.Handle := Cd.hdc;
              Cv.Brush.Color := FHeaderColor;
              Cv.Brush.Style := bsSolid;
              Cv.FillRect(Rect(X, HR.Top, HR.Right, HR.Bottom));
              Cv.Pen.Color := FHeaderLineColor;
              Cv.MoveTo(X, HR.Bottom - 1);
              Cv.LineTo(HR.Right, HR.Bottom - 1);
            finally
              Cv.Handle := 0;
              Cv.Free;
            end;
          end;
        end;
        AMessage.Result := CDRF_DODEFAULT;
      end;
    CDDS_ITEMPREPAINT:
      begin
        PaintHeaderItem(Cd.hdc, Cd.dwItemSpec, Cd.rc,
          (Cd.uItemState and CDIS_HOT) <> 0);
        AMessage.Result := CDRF_SKIPDEFAULT;
      end;
  end;
end;

type
  /// <summary>Color and Font are protected on TControl - every control we
  ///  theme here has them, whatever it publishes.</summary>
  TControlColors = class(TControl)
  public
    property Color;
    property Font;
  end;

function TouchControl(AControl: TControl; ABack, AText: TColor): Boolean;
begin
  Result := False;
  if TControlColors(AControl).Color <> ABack then
  begin
    TControlColors(AControl).Color := ABack;
    Result := True;
  end;
  if TControlColors(AControl).Font.Color <> AText then
  begin
    TControlColors(AControl).Font.Color := AText;
    Result := True;
  end;
end;

{$IFDEF STANDALONE_BUILD}
// In standalone, theming is a no-op: the IDE theme service is not
// available and our VCL forms use their own colors.
function IdeThemesEnabled: Boolean;
begin Result := False; end;
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
{$IF CompilerVersion >= 37}
      Result := Service.ActiveTheme.Contains('Dark', True);
{$ELSE}
      Result := UpperCase(Service.ActiveTheme).Contains('DARK');
{$IFEND}
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

function ApplyThemeToListView(ALV: TListView): Boolean;
begin
  Result := False;
  // Without a theme the NATIVE look is the right one - do not repaint a
  // perfectly good light-mode list in colours of our own.
  if (ALV = nil) or not IdeThemesEnabled then Exit;
  if ALV is TThemedListView then
    Result := TThemedListView(ALV).ApplyColors(
      GetThemedColor(clWindow), GetThemedColor(clWindowText))
  else
    Result := TouchControl(ALV, GetThemedColor(clWindow),
      GetThemedColor(clWindowText));
end;

function ApplyThemeToControls(AParent: TWinControl): Boolean;
var
  Back, Text, Face: TColor;
  I: Integer;
  C: TControl;
begin
  Result := False;
  if (AParent = nil) or not IdeThemesEnabled then Exit;
  Back := GetThemedColor(clWindow);
  Text := GetThemedColor(clWindowText);
  Face := GetThemedColor(clBtnFace);
  if TouchControl(AParent, Face, Text) then Result := True;
  for I := 0 to AParent.ControlCount - 1 do
  begin
    C := AParent.Controls[I];
    if C is TListView then
    begin
      if ApplyThemeToListView(TListView(C)) then Result := True;
    end
    // data areas keep the window colour, everything around them the face
    else if (C is TCustomEdit) or (C is TCustomListBox) or
            (C is TCustomComboBox) then
    begin
      if TouchControl(C, Back, Text) then Result := True;
    end
    else if (C is TCustomLabel) or (C is TSplitter) then
    begin
      if TouchControl(C, Face, Text) then Result := True;
    end
    else if C is TWinControl then
    begin
      if ApplyThemeToControls(TWinControl(C)) then Result := True;
    end;
  end;
end;

end.
