(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.BlameGutter;

{
  Live blame in the editor (IDE only).

  Two things are painted, both through the OFFICIAL painting API
  (INTACodeEditorServices.AddEditorEventsNotifier):

  * GUTTER: a narrow age stripe per line at the right edge of the gutter
    (fresh = saturated, old = faint), plus the revision and author as
    TEXT in a column of our own. The painting API cannot create space,
    but IOTABufferOptions.LeftGutterWidth can - see EnsureGutterWidth,
    including the rules for sharing the gutter with other add-ons.
    Column width 0 = stripe only, do not touch the gutter at all.
  * CARET LINE: the full annotation, dimmed, behind the end of the code:
    "Author - 3 day(s) ago - summary". That is where a blame is actually
    read, and it costs no width anywhere else.

  RULES OBEYED HERE (see CLAUDE.md):
  * The paint handler READS state only - no git, no file I/O, no ToolsAPI
    round trips beyond the context it is handed.
  * The data is fetched by a 1 s timer tick (WM_TIMER), never from a
    notifier or a queue callback.
  * While the buffer is MODIFIED the blame is REMAPPED, not hidden: the
    lines that survived the edit keep their author, everything inside the
    edited span shows as "not committed". The mapping is a prefix/suffix
    comparison against the file on disk - deliberately coarse, because a
    wrong author is worse than an honest "you touched this".
}

interface

/// <summary>Installs the editor notifier. Safe to call twice.</summary>
procedure InstallBlameGutter;
/// <summary>MUST run before the BPL unloads.</summary>
procedure UninstallBlameGutter;

/// <summary>On/off (persisted by the caller). Off = nothing is painted
///  and no git process is started.</summary>
function BlameEnabled: Boolean;
procedure SetBlameEnabled(AValue: Boolean);

/// <summary>Status row: what the feature is currently doing.</summary>
function BlameGutterStatus: string;

/// <summary>Re-reads the options (on/off, column width, what to show) and
///  applies them at once - the width is restored first, so switching the
///  column off never leaves a widened gutter behind.</summary>
procedure ApplyBlameSettings;

implementation

uses
  Expert.ResourceMonitor,
  System.SysUtils, System.Classes, System.Math, System.DateUtils,
  System.Generics.Collections,
  System.IOUtils, System.Hash,
  Winapi.Windows, Vcl.Graphics, Vcl.ExtCtrls, Vcl.Forms,
  ToolsAPI, ToolsAPI.Editor,
  Expert.EditorHelperIntf, Expert.VcsBlame, Expert.IdeThemes,
  Expert.PluginSettings;

procedure InvalidateTopEditor; forward;

type
  TBlameNotifier = class(TNTACodeEditorNotifier)
  private
    procedure HandlePaintLine(const Rect: TRect; const Stage: TPaintLineStage;
      const BeforeEvent: Boolean; var AllowDefaultPainting: Boolean;
      const Context: INTACodeEditorPaintContext);
    procedure HandlePaintGutter(const Rect: TRect; const Stage: TPaintGutterStage;
      const BeforeEvent: Boolean; var AllowDefaultPainting: Boolean;
      const Context: INTACodeEditorPaintContext);
  protected
    function AllowedEvents: TCodeEditorEvents; override;
    function AllowedLineStages: TPaintLineStages; override;
    function AllowedGutterStages: TPaintGutterStages; override;
  public
    constructor Create;
  end;

var
  GNotifier: TBlameNotifier = nil;
  GNotifierIndex: Integer = -1;
  GTimer: TTimer = nil;
  GEnabled: Boolean = False;
  // State the painter reads. Filled by the timer tick ONLY.
  GFile: string = '';
  GLines: TBlameLines = nil;
  GDirty: Boolean = False;        // buffer modified (lines remapped)
  // While the buffer is modified the blame is shown REMAPPED: the disk
  // lines that survived keep their author, the edited span shows as
  // uncommitted. GMapped holds that per-buffer-line copy, keyed by the
  // content hash so it is rebuilt only when the text really changed.
  GMapped: TBlameLines = nil;
  GMapHash: Integer = 0;
  GMapFile: string = '';
  // The disk side of the comparison, cached per (file, mtime).
  GDiskLines: TArray<string> = nil;
  GDiskFile: string = '';
  GDiskStamp: TDateTime = 0;
  GCaretLine: Integer = 0;        // 1-based
  GStatus: string = 'off';
  // "I see nothing in the gutter" must be answerable without guessing:
  // did the paint event arrive, and did we have data for those lines?
  GGutterEvents: Integer = 0;
  GStripesPainted: Integer = 0;
  GTextsPainted: Integer = 0;
  // Last seen gutter geometry - so "the space stays empty" can be READ
  // instead of guessed (the IDE does not document where the extra width
  // from LeftGutterWidth actually appears).
  GLastGutter: TRect;
  GLastData: TRect;
  // Gutter widening: the IDE lets us set the width per BUFFER
  // (IOTABufferOptions.LeftGutterWidth), which is the only way to get
  // room for text - the paint API alone cannot create space.
  GColumnWidth: Integer = 150;   // px reserved for the blame text
  GInfoLevel: Integer = 1;       // 0 = revision, 1 = + author, 2 = + age
  GColumnOffset: Integer = 0;    // px from the data area's left edge
  GSavedWidth: Integer = -1;     // the buffer's own width, for restoring
  // The ORIGINAL width per buffer. Remembered, never re-derived - see
  // EnsureGutterWidth for why (the gutter grew on every switch).
  GOrigWidths: TDictionary<string, Integer> = nil;
  GLastOrig: Integer = -1;
  GRestorePending: Boolean = False;
  GSetWidth: Integer = -1;       // what WE set it to - see RestoreGutterWidth
  GWidenedFile: string = '';

function BlameEnabled: Boolean;
begin
  Result := GEnabled;
end;

function BlameGutterStatus: string;
begin
  if not GEnabled then Exit('off');
  if GDirty then
    Exit(Format('buffer modified - %d line(s) remapped against the file on disk',
      [Length(GMapped)]));
  if Length(GLines) = 0 then Exit(BlameStatus);
  Result := Format(
    '%d line(s) for %s | %d event(s), %d stripe(s), %d text(s) | ' +
    'gutter %d..%d, data %d..%d, column %d px at +%d',
    [Length(GLines), ExtractFileName(GFile), GGutterEvents, GStripesPainted,
     GTextsPainted, GLastGutter.Left, GLastGutter.Right,
     GLastData.Left, GLastData.Right, GColumnWidth, GColumnOffset]);
end;

const
  // Font.Height is NEGATIVE (character height), so a factor < 1 shrinks.
  BlameFontScale = 0.78;
  EMPTYFILE = '';

// Buffer paths come from different ToolsAPI calls - compare normalised.
function NormFile(const AFile: string): string;
begin
  Result := LowerCase(AFile);
end;

// ---------------------------------------------------------------------------
//  Colours
// ---------------------------------------------------------------------------

// Age ramp: fresh = saturated, old = faint. Derived from the THEME's text
// colour so it works in both light and dark (the editor background is not
// available here, but the contrast direction is).
function AgeColor(const AWhen: TDateTime; AUncommitted: Boolean): TColor;
var
  Days: Double;
  T: Double;
  R, G, B: Byte;
begin
  if AUncommitted then Exit(TColor($3CB4FF));      // orange-ish: your edit
  if AWhen <= 0 then Exit(clNone);
  Days := Max(0, DaySpan(Now, AWhen));
  // 0 days -> 1.0, one year and older -> 0.0 (logarithmic: most commits
  // are recent, so a linear ramp would make everything look identical).
  T := 1 - Min(1, Ln(1 + Days) / Ln(1 + 365));
  R := Round(90 + 100 * T);
  G := Round(140 + 60 * T);
  B := Round(90 + 40 * T);
  Result := TColor(B shl 16 or G shl 8 or R);
end;

// APercent of the way from AColor to ATowards.
function Blend(AColor, ATowards: TColor; APercent: Integer): TColor;
var
  C1, C2: LongInt;
  R, G, B: Integer;
begin
  C1 := ColorToRGB(AColor);
  C2 := ColorToRGB(ATowards);
  APercent := EnsureRange(APercent, 0, 100);
  R := (GetRValue(C1) * (100 - APercent) + GetRValue(C2) * APercent) div 100;
  G := (GetGValue(C1) * (100 - APercent) + GetGValue(C2) * APercent) div 100;
  B := (GetBValue(C1) * (100 - APercent) + GetBValue(C2) * APercent) div 100;
  Result := TColor(RGB(R, G, B));
end;

// One light tint per COMMIT: consecutive lines of the same commit form a
// visible block, and neighbouring commits are told apart at a glance.
// A fixed palette (not a hash-to-RGB) keeps the colours pleasant, and the
// heavy blend towards the actual background keeps them a hint, not paint.
const
  CommitHues: array[0..11] of TColor = (
    TColor($4646C8), TColor($46A0C8), TColor($46C88C), TColor($64C846),
    TColor($C8C846), TColor($C89646), TColor($C85A46), TColor($C846A0),
    TColor($A046C8), TColor($6446C8), TColor($46C8C8), TColor($8CC846));

function CommitTint(const AHash: string; ABack: TColor;
  AUncommitted: Boolean): TColor;
var
  H, I: Cardinal;
begin
  // Not committed yet: one recognisable tint of its own rather than the
  // plain background - "this is mine, and it is not in yet". Decided by
  // the FLAG, not by the shape of the hash: git writes all zeros there,
  // svn writes nothing at all.
  if AUncommitted or (AHash = '') then
    Exit(Blend(TColor($3CB4FF), ABack, 86));
  H := 2166136261;                       // FNV-1a: stable across sessions
  for I := 1 to Length(AHash) do
  begin
    H := H xor Cardinal(Ord(AHash[I]));
    H := H * 16777619;
  end;
  // 86% towards the background: readable text stays the main thing.
  Result := Blend(CommitHues[H mod Length(CommitHues)], ABack, 86);
end;

// "Sebastian Jänicke" -> "Sebastian": the gutter has room for one word.
function FirstNamePart(const AAuthor: string): string;
var
  P: Integer;
begin
  Result := Trim(AAuthor);
  P := Pos(' ', Result);
  if P > 1 then Result := Copy(Result, 1, P - 1);
  P := Pos('@', Result);            // an e-mail address as the author
  if P > 1 then Result := Copy(Result, 1, P - 1);
end;

// ---------------------------------------------------------------------------
//  Painting
// ---------------------------------------------------------------------------

constructor TBlameNotifier.Create;
begin
  inherited Create;
  OnEditorPaintLine := HandlePaintLine;
  OnEditorPaintGutter := HandlePaintGutter;
end;

function TBlameNotifier.AllowedEvents: TCodeEditorEvents;
begin
  Result := [cevPaintLineEvents, cevPaintGutterEvents];
end;

function TBlameNotifier.AllowedLineStages: TPaintLineStages;
begin
  Result := [plsEndPaint];
end;

function TBlameNotifier.AllowedGutterStages: TPaintGutterStages;
begin
  // NOT pgsEndPaint: per the ToolsAPI docs that is the end of the whole
  // gutter paint, not a per-line stage, so LineState need not refer to
  // the line we want. These two are per line; painting the same opaque
  // stripe in both is harmless and makes us independent of which one a
  // given IDE build actually raises.
  Result := [pgsAnnotate, pgsLineNumber];
end;

function LineInfo(ALogical1Based: Integer; out AInfo: TBlameLine): Boolean;
var
  Src: TBlameLines;
begin
  if GDirty then Src := GMapped else Src := GLines;
  Result := GEnabled and (ALogical1Based >= 1)
    and (ALogical1Based <= Length(Src));
  if Result then
  begin
    AInfo := Src[ALogical1Based - 1];
    // NOT "has a hash": a line that is not committed yet HAS no revision
    // in svn (git gives it an all-zero hash instead, which is why this
    // only showed up on the svn side - those lines stayed blank). The
    // parser stamps Kind on every entry it emits, so that is the honest
    // test for "we know something about this line".
    Result := AInfo.Kind <> vcsNone;
  end;
end;

procedure TBlameNotifier.HandlePaintGutter(const Rect: TRect;
  const Stage: TPaintGutterStage; const BeforeEvent: Boolean;
  var AllowDefaultPainting: Boolean; const Context: INTACodeEditorPaintContext);
var
  Info: TBlameLine;
  R: TRect;
  C: TColor;
  SavedFontHeight, SavedFontColor, SavedBrush: Integer;
  SavedStyle: TBrushStyle;
  Cv: TCanvas;
begin
  // GDI objects this call leaves behind are booked per subsystem -
  // the status window shows the balance (Expert.ResourceMonitor).
  var GdiG := GdiGuard(gsBlamePaint, 32);
  if BeforeEvent or (Context = nil) then Exit;
  Cv := nil;
  SavedFontHeight := 0;
  SavedFontColor := 0;
  SavedBrush := 0;
  SavedStyle := bsSolid;
  try
    Inc(GGutterEvents);          // diagnosis: does the event arrive at all?
    if Context.LineState = nil then Exit;
    if not LineInfo(Context.LineState.LogicalLineNum, Info) then Exit;
    // THE CANVAS IS SHARED with the IDE's own gutter painting, which
    // continues after us. Whatever we change here has to go back, or the
    // line numbers end up in our (smaller) font.
    Cv := Context.Canvas;
    if Cv <> nil then
    begin
      SavedFontHeight := Cv.Font.Height;
      SavedFontColor := Integer(Cv.Font.Color);
      SavedBrush := Integer(Cv.Brush.Color);
      SavedStyle := Cv.Brush.Style;
    end;
    Inc(GStripesPainted);
    C := AgeColor(Info.AuthorTime, Info.IsUncommitted);
    if C = clNone then Exit;
    R := Context.LineState.GutterRect;

    // TEXT first, into the space the IDE does NOT use itself:
    // GutterLineDataRect is where line numbers and marks go, so whatever
    // lies left of it is ours (that is the room LeftGutterWidth added).
    // WHICH PART OF THE GUTTER IS OURS? Only what we ADDED. The rest
    // belonged to the IDE - and possibly to another add-on - before we
    // touched anything, so claiming all the free space would mean
    // painting over a neighbour. GutterLineDataRect marks where the IDE
    // itself draws (numbers, marks); the free strip beside it is a
    // candidate, and it is then clipped to our own column width.
    var DataRect := Context.LineState.GutterLineDataRect;
    GLastGutter := R;
    GLastData := DataRect;
    var Mine := Max(0, GColumnWidth);
    var TextRect := TRect.Empty;
    // MEASURED on the tester's IDE (the status row was built for exactly
    // this question): GutterRect is the narrow MARKER strip (0..20, where
    // breakpoints and bookmarks sit) and GutterLineDataRect is the WIDE
    // area (20..239) that holds the fold column and the line numbers -
    // and that is also where the width we added shows up. So our column
    // is the LEFT part of the DATA rect; the IDE keeps drawing its
    // numbers right-aligned next to the code, well clear of it.
    if (Mine > 0) and (DataRect.Width >= Mine + GColumnOffset + 24) then
      // The offset lets the column step aside for a neighbour: the
      // bundled Parnassus Navigator draws ITS marks at the very left of
      // this same area, and nothing in the API divides the space up.
      TextRect := TRect.Create(DataRect.Left + GColumnOffset + 2, R.Top,
        DataRect.Left + GColumnOffset + Mine - 4, R.Bottom)
    else if (Mine > 0) and (DataRect.Left - R.Left >= 24) then
      TextRect := TRect.Create(R.Left + 2, R.Top,
        Min(DataRect.Left - 2, R.Left + Mine), R.Bottom)
    else if (Mine > 0) and (R.Right - DataRect.Right >= 24) then
      TextRect := TRect.Create(Max(DataRect.Right + 2, R.Right - Mine), R.Top,
        R.Right - 4, R.Bottom)
    else if (Mine > 0) and (R.Width >= Mine + 16) then
      // FALLBACK - the tester's case: GutterLineDataRect gave no usable
      // free strip (it can cover the whole gutter), yet the gutter IS
      // wider than before because WE widened it. Then our own column is
      // the leftmost GColumnWidth pixels: that is precisely the space
      // that did not exist until we asked for it.
      TextRect := TRect.Create(R.Left + 2, R.Top, R.Left + Mine - 4, R.Bottom);
    if TextRect.Width < 24 then TextRect := TRect.Empty;

    if not TextRect.IsEmpty then
    begin
      var Txt := Info.ShortHash;
      if Info.IsUncommitted then
        Txt := 'not committed'
      else
      begin
        if (GInfoLevel >= 1) and (Info.Author <> '') then
          Txt := Txt + ' ' + FirstNamePart(Info.Author);
        if GInfoLevel >= 2 then
          Txt := Txt + ' ' + HumanAge(Info.AuthorTime, Now);
      end;
      // The BACKGROUND the IDE just painted here is the honest reference
      // for blending - it adapts to the theme without asking which one.
      var BackCol := Context.Canvas.Brush.Color;
      if BackCol = clNone then BackCol := GetThemedColor(clWindow);

      Context.Canvas.Brush.Style := bsSolid;
      Context.Canvas.Brush.Color :=
        CommitTint(Info.Hash, BackCol, Info.IsUncommitted);
      Context.Canvas.FillRect(TextRect);

      // Clearly smaller than the code: the column is an aside, and the
      // smaller face also fits revision AND author into the same width.
      Context.Canvas.Font.Height :=
        Min(-9, Round(Context.Canvas.Font.Height * BlameFontScale));
      Context.Canvas.Brush.Style := bsClear;
      Context.Canvas.Font.Color := Blend(C, GetThemedColor(clWindowText), 45);
      // TextRect clips and ellipsises, so a long name can never bleed
      // into the code area.
      Context.Canvas.TextRect(TextRect, Txt,
        [tfSingleLine, tfVerticalCenter, tfEndEllipsis]);
      Inc(GTextsPainted);
    end;

    // A 3 px stripe hugging the code, so the age stays readable even when
    // the gutter is too narrow for text.
    R.Left := Max(R.Left, R.Right - 3);
    Context.Canvas.Brush.Style := bsSolid;
    Context.Canvas.Brush.Color := C;
    Context.Canvas.FillRect(R);
  except
    // Never let anything escape into the editor's paint cycle.
  end;
  if Cv <> nil then
    try
      Cv.Font.Height := SavedFontHeight;
      Cv.Font.Color := TColor(SavedFontColor);
      Cv.Brush.Color := TColor(SavedBrush);
      Cv.Brush.Style := SavedStyle;
    except
    end;
end;

procedure TBlameNotifier.HandlePaintLine(const Rect: TRect;
  const Stage: TPaintLineStage; const BeforeEvent: Boolean;
  var AllowDefaultPainting: Boolean; const Context: INTACodeEditorPaintContext);
var
  Info: TBlameLine;
  Canvas: TCanvas;
  CodeRect: TRect;
  Text, Age: string;
  X, CharW: Integer;
  Line: Integer;
begin
  // GDI objects this call leaves behind are booked per subsystem -
  // the status window shows the balance (Expert.ResourceMonitor).
  var GdiG := GdiGuard(gsBlamePaint, 32);
  if BeforeEvent or (Stage <> plsEndPaint) or (Context = nil) then Exit;
  try
    Line := Context.LineState.LogicalLineNum;
    if Line <> GCaretLine then Exit;          // annotation: caret line only
    if not LineInfo(Line, Info) then Exit;

    Age := HumanAge(Info.AuthorTime, Now);
    if Info.IsUncommitted then
      Text := 'not committed yet'
    else
    begin
      Text := Info.Author;
      if Age <> '' then Text := Text + ' - ' + Age;
      if Info.Summary <> '' then Text := Text + ' - ' + Info.Summary;
      Text := Text + '  (' + Info.ShortHash + ')';
    end;
    if Text = '' then Exit;

    Canvas := Context.Canvas;
    if (Canvas = nil) or (Context.LineState = nil) or (Context.EditorState = nil) then
      Exit;
    CodeRect := Context.LineState.CodeRect;
    CharW := Max(1, Context.EditorState.CharWidth);

    // Three columns of air after the code, so the annotation cannot be
    // mistaken for part of the line.
    X := CodeRect.Left
      + (Length(Context.LineState.Text) - (Context.EditorState.LeftColumn - 1) + 3) * CharW;
    if X < CodeRect.Left then Exit;
    if X > CodeRect.Right - 20 then Exit;     // no room - stay quiet

    // Shared canvas again: put back what we change (see HandlePaintGutter).
    var SavedColor := Canvas.Font.Color;
    var SavedStyle := Canvas.Brush.Style;
    var SavedHeight := Canvas.Font.Height;
    try
      Canvas.Brush.Style := bsClear;
      // Same smaller face as the gutter column - it is an aside here too.
      Canvas.Font.Height := Min(-9, Round(Canvas.Font.Height * BlameFontScale));
      // Dimmed against the editor text: the annotation must be readable
      // but must never compete with the code itself. Blending TOWARDS the
      // background works in both themes without asking which is active.
      Canvas.Font.Color := Blend(SavedColor, GetThemedColor(clWindow), 55);
      // Vertically centred: the smaller text would otherwise cling to the
      // top of the line box.
      Canvas.TextOut(X,
        CodeRect.Top + Max(0, (CodeRect.Height - Canvas.TextHeight(Text)) div 2),
        Text);
    finally
      Canvas.Font.Height := SavedHeight;
      Canvas.Font.Color := SavedColor;
      Canvas.Brush.Style := SavedStyle;
    end;
  except
    // see above
  end;
end;

// ---------------------------------------------------------------------------
//  State updates (WM_TIMER only)
// ---------------------------------------------------------------------------

// The ONLY way to make room in the gutter: the buffer's own option. The
// paint API can draw, but it cannot create space. The original width is
// remembered and put back when the feature is switched off or the buffer
// changes - we must not leave the user's editor altered.
function TopBufferOptions: IOTABufferOptions;
var
  ES: IOTAEditorServices;
begin
  Result := nil;
  if not Supports(BorlandIDEServices, IOTAEditorServices, ES) then Exit;
  if ES.TopBuffer = nil then Exit;
  Result := ES.TopBuffer.BufferOptions;
end;

procedure RestoreGutterWidth;
var
  Opt: IOTABufferOptions;
  Orig: Integer;
begin
  // Restores the CURRENT buffer only - that is all the API reaches. Other
  // buffers we widened are put back the moment they become active again
  // (GRestorePending), so nothing stays altered behind our back.
  try
    Opt := TopBufferOptions;
    if Opt = nil then Exit;
    if not GOrigWidths.TryGetValue(NormFile(GWidenedFile), Orig) then Exit;
    // WE ARE NOT ALONE IN THE GUTTER. Other add-ons set this same option,
    // and the IDE has its own page for it - so put our value back ONLY
    // while the width is still exactly what we set. If it differs,
    // somebody changed it after us and now owns it.
    if Opt.LeftGutterWidth = Orig + GColumnWidth then
      Opt.LeftGutterWidth := Orig;
    GOrigWidths.Remove(NormFile(GWidenedFile));
  except
    // the buffer may already be gone - nothing to restore then
  end;
  GSavedWidth := -1;
  GSetWidth := -1;
  GWidenedFile := EMPTYFILE;
end;

procedure EnsureGutterWidth(const AFile: string);
var
  Opt: IOTABufferOptions;
  Key: string;
  Cur, Orig, Want: Integer;
begin
  if (AFile = EMPTYFILE) or not GEnabled or (GColumnWidth <= 0) then Exit;
  Key := NormFile(AFile);
  try
    Opt := TopBufferOptions;
    if Opt = nil then Exit;
    Cur := Opt.LeftGutterWidth;

    // THE ORIGINAL WIDTH IS REMEMBERED PER BUFFER AND NEVER RE-DERIVED.
    // Reading the current width as the new base is what made the gutter
    // grow by our column on EVERY buffer switch (tester): after a switch
    // the value we read was already widened, so we added another column
    // on top of our own - and again, and again.
    if not GOrigWidths.TryGetValue(Key, Orig) then
    begin
      // A width that already CONTAINS our column must not become the new
      // base either - in some IDE versions this option is shared between
      // buffers, so a freshly activated file can arrive pre-widened.
      if (GLastOrig >= 0) and (Cur = GLastOrig + GColumnWidth) then
        Orig := GLastOrig
      else
        Orig := Cur;
      GOrigWidths.Add(Key, Orig);
    end;
    GLastOrig := Orig;

    Want := Orig + GColumnWidth;
    if Cur <> Want then
      Opt.LeftGutterWidth := Want;
    GSavedWidth := Orig;
    GSetWidth := Want;
    GWidenedFile := AFile;
  except
    GSavedWidth := -1;
    GSetWidth := -1;
  end;
end;

// Switching off cannot reach every buffer we widened, so the ones that
// are not active right now are put back when they next become active.
procedure RestorePendingForActive(const AFile: string);
var
  Opt: IOTABufferOptions;
  Orig: Integer;
begin
  if GOrigWidths.Count = 0 then
  begin
    GRestorePending := False;
    Exit;
  end;
  if not GOrigWidths.TryGetValue(NormFile(AFile), Orig) then Exit;
  try
    Opt := TopBufferOptions;
    if (Opt <> nil) and (Opt.LeftGutterWidth = Orig + GColumnWidth) then
      Opt.LeftGutterWidth := Orig;
  except
  end;
  GOrigWidths.Remove(NormFile(AFile));
end;

function ActiveFileIsModified(const AFile: string): Boolean;
var
  MS: IOTAModuleServices;
  Module: IOTAModule;
  I: Integer;
begin
  Result := False;
  if not Supports(BorlandIDEServices, IOTAModuleServices, MS) then Exit;
  Module := MS.FindModule(AFile);
  if Module = nil then Exit;
  try
    for I := 0 to Module.GetModuleFileCount - 1 do
    begin
      var Ed := Module.GetModuleFileEditor(I);
      if (Ed <> nil) and SameText(Ed.FileName, AFile) then
        Exit(Ed.Modified);
    end;
  except
    Result := False;
  end;
end;

// Builds GMapped for the current buffer content. Returns True when it
// actually changed something (the caller only repaints then).
function RebuildMapping(const AFile: string): Boolean;
var
  Content, DiskText: string;
  Hash, I, Src: Integer;
  DiskStamp: TDateTime;
  BufLines, DiskLines: TArray<string>;
  Map: TArray<Integer>;
  Res: TBlameLines;
begin
  Result := False;
  if not Editor.ReadEditorContent(AFile, Content) then Exit;
  Hash := THashBobJenkins.GetHashValue(Content);
  if (Hash = GMapHash) and SameText(AFile, GMapFile) then Exit;

  // The file on disk does NOT change while the buffer is dirty, so read
  // and split it once per (file, mtime) instead of on every keystroke
  // pause.
  try
    DiskStamp := TFile.GetLastWriteTime(AFile);
    if (GDiskLines = nil) or not SameText(AFile, GDiskFile)
      or (DiskStamp <> GDiskStamp) then
    begin
      DiskText := TFile.ReadAllText(AFile);
      GDiskLines := DiskText.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
      GDiskFile := AFile;
      GDiskStamp := DiskStamp;
    end;
  except
    Exit;
  end;

  BufLines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  DiskLines := GDiskLines;
  Map := MapBufferToDiskLines(DiskLines, BufLines);

  SetLength(Res, Length(Map));
  for I := 0 to High(Map) do
  begin
    Src := Map[I];
    if (Src >= 1) and (Src <= Length(GLines)) then
      Res[I] := GLines[Src - 1]
    else
    begin
      // Edited or brand new: same shape as a VCS "not committed yet"
      // line, so the painter and the dialogs need no special case.
      Res[I] := Default(TBlameLine);
      Res[I].Kind := DetectVcs(AFile);
      Res[I].Hash := '';
    end;
  end;

  GMapped := Res;
  GMapHash := Hash;
  GMapFile := AFile;
  Result := True;
end;

type
  TBlameTicker = class
    procedure Tick(Sender: TObject);
  end;

var
  GTicker: TBlameTicker = nil;

procedure TBlameTicker.Tick(Sender: TObject);
var
  F: string;
  Line, Col: Integer;
  NewLines: TBlameLines;
  Changed: Boolean;
begin
  // GDI objects this call leaves behind are booked per subsystem -
  // the status window shows the balance (Expert.ResourceMonitor).
  var GdiG := GdiGuard(gsBlameTick);
  if not GEnabled then
  begin
    // Switched off, but buffers we widened are still out there: only the
    // ACTIVE one can be reached, so the rest are handed back as they come
    // up. The tick stays alive exactly as long as that list is not empty.
    if GRestorePending then
      try
        RestorePendingForActive(Editor.GetActiveFileName);
        if GOrigWidths.Count = 0 then
        begin
          GRestorePending := False;
          if GTimer <> nil then GTimer.Enabled := False;
        end;
      except
      end
    else if GTimer <> nil then
      GTimer.Enabled := False;
    Exit;
  end;
  if Application.ModalLevel > 0 then Exit;
  try
    // Cheap getters only - never GetCurrentContext from a timer.
    F := Editor.GetActiveFileName;
    if (F = '') or not SameText(ExtractFileExt(F), '.pas') then
    begin
      Changed := (GFile <> '') or (Length(GLines) > 0);
      GFile := '';
      GLines := nil;
      if Changed then InvalidateTopEditor;
      Exit;
    end;

    Changed := False;
    if not SameText(F, GFile) then
    begin
      GFile := F;
      GLines := nil;
      Changed := True;
    end;

    var Dirty := ActiveFileIsModified(F);
    if Dirty <> GDirty then
    begin
      GDirty := Dirty;
      Changed := True;
    end;

    // A modified buffer no longer matches the blame line for line - but
    // most of it still does. Remap instead of going dark: what survived
    // keeps its author, the edited span becomes "not committed".
    if GDirty and (Length(GLines) > 0) then
    begin
      if RebuildMapping(F) then Changed := True;
    end
    else if not GDirty and (Length(GMapped) > 0) then
    begin
      GMapped := nil;
      GMapHash := 0;
      GMapFile := '';
      Changed := True;
    end;

    if GRestorePending then RestorePendingForActive(F);
    EnsureGutterWidth(F);

    if not GDirty then
    begin
      RequestBlame(F);                     // no-op when current or running
      if BlameForFile(F, NewLines) then
      begin
        if Length(NewLines) <> Length(GLines) then Changed := True;
        GLines := NewLines;
      end
      else if Length(GLines) > 0 then
      begin
        GLines := nil;
        Changed := True;
      end;
    end;

    if Editor.GetCaretLineCol(Line, Col) and (Line <> GCaretLine) then
    begin
      GCaretLine := Line;
      Changed := True;
    end;

    // A repaint is a WINDOW operation - only when something changed.
    if Changed then InvalidateTopEditor;
  except
    // a tick must never raise
  end;
end;

procedure InvalidateTopEditor;
var
  SV: INTACodeEditorServices;
begin
  if Supports(BorlandIDEServices, INTACodeEditorServices, SV) then
    SV.InvalidateTopEditor;
end;

procedure SetBlameEnabled(AValue: Boolean);
begin
  if AValue = GEnabled then Exit;
  GEnabled := AValue;
  if GEnabled then
  begin
    GFile := '';
    GLines := nil;
    if GTimer <> nil then GTimer.Enabled := True;
  end
  else
  begin
    RestoreGutterWidth;      // the active buffer, right now
    GRestorePending := GOrigWidths.Count > 0;   // the rest, when visited
    // The tick keeps running while restores are pending - stopping it
    // would leave those buffers widened for the rest of the session.
    if (GTimer <> nil) and not GRestorePending then
      GTimer.Enabled := False;
    GLines := nil;
    GFile := '';
    InvalidateBlame('');
  end;
  InvalidateTopEditor;
end;

procedure ApplyBlameSettings;
begin
  // ORDER MATTERS: hand the gutter back with the width we are STILL
  // holding, and only then take the new numbers. The other way round the
  // restore compares against the NEW column width, does not recognise its
  // own value, and leaves our column in place - which the next claim then
  // treats as the buffer's own width. That is how a width change in the
  // live adjuster would grow the gutter, the same way switching buffers
  // once did.
  RestoreGutterWidth;

  GColumnWidth := EnsureRange(TPluginSettings.BlameColumnWidth, 0, 600);
  GColumnOffset := EnsureRange(TPluginSettings.BlameColumnOffset, 0, 600);
  GInfoLevel := EnsureRange(TPluginSettings.BlameInfo, 0, 2);
  SetBlameEnabled(TPluginSettings.LiveBlame);
  if GEnabled then
    try
      EnsureGutterWidth(Editor.GetActiveFileName);
    except
    end;
  InvalidateTopEditor;
end;

procedure InstallBlameGutter;
var
  SV: INTACodeEditorServices;
begin
  if GNotifierIndex >= 0 then Exit;
  if not Supports(BorlandIDEServices, INTACodeEditorServices, SV) then Exit;
  GNotifier := TBlameNotifier.Create;
  GNotifierIndex := SV.AddEditorEventsNotifier(GNotifier);
  if GNotifierIndex < 0 then
  begin
    GNotifier := nil;   // interface refcount frees it
    Exit;
  end;
  GOrigWidths := TDictionary<string, Integer>.Create;
  GTicker := TBlameTicker.Create;
  GTimer := TTimer.Create(nil);
  GTimer.Interval := 1000;
  GTimer.OnTimer := GTicker.Tick;
  GTimer.Enabled := GEnabled;
end;

procedure UninstallBlameGutter;
var
  SV: INTACodeEditorServices;
begin
  if GTimer <> nil then
  begin
    GTimer.Enabled := False;
    FreeAndNil(GTimer);
  end;
  RestoreGutterWidth;
  FreeAndNil(GTicker);
  if (GNotifierIndex >= 0)
    and Supports(BorlandIDEServices, INTACodeEditorServices, SV) then
    try
      SV.RemoveEditorEventsNotifier(GNotifierIndex);
    except
    end;
  GNotifierIndex := -1;
  GNotifier := nil;
  FreeAndNil(GOrigWidths);
  ShutdownBlame;
end;

end.
