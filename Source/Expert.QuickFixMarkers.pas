(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.QuickFixMarkers;

// Permanent editor markers for the live quick fixes: every source location
// with an applicable fix gets a dotted orange underline (the IDE itself
// underlines errors, but shows nothing for hints like H2443/H2164 - and
// our fixes should be discoverable without parking the caret on the line).
//
// Uses the official code-editor paint API (ToolsAPI.Editor, Delphi 11+):
// INTACodeEditorServices.AddEditorEventsNotifier delivers PaintLine events
// with the line's canvas + geometry; painting happens strictly inside the
// IDE's own paint cycle (main thread), reading only the live checker's
// published state (Expert.AutoImport.LiveFixesForLine). Repaints are
// triggered by Expert.AutoImport's poll tick via GLiveRepaintHook.
//
// IDE-only unit: do not add to the standalone project.

interface

procedure InstallQuickFixMarkers;
procedure UninstallQuickFixMarkers;

implementation

uses
  System.SysUtils, System.Types, System.Math,
  Vcl.Graphics, Vcl.Controls,
  ToolsAPI, ToolsAPI.Editor,
  Expert.AutoImport;

const
  MarkerColor = TColor($0000A5FF);   // orange (BGR)

type
  TQuickFixMarkerNotifier = class(TNTACodeEditorNotifier)
  private
    procedure HandlePaintLine(const Rect: TRect; const Stage: TPaintLineStage;
      const BeforeEvent: Boolean; var AllowDefaultPainting: Boolean;
      const Context: INTACodeEditorPaintContext);
  protected
    function AllowedEvents: TCodeEditorEvents; override;
    function AllowedLineStages: TPaintLineStages; override;
  public
    constructor Create;
  end;

constructor TQuickFixMarkerNotifier.Create;
begin
  inherited Create;
  OnEditorPaintLine := HandlePaintLine;
end;

function TQuickFixMarkerNotifier.AllowedEvents: TCodeEditorEvents;
begin
  Result := [cevPaintLineEvents];
end;

function TQuickFixMarkerNotifier.AllowedLineStages: TPaintLineStages;
begin
  Result := [plsEndPaint];
end;

procedure TQuickFixMarkerNotifier.HandlePaintLine(const Rect: TRect;
  const Stage: TPaintLineStage; const BeforeEvent: Boolean;
  var AllowDefaultPainting: Boolean; const Context: INTACodeEditorPaintContext);
var
  Fixes: TArray<TQuickFix>;
  Canvas: TCanvas;
  CodeRect: TRect;
  LineText: string;
  CharW, LeftCol, X1, X2, Y, Indent, WordLen: Integer;
begin
  // Draw AFTER the IDE finished the line (Stage = plsEndPaint, after-event).
  if BeforeEvent or (Stage <> plsEndPaint) or (Context = nil) then Exit;
  try
    Fixes := LiveFixesForLine(Context.FileName, Context.LogicalLineNum - 1);
    if Length(Fixes) = 0 then Exit;
    Canvas := Context.Canvas;
    if (Canvas = nil) or (Context.LineState = nil) or (Context.EditorState = nil) then Exit;

    CodeRect := Context.LineState.CodeRect;
    CharW := Max(1, Context.EditorState.CharWidth);
    LeftCol := Max(1, Context.EditorState.LeftColumn);
    Y := CodeRect.Bottom - 2;

    Canvas.Pen.Style := psDot;
    Canvas.Pen.Width := 1;
    Canvas.Pen.Color := MarkerColor;

    for var F in Fixes do
    begin
      if F.TokenLen > 0 then
      begin
        // Underline the affected token (columns are 0-based, LeftCol is
        // the 1-based first visible column - horizontal scroll offset).
        X1 := CodeRect.Left + (F.Col - (LeftCol - 1)) * CharW;
        X2 := X1 + F.TokenLen * CharW;
      end
      else
      begin
        // Line-level fix without a token anchor: underline the line's
        // FIRST WORD - a marker under the leading indentation would sit
        // detached from the code and just look broken.
        LineText := Context.LineState.Text;
        Indent := 0;
        while (Indent < Length(LineText)) and (LineText[Indent + 1] <= ' ') do
          Inc(Indent);
        WordLen := 0;
        while (Indent + WordLen < Length(LineText))
          and (LineText[Indent + WordLen + 1] > ' ') do
          Inc(WordLen);
        if WordLen = 0 then WordLen := 5;
        X1 := CodeRect.Left + (Indent - (LeftCol - 1)) * CharW;
        X2 := X1 + WordLen * CharW;
      end;
      X1 := Max(X1, CodeRect.Left);
      X2 := Min(X2, CodeRect.Right);
      if X2 <= X1 then Continue;
      Canvas.MoveTo(X1, Y);
      Canvas.LineTo(X2, Y);
    end;
  except
    // Never let anything escape into the editor's paint cycle.
  end;
end;

var
  GNotifier: TQuickFixMarkerNotifier = nil;
  GNotifierIndex: Integer = -1;

procedure DoRepaint;
var
  SV: INTACodeEditorServices;
begin
  if Supports(BorlandIDEServices, INTACodeEditorServices, SV) then
    SV.InvalidateTopEditor;
end;

procedure InstallQuickFixMarkers;
var
  SV: INTACodeEditorServices;
begin
  if GNotifierIndex >= 0 then Exit;
  if not Supports(BorlandIDEServices, INTACodeEditorServices, SV) then Exit;
  GNotifier := TQuickFixMarkerNotifier.Create;
  GNotifierIndex := SV.AddEditorEventsNotifier(GNotifier);
  if GNotifierIndex < 0 then
    GNotifier := nil   // interface refcount frees it
  else
    GLiveRepaintHook := DoRepaint;
end;

procedure UninstallQuickFixMarkers;
var
  SV: INTACodeEditorServices;
begin
  GLiveRepaintHook := nil;
  if (GNotifierIndex >= 0)
    and Supports(BorlandIDEServices, INTACodeEditorServices, SV) then
    try
      SV.RemoveEditorEventsNotifier(GNotifierIndex);
    except
    end;
  GNotifierIndex := -1;
  GNotifier := nil;
end;

end.
