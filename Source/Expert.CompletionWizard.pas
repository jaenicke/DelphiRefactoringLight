(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.CompletionWizard;

interface

uses
  System.SysUtils, System.Types, System.Classes, System.JSON, Winapi.Windows, Vcl.Forms, Vcl.Controls, Vcl.ExtCtrls, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF}
  Expert.EditorHelperIntf, Expert.LspManager, Expert.CompletionPopup, Lsp.Client;

type
  TLspCompletionWizard = class
  private
    FPopup: TCompletionPopup;
    /// <summary>Monotonic counter incremented at every Execute. The
    ///  background LSP call captures the value at start; the queued
    ///  result handler ignores itself if a newer call has been started
    ///  in the meantime. Prevents a slow first request from filling a
    ///  popup that the user has long since dismissed (or replaced
    ///  with a fresh trigger).</summary>
    FCallSeq: Integer;
    /// <summary>Watches the editor while the popup is visible (the popup
    ///  is WS_EX_NOACTIVATE, so it never learns about clicks/scrolling
    ///  elsewhere by itself): keeps the filter prefix live while the
    ///  user types, moves the popup along when the editor scrolls, and
    ///  dismisses it when the caret leaves the trigger word, the line
    ///  changes, or the editor loses focus. All window operations happen
    ///  here - a plain WM_TIMER tick (safe context).</summary>
    FPoll: TTimer;
    FShowFile: string;
    FShowLine: Integer;        // 1-based line the popup was opened on
    FShowWordStart: Integer;   // 1-based start column of the trigger word
    FLastPos: TPoint;
    procedure DoPollTick(Sender: TObject);
    procedure DoInsert(const AText: string);
    function GetCaretScreenPos: TPoint;
    function TryCaretScreenPos(out APos: TPoint): Boolean;
    /// <summary>Walks left from the current editor cursor collecting
    ///  word characters until a non-word character (or column 1) is
    ///  reached. Returns the collected characters as a prefix string.
    ///  Empty if the cursor is not immediately after a word character.</summary>
    function GetCurrentWordPrefix: string;
  public
    destructor Destroy; override;
    procedure Execute;

    // ---- Facade for the host editor ----
    // The popup never takes focus; the editor keeps every keystroke.
    // While the popup is visible the host intercepts Up/Down/Enter/
    // Escape and routes to these methods; on every other character
    // the host calls SetPrefix with the current word at the caret so
    // the popup filters live.
    function IsPopupVisible: Boolean;
    /// <summary>True only once items have arrived. Used by the host to
    ///  decide whether Up/Down/Enter should navigate the popup or be
    ///  passed through to the editor (during the Loading state, key
    ///  input should still feel normal).</summary>
    function IsPopupActive: Boolean;
    procedure SetPrefix(const APrefix: string);
    procedure MoveSelection(ADelta: Integer);
    procedure InsertSelected;
    procedure HidePopup;
  end;

var
  CompletionWizardInstance: TLspCompletionWizard;

implementation

uses
  Delphi.FileEncoding;

function TLspCompletionWizard.GetCurrentWordPrefix: string;
// Walks left from the caret to extract the partial identifier the user
// has already typed. Reads the line from the IEditorHelper - which
// reaches into the live editor buffer in the IDE plugin and the
// in-memory Memo buffer in standalone - so both hosts see the same
// thing the user sees on screen.
var
  Ctx: TEditorContext;
  Content, Line: string;
  Lines: TArray<string>;
  Col, Start: Integer;
begin
  Result := '';
  Ctx := Editor.GetCurrentContext;
  if not Ctx.IsValid then Exit;
  if not Editor.ReadEditorContent(Ctx.FileName, Content) then Exit;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  if (Ctx.Line < 1) or (Ctx.Line > Length(Lines)) then Exit;
  Line := Lines[Ctx.Line - 1];
  // Ctx.Column is 1-based. The character at Column belongs after the
  // caret (LSP / IDE convention), so we look at characters at positions
  // Column-1, Column-2, ... and stop at the first non-word char.
  Col := Ctx.Column - 1;
  if Col > Length(Line) then Col := Length(Line);
  Start := Col;
  while (Start >= 1) and CharInSet(Line[Start], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(Start);
  Result := Copy(Line, Start + 1, Col - Start);
end;

destructor TLspCompletionWizard.Destroy;
begin
  FPoll.Free;
  FreeAndNil(FPopup);
  inherited;
end;

// Caret position in SCREEN coordinates - only when the focused window is
// an EDITOR control and the caret sits inside its client area (a caret
// scrolled out of view must not reposition the popup).
function TLspCompletionWizard.TryCaretScreenPos(out APos: TPoint): Boolean;
var
  FocusHwnd: HWND;
  CaretPos: TPoint;
  R: TRect;
  Cls: array[0..63] of Char;
begin
  Result := False;
  FocusHwnd := GetFocus;
  if FocusHwnd = 0 then Exit;
  GetClassName(FocusHwnd, Cls, Length(Cls));
  if (StrIComp(Cls, 'TEditControl') <> 0) and (StrIComp(Cls, 'TMemo') <> 0) then
    Exit;
  if not GetCaretPos(CaretPos) then Exit;
  if not GetClientRect(FocusHwnd, R) then Exit;
  if (CaretPos.X < 0) or (CaretPos.Y < 0)
    or (CaretPos.X > R.Right) or (CaretPos.Y > R.Bottom) then Exit;
  ClientToScreen(FocusHwnd, CaretPos);
  APos := CaretPos;
  Inc(APos.Y, 20);
  Result := True;
end;

function TLspCompletionWizard.GetCaretScreenPos: TPoint;
begin
  if TryCaretScreenPos(Result) then Exit;
  GetCursorPos(Result);
  Inc(Result.Y, 20);
end;

// 1-based start column of the identifier that ends just before ACol1.
function WordStartAt(const ALine: string; ACol1: Integer): Integer;
var
  Col: Integer;
begin
  Col := ACol1 - 1;
  if Col > Length(ALine) then Col := Length(ALine);
  Result := Col;
  while (Result >= 1) and CharInSet(ALine[Result], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(Result);
  Inc(Result);   // first char OF the word (= ACol1 when there is no word)
end;

procedure TLspCompletionWizard.DoPollTick(Sender: TObject);
// The popup never takes focus and never sees editor events - this tick
// is its senses: live filter, follow-scroll, and every dismiss reason
// (line change, caret left the trigger word, focus lost, buffer switch).
var
  Line, Col, Start: Integer;
  Content, LineText: string;
  Lines: TArray<string>;
  Pt: TPoint;
begin
  if (FPopup = nil) or not FPopup.IsOnScreen then
  begin
    FPoll.Enabled := False;
    Exit;
  end;

  // NEVER GetCurrentContext in a timer (it moves the IDE caret) - the
  // cheap TopBuffer-based queries are safe.
  if not SameText(Editor.GetActiveFileName, FShowFile)
    or not Editor.GetCaretLineCol(Line, Col) then
  begin
    HidePopup;
    Exit;
  end;
  if Line <> FShowLine then
  begin
    HidePopup;   // Enter / click on another line
    Exit;
  end;

  // Focus / caret checks double as the "clicked somewhere else" and
  // "scrolled out of view" dismissal.
  if not TryCaretScreenPos(Pt) then
  begin
    HidePopup;
    Exit;
  end;

  if not Editor.ReadEditorContent(FShowFile, Content) then
  begin
    HidePopup;
    Exit;
  end;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  if (Line < 1) or (Line > Length(Lines)) then
  begin
    HidePopup;
    Exit;
  end;
  LineText := Lines[Line - 1];
  Start := WordStartAt(LineText, Col);
  if Start <> FShowWordStart then
  begin
    HidePopup;   // caret left the trigger word (click, arrow keys, ';', ...)
    Exit;
  end;

  // Live filter with whatever the user typed since the last tick.
  FPopup.SetPrefix(Copy(LineText, Start, Col - Start));

  // Follow the caret when the editor scrolls under the popup.
  if (Pt.X <> FLastPos.X) or (Pt.Y <> FLastPos.Y) then
  begin
    FLastPos := Pt;
    if Pt.X + FPopup.Width > Screen.Width then Pt.X := Screen.Width - FPopup.Width;
    if Pt.Y + FPopup.Height > Screen.Height then Pt.Y := Pt.Y - FPopup.Height - 40;
    SetWindowPos(FPopup.Handle, 0, Pt.X, Pt.Y, 0, 0,
      SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE);
  end;
end;

procedure TLspCompletionWizard.Execute;
// Async: the popup is shown immediately with a "Loading..." state,
// then the LSP roundtrip runs on a background thread. When the
// response (or error) arrives, the result is marshalled back to the
// UI thread which fills the popup.
//
// Why: Client.GetCompletion is a synchronous JSON-RPC call. During
// the LSP cold start it can block for 10-30 seconds while DelphiLSP
// builds its index. Calling it on the main thread froze the entire
// IDE / standalone window for that whole period.
//
// Stale-response handling: FCallSeq is bumped on every Execute. The
// background thread captures the value at start; the queued result
// handler discards itself if FCallSeq has moved on (user triggered
// a fresher completion). Without that, a slow first response would
// blow away the popup the user is already using.
var
  Context: TEditorContext;
  DelphiLspJson: string;
  RootPath: string;
  PopupPos: TPoint;
  Prefix: string;
  QueryCol: Integer;
  MySeq: Integer;
begin
  Context := Editor.GetCurrentContext;
  if not Context.IsValid then
    Exit;

  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
    Exit;

  RootPath := Context.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(Context.FileName);

  // If the cursor sits inside a partially typed identifier (e.g. "Sho"
  // in "foo.Sho|"), extract that prefix and query LSP at the START of
  // the word. That way LSP returns the full list (e.g. all members of
  // foo when preceded by a dot), and the popup filters locally by
  // prefix - instant response, and the user can backspace without
  // losing context.
  Prefix := GetCurrentWordPrefix;
  QueryCol := Context.Column - 1 - Length(Prefix); // 0-based

  PopupPos := GetCaretScreenPos;

  FreeAndNil(FPopup);

  FPopup := TCompletionPopup.CreatePopup(Application.MainForm);
  FPopup.OnInsert := DoInsert;
  // Feeds the live LSP warmup status into the popup's "Loading..."
  // line. Without this the user stares at a static "Loading..."
  // while LSP burns 25 s building its index.
  FPopup.OnLoadingStatus :=
    function: string
    begin
      Result := TLspManager.Instance.GetWarmupStatusLine;
    end;
  FPopup.ShowLoading(PopupPos.X, PopupPos.Y);

  // Arm the watcher: it keeps the filter live, follows editor scrolling
  // and dismisses the popup when the caret leaves the trigger word.
  FShowFile := Context.FileName;
  FShowLine := Context.Line;
  FShowWordStart := Context.Column - Length(Prefix);
  FLastPos := PopupPos;
  if FPoll = nil then
  begin
    FPoll := TTimer.Create(nil);
    FPoll.Interval := 150;
    FPoll.OnTimer := DoPollTick;
  end;
  FPoll.Enabled := True;

  Inc(FCallSeq);
  MySeq := FCallSeq;

  TThread.CreateAnonymousThread(
    procedure
    var
      Client: TLspClient;
      CompResponse: TJSONObject;
      Items: TArray<TCompletionItem>;
      Err: string;
    begin
      Items := nil;
      Err := '';
      try
        Client := TLspManager.Instance.GetClient(
          RootPath, Context.ProjectFile, DelphiLspJson);
        Client.RefreshDocument(Context.FileName);
        CompResponse := Client.GetCompletion(
          Context.FileName, Context.Line - 1, QueryCol);
        try
          Items := TCompletionItems.Parse(CompResponse);
        finally
          CompResponse.Free;
        end;
      except
        on E: Exception do
          Err := E.Message;
      end;

      TThread.Queue(nil,
        procedure
        begin
          // Drop stale responses - either the user dismissed the popup
          // with Escape (then it is still our FPopup but no longer
          // visible) or triggered another completion (FCallSeq bumped,
          // possibly with a fresh FPopup). In both cases, do not push
          // items into a popup the user is no longer waiting on.
          if MySeq <> FCallSeq then Exit;
          if (FPopup = nil) or not FPopup.IsOnScreen then Exit;
          if Err <> '' then
            FPopup.ShowMessage('Error: ' + Err)
          else if Length(Items) = 0 then
            FPopup.ShowMessage('No suggestions.')
          else
            FPopup.ShowItems(Items, Prefix);
        end);
    end).Start;
end;

procedure TLspCompletionWizard.DoInsert(const AText: string);
// Replace the partial identifier under the caret with AText. Word
// bounds are computed from the line text, then the replacement is
// pushed through IEditorHelper so it runs through IOTAEditWriter in
// the IDE plugin and a buffer rewrite in standalone.
var
  Ctx: TEditorContext;
  Content, Line: string;
  Lines: TArray<string>;
  Col, StartCol, EndCol: Integer;
begin
  Ctx := Editor.GetCurrentContext;
  if not Ctx.IsValid then Exit;
  if not Editor.ReadEditorContent(Ctx.FileName, Content) then Exit;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  if (Ctx.Line < 1) or (Ctx.Line > Length(Lines)) then Exit;
  Line := Lines[Ctx.Line - 1];
  Col := Ctx.Column;
  if Col > Length(Line) + 1 then Col := Length(Line) + 1;
  // Walk left to find the start of the word.
  StartCol := Col;
  while (StartCol > 1) and CharInSet(Line[StartCol - 1], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(StartCol);
  // Walk right to extend past anything still belonging to the same
  // identifier (rare - normally the caret sits right after the prefix).
  EndCol := Col;
  while (EndCol <= Length(Line)) and CharInSet(Line[EndCol], ['A'..'Z','a'..'z','0'..'9','_']) do
    Inc(EndCol);
  Editor.ReplaceSelection(Ctx.FileName, Ctx.Line, StartCol, Ctx.Line, EndCol, AText);
end;

function TLspCompletionWizard.IsPopupVisible: Boolean;
begin
  // IsOnScreen, never the VCL Visible flag - see TCompletionPopup.
  Result := (FPopup <> nil) and FPopup.IsOnScreen;
end;

function TLspCompletionWizard.IsPopupActive: Boolean;
begin
  Result := (FPopup <> nil) and FPopup.IsActive;
end;

procedure TLspCompletionWizard.SetPrefix(const APrefix: string);
begin
  if FPopup <> nil then FPopup.SetPrefix(APrefix);
end;

procedure TLspCompletionWizard.MoveSelection(ADelta: Integer);
begin
  if FPopup <> nil then FPopup.MoveSelection(ADelta);
end;

procedure TLspCompletionWizard.InsertSelected;
begin
  if FPopup <> nil then FPopup.InsertSelected;
end;

procedure TLspCompletionWizard.HidePopup;
begin
  if FPoll <> nil then FPoll.Enabled := False;
  if FPopup <> nil then FPopup.HidePopup;
end;

end.
