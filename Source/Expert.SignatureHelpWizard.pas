(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SignatureHelpWizard;

// Parameter hint popup.
//
// Triggered when the user types `(` after an identifier. Queries
// DelphiLSP's textDocument/signatureHelp (falls back to hover if the
// server does not implement signatureHelp). Shows a tooltip-style
// floating window with the signature label - and bolds the parameter
// the caret is currently on, IDE-style.
//
// The popup is non-modal and never takes focus: typing continues in
// the editor. The wizard also tracks subsequent caret movement via
// UpdateCursor; once the caret leaves the call's parenthesis range
// the popup hides itself.

interface

uses
  System.SysUtils, System.Classes, System.JSON,
  Winapi.Windows, Winapi.Messages,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Graphics, Vcl.ExtCtrls,
  Expert.EditorHelperIntf;

type
  /// <summary>(start, end) offsets into the signature label string,
  ///  identifying one parameter. End is exclusive.</summary>
  TSigParamRange = record
    StartOff, EndOff: Integer;
  end;

  TSignatureHelpPopup = class(TForm)
  private
    FFullLabel: string;
    FParams: TArray<TSigParamRange>;
    FActive: Integer;
    FPaintBox: TPaintBox;
    // Cached layout (recomputed on every ShowSignature / SetActive):
    FLayoutLines: TArray<string>;        // pre-wrapped visual lines
    FActiveLineIdx: Integer;             // -1 if no highlight
    FActiveStartX: Integer;              // pixel offset of bold overlay
    FActiveText: string;                 // text of the bold overlay
    procedure DoPaintBox(Sender: TObject);
    procedure RebuildLayout(AMaxPixelWidth: Integer);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    constructor CreatePopup(AOwner: TComponent);
    procedure ShowSignature(const AFullLabel: string;
      const AParams: TArray<TSigParamRange>; AActive: Integer;
      AScreenX, ALineTopY: Integer);
    procedure SetActive(AActive: Integer);
    procedure HidePopup;
  end;

  TLspSignatureHelpWizard = class
  private
    FPopup: TSignatureHelpPopup;
    FCallSeq: Integer;
    // Latched once signatureHelp resolves; used by UpdateCursor to
    // re-render the popup with a new active parameter without
    // re-querying the LSP.
    FCachedLabel: string;
    FCachedParams: TArray<TSigParamRange>;
    // Trigger position (where `(` was typed), 1-based.
    FTriggerFile: string;
    FTriggerLine: Integer;
    FTriggerCol: Integer;
    FActive: Boolean;       // true while a call is being tracked
    FLastScreenX: Integer;
    FLastScreenY: Integer;
    function FindIdentBeforeOpenParen(const ALine: string; ACol: Integer;
      out AIdentCol: Integer): string;
    function ExtractFromHover(const AHover: string): string;
    function ExtractParamRanges(const ASigLabel: string;
      ASigObj: TJSONObject): TArray<TSigParamRange>;
    function ComputeActiveParam(const ALines: TArray<string>;
      ACurLine, ACurCol: Integer; out AStillInside: Boolean): Integer;
    function FindEnclosingOpenParen(const ALines: TArray<string>;
      ACurLine, ACurCol: Integer; out AOpenLine, AOpenCol: Integer;
      out AIdent: string; out AIdentCol: Integer): Boolean;
    procedure StartQuery(const AFile: string;
      AOpenLine, AOpenCol, AIdentCol: Integer; const AIdent: string;
      const ARootPath, AProjectFile: string;
      AScreenX, AScreenY: Integer);
  public
    procedure ExecuteAt(AScreenX, ALineTopY: Integer);
    /// <summary>Manual trigger (Ctrl+Shift+Space): walks back from the
    ///  caret to find the enclosing unbalanced `(`, then runs
    ///  ExecuteAt as if the user had just typed it. Multi-line calls
    ///  are supported.
    ///
    ///  AGetLineScreenY is asked for the screen Y of the line that
    ///  actually contains the `(` (which may be many lines above the
    ///  current caret). Without it the popup would float above the
    ///  current argument line instead of above the function name.</summary>
    procedure TriggerManually(AScreenX: Integer;
      AGetLineScreenY: TFunc<Integer, Integer>);
    /// <summary>Called on every editor caret movement / keystroke after
    ///  a successful Execute. Hides the popup when the caret leaves
    ///  the tracked call; otherwise re-renders with the new active
    ///  parameter index.</summary>
    procedure UpdateCursor(ALine, ACol: Integer);
    procedure HidePopup;
  end;

implementation

uses
  System.StrUtils, System.Math,
  Expert.LspManager, Lsp.Client;

const
  PadX = 8;
  PadY = 4;

{ TSignatureHelpPopup }

constructor TSignatureHelpPopup.CreatePopup(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  BorderStyle := bsNone;
  FormStyle := fsStayOnTop;
  Color := $E0FFFF;
  Width := 100;
  Height := 24;
  FPaintBox := TPaintBox.Create(Self);
  FPaintBox.Parent := Self;
  FPaintBox.Align := alClient;
  FPaintBox.OnPaint := DoPaintBox;
end;

procedure TSignatureHelpPopup.CreateParams(var Params: TCreateParams);
begin
  inherited;
  Params.Style := (Params.Style and (not WS_OVERLAPPEDWINDOW)) or WS_POPUP or WS_BORDER;
  Params.ExStyle := Params.ExStyle or WS_EX_TOPMOST or WS_EX_NOACTIVATE
    or WS_EX_TOOLWINDOW;
end;

procedure TSignatureHelpPopup.RebuildLayout(AMaxPixelWidth: Integer);
// Decides single-line vs multi-line and pre-computes the bold-overlay
// position for the active parameter. Called from ShowSignature and
// SetActive whenever the displayed state changes.
//
// Single line: render the entire label as-is, overlay the active
// parameter substring in bold. Pixel-exact because we measure the
// prefix.
//
// Multi-line (when single-line would exceed AMaxPixelWidth and there
// are >= 2 parameters): emit
//   prefix         (e.g. "function CreateEvent(")
//   "  " + param1 + ";"
//   "  " + param2 + ";"   <-- bolded if active
//   ...
//   suffix         (e.g. "): THandle;")
// The active line is bold-overdrawn at its left padding.
var
  TmpBmp: TBitmap;
  C: TCanvas;
  FullWidth, PrefixEnd, SuffixStart, I: Integer;
  Prefix, Param, Suffix, Sep: string;
begin
  FActiveLineIdx := -1;
  FActiveStartX := 0;
  FActiveText := '';
  SetLength(FLayoutLines, 0);
  if FFullLabel = '' then Exit;

  TmpBmp := TBitmap.Create;
  try
    C := TmpBmp.Canvas;
    C.Font.Name := 'Consolas';
    C.Font.Size := 9;
    FullWidth := C.TextWidth(FFullLabel);

    if (FullWidth <= AMaxPixelWidth) or (Length(FParams) <= 1) then
    begin
      // ---- single line ----
      SetLength(FLayoutLines, 1);
      FLayoutLines[0] := FFullLabel;
      if (FActive >= 0) and (FActive < Length(FParams)) then
      begin
        var R := FParams[FActive];
        if (R.StartOff >= 0) and (R.EndOff > R.StartOff)
           and (R.EndOff <= Length(FFullLabel)) then
        begin
          FActiveLineIdx := 0;
          FActiveText := Copy(FFullLabel, R.StartOff + 1, R.EndOff - R.StartOff);
          FActiveStartX := PadX + C.TextWidth(Copy(FFullLabel, 1, R.StartOff));
        end;
      end;
      Exit;
    end;

    // ---- multi line ----
    PrefixEnd := FParams[0].StartOff;
    SuffixStart := FParams[High(FParams)].EndOff;
    Prefix := Copy(FFullLabel, 1, PrefixEnd);
    Suffix := Copy(FFullLabel, SuffixStart + 1, Length(FFullLabel) - SuffixStart);

    SetLength(FLayoutLines, 2 + Length(FParams));
    FLayoutLines[0] := Prefix;
    for I := 0 to High(FParams) do
    begin
      Param := Copy(FFullLabel, FParams[I].StartOff + 1,
        FParams[I].EndOff - FParams[I].StartOff);
      if I < High(FParams) then Sep := ';' else Sep := '';
      FLayoutLines[I + 1] := '  ' + Param + Sep;
    end;
    FLayoutLines[High(FLayoutLines)] := Suffix;

    if (FActive >= 0) and (FActive < Length(FParams)) then
    begin
      Param := Copy(FFullLabel, FParams[FActive].StartOff + 1,
        FParams[FActive].EndOff - FParams[FActive].StartOff);
      FActiveLineIdx := FActive + 1;
      FActiveText := '  ' + Param;
      FActiveStartX := PadX;
    end;
  finally
    TmpBmp.Free;
  end;
end;

procedure TSignatureHelpPopup.DoPaintBox(Sender: TObject);
var
  C: TCanvas;
  LineHeight, Y, I: Integer;
begin
  C := FPaintBox.Canvas;
  C.Brush.Color := Color;
  C.FillRect(FPaintBox.ClientRect);
  C.Font.Name := 'Consolas';
  C.Font.Size := 9;
  C.Font.Color := clWindowText;
  C.Font.Style := [];
  if Length(FLayoutLines) = 0 then Exit;
  LineHeight := Abs(C.Font.Height) + 4;
  if LineHeight <= 0 then LineHeight := 16;
  Y := PadY;
  for I := 0 to High(FLayoutLines) do
  begin
    C.TextOut(PadX, Y, FLayoutLines[I]);
    Inc(Y, LineHeight);
  end;
  if (FActiveLineIdx >= 0) and (FActiveLineIdx < Length(FLayoutLines)) then
  begin
    C.Font.Style := [fsBold];
    C.Font.Color := clNavy;
    C.TextOut(FActiveStartX, PadY + FActiveLineIdx * LineHeight, FActiveText);
  end;
end;

procedure TSignatureHelpPopup.ShowSignature(const AFullLabel: string;
  const AParams: TArray<TSigParamRange>; AActive: Integer;
  AScreenX, ALineTopY: Integer);
const
  MaxPixelWidth = 900;
var
  TmpBmp: TBitmap;
  C: TCanvas;
  LineHeight, MaxW, W, I: Integer;
begin
  FFullLabel := AFullLabel;
  FParams := AParams;
  FActive := AActive;
  RebuildLayout(MaxPixelWidth);

  TmpBmp := TBitmap.Create;
  try
    C := TmpBmp.Canvas;
    C.Font.Name := 'Consolas';
    C.Font.Size := 9;
    LineHeight := Abs(C.Font.Height) + 4;
    if LineHeight <= 0 then LineHeight := 16;
    MaxW := 0;
    for I := 0 to High(FLayoutLines) do
    begin
      W := C.TextWidth(FLayoutLines[I]);
      if W > MaxW then MaxW := W;
    end;
    ClientWidth := MaxW + 2 * PadX;
    ClientHeight := Length(FLayoutLines) * LineHeight + 2 * PadY;
  finally
    TmpBmp.Free;
  end;

  Left := AScreenX;
  Top := ALineTopY - ClientHeight - 2;
  if Left < 0 then Left := 0;
  if Left + Width > Screen.Width then Left := Screen.Width - Width;
  if Top < 0 then Top := ALineTopY + 22;
  ShowWindow(Handle, SW_SHOWNOACTIVATE);
  Visible := True;
  FPaintBox.Invalidate;
end;

procedure TSignatureHelpPopup.SetActive(AActive: Integer);
const
  MaxPixelWidth = 900;
begin
  if AActive = FActive then Exit;
  FActive := AActive;
  RebuildLayout(MaxPixelWidth);
  if Visible then FPaintBox.Invalidate;
end;

procedure TSignatureHelpPopup.HidePopup;
begin
  if Visible then Hide;
end;

{ TLspSignatureHelpWizard }

function TLspSignatureHelpWizard.FindIdentBeforeOpenParen(const ALine: string;
  ACol: Integer; out AIdentCol: Integer): string;
var
  P, EndP: Integer;
begin
  Result := '';
  AIdentCol := 0;
  P := ACol - 2;
  while (P >= 1) and CharInSet(ALine[P], [' ', #9]) do Dec(P);
  if (P < 1) or not CharInSet(ALine[P], ['A'..'Z','a'..'z','0'..'9','_']) then Exit;
  EndP := P;
  while (P >= 1) and CharInSet(ALine[P], ['A'..'Z','a'..'z','0'..'9','_']) do Dec(P);
  Inc(P);
  AIdentCol := P;
  Result := Copy(ALine, P, EndP - P + 1);
end;

function TLspSignatureHelpWizard.ExtractFromHover(const AHover: string): string;
var
  Lines: TArray<string>;
  L, T: string;
begin
  Result := '';
  if AHover = '' then Exit;
  Lines := AHover.Split([#10]);
  for L in Lines do
  begin
    T := Trim(L).Replace(#13, '');
    if T = '' then Continue;
    if T.StartsWith('```') then Continue;
    Exit(T);
  end;
end;

function TLspSignatureHelpWizard.ExtractParamRanges(const ASigLabel: string;
  ASigObj: TJSONObject): TArray<TSigParamRange>;
// Resolves parameters[].label entries into (start, end) offsets in
// ASigLabel. Each parameter label is either:
//   * a string -> we Pos() it in the signature
//   * a [start, end] number pair -> use as-is
// LSP offsets are UTF-16 code units, which for ASCII Pascal signatures
// equals plain character offsets.
var
  ParamsArr: TJSONArray;
  Item: TJSONValue;
  ParamObj: TJSONObject;
  LblVal: TJSONValue;
  S: string;
  P: Integer;
  R: TSigParamRange;
  Found: Integer;
begin
  Result := nil;
  if (ASigObj = nil) or (ASigLabel = '') then Exit;
  if not ASigObj.TryGetValue<TJSONArray>('parameters', ParamsArr) then Exit;
  if (ParamsArr = nil) or (ParamsArr.Count = 0) then Exit;
  Found := 0;
  for Item in ParamsArr do
  begin
    if not (Item is TJSONObject) then Continue;
    ParamObj := TJSONObject(Item);
    LblVal := ParamObj.GetValue('label');
    R.StartOff := -1; R.EndOff := -1;
    if LblVal is TJSONString then
    begin
      S := TJSONString(LblVal).Value;
      // Search from the previous param's end - same-name params don't
      // bother us this way.
      var SearchFrom: Integer := Found;
      P := PosEx(S, ASigLabel, SearchFrom + 1);
      if P > 0 then
      begin
        R.StartOff := P - 1;
        R.EndOff := P - 1 + Length(S);
        Found := R.EndOff;
      end;
    end
    else if LblVal is TJSONArray then
    begin
      var Arr := TJSONArray(LblVal);
      if Arr.Count >= 2 then
      begin
        R.StartOff := Arr.Items[0].AsType<Integer>;
        R.EndOff := Arr.Items[1].AsType<Integer>;
      end;
    end;
    SetLength(Result, Length(Result) + 1);
    Result[High(Result)] := R;
  end;
end;

function TLspSignatureHelpWizard.ComputeActiveParam(const ALines: TArray<string>;
  ACurLine, ACurCol: Integer; out AStillInside: Boolean): Integer;
// Pascal-aware scan from (FTriggerLine, FTriggerCol) up to but not
// including (ACurLine, ACurCol). Tracks parenthesis depth, returns
// the count of depth-1 commas seen so far (= active parameter index).
// AStillInside reflects whether the caret is still within the same
// unmatched `(` we are tracking. Multi-line aware.
var
  L, Col, EndCol, Depth, Commas: Integer;
  Line: string;
  InStr, InComment: Boolean;
  Ch: Char;
begin
  Result := 0;
  AStillInside := False;
  if FTriggerLine < 1 then Exit;
  // Caret behind the trigger? That should not happen but be safe.
  if (ACurLine < FTriggerLine) or
     ((ACurLine = FTriggerLine) and (ACurCol < FTriggerCol)) then Exit;

  Depth := 0;
  Commas := 0;
  InStr := False;
  InComment := False;

  L := FTriggerLine;
  Col := FTriggerCol;
  while L <= ACurLine do
  begin
    if (L < 1) or (L > Length(ALines)) then Break;
    Line := ALines[L - 1];
    if L = ACurLine then
      EndCol := Min(ACurCol - 1, Length(Line))
    else
      EndCol := Length(Line);

    while Col <= EndCol do
    begin
      Ch := Line[Col];
      if InComment then
      begin
        if Ch = '}' then InComment := False;
      end
      else if InStr then
      begin
        if Ch = '''' then InStr := False;
      end
      else
        case Ch of
          '''': InStr := True;
          '{':  InComment := True;
          '(':  Inc(Depth);
          ')':
            begin
              Dec(Depth);
              if Depth <= 0 then Exit;
            end;
          ',':
            if Depth = 1 then Inc(Commas);
        end;
      Inc(Col);
    end;

    // End of line: Pascal strings do not span lines; block comments do.
    InStr := False;
    Inc(L);
    Col := 1;
  end;

  AStillInside := Depth >= 1;
  Result := Commas;
end;

function TLspSignatureHelpWizard.FindEnclosingOpenParen(
  const ALines: TArray<string>; ACurLine, ACurCol: Integer;
  out AOpenLine, AOpenCol: Integer;
  out AIdent: string; out AIdentCol: Integer): Boolean;
// Walks backward from (ACurLine, ACurCol-1) until it finds the `(`
// that encloses the caret (i.e. an `(` with no matching `)` between
// it and the caret). Returns its position plus the identifier
// immediately to its left.
//
// Used by the manual Ctrl+Shift+Space trigger when there is no fresh
// `(` keystroke to start tracking from.
var
  L, Col, CloseDepth: Integer;
  Line: string;
  Ch: Char;
begin
  Result := False;
  AOpenLine := 0; AOpenCol := 0;
  AIdent := ''; AIdentCol := 0;
  if (ACurLine < 1) or (ACurLine > Length(ALines)) then Exit;
  CloseDepth := 0;
  L := ACurLine;
  Col := ACurCol - 1;
  while L >= 1 do
  begin
    Line := ALines[L - 1];
    if Col > Length(Line) then Col := Length(Line);
    while Col >= 1 do
    begin
      Ch := Line[Col];
      case Ch of
        ')': Inc(CloseDepth);
        '(':
          if CloseDepth = 0 then
          begin
            AOpenLine := L;
            AOpenCol := Col;
            AIdent := FindIdentBeforeOpenParen(Line, Col + 1, AIdentCol);
            Result := AIdent <> '';
            Exit;
          end
          else
            Dec(CloseDepth);
      end;
      Dec(Col);
    end;
    Dec(L);
    if L >= 1 then Col := Length(ALines[L - 1]);
  end;
end;

procedure TLspSignatureHelpWizard.ExecuteAt(AScreenX, ALineTopY: Integer);
// Auto trigger - called from the editor right after the user typed
// `(`. The trigger column is Context.Column - 1 (where the '(' sits).
var
  Context: TEditorContext;
  Content, Line, Ident: string;
  Lines: TArray<string>;
  IdentCol: Integer;
  RootPath: string;
begin
  Context := Editor.GetCurrentContext;
  if not Context.IsValid then Exit;
  if not Editor.ReadEditorContent(Context.FileName, Content) then Exit;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  if (Context.Line < 1) or (Context.Line > Length(Lines)) then Exit;
  Line := Lines[Context.Line - 1];
  Ident := FindIdentBeforeOpenParen(Line, Context.Column, IdentCol);
  if Ident = '' then Exit;
  RootPath := Context.ProjectRoot;
  if RootPath = '' then RootPath := ExtractFilePath(Context.FileName);
  StartQuery(Context.FileName, Context.Line, Context.Column - 1,
    IdentCol, Ident, RootPath, Context.ProjectFile,
    AScreenX, ALineTopY);
end;

procedure TLspSignatureHelpWizard.StartQuery(const AFile: string;
  AOpenLine, AOpenCol, AIdentCol: Integer; const AIdent: string;
  const ARootPath, AProjectFile: string;
  AScreenX, AScreenY: Integer);
// Common path for both auto and manual trigger. Anchors tracking at
// the discovered '(' position and kicks off the LSP query on a
// background thread.
var
  DelphiLspJson: string;
  QueryLineSig, QueryColSig: Integer;
  QueryLineHover, QueryColHover: Integer;
  MySeq: Integer;
begin
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then Exit;

  // LSP 0-based positions.
  // signatureHelp: inside the open paren (just after '(').
  // hover: on the identifier name.
  QueryLineSig := AOpenLine - 1;
  QueryColSig := AOpenCol;          // 0-based, char index right after '('
  QueryLineHover := AOpenLine - 1;
  QueryColHover := AIdentCol - 1;

  if FPopup = nil then
    FPopup := TSignatureHelpPopup.CreatePopup(Application.MainForm);
  FPopup.HidePopup;

  FTriggerFile := AFile;
  FTriggerLine := AOpenLine;
  FTriggerCol := AOpenCol;            // column of '('
  FActive := True;
  FCachedLabel := '';
  SetLength(FCachedParams, 0);
  FLastScreenX := AScreenX;
  FLastScreenY := AScreenY;

  Inc(FCallSeq);
  MySeq := FCallSeq;

  TThread.CreateAnonymousThread(
    procedure
    var
      Client: TLspClient;
      SigResult, ActiveSig: TJSONObject;
      Sigs: TJSONArray;
      ActiveSigIdx, ActiveParamIdx: Integer;
      Item: TJSONValue;
      SigLabel, Hover: string;
      ParamRanges: TArray<TSigParamRange>;
    begin
      SigLabel := '';
      ActiveParamIdx := 0;
      try
        Client := TLspManager.Instance.GetClient(ARootPath, AProjectFile, DelphiLspJson);
        // Push the current buffer contents to the LSP before querying.
        // Otherwise signatureHelp / hover see only what was on disk at
        // didOpen time - a function the user JUST typed in this edit
        // session is invisible to them. Reads through the IEditorHelper,
        // so it picks up the live in-memory buffer.
        try
          Client.RefreshDocument(AFile);
        except
          // RefreshDocument failure is non-fatal - we still try the
          // query; worst case the LSP answers stale.
        end;
        try
          SigResult := Client.GetSignatureHelp(AFile, QueryLineSig, QueryColSig);
          if SigResult <> nil then
          try
            if SigResult.TryGetValue<TJSONArray>('signatures', Sigs)
               and (Sigs <> nil) and (Sigs.Count > 0) then
            begin
              ActiveSigIdx := SigResult.GetValue<Integer>('activeSignature', 0);
              if (ActiveSigIdx < 0) or (ActiveSigIdx >= Sigs.Count) then ActiveSigIdx := 0;
              Item := Sigs.Items[ActiveSigIdx];
              if Item is TJSONObject then
              begin
                ActiveSig := TJSONObject(Item);
                ActiveSig.TryGetValue<string>('label', SigLabel);
                SigLabel := Trim(SigLabel);
                ParamRanges := ExtractParamRanges(SigLabel, ActiveSig);
                ActiveParamIdx := SigResult.GetValue<Integer>('activeParameter', 0);
              end;
            end;
          finally
            SigResult.Free;
          end;
        except
          SigLabel := '';
        end;
        // Hover fallback: one-liner, no parameter ranges (no
        // highlighting).
        if SigLabel = '' then
        try
          Hover := Client.GetHover(AFile, QueryLineHover, QueryColHover);
          SigLabel := ExtractFromHover(Hover);
          ParamRanges := nil;
        except
          SigLabel := '';
        end;
      except
        // Network / LSP error - show nothing.
      end;

      TThread.Queue(nil,
        procedure
        begin
          if MySeq <> FCallSeq then Exit;
          if FPopup = nil then Exit;
          if SigLabel = '' then
          begin
            FPopup.HidePopup;
            FActive := False;
            Exit;
          end;
          FCachedLabel := SigLabel;
          FCachedParams := ParamRanges;
          FPopup.ShowSignature(SigLabel, ParamRanges, ActiveParamIdx,
            FLastScreenX, FLastScreenY);
        end);
    end).Start;
end;

procedure TLspSignatureHelpWizard.UpdateCursor(ALine, ACol: Integer);
var
  Content: string;
  Lines: TArray<string>;
  StillInside: Boolean;
  NewActive: Integer;
begin
  if not FActive then Exit;
  // If we have not received the async response yet, leave the popup
  // alone - it will catch up via the queued result handler.
  if not Editor.ReadEditorContent(FTriggerFile, Content) then Exit;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  NewActive := ComputeActiveParam(Lines, ALine, ACol, StillInside);
  if not StillInside then
  begin
    HidePopup;
    Exit;
  end;
  if (FCachedLabel <> '') and (FPopup <> nil) and FPopup.Visible then
    FPopup.SetActive(NewActive);
end;

procedure TLspSignatureHelpWizard.TriggerManually(AScreenX: Integer;
  AGetLineScreenY: TFunc<Integer, Integer>);
var
  Context: TEditorContext;
  Content: string;
  Lines: TArray<string>;
  OpenLine, OpenCol, IdentCol, ScreenY: Integer;
  Ident, RootPath: string;
begin
  Context := Editor.GetCurrentContext;
  if not Context.IsValid then Exit;
  if not Editor.ReadEditorContent(Context.FileName, Content) then Exit;
  Lines := Content.Split([sLineBreak], TStringSplitOptions.None);
  if not FindEnclosingOpenParen(Lines, Context.Line, Context.Column,
       OpenLine, OpenCol, Ident, IdentCol) then
  begin
    HidePopup;
    Exit;
  end;
  RootPath := Context.ProjectRoot;
  if RootPath = '' then RootPath := ExtractFilePath(Context.FileName);
  // Ask the host for the screen Y of the open-paren's line so the
  // popup floats above the function NAME, not above the argument
  // line the caret currently sits in.
  if Assigned(AGetLineScreenY) then
    ScreenY := AGetLineScreenY(OpenLine)
  else
    ScreenY := 0;
  StartQuery(Context.FileName, OpenLine, OpenCol, IdentCol, Ident,
    RootPath, Context.ProjectFile, AScreenX, ScreenY);
end;

procedure TLspSignatureHelpWizard.HidePopup;
begin
  FActive := False;
  Inc(FCallSeq);  // cancel in-flight async
  if FPopup <> nil then FPopup.HidePopup;
end;

end.
