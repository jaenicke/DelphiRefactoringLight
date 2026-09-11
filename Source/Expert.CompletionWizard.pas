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
  Expert.EditorHelperIntf, Expert.LspManager, Expert.CompletionPopup, Lsp.Client,
  Expert.EventGen;

type
  /// <summary>A generated completion entry: an anonymous method or a new
  ///  method for a procedural type expected at the caret.</summary>
  TGenOffer = record
    Anonymous: Boolean;
    Info: TProcTypeInfo;
    SuggestedName: string;
    Caption: string;
    Detail: string;
  end;

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
    // Generated entries of the CURRENT popup (index = GenIndex - 1) and the
    // deferred execution of a picked one: the pick arrives from the
    // keyboard hook or a list click, and neither may open a modal dialog
    // or start a chain of edits - a one-shot WM_TIMER does both safely.
    FGenOffers: TArray<TGenOffer>;
    FGenTimer: TTimer;
    FGenPending: Integer;
    procedure DoGenerate(AGenIndex: Integer);
    procedure DoGenTick(Sender: TObject);
    procedure RunGenerated(const AOffer: TGenOffer);
    class function BuildGenOffers(AClient: TLspClient; const AFile,
      ALiveContent: string; const ACtx: TGenContext; ALine0,
      ACol0: Integer; out ANote: string): TArray<TGenOffer>; static;
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

/// <summary>What the last completion call decided about GENERATED entries
///  (event handler / anonymous method) and why - no site at the caret, no
///  signature help, a type that did not resolve (with the files looked
///  at), or how many entries were offered. Shown in the status window and
///  appended to %TEMP%\RefactoringLight-completion.log: a missing entry is
///  otherwise indistinguishable from a broken feature.</summary>
function CompletionGenerationNote: string;

implementation

uses
  System.Generics.Collections, System.StrUtils, System.Math, System.IOUtils,
  Delphi.FileEncoding, Expert.UnitIndex, Expert.DialogHelper, Lsp.Uri;

var
  GGenNote: string = 'no completion call yet';   // main thread only
  GGenLogChecked: Boolean = False;

function CompletionGenerationNote: string;
begin
  Result := GGenNote;
end;

// MAIN THREAD. Keeps the note for the status window and appends it to a
// small log, so the decision can be read after the fact.
procedure SetGenNote(const ANote: string);
var
  LogFile: string;
begin
  GGenNote := ANote;
  try
    LogFile := TPath.Combine(TPath.GetTempPath, 'RefactoringLight-completion.log');
    if not GGenLogChecked then
    begin
      GGenLogChecked := True;
      if TFile.Exists(LogFile) and (TFile.GetSize(LogFile) > 256 * 1024) then
        TFile.Delete(LogFile);
    end;
    TFile.AppendAllText(LogFile, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) +
      '  ' + ANote + sLineBreak, TEncoding.UTF8);
  except
    // a diagnostics log must never disturb completion
  end;
end;

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
  // Caret in virtual space beyond the line end: nothing typed there.
  if Col > Length(Line) then Exit;
  Start := Col;
  while (Start >= 1) and CharInSet(Line[Start], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(Start);
  Result := Copy(Line, Start + 1, Col - Start);
end;

destructor TLspCompletionWizard.Destroy;
begin
  FGenTimer.Free;
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
  // Caret in VIRTUAL space beyond the line end (the IDE keeps no trailing
  // blanks: "OnClick := |" is stored as "OnClick :=" with the caret two
  // columns further). No word ends there - answering the clamped line end
  // instead made the watcher see "caret left the trigger word" on its
  // first tick and close the popup before the result arrived (tester).
  if Col > Length(ALine) then Exit(ACol1);
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
  // NOT Context.IsValid: in the IDE that also demands a word under the
  // caret - and the most useful places for completion have none
  // ("Button1.OnClick := |", "TThread.Queue(nil, |"). Tester: "at that
  // place completion does not open at all". A file is all we need.
  if Context.FileName = '' then
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
  FPopup.OnGenerate := DoGenerate;
  FGenOffers := nil;
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

  // The buffer must be captured HERE, on the main thread: reading it
  // inside the worker means ToolsAPI from a foreign thread, which races
  // the IDE's own editing/parsing (IDE hangs). Empty = the file is not
  // open; the worker then falls back to disk.
  var LiveContent := '';
  if (Editor = nil) or not Editor.ReadEditorContent(Context.FileName, LiveContent) then
    LiveContent := '';

  // Is a procedural type expected here ("X.OnClick := |", "Foo(nil, |")?
  // Decided from the buffer text on the main thread; the LSP questions
  // (hover / signature help) run in the worker below.
  var GenCtx: TGenContext := Default(TGenContext);
  if LiveContent <> '' then
  begin
    var GenLines := LiveContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
    GenCtx := DetectGenContext(GenLines, Context.Line - 1, Context.Column - 1);
    if GenCtx.Site = gsNone then
    begin
      var Left := '';
      if (Context.Line >= 1) and (Context.Line <= Length(GenLines)) then
        Left := Copy(GenLines[Context.Line - 1], 1, Context.Column - 1);
      SetGenNote(Format('no assignment / call argument at %d:%d ("...%s|")',
        [Context.Line, Context.Column, Copy(Left, Max(1, Length(Left) - 40), 41)]));
    end;
  end
  else
    SetGenNote(Format('buffer of %s not readable - nothing generated',
      [ExtractFileName(Context.FileName)]));

  TThread.CreateAnonymousThread(
    procedure
    var
      Client: TLspClient;
      CompResponse: TJSONObject;
      Items: TArray<TCompletionItem>;
      Offers: TArray<TGenOffer>;
      GenNote: string;
      Err: string;
    begin
      Items := nil;
      Offers := nil;
      Err := '';
      try
        Client := TLspManager.Instance.GetClient(
          RootPath, Context.ProjectFile, DelphiLspJson);
        if LiveContent <> '' then
          Client.RefreshDocumentWith(Context.FileName, LiveContent)
        else
          Client.RefreshDocument(Context.FileName);
        CompResponse := Client.GetCompletion(
          Context.FileName, Context.Line - 1, QueryCol);
        try
          Items := TCompletionItems.Parse(CompResponse);
        finally
          CompResponse.Free;
        end;
        if GenCtx.Site <> gsNone then
        try
          Offers := BuildGenOffers(Client, Context.FileName, LiveContent,
            GenCtx, Context.Line - 1, Context.Column - 1, GenNote);
        except
          on E: Exception do
          begin
            Offers := nil;   // generation is a bonus - never break completion
            GenNote := 'generation failed: ' + E.ClassName + ': ' + E.Message;
          end;
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
          // The decision is recorded FIRST: a result that arrives for a
          // popup that is gone was silently dropped before, leaving no
          // trace in the status window or the log.
          if GenCtx.Site <> gsNone then
          begin
            if (GenNote = '') and (Err <> '') then
              GenNote := 'completion request failed before generation: ' + Err;
            var Stale := '';
            if MySeq <> FCallSeq then
              Stale := '  [dropped: superseded by a newer completion call]'
            else if (FPopup = nil) or not FPopup.IsOnScreen then
              Stale := '  [dropped: the popup was already closed]';
            SetGenNote(Format('%s:%d:%d  ', [ExtractFileName(Context.FileName),
              Context.Line, Context.Column]) + GenNote + Stale);
          end;
          if MySeq <> FCallSeq then Exit;
          if (FPopup = nil) or not FPopup.IsOnScreen then Exit;
          // Generated entries go FIRST - at such a place they are what
          // the user most likely wants.
          FGenOffers := Offers;
          var All: TArray<TCompletionItem> := nil;
          for var I := 0 to High(Offers) do
          begin
            var G: TCompletionItem;
            G.Label_ := Offers[I].Caption;
            G.Detail := Offers[I].Detail;
            G.Kind := CompletionKindGenerated;
            G.SortText := '';
            G.GenIndex := I + 1;
            All := All + [G];
          end;
          All := All + Items;
          if (Err <> '') and (Length(All) = 0) then
            FPopup.ShowMessage('Error: ' + Err)
          else if Length(All) = 0 then
            FPopup.ShowMessage('No suggestions.')
          else
            FPopup.ShowItems(All, Prefix);
        end);
    end).Start;
end;

class function TLspCompletionWizard.BuildGenOffers(AClient: TLspClient;
  const AFile, ALiveContent: string; const ACtx: TGenContext; ALine0,
  ACol0: Integer; out ANote: string): TArray<TGenOffer>;
// WORKER THREAD: LSP calls, the immutable index snapshot and DISK reads
// only - never the editor.
var
  Files: TDictionary<string, string>;
  Source: TGenTypeSource;
  Found: TList<TProcTypeInfo>;
  ParamNames: TList<string>;
  Lines: TArray<string>;
  Looked: TStringList;
  Note: string;
  IndexMissing: Boolean;

  procedure AddInfo(const AInfo: TProcTypeInfo; const AParamName: string);
  begin
    for var X in Found do
      if (X.Kind = AInfo.Kind) and (X.IsFunction = AInfo.IsFunction)
        and SameText(X.Params, AInfo.Params)
        and SameText(X.ResultType, AInfo.ResultType) then Exit;
    Found.Add(AInfo);
    ParamNames.Add(AParamName);
  end;

  function KindText(const AInfo: TProcTypeInfo): string;
  begin
    case AInfo.Kind of
      pkMethod:    Result := 'of object';
      pkReference: Result := 'reference to';
      pkPlain:     Result := 'plain procedural (nothing to generate)';
    else
      Result := '?';
    end;
  end;

  procedure Add(const S: string);
  begin
    if Note <> '' then Note := Note + ' | ';
    Note := Note + S;
  end;

  // The unit a hover's declaration link points to - where the hovered
  // symbol is DECLARED, i.e. the place its type names are resolved from.
  function FileFromHover(const AHover: string): string;
  var
    Uri: string;
  begin
    Result := AFile;
    Uri := HoverDeclUri(AHover);
    if Uri = '' then Exit;
    try
      Result := TLspUri.FileUriToPath(Uri);
    except
      Result := AFile;
    end;
  end;

begin
  Result := nil;
  Note := '';
  IndexMissing := False;
  Lines := ALiveContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  Files := TDictionary<string, string>.Create;
  Found := TList<TProcTypeInfo>.Create;
  ParamNames := TList<string>.Create;
  Looked := TStringList.Create;
  Looked.Sorted := True;
  Looked.Duplicates := dupIgnore;
  try
    // Declarations as the COMPILER sees them from the declaring unit:
    // Expert.EventGen.MakeIndexTypeSource ranks the index's candidates by
    // that unit's uses clause (a same-named TProc in IBObjects must not
    // shadow System.SysUtils' for System.Classes). Files are read once.
    var ReadFile: TFunc<string, string> :=
      function(APath: string): string
      begin
        var Key := UpperCase(APath);
        if not Files.TryGetValue(Key, Result) then
        begin
          try
            Result := TDelphiFileEncoding.ReadAll(APath);
          except
            Result := '';
          end;
          Files.Add(Key, Result);
          Looked.Add(ExtractFileName(APath));
        end;
      end;
    var Lookup: TFunc<string, TArray<TFindUnitHit>> :=
      function(AIdent: string): TArray<TFindUnitHit>
      begin
        Result := nil;
        var Snap := TUnitIndex.Instance.Snapshot;
        if Snap = nil then
        begin
          IndexMissing := True;
          Exit;
        end;
        Result := Snap.Lookup(AIdent);
      end;
    Source := MakeIndexTypeSource(Lookup, ReadFile, AFile, ALiveContent);

    var Info: TProcTypeInfo;
    case ACtx.Site of
      gsAssign:
        begin
          var HoverText := AClient.GetHover(AFile, ACtx.TargetLine, ACtx.TargetCol);
          var Owner: string;
          var TypeName := TypeFromHover(HoverText, Owner);
          var CtxFile := FileFromHover(HoverText);
          Add(Format('assignment to %s (declared in %s): hover type "%s", owner "%s"',
            [ACtx.TargetName, ExtractFileName(CtxFile), TypeName, Owner]));
          if (TypeName <> '') and ResolveProcType(TypeName, CtxFile, Source, Info) then
          begin
            Add(TypeName + ' = ' + KindText(Info));
            AddInfo(Info, '');
          end
          else if (TypeName = '') and (Owner <> '')
            and ResolveMemberProcType(Owner, ACtx.TargetName, CtxFile, Source, Info) then
          begin
            Add(Owner + '.' + ACtx.TargetName + ' = ' + KindText(Info));
            AddInfo(Info, '');
          end
          else
            Add('type not resolved');
        end;
      gsArgument:
        begin
          // parameter types are resolved in the unit that DECLARES the
          // called routine - hover its name to learn which one that is
          var CtxFile := AFile;
          if ACtx.CallLine >= 0 then
            CtxFile := FileFromHover(AClient.GetHover(AFile, ACtx.CallLine, ACtx.CallCol));
          var Sig := AClient.GetSignatureHelp(AFile, ALine0, ACol0);
          if Sig = nil then
            Add(Format('call argument at %d:%d: no signature help', [ALine0 + 1, ACol0 + 1]))
          else
          try
            Add('call declared in ' + ExtractFileName(CtxFile));
            var Active := Sig.GetValue<Integer>('activeParameter', 0);
            var Sigs: TJSONArray;
            if not Sig.TryGetValue<TJSONArray>('signatures', Sigs) or (Sigs.Count = 0) then
              Add('signature help without signatures')
            else
              // every overload whose parameter at this position is
              // procedural (TThread.Queue: TThreadMethod AND
              // TThreadProcedure)
              for var S in Sigs do
              begin
                var Params: TJSONArray;
                if not (S is TJSONObject)
                  or not TJSONObject(S).TryGetValue<TJSONArray>('parameters', Params) then
                  Continue;
                if (Active < 0) or (Active >= Params.Count) then
                begin
                  Add(Format('parameter %d not in this overload', [Active + 1]));
                  Continue;
                end;
                var ParamLabel := '';
                if Params[Active] is TJSONObject then
                  TJSONObject(Params[Active]).TryGetValue<string>('label', ParamLabel);
                var TypeName := TypeFromParamLabel(ParamLabel);
                if (TypeName <> '') and ResolveProcType(TypeName, CtxFile, Source, Info) then
                begin
                  Add('"' + ParamLabel + '" = ' + KindText(Info));
                  AddInfo(Info, ParamNameFromLabel(ParamLabel));
                end
                else
                  Add('"' + ParamLabel + '": type not resolved');
              end;
          finally
            Sig.Free;
          end;
        end;
    end;

    // One METHOD entry per signature: TThread.Queue's two overloads
    // (TThreadMethod, TThreadProcedure) both take "procedure;" - a method
    // fits either, two identical entries would only confuse.
    var MethodSigs := TList<string>.Create;
    try
      for var I := 0 to Found.Count - 1 do
      begin
        var FI := Found[I];
        if FI.Kind = pkReference then
        begin
          var O: TGenOffer;
          O.Anonymous := True;
          O.Info := FI;
          O.SuggestedName := '';
          O.Caption := ProcHeadText(FI) + ' begin ' + #$2026 + ' end';
          O.Detail := 'insert an anonymous method (' + FI.TypeName + ')';
          Result := Result + [O];
        end;
        var SigKey := UpperCase(ProcHeadText(FI));
        if (FI.Kind in [pkMethod, pkReference]) and not MethodSigs.Contains(SigKey) then
        begin
          MethodSigs.Add(SigKey);
          // Only offer what can actually be generated here: the caret must
          // sit in a method of a class declared in this unit. A taken name
          // gets a number.
          var Base := SuggestHandlerName(ACtx, ParamNames[I]);
          var Name := Base;
          var Plan := PlanEventHandler(Lines, ALine0, Name, FI);
          var N := 2;
          while not Plan.Ok and (Pos('already has a member', Plan.Reason) > 0) and (N < 10) do
          begin
            Name := Base + IntToStr(N);
            Plan := PlanEventHandler(Lines, ALine0, Name, FI);
            Inc(N);
          end;
          if Plan.Ok then
          begin
            var O: TGenOffer;
            O.Anonymous := False;
            O.Info := FI;
            O.SuggestedName := Name;
            O.Caption := Name;
            O.Detail := 'create method ' + ProcHeadText(FI, Plan.ClassName + '.' + Name) +
              ' (' + FI.TypeName + ')';
            Result := Result + [O];
          end
          else
            Add('no method: ' + Plan.Reason);
        end;
      end;
    finally
      MethodSigs.Free;
    end;

    if IndexMissing then Add('identifier index not ready');
    if Looked.Count > 0 then
      Add('declarations read from: ' + Looked.CommaText);
    Add(Format('%d generated entr%s', [Length(Result),
      IfThen(Length(Result) = 1, 'y', 'ies')]));
  finally
    ANote := Note;
    Looked.Free;
    ParamNames.Free;
    Found.Free;
    Files.Free;
  end;
end;

procedure TLspCompletionWizard.DoGenerate(AGenIndex: Integer);
begin
  if (AGenIndex < 1) or (AGenIndex > Length(FGenOffers)) then Exit;
  FGenPending := AGenIndex;
  if FGenTimer = nil then
  begin
    FGenTimer := TTimer.Create(nil);
    FGenTimer.Enabled := False;
    FGenTimer.Interval := 30;
    FGenTimer.OnTimer := DoGenTick;
  end;
  FGenTimer.Enabled := True;
end;

procedure TLspCompletionWizard.DoGenTick(Sender: TObject);
var
  Idx: Integer;
begin
  FGenTimer.Enabled := False;
  Idx := FGenPending;
  FGenPending := 0;
  if (Idx < 1) or (Idx > Length(FGenOffers)) then Exit;
  try
    RunGenerated(FGenOffers[Idx - 1]);
  except
    on E: Exception do
      ShowThemedMessage('Could not generate the code: ' + E.Message);
  end;
end;

procedure TLspCompletionWizard.RunGenerated(const AOffer: TGenOffer);
var
  F, Content, Name, Text, Indent: string;
  Lines: TArray<string>;
  Line, Col, BodyOffset, BodyCol: Integer;
  Ctx: TGenContext;
  Plan: TGenPlan;
begin
  if Editor = nil then Exit;
  F := Editor.GetActiveFileName;
  if (F = '') or not SameText(F, FShowFile) then Exit;
  if not Editor.GetCaretLineCol(Line, Col) then Exit;
  if not Editor.ReadEditorContent(F, Content) then Exit;
  Lines := Content.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
  // The caret must still be at such a place - the user may have moved on.
  Ctx := DetectGenContext(Lines, Line - 1, Col - 1);
  if Ctx.Site = gsNone then Exit;

  // Caret in virtual space beyond the line end: the blanks do not exist
  // in the buffer - insert them together with the text, and never address
  // a column past the end of the line.
  var LineLen := Length(Lines[Line - 1]);
  var Pad := '';
  var EndCol := Col;
  if Col > LineLen + 1 then
  begin
    Pad := StringOfChar(' ', Col - (LineLen + 1));
    EndCol := LineLen + 1;
  end;
  var StartCol := Min(Ctx.PrefixStartCol + 1, EndCol);

  if AOffer.Anonymous then
  begin
    Indent := '';
    for var C in Lines[Line - 1] do
      if CharInSet(C, [' ', #9]) then Indent := Indent + C else Break;
    Indent := Indent + '  ';
    Text := AnonymousMethodText(AOffer.Info, Indent, BodyOffset, BodyCol);
    if Editor.ReplaceSelection(F, Line, StartCol, Line, EndCol, Pad + Text) then
      Editor.GotoLocation(F, Line - 1 + BodyOffset, BodyCol);
    Exit;
  end;

  // A method: ask for the name (prefilled), validated against the unit.
  Name := AOffer.SuggestedName;
  var Info := AOffer.Info;
  var CaretLine0 := Line - 1;
  if not AskThemedText('Create method', 'Name of the new method:', Name,
    // no 'const': TFunc<string, string> is "reference to function (Arg1: string)"
    function(AValue: string): string
    begin
      Result := '';
      if not IsPascalIdentifier(AValue) then
        Exit('not a valid identifier');
      var P := PlanEventHandler(Lines, CaretLine0, AValue, Info);
      if not P.Ok then Result := P.Reason;
    end) then
    Exit;

  Plan := PlanEventHandler(Lines, CaretLine0, Name, Info);
  if not Plan.Ok then
  begin
    ShowThemedMessage('The method cannot be created: ' + Plan.Reason);
    Exit;
  end;
  // Bottom-up: the implementation lies below the caret, the declaration
  // above it - so every position stays valid until it is used.
  if not Editor.InsertTextAtLineStart(F, Plan.ImplLine0 + 1, Plan.ImplText) then
  begin
    ShowThemedMessage('The implementation could not be inserted.');
    Exit;
  end;
  Editor.ReplaceSelection(F, Line, StartCol, Line, EndCol, Pad + Name);
  Editor.InsertTextAtLineStart(F, Plan.DeclLine0 + 1, Plan.DeclText);
  Editor.GotoLocation(F, Plan.ImplBodyLine0 + Plan.DeclLines, 2);
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
