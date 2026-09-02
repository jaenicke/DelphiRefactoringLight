(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.StructureErrors;

// Taps the IDE's OWN live error evaluation (Error Insight) - no second LSP
// analysis needed: BorlandIDEServices implements IOTAStructureView (see
// StructureViewAPI), and its notifier fires StructureChanged whenever the
// Structure pane content is rebuilt, which includes the "Errors" root the
// IDE fills from its internal DelphiLSP as the user types.
//
// On each change we walk the structure context, pick the error entries -
// the caption pattern is "<code> <localized text> (line:col)", where the
// CODE prefix (E2003) and the "(L:C)" tail are language independent - and
// extract the identifier from the BUFFER at that position (never from the
// localized message text). E2003 entries are converted to TLspErrorDiag
// records and handed to Expert.AutoImport.LiveReportErrorDiags, which
// resolves them against the unit index and drives the caret hint.
//
// While this source delivers, the LSP polling fallback in Expert.AutoImport
// stays quiet (it sees the buffer state as answered). It remains in charge
// for the standalone exe and for IDE setups where Error Insight or the
// Structure pane does not deliver.
//
// IDE-only unit: do not add to the standalone project.

interface

procedure InstallStructureErrorSource;
procedure UninstallStructureErrorSource;

/// <summary>What this source is doing, for the status window: is the
///  notifier installed, how often has the IDE called it, and what did the
///  LAST call do (nodes walked, diagnostics parsed, or the reason it
///  bailed out). "0 fix(es) although the Structure view shows errors" is
///  otherwise impossible to tell apart from "never fired".</summary>
procedure StructureSourceStats(out AInstalled: Boolean;
  out AFires, ANodes, ADiags: Integer; out AReason: string);

implementation

uses
  System.SysUtils, System.Classes, System.Character,
  Vcl.ExtCtrls,
  ToolsAPI, StructureViewAPI,
  Lsp.Protocol, Expert.EditorHelperIntf, Expert.AutoImport;

type
  TStructureErrorNotifier = class(TNotifierObject, IOTAStructureNotifier)
  public
    // IOTAStructureNotifier
    procedure StructureChanged(const Context: IOTAStructureContext);
    procedure NodeEdited(const Node: IOTAStructureNode);
    procedure NodeFocused(const Node: IOTAStructureNode);
    procedure NodeSelected(const Node: IOTAStructureNode);
    procedure DefaultNodeAction(const Node: IOTAStructureNode);
    procedure VisibleChanged(Visible: WordBool);
  end;

  // One-shot timer deferral for the double-click quick fix: the popup must
  // open from a plain WM_TIMER (message loop), never from the notifier
  // dispatch or a CheckSynchronize context (deadlock risk - see
  // DefaultNodeAction).
  TShowFixDefer = class
  private
    FTimer: TTimer;
    FFile: string;
    FTriesLeft: Integer;
    procedure Tick(Sender: TObject);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Queue(const AFile: string);
  end;

var
  GShowFixDefer: TShowFixDefer = nil;

constructor TShowFixDefer.Create;
begin
  inherited;
  FTimer := TTimer.Create(nil);
  FTimer.Enabled := False;
  FTimer.Interval := 250;   // lets the IDE finish navigating to the error
  FTimer.OnTimer := Tick;
end;

destructor TShowFixDefer.Destroy;
begin
  FreeAndNil(FTimer);
  inherited;
end;

procedure TShowFixDefer.Queue(const AFile: string);
begin
  FFile := AFile;
  FTriesLeft := 12;          // resolve may still be running - retry ~3 s
  FTimer.Enabled := False;   // restart on rapid double-clicks
  FTimer.Enabled := True;
end;

procedure TShowFixDefer.Tick(Sender: TObject);
begin
  FTimer.Enabled := False;
  try
    Dec(FTriesLeft);
    if LiveShowFixAtCaret(FFile) then Exit;
    // Not ready yet? The background resolver may still be chewing on the
    // fuzzy scan - retry until it published, then give up quietly (False
    // with an idle resolver really means "no fix on this line").
    if (FTriesLeft > 0) and LiveResolveBusy then
      FTimer.Enabled := True;
  except
  end;
end;

procedure QueueShowFix(const AFile: string);
begin
  if GShowFixDefer = nil then
    GShowFixDefer := TShowFixDefer.Create;
  GShowFixDefer.Queue(AFile);
end;

// "E2003 Undeklarierter Bezeichner: 'X' in Zeile 27 (27:6)"  (German IDE)
// "E2003 Undeclared identifier: 'X' at line 27 (27:6)"       (English IDE)
// Language independent parts: the code token up front, "(line:col)" at the
// end (1-based, as shown in the IDE status bar).
function ParseErrorCaption(const ACaption: string; out ACode: string;
  out ALine, ACol: Integer): Boolean;
var
  S: string;
  P, Q, Colon: Integer;
  LinePart, ColPart: string;
begin
  Result := False;
  ACode := ''; ALine := 0; ACol := 0;
  S := Trim(ACaption);
  if Length(S) < 8 then Exit;

  // Code token: one letter + digits, followed by a space.
  if not CharInSet(S[1], ['E', 'W', 'H', 'F']) then Exit;
  P := 2;
  while (P <= Length(S)) and S[P].IsDigit do Inc(P);
  if (P < 4) or (P > Length(S)) or (S[P] <> ' ') then Exit;
  ACode := Copy(S, 1, P - 1);

  // "(digits:digits)" tail.
  if S[Length(S)] <> ')' then Exit;
  Q := Length(S) - 1;
  while (Q > 1) and (S[Q] <> '(') do Dec(Q);
  if Q <= 1 then Exit;
  Colon := Pos(':', S, Q);
  if (Colon <= Q) or (Colon >= Length(S)) then Exit;
  LinePart := Copy(S, Q + 1, Colon - Q - 1);
  ColPart := Copy(S, Colon + 1, Length(S) - Colon - 1);
  Result := TryStrToInt(LinePart, ALine) and TryStrToInt(ColPart, ACol)
    and (ALine > 0) and (ACol > 0);
end;

function SplitLines(const AContent: string): TArray<string>;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

var
  GFires: Integer;      // StructureChanged calls seen
  GLastNodes: Integer;  // nodes walked in the last call
  GLastDiags: Integer;  // diagnostics parsed in the last call
  GLastReason: string;  // why the last call did not report
  GStructInstalled: Boolean;

procedure StructureSourceStats(out AInstalled: Boolean;
  out AFires, ANodes, ADiags: Integer; out AReason: string);
begin
  AInstalled := GStructInstalled;
  AFires := GFires;
  ANodes := GLastNodes;
  ADiags := GLastDiags;
  AReason := GLastReason;
end;

procedure TStructureErrorNotifier.StructureChanged(const Context: IOTAStructureContext);
var
  FileName, Content: string;
  Lines: TArray<string>;
  Diags: TArray<TLspErrorDiag>;
  DiagCount: Integer;

  procedure CollectFrom(const ANode: IOTAStructureNode; ADepth: Integer);
  var
    I, ErrLine, ErrCol: Integer;
    Code: string;
    D: TLspErrorDiag;
  begin
    if (ANode = nil) or (ADepth > 3) then Exit;
    Inc(GLastNodes);
    if ParseErrorCaption(ANode.Caption, Code, ErrLine, ErrCol) then
    begin
      // Forward EVERY parsed entry (E2003, F2613, E2037, hints, ...) -
      // the quick-fix resolver in Expert.AutoImport decides which codes
      // it can act on. The range is zero-length at the reported position
      // (0-based); the resolver extracts the affected token from the
      // buffer content itself.
      if ErrLine <= Length(Lines) then
      begin
        D := Default(TLspErrorDiag);
        D.Code := Code;
        D.Severity := 1;
        D.Message := ANode.Caption;
        D.Range.Start.Line := ErrLine - 1;
        D.Range.Start.Character := ErrCol - 1;
        D.Range.End_ := D.Range.Start;
        if DiagCount >= Length(Diags) then
          SetLength(Diags, Length(Diags) + 8);
        Diags[DiagCount] := D;
        Inc(DiagCount);
      end;
    end
    else
      for I := 0 to ANode.ChildCount - 1 do
        CollectFrom(ANode.Child[I], ADepth + 1);
  end;

var
  I: Integer;
begin
  try
    Inc(GFires);
    GLastNodes := 0;
    GLastDiags := 0;
    GLastReason := '';
    if (Context = nil) or (Editor = nil) then
    begin
      GLastReason := 'no context / no editor helper';
      Exit;
    end;
    // Only the source-code structure carries the error nodes; designer /
    // other contexts must not count as an "answer" for the live checker.
    if not SameText(Context.StructureType, SourceCodeStructureType) then
    begin
      GLastReason := 'not a source-code structure (' + Context.StructureType + ')';
      Exit;
    end;

    FileName := Editor.GetActiveFileName;
    if FileName = '' then
    begin
      GLastReason := 'no active file';
      Exit;
    end;
    if not SameText(ExtractFileExt(FileName), '.pas') then
    begin
      GLastReason := 'active file is not a .pas (' + ExtractFileExt(FileName) + ')';
      Exit;
    end;
    if not Editor.ReadEditorContent(FileName, Content) then
    begin
      GLastReason := 'buffer not readable';
      Exit;
    end;
    Lines := SplitLines(Content);

    DiagCount := 0;
    for I := 0 to Context.RootNodeCount - 1 do
      CollectFrom(Context.GetRootStructureNode(I), 0);
    SetLength(Diags, DiagCount);
    GLastDiags := DiagCount;
    if DiagCount = 0 then
      GLastReason := Format('%d node(s), none parsed as an error caption',
        [GLastNodes]);

    // Empty is a valid answer ("no errors") - it hides the hint and keeps
    // the LSP fallback from re-analysing this buffer state.
    LiveReportErrorDiags(FileName, Content, Diags);
  except
    on E: Exception do
      GLastReason := E.ClassName + ': ' + E.Message;
    // Never let anything escape into the IDE's notifier dispatch.
  end;
end;

procedure TStructureErrorNotifier.NodeEdited(const Node: IOTAStructureNode);
begin
end;

procedure TStructureErrorNotifier.NodeFocused(const Node: IOTAStructureNode);
begin
end;

procedure TStructureErrorNotifier.NodeSelected(const Node: IOTAStructureNode);
begin
end;

// Double-click on an error entry in the Structure pane: the IDE's default
// action jumps to the error position; right after that we open our quick-fix
// chooser at the caret (a no-op when the live checker has no fix there).
procedure TStructureErrorNotifier.DefaultNodeAction(const Node: IOTAStructureNode);
var
  Code, FileName: string;
  ErrLine, ErrCol: Integer;
begin
  try
    if (Node = nil) or (Editor = nil) then Exit;
    if not ParseErrorCaption(Node.Caption, Code, ErrLine, ErrCol) then Exit;
    FileName := Editor.GetActiveFileName;
    if FileName = '' then Exit;
    // Defer via a one-shot TIMER (plain WM_TIMER from the message loop):
    // 1. it waits for the IDE to finish its own default action (navigating
    //    to the error), so the popup anchors at the NEW caret position;
    // 2. it gets us OUT of the notifier dispatch - opening an activating
    //    window inside it (or inside CheckSynchronize, where
    //    TThread.ForceQueue procs run) can deadlock the IDE against its
    //    parser thread via TEditWindow.ActivateModule -> CancelAndLock.
    QueueShowFix(FileName);
  except
    // Never let anything escape into the IDE's notifier dispatch.
  end;
end;

procedure TStructureErrorNotifier.VisibleChanged(Visible: WordBool);
begin
end;

// ---------------------------------------------------------------------------
//  Install / uninstall
// ---------------------------------------------------------------------------

type
  // TTimer needs an instance method; this tiny helper retries the service
  // lookup a few times in case the plugin loads before the Structure view
  // service is registered.
  TRetryHelper = class
    procedure Tick(Sender: TObject);
  end;

var
  GNotifierIndex: Integer = -1;
  GRetryTimer: TTimer = nil;
  GRetryHelper: TRetryHelper = nil;
  GRetryCount: Integer = 0;

procedure TryInstall;
var
  SV: IOTAStructureView;
begin
  if GNotifierIndex >= 0 then Exit;
  if Supports(BorlandIDEServices, IOTAStructureView, SV) then
    GNotifierIndex := SV.AddNotifier(TStructureErrorNotifier.Create);
  GStructInstalled := GNotifierIndex >= 0;
end;

procedure TRetryHelper.Tick(Sender: TObject);
begin
  Inc(GRetryCount);
  TryInstall;
  if (GNotifierIndex >= 0) or (GRetryCount >= 10) then
  begin
    GRetryTimer.Enabled := False;
    FreeAndNil(GRetryTimer);
    FreeAndNil(GRetryHelper);
  end;
end;

procedure InstallStructureErrorSource;
begin
  TryInstall;
  if GNotifierIndex >= 0 then Exit;
  if GRetryTimer <> nil then Exit;
  GRetryHelper := TRetryHelper.Create;
  GRetryTimer := TTimer.Create(nil);
  GRetryTimer.Interval := 2000;
  GRetryTimer.OnTimer := GRetryHelper.Tick;
  GRetryTimer.Enabled := True;
end;

procedure UninstallStructureErrorSource;
var
  SV: IOTAStructureView;
begin
  FreeAndNil(GShowFixDefer);
  if GRetryTimer <> nil then
  begin
    GRetryTimer.Enabled := False;
    FreeAndNil(GRetryTimer);
    FreeAndNil(GRetryHelper);
  end;
  if (GNotifierIndex >= 0)
    and Supports(BorlandIDEServices, IOTAStructureView, SV) then
    try
      SV.RemoveNotifier(GNotifierIndex);
    except
    end;
  GNotifierIndex := -1;
end;

end.
