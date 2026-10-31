(*
  Expert.IntfVisualizerIde - the IDE half of the interface visualizer.

  Expert.IntfVisualizer resolves an interface reference from the RTTI bytes
  and is pure: it takes a memory reader and knows nothing about the IDE.
  This unit supplies the two things only the IDE has - the inspected
  process's memory and the ADDRESS of the reference - and offers the result
  through both surfaces the debugger has:

    * IOTADebuggerVisualizerExternalViewer - a window with the hierarchy.
      ON by default: it is opt-in per use anyway (the debugger lists it for
      the type and the user picks it), and it writes nothing.
    * IOTADebuggerVisualizerValueReplacer - one line in every tooltip,
      watch, local view and inspector. OFF by default, because ToolsAPI
      says "There can be only one active IOTADebuggerVisualizerValueReplacer
      per type": claiming IInterface with AllDescendants takes over the
      display of every interface variable in the debugger.

  HOW THE ADDRESS IS FOUND - measured in ToolsAPI, because the surfaces
  hand us nothing but STRINGS (GetReplacementValue / Show receive
  Expression, TypeName, EvalResult):
    IOTAThread.Evaluate's doc says "ResultAddr - Will be set if the
    expression evaluates to an address within the process", so for an
    interface VARIABLE that is the address of the variable and ResultSize
    is its size. Reading PointerSize bytes there gives the reference
    itself - no string parsing at all. An expression that is no lvalue
    (a function result, a property) has no such address, and for those the
    evaluator's own text is the only source, so ParseEvalPointer reads it.
  WHAT COULD NOT BE MEASURED HERE, and the reason the window carries a
  diagnosis pane: none of this can be exercised without a debug session, so
  the first live run has to be able to say WHICH route produced the address
  and where it stopped. That is the same reason the status rows exist.
*)
unit Expert.IntfVisualizerIde;

interface

uses
  System.Classes, Vcl.Forms,
  Expert.IntfVisualizer;

/// <summary>Builds the window and fills it with a resolution that was
///  produced elsewhere. The IDE path does not use this - the seam exists
///  so the layout can be LOOKED AT before anyone clicks it, which has
///  found a defect in every dialog of this plugin it was used on. The
///  caller frees the form.</summary>
function CreateIntfViewerWindow(AOwner: TComponent;
  const AExpression, ATypeName, AEvalResult, ARoute, AProblem: string;
  const ARes: TIntfRefResolution): TForm;

/// <summary>Registers the surfaces the settings ask for. Called from
///  Expert.Registration; safe to call again (it only registers what is
///  not registered yet).</summary>
procedure ApplyIntfVisualizerSettings;

/// <summary>Takes both surfaces back. MANDATORY in the package's
///  finalization: a visualizer the IDE still holds would be called into
///  unmapped code after the BPL is gone.</summary>
procedure UnregisterIntfVisualizers;

/// <summary>What the status window says about this feature.</summary>
function IntfVisualizerStatus: string;

implementation

uses
  System.SysUtils, Winapi.Windows,
  Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Graphics, Vcl.Clipbrd,
  ToolsAPI, Expert.PluginSettings,
  Expert.DialogHelper, Expert.IdeThemes;

const
  /// <summary>The search budget the LINE replacer uses. It runs per watch
  ///  row rather than per user action, and a failed resolution costs
  ///  2 * budget / PointerSize reads from the inspected process - so the
  ///  hopeless case (a COM object, a delegating implementation) must not
  ///  cost 2048 of them per row. A reference whose object sits further
  ///  back than this answers "not proven", never wrongly.</summary>
  LineBudget = 512;

  /// <summary>How much of the evaluator's answer we keep for the
  ///  diagnosis pane.</summary>
  MaxEvalText = 512;

type
  /// <summary>Everything one resolution attempt learned, including the
  ///  route it took - the window shows it, because nothing here can be
  ///  measured without a live debug session.</summary>
  TIntfResolveAttempt = record
    Expression: string;
    TypeName: string;
    EvalResult: string;
    Route: string;
    Reference: UInt64;
    ProcessType: string;
    Ok: Boolean;
    Problem: string;
    Res: TIntfRefResolution;
  end;

{ ----------------------------------------------------- the IDE's memory --- }

function DebuggerServices: IOTADebuggerServices;
begin
  if not Supports(BorlandIDEServices, IOTADebuggerServices, Result) then
    Result := nil;
end;

/// <summary>A reader over the INSPECTED process. "not readable" has to be
///  a normal answer rather than an exception: the resolver deliberately
///  asks for addresses that need not be mapped.</summary>
function ProcessReader(const AProcess: IOTAProcess): TDebugMemReader;
begin
  Result :=
    function (AAddress: UInt64; ACount: Integer; ABuffer: Pointer): Boolean
    var
      Got: Integer;
    begin
      Result := False;
      if (AAddress = 0) or (ACount <= 0) or (ABuffer = nil) then Exit;
      try
        Got := AProcess.ReadProcessMemory(TOTAAddress(AAddress), ACount,
          PByte(ABuffer)^);
      except
        Exit;
      end;
      Result := Got = ACount;
    end;
end;

function ProcessTypeText(const AProcess: IOTAProcess): string;
begin
  try
    case AProcess.GetProcessType of
      optWin32: Result := 'Win32';
      optWin64: Result := 'Win64';
      optWin64x: Result := 'Win64 (modern)';
    else
      Result := 'other';
    end;
  except
    Result := 'unknown';
  end;
end;

/// <summary>The pointer size the inspected process PROBABLY has. Only a
///  first guess: the resolver proves the size itself through the VMT's
///  self-pointer slot, so a process type this list does not name costs
///  nothing.</summary>
function GuessedPointerSize(const AProcess: IOTAProcess): Integer;
begin
  Result := 0;
  try
    case AProcess.GetProcessType of
      optWin32, optOSX32, optiOS32, optAndroid: Result := 4;
      optWin64, optWin64x, optiOS64, optLinux64, optOSX64, optAndroid64,
      optOSXArm64: Result := 8;
    end;
  except
    Result := 0;
  end;
end;

/// <summary>Reads the reference VALUE out of the inspected process.
///  AEvalAddr is what Evaluate reported, i.e. the address OF the variable;
///  ASize is its size when the evaluator gave one. Both sizes are tried
///  when it did not, and the resolver decides which one is real.</summary>
function ReadReferenceAt(const ARead: TDebugMemReader; AEvalAddr: UInt64;
  ASize: Integer; out AValue: UInt64): Boolean;
var
  V32: UInt32;
  V64: UInt64;
begin
  AValue := 0;
  Result := False;
  if AEvalAddr = 0 then Exit;
  if (ASize = 8) or (ASize = 0) then
    if ARead(AEvalAddr, SizeOf(V64), @V64) and (V64 <> 0) then
    begin
      AValue := V64;
      Exit(True);
    end;
  if (ASize = 4) or (ASize = 0) then
    if ARead(AEvalAddr, SizeOf(V32), @V32) and (V32 <> 0) then
    begin
      AValue := V32;
      Exit(True);
    end;
end;

/// <summary>The whole chain for one expression: ask the evaluator, read
///  the reference, resolve it. Every exit records its ROUTE, because this
///  is the part no test here can reach.</summary>
function ResolveExpression(const AExpr, ATypeName, AEvalResult: string;
  AMaxOffset: Integer): TIntfResolveAttempt;
var
  DS: IOTADebuggerServices;
  Proc: IOTAProcess;
  Thread: IOTAThread;
  Read: TDebugMemReader;
  Buf: array[0..1023] of Char;
  CanModify: Boolean;
  ResAddr: TOTAAddress;
  ResSize, ResVal: LongWord;
  Eval: TOTAEvaluateResult;
  Ref: UInt64;
  PtrSize: Integer;
begin
  Result := Default(TIntfResolveAttempt);
  Result.Expression := AExpr;
  Result.TypeName := ATypeName;
  Result.EvalResult := Copy(AEvalResult, 1, MaxEvalText);

  DS := DebuggerServices;
  if DS = nil then
  begin
    Result.Problem := 'the IDE offers no debugger services';
    Exit;
  end;
  Proc := DS.CurrentProcess;
  if Proc = nil then
  begin
    Result.Problem := 'no process is being debugged right now';
    Exit;
  end;
  Result.ProcessType := ProcessTypeText(Proc);
  Read := ProcessReader(Proc);
  PtrSize := GuessedPointerSize(Proc);

  Ref := 0;
  // CurrentThread belongs to the PROCESS, not to the debugger services
  // (measured in ToolsAPI: IOTAProcess60 declares it) - and that is also
  // the right owner, because the frame an evaluation uses is that
  // process's.
  Thread := Proc.CurrentThread;
  if Thread <> nil then
  begin
    ResAddr := 0;
    ResSize := 0;
    ResVal := 0;
    Buf[0] := #0;
    try
      // AllowSideEffects = False: a visualizer must never make the
      // inspected process run code to answer a tooltip.
      Eval := Thread.Evaluate(AExpr, @Buf[0], Length(Buf), CanModify,
        False, nil, ResAddr, ResSize, ResVal);
    except
      Eval := erError;
    end;
    case Eval of
      erOK:
        if ReadReferenceAt(Read, ResAddr, Integer(ResSize), Ref) then
          Result.Route := Format('Evaluate -> variable at $%s (%d bytes), ' +
            'read the reference from there', [IntToHex(UInt64(ResAddr), 0),
            ResSize]);
      erDeferred:
        Result.Problem := 'the evaluator had to call into the process and ' +
          'answered deferred - a visualizer does not allow that';
      erBusy:
        Result.Problem := 'the evaluator is busy';
      erError:
        Result.Problem := 'the evaluator refused the expression: ' +
          string(Buf);
    end;
    if (Ref = 0) and (Eval = erOK) then
    begin
      // No lvalue (a function result, a property): the evaluator's own
      // text is all there is.
      Ref := ParseEvalPointer(string(Buf));
      if Ref <> 0 then
        Result.Route := 'Evaluate -> no address for the expression, the ' +
          'reference read out of its result text "' + Trim(string(Buf)) + '"';
    end;
  end
  else
    Result.Problem := 'no current thread - the process is running';

  if Ref = 0 then
  begin
    Ref := ParseEvalPointer(AEvalResult);
    if Ref <> 0 then
      Result.Route := 'the reference read out of the value the debugger ' +
        'had already shown';
  end;

  if Ref = 0 then
  begin
    if Result.Problem = '' then
      Result.Problem := 'no address could be derived for this expression';
    Exit;
  end;

  Result.Reference := Ref;
  Result.Res := ResolveInterfaceReference(Ref, Read, AMaxOffset);
  Result.Ok := Result.Res.Resolved;
  if not Result.Ok then
    Result.Problem := Result.Res.Reason;
  if PtrSize <> 0 then
    Result.Route := Result.Route + Format(' (the process is %d-bit; the ' +
      'resolver proves the pointer size itself)', [PtrSize * 8]);
end;

{ ------------------------------------------------------------ the window - }

type
  TIntfViewerForm = class;

  /// <summary>What the debugger talks to. It must outlive the form OR
  ///  tolerate its absence, so the two know about each other and each
  ///  one clears the other's reference before it goes.</summary>
  TIntfViewerUpdater = class(TInterfacedObject,
    IOTADebuggerVisualizerExternalViewerUpdater)
  strict private
    FForm: TIntfViewerForm;
    FClosed: TOTAVisualizerClosedProcedure;
  public
    constructor Create(AForm: TIntfViewerForm);
    destructor Destroy; override;
    procedure CloseVisualizer;
    procedure MarkUnavailable(Reason: TOTAVisualizerUnavailableReason);
    procedure RefreshVisualizer(const Expression, TypeName,
      EvalResult: string);
    procedure SetClosedCallback(ClosedProc: TOTAVisualizerClosedProcedure);
    /// <summary>The form is going away - stop refreshing it, and tell the
    ///  debugger so it stops asking.</summary>
    procedure FormDestroyed;
  end;

  TIntfViewerForm = class(TForm)
  strict private
    FUpdater: TIntfViewerUpdater;
    FHeader: TLabel;
    FTree: TTreeView;
    FDiag: TMemo;
    FViewerBox: TCheckBox;
    FLineBox: TCheckBox;
    FFilling: Boolean;
    FLast: TIntfResolveAttempt;
    procedure DoFormClose(Sender: TObject; var Action: TCloseAction);
    procedure DoToggle(Sender: TObject);
    procedure DoCopy(Sender: TObject);
    function ReportText: string;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ShowAttempt(const AAttempt: TIntfResolveAttempt);
    procedure MarkUnavailable(const AText: string);
    procedure DetachUpdater;
    property Updater: TIntfViewerUpdater read FUpdater write FUpdater;
  end;

{ TIntfViewerForm }

constructor TIntfViewerForm.Create(AOwner: TComponent);
var
  Bar: TPanel;
  Opt: TPanel;
  Btn: TButton;
  Hint: TLabel;
begin
  inherited CreateNew(AOwner);
  OnClose := DoFormClose;
  Caption := 'Interface reference';
  Width := 820;
  Height := 620;
  Position := poScreenCenter;
  BorderStyle := bsSizeable;

  // alBottom siblings are ordered by their TOP, never by the order they
  // were created in - so each one says where it belongs.
  Bar := TPanel.Create(Self);
  Bar.Parent := Self;
  Bar.Align := alBottom;
  Bar.Top := 30000;
  Bar.Height := 42;
  Bar.BevelOuter := bvNone;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := '&Close';
  Btn.SetBounds(Bar.Width - 110, 8, 95, 26);
  Btn.Anchors := [akRight, akTop];
  Btn.Cancel := True;
  Btn.OnClick := DoToggle;
  Btn.Tag := -1;

  Btn := TButton.Create(Self);
  Btn.Parent := Bar;
  Btn.Caption := 'Copy &report';
  Btn.SetBounds(12, 8, 110, 26);
  Btn.OnClick := DoCopy;

  FDiag := TMemo.Create(Self);
  FDiag.Parent := Self;
  FDiag.Align := alBottom;
  FDiag.Top := 20000;
  FDiag.Height := 92;
  FDiag.ReadOnly := True;
  FDiag.ScrollBars := ssVertical;
  FDiag.WordWrap := True;

  Opt := TPanel.Create(Self);
  Opt.Parent := Self;
  Opt.Align := alBottom;
  Opt.Top := 10000;
  Opt.Height := 74;
  Opt.BevelOuter := bvNone;

  FViewerBox := TCheckBox.Create(Self);
  FViewerBox.Parent := Opt;
  FViewerBox.SetBounds(12, 6, 760, 19);
  FViewerBox.Caption := 'Offer this window for interface references while ' +
    'debugging';
  FViewerBox.OnClick := DoToggle;
  FViewerBox.Tag := 1;

  FLineBox := TCheckBox.Create(Self);
  FLineBox.Parent := Opt;
  FLineBox.SetBounds(12, 28, 760, 19);
  FLineBox.Caption := 'Also replace the value shown for every interface ' +
    'reference with one line';
  FLineBox.OnClick := DoToggle;
  FLineBox.Tag := 2;

  Hint := TLabel.Create(Self);
  Hint.Parent := Opt;
  Hint.SetBounds(30, 49, 760, 17);
  Hint.Caption := 'The IDE allows only one such replacement per type, so ' +
    'the second box takes over the display of every interface variable.';
  Hint.Font.Color := clGrayText;

  FHeader := TLabel.Create(Self);
  FHeader.Parent := Self;
  FHeader.Align := alTop;
  // WITH MARGINS: an alTop label fills the client edge to edge, so its
  // first character sits ON the frame and a second line is clipped at the
  // top - which is exactly what the headless render showed.
  FHeader.AlignWithMargins := True;
  FHeader.Margins.SetBounds(12, 10, 12, 6);
  FHeader.AutoSize := False;
  FHeader.WordWrap := True;
  FHeader.Height := 38;
  FHeader.Caption := '';

  FTree := TTreeView.Create(Self);
  FTree.Parent := Self;
  FTree.Align := alClient;
  FTree.ReadOnly := True;
  FTree.HideSelection := False;
  FTree.ShowRoot := True;

  PrepareDialog(Self, AOwner);
  EnableThemes(Self);
end;

destructor TIntfViewerForm.Destroy;
begin
  if FUpdater <> nil then
  begin
    FUpdater.FormDestroyed;
    FUpdater := nil;
  end;
  inherited;
end;

procedure TIntfViewerForm.DetachUpdater;
begin
  FUpdater := nil;
end;

procedure TIntfViewerForm.DoFormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TIntfViewerForm.DoToggle(Sender: TObject);
begin
  if FFilling then Exit;
  case (Sender as TControl).Tag of
    -1: Close;
    1:
      begin
        TPluginSettings.IntfVisualizerViewer := FViewerBox.Checked;
        TPluginSettings.Save;
        ApplyIntfVisualizerSettings;
      end;
    2:
      begin
        TPluginSettings.IntfVisualizerLine := FLineBox.Checked;
        TPluginSettings.Save;
        ApplyIntfVisualizerSettings;
      end;
  end;
end;

function TIntfViewerForm.ReportText: string;
var
  SL: TStringList;
  L: TIntfClassLevel;
  E: TIntfEntryInfo;
  S: string;
begin
  SL := TStringList.Create;
  try
    SL.Add('Expression: ' + FLast.Expression);
    SL.Add('Static type: ' + FLast.TypeName);
    SL.Add('Debugger showed: ' + FLast.EvalResult);
    SL.Add('Process: ' + FLast.ProcessType);
    SL.Add('Route: ' + FLast.Route);
    if FLast.Problem <> '' then
      SL.Add('Problem: ' + FLast.Problem);
    SL.Add('');
    if FLast.Res.Resolved then
    begin
      SL.Add(Format('reference $%s -> object $%s, field offset %d, %d-bit',
        [IntToHex(FLast.Res.Reference, 0), IntToHex(FLast.Res.ObjectAddr, 0),
         FLast.Res.IOffset, FLast.Res.PointerSize * 8]));
      for L in FLast.Res.Levels do
      begin
        SL.Add(Format('%s (%s), %d bytes', [L.ClassName, L.UnitName,
          L.InstanceSize]));
        for E in L.Interfaces do
        begin
          S := Format('    %s %s offset %d', [E.Name, GUIDToString(E.IID),
            E.IOffset]);
          if E.IsTheReference then S := S + '   <== the reference';
          SL.Add(S);
        end;
      end;
    end;
    Result := SL.Text;
  finally
    SL.Free;
  end;
end;

procedure TIntfViewerForm.DoCopy(Sender: TObject);
begin
  Clipboard.AsText := ReportText;
end;

procedure TIntfViewerForm.MarkUnavailable(const AText: string);
begin
  FTree.Items.Clear;
  FHeader.Caption := AText;
end;

procedure TIntfViewerForm.ShowAttempt(const AAttempt: TIntfResolveAttempt);
var
  Node: TTreeNode;
  L: TIntfClassLevel;
  E: TIntfEntryInfo;
  S: string;
begin
  FLast := AAttempt;

  FFilling := True;
  try
    FViewerBox.Checked := TPluginSettings.IntfVisualizerViewer;
    FLineBox.Checked := TPluginSettings.IntfVisualizerLine;
  finally
    FFilling := False;
  end;

  FDiag.Lines.BeginUpdate;
  try
    FDiag.Lines.Clear;
    if AAttempt.Route <> '' then
      FDiag.Lines.Add('How the address was found: ' + AAttempt.Route);
    if AAttempt.Problem <> '' then
      FDiag.Lines.Add(AAttempt.Problem);
    if AAttempt.EvalResult <> '' then
      FDiag.Lines.Add('The debugger itself shows: ' + AAttempt.EvalResult);
  finally
    FDiag.Lines.EndUpdate;
  end;

  FTree.Items.BeginUpdate;
  try
    FTree.Items.Clear;
    if not AAttempt.Res.Resolved then
    begin
      FHeader.Caption := Format('%s - nothing could be proven behind this ' +
        'reference.', [AAttempt.Expression]);
      // An EMPTY tree reads as a broken window rather than as an answer
      // (the headless render), so the reason goes where the content would
      // have been as well as into the note below.
      S := AAttempt.Problem;
      if S = '' then S := AAttempt.Res.Reason;
      if S = '' then S := 'No reason was recorded.';
      FTree.Items.AddChild(nil, S);
      Exit;
    end;

    S := AAttempt.Res.ReferenceInterface;
    if S = '' then S := '<unnamed entry>';
    FHeader.Caption := Format('%s  is  %s  of  %s' + sLineBreak +
      'reference $%s -> object $%s, field offset %d, %d-bit, %d level(s)',
      [AAttempt.Expression, S, AAttempt.Res.Levels[0].ClassName,
       IntToHex(AAttempt.Res.Reference, 0),
       IntToHex(AAttempt.Res.ObjectAddr, 0), AAttempt.Res.IOffset,
       AAttempt.Res.PointerSize * 8, Length(AAttempt.Res.Levels)]);

    for L in AAttempt.Res.Levels do
    begin
      S := L.ClassName;
      if L.UnitName <> '' then S := S + '   (' + L.UnitName + ')';
      S := S + Format('   %d bytes', [L.InstanceSize]);
      Node := FTree.Items.AddChild(nil, S);
      if Length(L.Interfaces) = 0 then
        FTree.Items.AddChild(Node, 'declares no interface of its own')
      else
        for E in L.Interfaces do
        begin
          S := E.Name;
          if S = '' then S := '<unnamed entry>';
          S := S + '   ' + GUIDToString(E.IID) +
            Format('   offset %d', [E.IOffset]);
          if E.IsTheReference then S := S + '   <== the reference';
          FTree.Items.AddChild(Node, S);
        end;
      Node.Expand(False);
    end;
  finally
    FTree.Items.EndUpdate;
  end;
end;

{ TIntfViewerUpdater }

constructor TIntfViewerUpdater.Create(AForm: TIntfViewerForm);
begin
  inherited Create;
  FForm := AForm;
  if FForm <> nil then FForm.Updater := Self;
end;

destructor TIntfViewerUpdater.Destroy;
begin
  if FForm <> nil then
  begin
    FForm.DetachUpdater;
    FForm := nil;
  end;
  inherited;
end;

procedure TIntfViewerUpdater.FormDestroyed;
begin
  FForm := nil;
  if Assigned(FClosed) then FClosed;
end;

procedure TIntfViewerUpdater.CloseVisualizer;
var
  F: TIntfViewerForm;
begin
  if FForm = nil then Exit;
  F := FForm;
  FForm := nil;
  F.DetachUpdater;   // its destructor must not call back into us
  F.Free;
end;

procedure TIntfViewerUpdater.MarkUnavailable(
  Reason: TOTAVisualizerUnavailableReason);
begin
  if FForm = nil then Exit;
  case Reason of
    ovurProcessRunning: FForm.MarkUnavailable('Process not accessible - it ' +
      'is running.');
    ovurOutOfScope: FForm.MarkUnavailable('Out of scope.');
  else
    FForm.MarkUnavailable('Not available.');
  end;
end;

procedure TIntfViewerUpdater.RefreshVisualizer(const Expression, TypeName,
  EvalResult: string);
begin
  if FForm = nil then Exit;
  FForm.ShowAttempt(ResolveExpression(Expression, TypeName, EvalResult,
    MaxIntfFieldOffset));
end;

procedure TIntfViewerUpdater.SetClosedCallback(
  ClosedProc: TOTAVisualizerClosedProcedure);
begin
  FClosed := ClosedProc;
end;

{ ------------------------------------------------------- the visualizers - }

type
  /// <summary>What both surfaces answer about themselves. The supported
  ///  type is IInterface with AllDescendants, which is the only way to
  ///  reach an arbitrary interface - whether the debugger really matches
  ///  every descendant by that, or only the STATIC type name, could not
  ///  be measured without a live session.</summary>
  TIntfVisualizerBase = class(TInterfacedObject, IOTADebuggerVisualizer,
    IOTADebuggerVisualizer250)
  public
    function GetSupportedTypeCount: Integer;
    procedure GetSupportedType(Index: Integer; var TypeName: string;
      var AllDescendants: Boolean); overload;
    procedure GetSupportedType(Index: Integer; var TypeName: string;
      var AllDescendants: Boolean; var IsGeneric: Boolean); overload;
    function GetVisualizerIdentifier: string; virtual; abstract;
    function GetVisualizerName: string; virtual; abstract;
    function GetVisualizerDescription: string; virtual; abstract;
  end;

  TIntfViewerVisualizer = class(TIntfVisualizerBase,
    IOTADebuggerVisualizerExternalViewer)
  public
    function GetVisualizerIdentifier: string; override;
    function GetVisualizerName: string; override;
    function GetVisualizerDescription: string; override;
    function GetMenuText: string;
    function Show(const Expression, TypeName, EvalResult: string;
      SuggestedLeft, SuggestedTop: Integer):
      IOTADebuggerVisualizerExternalViewerUpdater;
  end;

  TIntfLineVisualizer = class(TIntfVisualizerBase,
    IOTADebuggerVisualizerValueReplacer)
  public
    function GetVisualizerIdentifier: string; override;
    function GetVisualizerName: string; override;
    function GetVisualizerDescription: string; override;
    function GetReplacementValue(const Expression, TypeName,
      EvalResult: string): string;
  end;

{ TIntfVisualizerBase }

function TIntfVisualizerBase.GetSupportedTypeCount: Integer;
begin
  Result := 1;
end;

procedure TIntfVisualizerBase.GetSupportedType(Index: Integer;
  var TypeName: string; var AllDescendants: Boolean);
begin
  TypeName := 'IInterface';
  AllDescendants := True;
end;

procedure TIntfVisualizerBase.GetSupportedType(Index: Integer;
  var TypeName: string; var AllDescendants: Boolean;
  var IsGeneric: Boolean);
begin
  GetSupportedType(Index, TypeName, AllDescendants);
  IsGeneric := False;
end;

{ TIntfViewerVisualizer }

function TIntfViewerVisualizer.GetVisualizerIdentifier: string;
begin
  Result := 'RefactoringLightInterfaceViewer';
end;

function TIntfViewerVisualizer.GetVisualizerName: string;
begin
  Result := 'Refactoring Light: interface reference';
end;

function TIntfViewerVisualizer.GetVisualizerDescription: string;
begin
  Result := 'Resolves an interface reference to the object behind it and ' +
    'lists the class hierarchy with the interfaces each level implements, ' +
    'read from the RTTI bytes.';
end;

function TIntfViewerVisualizer.GetMenuText: string;
begin
  Result := 'Interface reference';
end;

function TIntfViewerVisualizer.Show(const Expression, TypeName,
  EvalResult: string; SuggestedLeft, SuggestedTop: Integer):
  IOTADebuggerVisualizerExternalViewerUpdater;
var
  Form: TIntfViewerForm;
  A: TIntfResolveAttempt;
begin
  Result := nil;
  // AN EXCEPTION HERE WOULD BE THE ANSWER: the IDE shows its exception box
  // to somebody who is debugging. So the resolution runs FIRST and into a
  // record - a failure becomes a diagnosis the window can show - and what
  // is created afterwards is unwound, because an orphaned invisible form
  // is worse than no window at all.
  try
    A := ResolveExpression(Expression, TypeName, EvalResult,
      MaxIntfFieldOffset);
  except
    on E: Exception do
    begin
      A := Default(TIntfResolveAttempt);
      A.Expression := Expression;
      A.TypeName := TypeName;
      A.EvalResult := EvalResult;
      A.Problem := 'the resolution raised: ' + E.Message +
        ' - please report this';
    end;
  end;

  Form := TIntfViewerForm.Create(Application);
  try
    if (SuggestedLeft > 0) or (SuggestedTop > 0) then
    begin
      Form.Position := poDesigned;
      Form.Left := SuggestedLeft;
      Form.Top := SuggestedTop;
    end;
    Form.ShowAttempt(A);
    Result := TIntfViewerUpdater.Create(Form);
    Form.Show;
  except
    Result := nil;
    Form.Free;
    raise;
  end;
end;

{ TIntfLineVisualizer }

function TIntfLineVisualizer.GetVisualizerIdentifier: string;
begin
  Result := 'RefactoringLightInterfaceLine';
end;

function TIntfLineVisualizer.GetVisualizerName: string;
begin
  Result := 'Refactoring Light: interface reference (one line)';
end;

function TIntfLineVisualizer.GetVisualizerDescription: string;
begin
  Result := 'Shows the class behind an interface reference in every ' +
    'tooltip, watch, local view and inspector.';
end;

function TIntfLineVisualizer.GetReplacementValue(const Expression, TypeName,
  EvalResult: string): string;
var
  A: TIntfResolveAttempt;
begin
  // WHAT A FAILURE MUST DO: hand the debugger back exactly what it had.
  // This runs for every interface row in every watch and tooltip, so a
  // failure of ours must never become the value somebody reads.
  Result := EvalResult;
  try
    A := ResolveExpression(Expression, TypeName, EvalResult, LineBudget);
    if A.Ok then
      Result := IntfOneLine(A.Res);
  except
    Result := EvalResult;
  end;
end;

function CreateIntfViewerWindow(AOwner: TComponent;
  const AExpression, ATypeName, AEvalResult, ARoute, AProblem: string;
  const ARes: TIntfRefResolution): TForm;
var
  Form: TIntfViewerForm;
  A: TIntfResolveAttempt;
begin
  A := Default(TIntfResolveAttempt);
  A.Expression := AExpression;
  A.TypeName := ATypeName;
  A.EvalResult := AEvalResult;
  A.Route := ARoute;
  A.Problem := AProblem;
  A.ProcessType := 'Win32';
  A.Reference := ARes.Reference;
  A.Res := ARes;
  A.Ok := ARes.Resolved;
  Form := TIntfViewerForm.Create(AOwner);
  Form.ShowAttempt(A);
  Result := Form;
end;

{ ------------------------------------------------------- registration ---- }

var
  GViewer: IOTADebuggerVisualizer;
  GLine: IOTADebuggerVisualizer;
  GProblem: string;

procedure ApplyIntfVisualizerSettings;
var
  DS: IOTADebuggerServices;
begin
  DS := DebuggerServices;
  if DS = nil then
  begin
    GProblem := 'the IDE offers no debugger services';
    Exit;
  end;
  GProblem := '';
  try
    if TPluginSettings.IntfVisualizerViewer and (GViewer = nil) then
    begin
      GViewer := TIntfViewerVisualizer.Create;
      DS.RegisterDebugVisualizer(GViewer);
    end
    else if not TPluginSettings.IntfVisualizerViewer and (GViewer <> nil) then
    begin
      DS.UnregisterDebugVisualizer(GViewer);
      GViewer := nil;
    end;

    if TPluginSettings.IntfVisualizerLine and (GLine = nil) then
    begin
      GLine := TIntfLineVisualizer.Create;
      DS.RegisterDebugVisualizer(GLine);
    end
    else if not TPluginSettings.IntfVisualizerLine and (GLine <> nil) then
    begin
      DS.UnregisterDebugVisualizer(GLine);
      GLine := nil;
    end;
  except
    on E: Exception do
      GProblem := E.Message;
  end;
end;

procedure UnregisterIntfVisualizers;
var
  DS: IOTADebuggerServices;
begin
  DS := DebuggerServices;
  try
    if (DS <> nil) and (GViewer <> nil) then
      DS.UnregisterDebugVisualizer(GViewer);
    if (DS <> nil) and (GLine <> nil) then
      DS.UnregisterDebugVisualizer(GLine);
  except
    // A finalization must not raise; what we can no longer unregister is
    // the IDE's problem and ours is to let go of it.
  end;
  GViewer := nil;
  GLine := nil;
end;

function IntfVisualizerStatus: string;
begin
  if GProblem <> '' then
    Exit('NOT available: ' + GProblem);
  if (GViewer <> nil) and (GLine <> nil) then
    Exit('window and one-line replacement registered');
  if GViewer <> nil then
    Exit('window registered (the one-line replacement is off)');
  if GLine <> nil then
    Exit('one-line replacement registered (the window is off)');
  Result := 'off - neither surface is registered';
end;

end.
