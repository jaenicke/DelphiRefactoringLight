(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UnitReferencesWizard;

{
  "Find unit references" - given the active .pas unit, find every
  project file whose uses clause lists this unit, and within each
  such file enumerate every identifier of the target unit that is
  actually used. Files with a "dead" uses entry (listed but no
  symbols used) get a single placeholder row so the user can see
  candidates for removal at a glance.

  Algorithm:
    1. Determine which project files have the target unit in any of
       their uses clauses. Done with a comment-/string-aware text
       scan - deterministic and does not depend on LSP being awake.
    2. Get the names of every symbol declared in the target unit via
       LSP documentSymbol. We use ONLY the names, not positions: in
       practice DelphiLSP's findReferences ignores positions that
       come from documentSymbol's selectionRange and consistently
       returns empty arrays - the working FindReferences wizard only
       succeeds because it uses the editor's actual cursor position.
    3. For each using-file from step 1: scan the file textually for
       identifier tokens that appear in the name set built in step 2,
       skipping comments, strings and the uses clauses themselves.
       Each match becomes a row; files with zero matches become a
       single "dead reference" row.
}

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.UITypes, System.JSON,
  System.StrUtils, System.DateUtils, System.SyncObjs,
  System.Generics.Collections, System.Generics.Defaults,
  Vcl.Forms, Vcl.Dialogs, {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF} 
  Expert.EditorHelperIntf, Expert.UnitReferencesDialog, Expert.LspManager,
  Lsp.Uri, Lsp.Protocol, Lsp.Client, Delphi.FileEncoding;

type
  TLspFindUnitReferencesWizard = class{$IFNDEF STANDALONE_BUILD}(TNotifierObject, IOTAWizard, IOTAMenuWizard){$ENDIF}
  private type
    TSymbolPos = record
      Name: string;
      Line: Integer; // 0-based (LSP)
      Col: Integer;  // 0-based
      Kind: Integer; // LSP SymbolKind
    end;
  private
    FDialog: TUnitReferencesDialog;
    FContext: TEditorContext;
    FTraceFile: TStreamWriter;
    FTraceLock: TCriticalSection;
    FTraceStart: TDateTime;
    procedure DoGotoLocation(AItem: TUnitRefItem);
    procedure DoDialogClose(Sender: TObject);
    procedure SearchAndShow;
    procedure OpenTrace(const ARootPath: string);
    procedure CloseTrace;
    procedure TraceLine(const ADirection, AMethod, ABody: string);
    procedure TraceNote(const AText: string);
    procedure CollectSymbols(AArr: TJSONArray; AList: TList<TSymbolPos>;
      ASeen: TDictionary<string, Boolean>);
    function ReadCachedLines(AFile: string;
      ACache: TDictionary<string, TArray<string>>): TArray<string>;
  public
    {$IFNDEF STANDALONE_BUILD}

    // IOTAWizard / IOTAMenuWizard / IOTANotifier - IDE plugin only.
    procedure AfterSave;
    procedure BeforeSave;
    procedure Destroyed;
    procedure Modified;
    function GetIDString: string;
    function GetName: string;
    function GetState: TWizardState;
    function GetMenuText: string;

    {$ENDIF}
    procedure Execute;
  end;

var
  UnitReferencesInstance: TLspFindUnitReferencesWizard;

implementation

{$IFNDEF STANDALONE_BUILD}
{ TLspFindUnitReferencesWizard - IOTAWizard / IOTAMenuWizard / IOTANotifier glue.
  Only compiled into the IDE plugin; the standalone build does not
  inherit from TNotifierObject and never needs these. }

procedure TLspFindUnitReferencesWizard.AfterSave; begin end;
procedure TLspFindUnitReferencesWizard.BeforeSave; begin end;
procedure TLspFindUnitReferencesWizard.Destroyed; begin end;
procedure TLspFindUnitReferencesWizard.Modified; begin end;

function TLspFindUnitReferencesWizard.GetIDString: string;
begin Result := 'DelphiRefactoringLight.UnitReferencesWizard'; end;

function TLspFindUnitReferencesWizard.GetName: string;
begin Result := 'Delphi Refactoring Light - Find Unit References'; end;

function TLspFindUnitReferencesWizard.GetState: TWizardState;
begin Result := [wsEnabled]; end;

function TLspFindUnitReferencesWizard.GetMenuText: string;
begin Result := 'Find unit references...'; end;
{$ENDIF}
procedure TLspFindUnitReferencesWizard.OpenTrace(const ARootPath: string);
var
  Path: string;
  Stamp: string;
begin
  CloseTrace;
  FTraceLock := TCriticalSection.Create;
  FTraceStart := Now;
  Stamp := FormatDateTime('yyyymmdd_hhnnss', FTraceStart);
  Path := IncludeTrailingPathDelimiter(ARootPath) + 'UnitRefsTrace_' + Stamp + '.log';
  try
    FTraceFile := TStreamWriter.Create(Path, False, TEncoding.UTF8);
    FTraceFile.AutoFlush := True;
    FTraceFile.WriteLine(Format('# UnitRefs trace started %s',
      [FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz', FTraceStart)]));
  except
    FreeAndNil(FTraceFile);
  end;
end;

procedure TLspFindUnitReferencesWizard.CloseTrace;
begin
  if FTraceFile <> nil then
  begin
    try
      FTraceFile.WriteLine(Format('# trace closed %s, elapsed %.3fs',
        [FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz', Now),
         (Now - FTraceStart) * SecsPerDay]));
    except end;
    FreeAndNil(FTraceFile);
  end;
  FreeAndNil(FTraceLock);
end;

procedure TLspFindUnitReferencesWizard.TraceLine(const ADirection, AMethod, ABody: string);
var
  Elapsed: Double;
begin
  if FTraceFile = nil then Exit;
  Elapsed := (Now - FTraceStart) * SecsPerDay;
  FTraceLock.Enter;
  try
    try
      FTraceFile.WriteLine(Format('[%9.3f] %s %s', [Elapsed, ADirection, AMethod]));
      if ABody <> '' then
        FTraceFile.WriteLine('           ' + ABody);
    except end;
  finally
    FTraceLock.Leave;
  end;
end;

procedure TLspFindUnitReferencesWizard.TraceNote(const AText: string);
begin
  TraceLine('::', AText, '');
end;

procedure TLspFindUnitReferencesWizard.Execute;
var
  UnitName: string;
begin
  FContext := Editor.GetCurrentContext;

  if (FContext.FileName = '') or
     not SameText(ExtractFileExt(FContext.FileName), '.pas') then
  begin
    MessageDlg('Please open a Delphi unit (.pas) first.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  UnitName := ChangeFileExt(ExtractFileName(FContext.FileName), '');

  FDialog := TUnitReferencesDialog.CreateDialog(Application.MainForm, UnitName);
  FDialog.OnGotoLocation := DoGotoLocation;
  FDialog.OnDialogClose := DoDialogClose;
  TLspManager.Instance.ApplyStatusToCaption(FDialog);
  FDialog.Show;
  try
    Application.ProcessMessages;
    SearchAndShow;
  except
    on E: Exception do
      if FDialog <> nil then
        FDialog.SetStatus('Error: ' + E.Message);
  end;
  // Hand off ownership: from now on closing the dialog frees it.
  if FDialog <> nil then
    FDialog.SetClosable;
end;

procedure TLspFindUnitReferencesWizard.DoDialogClose(Sender: TObject);
begin
  // Called from the dialog's OnClose right before it frees itself.
  // Detach our log hook from the shared client and close the trace
  // before letting go of the dialog reference.
  if TLspManager.Instance.IsAlive then
  try
    var C := TLspManager.Instance.GetClient(
      FContext.ProjectRoot, FContext.ProjectFile,
      Editor.FindDelphiLspJson);
    C.Verbose := False;
    C.OnLog := nil;
  except end;
  CloseTrace;
  FDialog := nil;
end;

procedure TLspFindUnitReferencesWizard.DoGotoLocation(AItem: TUnitRefItem);
begin
  Editor.GotoLocation(AItem.FilePath, AItem.Line, AItem.Col, AItem.Length);
end;

procedure TLspFindUnitReferencesWizard.CollectSymbols(AArr: TJSONArray;
  AList: TList<TSymbolPos>; ASeen: TDictionary<string, Boolean>);
var
  Val: TJSONValue;
  Obj, RangeObj, StartObj, LocObj: TJSONObject;
  ChildArr: TJSONArray;
  Sym: TSymbolPos;
  Key, Detail: string;
  Children: TJSONValue;
begin
  if AArr = nil then Exit;
  for Val in AArr do
  begin
    if not (Val is TJSONObject) then Continue;
    Obj := TJSONObject(Val);

    Sym.Name := Obj.GetValue<string>('name', '');
    Sym.Kind := Obj.GetValue<Integer>('kind', 0);
    Detail := Obj.GetValue<string>('detail', '');
    if Sym.Name = '' then Continue;

    // Skip the "uses" section entirely - DelphiLsp emits imported
    // unit names as kind=1 (File) children of a node named "uses".
    // Those are not declarations of THIS unit, so we must not feed
    // them to the name set (e.g. "Classes" from "System.Classes"
    // would otherwise match every Classes token in the project).
    // `Continue` here also prevents recursing into children below.
    if SameText(Sym.Name, 'uses') then Continue;

    // DocumentSymbol form: selectionRange.start (preferred for clicks),
    // else range.start. Fallback to SymbolInformation: location.range.start.
    RangeObj := nil;
    if Obj.GetValue('selectionRange') is TJSONObject then
      RangeObj := Obj.GetValue('selectionRange') as TJSONObject
    else if Obj.GetValue('range') is TJSONObject then
      RangeObj := Obj.GetValue('range') as TJSONObject
    else if Obj.GetValue('location') is TJSONObject then
    begin
      LocObj := Obj.GetValue('location') as TJSONObject;
      if LocObj.GetValue('range') is TJSONObject then
        RangeObj := LocObj.GetValue('range') as TJSONObject;
    end;

    if RangeObj <> nil then
    begin
      if RangeObj.GetValue('start') is TJSONObject then
      begin
        StartObj := RangeObj.GetValue('start') as TJSONObject;
        Sym.Line := StartObj.GetValue<Integer>('line', -1);
        Sym.Col := StartObj.GetValue<Integer>('character', -1);

        if (Sym.Line >= 0) and (Sym.Col >= 0) then
        begin
          Key := Format('%d:%d:%s', [Sym.Line, Sym.Col, Sym.Name]);
          if not ASeen.ContainsKey(Key) then
          begin
            ASeen.Add(Key, True);
            AList.Add(Sym);
          end;
        end;
      end;
    end;

    // Recurse into children (DocumentSymbol form).
    Children := Obj.GetValue('children');
    if Children is TJSONArray then
    begin
      ChildArr := TJSONArray(Children);
      CollectSymbols(ChildArr, AList, ASeen);
    end;
  end;
end;

function TLspFindUnitReferencesWizard.ReadCachedLines(AFile: string;
  ACache: TDictionary<string, TArray<string>>): TArray<string>;
var
  Key: string;
begin
  Key := AnsiUpperCase(AFile);
  if ACache.TryGetValue(Key, Result) then Exit;
  try
    Result := ReadDelphiFileLines(AFile);
  except
    SetLength(Result, 0);
  end;
  ACache.Add(Key, Result);
end;

function StripPascalCommentsAndStrings(const ASource: string): string;
// Replaces { ... }, (* ... *), // ... , and 'string literals' with
// spaces (preserving line breaks) so the rest of the parser can
// safely look for "uses" without false hits inside comments/strings.
var
  I, N: Integer;
  C: Char;

  procedure Blank(AIdx: Integer);
  begin
    if (ASource[AIdx] = #10) or (ASource[AIdx] = #13) then
      Result[AIdx] := ASource[AIdx]
    else
      Result[AIdx] := ' ';
  end;

begin
  N := Length(ASource);
  SetLength(Result, N);
  I := 1;
  while I <= N do
  begin
    C := ASource[I];
    if (C = '/') and (I < N) and (ASource[I + 1] = '/') then
    begin
      while (I <= N) and (ASource[I] <> #10) and (ASource[I] <> #13) do
      begin
        Result[I] := ' ';
        Inc(I);
      end;
    end
    else if C = '{' then
    begin
      while (I <= N) and (ASource[I] <> '}') do
      begin
        Blank(I);
        Inc(I);
      end;
      if I <= N then begin Result[I] := ' '; Inc(I); end;
    end
    else if (C = '(') and (I < N) and (ASource[I + 1] = '*') then
    begin
      Result[I] := ' '; Inc(I);
      Result[I] := ' '; Inc(I);
      while I <= N do
      begin
        if (ASource[I] = '*') and (I < N) and (ASource[I + 1] = ')') then
        begin
          Result[I] := ' '; Inc(I);
          Result[I] := ' '; Inc(I);
          Break;
        end;
        Blank(I);
        Inc(I);
      end;
    end
    else if C = '''' then
    begin
      Result[I] := ' '; Inc(I);
      while I <= N do
      begin
        if ASource[I] = '''' then
        begin
          Result[I] := ' '; Inc(I);
          if (I <= N) and (ASource[I] = '''') then
          begin
            Result[I] := ' '; Inc(I);
            Continue;
          end;
          Break;
        end;
        Blank(I);
        Inc(I);
      end;
    end
    else
    begin
      Result[I] := C;
      Inc(I);
    end;
  end;
end;

function IsPascalIdentChar(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
end;

function ExtractUsedUnitNames(const ASource: string): TArray<string>;
// Returns every identifier from every uses clause in ASource.
// Dotted names (System.SysUtils) are kept as one entry. The "in '...'"
// suffix is stripped.
var
  Clean, Lower, Body, Name: string;
  P, EndPos, K, Ch: Integer;
  Names: TList<string>;
  Parts: TArray<string>;
begin
  Clean := StripPascalCommentsAndStrings(ASource);
  Lower := AnsiLowerCase(Clean);
  Names := TList<string>.Create;
  try
    P := 1;
    while P <= Length(Lower) - 3 do
    begin
      P := Pos('uses', Lower, P);
      if P = 0 then Break;

      // word-boundary check
      if ((P > 1) and IsPascalIdentChar(Lower[P - 1])) or
         ((P + 4 <= Length(Lower)) and IsPascalIdentChar(Lower[P + 4])) then
      begin
        Inc(P);
        Continue;
      end;

      EndPos := Pos(';', Clean, P + 4);
      if EndPos = 0 then Break;

      Body := Copy(Clean, P + 4, EndPos - (P + 4));
      Parts := Body.Split([',']);
      for K := 0 to High(Parts) do
      begin
        Name := Trim(Parts[K]);
        // strip "in '...'" - cut at first whitespace
        for Ch := 1 to Length(Name) do
          if (Name[Ch] = ' ') or (Name[Ch] = #9) or
             (Name[Ch] = #10) or (Name[Ch] = #13) then
          begin
            Name := Copy(Name, 1, Ch - 1);
            Break;
          end;
        Name := Trim(Name);
        if Name <> '' then
          Names.Add(Name);
      end;

      P := EndPos + 1;
    end;
    Result := Names.ToArray;
  finally
    Names.Free;
  end;
end;

type
  TUsageHit = record
    Line: Integer; // 0-based
    Col: Integer;  // 0-based
    NameLen: Integer;
    Name: string;
  end;

function BuildLineStarts(const ASource: string): TArray<Integer>;
// Returns a 0-based array of 1-based positions where each line starts
// in ASource. LineStarts[0] = 1 (the start of line 0). A line break is
// any of: CRLF (counted once), lone LF, lone CR. The position stored
// for line k is the index of the first character of line k.
var
  N, I: Integer;
  Starts: TList<Integer>;
begin
  Starts := TList<Integer>.Create;
  try
    Starts.Add(1);
    N := Length(ASource);
    I := 1;
    while I <= N do
    begin
      if ASource[I] = #13 then
      begin
        if (I < N) and (ASource[I + 1] = #10) then
          Inc(I, 2)
        else
          Inc(I);
        Starts.Add(I);
      end
      else if ASource[I] = #10 then
      begin
        Inc(I);
        Starts.Add(I);
      end
      else
        Inc(I);
    end;
    Result := Starts.ToArray;
  finally
    Starts.Free;
  end;
end;

function PositionToLine(const ALineStarts: TArray<Integer>; APos: Integer): Integer;
// Binary search: which line (0-based) contains the 1-based position APos?
var
  Lo, Hi, Mid: Integer;
begin
  Lo := 0;
  Hi := High(ALineStarts);
  Result := 0;
  while Lo <= Hi do
  begin
    Mid := (Lo + Hi) shr 1;
    if ALineStarts[Mid] <= APos then
    begin
      Result := Mid;
      Lo := Mid + 1;
    end
    else
      Hi := Mid - 1;
  end;
end;

function FindUnitUsages(const ASource: string;
  const ANameSet: TDictionary<string, Boolean>): TArray<TUsageHit>;
// Returns every position in ASource where an identifier token from
// ANameSet appears outside comments, string literals and uses
// clauses. Line/Col come from a dedicated index over the ORIGINAL
// source (independent of comment-stripping) so they never drift.
var
  Clean: string;
  LineStarts: TArray<Integer>;
  P, N, IdentStart, LineNo: Integer;
  Token, TokenUpper: string;
  InUses: Boolean;
  Hits: TList<TUsageHit>;
  Hit: TUsageHit;

  function IsIdStart(C: Char): Boolean; inline;
  begin
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '_']);
  end;

  function IsIdCont(C: Char): Boolean; inline;
  begin
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
  end;

begin
  Clean := StripPascalCommentsAndStrings(ASource);
  N := Length(Clean);
  // Build line table from the ORIGINAL source. Comment stripping
  // preserves CR/LF in theory, but using the original removes any
  // chance of a stripping edge case shifting line counts.
  LineStarts := BuildLineStarts(ASource);
  Hits := TList<TUsageHit>.Create;
  try
    InUses := False;
    P := 1;
    while P <= N do
    begin
      // Reset uses-state at semicolons; line breaks no longer reset
      // anything (we look the line up later).
      if Clean[P] = ';' then
      begin
        if InUses then InUses := False;
        Inc(P);
        Continue;
      end;

      if IsIdStart(Clean[P]) then
      begin
        IdentStart := P;
        while (P <= N) and IsIdCont(Clean[P]) do Inc(P);
        Token := Copy(Clean, IdentStart, P - IdentStart);
        TokenUpper := AnsiUpperCase(Token);

        if TokenUpper = 'USES' then
          InUses := True
        else if (not InUses) and ANameSet.ContainsKey(TokenUpper) then
        begin
          LineNo := PositionToLine(LineStarts, IdentStart);
          Hit.Line := LineNo;
          Hit.Col := IdentStart - LineStarts[LineNo];
          Hit.NameLen := Length(Token);
          Hit.Name := Token;
          Hits.Add(Hit);
        end;
        Continue;
      end;

      Inc(P);
    end;
    Result := Hits.ToArray;
  finally
    Hits.Free;
  end;
end;

function FileUsesUnit(const AFile, AUnitName: string): Boolean;
// True if AFile's uses clauses list AUnitName (case-insensitive,
// matches dotted names by their full identifier).
var
  Source: string;
  Names: TArray<string>;
  N: string;
begin
  Result := False;
  try
    Source := TDelphiFileEncoding.ReadAll(AFile);
  except
    Exit;
  end;
  Names := ExtractUsedUnitNames(Source);
  for N in Names do
    if SameText(N, AUnitName) then
      Exit(True);
end;


procedure TLspFindUnitReferencesWizard.SearchAndShow;
var
  DelphiLspJson, RootPath, TargetExpanded, TargetUnitName: string;
  ProjFiles: TArray<string>;
  UsingFiles: TList<string>;        // expanded paths of files that use the target
  UsingFileSet: TDictionary<string, Boolean>; // upper(path) -> True
  Symbols: TList<TSymbolPos>;
  SymSeen: TDictionary<string, Boolean>;
  HitsByFile: TObjectDictionary<string, TList<TUnitRefItem>>;
  ResultSeen: TDictionary<string, Boolean>;
  LineCache: TDictionary<string, TArray<string>>;
  Client: TLspClient;
  SymbolsJson: TJSONArray;
  I: Integer;
  Sym: TSymbolPos;
  LspLocations: TArray<TLspLocation>;
  Loc: TLspLocation;
  LocPath, LocPathExpanded, LocKey, UpKey: string;
  Lines: TArray<string>;
  Item: TUnitRefItem;
  WasRunning: Boolean;
  RawHitTotal, DroppedTarget, DroppedNonProject, DroppedDup: Integer;
  FinalItems: TList<TUnitRefItem>;
  FileList: TList<TUnitRefItem>;
begin
  DelphiLspJson := Editor.FindDelphiLspJson;
  if DelphiLspJson = '' then
  begin
    FDialog.SetStatus('No .delphilsp.json found - enable Tools > Options > '
      + 'Editor > Language > Code Insight > "Generate LSP Config".');
    Exit;
  end;

  RootPath := FContext.ProjectRoot;
  if RootPath = '' then
    RootPath := ExtractFilePath(FContext.FileName);

  TargetExpanded := ExpandFileName(FContext.FileName);

  // Save editor changes so the LSP sees the same content.
  FDialog.SetStatus('Saving all files...');
  Editor.SaveAllFiles;

  WasRunning := TLspManager.Instance.IsAlive;
  if WasRunning then
    FDialog.SetStatus('LSP already running. Opening file...')
  else
    FDialog.SetStatus('Starting LSP server (one-time)...');

  Client := TLspManager.Instance.GetClient(
    RootPath, FContext.ProjectFile, DelphiLspJson);

  // Hook the LSP traffic logger for the duration of this search.
  OpenTrace(RootPath);
  TraceNote(Format('target unit: %s', [FContext.FileName]));
  TraceNote(Format('project root: %s', [RootPath]));
  TraceNote(Format('LSP already running: %s', [BoolToStr(WasRunning, True)]));
  Client.OnLog := procedure(const ADirection, AMethod, ABody: string)
    begin
      TraceLine(ADirection, AMethod, ABody);
    end;
  Client.Verbose := True;

  TraceNote('RefreshDocument(target)');
  Client.RefreshDocument(FContext.FileName);

  // Wait for the lexical index (documentSymbol). We don't need the
  // cross-reference index at all - usages are resolved via a
  // textual identifier scan against the symbol-name set.
  if not WasRunning then
  begin
    for var Retry := 1 to 30 do
    begin
      FDialog.SetStatus(Format('Waiting for LSP indexing... (%d/30)', [Retry]));
      Application.ProcessMessages;
      var ProbeOk := False;
      try
        var ProbeJson := Client.GetDocumentSymbols(FContext.FileName);
        try
          ProbeOk := (ProbeJson <> nil) and (ProbeJson.Count > 0);
        finally
          ProbeJson.Free;
        end;
      except end;
      if ProbeOk then Break;
      Sleep(1000);
    end;
  end
  else
    Sleep(300);

  TargetUnitName := ChangeFileExt(ExtractFileName(FContext.FileName), '');

  // ============================================================
  // Pass 1 (deterministic): scan all project files' uses clauses
  // to find every file that lists the target unit. This works
  // even when LSP can't resolve a symbol and finds dead uses.
  // ============================================================
  FDialog.SetStatus('Scanning project files for uses of ' + TargetUnitName + '...');
  Application.ProcessMessages;

  ProjFiles := Editor.GetProjectSourceFiles;
  UsingFiles := TList<string>.Create;
  UsingFileSet := TDictionary<string, Boolean>.Create;
  Symbols := TList<TSymbolPos>.Create;
  SymSeen := TDictionary<string, Boolean>.Create;
  HitsByFile := TObjectDictionary<string, TList<TUnitRefItem>>.Create([doOwnsValues]);
  ResultSeen := TDictionary<string, Boolean>.Create;
  LineCache := TDictionary<string, TArray<string>>.Create;
  FinalItems := TList<TUnitRefItem>.Create;
  try
    FDialog.SetProgress(0, System.Length(ProjFiles));
    for I := 0 to High(ProjFiles) do
    begin
      var ExpProj := ExpandFileName(ProjFiles[I]);
      FDialog.SetProgress(I + 1, System.Length(ProjFiles));
      FDialog.SetStatus(Format('Scanning uses clauses %d/%d...',
        [I + 1, System.Length(ProjFiles)]));
      Application.ProcessMessages;

      // Skip the target file itself.
      if SameText(ExpProj, TargetExpanded) then Continue;

      // Only .pas files have uses clauses.
      if not SameText(ExtractFileExt(ExpProj), '.pas') then Continue;

      if FileUsesUnit(ExpProj, TargetUnitName) then
      begin
        UpKey := AnsiUpperCase(ExpProj);
        if not UsingFileSet.ContainsKey(UpKey) then
        begin
          UsingFileSet.Add(UpKey, True);
          UsingFiles.Add(ExpProj);
        end;
      end;
    end;

    if UsingFiles.Count = 0 then
    begin
      FDialog.SetItems(nil);
      FDialog.SetStatus(Format(
        'No project file uses %s. (scanned %d file(s))',
        [TargetUnitName, System.Length(ProjFiles)]));
      FDialog.SetProgress(0, 1);
      Exit;
    end;

    // ============================================================
    // Pass 2 (LSP documentSymbol): obtain the names of all symbols
    // declared in the target unit. We use only the names; positions
    // are not used because DelphiLSP's findReferences doesn't react
    // to positions derived from selectionRange.
    // ============================================================
    FDialog.SetStatus('Querying unit symbols...');
    Application.ProcessMessages;

    SymbolsJson := nil;
    try
      SymbolsJson := Client.GetDocumentSymbols(FContext.FileName);
    except
      on E: Exception do
        FDialog.SetStatus('LSP error on documentSymbol: ' + E.Message);
    end;

    try
      CollectSymbols(SymbolsJson, Symbols, SymSeen);
    finally
      SymbolsJson.Free;
    end;

    // Build an uppercased name set; exclude the unit's own name to
    // avoid matching the bare unit identifier on the uses line of
    // an indirect file (the uses-line is filtered anyway, but a name
    // collision elsewhere - e.g. a variable of the same name - is
    // ambiguous and would just produce noise).
    var NameSet := TDictionary<string, Boolean>.Create;
    var TargetUpper := AnsiUpperCase(TargetUnitName);
    try
      // documentSymbol returns names that include signatures and class
      // qualifiers, e.g.
      //   "IsValidIdentifier(const AName: string): Boolean"   (method)
      //   "TFoo.IsValidIdentifier(const AName: string): ..."  (impl)
      //   "FProjectFiles: TArray<string>"                     (field)
      //   "ResultFiles: TArray<string>"                       (property)
      // We need just the bare identifier so the token-scanner can match
      // it against words in the using files. Cut at '(' (signature) and
      // at ':' (typed field/property), then take the last segment after
      // '.' (Class.Method -> Method).
      for Sym in Symbols do
      begin
        // Only accept LSP SymbolKinds that correspond to real declarations.
        // Reject File(1), Module/section(2), Package(4), Namespace(3),
        // String(15), Number(16), Boolean(17), Array(18), Null(21), Key(20),
        // EnumMember(22), Object(19), Event(24), TypeParameter(26).
        // Accept: Class(5), Method(6), Property(7), Field(8), Constructor(9),
        // Enum(10), Interface(11), Function(12), Variable(13), Constant(14),
        // Struct/Record(23), Operator(25).
        case Sym.Kind of
          5,6,7,8,9,10,11,12,13,14,23,25: ; // accepted declaration kinds
        else
          Continue;
        end;

        var Bare := Sym.Name;
        var CutAt := Pos('(', Bare);
        if CutAt > 0 then Bare := Copy(Bare, 1, CutAt - 1);
        CutAt := Pos(':', Bare);
        if CutAt > 0 then Bare := Copy(Bare, 1, CutAt - 1);
        CutAt := Pos('<', Bare);
        if CutAt > 0 then Bare := Copy(Bare, 1, CutAt - 1);
        var DotPos2 := LastDelimiter('.', Bare);
        if DotPos2 > 0 then Bare := Copy(Bare, DotPos2 + 1, MaxInt);
        Bare := Trim(Bare);
        if (Bare = '') or not IsPascalIdentChar(Bare[1]) then Continue;
        // skip section markers ("interface", "implementation", "uses")
        // and the unit's own name.
        if SameText(Bare, 'interface') or SameText(Bare, 'implementation') or
           SameText(Bare, 'uses') or SameText(Bare, 'initialization') or
           SameText(Bare, 'finalization') or SameText(Bare, TargetUnitName) then
          Continue;
        NameSet.AddOrSetValue(AnsiUpperCase(Bare), True);
      end;

      // Log the final name set (sorted) so we can verify exactly what
      // the text-scan will match against.
      var Names := TList<string>.Create;
      try
        for var N in NameSet.Keys do Names.Add(N);
        Names.Sort;
        TraceNote(Format('NameSet has %d entries:', [Names.Count]));
        for var N in Names do TraceNote('  ' + N);
      finally
        Names.Free;
      end;

      // Also handle dotted unit names: the last segment is what
      // appears in code (`Foo.Bar` → bare `Bar` calls don't make
      // sense, but we still want to be safe).
      var DotPos := LastDelimiter('.', TargetUpper);
      if DotPos > 0 then
        NameSet.Remove(Copy(TargetUpper, DotPos + 1, MaxInt))
      else
        NameSet.Remove(TargetUpper);

      RawHitTotal := 0;
      DroppedTarget := 0;
      DroppedNonProject := 0;
      DroppedDup := 0;

      // ============================================================
      // Pass 3: per using-file, scan the source for identifiers
      // that match any name in NameSet, then verify each candidate
      // via LSP GotoDefinition. A candidate counts as a real usage
      // only when the LSP resolves it to a declaration inside the
      // target unit. This filters false positives caused by common
      // identifier names (Create, Free, etc.) that may resolve to
      // a different unit's symbol.
      // ============================================================
      FDialog.SetProgress(0, UsingFiles.Count);

      // Collect all candidates first, then verify with progress so
      // the user sees per-candidate feedback during the LSP roundtrips.
      var AllCandidates := TList<TUnitRefItem>.Create;
      try
        for I := 0 to UsingFiles.Count - 1 do
        begin
          var UF := UsingFiles[I];
          FDialog.SetProgress(I + 1, UsingFiles.Count);
          FDialog.SetStatus(Format('Scanning %d/%d: %s',
            [I + 1, UsingFiles.Count, ExtractFileName(UF)]));
          Application.ProcessMessages;

          var FileSource: string := '';
          try
            FileSource := TDelphiFileEncoding.ReadAll(UF);
          except
            FileSource := '';
          end;
          if FileSource = '' then Continue;

          var Usages := FindUnitUsages(FileSource, NameSet);
          Inc(RawHitTotal, System.Length(Usages));
          if System.Length(Usages) = 0 then Continue;

          Lines := ReadCachedLines(UF, LineCache);
          for var U in Usages do
          begin
            Item := Default(TUnitRefItem);
            Item.Identifier := U.Name;
            Item.FilePath := UF;
            Item.Line := U.Line;
            Item.Col := U.Col;
            Item.Length := U.NameLen;
            if (U.Line >= 0) and (U.Line < System.Length(Lines)) then
              Item.Preview := Trim(Lines[U.Line])
            else
              Item.Preview := '';
            AllCandidates.Add(Item);
          end;
        end;

        // Verify each candidate via GotoDefinition serially against the
        // main LSP client. The parallel worker-pool path was removed
        // because DelphiLSP doesn't tolerate concurrent load (server
        // not responding / internal errors / request removed).
        FDialog.SetProgress(0, AllCandidates.Count);
        var DroppedNotResolving: Integer := 0;

        var LastRefreshedFile: string := '';
        for I := 0 to AllCandidates.Count - 1 do
        begin
          if FDialog.CloseRequested then Break;
          Item := AllCandidates[I];
          FDialog.SetProgress(I + 1, AllCandidates.Count);
          FDialog.SetStatus(Format('Verifying %d/%d (%s)...',
            [I + 1, AllCandidates.Count, Item.Identifier]));
          Application.ProcessMessages;

          if not SameText(Item.FilePath, LastRefreshedFile) then
          begin
            TraceNote(Format('RefreshDocument: %s', [Item.FilePath]));
            try
              Client.RefreshDocument(Item.FilePath);
            except
              on E: Exception do
                TraceNote('RefreshDocument FAILED: ' + E.Message);
            end;
            LastRefreshedFile := Item.FilePath;
          end;

          var Resolves := False;
          var T0 := Now;
          var DefCount := 0;
          var DefPathLog := '';
          var ErrLog := '';
          try
            var Defs := Client.GotoDefinition(Item.FilePath, Item.Line, Item.Col);
            DefCount := System.Length(Defs);
            if DefCount > 0 then
            begin
              var DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
              DefPathLog := DefPath;
              if (DefPath <> '') and
                 SameText(ExpandFileName(DefPath), TargetExpanded) then
                Resolves := True;
            end;
          except
            on E: Exception do
              ErrLog := E.ClassName + ': ' + E.Message;
          end;
          var Elapsed := (Now - T0) * SecsPerDay * 1000;
          TraceNote(Format(
            'GotoDef cand %d/%d: %s @ %s:%d:%d -> defs=%d match=%s elapsed=%.0fms %s',
            [I + 1, AllCandidates.Count, Item.Identifier,
             ExtractFileName(Item.FilePath), Item.Line, Item.Col,
             DefCount, BoolToStr(Resolves, True), Elapsed,
             IfThen(ErrLog <> '', '[ERR ' + ErrLog + ']',
                    IfThen(DefPathLog <> '',
                           '-> ' + ExtractFileName(DefPathLog), ''))]));

          if not Resolves then
          begin
            Inc(DroppedNotResolving);
            Continue;
          end;

          UpKey := AnsiUpperCase(Item.FilePath);
          LocKey := Format('%s|%d|%d', [UpKey, Item.Line, Item.Col]);
          if ResultSeen.ContainsKey(LocKey) then
          begin
            Inc(DroppedDup);
            Continue;
          end;
          ResultSeen.Add(LocKey, True);

          if not HitsByFile.TryGetValue(UpKey, FileList) then
          begin
            FileList := TList<TUnitRefItem>.Create;
            HitsByFile.Add(UpKey, FileList);
          end;
          FileList.Add(Item);
        end;
        // Stash the count for the final status line via the existing
        // counter (repurposed: "non-using" no longer applies).
        DroppedNonProject := DroppedNotResolving;
      finally
        AllCandidates.Free;
      end;
    finally
      NameSet.Free;
    end;

    // ============================================================
    // Build final list: for each using-file emit either its hits
    // (sorted) or a single "dead" placeholder row.
    // ============================================================
    UsingFiles.Sort(TComparer<string>.Construct(
      function(const A, B: string): Integer
      begin
        Result := CompareText(A, B);
      end));

    var DeadCount: Integer := 0;
    for var UF in UsingFiles do
    begin
      UpKey := AnsiUpperCase(UF);
      if HitsByFile.TryGetValue(UpKey, FileList) and (FileList.Count > 0) then
      begin
        FileList.Sort(TComparer<TUnitRefItem>.Construct(
          function(const A, B: TUnitRefItem): Integer
          begin
            Result := A.Line - B.Line;
            if Result = 0 then Result := A.Col - B.Col;
          end));
        for var H in FileList do
          FinalItems.Add(H);
      end
      else
      begin
        Item := Default(TUnitRefItem);
        Item.IsDead := True;
        Item.Identifier := '';
        Item.FilePath := UF;
        Item.Line := 0;
        Item.Col := 0;
        Item.Length := 0;
        Item.Preview := Format(
          '%s is listed in the uses clause but no symbols of it are used here.',
          [TargetUnitName]);
        FinalItems.Add(Item);
        Inc(DeadCount);
      end;
    end;

    FDialog.SetItems(FinalItems.ToArray);

    var LiveCount: Integer := UsingFiles.Count - DeadCount;
    FDialog.SetStatus(Format(
      '%d using unit(s): %d active, %d dead.  ' +
      '[symbols=%d, raw matches=%d, verified-elsewhere=%d]',
      [UsingFiles.Count, LiveCount, DeadCount,
       Symbols.Count, RawHitTotal, DroppedNonProject]));
  finally
    FinalItems.Free;
    HitsByFile.Free;
    ResultSeen.Free;
    LineCache.Free;
    UsingFiles.Free;
    UsingFileSet.Free;
    Symbols.Free;
    SymSeen.Free;
  end;
end;

end.
