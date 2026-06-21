(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Standalone.EditorHelper;

// IEditorHelper implementation for the standalone executable.
//
// Backed by TStandaloneProjectState: a simple in-memory model of the
// currently open project (the .dpr / .dproj path plus the list of
// .pas files) and an "active file + cursor" pair set by the main
// form whenever the user switches editor tabs.
//
// The IDE-niceties (SaveAllFiles, ReloadModifiedFiles, NotifyClass-
// StructureChanged) collapse to no-ops here because there is no IDE
// editor on the other side. ReplaceFileContent writes straight to
// disk, ReadEditorContent returns the editor's live buffer if the
// file is open in a tab.

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.StrUtils,
  System.Generics.Collections,
  Expert.EditorHelperIntf;

type
  /// <summary>Snapshot of the editor's current selection. Filled by
  ///  the main form on demand from the Memo's SelStart / SelLength.
  ///  Line / Col are 1-based; AEndLine / AEndCol point one past the
  ///  last character (LSP-range-end style).</summary>
  TStandaloneSelection = record
    HasSelection: Boolean;
    FileName: string;
    StartLine, StartCol: Integer;
    EndLine, EndCol: Integer;
    Text: string;
  end;

  /// <summary>Callback the standalone main form installs so the
  ///  IEditorHelper can synchronously ask "what's selected right now".
  ///  The Memo is a VCL control, so the form is the only object that
  ///  can read SelStart / SelLength.</summary>
  TGetSelectionFunc = reference to function: TStandaloneSelection;

  /// <summary>Minimal "open project" model used by the standalone
  ///  executable. The main form owns one of these and feeds it to
  ///  TStandaloneEditorHelper. Wizards call methods on Editor; Editor
  ///  reads from this state.</summary>
  TStandaloneProjectState = class
  private
    FProjectFile: string;
    FProjectRoot: string;
    FSourceFiles: TArray<string>;
    FSearchPaths: string;
    FActiveFile: string;
    FActiveLine, FActiveCol: Integer;
    /// <summary>Map: filename (case-insensitive) -> live content of
    ///  that tab's editor buffer. The main form populates and updates
    ///  this whenever a tab becomes dirty.</summary>
    FOpenBuffers: TDictionary<string, string>;
    FGetSelection: TGetSelectionFunc;
    FOnExternalChange: TProc<string>;
    function NormKey(const AFile: string): string;
  public
    constructor Create;
    destructor Destroy; override;
    procedure LoadFromDproj(const ADprojPath: string);
    procedure SetActiveFile(const AFile: string; ALine, ACol: Integer);
    /// <summary>Updates the cursor (Line/Col) only if AFile is the
    ///  currently active file. Used by the standalone editor helper
    ///  after a write: the cursor should sit at the end of the just-
    ///  inserted text. We do NOT switch the active file - editing a
    ///  background file (e.g. cross-file rename) must not pull focus.</summary>
    procedure UpdateCursorIfActive(const AFile: string; ALine, ACol: Integer);
    procedure UpdateBuffer(const AFile, AContent: string);
    procedure CloseBuffer(const AFile: string);
    /// <summary>Tells subscribers (typically the main form) that AFile
    ///  was changed from outside the editor control - i.e. a wizard
    ///  wrote through Editor.ReplaceFileContent. Allows the main form
    ///  to refresh its Memo even when the change happened
    ///  asynchronously (e.g. after the completion popup committed an
    ///  insertion, long after the wizard's Execute had returned).
    ///
    ///  UpdateBuffer is intentionally silent so the form's Memo
    ///  OnChange handler can keep calling it on every keystroke
    ///  without re-entering the reload path.</summary>
    procedure NotifyExternalChange(const AFile: string);

    property ProjectFile: string read FProjectFile;
    property ProjectRoot: string read FProjectRoot;
    property SourceFiles: TArray<string> read FSourceFiles;
    property SearchPaths: string read FSearchPaths;
    property ActiveFile: string read FActiveFile;
    property ActiveLine: Integer read FActiveLine;
    property ActiveCol: Integer read FActiveCol;
    property GetSelectionFunc: TGetSelectionFunc read FGetSelection write FGetSelection;
    property OnExternalChange: TProc<string> read FOnExternalChange write FOnExternalChange;
    function TryGetBuffer(const AFile: string; out AContent: string): Boolean;
  end;

  TStandaloneEditorHelper = class(TInterfacedObject, IEditorHelper)
  private
    FState: TStandaloneProjectState;
  public
    constructor Create(AState: TStandaloneProjectState);

    function GetCurrentContext: TEditorContext;
    function GetCurrentProjectDproj: string;
    function GetProjectRoot: string;
    function GetProjectSearchPaths: string;
    function GetProjectSourceFiles: TArray<string>;
    function BuildSearchPathFromProject(
      const ADprojPath, ARootPath: string): string;
    function FindDelphiLspJson: string;

    function ReadEditorContent(const AFilePath: string; out AContent: string): Boolean;
    function ReplaceFileContent(const AFilePath: string;
      const ANewContent: string): Boolean;
    function ReplaceSelection(const AFilePath: string;
      AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      const ANewText: string): Boolean;
    function ReplaceLineAt(const AFilePath: string; ALine: Integer;
      const ANewContent: string): Boolean;
    function DeleteLineAt(const AFilePath: string; ALine: Integer): Boolean;
    function InsertTextAtLineStart(const AFilePath: string;
      ALine: Integer; const AText: string): Boolean;
    function ApplyEditViaEditor(const AFilePath: string;
      ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;

    procedure SaveAllFiles;
    procedure ReloadModifiedFiles(const FilePaths: TArray<string>);
    procedure NotifyClassStructureChanged(const AFilePath: string);
    function GotoLocation(const AFilePath: string;
      ALine, ACol: Integer; AHighlightLen: Integer = 0): Boolean;
    function AddFileToActiveProject(const AFilePath: string): Boolean;
    function GetSelection(out AFilePath: string;
      out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
      out AText: string): Boolean;
  end;

implementation

uses
  Xml.XMLDoc, Xml.XMLIntf, Delphi.FileEncoding;

{ TStandaloneProjectState }

constructor TStandaloneProjectState.Create;
begin
  inherited Create;
  FOpenBuffers := TDictionary<string, string>.Create;
  FActiveLine := 1;
  FActiveCol := 1;
end;

destructor TStandaloneProjectState.Destroy;
begin
  FOpenBuffers.Free;
  inherited;
end;

function TStandaloneProjectState.NormKey(const AFile: string): string;
begin
  Result := UpperCase(AFile);
end;

procedure TStandaloneProjectState.LoadFromDproj(const ADprojPath: string);
var
  Doc: IXMLDocument;
  Root, Node, Cfg, Group, Item: IXMLNode;
  Files: TList<string>;
  I, J: Integer;
  Filename: string;
begin
  FProjectFile := ADprojPath;
  FProjectRoot := ExtractFilePath(ADprojPath);
  FSourceFiles := nil;
  FSearchPaths := '';
  if not TFile.Exists(ADprojPath) then Exit;

  Files := TList<string>.Create;
  Doc := TXMLDocument.Create(nil);
  try
    Doc.LoadFromFile(ADprojPath);
    Root := Doc.DocumentElement;
    if Root = nil then Exit;
    // Iterate <ItemGroup> with <DCCReference Include="..."/>
    for I := 0 to Root.ChildNodes.Count - 1 do
    begin
      Node := Root.ChildNodes[I];
      if not SameText(Node.NodeName, 'ItemGroup') then Continue;
      for J := 0 to Node.ChildNodes.Count - 1 do
      begin
        Item := Node.ChildNodes[J];
        if SameText(Item.NodeName, 'DCCReference') then
        begin
          Filename := Item.Attributes['Include'];
          if Filename = '' then Continue;
          if TPath.IsRelativePath(Filename) then
            Filename := TPath.GetFullPath(
              IncludeTrailingPathDelimiter(FProjectRoot) + Filename);
          if EndsText('.pas', Filename) or EndsText('.dpr', Filename) or
             EndsText('.dpk', Filename) then
            Files.Add(Filename);
        end;
      end;
    end;
    // Pull DCC_UnitSearchPath from <PropertyGroup>s; cheap approach -
    // grep the whole XML text for the value of that element. A full
    // PropertyGroup walk would be more correct but adds complexity.
    Cfg := Root;
    Group := nil;
    for I := 0 to Cfg.ChildNodes.Count - 1 do
      if SameText(Cfg.ChildNodes[I].NodeName, 'PropertyGroup') then
        for J := 0 to Cfg.ChildNodes[I].ChildNodes.Count - 1 do
          if SameText(Cfg.ChildNodes[I].ChildNodes[J].NodeName, 'DCC_UnitSearchPath') then
            FSearchPaths := FSearchPaths + ';' + Cfg.ChildNodes[I].ChildNodes[J].Text;
    FSourceFiles := Files.ToArray;
  finally
    Files.Free;
  end;
  if Group <> nil then ;
end;

procedure TStandaloneProjectState.SetActiveFile(const AFile: string;
  ALine, ACol: Integer);
begin
  FActiveFile := AFile;
  FActiveLine := ALine;
  FActiveCol := ACol;
end;

procedure TStandaloneProjectState.UpdateCursorIfActive(const AFile: string;
  ALine, ACol: Integer);
begin
  if SameText(AFile, FActiveFile) then
  begin
    FActiveLine := ALine;
    FActiveCol := ACol;
  end;
end;

procedure TStandaloneProjectState.UpdateBuffer(const AFile, AContent: string);
begin
  FOpenBuffers.AddOrSetValue(NormKey(AFile), AContent);
end;

procedure TStandaloneProjectState.CloseBuffer(const AFile: string);
begin
  FOpenBuffers.Remove(NormKey(AFile));
end;

procedure TStandaloneProjectState.NotifyExternalChange(const AFile: string);
begin
  if Assigned(FOnExternalChange) then
    FOnExternalChange(AFile);
end;

function TStandaloneProjectState.TryGetBuffer(const AFile: string;
  out AContent: string): Boolean;
begin
  Result := FOpenBuffers.TryGetValue(NormKey(AFile), AContent);
end;

{ TStandaloneEditorHelper }

constructor TStandaloneEditorHelper.Create(AState: TStandaloneProjectState);
begin
  inherited Create;
  FState := AState;
end;

function TStandaloneEditorHelper.GetCurrentContext: TEditorContext;
begin
  Result := Default(TEditorContext);
  Result.FileName := FState.ActiveFile;
  Result.Line := FState.ActiveLine;
  Result.Column := FState.ActiveCol;
  Result.ProjectFile := FState.ProjectFile;
  Result.ProjectRoot := FState.ProjectRoot;
  Result.IsValid := (Result.FileName <> '') and TFile.Exists(Result.FileName);
end;

function TStandaloneEditorHelper.GetCurrentProjectDproj: string;
begin
  Result := FState.ProjectFile;
end;

function TStandaloneEditorHelper.GetProjectRoot: string;
begin
  Result := FState.ProjectRoot;
end;

function TStandaloneEditorHelper.GetProjectSearchPaths: string;
begin
  Result := FState.SearchPaths;
end;

function TStandaloneEditorHelper.GetProjectSourceFiles: TArray<string>;
begin
  Result := FState.SourceFiles;
end;

function TStandaloneEditorHelper.BuildSearchPathFromProject(
  const ADprojPath, ARootPath: string): string;
begin
  // For standalone we just return the search paths already extracted
  // from the .dproj at load time. A more thorough implementation would
  // walk every directory and add .pas-containing folders, but the
  // simple version is enough for first runs.
  Result := FState.SearchPaths;
end;

function TStandaloneEditorHelper.FindDelphiLspJson: string;
var
  Candidate: string;
begin
  Result := '';
  if FState.ProjectRoot = '' then Exit;
  // Convention: <projectname>.delphilsp.json next to the .dpr/.dproj.
  Candidate := ChangeFileExt(FState.ProjectFile, '.delphilsp.json');
  if TFile.Exists(Candidate) then Exit(Candidate);
  Candidate := IncludeTrailingPathDelimiter(FState.ProjectRoot) + '.delphilsp.json';
  if TFile.Exists(Candidate) then Exit(Candidate);
end;

function TStandaloneEditorHelper.ReadEditorContent(const AFilePath: string;
  out AContent: string): Boolean;
begin
  // True only when the file is currently open in a tab; otherwise
  // tell the caller to fall back to a disk read.
  Result := FState.TryGetBuffer(AFilePath, AContent);
end;

function TStandaloneEditorHelper.ReplaceFileContent(const AFilePath: string;
  const ANewContent: string): Boolean;
var
  Enc: TEncoding;
begin
  try
    if TFile.Exists(AFilePath) then Enc := TDelphiFileEncoding.Detect(AFilePath)
    else Enc := TEncoding.UTF8;
    TDelphiFileEncoding.WriteAll(AFilePath, ANewContent, Enc);
    // Keep the buffer in sync so the main form's editor reloads it.
    FState.UpdateBuffer(AFilePath, ANewContent);
    Result := True;
    // Tell the main form that this file changed under it. The form
    // reloads the Memo from the buffer if AFilePath is the active
    // file. This is what lets the completion popup's
    // Enter-to-accept actually show up in the editor immediately,
    // instead of only after a tab switch.
    FState.NotifyExternalChange(AFilePath);
  except
    Result := False;
  end;
end;

function TStandaloneEditorHelper.ReplaceSelection(const AFilePath: string;
  AStartLine, AStartCol, AEndLine, AEndCol: Integer;
  const ANewText: string): Boolean;
var
  Content: string;
  Lines: TStringList;
  EndLine, EndCol, P, LastBreak: Integer;
begin
  // Standalone v1: round-trip via TStringList. A direct char-offset
  // edit would be faster but the simpler version is enough for now.
  Result := False;
  if not ReadEditorContent(AFilePath, Content) then
    if TFile.Exists(AFilePath) then Content := TDelphiFileEncoding.ReadAll(AFilePath)
    else Exit;
  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    // 1-based -> 0-based for TStringList indexing
    if (AStartLine < 1) or (AStartLine > Lines.Count) then Exit;
    var L: string := Lines[AStartLine - 1];
    if (AStartCol < 1) or (AStartCol > Length(L) + 1) then Exit;
    var Before: string := Copy(L, 1, AStartCol - 1);
    var After: string := '';
    if AStartLine = AEndLine then
      After := Copy(L, AEndCol, MaxInt)
    else if (AEndLine >= 1) and (AEndLine <= Lines.Count) then
      After := Copy(Lines[AEndLine - 1], AEndCol, MaxInt);
    var NewLine: string := Before + ANewText + After;
    // Replace start line, drop the lines between, splice NewLine
    Lines[AStartLine - 1] := NewLine;
    var I: Integer := AStartLine;
    while (I < AEndLine) and (AStartLine < Lines.Count) do
    begin
      Lines.Delete(AStartLine);
      Inc(I);
    end;

    // Compute the (1-based) end position of the inserted text. The
    // caret should land here so the user can keep typing immediately
    // after a code-completion accept. We update the state's cursor
    // BEFORE calling ReplaceFileContent because ReplaceFileContent
    // fires NotifyExternalChange -> main form ReloadActiveFile, which
    // reads exactly these two fields to position the Memo's caret.
    EndLine := AStartLine;
    LastBreak := 0;
    for P := 1 to Length(ANewText) do
      if ANewText[P] = #10 then
      begin
        Inc(EndLine);
        LastBreak := P;
      end;
    if LastBreak = 0 then
      EndCol := AStartCol + Length(ANewText)
    else
      EndCol := Length(ANewText) - LastBreak + 1;
    FState.UpdateCursorIfActive(AFilePath, EndLine, EndCol);

    Result := ReplaceFileContent(AFilePath, Lines.Text);
  finally
    Lines.Free;
  end;
end;

function TStandaloneEditorHelper.ReplaceLineAt(const AFilePath: string;
  ALine: Integer; const ANewContent: string): Boolean;
var
  Content: string;
  Lines: TStringList;
begin
  Result := False;
  if not ReadEditorContent(AFilePath, Content) then
    if TFile.Exists(AFilePath) then Content := TDelphiFileEncoding.ReadAll(AFilePath)
    else Exit;
  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    if (ALine < 1) or (ALine > Lines.Count) then Exit;
    Lines[ALine - 1] := ANewContent;
    Result := ReplaceFileContent(AFilePath, Lines.Text);
  finally
    Lines.Free;
  end;
end;

function TStandaloneEditorHelper.DeleteLineAt(const AFilePath: string;
  ALine: Integer): Boolean;
var
  Content: string;
  Lines: TStringList;
begin
  Result := False;
  if not ReadEditorContent(AFilePath, Content) then
    if TFile.Exists(AFilePath) then Content := TDelphiFileEncoding.ReadAll(AFilePath)
    else Exit;
  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    if (ALine < 1) or (ALine > Lines.Count) then Exit;
    Lines.Delete(ALine - 1);
    Result := ReplaceFileContent(AFilePath, Lines.Text);
  finally
    Lines.Free;
  end;
end;

function TStandaloneEditorHelper.InsertTextAtLineStart(const AFilePath: string;
  ALine: Integer; const AText: string): Boolean;
var
  Content: string;
  Lines: TStringList;
begin
  Result := False;
  if not ReadEditorContent(AFilePath, Content) then
    if TFile.Exists(AFilePath) then Content := TDelphiFileEncoding.ReadAll(AFilePath)
    else Exit;
  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    if (ALine < 1) or (ALine > Lines.Count + 1) then Exit;
    Lines.Insert(ALine - 1, AText);
    Result := ReplaceFileContent(AFilePath, Lines.Text);
  finally
    Lines.Free;
  end;
end;

function TStandaloneEditorHelper.ApplyEditViaEditor(const AFilePath: string;
  ALine, ACol: Integer; const AOldText, ANewText: string): Boolean;
var
  Content: string;
  Lines: TStringList;
  L: string;
  P: Integer;
begin
  // ALine / ACol are 0-based here (interface contract).
  Result := False;
  if not ReadEditorContent(AFilePath, Content) then
    if TFile.Exists(AFilePath) then Content := TDelphiFileEncoding.ReadAll(AFilePath)
    else Exit;
  Lines := TStringList.Create;
  try
    Lines.Text := Content;
    if (ALine < 0) or (ALine >= Lines.Count) then Exit;
    L := Lines[ALine];
    P := ACol + 1; // 1-based for Copy/Pos
    if (P < 1) or (P > Length(L) + 1) then Exit;
    if (AOldText <> '') and (Copy(L, P, Length(AOldText)) <> AOldText) then Exit;
    L := Copy(L, 1, P - 1) + ANewText + Copy(L, P + Length(AOldText), MaxInt);
    Lines[ALine] := L;
    Result := ReplaceFileContent(AFilePath, Lines.Text);
  finally
    Lines.Free;
  end;
end;

procedure TStandaloneEditorHelper.SaveAllFiles;
// Standalone keeps buffers and disk in sync on every write, so there
// is nothing to save explicitly. No-op.
begin end;

procedure TStandaloneEditorHelper.ReloadModifiedFiles(const FilePaths: TArray<string>);
// The main form refreshes editor tabs from disk after every successful
// write; the wizard does not need to trigger anything here. No-op.
begin end;

procedure TStandaloneEditorHelper.NotifyClassStructureChanged(const AFilePath: string);
// There is no form designer to notify in standalone. No-op.
begin end;

function TStandaloneEditorHelper.GotoLocation(const AFilePath: string;
  ALine, ACol: Integer; AHighlightLen: Integer): Boolean;
begin
  // Update the active-file pointer so the main form (which observes
  // the state) can switch tabs and position the caret in its editor
  // on the next event-pump cycle.
  // Convert from LSP 0-based to our 1-based.
  FState.SetActiveFile(AFilePath, ALine + 1, ACol + 1);
  Result := True;
end;

function TStandaloneEditorHelper.AddFileToActiveProject(const AFilePath: string): Boolean;
// Adds AFilePath as a new <DCCReference Include="..."/> to the active
// .dproj. Idempotent (re-add is detected case-insensitively against
// existing entries, including those stored as relative paths).
//
// We rewrite the file via XmlDoc so existing formatting / encoding is
// preserved as much as Delphi's XML serializer allows. The reloaded
// state's source-file list is refreshed so the file tree picks up the
// new entry on the next RefreshTree call.
var
  Doc: IXMLDocument;
  Root, Group, Item, NewItem, FirstReferenceGroup, Existing: IXMLNode;
  I, J: Integer;
  RelTarget, AbsCandidate, ExistingInclude: string;
  Dproj: string;
begin
  Result := False;
  Dproj := FState.ProjectFile;
  if (Dproj = '') or not TFile.Exists(Dproj) then Exit;

  Doc := TXMLDocument.Create(nil);
  try
    Doc.LoadFromFile(Dproj);
    Root := Doc.DocumentElement;
    if Root = nil then Exit;

    // Build the relative path of AFilePath against the .dproj directory.
    // The IDE itself stores file references relatively when they sit
    // beneath the project; we follow the same convention so the
    // resulting .dproj stays diff-friendly.
    if SameText(ExtractFilePath(AFilePath),
                IncludeTrailingPathDelimiter(FState.ProjectRoot)) then
      RelTarget := ExtractFileName(AFilePath)
    else
      RelTarget := ExtractRelativePath(
        IncludeTrailingPathDelimiter(FState.ProjectRoot), AFilePath);

    // Idempotency check + locate first ItemGroup that already holds
    // DCCReference children. We append to that group so file refs
    // stay co-located.
    FirstReferenceGroup := nil;
    for I := 0 to Root.ChildNodes.Count - 1 do
    begin
      Group := Root.ChildNodes[I];
      if not SameText(Group.NodeName, 'ItemGroup') then Continue;
      for J := 0 to Group.ChildNodes.Count - 1 do
      begin
        Item := Group.ChildNodes[J];
        if not SameText(Item.NodeName, 'DCCReference') then Continue;
        if FirstReferenceGroup = nil then FirstReferenceGroup := Group;
        ExistingInclude := Item.Attributes['Include'];
        if ExistingInclude = '' then Continue;
        if SameText(ExistingInclude, RelTarget) or
           SameText(ExistingInclude, AFilePath) then
          Exit(True); // already present
        // Also handle case where the existing entry is relative and
        // resolves to the same absolute path.
        if TPath.IsRelativePath(ExistingInclude) then
        begin
          AbsCandidate := TPath.GetFullPath(
            IncludeTrailingPathDelimiter(FState.ProjectRoot) + ExistingInclude);
          if SameText(AbsCandidate, AFilePath) then Exit(True);
        end;
      end;
    end;

    // No existing group? Create one as last child.
    if FirstReferenceGroup = nil then
      FirstReferenceGroup := Root.AddChild('ItemGroup');

    NewItem := FirstReferenceGroup.AddChild('DCCReference');
    NewItem.Attributes['Include'] := RelTarget;

    Doc.SaveToFile(Dproj);
    Result := True;
  finally
    Doc := nil;
  end;

  // Refresh the in-memory project model so the file tree shows the
  // new file when the caller refreshes it.
  if Result then
    FState.LoadFromDproj(Dproj);
  if Existing <> nil then ;  // suppress unused warning
end;

function TStandaloneEditorHelper.GetSelection(out AFilePath: string;
  out AStartLine, AStartCol, AEndLine, AEndCol: Integer;
  out AText: string): Boolean;
// Delegates to the main form via the callback the form installed at
// startup. The form is the only object that can read Memo.SelStart /
// SelLength synchronously, so the call goes through there. If no
// callback is installed (e.g. tests, headless), returns False; the
// wizard will then show "Please select code first" and abort.
var
  Sel: TStandaloneSelection;
begin
  AFilePath := ''; AText := '';
  AStartLine := 0; AStartCol := 0; AEndLine := 0; AEndCol := 0;
  Result := False;
  if not Assigned(FState.GetSelectionFunc) then Exit;
  Sel := FState.GetSelectionFunc();
  if not Sel.HasSelection then Exit;
  AFilePath := Sel.FileName;
  AStartLine := Sel.StartLine;
  AStartCol := Sel.StartCol;
  AEndLine := Sel.EndLine;
  AEndCol := Sel.EndCol;
  AText := Sel.Text;
  Result := True;
end;

end.
