(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ExtractInterfaceWizard;

{
  Wizard plumbing for "Extract interface from class" and "Add to existing
  interface". Both entry points share the same dialog.

  Flow:
    1. Read editor cursor; parse the enclosing class via Expert.Extract-
       Interface (text-based).
    2. For Add-to-existing: scan all project units for IXxx = interface
       declarations and pass them as a picker.
    3. Open the modal dialog. User selects members + interface name +
       target file (extract) or target interface (add).
    4. On OK:
         - For ExtractNew: write the new unit file to disk; insert a
           Forms/uses reference to it into the source unit and add
           Interface to the class's ancestor list; for any selected
           field, synthesise property + getter + setter on the class.
         - For AddToExisting: splice the new declarations into the
           existing interface body; rewrite the class same way as above.
    5. All edits to the source unit run through IOTAEditWriter so they
       are undoable.
}

interface

procedure ExtractInterfaceFromClass;
procedure AddToExistingInterface;
/// <summary>For a class that does NOT descend from TInterfacedObject,
///  adds full IInterface support directly to the class: `IInterface`
///  in the ancestor list, FRefCount field, NewInstance + After-
///  Construction overrides (to mirror TInterfacedObject's
///  initial-refcount-1 trick) and the three IInterface methods
///  (QueryInterface / _AddRef / _Release). Once applied the instance
///  is held purely as an IInterface and frees itself when the last
///  reference drops.</summary>
procedure DelegateInterfaceImplementation;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.StrUtils,
  System.Character, System.Generics.Collections,
  System.UITypes,
  Vcl.Dialogs, Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ExtCtrls,
  {$IFNDEF STANDALONE_BUILD}ToolsAPI,{$ENDIF} 
  Expert.EditorHelperIntf, Expert.ExtractInterface, Expert.ExtractInterfaceDialog,
  Expert.LspManager, Lsp.Client, Lsp.Uri, Lsp.Protocol,
  Delphi.FileEncoding;

function StringListToArray(ASL: TStringList): TArray<string>;
var
  I: Integer;
begin
  SetLength(Result, ASL.Count);
  for I := 0 to ASL.Count - 1 do Result[I] := ASL[I];
end;

procedure WriteLinesEnc(const AFile: string; const ALines: TArray<string>);
var
  Enc: TEncoding;
  Content: string;
begin
  if TFile.Exists(AFile) then
    Enc := TDelphiFileEncoding.Detect(AFile)
  else
    Enc := TEncoding.UTF8;
  Content := string.Join(sLineBreak, ALines);
  TDelphiFileEncoding.WriteAll(AFile, Content, Enc);
end;

procedure WriteTextEnc(const AFile, AContent: string);
var
  Enc: TEncoding;
begin
  if TFile.Exists(AFile) then
    Enc := TDelphiFileEncoding.Detect(AFile)
  else
    Enc := TEncoding.UTF8;
  TDelphiFileEncoding.WriteAll(AFile, AContent, Enc);
end;

/// <summary>Reads the lines of a unit, preferring the IDE's live editor
///  buffer over the on-disk file. Ensures we modify the user's current
///  unsaved state, not a stale copy.</summary>
function ReadSourceLines(const AFile: string): TArray<string>;
var
  Content: string;
begin
  if Editor.ReadEditorContent(AFile, Content) then
    Result := Content.Split([sLineBreak], TStringSplitOptions.None)
  else
    Result := TDelphiFileEncoding.ReadLines(AFile);
end;

/// <summary>Writes lines back to a unit. Pushes through IOTAEditWriter
///  via Editor.ReplaceFileContent so the IDE sees the change
///  instantly and it is part of the undo stack. Falls back to a direct
///  disk write only if the file is not open in the IDE.</summary>
procedure WriteSourceLines(const AFile: string; const ALines: TArray<string>);
var
  NewContent: string;
begin
  NewContent := string.Join(sLineBreak, ALines);
  if not Editor.ReplaceFileContent(AFile, NewContent) then
    WriteLinesEnc(AFile, ALines);
end;

/// <summary>Adds AFile to the active project. Goes through the
///  IEditorHelper, which uses IOTAProject.AddFile in the IDE and a
///  direct .dproj XML edit in standalone. Idempotent.</summary>
procedure AddFileToActiveProject(const AFile: string);
begin
  Editor.AddFileToActiveProject(AFile);
end;

function IsIdentCh(C: Char): Boolean; inline;
begin
  Result := C.IsLetterOrDigit or (C = '_');
end;

function StripLineCommentLocal(const ALine: string): string;
var
  P: Integer;
begin
  Result := ALine;
  P := Pos('//', Result);
  if P > 0 then Result := Copy(Result, 1, P - 1);
end;

function IsPrimitiveType(const AUpper: string): Boolean;
begin
  Result :=
    (AUpper = 'INTEGER') or (AUpper = 'STRING') or (AUpper = 'BOOLEAN') or
    (AUpper = 'BYTE') or (AUpper = 'WORD') or (AUpper = 'CARDINAL') or
    (AUpper = 'INT64') or (AUpper = 'UINT64') or
    (AUpper = 'CHAR') or (AUpper = 'ANSICHAR') or (AUpper = 'WIDECHAR') or
    (AUpper = 'DOUBLE') or (AUpper = 'EXTENDED') or (AUpper = 'SINGLE') or
    (AUpper = 'REAL') or (AUpper = 'CURRENCY') or
    (AUpper = 'POINTER') or (AUpper = 'VARIANT') or (AUpper = 'PCHAR') or
    (AUpper = 'PANSICHAR') or (AUpper = 'PWIDECHAR') or
    (AUpper = 'NATIVEINT') or (AUpper = 'NATIVEUINT') or
    (AUpper = 'SHORTINT') or (AUpper = 'SMALLINT') or (AUpper = 'LONGINT') or
    (AUpper = 'TOBJECT');
end;

procedure CollectTypesFromSignature(const ASig: string;
  ATypes: TDictionary<string, Boolean>);
// Walks a Pascal signature, harvesting every identifier that appears
// in a type position: after a ':', inside '<...>' generic angles.
var
  I, Start: Integer;
  Ident, U: string;
begin
  I := 1;
  while I <= Length(ASig) do
  begin
    if (ASig[I] = ':') or (ASig[I] = '<') or (ASig[I] = ',') then
    begin
      Inc(I);
      while (I <= Length(ASig)) and (ASig[I] = ' ') do Inc(I);
      Start := I;
      while (I <= Length(ASig)) and (IsIdentCh(ASig[I]) or (ASig[I] = '.')) do Inc(I);
      if I > Start then
      begin
        Ident := Copy(ASig, Start, I - Start);
        U := UpperCase(Ident);
        if (Ident <> '') and Ident[1].IsLetter and not IsPrimitiveType(U) then
          ATypes.AddOrSetValue(Ident, True);
      end;
    end
    else
      Inc(I);
  end;
end;

/// <summary>LSP-based: for every type identifier referenced by the
///  selected interface members, calls textDocument/definition on a
///  position in the source file where that identifier occurs and maps
///  the resolved file path back to a unit name. Returns the dedup'd
///  list of unit names. Used to seed the new interface unit's uses
///  with exactly the units that are actually needed - not everything
///  the source unit happened to bring along.</summary>
/// <summary>Stats from one ResolveUsesViaLsp run, populated even on
///  partial failure. Used by callers to surface what actually happened
///  in the success message (so the user can tell whether LSP ran).</summary>
type
  TResolveUsesStats = record
    TypesAttempted: Integer;
    TypesResolved: Integer;
    LspAvailable: Boolean;
    LspError: string;
  end;

type
  TTypeRef = record
    TypeName: string;
    /// <summary>1-based source line where the member that referenced
    ///  this type lives. We search inside this line (and a couple of
    ///  surrounding lines for multi-line method signatures) instead of
    ///  walking the whole file.</summary>
    SourceLine: Integer;
  end;

function ResolveUsesViaLsp(const ASourceFile: string;
  const AMembers: TArray<TClassMember>;
  AStatusCallback: TProc<string>;
  out AStats: TResolveUsesStats): TArray<string>;
var
  TypeRefs: TList<TTypeRef>;
  SeenTypes: TDictionary<string, Boolean>;
  Units: TDictionary<string, Boolean>;
  Result_: TList<string>;
  SL: TStringList;
  Lines: TArray<string>;
  Client: TLspClient;
  M: TClassMember;
  Ref: TTypeRef;
  Content, Dproj, Pr, Dj: string;
  procedure Status(const S: string);
  begin
    if Assigned(AStatusCallback) then AStatusCallback(S);
  end;
  procedure AddTypeRef(const ATypeName: string; ALine: Integer);
  var
    Sig: TTypeRef;
    Key: string;
  begin
    if (ATypeName = '') or IsPrimitiveType(UpperCase(ATypeName)) then Exit;
    Key := UpperCase(ATypeName);
    if SeenTypes.ContainsKey(Key) then Exit;
    SeenTypes.Add(Key, True);
    Sig.TypeName := ATypeName;
    Sig.SourceLine := ALine;
    TypeRefs.Add(Sig);
  end;
  procedure CollectTypesAndLines(const M2: TClassMember);
  var
    Types: TDictionary<string, Boolean>;
    Tn: string;
  begin
    Types := TDictionary<string, Boolean>.Create;
    try
      CollectTypesFromSignature(M2.Signature, Types);
      if (M2.TypeName <> '') then Types.AddOrSetValue(M2.TypeName, True);
      for Tn in Types.Keys do
        AddTypeRef(Tn, M2.LineStart);
    finally
      Types.Free;
    end;
  end;
begin
  Result := nil;
  AStats := Default(TResolveUsesStats);
  TypeRefs := TList<TTypeRef>.Create;
  SeenTypes := TDictionary<string, Boolean>.Create;
  Units := TDictionary<string, Boolean>.Create;
  Result_ := TList<string>.Create;
  try
    for M in AMembers do
      if M.Selected then CollectTypesAndLines(M);
    AStats.TypesAttempted := TypeRefs.Count;
    if TypeRefs.Count = 0 then Exit;

    try
      Dproj := Editor.GetCurrentProjectDproj;
      Pr := Editor.GetProjectRoot;
      if Pr = '' then Pr := ExtractFilePath(ASourceFile);
      Dj := Editor.FindDelphiLspJson;
      Client := TLspManager.Instance.GetClient(Pr, Dproj, Dj);
      AStats.LspAvailable := True;
      Status('LSP: warming up source file for type resolution...');
      Client.EnsureFileAnalysed(ASourceFile, 30000, 10000, AStatusCallback);
    except
      on E: Exception do
      begin
        AStats.LspError := E.ClassName + ': ' + E.Message;
        Exit;
      end;
    end;

    // Use TStringList to split content - it handles CR / LF / CRLF
    // uniformly, where Split([sLineBreak]) misbehaves on mixed line
    // endings (which is exactly how the off-by-line bugs sneaked in
    // earlier: positions ended up pointing to the wrong line).
    SL := TStringList.Create;
    try
      if not Editor.ReadEditorContent(ASourceFile, Content) then
        Content := TDelphiFileEncoding.ReadAll(ASourceFile);
      SL.Text := Content;
      SetLength(Lines, SL.Count);
      for var I := 0 to SL.Count - 1 do Lines[I] := SL[I];
    finally
      SL.Free;
    end;

    for Ref in TypeRefs do
    begin
      // Search the member's known source line first. If we don't find
      // the identifier there (multi-line method signatures, comments,
      // ...) fall back to a small window around the member's line.
      var FoundLine: Integer := -1;
      var FoundCol: Integer := -1;
      var Tn := Ref.TypeName;
      var SearchStart: Integer := Ref.SourceLine - 1;       // 0-based
      var SearchEnd: Integer := Ref.SourceLine + 3;         // small look-ahead window
      if SearchStart < 0 then SearchStart := 0;
      if SearchEnd > High(Lines) then SearchEnd := High(Lines);
      for var I := SearchStart to SearchEnd do
      begin
        var L := Lines[I];
        var P: Integer := 1;
        while P > 0 do
        begin
          P := PosEx(Tn, L, P);
          if P = 0 then Break;
          var BeforeOK := (P = 1) or not IsIdentCh(L[P - 1]);
          var AfterOK := (P + Length(Tn) - 1 >= Length(L))
            or not IsIdentCh(L[P + Length(Tn)]);
          if BeforeOK and AfterOK then
          begin
            FoundLine := I;
            FoundCol := P - 1;
            Break;
          end;
          P := P + 1;
        end;
        if FoundLine >= 0 then Break;
      end;
      if FoundLine < 0 then Continue;

      Status('LSP: resolving ' + Tn + '...');
      try
        var Locs := Client.GotoDefinition(ASourceFile, FoundLine, FoundCol);
        if Length(Locs) > 0 then
        begin
          Inc(AStats.TypesResolved);
          var Path := TLspUri.FileUriToPath(Locs[0].Uri);
          if Path = '' then Continue;
          if SameText(ExtractFileName(Path), ExtractFileName(ASourceFile)) then Continue;
          var UnitName := ChangeFileExt(ExtractFileName(Path), '');
          var UU := UpperCase(UnitName);
          if (UnitName <> '') and not Units.ContainsKey(UU) then
          begin
            Units.Add(UU, True);
            Result_.Add(UnitName);
          end;
        end;
      except
        // single-type resolve failure is non-fatal
      end;
    end;
    Result := Result_.ToArray;
  finally
    TypeRefs.Free;
    SeenTypes.Free;
    Units.Free;
    Result_.Free;
  end;
end;

/// <summary>Extracts the interface-section uses clause of AFile as a
///  list of unit names (no trailing comma/semicolon, whitespace
///  trimmed). Used as a fallback when LSP resolution fails.</summary>
/// <summary>Returns A unioned with the entries of B that are not
///  already in A (case-insensitive). Used to top off LSP-resolved
///  uses with the source unit's own interface-uses as a safety net
///  when LSP could not resolve every type.</summary>
function UnionUnitNames(const A, B: TArray<string>): TArray<string>;
var
  Seen: TDictionary<string, Boolean>;
  L: TList<string>;
  S: string;
begin
  Seen := TDictionary<string, Boolean>.Create;
  L := TList<string>.Create;
  try
    for S in A do
      if S <> '' then
      begin
        if not Seen.ContainsKey(UpperCase(S)) then
        begin
          Seen.Add(UpperCase(S), True);
          L.Add(S);
        end;
      end;
    for S in B do
      if S <> '' then
      begin
        if not Seen.ContainsKey(UpperCase(S)) then
        begin
          Seen.Add(UpperCase(S), True);
          L.Add(S);
        end;
      end;
    Result := L.ToArray;
  finally
    L.Free;
    Seen.Free;
  end;
end;

function ExtractInterfaceUses(const AFile: string): TArray<string>;
var
  Lines: TArray<string>;
  I, Idx: Integer;
  U, L: string;
  InInterface: Boolean;
  Acc: string;
  Token: string;
  K: Integer;
  List: TList<string>;
begin
  Result := nil;
  Lines := TDelphiFileEncoding.ReadLines(AFile);
  InInterface := False;
  Idx := -1;
  for I := 0 to High(Lines) do
  begin
    U := UpperCase(Trim(Lines[I]));
    if U = 'INTERFACE' then InInterface := True
    else if U = 'IMPLEMENTATION' then Break;
    if InInterface and StartsText('USES', U) then begin Idx := I; Break; end;
  end;
  if Idx < 0 then Exit;
  // Collect text until ';'.
  Acc := '';
  for I := Idx to High(Lines) do
  begin
    Acc := Acc + ' ' + Lines[I];
    if Pos(';', Lines[I]) > 0 then Break;
  end;
  // Strip leading 'uses', trailing ';' and split by ','.
  L := Trim(Acc);
  if StartsText('uses', L) then Delete(L, 1, 4);
  L := Trim(L);
  if (L <> '') and (L[Length(L)] = ';') then
    L := Trim(Copy(L, 1, Length(L) - 1));
  List := TList<string>.Create;
  try
    for Token in L.Split([',']) do
    begin
      var T := Trim(Token);
      // Strip "in '<file>'" suffix on dpr-style entries.
      K := Pos(' in ', T);
      if K > 0 then T := Trim(Copy(T, 1, K - 1));
      if T <> '' then List.Add(T);
    end;
    Result := List.ToArray;
  finally
    List.Free;
  end;
end;

/// <summary>Opens AFile in the IDE so the user sees the freshly written
///  unit immediately.</summary>
procedure OpenModuleInIDE(const AFile: string);
{$IFNDEF STANDALONE_BUILD}
var
  MS: IOTAModuleServices;
{$ENDIF}
begin
{$IFNDEF STANDALONE_BUILD}
  if Supports(BorlandIDEServices, IOTAModuleServices, MS) then
    MS.OpenModule(AFile);
{$ELSE}
  // Standalone has no IDE module manager. The new file lives on disk;
  // the user can open it manually. No-op.
{$ENDIF}
end;

function GetEditorCursorLine(out AFile: string; out ALine: Integer): Boolean;
var
  Ctx: TEditorContext;
begin
  Ctx := Editor.GetCurrentContext;
  AFile := Ctx.FileName;
  ALine := Ctx.Line;
  Result := (AFile <> '') and (ALine > 0);
end;

procedure ReplaceLinesInFile(const AFile: string; AStartLine, AEndLine: Integer;
  const ANewText: string);
var
  Lines: TArray<string>;
  Buf: TStringList;
  I: Integer;
begin
  Lines := TDelphiFileEncoding.ReadLines(AFile);
  Buf := TStringList.Create;
  try
    for I := 0 to AStartLine - 2 do Buf.Add(Lines[I]);
    Buf.Add(ANewText);
    for I := AEndLine to High(Lines) do Buf.Add(Lines[I]);
    WriteLinesEnc(AFile, StringListToArray(Buf));
  finally
    Buf.Free;
  end;
end;

procedure InsertAtLineInFile(const AFile: string; AAtLine: Integer;
  const ANewText: string);
var
  Lines: TArray<string>;
  Buf: TStringList;
  I: Integer;
begin
  Lines := TDelphiFileEncoding.ReadLines(AFile);
  Buf := TStringList.Create;
  try
    for I := 0 to AAtLine - 2 do Buf.Add(Lines[I]);
    Buf.Add(ANewText);
    for I := AAtLine - 1 to High(Lines) do Buf.Add(Lines[I]);
    WriteLinesEnc(AFile, StringListToArray(Buf));
  finally
    Buf.Free;
  end;
end;

procedure AddAncestorToClassLines(var ALines: TArray<string>; ADeclLine: Integer;
  const AInterfaceName: string);
var
  Line, U, NewLine: string;
  ParenOpen, ParenClose, EqPos, ClassPos: Integer;
begin
  if (ADeclLine < 1) or (ADeclLine > Length(ALines)) then Exit;
  Line := ALines[ADeclLine - 1];
  U := UpperCase(Line);
  EqPos := Pos('=', Line);
  ClassPos := Pos('CLASS', U);
  if (EqPos = 0) or (ClassPos = 0) then Exit;

  ParenOpen := Pos('(', Copy(Line, ClassPos, MaxInt));
  if ParenOpen > 0 then
  begin
    var Abs_ := ClassPos + ParenOpen - 1;
    ParenClose := Pos(')', Copy(Line, Abs_, MaxInt));
    if ParenClose = 0 then Exit;
    var Body := Copy(Line, Abs_ + 1, ParenClose - 2);
    // Skip if interface already in list.
    if Pos(UpperCase(AInterfaceName), UpperCase(Body)) > 0 then Exit;
    var Insert := Body + ', ' + AInterfaceName;
    NewLine := Copy(Line, 1, Abs_) + Insert + Copy(Line, Abs_ + ParenClose - 1, MaxInt);
  end
  else
  begin
    // No ancestor list yet: "TFoo = class" -> "TFoo = class(TObject, IFoo)"
    var Tail := Trim(Copy(Line, ClassPos + 5, MaxInt));
    if (Tail = '') or (Tail[1] = ';') then
      NewLine := Copy(Line, 1, ClassPos + 4) + '(TObject, ' + AInterfaceName + ')'
        + Copy(Line, ClassPos + 5, MaxInt)
    else
      NewLine := Copy(Line, 1, ClassPos + 4) + '(TObject, ' + AInterfaceName + ')'
        + ' ' + Tail;
  end;

  ALines[ADeclLine - 1] := NewLine;
end;

procedure EnsureUsesContainsInLines(var ALines: TArray<string>;
  const AUnitName: string);
var
  I, Idx: Integer;
  L, U: string;
  InInterface: Boolean;
begin
  InInterface := False;
  Idx := -1;
  for I := 0 to High(ALines) do
  begin
    U := UpperCase(Trim(ALines[I]));
    if U = 'INTERFACE' then InInterface := True
    else if U = 'IMPLEMENTATION' then InInterface := False;
    if InInterface and StartsText('USES', U) then begin Idx := I; Break; end;
  end;

  if Idx < 0 then
  begin
    // Source unit has no interface-section uses clause. Inject one
    // right after the "interface" keyword line.
    for I := 0 to High(ALines) do
    begin
      U := UpperCase(Trim(ALines[I]));
      if U = 'INTERFACE' then
      begin
        // Insert a uses clause after this line.
        var NewLines: TArray<string>;
        SetLength(NewLines, Length(ALines) + 3);
        var J: Integer;
        for J := 0 to I do NewLines[J] := ALines[J];
        NewLines[I + 1] := '';
        NewLines[I + 2] := 'uses';
        NewLines[I + 3] := '  ' + AUnitName + ';';
        for J := I + 1 to High(ALines) do NewLines[J + 3] := ALines[J];
        ALines := NewLines;
        Exit;
      end;
    end;
    Exit;
  end;

  // Walk forward until ';' that ends uses, search for the unit name
  // and remember the line that carries the closing semicolon.
  var Acc := '';
  var EndIdx := Idx;
  for I := Idx to High(ALines) do
  begin
    Acc := Acc + ' ' + ALines[I];
    if Pos(';', ALines[I]) > 0 then begin EndIdx := I; Break; end;
  end;
  if Pos(UpperCase(AUnitName), UpperCase(Acc)) > 0 then Exit; // already present

  // Append at the END of the uses clause (Delphi convention). Find
  // the ';' on EndIdx and turn ", <newunit>;" in front of it.
  L := ALines[EndIdx];
  var SemPos := Pos(';', L);
  if SemPos = 0 then Exit;
  // Drop trailing spaces immediately before ';' so the insert is clean.
  var Before := Copy(L, 1, SemPos - 1);
  while (Before <> '') and (Before[Length(Before)] = ' ') do
    Before := Copy(Before, 1, Length(Before) - 1);
  // If the uses list spans a single line ("uses A, B;") put the new
  // unit on that same line. If it spans multiple lines (each unit on
  // its own line), put the new unit on its own new line with the same
  // indent that the previous unit on EndIdx used.
  if EndIdx = Idx then
    L := Before + ', ' + AUnitName + Copy(L, SemPos, MaxInt)
  else
  begin
    // Multi-line: turn EndIdx into "... <prev>," and add a new line
    // "  <newunit>;" with the same leading indentation as EndIdx.
    var Indent := '';
    var K: Integer := 1;
    while (K <= Length(L)) and (L[K] = ' ') do begin Indent := Indent + ' '; Inc(K); end;
    if Indent = '' then Indent := '  ';
    ALines[EndIdx] := Before + ',';
    // Splice a new line right after EndIdx.
    var NewLines: TArray<string>;
    SetLength(NewLines, Length(ALines) + 1);
    for K := 0 to EndIdx do NewLines[K] := ALines[K];
    NewLines[EndIdx + 1] := Indent + AUnitName + Copy(L, SemPos, MaxInt);
    for K := EndIdx + 1 to High(ALines) do NewLines[K + 1] := ALines[K];
    ALines := NewLines;
    Exit;
  end;
  ALines[EndIdx] := L;
end;


function FindLastPrivateSectionInClass(const ALines: TArray<string>;
  AClassDeclLine, AClassEndLine: Integer; out AInsertBeforeLine: Integer): Boolean;
// Looks for an existing 'private' (or 'strict private') section inside
// the class body. Returns the LAST such section (closest to 'end;') so
// new synth declarations get appended to the user's intended private
// area instead of growing a duplicate one at the bottom.
//
// On success, AInsertBeforeLine is the 1-based line index of the FIRST
// line that follows the private section's last member - i.e. the next
// visibility keyword ('protected'/'public'/'published'/'private') or
// the class's terminating 'end;'. The caller inserts new lines just
// before this index.
var
  I: Integer;
  T, U: string;
  LastPrivateLine: Integer;
begin
  Result := False;
  AInsertBeforeLine := -1;
  LastPrivateLine := -1;
  // 1-based class body: between (AClassDeclLine + 1) and (AClassEndLine - 1).
  for I := AClassDeclLine to AClassEndLine - 2 do
  begin
    if (I < 0) or (I >= Length(ALines)) then Continue;
    U := UpperCase(Trim(ALines[I]));
    // Strip trailing ';' if any (rare on visibility keywords, but safe).
    if (U <> '') and (U[Length(U)] = ';') then
      U := Trim(Copy(U, 1, Length(U) - 1));
    if (U = 'PRIVATE') or (U = 'STRICT PRIVATE') then
      LastPrivateLine := I;
  end;
  if LastPrivateLine < 0 then Exit;

  // Walk forward from the private keyword to find where the section
  // ends: either the next visibility keyword or the class end.
  for I := LastPrivateLine + 1 to AClassEndLine - 1 do
  begin
    if I >= Length(ALines) then Break;
    T := UpperCase(Trim(ALines[I]));
    if (T <> '') and (T[Length(T)] = ';') then
      T := Trim(Copy(T, 1, Length(T) - 1));
    if (T = 'PRIVATE') or (T = 'STRICT PRIVATE') or
       (T = 'PROTECTED') or (T = 'STRICT PROTECTED') or
       (T = 'PUBLIC') or (T = 'PUBLISHED') then
    begin
      AInsertBeforeLine := I + 1;
      Exit(True);
    end;
  end;
  // No following visibility keyword: insert right before the class end.
  AInsertBeforeLine := AClassEndLine;
  Result := True;
end;

procedure SynthesiseAllPropertiesOnLines(var ALines: TArray<string>;
  AClassDeclLine, AClassEndLine: Integer;
  const AMembers: TArray<TClassMember>);
// Splices all synthetic Get/Set declarations for AMembers into the
// class body. Prefers appending to an existing 'private' section so we
// do not introduce a duplicate one. Falls back to a fresh 'private'
// section just before the class 'end;' if no private section exists.
//
// We do NOT add a class property X - Pascal binds the interface's
// "property X read GetX write SetX" directly to those accessor methods,
// and a same-name class property would collide with an existing
// like-named field (e.g. a VCL component "Button1: TButton").
var
  Buf: TStringList;
  M: TClassMember;
  I, InsertBefore: Integer;
  Indent, Pn, T, Get, Setter: string;
  HasPrivate: Boolean;
  EmitLines: TList<string>;
begin
  if Length(AMembers) = 0 then Exit;
  if (AClassEndLine < 2) or (AClassEndLine > Length(ALines)) then Exit;

  Indent := '  ';

  // Build the lines we want to emit (just the method declarations -
  // the 'private' keyword is conditional on whether we're appending
  // to an existing section or opening a new one).
  EmitLines := TList<string>.Create;
  try
    for M in AMembers do
    begin
      Pn := M.ExposedName;
      T := M.TypeName;
      Get := 'Get' + Pn;
      Setter := 'Set' + Pn;
      EmitLines.Add(Indent + Indent + 'function ' + Get + ': ' + T + ';');
      if not (M.IsReadOnly and (M.Kind <> mkField)) then
        EmitLines.Add(Indent + Indent + 'procedure ' + Setter + '(const AValue: ' + T + ');');
    end;

    HasPrivate := FindLastPrivateSectionInClass(ALines,
      AClassDeclLine, AClassEndLine, InsertBefore);

    Buf := TStringList.Create;
    try
      if HasPrivate then
      begin
        for I := 0 to InsertBefore - 2 do Buf.Add(ALines[I]);
        for var SL in EmitLines do Buf.Add(SL);
        for I := InsertBefore - 1 to High(ALines) do Buf.Add(ALines[I]);
      end
      else
      begin
        // No existing private section - open a fresh one before 'end;'.
        for I := 0 to AClassEndLine - 2 do Buf.Add(ALines[I]);
        Buf.Add(Indent + 'private');
        for var SL in EmitLines do Buf.Add(SL);
        for I := AClassEndLine - 1 to High(ALines) do Buf.Add(ALines[I]);
      end;
      ALines := StringListToArray(Buf);
    finally
      Buf.Free;
    end;
  finally
    EmitLines.Free;
  end;
end;

function SynthesiseImplementations(const AClassName: string;
  const M: TClassMember; const AFieldName: string): string;
var
  Pn, T, Get, Setter, Backing: string;
  SB: TStringBuilder;
begin
  Pn := M.ExposedName;
  T := M.TypeName;
  Get := 'Get' + Pn;
  Setter := 'Set' + Pn;
  Backing := AFieldName;
  SB := TStringBuilder.Create;
  try
    SB.AppendLine;
    SB.Append('function ').Append(AClassName).Append('.').Append(Get)
      .Append(': ').Append(T).Append(';').AppendLine;
    SB.Append('begin').AppendLine;
    SB.Append('  Result := ').Append(Backing).Append(';').AppendLine;
    SB.Append('end;').AppendLine;
    if not M.IsReadOnly then
    begin
      SB.AppendLine;
      SB.Append('procedure ').Append(AClassName).Append('.').Append(Setter)
        .Append('(const AValue: ').Append(T).Append(');').AppendLine;
      SB.Append('begin').AppendLine;
      SB.Append('  ').Append(Backing).Append(' := AValue;').AppendLine;
      SB.Append('end;').AppendLine;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure AppendImplementationsInLines(var ALines: TArray<string>;
  const AClassName: string; const AMembers: TArray<TClassMember>);
var
  I, ImplLine: Integer;
  Buf: TStringList;
  S: string;
  M: TClassMember;
  Add: string;
begin
  ImplLine := -1;
  for I := High(ALines) downto 0 do
  begin
    S := UpperCase(Trim(ALines[I]));
    if (S = 'END.') then begin ImplLine := I + 1; Break; end;
  end;
  if ImplLine < 1 then Exit;

  Buf := TStringList.Create;
  try
    for I := 0 to ImplLine - 2 do Buf.Add(ALines[I]);
    for M in AMembers do
    begin
      if not M.Selected then Continue;
      Add := '';
      if M.Kind = mkField then
        Add := SynthesiseImplementations(AClassName, M, M.Name)
      else if (M.Kind = mkProperty) and M.NeedsSynthAccessors then
        Add := SynthesiseImplementations(AClassName, M, 'F' + M.ExposedName);
      if Add <> '' then
        for var L in Add.Split([sLineBreak]) do
          Buf.Add(L);
    end;
    for I := ImplLine - 1 to High(ALines) do Buf.Add(ALines[I]);
    ALines := StringListToArray(Buf);
  finally
    Buf.Free;
  end;
end;

procedure ApplyExtractNew(const AInfo: TExtractInterfaceInfo;
  out AStats: TResolveUsesStats; AStatusCallback: TProc<string>);
var
  UnitText, UnitName: string;
  M: TClassMember;
  NeedSynth: TArray<TClassMember>;
  Lines: TArray<string>;
begin
  AStats := Default(TResolveUsesStats);
  if TFile.Exists(AInfo.TargetFile) then
    if MessageDlg('Target file already exists. Overwrite?' + sLineBreak +
        AInfo.TargetFile, mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      Exit;

  // 1. Create the new interface unit on disk, register it with the
  //    active project and open it in the IDE so the user can see it.
  //    Seed the new unit's uses clause with the source unit's
  //    interface-uses so type references (TButton, TListView, ...)
  //    resolve out of the box.
  begin
    var InfoWithUses: TExtractInterfaceInfo := AInfo;
    // LSP-first: resolve only the units that the chosen interface
    // members actually need (via textDocument/definition on every
    // referenced type identifier). Falls back to the source unit's
    // full interface-uses when LSP delivers nothing.
    InfoWithUses.NewUnitUses := ResolveUsesViaLsp(AInfo.SourceFile, AInfo.Members,
      AStatusCallback, AStats);
    // Safety net: if LSP did not (or could not) resolve every type we
    // collected, top off the uses with the source unit's own
    // interface-uses. The list stays minimal when LSP resolved
    // everything, but does not leave the new unit uncompilable when
    // resolution was partial (e.g. because the source still has its
    // own uses-issue that prevents LSP from resolving from its
    // perspective).
    if (not AStats.LspAvailable) or
       (AStats.TypesResolved < AStats.TypesAttempted) then
      InfoWithUses.NewUnitUses := UnionUnitNames(
        InfoWithUses.NewUnitUses,
        ExtractInterfaceUses(AInfo.SourceFile));
    UnitText := TExtractInterfaceEngine.BuildNewUnitText(InfoWithUses);
  end;
  WriteTextEnc(AInfo.TargetFile, UnitText);
  AddFileToActiveProject(AInfo.TargetFile);
  OpenModuleInIDE(AInfo.TargetFile);

  UnitName := ChangeFileExt(ExtractFileName(AInfo.TargetFile), '');

  // 2. Update the source unit through ONE buffer. Reading from the
  //    editor buffer (if open) and pushing back via IOTAEditWriter
  //    means the change is undoable and shows up in the IDE without
  //    a manual reload. All transformations operate on the in-memory
  //    line array so there is no line-number drift between steps.
  Lines := ReadSourceLines(AInfo.SourceFile);

  NeedSynth := nil;
  for M in AInfo.Members do
    if M.Selected and
       ((M.Kind = mkField) or
        ((M.Kind = mkProperty) and M.NeedsSynthAccessors)) then
      NeedSynth := NeedSynth + [M];

  // Order matters: synth-decls go BEFORE the class-end line. Since the
  // class-decl line sits above the synth area, it stays valid. The
  // ancestor edit modifies that line in place. AppendImplementations
  // finds the trailing 'end.' on its own, so it tolerates the shift
  // caused by the synth splice.
  if Length(NeedSynth) > 0 then
  begin
    SynthesiseAllPropertiesOnLines(Lines, AInfo.ClassDeclLine, AInfo.ClassEndLine, NeedSynth);
    AppendImplementationsInLines(Lines, AInfo.ClassName, NeedSynth);
  end;
  AddAncestorToClassLines(Lines, AInfo.ClassDeclLine, AInfo.InterfaceName);
  EnsureUsesContainsInLines(Lines, UnitName);

  WriteSourceLines(AInfo.SourceFile, Lines);
end;

procedure ApplyAddToExisting(const AInfo: TExtractInterfaceInfo;
  out AStats: TResolveUsesStats; AStatusCallback: TProc<string>);

  /// <summary>Builds a per-target FilteredInfo from AInfo, where only
  ///  the members selected for THIS target are marked Selected, plus
  ///  applies clash filtering against the target's existing interface
  ///  declarations.</summary>
  function BuildFilteredInfoForTarget(const ATarget: TInterfaceTarget): TExtractInterfaceInfo;
  var
    K: Integer;
    Existing: TArray<string>;
  begin
    Result := AInfo;
    Result.InterfaceName := ATarget.InterfaceName;
    Result.ExistingFile := ATarget.FileName;
    Result.ExistingDeclLine := ATarget.DeclLine;
    Result.ExistingEndLine := ATarget.EndLine;
    // Take selections from the target only - ignore the .Selected
    // flags carried by AInfo.Members (which reflect the last UI view).
    for K := 0 to High(Result.Members) do
      Result.Members[K].Selected :=
        (K <= High(ATarget.MemberSelected)) and ATarget.MemberSelected[K];
    Existing := TExtractInterfaceEngine.ParseExistingInterfaceNames(
      ATarget.FileName, ATarget.DeclLine, ATarget.EndLine);
    for K := 0 to High(Result.Members) do
      if Result.Members[K].Selected and
         TExtractInterfaceEngine.ClashesWithExisting(Result.Members[K], Existing) then
        Result.Members[K].Selected := False;
  end;

  /// <summary>Splices the body of AFilteredInfo's interface text into
  ///  the target file just before its terminating 'end;', AND extends
  ///  the target file's uses clause with the type units AFilteredInfo
  ///  needs. Stats are accumulated into AStats across all targets.</summary>
  procedure SpliceIntoOneTarget(const AFilteredInfo: TExtractInterfaceInfo);
  var
    ExistingLines: TArray<string>;
    EndLineIdx, I: Integer;
    InterfaceText, Splice: string;
    PartialStats: TResolveUsesStats;
    Buf: TStringList;
  begin
    ExistingLines := ReadSourceLines(AFilteredInfo.ExistingFile);
    EndLineIdx := AFilteredInfo.ExistingEndLine - 1;
    if (EndLineIdx < 0) or (EndLineIdx >= Length(ExistingLines)) then Exit;

    // Type resolution for THIS target's selected members.
    var NeededUses: TArray<string> :=
      ResolveUsesViaLsp(AFilteredInfo.SourceFile, AFilteredInfo.Members,
        AStatusCallback, PartialStats);
    AStats.TypesAttempted := AStats.TypesAttempted + PartialStats.TypesAttempted;
    AStats.TypesResolved := AStats.TypesResolved + PartialStats.TypesResolved;
    if PartialStats.LspAvailable then AStats.LspAvailable := True;
    if (AStats.LspError = '') and (PartialStats.LspError <> '') then
      AStats.LspError := PartialStats.LspError;
    if (not PartialStats.LspAvailable) or
       (PartialStats.TypesResolved < PartialStats.TypesAttempted) then
      NeededUses := UnionUnitNames(NeededUses,
        ExtractInterfaceUses(AFilteredInfo.SourceFile));

    var ExistingUnitName := UpperCase(
      ChangeFileExt(ExtractFileName(AFilteredInfo.ExistingFile), ''));
    for var UN in NeededUses do
      if not SameText(UN, ExistingUnitName) then
        EnsureUsesContainsInLines(ExistingLines, UN);

    // Re-locate the interface's 'end;' line after the uses edits.
    begin
      var Depth: Integer := 1;
      var FoundEnd: Integer := -1;
      for I := AFilteredInfo.ExistingDeclLine to High(ExistingLines) do
      begin
        var U := UpperCase(StripLineCommentLocal(ExistingLines[I]));
        var P := 1;
        while P <= Length(U) do
        begin
          if U[P].IsLetter or (U[P] = '_') then
          begin
            var Q := P;
            while (Q <= Length(U)) and (U[Q].IsLetterOrDigit or (U[Q] = '_')) do Inc(Q);
            var W := Copy(U, P, Q - P);
            if (W = 'RECORD') then Inc(Depth)
            else if (W = 'END') then
            begin
              Dec(Depth);
              if Depth = 0 then begin FoundEnd := I + 1; Break; end;
            end;
            P := Q;
          end
          else
            Inc(P);
        end;
        if FoundEnd > 0 then Break;
      end;
      if FoundEnd > 0 then EndLineIdx := FoundEnd - 1;
    end;

    InterfaceText := TExtractInterfaceEngine.BuildInterfaceText(AFilteredInfo);
    var Body: TStringList := TStringList.Create;
    try
      Body.Text := InterfaceText;
      if Body.Count > 3 then
      begin
        Body.Delete(0); Body.Delete(0);
        Body.Delete(Body.Count - 1);
      end;
      Splice := Body.Text;
    finally
      Body.Free;
    end;

    Buf := TStringList.Create;
    try
      for I := 0 to EndLineIdx - 1 do Buf.Add(ExistingLines[I]);
      for var SL in Splice.Split([sLineBreak]) do
        if SL <> '' then Buf.Add(SL);
      for I := EndLineIdx to High(ExistingLines) do Buf.Add(ExistingLines[I]);
      WriteSourceLines(AFilteredInfo.ExistingFile, StringListToArray(Buf));
    finally
      Buf.Free;
    end;
  end;

var
  ExistingLines, SrcLines: TArray<string>;
  InterfaceText, Splice: string;
  EndLineIdx: Integer;
  M: TClassMember;
  NeedSynth: TArray<TClassMember>;
  Buf: TStringList;
  I: Integer;
begin
  AStats := Default(TResolveUsesStats);

  // Multi-target path: iterate every interface the user picked members
  // for. Each gets its own splice + uses edit; the class-side rewrite
  // happens ONCE afterwards on the union of all targets' synth-able
  // members so we don't synthesise duplicate Get/Set methods.
  if Length(AInfo.Targets) > 0 then
  begin
    // Union of class-side selections across all targets, then class-
    // side rewrite below.
    var UnionMembers: TArray<TClassMember> := AInfo.Members;
    for I := 0 to High(UnionMembers) do
      UnionMembers[I].Selected := False;

    for var Target in AInfo.Targets do
    begin
      var FilteredInfo := BuildFilteredInfoForTarget(Target);
      SpliceIntoOneTarget(FilteredInfo);
      for I := 0 to High(FilteredInfo.Members) do
        if FilteredInfo.Members[I].Selected then
          UnionMembers[I].Selected := True;
    end;

    // Class-side: one rewrite covering the union, plus an ancestor
    // entry and uses entry per target interface.
    NeedSynth := nil;
    for M in UnionMembers do
      if M.Selected and
         ((M.Kind = mkField) or
          ((M.Kind = mkProperty) and M.NeedsSynthAccessors)) then
        NeedSynth := NeedSynth + [M];
    SrcLines := ReadSourceLines(AInfo.SourceFile);
    if Length(NeedSynth) > 0 then
    begin
      SynthesiseAllPropertiesOnLines(SrcLines, AInfo.ClassDeclLine, AInfo.ClassEndLine, NeedSynth);
      AppendImplementationsInLines(SrcLines, AInfo.ClassName, NeedSynth);
    end;
    for var Target in AInfo.Targets do
    begin
      AddAncestorToClassLines(SrcLines, AInfo.ClassDeclLine, Target.InterfaceName);
      EnsureUsesContainsInLines(SrcLines,
        ChangeFileExt(ExtractFileName(Target.FileName), ''));
    end;
    WriteSourceLines(AInfo.SourceFile, SrcLines);
    Exit;
  end;

  // Single-target legacy path (Targets empty -> falls through to the
  // original ExistingFile/DeclLine/EndLine fields).
  if (AInfo.ExistingFile = '') or (AInfo.ExistingEndLine < 1) then
  begin
    ShowMessage('No target interface selected.');
    Exit;
  end;

  // Deselect any member whose would-emit names are already declared
  // in the target interface - otherwise we would splice duplicates and
  // the unit would not compile. We work on a local copy of AInfo so
  // the caller's selection state is left intact.
  var FilteredInfo: TExtractInterfaceInfo := AInfo;
  begin
    var Existing := TExtractInterfaceEngine.ParseExistingInterfaceNames(
      AInfo.ExistingFile, AInfo.ExistingDeclLine, AInfo.ExistingEndLine);
    for I := 0 to High(FilteredInfo.Members) do
      if FilteredInfo.Members[I].Selected and
         TExtractInterfaceEngine.ClashesWithExisting(
           FilteredInfo.Members[I], Existing) then
        FilteredInfo.Members[I].Selected := False;
  end;

  // 1. Splice the new declarations into the existing interface unit
  //    AND ensure all type identifiers used by the new members resolve
  //    via the existing unit's interface-uses clause. Both edits run
  //    on one in-memory buffer so they go through the editor in a
  //    single round-trip.
  ExistingLines := ReadSourceLines(AInfo.ExistingFile);
  EndLineIdx := AInfo.ExistingEndLine - 1;
  if (EndLineIdx < 0) or (EndLineIdx >= Length(ExistingLines)) then Exit;

  // 1a. Resolve types via LSP and add any missing units to the
  //     existing interface unit's uses clause. (Without this step the
  //     unit would not compile after the splice when a referenced type
  //     - e.g. TButton - was not previously used by the interface
  //     unit.) If LSP returns nothing we fall back to seeding from the
  //     source unit's interface-uses.
  var NeededUses: TArray<string> :=
    ResolveUsesViaLsp(AInfo.SourceFile, FilteredInfo.Members,
      AStatusCallback, AStats);
  // Same safety net as ApplyExtractNew: top off with source uses when
  // LSP did not resolve every type.
  if (not AStats.LspAvailable) or
     (AStats.TypesResolved < AStats.TypesAttempted) then
    NeededUses := UnionUnitNames(NeededUses,
      ExtractInterfaceUses(AInfo.SourceFile));
  var ExistingUnitName := UpperCase(
    ChangeFileExt(ExtractFileName(AInfo.ExistingFile), ''));
  for var UN in NeededUses do
    // Don't add the existing interface unit to its own uses.
    if not SameText(UN, ExistingUnitName) then
      EnsureUsesContainsInLines(ExistingLines, UN);

  // After the uses edit ExistingLines may have grown - re-locate the
  // interface's end-line by recomputing from AInfo.ExistingEndLine plus
  // the line-count delta introduced by the new uses entries.
  // Simpler and robust: re-find the 'end;' line by walking from
  // AInfo.ExistingDeclLine downward, tracking nesting.
  begin
    var Depth: Integer := 1;
    var FoundEnd: Integer := -1;
    for I := AInfo.ExistingDeclLine - 1 + 1 to High(ExistingLines) do
    begin
      var U := UpperCase(StripLineCommentLocal(ExistingLines[I]));
      var P := 1;
      while P <= Length(U) do
      begin
        if U[P].IsLetter or (U[P] = '_') then
        begin
          var Q := P;
          while (Q <= Length(U)) and (U[Q].IsLetterOrDigit or (U[Q] = '_')) do Inc(Q);
          var W := Copy(U, P, Q - P);
          if (W = 'RECORD') then Inc(Depth)
          else if (W = 'END') then
          begin
            Dec(Depth);
            if Depth = 0 then begin FoundEnd := I + 1; Break; end;
          end;
          P := Q;
        end
        else
          Inc(P);
      end;
      if FoundEnd > 0 then Break;
    end;
    if FoundEnd > 0 then EndLineIdx := FoundEnd - 1;
  end;

  // 1b. Build the body to splice (drop the synthetic "IFoo = interface"
  //     header lines and the trailing 'end;' - we only need the inner
  //     declarations).
  InterfaceText := TExtractInterfaceEngine.BuildInterfaceText(FilteredInfo);
  var Body: TStringList := TStringList.Create;
  try
    Body.Text := InterfaceText;
    if Body.Count > 3 then
    begin
      Body.Delete(0); Body.Delete(0);
      Body.Delete(Body.Count - 1);
    end;
    Splice := Body.Text;
  finally
    Body.Free;
  end;

  Buf := TStringList.Create;
  try
    for I := 0 to EndLineIdx - 1 do Buf.Add(ExistingLines[I]);
    for var SL in Splice.Split([sLineBreak]) do
      if SL <> '' then Buf.Add(SL);
    for I := EndLineIdx to High(ExistingLines) do Buf.Add(ExistingLines[I]);
    WriteSourceLines(AInfo.ExistingFile, StringListToArray(Buf));
  finally
    Buf.Free;
  end;

  // 2. Class-side: ancestor + synth + impls + uses, all via one
  //    in-memory buffer pushed back through the IDE. Use FilteredInfo
  //    so synthesised accessors are NOT emitted for members that
  //    already live in the target interface.
  NeedSynth := nil;
  for M in FilteredInfo.Members do
    if M.Selected and
       ((M.Kind = mkField) or
        ((M.Kind = mkProperty) and M.NeedsSynthAccessors)) then
      NeedSynth := NeedSynth + [M];

  SrcLines := ReadSourceLines(AInfo.SourceFile);
  if Length(NeedSynth) > 0 then
  begin
    SynthesiseAllPropertiesOnLines(SrcLines, AInfo.ClassDeclLine, AInfo.ClassEndLine, NeedSynth);
    AppendImplementationsInLines(SrcLines, AInfo.ClassName, NeedSynth);
  end;
  AddAncestorToClassLines(SrcLines, AInfo.ClassDeclLine, AInfo.InterfaceName);
  EnsureUsesContainsInLines(SrcLines,
    ChangeFileExt(ExtractFileName(AInfo.ExistingFile), ''));
  WriteSourceLines(AInfo.SourceFile, SrcLines);
end;

procedure RunWizard(AMode: TInterfaceMode);
var
  Src: string;
  CurLine: Integer;
  Lines: TArray<string>;
  Info, Result_: TExtractInterfaceInfo;
  Existing: TArray<TInterfaceDeclLocation>;
  M: TClassMember;
  ProjFiles: TArray<string>;
begin
  Editor.SaveAllFiles;
  if not GetEditorCursorLine(Src, CurLine) then
  begin
    ShowMessage('No editor file at cursor.');
    Exit;
  end;
  Lines := TDelphiFileEncoding.ReadLines(Src);
  if not TExtractInterfaceEngine.ParseClassAtLine(Lines, Src, CurLine, Info) then
  begin
    ShowMessage('No class declaration found around the cursor.');
    Exit;
  end;

  Info.Mode := AMode;

  // Default selection: all public + published methods/properties.
  for var I := 0 to High(Info.Members) do
  begin
    M := Info.Members[I];
    Info.Members[I].Selected :=
      (M.Visibility in [mvPublic, mvPublished]) and (M.Kind <> mkField);
  end;

  Existing := nil;
  if AMode = eimAddToExisting then
  begin
    ProjFiles := Editor.GetProjectSourceFiles;
    var AllProjectInterfaces := TProjectInterfaceScanner.ScanProject(ProjFiles);
    if Length(AllProjectInterfaces) = 0 then
    begin
      ShowMessage('No interface declarations found in the project.');
      Exit;
    end;

    // Filter: prefer interfaces the class already implements (its
    // ancestor list past the base class). Fallback: all project
    // interfaces, in case the class implements none yet.
    var BaseClass: string;
    var Implemented := TExtractInterfaceEngine.ParseClassAncestors(
      Lines[Info.ClassDeclLine - 1], BaseClass);
    if Length(Implemented) > 0 then
    begin
      var Filtered: TArray<TInterfaceDeclLocation>;
      for var Loc in AllProjectInterfaces do
        for var IfName in Implemented do
          if SameText(Loc.InterfaceName, IfName) then
            Filtered := Filtered + [Loc];
      if Length(Filtered) > 0 then
        Existing := Filtered
      else
        Existing := AllProjectInterfaces;
    end
    else
      Existing := AllProjectInterfaces;
  end;

  if not TExtractInterfaceDialog.Choose(Application.MainForm, AMode, Info,
    Existing, Result_) then Exit;

  // Visible status feedback during apply. The Application title gets
  // updated for each LSP step so the user sees "warming up..." /
  // "resolving TButton..." pass by even though there is no modal
  // status dialog.
  var OriginalTitle: string := Application.Title;
  var StatusCb: TProc<string> :=
    procedure(S: string)
    begin
      Application.Title := 'Refactoring Light: ' + S;
      Application.ProcessMessages;
    end;

  var Stats: TResolveUsesStats;
  try
    Screen.Cursor := crHourGlass;
    try
      if AMode = eimExtractNew then
        ApplyExtractNew(Result_, Stats, StatusCb)
      else
        ApplyAddToExisting(Result_, Stats, StatusCb);
    finally
      Screen.Cursor := crDefault;
      Application.Title := OriginalTitle;
    end;

    // Build a diagnostic message so the user can verify what LSP did.
    var Diag: string := sLineBreak + sLineBreak + 'LSP type resolution:' + sLineBreak;
    if not Stats.LspAvailable then
      Diag := Diag + '  - LSP client could not be acquired (' +
        IfThen(Stats.LspError = '', 'no .delphilsp.json?', Stats.LspError) + ').' + sLineBreak +
        '  - Fell back to the source unit''s full interface-uses.'
    else
      Diag := Diag + Format(
        '  - %d distinct type(s) collected from selected members.' + sLineBreak +
        '  - %d resolved via textDocument/definition.' + sLineBreak +
        '  - %d not found in the project / DCU search path.',
        [Stats.TypesAttempted, Stats.TypesResolved,
         Stats.TypesAttempted - Stats.TypesResolved]);

    ShowMessage('Interface ' +
      IfThen(AMode = eimExtractNew, 'extracted to ' + Result_.TargetFile,
                                     'extended in ' + Result_.ExistingFile)
      + '.' + Diag);
  except
    on E: Exception do
    begin
      Application.Title := OriginalTitle;
      Screen.Cursor := crDefault;
      ShowMessage('Failed: ' + E.ClassName + ': ' + E.Message);
    end;
  end;
end;

procedure ExtractInterfaceFromClass;
begin
  RunWizard(eimExtractNew);
end;

procedure AddToExistingInterface;
begin
  RunWizard(eimAddToExisting);
end;

{ ---------- IInterface support flow ----------

  Adds IInterface compatibility directly to the class at the cursor:
  declares IInterface in the ancestor list, adds FRefCount, overrides
  NewInstance / AfterConstruction (to mirror TInterfacedObject's
  initial-refcount-1 trick so an interface cast during the constructor
  cannot cause premature self-Destroy), and implements QueryInterface /
  _AddRef / _Release. Once applied, the class can be held purely as an
  IInterface and the instance frees itself when the last reference is
  released.

  Two emission modes depending on the base class:
    * TObject / TPersistent / unknown:  fresh implementations (no
      `override` on the IInterface methods).
    * TComponent or descendant:         `override` (TComponent declares
      _AddRef / _Release / QueryInterface as virtual stdcall).
}

function InjectStmtBeforeMethodBodyEnd(var ALines: TArray<string>;
  const AClassName, AMethodName, AStmtToInject: string): Boolean;
// Locates the implementation of AClassName.AMethodName and splices
// AStmtToInject into its body just before the matching 'end;'. Tracks
// nested begin/try/case/record blocks. Returns False when the impl
// cannot be located, so the caller can warn the user instead of
// silently dropping the cleanup.
var
  I, J, K, BeginLine, EndLine, Depth: Integer;
  Trimmed, Upper, Indent, HeaderUpper: string;
  Buf: TStringList;
begin
  Result := False;
  HeaderUpper := UpperCase(AClassName) + '.' + UpperCase(AMethodName);

  for I := 0 to High(ALines) do
  begin
    Trimmed := Trim(StripLineCommentLocal(ALines[I]));
    Upper := UpperCase(Trimmed);
    if not (StartsText('PROCEDURE ', Upper) or
            StartsText('FUNCTION ', Upper) or
            StartsText('CLASS PROCEDURE ', Upper) or
            StartsText('CLASS FUNCTION ', Upper)) then Continue;
    if Pos(HeaderUpper, Upper) = 0 then Continue;

    // Find 'begin' that opens this routine.
    BeginLine := -1;
    for J := I + 1 to High(ALines) do
    begin
      Upper := UpperCase(Trim(StripLineCommentLocal(ALines[J])));
      if (Upper = 'BEGIN') or StartsText('BEGIN ', Upper) then
      begin
        BeginLine := J; Break;
      end;
    end;
    if BeginLine < 0 then Exit;

    // Track depth to the matching closing 'end;'.
    Depth := 1;
    EndLine := -1;
    for J := BeginLine + 1 to High(ALines) do
    begin
      Trimmed := Trim(StripLineCommentLocal(ALines[J]));
      Upper := UpperCase(Trimmed);
      while (Upper <> '') and (Upper[Length(Upper)] = ';') do
        Upper := Trim(Copy(Upper, 1, Length(Upper) - 1));

      if (Upper = 'BEGIN') or StartsText('BEGIN ', Upper) or
         (Upper = 'TRY') or StartsText('TRY ', Upper) or
         (Upper = 'CASE') or StartsText('CASE ', Upper) or
         (Upper = 'RECORD') or StartsText('RECORD ', Upper) then
        Inc(Depth)
      else if (Upper = 'END') or StartsText('END ', Upper) then
      begin
        Dec(Depth);
        if Depth = 0 then begin EndLine := J; Break; end;
      end;
    end;
    if EndLine < 0 then Exit;

    // Pick indent from the closing 'end;' line.
    Indent := '';
    for K := 1 to Length(ALines[EndLine]) do
      if ALines[EndLine][K] = ' ' then Indent := Indent + ' '
      else Break;
    if Indent = '' then Indent := '  ';
    // Method bodies are typically '  ' indented; the body content is
    // one level deeper.
    Indent := Indent + '  ';

    Buf := TStringList.Create;
    try
      for K := 0 to EndLine - 1 do Buf.Add(ALines[K]);
      Buf.Add(Indent + AStmtToInject);
      for K := EndLine to High(ALines) do Buf.Add(ALines[K]);
      ALines := StringListToArray(Buf);
    finally
      Buf.Free;
    end;
    Result := True;
    Exit;
  end;
end;

procedure DelegateInterfaceImplementation;
var
  Src, BaseClass: string;
  CurLine, InsertBefore, ClassEndAfter, I: Integer;
  Lines: TArray<string>;
  Info: TExtractInterfaceInfo;
  DeclBlock, ImplBlock: string;
  BaseIsTObjectLike, HasNewInst, HasAfterCtor: Boolean;
  OverrideKW, BaseDesc, ExtraNote: string;
  InjectNewInstOK, InjectAfterCtorOK: Boolean;
begin
  Editor.SaveAllFiles;
  if not GetEditorCursorLine(Src, CurLine) then
  begin
    ShowMessage('No editor file at cursor.'); Exit;
  end;
  Lines := ReadSourceLines(Src);
  if not TExtractInterfaceEngine.ParseClassAtLine(Lines, Src, CurLine, Info) then
  begin
    ShowMessage('No class declaration found around the cursor.'); Exit;
  end;

  // Skip if the class already lists IInterface / IUnknown.
  var Ancestors := TExtractInterfaceEngine.ParseClassAncestors(
    Lines[Info.ClassDeclLine - 1], BaseClass);
  for var A in Ancestors do
    if SameText(A, 'IInterface') or SameText(A, 'IUnknown') then
    begin
      ShowMessage(Format(
        'Class %s already lists %s in its ancestor list - nothing to do.',
        [Info.ClassName, A]));
      Exit;
    end;

  // Heuristic on base class. TObject / TPersistent have no virtual
  // IInterface methods to override; everything else we treat as a
  // TComponent descendant. Users with an exotic base can adjust the
  // generated `override` keywords by hand.
  BaseIsTObjectLike := (BaseClass = '') or
    SameText(BaseClass, 'TObject') or SameText(BaseClass, 'TPersistent');
  if BaseIsTObjectLike then OverrideKW := ''
  else OverrideKW := 'override; ';

  // Detect whether the class already declares NewInstance and/or
  // AfterConstruction. We must NOT re-declare them (would trip E2007),
  // and we must NOT re-emit fresh implementations - instead the
  // existing bodies get the relevant refcount statement spliced in.
  HasNewInst := TExtractInterfaceEngine.ClassHasMethodDecl(
    Lines, Info.ClassDeclLine, Info.ClassEndLine, 'NewInstance');
  HasAfterCtor := TExtractInterfaceEngine.ClassHasMethodDecl(
    Lines, Info.ClassDeclLine, Info.ClassEndLine, 'AfterConstruction');

  DeclBlock :=
    '    FRefCount: Integer;' + sLineBreak +
    '  protected' + sLineBreak +
    '    function QueryInterface(const IID: TGUID; out Obj): HResult; ' + OverrideKW + 'stdcall;' + sLineBreak +
    '    function _AddRef: Integer; ' + OverrideKW + 'stdcall;' + sLineBreak +
    '    function _Release: Integer; ' + OverrideKW + 'stdcall;';
  if (not HasAfterCtor) or (not HasNewInst) then
    DeclBlock := DeclBlock + sLineBreak + '  public';
  if not HasAfterCtor then
    DeclBlock := DeclBlock + sLineBreak +
      '    procedure AfterConstruction; override;';
  if not HasNewInst then
    DeclBlock := DeclBlock + sLineBreak +
      '    class function NewInstance: TObject; override;';

  Lines := ReadSourceLines(Src);

  if FindLastPrivateSectionInClass(Lines, Info.ClassDeclLine,
    Info.ClassEndLine, InsertBefore) then
  begin
    var Buf: TStringList := TStringList.Create;
    try
      for I := 0 to InsertBefore - 2 do Buf.Add(Lines[I]);
      for var BL in DeclBlock.Split([sLineBreak]) do Buf.Add(BL);
      for I := InsertBefore - 1 to High(Lines) do Buf.Add(Lines[I]);
      Lines := StringListToArray(Buf);
    finally
      Buf.Free;
    end;
  end
  else
  begin
    var Buf: TStringList := TStringList.Create;
    try
      for I := 0 to Info.ClassEndLine - 2 do Buf.Add(Lines[I]);
      Buf.Add('  private');
      for var BL in DeclBlock.Split([sLineBreak]) do Buf.Add(BL);
      for I := Info.ClassEndLine - 1 to High(Lines) do Buf.Add(Lines[I]);
      Lines := StringListToArray(Buf);
    finally
      Buf.Free;
    end;
  end;

  // Ancestor list: explicit IInterface. For TObject-like base this is
  // required (TObject does not implement IInterface). For TComponent-
  // like base it is redundant but harmless and makes intent explicit.
  AddAncestorToClassLines(Lines, Info.ClassDeclLine, 'IInterface');

  // 1. Splice into existing NewInstance / AfterConstruction bodies
  //    (if the user has them). Do this BEFORE appending our fresh
  //    impls so the find-by-name search in InjectStmtBeforeMethodBodyEnd
  //    sees only the user's pre-existing impl, not our about-to-be-
  //    appended one.
  InjectNewInstOK := True;
  InjectAfterCtorOK := True;
  if HasNewInst then
    InjectNewInstOK := InjectStmtBeforeMethodBodyEnd(Lines,
      Info.ClassName, 'NewInstance',
      Info.ClassName + '(Result).FRefCount := 1;');
  if HasAfterCtor then
    InjectAfterCtorOK := InjectStmtBeforeMethodBodyEnd(Lines,
      Info.ClassName, 'AfterConstruction', 'AtomicDecrement(FRefCount);');

  // 2. Build the impl block, omitting NewInstance / AfterConstruction
  //    when the user already has them.
  ImplBlock := '';
  if not HasNewInst then
    ImplBlock := ImplBlock + sLineBreak +
      'class function ' + Info.ClassName + '.NewInstance: TObject;' + sLineBreak +
      'begin' + sLineBreak +
      '  // Start at refcount 1 so an interface cast during the' + sLineBreak +
      '  // constructor cannot drop the count to 0 and self-Destroy.' + sLineBreak +
      '  // AfterConstruction below decrements it back to 0.' + sLineBreak +
      '  Result := inherited NewInstance;' + sLineBreak +
      '  ' + Info.ClassName + '(Result).FRefCount := 1;' + sLineBreak +
      'end;' + sLineBreak;
  if not HasAfterCtor then
    ImplBlock := ImplBlock + sLineBreak +
      'procedure ' + Info.ClassName + '.AfterConstruction;' + sLineBreak +
      'begin' + sLineBreak +
      '  inherited;' + sLineBreak +
      '  AtomicDecrement(FRefCount);' + sLineBreak +
      'end;' + sLineBreak;
  ImplBlock := ImplBlock + sLineBreak +
    'function ' + Info.ClassName + '.QueryInterface(const IID: TGUID; out Obj): HResult;' + sLineBreak +
    'begin' + sLineBreak +
    '  if GetInterface(IID, Obj) then' + sLineBreak +
    '    Result := S_OK' + sLineBreak +
    '  else' + sLineBreak +
    '    Result := E_NOINTERFACE;' + sLineBreak +
    'end;' + sLineBreak + sLineBreak +
    'function ' + Info.ClassName + '._AddRef: Integer;' + sLineBreak +
    'begin' + sLineBreak +
    '  Result := AtomicIncrement(FRefCount);' + sLineBreak +
    'end;' + sLineBreak + sLineBreak +
    'function ' + Info.ClassName + '._Release: Integer;' + sLineBreak +
    'begin' + sLineBreak +
    '  Result := AtomicDecrement(FRefCount);' + sLineBreak +
    '  if Result = 0 then' + sLineBreak +
    '    Destroy;' + sLineBreak +
    'end;';

  ClassEndAfter := -1;
  for I := High(Lines) downto 0 do
    if UpperCase(Trim(Lines[I])) = 'END.' then
    begin
      ClassEndAfter := I; Break;
    end;
  if ClassEndAfter >= 0 then
  begin
    var Buf: TStringList := TStringList.Create;
    try
      for I := 0 to ClassEndAfter - 1 do Buf.Add(Lines[I]);
      // Normalise spacing: strip any blank lines that already trail
      // the prior content so we never end up with three blank lines
      // before our first method (user's two + our leading newline).
      while (Buf.Count > 0) and (Trim(Buf[Buf.Count - 1]) = '') do
        Buf.Delete(Buf.Count - 1);
      // One blank line before our first inserted method.
      Buf.Add('');
      // Strip the leading newline ImplBlock starts with (it was there
      // so the old splice produced *some* gap; the new splice manages
      // the gap explicitly).
      var BlockText: string := ImplBlock;
      while StartsText(sLineBreak, BlockText) do
        BlockText := Copy(BlockText, Length(sLineBreak) + 1, MaxInt);
      for var BL in BlockText.Split([sLineBreak]) do Buf.Add(BL);
      // One blank line after the last inserted method, before 'end.'.
      Buf.Add('');
      for I := ClassEndAfter to High(Lines) do Buf.Add(Lines[I]);
      Lines := StringListToArray(Buf);
    finally
      Buf.Free;
    end;
  end;

  WriteSourceLines(Src, Lines);

  if BaseIsTObjectLike then
    BaseDesc := 'TObject-like base (' +
      IfThen(BaseClass = '', 'TObject', BaseClass) +
      ') - fresh IInterface implementation'
  else
    BaseDesc := 'TComponent-style base (' + BaseClass +
      ') - overrides the inherited virtual IInterface methods';

  ExtraNote := '';
  if HasNewInst then
    if InjectNewInstOK then
      ExtraNote := ExtraNote + sLineBreak +
        '- Spliced ' + Info.ClassName + '(Result).FRefCount := 1; into the existing NewInstance.'
    else
      ExtraNote := ExtraNote + sLineBreak +
        '- WARNING: could not locate the existing NewInstance body. ' +
        'Add  ' + Info.ClassName + '(Result).FRefCount := 1;  manually.';
  if HasAfterCtor then
    if InjectAfterCtorOK then
      ExtraNote := ExtraNote + sLineBreak +
        '- Spliced AtomicDecrement(FRefCount); into the existing AfterConstruction.'
    else
      ExtraNote := ExtraNote + sLineBreak +
        '- WARNING: could not locate the existing AfterConstruction body. ' +
        'Add  AtomicDecrement(FRefCount);  manually.';

  ShowMessage(Format(
    'Class %s now implements IInterface (%s).' + sLineBreak +
    'The instance frees itself when the last interface reference is dropped.%s',
    [Info.ClassName, BaseDesc, ExtraNote]));
end;

end.
