(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SignatureCheck;

{
  Collects all declarations and the implementation of a Delphi method
  via the LSP server's textDocument/documentSymbol response and compares
  their signatures.

  For a method "TFoo.Bar" the following entries may exist in a project:

    * One declaration inside a class definition   (role = srClassDecl)
    * One declaration inside an interface type    (role = srInterfaceDecl)
    * The implementation in the implementation
      section of a unit                            (role = srImplementation)

  All three need to have the same parameter list and return type.
  This unit extracts them from the documentSymbol tree and produces a
  normalized form that can be compared.
}

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.Generics.Collections,
  Lsp.Client;

type
  TSignatureRole = (srUnknown, srInterfaceDecl, srClassDecl, srImplementation);

  TSignatureEntry = record
    FilePath: string;
    Line: Integer;        // 0-based
    Col: Integer;         // 0-based
    Role: TSignatureRole;
    Container: string;    // containing class / interface name (may be empty)
    Name: string;         // method name only (e.g. "Bar")
    RawSignature: string; // as reported by the LSP (name + detail)
    Normalized: string;   // whitespace / case normalized for comparing
  end;

  TSignatureEntries = TArray<TSignatureEntry>;

  TSignatureChecker = class
  private
    class function IsMethodKind(AKind: Integer): Boolean; static;
    class function IsTypeKind(AKind: Integer): Boolean; static;
    class function ExtractMethodName(const AFullName: string): string; static;
    class procedure WalkSymbols(const ASymbols: TJSONArray;
      const AFilePath, AContainer: string; const ATarget: string;
      AList: TList<TSignatureEntry>); static;
    class function LooksLikeInterfaceName(const AName: string): Boolean; static;
    class function ReadSignatureFromSource(const AFilePath: string;
      ALine: Integer; const AMethodName: string): string; static;
  public
    /// <summary>Collects signature entries for AMethodName from AFilePath.
    ///  If ARole is srImplementation this is the implementation file; if
    ///  it is a declaration the implementation will be resolved via
    ///  GotoDefinition and queried as well.</summary>
    class function Collect(AClient: TLspClient;
      const ACurrentFile, AMethodName: string): TSignatureEntries;

    /// <summary>Normalizes whitespace and case for comparison. Keeps the
    ///  parameter list order intact.</summary>
    class function Normalize(const ASignature: string): string;

    /// <summary>True if all entries share the same normalized signature.</summary>
    class function AllEqual(const AEntries: TSignatureEntries): Boolean;

    class function RoleToString(ARole: TSignatureRole): string;
  end;

implementation

uses
  System.StrUtils, System.Character, System.Math, Lsp.Uri, Delphi.FileEncoding;

const
  // Subset of LSP SymbolKind that we care about
  skInterface = 11;
  skClass     = 5;
  skRecord    = 23;
  skMethod    = 6;
  skFunction  = 12;
  skConstructor = 9;
  skProperty  = 7;

class function TSignatureChecker.IsMethodKind(AKind: Integer): Boolean;
begin
  Result := (AKind = skMethod) or (AKind = skFunction) or (AKind = skConstructor);
end;

class function TSignatureChecker.IsTypeKind(AKind: Integer): Boolean;
begin
  Result := (AKind = skClass) or (AKind = skInterface) or (AKind = skRecord);
end;

class function TSignatureChecker.LooksLikeInterfaceName(const AName: string): Boolean;
begin
  Result := (Length(AName) >= 2) and (UpCase(AName[1]) = 'I') and
            CharInSet(AName[2], ['A'..'Z']);
end;

class function TSignatureChecker.ExtractMethodName(const AFullName: string): string;
var
  DotPos, ParenPos, I: Integer;
  S: string;
begin
  // Can look like "TFoo.Bar(A: Integer): Boolean" or just "Bar"
  S := AFullName;
  ParenPos := Pos('(', S);
  if ParenPos > 0 then
    S := Copy(S, 1, ParenPos - 1);
  // Strip a trailing ": ReturnType" if present (rare without parens, but safe)
  I := Pos(':', S);
  if I > 0 then
    S := Copy(S, 1, I - 1);
  DotPos := LastDelimiter('.', S);
  if DotPos > 0 then
    S := Copy(S, DotPos + 1, MaxInt);
  Result := Trim(S);
end;

class function TSignatureChecker.ReadSignatureFromSource(const AFilePath: string;
  ALine: Integer; const AMethodName: string): string;
// Reads the actual signature text from disk. We cannot trust the LSP's
// "name" field of documentSymbol because DelphiLSP serves it from its
// symbol cache, which lags behind editor edits.
//
// Strategy: starting at ALine (the line containing the method name),
// scan back until we find the procedure/function/constructor/destructor
// keyword (or the start of the line if it's a class member without a
// keyword). Then concatenate forward across lines until we hit ';' at
// paren-depth 0.
const
  CKeywords: array[0..3] of string = ('PROCEDURE', 'FUNCTION', 'CONSTRUCTOR', 'DESTRUCTOR');
var
  Lines: TArray<string>;
  L: string;
  StartLine: Integer;
  StartCol: Integer;
  SB: TStringBuilder;
  I, J: Integer;
  KW: string;
  KPos, BestPos: Integer;
  Depth: Integer;
  CommentPos: Integer;
  Done: Boolean;
begin
  Result := '';
  try
    Lines := ReadDelphiFileLines(AFilePath);
  except
    Exit;
  end;
  if (ALine < 0) or (ALine >= Length(Lines)) then Exit;

  // Find the keyword on ALine or earlier
  StartLine := ALine;
  StartCol := 1;
  BestPos := 0;
  for I := ALine downto Max(0, ALine - 5) do
  begin
    L := UpperCase(Lines[I]);
    BestPos := 0;
    for KW in CKeywords do
    begin
      KPos := Pos(KW, L);
      // make sure it's word-bounded
      while KPos > 0 do
      begin
        var BeforeOk := (KPos = 1) or
          not CharInSet(L[KPos - 1], ['A'..'Z','0'..'9','_']);
        var AfterPos := KPos + Length(KW);
        var AfterOk := (AfterPos > Length(L)) or
          not CharInSet(L[AfterPos], ['A'..'Z','0'..'9','_']);
        if BeforeOk and AfterOk then
        begin
          if (BestPos = 0) or (KPos < BestPos) then
            BestPos := KPos;
          Break;
        end;
        var Next := Pos(KW, Copy(L, KPos + 1));
        if Next = 0 then KPos := 0 else KPos := KPos + Next;
      end;
    end;
    if BestPos > 0 then
    begin
      StartLine := I;
      StartCol := BestPos;
      Break;
    end;
  end;

  if BestPos = 0 then
  begin
    // No keyword found - probably a property; just use ALine from col 1.
    StartLine := ALine;
    StartCol := 1;
  end;

  SB := TStringBuilder.Create;
  try
    Depth := 0;
    Done := False;
    for I := StartLine to Min(StartLine + 30, High(Lines)) do
    begin
      L := Lines[I];
      // Strip line comments
      CommentPos := Pos('//', L);
      if CommentPos > 0 then L := Copy(L, 1, CommentPos - 1);

      var FromCol: Integer;
      if I = StartLine then FromCol := StartCol else FromCol := 1;

      for J := FromCol to Length(L) do
      begin
        if L[J] = '(' then Inc(Depth)
        else if L[J] = ')' then Dec(Depth);
        SB.Append(L[J]);
        if (L[J] = ';') and (Depth <= 0) then
        begin
          Done := True;
          Break;
        end;
      end;
      if Done then Break;
      SB.Append(' ');
    end;
    Result := Trim(SB.ToString);
  finally
    SB.Free;
  end;

  // Sanity check: result must contain the method name. If not, the line
  // we read does not match - fall back to original LSP value.
  if (AMethodName <> '') and (Pos(UpperCase(AMethodName), UpperCase(Result)) = 0) then
    Result := '';
end;

class procedure TSignatureChecker.WalkSymbols(const ASymbols: TJSONArray;
  const AFilePath, AContainer: string; const ATarget: string;
  AList: TList<TSignatureEntry>);
var
  Sym: TJSONValue;
  Obj: TJSONObject;
  Name, Detail, ChildContainer: string;
  Kind: Integer;
  Entry: TSignatureEntry;
  Children: TJSONArray;
  Range, Start: TJSONObject;
  MethodName: string;
begin
  if ASymbols = nil then Exit;
  for Sym in ASymbols do
  begin
    if not (Sym is TJSONObject) then Continue;
    Obj := TJSONObject(Sym);

    Name := '';
    Detail := '';
    Kind := 0;
    if Obj.GetValue('name') <> nil then Name := Obj.GetValue('name').Value;
    if Obj.GetValue('detail') <> nil then Detail := Obj.GetValue('detail').Value;
    if Obj.GetValue('kind') is TJSONNumber then
      Kind := TJSONNumber(Obj.GetValue('kind')).AsInt;

    if IsMethodKind(Kind) then
    begin
      MethodName := ExtractMethodName(Name);
      if SameText(MethodName, ATarget) then
      begin
        Entry := Default(TSignatureEntry);
        Entry.FilePath := AFilePath;
        Entry.Name := MethodName;
        Entry.Container := AContainer;
        Entry.RawSignature := Trim(Name); // LSP fallback (may be stale)

        // Position
        if Obj.GetValue('selectionRange') is TJSONObject then
          Range := TJSONObject(Obj.GetValue('selectionRange'))
        else if Obj.GetValue('range') is TJSONObject then
          Range := TJSONObject(Obj.GetValue('range'))
        else
          Range := nil;
        if (Range <> nil) and (Range.GetValue('start') is TJSONObject) then
        begin
          Start := TJSONObject(Range.GetValue('start'));
          if Start.GetValue('line') is TJSONNumber then
            Entry.Line := TJSONNumber(Start.GetValue('line')).AsInt;
          if Start.GetValue('character') is TJSONNumber then
            Entry.Col := TJSONNumber(Start.GetValue('character')).AsInt;
        end;

        // Role is derived by caller; here we tag by context hints.
        if AContainer = '' then
          Entry.Role := srImplementation
        else if LooksLikeInterfaceName(AContainer) then
          Entry.Role := srInterfaceDecl
        else
          Entry.Role := srClassDecl;

        // Override the signature with the actual source text. The LSP's
        // documentSymbol "name" comes from a cached symbol DB and lags
        // behind editor edits, so we cannot rely on it for comparison.
        var SrcSig := ReadSignatureFromSource(AFilePath, Entry.Line, MethodName);
        if SrcSig <> '' then
          Entry.RawSignature := SrcSig;

        Entry.Normalized := Normalize(Entry.RawSignature);
        AList.Add(Entry);
      end;
    end;

    // Recurse into children - classes / interfaces contain methods,
    // but methods themselves may also have nested symbols we want to skip.
    Children := nil;
    if Obj.GetValue('children') is TJSONArray then
      Children := TJSONArray(Obj.GetValue('children'));
    if Children <> nil then
    begin
      if IsTypeKind(Kind) then
        ChildContainer := Name
      else
        ChildContainer := AContainer;
      WalkSymbols(Children, AFilePath, ChildContainer, ATarget, AList);
    end;
  end;
end;

class function TSignatureChecker.Collect(AClient: TLspClient;
  const ACurrentFile, AMethodName: string): TSignatureEntries;
var
  List: TList<TSignatureEntry>;
  Arr: TJSONArray;
  SeenFiles: TStringList;

  procedure Query(const AFilePath: string);
  begin
    if SeenFiles.IndexOf(LowerCase(AFilePath)) >= 0 then Exit;
    SeenFiles.Add(LowerCase(AFilePath));
    try
      AClient.RefreshDocument(AFilePath);
    except
      // ignore - may already be open
    end;
    Arr := nil;
    try
      Arr := AClient.GetDocumentSymbols(AFilePath);
      if Arr <> nil then
        WalkSymbols(Arr, AFilePath, '', AMethodName, List);
    finally
      Arr.Free;
    end;
  end;

var
  I: Integer;
  ImplFile: string;
  InitialCount: Integer;
begin
  Result := nil;
  List := TList<TSignatureEntry>.Create;
  SeenFiles := TStringList.Create;
  try
    Query(ACurrentFile);

    InitialCount := List.Count;

    // For every entry in the current file, try to resolve the
    // counterpart via GotoDefinition and query that file too.
    for I := 0 to InitialCount - 1 do
    begin
      try
        var Defs := AClient.GotoDefinition(List[I].FilePath, List[I].Line, List[I].Col);
        for var D in Defs do
        begin
          ImplFile := TLspUri.FileUriToPath(D.Uri);
          if ImplFile <> '' then
            Query(ImplFile);
        end;
      except
        // ignore - LSP may not resolve every entry
      end;
    end;

    Result := List.ToArray;
  finally
    List.Free;
    SeenFiles.Free;
  end;
end;

class function TSignatureChecker.Normalize(const ASignature: string): string;
var
  I: Integer;
  C: Char;
  SB: TStringBuilder;
  PrevSpace: Boolean;
begin
  SB := TStringBuilder.Create;
  try
    PrevSpace := True; // swallow leading whitespace
    for I := 1 to Length(ASignature) do
    begin
      C := ASignature[I];
      if C.IsWhiteSpace then
      begin
        if not PrevSpace then
        begin
          SB.Append(' ');
          PrevSpace := True;
        end;
      end
      else if CharInSet(C, [';', ',', ':', '(', ')', '[', ']', '=']) then
      begin
        // Drop whitespace around structural tokens
        if (SB.Length > 0) and (SB.Chars[SB.Length - 1] = ' ') then
          SB.Length := SB.Length - 1;
        SB.Append(C);
        PrevSpace := True;
      end
      else
      begin
        SB.Append(C.ToLower);
        PrevSpace := False;
      end;
    end;
    // Drop trailing semicolon / space
    while (SB.Length > 0) and ((SB.Chars[SB.Length - 1] = ' ') or
          (SB.Chars[SB.Length - 1] = ';')) do
      SB.Length := SB.Length - 1;
    Result := SB.ToString;
  finally
    SB.Free;
  end;

  // Strip a leading "TClass." prefix so declarations and implementations
  // compare as equal regardless of qualification.
  var Dot := Pos('.', Result);
  var ParenPos := Pos('(', Result);
  if (Dot > 0) and ((ParenPos = 0) or (Dot < ParenPos)) then
    Result := Copy(Result, Dot + 1, MaxInt);
end;

class function TSignatureChecker.AllEqual(const AEntries: TSignatureEntries): Boolean;
var
  I: Integer;
begin
  Result := True;
  if Length(AEntries) < 2 then Exit;
  for I := 1 to High(AEntries) do
    if AEntries[I].Normalized <> AEntries[0].Normalized then
      Exit(False);
end;

class function TSignatureChecker.RoleToString(ARole: TSignatureRole): string;
begin
  case ARole of
    srInterfaceDecl: Result := 'Interface decl.';
    srClassDecl:     Result := 'Class decl.';
    srImplementation: Result := 'Implementation';
  else
    Result := 'Unknown';
  end;
end;

end.
