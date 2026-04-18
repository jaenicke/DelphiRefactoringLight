(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ImplementationFinder;

{
  Gemeinsamer Finder fuer Interface-/Virtual-Method-Implementierungen.

  Nutzt LSP's textDocument/implementation. Das Ergebnis kommt als
  TFindReferenceItems (wie bei Find References), damit es direkt vom
  existierenden Ergebnis-Dialog angezeigt werden kann und auch vom
  Rename-Wizard wiederverwendet wird.
}

interface

uses
  System.SysUtils,
  Lsp.Client,
  Expert.FindReferencesDialog;

type
  /// <summary>Callback fuer Fortschritt beim Datei-Scan.
  ///  ACurrent/ATotal sind 1-basiert.</summary>
  TImplementationScanProgress = reference to procedure(ACurrent, ATotal: Integer);

  TImplementationFinder = class
  public
    /// <summary>
    /// Durchsucht alle uebergebenen Projektdateien nach Zeilen, die wie
    /// eine Klassen-Methoden-Implementierung aussehen:
    ///   procedure TFoo.Method;
    ///   function   TFoo.Method(...): ...;
    ///   constructor TFoo.Create;
    ///   destructor  TFoo.Destroy;
    ///   class procedure TFoo.Method;
    /// usw.
    ///
    /// Vorgehen analog zum Rename-Wizard:
    /// 1. Text-Suche mit Wort-Grenzen ueber alle Dateien.
    /// 2. Filter auf Methoden-Implementierungs-Syntax (Keyword am Zeilenanfang,
    ///    Punkt direkt vor dem Bezeichner).
    /// 3. Wenn AExpectedOwnerType gesetzt ist: Klasse der Impl-Zeile
    ///    (vor dem Punkt) extrahieren und pruefen ob sie direkt oder via
    ///    Vererbung diesen Typ implementiert. Dadurch werden Klassen mit
    ///    zufaellig gleichem Methodennamen, aber anderem Interface, gefiltert.
    /// </summary>
    class function FindByProjectScan(const AProjectFiles: TArray<string>; const AIdentifier: string;
      const AExpectedOwnerType: string; AProgress: TImplementationScanProgress = nil): TFindReferenceItems;

    /// <summary>Scant rueckwaerts von AStartLine nach einer Zeile die wie
    ///  'TName = interface' oder 'TName = class' aussieht und liefert
    ///  den Typ-Namen. Damit findet man den Container-Typ in dem eine
    ///  Methode deklariert ist. Liefert '' wenn nichts gefunden.</summary>
    class function FindContainingType(const AFilePath: string; AStartLine: Integer): string;

    /// <summary>Prueft rekursiv ob die Klasse AClassName den Typ
    ///  ATargetType implementiert oder von ihm erbt. Parst hierzu die
    ///  Parent-Listen der Klassen-Deklarationen im Projekt
    ///  (z.B. 'TFoo = class(TBase, IFoo, IBar)').</summary>
    class function ClassImplementsType(const AProjectFiles: TArray<string>; const AClassName, ATargetType: string): Boolean;

    /// <summary>
    /// LSP-basiert (textDocument/implementation an einer Position).
    /// Funktioniert nur wenn der LSP-Server alle relevanten Dateien bereits
    /// indexiert hat. Als Fallback / Ergaenzung zu FindByProjectScan gedacht.
    /// </summary>
    class function FindAt(AClient: TLspClient; const AFilePath: string;
      ALine, ACol: Integer; const AIdentifier: string): TFindReferenceItems;

    /// <summary>
    /// LSP-basiert (robustere Variante von FindAt):
    /// probiert erst die Cursor-Position, dann (falls ohne Treffer) die
    /// Deklarationsstelle.
    /// </summary>
    class function FindForIdentifierAt(AClient: TLspClient; const AFilePath: string; ALine, ACol: Integer;
      const AIdentifier: string): TFindReferenceItems;
  end;

implementation

uses
  System.Classes, System.Generics.Collections, System.StrUtils, Winapi.Windows, Lsp.Protocol, Lsp.Uri, Delphi.FileEncoding;

{ File-private helpers for parsing Pascal class/interface headers
  and class-method implementation lines. Grouped into a class to keep
  the unit free of global routines. }

type
  TImplFinderHelper = class
  public
    class function IsClassMethodImplLine(const ALine: string; AIdentifierCol: Integer): Boolean; static;
    class function ExtractClassNameBeforeIdentifier(const ALine: string; AIdentifierCol: Integer): string; static;
    class function ExtractTypeHeaderName(const ALine: string): string; static;
    class function ExtractClassParents(const ALines: TArray<string>; AStartLine: Integer): TArray<string>; static;
    class function FindClassParents(const AProjectFiles: TArray<string>; const AClassName: string): TArray<string>; static;
    class function ClassImplementsTypeRec(const AProjectFiles: TArray<string>; const AClassName, ATargetType: string;
      AVisited: TStringList): Boolean; static;
  end;

/// <summary>True if the line is a class-method implementation header
///  (e.g. 'procedure TFoo.Bar;') and AIdentifierCol points at the method
///  name. Conditions:
///    * Line (left-trimmed) starts with procedure / function /
///      constructor / destructor / operator (optionally 'class '
///      prefixed).
///    * A dot sits directly before the identifier (ClassName.Identifier).</summary>
class function TImplFinderHelper.IsClassMethodImplLine(const ALine: string; AIdentifierCol: Integer): Boolean;
var
  Trimmed: string;
begin
  Result := False;

  // Dot immediately before the identifier? (ClassName.Identifier)
  if AIdentifierCol <= 1 then Exit;
  if ALine[AIdentifierCol - 1] <> '.' then Exit;

  // Line must start with a method keyword.
  Trimmed := LowerCase(TrimLeft(ALine));
  Result :=
    StartsStr('procedure ',    Trimmed) or
    StartsStr('function ',     Trimmed) or
    StartsStr('constructor ',  Trimmed) or
    StartsStr('destructor ',   Trimmed) or
    StartsStr('operator ',     Trimmed) or
    StartsStr('class procedure ',   Trimmed) or
    StartsStr('class function ',    Trimmed) or
    StartsStr('class constructor ', Trimmed) or
    StartsStr('class destructor ',  Trimmed) or
    StartsStr('class operator ',    Trimmed);
end;

/// <summary>Extracts the class name that sits directly before the dot
///  preceding AIdentifierCol. Example: in 'procedure TFoo.Bar;' with
///  AIdentifierCol pointing at 'B' in 'Bar' -> returns 'TFoo'.</summary>
class function TImplFinderHelper.ExtractClassNameBeforeIdentifier(const ALine: string; AIdentifierCol: Integer): string;
var
  EndPos, StartPos: Integer;
begin
  Result := '';
  if AIdentifierCol <= 2 then Exit;
  if ALine[AIdentifierCol - 1] <> '.' then Exit;

  EndPos := AIdentifierCol - 2; // the character immediately before the dot
  StartPos := EndPos;
  while (StartPos >= 1) and
        CharInSet(ALine[StartPos], ['A'..'Z','a'..'z','0'..'9','_']) do
    Dec(StartPos);
  Inc(StartPos);

  if StartPos <= EndPos then
    Result := Copy(ALine, StartPos, EndPos - StartPos + 1);
end;

/// <summary>Parses a line like 'TName = class(...)' or
///  'TName = interface(...)' and returns the type name. Returns '' if
///  the line does not contain such a header.</summary>
class function TImplFinderHelper.ExtractTypeHeaderName(const ALine: string): string;
var
  Trimmed, Lower, Before, After: string;
  EqPos, GenPos: Integer;
begin
  Result := '';
  Trimmed := Trim(ALine);
  if Trimmed = '' then Exit;
  Lower := LowerCase(Trimmed);

  EqPos := Pos('=', Lower);
  if EqPos = 0 then Exit;

  Before := TrimRight(Copy(Trimmed, 1, EqPos - 1));
  After  := TrimLeft(Copy(Lower, EqPos + 1));

  // 'packed class'
  if StartsStr('packed ', After) then
    After := TrimLeft(Copy(After, Length('packed ') + 1));

  // Must start with 'class' or 'interface'.
  if not (StartsStr('class', After) or StartsStr('interface', After)) then Exit;

  // 'class of X' is a metaclass type, not a class declaration.
  if StartsStr('class of ', After) then Exit;

  // 'class procedure/function ...' is a method header, not a class
  // declaration (can't really appear here without '=', but defensive).
  if StartsStr('class procedure', After) or
     StartsStr('class function', After) or
     StartsStr('class constructor', After) or
     StartsStr('class destructor', After) or
     StartsStr('class operator', After) then Exit;

  // Strip generics from the type name: 'TList<T>' -> 'TList'.
  GenPos := Pos('<', Before);
  if GenPos > 0 then
    Before := TrimRight(Copy(Before, 1, GenPos - 1));

  if Before = '' then Exit;
  if not CharInSet(Before[1], ['A'..'Z','a'..'z','_']) then Exit;

  Result := Before;
end;

/// <summary>Extracts the parent list (paren contents) from a class
///  declaration. Combines follow-up lines if necessary until the closing
///  ')' is found. Generics are stripped from each entry. Returns an
///  empty array if there are no parens (e.g. "TFoo = class"), or if
///  the line is not a class/interface header.</summary>
class function TImplFinderHelper.ExtractClassParents(const ALines: TArray<string>; AStartLine: Integer): TArray<string>;
var
  Combined, Inside, Part, P: string;
  OpenParen, CloseParen, I, GenPos: Integer;
  Parts: TArray<string>;
begin
  Result := nil;
  if (AStartLine < 0) or (AStartLine > High(ALines)) then Exit;

  Combined := ALines[AStartLine];
  OpenParen := Pos('(', Combined);
  if OpenParen = 0 then Exit; // e.g. "TFoo = class" (implicitly inherits TObject)

  CloseParen := 0;
  I := AStartLine;
  // Append follow-up lines until ')' is found - cap at 20 lines.
  while CloseParen = 0 do
  begin
    CloseParen := PosEx(')', Combined, OpenParen + 1);
    if CloseParen = 0 then
    begin
      Inc(I);
      if (I > High(ALines)) or (I - AStartLine > 20) then Exit;
      Combined := Combined + ' ' + ALines[I];
    end;
  end;

  Inside := Copy(Combined, OpenParen + 1, CloseParen - OpenParen - 1);
  Parts := Inside.Split([','], TStringSplitOptions.ExcludeEmpty);

  SetLength(Result, 0);
  for Part in Parts do
  begin
    P := Trim(Part);
    // Strip generics: 'IFoo<T>' -> 'IFoo'.
    GenPos := Pos('<', P);
    if GenPos > 0 then
      P := TrimRight(Copy(P, 1, GenPos - 1));
    if P <> '' then
    begin
      SetLength(Result, Length(Result) + 1);
      Result[High(Result)] := P;
    end;
  end;
end;

/// <summary>Searches AProjectFiles for the declaration of AClassName
///  and returns its parent list. Returns nil if not found.</summary>
class function TImplFinderHelper.FindClassParents(const AProjectFiles: TArray<string>; const AClassName: string): TArray<string>;
var
  Lines: TArray<string>;
  I: Integer;
  RawContent: string;
begin
  Result := nil;
  if AClassName = '' then Exit;

  for var F in AProjectFiles do
  begin
    // Coarse file-level check first.
    try
      RawContent := TDelphiFileEncoding.ReadAll(F);
      if Pos(UpperCase(AClassName), UpperCase(RawContent)) = 0 then Continue;
      Lines := TDelphiFileEncoding.ReadLines(F);
    except
      Continue;
    end;

    for I := 0 to High(Lines) do
    begin
      if SameText(ExtractTypeHeaderName(Lines[I]), AClassName) then
      begin
        Result := ExtractClassParents(Lines, I);
        Exit;
      end;
    end;
  end;
end;

/// <summary>Recursively checks whether AClassName implements ATargetType
///  either directly or via an inheritance/implementation chain.</summary>
class function TImplFinderHelper.ClassImplementsTypeRec(const AProjectFiles: TArray<string>;
  const AClassName, ATargetType: string; AVisited: TStringList): Boolean;
var
  Parents: TArray<string>;
begin
  Result := False;
  if AClassName = '' then Exit;
  if SameText(AClassName, ATargetType) then Exit(True);

  // Cycle guard
  if AVisited.IndexOf(UpperCase(AClassName)) >= 0 then Exit;
  AVisited.Add(UpperCase(AClassName));

  Parents := FindClassParents(AProjectFiles, AClassName);
  for var Parent in Parents do
  begin
    if SameText(Parent, ATargetType) then Exit(True);
    if ClassImplementsTypeRec(AProjectFiles, Parent, ATargetType, AVisited) then
      Exit(True);
  end;
end;

{ TImplementationFinder }

class function TImplementationFinder.FindContainingType(const AFilePath: string; AStartLine: Integer): string;
var
  Lines: TArray<string>;
  I: Integer;
  TypeName: string;
begin
  Result := '';
  try
    Lines := ReadDelphiFileLines(AFilePath);
  except
    Exit;
  end;

  if AStartLine < 0 then AStartLine := 0;
  if AStartLine > High(Lines) then AStartLine := High(Lines);

  for I := AStartLine downto 0 do
  begin
    TypeName := TImplFinderHelper.ExtractTypeHeaderName(Lines[I]);
    if TypeName <> '' then
      Exit(TypeName);
  end;
end;

class function TImplementationFinder.ClassImplementsType(const AProjectFiles: TArray<string>;
  const AClassName, ATargetType: string): Boolean;
var
  Visited: TStringList;
begin
  Result := False;
  if (AClassName = '') or (ATargetType = '') then Exit;

  Visited := TStringList.Create;
  try
    Result := TImplFinderHelper.ClassImplementsTypeRec(AProjectFiles, AClassName, ATargetType, Visited);
  finally
    Visited.Free;
  end;
end;

class function TImplementationFinder.FindByProjectScan(const AProjectFiles: TArray<string>; const AIdentifier: string;
  const AExpectedOwnerType: string; AProgress: TImplementationScanProgress): TFindReferenceItems;
var
  ResultList: TList<TFindReferenceItem>;
  Item: TFindReferenceItem;
  RawContent, LineStr: string;
  Lines: TArray<string>;
  UpperId: string;
  LineIdx, SearchPos, FoundPos, AfterPos: Integer;
  BeforeOk, AfterOk, Keep: Boolean;
  ClassCache: TDictionary<string, Boolean>;
  VerifyOwner: Boolean;
  ClassName: string;
begin
  Result := nil;
  if AIdentifier = '' then Exit;
  if System.Length(AProjectFiles) = 0 then Exit;

  UpperId := UpperCase(AIdentifier);
  VerifyOwner := AExpectedOwnerType <> '';

  ResultList := TList<TFindReferenceItem>.Create;
  ClassCache := TDictionary<string, Boolean>.Create;
  try
    for var FileIdx := 0 to High(AProjectFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(FileIdx + 1, System.Length(AProjectFiles));

      var F := AProjectFiles[FileIdx];

      // Erst den Rohinhalt pruefen - spart Arbeit bei Dateien ohne Treffer
      try
        RawContent := ReadDelphiFile(F);
        if Pos(UpperId, UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(F);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        LineStr := Lines[LineIdx];
        SearchPos := 1;
        while SearchPos <= System.Length(LineStr) do
        begin
          FoundPos := Pos(UpperId, UpperCase(Copy(LineStr, SearchPos)));
          if FoundPos = 0 then Break;
          FoundPos := SearchPos + FoundPos - 1;

          // Wort-Grenzen pruefen
          BeforeOk := (FoundPos = 1) or not CharInSet(LineStr[FoundPos - 1], ['A'..'Z','a'..'z','0'..'9','_']);
          AfterPos := FoundPos + System.Length(AIdentifier);
          AfterOk := (AfterPos > System.Length(LineStr)) or not CharInSet(LineStr[AfterPos], ['A'..'Z','a'..'z','0'..'9','_']);

          if BeforeOk and AfterOk and TImplFinderHelper.IsClassMethodImplLine(LineStr, FoundPos) then
          begin
            Keep := True;

            if VerifyOwner then
            begin
              // Klassenname vor dem Punkt extrahieren und pruefen ob
              // diese Klasse den erwarteten Owner-Typ implementiert.
              ClassName := TImplFinderHelper.ExtractClassNameBeforeIdentifier(LineStr, FoundPos);
              if ClassName = '' then
                Keep := False
              else if not ClassCache.TryGetValue(UpperCase(ClassName), Keep) then
              begin
                Keep := SameText(ClassName, AExpectedOwnerType) or ClassImplementsType(AProjectFiles, ClassName, AExpectedOwnerType);
                ClassCache.Add(UpperCase(ClassName), Keep);
              end;
            end;

            if Keep then
            begin
              Item.FilePath := F;
              Item.Line := LineIdx;
              Item.Col := FoundPos - 1;
              Item.Length := System.Length(AIdentifier);
              Item.Preview := Trim(LineStr);
              ResultList.Add(Item);
            end;
          end;

          SearchPos := FoundPos + System.Length(AIdentifier);
        end;
      end;
    end;

    Result := ResultList.ToArray;
  finally
    ClassCache.Free;
    ResultList.Free;
  end;
end;

class function TImplementationFinder.FindAt(AClient: TLspClient; const AFilePath: string; ALine, ACol: Integer;
  const AIdentifier: string): TFindReferenceItems;
var
  ImplLocs: TArray<TLspLocation>;
  ResultList: TList<TFindReferenceItem>;
  LineCache: TDictionary<string, TArray<string>>;
  Item: TFindReferenceItem;
  UpperId: string;
begin
  Result := nil;
  if AIdentifier = '' then Exit;

  ResultList := TList<TFindReferenceItem>.Create;
  LineCache := TDictionary<string, TArray<string>>.Create;
  try
    try
      AClient.RefreshDocument(AFilePath);
      Sleep(300);
      ImplLocs := AClient.GotoImplementation(AFilePath, ALine, ACol);
    except
      Exit(nil);
    end;

    UpperId := UpperCase(AIdentifier);
    for var Loc in ImplLocs do
    begin
      var FilePath := TLspUri.FileUriToPath(Loc.Uri);
      var ImplLine := Loc.Range.Start.Line;

      // Gleiche Zeile wie Anfrage ueberspringen (das waere das Interface selbst)
      if SameText(ExpandFileName(FilePath), ExpandFileName(AFilePath)) and (ImplLine = ALine) then
        Continue;

      // Zeilen mit Cache lesen (vermeidet mehrfaches Lesen gleicher Dateien)
      var Lines: TArray<string>;
      if not LineCache.TryGetValue(FilePath, Lines) then
      begin
        try
          Lines := ReadDelphiFileLines(FilePath);
          LineCache.Add(FilePath, Lines);
        except
          Continue;
        end;
      end;

      if (ImplLine < 0) or (ImplLine >= System.Length(Lines)) then
        Continue;

      var LineStr := Lines[ImplLine];

      // LSP zeigt teilweise auf den Beginn der Signatur, nicht den Bezeichner.
      // Wir suchen den Bezeichner innerhalb der Zeile.
      var FoundPos := System.Pos(UpperId, UpperCase(LineStr));
      if FoundPos = 0 then Continue;

      Item.FilePath := FilePath;
      Item.Line := ImplLine;
      Item.Col := FoundPos - 1;
      Item.Length := System.Length(AIdentifier);
      Item.Preview := Trim(LineStr);
      ResultList.Add(Item);
    end;

    Result := ResultList.ToArray;
  finally
    LineCache.Free;
    ResultList.Free;
  end;
end;

class function TImplementationFinder.FindForIdentifierAt(AClient: TLspClient; const AFilePath: string; ALine, ACol: Integer;
  const AIdentifier: string): TFindReferenceItems;
var
  Defs: TArray<TLspLocation>;
  DefPath: string;
  DefLine, DefCol: Integer;
begin
  // 1. Versuch: Cursor-Position
  Result := FindAt(AClient, AFilePath, ALine, ACol, AIdentifier);
  if System.Length(Result) > 0 then Exit;

  // 2. Versuch: Deklarationsstelle
  try
    Defs := AClient.GotoDefinition(AFilePath, ALine, ACol);
  except
    Exit;
  end;

  if System.Length(Defs) = 0 then Exit;

  DefPath := TLspUri.FileUriToPath(Defs[0].Uri);
  DefLine := Defs[0].Range.Start.Line;
  DefCol := Defs[0].Range.Start.Character;

  // Nicht nochmal probieren wenn Deklaration = Cursor
  if SameText(ExpandFileName(DefPath), ExpandFileName(AFilePath)) and
     (DefLine = ALine) and (DefCol = ACol) then
    Exit;

  Result := FindAt(AClient, DefPath, DefLine, DefCol, AIdentifier);
end;

end.
