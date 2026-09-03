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
    /// <summary>Parses a property declaration line for APropName and
    ///  returns its read/write accessors ('' when the clause is missing).
    ///  Handles array properties and the trailing specifiers
    ///  (index/default/stored/nodefault/implements). False when ALine is
    ///  not a declaration of that property.</summary>
    class function ParsePropertyAccessors(const ALine, APropName: string;
      out AGetter, ASetter: string): Boolean; static;

    /// <summary>Implementations for a PROPERTY: a property has no
    ///  implementation line of its own, so the result is every
    ///  'property APropName ...' declaration in the project plus the
    ///  implementations of its read/write ACCESSOR methods.</summary>
    class function FindPropertyImplementations(const AProjectFiles: TArray<string>;
      const APropName: string; AProgress: TImplementationScanProgress): TFindReferenceItems;

    /// <summary>Implementations of a TYPE: every class in the project
    ///  that implements the interface ATypeName or descends from the class
    ///  ATypeName - the answer when the cursor sits on the type NAME
    ///  ('ITest = interface') instead of on one of its methods. Result
    ///  rows point at the class declarations.</summary>
    class function FindTypeImplementations(const AProjectFiles: TArray<string>;
      const ATypeName: string; AProgress: TImplementationScanProgress): TFindReferenceItems;

    class function FindByProjectScan(const AProjectFiles: TArray<string>; const AIdentifier: string;
      const AExpectedOwnerType: string; AProgress: TImplementationScanProgress = nil): TFindReferenceItems;

    /// <summary>Scant rueckwaerts von AStartLine nach einer Zeile die wie
    ///  'TName = interface' oder 'TName = class' aussieht und liefert
    ///  den Typ-Namen. Damit findet man den Container-Typ in dem eine
    ///  Methode deklariert ist. Liefert '' wenn nichts gefunden.</summary>
    class function FindContainingType(const AFilePath: string; AStartLine: Integer): string;

    /// <summary>Same, but on already-read lines - pure and testable.
    ///  IMPORTANT over a naive backward scan: it verifies that AStartLine
    ///  really is INSIDE the type it finds. A line in the implementation
    ///  section, inside a routine body, or after the closing 'end;' of the
    ///  preceding type yields '' instead of the nearest header above it -
    ///  a wrong owner type filters away EVERY implementation later on.
    ///  When AStartLine is itself a method implementation header
    ///  ('procedure TFoo.Bar;'), its qualifier is the answer.</summary>
    class function FindContainingTypeInLines(const ALines: TArray<string>;
      AStartLine: Integer): string; static;

    /// <summary>Owner type of a method IMPLEMENTATION header:
    ///  'procedure TFoo.Bar;' -> 'TFoo', 'function TFoo&lt;T&gt;.Get: T' ->
    ///  'TFoo', 'procedure TOuter.TInner.Baz;' -> 'TInner' (the qualifier
    ///  directly before the method name, which is what the scan compares
    ///  against). '' when the line is not such a header.</summary>
    class function OwnerTypeFromImplLine(const ALine: string): string; static;

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

class function TImplementationFinder.OwnerTypeFromImplLine(
  const ALine: string): string;
var
  L, U, Rest, Name: string;
  P, I, Depth: Integer;
  Parts: TArray<string>;
begin
  Result := '';
  L := Trim(ALine);
  P := Pos('//', L);
  if P > 0 then L := TrimRight(Copy(L, 1, P - 1));
  if L = '' then Exit;
  U := UpperCase(L);
  if StartsStr('CLASS ', U) then
  begin
    L := TrimLeft(Copy(L, Length('CLASS ') + 1, MaxInt));
    U := UpperCase(L);
  end;
  if not (StartsStr('PROCEDURE ', U) or StartsStr('FUNCTION ', U)
    or StartsStr('CONSTRUCTOR ', U) or StartsStr('DESTRUCTOR ', U)
    or StartsStr('OPERATOR ', U)) then Exit;

  P := Pos(' ', L);
  Rest := TrimLeft(Copy(L, P + 1, MaxInt));

  // Qualified name up to '(' / ':' / ';'; generic arguments are skipped
  // so 'TFoo<T>.Get' still splits at the right dot.
  Name := '';
  Depth := 0;
  I := 1;
  while I <= Length(Rest) do
  begin
    if Rest[I] = '<' then
      Inc(Depth)
    else if Rest[I] = '>' then
      Dec(Depth)
    else if Depth = 0 then
    begin
      if CharInSet(Rest[I], ['(', ':', ';', ' ', #9, '=']) then Break;
      Name := Name + Rest[I];
    end;
    Inc(I);
  end;

  Parts := Name.Split(['.']);
  if Length(Parts) >= 2 then
    Result := Trim(Parts[High(Parts) - 1]);
end;

class function TImplementationFinder.FindContainingTypeInLines(
  const ALines: TArray<string>; AStartLine: Integer): string;
var
  I: Integer;
  L, U, TypeName: string;

  function Clean(const S: string): string;
  var
    P: Integer;
  begin
    Result := Trim(S);
    if StartsStr('//', Result) then Exit('');
    P := Pos('//', Result);
    if P > 0 then Result := TrimRight(Copy(Result, 1, P - 1));
  end;

begin
  Result := '';
  if Length(ALines) = 0 then Exit;
  if AStartLine < 0 then AStartLine := 0;
  if AStartLine > High(ALines) then AStartLine := High(ALines);

  // The line itself may already BE the implementation header - that is
  // where DelphiLSP likes to land when asked for a method's definition.
  Result := OwnerTypeFromImplLine(ALines[AStartLine]);
  if Result <> '' then Exit;

  for I := AStartLine downto 0 do
  begin
    L := Clean(ALines[I]);
    if L = '' then Continue;
    U := UpperCase(L);

    if I < AStartLine then
    begin
      // Everything below is a boundary the search must not cross. Being
      // conservative here is deliberate: an empty result means "search
      // unfiltered", a WRONG type means "find nothing at all".
      // Reaching a method implementation header means the start line
      // sits INSIDE that method - its qualifier is the containing type.
      TypeName := OwnerTypeFromImplLine(L);
      if TypeName <> '' then Exit(TypeName);
      if (U = 'IMPLEMENTATION') or (U = 'END;') or (U = 'END') then
        Exit('');
    end;

    TypeName := TImplFinderHelper.ExtractTypeHeaderName(L);
    if TypeName <> '' then
    begin
      // 'TFoo = class;' / 'TFoo = class(TBase);' is a FORWARD declaration
      // - it has no body, so nothing can be inside it.
      if EndsStr(';', L) then Exit('');
      Exit(TypeName);
    end;
  end;
end;

class function TImplementationFinder.FindContainingType(const AFilePath: string; AStartLine: Integer): string;
var
  Lines: TArray<string>;
begin
  Result := '';
  try
    Lines := ReadDelphiFileLines(AFilePath);
  except
    Exit;
  end;
  Result := FindContainingTypeInLines(Lines, AStartLine);
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

class function TImplementationFinder.ParsePropertyAccessors(const ALine,
  APropName: string; out AGetter, ASetter: string): Boolean;
var
  L, U, Name: string;
  P, I: Integer;

  // Identifier following the keyword AKey (word-bounded), '' when absent.
  function AccessorAfter(const AKey: string): string;
  var
    Q, Start: Integer;
  begin
    Result := '';
    Q := Pos(' ' + AKey + ' ', ' ' + U + ' ');
    if Q = 0 then Exit;
    // Q is 1-based in the padded string -> position of the keyword in U.
    Start := Q + Length(AKey) + 1;
    while (Start <= Length(L)) and CharInSet(L[Start], [' ', #9]) do Inc(Start);
    Q := Start;
    while (Q <= Length(L)) and CharInSet(L[Q],
      ['A'..'Z', 'a'..'z', '0'..'9', '_']) do Inc(Q);
    Result := Copy(L, Start, Q - Start);
  end;

begin
  Result := False;
  AGetter := '';
  ASetter := '';
  L := Trim(ALine);
  P := Pos('//', L);
  if P > 0 then L := TrimRight(Copy(L, 1, P - 1));
  if L = '' then Exit;
  U := UpperCase(L);
  if not (U.StartsWith('PROPERTY ') or U.StartsWith('CLASS PROPERTY ')) then Exit;

  // Name = first identifier after the keyword, up to ':' / '[' / space.
  P := Pos('PROPERTY ', U) + Length('PROPERTY ');
  while (P <= Length(L)) and CharInSet(L[P], [' ', #9]) do Inc(P);
  I := P;
  while (I <= Length(L)) and CharInSet(L[I],
    ['A'..'Z', 'a'..'z', '0'..'9', '_']) do Inc(I);
  Name := Copy(L, P, I - P);
  if not SameText(Name, APropName) then Exit;

  AGetter := AccessorAfter('READ');
  ASetter := AccessorAfter('WRITE');
  Result := True;
end;

class function TImplementationFinder.FindPropertyImplementations(
  const AProjectFiles: TArray<string>; const APropName: string;
  AProgress: TImplementationScanProgress): TFindReferenceItems;
var
  ResultList: TList<TFindReferenceItem>;
  Accessors: TStringList;
  Item: TFindReferenceItem;
  RawContent, LineStr, Getter, Setter: string;
  Lines: TArray<string>;
  UpperId: string;
  LineIdx, Col: Integer;

  function AlreadyListed(const AFile: string; ALine: Integer): Boolean;
  var
    K: Integer;
  begin
    Result := False;
    for K := 0 to ResultList.Count - 1 do
      if (ResultList[K].Line = ALine)
        and SameText(ResultList[K].FilePath, AFile) then
        Exit(True);
  end;

begin
  Result := nil;
  if (APropName = '') or (System.Length(AProjectFiles) = 0) then Exit;
  UpperId := UpperCase(APropName);

  ResultList := TList<TFindReferenceItem>.Create;
  Accessors := TStringList.Create;
  try
    Accessors.Duplicates := dupIgnore;
    Accessors.Sorted := True;

    // Pass 1: every declaration of that property + its accessors.
    for var FileIdx := 0 to High(AProjectFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(FileIdx + 1, System.Length(AProjectFiles));
      try
        RawContent := ReadDelphiFile(AProjectFiles[FileIdx]);
        if Pos(UpperId, UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(AProjectFiles[FileIdx]);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        LineStr := Lines[LineIdx];
        if not ParsePropertyAccessors(LineStr, APropName, Getter, Setter) then
          Continue;
        Col := Pos(UpperId, UpperCase(LineStr)) - 1;
        if Col < 0 then Col := 0;
        Item.FilePath := AProjectFiles[FileIdx];
        Item.Line := LineIdx;
        Item.Col := Col;
        Item.Length := System.Length(APropName);
        Item.Preview := Trim(LineStr);
        ResultList.Add(Item);
        if Getter <> '' then Accessors.Add(Getter);
        if Setter <> '' then Accessors.Add(Setter);
      end;
    end;

    // Pass 2: the accessor METHODS' implementations (a field accessor
    // simply has none - its declaration is already covered above).
    for var A in Accessors do
      for var Acc in FindByProjectScan(AProjectFiles, A, '', nil) do
        if not AlreadyListed(Acc.FilePath, Acc.Line) then
          ResultList.Add(Acc);

    Result := ResultList.ToArray;
  finally
    Accessors.Free;
    ResultList.Free;
  end;
end;

class function TImplementationFinder.FindTypeImplementations(
  const AProjectFiles: TArray<string>; const ATypeName: string;
  AProgress: TImplementationScanProgress): TFindReferenceItems;
type
  TDeclInfo = record
    FilePath: string;
    Line, Col: Integer;
    Preview: string;
    Parents: TArray<string>;
  end;
var
  Decls: TDictionary<string, TDeclInfo>;   // UPPER(name) -> declaration
  Order: TList<string>;                    // stable output order
  ResultList: TList<TFindReferenceItem>;
  Item: TFindReferenceItem;
  Info: TDeclInfo;
  Lines: TArray<string>;
  RawContent, TypeName: string;
  LineIdx, P: Integer;

  // Walks the parent chain in memory - no further file reads.
  function Descends(const AName: string; ADepth: Integer): Boolean;
  var
    D: TDeclInfo;
  begin
    Result := False;
    if (AName = '') or (ADepth > 32) then Exit;
    if SameText(AName, ATypeName) then Exit(True);
    if not Decls.TryGetValue(UpperCase(AName), D) then Exit;
    for var Par in D.Parents do
      if SameText(Par, ATypeName) or Descends(Par, ADepth + 1) then
        Exit(True);
  end;

begin
  Result := nil;
  if (ATypeName = '') or (System.Length(AProjectFiles) = 0) then Exit;

  Decls := TDictionary<string, TDeclInfo>.Create;
  Order := TList<string>.Create;
  ResultList := TList<TFindReferenceItem>.Create;
  try
    // Pass 1: collect EVERY class/interface declaration with its parents.
    for var FileIdx := 0 to High(AProjectFiles) do
    begin
      if Assigned(AProgress) then
        AProgress(FileIdx + 1, System.Length(AProjectFiles));
      try
        RawContent := ReadDelphiFile(AProjectFiles[FileIdx]);
        if Pos('CLASS', UpperCase(RawContent)) = 0 then Continue;
        Lines := ReadDelphiFileLines(AProjectFiles[FileIdx]);
      except
        Continue;
      end;

      for LineIdx := 0 to High(Lines) do
      begin
        TypeName := TImplFinderHelper.ExtractTypeHeaderName(Lines[LineIdx]);
        if TypeName = '' then Continue;
        // A forward declaration carries no parents - never let it shadow
        // the real one.
        if EndsStr(';', Trim(Lines[LineIdx]))
          and Decls.ContainsKey(UpperCase(TypeName)) then Continue;

        Info := Default(TDeclInfo);
        Info.FilePath := AProjectFiles[FileIdx];
        Info.Line := LineIdx;
        P := Pos(UpperCase(TypeName), UpperCase(Lines[LineIdx]));
        if P < 1 then P := 1;
        Info.Col := P - 1;
        Info.Preview := Trim(Lines[LineIdx]);
        Info.Parents := TImplFinderHelper.ExtractClassParents(Lines, LineIdx);

        if not Decls.ContainsKey(UpperCase(TypeName)) then
          Order.Add(UpperCase(TypeName));
        Decls.AddOrSetValue(UpperCase(TypeName), Info);
      end;
    end;

    // Pass 2: keep the ones that reach ATypeName (excluding itself).
    for var Key in Order do
    begin
      Info := Decls[Key];
      if SameText(Key, ATypeName) then Continue;
      if not Descends(Key, 0) then Continue;
      Item.FilePath := Info.FilePath;
      Item.Line := Info.Line;
      Item.Col := Info.Col;
      Item.Length := System.Length(Key);
      Item.Preview := Info.Preview;
      ResultList.Add(Item);
    end;

    Result := ResultList.ToArray;
  finally
    ResultList.Free;
    Order.Free;
    Decls.Free;
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
