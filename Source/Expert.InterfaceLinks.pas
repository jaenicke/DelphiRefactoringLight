(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.InterfaceLinks;

// Interface <-> class links of a METHOD, for find references (user request
// 2026-09-19): a class method that implements an interface method is "used"
// by that interface declaration - even when nobody ever calls it through the
// interface - and calls through the interface reach it. The other way round,
// an interface method is implemented (and reached) by the implementing class
// methods.
//
//   InterfaceMethodsImplementedBy('TFoo', 'Bar') -> IFoo.Bar, IBase.Bar ...
//     every interface in TFoo's parent list, plus the interfaces THEY inherit
//     from (IFoo = interface(IBase)), that declares Bar;
//   ClassMethodsImplementing('IFoo', 'Bar') -> TFoo.Bar, TOther.Bar ...
//     every class whose parent list reaches IFoo (directly or through an
//     interface descending from it) and declares Bar.
//
// The type graph covers the scanned files; interfaces declared elsewhere
// (RTL/VCL, libraries) are loaded on demand through the identifier index.
// Text level (comments and strings masked), no DelphiLSP: the scans verify
// the call sites through DelphiLSP afterwards, with these positions as
// additional targets.

interface

uses
  System.SysUtils, System.Generics.Collections, Expert.IncludeExpansion;

type
  TMemberLink = record
    TypeName: string;       // 'IFoo' / 'TFoo'
    IsInterface: Boolean;   // the TYPE of this link
    FilePath: string;
    Line: Integer;          // 0-based line of the member declaration
    Col: Integer;           // 0-based column of the member NAME there
    ImplLine: Integer;      // class method implementation header, -1 = none
    Text: string;           // the declaration line (trimmed)
  end;

  TTypeDecl = record
    Name: string;
    FilePath: string;
    Line: Integer;
    IsInterface: Boolean;
    Parents: TArray<string>;
  end;

  TTypeGraph = class
  private
    FReader: TIncludeReader;
    FTypes: TDictionary<string, TTypeDecl>;     // UPPER name -> declaration
    FContents: TDictionary<string, string>;     // UPPER path -> content
    FLoadedFiles: TDictionary<string, Boolean>;
    FGlobals: TDictionary<string, string>;      // UPPER name -> declared type
    FTypeMiss: TDictionary<string, Boolean>;    // names the index could not resolve
    FMasked: TDictionary<string, TArray<string>>;
    FRawLines: TDictionary<string, TArray<string>>;
    FUseIndex: Boolean;
    procedure LoadFile(const AFile: string);
    function Content(const AFile: string): string;
    function TryType(const AName: string; out ADecl: TTypeDecl): Boolean;
    function MemberLink(const ADecl: TTypeDecl; const AMember: string;
      out ALink: TMemberLink): Boolean;
  public
    /// <summary>AReader: content source (buffer or disk); nil = disk.
    ///  AUseIndex: resolve types outside AFiles through the identifier
    ///  index (off in tests).</summary>
    constructor Create(const AFiles: TArray<string>; const AReader: TIncludeReader;
      AUseIndex: Boolean = True);
    destructor Destroy; override;
    /// <summary>Interface methods AMember of every interface AClassName
    ///  implements (its parent list and their interface ancestors).</summary>
    function InterfaceMethodsImplementedBy(const AClassName,
      AMember: string): TArray<TMemberLink>;
    /// <summary>Class methods AMember of every scanned class that
    ///  implements AInterfaceName (directly or via a descendant interface).</summary>
    function ClassMethodsImplementing(const AInterfaceName,
      AMember: string): TArray<TMemberLink>;
    /// <summary>The declaration of AName, if known.</summary>
    function FindType(const AName: string; out ADecl: TTypeDecl): Boolean;
    /// <summary>The declared type of a GLOBAL variable or typed constant
    ///  AName, looked up through the identifier index (which holds the
    ///  top-level declarations of every reachable unit) - for a use site
    ///  whose qualifier lives in another unit. '' when unknown, and always
    ///  '' without the index (tests).</summary>
    function TypeOfGlobal(const AName: string): string;
    /// <summary>The declaration of AMember in ATypeName or, when that type
    ///  does not declare it, in its nearest ANCESTOR that does (classes and
    ///  interfaces alike). This is the member lookup the compiler does and
    ///  DelphiLSP sometimes refuses to (a private/public overload pair
    ///  makes it answer nothing at all) - resolving a use site through the
    ///  declared type of its qualifier is then the only way left.
    ///  AAmbiguous says the type declares the member MORE THAN ONCE
    ///  (overloads): the position is then one of several and callers must
    ///  not treat it as THE declaration.</summary>
    function FindMember(const ATypeName, AMember: string;
      out ALink: TMemberLink; out AAmbiguous: Boolean): Boolean;
    /// <summary>EVERY declaration of AMember in ATypeName (or, when that
    ///  type does not declare it, in its nearest ancestor that does):
    ///  overloads come back as several links, which is what lets a call be
    ///  attributed by its ARGUMENT COUNT.</summary>
    function FindMembers(const ATypeName, AMember: string): TArray<TMemberLink>;
    /// <summary>Masked lines of AFile, cached - a scan asks for the same
    ///  file once per candidate, and masking a unit is not cheap.</summary>
    function MaskedLines(const AFile, AContent: string): TArray<string>;
    /// <summary>The virtual / override chain of AMember around the class
    ///  AClassName: up through the ancestors as long as they declare the
    ///  member (the topmost one introduces it), then every SCANNED class
    ///  descending from that one which declares it again. The start class
    ///  itself is part of the result.</summary>
    function ClassHierarchyMembers(const AClassName, AMember: string): TArray<TMemberLink>;
  end;

/// <summary>Type declarations in ALines (masked text is parsed, AFile only
///  labels them): "X = class(...)", "X = interface(...)", generic names
///  stripped to their base name; forward declarations are skipped.</summary>
function ParseTypeDecls(const AFile: string; const ALines: TArray<string>): TArray<TTypeDecl>;

/// <summary>Parent list of a type header at ALine0 (may wrap lines): unit
///  qualifiers and generic arguments removed ('System.IInterface' ->
///  'IInterface', 'IFoo<T>' -> 'IFoo').</summary>
function TypeHeaderParents(const ALines: TArray<string>; ALine0: Integer): TArray<string>;

/// <summary>0-based line of the implementation header "AType.AMember" in
///  ALines ('procedure TFoo.Bar', also 'function TFoo<T>.Bar'), -1 = none.</summary>
function FindMethodImplLine(const ALines: TArray<string>; const AType, AMember: string): Integer;

/// <summary>ALines[AFrom] plus the following lines while its brackets are
///  still open, joined by a blank - a call or a declaration whose parameter
///  list WRAPS ("Self.Init( nil," / "True," / ...). Positions on the first
///  line keep their column. Line comments are cut, string literals kept.</summary>
function JoinOpenParenLines(const ALines: TArray<string>; AFrom: Integer;
  AMaxLines: Integer = 40): string;

type
  /// <summary>Positions of a symbol's relatives with a label each ("declared
  ///  in interface IFoo", "implemented by TFoo") - used by the find
  ///  references scans to tell WHY a hit belongs to the symbol.</summary>
  TLinkedTargets = class
  private
    FTypes: TDictionary<string, string>;   // key -> 'I:IFoo' / 'C:TFoo'
    class function Key(const AFile: string; ALine: Integer): string; static;
  public
    constructor Create;
    destructor Destroy; override;
    /// <summary>A position of a relative: the member in interface / class
    ///  ATypeName.</summary>
    procedure Add(const AFile: string; ALine: Integer; AIsInterface: Boolean;
      const ATypeName: string);
    function Contains(const AFile: string; ALine: Integer): Boolean;
    /// <summary>Any linked position in that FILE - the coarse test for
    ///  cases where a line cannot be pinned down (overloads).</summary>
    function ContainsFile(const AFile: string): Boolean;
    /// <summary>Is ATypeName one of the linked types (an implementer of
    ///  the searched interface method, or the interface of a class
    ///  method)? The test for overloads that cannot be pinned to a line.</summary>
    function HasType(const ATypeName: string): Boolean;
    /// <summary>For a hit AT the position: "declared in interface IFoo" /
    ///  "implemented by TFoo"; '' when not linked.</summary>
    function DeclLabel(const AFile: string; ALine: Integer): string;
    /// <summary>For a hit that LEADS to the position: "call via interface
    ///  IFoo" / "call via class TFoo"; '' when not linked.</summary>
    function CallLabel(const AFile: string; ALine: Integer): string;
    function Count: Integer;
    /// <summary>Every linked position with its type, for diagnostics.</summary>
    function Text: string;
  end;

/// <summary>Fills ATargets for the symbol AMember of AOwnerType: when
///  AOwnerType is a CLASS, the interface declarations it implements
///  ("declared in interface IFoo"); when it is an INTERFACE, the implementing
///  class methods' declarations and implementations ("implemented by
///  TFoo"). Returns the links (for rows the scans add themselves - e.g. an
///  interface outside the scanned files).</summary>
function CollectLinkedTargets(AGraph: TTypeGraph; const AOwnerType, AMember: string;
  ATargets: TLinkedTargets): TArray<TMemberLink>;

type
  /// <summary>Outcome of ResolveMemberUse. murAmbiguous carries a position
  ///  too, but the type declares the member several times (overloads), so
  ///  it says "this type" and not "this declaration".</summary>
  TMemberUseResult = (murNone, murResolved, murAmbiguous);

/// <summary>Which member declaration a DOTTED use site refers to, decided
///  WITHOUT DelphiLSP: the qualifier before the caret is resolved to its
///  declared type (DeclaredTypeOfIdentifier) and that type - or its
///  nearest ancestor declaring it - gives the member. ALine0/ACol0 are the
///  0-based position of AMember in AContent.
///  This is the answer for the use sites DelphiLSP refuses: since Delphi
///  13.1 a private/public overload pair makes definition and completion
///  return nothing at all (RSS-5463), and a text scan alone cannot tell
///  whose member it is.</summary>
function ResolveMemberUse(AGraph: TTypeGraph; const AFile, AContent: string;
  ALine0, ACol0: Integer; const AMember: string;
  out ALink: TMemberLink): TMemberUseResult;

type
  /// <summary>What an occurrence DelphiLSP gave NO answer for turned out
  ///  to be. uuOtherSymbol is the valuable one: it belongs to a different
  ///  type, so the scans can drop it instead of listing it as unverified
  ///  noise. uuOverloaded = the right type, but which overload cannot be
  ///  decided from the text.</summary>
  TUnansweredUse = (uuUnknown, uuOurs, uuOtherSymbol, uuOverloaded);

/// <summary>Classifies an occurrence DelphiLSP did not resolve, using the
///  declared type of its qualifier (see ResolveMemberUse). AIsTargetLine
///  answers whether a position is one of the searched symbol's
///  declarations; AHasTargetInFile whether a file holds any of them - an
///  overloaded member is decided at FILE level, because the text cannot
///  tell its overloads apart.</summary>
function ClassifyUnansweredUse(AGraph: TTypeGraph; const AFile, AContent: string;
  ALine0, ACol0: Integer; const AMember: string;
  const AIsTargetLine: TFunc<string, Integer, Boolean>;
  const AIsOurType: TFunc<string, Boolean>;
  out ALink: TMemberLink): TUnansweredUse;

implementation

uses
  System.StrUtils, Expert.PascalScanner, Expert.UnitIndex, Delphi.FileEncoding,
  Expert.SignatureEdit;

function SplitContentLinesLocal(const AContent: string): TArray<string>;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

function StripGenericAndUnit(const AName: string): string;
begin
  Result := Trim(AName);
  var LT := Pos('<', Result);
  if LT > 0 then Result := Trim(Copy(Result, 1, LT - 1));
  var Dot := LastDelimiter('.', Result);
  if Dot > 0 then Result := Copy(Result, Dot + 1, MaxInt);
end;

function TypeHeaderParents(const ALines: TArray<string>; ALine0: Integer): TArray<string>;
var
  Combined: string;
  Open, Close, Depth, I: Integer;
begin
  Result := nil;
  if (ALine0 < 0) or (ALine0 > High(ALines)) then Exit;
  Combined := ALines[ALine0];
  var Eq := Pos('=', Combined);
  Open := Pos('(', Combined, Eq + 1);
  if (Eq = 0) or (Open = 0) then Exit;
  // the '(' must follow the class / interface keyword directly
  var Between := UpperCase(Trim(Copy(Combined, Eq + 1, Open - Eq - 1)));
  if Between.StartsWith('PACKED ') then Between := TrimLeft(Copy(Between, 8, MaxInt));
  // class / class abstract / class sealed / interface / dispinterface
  if (Between <> 'CLASS') and (Between <> 'INTERFACE') and (Between <> 'DISPINTERFACE') and
     not Between.StartsWith('CLASS ') then Exit;
  I := ALine0;
  Close := 0;
  while Close = 0 do
  begin
    Depth := 0;
    for var K := Open to Length(Combined) do
      case Combined[K] of
        '(', '<': Inc(Depth);
        '>': Dec(Depth);
        ')':
          begin
            Dec(Depth);
            if Depth = 0 then begin Close := K; Break; end;
          end;
      end;
    if Close = 0 then
    begin
      Inc(I);
      if (I > High(ALines)) or (I - ALine0 > 20) then Exit;
      Combined := Combined + ' ' + ALines[I];
    end;
  end;
  // split at top-level commas (generic arguments may contain commas)
  var Inside := Copy(Combined, Open + 1, Close - Open - 1);
  var Part := '';
  Depth := 0;
  for var K := 1 to Length(Inside) + 1 do
  begin
    if (K > Length(Inside)) or ((Inside[K] = ',') and (Depth = 0)) then
    begin
      Part := StripGenericAndUnit(Part);
      if Part <> '' then Result := Result + [Part];
      Part := '';
      Continue;
    end;
    case Inside[K] of
      '<': Inc(Depth);
      '>': Dec(Depth);
    end;
    Part := Part + Inside[K];
  end;
end;

function ParseTypeDecls(const AFile: string; const ALines: TArray<string>): TArray<TTypeDecl>;
var
  M: TArray<string>;
begin
  Result := nil;
  M := MaskCommentsAndStrings(ALines);
  for var L := 0 to High(M) do
  begin
    var T := Trim(M[L]);
    var Eq := Pos('=', T);
    if Eq < 2 then Continue;
    var Name := Trim(Copy(T, 1, Eq - 1));
    // "type TFoo = class" on one line
    if UpperCase(Name).StartsWith('TYPE ') then Name := Trim(Copy(Name, 6, MaxInt));
    Name := StripGenericAndUnit(Name);
    if not IsIdentifier(Name) then Continue;
    var Rest := UpperCase(Trim(Copy(T, Eq + 1, MaxInt)));
    if Rest.StartsWith('PACKED ') then Rest := TrimLeft(Copy(Rest, 8, MaxInt));
    var IsIntf := Rest.StartsWith('INTERFACE') or Rest.StartsWith('DISPINTERFACE');
    // "class", "class(", "class abstract", "class sealed" - but not "class of"
    var IsClass := Rest.StartsWith('CLASS') and not Rest.StartsWith('CLASS OF');
    // RECORDs and old-style OBJECTs have methods too, and a use site does
    // not care which kind its type is ("Idx.Init(Src)"). They were missing
    // here, which made every record member unresolvable - the forum
    // example of post #150 is records throughout.
    var IsRecord := (Rest.StartsWith('RECORD') or Rest.StartsWith('OBJECT'))
      and not Rest.StartsWith('RECORD HELPER') and not Rest.StartsWith('OBJECT OF');
    if not (IsIntf or IsClass or IsRecord) then Continue;
    // a forward declaration ("TFoo = class;" / "IFoo = interface;") has
    // neither parents nor a body
    var Word := 'CLASS';
    if IsIntf then
      Word := IfThen(Rest.StartsWith('DISP'), 'DISPINTERFACE', 'INTERFACE')
    else if IsRecord then
      Word := IfThen(Rest.StartsWith('OBJECT'), 'OBJECT', 'RECORD');
    if Trim(Copy(Rest, Length(Word) + 1, MaxInt)) = ';' then Continue;
    var D: TTypeDecl;
    D.Name := Name;
    D.FilePath := AFile;
    D.Line := L;
    D.IsInterface := IsIntf;
    D.Parents := TypeHeaderParents(M, L);
    Result := Result + [D];
  end;
end;

function FindMethodImplLine(const ALines: TArray<string>; const AType, AMember: string): Integer;
var
  M: TArray<string>;
  U, Want: string;
begin
  Result := -1;
  M := MaskCommentsAndStrings(ALines);
  Want := UpperCase(AMember);
  for var L := 0 to High(M) do
  begin
    U := UpperCase(TrimLeft(M[L]));
    if U.StartsWith('CLASS ') then U := TrimLeft(Copy(U, 7, MaxInt));
    var KW := '';
    for var K in ['PROCEDURE ', 'FUNCTION ', 'CONSTRUCTOR ', 'DESTRUCTOR '] do
      if U.StartsWith(K) then KW := K;
    if KW = '' then Continue;
    U := TrimLeft(Copy(U, Length(KW) + 1, MaxInt));
    // qualified name up to '(' ';' ':' or blank; generic arguments removed
    var I := 1;
    var Q := '';
    var Depth := 0;
    while I <= Length(U) do
    begin
      var C := U[I];
      if C = '<' then Inc(Depth)
      else if C = '>' then Dec(Depth)
      else if (Depth = 0) and CharInSet(C, ['(', ';', ':', ' ']) then Break
      else if Depth = 0 then Q := Q + C;
      Inc(I);
    end;
    var Dot := LastDelimiter('.', Q);
    if Dot = 0 then Continue;
    if not SameText(Copy(Q, Dot + 1, MaxInt), Want) then Continue;
    var Owner := Copy(Q, 1, Dot - 1);
    var OD := LastDelimiter('.', Owner);
    Owner := Copy(Owner, OD + 1, MaxInt);
    if SameText(Owner, AType) then Exit(L);
  end;
end;

{ TTypeGraph }

constructor TTypeGraph.Create(const AFiles: TArray<string>; const AReader: TIncludeReader;
  AUseIndex: Boolean);
begin
  inherited Create;
  FReader := AReader;
  FUseIndex := AUseIndex;
  FTypes := TDictionary<string, TTypeDecl>.Create;
  FContents := TDictionary<string, string>.Create;
  FLoadedFiles := TDictionary<string, Boolean>.Create;
  FGlobals := TDictionary<string, string>.Create;
  FTypeMiss := TDictionary<string, Boolean>.Create;
  FMasked := TDictionary<string, TArray<string>>.Create;
  FRawLines := TDictionary<string, TArray<string>>.Create;
  for var F in AFiles do LoadFile(F);
end;

destructor TTypeGraph.Destroy;
begin
  FLoadedFiles.Free;
  FGlobals.Free;
  FTypeMiss.Free;
  FMasked.Free;
  FRawLines.Free;
  FContents.Free;
  FTypes.Free;
  inherited;
end;

function TTypeGraph.Content(const AFile: string): string;
begin
  if FContents.TryGetValue(UpperCase(AFile), Result) then Exit;
  Result := '';
  if Assigned(FReader) then
  begin
    if not FReader(AFile, Result) then Result := '';
  end
  else
    try
      Result := ReadDelphiFile(AFile);
    except
      Result := '';
    end;
  FContents.AddOrSetValue(UpperCase(AFile), Result);
end;

procedure TTypeGraph.LoadFile(const AFile: string);
begin
  var Key := UpperCase(ExpandFileName(AFile));
  if FLoadedFiles.ContainsKey(Key) then Exit;
  FLoadedFiles.Add(Key, True);
  var C := Content(AFile);
  if (C = '') or ((Pos('CLASS', UpperCase(C)) = 0) and (Pos('INTERFACE', UpperCase(C)) = 0)) then
    Exit;
  for var D in ParseTypeDecls(AFile, SplitContentLinesLocal(C)) do
    // first real declaration wins (a scan file before a library one)
    if not FTypes.ContainsKey(UpperCase(D.Name)) then
      FTypes.Add(UpperCase(D.Name), D);
end;

function TTypeGraph.TryType(const AName: string; out ADecl: TTypeDecl): Boolean;
const
  // Loading files for a name the index knows from many units (a common
  // identifier, or one that is no type at all) used to read and parse
  // whatever the index offered - per occurrence. A scan then spent its
  // time in RTL/VCL sources instead of asking DelphiLSP (tester: "find
  // references got noticeably slower", 2026-09-20).
  MaxOnDemandFiles = 4;
begin
  if FTypes.TryGetValue(UpperCase(AName), ADecl) then Exit(True);
  Result := False;
  if not FUseIndex then Exit;
  if FTypeMiss.ContainsKey(UpperCase(AName)) then Exit;
  var Snap := TUnitIndex.Instance.Snapshot;
  if Snap = nil then Exit;
  var Loaded := 0;
  for var H in Snap.Lookup(AName) do
  begin
    if Loaded >= MaxOnDemandFiles then Break;
    if FLoadedFiles.ContainsKey(UpperCase(ExpandFileName(H.Path))) then Continue;
    LoadFile(H.Path);
    Inc(Loaded);
    if FTypes.ContainsKey(UpperCase(AName)) then Break;
  end;
  Result := FTypes.TryGetValue(UpperCase(AName), ADecl);
  // remember the misses too: the same qualifier appears again and again
  if not Result then FTypeMiss.AddOrSetValue(UpperCase(AName), True);
end;

function TTypeGraph.FindType(const AName: string; out ADecl: TTypeDecl): Boolean;
begin
  Result := TryType(StripGenericAndUnit(AName), ADecl);
end;

function TTypeGraph.TypeOfGlobal(const AName: string): string;
begin
  Result := '';
  if not FUseIndex then Exit;
  if FGlobals.TryGetValue(UpperCase(AName), Result) then Exit;
  var Snap := TUnitIndex.Instance.Snapshot;
  if Snap <> nil then
    for var H in Snap.Lookup(AName) do
    begin
      if H.Path = '' then Continue;
      var C := Content(H.Path);
      if C = '' then Continue;
      var L := FindDeclarationLine(C, AName);
      if L < 0 then Continue;
      Result := DeclaredTypeOfIdentifier(C, L, AName);
      if Result <> '' then Break;
    end;
  // '' is cached too: an identifier that is no global must not be looked
  // up again for every one of its occurrences
  FGlobals.AddOrSetValue(UpperCase(AName), Result);
end;

// Arguments of the call that starts right after the member name at ACol0
// (0-based): 0 for "Foo;" / "Foo.Bar", -1 when the list is not closed on
// this line (then the text cannot count reliably). Nested calls, strings
// and brackets are skipped, so "Foo(A, B(C, D))" counts 2.
function JoinOpenParenLines(const ALines: TArray<string>; AFrom: Integer;
  AMaxLines: Integer): string;

  function DepthOf(const S: string; ADepth: Integer): Integer;
  begin
    Result := ADepth;
    var I := 1;
    while I <= Length(S) do
    begin
      case S[I] of
        #39:
          begin
            Inc(I);
            while (I <= Length(S)) and (S[I] <> #39) do Inc(I);
          end;
        '(', '[': Inc(Result);
        ')', ']': if Result > 0 then Dec(Result);
      end;
      Inc(I);
    end;
  end;

begin
  Result := '';
  if (AFrom < 0) or (AFrom > High(ALines)) then Exit;
  Result := StripLineComment(ALines[AFrom]);
  var Depth := DepthOf(Result, 0);
  var I := AFrom + 1;
  while (Depth > 0) and (I <= High(ALines)) and (I - AFrom <= AMaxLines) do
  begin
    var S := StripLineComment(ALines[I]);
    Result := Result + ' ' + S;
    Depth := DepthOf(S, Depth);
    Inc(I);
  end;
end;

function CallArgumentCount(const ALine: string; ACol0: Integer): Integer;
var
  I, Depth, Count: Integer;
  Empty: Boolean;
begin
  I := ACol0 + 1;                      // 1-based, first char after the name
  while (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  if I > Length(ALine) then Exit(0);   // name at the line end: no arguments
  if ALine[I] <> '(' then
  begin
    if CharInSet(ALine[I], [';', ',', ')', ']', '.', '=', '<', '>', '+', '-', '*', '/']) then
      Exit(0);                         // a call without a list, or a use
    Exit(-1);
  end;
  Inc(I);
  Depth := 1;
  Count := 1;
  Empty := True;
  while I <= Length(ALine) do
  begin
    var C := ALine[I];
    if C = #39 then                    // a string literal
    begin
      Empty := False;
      Inc(I);
      while (I <= Length(ALine)) and (ALine[I] <> #39) do Inc(I);
    end
    else if CharInSet(C, ['(', '[']) then
    begin
      Inc(Depth);
      Empty := False;
    end
    else if CharInSet(C, [')', ']']) then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        if Empty and (Count = 1) then Count := 0;
        Exit(Count);
      end;
    end
    else if (C = ',') and (Depth = 1) then Inc(Count)
    else if not CharInSet(C, [' ', #9]) then Empty := False;
    Inc(I);
  end;
  Result := -1;                        // list not closed on this line
end;

// Can ADecl ("procedure Load(const AName: string; AFlag: Boolean = False);")
// be called with ACount arguments? Parameters with a default value lower
// the minimum.
function MemberAcceptsArgCount(const ADecl: string; ACount: Integer): Boolean;
var
  Open, Close, Depth, I, MinCount, MaxCount: Integer;
begin
  Open := 0;
  Close := 0;
  Depth := 0;
  for I := 1 to Length(ADecl) do
    if ADecl[I] = '(' then
    begin
      Inc(Depth);
      if Depth = 1 then Open := I;
    end
    else if ADecl[I] = ')' then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        Close := I;
        Break;
      end;
    end;
  if (Open = 0) or (Close = 0) then Exit(ACount = 0);   // no parameter list
  MinCount := 0;
  MaxCount := 0;
  for var P in ParseParamList(Copy(ADecl, Open + 1, Close - Open - 1)) do
  begin
    Inc(MaxCount);
    if P.DefaultText = '' then Inc(MinCount);
  end;
  Result := (ACount >= MinCount) and (ACount <= MaxCount);
end;

// The ARGUMENT TEXTS of the call after the member name at ACol0, or nil
// when there is no argument list on this line. "Foo(A, B(C, D))" -> two.
function CallArgumentTexts(const ALine: string; ACol0: Integer): TArray<string>;
var
  I, Depth, Start: Integer;
begin
  Result := nil;
  I := ACol0 + 1;
  while (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Inc(I);
  if (I > Length(ALine)) or (ALine[I] <> '(') then Exit;
  Inc(I);
  Depth := 1;
  Start := I;
  while I <= Length(ALine) do
  begin
    var C := ALine[I];
    if C = #39 then
    begin
      Inc(I);
      while (I <= Length(ALine)) and (ALine[I] <> #39) do Inc(I);
    end
    else if CharInSet(C, ['(', '[']) then Inc(Depth)
    else if CharInSet(C, [')', ']']) then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        if Trim(Copy(ALine, Start, I - Start)) <> '' then
          Result := Result + [Trim(Copy(ALine, Start, I - Start))];
        Exit;
      end;
    end
    else if (C = ',') and (Depth = 1) then
    begin
      Result := Result + [Trim(Copy(ALine, Start, I - Start))];
      Start := I + 1;
    end;
    Inc(I);
  end;
  Result := nil;                       // list not closed on this line
end;

// The TYPE of one argument, as far as the TEXT can tell: an explicit
// typecast "TMyListA(X)", a literal, or a plain identifier whose
// declaration is in this file. '' = unknown, and an unknown argument
// never disqualifies a candidate.
function ArgumentTypeName(AGraph: TTypeGraph; const ALines, AMasked: TArray<string>;
  ALine0: Integer; const AArg: string): string;
var
  P: Integer;
  Head: string;
  Decl: TTypeDecl;
begin
  Result := '';
  if AArg = '' then Exit;
  if AArg[1] = #39 then Exit('string');
  if SameText(AArg, 'True') or SameText(AArg, 'False') then Exit('Boolean');
  if SameText(AArg, 'nil') then Exit('');          // fits every class type
  // "TMyListA(AListe)" - a typecast, and the most explicit answer there is
  P := Pos('(', AArg);
  if (P > 1) and AArg.EndsWith(')') then
  begin
    Head := Trim(Copy(AArg, 1, P - 1));
    if IsIdentifier(Head) and AGraph.FindType(Head, Decl) then Exit(Head);
  end;
  if IsIdentifier(AArg) then
    Result := DeclaredTypeOfIdentifierIn(ALines, AMasked, ALine0, AArg);
end;

// Does ADecl's parameter list accept an argument of type ATypeName at
// position AIndex? Unknown ('') always fits; a DESCENDANT fits its
// ancestor's parameter type.
function ParamTypeAccepts(AGraph: TTypeGraph; const ADecl: string;
  AIndex: Integer; const ATypeName: string): Boolean;
var
  Open, Close, Depth, I: Integer;
begin
  Result := True;
  if ATypeName = '' then Exit;
  Open := 0;
  Close := 0;
  Depth := 0;
  for I := 1 to Length(ADecl) do
    if ADecl[I] = '(' then
    begin
      Inc(Depth);
      if Depth = 1 then Open := I;
    end
    else if ADecl[I] = ')' then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        Close := I;
        Break;
      end;
    end;
  if (Open = 0) or (Close = 0) then Exit;
  var Params := ParseParamList(Copy(ADecl, Open + 1, Close - Open - 1));
  if (AIndex < 0) or (AIndex > High(Params)) then Exit;
  var Want := StripGenericAndUnit(Trim(Params[AIndex].TypeText));
  if (Want = '') or SameText(Want, ATypeName) then Exit;
  // a descendant may be passed where the ancestor is expected
  var Name := ATypeName;
  var D: TTypeDecl;
  for var Step := 0 to 16 do
  begin
    if not AGraph.FindType(Name, D) or (Length(D.Parents) = 0) then Break;
    Name := StripGenericAndUnit(D.Parents[0]);
    if SameText(Name, Want) then Exit;
  end;
  Result := False;
end;

function ResolveMemberUse(AGraph: TTypeGraph; const AFile, AContent: string;
  ALine0, ACol0: Integer; const AMember: string;
  out ALink: TMemberLink): TMemberUseResult;
var
  Lines: TArray<string>;
  Decl: TTypeDecl;
  TypeName, Qualifier: string;
begin
  Result := murNone;
  ALink := Default(TMemberLink);
  if (AGraph = nil) or (AContent = '') or (AMember = '') then Exit;
  Lines := SplitContentLinesLocal(AContent);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then Exit;
  var Masked := AGraph.MaskedLines(AFile, AContent);
  Qualifier := QualifierBefore(Lines[ALine0], ACol0);
  if SameText(Qualifier, 'Self') then
  begin
    // "Self.Init(...)" is the enclosing type's member, like the unqualified
    // form - but without its shadowing question: a local variable can never
    // hide Self.X. The qualifier is no declared identifier, so the lookup
    // found nothing and the occurrence stayed UNVERIFIED (forum,
    // 2026-09-20).
    var SFirst, SLast: Integer;
    if not FindEnclosingRoutineRangeIn(Lines, ALine0, SFirst, SLast) then Exit;
    TypeName := OwnerTypeOfImplHeader(Lines[SFirst]);
  end
  else if Qualifier = '' then
  begin
    // UNQUALIFIED use ("Init(TMyListA(AListe), ASpur);" inside another
    // method of the same type): Delphi resolves it against the enclosing
    // type, so we do too - unless the name is a local variable, which
    // would shadow the member (tester 2026-09-20, forum #156).
    var First, Last: Integer;
    if not FindEnclosingRoutineRangeIn(Lines, ALine0, First, Last) then Exit;
    if DeclaredTypeOfIdentifierIn(Lines, Masked, ALine0, AMember) <> '' then Exit;
    TypeName := OwnerTypeOfImplHeader(Lines[First]);
    if TypeName = '' then Exit;
  end
  else
  begin
    // The qualifier is either a TYPE itself ("TMyClass.Create") or an
    // expression whose declared type we have to look up ("lMyClassA.Init").
    if AGraph.FindType(Qualifier, Decl) then
      TypeName := Decl.Name
    else
      TypeName := DeclaredTypeOfIdentifierIn(Lines, Masked, ALine0, Qualifier);
    // Not declared in this file? Then it is a global of ANOTHER unit
    // ("WizardInstance.Execute"), which the identifier index can point to.
    if TypeName = '' then
      TypeName := AGraph.TypeOfGlobal(Qualifier);
  end;
  if TypeName = '' then Exit;

  var Links := AGraph.FindMembers(TypeName, AMember);
  if Length(Links) = 0 then Exit;
  ALink := Links[0];
  if Length(Links) = 1 then Exit(murResolved);

  // OVERLOADS: the text cannot compare parameter TYPES, but it can count.
  // A call with N arguments can only reach a declaration that accepts N -
  // when exactly one does, the use site is resolved after all. That is the
  // "one of the overloads is strict private" case the tester still saw as
  // UNVERIFIED (2026-09-20).
  // the call may WRAP over several lines - then its argument list is not
  // closed on the use line and nothing could be counted (forum, 2026-09-20)
  var CallText := JoinOpenParenLines(Lines, ALine0, 40);
  var ArgCount := CallArgumentCount(CallText, ACol0 + Length(AMember));
  if ArgCount >= 0 then
  begin
    var Fits: TArray<TMemberLink> := nil;
    for var L in Links do
      if MemberAcceptsArgCount(L.Text, ArgCount) then Fits := Fits + [L];
    if Length(Fits) = 1 then
    begin
      ALink := Fits[0];
      Exit(murResolved);
    end;
    // Same argument COUNT in several overloads - then their TYPES decide,
    // as far as the text shows them: an explicit typecast
    // "Init(TMyListA(AListe), ASpur)" names the parameter type outright
    // (the tester's record with two 2-parameter overloads, 2026-09-20).
    if Length(Fits) > 1 then
    begin
      var Args := CallArgumentTexts(CallText, ACol0 + Length(AMember));
      if Length(Args) > 0 then
      begin
        var ByType: TArray<TMemberLink> := nil;
        for var L in Fits do
        begin
          var Ok := True;
          for var I := 0 to High(Args) do
            if not ParamTypeAccepts(AGraph, L.Text, I,
              ArgumentTypeName(AGraph, Lines, Masked, ALine0, Args[I])) then
            begin
              Ok := False;
              Break;
            end;
          if Ok then ByType := ByType + [L];
        end;
        if Length(ByType) = 1 then
        begin
          ALink := ByType[0];
          Exit(murResolved);
        end;
      end;
    end;
  end;
  Result := murAmbiguous;
end;

function ClassifyUnansweredUse(AGraph: TTypeGraph; const AFile, AContent: string;
  ALine0, ACol0: Integer; const AMember: string;
  const AIsTargetLine: TFunc<string, Integer, Boolean>;
  const AIsOurType: TFunc<string, Boolean>;
  out ALink: TMemberLink): TUnansweredUse;
begin
  Result := uuUnknown;
  case ResolveMemberUse(AGraph, AFile, AContent, ALine0, ACol0, AMember, ALink) of
    murResolved:
      if Assigned(AIsTargetLine) and AIsTargetLine(ALink.FilePath, ALink.Line) then
        Result := uuOurs
      else
        Result := uuOtherSymbol;
    murAmbiguous:
      // overloads the argument count could not tell apart: only the TYPE
      // can decide. Comparing FILES was wrong - an unrelated class in the
      // same unit counted as ours then (tester, 2026-09-20: "a method that
      // is only in the strict private part, twice overloaded, is shown
      // although it has nothing to do with the search").
      if Assigned(AIsOurType) and AIsOurType(ALink.TypeName) then
        Result := uuOverloaded
      else
        Result := uuOtherSymbol;
  end;
end;


function TTypeGraph.MaskedLines(const AFile, AContent: string): TArray<string>;
begin
  var Key := UpperCase(ExpandFileName(AFile));
  if (AFile <> '') and FMasked.TryGetValue(Key, Result) then Exit;
  var Lines := SplitContentLinesLocal(AContent);
  Result := MaskCommentsAndStrings(Lines);
  if AFile <> '' then
  begin
    FMasked.AddOrSetValue(Key, Result);
    FRawLines.AddOrSetValue(Key, Lines);
  end;
end;

function TTypeGraph.FindMembers(const ATypeName, AMember: string): TArray<TMemberLink>;
var
  Link: TMemberLink;
  Ambiguous: Boolean;
begin
  Result := nil;
  if not FindMember(ATypeName, AMember, Link, Ambiguous) then Exit;
  Result := [Link];
  if not Ambiguous then Exit;
  // overloads: EVERY declaration of that name in the same type body, so a
  // call can be attributed by its argument count
  var C := Content(Link.FilePath);
  if C = '' then Exit;
  var Lines := SplitContentLinesLocal(C);
  var All: TArray<TMemberLink> := nil;
  for var L in FindMemberDeclarationLines(C, Link.TypeName, AMember) do
  begin
    if (L < 0) or (L > High(Lines)) then Continue;
    var L2 := Link;
    L2.Line := L;
    // a WRAPPED parameter list belongs to the declaration, or its
    // argument count cannot be compared (forum, 2026-09-20)
    L2.Text := Trim(JoinOpenParenLines(Lines, L, 40));
    var P := Pos(UpperCase(AMember), UpperCase(Lines[L]));
    if P > 0 then L2.Col := P - 1;
    All := All + [L2];
  end;
  if Length(All) > 0 then Result := All;
end;


function TTypeGraph.FindMember(const ATypeName, AMember: string;
  out ALink: TMemberLink; out AAmbiguous: Boolean): Boolean;
var
  Queue: TArray<string>;
  Seen: TArray<string>;
  Decl: TTypeDecl;

  function AlreadySeen(const AName: string): Boolean;
  begin
    Result := False;
    for var S in Seen do
      if SameText(S, AName) then Exit(True);
  end;

begin
  Result := False;
  AAmbiguous := False;
  ALink := Default(TMemberLink);
  if (ATypeName = '') or (AMember = '') then Exit;
  Queue := [StripGenericAndUnit(ATypeName)];
  // breadth first through the parents, so the type's OWN declaration wins
  // over an inherited one; 64 steps are far more than any real hierarchy
  for var Step := 0 to 63 do
  begin
    if Length(Queue) = 0 then Exit;
    var Name := Queue[0];
    Delete(Queue, 0, 1);
    if AlreadySeen(Name) then Continue;
    Seen := Seen + [Name];
    if not TryType(Name, Decl) then Continue;
    if MemberLink(Decl, AMember, ALink) then
    begin
      // Overloads: the type declares the member more than once, so this
      // position is one of several - the caller must not take it as THE
      // declaration (and that is exactly the shape DelphiLSP chokes on).
      var U := ' ' + UpperCase(ALink.Text) + ' ';
      AAmbiguous := (Pos(' OVERLOAD;', U) > 0) or (Pos(' OVERLOAD ', U) > 0);
      Exit(True);
    end;
    for var P in Decl.Parents do
      Queue := Queue + [StripGenericAndUnit(P)];
  end;
end;

function TTypeGraph.MemberLink(const ADecl: TTypeDecl; const AMember: string;
  out ALink: TMemberLink): Boolean;
begin
  Result := False;
  var C := Content(ADecl.FilePath);
  var L := FindMemberDeclarationLine(C, ADecl.Name, AMember);
  if L < 0 then Exit;
  var Lines := SplitContentLinesLocal(C);
  if L > High(Lines) then Exit;
  ALink := Default(TMemberLink);
  ALink.TypeName := ADecl.Name;
  ALink.IsInterface := ADecl.IsInterface;
  ALink.FilePath := ADecl.FilePath;
  ALink.Line := L;
  ALink.Text := Trim(JoinOpenParenLines(Lines, L, 40));
  // the member NAME as a whole word on that line
  var U := UpperCase(Lines[L]);
  var W := UpperCase(AMember);
  var P := Pos(W, U);
  ALink.Col := 0;
  while P > 0 do
  begin
    var E := P + Length(W);
    if ((P = 1) or not IsIdentChar(U[P - 1])) and ((E > Length(U)) or not IsIdentChar(U[E])) then
    begin
      ALink.Col := P - 1;
      Break;
    end;
    P := Pos(W, U, P + 1);
  end;
  ALink.ImplLine := -1;
  if not ADecl.IsInterface then
    ALink.ImplLine := FindMethodImplLine(Lines, ADecl.Name, AMember);
  Result := True;
end;

function TTypeGraph.InterfaceMethodsImplementedBy(const AClassName,
  AMember: string): TArray<TMemberLink>;
var
  Seen: TDictionary<string, Boolean>;

  procedure Walk(const AIntf: string; ADepth: Integer);
  var
    D: TTypeDecl;
    Link: TMemberLink;
  begin
    if (ADepth > 16) or Seen.ContainsKey(UpperCase(AIntf)) then Exit;
    Seen.Add(UpperCase(AIntf), True);
    if not TryType(AIntf, D) or not D.IsInterface then Exit;
    if MemberLink(D, AMember, Link) then Result := Result + [Link];
    for var P in D.Parents do Walk(P, ADepth + 1);
  end;

var
  Cls: TTypeDecl;
begin
  Result := nil;
  if not TryType(StripGenericAndUnit(AClassName), Cls) or Cls.IsInterface then Exit;
  Seen := TDictionary<string, Boolean>.Create;
  try
    for var P in Cls.Parents do Walk(P, 0);
  finally
    Seen.Free;
  end;
end;

function TTypeGraph.ClassMethodsImplementing(const AInterfaceName,
  AMember: string): TArray<TMemberLink>;
var
  Target: string;

  // AName (an interface) is ATarget or descends from it
  function Reaches(const AName: string; ADepth: Integer): Boolean;
  var
    D: TTypeDecl;
  begin
    if SameText(AName, Target) then Exit(True);
    if ADepth > 16 then Exit(False);
    Result := False;
    if not TryType(AName, D) or not D.IsInterface then Exit;
    for var P in D.Parents do
      if Reaches(P, ADepth + 1) then Exit(True);
  end;

begin
  Result := nil;
  Target := StripGenericAndUnit(AInterfaceName);
  // only the SCANNED types are candidates (FTypes before any index loads)
  var Classes: TArray<TTypeDecl> := nil;
  for var D in FTypes.Values do
    if not D.IsInterface then Classes := Classes + [D];
  for var D in Classes do
  begin
    var Hit := False;
    for var P in D.Parents do
      if Reaches(P, 0) then begin Hit := True; Break; end;
    if not Hit then Continue;
    var Link: TMemberLink;
    if MemberLink(D, AMember, Link) then Result := Result + [Link];
  end;
end;

function TTypeGraph.ClassHierarchyMembers(const AClassName, AMember: string): TArray<TMemberLink>;
var
  Root: TTypeDecl;
  Link: TMemberLink;

  // the class parent of ADecl (the first parent that is a known class)
  function ClassParent(const ADecl: TTypeDecl; out AParent: TTypeDecl): Boolean;
  begin
    Result := False;
    if Length(ADecl.Parents) = 0 then Exit;
    Result := TryType(ADecl.Parents[0], AParent) and not AParent.IsInterface;
  end;

  // AName descends from Root (or is it)
  function DescendsFromRoot(const ADecl: TTypeDecl): Boolean;
  var
    D, P: TTypeDecl;
  begin
    D := ADecl;
    for var Depth := 0 to 32 do
    begin
      if SameText(D.Name, Root.Name) then Exit(True);
      if not ClassParent(D, P) then Exit(False);
      D := P;
    end;
    Result := False;
  end;

var
  Cur, P: TTypeDecl;
begin
  Result := nil;
  if not TryType(StripGenericAndUnit(AClassName), Cur) or Cur.IsInterface then Exit;
  if not MemberLink(Cur, AMember, Link) then Exit;
  Root := Cur;
  for var Depth := 0 to 32 do
  begin
    if not ClassParent(Root, P) or not MemberLink(P, AMember, Link) then Break;
    Root := P;
  end;
  // the scanned classes (plus the ancestors just loaded) below Root
  var Classes: TArray<TTypeDecl> := nil;
  for var D in FTypes.Values do
    if not D.IsInterface then Classes := Classes + [D];
  for var D in Classes do
    if DescendsFromRoot(D) and MemberLink(D, AMember, Link) then
      Result := Result + [Link];
end;

{ TLinkedTargets }

class function TLinkedTargets.Key(const AFile: string; ALine: Integer): string;
begin
  Result := UpperCase(ExpandFileName(AFile)) + '|' + IntToStr(ALine);
end;

constructor TLinkedTargets.Create;
begin
  inherited;
  FTypes := TDictionary<string, string>.Create;
end;

destructor TLinkedTargets.Destroy;
begin
  FTypes.Free;
  inherited;
end;

procedure TLinkedTargets.Add(const AFile: string; ALine: Integer; AIsInterface: Boolean;
  const ATypeName: string);
begin
  if (AFile <> '') and (ALine >= 0) and not FTypes.ContainsKey(Key(AFile, ALine)) then
    FTypes.Add(Key(AFile, ALine), IfThen(AIsInterface, 'I:', 'C:') + ATypeName);
end;

function TLinkedTargets.Contains(const AFile: string; ALine: Integer): Boolean;
begin
  Result := FTypes.ContainsKey(Key(AFile, ALine));
end;

function TLinkedTargets.ContainsFile(const AFile: string): Boolean;
begin
  Result := False;
  if AFile = '' then Exit;
  var Prefix := UpperCase(ExpandFileName(AFile)) + '|';
  for var K in FTypes.Keys do
    if K.StartsWith(Prefix) then Exit(True);
end;

function TLinkedTargets.HasType(const ATypeName: string): Boolean;
begin
  Result := False;
  if ATypeName = '' then Exit;
  for var V in FTypes.Values do
    if SameText(Copy(V, 3, MaxInt), ATypeName) then Exit(True);   // 'I:' / 'C:' prefix
end;

function TLinkedTargets.DeclLabel(const AFile: string; ALine: Integer): string;
var
  V: string;
begin
  Result := '';
  if not FTypes.TryGetValue(Key(AFile, ALine), V) then Exit;
  if V.StartsWith('I:') then
    Result := 'declared in interface ' + Copy(V, 3, MaxInt)
  else
    Result := 'implemented by ' + Copy(V, 3, MaxInt);
end;

function TLinkedTargets.CallLabel(const AFile: string; ALine: Integer): string;
var
  V: string;
begin
  Result := '';
  if not FTypes.TryGetValue(Key(AFile, ALine), V) then Exit;
  if V.StartsWith('I:') then
    Result := 'call via interface ' + Copy(V, 3, MaxInt)
  else
    Result := 'call via class ' + Copy(V, 3, MaxInt);
end;

function TLinkedTargets.Count: Integer;
begin
  Result := FTypes.Count;
end;

function TLinkedTargets.Text: string;
begin
  Result := '';
  for var P in FTypes do
  begin
    if Result <> '' then Result := Result + '; ';
    Result := Result + P.Key + ' (' + P.Value + ')';
  end;
end;

function CollectLinkedTargets(AGraph: TTypeGraph; const AOwnerType, AMember: string;
  ATargets: TLinkedTargets): TArray<TMemberLink>;
var
  D: TTypeDecl;
begin
  Result := nil;
  if (AOwnerType = '') or (AMember = '') or not AGraph.FindType(AOwnerType, D) then Exit;
  if D.IsInterface then
  begin
    Result := AGraph.ClassMethodsImplementing(D.Name, AMember);
    for var L in Result do
    begin
      ATargets.Add(L.FilePath, L.Line, False, L.TypeName);
      if L.ImplLine >= 0 then
        ATargets.Add(L.FilePath, L.ImplLine, False, L.TypeName);
    end;
  end
  else
  begin
    Result := AGraph.InterfaceMethodsImplementedBy(D.Name, AMember);
    for var L in Result do
      ATargets.Add(L.FilePath, L.Line, True, L.TypeName);
  end;
end;

end.
