(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ReferenceKind;

// What KIND of use a reference is (issue #11, "Find References kind
// column" by Ian Branch): declaration, implementation, call, inherited
// call, write, read, method reference, address (@), property accessor,
// type use, uses clause. Pure text classification of positions that were
// already VERIFIED to belong to the symbol - the question here is only
// "how is it used at this place", never "is it the symbol".
//
// The symbol's own kind (procedure / function / data / type, from its
// declaration line) decides the ambiguous shapes: "X := Foo;" is a CALL
// of a function, a METHOD REFERENCE of a procedure and a READ of a
// variable.

interface

uses
  System.SysUtils, System.Types;

type
  TRefSymbolKind = (rsUnknown, rsProcedure, rsFunction, rsData, rsType);

  TRefKind = (rkUse, rkDeclaration, rkImplementation, rkCall, rkInheritedCall,
    rkWrite, rkRead, rkMethodReference, rkAddress, rkAccessor, rkTypeUse,
    rkUsesClause);

function RefKindText(AKind: TRefKind): string;

/// <summary>The symbol's kind from the line that declares it ("procedure
///  X", "function X", "property X", "X: Integer", "TX = class",
///  "cMax = 10"). rsUnknown when the line says nothing clear.</summary>
function SymbolKindFromDeclLine(const ALine, AName: string): TRefSymbolKind;

/// <summary>The kind of every occurrence of a name of length ANameLen at
///  APositions (X = column, Y = line, both 0-based) in AContent - one
///  masking pass for the file.</summary>
function ClassifyReferences(const AContent: string; const APositions: TArray<TPoint>;
  ANameLen: Integer; ASymbol: TRefSymbolKind): TArray<TRefKind>;

implementation

uses
  System.StrUtils, System.RegularExpressions, Expert.PascalScanner, Expert.SignatureEdit;

function RefKindText(AKind: TRefKind): string;
const
  Texts: array[TRefKind] of string = ('Use', 'Declaration', 'Implementation', 'Call',
    'Inherited call', 'Write', 'Read', 'Method reference', 'Address (@)',
    'Property accessor', 'Type use', 'Uses clause');
begin
  Result := Texts[AKind];
end;

function StartsWithWord(const S, W: string): Boolean;
begin
  Result := S.StartsWith(W) and ((Length(S) = Length(W)) or not IsIdentChar(S[Length(W) + 1]));
end;

function StripWord(const S, W: string): string;
begin
  if StartsWithWord(S, W) then Result := TrimLeft(Copy(S, Length(W) + 1, MaxInt))
  else Result := S;
end;

function SymbolKindFromDeclLine(const ALine, AName: string): TRefSymbolKind;
var
  T, N, Rhs: string;
begin
  Result := rsUnknown;
  T := UpperCase(Trim(StripLineComment(ALine)));
  N := UpperCase(AName);
  T := StripWord(T, 'CLASS');
  if StartsWithWord(T, 'PROCEDURE') or StartsWithWord(T, 'CONSTRUCTOR') or
     StartsWithWord(T, 'DESTRUCTOR') then Exit(rsProcedure);
  if StartsWithWord(T, 'FUNCTION') or StartsWithWord(T, 'OPERATOR') then Exit(rsFunction);
  if StartsWithWord(T, 'PROPERTY') then Exit(rsData);
  for var K in ['TYPE', 'VAR', 'CONST', 'THREADVAR', 'RESOURCESTRING'] do
    T := StripWord(T, K);
  // "A, B, X: Integer" - the name somewhere in a leading name list
  var M := TRegEx.Match(T, '^([A-Z_0-9]+\s*,\s*)*' + TRegEx.Escape(N) + '\b\s*(<[^>]*>)?\s*([:=,])');
  if not M.Success then Exit;
  if M.Groups[3].Value <> '=' then Exit(rsData);
  Rhs := TrimLeft(Copy(T, M.Index + M.Length, MaxInt));
  Rhs := StripWord(Rhs, 'PACKED');
  Rhs := StripWord(Rhs, 'TYPE');
  for var K in ['CLASS', 'RECORD', 'INTERFACE', 'DISPINTERFACE', 'OBJECT', 'SET', 'ARRAY',
    'REFERENCE', 'PROCEDURE', 'FUNCTION', 'FILE'] do
    if StartsWithWord(Rhs, K) then Exit(rsType);
  if Rhs.StartsWith('(') or Rhs.StartsWith('^') then Exit(rsType);   // enum / pointer
  // "TMyInt = Integer;" (alias) vs "cMax = OtherConst;": the T-prefix
  // convention decides - a plain value is a constant
  var Ident := Rhs.EndsWith(';') and TRegEx.IsMatch(Rhs, '^[A-Z_][A-Z_0-9.]*\s*(<.*>)?\s*;$');
  if Ident and (Length(AName) > 1) and (AName[1] = 'T') and CharInSet(AName[2], ['A'..'Z']) then
    Exit(rsType);
  Result := rsData;
end;

function ClassifyReferences(const AContent: string; const APositions: TArray<TPoint>;
  ANameLen: Integer; ASymbol: TRefSymbolKind): TArray<TRefKind>;
var
  Lines, Masked: TArray<string>;
  MJ: string;
  Starts: TArray<Integer>;
  ImplLine: Integer;
  InCode: TArray<Boolean>;   // the line starts inside begin/try/case/asm ... end

  // One token pass: a stack of open blocks, 'C' = code (begin, try, case
  // in code, asm), 'D' = declaration body (record, class/interface/object
  // with a body). Needed because "cRed: DoIt;" (a case label) and
  // "FRed: Integer;" (a field) look the same on their own line.
  procedure ComputeCodeLines;
  var
    Sc: TPascalScanner;
    T: TPasToken;
  begin
    SetLength(InCode, Length(Lines));
    var Stack := '';
    var LastLine := -1;
    var Prev1 := '';
    var Prev2 := '';
    var Pending := False;      // after "= class": does a body follow?
    var ParenDepth := 0;       // skipping "class(TBase, IFoo)"
    Sc := TPascalScanner.Create(MJ);
    try
      while Sc.Next(T) do
      begin
        while LastLine < T.Line do
        begin
          Inc(LastLine);
          if LastLine <= High(InCode) then
            InCode[LastLine] := (Stack <> '') and (Stack[Length(Stack)] = 'C');
        end;
        var U := T.Upper;
        if Pending then
        begin
          if ParenDepth > 0 then
          begin
            if T.IsSymbol('(') then Inc(ParenDepth)
            else if T.IsSymbol(')') then Dec(ParenDepth);
            Prev2 := Prev1; Prev1 := U;
            Continue;
          end;
          if T.IsSymbol('(') and (Prev1 <> ')') then
          begin
            ParenDepth := 1;
            Prev2 := Prev1; Prev1 := U;
            Continue;
          end;
          Pending := False;
          // "TFoo = class;" / "class of" / "TFoo = class(TBar);" open nothing
          if not (T.IsSymbol(';') or T.IsWord('of')) then Stack := Stack + 'D';
        end;
        if T.Kind = ptIdent then
        begin
          if (U = 'BEGIN') or (U = 'TRY') or (U = 'ASM') then
            Stack := Stack + 'C'
          else if U = 'CASE' then
          begin
            // a variant record's "case" has no own end
            if (Stack <> '') and (Stack[Length(Stack)] = 'C') then Stack := Stack + 'C';
          end
          else if U = 'RECORD' then
            Stack := Stack + 'D'
          else if ((U = 'CLASS') or (U = 'OBJECT') or (U = 'INTERFACE') or
            (U = 'DISPINTERFACE')) and ((Prev1 = '=') or ((Prev1 = 'PACKED') and (Prev2 = '='))) then
            Pending := True
          else if U = 'END' then
          begin
            if Stack <> '' then SetLength(Stack, Length(Stack) - 1);
          end;
        end;
        Prev2 := Prev1;
        Prev1 := U;
      end;
    finally
      Sc.Free;
    end;
    for var L := LastLine + 1 to High(InCode) do
      InCode[L] := (Stack <> '') and (Stack[Length(Stack)] = 'C');
  end;

  function PrevCodePos(AFrom: Integer): Integer;   // last code char before AFrom
  begin
    Result := AFrom - 1;
    while (Result >= 1) and CharInSet(MJ[Result], [' ', #9, #10, #13]) do Dec(Result);
  end;

  function NextCodePos(AFrom: Integer): Integer;   // first code char at/after AFrom
  begin
    Result := AFrom;
    while (Result <= Length(MJ)) and CharInSet(MJ[Result], [' ', #9, #10, #13]) do Inc(Result);
  end;

  function WordEndingAt(AEnd: Integer): string;
  begin
    var B := AEnd;
    while (B >= 1) and IsIdentChar(MJ[B]) do Dec(B);
    Result := UpperCase(Copy(MJ, B + 1, AEnd - B));
  end;

  // the uses clause: back over names, dots, commas and "in" to USES
  function InUsesClause(AStart: Integer): Boolean;
  begin
    var B := AStart - 1;
    while (B >= 1) and (IsIdentChar(MJ[B]) or CharInSet(MJ[B], ['.', ',', ' ', #9, #10, #13])) do
      Dec(B);
    // the walk may run on over "interface" / "implementation" - the clause
    // keyword just has to be IN the walked span
    var S := ' ' + UpperCase(Copy(MJ, B + 1, AStart - B - 1)) + ' ';
    Result := TRegEx.IsMatch(S, '[^A-Z_0-9.](USES|CONTAINS|REQUIRES)[^A-Z_0-9.]');
  end;

  // the routine / property NAME of a header line, and whether it is
  // qualified ("TFoo.Bar")
  function HeaderNameCheck(ALine, ACol: Integer; out AQualified, AIsProperty: Boolean): Boolean;
  var
    S, U: string;
    P: Integer;
  begin
    Result := False;
    AQualified := False;
    AIsProperty := False;
    S := Masked[ALine];
    U := UpperCase(S);
    P := 1;
    while (P <= Length(U)) and CharInSet(U[P], [' ', #9]) do Inc(P);
    var Rest := Copy(U, P, MaxInt);
    if StartsWithWord(Rest, 'CLASS') then
    begin
      Inc(P, 5);
      while (P <= Length(U)) and CharInSet(U[P], [' ', #9]) do Inc(P);
      Rest := Copy(U, P, MaxInt);
    end;
    var KW := '';
    for var K in ['PROCEDURE', 'FUNCTION', 'CONSTRUCTOR', 'DESTRUCTOR', 'OPERATOR', 'PROPERTY'] do
      if StartsWithWord(Rest, K) then KW := K;
    if KW = '' then Exit;
    AIsProperty := KW = 'PROPERTY';
    Inc(P, Length(KW));
    while (P <= Length(U)) and CharInSet(U[P], [' ', #9]) do Inc(P);
    // the (qualified, maybe generic) name runs to '(' ':' ';' or a blank
    var Q := P;
    var Depth := 0;
    while Q <= Length(U) do
    begin
      if U[Q] = '<' then Inc(Depth)
      else if U[Q] = '>' then Dec(Depth)
      else if (Depth = 0) and CharInSet(U[Q], ['(', ':', ';', ' ', #9, '[']) then Break;
      Inc(Q);
    end;
    // ACol (0-based) must start the LAST dotted segment of that name
    var Seg := P;
    Depth := 0;
    for var K := P to Q - 1 do
      if U[K] = '<' then Inc(Depth)
      else if U[K] = '>' then Dec(Depth)
      else if (Depth = 0) and (U[K] = '.') then
      begin
        Seg := K + 1;
        AQualified := True;
      end;
    Result := ACol + 1 = Seg;
  end;

  // "A, B, X: T" / "X = ..." as a declaration at the start of a line -
  // never inside a code block, where the same shape is a case label
  function LeadingDeclaration(ALine, ACol: Integer): Boolean;
  var
    U: string;
  begin
    Result := False;
    if InCode[ALine] then Exit;
    U := UpperCase(Masked[ALine]);
    var Head := UpperCase(Copy(U, 1, ACol));
    Head := Trim(Head);
    for var K in ['CLASS VAR', 'CLASS', 'VAR', 'CONST', 'THREADVAR', 'TYPE', 'OUT', 'CONSTREF'] do
      if Head = K then Head := ''
      else if Head.StartsWith(K + ' ') then Head := TrimLeft(Copy(Head, Length(K) + 1, MaxInt));
    // only names and commas in front. NOTE: TRegEx.IsMatch answers False
    // for an EMPTY subject even when the pattern matches '' - and nothing
    // in front of the name is the normal case
    if (Head <> '') and not TRegEx.IsMatch(Head, '^([A-Z_0-9]+\s*,\s*)*$') then Exit;
    var After := TrimLeft(Copy(U, ACol + ANameLen + 1, MaxInt));
    if After.StartsWith('<') and (Pos('>', After) > 0) then
      After := TrimLeft(Copy(After, Pos('>', After) + 1, MaxInt));
    if After.StartsWith(':=') then Exit;
    if After.StartsWith(':') or After.StartsWith(',') then Exit(True);
    // "TFoo = class ..." / "cMax = 10;" - a type or constant declaration
    // (outside code, a line cannot start with a comparison)
    if After.StartsWith('=') then Result := True;
  end;

var
  Rx: TArray<TRefKind>;
begin
  Result := nil;
  MJ := SigJoinedText(AContent);
  Lines := MJ.Split([#10]);
  Masked := MaskCommentsAndStrings(Lines);
  MJ := string.Join(#10, Masked);
  SetLength(Starts, Length(Lines));
  var Ofs := 1;
  for var L := 0 to High(Lines) do
  begin
    Starts[L] := Ofs;
    Inc(Ofs, Length(Lines[L]) + 1);
  end;
  ComputeCodeLines;
  ImplLine := MaxInt;
  for var L := 0 to High(Masked) do
    if SameText(Trim(Masked[L]), 'implementation') then
    begin
      ImplLine := L;
      Break;
    end;

  SetLength(Rx, Length(APositions));
  for var I := 0 to High(APositions) do
  begin
    var Line := APositions[I].Y;
    var Col := APositions[I].X;
    Rx[I] := rkUse;
    if (Line < 0) or (Line > High(Lines)) then Continue;
    var S0 := Starts[Line] + Col;
    var E := S0 + ANameLen - 1;
    if (E > Length(MJ)) or (S0 < 1) then Continue;

    if InUsesClause(S0) then begin Rx[I] := rkUsesClause; Continue; end;

    var Qualified, IsProp: Boolean;
    if HeaderNameCheck(Line, Col, Qualified, IsProp) then
    begin
      if IsProp then Rx[I] := rkDeclaration
      else if Qualified or ((Line > ImplLine) and (Lines[Line] = TrimLeft(Lines[Line]))) then
        Rx[I] := rkImplementation
      else
        Rx[I] := rkDeclaration;
      Continue;
    end;

    var P := PrevCodePos(S0);
    var WB := '';
    if (P >= 1) and IsIdentChar(MJ[P]) then WB := WordEndingAt(P);
    // "var X", "for var X", "const X" (inline), "out X" (parameter)
    if (WB = 'VAR') or (WB = 'CONST') or (WB = 'OUT') or (WB = 'THREADVAR') or
       (WB = 'CONSTREF') or LeadingDeclaration(Line, Col) then
    begin
      Rx[I] := rkDeclaration;
      Continue;
    end;

    if ASymbol = rsType then begin Rx[I] := rkTypeUse; Continue; end;

    var O, C: Integer;
    var Ctx := CallContextAt(MJ, S0, E, O, C);
    if Ctx = ccAccessor then begin Rx[I] := rkAccessor; Continue; end;
    // a type position (": T", "is T", "as T", "of T") of a symbol whose kind
    // is unknown
    if (ASymbol = rsUnknown) and (((P >= 1) and (MJ[P] = ':') and
       ((P + 1 > Length(MJ)) or (MJ[P + 1] <> '='))) or (WB = 'IS') or (WB = 'AS') or (WB = 'OF')) then
    begin
      Rx[I] := rkTypeUse;
      Continue;
    end;
    var AtSign := (P >= 1) and (MJ[P] = '@');
    var IsInherited := WB = 'INHERITED';
    // the next code after the name (and an index "[...]")
    var N := NextCodePos(E + 1);
    if (N <= Length(MJ)) and (MJ[N] = '[') then
    begin
      var CB := MatchingBracket(MJ, N);
      if CB > 0 then N := NextCodePos(CB + 1);
    end;
    var IsAssign := (N < Length(MJ)) and (MJ[N] = ':') and (MJ[N + 1] = '=');

    case ASymbol of
      rsProcedure, rsFunction:
        if AtSign then Rx[I] := rkMethodReference
        else if Ctx in [ccArgs, ccStatement, ccExpression] then
        begin
          if IsInherited then Rx[I] := rkInheritedCall else Rx[I] := rkCall;
        end
        else if ASymbol = rsFunction then
          Rx[I] := rkCall                 // "X := Foo;" / "Bar(Foo)" call a function
        else
          Rx[I] := rkMethodReference;     // a procedure there is a reference
      rsData:
        if AtSign then Rx[I] := rkAddress
        else if IsAssign then Rx[I] := rkWrite
        else Rx[I] := rkRead;
    else
      if AtSign then Rx[I] := rkAddress
      else if IsAssign then Rx[I] := rkWrite
      else if IsInherited then Rx[I] := rkInheritedCall
      else if Ctx = ccArgs then Rx[I] := rkCall
      else Rx[I] := rkUse;
    end;
  end;
  Result := Rx;
end;

end.
