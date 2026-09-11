(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.EventGen;

// Generating event handlers and anonymous methods from code completion
// (forum request: "look up the event type, then build the method by hand -
// every single time").
//
// At two kinds of places the EXPECTED type is known, and when it is a
// procedural type the matching code can be written for the user:
//   Button1.OnClick := |       an assignment - the target's type
//   TThread.Queue(nil, |)      a call argument - the parameter's type
// Completion then offers, at the top of its list:
//   * "procedure(Sender: TObject) begin ... end" - an anonymous method,
//     inserted in place (for "reference to" types);
//   * "Button1Click" - a new method in the current class, declared in its
//     private section, implemented right after the current method and
//     inserted at the caret (for "of object" types, and as an alternative
//     for "reference to" types, which accept methods too).
//
// Measured with DelphiLSP before designing this (scratchpad lspprobe):
//   * signatureHelp gives parameter labels WITH the type
//     ("AThreadProc: TThreadProcedure"), overloads included;
//   * hover on a variable/field gives "var TProbe.FCallback: TProc2";
//   * hover on a REPUBLISHED property gives only "property TButton.OnClick"
//     - no type. The type then comes from walking the class chain in the
//     sources (TButton -> TCustomButton -> TButtonControl -> TWinControl ->
//     TControl, "property OnClick: TNotifyEvent ...").
// Type names are resolved through source text (identifier index + file
// contents, injected as TGenTypeSource) - aliases, generic arity
// (TProc<T1,T2> is a different declaration than TProc<T>) and type
// parameter substitution included.
//
// Everything here is PURE (no editor, no LSP) and covered by the console
// suite; Expert.CompletionWizard does the LSP calls and the edits.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Expert.UnitIndex;

type
  TGenSite = (gsNone, gsAssign, gsArgument);

  TGenContext = record
    Site: TGenSite;
    /// <summary>gsAssign: the LAST identifier of the assignment target
    ///  ("OnClick" in "Button1.OnClick := "), 0-based position - the place
    ///  to ask the LSP for its type.</summary>
    TargetName: string;
    TargetLine: Integer;
    TargetCol: Integer;
    /// <summary>The identifier before the target's dot ("Button1"), ''
    ///  when there is none - it names the handler.</summary>
    QualifierName: string;
    /// <summary>Partially typed identifier at the caret and its 0-based
    ///  start column - replaced by the insertion.</summary>
    Prefix: string;
    PrefixStartCol: Integer;
    /// <summary>gsArgument: 0-based position of the CALLED routine's name
    ///  ("CreateAnonymousThread") - hovering it names the unit that
    ///  declares the routine, i.e. where its parameter types are resolved.
    ///  -1 when unknown (a generic call "Foo<T>(").</summary>
    CallLine: Integer;
    CallCol: Integer;
  end;

  TProcKind = (pkNone, pkMethod, pkReference, pkPlain);

  TProcTypeInfo = record
    Kind: TProcKind;       // of object / reference to / plain procedural
    IsFunction: Boolean;
    Params: string;        // 'Sender: TObject', '' without parameters
    ResultType: string;
    TypeName: string;      // as it was asked for, for display
  end;

  /// <summary>One candidate declaration source: a unit's path and the part
  ///  of its text that is visible for the lookup.</summary>
  TGenDecl = record
    Path: string;
    Content: string;
  end;

  /// <summary>Candidate declarations of AIdentifier, ORDERED BY VISIBILITY
  ///  from AContextFile ('' = the current unit): the context unit itself,
  ///  then units in its uses clause (last one first), then the rest. The
  ///  first declaration found decides - exactly like the compiler.</summary>
  TGenTypeSource = reference to function(const AIdentifier, AContextFile: string): TArray<TGenDecl>;

  TGenPlan = record
    Ok: Boolean;
    Reason: string;
    ClassName: string;
    /// <summary>Insert DeclText at the START of 0-based line DeclLine0
    ///  (it ends with a line break); DeclLines = lines it adds.</summary>
    DeclLine0: Integer;
    DeclText: string;
    DeclLines: Integer;
    ImplLine0: Integer;
    ImplText: string;
    /// <summary>0-based line of the empty body line inside ImplText,
    ///  relative to the ORIGINAL numbering (add DeclLines once the
    ///  declaration above has been inserted).</summary>
    ImplBodyLine0: Integer;
  end;

function DetectGenContext(const ALines: TArray<string>; ALine0, ACol0: Integer): TGenContext;

/// <summary>Type named by a hover text ("var TProbe.FCallback: TProc2" ->
///  'TProc2'). '' when the hover names no type; AOwnerClass is the class
///  in front of the member name ("property TButton.OnClick" -> 'TButton').</summary>
function TypeFromHover(const AHover: string; out AOwnerClass: string): string;

/// <summary>'const AProc: TProc<Integer> = nil' -> 'TProc<Integer>'.</summary>
function TypeFromParamLabel(const ALabel: string): string;
/// <summary>'const AProc: TProc<Integer>' -> 'AProc'.</summary>
function ParamNameFromLabel(const ALabel: string): string;

/// <summary>Parses the right-hand side of a procedural type declaration:
///  'procedure(Sender: TObject) of object', 'reference to function: T', ...</summary>
function ParseProcTypeText(const AText: string; out AInfo: TProcTypeInfo): Boolean;

/// <summary>'TFunc<Integer, TList<string>>' -> base 'TFunc', args
///  ['Integer', 'TList<string>'].</summary>
function SplitGenericArgs(const ATypeName: string; out ABase: string): TArray<string>;

/// <summary>Whole-word replacement of type parameter names in one pass.</summary>
function SubstituteTypeParams(const AText: string;
  const ANames, AValues: TArray<string>): string;

/// <summary>Finds "ABase<P1,..> = <rhs>;" with AArity type parameters in
///  AContent and returns the right-hand side (joined across lines).</summary>
function FindTypeDeclText(const AContent, ABase: string; AArity: Integer;
  out ATypeParams: TArray<string>; out ARhs: string): Boolean;

/// <summary>Resolves a procedural type name as seen from AContextFile.</summary>
function ResolveProcType(const ATypeName, AContextFile: string;
  const ASource: TGenTypeSource; out AInfo: TProcTypeInfo): Boolean;

/// <summary>Type of member AMember of class AClass, walking the parent
///  chain for republished properties ("property OnClick;").</summary>
function ResolveMemberProcType(const AClass, AMember, AContextFile: string;
  const ASource: TGenTypeSource; out AInfo: TProcTypeInfo): Boolean;

/// <summary>"procedure(Sender: TObject)" / "function(const A: T): Boolean".</summary>
function ProcHeadText(const AInfo: TProcTypeInfo; const AName: string = ''): string;

/// <summary>The anonymous method to insert at the caret. The first line
///  continues the caret's line; the body is ABodyLineOffset lines further
///  down, at column ABodyCol (0-based).</summary>
function AnonymousMethodText(const AInfo: TProcTypeInfo; const AIndent: string;
  out ABodyLineOffset, ABodyCol: Integer): string;

/// <summary>Handler name after the IDE's convention: Button1.OnClick ->
///  Button1Click; FTimer.OnTimer -> TimerTimer; bare FOnChange -> DoChange;
///  an argument AOnDone -> HandleOnDone ... -> HandleDone.</summary>
function SuggestHandlerName(const ACtx: TGenContext; const AParamName: string): string;

/// <summary>Where and what to insert for a new method AName of the class
///  whose method contains ACaretLine0. Ok = False with a Reason when that
///  is not possible (no enclosing method, class not in this unit, name
///  taken).</summary>
function PlanEventHandler(const ALines: TArray<string>; ACaretLine0: Integer;
  const AName: string; const AInfo: TProcTypeInfo): TGenPlan;

function IsPascalIdentifier(const S: string): Boolean;

/// <summary>Indices into ACandidates (unit names), most visible first:
///  units named in AUses in REVERSE order (the last used unit wins), then
///  the rest in their original order. 'SysUtils' matches 'System.SysUtils'.</summary>
function RankUnitsByUses(const ACandidates, AUses: TArray<string>): TArray<Integer>;

/// <summary>The file URI of the declaration link DelphiLSP puts in front of
///  a hover ("[Vcl.StdCtrls.pas:1152](file:///...#1152)"), '' when none.</summary>
function HoverDeclUri(const AHover: string): string;

/// <summary>The type source the wizard uses: the current buffer, the
///  context unit's interface, and the identifier index's declarations
///  ranked by the context unit's uses clause. ALookup = index lookup,
///  AReadFile = file text (the caller caches).</summary>
function MakeIndexTypeSource(const ALookup: TFunc<string, TArray<TFindUnitHit>>;
  const AReadFile: TFunc<string, string>;
  const ACurrentFile, ACurrentContent: string): TGenTypeSource;

implementation

uses
  System.StrUtils, System.Character, System.Math, Expert.DfmRename, Expert.UsesGraph;

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := C.IsLetterOrDigit or (C = '_');
end;

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := C.IsLetter or (C = '_');
end;

function IsPascalIdentifier(const S: string): Boolean;
var
  I: Integer;
begin
  Result := (S <> '') and IsIdentStart(S[1]);
  if Result then
    for I := 2 to Length(S) do
      if not IsIdentChar(S[I]) then Exit(False);
end;

function LeadingSpaces(const S: string): string;
var
  I: Integer;
begin
  I := 1;
  while (I <= Length(S)) and CharInSet(S[I], [' ', #9]) do Inc(I);
  Result := Copy(S, 1, I - 1);
end;

// ---------------------------------------------------------------------------
//  Where is the caret?
// ---------------------------------------------------------------------------

const
  NonCallWords: array[0..16] of string = ('if', 'while', 'until', 'and', 'or',
    'not', 'xor', 'in', 'of', 'case', 'then', 'do', 'else', 'repeat', 'div',
    'mod', 'with');

// Walks back from (ALine0, 1-based APos) to the unmatched '(' that opens
// the argument list the caret is in. True when that paren belongs to a
// call (an identifier or a generic '>' directly before it).
function InsideCallArgs(const ALines: TArray<string>; ALine0, APos: Integer;
  out ANameLine, ANameCol: Integer): Boolean;
var
  L, I, Depth, BDepth, Budget: Integer;
  S: string;
begin
  Result := False;
  ANameLine := -1;
  ANameCol := -1;
  Depth := 0;
  BDepth := 0;
  Budget := 0;
  L := ALine0;
  I := APos;
  while (L >= 0) and (Budget < 20) do
  begin
    S := ALines[L];
    if L <> ALine0 then I := Length(S);
    while I >= 1 do
    begin
      case S[I] of
        ')': Inc(Depth);
        ']': Inc(BDepth);
        '[': if BDepth = 0 then Exit(False) else Dec(BDepth);
        ';': if (Depth = 0) and (BDepth = 0) then Exit(False);
        '''':
          begin
            // skip back over a string literal on this line
            Dec(I);
            while (I >= 1) and (S[I] <> '''') do Dec(I);
          end;
        '(':
          if Depth > 0 then
            Dec(Depth)
          else
          begin
            var J := I - 1;
            while (J >= 1) and CharInSet(S[J], [' ', #9]) do Dec(J);
            if J < 1 then Exit(False);
            if S[J] = '>' then Exit(True);   // Foo<T>(
            if not IsIdentChar(S[J]) then Exit(False);
            var K := J;
            while (K >= 1) and IsIdentChar(S[K]) do Dec(K);
            var W := LowerCase(Copy(S, K + 1, J - K));
            for var NW in NonCallWords do
              if W = NW then Exit(False);
            ANameLine := L;
            ANameCol := K;
            Exit(True);
          end;
      end;
      Dec(I);
    end;
    Dec(L);
    Inc(Budget);
  end;
end;

function DetectGenContext(const ALines: TArray<string>; ALine0, ACol0: Integer): TGenContext;
var
  S, Before, L: string;
  P, E, Q: Integer;
begin
  Result := Default(TGenContext);
  Result.CallLine := -1;
  Result.CallCol := -1;
  if (ALine0 < 0) or (ALine0 > High(ALines)) then Exit;
  S := ALines[ALine0];
  if ACol0 < 0 then Exit;
  // Caret in VIRTUAL space beyond the line end: the IDE stores no trailing
  // blanks, so "OnClick := |" arrives as "OnClick :=" with the caret two
  // columns further. Pad, so prefix and positions refer to the caret.
  if ACol0 > Length(S) then
    S := S + StringOfChar(' ', ACol0 - Length(S));
  S := Copy(S, 1, ACol0);
  P := Length(S);
  while (P >= 1) and IsIdentChar(S[P]) do Dec(P);
  Result.Prefix := Copy(S, P + 1, MaxInt);
  Result.PrefixStartCol := P;
  if (Result.Prefix <> '') and not IsIdentStart(Result.Prefix[1]) then Exit;
  // Inside a comment or after a string start on this line: no site.
  if (Pos('//', S) > 0) or (Pos('{', S) > 0) then Exit;
  if Odd(S.CountChar('''')) then Exit;
  Before := TrimRight(Copy(S, 1, P));

  if Before.EndsWith(':=') then
  begin
    L := TrimRight(Copy(Before, 1, Length(Before) - 2));
    E := Length(L);
    Q := E;
    while (Q >= 1) and IsIdentChar(L[Q]) do Dec(Q);
    if (Q = E) or not IsIdentStart(L[Q + 1]) then Exit;
    Result.TargetName := Copy(L, Q + 1, E - Q);
    Result.TargetLine := ALine0;
    Result.TargetCol := Q;
    if (Q >= 1) and (L[Q] = '.') then
    begin
      var R := Q - 1;
      var Z := R;
      while (Z >= 1) and IsIdentChar(L[Z]) do Dec(Z);
      if Z < R then Result.QualifierName := Copy(L, Z + 1, R - Z);
    end;
    Result.Site := gsAssign;
    Exit;
  end;

  var NameLine, NameCol: Integer;
  if (Before.EndsWith('(') or Before.EndsWith(','))
    and InsideCallArgs(ALines, ALine0, Length(Before), NameLine, NameCol) then
  begin
    Result.Site := gsArgument;
    Result.CallLine := NameLine;
    Result.CallCol := NameCol;
  end;
end;

// ---------------------------------------------------------------------------
//  Reading types out of LSP text
// ---------------------------------------------------------------------------

// Position (1-based) of the first ':' outside ( ) [ ] < >, 0 = none.
function TopLevelColon(const S: string): Integer;
var
  I, D: Integer;
begin
  D := 0;
  for I := 1 to Length(S) do
    case S[I] of
      '(', '[', '<': Inc(D);
      ')', ']', '>': if D > 0 then Dec(D);
      ':': if (D = 0) and ((I = Length(S)) or (S[I + 1] <> '=')) then Exit(I);
    end;
  Result := 0;
end;

// Cuts a type expression at the first property clause / terminator.
function CutTypeTail(const S: string): string;
var
  U: string;
  P: Integer;
begin
  Result := Trim(S);
  U := ' ' + LowerCase(Result) + ' ';
  for var W in [' read ', ' write ', ' stored ', ' default ', ' nodefault ',
    ' index ', ' implements ', ' deprecated ', ' platform '] do
  begin
    P := Pos(W, U);
    if P > 0 then
    begin
      Result := Trim(Copy(Result, 1, P - 2));
      U := ' ' + LowerCase(Result) + ' ';
    end;
  end;
  P := Pos(';', Result);
  if P > 0 then Result := Trim(Copy(Result, 1, P - 1));
  P := Pos('=', Result);
  if P > 0 then Result := Trim(Copy(Result, 1, P - 1));   // default value
end;

function TypeFromHover(const AHover: string; out AOwnerClass: string): string;
var
  T, W, Rest, NamePart: string;
  C, P: Integer;
begin
  Result := '';
  AOwnerClass := '';
  for var Raw in AHover.Replace(#13, '').Split([#10]) do
  begin
    T := Trim(Raw);
    if (T = '') or StartsStr('```', T) or StartsStr('[', T) then Continue;
    P := Pos(' ', T);
    if P = 0 then Continue;
    W := LowerCase(Copy(T, 1, P - 1));
    if (W <> 'property') and (W <> 'var') and (W <> 'field') and (W <> 'const')
      and (W <> 'param') and (W <> 'threadvar') then Continue;
    Rest := Trim(Copy(T, P + 1, MaxInt));
    C := TopLevelColon(Rest);
    if C > 0 then
    begin
      NamePart := Trim(Copy(Rest, 1, C - 1));
      Result := CutTypeTail(Copy(Rest, C + 1, MaxInt));
    end
    else
      NamePart := CutTypeTail(Rest);
    // array property: Items[Index: Integer]
    P := Pos('[', NamePart);
    if P > 0 then NamePart := Copy(NamePart, 1, P - 1);
    P := LastDelimiter('.', NamePart);
    if P > 0 then
    begin
      AOwnerClass := Copy(NamePart, 1, P - 1);
      // a unit-qualified owner: keep the class (last dotted segment)
      var Q := LastDelimiter('.', AOwnerClass);
      if Q > 0 then AOwnerClass := Copy(AOwnerClass, Q + 1, MaxInt);
    end;
    Exit;
  end;
end;

function TypeFromParamLabel(const ALabel: string): string;
var
  C: Integer;
begin
  C := TopLevelColon(ALabel);
  if C = 0 then Exit('');
  Result := CutTypeTail(Copy(ALabel, C + 1, MaxInt));
end;

function ParamNameFromLabel(const ALabel: string): string;
var
  C: Integer;
  S: string;
begin
  S := Trim(ALabel);
  C := TopLevelColon(S);
  if C > 0 then S := Trim(Copy(S, 1, C - 1));
  // attributes like [Ref]
  if StartsStr('[', S) and (Pos(']', S) > 0) then S := Trim(Copy(S, Pos(']', S) + 1, MaxInt));
  for var W in ['const ', 'var ', 'out ', 'constref '] do
    if StartsText(W, S) then S := Trim(Copy(S, Length(W) + 1, MaxInt));
  C := Pos(',', S);
  if C > 0 then S := Trim(Copy(S, 1, C - 1));
  Result := S;
end;

function ParseProcTypeText(const AText: string; out AInfo: TProcTypeInfo): Boolean;
var
  T, U: string;
  P, D, I: Integer;
begin
  AInfo := Default(TProcTypeInfo);
  Result := False;
  T := Trim(AText);
  while T.EndsWith(';') do T := TrimRight(Copy(T, 1, Length(T) - 1));
  U := LowerCase(T);
  if StartsStr('reference to ', U) then
  begin
    AInfo.Kind := pkReference;
    T := Trim(Copy(T, Length('reference to ') + 1, MaxInt));
    U := LowerCase(T);
  end;
  if StartsStr('procedure', U) and ((Length(U) = 9) or not IsIdentChar(U[10])) then
    T := Trim(Copy(T, 10, MaxInt))
  else if StartsStr('function', U) and ((Length(U) = 8) or not IsIdentChar(U[9])) then
  begin
    AInfo.IsFunction := True;
    T := Trim(Copy(T, 9, MaxInt));
  end
  else
    Exit;

  if StartsStr('(', T) then
  begin
    D := 0;
    P := 0;
    for I := 1 to Length(T) do
    begin
      if T[I] = '(' then Inc(D)
      else if T[I] = ')' then
      begin
        Dec(D);
        if D = 0 then begin P := I; Break; end;
      end;
    end;
    if P = 0 then Exit;
    AInfo.Params := Trim(Copy(T, 2, P - 2));
    T := Trim(Copy(T, P + 1, MaxInt));
  end;

  U := LowerCase(T);
  if U.EndsWith('of object') then
  begin
    if AInfo.Kind = pkReference then Exit;   // not a valid combination
    AInfo.Kind := pkMethod;
    T := TrimRight(Copy(T, 1, Length(T) - Length('of object')));
  end
  else if AInfo.Kind = pkNone then
    AInfo.Kind := pkPlain;

  T := Trim(T);
  if AInfo.IsFunction then
  begin
    if not StartsStr(':', T) then Exit;
    AInfo.ResultType := Trim(Copy(T, 2, MaxInt));
    if AInfo.ResultType = '' then Exit;
  end
  else if T <> '' then
    Exit;   // something we do not understand after the parameters
  // calling conventions are not part of what we generate
  Result := True;
end;

function SplitTopLevel(const S: string; ASep: Char): TArray<string>;
var
  I, D, Start: Integer;
begin
  Result := nil;
  D := 0;
  Start := 1;
  for I := 1 to Length(S) do
    case S[I] of
      '<', '(', '[': Inc(D);
      '>', ')', ']': if D > 0 then Dec(D);
    else
      if (S[I] = ASep) and (D = 0) then
      begin
        Result := Result + [Trim(Copy(S, Start, I - Start))];
        Start := I + 1;
      end;
    end;
  if Trim(Copy(S, Start, MaxInt)) <> '' then
    Result := Result + [Trim(Copy(S, Start, MaxInt))];
end;

function SplitGenericArgs(const ATypeName: string; out ABase: string): TArray<string>;
var
  T: string;
  P: Integer;
begin
  Result := nil;
  T := Trim(ATypeName);
  P := Pos('<', T);
  if (P > 0) and T.EndsWith('>') then
  begin
    ABase := Trim(Copy(T, 1, P - 1));
    Result := SplitTopLevel(Copy(T, P + 1, Length(T) - P - 1), ',');
  end
  else
    ABase := T;
  P := LastDelimiter('.', ABase);
  if P > 0 then ABase := Copy(ABase, P + 1, MaxInt);
end;

function SubstituteTypeParams(const AText: string;
  const ANames, AValues: TArray<string>): string;
var
  I, J, N: Integer;
  W: string;
  Hit: Boolean;
begin
  if Length(ANames) = 0 then Exit(AText);
  Result := '';
  I := 1;
  while I <= Length(AText) do
  begin
    if IsIdentStart(AText[I]) then
    begin
      J := I;
      while (J <= Length(AText)) and IsIdentChar(AText[J]) do Inc(J);
      W := Copy(AText, I, J - I);
      Hit := False;
      for N := 0 to Min(High(ANames), High(AValues)) do
        if SameText(W, ANames[N]) then
        begin
          Result := Result + AValues[N];
          Hit := True;
          Break;
        end;
      if not Hit then Result := Result + W;
      I := J;
    end
    else
    begin
      Result := Result + AText[I];
      Inc(I);
    end;
  end;
end;

function FindTypeDeclText(const AContent, ABase: string; AArity: Integer;
  out ATypeParams: TArray<string>; out ARhs: string): Boolean;
var
  Lines: TArray<string>;
  LineNo, P, Q, D, I, K: Integer;
  S, Acc: string;
begin
  Result := False;
  ATypeParams := nil;
  ARhs := '';
  if ABase = '' then Exit;
  Lines := AContent.Replace(#13, '').Split([#10]);
  for LineNo := 0 to High(Lines) do
  begin
    S := Lines[LineNo];
    P := 1;
    while (P <= Length(S)) and CharInSet(S[P], [' ', #9]) do Inc(P);
    if not SameText(Copy(S, P, Length(ABase)), ABase) then Continue;
    Q := P + Length(ABase);
    if (Q <= Length(S)) and IsIdentChar(S[Q]) then Continue;
    while (Q <= Length(S)) and CharInSet(S[Q], [' ', #9]) do Inc(Q);
    var Params: TArray<string> := nil;
    if (Q <= Length(S)) and (S[Q] = '<') then
    begin
      K := Pos('>', S, Q);
      if K = 0 then Continue;
      for var Tp in SplitTopLevel(Copy(S, Q + 1, K - Q - 1), ',') do
      begin
        var Name := Tp;
        var C := Pos(':', Name);          // constraint: T: class
        if C > 0 then Name := Trim(Copy(Name, 1, C - 1));
        Params := Params + [Name];
      end;
      Q := K + 1;
      while (Q <= Length(S)) and CharInSet(S[Q], [' ', #9]) do Inc(Q);
    end;
    if Length(Params) <> AArity then Continue;
    if (Q > Length(S)) or (S[Q] <> '=') then Continue;
    // '=' of a type declaration, not of a comparison inside code
    Inc(Q);
    // Right-hand side up to ';' outside parentheses, at most 8 lines.
    Acc := '';
    D := 0;
    for K := LineNo to Min(LineNo + 7, High(Lines)) do
    begin
      var Part: string;
      if K = LineNo then Part := Copy(S, Q, MaxInt) else Part := Lines[K];
      var Cm := Pos('//', Part);
      if Cm > 0 then Part := Copy(Part, 1, Cm - 1);
      for I := 1 to Length(Part) do
      begin
        case Part[I] of
          '(': Inc(D);
          ')': if D > 0 then Dec(D);
          ';':
            if D = 0 then
            begin
              ARhs := Trim(Acc + ' ' + Copy(Part, 1, I - 1));
              ATypeParams := Params;
              Exit(True);
            end;
        end;
      end;
      Acc := Acc + ' ' + Part;
    end;
    // no terminator (a class body follows): the first line is enough to
    // tell that it is not a procedural type
    ARhs := Trim(Copy(S, Q, MaxInt));
    ATypeParams := Params;
    Exit(True);
  end;
end;

function IsTypeReference(const S: string): Boolean;
begin
  Result := S <> '';
  for var C in S do
    if not (IsIdentChar(C) or CharInSet(C, ['.', '<', '>', ',', ' '])) then
      Exit(False);
  Result := Result and IsIdentStart(S[1]);
end;

function ResolveProcTypeDepth(const ATypeName, AContextFile: string;
  const ASource: TGenTypeSource; out AInfo: TProcTypeInfo; ADepth: Integer): Boolean;
var
  Base, Rhs, A: string;
  Args, TypeParams: TArray<string>;
begin
  Result := False;
  AInfo := Default(TProcTypeInfo);
  if ADepth > 6 then Exit;
  // An inline procedural type ("reference to procedure" in a parameter
  // label) needs no lookup at all.
  if ParseProcTypeText(ATypeName, AInfo) then
  begin
    AInfo.TypeName := Trim(ATypeName);
    Exit(True);
  end;
  Args := SplitGenericArgs(ATypeName, Base);
  if (Base = '') or not Assigned(ASource) then Exit;
  // The candidates arrive ORDERED BY VISIBILITY from AContextFile, so the
  // first declaration with this name and arity is the one the compiler
  // sees there - and it decides. (Tester: IBObjects' IB_Utils declares
  // "TProc = procedure;" in its interface; taken as "the first TProc in
  // the index" it turned TThread.CreateAnonymousThread's "reference to
  // procedure" into a plain procedure, and nothing was offered.)
  for var Decl in ASource(Base, AContextFile) do
  begin
    if not FindTypeDeclText(Decl.Content, Base, Length(Args), TypeParams, Rhs) then
      Continue;
    Rhs := SubstituteTypeParams(Rhs, TypeParams, Args);
    if ParseProcTypeText(Rhs, AInfo) then
    begin
      AInfo.TypeName := Trim(ATypeName);
      Exit(True);
    end;
    // alias: TMyEvent = TNotifyEvent; / TMyEvent = type TNotifyEvent;
    // resolved as seen from the unit that declares the alias
    A := Trim(Rhs);
    if StartsText('type ', A) then A := Trim(Copy(A, 6, MaxInt));
    if IsTypeReference(A) and not SameText(A, Trim(ATypeName)) then
      if ResolveProcTypeDepth(A, Decl.Path, ASource, AInfo, ADepth + 1) then
      begin
        AInfo.TypeName := Trim(ATypeName);
        Exit(True);
      end;
    Exit(False);   // the visible declaration is not procedural
  end;
end;

function ResolveProcType(const ATypeName, AContextFile: string;
  const ASource: TGenTypeSource; out AInfo: TProcTypeInfo): Boolean;
begin
  Result := ResolveProcTypeDepth(ATypeName, AContextFile, ASource, AInfo, 0);
end;

// 'property OnClick: TNotifyEvent read ...' / 'FCallback: TProc2;' ->
// the type; '' for a republished 'property OnClick;'.
function MemberTypeFromDeclLine(const ALine: string): string;
var
  T: string;
  P: Integer;
begin
  T := Trim(ALine);
  if StartsText('class ', T) then T := Trim(Copy(T, 7, MaxInt));
  if StartsText('property ', T) then T := Trim(Copy(T, 10, MaxInt));
  // skip the name and an index parameter list
  P := 1;
  while (P <= Length(T)) and IsIdentChar(T[P]) do Inc(P);
  if (P <= Length(T)) and (T[P] = '[') then
  begin
    var D := 0;
    while P <= Length(T) do
    begin
      if T[P] = '[' then Inc(D)
      else if T[P] = ']' then
      begin
        Dec(D);
        if D = 0 then begin Inc(P); Break; end;
      end;
      Inc(P);
    end;
  end;
  T := Trim(Copy(T, P, MaxInt));
  if not StartsStr(':', T) then Exit('');
  Result := CutTypeTail(Copy(T, 2, MaxInt));
end;

function ResolveMemberProcType(const AClass, AMember, AContextFile: string;
  const ASource: TGenTypeSource; out AInfo: TProcTypeInfo): Boolean;
var
  Cls, Context, Next, NextContext, TypeName: string;
  Depth, Line: Integer;
  Parents: TDictionary<string, string>;
begin
  Result := False;
  AInfo := Default(TProcTypeInfo);
  if not Assigned(ASource) then Exit;
  SplitGenericArgs(AClass, Cls);
  Context := AContextFile;
  Parents := TDictionary<string, string>.Create;
  try
    for Depth := 0 to 24 do
    begin
      Next := '';
      NextContext := '';
      for var Decl in ASource(Cls, Context) do
      begin
        var Lines := Decl.Content.Replace(#13, '').Split([#10]);
        Line := FindMemberDeclarationLine(Decl.Content, Cls, AMember);
        if (Line >= 0) and (Line <= High(Lines)) then
        begin
          TypeName := MemberTypeFromDeclLine(Lines[Line]);
          if TypeName <> '' then
            // the member's type as seen from the unit declaring the class
            Exit(ResolveProcType(TypeName, Decl.Path, ASource, AInfo));
        end;
        Parents.Clear;
        CollectClassParents(Lines, Parents);
        if Parents.TryGetValue(UpperCase(Cls), Next) then
        begin
          NextContext := Decl.Path;
          Break;
        end;
      end;
      if Next = '' then Exit;
      Cls := Next;
      Context := NextContext;
    end;
  finally
    Parents.Free;
  end;
end;

function RankUnitsByUses(const ACandidates, AUses: TArray<string>): TArray<Integer>;

  function Matches(const ACand, AUse: string): Boolean;
  begin
    // 'SysUtils' in a uses clause means 'System.SysUtils' (unit scopes)
    Result := SameText(ACand, AUse) or EndsText('.' + AUse, ACand);
  end;

var
  Taken: TArray<Boolean>;
  I, J: Integer;
begin
  Result := nil;
  SetLength(Taken, Length(ACandidates));
  // the LAST unit in the uses clause wins in Delphi
  for J := High(AUses) downto 0 do
    for I := 0 to High(ACandidates) do
      if not Taken[I] and Matches(ACandidates[I], AUses[J]) then
      begin
        Result := Result + [I];
        Taken[I] := True;
      end;
  for I := 0 to High(ACandidates) do
    if not Taken[I] then
      Result := Result + [I];
end;

function HoverDeclUri(const AHover: string): string;
var
  P, Q, H: Integer;
begin
  Result := '';
  P := Pos('](', AHover);
  if P = 0 then Exit;
  Q := Pos(')', AHover, P + 2);
  if Q = 0 then Exit;
  Result := Copy(AHover, P + 2, Q - P - 2);
  H := Pos('#', Result);
  if H > 0 then Result := Copy(Result, 1, H - 1);
  if not StartsText('file:', Result) then Result := '';
end;

// A unit's text up to its 'implementation' keyword - what OTHER units see.
function InterfacePartOf(const AContent: string): string;
var
  Lines: TArray<string>;
begin
  Lines := AContent.Replace(#13, '').Split([#10]);
  for var I := 0 to High(Lines) do
    if SameText(Trim(Lines[I]), 'implementation') then
      Exit(string.Join(#10, Copy(Lines, 0, I)));
  Result := AContent;
end;

function MakeIndexTypeSource(const ALookup: TFunc<string, TArray<TFindUnitHit>>;
  const AReadFile: TFunc<string, string>;
  const ACurrentFile, ACurrentContent: string): TGenTypeSource;
begin
  Result :=
    function(const AIdent, AContextFile: string): TArray<TGenDecl>
    var
      IsCurrent: Boolean;
      CtxFull: string;
      D: TGenDecl;
      Paths, Names, UsesNames: TArray<string>;
    begin
      Result := nil;
      IsCurrent := (AContextFile = '') or SameText(AContextFile, ACurrentFile);
      if IsCurrent then
        CtxFull := ACurrentContent
      else if Assigned(AReadFile) then
        CtxFull := AReadFile(AContextFile)
      else
        CtxFull := '';

      // 1. the context unit itself: the current buffer as a whole (local
      //    types of its own implementation are visible there), any other
      //    unit only up to 'implementation'
      if CtxFull <> '' then
      begin
        if IsCurrent then
        begin
          D.Path := ACurrentFile;
          D.Content := CtxFull;
        end
        else
        begin
          D.Path := AContextFile;
          D.Content := InterfacePartOf(CtxFull);
        end;
        Result := [D];
      end;
      if not Assigned(ALookup) then Exit;

      // 2. every indexed declaration, ranked by the context's uses clause
      for var H in ALookup(AIdent) do
      begin
        if SameText(H.Path, ACurrentFile) or SameText(H.Path, AContextFile) then
          Continue;
        var Dup := False;
        for var X in Paths do
          if SameText(X, H.Path) then Dup := True;
        if Dup then Continue;
        Paths := Paths + [H.Path];
        Names := Names + [H.UnitName];
      end;
      for var E in TUsesGraphAnalyzer.ParseUsesEntries(CtxFull) do
        UsesNames := UsesNames + [E.UnitName];
      for var Idx in RankUnitsByUses(Names, UsesNames) do
      begin
        var Text := '';
        if Assigned(AReadFile) then Text := AReadFile(Paths[Idx]);
        if Text = '' then Continue;
        D.Path := Paths[Idx];
        D.Content := InterfacePartOf(Text);
        Result := Result + [D];
      end;
    end;
end;

// ---------------------------------------------------------------------------
//  Generating text
// ---------------------------------------------------------------------------

function ProcHeadText(const AInfo: TProcTypeInfo; const AName: string): string;
begin
  if AInfo.IsFunction then Result := 'function' else Result := 'procedure';
  if AName <> '' then Result := Result + ' ' + AName;
  if AInfo.Params <> '' then Result := Result + '(' + AInfo.Params + ')';
  if AInfo.IsFunction and (AInfo.ResultType <> '') then
    Result := Result + ': ' + AInfo.ResultType;
end;

function AnonymousMethodText(const AInfo: TProcTypeInfo; const AIndent: string;
  out ABodyLineOffset, ABodyCol: Integer): string;
begin
  Result := ProcHeadText(AInfo) + sLineBreak +
    AIndent + 'begin' + sLineBreak +
    AIndent + '  ' + sLineBreak +
    AIndent + 'end';
  ABodyLineOffset := 2;
  ABodyCol := Length(AIndent) + 2;
end;

function StripF(const S: string): string;
begin
  if (Length(S) >= 2) and CharInSet(S[1], ['F', 'f']) and S[2].IsUpper then
    Result := Copy(S, 2, MaxInt)
  else
    Result := S;
end;

function StripOn(const S: string): string;
begin
  if (Length(S) > 2) and StartsText('On', S) and S[3].IsUpper then
    Result := Copy(S, 3, MaxInt)
  else
    Result := S;
end;

function StripA(const S: string): string;
begin
  if (Length(S) >= 2) and (S[1] = 'A') and S[2].IsUpper then
    Result := Copy(S, 2, MaxInt)
  else
    Result := S;
end;

function SuggestHandlerName(const ACtx: TGenContext; const AParamName: string): string;
begin
  case ACtx.Site of
    gsAssign:
      if ACtx.QualifierName <> '' then
        Result := StripF(ACtx.QualifierName) + StripOn(StripF(ACtx.TargetName))
      else
        Result := 'Do' + StripOn(StripF(ACtx.TargetName));
  else
    if AParamName <> '' then
      Result := 'Handle' + StripOn(StripA(AParamName))
    else
      Result := 'HandleEvent';
  end;
  if not IsPascalIdentifier(Result) then Result := 'HandleEvent';
end;

function PlanEventHandler(const ALines: TArray<string>; ACaretLine0: Integer;
  const AName: string; const AInfo: TProcTypeInfo): TGenPlan;
var
  Content, Hdr, Rest, Qualified, Cls, ClassIndent, MemberIndent, Head: string;
  First, Last, ClsLine, EndLine, PrivLine, I, P: Integer;
begin
  Result := Default(TGenPlan);
  Content := string.Join(sLineBreak, ALines);
  if not IsPascalIdentifier(AName) then
  begin
    Result.Reason := Format('"%s" is not a valid identifier', [AName]);
    Exit;
  end;
  if not FindEnclosingRoutineRange(Content, ACaretLine0, First, Last) then
  begin
    Result.Reason := 'the caret is not inside a method';
    Exit;
  end;

  // The enclosing routine's qualified name: "procedure TForm1.Test(...)".
  Hdr := Trim(ALines[First]);
  if StartsText('class ', Hdr) then Hdr := Trim(Copy(Hdr, 7, MaxInt));
  P := Pos(' ', Hdr);
  Rest := Trim(Copy(Hdr, P + 1, MaxInt));
  I := 1;
  while (I <= Length(Rest)) and not CharInSet(Rest[I], ['(', ';', ':']) do Inc(I);
  Qualified := Trim(Copy(Rest, 1, I - 1));
  P := LastDelimiter('.', Qualified);
  if P = 0 then
  begin
    Result.Reason := 'the caret is not inside a METHOD - an "of object" ' +
      'event needs a method of a class';
    Exit;
  end;
  Qualified := Copy(Qualified, 1, P - 1);          // 'TForm1' / 'TOuter.TInner' / 'TFoo<T>'
  Cls := Qualified;
  P := LastDelimiter('.', Cls);
  if P > 0 then Cls := Copy(Cls, P + 1, MaxInt);
  P := Pos('<', Cls);
  if P > 0 then Cls := Copy(Cls, 1, P - 1);
  Result.ClassName := Cls;

  // The class declaration in this unit (not a forward declaration).
  ClsLine := -1;
  for I := 0 to High(ALines) do
  begin
    var T := Trim(ALines[I]);
    if not StartsText(Cls, T) then Continue;
    var R := Copy(T, Length(Cls) + 1, MaxInt);
    if (R <> '') and IsIdentChar(R[1]) then Continue;
    if StartsStr('<', R) and (Pos('>', R) > 0) then R := Copy(R, Pos('>', R) + 1, MaxInt);
    R := Trim(R);
    if not StartsStr('=', R) then Continue;
    R := LowerCase(Trim(Copy(R, 2, MaxInt)));
    if not StartsStr('class', R) then Continue;
    if (R = 'class;') or StartsStr('class of', R) then Continue;
    ClsLine := I;
    Break;
  end;
  if ClsLine < 0 then
  begin
    Result.Reason := Format('the declaration of %s is not in this unit', [Cls]);
    Exit;
  end;
  ClassIndent := LeadingSpaces(ALines[ClsLine]);
  EndLine := -1;
  for I := ClsLine + 1 to High(ALines) do
  begin
    var T := LowerCase(Trim(ALines[I]));
    if ((T = 'end;') or (T = 'end')) and
      (Length(LeadingSpaces(ALines[I])) <= Length(ClassIndent)) then
    begin
      EndLine := I;
      Break;
    end;
  end;
  if EndLine < 0 then
  begin
    Result.Reason := Format('the end of the declaration of %s was not found', [Cls]);
    Exit;
  end;
  if FindMemberDeclarationLine(Content, Cls, AName) >= 0 then
  begin
    Result.Reason := Format('%s already has a member named %s', [Cls, AName]);
    Exit;
  end;

  MemberIndent := ClassIndent + '  ';
  PrivLine := -1;
  for I := ClsLine + 1 to EndLine - 1 do
  begin
    var T := LowerCase(Trim(ALines[I]));
    if (T = 'private') or (T = 'strict private') then
    begin
      PrivLine := I;
      Break;
    end;
  end;

  Head := ProcHeadText(AInfo, AName) + ';';
  if PrivLine >= 0 then
  begin
    Result.DeclLine0 := PrivLine + 1;
    Result.DeclText := MemberIndent + Head + sLineBreak;
    Result.DeclLines := 1;
  end
  else
  begin
    Result.DeclLine0 := EndLine;
    Result.DeclText := ClassIndent + 'private' + sLineBreak + MemberIndent + Head + sLineBreak;
    Result.DeclLines := 2;
  end;

  if Last + 1 > High(ALines) then
  begin
    Result.Reason := 'no room after the current method';
    Exit;
  end;
  Result.ImplLine0 := Last + 1;
  Result.ImplText := sLineBreak +
    ProcHeadText(AInfo, Qualified + '.' + AName) + ';' + sLineBreak +
    'begin' + sLineBreak +
    '  ' + sLineBreak +
    'end;' + sLineBreak;
  Result.ImplBodyLine0 := Result.ImplLine0 + 3;
  Result.Ok := True;
end;

end.
