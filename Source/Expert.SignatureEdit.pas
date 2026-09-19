(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SignatureEdit;

// "Change method signature" (issue #11, suggestion 9 by Ian Branch), the
// pure text half: parameter lists, argument lists and how a call site or a
// header is rewritten for a new parameter list. The scan that finds the
// headers and calls (DelphiLSP-verified) lives in Expert.ChangeSignature.
//
// The argument splitter honours nested ( ) and [ ], string literals
// (also '' escapes), comments, and generic type arguments - '<' counts as
// a bracket only when it directly follows an identifier and a matching '>'
// comes before the argument ends, so "Foo(a < b, c > d)" stays two
// comparisons while "Foo(TDictionary<string, Integer>.Create)" is ONE
// argument.

interface

uses
  System.SysUtils, System.Types;

type
  TSigParam = record
    Modifier: string;     // '', 'const', 'var', 'out', 'constref'
    Name: string;
    TypeText: string;     // '' for an untyped var/const/out parameter
    DefaultText: string;  // '' = no default value
  end;

  /// <summary>One parameter of the NEW list: where it comes from and, for a
  ///  new one, what existing calls pass.</summary>
  TNewParam = record
    Param: TSigParam;
    OldIndex: Integer;    // index in the OLD list, -1 = new parameter
    CallValue: string;    // new parameter: the argument existing calls get
                          // ('' = rely on its default value)
  end;

  TArgSpan = record
    Start, Stop: Integer; // 1-based, inclusive, INSIDE the parentheses
    Text: string;         // trimmed
  end;

  /// <summary>One routine header of the family (declaration, implementation,
  ///  interface method, override). Offsets are 1-based into the file's text
  ///  with #10 line breaks (SigJoinedText).</summary>
  TSigHeader = record
    FileIndex: Integer;       // into the TSigSource array
    Line: Integer;            // 0-based line of the NAME
    Caption: string;          // 'TFoo.Bar (declaration)'
    IsDecl: Boolean;          // declaration (default values live here)
    NameStart, NameEnd: Integer;
    Open, Close: Integer;     // the parameter list's parentheses, 0 = none
    Params: TArray<TSigParam>;
    HadDefaults: Boolean;
    BodyFirst, BodyLast: Integer; // implementation body lines, -1 = none
  end;

  TSigSource = record
    FilePath: string;
    Content: string;          // as read (hash check before applying)
  end;

  /// <summary>How a name is used at an occurrence.</summary>
  TCallContext = (
    ccArgs,        // Foo(...) - AOpen / AClose are the parentheses
    ccStatement,   // Foo;  Obj.Foo;  inherited Foo;  - a call without arguments
    ccExpression,  // if Foo then, X := Foo + 1 - a function call without arguments
    ccAccessor,    // property ... read Foo / write Foo / stored Foo
    ccReference);  // @Foo, X := Foo;  Bar(Foo) - a call OR a method
                   // reference, the text cannot tell

/// <summary>The parameters of a parameter list (the text between the
///  parentheses). "A, B: Integer" gives two parameters.</summary>
function ParseParamList(const AText: string): TArray<TSigParam>;

/// <summary>"const A: Integer; B: string = ''" - one parameter per group.
///  AWithDefaults: False leaves the default values out (implementation
///  headers that did not repeat them).</summary>
function FormatParamList(const AParams: TArray<TSigParam>; AWithDefaults: Boolean = True): string;

/// <summary>The parameters AParams as TNewParams that keep them unchanged.</summary>
function UnchangedSignature(const AParams: TArray<TSigParam>): TArray<TNewParam>;

/// <summary>'' when ANew is a valid Delphi parameter list for a change of
///  AOld, else the reason: empty / duplicate / invalid names, a new
///  parameter without a default AND without a call value, a parameter
///  without a default after one with a default, a default on var/out or an
///  untyped parameter, an old parameter used twice.</summary>
function ValidateSignatureChange(const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>): string;

/// <summary>True when existing parameters are removed or change their
///  relative order (what a bare "inherited;" would forward wrongly).</summary>
function ExistingOrderChanges(const AOld: TArray<TSigParam>; const ANew: TArray<TNewParam>): Boolean;

/// <summary>Old name -> new name for renamed parameters.</summary>
function RenamedParams(const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>): TArray<TArray<string>>;

/// <summary>1-based position of the ')' / ']' matching the bracket at
///  AOpen; strings and comments are skipped. 0 = none.</summary>
function MatchingBracket(const S: string; AOpen: Integer): Integer;

/// <summary>The arguments between AOpen ('(') and AClose (')').</summary>
function SplitArguments(const S: string; AOpen, AClose: Integer): TArray<TArgSpan>;

/// <summary>The argument text for a call that passed AOldArgs (possibly
///  fewer than AOld - the rest took their defaults). Positions that must be
///  filled explicitly get the old default value or the new call value;
///  trailing arguments that the new list covers by a default are left out.
///  ANote names arguments that were DROPPED although they contain a call
///  (side effects). False with AError when an argument cannot be supplied.</summary>
function RewriteArguments(const AOldArgs: TArray<string>; const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>; out AText, ANote, AError: string): Boolean;

/// <summary>How the name at ANameStart..ANameEnd (1-based, inclusive) is
///  used. AMasked is the text with comments and strings BLANKED
///  (MaskCommentsAndStrings, lines joined with #10) - positions equal those
///  of the real text. For ccArgs, AOpen / AClose are the parentheses
///  (generic arguments "Foo<T>(" are skipped).</summary>
function CallContextAt(const AMasked: string; ANameStart, ANameEnd: Integer;
  out AOpen, AClose: Integer): TCallContext;

type
  TSigCall = record
    FileIndex: Integer;
    Line, Col: Integer;       // 0-based position of the name
    NameStart, NameEnd: Integer;
    Context: TCallContext;
    Open, Close: Integer;     // ccArgs: the parentheses
    Text: string;             // the source line (trimmed)
  end;

  /// <summary>Replace Start..Stop (1-based, inclusive; Stop = Start - 1 is
  ///  an insertion before Start) in the file's SigJoinedText.</summary>
  TSigEdit = record
    FileIndex: Integer;
    Start, Stop: Integer;
    NewText: string;
    Line: Integer;            // 0-based, for the preview
    What: string;             // 'header', 'call', 'body'
  end;

  TSigPlan = record
    Errors: TArray<string>;   // blocking
    Warnings: TArray<string>;
    Edits: TArray<TSigEdit>;
    function Ok: Boolean;
  end;

/// <summary>AContent with every line break turned into #10 - the text all
///  TSigHeader / TSigCall / TSigEdit offsets refer to.</summary>
function SigJoinedText(const AContent: string): string;

/// <summary>1-based offset of (ALine0, ACol0) in SigJoinedText(AContent).</summary>
function SigOffsetOf(const AContent: string; ALine0, ACol0: Integer): Integer;

/// <summary>The routine header on ALine0 declaring AName: name position,
///  parameter list and parameters; for an implementation (not AIsDecl, or
///  AWithBody for a routine that exists only in the implementation) also
///  the body lines. False with AWhy when the line is no such header or the
///  parameter list holds a comment / directive (not rewritten
///  automatically).</summary>
function LocateSigHeader(const AContent: string; ALine0: Integer; const AName: string;
  AIsDecl: Boolean; out AHeader: TSigHeader; out AWhy: string;
  AWithBody: Boolean = False): Boolean;

/// <summary>The occurrences of a name of length ANameLen at APositions
///  (X = column, Y = line, both 0-based), classified - one masking pass for
///  the whole file.</summary>
function LocateSigCalls(const AContent: string; const APositions: TArray<TPoint>;
  ANameLen, AFileIndex: Integer): TArray<TSigCall>;

/// <summary>The occurrence of the name at (ALine0, ACol0), classified.</summary>
function LocateSigCall(const AContent: string; ALine0, ACol0, ANameLen: Integer;
  out ACall: TSigCall): Boolean;

/// <summary>Every edit for the new parameter list ANew of the family
///  AHeaders (AHeaders[0] = the declaration the user edited) and the
///  verified occurrences ACalls. Pure.</summary>
function PlanSignatureEdits(const ASources: TArray<TSigSource>;
  const AHeaders: TArray<TSigHeader>; const ACalls: TArray<TSigCall>;
  const ANew: TArray<TNewParam>): TSigPlan;

/// <summary>SigJoinedText(AContent) with the edits of AFileIndex applied.</summary>
function ApplySigEdits(const AContent: string; const AEdits: TArray<TSigEdit>;
  AFileIndex: Integer): string;

/// <summary>ALines[AFirst..ALast]: every CODE occurrence of a name in
///  APairs[i][0] as a bare identifier (not after a '.') renamed to
///  APairs[i][1] - in ONE pass, so swapping two names works.</summary>
function RenameIdentifiersInRange(const ALines: TArray<string>; AFirst, ALast: Integer;
  const APairs: TArray<TArray<string>>; out ACount: Integer): TArray<string>;

/// <summary>ALines[AFirst..ALast] with every CODE occurrence of AOld as a
///  bare identifier (not after a '.') renamed to ANew. Used for renamed
///  parameters inside a method body. ACount = replacements.</summary>
function RenameIdentifierInRange(const ALines: TArray<string>; AFirst, ALast: Integer;
  const AOld, ANew: string; out ACount: Integer): TArray<string>;

implementation

uses
  System.StrUtils, System.Math, System.Generics.Collections, System.Generics.Defaults,
  Expert.PascalScanner, Expert.UnitIndex;

function ParseParamList(const AText: string): TArray<TSigParam>;
var
  Groups: TArray<string>;
  Depth: Integer;
  Cur: string;
begin
  Result := nil;
  // groups are separated by ';' at depth 0 (default values may contain
  // brackets and strings)
  Groups := nil;
  Cur := '';
  Depth := 0;
  var InStr := False;
  for var I := 1 to Length(AText) do
  begin
    var C := AText[I];
    if C = '''' then InStr := not InStr;
    if not InStr then
      case C of
        '(', '[': Inc(Depth);
        ')', ']': Dec(Depth);
      end;
    if (C = ';') and (Depth = 0) and not InStr then
    begin
      Groups := Groups + [Cur];
      Cur := '';
    end
    else
      Cur := Cur + C;
  end;
  if Trim(Cur) <> '' then Groups := Groups + [Cur];
  for var G in Groups do
  begin
    var T := Trim(G);
    if T = '' then Continue;
    var Modifier := '';
    // attributes like [ref] are kept with the modifier
    var Attr := '';
    while T.StartsWith('[') do
    begin
      var E := Pos(']', T);
      if E = 0 then Break;
      Attr := Attr + Copy(T, 1, E) + ' ';
      T := TrimLeft(Copy(T, E + 1, MaxInt));
    end;
    for var M in ['constref', 'const', 'var', 'out'] do
      if T.ToLower.StartsWith(M + ' ') then
      begin
        Modifier := M;
        T := TrimLeft(Copy(T, Length(M) + 1, MaxInt));
        Break;
      end;
    Modifier := Trim(Attr + Modifier);
    var Names := T;
    var TypeText := '';
    var Def := '';
    // "A, B: Integer = 5": the first ':' outside strings separates the names
    var Colon := Pos(':', T);
    if Colon > 0 then
    begin
      Names := Copy(T, 1, Colon - 1);
      var Rest := Copy(T, Colon + 1, MaxInt);
      // the default: the first '=' at depth 0 outside strings
      var D := 0;
      var S := False;
      var EqPos := 0;
      for var K := 1 to Length(Rest) do
      begin
        var C := Rest[K];
        if C = '''' then S := not S;
        if S then Continue;
        case C of
          '(', '[', '<': Inc(D);
          ')', ']', '>': Dec(D);
          '=': if D = 0 then begin EqPos := K; Break; end;
        end;
      end;
      if EqPos > 0 then
      begin
        TypeText := Trim(Copy(Rest, 1, EqPos - 1));
        Def := Trim(Copy(Rest, EqPos + 1, MaxInt));
      end
      else
        TypeText := Trim(Rest);
    end;
    for var N in Names.Split([',']) do
    begin
      var P := Default(TSigParam);
      P.Modifier := Modifier;
      P.Name := Trim(N);
      P.TypeText := TypeText;
      P.DefaultText := Def;
      if P.Name <> '' then Result := Result + [P];
    end;
  end;
end;

function FormatParamList(const AParams: TArray<TSigParam>; AWithDefaults: Boolean): string;
begin
  Result := '';
  for var P in AParams do
  begin
    var S := '';
    if P.Modifier <> '' then S := P.Modifier + ' ';
    S := S + P.Name;
    if P.TypeText <> '' then S := S + ': ' + P.TypeText;
    if AWithDefaults and (P.DefaultText <> '') then S := S + ' = ' + P.DefaultText;
    if Result <> '' then Result := Result + '; ';
    Result := Result + S;
  end;
end;

function UnchangedSignature(const AParams: TArray<TSigParam>): TArray<TNewParam>;
begin
  SetLength(Result, Length(AParams));
  for var I := 0 to High(AParams) do
  begin
    Result[I] := Default(TNewParam);
    Result[I].Param := AParams[I];
    Result[I].OldIndex := I;
  end;
end;

function ValidateSignatureChange(const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>): string;
var
  Seen, SeenOld: TDictionary<string, Boolean>;
begin
  Result := '';
  Seen := TDictionary<string, Boolean>.Create;
  SeenOld := TDictionary<string, Boolean>.Create;
  try
    var HadDefault := False;
    for var I := 0 to High(ANew) do
    begin
      var P := ANew[I].Param;
      if not IsIdentifier(P.Name) then
        Exit(Format('parameter %d: "%s" is not a valid name', [I + 1, P.Name]));
      if Seen.ContainsKey(UpperCase(P.Name)) then
        Exit('parameter "' + P.Name + '" appears twice');
      Seen.Add(UpperCase(P.Name), True);
      if (ANew[I].OldIndex < -1) or (ANew[I].OldIndex > High(AOld)) then
        Exit('parameter "' + P.Name + '": invalid source parameter');
      if ANew[I].OldIndex >= 0 then
      begin
        if SeenOld.ContainsKey(IntToStr(ANew[I].OldIndex)) then
          Exit('old parameter "' + AOld[ANew[I].OldIndex].Name + '" is used twice');
        SeenOld.Add(IntToStr(ANew[I].OldIndex), True);
      end;
      var M := LowerCase(P.Modifier);
      if (P.TypeText = '') and not ((M = 'var') or (M = 'const') or (M = 'out') or
         M.EndsWith(' var') or M.EndsWith(' const') or M.EndsWith(' out')) then
        Exit('parameter "' + P.Name + '" needs a type');
      if P.DefaultText <> '' then
      begin
        if (M = 'var') or (M = 'out') or M.EndsWith(' var') or M.EndsWith(' out') then
          Exit('parameter "' + P.Name + '": a var/out parameter cannot have a default value');
        if P.TypeText = '' then
          Exit('parameter "' + P.Name + '": an untyped parameter cannot have a default value');
        HadDefault := True;
      end
      else if HadDefault then
        Exit('parameter "' + P.Name + '" has no default value but follows one that has - ' +
          'default values must come last');
      if (ANew[I].OldIndex < 0) and (ANew[I].CallValue = '') and (P.DefaultText = '') then
        Exit('new parameter "' + P.Name + '" needs a default value or a value for the ' +
          'existing calls');
      if (ANew[I].OldIndex < 0) and ((M = 'var') or (M = 'out')) and (ANew[I].CallValue <> '') and
         not IsIdentifier(ANew[I].CallValue) then
        Exit('new var/out parameter "' + P.Name + '": the value for existing calls must be ' +
          'a variable');
    end;
  finally
    SeenOld.Free;
    Seen.Free;
  end;
end;

function ExistingOrderChanges(const AOld: TArray<TSigParam>; const ANew: TArray<TNewParam>): Boolean;
begin
  var Last := -1;
  var Kept := 0;
  for var N in ANew do
    if N.OldIndex >= 0 then
    begin
      if N.OldIndex < Last then Exit(True);
      Last := N.OldIndex;
      Inc(Kept);
    end;
  Result := Kept <> Length(AOld);
end;

function RenamedParams(const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>): TArray<TArray<string>>;
begin
  Result := nil;
  for var N in ANew do
    if (N.OldIndex >= 0) and not SameText(AOld[N.OldIndex].Name, N.Param.Name) then
      Result := Result + [[AOld[N.OldIndex].Name, N.Param.Name]];
end;

// ---------------------------------------------------------------------------

// Skips a string literal / comment starting at I; returns the index after
// it, or I when there is none.
function SkipNonCode(const S: string; I: Integer): Integer;
var
  N: Integer;
begin
  N := Length(S);
  Result := I;
  if I > N then Exit;
  case S[I] of
    '''':
      begin
        var K := I + 1;
        while K <= N do
        begin
          if S[K] = '''' then
          begin
            if (K < N) and (S[K + 1] = '''') then Inc(K, 2)
            else Exit(K + 1);
          end
          else
            Inc(K);
        end;
        Exit(N + 1);
      end;
    '{':
      begin
        var E := PosEx('}', S, I + 1);
        if E = 0 then Exit(N + 1);
        Exit(E + 1);
      end;
    '(':
      if (I < N) and (S[I + 1] = '*') then
      begin
        var E := PosEx('*)', S, I + 2);
        if E = 0 then Exit(N + 1);
        Exit(E + 2);
      end;
    '/':
      if (I < N) and (S[I + 1] = '/') then
      begin
        var K := I;
        while (K <= N) and not CharInSet(S[K], [#10, #13]) do Inc(K);
        Exit(K);
      end;
  end;
end;

function MatchingBracket(const S: string; AOpen: Integer): Integer;
var
  Stack: string;
  I, N: Integer;
begin
  Result := 0;
  N := Length(S);
  if (AOpen < 1) or (AOpen > N) or not CharInSet(S[AOpen], ['(', '[']) then Exit;
  Stack := '';
  I := AOpen;
  while I <= N do
  begin
    var J := SkipNonCode(S, I);
    if J <> I then
    begin
      I := J;
      Continue;
    end;
    case S[I] of
      '(', '[': Stack := Stack + S[I];
      ')', ']':
        begin
          if Stack = '' then Exit;
          SetLength(Stack, Length(Stack) - 1);
          if Stack = '' then Exit(I);
        end;
    end;
    Inc(I);
  end;
end;

// '<' at APos starts GENERIC arguments: it follows an identifier directly,
// and a matching '>' comes before the argument ends (',' / ')' at depth 0),
// with only type-ish characters in between.
function GenericClose(const S: string; APos, ALimit: Integer): Integer;
begin
  Result := 0;
  if (APos <= 1) or not IsIdentChar(S[APos - 1]) then Exit;
  var Depth := 0;
  for var K := APos to ALimit do
  begin
    var C := S[K];
    case C of
      '<': Inc(Depth);
      '>':
        begin
          Dec(Depth);
          if Depth = 0 then Exit(K);
        end;
      ',', ' ', '.', '_', #9: ;
    else
      if not IsIdentChar(C) then Exit(0);   // an operator: a comparison
    end;
  end;
end;

function SplitArguments(const S: string; AOpen, AClose: Integer): TArray<TArgSpan>;
var
  Depth, I, Start: Integer;

  procedure Emit(AStop: Integer);
  var
    A: TArgSpan;
  begin
    A.Start := Start;
    A.Stop := AStop;
    A.Text := Trim(Copy(S, Start, AStop - Start + 1));
    Result := Result + [A];
  end;

begin
  Result := nil;
  if (AOpen < 1) or (AClose <= AOpen) then Exit;
  if Trim(Copy(S, AOpen + 1, AClose - AOpen - 1)) = '' then Exit;   // "()"
  Depth := 0;
  Start := AOpen + 1;
  I := AOpen + 1;
  while I < AClose do
  begin
    var J := SkipNonCode(S, I);
    if J <> I then
    begin
      I := J;
      Continue;
    end;
    case S[I] of
      '(', '[': Inc(Depth);
      ')', ']': Dec(Depth);
      '<':
        begin
          var G := GenericClose(S, I, AClose - 1);
          if G > 0 then
          begin
            I := G + 1;
            Continue;
          end;
        end;
      ',':
        if Depth = 0 then
        begin
          Emit(I - 1);
          Start := I + 1;
        end;
    end;
    Inc(I);
  end;
  Emit(AClose - 1);
end;

function RewriteArguments(const AOldArgs: TArray<string>; const AOld: TArray<TSigParam>;
  const ANew: TArray<TNewParam>; out AText, ANote, AError: string): Boolean;
var
  Vals: TArray<string>;
  Absent: TArray<Boolean>;
begin
  Result := False;
  AText := '';
  ANote := '';
  AError := '';
  if Length(AOldArgs) > Length(AOld) then
  begin
    AError := Format('the call passes %d argument(s), the method has %d parameter(s)',
      [Length(AOldArgs), Length(AOld)]);
    Exit;
  end;
  SetLength(Vals, Length(ANew));
  SetLength(Absent, Length(ANew));
  for var K := 0 to High(ANew) do
  begin
    var O := ANew[K].OldIndex;
    if O >= 0 then
    begin
      if O < Length(AOldArgs) then
        Vals[K] := AOldArgs[O]
      else
      begin
        Absent[K] := True;
        Vals[K] := AOld[O].DefaultText;   // what the call relied on
      end;
    end
    else if ANew[K].CallValue <> '' then
      Vals[K] := ANew[K].CallValue
    else
    begin
      Absent[K] := True;
      Vals[K] := ANew[K].Param.DefaultText;
    end;
  end;
  // trailing absent arguments stay away only when the NEW parameter has a
  // default; everything before the last written one must be explicit
  var Last := High(ANew);
  while (Last >= 0) and Absent[Last] and (ANew[Last].Param.DefaultText <> '') and
    ((ANew[Last].OldIndex < 0) or
     (Trim(ANew[Last].Param.DefaultText) = Trim(AOld[ANew[Last].OldIndex].DefaultText))) do
    Dec(Last);
  var Parts: TArray<string> := nil;
  for var K := 0 to Last do
  begin
    if Vals[K] = '' then
    begin
      AError := Format('no value for parameter "%s" at this call', [ANew[K].Param.Name]);
      Exit;
    end;
    Parts := Parts + [Vals[K]];
  end;
  AText := string.Join(', ', Parts);
  // dropped arguments that contain a call may have had side effects
  for var O := 0 to High(AOldArgs) do
  begin
    var Kept := False;
    for var N in ANew do
      if N.OldIndex = O then Kept := True;
    if not Kept and (Pos('(', AOldArgs[O]) > 0) then
      ANote := ANote + IfThen(ANote <> '', '; ', '') + 'dropped argument with a call: ' + AOldArgs[O];
  end;
  Result := True;
end;

// masked text: only blanks and line breaks need skipping
function NextCode(const S: string; I: Integer): Integer;
begin
  Result := I;
  while (Result <= Length(S)) and CharInSet(S[Result], [' ', #9, #10, #13]) do Inc(Result);
end;

function PrevCode(const S: string; I: Integer): Integer;
begin
  Result := I - 1;
  while (Result >= 1) and CharInSet(S[Result], [' ', #9, #10, #13]) do Dec(Result);
end;

function WordEndingAt(const S: string; AEnd: Integer): string;
begin
  var B := AEnd;
  while (B >= 1) and IsIdentChar(S[B]) do Dec(B);
  Result := UpperCase(Copy(S, B + 1, AEnd - B));
end;

function WordStartingAt(const S: string; AStart: Integer): string;
begin
  var E := AStart;
  while (E <= Length(S)) and IsIdentChar(S[E]) do Inc(E);
  Result := UpperCase(Copy(S, AStart, E - AStart));
end;

function CallContextAt(const AMasked: string; ANameStart, ANameEnd: Integer;
  out AOpen, AClose: Integer): TCallContext;
const
  StatementStarters: array[0..9] of string = ('BEGIN', 'THEN', 'ELSE', 'DO', 'TRY',
    'FINALLY', 'EXCEPT', 'REPEAT', 'OF', 'INITIALIZATION');
  StatementEnders: array[0..5] of string = ('END', 'ELSE', 'UNTIL', 'EXCEPT', 'FINALLY',
    'FINALIZATION');
var
  S: string;
begin
  S := AMasked;
  AOpen := 0;
  AClose := 0;
  var N := NextCode(S, ANameEnd + 1);
  // generic arguments: Foo<T>(...)
  if (N <= Length(S)) and (S[N] = '<') and (N = ANameEnd + 1) then
  begin
    var G := GenericClose(S, N, Length(S));
    if G > 0 then N := NextCode(S, G + 1);
  end;
  if (N <= Length(S)) and (S[N] = '(') then
  begin
    var C := MatchingBracket(S, N);
    if C > 0 then
    begin
      AOpen := N;
      AClose := C;
      Exit(ccArgs);
    end;
  end;
  // what stands before the (qualified) name: back over "X.", "X[..].",
  // "X(..)." chains and the inherited keyword
  var P := PrevCode(S, ANameStart);
  while (P >= 1) and (S[P] = '.') do
  begin
    P := PrevCode(S, P);
    while (P >= 1) and CharInSet(S[P], [')', ']']) do
    begin
      var Depth := 0;
      while P >= 1 do
      begin
        if CharInSet(S[P], [')', ']']) then Inc(Depth)
        else if CharInSet(S[P], ['(', '[']) then
        begin
          Dec(Depth);
          if Depth = 0 then Break;
        end;
        Dec(P);
      end;
      P := PrevCode(S, P);
    end;
    while (P >= 1) and IsIdentChar(S[P]) do Dec(P);
    P := PrevCode(S, P + 1);
  end;
  var W := WordEndingAt(S, P);
  if W = 'INHERITED' then
  begin
    P := PrevCode(S, P - Length(W) + 1);
    W := WordEndingAt(S, P);
  end;
  if (W = 'READ') or (W = 'WRITE') or (W = 'STORED') then Exit(ccAccessor);
  if (P >= 1) and (S[P] = '@') then Exit(ccReference);
  var Starts := (P < 1) or (S[P] = ';') or
    ((S[P] = ':') and ((P + 1 > Length(S)) or (S[P + 1] <> '=')));
  for var K in StatementStarters do
    if W = K then Starts := True;
  var W2 := '';
  if N <= Length(S) then W2 := WordStartingAt(S, N);
  var Ends := (N > Length(S)) or (S[N] = ';');
  for var K in StatementEnders do
    if W2 = K then Ends := True;
  if Starts and Ends then Exit(ccStatement);
  // "X := Foo;" and a lone argument "Bar(Foo)" / "Bar(A, Foo)" may be a
  // method reference
  var Assigned := (P >= 2) and (S[P] = '=') and (S[P - 1] = ':');
  if Assigned and Ends then Exit(ccReference);
  if (P >= 1) and CharInSet(S[P], ['(', ',']) and (N <= Length(S)) and
     CharInSet(S[N], [')', ',']) then
    Exit(ccReference);
  Result := ccExpression;
end;

type
  TWordHit = record
    Line, Col: Integer;   // 0-based
    Len: Integer;
    NewText: string;
  end;

// the CODE words of ALines[AFirst..ALast] found in APairs (not after '.'
// or '&'), left to right
function RenameHits(const ALines: TArray<string>; AFirst, ALast: Integer;
  const APairs: TArray<TArray<string>>): TArray<TWordHit>;
var
  M: TArray<string>;
  Map: TDictionary<string, string>;
begin
  Result := nil;
  if Length(APairs) = 0 then Exit;
  Map := TDictionary<string, string>.Create;
  try
    for var P in APairs do
      if (Length(P) = 2) and (P[0] <> '') and (P[0] <> P[1]) then
        Map.AddOrSetValue(UpperCase(P[0]), P[1]);
    if Map.Count = 0 then Exit;
    M := MaskCommentsAndStrings(ALines);
    for var L := Max(AFirst, 0) to Min(ALast, High(M)) do
    begin
      var S := M[L];
      var I := 1;
      while I <= Length(S) do
      begin
        if IsIdentStart(S[I]) and ((I = 1) or not IsIdentChar(S[I - 1])) then
        begin
          var E := I;
          while (E < Length(S)) and IsIdentChar(S[E + 1]) do Inc(E);
          var NewName: string;
          if Map.TryGetValue(UpperCase(Copy(S, I, E - I + 1)), NewName) then
          begin
            var B := I - 1;
            while (B >= 1) and CharInSet(S[B], [' ', #9]) do Dec(B);
            if not ((B >= 1) and CharInSet(S[B], ['.', '&'])) then
            begin
              var H: TWordHit;
              H.Line := L; H.Col := I - 1; H.Len := E - I + 1; H.NewText := NewName;
              Result := Result + [H];
            end;
          end;
          I := E + 1;
        end
        else
          Inc(I);
      end;
    end;
  finally
    Map.Free;
  end;
end;

function RenameIdentifiersInRange(const ALines: TArray<string>; AFirst, ALast: Integer;
  const APairs: TArray<TArray<string>>; out ACount: Integer): TArray<string>;
begin
  Result := Copy(ALines);
  var Hits := RenameHits(ALines, AFirst, ALast, APairs);
  ACount := Length(Hits);
  // right to left (the hits come left to right, line by line)
  for var K := High(Hits) downto 0 do
    Result[Hits[K].Line] := Copy(Result[Hits[K].Line], 1, Hits[K].Col) + Hits[K].NewText +
      Copy(Result[Hits[K].Line], Hits[K].Col + Hits[K].Len + 1, MaxInt);
end;

function RenameIdentifierInRange(const ALines: TArray<string>; AFirst, ALast: Integer;
  const AOld, ANew: string; out ACount: Integer): TArray<string>;
begin
  if SameText(AOld, ANew) then
  begin
    ACount := 0;
    Exit(Copy(ALines));
  end;
  Result := RenameIdentifiersInRange(ALines, AFirst, ALast, [[AOld, ANew]], ACount);
end;

// ---------------------------------------------------------------------------
//  Headers, occurrences, the plan
// ---------------------------------------------------------------------------

function SigJoinedText(const AContent: string): string;
begin
  Result := AContent.Replace(#13#10, #10).Replace(#13, #10);
end;

function LineStartsOf(const AJoined: string): TArray<Integer>;
begin
  var N := 1;
  for var I := 1 to Length(AJoined) do
    if AJoined[I] = #10 then Inc(N);
  SetLength(Result, N);
  Result[0] := 1;
  N := 1;
  for var I := 1 to Length(AJoined) do
    if AJoined[I] = #10 then
    begin
      Result[N] := I + 1;
      Inc(N);
    end;
end;

function LineOfOffset(const AStarts: TArray<Integer>; AOffset: Integer): Integer;
begin
  Result := 0;
  var Lo := 0;
  var Hi := High(AStarts);
  while Lo <= Hi do
  begin
    var Mid := (Lo + Hi) div 2;
    if AStarts[Mid] <= AOffset then
    begin
      Result := Mid;
      Lo := Mid + 1;
    end
    else
      Hi := Mid - 1;
  end;
end;

function SigOffsetOf(const AContent: string; ALine0, ACol0: Integer): Integer;
begin
  var S := LineStartsOf(SigJoinedText(AContent));
  if (ALine0 < 0) or (ALine0 > High(S)) then Exit(0);
  Result := S[ALine0] + ACol0;
end;

function MaskedJoined(const ALines: TArray<string>): string;
begin
  Result := string.Join(#10, MaskCommentsAndStrings(ALines));
end;

function LocateSigHeader(const AContent: string; ALine0: Integer; const AName: string;
  AIsDecl: Boolean; out AHeader: TSigHeader; out AWhy: string; AWithBody: Boolean): Boolean;
var
  J, MJ, S, U, W: string;
  Lines: TArray<string>;
  Starts: TArray<Integer>;
begin
  Result := False;
  AWhy := '';
  AHeader := Default(TSigHeader);
  AHeader.Line := ALine0;
  AHeader.IsDecl := AIsDecl;
  AHeader.BodyFirst := -1;
  AHeader.BodyLast := -1;
  J := SigJoinedText(AContent);
  Lines := J.Split([#10]);
  if (ALine0 < 0) or (ALine0 > High(Lines)) then
  begin
    AWhy := Format('line %d does not exist', [ALine0 + 1]);
    Exit;
  end;
  MJ := MaskedJoined(Lines);
  Starts := LineStartsOf(J);
  S := Copy(MJ, Starts[ALine0], Length(Lines[ALine0]));
  U := UpperCase(TrimLeft(S));
  if U.StartsWith('CLASS ') then U := TrimLeft(Copy(U, 7, MaxInt));
  var IsHdr := False;
  for var K in ['PROCEDURE ', 'FUNCTION ', 'CONSTRUCTOR ', 'DESTRUCTOR '] do
    if U.StartsWith(K) then IsHdr := True;
  if not IsHdr then
  begin
    AWhy := Format('line %d is no routine header', [ALine0 + 1]);
    Exit;
  end;
  // the name: after a '.' (qualified implementation header) wins, else the
  // first whole word after the keyword
  W := UpperCase(AName);
  U := UpperCase(S);
  var Col := 0;
  var P := Pos(W, U);
  while P > 0 do
  begin
    var E := P + Length(W);
    if ((P = 1) or not IsIdentChar(U[P - 1])) and ((E > Length(U)) or not IsIdentChar(U[E])) then
    begin
      var B := P - 1;
      while (B >= 1) and CharInSet(U[B], [' ', #9]) do Dec(B);
      if (B >= 1) and (U[B] = '.') then
      begin
        Col := P;
        Break;
      end;
      if (Col = 0) and (B >= 1) and IsIdentChar(U[B]) then Col := P;
    end;
    P := Pos(W, U, P + 1);
  end;
  if Col = 0 then
  begin
    AWhy := Format('line %d does not declare %s', [ALine0 + 1, AName]);
    Exit;
  end;
  AHeader.NameStart := Starts[ALine0] + Col - 1;
  var N := AHeader.NameStart + Length(AName);
  // generic parameters of the method ("Foo<T: class>") - constraints may
  // hold ':' and ',', so plain angle-bracket counting
  if (N <= Length(MJ)) and (MJ[N] = '<') then
  begin
    var Depth := 0;
    var K := N;
    while K <= Length(MJ) do
    begin
      if MJ[K] = '<' then Inc(Depth)
      else if MJ[K] = '>' then
      begin
        Dec(Depth);
        if Depth = 0 then Break;
      end
      else if CharInSet(MJ[K], ['(', ';']) then Break;
      Inc(K);
    end;
    if (K <= Length(MJ)) and (MJ[K] = '>') then N := K + 1;
  end;
  AHeader.NameEnd := N - 1;
  N := NextCode(MJ, N);
  var BodyFrom := ALine0;
  if (N <= Length(MJ)) and (MJ[N] = '(') then
  begin
    var C := MatchingBracket(MJ, N);
    if C = 0 then
    begin
      AWhy := Format('the parameter list on line %d is not closed', [ALine0 + 1]);
      Exit;
    end;
    AHeader.Open := N;
    AHeader.Close := C;
    var PT := Copy(J, N + 1, C - N - 1);
    var Sc := TPascalScanner.Create(PT, True, True);
    try
      var T: TPasToken;
      while Sc.Next(T) do
        if T.Kind in [ptComment, ptDirective] then
        begin
          AWhy := Format('the parameter list on line %d contains a comment or a compiler ' +
            'directive - change it by hand', [ALine0 + 1]);
          Exit;
        end;
    finally
      Sc.Free;
    end;
    AHeader.Params := ParseParamList(PT);
    for var Prm in AHeader.Params do
      if Prm.DefaultText <> '' then AHeader.HadDefaults := True;
    BodyFrom := LineOfOffset(Starts, C);
  end;
  if (not AIsDecl) or AWithBody then
  begin
    var F, L: Integer;
    if FindEnclosingRoutineRange(J, ALine0, F, L) and (F <= ALine0) and (L > BodyFrom) then
    begin
      AHeader.BodyFirst := BodyFrom + 1;
      AHeader.BodyLast := L;
    end;
  end;
  Result := True;
end;

function LocateSigCalls(const AContent: string; const APositions: TArray<TPoint>;
  ANameLen, AFileIndex: Integer): TArray<TSigCall>;
var
  J, MJ: string;
  Lines: TArray<string>;
  Starts: TArray<Integer>;
begin
  Result := nil;
  J := SigJoinedText(AContent);
  Lines := J.Split([#10]);
  MJ := MaskedJoined(Lines);
  Starts := LineStartsOf(J);
  for var P in APositions do
  begin
    if (P.Y < 0) or (P.Y > High(Lines)) then Continue;
    var C := Default(TSigCall);
    C.FileIndex := AFileIndex;
    C.Line := P.Y;
    C.Col := P.X;
    C.NameStart := Starts[P.Y] + P.X;
    C.NameEnd := C.NameStart + ANameLen - 1;
    C.Context := CallContextAt(MJ, C.NameStart, C.NameEnd, C.Open, C.Close);
    C.Text := Trim(Lines[P.Y]);
    Result := Result + [C];
  end;
end;

function LocateSigCall(const AContent: string; ALine0, ACol0, ANameLen: Integer;
  out ACall: TSigCall): Boolean;
begin
  var R := LocateSigCalls(AContent, [Point(ACol0, ALine0)], ANameLen, 0);
  Result := Length(R) = 1;
  if Result then ACall := R[0] else ACall := Default(TSigCall);
end;

function TSigPlan.Ok: Boolean;
begin
  Result := Length(Errors) = 0;
end;

function TypeKey(const S: string): string;
begin
  Result := UpperCase(S.Replace(' ', '').Replace(#9, ''));
end;

function PlanSignatureEdits(const ASources: TArray<TSigSource>;
  const AHeaders: TArray<TSigHeader>; const ACalls: TArray<TSigCall>;
  const ANew: TArray<TNewParam>): TSigPlan;
var
  Plan: TSigPlan;
  Origin: TArray<TSigParam>;
  Joined: TArray<string>;
  LinesOf: TArray<TArray<string>>;
  StartsOf: TArray<TArray<Integer>>;
  BodyPairs: TArray<TArray<TArray<string>>>;
  BodyNames: TArray<TArray<string>>;   // per header: the NEW parameter names
  CallSpans: TArray<TSigEdit>;

  procedure Err(const S: string);
  begin
    Plan.Errors := Plan.Errors + [S];
  end;

  procedure Warn(const S: string);
  begin
    Plan.Warnings := Plan.Warnings + [S];
  end;

  procedure AddEdit(AFile, AStart, AStop: Integer; const AText: string; ALine: Integer;
    const AWhat: string);
  var
    E: TSigEdit;
  begin
    E.FileIndex := AFile;
    E.Start := AStart;
    E.Stop := AStop;
    E.NewText := AText;
    E.Line := ALine;
    E.What := AWhat;
    Plan.Edits := Plan.Edits + [E];
  end;

  function Loc(AFile, ALine: Integer): string;
  begin
    Result := Format('%s(%d)', [ExtractFileName(ASources[AFile].FilePath), ALine + 1]);
  end;

  // "inherited Foo(...)" inside a family body hands the NEW parameters on
  // by name - neither the value for existing calls nor the default is
  // what the override received
  function CallParams(const ACall: TSigCall): TArray<TNewParam>;
  begin
    Result := ANew;
    var J := Joined[ACall.FileIndex];
    var P := ACall.NameStart - 1;
    while (P >= 1) and CharInSet(J[P], [' ', #9, #10]) do Dec(P);
    if not ((P >= 9) and SameText(Copy(J, P - 8, 9), 'inherited') and
       ((P = 9) or not IsIdentChar(J[P - 9]))) then Exit;
    for var HI := 0 to High(AHeaders) do
      if (AHeaders[HI].FileIndex = ACall.FileIndex) and (AHeaders[HI].BodyFirst >= 0) and
         (ACall.Line >= AHeaders[HI].BodyFirst) and (ACall.Line <= AHeaders[HI].BodyLast) and
         (Length(BodyNames[HI]) = Length(ANew)) then
      begin
        Result := Copy(ANew);
        for var K := 0 to High(Result) do
          if Result[K].OldIndex < 0 then Result[K].CallValue := BodyNames[HI][K];
        Exit;
      end;
  end;

  function BodyUses(const AHeader: TSigHeader; const AName: string): Integer;
  begin
    Result := -1;
    var Hits := RenameHits(LinesOf[AHeader.FileIndex], AHeader.BodyFirst, AHeader.BodyLast,
      [[AName, AName + '_']]);
    if Length(Hits) > 0 then Result := Hits[0].Line;
  end;

begin
  Plan := Default(TSigPlan);
  if Length(AHeaders) = 0 then
  begin
    Err('the declaration was not found');
    Exit(Plan);
  end;
  Origin := AHeaders[0].Params;
  var V := ValidateSignatureChange(Origin, ANew);
  if V <> '' then
  begin
    Err(V);
    Exit(Plan);
  end;
  SetLength(Joined, Length(ASources));
  SetLength(LinesOf, Length(ASources));
  SetLength(StartsOf, Length(ASources));
  for var I := 0 to High(ASources) do
  begin
    Joined[I] := SigJoinedText(ASources[I].Content);
    LinesOf[I] := Joined[I].Split([#10]);
    StartsOf[I] := LineStartsOf(Joined[I]);
  end;

  // nothing but names changed: the calls stay as they are
  var OnlyNames := Length(ANew) = Length(Origin);
  if OnlyNames then
    for var I := 0 to High(ANew) do
      if (ANew[I].OldIndex <> I) or
         not SameText(Trim(ANew[I].Param.Modifier), Trim(Origin[I].Modifier)) or
         (TypeKey(ANew[I].Param.TypeText) <> TypeKey(Origin[I].TypeText)) or
         (Trim(ANew[I].Param.DefaultText) <> Trim(Origin[I].DefaultText)) then
        OnlyNames := False;

  // 1. the headers of the family and the parameter names in their bodies
  SetLength(BodyPairs, Length(AHeaders));
  SetLength(BodyNames, Length(AHeaders));
  for var HI := 0 to High(AHeaders) do
  begin
    var H := AHeaders[HI];
    // an implementation header may leave out the parameter list
    var Short := (not H.IsDecl) and (H.Open = 0);
    var Own := H.Params;
    if Short then Own := Origin;
    if Length(Own) <> Length(Origin) then
    begin
      Err(Format('%s at %s has %d parameter(s), the declaration %d - align the ' +
        'signatures first', [H.Caption, Loc(H.FileIndex, H.Line), Length(Own), Length(Origin)]));
      Continue;
    end;
    var NewList: TArray<TSigParam> := nil;
    var Pairs: TArray<TArray<string>> := nil;
    for var NP in ANew do
    begin
      var P := NP.Param;
      if NP.OldIndex >= 0 then
      begin
        var O := Own[NP.OldIndex];
        var Org := Origin[NP.OldIndex];
        if SameText(P.Name, Org.Name) then
          P.Name := O.Name                 // not renamed: keep this header's spelling
        else if not SameText(O.Name, Org.Name) then
        begin
          Warn(Format('%s calls the parameter "%s" (the declaration "%s") - its name is kept',
            [H.Caption, O.Name, Org.Name]));
          P.Name := O.Name;
        end;
        if SameText(Trim(P.Modifier), Trim(Org.Modifier)) then P.Modifier := O.Modifier;
        if TypeKey(P.TypeText) = TypeKey(Org.TypeText) then P.TypeText := O.TypeText;
        if Trim(P.DefaultText) = Trim(Org.DefaultText) then P.DefaultText := O.DefaultText;
        if O.Name <> P.Name then Pairs := Pairs + [[O.Name, P.Name]];
      end;
      if not H.IsDecl and not H.HadDefaults then P.DefaultText := '';
      NewList := NewList + [P];
    end;
    BodyPairs[HI] := Pairs;
    for var P in NewList do BodyNames[HI] := BodyNames[HI] + [P.Name];

    // the body: removed parameters must be unused, new names must be free
    if (not H.IsDecl) or (H.BodyFirst >= 0) then
    begin
      if H.BodyFirst < 0 then
      begin
        if Length(Pairs) > 0 then
          Err(Format('%s at %s: the body could not be delimited - the renamed parameters ' +
            'cannot be followed', [H.Caption, Loc(H.FileIndex, H.Line)]));
      end
      else
      begin
        for var O := 0 to High(Own) do
        begin
          var Kept := False;
          for var NP in ANew do
            if NP.OldIndex = O then Kept := True;
          if Kept then Continue;
          var UL := BodyUses(H, Own[O].Name);
          if UL >= 0 then
            Err(Format('%s: the removed parameter "%s" is still used in its body (%s)',
              [H.Caption, Own[O].Name, Loc(H.FileIndex, UL)]));
        end;
        for var K := 0 to High(NewList) do
        begin
          // a name that is new to THIS header
          var IsNew := True;
          for var O in Own do
            if SameText(O.Name, NewList[K].Name) then IsNew := False;
          if not IsNew then Continue;
          var UL := BodyUses(H, NewList[K].Name);
          if UL >= 0 then
            Err(Format('%s: the name "%s" already means something else in its body (%s)',
              [H.Caption, NewList[K].Name, Loc(H.FileIndex, UL)]));
        end;
        for var Hit in RenameHits(LinesOf[H.FileIndex], H.BodyFirst, H.BodyLast, Pairs) do
          AddEdit(H.FileIndex, StartsOf[H.FileIndex][Hit.Line] + Hit.Col,
            StartsOf[H.FileIndex][Hit.Line] + Hit.Col + Hit.Len - 1, Hit.NewText, Hit.Line, 'body');
      end;
    end;

    // the header itself
    var NewText := FormatParamList(NewList, True);
    if Short or (NewText = FormatParamList(Own, True)) then Continue;
    if H.Open > 0 then
    begin
      if NewText = '' then
        AddEdit(H.FileIndex, H.Open, H.Close, '', H.Line, 'header')
      else
        AddEdit(H.FileIndex, H.Open + 1, H.Close - 1, NewText, H.Line, 'header');
    end
    else if NewText <> '' then
      AddEdit(H.FileIndex, H.NameEnd + 1, H.NameEnd, '(' + NewText + ')', H.Line, 'header');
  end;

  // 2. the calls
  CallSpans := nil;
  if not OnlyNames then
    for var C in ACalls do
    begin
      var Where := Loc(C.FileIndex, C.Line);
      case C.Context of
        ccReference:
          Err(Where + ': used as a method reference (' + C.Text + ') - it would no longer ' +
            'match the procedure type it is assigned to');
        ccAccessor:
          Err(Where + ': property accessor (' + C.Text + ') - change the property by hand');
        ccArgs:
          begin
            var J := Joined[C.FileIndex];
            var Args: TArray<string> := nil;
            for var A in SplitArguments(J, C.Open, C.Close) do Args := Args + [A.Text];
            var T, Note, E: string;
            if not RewriteArguments(Args, Origin, CallParams(C), T, Note, E) then
            begin
              Err(Where + ': ' + E);
              Continue;
            end;
            if Note <> '' then Warn(Where + ': ' + Note);
            if T = string.Join(', ', Args) then Continue;
            // a call inside a body whose parameters are renamed
            for var HI := 0 to High(AHeaders) do
              if (AHeaders[HI].FileIndex = C.FileIndex) and (Length(BodyPairs[HI]) > 0) and
                 (C.Line >= AHeaders[HI].BodyFirst) and (C.Line <= AHeaders[HI].BodyLast) then
              begin
                var Cnt: Integer;
                T := RenameIdentifiersInRange([T], 0, 0, BodyPairs[HI], Cnt)[0];
              end;
            AddEdit(C.FileIndex, C.Open + 1, C.Close - 1, T, C.Line, 'call');
            CallSpans := CallSpans + [Plan.Edits[High(Plan.Edits)]];
          end;
        ccStatement, ccExpression:
          begin
            var T, Note, E: string;
            if not RewriteArguments([], Origin, CallParams(C), T, Note, E) then
            begin
              Err(Where + ': ' + E);
              Continue;
            end;
            if T <> '' then
              AddEdit(C.FileIndex, C.NameEnd + 1, C.NameEnd, '(' + T + ')', C.Line, 'call');
          end;
      end;
    end;

  // 3. body renames inside a rewritten argument list are part of it now
  var KeptEdits: TArray<TSigEdit> := nil;
  for var Ed in Plan.Edits do
  begin
    var Inside := False;
    if Ed.What = 'body' then
      for var Sp in CallSpans do
        if (Sp.FileIndex = Ed.FileIndex) and (Ed.Start >= Sp.Start) and (Ed.Stop <= Sp.Stop) then
          Inside := True;
    if not Inside then KeptEdits := KeptEdits + [Ed];
  end;
  Plan.Edits := KeptEdits;

  // 4. sorted; overlapping edits would mean a planning error
  TArray.Sort<TSigEdit>(Plan.Edits, TComparer<TSigEdit>.Construct(
    function(const A, B: TSigEdit): Integer
    begin
      Result := A.FileIndex - B.FileIndex;
      if Result = 0 then Result := A.Start - B.Start;
    end));
  for var I := 1 to High(Plan.Edits) do
    if (Plan.Edits[I].FileIndex = Plan.Edits[I - 1].FileIndex) and
       (Plan.Edits[I].Start <= Plan.Edits[I - 1].Stop) then
      Err(Format('overlapping edits at %s - please report this', [Loc(Plan.Edits[I].FileIndex,
        Plan.Edits[I].Line)]));
  Result := Plan;
end;

function ApplySigEdits(const AContent: string; const AEdits: TArray<TSigEdit>;
  AFileIndex: Integer): string;
var
  Mine: TArray<TSigEdit>;
begin
  Result := SigJoinedText(AContent);
  Mine := nil;
  for var E in AEdits do
    if E.FileIndex = AFileIndex then Mine := Mine + [E];
  TArray.Sort<TSigEdit>(Mine, TComparer<TSigEdit>.Construct(
    function(const A, B: TSigEdit): Integer
    begin
      Result := B.Start - A.Start;
    end));
  for var E in Mine do
    Result := Copy(Result, 1, E.Start - 1) + E.NewText + Copy(Result, E.Stop + 1, MaxInt);
end;

end.
