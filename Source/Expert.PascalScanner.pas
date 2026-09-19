(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.PascalScanner;

// THE one Pascal lexer of the plugin (issue #11, suggestion 7 by Ian
// Branch): comment-, directive- and string-aware, with token kinds and
// positions. Before, TPascalScanner sat hidden in the implementation of
// Expert.SelectionValidator and every other feature carried its own
// "is this an identifier character" / "strip the // comment" / "skip the
// string" code - about twenty copies, with different answers:
//   * Unicode identifiers ('Größe' is a valid Delphi identifier) were
//     split by the ASCII-only copies,
//   * three StripLineComment copies cut at the first '//' even inside a
//     string ('http://...'),
//   * Delphi 12 multi-line strings (''' ... ''') were scanned as CODE by
//     everything except the with-scanner.
// Pure unit: no VCL, no ToolsAPI - usable from worker threads and tests.

interface

uses
  System.SysUtils;

type
  TPasTokenKind = (
    ptEnd,        // no more tokens
    ptIdent,      // identifier or keyword ('&begin' = escaped identifier)
    ptNumber,     // 123, 1_000, 1.5e3, $FF, %1010
    ptString,     // 'text', a multi-line '''...''' literal, #13, #$0D
    ptSymbol,     // ( ) [ ] ; , . : = + - * / ^ @ < > := <= >= <> .. (. .)
    ptComment,    // // ..., { ... }, (* ... *)   (only when requested)
    ptDirective); // {$...}, (*$...*)             (only when requested)

  TPasToken = record
    Kind: TPasTokenKind;
    Text: string;     // the token exactly as in the source
    Pos: Integer;     // 1-based offset into the scanned text
    Line: Integer;    // 0-based line of the first character
    Col: Integer;     // 0-based column (UTF-16 code units) of it
    /// <summary>Text in upper case (keywords compare against this).</summary>
    function Upper: string;
    /// <summary>A NON-escaped identifier equal to AWord (case-insensitive).
    ///  '&begin' is not the keyword begin.</summary>
    function IsWord(const AWord: string): Boolean;
    function IsSymbol(const ASymbol: string): Boolean;
  end;

  TPascalScanner = class
  private
    FText: string;
    FLen: Integer;
    FPos: Integer;
    FLine: Integer;
    FLineStart: Integer;
    FWantComments: Boolean;
    FWantDirectives: Boolean;
    function Peek(AOffset: Integer = 0): Char; inline;
    procedure Step;
    procedure StepN(N: Integer);
    procedure ScanStringLiteral;
  public
    /// <summary>AWantComments / AWantDirectives: return those as tokens
    ///  instead of skipping them.</summary>
    constructor Create(const AText: string; AWantComments: Boolean = False;
      AWantDirectives: Boolean = False);
    /// <summary>The next token; False (and Kind = ptEnd) at the end.</summary>
    function Next(out AToken: TPasToken): Boolean;
    /// <summary>The token stream the selection validator was built on:
    ///  identifiers in UPPER case and the brackets ( ) [ ] plus ; and . -
    ///  everything else is skipped. '(.' / '.)' come back as '[' / ']'.</summary>
    function NextToken(out AToken: string): Boolean;
    property Position: Integer read FPos;
  end;

/// <summary>Delphi's identifier rules: a letter (Unicode) or '_' starts
///  one, letters, digits and '_' continue it.</summary>
function IsIdentStart(C: Char): Boolean;
function IsIdentChar(C: Char): Boolean;

/// <summary>S is a single identifier (no dots, not empty).</summary>
function IsIdentifier(const S: string): Boolean;

/// <summary>ALine without a trailing '// ...' comment (right-trimmed when
///  one was cut). A '//' inside a string literal stays.</summary>
function StripLineComment(const ALine: string): string;

/// <summary>The qualifier directly before the identifier at the 0-based
///  column ACol0: for "lMyClassA.Init" with the caret on Init this is
///  'lMyClassA'. '' when the use site is not dot-qualified. Whitespace
///  around the dot is allowed; only the LAST segment is returned
///  ("A.B.Init" -&gt; 'B').</summary>
function QualifierBefore(const ALine: string; ACol0: Integer): string;

/// <summary>ALines with every comment, directive and string literal
///  character replaced by a blank - line count and line lengths are kept,
///  so a position in the result IS the position in the source. Block
///  comments and multi-line strings carry across lines. Char literals
///  (#13) stay: they are not text anybody searches in. A hit at
///  Masked[L][P] is code iff Masked[L][P] = ALines[L][P].</summary>
function MaskCommentsAndStrings(const ALines: TArray<string>): TArray<string>;

/// <summary>Quote count of a Delphi 12 multi-line string opener at AIndex
///  (an odd run of >= 3 apostrophes followed only by blanks up to the line
///  end), else 0.</summary>
function MultiLineStringOpener(const AText: string; AIndex: Integer): Integer;

implementation

uses
  System.Character;

// ASCII first (the common case, no table lookup), Unicode letters after.
function IsIdentStart(C: Char): Boolean;
begin
  if Ord(C) < 128 then
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '_'])
  else
    Result := C.IsLetter;
end;

function IsIdentChar(C: Char): Boolean;
begin
  if Ord(C) < 128 then
    Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_'])
  else
    Result := C.IsLetterOrDigit;
end;

function IsIdentifier(const S: string): Boolean;
begin
  Result := (S <> '') and IsIdentStart(S[1]);
  if Result then
    for var I := 2 to Length(S) do
      if not IsIdentChar(S[I]) then Exit(False);
end;

function QualifierBefore(const ALine: string; ACol0: Integer): string;
var
  I, EndP: Integer;
begin
  Result := '';
  I := ACol0;                      // 1-based index of the char BEFORE it
  while (I >= 1) and (I <= Length(ALine)) and CharInSet(ALine[I], [' ', #9]) do Dec(I);
  if (I < 1) or (I > Length(ALine)) or (ALine[I] <> '.') then Exit;
  Dec(I);
  while (I >= 1) and CharInSet(ALine[I], [' ', #9]) do Dec(I);
  EndP := I;
  while (I >= 1) and IsIdentChar(ALine[I]) do Dec(I);
  if EndP > I then Result := Copy(ALine, I + 1, EndP - I);
end;

function StripLineComment(const ALine: string): string;
var
  I, N: Integer;
  InStr: Boolean;
begin
  N := Length(ALine);
  InStr := False;
  I := 1;
  while I <= N do
  begin
    if ALine[I] = '''' then
      InStr := not InStr   // '' inside a string toggles twice - still inside
    else if (not InStr) and (ALine[I] = '/') and (I < N) and (ALine[I + 1] = '/') then
      Exit(TrimRight(Copy(ALine, 1, I - 1)));
    Inc(I);
  end;
  Result := ALine;
end;

function MultiLineStringOpener(const AText: string; AIndex: Integer): Integer;
var
  Q, K, N: Integer;
begin
  Result := 0;
  N := Length(AText);
  Q := 0;
  while (AIndex + Q <= N) and (AText[AIndex + Q] = '''') do Inc(Q);
  if (Q < 3) or not Odd(Q) then Exit;
  K := AIndex + Q;
  while (K <= N) and CharInSet(AText[K], [' ', #9]) do Inc(K);
  if (K > N) or CharInSet(AText[K], [#13, #10]) then
    Result := Q;
end;

{ TPasToken }

function TPasToken.Upper: string;
begin
  Result := UpperCase(Text);
end;

function TPasToken.IsWord(const AWord: string): Boolean;
begin
  Result := (Kind = ptIdent) and SameText(Text, AWord);
end;

function TPasToken.IsSymbol(const ASymbol: string): Boolean;
begin
  Result := (Kind = ptSymbol) and (Text = ASymbol);
end;

{ TPascalScanner }

constructor TPascalScanner.Create(const AText: string; AWantComments,
  AWantDirectives: Boolean);
begin
  inherited Create;
  FText := AText;
  FLen := Length(AText);
  FPos := 1;
  FLine := 0;
  FLineStart := 1;
  FWantComments := AWantComments;
  FWantDirectives := AWantDirectives;
end;

function TPascalScanner.Peek(AOffset: Integer): Char;
begin
  if FPos + AOffset <= FLen then
    Result := FText[FPos + AOffset]
  else
    Result := #0;
end;

// One character forward, counting line breaks (CRLF, LF and a lone CR).
procedure TPascalScanner.Step;
begin
  if FPos > FLen then Exit;
  if (FText[FPos] = #10) or ((FText[FPos] = #13) and (Peek(1) <> #10)) then
  begin
    Inc(FPos);
    Inc(FLine);
    FLineStart := FPos;
  end
  else
    Inc(FPos);
end;

procedure TPascalScanner.StepN(N: Integer);
begin
  for var I := 1 to N do Step;
end;

procedure TPascalScanner.ScanStringLiteral;
var
  Q, K: Integer;
begin
  Q := MultiLineStringOpener(FText, FPos);
  if Q > 0 then
  begin
    // ends at a line whose first non-blank characters are the same run
    StepN(Q);
    while FPos <= FLen do
    begin
      while (FPos <= FLen) and not CharInSet(FText[FPos], [#13, #10]) do Step;
      if FPos > FLen then Exit;
      if (FText[FPos] = #13) and (Peek(1) = #10) then Step;
      Step;   // the line break
      while CharInSet(Peek, [' ', #9]) do Step;
      K := 0;
      while Peek(K) = '''' do Inc(K);
      if K = Q then
      begin
        StepN(Q);
        Exit;
      end;
    end;
    Exit;
  end;
  // 'text' with '' as the escaped apostrophe; never spans a line
  Step;
  while FPos <= FLen do
  begin
    case FText[FPos] of
      '''':
        if Peek(1) = '''' then
          StepN(2)
        else
        begin
          Step;
          Exit;
        end;
      #13, #10:
        Exit;   // unterminated
    else
      Step;
    end;
  end;
end;

function TPascalScanner.Next(out AToken: TPasToken): Boolean;
var
  C, D: Char;
  Start: Integer;

  procedure Emit(AKind: TPasTokenKind);
  begin
    AToken.Kind := AKind;
    AToken.Text := Copy(FText, Start, FPos - Start);
  end;

begin
  AToken := Default(TPasToken);
  while FPos <= FLen do
  begin
    C := FText[FPos];
    if C <= ' ' then
    begin
      Step;
      Continue;
    end;
    Start := FPos;
    AToken.Pos := FPos;
    AToken.Line := FLine;
    AToken.Col := FPos - FLineStart;
    D := Peek(1);

    // comments and directives
    if (C = '/') and (D = '/') then
    begin
      while (FPos <= FLen) and not CharInSet(FText[FPos], [#13, #10]) do Step;
      if FWantComments then begin Emit(ptComment); Exit(True); end;
      Continue;
    end;
    if C = '{' then
    begin
      var IsDir := D = '$';
      while (FPos <= FLen) and (FText[FPos] <> '}') do Step;
      Step;   // the '}'
      if IsDir and FWantDirectives then begin Emit(ptDirective); Exit(True); end;
      if (not IsDir) and FWantComments then begin Emit(ptComment); Exit(True); end;
      Continue;
    end;
    if (C = '(') and (D = '*') then
    begin
      var IsDir := Peek(2) = '$';
      StepN(2);
      while (FPos <= FLen) and not ((FText[FPos] = '*') and (Peek(1) = ')')) do Step;
      StepN(2);
      if IsDir and FWantDirectives then begin Emit(ptDirective); Exit(True); end;
      if (not IsDir) and FWantComments then begin Emit(ptComment); Exit(True); end;
      Continue;
    end;

    // strings and char literals
    if C = '''' then
    begin
      ScanStringLiteral;
      Emit(ptString);
      Exit(True);
    end;
    if (C = '#') and (CharInSet(D, ['0'..'9']) or
      ((D = '$') and CharInSet(Peek(2), ['0'..'9', 'A'..'F', 'a'..'f']))) then
    begin
      Step;
      if Peek = '$' then
      begin
        Step;
        while CharInSet(Peek, ['0'..'9', 'A'..'F', 'a'..'f', '_']) do Step;
      end
      else
        while CharInSet(Peek, ['0'..'9', '_']) do Step;
      Emit(ptString);
      Exit(True);
    end;

    // identifiers ('&' escapes a reserved word: &type, &begin)
    if IsIdentStart(C) or ((C = '&') and IsIdentStart(D)) then
    begin
      Step;
      while (FPos <= FLen) and IsIdentChar(FText[FPos]) do Step;
      Emit(ptIdent);
      Exit(True);
    end;

    // numbers
    if CharInSet(C, ['0'..'9']) then
    begin
      while CharInSet(Peek, ['0'..'9', '_']) do Step;
      // a fraction only when a digit follows the dot - '1..5' is a range
      if (Peek = '.') and CharInSet(Peek(1), ['0'..'9']) then
      begin
        Step;
        while CharInSet(Peek, ['0'..'9', '_']) do Step;
      end;
      if CharInSet(Peek, ['e', 'E']) and (CharInSet(Peek(1), ['0'..'9']) or
        (CharInSet(Peek(1), ['+', '-']) and CharInSet(Peek(2), ['0'..'9']))) then
      begin
        StepN(2);
        while CharInSet(Peek, ['0'..'9']) do Step;
      end;
      Emit(ptNumber);
      Exit(True);
    end;
    if (C = '$') and CharInSet(D, ['0'..'9', 'A'..'F', 'a'..'f']) then
    begin
      Step;
      while CharInSet(Peek, ['0'..'9', 'A'..'F', 'a'..'f', '_']) do Step;
      Emit(ptNumber);
      Exit(True);
    end;
    if (C = '%') and CharInSet(D, ['0', '1']) then
    begin
      Step;
      while CharInSet(Peek, ['0', '1', '_']) do Step;
      Emit(ptNumber);
      Exit(True);
    end;
    if (C = '&') and CharInSet(D, ['0'..'7']) then
    begin
      Step;
      while CharInSet(Peek, ['0'..'7', '_']) do Step;
      Emit(ptNumber);
      Exit(True);
    end;

    // symbols: the two-character ones first
    if ((C = ':') and (D = '=')) or ((C = '<') and CharInSet(D, ['=', '>'])) or
       ((C = '>') and (D = '=')) or ((C = '.') and CharInSet(D, ['.', ')'])) or
       ((C = '(') and (D = '.')) then
      StepN(2)
    else
      Step;
    Emit(ptSymbol);
    Exit(True);
  end;
  AToken.Kind := ptEnd;
  AToken.Pos := FPos;
  AToken.Line := FLine;
  AToken.Col := FPos - FLineStart;
  Result := False;
end;

function TPascalScanner.NextToken(out AToken: string): Boolean;
var
  T: TPasToken;
begin
  AToken := '';
  while Next(T) do
    case T.Kind of
      ptIdent:
        begin
          AToken := T.Upper;
          Exit(True);
        end;
      ptSymbol:
        begin
          if T.Text = '(.' then AToken := '['
          else if T.Text = '.)' then AToken := ']'
          else if (Length(T.Text) = 1) and CharInSet(T.Text[1], ['(', ')', '[', ']', ';', '.']) then
            AToken := T.Text
          else
            Continue;
          Exit(True);
        end;
    end;
  Result := False;
end;

function MaskCommentsAndStrings(const ALines: TArray<string>): TArray<string>;
var
  Text: string;
  Starts: TArray<Integer>;
  S: TPascalScanner;
  T: TPasToken;
  Buf: string;
  Masked: Boolean;
begin
  Result := Copy(ALines);
  if Length(ALines) = 0 then Exit;
  // one text with #10 separators; Starts[L] = 1-based offset of line L
  SetLength(Starts, Length(ALines));
  var Total := 0;
  for var L := 0 to High(ALines) do
  begin
    Starts[L] := Total + 1;
    Inc(Total, Length(ALines[L]) + 1);
  end;
  Text := string.Join(#10, ALines);
  Buf := Text;
  Masked := False;
  S := TPascalScanner.Create(Text, True, True);
  try
    while S.Next(T) do
      if (T.Kind in [ptComment, ptDirective]) or
         ((T.Kind = ptString) and (T.Text[1] = '''')) then
      begin
        for var I := T.Pos to T.Pos + Length(T.Text) - 1 do
          if Buf[I] <> #10 then
          begin
            Buf[I] := ' ';
            Masked := True;
          end;
      end;
  finally
    S.Free;
  end;
  if not Masked then Exit;
  for var L := 0 to High(ALines) do
    if Copy(Buf, Starts[L], Length(ALines[L])) <> ALines[L] then
      Result[L] := Copy(Buf, Starts[L], Length(ALines[L]));
end;

end.
