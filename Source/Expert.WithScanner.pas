(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.WithScanner;

// Pure-text scanner for "with X, Y, ... do ..." statements in Pascal
// source. No LSP dependency; produces a list of occurrences with source
// ranges that the rewriter / dialog consume.
//
// Output positions are 1-based line/column to match what the rest of
// the refactoring code uses internally; the LSP layer (0-based)
// converts at the edge.
//
// What this scanner DOES:
//  - skips string literals, curly comments, paren-star comments and
//    line comments
//  - finds the keyword "with" at a statement boundary
//  - parses the target list separated by "," until the matching "do",
//    respecting nested parentheses and brackets so commas inside calls
//    do not split targets
//  - locates the body: either a begin..end block, or a single
//    statement up to the next semicolon at the same nesting level
//  - tracks block depth using BEGIN/TRY/CASE/ASM as openers and END as
//    closer (same convention used in Expert.ExtractMethod.FindMethodEnd)
//
// What this scanner does NOT do:
//  - resolve identifiers (that is the LSP rewriter's job)
//  - judge whether a target expression is simple enough to inline
//    (the rewriter classifies it)

interface

uses
  System.Classes, System.SysUtils;

type
  /// <summary>1-based source position. Col counts code units (chars), not
  ///  bytes. Line 1 = first line of the file.</summary>
  TWithSourcePos = record
    Line: Integer;
    Col: Integer;
  end;

  /// <summary>An inclusive 1-based source range.</summary>
  TWithSourceRange = record
    StartPos: TWithSourcePos;
    EndPos: TWithSourcePos;
  end;

  /// <summary>One target inside the with-clause, e. g. "FFoo.Bar" or
  ///  "GetThing(42)". The textual expression is preserved verbatim
  ///  (including whitespace inside parentheses) so the rewriter can
  ///  decide whether to inline a temp var.</summary>
  TWithTarget = record
    /// <summary>Raw source text of the expression.</summary>
    Expression: string;
    /// <summary>Source range of the expression (without surrounding
    ///  whitespace or commas).</summary>
    Range: TWithSourceRange;
  end;

  /// <summary>Body shape of a with statement.</summary>
  TWithBodyKind = (
    /// <summary>Body is a begin..end block. BodyRange covers the begin
    ///  through end keyword, BodyInnerRange is what's between them.</summary>
    wbkBeginEnd,
    /// <summary>Body is a single statement terminated by ';' or eof.</summary>
    wbkSingle
  );

  TWithOccurrence = record
    /// <summary>Position of the 'with' keyword itself (1-based).</summary>
    KeywordPos: TWithSourcePos;
    /// <summary>Position of the matching 'do' keyword (1-based,
    ///  points to the 'd' of 'do').</summary>
    DoPos: TWithSourcePos;
    /// <summary>One or more targets parsed between 'with' and 'do'.</summary>
    Targets: TArray<TWithTarget>;
    BodyKind: TWithBodyKind;
    /// <summary>Whole body range, including 'begin' and 'end' if present,
    ///  or the single statement terminator.</summary>
    BodyRange: TWithSourceRange;
    /// <summary>Inner body range — for begin..end, between begin and end
    ///  (exclusive); for single statement, identical to BodyRange.</summary>
    BodyInnerRange: TWithSourceRange;
  end;

  TWithScanner = class
  public
    /// <summary>Scans the given source and returns one entry per
    ///  with-statement found. Malformed with-statements (no 'do', no
    ///  body) are silently skipped; the scanner never raises.</summary>
    class function ScanSource(const ASource: string): TArray<TWithOccurrence>; static;

    /// <summary>Convenience overload reading from a file.</summary>
    class function ScanFile(const AFileName: string): TArray<TWithOccurrence>; static;
  end;

implementation

uses
  System.Character, System.IOUtils, System.Generics.Collections;

type
  /// <summary>Cursor over the source string with 1-based line/col tracking.</summary>
  TCursor = record
    Src: string;
    Idx: Integer;       // 1-based string index into Src
    Line: Integer;      // 1-based current line
    Col: Integer;       // 1-based current column
    function Eof: Boolean; inline;
    function Peek: Char; inline;
    function PeekAt(AOffset: Integer): Char; inline;
    procedure Advance; inline;
    function Pos: TWithSourcePos; inline;
  end;

function TCursor.Eof: Boolean;
begin
  Result := Idx > Length(Src);
end;

function TCursor.Peek: Char;
begin
  if Idx > Length(Src) then
    Result := #0
  else
    Result := Src[Idx];
end;

function TCursor.PeekAt(AOffset: Integer): Char;
var
  P: Integer;
begin
  P := Idx + AOffset;
  if (P < 1) or (P > Length(Src)) then
    Result := #0
  else
    Result := Src[P];
end;

procedure TCursor.Advance;
var
  Ch: Char;
begin
  if Idx > Length(Src) then Exit;
  Ch := Src[Idx];
  Inc(Idx);
  if Ch = #10 then
  begin
    Inc(Line);
    Col := 1;
  end
  else if Ch = #13 then
  begin
    // CRLF: don't double-bump on the LF that follows
    Inc(Line);
    Col := 1;
    if (Idx <= Length(Src)) and (Src[Idx] = #10) then
      Inc(Idx);
  end
  else
    Inc(Col);
end;

function TCursor.Pos: TWithSourcePos;
begin
  Result.Line := Line;
  Result.Col := Col;
end;

{ Lexical helpers }

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := (C = '_') or C.IsLetter;
end;

function IsIdentCont(C: Char): Boolean; inline;
begin
  Result := (C = '_') or C.IsLetterOrDigit;
end;

/// <summary>Skips one whitespace/comment token if at one. Returns True
///  when something was skipped. Strings are NOT skipped here — the
///  outer state machine treats string literals separately because
///  identifiers cannot appear inside them.</summary>
function SkipTrivia(var Cur: TCursor): Boolean;
var
  Started: Boolean;
begin
  Started := False;
  while not Cur.Eof do
  begin
    case Cur.Peek of
      #9, #10, #11, #12, #13, ' ':
        begin
          Cur.Advance;
          Started := True;
        end;
      '/':
        begin
          if Cur.PeekAt(1) = '/' then
          begin
            // line comment to EOL
            while (not Cur.Eof) and (Cur.Peek <> #10) and (Cur.Peek <> #13) do
              Cur.Advance;
            Started := True;
          end
          else
            Break;
        end;
      '{':
        begin
          // curly comment (incl. compiler directives — we treat them as
          // comments for scanning purposes; that's fine because a
          // directive cannot start a 'with'-statement)
          Cur.Advance;
          while (not Cur.Eof) and (Cur.Peek <> '}') do
            Cur.Advance;
          if not Cur.Eof then
            Cur.Advance; // consume '}'
          Started := True;
        end;
      '(':
        begin
          if Cur.PeekAt(1) = '*' then
          begin
            Cur.Advance; // (
            Cur.Advance; // *
            while not Cur.Eof do
            begin
              if (Cur.Peek = '*') and (Cur.PeekAt(1) = ')') then
              begin
                Cur.Advance;
                Cur.Advance;
                Break;
              end;
              Cur.Advance;
            end;
            Started := True;
          end
          else
            Break;
        end;
    else
      Break;
    end;
  end;
  Result := Started;
end;

/// <summary>Skips a Pascal string literal '...''...'. Doubled apostrophes
///  inside are escapes. Cursor must be at the opening apostrophe.</summary>
procedure SkipString(var Cur: TCursor);
begin
  if Cur.Peek <> '''' then Exit;
  Cur.Advance;
  while not Cur.Eof do
  begin
    if Cur.Peek = '''' then
    begin
      if Cur.PeekAt(1) = '''' then
      begin
        Cur.Advance;
        Cur.Advance;
      end
      else
      begin
        Cur.Advance;
        Exit;
      end;
    end
    else if (Cur.Peek = #10) or (Cur.Peek = #13) then
      // unterminated string — bail
      Exit
    else
      Cur.Advance;
  end;
end;

/// <summary>Skips a #nnnn or #$nn character constant.</summary>
procedure SkipCharConst(var Cur: TCursor);
begin
  if Cur.Peek <> '#' then Exit;
  Cur.Advance;
  if Cur.Peek = '$' then
  begin
    Cur.Advance;
    while CharInSet(Cur.Peek, ['0'..'9', 'a'..'f', 'A'..'F']) do
      Cur.Advance;
  end
  else
    while CharInSet(Cur.Peek, ['0'..'9']) do
      Cur.Advance;
end;

/// <summary>Reads an identifier starting at the cursor. Cursor must be
///  at an identifier-start char. Returns the identifier text.</summary>
function ReadIdent(var Cur: TCursor; out AStart: TWithSourcePos): string;
var
  SB: TStringBuilder;
begin
  AStart := Cur.Pos;
  SB := TStringBuilder.Create;
  try
    while (not Cur.Eof) and IsIdentCont(Cur.Peek) do
    begin
      SB.Append(Cur.Peek);
      Cur.Advance;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function SameKeyword(const AIdent, AKeyword: string): Boolean; inline;
begin
  Result := SameText(AIdent, AKeyword);
end;

{ Target parsing }

/// <summary>Parses one with-target: an expression up to a top-level ','
///  or 'do'. Top-level means: not inside () or []. The expression text
///  is captured verbatim (incl. internal whitespace and comments) so
///  the rewriter sees what the user wrote.</summary>
function ReadOneTarget(var Cur: TCursor; out ATarget: TWithTarget): Boolean;
var
  StartIdx, EndIdx: Integer;
  StartPos, LastNonWsPos: TWithSourcePos;
  ParenDepth, BrackDepth: Integer;
  Tail: string;
  Ident: string;
  IdentStart: TWithSourcePos;
begin
  Result := False;
  // Skip leading whitespace/comments before the target so StartPos
  // points at the first meaningful char.
  SkipTrivia(Cur);
  if Cur.Eof then Exit;

  StartIdx := Cur.Idx;
  StartPos := Cur.Pos;
  LastNonWsPos := StartPos;
  EndIdx := StartIdx;
  ParenDepth := 0;
  BrackDepth := 0;

  while not Cur.Eof do
  begin
    case Cur.Peek of
      '(':
        begin
          Inc(ParenDepth);
          LastNonWsPos := Cur.Pos;
          EndIdx := Cur.Idx;
          Cur.Advance;
        end;
      ')':
        begin
          if ParenDepth = 0 then
            // unmatched close-paren — let the outer loop see it as garbage
            Break;
          Dec(ParenDepth);
          LastNonWsPos := Cur.Pos;
          EndIdx := Cur.Idx;
          Cur.Advance;
        end;
      '[':
        begin
          Inc(BrackDepth);
          LastNonWsPos := Cur.Pos;
          EndIdx := Cur.Idx;
          Cur.Advance;
        end;
      ']':
        begin
          if BrackDepth = 0 then Break;
          Dec(BrackDepth);
          LastNonWsPos := Cur.Pos;
          EndIdx := Cur.Idx;
          Cur.Advance;
        end;
      ',':
        if (ParenDepth = 0) and (BrackDepth = 0) then
          Break
        else
        begin
          LastNonWsPos := Cur.Pos;
          EndIdx := Cur.Idx;
          Cur.Advance;
        end;
      '''':
        begin
          LastNonWsPos := Cur.Pos;
          SkipString(Cur);
          EndIdx := Cur.Idx - 1;
        end;
      '#':
        begin
          LastNonWsPos := Cur.Pos;
          SkipCharConst(Cur);
          EndIdx := Cur.Idx - 1;
        end;
      '/':
        begin
          if Cur.PeekAt(1) = '/' then
            SkipTrivia(Cur)
          else
          begin
            LastNonWsPos := Cur.Pos;
            EndIdx := Cur.Idx;
            Cur.Advance;
          end;
        end;
      '{':
        SkipTrivia(Cur);
      ' ', #9, #10, #11, #12, #13:
        Cur.Advance;
    else
      // Could be an identifier or a punctuation char — but we also need
      // to detect 'do' at top level to terminate.
      if IsIdentStart(Cur.Peek) and (ParenDepth = 0) and (BrackDepth = 0) then
      begin
        // Peek the identifier without consuming
        var SaveIdx := Cur.Idx;
        var SaveLine := Cur.Line;
        var SaveCol := Cur.Col;
        Ident := ReadIdent(Cur, IdentStart);
        if SameKeyword(Ident, 'do') then
        begin
          // Roll back so the outer parser sees 'do'
          Cur.Idx := SaveIdx;
          Cur.Line := SaveLine;
          Cur.Col := SaveCol;
          Break;
        end;
        // It's part of the target expression
        LastNonWsPos.Line := IdentStart.Line;
        LastNonWsPos.Col := IdentStart.Col + Length(Ident) - 1;
        EndIdx := Cur.Idx - 1;
      end
      else
      begin
        LastNonWsPos := Cur.Pos;
        EndIdx := Cur.Idx;
        Cur.Advance;
      end;
    end;
  end;

  if EndIdx < StartIdx then Exit;
  Tail := Copy(Cur.Src, StartIdx, EndIdx - StartIdx + 1);
  Tail := Tail.TrimRight;
  if Tail = '' then Exit;

  ATarget.Expression := Tail;
  ATarget.Range.StartPos := StartPos;
  ATarget.Range.EndPos := LastNonWsPos;
  Result := True;
end;

{ Body parsing }

/// <summary>From the current cursor position (just after 'do' + trivia),
///  determines the body shape and the body range.</summary>
function ReadBody(var Cur: TCursor; var AOcc: TWithOccurrence): Boolean;
var
  Ident: string;
  IdentStart: TWithSourcePos;
  BlockDepth: Integer;
  StartIdx, BodyStartIdx: Integer;
  StartPos, EndPos, InnerStart, InnerEnd: TWithSourcePos;
  SaveIdx: Integer;
  SaveLine, SaveCol: Integer;
begin
  Result := False;
  SkipTrivia(Cur);
  if Cur.Eof then Exit;

  StartIdx := Cur.Idx;
  StartPos := Cur.Pos;

  if IsIdentStart(Cur.Peek) then
  begin
    SaveIdx := Cur.Idx;
    SaveLine := Cur.Line;
    SaveCol := Cur.Col;
    Ident := ReadIdent(Cur, IdentStart);
    if SameKeyword(Ident, 'begin') then
    begin
      // begin..end body. Walk until block depth balances.
      AOcc.BodyKind := wbkBeginEnd;
      InnerStart := Cur.Pos;
      BlockDepth := 1;
      while (not Cur.Eof) and (BlockDepth > 0) do
      begin
        SkipTrivia(Cur);
        if Cur.Eof then Break;
        case Cur.Peek of
          '''': SkipString(Cur);
          '#':  SkipCharConst(Cur);
        else
          if IsIdentStart(Cur.Peek) then
          begin
            var WordStart := Cur.Idx;
            Ident := ReadIdent(Cur, IdentStart);
            if SameKeyword(Ident, 'begin') or SameKeyword(Ident, 'try')
              or SameKeyword(Ident, 'case') or SameKeyword(Ident, 'asm') then
              Inc(BlockDepth)
            else if SameKeyword(Ident, 'end') then
            begin
              Dec(BlockDepth);
              if BlockDepth = 0 then
              begin
                EndPos.Line := IdentStart.Line;
                EndPos.Col := IdentStart.Col + 2; // 'end' is 3 chars; last col = start+2
                InnerEnd.Line := IdentStart.Line;
                InnerEnd.Col := IdentStart.Col - 1;
                AOcc.BodyRange.StartPos := StartPos;
                AOcc.BodyRange.EndPos := EndPos;
                AOcc.BodyInnerRange.StartPos := InnerStart;
                AOcc.BodyInnerRange.EndPos := InnerEnd;
                Exit(True);
              end;
            end;
            // otherwise: just an identifier, keep going
            // (suppress hint about WordStart being unused in some configs)
            if WordStart < 0 then ;
          end
          else
            Cur.Advance;
        end;
      end;
      // unbalanced — bail
      Exit(False);
    end
    else
    begin
      // Not 'begin' — single-statement body. Roll back and consume up to ';'
      Cur.Idx := SaveIdx;
      Cur.Line := SaveLine;
      Cur.Col := SaveCol;
    end;
  end;

  // Single-statement form. Read until ';' at outer depth, or until 'end'
  // / 'else' / EOF (a single-statement with at the tail of a block has
  // no ';' before the enclosing 'end').
  AOcc.BodyKind := wbkSingle;
  BodyStartIdx := Cur.Idx;
  if BodyStartIdx < 0 then ; // suppress unused-warning in some configs
  EndPos := StartPos;
  var ParenDepth := 0;
  var BrackDepth := 0;
  var LastNonWsPos := StartPos;
  while not Cur.Eof do
  begin
    case Cur.Peek of
      '(': begin Inc(ParenDepth); LastNonWsPos := Cur.Pos; Cur.Advance; end;
      ')': begin if ParenDepth > 0 then Dec(ParenDepth); LastNonWsPos := Cur.Pos; Cur.Advance; end;
      '[': begin Inc(BrackDepth); LastNonWsPos := Cur.Pos; Cur.Advance; end;
      ']': begin if BrackDepth > 0 then Dec(BrackDepth); LastNonWsPos := Cur.Pos; Cur.Advance; end;
      '''': begin LastNonWsPos := Cur.Pos; SkipString(Cur); end;
      '#':  begin LastNonWsPos := Cur.Pos; SkipCharConst(Cur); end;
      ';':
        if (ParenDepth = 0) and (BrackDepth = 0) then
        begin
          LastNonWsPos := Cur.Pos;
          Cur.Advance;
          Break;
        end
        else
        begin
          LastNonWsPos := Cur.Pos;
          Cur.Advance;
        end;
      '/':
        if Cur.PeekAt(1) = '/' then SkipTrivia(Cur)
        else begin LastNonWsPos := Cur.Pos; Cur.Advance; end;
      '{': SkipTrivia(Cur);
      ' ', #9, #10, #11, #12, #13: Cur.Advance;
    else
      if IsIdentStart(Cur.Peek) and (ParenDepth = 0) and (BrackDepth = 0) then
      begin
        SaveIdx := Cur.Idx; SaveLine := Cur.Line; SaveCol := Cur.Col;
        Ident := ReadIdent(Cur, IdentStart);
        if SameKeyword(Ident, 'end') or SameKeyword(Ident, 'else')
          or SameKeyword(Ident, 'until') or SameKeyword(Ident, 'finally')
          or SameKeyword(Ident, 'except') then
        begin
          // body terminated by enclosing structure — roll back, do not
          // consume the keyword
          Cur.Idx := SaveIdx; Cur.Line := SaveLine; Cur.Col := SaveCol;
          Break;
        end;
        LastNonWsPos.Line := IdentStart.Line;
        LastNonWsPos.Col := IdentStart.Col + Length(Ident) - 1;
      end
      else
      begin
        LastNonWsPos := Cur.Pos;
        Cur.Advance;
      end;
    end;
  end;

  AOcc.BodyRange.StartPos := StartPos;
  AOcc.BodyRange.EndPos := LastNonWsPos;
  AOcc.BodyInnerRange := AOcc.BodyRange;
  Result := True;
end;

{ Top-level scan }

class function TWithScanner.ScanSource(const ASource: string): TArray<TWithOccurrence>;
var
  Cur: TCursor;
  Results: TList<TWithOccurrence>;
  Occ: TWithOccurrence;
  Ident: string;
  IdentStart: TWithSourcePos;
  AtStmtBoundary: Boolean;
  PrevSig: Char;     // last significant non-trivia char
  Targets: TList<TWithTarget>;
  Target: TWithTarget;
  SaveIdx: Integer;
  SaveLine, SaveCol: Integer;
begin
  Results := TList<TWithOccurrence>.Create;
  try
    Cur.Src := ASource;
    Cur.Idx := 1;
    Cur.Line := 1;
    Cur.Col := 1;
    PrevSig := #0;
    AtStmtBoundary := True;

    while not Cur.Eof do
    begin
      if SkipTrivia(Cur) then
        Continue;
      if Cur.Eof then Break;

      case Cur.Peek of
        '''':
          begin
            SkipString(Cur);
            PrevSig := '''';
            AtStmtBoundary := False;
          end;
        '#':
          begin
            SkipCharConst(Cur);
            PrevSig := '#';
            AtStmtBoundary := False;
          end;
        ';':
          begin
            Cur.Advance;
            PrevSig := ';';
            AtStmtBoundary := True;
          end;
      else
        if IsIdentStart(Cur.Peek) then
        begin
          SaveIdx := Cur.Idx;
          SaveLine := Cur.Line;
          SaveCol := Cur.Col;
          Ident := ReadIdent(Cur, IdentStart);

          if AtStmtBoundary and SameKeyword(Ident, 'with') then
          begin
            Occ := Default(TWithOccurrence);
            Occ.KeywordPos := IdentStart;

            // Parse one or more comma-separated targets
            Targets := TList<TWithTarget>.Create;
            try
              while True do
              begin
                if not ReadOneTarget(Cur, Target) then Break;
                Targets.Add(Target);
                SkipTrivia(Cur);
                if Cur.Peek = ',' then
                begin
                  Cur.Advance;
                  Continue;
                end;
                Break;
              end;
              Occ.Targets := Targets.ToArray;
            finally
              Targets.Free;
            end;

            SkipTrivia(Cur);
            if (Length(Occ.Targets) > 0) and IsIdentStart(Cur.Peek) then
            begin
              SaveIdx := Cur.Idx;
              SaveLine := Cur.Line;
              SaveCol := Cur.Col;
              Ident := ReadIdent(Cur, IdentStart);
              if SameKeyword(Ident, 'do') then
              begin
                Occ.DoPos := IdentStart;
                if ReadBody(Cur, Occ) then
                  Results.Add(Occ);
              end
              else
              begin
                // Not a 'do' — malformed, roll back and let outer loop
                // continue from after the keyword.
                Cur.Idx := SaveIdx;
                Cur.Line := SaveLine;
                Cur.Col := SaveCol;
              end;
            end;
            // After processing a with (or skipping a malformed one),
            // we are no longer at a statement boundary unless the next
            // token says so.
            AtStmtBoundary := False;
            PrevSig := 'a';
            Continue;
          end;

          // Mark statement boundaries after structural keywords. We are
          // permissive here — false positives only mean we'll consider
          // a 'with' that wasn't actually at a boundary; the parser
          // rejects such cases when target parsing fails.
          if SameKeyword(Ident, 'begin') or SameKeyword(Ident, 'do')
            or SameKeyword(Ident, 'then') or SameKeyword(Ident, 'else')
            or SameKeyword(Ident, 'of') or SameKeyword(Ident, 'try')
            or SameKeyword(Ident, 'finally') or SameKeyword(Ident, 'except')
            or SameKeyword(Ident, 'repeat') then
            AtStmtBoundary := True
          else
            AtStmtBoundary := False;

          PrevSig := 'a';

          // Suppress "value never used" warnings in some configs
          if SaveIdx < 0 then Cur.Idx := SaveLine + SaveCol;
        end
        else
        begin
          // Punctuation — only ';' counts as statement boundary, handled
          // above. Other punctuation does not change boundary state.
          PrevSig := Cur.Peek;
          Cur.Advance;
          AtStmtBoundary := False;
        end;
      end;
    end;

    Result := Results.ToArray;
  finally
    Results.Free;
  end;

  // Suppress hint about PrevSig being assigned but not read
  if PrevSig = #1 then ;
end;

class function TWithScanner.ScanFile(const AFileName: string): TArray<TWithOccurrence>;
var
  Source: string;
begin
  Source := TFile.ReadAllText(AFileName);
  Result := ScanSource(Source);
end;

end.
