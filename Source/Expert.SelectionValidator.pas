(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SelectionValidator;

interface

uses
  System.SysUtils, System.Classes, System.Math, System.Generics.Collections;

type
  TValidationLevel = (vlOk, vlWarning, vlError);

  TValidationIssue = record
    Level: TValidationLevel;
    Message: string;
  end;

  TValidationResult = record
    Issues: TArray<TValidationIssue>;
    function IsValid: Boolean;
    function HasErrors: Boolean;
    function HasWarnings: Boolean;
    function ErrorCount: Integer;
    function WarningCount: Integer;
    function FormatIssues: string;
  end;

  TSelectionValidator = class
  public
    /// <summary>Checks the selected code block for semantic problems.
    ///  ASelectedText is the pure block text (without context).
    ///  AFileLines is the whole file (used for context checks).
    ///  AStartLine/AEndLine are 1-based.</summary>
    class function Validate(const ASelectedText: string; const AFileLines: TArray<string>;
      AStartLine, AEndLine, AInsertLine: Integer; const AEnclosingClass: string): TValidationResult;
  end;

/// <summary>Every identifier of ALINES that the text shows being
///  WRITTEN: an assignment at a statement boundary (not only at the
///  start of a line), a for-loop variable, and an argument of a routine
///  whose parameter is known to be var (Inc, Dec, Read, ReadLn, Val,
///  FillChar, SetLength, ...). Extract method only recognised
///  "^identifier :=", so a block with Inc(N) or "for I := 0 to 9" made
///  N / I a CONST parameter and the result did not compile
///  (audit #39, L7f). Comments and string literals are masked first, so
///  nothing inside them counts. Upper case, no duplicates.</summary>
function WrittenIdentifiersIn(const ALines: TArray<string>): TArray<string>;

/// <summary>'' when ANAME is usable as a method name, else why not: empty,
///  not a Pascal identifier ('2Foo', 'My Method') or a reserved word
///  ('begin'). Extract method accepted any non-empty text and produced
///  code that does not compile (audit #39, L7e).</summary>
function MethodNameProblem(const AName: string): string;

implementation

uses
  System.StrUtils, Expert.PascalScanner, Expert.IdentifierCheck;

{ TValidationResult }

function TValidationResult.IsValid: Boolean;
begin
  Result := not HasErrors;
end;

function TValidationResult.HasErrors: Boolean;
begin
  Result := ErrorCount > 0;
end;

function TValidationResult.HasWarnings: Boolean;
begin
  Result := WarningCount > 0;
end;

function TValidationResult.ErrorCount: Integer;
begin
  Result := 0;
  for var Issue in Issues do
    if Issue.Level = vlError then Inc(Result);
end;

function TValidationResult.WarningCount: Integer;
begin
  Result := 0;
  for var Issue in Issues do
    if Issue.Level = vlWarning then Inc(Result);
end;

function TValidationResult.FormatIssues: string;
var
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    for var Issue in Issues do
    begin
      case Issue.Level of
        vlError:   SB.Append('  [ERROR] ');
        vlWarning: SB.Append('  [WARNING] ');
        vlOk:      SB.Append('  [OK] ');
      end;
      SB.AppendLine(Issue.Message);
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

{ TSelectionValidator }

class function TSelectionValidator.Validate(const ASelectedText: string; const AFileLines: TArray<string>;
  AStartLine, AEndLine, AInsertLine: Integer; const AEnclosingClass: string): TValidationResult;
var
  Issues: TList<TValidationIssue>;

  procedure AddError(const AMsg: string);
  var
    Issue: TValidationIssue;
  begin
    Issue.Level := vlError;
    Issue.Message := AMsg;
    Issues.Add(Issue);
  end;

  procedure AddWarning(const AMsg: string);
  var
    Issue: TValidationIssue;
  begin
    Issue.Level := vlWarning;
    Issue.Message := AMsg;
    Issues.Add(Issue);
  end;

  procedure AddOk(const AMsg: string);
  var
    Issue: TValidationIssue;
  begin
    Issue.Level := vlOk;
    Issue.Message := AMsg;
    Issues.Add(Issue);
  end;

var
  Scanner: TPascalScanner;
  Token: string;
  BeginEndDepth: Integer;
  TryDepth: Integer;
  CaseDepth: Integer;
  ParenDepth, BracketDepth: Integer;
  FirstToken, LastToken: string;
  TokenCount: Integer;
  // Tracking for if/then/else
  PendingThen: Integer;       // open if (waiting for then)
  PendingElse: Integer;       // how many else could still come without an if
  OrphanElseFound: Boolean;
  OrphanUntilFound: Boolean;
  RepeatDepth: Integer;
  ForeignFinallyExceptFound: Boolean;
  InsideCase: Boolean;
  // Control-flow tracking (Exit/Break/Continue/goto/raise)
  BlockStack: TList<Integer>;
  LoopHeaderPending: Boolean;  // saw FOR/WHILE, waiting for DO
  ExpectLoopBody: Boolean;     // saw DO, next BEGIN opens a loop body
  ExitCount, GotoCount: Integer;
  BreakOutsideLoop, ContinueOutsideLoop: Integer;
  BareRaiseOutsideExcept: Integer;
  PendingBareRaise: Boolean;
  PendingBareRaiseInExcept: Boolean;

const
  bkBegin = 0; bkLoopBegin = 1; bkRepeat = 2;
  bkTry = 3;   bkExcept = 4;    bkFinally = 5; bkCase = 6;

  function StackContains(AKind: Integer): Boolean;
  var K: Integer;
  begin
    Result := False;
    for K in BlockStack do
      if K = AKind then Exit(True);
  end;

  function InLoop: Boolean;
  begin
    Result := StackContains(bkLoopBegin) or StackContains(bkRepeat);
  end;

  function InExcept: Boolean;
  begin
    Result := StackContains(bkExcept);
  end;

  /// <summary>Pre-scans the file text BEFORE the selection to determine
  ///  the innermost open structural block (begin / case / try / repeat)
  ///  at the selection's starting line. Returns True if a 'case ... of'
  ///  is the closest enclosing block, meaning the selection sits between
  ///  case branches - extracting a single branch would break the case.
  ///  A 'begin ... end' inside a branch pushes onto the stack above
  ///  'case', so this correctly returns False for a begin-block that
  ///  happens to be inside a branch.</summary>
  function IsSelectionInsideCase: Boolean;
  const
    skBegin  = 0;
    skCase   = 1;
    skTry    = 2;
    skRepeat = 3;
  var
    Stack: TStack<Integer>;
    PreText, Tok: string;
    L: Integer;
    PreScanner: TPascalScanner;
  begin
    Result := False;
    if Length(AFileLines) = 0 then Exit;
    if AStartLine <= 1 then Exit;

    Stack := TStack<Integer>.Create;
    try
      // Concatenate every line BEFORE the selection (AStartLine is 1-based)
      PreText := '';
      for L := 0 to Min(AStartLine - 2, High(AFileLines)) do
        PreText := PreText + AFileLines[L] + #10;

      PreScanner := TPascalScanner.Create(PreText);
      try
        while PreScanner.NextToken(Tok) do
        begin
          if Tok = 'BEGIN' then Stack.Push(skBegin)
          else if Tok = 'CASE' then Stack.Push(skCase)
          else if Tok = 'TRY' then Stack.Push(skTry)
          else if Tok = 'REPEAT' then Stack.Push(skRepeat)
          else if Tok = 'END' then
          begin
            // 'end' closes the topmost begin/case/try (not repeat).
            if Stack.Count > 0 then Stack.Pop;
          end
          else if Tok = 'UNTIL' then
          begin
            if (Stack.Count > 0) and (Stack.Peek = skRepeat) then
              Stack.Pop;
          end;
          // EXCEPT / FINALLY do NOT pop - the try block continues until 'end'.
        end;
      finally
        PreScanner.Free;
      end;

      Result := (Stack.Count > 0) and (Stack.Peek = skCase);
    finally
      Stack.Free;
    end;
  end;
begin
  Issues := TList<TValidationIssue>.Create;
  try
    // 1. Check selection bounds
    if (AStartLine < 1) or (AEndLine < AStartLine) then
    begin
      AddError('Invalid selection range.');
      Result.Issues := Issues.ToArray;
      Exit;
    end;

    // 2. Enclosing method found?
    if AInsertLine <= 0 then
      AddWarning('No enclosing method detected - block does not appear to be inside a procedure/function.');

    // 3. Block content must be non-trivial
    if Trim(ASelectedText) = '' then
    begin
      AddError('Selected block is empty.');
      Result.Issues := Issues.ToArray;
      Exit;
    end;

    // 4. Token analysis: structural balance
    BeginEndDepth := 0;
    TryDepth := 0;
    CaseDepth := 0;
    ParenDepth := 0;
    BracketDepth := 0;
    FirstToken := '';
    LastToken := '';
    TokenCount := 0;

    PendingThen := 0;
    PendingElse := 0;
    OrphanElseFound := False;
    OrphanUntilFound := False;
    RepeatDepth := 0;
    ForeignFinallyExceptFound := False;

    BlockStack := TList<Integer>.Create;
    LoopHeaderPending := False;
    ExpectLoopBody := False;
    ExitCount := 0;
    GotoCount := 0;
    BreakOutsideLoop := 0;
    ContinueOutsideLoop := 0;
    BareRaiseOutsideExcept := 0;
    PendingBareRaise := False;
    PendingBareRaiseInExcept := False;

    Scanner := TPascalScanner.Create(ASelectedText);
    try
      while Scanner.NextToken(Token) do
      begin
        Inc(TokenCount);
        if FirstToken = '' then FirstToken := Token;
        LastToken := Token;

        // Every block that "end" closes counts in BeginEndDepth: begin,
        // try, case and asm - counting only begin refused every complete
        // try..end and case..end
        if (Token = 'BEGIN') or (Token = 'ASM') then Inc(BeginEndDepth)
        else if Token = 'END' then Dec(BeginEndDepth)
        else if Token = 'TRY' then
        begin
          Inc(TryDepth);
          Inc(BeginEndDepth);
        end
        else if (Token = 'EXCEPT') or (Token = 'FINALLY') then
        begin
          if TryDepth > 0 then
            Dec(TryDepth)  // closes an open try
          else
            ForeignFinallyExceptFound := True;
        end
        else if Token = 'CASE' then
        begin
          Inc(CaseDepth);
          Inc(BeginEndDepth);
        end
        else if Token = 'IF' then
          Inc(PendingThen)
        else if Token = 'THEN' then
        begin
          if PendingThen > 0 then
          begin
            Dec(PendingThen);
            Inc(PendingElse); // this if-then allows an else
          end;
        end
        else if Token = 'ELSE' then
        begin
          if PendingElse > 0 then
            Dec(PendingElse)
          else if (CaseDepth = 0) then
            // else without if-then and not inside case -> orphan
            OrphanElseFound := True;
        end
        else if Token = 'REPEAT' then Inc(RepeatDepth)
        else if Token = 'UNTIL' then
        begin
          if RepeatDepth > 0 then
            Dec(RepeatDepth)
          else
            OrphanUntilFound := True;
        end
        else if Token = '(' then Inc(ParenDepth)
        else if Token = ')' then Dec(ParenDepth)
        else if Token = '[' then Inc(BracketDepth)
        else if Token = ']' then Dec(BracketDepth);

        // --- Control-flow state machine (independent of existing depth counters) ---

        // Close any pending bare-raise check: if the token after RAISE is a
        // statement terminator, the raise was bare ("raise;"); otherwise
        // it's "raise SomeException.Create(...)" which is always fine.
        if PendingBareRaise then
        begin
          if (Token = ';') or (Token = 'END') or (Token = 'ELSE') or
             (Token = 'UNTIL') or (Token = 'FINALLY') or (Token = 'EXCEPT') then
          begin
            if not PendingBareRaiseInExcept then
              Inc(BareRaiseOutsideExcept);
          end;
          PendingBareRaise := False;
        end;

        // Block-stack maintenance
        if Token = 'FOR' then LoopHeaderPending := True
        else if Token = 'WHILE' then LoopHeaderPending := True
        else if Token = 'DO' then
        begin
          if LoopHeaderPending then
          begin
            ExpectLoopBody := True;
            LoopHeaderPending := False;
          end;
        end
        else if Token = 'ASM' then BlockStack.Add(bkBegin)
        else if Token = 'BEGIN' then
        begin
          if ExpectLoopBody then
          begin
            BlockStack.Add(bkLoopBegin);
            ExpectLoopBody := False;
          end
          else
            BlockStack.Add(bkBegin);
        end
        else if Token = 'REPEAT' then BlockStack.Add(bkRepeat)
        else if Token = 'TRY' then BlockStack.Add(bkTry)
        else if Token = 'CASE' then BlockStack.Add(bkCase)
        else if Token = 'EXCEPT' then
        begin
          if (BlockStack.Count > 0) and (BlockStack.Last = bkTry) then
            BlockStack[BlockStack.Count - 1] := bkExcept;
        end
        else if Token = 'FINALLY' then
        begin
          if (BlockStack.Count > 0) and (BlockStack.Last = bkTry) then
            BlockStack[BlockStack.Count - 1] := bkFinally;
        end
        else if Token = 'END' then
        begin
          if BlockStack.Count > 0 then
          begin
            // the case is closed: an "else" after it is an orphan again
            if (BlockStack.Last = bkCase) and (CaseDepth > 0) then
              Dec(CaseDepth);
            BlockStack.Delete(BlockStack.Count - 1);
          end;
        end
        else if Token = 'UNTIL' then
        begin
          if (BlockStack.Count > 0) and (BlockStack.Last = bkRepeat) then
            BlockStack.Delete(BlockStack.Count - 1);
        end
        // Control-flow keywords
        else if Token = 'EXIT' then Inc(ExitCount)
        else if Token = 'BREAK' then
        begin
          if not InLoop then Inc(BreakOutsideLoop);
        end
        else if Token = 'CONTINUE' then
        begin
          if not InLoop then Inc(ContinueOutsideLoop);
        end
        else if Token = 'GOTO' then Inc(GotoCount)
        else if Token = 'RAISE' then
        begin
          PendingBareRaise := True;
          PendingBareRaiseInExcept := InExcept;
        end;

        // Cancel ExpectLoopBody if the token after DO isn't BEGIN
        // (single-statement loop body - we don't track inside it).
        if ExpectLoopBody and (Token <> 'DO') and (Token <> 'BEGIN') then
          ExpectLoopBody := False;
      end;

      // End-of-stream: a pending bare raise with nothing after is bare too.
      if PendingBareRaise and not PendingBareRaiseInExcept then
        Inc(BareRaiseOutsideExcept);
    finally
      Scanner.Free;
      BlockStack.Free;
    end;

    // Structural errors
    if OrphanElseFound then
      AddError('"else" without matching "if..then" inside the block - block is part of an if-statement.');
    if OrphanUntilFound then
      AddError('"until" without matching "repeat" inside the block - block is part of a repeat-statement.');
    if ForeignFinallyExceptFound then
      AddError('"except" or "finally" without matching "try" inside the block - block is part of a try-statement.');

    // Control-flow issues: extracting would break compile or change semantics.
    if ExitCount > 0 then
      AddError(Format('Selection contains %d "Exit" call(s) - extraction would change semantics ' +
        '(Exit would leave only the extracted method, not the original enclosing one).', [ExitCount]));
    if BreakOutsideLoop > 0 then
      AddError(Format('Selection contains %d "Break" without an enclosing loop inside the selection - ' +
        'would not compile after extraction.', [BreakOutsideLoop]));
    if ContinueOutsideLoop > 0 then
      AddError(Format('Selection contains %d "Continue" without an enclosing loop inside the selection - ' +
        'would not compile after extraction.', [ContinueOutsideLoop]));
    if BareRaiseOutsideExcept > 0 then
      AddError(Format('Selection contains %d bare "raise;" outside of an except-block inside the selection - ' +
        'would not compile after extraction.', [BareRaiseOutsideExcept]));
    if GotoCount > 0 then
      AddWarning(Format('Selection contains %d "goto" - label targets are not verified; ' +
        'ensure all referenced labels stay in scope after extraction.', [GotoCount]));

    // 5. Bracket balance
    if ParenDepth > 0 then
      AddError(Format('%d unclosed round bracket(s) "(" in the block.', [ParenDepth]))
    else if ParenDepth < 0 then
      AddError(Format('%d extra closing round bracket(s) ")" in the block.', [-ParenDepth]));

    if BracketDepth > 0 then
      AddError(Format('%d unclosed square bracket(s) "[" in the block.', [BracketDepth]))
    else if BracketDepth < 0 then
      AddError(Format('%d extra closing square bracket(s) "]" in the block.', [-BracketDepth]));

    // 6. Begin/End balance
    if BeginEndDepth > 0 then
      AddError(Format('%d "begin" without matching "end" in the block.', [BeginEndDepth]))
    else if BeginEndDepth < 0 then
      AddError(Format('%d "end" without matching "begin"/"try"/"case" in the block.',
        [-BeginEndDepth]));

    // 7. Pre-scan: is the selection inside a case statement's branches?
    // If so, extracting a single case branch would break the case
    // structure (the extracted method call would appear bare where a
    // case label is expected). The only exception: if the selection is
    // a balanced begin..end block starting with 'begin', it can be
    // extracted cleanly - the call replaces the begin..end in the branch.
    InsideCase := IsSelectionInsideCase;
    if InsideCase and (FirstToken <> 'BEGIN') then
      AddError('Selection is inside a case statement''s branches - ' +
        'a single case entry cannot be extracted as a method. ' +
        'Select the whole begin..end of a branch instead.');

    // 8. Isolated tokens at the start
    if (FirstToken = 'ELSE') then
      AddError('Block starts with "else" - that is part of an if-statement and should not be extracted on its own.');
    if (FirstToken = 'UNTIL') then
      AddError('Block starts with "until" - that is part of a repeat-statement.');
    if (FirstToken = 'EXCEPT') or (FirstToken = 'FINALLY') then
      AddError(Format('Block starts with "%s" - that is part of a try-statement.',
        [LowerCase(FirstToken)]));
    if (FirstToken = 'OF') then
      AddError('Block starts with "of" - that is part of a case-statement.');
    if (FirstToken = 'THEN') then
      AddError('Block starts with "then" - that is part of an if-statement.');
    if (FirstToken = 'DO') then
      AddError('Block starts with "do" - that is part of a for/while-statement.');

    // 8. If block contains a single expression (no statement), warn
    if (TokenCount < 2) then
      AddWarning('Block contains only a single token - probably not a complete statement.');

    // 9. Block should not cross method boundaries
    if Length(AFileLines) > 0 then
    begin
      var ProcCount := 0;
      for var I := AStartLine - 1 to AEndLine - 1 do
      begin
        if (I < 0) or (I >= Length(AFileLines)) then Continue;
        var Upper := UpperCase(Trim(AFileLines[I]));
        if Upper.StartsWith('PROCEDURE ') or Upper.StartsWith('FUNCTION ') or
           Upper.StartsWith('CONSTRUCTOR ') or Upper.StartsWith('DESTRUCTOR ') then
          Inc(ProcCount);
      end;
      if ProcCount > 0 then
        AddError('Block contains a procedure/function declaration - block crosses method boundaries.');
    end;

    // 10. Block should lie inside a method (between begin..end).
    //     Heuristic: search backwards from AStartLine for begin,
    //     forwards for end.
    if Length(AFileLines) > 0 then
    begin
      var FoundBegin := False;
      for var I := AStartLine - 2 downto 0 do
      begin
        if I >= Length(AFileLines) then Continue;
        var Upper := UpperCase(Trim(AFileLines[I]));
        if (Upper = 'BEGIN') or Upper.StartsWith('BEGIN ') then
        begin
          FoundBegin := True;
          Break;
        end;
        if Upper.StartsWith('PROCEDURE ') or Upper.StartsWith('FUNCTION ') or
           Upper.StartsWith('CONSTRUCTOR ') or Upper.StartsWith('DESTRUCTOR ') then
          Break;
      end;
      if not FoundBegin then
        AddWarning('No "begin" found before the block - block may be in the interface or type section.');
    end;

    // If no issues: OK message
    if Issues.Count = 0 then
      AddOk('Block is syntactically balanced and can be safely extracted.');

    Result.Issues := Issues.ToArray;
  finally
    Issues.Free;
  end;
end;


// ---------------------------------------------------------------------------
//  Written identifiers / method name (audit #39, L7f + L7e)
// ---------------------------------------------------------------------------

// Routines whose FIRST argument the compiler passes as var - an
// identifier handed to one of them is written, whatever the text looks
// like. Deliberately a short, certain list: the point is to stop calling
// such a variable "const", and a name that is not on the list still goes
// through the assignment tests below.
function VarParamRoutine(const AUpperName: string): Boolean;
begin
  Result := (AUpperName = 'INC') or (AUpperName = 'DEC') or
    (AUpperName = 'READ') or (AUpperName = 'READLN') or
    (AUpperName = 'VAL') or (AUpperName = 'FILLCHAR') or
    (AUpperName = 'SETLENGTH') or (AUpperName = 'SETSTRING') or
    (AUpperName = 'FREEANDNIL') or (AUpperName = 'MOVE') or
    (AUpperName = 'NEW') or (AUpperName = 'GETMEM') or
    (AUpperName = 'REALLOCMEM') or (AUpperName = 'FREEMEM') or
    (AUpperName = 'INSERT') or (AUpperName = 'DELETE') or
    (AUpperName = 'EXCLUDE') or (AUpperName = 'INCLUDE');
end;

// A statement can also start right after one of these, so
// "if Flag then Count := 1;" writes Count just as much as a line of its
// own does.
function StatementStarter(const AWord: string): Boolean;
begin
  Result := SameText(AWord, 'then') or SameText(AWord, 'else') or
    SameText(AWord, 'do') or SameText(AWord, 'begin') or
    SameText(AWord, 'repeat') or SameText(AWord, 'of') or
    SameText(AWord, 'try') or SameText(AWord, 'finally') or
    SameText(AWord, 'except') or SameText(AWord, 'var');
end;

function WrittenIdentifiersIn(const ALines: TArray<string>): TArray<string>;
var
  Masked: TArray<string>;
  Found: TStringList;

  procedure Note(const AName: string);
  begin
    if (AName <> '') and (Found.IndexOf(UpperCase(AName)) < 0) then
      Found.Add(UpperCase(AName));
  end;

begin
  Result := nil;
  if Length(ALines) = 0 then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  Found := TStringList.Create;
  try
    Found.Sorted := True;
    for var L := 0 to High(Masked) do
    begin
      var S := Masked[L];
      var Prev := '';        // the previous identifier on this line
      var PrevPrev := '';
      var I := 1;
      var AtBoundary := True;  // start of line = start of a statement
      while I <= Length(S) do
      begin
        if IsIdentStart(S[I]) then
        begin
          var Start := I;
          while (I <= Length(S)) and IsIdentChar(S[I]) do Inc(I);
          var W := Copy(S, Start, I - Start);
          // Skip blanks and look at what follows.
          var J := I;
          while (J <= Length(S)) and (S[J] = ' ') do Inc(J);
          // X := ... at a statement boundary, and "for X :=" / "for X in"
          if (J + 1 <= Length(S)) and (S[J] = ':') and (S[J + 1] = '=') then
          begin
            if AtBoundary or SameText(Prev, 'for') or SameText(Prev, 'with') then
              Note(W);
          end
          else if SameText(Prev, 'for') and SameText(W, 'var') then
            // "for var I := ..." - the name follows, handled next round
          else if SameText(PrevPrev, 'for') and SameText(Prev, 'var') then
            Note(W)
          else if SameText(Prev, 'for') and (J <= Length(S)) and
            (Copy(S, J, 3) = 'in ') then
            Note(W)
          else if (J <= Length(S)) and (S[J] = '(') and VarParamRoutine(UpperCase(W)) then
          begin
            // ... and the identifiers inside its argument list
            var Depth := 0;
            var K := J;
            while K <= Length(S) do
            begin
              if S[K] = '(' then Inc(Depth)
              else if S[K] = ')' then
              begin
                Dec(Depth);
                if Depth = 0 then Break;
              end
              else if IsIdentStart(S[K]) then
              begin
                var AS_ := K;
                while (K <= Length(S)) and IsIdentChar(S[K]) do Inc(K);
                Note(Copy(S, AS_, K - AS_));
                Continue;
              end;
              Inc(K);
            end;
          end;
          PrevPrev := Prev;
          Prev := W;
          AtBoundary := StatementStarter(W);
          Continue;
        end;
        if S[I] = ';' then
        begin
          AtBoundary := True;
          Prev := '';
          PrevPrev := '';
        end
        else if S[I] > ' ' then
          AtBoundary := False;
        Inc(I);
      end;
    end;
    Result := Found.ToStringArray;
  finally
    Found.Free;
  end;
end;

function MethodNameProblem(const AName: string): string;
begin
  if Trim(AName) = '' then Exit('Enter a method name.');
  if AName <> Trim(AName) then
    Exit('The method name must not start or end with a space.');
  if not IsIdentifier(AName) then
    Exit(Format('"%s" is not a valid Pascal identifier.', [AName]));
  if TIdentifierChecker.IsPascalKeyword(AName) then
    Exit(Format('"%s" is a reserved word.', [AName]));
end;

end.
