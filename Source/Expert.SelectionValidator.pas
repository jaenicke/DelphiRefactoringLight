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

implementation

uses
  System.StrUtils;

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

{ Helper functions for tokenising Pascal code }

type
  TPascalScanner = class
  private
    FText: string;
    FPos: Integer;
    function IsAtComment: Boolean;
    function IsAtString: Boolean;
    procedure SkipComment;
    procedure SkipString;
    procedure SkipWhitespace;
  public
    constructor Create(const AText: string);
    /// <summary>Returns the next Pascal token (identifier or special
    ///  character). Skips comments and strings.</summary>
    function NextToken(out AToken: string): Boolean;
  end;

constructor TPascalScanner.Create(const AText: string);
begin
  FText := AText;
  FPos := 1;
end;

function TPascalScanner.IsAtComment: Boolean;
begin
  Result := False;
  if FPos > Length(FText) then Exit;
  if (FText[FPos] = '/') and (FPos < Length(FText)) and (FText[FPos+1] = '/') then
    Exit(True);
  if FText[FPos] = '{' then Exit(True);
  if (FText[FPos] = '(') and (FPos < Length(FText)) and (FText[FPos+1] = '*') then
    Exit(True);
end;

function TPascalScanner.IsAtString: Boolean;
begin
  Result := (FPos <= Length(FText)) and (FText[FPos] = '''');
end;

procedure TPascalScanner.SkipComment;
begin
  if FPos > Length(FText) then Exit;
  if (FText[FPos] = '/') and (FPos < Length(FText)) and (FText[FPos+1] = '/') then
  begin
    // Until end of line
    while (FPos <= Length(FText)) and not CharInSet(FText[FPos], [#10, #13]) do
      Inc(FPos);
  end
  else if FText[FPos] = '{' then
  begin
    Inc(FPos);
    while (FPos <= Length(FText)) and (FText[FPos] <> '}') do
      Inc(FPos);
    if FPos <= Length(FText) then Inc(FPos);
  end
  else if (FText[FPos] = '(') and (FPos < Length(FText)) and (FText[FPos+1] = '*') then
  begin
    Inc(FPos, 2);
    while FPos < Length(FText) do
    begin
      if (FText[FPos] = '*') and (FText[FPos+1] = ')') then
      begin
        Inc(FPos, 2);
        Exit;
      end;
      Inc(FPos);
    end;
  end;
end;

procedure TPascalScanner.SkipString;
begin
  if (FPos > Length(FText)) or (FText[FPos] <> '''') then Exit;
  Inc(FPos);
  while FPos <= Length(FText) do
  begin
    if FText[FPos] = '''' then
    begin
      // Double apostrophe = escape
      if (FPos < Length(FText)) and (FText[FPos+1] = '''') then
        Inc(FPos, 2)
      else
      begin
        Inc(FPos);
        Exit;
      end;
    end
    else
      Inc(FPos);
  end;
end;

procedure TPascalScanner.SkipWhitespace;
begin
  while (FPos <= Length(FText)) and CharInSet(FText[FPos], [' ', #9, #10, #13]) do
    Inc(FPos);
end;

function TPascalScanner.NextToken(out AToken: string): Boolean;
begin
  AToken := '';
  while FPos <= Length(FText) do
  begin
    SkipWhitespace;
    if FPos > Length(FText) then Exit(False);

    if IsAtComment then
    begin
      SkipComment;
      Continue;
    end;

    if IsAtString then
    begin
      SkipString;
      Continue;
    end;

    // Identifier or keyword
    if CharInSet(FText[FPos], ['A'..'Z', 'a'..'z', '_']) then
    begin
      var Start := FPos;
      while (FPos <= Length(FText)) and
            CharInSet(FText[FPos], ['A'..'Z', 'a'..'z', '0'..'9', '_']) do
        Inc(FPos);
      AToken := UpperCase(Copy(FText, Start, FPos - Start));
      Exit(True);
    end;

    // Special characters
    if CharInSet(FText[FPos], ['(', ')', '[', ']', ';', '.']) then
    begin
      AToken := FText[FPos];
      Inc(FPos);
      Exit(True);
    end;

    // Number or other - skip
    Inc(FPos);
  end;
  Result := False;
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

    Scanner := TPascalScanner.Create(ASelectedText);
    try
      while Scanner.NextToken(Token) do
      begin
        Inc(TokenCount);
        if FirstToken = '' then FirstToken := Token;
        LastToken := Token;

        if Token = 'BEGIN' then Inc(BeginEndDepth)
        else if Token = 'END' then Dec(BeginEndDepth)
        else if Token = 'TRY' then Inc(TryDepth)
        else if (Token = 'EXCEPT') or (Token = 'FINALLY') then
        begin
          if TryDepth > 0 then
            Dec(TryDepth)  // closes an open try
          else
            ForeignFinallyExceptFound := True;
        end
        else if Token = 'CASE' then Inc(CaseDepth)
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
      end;
    finally
      Scanner.Free;
    end;

    // Structural errors
    if OrphanElseFound then
      AddError('"else" without matching "if..then" inside the block - block is part of an if-statement.');
    if OrphanUntilFound then
      AddError('"until" without matching "repeat" inside the block - block is part of a repeat-statement.');
    if ForeignFinallyExceptFound then
      AddError('"except" or "finally" without matching "try" inside the block - block is part of a try-statement.');

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

end.
