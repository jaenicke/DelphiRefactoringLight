(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.StatementRefactor;

// Two statement-level refactorings suggested in issue #11 (Ian Branch):
//
// EXTRACT VARIABLE - the selected expression becomes an inline variable
//   ("var LName := <expr>;") declared right before the STATEMENT that
//   contains it, and the selection is replaced by the name. Never at the
//   routine's begin: hoisting further would change how often and whether
//   the expression is evaluated. Refused where even the statement start is
//   too far: after a short-circuit operator ("if Assigned(X) and (X.Foo >
//   0)" - hoisting X.Foo would dereference nil), inside a branch or loop
//   body on the same line (then/else/do), in a loop condition, as the sole
//   statement of a branch (inserting before it would change the branch),
//   and inside a with (the names might bind to the with target).
//
// WRAP IN TRY..FINALLY - the selected statements go into a try block; the
//   cleanup is inferred from the statement right before them:
//   "X := TFoo.Create..." -> X.Free, "X.BeginUpdate" -> X.EndUpdate,
//   "X.Enter/Acquire/Lock" -> Leave/Release/Unlock, "TMonitor.Enter(X)" ->
//   TMonitor.Exit(X); otherwise a TODO comment. Only wrapper lines are
//   added, so a wrong guess is a compile error, never silent damage.
//
// The planners are pure (tested); the two entry points read the selection
// through the editor helper and write through ApplyLinesMinimal (undoable
// in the IDE).

interface

uses
  System.SysUtils;

type
  TExtractVarPlan = record
    StatementLine: Integer;   // 0-based line the declaration goes before
    DeclText: string;         // "  var LCount := Foo.Bar.Count;"
    NewLines: TArray<string>; // the whole unit after the edit
  end;

/// <summary>A readable name for a variable holding AExpr: the last
///  identifier before a call / index ("Foo.Bar.Count" -> "LCount",
///  "GetName(X)" -> "LName", "Items[I].Caption" -> "LCaption"), "LValue"
///  when there is none.</summary>
function SuggestVariableName(const AExpr: string): string;

/// <summary>Plans "extract variable" for the single-line selection
///  [AStartCol0, AEndCol0) on ALine0 (0-based, end exclusive). False with
///  AWhy when the extraction would not be safe (see the unit header).</summary>
function PlanExtractVariable(const ALines: TArray<string>; ALine0, AStartCol0,
  AEndCol0: Integer; const AName: string; out APlan: TExtractVarPlan;
  out AWhy: string): Boolean;

/// <summary>The cleanup statement for a try..finally around statements
///  that follow APrevStatement ('' when nothing can be inferred).</summary>
function InferCleanup(const APrevStatement: string): string;

/// <summary>ALines with the statements on lines AFirst0..ALast0 wrapped in
///  try..finally (cleanup = ACleanup, or a TODO comment when empty). False
///  with AWhy when the lines are not a sequence of complete statements
///  (unbalanced begin/end, try/end, case/end, parentheses).</summary>
function PlanWrapTryFinally(const ALines: TArray<string>; AFirst0, ALast0: Integer;
  const ACleanup: string; out ANewLines: TArray<string>; out AWhy: string): Boolean;

/// <summary>Editor entry points (IDE and standalone).</summary>
procedure ExtractVariableAtSelection;
procedure WrapSelectionInTryFinally;

implementation

uses
  System.Classes, System.StrUtils, System.Math, Vcl.Dialogs,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.UsesEditor,
  Expert.AutoImport, Expert.WithScanner, Expert.DialogHelper;

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
end;

function LeadingBlanks(const S: string): string;
var
  K: Integer;
begin
  K := 1;
  while (K <= Length(S)) and CharInSet(S[K], [' ', #9]) do Inc(K);
  Result := Copy(S, 1, K - 1);
end;

function HasWordCI(const S, AWord: string): Boolean;
var
  U, W: string;
  P, E: Integer;
begin
  U := UpperCase(S);
  W := UpperCase(AWord);
  P := Pos(W, U);
  while P > 0 do
  begin
    E := P + Length(W);
    if ((P = 1) or not IsIdentChar(U[P - 1])) and ((E > Length(U)) or not IsIdentChar(U[E])) then
      Exit(True);
    P := Pos(W, U, P + 1);
  end;
  Result := False;
end;

function SuggestVariableName(const AExpr: string): string;
var
  S, Id: string;
  I, Depth: Integer;
begin
  S := Trim(AExpr);
  // drop trailing (...) / [...] groups: "GetName(X)" -> "GetName"
  while (S <> '') and CharInSet(S[Length(S)], [')', ']']) do
  begin
    Depth := 0;
    I := Length(S);
    while I >= 1 do
    begin
      if CharInSet(S[I], [')', ']']) then Inc(Depth)
      else if CharInSet(S[I], ['(', '[']) then
      begin
        Dec(Depth);
        if Depth = 0 then Break;
      end;
      Dec(I);
    end;
    if I < 1 then Break;
    S := TrimRight(Copy(S, 1, I - 1));
  end;
  // last identifier
  I := Length(S);
  while (I >= 1) and IsIdentChar(S[I]) do Dec(I);
  Id := Copy(S, I + 1, MaxInt);
  if (Id = '') or CharInSet(Id[1], ['0'..'9']) then Exit('LValue');
  if (Length(Id) > 3) and StartsText('Get', Id) and CharInSet(Id[4], ['A'..'Z']) then
    Id := Copy(Id, 4, MaxInt);
  if (Length(Id) > 1) and CharInSet(Id[1], ['F', 'A']) and CharInSet(Id[2], ['A'..'Z']) then
    Id := Copy(Id, 2, MaxInt);   // FCount / ACount -> Count
  Result := 'L' + UpCase(Id[1]) + Copy(Id, 2, MaxInt);
end;

// Last code token of a masked line ('' for a blank one).
function LastCodeWord(const AMasked: string; out ALastChar: Char): string;
var
  S: string;
  I: Integer;
begin
  S := TrimRight(AMasked);
  ALastChar := #0;
  Result := '';
  if S = '' then Exit;
  ALastChar := S[Length(S)];
  I := Length(S);
  while (I >= 1) and IsIdentChar(S[I]) do Dec(I);
  Result := UpperCase(Copy(S, I + 1, MaxInt));
end;

// 0-based line where the statement containing ALine0 starts: the line after
// the nearest code line that ends a statement or opens a statement list.
function StatementStartLine(const AMasked: TArray<string>; ALine0: Integer): Integer;
var
  W: string;
  C: Char;
begin
  Result := ALine0;
  for var L := ALine0 - 1 downto 0 do
  begin
    W := LastCodeWord(AMasked[L], C);
    if C = #0 then Continue;   // blank / comment-only line
    if (C = ';') or (W = 'BEGIN') or (W = 'THEN') or (W = 'DO') or (W = 'ELSE')
      or (W = 'TRY') or (W = 'FINALLY') or (W = 'EXCEPT') or (W = 'REPEAT')
      or (W = 'OF') or ((C = ':') and not AMasked[L].TrimRight.EndsWith(':=')) then
      Exit(L + 1);
    Result := L;   // continuation of the statement
  end;
end;

function PlanExtractVariable(const ALines: TArray<string>; ALine0, AStartCol0,
  AEndCol0: Integer; const AName: string; out APlan: TExtractVarPlan;
  out AWhy: string): Boolean;
var
  M: TArray<string>;
  Sel, Before, FirstWord: string;
  Depth, StmtLine, I: Integer;
begin
  Result := False;
  APlan := Default(TExtractVarPlan);
  AWhy := '';
  if (ALine0 < 0) or (ALine0 > High(ALines)) or (AStartCol0 < 0)
    or (AEndCol0 <= AStartCol0) or (AEndCol0 > Length(ALines[ALine0])) then
  begin
    AWhy := 'select an expression on ONE line';
    Exit;
  end;
  if not IsValidIdent(AName) then
  begin
    AWhy := '"' + AName + '" is not a valid identifier';
    Exit;
  end;
  Sel := Copy(ALines[ALine0], AStartCol0 + 1, AEndCol0 - AStartCol0);
  if Trim(Sel) = '' then begin AWhy := 'the selection is empty'; Exit; end;
  Sel := Trim(Sel);
  M := MaskCommentsAndStrings(ALines);
  // the selection must be CODE at both ends (not inside a comment/string)
  // and balanced: brackets, and no statement separator inside
  var SelM := Copy(M[ALine0], AStartCol0 + 1, AEndCol0 - AStartCol0);
  // skip leading blanks of the selection; its first real character must be
  // code - or the opening quote of a string literal
  var FirstIdx := AStartCol0 + 1;
  while (FirstIdx <= AEndCol0) and CharInSet(ALines[ALine0][FirstIdx], [' ', #9]) do
    Inc(FirstIdx);
  if (M[ALine0][FirstIdx] <> ALines[ALine0][FirstIdx])
    and (ALines[ALine0][FirstIdx] <> '''') and (ALines[ALine0][FirstIdx] <> '#') then
  begin
    AWhy := 'the selection starts inside a comment or string';
    Exit;
  end;
  Depth := 0;
  for I := 1 to Length(SelM) do
    case SelM[I] of
      '(', '[': Inc(Depth);
      ')', ']':
        begin
          Dec(Depth);
          if Depth < 0 then Break;
        end;
      ';': begin AWhy := 'the selection contains a statement end ";"'; Exit; end;
    end;
  if Depth <> 0 then
  begin
    AWhy := 'the selection is not a complete expression (unbalanced brackets)';
    Exit;
  end;
  if Sel.Contains(':=') then
  begin
    AWhy := 'the selection contains an assignment - select an expression';
    Exit;
  end;

  StmtLine := StatementStartLine(M, ALine0);
  if IsSoleBranchStatement(ALines, StmtLine) then
  begin
    AWhy := 'the statement is the only one of a then/else/do branch - put it ' +
      'into begin..end first (a declaration before it would leave the branch)';
    Exit;
  end;
  // everything of the statement in front of the selection
  Before := '';
  for I := StmtLine to ALine0 - 1 do Before := Before + ' ' + M[I];
  Before := Before + ' ' + Copy(M[ALine0], 1, AStartCol0);
  FirstWord := '';
  var T := TrimLeft(Before);
  I := 1;
  while (I <= Length(T)) and IsIdentChar(T[I]) do Inc(I);
  FirstWord := UpperCase(Copy(T, 1, I - 1));
  // "1: Foo(X)" - a case branch or label: no declaration can go before it
  var ColonP := Pos(':', T);
  if (ColonP > 0) and ((ColonP = Length(T)) or (T[ColonP + 1] <> '=')) then
  begin
    AWhy := 'the statement is a case branch or carries a label - put it into ' +
      'begin..end first';
    Exit;
  end;
  if (FirstWord = 'WHILE') or (FirstWord = 'UNTIL') then
  begin
    AWhy := 'the expression is part of a loop condition - it is evaluated on ' +
      'every iteration';
    Exit;
  end;
  for var W in ['AND', 'OR', 'THEN', 'ELSE', 'DO', 'OF', 'EXCEPT', 'FINALLY'] do
    if HasWordCI(Before, W) then
    begin
      if (W = 'AND') or (W = 'OR') then
        AWhy := 'the expression follows a short-circuit "' + LowerCase(W) + '" - ' +
          'hoisting it would evaluate it unconditionally'
      else
        AWhy := 'the expression sits in a branch or loop body on the same line ' +
          '("' + LowerCase(W) + '") - put that part on its own line first';
      Exit;
    end;
  // inside a with? its names may bind to the with target
  var Src := string.Join(sLineBreak, ALines);
  for var Occ in TWithScanner.ScanSource(Src) do
    if (ALine0 + 1 >= Occ.BodyRange.StartPos.Line) and (ALine0 + 1 <= Occ.BodyRange.EndPos.Line) then
    begin
      AWhy := 'the selection is inside a "with" - remove the with first';
      Exit;
    end;
  // the name must be free in the routine
  var RF, RL: Integer;
  if FindEnclosingRoutineRange(Src, ALine0, RF, RL) then
  begin
    for var L in CodeWordLines(Copy(ALines, RF, RL - RF + 1), AName) do
    begin
      AWhy := Format('"%s" is already used in this routine (line %d)', [AName, RF + L + 1]);
      Exit;
    end;
  end
  else
  begin
    AWhy := 'the selection is not inside a routine';
    Exit;
  end;

  APlan.StatementLine := StmtLine;
  APlan.DeclText := LeadingBlanks(ALines[StmtLine]) + 'var ' + AName + ' := ' + Sel + ';';
  SetLength(APlan.NewLines, Length(ALines) + 1);
  for I := 0 to StmtLine - 1 do APlan.NewLines[I] := ALines[I];
  APlan.NewLines[StmtLine] := APlan.DeclText;
  for I := StmtLine to High(ALines) do APlan.NewLines[I + 1] := ALines[I];
  APlan.NewLines[ALine0 + 1] := Copy(ALines[ALine0], 1, AStartCol0) + AName +
    Copy(ALines[ALine0], AEndCol0 + 1, MaxInt);
  Result := True;
end;

function InferCleanup(const APrevStatement: string): string;
var
  S, U, Target: string;
  P: Integer;
begin
  Result := '';
  S := Trim(APrevStatement);
  while S.EndsWith(';') do S := TrimRight(Copy(S, 1, Length(S) - 1));
  if S = '' then Exit;
  U := UpperCase(S);
  // X := TFoo.Create(...) / X := TFoo.CreateFmt(...)
  P := Pos(':=', S);
  if P > 1 then
  begin
    Target := Trim(Copy(S, 1, P - 1));
    if StartsText('VAR ', UpperCase(Target)) then Target := Trim(Copy(Target, 5, MaxInt));
    var Colon := Pos(':', Target);
    if Colon > 0 then Target := Trim(Copy(Target, 1, Colon - 1));
    var Rhs := UpperCase(Trim(Copy(S, P + 2, MaxInt)));
    if IsValidIdent(Target, True) and (Pos('.CREATE', Rhs) > 0) then
      Exit(Target + '.Free');
    Exit;
  end;
  // TMonitor.Enter(X)
  if StartsText('TMONITOR.ENTER(', U) then
    Exit('TMonitor.Exit(' + Copy(S, Length('TMonitor.Enter(') + 1, MaxInt));
  // X.BeginUpdate / X.Enter / X.Acquire / X.Lock / X.BeginWrite / X.BeginRead
  P := LastDelimiter('.', S);
  if P > 1 then
  begin
    Target := Copy(S, 1, P - 1);
    var Meth := UpperCase(Copy(S, P + 1, MaxInt));
    if not IsValidIdent(Target, True) then Exit;
    if Meth = 'BEGINUPDATE' then Exit(Target + '.EndUpdate');
    if Meth = 'ENTER' then Exit(Target + '.Leave');
    if Meth = 'ACQUIRE' then Exit(Target + '.Release');
    if Meth = 'LOCK' then Exit(Target + '.Unlock');
    if Meth = 'BEGINWRITE' then Exit(Target + '.EndWrite');
    if Meth = 'BEGINREAD' then Exit(Target + '.EndRead');
  end;
end;

function PlanWrapTryFinally(const ALines: TArray<string>; AFirst0, ALast0: Integer;
  const ACleanup: string; out ANewLines: TArray<string>; out AWhy: string): Boolean;
var
  M: TArray<string>;
  Blocks, Parens, I: Integer;
  Indent, Cleanup: string;
  Res: TArray<string>;
begin
  Result := False;
  ANewLines := nil;
  AWhy := '';
  if (AFirst0 < 0) or (ALast0 > High(ALines)) or (ALast0 < AFirst0) then
  begin
    AWhy := 'select the statements to wrap';
    Exit;
  end;
  M := MaskCommentsAndStrings(ALines);
  // complete statements: block openers and ends balance, parentheses too,
  // and the last line ends a statement
  Blocks := 0;
  Parens := 0;
  for I := AFirst0 to ALast0 do
  begin
    var U := UpperCase(M[I]);
    var K := 1;
    while K <= Length(U) do
    begin
      if CharInSet(U[K], ['A'..'Z', '_']) and ((K = 1) or not IsIdentChar(U[K - 1])) then
      begin
        var S := K;
        while (K <= Length(U)) and IsIdentChar(U[K]) do Inc(K);
        var W := Copy(U, S, K - S);
        if (W = 'BEGIN') or (W = 'TRY') or (W = 'CASE') or (W = 'ASM') then Inc(Blocks)
        else if W = 'END' then Dec(Blocks);
        if Blocks < 0 then Break;
        Continue;
      end;
      if U[K] = '(' then Inc(Parens) else if U[K] = ')' then Dec(Parens);
      Inc(K);
    end;
    if Blocks < 0 then Break;
  end;
  if (Blocks <> 0) or (Parens <> 0) then
  begin
    AWhy := 'the selected lines are not complete statements (unbalanced ' +
      'begin/try/case..end or brackets)';
    Exit;
  end;
  var LastChar: Char;
  var LastW := LastCodeWord(M[ALast0], LastChar);
  if not ((LastChar = ';') or (LastW = 'END')) then
  begin
    AWhy := 'the selection does not end with a complete statement';
    Exit;
  end;
  Indent := LeadingBlanks(ALines[AFirst0]);
  Cleanup := ACleanup;
  if Cleanup = '' then Cleanup := '// TODO: release what the block acquired'
  else if not Cleanup.EndsWith(';') then Cleanup := Cleanup + ';';
  Res := Copy(ALines, 0, AFirst0);
  Res := Res + [Indent + 'try'];
  for I := AFirst0 to ALast0 do
    if Trim(ALines[I]) = '' then Res := Res + [ALines[I]]
    else Res := Res + ['  ' + ALines[I]];
  Res := Res + [Indent + 'finally', Indent + '  ' + Cleanup, Indent + 'end;'];
  Res := Res + Copy(ALines, ALast0 + 1, MaxInt);
  ANewLines := Res;
  Result := True;
end;

// ---------------------------------------------------------------------------
//  Editor entry points
// ---------------------------------------------------------------------------

function ReadLinesOf(const AFile: string; out AContent: string;
  out ALines: TArray<string>): Boolean;
var
  SL: TStringList;
begin
  Result := (Editor <> nil) and Editor.ReadEditorContent(AFile, AContent);
  if not Result then Exit;
  SL := TStringList.Create;
  try
    SL.Text := AContent;
    ALines := SL.ToStringArray;
  finally
    SL.Free;
  end;
end;

function WriteLines(const AFile, AOriginal: string; const ALines: TArray<string>): Boolean;
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    for var L in ALines do SL.Add(L);
    Result := ApplyLinesMinimal(AFile, SL, AOriginal);
  finally
    SL.Free;
  end;
end;

procedure ExtractVariableAtSelection;
var
  F, Text, Content, Name, Why: string;
  SL, SC, EL, EC: Integer;
  Lines: TArray<string>;
  Plan: TExtractVarPlan;
begin
  if (Editor = nil) or not Editor.GetSelection(F, SL, SC, EL, EC, Text) or (Trim(Text) = '') then
  begin
    ShowThemedMessage('Extract variable: select the expression first.');
    Exit;
  end;
  if (EL <> SL) or Text.Contains(#10) then
  begin
    ShowThemedMessage('Extract variable: select an expression on ONE line.');
    Exit;
  end;
  if not ReadLinesOf(F, Content, Lines) or (SL < 1) or (SL > Length(Lines)) then
  begin
    ShowThemedMessage('Extract variable: the file is not open in the editor.');
    Exit;
  end;
  // The selection's columns are editor columns; trust the selected TEXT and
  // find it on the line at (or right next to) the reported start.
  var Col0 := SC - 1;
  var LineText := Lines[SL - 1];
  if Copy(LineText, Col0 + 1, Length(Text)) <> Text then
  begin
    var P := Pos(Text, LineText, Max(1, Col0 - 1));
    if P = 0 then P := Pos(Text, LineText);
    if P = 0 then
    begin
      ShowThemedMessage('Extract variable: the selection could not be located on the line.');
      Exit;
    end;
    Col0 := P - 1;
  end;
  Name := SuggestVariableName(Text);
  if not AskThemedText('Extract variable', 'Name of the new variable (inline "var"):',
    Name,
    function(S: string): string
    begin
      if IsValidIdent(Trim(S)) then Result := '' else Result := 'not a valid identifier';
    end) then Exit;
  Name := Trim(Name);
  if not PlanExtractVariable(Lines, SL - 1, Col0, Col0 + Length(Text), Name, Plan, Why) then
  begin
    ShowThemedMessage('Extract variable not possible: ' + Why + '.');
    Exit;
  end;
  if not WriteLines(F, Content, Plan.NewLines) then
    ShowThemedMessage('Extract variable: writing the change failed.');
end;

procedure WrapSelectionInTryFinally;
var
  F, Text, Content, Why, Cleanup: string;
  SL, SC, EL, EC: Integer;
  Lines, NewLines: TArray<string>;
begin
  if (Editor = nil) or not Editor.GetSelection(F, SL, SC, EL, EC, Text) or (Trim(Text) = '') then
  begin
    ShowThemedMessage('Wrap in try..finally: select the statements first.');
    Exit;
  end;
  if EC = 1 then Dec(EL);   // a selection ending at a line start excludes that line
  if not ReadLinesOf(F, Content, Lines) or (SL < 1) or (EL > Length(Lines)) then
  begin
    ShowThemedMessage('Wrap in try..finally: the file is not open in the editor.');
    Exit;
  end;
  // the statement before the selection names what to release
  Cleanup := '';
  for var L := SL - 2 downto 0 do
    if Trim(Lines[L]) <> '' then
    begin
      Cleanup := InferCleanup(Lines[L]);
      Break;
    end;
  if not PlanWrapTryFinally(Lines, SL - 1, EL - 1, Cleanup, NewLines, Why) then
  begin
    ShowThemedMessage('Wrap in try..finally not possible: ' + Why + '.');
    Exit;
  end;
  if not WriteLines(F, Content, NewLines) then
    ShowThemedMessage('Wrap in try..finally: writing the change failed.');
end;

end.
