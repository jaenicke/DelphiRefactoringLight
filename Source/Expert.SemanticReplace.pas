(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SemanticReplace;

// Engine for the "Semantic Replace" workflow.
//
// Reads a list of find/replace rules from
// <project_root>/semantic-replace.json and applies them to project
// source files with:
//
//   - identifier-boundary matching (Find is matched only at
//     whole-token positions),
//   - comment- and string-aware scanning (matches inside //, brace
//     comments, paren-star comments, and string literals are ignored),
//   - automatic uses-clause augmentation: each rule lists units that
//     the Replace expression depends on; if a file ends up with at
//     least one edit, the wizard makes sure those units are present
//     in the file's interface uses,
//   - optional local-variable hoisting: when a rule fires more than
//     once inside the same routine AND the rule defines a localVar
//     triple (name, type, value), the wizard introduces
//        var Name: Type := Value;
//     right after the routine's BEGIN and uses ReplaceWhenLocalVar
//     instead of Replace for each occurrence inside that routine.

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.IOUtils,
  System.Character,
  System.Generics.Collections, System.Generics.Defaults, System.StrUtils;

type
  TSemanticReplaceRule = record
    Find: string;
    Replace: string;
    UsesToAdd: TArray<string>;
    /// <summary>Optional. When set, and the rule fires more than once
    ///  inside the same routine body, the wizard introduces
    ///  `var <Name>: <Type> := <Value>;` right after BEGIN and uses
    ///  ReplaceWhenLocalVar for each occurrence instead of Replace.</summary>
    LocalVarName: string;
    LocalVarType: string;
    LocalVarValue: string;
    ReplaceWhenLocalVar: string;
  end;

  TSemanticReplaceMatch = record
    RuleIdx: Integer;
    /// <summary>1-based character offset in the original text.</summary>
    Offset: Integer;
    /// <summary>Number of characters of the original token (= Length(Find)).</summary>
    Length: Integer;
  end;

  TSemanticReplaceStats = record
    /// <summary>Total number of textual replacements (occurrences) made.</summary>
    Occurrences: Integer;
    /// <summary>Number of routines where a local-var was hoisted because
    ///  a rule fired >= 2 times inside that routine.</summary>
    LocalVarsIntroduced: Integer;
    /// <summary>Rule indices that fired at least once.</summary>
    RuleHits: TArray<Integer>;
  end;

  TSemanticReplaceEngine = class
  public
    /// <summary>Reads rules from a JSON file. Returns an empty array
    ///  when the file does not exist. AError gets a human-readable
    ///  diagnostic on parse errors (and the result stays empty).</summary>
    class function LoadRules(const AJsonPath: string;
      out AError: string): TArray<TSemanticReplaceRule>;

    /// <summary>Writes a starter file with one fully-commented example
    ///  rule. Used when the user runs the wizard against a project
    ///  that has no rules file yet.</summary>
    class procedure WriteExampleRules(const AJsonPath: string);

    /// <summary>Applies all rules to AOriginal and returns the modified
    ///  text. Comment- and string-aware. Performs local-var hoisting
    ///  when a rule has a LocalVar* triple and fires >= 2 times inside
    ///  the same routine. Sets AStats with totals and the per-rule
    ///  hit list.</summary>
    class function ApplyToText(const AOriginal: string;
      const ARules: TArray<TSemanticReplaceRule>;
      out AStats: TSemanticReplaceStats): string;

    /// <summary>Writes the given rules back to a JSON file using the
    ///  same shape LoadRules expects.</summary>
    class procedure SaveRules(const AJsonPath: string;
      const ARules: TArray<TSemanticReplaceRule>);

    /// <summary>Returns a one-based (line, column) tuple for the given
    ///  1-based character offset inside AText. Used by preview formatting.</summary>
    class procedure OffsetToLineCol(const AText: string; AOffset: Integer;
      out ALine, ACol: Integer);

    /// <summary>Returns the entire line (without newline) containing
    ///  the 1-based AOffset.</summary>
    class function LineAtOffset(const AText: string; AOffset: Integer): string;

    /// <summary>Comment/string-aware scan that returns every whole-
    ///  identifier match of any rule's Find. Used by the preview to
    ///  list per-match before/after lines.</summary>
    class function FindAllMatches(const AText: string;
      const ARules: TArray<TSemanticReplaceRule>): TArray<TSemanticReplaceMatch>;
  end;

implementation

const
  CExampleJson =
    '{' + sLineBreak +
    '  "_comment": "Semantic replacements applied by Refactoring Light. ' +
    'Each rule rewrites a sequence of dotted identifiers; if `uses` is set, ' +
    'those units are added to any file that ends up with at least one edit. ' +
    'When a rule fires more than once inside the same routine AND a ' +
    '`localVar` block is set, a local var is hoisted right after BEGIN ' +
    'and `localVar.replace` is used inside that routine instead of `replace`.",' + sLineBreak +
    '  "rules": [' + sLineBreak +
    '    {' + sLineBreak +
    '      "find":    "Manager.Config.WriteConfig",' + sLineBreak +
    '      "replace": "TAppCentral.Get<IConfig>.WriteConfig",' + sLineBreak +
    '      "uses":    ["AppCentral.Core", "AppCentral.Config"],' + sLineBreak +
    '      "localVar": {' + sLineBreak +
    '        "name":    "LConfig",' + sLineBreak +
    '        "type":    "IConfig",' + sLineBreak +
    '        "value":   "TAppCentral.Get<IConfig>",' + sLineBreak +
    '        "replace": "LConfig.WriteConfig"' + sLineBreak +
    '      }' + sLineBreak +
    '    }' + sLineBreak +
    '  ]' + sLineBreak +
    '}' + sLineBreak;

class function TSemanticReplaceEngine.LoadRules(const AJsonPath: string;
  out AError: string): TArray<TSemanticReplaceRule>;
var
  Json: TJSONValue;
  Arr: TJSONArray;
  Item: TJSONValue;
  Obj, LocalObj: TJSONObject;
  UsesVal: TJSONValue;
  Rule: TSemanticReplaceRule;
  List: TList<TSemanticReplaceRule>;
begin
  Result := nil;
  AError := '';
  if not TFile.Exists(AJsonPath) then Exit;
  try
    Json := TJSONObject.ParseJSONValue(TFile.ReadAllText(AJsonPath));
  except
    on E: Exception do
    begin
      AError := 'I/O error: ' + E.Message; Exit;
    end;
  end;
  if Json = nil then
  begin
    AError := 'Not valid JSON.'; Exit;
  end;
  List := TList<TSemanticReplaceRule>.Create;
  try
    Arr := nil;
    if Json is TJSONObject then
      Arr := TJSONObject(Json).GetValue<TJSONArray>('rules', nil)
    else if Json is TJSONArray then
      Arr := TJSONArray(Json);
    if Arr = nil then
    begin
      AError := 'Expected a "rules" array.'; Exit;
    end;
    for Item in Arr do
    begin
      if not (Item is TJSONObject) then Continue;
      Obj := TJSONObject(Item);
      Rule := Default(TSemanticReplaceRule);
      Rule.Find := Obj.GetValue<string>('find', '');
      Rule.Replace := Obj.GetValue<string>('replace', '');
      UsesVal := Obj.GetValue('uses');
      if UsesVal is TJSONArray then
        for var U in TJSONArray(UsesVal) do
          if U is TJSONString then
            Rule.UsesToAdd := Rule.UsesToAdd + [TJSONString(U).Value];
      LocalObj := nil;
      if Obj.TryGetValue<TJSONObject>('localVar', LocalObj) and (LocalObj <> nil) then
      begin
        Rule.LocalVarName := LocalObj.GetValue<string>('name', '');
        Rule.LocalVarType := LocalObj.GetValue<string>('type', '');
        Rule.LocalVarValue := LocalObj.GetValue<string>('value', '');
        Rule.ReplaceWhenLocalVar := LocalObj.GetValue<string>('replace', '');
      end;
      if Rule.Find <> '' then List.Add(Rule);
    end;
    Result := List.ToArray;
  finally
    List.Free;
    Json.Free;
  end;
end;

class procedure TSemanticReplaceEngine.WriteExampleRules(const AJsonPath: string);
begin
  TFile.WriteAllText(AJsonPath, CExampleJson, TEncoding.UTF8);
end;

class procedure TSemanticReplaceEngine.SaveRules(const AJsonPath: string;
  const ARules: TArray<TSemanticReplaceRule>);
var
  Root, Obj, LocalObj: TJSONObject;
  Arr, UsesArr: TJSONArray;
  R: TSemanticReplaceRule;
  U: string;
begin
  Root := TJSONObject.Create;
  Arr := TJSONArray.Create;
  try
    Root.AddPair('rules', Arr);
    for R in ARules do
    begin
      Obj := TJSONObject.Create;
      Obj.AddPair('find', R.Find);
      Obj.AddPair('replace', R.Replace);
      if Length(R.UsesToAdd) > 0 then
      begin
        UsesArr := TJSONArray.Create;
        for U in R.UsesToAdd do UsesArr.Add(U);
        Obj.AddPair('uses', UsesArr);
      end;
      if (R.LocalVarName <> '') or (R.LocalVarType <> '') or
         (R.LocalVarValue <> '') or (R.ReplaceWhenLocalVar <> '') then
      begin
        LocalObj := TJSONObject.Create;
        LocalObj.AddPair('name', R.LocalVarName);
        LocalObj.AddPair('type', R.LocalVarType);
        LocalObj.AddPair('value', R.LocalVarValue);
        LocalObj.AddPair('replace', R.ReplaceWhenLocalVar);
        Obj.AddPair('localVar', LocalObj);
      end;
      Arr.AddElement(Obj);
    end;
    TFile.WriteAllText(AJsonPath, Root.Format(2), TEncoding.UTF8);
  finally
    Root.Free;
  end;
end;

class procedure TSemanticReplaceEngine.OffsetToLineCol(const AText: string;
  AOffset: Integer; out ALine, ACol: Integer);
var
  I: Integer;
begin
  ALine := 1;
  ACol := 1;
  for I := 1 to AOffset - 1 do
  begin
    if I > Length(AText) then Break;
    if AText[I] = #10 then
    begin
      Inc(ALine); ACol := 1;
    end
    else if AText[I] <> #13 then
      Inc(ACol);
  end;
end;

class function TSemanticReplaceEngine.LineAtOffset(const AText: string;
  AOffset: Integer): string;
var
  S, E: Integer;
begin
  Result := '';
  if (AOffset < 1) or (AOffset > Length(AText)) then Exit;
  S := AOffset;
  while (S > 1) and (AText[S - 1] <> #10) and (AText[S - 1] <> #13) do Dec(S);
  E := AOffset;
  while (E <= Length(AText)) and (AText[E] <> #10) and (AText[E] <> #13) do Inc(E);
  Result := Copy(AText, S, E - S);
end;

{ ---- internals ---- }

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := C.IsLetterOrDigit or (C = '_');
end;

/// <summary>One-pass comment/string-aware scanner. Calls AOnMatch for
///  every whole-identifier occurrence of any Rule.Find.</summary>
procedure ScanForMatches(const AText: string;
  const ARules: TArray<TSemanticReplaceRule>;
  AOnMatch: TProc<Integer, Integer>);
type
  TScanState = (sCode, sLineCmt, sBraceCmt, sParenCmt, sStr);
var
  S: TScanState;
  I, N: Integer;
  C: Char;
begin
  S := sCode;
  I := 1;
  N := Length(AText);
  while I <= N do
  begin
    C := AText[I];
    case S of
      sCode:
      begin
        if (C = '/') and (I < N) and (AText[I + 1] = '/') then
        begin
          S := sLineCmt; Inc(I, 2); Continue;
        end;
        if C = '{' then begin S := sBraceCmt; Inc(I); Continue; end;
        if (C = '(') and (I < N) and (AText[I + 1] = '*') then
        begin
          S := sParenCmt; Inc(I, 2); Continue;
        end;
        if C = '''' then begin S := sStr; Inc(I); Continue; end;

        if C.IsLetter or (C = '_') then
        begin
          // Left boundary check.
          if (I = 1) or not IsIdentChar(AText[I - 1]) then
          begin
            var Hit: Boolean := False;
            for var R := 0 to High(ARules) do
            begin
              var F: string := ARules[R].Find;
              var L: Integer := Length(F);
              if (L = 0) or (I + L - 1 > N) then Continue;
              if CompareStr(Copy(AText, I, L), F) <> 0 then Continue;
              if (I + L <= N) and IsIdentChar(AText[I + L]) then Continue;
              AOnMatch(R, I);
              Inc(I, L);
              Hit := True;
              Break;
            end;
            if Hit then Continue;
          end;
        end;
        Inc(I);
      end;

      sLineCmt:
      begin
        if (C = #10) or (C = #13) then S := sCode;
        Inc(I);
      end;

      sBraceCmt:
      begin
        if C = '}' then S := sCode;
        Inc(I);
      end;

      sParenCmt:
      begin
        if (C = '*') and (I < N) and (AText[I + 1] = ')') then
        begin
          S := sCode; Inc(I, 2);
        end
        else
          Inc(I);
      end;

      sStr:
      begin
        if C = '''' then
        begin
          if (I < N) and (AText[I + 1] = '''') then Inc(I, 2)
          else begin S := sCode; Inc(I); end;
        end
        else
          Inc(I);
      end;
    end;
  end;
end;

type
  TMethodBody = record
    /// <summary>1-based offsets of the routine body in the source text:
    ///  BeginOffset points at the 'b' of begin; EndOffset points one
    ///  past the matching 'end;' (so [BeginOffset .. EndOffset - 1]
    ///  spans the whole body including BEGIN and END).</summary>
    BeginOffset: Integer;
    /// <summary>1-based offset RIGHT AFTER the BEGIN keyword (so the
    ///  caller can splice a local-var declaration there).</summary>
    AfterBeginOffset: Integer;
    EndOffset: Integer;
  end;

/// <summary>Comment/string-aware scan that finds every routine body in
///  the file (anything whose header starts with procedure / function /
///  constructor / destructor / class procedure / class function and
///  ends with the matching 'end;'). Used to scope the local-var
///  hoisting per routine.</summary>
function FindMethodBodies(const AText: string): TArray<TMethodBody>;
type
  TScanState = (sCode, sLineCmt, sBraceCmt, sParenCmt, sStr);
var
  S: TScanState;
  I, N, Depth: Integer;
  C: Char;
  List: TList<TMethodBody>;
  PendingHeader: Boolean;
  HeaderStart, BeginPos, AfterBegin: Integer;

  function IsKeyword(const AKw: string): Boolean;
  begin
    Result := False;
    if I + Length(AKw) - 1 > N then Exit;
    if not SameText(Copy(AText, I, Length(AKw)), AKw) then Exit;
    if (I > 1) and IsIdentChar(AText[I - 1]) then Exit;
    if (I + Length(AKw) <= N) and IsIdentChar(AText[I + Length(AKw)]) then Exit;
    Result := True;
  end;

  function AtBegin: Boolean; begin Result := IsKeyword('begin'); end;
  function AtEnd: Boolean;   begin Result := IsKeyword('end');   end;
  function AtTry: Boolean;   begin Result := IsKeyword('try');   end;
  function AtCase: Boolean;  begin Result := IsKeyword('case');  end;
  function AtRecord: Boolean; begin Result := IsKeyword('record'); end;

  function AtRoutineKeyword: Boolean;
  begin
    Result := IsKeyword('procedure') or IsKeyword('function')
      or IsKeyword('constructor') or IsKeyword('destructor');
  end;
begin
  Result := nil;
  List := TList<TMethodBody>.Create;
  try
    S := sCode;
    I := 1;
    N := Length(AText);
    PendingHeader := False;
    HeaderStart := -1;
    BeginPos := -1;
    AfterBegin := -1;
    Depth := 0;

    while I <= N do
    begin
      C := AText[I];
      case S of
        sCode:
        begin
          if (C = '/') and (I < N) and (AText[I + 1] = '/') then
          begin
            S := sLineCmt; Inc(I, 2); Continue;
          end;
          if C = '{' then begin S := sBraceCmt; Inc(I); Continue; end;
          if (C = '(') and (I < N) and (AText[I + 1] = '*') then
          begin
            S := sParenCmt; Inc(I, 2); Continue;
          end;
          if C = '''' then begin S := sStr; Inc(I); Continue; end;

          // Outside a routine body: look for header keyword.
          if (Depth = 0) and not PendingHeader and AtRoutineKeyword then
          begin
            PendingHeader := True;
            HeaderStart := I;
            // skip over the keyword
            while (I <= N) and IsIdentChar(AText[I]) do Inc(I);
            Continue;
          end;

          // Already inside a header awaiting BEGIN.
          if PendingHeader and (Depth = 0) then
          begin
            if AtBegin then
            begin
              BeginPos := I;
              AfterBegin := I + 5;
              Inc(I, 5);
              Depth := 1;
              Continue;
            end;
            Inc(I);
            Continue;
          end;

          // Inside body - track nesting.
          if Depth > 0 then
          begin
            if AtBegin or AtTry or AtCase or AtRecord then
            begin
              var KwLen := 5;       // 'begin'
              if IsKeyword('try') then KwLen := 3
              else if IsKeyword('case') then KwLen := 4
              else if IsKeyword('record') then KwLen := 6;
              Inc(Depth);
              Inc(I, KwLen);
              Continue;
            end;
            if AtEnd then
            begin
              Dec(Depth);
              Inc(I, 3);
              if Depth = 0 then
              begin
                // skip the trailing ';' if present
                var TailEnd := I;
                while (TailEnd <= N) and (AText[TailEnd] = ' ') do Inc(TailEnd);
                if (TailEnd <= N) and (AText[TailEnd] = ';') then Inc(TailEnd);
                var Body: TMethodBody;
                Body.BeginOffset := BeginPos;
                Body.AfterBeginOffset := AfterBegin;
                Body.EndOffset := TailEnd;
                List.Add(Body);
                PendingHeader := False;
                BeginPos := -1;
                AfterBegin := -1;
                HeaderStart := -1;
              end;
              Continue;
            end;
            Inc(I);
            Continue;
          end;

          Inc(I);
        end;

        sLineCmt:
        begin
          if (C = #10) or (C = #13) then S := sCode;
          Inc(I);
        end;

        sBraceCmt:
        begin
          if C = '}' then S := sCode;
          Inc(I);
        end;

        sParenCmt:
        begin
          if (C = '*') and (I < N) and (AText[I + 1] = ')') then
          begin
            S := sCode; Inc(I, 2);
          end
          else
            Inc(I);
        end;

        sStr:
        begin
          if C = '''' then
          begin
            if (I < N) and (AText[I + 1] = '''') then Inc(I, 2)
            else begin S := sCode; Inc(I); end;
          end
          else
            Inc(I);
        end;
      end;
    end;
    Result := List.ToArray;
    if HeaderStart = 0 then ;
  finally
    List.Free;
  end;
end;

function ContainingBody(const ABodies: TArray<TMethodBody>;
  AOffset: Integer): Integer;
// Returns the index of the body that contains AOffset, or -1.
var I: Integer;
begin
  Result := -1;
  for I := 0 to High(ABodies) do
    if (AOffset >= ABodies[I].BeginOffset) and (AOffset < ABodies[I].EndOffset) then
      Exit(I);
end;

class function TSemanticReplaceEngine.ApplyToText(const AOriginal: string;
  const ARules: TArray<TSemanticReplaceRule>;
  out AStats: TSemanticReplaceStats): string;
type
  TPendingEdit = record
    Offset: Integer;       // 1-based, position in AOriginal
    OldLen: Integer;       // 0 = pure insertion
    NewText: string;
  end;
var
  Matches: TList<TSemanticReplaceMatch>;
  Bodies: TArray<TMethodBody>;
  Edits: TList<TPendingEdit>;
  RuleHit: TDictionary<Integer, Boolean>;
  BodyRuleCount: TDictionary<Int64, Integer>;
  HoistedKeys: TDictionary<Int64, Boolean>;
  M: TSemanticReplaceMatch;
  E: TPendingEdit;
  I: Integer;
  Body: Integer;
  Rule: TSemanticReplaceRule;
  SB: TStringBuilder;
  Cursor: Integer;
  Ed: TPendingEdit;

  procedure IncBodyRule(ABody, ARule: Integer);
  var
    K: Int64;
    V: Integer;
  begin
    K := Int64(ABody) * 100000 + ARule;
    if BodyRuleCount.TryGetValue(K, V) then BodyRuleCount[K] := V + 1
    else BodyRuleCount.Add(K, 1);
  end;
  function GetBodyRuleCount(ABody, ARule: Integer): Integer;
  var V: Integer;
  begin
    if BodyRuleCount.TryGetValue(Int64(ABody) * 100000 + ARule, V) then
      Result := V
    else
      Result := 0;
  end;
  function AlreadyHoisted(ABody, ARule: Integer): Boolean;
  begin
    Result := HoistedKeys.ContainsKey(Int64(ABody) * 100000 + ARule);
  end;
  procedure MarkHoisted(ABody, ARule: Integer);
  begin
    HoistedKeys.Add(Int64(ABody) * 100000 + ARule, True);
  end;
  function IndentAfterBegin(AAfterBeginOff: Integer): string;
  var
    J, K: Integer;
  begin
    Result := '  ';
    J := AAfterBeginOff;
    while (J <= Length(AOriginal)) and
          ((AOriginal[J] = #13) or (AOriginal[J] = #10)) do Inc(J);
    K := J;
    while (K <= Length(AOriginal)) and (AOriginal[K] = ' ') do Inc(K);
    if K > J then Result := Copy(AOriginal, J, K - J);
  end;
  function CompareEdits(const A, B: TPendingEdit): Integer;
  begin
    Result := A.Offset - B.Offset;
  end;
begin
  AStats := Default(TSemanticReplaceStats);
  Matches := TList<TSemanticReplaceMatch>.Create;
  Edits := TList<TPendingEdit>.Create;
  RuleHit := TDictionary<Integer, Boolean>.Create;
  BodyRuleCount := TDictionary<Int64, Integer>.Create;
  HoistedKeys := TDictionary<Int64, Boolean>.Create;
  try
    Bodies := FindMethodBodies(AOriginal);

    // Anonymous methods can't capture the nested IncBodyRule, so we
    // inline the dictionary increment here.
    ScanForMatches(AOriginal, ARules,
      procedure(ARuleIdx, AOffset: Integer)
      var
        Mt: TSemanticReplaceMatch;
        B: Integer;
        K: Int64;
        V: Integer;
      begin
        Mt.RuleIdx := ARuleIdx;
        Mt.Offset := AOffset;
        Mt.Length := Length(ARules[ARuleIdx].Find);
        Matches.Add(Mt);
        RuleHit.AddOrSetValue(ARuleIdx, True);
        B := ContainingBody(Bodies, AOffset);
        if B >= 0 then
        begin
          K := Int64(B) * 100000 + ARuleIdx;
          if BodyRuleCount.TryGetValue(K, V) then BodyRuleCount[K] := V + 1
          else BodyRuleCount.Add(K, 1);
        end;
      end);

    // Build edits.
    for I := 0 to Matches.Count - 1 do
    begin
      M := Matches[I];
      Rule := ARules[M.RuleIdx];
      Body := ContainingBody(Bodies, M.Offset);
      var UseLocalVar: Boolean :=
        (Rule.LocalVarName <> '') and (Rule.LocalVarType <> '') and
        (Rule.LocalVarValue <> '') and (Rule.ReplaceWhenLocalVar <> '') and
        (Body >= 0) and (GetBodyRuleCount(Body, M.RuleIdx) >= 2);
      E := Default(TPendingEdit);
      E.Offset := M.Offset;
      E.OldLen := M.Length;
      if UseLocalVar then
        E.NewText := Rule.ReplaceWhenLocalVar
      else
        E.NewText := Rule.Replace;
      Edits.Add(E);
      if UseLocalVar and not AlreadyHoisted(Body, M.RuleIdx) then
      begin
        MarkHoisted(Body, M.RuleIdx);
        var Indent: string := IndentAfterBegin(Bodies[Body].AfterBeginOffset);
        var Hoist: TPendingEdit;
        Hoist.Offset := Bodies[Body].AfterBeginOffset;
        Hoist.OldLen := 0;
        Hoist.NewText := sLineBreak + Indent +
          'var ' + Rule.LocalVarName + ': ' + Rule.LocalVarType +
          ' := ' + Rule.LocalVarValue + ';';
        Edits.Add(Hoist);
        Inc(AStats.LocalVarsIntroduced);
      end;
    end;

    Edits.Sort(TComparer<TPendingEdit>.Construct(
      function(const A, B: TPendingEdit): Integer
      begin
        Result := A.Offset - B.Offset;
      end));

    SB := TStringBuilder.Create;
    try
      Cursor := 1;
      for I := 0 to Edits.Count - 1 do
      begin
        Ed := Edits[I];
        if Ed.Offset > Cursor then
          SB.Append(Copy(AOriginal, Cursor, Ed.Offset - Cursor));
        SB.Append(Ed.NewText);
        Cursor := Ed.Offset + Ed.OldLen;
      end;
      if Cursor <= Length(AOriginal) then
        SB.Append(Copy(AOriginal, Cursor, MaxInt));
      Result := SB.ToString;
    finally
      SB.Free;
    end;

    AStats.Occurrences := Matches.Count;
    for var K in RuleHit.Keys do
      AStats.RuleHits := AStats.RuleHits + [K];
  finally
    HoistedKeys.Free;
    BodyRuleCount.Free;
    RuleHit.Free;
    Edits.Free;
    Matches.Free;
  end;
end;

class function TSemanticReplaceEngine.FindAllMatches(const AText: string;
  const ARules: TArray<TSemanticReplaceRule>): TArray<TSemanticReplaceMatch>;
var
  L: TList<TSemanticReplaceMatch>;
begin
  L := TList<TSemanticReplaceMatch>.Create;
  try
    ScanForMatches(AText, ARules,
      procedure(ARuleIdx, AOffset: Integer)
      var Mt: TSemanticReplaceMatch;
      begin
        Mt.RuleIdx := ARuleIdx;
        Mt.Offset := AOffset;
        Mt.Length := Length(ARules[ARuleIdx].Find);
        L.Add(Mt);
      end);
    Result := L.ToArray;
  finally
    L.Free;
  end;
end;

end.
