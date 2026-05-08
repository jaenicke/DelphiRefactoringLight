(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ExtractMethod;

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.Types, System.Generics.Collections, System.RegularExpressions,
  Winapi.Windows, Vcl.Forms, Vcl.Dialogs, ToolsAPI, Expert.EditorHelper, Expert.LspManager, Expert.ExtractMethodDialog,
  Expert.SelectionValidator, Lsp.Uri, Lsp.Protocol, Lsp.Client, Delphi.FileEncoding;

type
  TParamMode = (pmConst, pmVar, pmLocal);

  TExtractedParam = record
    Name: string;
    TypeName: string;
    Mode: TParamMode;
  end;

  TLocalVar = record
    Name: string;
    TypeName: string;
  end;

  TExtractMethodInfo = record
    SelectedText: string;
    StartLine, EndLine, StartCol, EndCol: Integer;
    FileName: string;
    Params: TArray<TExtractedParam>;
    LocalVars: TArray<TLocalVar>;
    MethodName: string;
    InsertLine, ClassDeclLine: Integer;
    Indent, EnclosingClass, DiagLog: string;
    /// <summary>If non-empty, the extracted routine becomes a function
    ///  whose Result replaces this variable in the body, and the call
    ///  site is prefixed with 'ReturnVarName := '. Set by AnalyzeVariables
    ///  by promoting the first var-parameter candidate (i.e. the first
    ///  variable that is both assigned inside the block and used outside
    ///  it).</summary>
    ReturnVarName: string;
    ReturnVarType: string;
  end;

  TLspExtractMethodWizard = class
  private
    FDialog: TExtractMethodDialog;
    FCurrentInfo: TExtractMethodInfo;
    FInfoReady: Boolean;
    procedure UpdatePreview;
    procedure OnMethodNameChange(Sender: TObject);
    function GetSelectedBlock(out AInfo: TExtractMethodInfo): Boolean;
    function AnalyzeVariables(var AInfo: TExtractMethodInfo; AClient: TLspClient): Boolean;
    function FindInsertPoint(var AInfo: TExtractMethodInfo): Boolean;
    function BuildParamSignature(const AInfo: TExtractMethodInfo): string;
    function GenerateMethod(const AInfo: TExtractMethodInfo): string;
    function GenerateCall(const AInfo: TExtractMethodInfo): string;
    function GenerateClassDeclaration(const AInfo: TExtractMethodInfo): string;
    /// <summary>Removes the moved variables from the var declaration
    ///  of the original method (between InsertLine and StartLine).</summary>
    procedure RemoveLocalVarsFromDeclaration(const AInfo: TExtractMethodInfo);
    procedure DoAnalyzeAndPreview(var AInfo: TExtractMethodInfo);
  public
    procedure Execute;
  end;

var
  ExtractMethodInstance: TLspExtractMethodWizard;

implementation

uses
  System.UITypes;

type
  TSymbolKind = (skUnknown, skVariable, skConst, skField, skProperty, skParameter, skProcedure, skFunction, skType, skUnit,
    skKeyword);

  TTokenInfo = record
    Ident: string;
    LspLine, LspCol, HoverId, DefId: Integer;
    HoverText, DefFile: string;
    DefLine: Integer;
  end;

  /// <summary>File-private helper functions grouped into a class to
  ///  keep the unit free of global routines. The methods are thin and
  ///  mostly do text parsing or LSP-payload shaping for the extract
  ///  method wizard.</summary>
  TExtractMethodHelper = class
  public
    class function ParamPrefix(const AName: string): string; static;
    class function ReplaceIdentifierInLine(const ALine, AIdent, ANew: string): string; static;
    class function ClassifyHover(const AHoverText: string): TSymbolKind; static;
    class function ExtractTypeFromHover(const AHoverText: string): string; static;
    class function IsPascalKeyword(const AUpper: string): Boolean; static;
    class function IsInsideStringOrComment(const ALine: string; APos: Integer): Boolean; static;
    class function IdentUsedOutsideBlock(const AFileLines: TArray<string>;
      const AIdent: string; AMethodStart, AMethodEnd, ABlockStart, ABlockEnd, ADefLine: Integer): Boolean; static;
    class function FindMethodEnd(const AFileLines: TArray<string>; AStartLine: Integer): Integer; static;
    class function MakeHoverParams(const AFileName: string; ALine, ACol: Integer): TJSONObject; static;
  end;

/// <summary>Prefixes the parameter name with 'A' (Delphi convention for
///  method parameters). If the name already starts with 'A' followed by an
///  uppercase letter, it is kept unchanged.</summary>
class function TExtractMethodHelper.ParamPrefix(const AName: string): string;
begin
  if (Length(AName) >= 2) and (AName[1] = 'A') and
     CharInSet(AName[2], ['A'..'Z']) then
    Result := AName
  else
    Result := 'A' + AName;
end;

/// <summary>Replaces every whole-word occurrence of AIdent with ANew in
///  ALine, skipping quoted strings, '//' line comments, and '{...}' block
///  comments on the same line. Block comments spanning multiple lines are
///  handled conservatively (may rename inside such comments - edge case).</summary>
class function TExtractMethodHelper.ReplaceIdentifierInLine(const ALine, AIdent, ANew: string): string;
var
  SB: TStringBuilder;
  I, L, EndBrace: Integer;
  InString: Boolean;
  Start: Integer;
  Word: string;
begin
  if (ALine = '') or (AIdent = '') then Exit(ALine);

  SB := TStringBuilder.Create;
  try
    I := 1;
    L := Length(ALine);
    InString := False;

    while I <= L do
    begin
      // Quoted string toggle
      if ALine[I] = '''' then
      begin
        InString := not InString;
        SB.Append(ALine[I]);
        Inc(I);
        Continue;
      end;

      // Inside string: pass through
      if InString then
      begin
        SB.Append(ALine[I]);
        Inc(I);
        Continue;
      end;

      // Line comment '//': copy rest of line verbatim
      if (I < L) and (ALine[I] = '/') and (ALine[I + 1] = '/') then
      begin
        SB.Append(Copy(ALine, I, MaxInt));
        Break;
      end;

      // Block comment '{...}': copy through closing '}'
      if ALine[I] = '{' then
      begin
        EndBrace := Pos('}', ALine, I);
        if EndBrace = 0 then
        begin
          SB.Append(Copy(ALine, I, MaxInt));
          Break;
        end;
        SB.Append(Copy(ALine, I, EndBrace - I + 1));
        I := EndBrace + 1;
        Continue;
      end;

      // Identifier?
      if CharInSet(ALine[I], ['A'..'Z', 'a'..'z', '_']) then
      begin
        Start := I;
        while (I <= L) and
              CharInSet(ALine[I], ['A'..'Z', 'a'..'z', '0'..'9', '_']) do
          Inc(I);
        Word := Copy(ALine, Start, I - Start);
        // Do not rename member accesses ('Log.Log' -> only the first
        // 'Log' is the variable; the second is the method name after the
        // dot and must stay as-is). '&' is the Delphi escape prefix for
        // reserved-word identifiers.
        var IsMemberAccess := (Start > 1) and
          ((ALine[Start - 1] = '.') or (ALine[Start - 1] = '&'));
        if SameText(Word, AIdent) and not IsMemberAccess then
          SB.Append(ANew)
        else
          SB.Append(Word);
      end
      else
      begin
        SB.Append(ALine[I]);
        Inc(I);
      end;
    end;

    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

class function TExtractMethodHelper.ClassifyHover(const AHoverText: string): TSymbolKind;
var
  Upper: string;
begin
  Result := skUnknown;
  if AHoverText = '' then Exit;
  Upper := UpperCase(AHoverText);
  if Pos('PROCEDURE ', Upper) > 0 then Exit(skProcedure);
  if Pos('FUNCTION ', Upper) > 0 then Exit(skFunction);
  if Pos('CONSTRUCTOR ', Upper) > 0 then Exit(skProcedure);
  if Pos('DESTRUCTOR ', Upper) > 0 then Exit(skProcedure);
  if (Pos(' = CLASS', Upper) > 0) or (Pos(' = INTERFACE', Upper) > 0) or
     (Pos(' = RECORD', Upper) > 0) or (Pos('TYPE ', Upper) > 0) then Exit(skType);
  if Pos('UNIT ', Upper) > 0 then Exit(skUnit);
  if Pos('PROPERTY ', Upper) > 0 then Exit(skProperty);
  if Pos('CONST ', Upper) > 0 then Exit(skConst);
  if Pos('VAR ', Upper) > 0 then Exit(skVariable);
  if Pos(': ', AHoverText) > 0 then Exit(skVariable);
end;

class function TExtractMethodHelper.ExtractTypeFromHover(const AHoverText: string): string;
begin
  Result := '';
  if AHoverText = '' then Exit;
  var M := TRegEx.Match(AHoverText, ':\s*([A-Za-z_][\w.]*(?:<[^>]+>)?)');
  if M.Success and (M.Groups.Count > 1) then
    Result := M.Groups[1].Value;
end;

class function TExtractMethodHelper.IsPascalKeyword(const AUpper: string): Boolean;
const
  Keywords: array[0..64] of string = (
    'AND','ARRAY','AS','ASM','BEGIN','CASE','CLASS','CONST','CONSTRUCTOR',
    'DESTRUCTOR','DISPINTERFACE','DIV','DO','DOWNTO','ELSE','END','EXCEPT',
    'EXPORTS','FILE','FINALIZATION','FINALLY','FOR','FUNCTION','GOTO','IF',
    'IMPLEMENTATION','IN','INHERITED','INITIALIZATION','INLINE','INTERFACE',
    'IS','LABEL','LIBRARY','MOD','NIL','NOT','OBJECT','OF','OR','OUT',
    'PACKED','PROCEDURE','PROGRAM','PROPERTY','RAISE','RECORD','REPEAT',
    'RESOURCESTRING','SET','SHL','SHR','STRING','THEN','THREADVAR','TO',
    'TRY','TYPE','UNIT','UNTIL','USES','VAR','WHILE','WITH','XOR');
begin
  for var KW in Keywords do
    if AUpper = KW then Exit(True);
  Result := False;
end;

// Forward declaration is no longer needed: both IsInsideStringOrComment and
// its callers now live in TExtractMethodHelper and resolve via the class.

/// <summary>Checks whether the identifier AIdent occurs in the source
///  code between AMethodStart and AMethodEnd (1-based, inclusive) outside
///  of the selected block (ABlockStart..ABlockEnd). ADefLine (1-based,
///  -1 if unknown) is also excluded because the declaration itself does
///  not count as a "use".</summary>
class function TExtractMethodHelper.IdentUsedOutsideBlock(const AFileLines: TArray<string>; const AIdent: string;
  AMethodStart, AMethodEnd, ABlockStart, ABlockEnd, ADefLine: Integer): Boolean;
var
  UpperName: string;
  I: Integer;
begin
  Result := False;
  UpperName := UpperCase(AIdent);
  for I := AMethodStart - 1 to AMethodEnd - 1 do
  begin
    if I < 0 then Continue;
    if I >= Length(AFileLines) then Break;
    // Skip lines INSIDE the block (compare 1-based with 0-based index)
    if (I + 1 >= ABlockStart) and (I + 1 <= ABlockEnd) then Continue;
    // Skip the declaration line itself (does not count as a use)
    if (ADefLine > 0) and (I + 1 = ADefLine) then Continue;

    var Line := AFileLines[I];
    // Whole-word search
    var Pos_ := System.Pos(UpperName, UpperCase(Line));
    while Pos_ > 0 do
    begin
      var BeforeOk := (Pos_ = 1) or
        not CharInSet(Line[Pos_ - 1], ['A'..'Z','a'..'z','0'..'9','_']);
      var AfterPos := Pos_ + Length(AIdent);
      var AfterOk := (AfterPos > Length(Line)) or
        not CharInSet(Line[AfterPos], ['A'..'Z','a'..'z','0'..'9','_']);
      if BeforeOk and AfterOk and not TExtractMethodHelper.IsInsideStringOrComment(Line, Pos_) then
        Exit(True);
      // Continue searching
      var NextPos := System.Pos(UpperName, UpperCase(Copy(Line, Pos_ + Length(AIdent))));
      if NextPos = 0 then Break;
      Pos_ := Pos_ + Length(AIdent) + NextPos - 1;
    end;
  end;
end;

/// <summary>Finds the end of the surrounding method (the line with the
///  matching "end;" of the procedure).</summary>
class function TExtractMethodHelper.FindMethodEnd(const AFileLines: TArray<string>; AStartLine: Integer): Integer;

  function HasKeyword(const AUpper, AKeyword: string): Boolean;
  // Whitespace-bounded keyword detection on a line. AUpper is already
  // uppercased and trimmed.
  begin
    Result :=
      (AUpper = AKeyword) or
      AUpper.StartsWith(AKeyword + ' ') or
      AUpper.EndsWith(' ' + AKeyword) or
      (Pos(' ' + AKeyword + ' ', ' ' + AUpper + ' ') > 0);
  end;

var
  BlockDepth: Integer;
  Upper: string;
  InBegin: Boolean;
begin
  // AStartLine is 1-based. Walk forward and balance opener/closer
  // keywords. Openers are BEGIN, TRY, CASE, ASM (they all close with
  // END). Without counting TRY/CASE we'd exit too early as soon as a
  // try-block inside the method closes more ENDs than there are BEGINs.
  BlockDepth := 0;
  InBegin := False;
  for var I := AStartLine - 1 to Length(AFileLines) - 1 do
  begin
    Upper := UpperCase(Trim(AFileLines[I]));
    // Next method or unit end reached? Bail out cautiously.
    if (not InBegin) and
       (Upper.StartsWith('PROCEDURE ') or Upper.StartsWith('FUNCTION ') or
        Upper.StartsWith('CONSTRUCTOR ') or Upper.StartsWith('DESTRUCTOR ')) and
       (I > AStartLine - 1) then
      Exit(I); // 0-based, returned as 1-based: I (line BEFORE next method)

    if HasKeyword(Upper, 'BEGIN') then
    begin
      Inc(BlockDepth);
      InBegin := True;
    end;
    if InBegin then
    begin
      // try / case / asm also close with end; count them so the
      // depth stays balanced.
      if HasKeyword(Upper, 'TRY') then Inc(BlockDepth);
      if HasKeyword(Upper, 'ASM') then Inc(BlockDepth);
      // 'case ... of' is the form that opens a block. A bare 'case'
      // also appears in variant records, but those live at type level
      // and never inside a method body, so counting any 'CASE' as
      // opener is safe here.
      if HasKeyword(Upper, 'CASE') then Inc(BlockDepth);
    end;

    if InBegin and ((Upper = 'END;') or (Upper = 'END.') or Upper.StartsWith('END ') or
                    Upper.StartsWith('END;') or Upper.StartsWith('END.')) then
    begin
      Dec(BlockDepth);
      if BlockDepth <= 0 then
        Exit(I + 1); // 1-based
    end;
  end;
  Result := Length(AFileLines);
end;

class function TExtractMethodHelper.IsInsideStringOrComment(const ALine: string; APos: Integer): Boolean;
var
  I: Integer;
  InString: Boolean;
begin
  // Line comment //
  for I := 1 to APos - 1 do
    if (I < Length(ALine)) and (ALine[I] = '/') and (ALine[I + 1] = '/') then
      Exit(True);
  // { } comment
  var BD := 0;
  for I := 1 to APos - 1 do
    if ALine[I] = '{' then Inc(BD)
    else if ALine[I] = '}' then Dec(BD);
  if BD > 0 then Exit(True);
  // String literal (count apostrophes)
  InString := False;
  for I := 1 to APos - 1 do
    if ALine[I] = '''' then InString := not InString;
  Result := InString;
end;

class function TExtractMethodHelper.MakeHoverParams(const AFileName: string; ALine, ACol: Integer): TJSONObject;
begin
  var TD := TJSONObject.Create;
  TD.AddPair('uri', TLspUri.PathToFileUri(ExpandFileName(AFileName)));
  var Pos := TJSONObject.Create;
  Pos.AddPair('line', TJSONNumber.Create(ALine));
  Pos.AddPair('character', TJSONNumber.Create(ACol));
  Result := TJSONObject.Create;
  Result.AddPair('textDocument', TD);
  Result.AddPair('position', Pos);
end;

function TLspExtractMethodWizard.GetSelectedBlock(out AInfo: TExtractMethodInfo): Boolean;
var
  ES: IOTAEditorServices;
  EB: IOTAEditBuffer;
  BL: IOTAEditBlock;
begin
  Result := False;
  AInfo := Default(TExtractMethodInfo);
  if not Supports(BorlandIDEServices, IOTAEditorServices, ES) then Exit;
  EB := ES.TopBuffer;
  if EB = nil then Exit;
  BL := EB.EditBlock;
  if (BL = nil) or not BL.IsValid then
  begin
    MessageDlg('Please select the code block to extract first.', mtInformation, [mbOK], 0);
    Exit;
  end;
  AInfo.SelectedText := BL.Text;
  AInfo.StartLine := BL.StartingRow;
  AInfo.StartCol := BL.StartingColumn;
  AInfo.EndLine := BL.EndingRow;
  AInfo.EndCol := BL.EndingColumn;
  // If the selection ends at the start of the next line (EndCol=1),
  // that line does NOT belong to the selection
  if AInfo.EndCol = 1 then Dec(AInfo.EndLine);
  if Trim(AInfo.SelectedText) = '' then begin MessageDlg('Selection is empty.', mtWarning, [mbOK], 0); Exit; end;
  var Mod_ := EB.Module;
  if Mod_ <> nil then AInfo.FileName := Mod_.FileName;
  var Lines := AInfo.SelectedText.Split([#10]);
  var MinI := MaxInt;
  for var L in Lines do begin var S := L.TrimRight([#13,#10]); if Trim(S)='' then Continue;
    var C := 0; while (C<Length(S)) and CharInSet(S[C+1],[' ',#9]) do Inc(C);
    if C<MinI then MinI:=C; end;
  if MinI=MaxInt then MinI:=0;
  AInfo.Indent := StringOfChar(' ', MinI);
  AInfo.MethodName := 'ExtractedMethod';
  Result := True;
end;

function TLspExtractMethodWizard.AnalyzeVariables(var AInfo: TExtractMethodInfo; AClient: TLspClient): Boolean;
var
  Lines: TArray<string>;
  Seen: TDictionary<string, Boolean>;
  Assigned_: TDictionary<string, Boolean>;
  Tokens: TList<TTokenInfo>;
  NeedDef: TList<Integer>;
  PL: TList<TExtractedParam>;
  LVL: TList<TLocalVar>;
  Line: string;
  I: Integer;
begin
  Result := True;
  Lines := AInfo.SelectedText.Split([#10]);
  Seen := TDictionary<string, Boolean>.Create;
  Assigned_ := TDictionary<string, Boolean>.Create;
  Tokens := TList<TTokenInfo>.Create;
  NeedDef := TList<Integer>.Create;
  PL := TList<TExtractedParam>.Create;
  LVL := TList<TLocalVar>.Create;
  try
    AInfo.DiagLog := AInfo.DiagLog + '=== Variable analysis (parallel) ===' + sLineBreak;
    // Collect identifiers
    if FDialog<>nil then FDialog.SetStatus('Collecting identifiers...');
    for I := 0 to High(Lines) do
    begin
      Line := Lines[I].TrimRight([#13,#10]);
      var LL := AInfo.StartLine-1+I;
      var AM := TRegEx.Match(Line, '^\s*(\w+)\s*:=');
      if AM.Success then Assigned_.AddOrSetValue(UpperCase(AM.Groups[1].Value), True);
      for var M in TRegEx.Matches(Line, '\b([A-Za-z_]\w*)\b') do
      begin
        var UI := UpperCase(M.Value);
        if Seen.ContainsKey(UI) then Continue;
        // Skip Pascal keywords (cost LSP time, yield nothing)
        if TExtractMethodHelper.IsPascalKeyword(UI) then Continue;
        // Skip tokens inside string/comment
        if TExtractMethodHelper.IsInsideStringOrComment(Line, M.Index) then Continue;
        Seen.Add(UI, True);
        var TI: TTokenInfo;
        TI.Ident:=M.Value; TI.LspLine:=LL; TI.LspCol:=M.Index-1;
        TI.HoverId:=-1; TI.DefId:=-1; TI.HoverText:=''; TI.DefFile:=''; TI.DefLine:=-1;
        Tokens.Add(TI);
      end;
    end;
    AInfo.DiagLog := AInfo.DiagLog + IntToStr(Tokens.Count)+' relevant identifiers (keywords/strings filtered)'+sLineBreak;
    // DelphiLSP does NOT support parallel requests -> batch size 1
    const BatchSize = 1;
    var HoverOk := 0; var HoverEmpty := 0; var HoverError := 0;
    if FDialog<>nil then FDialog.SetProgress(0, Tokens.Count*2);

    var BatchStart := 0;
    while BatchStart < Tokens.Count do
    begin
      var BatchEnd := BatchStart + BatchSize - 1;
      if BatchEnd >= Tokens.Count then BatchEnd := Tokens.Count - 1;

      if FDialog<>nil then
        FDialog.SetStatus(Format('Hover batch %d-%d of %d...', [BatchStart+1, BatchEnd+1, Tokens.Count]));

      // Send batch
      for I := BatchStart to BatchEnd do
      begin
        var TI := Tokens[I];
        TI.HoverId := AClient.SendRequestAsync('textDocument/hover',
          TExtractMethodHelper.MakeHoverParams(AInfo.FileName, TI.LspLine, TI.LspCol));
        Tokens[I] := TI;
      end;

      // Receive batch
      for I := BatchStart to BatchEnd do
      begin
        var TI := Tokens[I];
        if FDialog<>nil then FDialog.SetProgress(I+1, Tokens.Count*2);
        try
          var R := AClient.WaitForResponse(TI.HoverId, 30000);
          try
            var RO: TJSONObject;
            if R.TryGetValue<TJSONObject>('result', RO) then
            begin
              var C := RO.GetValue('contents');
              if C is TJSONObject then TI.HoverText := TJSONObject(C).GetValue<string>('value','')
              else if C is TJSONString then TI.HoverText := C.Value;
            end;
            if TI.HoverText<>'' then Inc(HoverOk) else Inc(HoverEmpty);
          finally R.Free; end;
        except
          on E: Exception do
          begin
            Inc(HoverError);
            AInfo.DiagLog := AInfo.DiagLog + Format('  [%d] %s: hover error: %s', [I, TI.Ident, E.Message]) + sLineBreak;
          end;
        end;
        Tokens[I] := TI;
      end;

      BatchStart := BatchEnd + 1;
    end;

    AInfo.DiagLog := AInfo.DiagLog + Format('Hover: %d OK, %d empty, %d errors',
      [HoverOk, HoverEmpty, HoverError]) + sLineBreak;
    // Only variables need a definition
    for I := 0 to Tokens.Count-1 do
    begin
      var TI := Tokens[I];
      if TI.HoverText='' then Continue;
      var K := TExtractMethodHelper.ClassifyHover(TI.HoverText);
      if not (K in [skVariable,skConst,skField,skProperty,skParameter]) then Continue;
      if TExtractMethodHelper.ExtractTypeFromHover(TI.HoverText)='' then Continue;
      NeedDef.Add(I);
    end;
    // Definitions in batches
    if FDialog <> nil then
      FDialog.SetStatus(Format('Resolving %d definitions...', [NeedDef.Count]));
    BatchStart := 0;
    while BatchStart < NeedDef.Count do
    begin
      var BatchEnd := BatchStart + BatchSize - 1;
      if BatchEnd >= NeedDef.Count then BatchEnd := NeedDef.Count - 1;

      if FDialog<>nil then
        FDialog.SetStatus(Format('Definition batch %d-%d of %d...', [BatchStart+1, BatchEnd+1, NeedDef.Count]));

      // Send batch
      for I := BatchStart to BatchEnd do
      begin
        var Idx := NeedDef[I];
        var TI := Tokens[Idx];
        TI.DefId := AClient.SendRequestAsync('textDocument/definition',
          TExtractMethodHelper.MakeHoverParams(AInfo.FileName, TI.LspLine, TI.LspCol));
        Tokens[Idx] := TI;
      end;

      // Receive batch
      for I := BatchStart to BatchEnd do
      begin
        var Idx := NeedDef[I];
        var TI := Tokens[Idx];
        if FDialog <> nil then
          FDialog.SetProgress(Tokens.Count+I+1, Tokens.Count*2);
        try
          var R := AClient.WaitForResponse(TI.DefId, 30000);
          try
            var RV := R.GetValue('result');
            var LO: TJSONObject := nil;
            if RV is TJSONObject then LO := TJSONObject(RV)
            else if (RV is TJSONArray) and (TJSONArray(RV).Count>0) then
              LO := TJSONArray(RV).Items[0] as TJSONObject;
            if LO<>nil then
            begin
              TI.DefFile := TLspUri.FileUriToPath(LO.GetValue<string>('uri',''));
              var Rng := LO.GetValue<TJSONObject>('range');
              if Rng<>nil then
                TI.DefLine := Rng.GetValue<TJSONObject>('start').GetValue<Integer>('line',-1)+1;
            end;
          finally R.Free; end;
        except end;
        Tokens[Idx] := TI;
      end;

      BatchStart := BatchEnd + 1;
    end;
    // Classify
    if FDialog<>nil then FDialog.SetStatus('Classifying...');
    var FileLines := TDelphiFileEncoding.ReadLines(AInfo.FileName);
    var MethodEnd := TExtractMethodHelper.FindMethodEnd(FileLines, AInfo.InsertLine + 1); // start after the method header
    AInfo.DiagLog := AInfo.DiagLog + Format('Method: line %d-%d, block: line %d-%d',
      [AInfo.InsertLine, MethodEnd, AInfo.StartLine, AInfo.EndLine]) + sLineBreak;

    for I := 0 to Tokens.Count-1 do
    begin
      var TI := Tokens[I];
      if TI.HoverText='' then
      begin
        AInfo.DiagLog := AInfo.DiagLog + Format('  %s: no hover',[TI.Ident])+sLineBreak;
        Continue;
      end;
      var K := TExtractMethodHelper.ClassifyHover(TI.HoverText);
      var TN := TExtractMethodHelper.ExtractTypeFromHover(TI.HoverText);
      AInfo.DiagLog := AInfo.DiagLog + Format('  %s: Kind=%d Type=%s Hover=%s',
        [TI.Ident,Ord(K),TN,Copy(TI.HoverText,1,60)])+sLineBreak;
      if not (K in [skVariable,skConst,skField,skProperty,skParameter]) then Continue;
      if TN='' then Continue;
      var DSF := SameText(ExpandFileName(TI.DefFile), ExpandFileName(AInfo.FileName));
      if DSF and (TI.DefLine>=AInfo.StartLine) and (TI.DefLine<=AInfo.EndLine) then
        begin AInfo.DiagLog:=AInfo.DiagLog+'    -> local'+sLineBreak; Continue; end;
      if DSF and (TI.DefLine>=AInfo.InsertLine) and (TI.DefLine<AInfo.StartLine) then
      begin
        // Local var of the method: is it used OUTSIDE the block?
        var UsedOutside := TExtractMethodHelper.IdentUsedOutsideBlock(FileLines, TI.Ident,
          AInfo.InsertLine, MethodEnd, AInfo.StartLine, AInfo.EndLine,
          TI.DefLine);
        if UsedOutside then
        begin
          // Variable is also used outside -> parameter
          var P: TExtractedParam; P.Name:=TI.Ident; P.TypeName:=TN;
          if Assigned_.ContainsKey(UpperCase(TI.Ident)) then P.Mode:=pmVar
          else P.Mode:=pmConst;
          AInfo.DiagLog:=AInfo.DiagLog+'    -> parameter (also used outside)'+sLineBreak;
          PL.Add(P);
        end
        else
        begin
          // Only inside the block -> move
          var LV: TLocalVar; LV.Name:=TI.Ident; LV.TypeName:=TN; LVL.Add(LV);
          AInfo.DiagLog:=AInfo.DiagLog+'    -> move (only inside block)'+sLineBreak;
        end;
        Continue;
      end;
      if not DSF then begin AInfo.DiagLog:=AInfo.DiagLog+'    -> external'+sLineBreak; Continue; end;
      if (AInfo.EnclosingClass<>'') and (TI.DefLine<AInfo.InsertLine) then
        begin AInfo.DiagLog:=AInfo.DiagLog+'    -> class member'+sLineBreak; Continue; end;
      if (AInfo.EnclosingClass='') and (TI.DefLine<AInfo.InsertLine) then
        begin AInfo.DiagLog:=AInfo.DiagLog+'    -> global'+sLineBreak; Continue; end;
      var P: TExtractedParam; P.Name:=TI.Ident; P.TypeName:=TN;
      if K=skConst then P.Mode:=pmConst
      else if Assigned_.ContainsKey(UpperCase(TI.Ident)) then P.Mode:=pmVar
      else P.Mode:=pmConst;
      AInfo.DiagLog:=AInfo.DiagLog+'    -> parameter'+sLineBreak;
      PL.Add(P);
    end;
    AInfo.Params := PL.ToArray;
    AInfo.LocalVars := LVL.ToArray;

    // Promote the first var-parameter (pmVar) to the function's Result:
    // when a value is assigned inside the block and read afterwards, it is
    // more idiomatic Delphi to return it as a function result than to pass
    // it back through a var parameter.
    AInfo.ReturnVarName := '';
    AInfo.ReturnVarType := '';
    for var RI := 0 to High(AInfo.Params) do
      if AInfo.Params[RI].Mode = pmVar then
      begin
        AInfo.ReturnVarName := AInfo.Params[RI].Name;
        AInfo.ReturnVarType := AInfo.Params[RI].TypeName;
        AInfo.DiagLog := AInfo.DiagLog +
          '  -> Return value: ' + AInfo.ReturnVarName +
          ' (' + AInfo.ReturnVarType + ')' + sLineBreak;

        // Remove the promoted parameter from the params array.
        var Rest := TList<TExtractedParam>.Create;
        try
          for var RJ := 0 to High(AInfo.Params) do
            if RJ <> RI then Rest.Add(AInfo.Params[RJ]);
          AInfo.Params := Rest.ToArray;
        finally
          Rest.Free;
        end;
        Break;
      end;

    if FDialog<>nil then FDialog.SetProgress(Tokens.Count*2, Tokens.Count*2);
  finally
    NeedDef.Free; Tokens.Free; LVL.Free; PL.Free; Assigned_.Free; Seen.Free;
  end;
end;

function TLspExtractMethodWizard.FindInsertPoint(var AInfo: TExtractMethodInfo): Boolean;
var
  FL: TArray<string>;
  I: Integer;
  U: string;
begin
  FL := TDelphiFileEncoding.ReadLines(AInfo.FileName);
  AInfo.EnclosingClass := ''; AInfo.InsertLine := AInfo.StartLine; AInfo.ClassDeclLine := -1;
  for I := AInfo.StartLine-2 downto 0 do
  begin
    U := UpperCase(Trim(FL[I]));
    if U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ') or
       U.StartsWith('CONSTRUCTOR ') or U.StartsWith('DESTRUCTOR ') then
    begin
      AInfo.InsertLine := I+1;
      var L := Trim(FL[I]); var DP := Pos('.', L);
      if DP>0 then begin var AK := Trim(Copy(L, Pos(' ',L)+1));
        AInfo.EnclosingClass := Copy(AK, 1, Pos('.',AK)-1); end;
      Break;
    end;
  end;
  if AInfo.EnclosingClass<>'' then
  begin
    var CP := UpperCase(AInfo.EnclosingClass+' = CLASS');
    for I := 0 to Length(FL)-1 do
    begin
      U := UpperCase(Trim(FL[I]));
      if Pos(CP, U)>0 then
      begin
        for var J := I+1 to Length(FL)-1 do
        begin
          var SU := UpperCase(Trim(FL[J]));
          if (SU='PRIVATE') or SU.StartsWith('PRIVATE ') then begin AInfo.ClassDeclLine:=J+2; Break; end;
          if (SU='PUBLIC') or SU.StartsWith('PUBLIC ') then begin AInfo.ClassDeclLine:=J+1; Break; end;
          if SU.StartsWith('END;') then begin AInfo.ClassDeclLine:=J+1; Break; end;
        end;
        Break;
      end;
    end;
  end;
  Result := True;
end;

function TLspExtractMethodWizard.BuildParamSignature(const AInfo: TExtractMethodInfo): string;
var
  SB: TStringBuilder;
  F: Boolean;
begin
  var HP := False;
  for var P in AInfo.Params do if P.Mode<>pmLocal then begin HP:=True; Break; end;
  if not HP then Exit('');
  SB := TStringBuilder.Create;
  try
    SB.Append('('); F:=True;
    for var P in AInfo.Params do begin if P.Mode=pmLocal then Continue;
      if not F then SB.Append('; ');
      case P.Mode of pmConst: SB.Append('const '); pmVar: SB.Append('var '); end;
      SB.Append(TExtractMethodHelper.ParamPrefix(P.Name)+': '+P.TypeName); F:=False; end;
    SB.Append(')'); Result := SB.ToString;
  finally SB.Free; end;
end;

function TLspExtractMethodWizard.GenerateMethod(const AInfo: TExtractMethodInfo): string;
var
  SB: TStringBuilder;
  FL: TArray<string>;
begin
  SB := TStringBuilder.Create;
  try
    var RoutineKind := 'procedure';
    if AInfo.ReturnVarName <> '' then RoutineKind := 'function';
    if AInfo.EnclosingClass<>'' then SB.Append(RoutineKind+' '+AInfo.EnclosingClass+'.'+AInfo.MethodName)
    else SB.Append(RoutineKind+' '+AInfo.MethodName);
    SB.Append(BuildParamSignature(AInfo));
    if AInfo.ReturnVarName <> '' then
      SB.Append(': '+AInfo.ReturnVarType);
    SB.AppendLine(';');
    if Length(AInfo.LocalVars) > 0 then
    begin
      SB.AppendLine('var');
      for var LVIdx := 0 to High(AInfo.LocalVars) do
      begin
        var LV := AInfo.LocalVars[LVIdx];
        SB.AppendLine('  ' + LV.Name + ': ' + LV.TypeName + ';');
      end;
    end;
    SB.AppendLine('begin');
    FL := TDelphiFileEncoding.ReadLines(AInfo.FileName);

    // Base indentation: minimum indent of all non-empty lines in the block
    var BIL := MaxInt;
    for var LI := AInfo.StartLine-1 to AInfo.EndLine-1 do
    begin
      if (LI<0) or (LI>=Length(FL)) then Continue;
      var S := FL[LI];
      if Trim(S)='' then Continue;
      var C := 0;
      while (C<Length(S)) and CharInSet(S[C+1],[' ',#9]) do Inc(C);
      if C<BIL then BIL:=C;
    end;
    if BIL=MaxInt then BIL:=0;

    // All lines: strip the base indent, then prefix with 2 spaces
    for var LI := AInfo.StartLine-1 to AInfo.EndLine-1 do
    begin
      if (LI<0) or (LI>=Length(FL)) then Continue;
      var S := FL[LI];
      if (BIL>0) and (Length(S)>=BIL) then
      begin
        var AllSpace := True;
        for var CI := 1 to BIL do
          if (CI>Length(S)) or not CharInSet(S[CI],[' ',#9]) then
          begin
            AllSpace := False; Break;
          end;
        if AllSpace then S := Copy(S, BIL+1);
      end;
      // Rename parameter references to the A-prefixed form inside the
      // extracted method body. Local variables keep their original names.
      for var PRI := 0 to High(AInfo.Params) do
        if AInfo.Params[PRI].Mode <> pmLocal then
          S := TExtractMethodHelper.ReplaceIdentifierInLine(S, AInfo.Params[PRI].Name,
            TExtractMethodHelper.ParamPrefix(AInfo.Params[PRI].Name));
      // If a return variable was promoted: replace its identifier with
      // Result inside the body. The assignment 'Result := ...' makes the
      // function return the value.
      if AInfo.ReturnVarName <> '' then
        S := TExtractMethodHelper.ReplaceIdentifierInLine(S, AInfo.ReturnVarName, 'Result');
      SB.AppendLine('  '+S);
    end;
    SB.AppendLine('end;'); SB.AppendLine('');
    Result := SB.ToString;
  finally SB.Free; end;
end;

function TLspExtractMethodWizard.GenerateCall(const AInfo: TExtractMethodInfo): string;
var
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    SB.Append(AInfo.Indent);
    // If the extracted routine returns a value, the call site assigns it
    // back to the original variable: 'b := Method(...);'.
    if AInfo.ReturnVarName <> '' then
      SB.Append(AInfo.ReturnVarName+' := ');
    SB.Append(AInfo.MethodName);
    var HP := False;
    for var P in AInfo.Params do if P.Mode<>pmLocal then begin HP:=True; Break; end;
    if HP then begin SB.Append('('); var F := True;
      for var P in AInfo.Params do begin if P.Mode=pmLocal then Continue;
        if not F then SB.Append(', '); SB.Append(P.Name); F:=False; end;
      SB.Append(')'); end;
    SB.Append(';'); Result := SB.ToString;
  finally SB.Free; end;
end;

function TLspExtractMethodWizard.GenerateClassDeclaration(const AInfo: TExtractMethodInfo): string;
begin
  if AInfo.ReturnVarName <> '' then
    Result := '    function '+AInfo.MethodName+BuildParamSignature(AInfo)+
      ': '+AInfo.ReturnVarType+';'+sLineBreak
  else
    Result := '    procedure '+AInfo.MethodName+BuildParamSignature(AInfo)+
      ';'+sLineBreak;
end;

procedure TLspExtractMethodWizard.RemoveLocalVarsFromDeclaration(
  const AInfo: TExtractMethodInfo);
var
  Content: string;
  FL: TArray<string>;
  SL: TStringList;
  ToRemove: TDictionary<string, Boolean>;
  I: Integer;
  InVar: Boolean;
begin
  // IMPORTANT: read the editor buffer (not the file on disk) so that the
  // newly inserted method and call are not lost.
  if not TEditorHelper.ReadEditorContent(AInfo.FileName, Content) then Exit;

  SL := TStringList.Create;
  try
    SL.Text := Content;
    SetLength(FL, SL.Count);
    for I := 0 to SL.Count - 1 do
      FL[I] := SL[I];
  finally
    SL.Free;
  end;

  ToRemove := TDictionary<string, Boolean>.Create;
  // Collect line changes: 1-based line -> NewText (#0 = delete)
  var LineChanges := TList<TPair<Integer, string>>.Create;
  try
    for var LV in AInfo.LocalVars do
      ToRemove.AddOrSetValue(UpperCase(LV.Name), True);

    // Walk all var sections in the file and remove affected names,
    // BUT: skip the var section of the NEWLY INSERTED method!
    var UpperMethodName := UpperCase(AInfo.MethodName);
    InVar := False;
    var IsInExtractedMethod := False;
    I := 0;
    while I < Length(FL) do
    begin
      var Line := FL[I];
      var U := UpperCase(Trim(Line));

      if U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ') or
         U.StartsWith('CONSTRUCTOR ') or U.StartsWith('DESTRUCTOR ') then
      begin
        InVar := False;
        // Check whether this is the header of the new method
        // e.g. "procedure TForm4.ExtractedMethod(...)" or "procedure ExtractedMethod(..)"
        IsInExtractedMethod := (Pos('.' + UpperMethodName, U) > 0) or
          (Pos('PROCEDURE ' + UpperMethodName, U) > 0) or
          (Pos('FUNCTION ' + UpperMethodName, U) > 0);
      end
      else if (U = 'VAR') or U.StartsWith('VAR ') then
        InVar := True
      else if (U = 'BEGIN') or U.StartsWith('BEGIN ') or
              (U = 'CONST') or U.StartsWith('CONST ') then
        InVar := False;

      // New method: skip
      if InVar and not IsInExtractedMethod then
      begin
        var ColPos := Pos(':', Line);
        if ColPos > 0 then
        begin
          var Before := Copy(Line, 1, ColPos - 1);
          var After := Copy(Line, ColPos);
          var Names := Before.Split([',']);
          var NewNames := TStringList.Create;
          try
            for var N in Names do
            begin
              var TrimN := Trim(N);
              if TrimN = '' then Continue;
              if not ToRemove.ContainsKey(UpperCase(TrimN)) then
                NewNames.Add(TrimN);
            end;
            if NewNames.Count = 0 then
              // All names removed -> delete the line
              LineChanges.Add(TPair<Integer, string>.Create(I + 1, #0))
            else if NewNames.Count < Length(Names) then
            begin
              var Indent := '';
              for var CI := 1 to Length(Line) do
                if CharInSet(Line[CI], [' ', #9]) then
                  Indent := Indent + Line[CI]
                else
                  Break;
              LineChanges.Add(TPair<Integer, string>.Create(I + 1,
                Indent + string.Join(', ', NewNames.ToStringArray) + After));
            end;
          finally
            NewNames.Free;
          end;
        end;
      end;

      Inc(I);
    end;

    // Apply changes in reverse so that line numbers stay stable
    for var J := LineChanges.Count - 1 downto 0 do
    begin
      var LineNum := LineChanges[J].Key;
      var NewText := LineChanges[J].Value;
      if NewText = #0 then
        TEditorHelper.DeleteLineAt(AInfo.FileName, LineNum)
      else
        TEditorHelper.ReplaceLineAt(AInfo.FileName, LineNum, NewText);
    end;
  finally
    LineChanges.Free;
    ToRemove.Free;
  end;
end;

procedure TLspExtractMethodWizard.UpdatePreview;
var
  NMC, CC, CDC, Preview: string;
begin
  if not FInfoReady then Exit;
  // Take MethodName from the current dialog value
  FCurrentInfo.MethodName := FDialog.GetMethodName;
  if FCurrentInfo.MethodName = '' then
  begin
    FDialog.SetPreviewText('Please enter a method name.');
    FDialog.EnableExtract(False);
    Exit;
  end;
  NMC := GenerateMethod(FCurrentInfo);
  CC := GenerateCall(FCurrentInfo);
  if FCurrentInfo.EnclosingClass <> '' then
    CDC := GenerateClassDeclaration(FCurrentInfo)
  else
    CDC := '';
  Preview := '--- New method (before line ' + IntToStr(FCurrentInfo.InsertLine) + ') ---' +
    sLineBreak + NMC + sLineBreak +
    '--- Call ---' + sLineBreak + CC + sLineBreak;
  if CDC <> '' then
    Preview := Preview + sLineBreak +
      '--- Class declaration (line ' + IntToStr(FCurrentInfo.ClassDeclLine) + ') ---' +
      sLineBreak + CDC;
  if Length(FCurrentInfo.LocalVars) > 0 then
  begin
    Preview := Preview + sLineBreak + '--- Moved variables ---' + sLineBreak;
    for var LV in FCurrentInfo.LocalVars do
      Preview := Preview + '  ' + LV.Name + ': ' + LV.TypeName + sLineBreak;
  end;
  Preview := Preview + sLineBreak + sLineBreak + FCurrentInfo.DiagLog;
  FDialog.SetPreviewText(Preview);
  FDialog.EnableExtract(True);
end;

procedure TLspExtractMethodWizard.OnMethodNameChange(Sender: TObject);
begin
  UpdatePreview;
end;

procedure TLspExtractMethodWizard.DoAnalyzeAndPreview(var AInfo: TExtractMethodInfo);
var
  Client: TLspClient;
  DJ, RP: string;
begin
  DJ := TEditorHelper.FindDelphiLspJson;
  if DJ='' then begin FDialog.SetPreviewText('No .delphilsp.json found.'); FDialog.SetStatus('Error'); Exit; end;
  RP := TEditorHelper.GetProjectRoot;
  if RP='' then RP := ExtractFilePath(AInfo.FileName);
  AInfo.MethodName := FDialog.GetMethodName;
  if AInfo.MethodName='' then begin FDialog.SetPreviewText('Please enter a method name.'); Exit; end;
  FDialog.SetBusy(True);
  try
    FDialog.SetStatus('Saving files...'); TEditorHelper.SaveAllFiles;
    FDialog.SetStatus('Connecting to LSP...');
    Client := TLspManager.Instance.GetClient(RP, TEditorHelper.GetCurrentProjectDproj, DJ);
    FDialog.SetStatus('Refreshing source file...'); Client.RefreshDocument(AInfo.FileName); Sleep(500);

    // LSP-Warmup: Wait until ready (first request often returns "Request removed")
    FDialog.SetStatus('Waiting for LSP readiness...');
    for var Retry := 1 to 15 do
    begin
      try
        var H := Client.GetHover(AInfo.FileName, AInfo.StartLine-1, 0);
        // Any response (even empty) without exception means the LSP is ready
        Break;
      except
        on E: Exception do
        begin
          FDialog.SetStatus(Format('LSP warmup %d/15: %s', [Retry, E.Message]));
          Sleep(500);
        end;
      end;
    end;

    FDialog.SetStatus('Finding insertion point...'); FindInsertPoint(AInfo);

    // Validation: is the selected block sensible?
    FDialog.SetStatus('Validating selection...');
    var FileLines := TDelphiFileEncoding.ReadLines(AInfo.FileName);
    var ValidResult := TSelectionValidator.Validate(AInfo.SelectedText,
      FileLines, AInfo.StartLine, AInfo.EndLine, AInfo.InsertLine,
      AInfo.EnclosingClass);
    AInfo.DiagLog := AInfo.DiagLog + sLineBreak + '=== Validation ===' + sLineBreak +
      ValidResult.FormatIssues + sLineBreak;
    if ValidResult.HasErrors then
    begin
      FDialog.SetPreviewText('The selection cannot be sensibly extracted:' + sLineBreak +
        sLineBreak + ValidResult.FormatIssues + sLineBreak +
        'Please correct the selection and try again.');
      FDialog.SetStatus(Format('Validation failed: %d error(s).',
        [ValidResult.ErrorCount]));
      FDialog.EnableExtract(False);
      FInfoReady := False;
      Exit;
    end;

    if AInfo.EnclosingClass<>'' then FDialog.SetStatus('Class: '+AInfo.EnclosingClass+'. Analyzing...')
    else FDialog.SetStatus('Analyzing...');
    AnalyzeVariables(AInfo, Client);
    FDialog.SetStatus('Generating code...');
    // Store results for later live updates
    FCurrentInfo := AInfo;
    FInfoReady := True;
    UpdatePreview;
    FDialog.EnableExtract(True);
    FDialog.SetStatus(Format('Ready: %d parameter(s), %d local variable(s).',[Length(AInfo.Params),Length(AInfo.LocalVars)]));
  except
    on E: Exception do begin FDialog.SetPreviewText('ERROR: '+E.ClassName+': '+E.Message+sLineBreak+sLineBreak+AInfo.DiagLog); FDialog.SetStatus('Error.'); end;
  end;
  FDialog.SetBusy(False);
end;

procedure TLspExtractMethodWizard.Execute;
var
  Info: TExtractMethodInfo;
  Client: TLspClient;
  NMC, CC, CDC: string;
  ES: IOTAEditorServices;
  EB: IOTAEditBuffer;
  EP: IOTAEditPosition;
begin
  if not GetSelectedBlock(Info) then Exit;
  FInfoReady := False;
  FDialog := TExtractMethodDialog.CreateDialog(Application.MainForm, 'ExtractedMethod');
  try
    FDialog.OnNameChanged := OnMethodNameChange;
    FDialog.SetCheckContext(Info.FileName, TEditorHelper.GetProjectSourceFiles);
    FDialog.Show;
    Application.ProcessMessages;
    DoAnalyzeAndPreview(Info);
    FDialog.Hide;
    Application.ProcessMessages; // make Hide actually take effect
    if FDialog.ShowModal<>mrOk then Exit;
    Info.MethodName := FDialog.GetMethodName;
    if Info.MethodName='' then Exit;
    NMC := GenerateMethod(Info); CC := GenerateCall(Info);
    if Info.EnclosingClass<>'' then CDC := GenerateClassDeclaration(Info) else CDC := '';

    if not Supports(BorlandIDEServices, IOTAEditorServices, ES) then Exit;
    EB := ES.TopBuffer; if EB=nil then Exit;
    EP := EB.EditPosition; if EP=nil then Exit;
    // Replace the block with the call, byte-exact via Writer (no auto-indent).
    // Delete from (StartLine, 1) to the start of the line after EndLine so that
    // indentation and the trailing newline of the block are removed too.
    TEditorHelper.ReplaceSelection(Info.FileName,
      Info.StartLine, 1, Info.EndLine + 1, 1,
      CC + #13#10);

    // Insert the new method and class declaration via a direct Writer
    // (avoids the IDE's auto-indent)
    TEditorHelper.InsertTextAtLineStart(Info.FileName, Info.InsertLine, NMC);
    if (Info.ClassDeclLine>0) and (CDC<>'') then
      TEditorHelper.InsertTextAtLineStart(Info.FileName, Info.ClassDeclLine, CDC);

    // Remove the moved variables from the var declaration of the old method
    if Length(Info.LocalVars)>0 then
      RemoveLocalVarsFromDeclaration(Info);

    // Notify the form designer about class changes
    TEditorHelper.NotifyClassStructureChanged(Info.FileName);

    try Client := TLspManager.Instance.GetClient(TEditorHelper.GetProjectRoot,
      TEditorHelper.GetCurrentProjectDproj, TEditorHelper.FindDelphiLspJson);
      Client.RefreshDocument(Info.FileName); except end;
    MessageDlg('Method "'+Info.MethodName+'" extracted. Ctrl+Z to undo.', mtInformation, [mbOK], 0);
  finally FDialog.Free; FDialog := nil; end;
end;

end.
