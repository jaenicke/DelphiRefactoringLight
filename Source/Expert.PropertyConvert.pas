(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.PropertyConvert;

// Property converter (user request 2026-09-19), the pure half: the selected
// properties of a class / record switch between direct FIELD access and
// getter / setter METHODS.
//
//   to accessors:  property Name: string read FName write FName;
//               -> property Name: string read GetName write SetName;
//                  + "function GetName: string;" / "procedure SetName(const
//                  Value: string);" in the private section, and
//                  "Result := FName;" / "FName := Value;" implementations
//                  after the class's last method implementation.
//   to fields:     the other way round - only for TRIVIAL accessors (the
//                  body is exactly "Result := FName;" / "Exit(FName);" /
//                  "FName := Value;"), which are then removed. Refused when
//                  the accessor is used anywhere else (in the unit; for a
//                  non-private one also elsewhere - AExternalUse decides),
//                  or is virtual / override / overload / message.
//
// Refused with a reason, per property: array and indexed properties, class
// properties, declarations spanning lines, a read/write that is neither a
// field nor a method of THIS type (inherited, a record path), an accessor
// name that is already taken. Every property gets a row with what happens
// or why not - the dialog and the MCP tool show them.

interface

uses
  System.SysUtils;

type
  TPropConvDirection = (pcToAccessors, pcToFields);

  TPropConvItem = record
    Name: string;
    Line: Integer;      // 0-based property line
    Before: string;     // the property line now
    After: string;      // ... and after the conversion ('' = unchanged)
    Ok: Boolean;
    Note: string;       // what happens / why not
  end;

  TPropConvPlan = record
    Container: string;
    Items: TArray<TPropConvItem>;
    NewLines: TArray<string>;   // the whole unit afterwards (= input when nothing changes)
    Changed: Boolean;
    Error: string;              // no property found / not inside a type
    function OkCount: Integer;
  end;

  /// <summary>True when AName is used OUTSIDE the unit being converted.</summary>
  TExternalUseCheck = reference to function(const AName: string): Boolean;

/// <summary>Plans the conversion of every property declared on the lines
///  AFirst..ALast (0-based) of ALines. AGetter / ASetter select the sides
///  for pcToAccessors (pcToFields always tries both). AExternalUse may be
///  nil (then non-private accessors are never removed).</summary>
function PlanPropertyConversion(const ALines: TArray<string>; AFirst, ALast: Integer;
  ADirection: TPropConvDirection; AGetter, ASetter: Boolean;
  const AExternalUse: TExternalUseCheck): TPropConvPlan;

implementation

uses
  System.Classes, System.StrUtils, System.Math, System.Generics.Collections,
  Expert.PascalScanner, Expert.AutoImport, Expert.UnitIndex, Expert.InterfaceLinks;

type
  TMemberKind = (mkField, mkMethod, mkProperty, mkOther);

  TMemberInfo = record
    Name: string;
    Kind: TMemberKind;
    Line: Integer;
    HdrEnd: Integer;       // methods: last header line
    Visibility: string;    // '' (default section), 'private', 'strict private', ...
    Directives: string;    // methods: text after the signature
    Params: string;        // methods: parameter text
  end;

  TContainerScan = record
    Name: string;          // qualified, generic args kept ('TFoo<T>', 'TOuter.TInner')
    HeaderLine: Integer;
    EndLine: Integer;      // the type's 'end;'
    Members: TArray<TMemberInfo>;
    PrivateEnd: Integer;   // insert new private members BEFORE this line, -1 = none
    FirstVisLine: Integer; // first visibility keyword line, -1 = none
    MemberIndent: string;
    function Find(const AName: string; out AInfo: TMemberInfo): Boolean;
  end;

  TEditMap = class
  public
    Deleted: TDictionary<Integer, Boolean>;
    Replaced: TDictionary<Integer, string>;
    InsertBefore: TObjectDictionary<Integer, TStringList>;
    constructor Create;
    destructor Destroy; override;
    procedure Insert(ALine: Integer; const ATexts: TArray<string>);
    function Apply(const ALines: TArray<string>): TArray<string>;
  end;

{ TPropConvPlan }

function TPropConvPlan.OkCount: Integer;
begin
  Result := 0;
  for var It in Items do
    if It.Ok then Inc(Result);
end;

{ TContainerScan }

function TContainerScan.Find(const AName: string; out AInfo: TMemberInfo): Boolean;
begin
  for var M in Members do
    if SameText(M.Name, AName) then
    begin
      AInfo := M;
      Exit(True);
    end;
  Result := False;
end;

{ TEditMap }

constructor TEditMap.Create;
begin
  inherited;
  Deleted := TDictionary<Integer, Boolean>.Create;
  Replaced := TDictionary<Integer, string>.Create;
  InsertBefore := TObjectDictionary<Integer, TStringList>.Create([doOwnsValues]);
end;

destructor TEditMap.Destroy;
begin
  InsertBefore.Free;
  Replaced.Free;
  Deleted.Free;
  inherited;
end;

procedure TEditMap.Insert(ALine: Integer; const ATexts: TArray<string>);
var
  SL: TStringList;
begin
  if not InsertBefore.TryGetValue(ALine, SL) then
  begin
    SL := TStringList.Create;
    InsertBefore.Add(ALine, SL);
  end;
  for var T in ATexts do SL.Add(T);
end;

function TEditMap.Apply(const ALines: TArray<string>): TArray<string>;
var
  Res: TList<string>;
  SL: TStringList;
begin
  Res := TList<string>.Create;
  try
    for var I := 0 to Length(ALines) do
    begin
      if InsertBefore.TryGetValue(I, SL) then
        for var T in SL do Res.Add(T);
      if I > High(ALines) then Break;
      if Deleted.ContainsKey(I) then Continue;
      var R: string;
      if Replaced.TryGetValue(I, R) then Res.Add(R) else Res.Add(ALines[I]);
    end;
    Result := Res.ToArray;
  finally
    Res.Free;
  end;
end;

// ---------------------------------------------------------------------------

function Indent(const S: string): string;
begin
  var I := 1;
  while (I <= Length(S)) and CharInSet(S[I], [' ', #9]) do Inc(I);
  Result := Copy(S, 1, I - 1);
end;

function FirstWordU(const S: string): string;
begin
  var T := TrimLeft(S);
  var I := 1;
  while (I <= Length(T)) and IsIdentChar(T[I]) do Inc(I);
  Result := UpperCase(Copy(T, 1, I - 1));
end;

function IsVisibilityLine(const AMasked: string; out AVis: string): Boolean;
begin
  var U := UpperCase(Trim(AMasked));
  if U.StartsWith('STRICT ') then
    U := 'STRICT ' + Trim(Copy(U, 8, MaxInt));
  Result := (U = 'PRIVATE') or (U = 'PROTECTED') or (U = 'PUBLIC') or (U = 'PUBLISHED') or
    (U = 'AUTOMATED') or (U = 'STRICT PRIVATE') or (U = 'STRICT PROTECTED');
  if Result then AVis := LowerCase(U);
end;

// the body of the type whose header is at AHeader: members at depth 1
function ScanContainer(const ALines, AMasked: TArray<string>; AHeader: Integer;
  const AName: string): TContainerScan;
var
  Depth: Integer;
  Vis, V, Kind, Q, P, R: string;
  IsCM: Boolean;
begin
  Result := Default(TContainerScan);
  Result.Name := AName;
  Result.HeaderLine := AHeader;
  Result.EndLine := -1;
  Result.PrivateEnd := -1;
  Result.FirstVisLine := -1;
  Depth := 1;
  Vis := '';
  var InPrivate := False;
  var L := AHeader + 1;
  while L <= High(AMasked) do
  begin
    var T := Trim(AMasked[L]);
    var U := UpperCase(T);
    if (ClassOpenerName(T) <> '') then
    begin
      Inc(Depth);
      Inc(L);
      Continue;
    end;
    if (U = 'END;') or (U = 'END') then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        Result.EndLine := L;
        if InPrivate then Result.PrivateEnd := L;
        Break;
      end;
      Inc(L);
      Continue;
    end;
    if Depth > 1 then begin Inc(L); Continue; end;
    if IsVisibilityLine(T, V) then
    begin
      if Result.FirstVisLine < 0 then Result.FirstVisLine := L;
      if InPrivate and (Result.PrivateEnd < 0) then Result.PrivateEnd := L;
      Vis := V;
      InPrivate := ((V = 'private') or (V = 'strict private')) and (Result.PrivateEnd < 0);
      Inc(L);
      Continue;
    end;
    if T = '' then begin Inc(L); Continue; end;
    var M := Default(TMemberInfo);
    M.Line := L;
    M.HdrEnd := L;
    M.Visibility := Vis;
    if Result.MemberIndent = '' then Result.MemberIndent := Indent(ALines[L]);
    var W := FirstWordU(T);
    if (W = 'PROPERTY') or ((W = 'CLASS') and U.StartsWith('CLASS PROPERTY')) then
    begin
      M.Kind := mkProperty;
      var S := TrimLeft(Copy(T, Pos('PROPERTY', U) + 8, MaxInt));
      var I := 1;
      while (I <= Length(S)) and IsIdentChar(S[I]) do Inc(I);
      M.Name := Copy(S, 1, I - 1);
    end
    else if IsHeaderLine(T, Kind, IsCM) then
    begin
      M.Kind := mkMethod;
      var Hdr := CollectHeader(ALines, L, M.HdrEnd);
      if (Hdr <> '') and ParseHeader(Hdr, Kind, Q, P, R) then
      begin
        M.Name := Copy(Q, LastDelimiter('.', Q) + 1, MaxInt);
        M.Params := P;
      end;
      // directives: after the signature's ';' on the last header line (and
      // on directive-only follow-up lines)
      var Last := AMasked[M.HdrEnd];
      var D := 0;
      var Semi := 0;
      for var K := 1 to Length(Last) do
        case Last[K] of
          '(', '[': Inc(D);
          ')', ']': Dec(D);
          ';': if D = 0 then begin Semi := K; Break; end;
        end;
      M.Directives := ' ' + Copy(Last, Semi + 1, MaxInt);
      L := M.HdrEnd;
    end
    else if Pos(':', T) > 0 then
    begin
      // "FA, FB: Integer;" / "class var FX: T;" / "var FX: T;"
      var S := T;
      if U.StartsWith('CLASS VAR ') then S := Trim(Copy(S, 11, MaxInt))
      else if U.StartsWith('VAR ') then S := Trim(Copy(S, 5, MaxInt));
      var Names := Copy(S, 1, Pos(':', S) - 1);
      for var N in Names.Split([',']) do
        if IsIdentifier(Trim(N)) then
        begin
          var F := M;
          F.Kind := mkField;
          F.Name := Trim(N);
          Result.Members := Result.Members + [F];
        end;
      Inc(L);
      Continue;
    end
    else
      M.Kind := mkOther;
    if M.Name <> '' then Result.Members := Result.Members + [M];
    Inc(L);
  end;
end;

type
  // one value replaced in the property line (applied right-to-left)
  TRepl = record
    P, L: Integer;
    S: string;
  end;

  TPropDecl = record
    Name, TypeText, ReadVal, WriteVal: string;
    ReadPos, WritePos: Integer;   // 1-based offsets of the values in the line
    Refuse: string;
  end;

// "property Name: T read R write W default X;" - one line
function ParsePropertyLine(const ALine, AMasked: string): TPropDecl;
var
  S: TPascalScanner;
  T: TPasToken;
  Toks: TArray<TPasToken>;
begin
  Result := Default(TPropDecl);
  S := TPascalScanner.Create(AMasked);
  try
    while S.Next(T) do Toks := Toks + [T];
  finally
    S.Free;
  end;
  if Length(Toks) = 0 then Exit;
  var I: Integer;
  if Toks[0].IsWord('class') then
  begin
    Result.Refuse := 'class property - not supported';
    Exit;
  end;
  if not Toks[0].IsWord('property') or (Length(Toks) < 2) then
  begin
    Result.Refuse := 'not a property';
    Exit;
  end;
  Result.Name := Toks[1].Text;
  if (Length(Toks) > 2) and (Toks[2].IsSymbol('[') or Toks[2].IsSymbol('(.')) then
  begin
    Result.Refuse := 'array property - not supported';
    Exit;
  end;
  if (Length(Toks) < 3) or not Toks[2].IsSymbol(':') then
  begin
    Result.Refuse := 'no type (a redeclared property) - nothing to convert';
    Exit;
  end;
  if not Toks[High(Toks)].IsSymbol(';') then
  begin
    Result.Refuse := 'the declaration spans several lines - not supported';
    Exit;
  end;
  // the type runs up to the first specifier keyword
  I := 3;
  var TypeEnd := Length(ALine) + 1;
  while I <= High(Toks) do
  begin
    var W := UpperCase(Toks[I].Text);
    if (Toks[I].Kind = ptIdent) and ((W = 'READ') or (W = 'WRITE') or (W = 'STORED') or
       (W = 'DEFAULT') or (W = 'NODEFAULT') or (W = 'INDEX') or (W = 'IMPLEMENTS') or
       (W = 'READONLY') or (W = 'WRITEONLY')) then
    begin
      TypeEnd := Toks[I].Pos;
      Break;
    end;
    Inc(I);
  end;
  Result.TypeText := Trim(Copy(ALine, Toks[2].Pos + 1, TypeEnd - Toks[2].Pos - 1));
  while I <= High(Toks) do
  begin
    var W := UpperCase(Toks[I].Text);
    if W = 'INDEX' then
    begin
      Result.Refuse := 'indexed property - not supported';
      Exit;
    end;
    if ((W = 'READ') or (W = 'WRITE')) and (I < High(Toks)) then
    begin
      // the value: one identifier; a path ("FRec.X") is not ours to convert
      var V := Toks[I + 1];
      var Path := (I + 2 <= High(Toks)) and Toks[I + 2].IsSymbol('.');
      var Val := V.Text;
      if Path then Val := '';
      if W = 'READ' then begin Result.ReadVal := Val; Result.ReadPos := V.Pos; end
      else begin Result.WriteVal := Val; Result.WritePos := V.Pos; end;
      if Path then
      begin
        Result.Refuse := LowerCase(W) + ' uses a field path - not supported';
        Exit;
      end;
      Inc(I, 2);
      Continue;
    end;
    Inc(I);
  end;
end;

// line with the value at APos (1-based, length ALen) replaced
function ReplaceAt(const ALine: string; APos, ALen: Integer; const ANew: string): string;
begin
  Result := Copy(ALine, 1, APos - 1) + ANew + Copy(ALine, APos + ALen, MaxInt);
end;

// the trivial body of an accessor: the one statement between begin and end
function TrivialBodyStatement(const ALines, AMasked: TArray<string>; AImplLine: Integer;
  out AStatement: string; out AFirst, ALast: Integer): Boolean;
begin
  Result := False;
  AStatement := '';
  if not FindEnclosingRoutineRange(string.Join(#10, ALines), AImplLine, AFirst, ALast) then Exit;
  var HdrEnd: Integer;
  if CollectHeader(ALines, AImplLine, HdrEnd) = '' then Exit;
  // the body must start right after the header: no var / const / nested routine
  var B := HdrEnd + 1;
  while (B < ALast) and (Trim(AMasked[B]) = '') do Inc(B);
  if not SameText(Trim(AMasked[B]), 'begin') then Exit;
  var S := '';
  for var L := B + 1 to ALast - 1 do
    S := S + Trim(AMasked[L]);
  S := UpperCase(S.Replace(' ', '').Replace(#9, ''));
  // the last line must close the routine
  if not UpperCase(Trim(AMasked[ALast])).StartsWith('END') then Exit;
  AStatement := S;
  Result := True;
end;

function PlanPropertyConversion(const ALines: TArray<string>; AFirst, ALast: Integer;
  ADirection: TPropConvDirection; AGetter, ASetter: Boolean;
  const AExternalUse: TExternalUseCheck): TPropConvPlan;
var
  Masked: TArray<string>;
  Scan: TContainerScan;
  Edits: TEditMap;
  NewDecls, NewImpls: TArray<string>;
  Info: TMemberInfo;
begin
  Result := Default(TPropConvPlan);
  Result.NewLines := ALines;
  Masked := MaskCommentsAndStrings(ALines);
  if (AFirst < 0) or (ALast > High(ALines)) or (ALast < AFirst) then
  begin
    Result.Error := 'the selection is outside the unit';
    Exit;
  end;
  // the container: the type whose body holds the first selected line
  var Container := EnclosingContainerName(ALines, AFirst);
  if Container = '' then
  begin
    Result.Error := 'the selection is not inside a class or record declaration';
    Exit;
  end;
  var Header := -1;
  var Inner := Container;
  if LastDelimiter('.', Inner) > 0 then Inner := Copy(Inner, LastDelimiter('.', Inner) + 1, MaxInt);
  // (a dot inside generic arguments is not expected here)
  for var L := AFirst downto 0 do
    if SameText(ClassOpenerName(Trim(Masked[L])), Inner) then
    begin
      Header := L;
      Break;
    end;
  if Header < 0 then
  begin
    Result.Error := 'the declaration of ' + Container + ' was not found';
    Exit;
  end;
  Scan := ScanContainer(ALines, Masked, Header, Container);
  Result.Container := Container;
  if Scan.EndLine < 0 then
  begin
    Result.Error := 'the end of ' + Container + ' was not found';
    Exit;
  end;

  // the qualifier of implementation headers: generic arguments kept
  var BaseName := Inner;
  if Pos('<', BaseName) > 0 then BaseName := Copy(BaseName, 1, Pos('<', BaseName) - 1);

  Edits := TEditMap.Create;
  try
    var Content := string.Join(#10, ALines);
    for var PL := AFirst to Min(ALast, Scan.EndLine) do
    begin
      if FirstWordU(Masked[PL]) <> 'PROPERTY' then
        if not UpperCase(Trim(Masked[PL])).StartsWith('CLASS PROPERTY') then Continue;
      var Item := Default(TPropConvItem);
      Item.Line := PL;
      Item.Before := ALines[PL];
      var PD := ParsePropertyLine(ALines[PL], Masked[PL]);
      Item.Name := PD.Name;
      if PD.Refuse <> '' then
      begin
        Item.Note := PD.Refuse;
        Result.Items := Result.Items + [Item];
        Continue;
      end;
      var NewLine := ALines[PL];
      var Repl: TArray<TRepl> := nil;
      var Notes: TArray<string> := nil;

      if ADirection = pcToAccessors then
      begin
        var Sides: TArray<Boolean> := [False, True];   // False = read, True = write
        for var IsWrite in Sides do
        begin
          if IsWrite and not ASetter then Continue;
          if not IsWrite and not AGetter then Continue;
          var Val := IfThen(IsWrite, PD.WriteVal, PD.ReadVal);
          var Side := IfThen(IsWrite, 'write', 'read');
          if Val = '' then begin Notes := Notes + ['no ' + Side]; Continue; end;
          if not Scan.Find(Val, Info) then
          begin
            Notes := Notes + [Side + ' ' + Val + ' is not declared in ' + Container];
            Continue;
          end;
          if Info.Kind <> mkField then
          begin
            Notes := Notes + [Side + ' already uses a method'];
            Continue;
          end;
          var Acc := IfThen(IsWrite, 'Set', 'Get') + PD.Name;
          if Scan.Find(Acc, Info) then
          begin
            Notes := Notes + [Acc + ' already exists'];
            Continue;
          end;
          // the declaration, the implementation, the property
          if IsWrite then
          begin
            NewDecls := NewDecls + [Scan.MemberIndent + 'procedure ' + Acc +
              '(const Value: ' + PD.TypeText + ');'];
            NewImpls := NewImpls + ['', 'procedure ' + Container + '.' + Acc +
              '(const Value: ' + PD.TypeText + ');', 'begin', '  ' + Val + ' := Value;', 'end;'];
            var R := Default(TRepl); R.P := PD.WritePos; R.L := Length(Val); R.S := Acc;
            Repl := Repl + [R];
          end
          else
          begin
            NewDecls := NewDecls + [Scan.MemberIndent + 'function ' + Acc + ': ' + PD.TypeText + ';'];
            NewImpls := NewImpls + ['', 'function ' + Container + '.' + Acc + ': ' + PD.TypeText + ';',
              'begin', '  Result := ' + Val + ';', 'end;'];
            var R := Default(TRepl); R.P := PD.ReadPos; R.L := Length(Val); R.S := Acc;
            Repl := Repl + [R];
          end;
          Notes := Notes + [Side + ' ' + Val + ' -> ' + Acc];
        end;
      end
      else
      begin
        // accessors -> fields: both sides, each only when trivial and unused
        var Sides: TArray<Boolean> := [False, True];
        for var IsWrite in Sides do
        begin
          var Val := IfThen(IsWrite, PD.WriteVal, PD.ReadVal);
          var Side := IfThen(IsWrite, 'write', 'read');
          if Val = '' then Continue;
          if not Scan.Find(Val, Info) then
          begin
            Notes := Notes + [Side + ' ' + Val + ' is not declared in ' + Container];
            Continue;
          end;
          if Info.Kind = mkField then Continue;   // already a field
          if Info.Kind <> mkMethod then
          begin
            Notes := Notes + [Side + ' ' + Val + ' is neither a field nor a method'];
            Continue;
          end;
          var Dir := UpperCase(Info.Directives);
          var Bad := '';
          for var KW in ['VIRTUAL', 'OVERRIDE', 'DYNAMIC', 'ABSTRACT', 'OVERLOAD', 'MESSAGE',
            'REINTRODUCE'] do
            if HasWholeWordCI(Dir, KW) then Bad := LowerCase(KW);
          if Bad <> '' then
          begin
            Notes := Notes + [Val + ' is ' + Bad + ' - kept'];
            Continue;
          end;
          if Info.HdrEnd <> Info.Line then
          begin
            Notes := Notes + [Val + ': multi-line declaration - kept'];
            Continue;
          end;
          var ImplL := FindMethodImplLine(ALines, BaseName, Val);
          if ImplL < 0 then
          begin
            Notes := Notes + [Val + ': no implementation found - kept'];
            Continue;
          end;
          var Stmt: string;
          var BF, BL: Integer;
          if not TrivialBodyStatement(ALines, Masked, ImplL, Stmt, BF, BL) then
          begin
            Notes := Notes + [Val + ': body is not trivial - kept'];
            Continue;
          end;
          // the field the trivial body touches
          var Field := '';
          if not IsWrite then
          begin
            if Stmt.StartsWith('RESULT:=') then Field := Copy(Stmt, 9, MaxInt)
            else if Stmt.StartsWith('EXIT(') then Field := Copy(Stmt, 6, MaxInt).Replace(')', '');
            Field := Field.Replace(';', '');
          end
          else
          begin
            // "FName:=Value;" - Value = the setter's parameter name
            var PName := Info.Params;
            var C := Pos(':', PName);
            if C > 0 then PName := Copy(PName, 1, C - 1);
            PName := Trim(PName);
            for var Pre in ['CONST ', 'VAR ', 'CONSTREF '] do
              if UpperCase(PName).StartsWith(Pre) then PName := Trim(Copy(PName, Length(Pre) + 1, MaxInt));
            var Asg := Pos(':=', Stmt);
            if (Asg > 0) and SameText(Copy(Stmt, Asg + 2, MaxInt).Replace(';', ''), PName) then
              Field := Copy(Stmt, 1, Asg - 1);
          end;
          var FieldInfo: TMemberInfo;
          if (Field = '') or not IsIdentifier(Field) or not Scan.Find(Field, FieldInfo) or
             (FieldInfo.Kind <> mkField) then
          begin
            Notes := Notes + [Val + ': body is not a plain field ' +
              IfThen(IsWrite, 'assignment', 'read') + ' - kept'];
            Continue;
          end;
          Field := FieldInfo.Name;   // declared spelling
          // used anywhere else in the unit?
          var Uses_ := 0;
          for var L in CodeWordLines(ALines, Val) do
            if (L <> Info.Line) and (L <> ImplL) and (L <> PL) then Inc(Uses_);
          if Uses_ > 0 then
          begin
            Notes := Notes + [Val + ' is used elsewhere in the unit - kept'];
            Continue;
          end;
          if not ((Info.Visibility = 'private') or (Info.Visibility = 'strict private')) then
            if not Assigned(AExternalUse) or AExternalUse(Val) then
            begin
              Notes := Notes + [Val + ' is ' + IfThen(Info.Visibility = '', 'visible',
                Info.Visibility) + ' and may be used in other units - kept'];
              Continue;
            end;
          // convert: the property, and the accessor goes (declaration +
          // implementation, with a blank line of the implementation)
          var R := Default(TRepl);
          R.P := IfThen(IsWrite, PD.WritePos, PD.ReadPos);
          R.L := Length(Val);
          R.S := Field;
          Repl := Repl + [R];
          Edits.Deleted.AddOrSetValue(Info.Line, True);
          var D1 := Info.Line;
          while (D1 > 0) and Trim(ALines[D1 - 1]).StartsWith('///') do
          begin
            Dec(D1);
            Edits.Deleted.AddOrSetValue(D1, True);
          end;
          var LastDel := BL;
          if (LastDel < High(ALines)) and (Trim(ALines[LastDel + 1]) = '') and (BF > 0) and
             (Trim(ALines[BF - 1]) = '') then Inc(LastDel);
          for var K := BF to LastDel do Edits.Deleted.AddOrSetValue(K, True);
          Notes := Notes + [Side + ' ' + Val + ' -> ' + Field + ' (' + Val + ' removed)'];
        end;
      end;

      // right-to-left, so an earlier offset stays valid
      for var A := 0 to High(Repl) - 1 do
        for var B := A + 1 to High(Repl) do
          if Repl[B].P > Repl[A].P then
          begin
            var Tmp := Repl[A]; Repl[A] := Repl[B]; Repl[B] := Tmp;
          end;
      for var R in Repl do
        NewLine := ReplaceAt(NewLine, R.P, R.L, R.S);
      if NewLine <> ALines[PL] then
      begin
        Item.Ok := True;
        Item.After := NewLine;
        Edits.Replaced.AddOrSetValue(PL, NewLine);
      end;
      if Length(Notes) = 0 then Notes := ['nothing to convert'];
      Item.Note := string.Join('; ', Notes);
      Result.Items := Result.Items + [Item];
    end;

    if Length(Result.Items) = 0 then
    begin
      Result.Error := 'no property declaration in the selection';
      Exit;
    end;

    // new members: into the private section (created where it cannot change
    // anybody else's visibility), implementations after the type's last one
    if Length(NewDecls) > 0 then
    begin
      if Scan.PrivateEnd >= 0 then
      begin
        // behind the section's last member, not behind trailing blank lines
        var At := Scan.PrivateEnd;
        while (At - 1 > Scan.HeaderLine) and (Trim(ALines[At - 1]) = '') do Dec(At);
        Edits.Insert(At, NewDecls);
      end
      else
      begin
        var VisIndent := Scan.MemberIndent;
        if Length(VisIndent) >= 2 then VisIndent := Copy(VisIndent, 1, Length(VisIndent) - 2);
        var At := Scan.FirstVisLine;
        if At < 0 then At := Scan.EndLine;
        Edits.Insert(At, [VisIndent + 'private'] + NewDecls);
      end;
      // implementations
      var LastImplEnd := -1;
      var ImplStart := ImplementationLineOf(ALines);
      if ImplStart <> MaxInt then
        for var L := ImplStart + 1 to High(ALines) do
        begin
          var U := UpperCase(TrimLeft(Masked[L]));
          if U.StartsWith('CLASS ') then U := TrimLeft(Copy(U, 7, MaxInt));
          if not (U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ') or
             U.StartsWith('CONSTRUCTOR ') or U.StartsWith('DESTRUCTOR ')) then Continue;
          var Q := TrimLeft(Copy(U, Pos(' ', U) + 1, MaxInt)).Replace(' ', '');
          var Q2 := Q;
          var LT := Pos('<', Q2);
          if LT > 0 then Q2 := Copy(Q2, 1, LT - 1) + Copy(Q2, Pos('>', Q2) + 1, MaxInt);
          if not Q2.StartsWith(UpperCase(BaseName) + '.') then Continue;
          var F, E: Integer;
          if FindEnclosingRoutineRange(Content, L, F, E) and (F = L) and (E > LastImplEnd) then
            LastImplEnd := E;
        end;
      if LastImplEnd >= 0 then
        Edits.Insert(LastImplEnd + 1, NewImpls)
      else
      begin
        var At := ImplInsertLine(ALines);
        if At < 0 then
        begin
          Result.Error := 'the unit has no implementation section for the accessors';
          Result.Items := nil;
          Exit;
        end;
        // symmetric with the way back: a blank line already standing before
        // 'end.' separates the first block, each block is followed by one
        var Ins := NewImpls;
        if (At > 0) and (Trim(ALines[At - 1]) = '') then Ins := Copy(Ins, 1, MaxInt);
        Edits.Insert(At, Ins + ['']);
      end;
    end;

    if Result.OkCount > 0 then
    begin
      Result.NewLines := Edits.Apply(ALines);
      Result.Changed := True;
    end;
  finally
    Edits.Free;
  end;
end;

end.
