(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.SafeDeletePlan;

// SAFE DELETE, the pure half (idea: Ian Branch, issue #11): WHAT would be
// deleted for the symbol declared at a given line, and which properties of
// the declaration forbid deleting it at all. The usage check (text scan +
// DelphiLSP verification, form files) lives in Expert.SafeDelete.
//
// Supported: methods (class / record / interface), free routines,
// fields, variables, constants, properties and single-line types.
// Refused (a "veto" with the reason): overloaded, virtual / dynamic /
// abstract / override / message methods (dispatch can reach them without
// a textual call), published members (streaming / RTTI), multi-line
// declarations of data, and types with a body.

interface

uses
  System.SysUtils;

type
  TSafeDeleteKind = (sdkUnknown, sdkMethod, sdkRoutine, sdkField, sdkProperty,
    sdkVariable, sdkConstant, sdkType);

  /// <summary>Whole lines FirstLine..LastLine (0-based, inclusive) are
  ///  removed - or, with HasReplacement, replaced by the ONE line
  ///  Replacement (a name dropped from "A, B, C: Integer;").</summary>
  TSafeDeleteEdit = record
    FirstLine, LastLine: Integer;
    HasReplacement: Boolean;
    Replacement: string;
  end;

  TSafeDeleteSymbol = record
    Name: string;
    Kind: TSafeDeleteKind;
    Container: string;        // owning class / record / interface ('' = unit level)
    ContainerLine: Integer;   // its header line, -1 = none
    IsInterfaceMember: Boolean;
    Parents: TArray<string>;  // the container's parent list (interface veto)
    DeclLine: Integer;        // the declaration
    ImplLine: Integer;        // implementation header, -1 = none
    LocalFirst, LocalLast: Integer;  // a routine-local symbol: its routine; else -1
    Edits: TArray<TSafeDeleteEdit>;
    Vetoes: TArray<string>;
  end;

/// <summary>Plans the deletion of AName declared on ADeclLine0 of ALines
///  (the declaring unit). The line may also be the IMPLEMENTATION header of
///  a method or routine - the declaration is then located. False with
///  AWhy when the line does not declare AName or the shape is not
///  supported; vetoes do NOT make it fail (ASym.Vetoes lists them).</summary>
function PlanSafeDeleteSymbol(const ALines: TArray<string>; ADeclLine0: Integer;
  const AName: string; out ASym: TSafeDeleteSymbol; out AWhy: string): Boolean;

/// <summary>ALine declares AName (routine header, property, "A, B: T",
///  "X = ..."). Used when DelphiLSP gives no definition at the caret.</summary>
function LineDeclaresName(const ALine, AName: string): Boolean;

/// <summary>ALines after AEdits (applied bottom-up; overlapping edits are
///  not expected - the planner never produces them).</summary>
function ApplySafeDeleteEdits(const ALines: TArray<string>;
  const AEdits: TArray<TSafeDeleteEdit>): TArray<string>;

/// <summary>0-based lines of a TEXT form file (.dfm/.fmx) that mention
///  AName as a whole word - component names, event handler bindings,
///  references. Deliberately text-level: a form binds by NAME.</summary>
function FormTextMentions(const AText, AName: string): TArray<Integer>;

/// <summary>Human-readable kind ("method", "field", ...).</summary>
function SafeDeleteKindText(AKind: TSafeDeleteKind): string;

implementation

uses
  System.Classes, System.StrUtils, System.Math, Expert.AutoImport, Expert.UsesEditor,
  Expert.UnitIndex;

function IsIdentCh(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_']);
end;

function SafeDeleteKindText(AKind: TSafeDeleteKind): string;
begin
  case AKind of
    sdkMethod: Result := 'method';
    sdkRoutine: Result := 'routine';
    sdkField: Result := 'field';
    sdkProperty: Result := 'property';
    sdkVariable: Result := 'variable';
    sdkConstant: Result := 'constant';
    sdkType: Result := 'type';
  else
    Result := 'symbol';
  end;
end;

// The line's CODE: comments ({ }, (* *), //) and string contents blanked
// (MaskCommentsAndStrings keeps the length), trimmed. Analysis only - edits
// always use the raw line.
function Code(const S: string): string;
begin
  Result := Trim(MaskCommentsAndStrings([S])[0]);
end;

function FirstWordU(const S: string): string;
var
  T: string;
  I: Integer;
begin
  T := TrimLeft(S);
  I := 1;
  while (I <= Length(T)) and IsIdentCh(T[I]) do Inc(I);
  Result := UpperCase(Copy(T, 1, I - 1));
end;

// Header line of the class / record / interface / object whose BODY
// contains ALine0 (-1 = none), with its name.
function EnclosingTypeHeader(const ALines: TArray<string>; ALine0: Integer;
  out AName: string; out AIsInterface: Boolean): Integer;
var
  Depth: Integer;
  T, U, Rest: string;
  P: Integer;
begin
  Result := -1;
  AName := '';
  AIsInterface := False;
  Depth := 0;
  for var L := ALine0 - 1 downto 0 do
  begin
    T := Code(ALines[L]);
    U := UpperCase(T);
    if (U = 'END;') or (U = 'END') then
    begin
      Inc(Depth);
      Continue;
    end;
    if (U = 'IMPLEMENTATION') or (U = 'INTERFACE') then Exit;
    P := Pos('=', T);
    if P > 1 then
    begin
      Rest := UpperCase(Trim(Copy(T, P + 1, MaxInt)));
      if Rest.StartsWith('PACKED ') then Rest := TrimLeft(Copy(Rest, 8, MaxInt));
      var Opens := (Rest.StartsWith('CLASS') and not Rest.StartsWith('CLASS OF'))
        or Rest.StartsWith('RECORD') or Rest.StartsWith('OBJECT')
        or Rest.StartsWith('INTERFACE') or Rest.StartsWith('DISPINTERFACE');
      if Opens and not Rest.EndsWith(';') then   // not a forward / alias
      begin
        if Depth = 0 then
        begin
          AName := Trim(Copy(T, 1, P - 1));
          var LT := Pos('<', AName);
          if LT > 0 then AName := Trim(Copy(AName, 1, LT - 1));
          AIsInterface := Rest.StartsWith('INTERFACE') or Rest.StartsWith('DISPINTERFACE');
          Exit(L);
        end;
        Dec(Depth);
      end;
    end;
  end;
end;

// Parent list of a type header line "TFoo = class(TBase, IA, IB)".
function HeaderParents(const ALine: string): TArray<string>;
var
  A, B: Integer;
begin
  Result := nil;
  A := Pos('(', ALine);
  B := Pos(')', ALine);
  if (A = 0) or (B <= A) then Exit;
  for var S in Copy(ALine, A + 1, B - A - 1).Split([',']) do
    if Trim(S) <> '' then Result := Result + [Trim(S)];
end;

// Implementation header of AQualified ("TFoo.Bar" or a plain name for a
// top-level routine), -1 = none. Nested routines (indented) never match a
// plain name.
function FindImplHeader(const ALines: TArray<string>; const AQualified: string): Integer;
var
  Kind, Hdr, Q, P, R: string;
  IsCM: Boolean;
  HdrEnd, Impl: Integer;
  Plain: Boolean;
begin
  Result := -1;
  Impl := ImplementationLineOf(ALines);
  if Impl = MaxInt then Exit;
  Plain := Pos('.', AQualified) = 0;
  for var L := Impl + 1 to High(ALines) do
  begin
    if not IsHeaderLine(StripLineComment(ALines[L]), Kind, IsCM) then Continue;
    if Plain and (ALines[L] <> TrimLeft(ALines[L])) then Continue;   // nested
    Hdr := CollectHeader(ALines, L, HdrEnd);
    if (Hdr = '') or not ParseHeader(Hdr, Kind, Q, P, R) then Continue;
    // "TOuter.TInner.Bar" also matches a container given as "TInner"
    var LT := Pos('<', Q);
    if LT > 0 then   // TList<T>.Add -> TList.Add
      Q := Copy(Q, 1, LT - 1) + Copy(Q, Pos('>', Q) + 1, MaxInt);
    if SameText(Q, AQualified) or (not Plain and Q.ToUpper.EndsWith('.' + AQualified.ToUpper)) then
      Exit(L);
  end;
end;

// Words after the signature's ';' on the header's last line (+ the lines
// that only hold directives): 'virtual', 'override', ...
function HeaderDirectives(const ALines: TArray<string>; AStart, AHdrEnd: Integer): string;
var
  S: string;
  Depth, P: Integer;
begin
  S := StripLineComment(ALines[AHdrEnd]);
  Depth := 0;
  P := 0;
  var From := 1;
  if AHdrEnd <> AStart then From := 1;
  for var I := From to Length(S) do
    case S[I] of
      '(', '[': Inc(Depth);
      ')', ']': Dec(Depth);
      ';': if Depth = 0 then begin P := I; Break; end;
    end;
  Result := ' ' + Copy(S, P + 1, MaxInt);
  // directive-only follow-up lines ("      override;")
  for var L := AHdrEnd + 1 to High(ALines) do
  begin
    var T := Code(ALines[L]);
    var W := FirstWordU(T);
    if (W = 'VIRTUAL') or (W = 'OVERRIDE') or (W = 'OVERLOAD') or (W = 'ABSTRACT') or
       (W = 'DYNAMIC') or (W = 'REINTRODUCE') or (W = 'STDCALL') or (W = 'CDECL') or
       (W = 'INLINE') or (W = 'MESSAGE') or (W = 'STATIC') or (W = 'FINAL') then
      Result := Result + ' ' + T
    else
      Break;
  end;
end;

// Visibility keyword in force at ALine0 inside the type body starting at
// AHeader ('' = the default section at the top).
function VisibilityAt(const ALines: TArray<string>; AHeader, ALine0: Integer): string;
var
  Depth: Integer;
begin
  Result := '';
  Depth := 0;
  for var L := AHeader + 1 to ALine0 - 1 do
  begin
    var T := Code(ALines[L]);
    var U := UpperCase(T);
    if (ClassOpenerName(T) <> '') or ((Pos('= RECORD', U) > 0) and not U.EndsWith(';')) then
      Inc(Depth)
    else if (U = 'END;') or (U = 'END') then
      Dec(Depth)
    else if Depth = 0 then
    begin
      if U.StartsWith('STRICT ') then U := TrimLeft(Copy(U, 8, MaxInt));
      if (U = 'PRIVATE') or (U = 'PROTECTED') or (U = 'PUBLIC') or (U = 'PUBLISHED') or
         (U = 'AUTOMATED') then
        Result := LowerCase(U);
    end;
  end;
end;

// The section keyword ('var', 'const', 'type', ...) that governs a unit-
// or routine-level declaration on ALine0 ('' = unknown).
function SectionAt(const ALines: TArray<string>; ALine0: Integer): string;
begin
  Result := '';
  for var L := ALine0 downto 0 do
  begin
    var W := FirstWordU(Code(ALines[L]));
    if (W = 'VAR') or (W = 'CONST') or (W = 'TYPE') or (W = 'THREADVAR') or
       (W = 'RESOURCESTRING') then
      Exit(LowerCase(W));
    if (W = 'BEGIN') or (W = 'IMPLEMENTATION') or (W = 'INTERFACE') or
       (W = 'PROCEDURE') or (W = 'FUNCTION') or (W = 'CONSTRUCTOR') or
       (W = 'DESTRUCTOR') then
      Exit;
  end;
end;

procedure AddEdit(var ASym: TSafeDeleteSymbol; AFirst, ALast: Integer;
  const ALines: TArray<string>);
var
  E: TSafeDeleteEdit;
begin
  // XML doc comments directly above belong to the declaration
  while (AFirst > 0) and Trim(ALines[AFirst - 1]).StartsWith('///') do Dec(AFirst);
  // no double blank line left behind
  if (AFirst > 0) and (ALast < High(ALines)) and (Trim(ALines[AFirst - 1]) = '')
    and (Trim(ALines[ALast + 1]) = '') then
    Inc(ALast);
  E := Default(TSafeDeleteEdit);
  E.FirstLine := AFirst;
  E.LastLine := ALast;
  ASym.Edits := ASym.Edits + [E];
end;

procedure Veto(var ASym: TSafeDeleteSymbol; const AWhy: string);
begin
  ASym.Vetoes := ASym.Vetoes + [AWhy];
end;

// Names of a data declaration line: "A, B: Integer;" / "X = 5;" /
// "var X: T;" / "class var F: T;". AColon receives the position of the
// ':' or '=' separating names and the rest, ANamesStart where the names
// begin.
function SplitDeclNames(const ALine: string; out ANamesStart, ASep: Integer;
  out AIsEquals: Boolean): TArray<string>;
var
  T, U: string;
  Lead, I, Depth: Integer;
begin
  Result := nil;
  ASep := 0;
  AIsEquals := False;
  T := StripLineComment(ALine);
  Lead := 1;
  while (Lead <= Length(T)) and CharInSet(T[Lead], [' ', #9]) do Inc(Lead);
  U := UpperCase(Copy(T, Lead, MaxInt));
  for var KW in ['CLASS VAR ', 'VAR ', 'CONST ', 'THREADVAR ', 'TYPE ', 'RESOURCESTRING '] do
    if U.StartsWith(KW) then
    begin
      Inc(Lead, Length(KW));
      while (Lead <= Length(T)) and CharInSet(T[Lead], [' ', #9]) do Inc(Lead);
      Break;
    end;
  ANamesStart := Lead;
  Depth := 0;
  for I := Lead to Length(T) do
  begin
    case T[I] of
      '(', '[', '<': Inc(Depth);
      ')', ']', '>': Dec(Depth);
      ':': if (Depth = 0) and ((I = Length(T)) or (T[I + 1] <> '=')) then begin ASep := I; Break; end;
      '=': if Depth = 0 then begin ASep := I; AIsEquals := True; Break; end;
    end;
  end;
  if ASep = 0 then Exit;
  for var S in Copy(T, Lead, ASep - Lead).Split([',']) do
  begin
    var N := Trim(S);
    var LT := Pos('<', N);
    if LT > 0 then N := Trim(Copy(N, 1, LT - 1));
    if (N = '') or not IsValidIdent(N) then Exit(nil);
    Result := Result + [N];
  end;
end;

function LineDeclaresName(const ALine, AName: string): Boolean;
var
  Kind, Hdr, Q, P, R: string;
  IsCM, IsEq: Boolean;
  NS, Sep, HdrEnd: Integer;
begin
  Result := False;
  if IsHeaderLine(StripLineComment(ALine), Kind, IsCM) then
  begin
    Hdr := CollectHeader([ALine], 0, HdrEnd);
    if Hdr = '' then Hdr := StripLineComment(ALine) + ';';
    if not ParseHeader(Hdr, Kind, Q, P, R) then Exit;
    var Dot := LastDelimiter('.', Q);
    Exit(SameText(Copy(Q, Dot + 1, MaxInt), AName));
  end;
  var T := Code(ALine);
  if T.ToUpper.StartsWith('CLASS PROPERTY ') then T := Trim(Copy(T, 7, MaxInt));
  if T.ToUpper.StartsWith('PROPERTY ') then
  begin
    var I := 10;
    while (I <= Length(T)) and CharInSet(T[I], [' ', #9]) do Inc(I);
    var S := I;
    while (I <= Length(T)) and IsIdentCh(T[I]) do Inc(I);
    Exit(SameText(Copy(T, S, I - S), AName));
  end;
  for var N in SplitDeclNames(ALine, NS, Sep, IsEq) do
    if SameText(N, AName) then Exit(True);
end;

function PlanSafeDeleteSymbol(const ALines: TArray<string>; ADeclLine0: Integer;
  const AName: string; out ASym: TSafeDeleteSymbol; out AWhy: string): Boolean;
var
  Content, Kind, Hdr, Qual, Params, Ret, T, U, Dirs, CName: string;
  IsCM, IsIntf: Boolean;
  Impl, HdrEnd, DeclLine, F, L: Integer;
begin
  Result := False;
  ASym := Default(TSafeDeleteSymbol);
  ASym.Name := AName;
  ASym.ImplLine := -1;
  ASym.ContainerLine := -1;
  ASym.LocalFirst := -1;
  ASym.LocalLast := -1;
  AWhy := '';
  if (ADeclLine0 < 0) or (ADeclLine0 > High(ALines)) then
  begin
    AWhy := 'the declaration line is outside the unit';
    Exit;
  end;
  Content := string.Join(sLineBreak, ALines);
  Impl := ImplementationLineOf(ALines);
  DeclLine := ADeclLine0;
  T := Code(ALines[DeclLine]);

  // ---- routines and methods ------------------------------------------------
  if IsHeaderLine(T, Kind, IsCM) then
  begin
    Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
    if (Hdr = '') or not ParseHeader(Hdr, Kind, Qual, Params, Ret) then
    begin
      AWhy := 'the routine header could not be read';
      Exit;
    end;
    var Dot := LastDelimiter('.', Qual);
    if not SameText(Copy(Qual, Dot + 1, MaxInt), AName) then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    if Dot > 0 then
    begin
      // an IMPLEMENTATION header "TFoo.Bar": the declaration is in the type
      var Owner := Copy(Qual, 1, Dot - 1);
      var OD := LastDelimiter('.', Owner);
      Owner := Copy(Owner, OD + 1, MaxInt);
      var LT := Pos('<', Owner);
      if LT > 0 then Owner := Copy(Owner, 1, LT - 1);
      DeclLine := FindMemberDeclarationLine(Content, Owner, AName);
      if DeclLine < 0 then
      begin
        AWhy := Format('the declaration of %s.%s was not found', [Owner, AName]);
        Exit;
      end;
      Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
      if (Hdr = '') or not IsHeaderLine(Code(ALines[DeclLine]), Kind, IsCM) then
      begin
        AWhy := 'the declaration of the method could not be read';
        Exit;
      end;
    end
    else if (DeclLine > Impl) and (ALines[DeclLine] = TrimLeft(ALines[DeclLine])) then
    begin
      // an unqualified top-level header in the implementation: either the
      // body of a routine declared in the interface, or impl-only
      for var I := 0 to Min(Impl, High(ALines)) - 1 do
        if IsHeaderLine(Code(ALines[I]), Kind, IsCM) and (ALines[I] = TrimLeft(ALines[I]))
          and LineDeclaresName(ALines[I], AName) then
        begin
          DeclLine := I;
          Hdr := CollectHeader(ALines, DeclLine, HdrEnd);
          Break;
        end;
    end;
    ASym.DeclLine := DeclLine;
    Dirs := HeaderDirectives(ALines, DeclLine, HdrEnd);
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    if ASym.ContainerLine >= 0 then
    begin
      ASym.Kind := sdkMethod;
      ASym.Container := CName;
      ASym.IsInterfaceMember := IsIntf;
      ASym.Parents := HeaderParents(Code(ALines[ASym.ContainerLine]));
      if not IsIntf then
        ASym.ImplLine := FindImplHeader(ALines, CName + '.' + AName);
    end
    else
    begin
      ASym.Kind := sdkRoutine;
      if DeclLine > Impl then
        ASym.ImplLine := DeclLine          // implementation-only (or nested)
      else
        ASym.ImplLine := FindImplHeader(ALines, AName);
      // a routine nested in another one is local to it
      if (ASym.ImplLine >= 0) and (ALines[ASym.ImplLine] <> TrimLeft(ALines[ASym.ImplLine])) then
      begin
        var OF_, OL_: Integer;
        if FindEnclosingRoutineRange(Content, ASym.ImplLine - 1, OF_, OL_) then
        begin
          ASym.LocalFirst := OF_;
          ASym.LocalLast := OL_;
        end;
      end;
    end;
    // what may reach a method WITHOUT a textual call
    for var W in ['overload', 'virtual', 'dynamic', 'abstract', 'override', 'message'] do
      if HasWholeWordCI(Dirs, W) then
        Veto(ASym, Format('the %s is declared "%s" - %s', [SafeDeleteKindText(ASym.Kind), W,
          IfThen(W = 'overload', 'another overload could silently take over its calls',
          'it can be reached through dispatch without a textual call')]));
    if HasWholeWordCI(Dirs, 'external') then ASym.ImplLine := -1;
    // edits: the declaration (when separate) and the implementation
    if ASym.ImplLine <> DeclLine then
      AddEdit(ASym, DeclLine, HdrEnd, ALines);
    if ASym.ImplLine >= 0 then
    begin
      if not FindEnclosingRoutineRange(Content, ASym.ImplLine, F, L) or (F <> ASym.ImplLine) then
      begin
        AWhy := 'the implementation block could not be delimited';
        Exit;
      end;
      AddEdit(ASym, F, L, ALines);
    end;
    // (declared but never implemented: deleting the declaration is enough)
  end
  // ---- properties ------------------------------------------------------------
  else if T.ToUpper.StartsWith('PROPERTY ') or T.ToUpper.StartsWith('CLASS PROPERTY ') then
  begin
    if not LineDeclaresName(ALines[DeclLine], AName) then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    ASym.Kind := sdkProperty;
    ASym.DeclLine := DeclLine;
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    ASym.Container := CName;
    ASym.IsInterfaceMember := IsIntf;
    if not T.EndsWith(';') then
      Veto(ASym, 'the property declaration spans several lines - not supported');
    AddEdit(ASym, DeclLine, DeclLine, ALines);
  end
  // ---- data: fields, variables, constants, single-line types -------------------
  else
  begin
    var NS, Sep: Integer;
    var IsEq: Boolean;
    var Names := SplitDeclNames(ALines[DeclLine], NS, Sep, IsEq);
    var Idx := -1;
    for var I := 0 to High(Names) do
      if SameText(Names[I], AName) then Idx := I;
    if Idx < 0 then
    begin
      AWhy := Format('line %d does not declare "%s"', [DeclLine + 1, AName]);
      Exit;
    end;
    ASym.DeclLine := DeclLine;
    ASym.ContainerLine := EnclosingTypeHeader(ALines, DeclLine, CName, IsIntf);
    ASym.Container := CName;
    U := UpperCase(Trim(Copy(StripLineComment(ALines[DeclLine]), Sep + 1, MaxInt)));
    if ASym.ContainerLine >= 0 then
      ASym.Kind := sdkField
    else
    begin
      var Sec := SectionAt(ALines, DeclLine);
      if (Sec = 'type') or (IsEq and (Sec = '') and (U.StartsWith('CLASS') or U.StartsWith('RECORD'))) then
        ASym.Kind := sdkType
      else if (Sec = 'const') or (Sec = 'resourcestring') or IsEq then
        ASym.Kind := sdkConstant
      else
        ASym.Kind := sdkVariable;
      var RF, RL: Integer;
      if FindEnclosingRoutineRange(Content, DeclLine, RF, RL) then
      begin
        ASym.LocalFirst := RF;
        ASym.LocalLast := RL;
      end;
    end;
    if (ASym.Kind = sdkType) and ((ClassOpenerName(Code(ALines[DeclLine])) <> '')
      or U.StartsWith('CLASS') or U.StartsWith('RECORD') or U.StartsWith('INTERFACE')
      or U.StartsWith('PACKED')) and not Code(ALines[DeclLine]).EndsWith(';') then
    begin
      Veto(ASym, 'types with a body are not supported yet');
    end
    else if not Code(ALines[DeclLine]).EndsWith(';') then
      Veto(ASym, 'the declaration spans several lines - not supported');
    if Length(Names) > 1 then
    begin
      // drop the name from the list, keep everything else as written
      var S := ALines[DeclLine];
      var Kept: TArray<string> := nil;
      for var I := 0 to High(Names) do
        if I <> Idx then Kept := Kept + [Names[I]];
      var E := Default(TSafeDeleteEdit);
      E.FirstLine := DeclLine;
      E.LastLine := DeclLine;
      E.HasReplacement := True;
      E.Replacement := Copy(S, 1, NS - 1) + string.Join(', ', Kept) +
        IfThen(IsEq, ' ', '') + Copy(S, Sep, MaxInt);
      ASym.Edits := ASym.Edits + [E];
    end
    else
    begin
      // the whole line - and a section keyword left without declarations
      F := DeclLine;
      L := DeclLine;
      var W := FirstWordU(Code(ALines[DeclLine]));
      var KeywordOnLine := (W = 'VAR') or (W = 'CONST') or (W = 'TYPE') or (W = 'THREADVAR');
      if not KeywordOnLine then
      begin
        var Prev := DeclLine - 1;
        while (Prev >= 0) and (Code(ALines[Prev]) = '') do Dec(Prev);
        var Next := DeclLine + 1;
        while (Next <= High(ALines)) and (Code(ALines[Next]) = '') do Inc(Next);
        var PW := UpperCase(Code(IfThen(Prev >= 0, ALines[Prev], '')));
        var NW := '';
        if Next <= High(ALines) then NW := FirstWordU(Code(ALines[Next]));
        if ((PW = 'VAR') or (PW = 'CONST') or (PW = 'TYPE') or (PW = 'THREADVAR') or
            (PW = 'CLASS VAR') or (PW = 'RESOURCESTRING'))
          and ((NW = 'BEGIN') or (NW = 'VAR') or (NW = 'CONST') or (NW = 'TYPE') or
            (NW = 'THREADVAR') or (NW = 'RESOURCESTRING') or (NW = 'LABEL') or
            (NW = 'PROCEDURE') or (NW = 'FUNCTION') or (NW = 'CONSTRUCTOR') or
            (NW = 'DESTRUCTOR') or (NW = 'CLASS') or (NW = 'IMPLEMENTATION') or
            (NW = 'INITIALIZATION') or (NW = 'FINALIZATION') or (NW = 'END') or
            (NW = 'PRIVATE') or (NW = 'PROTECTED') or (NW = 'PUBLIC') or
            (NW = 'PUBLISHED') or (NW = 'STRICT') or (NW = 'PROPERTY') or (NW = '')) then
          F := Prev;
      end;
      AddEdit(ASym, F, L, ALines);
    end;
  end;

  // ---- container checks -------------------------------------------------------
  if (ASym.ContainerLine >= 0) and not ASym.IsInterfaceMember then
  begin
    ASym.Parents := HeaderParents(Code(ALines[ASym.ContainerLine]));
    if VisibilityAt(ALines, ASym.ContainerLine, ASym.DeclLine) = 'published' then
      Veto(ASym, 'the member is PUBLISHED - it can be used by name (form streaming, RTTI)');
  end;
  Result := True;
end;

function ApplySafeDeleteEdits(const ALines: TArray<string>;
  const AEdits: TArray<TSafeDeleteEdit>): TArray<string>;
var
  Sorted: TArray<TSafeDeleteEdit>;
  Res: TArray<string>;
begin
  Sorted := Copy(AEdits);
  // bottom-up
  for var I := 0 to High(Sorted) - 1 do
    for var J := I + 1 to High(Sorted) do
      if Sorted[J].FirstLine > Sorted[I].FirstLine then
      begin
        var Tmp := Sorted[I];
        Sorted[I] := Sorted[J];
        Sorted[J] := Tmp;
      end;
  Res := Copy(ALines);
  for var E in Sorted do
  begin
    if (E.FirstLine < 0) or (E.LastLine > High(Res)) or (E.LastLine < E.FirstLine) then Continue;
    var Tail := Copy(Res, E.LastLine + 1, MaxInt);
    Res := Copy(Res, 0, E.FirstLine);
    if E.HasReplacement then Res := Res + [E.Replacement];
    Res := Res + Tail;
  end;
  Result := Res;
end;

function FormTextMentions(const AText, AName: string): TArray<Integer>;
var
  Lines: TArray<string>;
  U, W: string;
  P, E: Integer;
begin
  Result := nil;
  if AName = '' then Exit;
  W := UpperCase(AName);
  Lines := AText.Replace(#13#10, #10).Split([#10]);
  for var L := 0 to High(Lines) do
  begin
    U := UpperCase(Lines[L]);
    P := Pos(W, U);
    while P > 0 do
    begin
      E := P + Length(W);
      if ((P = 1) or not IsIdentCh(U[P - 1])) and ((E > Length(U)) or not IsIdentCh(U[E])) then
      begin
        Result := Result + [L];
        Break;
      end;
      P := Pos(W, U, P + 1);
    end;
  end;
end;

end.
