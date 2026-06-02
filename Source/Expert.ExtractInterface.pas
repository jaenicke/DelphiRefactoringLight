(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ExtractInterface;

{
  "Extract interface from class" engine.

  Walks a Pascal class declaration, collects all members (fields,
  methods, properties) with their visibility + signatures, and lets the
  caller compose a target interface from a freely chosen subset of them.

  When fields are picked, the engine generates a paired property +
  getter/setter triple in the original class so the field-based member
  is reachable through an interface (which can only carry method/
  property declarations).

  Two operating modes:
    eimExtractNew     - create a new unit IFoo.pas with a fresh
                        interface declaration; rewrite the class to
                        implement it.
    eimAddToExisting  - splice the chosen members into an existing
                        interface declared somewhere in the project;
                        rewrite the class to implement it (if not
                        already in the ancestor list).

  Limitations of v1:
    - Generic classes / generic interfaces are not specially handled.
    - Indexed properties / default array properties are emitted
      verbatim into the interface; no special validation.
    - Conditional defines inside the class body are preserved
      textually but the parser does not track which branch is active.
    - Overloaded methods are all emitted - no dedup.
}

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Generics.Collections,
  Delphi.FileEncoding;

type
  TInterfaceMode = (eimExtractNew, eimAddToExisting);

  TMemberKind = (mkField, mkProperty, mkMethod);

  TMemberVisibility = (mvUnknown, mvStrictPrivate, mvPrivate, mvStrictProtected,
    mvProtected, mvPublic, mvPublished);

  /// <summary>One declaration inside the class body. For fields, Signature
  ///  is e.g. "FFoo: TBar"; for methods, the full procedure/function
  ///  header without the trailing semicolon; for properties the full
  ///  "property X: T read ... write ..." text.</summary>
  TClassMember = record
    Name: string;
    Kind: TMemberKind;
    Visibility: TMemberVisibility;
    IsClassMember: Boolean;
    Signature: string;
    TypeName: string;
    /// <summary>1-based line numbers in the source file.</summary>
    LineStart, LineEnd: Integer;
    /// <summary>True iff this row is checked in the dialog.</summary>
    Selected: Boolean;
    /// <summary>For fields: the synthetic property name we plan to
    ///  expose (FFoo -> Foo). For properties: copy of Name.</summary>
    ExposedName: string;
    /// <summary>For methods we just take the signature verbatim. For
    ///  fields / property-without-accessors, the engine fills these
    ///  during code generation.</summary>
    GetterName, SetterName: string;
    /// <summary>True iff the original property had no explicit Get/Set
    ///  (i.e. it read/wrote a field directly) and we need to invent
    ///  accessors for the interface contract.</summary>
    NeedsSynthAccessors: Boolean;
    /// <summary>For properties: read-only flag (only "read" clause).
    ///  Affects whether we emit a setter.</summary>
    IsReadOnly: Boolean;
  end;

  TExtractInterfaceInfo = record
    Mode: TInterfaceMode;
    SourceFile: string;
    ClassName: string;
    ClassDeclLine: Integer;          // line of "TFoo = class[...]"
    ClassEndLine: Integer;           // line of the matching 'end;'
    AncestorList: string;            // raw "(TBase, IOther)" or ''
    Members: TArray<TClassMember>;
    InterfaceName: string;
    Guid: string;                    // "{XXXXXXXX-...}"
    /// <summary>For eimExtractNew: absolute path of the new unit file.
    ///  Defaults to the source file's directory with name <Interface>.pas
    ///  (T-prefix stripped for the class name, I-prefix prepended).
    ///  For eimAddToExisting: ignored.</summary>
    TargetFile: string;
    /// <summary>For eimAddToExisting: source file containing the
    ///  existing interface, plus 1-based start/end lines of its
    ///  declaration (from "IXxx = interface" up to the matching "end;").
    ///  For eimExtractNew: ignored.</summary>
    ExistingFile: string;
    ExistingDeclLine, ExistingEndLine: Integer;
    /// <summary>Unit names to include in the new unit's interface-uses
    ///  clause. The wizard typically fills this with the source unit's
    ///  own interface-uses so any type referenced by the synthesised
    ///  interface members (TButton, TListView, ...) resolves out of the
    ///  box. eimAddToExisting ignores this.</summary>
    NewUnitUses: TArray<string>;
  end;

  TExtractInterfaceEngine = class
  public
    /// <summary>Parses the class whose declaration line is closest to,
    ///  and at-or-before, ALine in AFileLines. Returns False if no class
    ///  was found around the position.</summary>
    class function ParseClassAtLine(const AFileLines: TArray<string>;
      const AFile: string; ALine: Integer;
      out AInfo: TExtractInterfaceInfo): Boolean;

    /// <summary>Builds the text of the new interface from AInfo. Format
    ///  is fixed: 2-space indent under "type", GUID line, then each
    ///  selected member on its own line. Field members are emitted as a
    ///  property triple (Getter, Setter when not read-only, property).
    ///  Property members with NeedsSynthAccessors get the same triple;
    ///  otherwise the property is emitted verbatim.</summary>
    class function BuildInterfaceText(const AInfo: TExtractInterfaceInfo): string;

    /// <summary>Builds the full text of the new unit file (header,
    ///  uses-clause guess, type section with interface, "end."). Only
    ///  used for eimExtractNew.</summary>
    class function BuildNewUnitText(const AInfo: TExtractInterfaceInfo): string;

    /// <summary>Returns a fresh Pascal GUID literal of the form
    ///  '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'.</summary>
    class function NewGuidLiteral: string;

    /// <summary>Suggests an interface name from a class name: TFoo ->
    ///  IFoo, FooBar -> IFooBar, TForm1 -> IForm1.</summary>
    class function SuggestInterfaceName(const AClassName: string): string;

    /// <summary>Suggests a unit / file name for the new unit. Strips a
    ///  leading I (the interface naming convention) only for the .pas
    ///  base name? No - we keep the interface name as the unit base
    ///  name: IFoo -> IFoo.pas in the same dir as ASourceFile.</summary>
    class function SuggestTargetFile(const ASourceFile, AInterfaceName: string): string;

    /// <summary>Parses an interface declaration (from ADeclLine to
    ///  AEndLine in AFile) and returns the names of every method,
    ///  property, and accessor (the identifiers appearing after
    ///  'read' / 'write') it declares. Used for clash detection when
    ///  adding members to an existing interface.</summary>
    class function ParseExistingInterfaceNames(const AFile: string;
      ADeclLine, AEndLine: Integer): TArray<string>;

    /// <summary>Returns the list of identifier names a single class
    ///  member would emit into the interface (method name; or property
    ///  name + Get/Set names; or for a field the synthetic
    ///  Get/Set/property names). Used together with
    ///  ParseExistingInterfaceNames to detect clashes.</summary>
    class function WouldEmitNames(const M: TClassMember): TArray<string>;

    /// <summary>True iff at least one of the identifiers M would emit
    ///  is already in AExistingUpper (which must contain UPPERCASE
    ///  names).</summary>
    class function ClashesWithExisting(const M: TClassMember;
      const AExistingUpper: TArray<string>): Boolean;
  end;

  /// <summary>Tiny utility: scans an entire Pascal source file and yields
  ///  every interface declaration ("IFoo = interface ['{guid}']") it
  ///  finds, together with the file path and the line numbers. Used by
  ///  the "Add to existing interface" picker.</summary>
  TInterfaceDeclLocation = record
    FileName: string;
    InterfaceName: string;
    DeclLine: Integer;          // line of "IFoo = interface[(...)]"
    EndLine: Integer;           // line of matching 'end;'
  end;

  TProjectInterfaceScanner = class
  public
    class function ScanFile(const AFileName: string): TArray<TInterfaceDeclLocation>;
    class function ScanProject(const AFiles: TArray<string>): TArray<TInterfaceDeclLocation>;
  end;

implementation

uses
  System.StrUtils, System.Character, System.Math, Winapi.ActiveX, Winapi.Windows;

{ ---------- helpers ---------- }

function StripLineComment(const ALine: string): string;
var
  P: Integer;
begin
  Result := ALine;
  P := Pos('//', Result);
  if P > 0 then Result := Copy(Result, 1, P - 1);
end;

function IsIdent(C: Char): Boolean; inline;
begin
  Result := C.IsLetterOrDigit or (C = '_');
end;

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := C.IsLetter or (C = '_');
end;

function FirstWordUpper(const ALine: string; out APosAfter: Integer): string;
var
  I, J: Integer;
begin
  Result := '';
  APosAfter := 0;
  I := 1;
  while (I <= Length(ALine)) and ALine[I].IsWhiteSpace do Inc(I);
  if I > Length(ALine) then Exit;
  if not IsIdentStart(ALine[I]) then Exit;
  J := I;
  while (J <= Length(ALine)) and IsIdent(ALine[J]) do Inc(J);
  Result := UpperCase(Copy(ALine, I, J - I));
  APosAfter := J;
end;

function ExtractIdentAfter(const ALine: string; AFromPos: Integer;
  out AEndPos: Integer): string;
var
  I: Integer;
begin
  Result := '';
  AEndPos := AFromPos;
  I := AFromPos;
  while (I <= Length(ALine)) and ALine[I].IsWhiteSpace do Inc(I);
  if (I > Length(ALine)) or not IsIdentStart(ALine[I]) then Exit;
  AEndPos := I;
  while (AEndPos <= Length(ALine)) and IsIdent(ALine[AEndPos]) do Inc(AEndPos);
  Result := Copy(ALine, I, AEndPos - I);
end;

{ ---------- TExtractInterfaceEngine ---------- }

class function TExtractInterfaceEngine.NewGuidLiteral: string;
var
  G: TGUID;
begin
  CoCreateGuid(G);
  Result := GUIDToString(G);
end;

class function TExtractInterfaceEngine.SuggestInterfaceName(const AClassName: string): string;
var
  Bare: string;
begin
  Bare := AClassName;
  if (Length(Bare) > 1) and (Bare[1].ToUpper = 'T') then
    Bare := Copy(Bare, 2, MaxInt);
  Result := 'I' + Bare;
end;

class function TExtractInterfaceEngine.SuggestTargetFile(
  const ASourceFile, AInterfaceName: string): string;
var
  Bare: string;
begin
  // Convention: "IBlub" -> "Interfaces.Blub.pas" in the source file's
  // directory. Strip the leading 'I' if present; otherwise use the name
  // verbatim as the dotted suffix.
  Bare := AInterfaceName;
  if (Length(Bare) > 1) and (Bare[1].ToUpper = 'I') and IsIdentStart(Bare[2]) then
    Bare := Copy(Bare, 2, MaxInt);
  Result := IncludeTrailingPathDelimiter(ExtractFilePath(ASourceFile)) +
    'Interfaces.' + Bare + '.pas';
end;

class function TExtractInterfaceEngine.ParseClassAtLine(
  const AFileLines: TArray<string>; const AFile: string; ALine: Integer;
  out AInfo: TExtractInterfaceInfo): Boolean;
var
  I, P, J: Integer;
  Line, Trimmed, Upper, FirstWord: string;
  DeclLine, EndLine: Integer;
  Depth: Integer;
  CurClassName, AncestorRaw: string;
  Members: TList<TClassMember>;
  CurVis: TMemberVisibility;
  CurIsClass: Boolean;
  procedure ResetModifiers; begin CurIsClass := False; end;

  function MatchClassHeader(const ALine: string;
    out AClassName, AAncestors: string): Boolean;
  var
    EqPos, ClassPos, ParenPos: Integer;
    L, U, Name, AfterClass: string;
  begin
    Result := False;
    L := StripLineComment(ALine);
    U := UpperCase(L);
    EqPos := Pos('=', L);
    if EqPos = 0 then Exit;
    ClassPos := Pos('CLASS', U);
    if (ClassPos = 0) or (ClassPos < EqPos) then Exit;
    // exclude 'class procedure', 'class function', 'class var', etc.
    AfterClass := Trim(Copy(L, ClassPos + 5, MaxInt));
    if (AfterClass <> '') then
    begin
      var AfterUp := UpperCase(AfterClass);
      if StartsText('PROCEDURE', AfterUp) or StartsText('FUNCTION', AfterUp) or
         StartsText('VAR', AfterUp) or StartsText('CONST', AfterUp) or
         StartsText('THREADVAR', AfterUp) or StartsText('OPERATOR', AfterUp) or
         StartsText('PROPERTY', AfterUp) or StartsText('OF ', AfterUp) or
         (AfterUp = 'OF') then Exit;
      // class reference type: "= class of TFoo"
      if StartsText('OF', AfterUp) and ((Length(AfterUp) = 2) or (AfterUp[3] = ' ')) then
        Exit;
    end;
    // forward declaration "TFoo = class;" - skip
    if Pos(';', L) > 0 then
    begin
      // Only if the ';' comes right after CLASS (no body opens this line).
      var TailUp := UpperCase(Trim(Copy(L, ClassPos + 5, MaxInt)));
      if (TailUp = ';') or StartsText(';', TailUp) then Exit;
    end;
    Name := Trim(Copy(L, 1, EqPos - 1));
    // drop generic <T> suffix if present (not v1-specially-handled but
    // accepted as a name).
    var GenPos := Pos('<', Name);
    if GenPos > 0 then Name := Trim(Copy(Name, 1, GenPos - 1));
    AClassName := Name;
    AAncestors := '';
    ParenPos := Pos('(', Copy(L, ClassPos, MaxInt));
    if ParenPos > 0 then
    begin
      var Abs_ := ClassPos + ParenPos - 1;
      var Close_ := Pos(')', Copy(L, Abs_, MaxInt));
      if Close_ > 0 then
        AAncestors := Trim(Copy(L, Abs_, Close_));
    end;
    Result := True;
  end;

begin
  Result := False;
  AInfo := Default(TExtractInterfaceInfo);
  if Length(AFileLines) = 0 then Exit;

  // Walk upward from ALine to find a class header. If we don't find
  // one upward within 2000 lines, walk downward as fallback (cursor may
  // be before the class).
  DeclLine := -1;
  for I := Min(ALine, Length(AFileLines)) - 1 downto Max(0, ALine - 2000) do
  begin
    if MatchClassHeader(AFileLines[I], CurClassName, AncestorRaw) then
    begin
      DeclLine := I + 1;
      Break;
    end;
  end;
  if DeclLine < 1 then
  begin
    for I := Max(0, ALine - 1) to Min(Length(AFileLines) - 1, ALine + 2000) do
    begin
      if MatchClassHeader(AFileLines[I], CurClassName, AncestorRaw) then
      begin
        DeclLine := I + 1;
        Break;
      end;
    end;
  end;
  if DeclLine < 1 then Exit;

  // Find end of class declaration. Track 'record'/'class' depth so a
  // nested record/class doesn't close us prematurely.
  EndLine := -1;
  Depth := 1;
  for I := DeclLine to Length(AFileLines) - 1 do
  begin
    Line := StripLineComment(AFileLines[I]);
    Upper := UpperCase(Line);
    // Tokenize words and react to 'record'/'class' (opens) and 'end'
    // (closes). String-content is ignored for simplicity (v1).
    var Pos_ := 1;
    while Pos_ <= Length(Upper) do
    begin
      if IsIdentStart(Upper[Pos_]) then
      begin
        var Q := Pos_;
        while (Q <= Length(Upper)) and IsIdent(Upper[Q]) do Inc(Q);
        var W := Copy(Upper, Pos_, Q - Pos_);
        if (W = 'RECORD') then
          Inc(Depth)
        else if (W = 'END') then
        begin
          Dec(Depth);
          if Depth = 0 then begin EndLine := I + 1; Break; end;
        end;
        Pos_ := Q;
      end
      else
        Inc(Pos_);
    end;
    if EndLine > 0 then Break;
  end;
  if EndLine < 1 then Exit;

  AInfo.SourceFile := AFile;
  AInfo.ClassName := CurClassName;
  AInfo.ClassDeclLine := DeclLine;
  AInfo.ClassEndLine := EndLine;
  AInfo.AncestorList := AncestorRaw;

  // Parse members between DeclLine+1 and EndLine-1.
  Members := TList<TClassMember>.Create;
  // Members declared before any visibility keyword default to published
  // for $M+ classes (TPersistent / TComponent / TForm descendants - the
  // common case in VCL/FMX projects) and to public for plain TObject
  // descendants. We default to mvPublished here because the published
  // case is by far the dominant one in practice; in the rare pure-
  // TObject case the user can re-tick manually.
  CurVis := mvPublished;
  ResetModifiers;
  try
    I := DeclLine;
    while I < EndLine - 1 do
    begin
      Line := StripLineComment(AFileLines[I]);
      Trimmed := Trim(Line);
      if Trimmed = '' then begin Inc(I); Continue; end;

      FirstWord := FirstWordUpper(Line, P);

      // Visibility section?
      if (FirstWord = 'PRIVATE') then begin CurVis := mvPrivate; ResetModifiers; Inc(I); Continue end
      else if (FirstWord = 'PROTECTED') then begin CurVis := mvProtected; ResetModifiers; Inc(I); Continue end
      else if (FirstWord = 'PUBLIC') then begin CurVis := mvPublic; ResetModifiers; Inc(I); Continue end
      else if (FirstWord = 'PUBLISHED') then begin CurVis := mvPublished; ResetModifiers; Inc(I); Continue end
      else if (FirstWord = 'STRICT') then
      begin
        var Rest := UpperCase(Trim(Copy(Line, P, MaxInt)));
        if StartsText('PRIVATE', Rest) then CurVis := mvStrictPrivate
        else if StartsText('PROTECTED', Rest) then CurVis := mvStrictProtected;
        ResetModifiers; Inc(I); Continue;
      end;

      // CLASS modifier (class procedure / class var / class property)?
      if (FirstWord = 'CLASS') then
      begin
        var Rest := UpperCase(Trim(Copy(Line, P, MaxInt)));
        if StartsText('VAR', Rest) or StartsText('CONST', Rest) or
           StartsText('THREADVAR', Rest) then
        begin
          // "class var" introduces a new sub-section; treat like
          // visibility-stays-the-same but mark class-member. Just continue.
          CurIsClass := True; Inc(I); Continue;
        end;
        // class procedure / class function / class property: mark and
        // re-parse the rest of the line as the actual member.
        CurIsClass := True;
        Line := Trim(Copy(Line, P, MaxInt));
        Trimmed := Line;
        FirstWord := FirstWordUpper(Line, P);
      end;

      // Method?
      if (FirstWord = 'PROCEDURE') or (FirstWord = 'FUNCTION') or
         (FirstWord = 'CONSTRUCTOR') or (FirstWord = 'DESTRUCTOR') or
         (FirstWord = 'OPERATOR') then
      begin
        // Collect the full signature: may span multiple lines until ';'
        var SigStart := I;
        var Acc := Line;
        while Pos(';', StripLineComment(Acc)) = 0 do
        begin
          Inc(I);
          if I >= EndLine - 1 then Break;
          Acc := Acc + ' ' + Trim(StripLineComment(AFileLines[I]));
        end;
        // Stop accumulating after first ';' that closes the header.
        var SigEnd := I + 1;
        // After ';', directives (overload; virtual; abstract; ...) may
        // follow on the same or next lines. Keep grabbing until we
        // hit a line whose trimmed first word isn't a known directive
        // and isn't blank.
        Inc(I);
        while I < EndLine - 1 do
        begin
          var DLine := Trim(StripLineComment(AFileLines[I]));
          if DLine = '' then begin Inc(I); Continue; end;
          var DUpper := UpperCase(DLine);
          // strip trailing ';'
          while (DUpper <> '') and (DUpper[Length(DUpper)] = ';') do
            DUpper := Copy(DUpper, 1, Length(DUpper) - 1);
          if (DUpper = 'OVERLOAD') or (DUpper = 'OVERRIDE') or
             (DUpper = 'VIRTUAL') or (DUpper = 'ABSTRACT') or
             (DUpper = 'REINTRODUCE') or (DUpper = 'STATIC') or
             (DUpper = 'INLINE') or (DUpper = 'CDECL') or
             (DUpper = 'STDCALL') or (DUpper = 'SAFECALL') or
             (DUpper = 'PASCAL') or (DUpper = 'REGISTER') or
             (DUpper = 'FINAL') or (DUpper = 'MESSAGE') or
             StartsText('MESSAGE ', DUpper) or
             StartsText('DEPRECATED', DUpper) or
             (DUpper = 'EXPERIMENTAL') or (DUpper = 'PLATFORM') or
             (DUpper = 'LIBRARY') then
          begin
            Acc := Acc + ' ' + DLine;
            SigEnd := I + 1;
            Inc(I); Continue;
          end;
          Break;
        end;

        // Extract method name = first ident after PROCEDURE/FUNCTION/...
        var After := Length(FirstWord) + 1; // FirstWord found at start of (trimmed) Acc
        var TrimmedAcc := Trim(Acc);
        var AfterKw := Pos(' ', TrimmedAcc);
        var NameStart := AfterKw + 1;
        var NameEnd := NameStart;
        while (NameEnd <= Length(TrimmedAcc)) and IsIdent(TrimmedAcc[NameEnd]) do Inc(NameEnd);
        var MName := Copy(TrimmedAcc, NameStart, NameEnd - NameStart);
        if After = 0 then ; // silence hint

        var M: TClassMember;
        M := Default(TClassMember);
        M.Name := MName;
        M.Kind := mkMethod;
        M.Visibility := CurVis;
        M.IsClassMember := CurIsClass;
        M.Signature := Trim(Acc);
        // Strip a trailing ';' from signature for emission control.
        while (M.Signature <> '') and (M.Signature[Length(M.Signature)] = ';') do
          M.Signature := Trim(Copy(M.Signature, 1, Length(M.Signature) - 1));
        M.LineStart := SigStart + 1;
        M.LineEnd := SigEnd;
        M.ExposedName := MName;
        Members.Add(M);
        ResetModifiers;
        Continue;
      end;

      // Property?
      if FirstWord = 'PROPERTY' then
      begin
        var SigStart := I;
        var Acc := Line;
        while Pos(';', StripLineComment(Acc)) = 0 do
        begin
          Inc(I);
          if I >= EndLine - 1 then Break;
          Acc := Acc + ' ' + Trim(StripLineComment(AFileLines[I]));
        end;
        var SigEnd := I + 1;
        Inc(I);
        // Detect default-array / index suffixes? For v1 we just keep them.
        var TrimAcc := Trim(Acc);
        // name after "property "
        var Pn := ExtractIdentAfter(TrimAcc, Length('property') + 1, J);
        // type after first ':'
        var ColonP := Pos(':', TrimAcc);
        var TName := '';
        if ColonP > 0 then
        begin
          var T := Trim(Copy(TrimAcc, ColonP + 1, MaxInt));
          // up to next space or ';' or 'read' keyword
          var K := 1;
          while (K <= Length(T)) and not T[K].IsWhiteSpace and (T[K] <> ';') do Inc(K);
          TName := Copy(T, 1, K - 1);
        end;
        var UpAcc := UpperCase(TrimAcc);
        var HasWrite := (Pos(' WRITE ', ' ' + UpAcc + ' ') > 0);
        var HasRead  := (Pos(' READ ', ' ' + UpAcc + ' ') > 0);

        // Detect "naked" accessors that name a field (no method call):
        // word after READ / WRITE that does NOT start with Get/Set we
        // treat as field-style (conservative - rather emit a redundant
        // accessor than miss one).
        var AccReadIsField := False;
        var AccWriteIsField := False;
        for var KW in ['READ', 'WRITE'] do
        begin
          var KU: string := ' ' + UpAcc + ' ';
          var KP: Integer := Pos(' ' + KW + ' ', KU);
          if KP = 0 then Continue;
          // Step past the keyword in UpAcc itself (KP is 1-based into KU).
          var Idx: Integer := KP + Length(KW);
          while (Idx <= Length(KU)) and (KU[Idx] = ' ') do Inc(Idx);
          var Start: Integer := Idx;
          while (Idx <= Length(KU)) and (IsIdent(KU[Idx]) or (KU[Idx] = '.')) do Inc(Idx);
          var W: string := Copy(KU, Start, Idx - Start);
          var Field: Boolean := (W <> '') and
            not StartsText('GET', W) and not StartsText('SET', W);
          if KW = 'READ' then AccReadIsField := Field
          else AccWriteIsField := Field;
        end;

        var M: TClassMember;
        M := Default(TClassMember);
        M.Name := Pn;
        M.Kind := mkProperty;
        M.Visibility := CurVis;
        M.IsClassMember := CurIsClass;
        M.Signature := Trim(Acc);
        while (M.Signature <> '') and (M.Signature[Length(M.Signature)] = ';') do
          M.Signature := Trim(Copy(M.Signature, 1, Length(M.Signature) - 1));
        M.TypeName := TName;
        M.LineStart := SigStart + 1;
        M.LineEnd := SigEnd;
        M.ExposedName := Pn;
        M.IsReadOnly := HasRead and not HasWrite;
        M.NeedsSynthAccessors :=
          (HasRead and AccReadIsField) or (HasWrite and AccWriteIsField);
        Members.Add(M);
        ResetModifiers;
        Continue;
      end;

      // Otherwise: assume a field declaration "Name[, Name2]: Type;".
      var Sem := Pos(';', Trimmed);
      var FieldDecl := Trimmed;
      if Sem > 0 then FieldDecl := Copy(Trimmed, 1, Sem - 1);
      var CP := Pos(':', FieldDecl);
      if CP > 0 then
      begin
        var NamesPart := Trim(Copy(FieldDecl, 1, CP - 1));
        var TypePart := Trim(Copy(FieldDecl, CP + 1, MaxInt));
        for var NItem in NamesPart.Split([',']) do
        begin
          var FN := Trim(NItem);
          if FN = '' then Continue;
          if not IsIdentStart(FN[1]) then Continue;
          var M: TClassMember;
          M := Default(TClassMember);
          M.Name := FN;
          M.Kind := mkField;
          M.Visibility := CurVis;
          M.IsClassMember := CurIsClass;
          M.Signature := FN + ': ' + TypePart;
          M.TypeName := TypePart;
          M.LineStart := I + 1;
          M.LineEnd := I + 1;
          // ExposedName: F-prefix stripped if present.
          if (Length(FN) >= 2) and (FN[1].ToUpper = 'F') and IsIdentStart(FN[2]) then
            M.ExposedName := Copy(FN, 2, MaxInt)
          else
            M.ExposedName := FN;
          Members.Add(M);
        end;
      end;
      ResetModifiers;
      Inc(I);
    end;

    AInfo.Members := Members.ToArray;
  finally
    Members.Free;
  end;

  AInfo.InterfaceName := SuggestInterfaceName(AInfo.ClassName);
  AInfo.Guid := NewGuidLiteral;
  AInfo.TargetFile := SuggestTargetFile(AFile, AInfo.InterfaceName);
  Result := True;
end;

class function TExtractInterfaceEngine.BuildInterfaceText(
  const AInfo: TExtractInterfaceInfo): string;
var
  SB: TStringBuilder;
  M: TClassMember;
  Pn, Getter, Setter, TN, Indent: string;
begin
  Indent := '    ';
  SB := TStringBuilder.Create;
  try
    SB.Append('  ').Append(AInfo.InterfaceName).Append(' = interface').AppendLine;
    SB.Append(Indent).Append('[''').Append(AInfo.Guid).Append(''']').AppendLine;
    for M in AInfo.Members do
    begin
      if not M.Selected then Continue;
      case M.Kind of
        mkMethod:
          SB.Append(Indent).Append(M.Signature).Append(';').AppendLine;

        mkField:
        begin
          Pn := M.ExposedName;
          Getter := 'Get' + Pn;
          Setter := 'Set' + Pn;
          TN := M.TypeName;
          SB.Append(Indent).Append('function ').Append(Getter)
            .Append(': ').Append(TN).Append(';').AppendLine;
          SB.Append(Indent).Append('procedure ').Append(Setter)
            .Append('(const AValue: ').Append(TN).Append(');').AppendLine;
          SB.Append(Indent).Append('property ').Append(Pn)
            .Append(': ').Append(TN)
            .Append(' read ').Append(Getter)
            .Append(' write ').Append(Setter).Append(';').AppendLine;
        end;

        mkProperty:
        begin
          if not M.NeedsSynthAccessors then
            SB.Append(Indent).Append(M.Signature).Append(';').AppendLine
          else
          begin
            Pn := M.ExposedName;
            Getter := 'Get' + Pn;
            TN := M.TypeName;
            SB.Append(Indent).Append('function ').Append(Getter)
              .Append(': ').Append(TN).Append(';').AppendLine;
            if not M.IsReadOnly then
            begin
              Setter := 'Set' + Pn;
              SB.Append(Indent).Append('procedure ').Append(Setter)
                .Append('(const AValue: ').Append(TN).Append(');').AppendLine;
              SB.Append(Indent).Append('property ').Append(Pn)
                .Append(': ').Append(TN)
                .Append(' read ').Append(Getter)
                .Append(' write ').Append(Setter).Append(';').AppendLine;
            end
            else
              SB.Append(Indent).Append('property ').Append(Pn)
                .Append(': ').Append(TN)
                .Append(' read ').Append('Get' + Pn).Append(';').AppendLine;
          end;
        end;
      end;
    end;
    SB.Append('  end;');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

class function TExtractInterfaceEngine.BuildNewUnitText(
  const AInfo: TExtractInterfaceInfo): string;
var
  UName: string;
  SB: TStringBuilder;
  I: Integer;
begin
  UName := ChangeFileExt(ExtractFileName(AInfo.TargetFile), '');
  SB := TStringBuilder.Create;
  try
    SB.Append('unit ').Append(UName).Append(';').AppendLine.AppendLine;
    SB.Append('interface').AppendLine.AppendLine;
    SB.Append('uses').AppendLine;
    if Length(AInfo.NewUnitUses) > 0 then
    begin
      // Emit comma-separated on one or more wrapped lines (Delphi
      // convention) instead of one-unit-per-line. Wrap when the
      // current line would exceed ~75 chars including the trailing
      // separator.
      var Line: string := '  ';
      for I := 0 to High(AInfo.NewUnitUses) do
      begin
        var Sep: string;
        if I = High(AInfo.NewUnitUses) then Sep := ';' else Sep := ', ';
        var Add: string := AInfo.NewUnitUses[I] + Sep;
        // If Line already has content beyond the initial indent and
        // adding Add would overflow, flush Line and start a new one.
        if (Length(Line) > 2) and (Length(Line) + Length(Add) > 75) then
        begin
          SB.Append(TrimRight(Line)).AppendLine;
          Line := '  ';
        end;
        Line := Line + Add;
      end;
      if Length(Line) > 2 then
        SB.Append(TrimRight(Line)).AppendLine;
    end
    else
      SB.Append('  System.Classes, System.SysUtils;').AppendLine;
    SB.AppendLine;
    SB.Append('type').AppendLine;
    SB.Append(BuildInterfaceText(AInfo)).AppendLine.AppendLine;
    SB.Append('implementation').AppendLine.AppendLine;
    SB.Append('end.').AppendLine;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

class function TExtractInterfaceEngine.ParseExistingInterfaceNames(
  const AFile: string; ADeclLine, AEndLine: Integer): TArray<string>;
var
  Lines: TArray<string>;
  L: TList<string>;
  I, P, Q: Integer;
  Line, Trimmed, Upper: string;
  procedure AddIdentAfter(const AKeyword: string; const ATrimmedLine: string);
  var
    UpLine: string;
    Idx, Start: Integer;
  begin
    UpLine := UpperCase(ATrimmedLine);
    if not StartsText(AKeyword + ' ', UpLine) then Exit;
    Idx := Length(AKeyword) + 1;
    while (Idx <= Length(ATrimmedLine)) and (ATrimmedLine[Idx] = ' ') do Inc(Idx);
    Start := Idx;
    while (Idx <= Length(ATrimmedLine)) and
          (ATrimmedLine[Idx].IsLetterOrDigit or (ATrimmedLine[Idx] = '_')) do
      Inc(Idx);
    if Idx > Start then L.Add(Copy(ATrimmedLine, Start, Idx - Start));
  end;
  procedure CollectAfterKeyword(const AKW, AUpAcc, AOrigAcc: string);
  var
    KP, Start: Integer;
    Padded: string;
  begin
    Padded := ' ' + AUpAcc + ' ';
    KP := Pos(' ' + AKW + ' ', Padded);
    if KP = 0 then Exit;
    KP := KP + Length(AKW); // index in Padded after the keyword
    while (KP <= Length(Padded)) and (Padded[KP] = ' ') do Inc(KP);
    Start := KP;
    while (KP <= Length(Padded)) and
          (Padded[KP].IsLetterOrDigit or (Padded[KP] = '_')) do
      Inc(KP);
    if KP > Start then
      L.Add(Copy(AOrigAcc, Start - 1, KP - Start));
  end;
begin
  Result := nil;
  Lines := TDelphiFileEncoding.ReadLines(AFile);
  L := TList<string>.Create;
  try
    for I := ADeclLine to AEndLine - 1 do
    begin
      if (I < 1) or (I > Length(Lines)) then Continue;
      Line := StripLineComment(Lines[I - 1]);
      Trimmed := Trim(Line);
      if Trimmed = '' then Continue;
      Upper := UpperCase(Trimmed);

      if StartsText('PROCEDURE ', Upper) or StartsText('FUNCTION ', Upper) then
      begin
        AddIdentAfter('PROCEDURE', Trimmed);
        AddIdentAfter('FUNCTION', Trimmed);
      end
      else if StartsText('PROPERTY ', Upper) then
      begin
        AddIdentAfter('PROPERTY', Trimmed);
        CollectAfterKeyword('READ', Upper, Trimmed);
        CollectAfterKeyword('WRITE', Upper, Trimmed);
      end;

      if P = 0 then ;
      if Q = 0 then ;
    end;
    Result := L.ToArray;
  finally
    L.Free;
  end;
end;

class function TExtractInterfaceEngine.WouldEmitNames(
  const M: TClassMember): TArray<string>;
var
  L: TList<string>;
begin
  L := TList<string>.Create;
  try
    case M.Kind of
      mkMethod: L.Add(M.Name);
      mkField:
      begin
        L.Add(M.ExposedName);          // interface property name
        L.Add('Get' + M.ExposedName);
        L.Add('Set' + M.ExposedName);
      end;
      mkProperty:
      begin
        L.Add(M.ExposedName);
        if M.NeedsSynthAccessors then
        begin
          L.Add('Get' + M.ExposedName);
          if not M.IsReadOnly then L.Add('Set' + M.ExposedName);
        end;
      end;
    end;
    Result := L.ToArray;
  finally
    L.Free;
  end;
end;

class function TExtractInterfaceEngine.ClashesWithExisting(
  const M: TClassMember; const AExistingUpper: TArray<string>): Boolean;
var
  Mine: TArray<string>;
  N, E: string;
begin
  Result := False;
  Mine := WouldEmitNames(M);
  for N in Mine do
    for E in AExistingUpper do
      if SameText(N, E) then Exit(True);
end;

{ ---------- TProjectInterfaceScanner ---------- }

class function TProjectInterfaceScanner.ScanFile(
  const AFileName: string): TArray<TInterfaceDeclLocation>;
var
  Lines: TArray<string>;
  I, J: Integer;
  L, Trimmed, Upper, Name: string;
  EqPos, IfacePos, Depth: Integer;
  List: TList<TInterfaceDeclLocation>;
begin
  Result := nil;
  if not TFile.Exists(AFileName) then Exit;
  try
    Lines := TDelphiFileEncoding.ReadLines(AFileName);
  except
    Exit;
  end;
  List := TList<TInterfaceDeclLocation>.Create;
  try
    I := 0;
    while I < Length(Lines) do
    begin
      L := StripLineComment(Lines[I]);
      Trimmed := Trim(L);
      Upper := UpperCase(Trimmed);
      EqPos := Pos('=', Trimmed);
      if EqPos > 0 then
      begin
        IfacePos := Pos('INTERFACE', Upper);
        if (IfacePos > EqPos) then
        begin
          // exclude the unit's "interface" section keyword
          var Name_ := Trim(Copy(Trimmed, 1, EqPos - 1));
          // exclude generic-suffix
          var GenPos := Pos('<', Name_);
          if GenPos > 0 then Name_ := Trim(Copy(Name_, 1, GenPos - 1));
          if (Name_ <> '') and IsIdentStart(Name_[1]) then
          begin
            // Find the matching 'end;' tracking record/class nesting.
            Depth := 1;
            var EndLineFound := -1;
            J := I;
            while J < Length(Lines) do
            begin
              var LineU := UpperCase(StripLineComment(Lines[J]));
              var Pos_ := 1;
              while Pos_ <= Length(LineU) do
              begin
                if IsIdentStart(LineU[Pos_]) then
                begin
                  var Q := Pos_;
                  while (Q <= Length(LineU)) and IsIdent(LineU[Q]) do Inc(Q);
                  var W := Copy(LineU, Pos_, Q - Pos_);
                  if (W = 'RECORD') and (J > I) then Inc(Depth)
                  else if (W = 'END') then
                  begin
                    Dec(Depth);
                    if Depth = 0 then begin EndLineFound := J + 1; Break; end;
                  end;
                  Pos_ := Q;
                end
                else
                  Inc(Pos_);
              end;
              if EndLineFound > 0 then Break;
              Inc(J);
            end;
            if EndLineFound > 0 then
            begin
              var Loc: TInterfaceDeclLocation;
              Loc.FileName := AFileName;
              Loc.InterfaceName := Name_;
              Loc.DeclLine := I + 1;
              Loc.EndLine := EndLineFound;
              List.Add(Loc);
              I := EndLineFound;
              Continue;
            end;
          end;
        end;
      end;
      Inc(I);
    end;
    Result := List.ToArray;
  finally
    List.Free;
  end;
end;

class function TProjectInterfaceScanner.ScanProject(
  const AFiles: TArray<string>): TArray<TInterfaceDeclLocation>;
var
  All: TList<TInterfaceDeclLocation>;
  F: string;
begin
  All := TList<TInterfaceDeclLocation>.Create;
  try
    for F in AFiles do
      All.AddRange(ScanFile(F));
    Result := All.ToArray;
  finally
    All.Free;
  end;
end;

end.
