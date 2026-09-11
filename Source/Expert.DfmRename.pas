(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DfmRename;

// Rename support for form files (.dfm / .fmx, TEXT format).
//
// A rename that only touches the .pas breaks the form: renaming the field
// Button1 leaves "object Button1: TButton" behind (-> "field not found" at
// load time), renaming a method leaves "OnClick = Button1Click" pointing at
// nothing. The LSP knows nothing about form files, so they are handled
// here, textually - but with the SAME question the LSP answers for code:
// does this occurrence refer to the renamed symbol?
//
// That is decided by the class that OWNS a name at a given place:
//  * a component name in an object header is a field of the class whose
//    components are being streamed there - the root form's class, or,
//    inside an "inline" section, the frame's class;
//  * a bare identifier value ("PopupMenu = PopupMenu1") references a
//    component of that owner (or of the root);
//  * an event handler ("OnClick = Button1Click") is a method of the ROOT
//    class - handlers set on an inline frame's controls live in the host;
//  * "DataModule1.DataSource1" qualifies with another root (or a local
//    component such as an inline frame);
//  * the left-hand property name of an object belongs to that object's
//    class (renaming a published property of an own component);
//  * a class name in a header ("object Button1: TMyButton").
// The caller supplies "does class X match the renamed owner type (or
// descend from it)?" - so Button1 of an unrelated form is never touched.
//
// Binary form files are not supported (IsTextFormFile says so).

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TDfmTargetKind = (
    dtField,     // a component field of a form/frame/data module class
    dtMethod,    // a method - event handlers
    dtProperty,  // a published property of a component class
    dtType);     // a class name

  TDfmHitKind = (dhComponentName, dhComponentRef, dhQualifiedRef,
    dhEventHandler, dhPropertyName, dhClassName);

  TDfmHit = record
    Line: Integer;   // 0-based
    Col: Integer;    // 0-based
    Len: Integer;
    Kind: TDfmHitKind;
  end;

  TDfmRoot = record
    Name: string;
    ClassName: string;
  end;

  /// <summary>True when AClass is the renamed symbol's owner type or
  ///  descends from it.</summary>
  TDfmClassMatch = reference to function(const AClass: string): Boolean;
  /// <summary>True when AQualifier names another root (a data module, a
  ///  form) whose class matches the owner type.</summary>
  TDfmQualifierMatch = reference to function(const AQualifier: string): Boolean;

/// <summary>Root object of a text form file ('' fields when not found).</summary>
function ReadDfmRoot(const AText: string): TDfmRoot;

/// <summary>Every place in AText that refers to AOldName in the sense of
///  ATarget. AClassMatch decides ownership (see the unit header);
///  AQualifierMatch may be nil.</summary>
function FindDfmRenameHits(const AText, AOldName: string;
  ATarget: TDfmTargetKind; const AClassMatch: TDfmClassMatch;
  const AQualifierMatch: TDfmQualifierMatch): TArray<TDfmHit>;

/// <summary>Applies the hits bottom-up. A hit whose text no longer is
///  AOldName (case-insensitively) is skipped and counted in ASkipped.</summary>
function ApplyDfmHits(const AText, AOldName, ANewName: string;
  const AHits: TArray<TDfmHit>; out ASkipped: Integer): string;

/// <summary>Text form files start with object / inherited / inline;
///  binary ones with a resource header or 'TPF0'.</summary>
function IsTextFormFile(const AFileName: string): Boolean;

/// <summary>The .dfm or .fmx next to a unit, '' when there is none.</summary>
function FormFileOf(const APasFile: string): string;

function DfmHitKindText(AKind: TDfmHitKind): string;

/// <summary>UPPER class name -> UPPER direct parent (unit prefix cut),
///  from "TFoo = class(TBar" declarations in the given source lines.</summary>
procedure CollectClassParents(const ALines: TArray<string>;
  AMap: TDictionary<string, string>);

/// <summary>AClass = AType, or AClass descends from AType according to
///  AParents (case-insensitive).</summary>
function ClassMatchesType(AParents: TDictionary<string, string>;
  const AClass, AType: string): Boolean;

implementation

uses
  System.IOUtils, System.StrUtils, System.Character, System.Generics.Defaults;

function IsIdentStart(C: Char): Boolean; inline;
begin
  Result := C.IsLetter or (C = '_');
end;

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := C.IsLetterOrDigit or (C = '_');
end;

// Reads an identifier starting at 1-based P; returns it and moves P behind.
function ReadIdent(const S: string; var P: Integer): string;
var
  Start: Integer;
begin
  Start := P;
  while (P <= Length(S)) and IsIdentChar(S[P]) do Inc(P);
  Result := Copy(S, Start, P - Start);
end;

procedure SkipSpaces(const S: string; var P: Integer); inline;
begin
  while (P <= Length(S)) and ((S[P] = ' ') or (S[P] = #9)) do Inc(P);
end;

// Position (1-based) of an unquoted AChar in S at or after AFrom, 0 = none.
function UnquotedPos(AChar: Char; const S: string; AFrom: Integer): Integer;
var
  I: Integer;
  InStr: Boolean;
begin
  InStr := False;
  for I := AFrom to Length(S) do
  begin
    if S[I] = '''' then InStr := not InStr
    else if (not InStr) and (S[I] = AChar) then Exit(I);
  end;
  Result := 0;
end;

type
  THeader = record
    Ok: Boolean;
    Keyword: string;     // object / inherited / inline (lower case)
    Name: string;
    NameCol: Integer;    // 1-based, 0 = unnamed
    ClassName: string;
    ClassCol: Integer;   // 1-based
  end;

function ParseHeader(const ALine: string): THeader;
var
  P: Integer;
  W: string;
begin
  Result := Default(THeader);
  P := 1;
  SkipSpaces(ALine, P);
  if (P > Length(ALine)) or not IsIdentStart(ALine[P]) then Exit;
  W := LowerCase(ReadIdent(ALine, P));
  if (W <> 'object') and (W <> 'inherited') and (W <> 'inline') then Exit;
  Result.Keyword := W;
  SkipSpaces(ALine, P);
  if (P > Length(ALine)) or not IsIdentStart(ALine[P]) then Exit;
  var FirstCol := P;
  var First := ReadIdent(ALine, P);
  SkipSpaces(ALine, P);
  if (P <= Length(ALine)) and (ALine[P] = ':') then
  begin
    // "object Name: TClass"
    Inc(P);
    SkipSpaces(ALine, P);
    if (P > Length(ALine)) or not IsIdentStart(ALine[P]) then Exit;
    Result.Name := First;
    Result.NameCol := FirstCol;
    Result.ClassCol := P;
    Result.ClassName := ReadIdent(ALine, P);
  end
  else
  begin
    // unnamed: "object TClass"
    Result.ClassCol := FirstCol;
    Result.ClassName := First;
  end;
  Result.Ok := True;
end;

function ReadDfmRoot(const AText: string): TDfmRoot;
begin
  Result := Default(TDfmRoot);
  for var L in AText.Split([#10]) do
  begin
    var S := L.TrimRight([#13]);
    if Trim(S) = '' then Continue;
    var H := ParseHeader(S);
    if H.Ok then
    begin
      Result.Name := H.Name;
      Result.ClassName := H.ClassName;
    end;
    Exit;   // the first non-empty line decides
  end;
end;

type
  TDfmFrameKind = (fkObject, fkCollection, fkItem);
  TDfmFrame = record
    Kind: TDfmFrameKind;
    ChildOwner: string;   // class owning component names below this object
    ObjClass: string;     // the object's own class (for property names)
  end;

function FindDfmRenameHits(const AText, AOldName: string;
  ATarget: TDfmTargetKind; const AClassMatch: TDfmClassMatch;
  const AQualifierMatch: TDfmQualifierMatch): TArray<TDfmHit>;
var
  Lines: TArray<string>;
  Stack: TList<TDfmFrame>;
  Hits: TList<TDfmHit>;
  LocalClass: TDictionary<string, string>;   // UPPER component name -> class
  RootClass: string;
  InBinary, InList: Boolean;

  procedure Hit(ALine, ACol1, ALen: Integer; AKind: TDfmHitKind);
  var
    H: TDfmHit;
  begin
    H.Line := ALine;
    H.Col := ACol1 - 1;
    H.Len := ALen;
    H.Kind := AKind;
    Hits.Add(H);
  end;

  function Matches(const AClass: string): Boolean;
  begin
    Result := (AClass <> '') and Assigned(AClassMatch) and AClassMatch(AClass);
  end;

  function CurrentOwner: string;
  var
    I: Integer;
  begin
    for I := Stack.Count - 1 downto 0 do
      if Stack[I].Kind = fkObject then Exit(Stack[I].ChildOwner);
    Result := RootClass;
  end;

  function CurrentObjClass: string;
  var
    I: Integer;
  begin
    for I := Stack.Count - 1 downto 0 do
      if Stack[I].Kind = fkObject then Exit(Stack[I].ObjClass);
    Result := '';
  end;

var
  LineNo, P, Q, EqPos: Integer;
  S, T, W: string;
  H: THeader;
  F: TDfmFrame;
begin
  Result := nil;
  if AOldName = '' then Exit;
  Lines := AText.Split([#10]);
  Stack := TList<TDfmFrame>.Create;
  Hits := TList<TDfmHit>.Create;
  LocalClass := TDictionary<string, string>.Create;
  try
    // Pre-pass: component name -> class, for qualifiers such as
    // "Frame21.Edit1" that point INTO a local component.
    for S in Lines do
    begin
      H := ParseHeader(S.TrimRight([#13]));
      if H.Ok and (H.Name <> '') then
        LocalClass.AddOrSetValue(UpperCase(H.Name), H.ClassName);
    end;

    RootClass := '';
    InBinary := False;
    InList := False;
    for LineNo := 0 to High(Lines) do
    begin
      S := Lines[LineNo].TrimRight([#13]);
      T := Trim(S);
      if T = '' then Continue;

      // Multi-line binary data { ... } and lists ( ... ).
      if InBinary then
      begin
        if Pos('}', T) > 0 then InBinary := False;
        Continue;
      end;
      if InList then
      begin
        if UnquotedPos(')', S, 1) > 0 then InList := False;
        Continue;
      end;
      // String continuation lines.
      if CharInSet(T[1], ['''', '#', '+']) then Continue;

      H := ParseHeader(S);
      if H.Ok then
      begin
        F.Kind := fkObject;
        F.ObjClass := H.ClassName;
        if Stack.Count = 0 then
        begin
          RootClass := H.ClassName;
          F.ChildOwner := H.ClassName;
        end
        else
        begin
          // A component name below the root is a FIELD of the current
          // owner. The root's own name is not a field of its class.
          if (ATarget = dtField) and SameText(H.Name, AOldName)
            and Matches(CurrentOwner) then
            Hit(LineNo, H.NameCol, Length(H.Name), dhComponentName);
          if H.Keyword = 'inline' then
            F.ChildOwner := H.ClassName   // a frame streams its own fields
          else
            F.ChildOwner := CurrentOwner;
        end;
        if (ATarget = dtType) and SameText(H.ClassName, AOldName) then
          Hit(LineNo, H.ClassCol, Length(H.ClassName), dhClassName);
        Stack.Add(F);
        Continue;
      end;

      W := LowerCase(T);
      if (W = 'end') or (W = 'end>') then
      begin
        // "end" closes an item or an object, "end>" an item AND its
        // collection.
        if Stack.Count > 0 then Stack.Delete(Stack.Count - 1);
        if (W = 'end>') and (Stack.Count > 0)
          and (Stack.Last.Kind = fkCollection) then
          Stack.Delete(Stack.Count - 1);
        Continue;
      end;
      if W = 'item' then
      begin
        F.Kind := fkItem;
        F.ChildOwner := '';
        F.ObjClass := '';
        Stack.Add(F);
        Continue;
      end;
      if W = '>' then
      begin
        if (Stack.Count > 0) and (Stack.Last.Kind = fkCollection) then
          Stack.Delete(Stack.Count - 1);
        Continue;
      end;

      // Property line: Path = Value
      EqPos := UnquotedPos('=', S, 1);
      if EqPos = 0 then Continue;

      // Left side: the first path segment is a property of the object.
      P := 1;
      SkipSpaces(S, P);
      if (ATarget = dtProperty) and (P <= Length(S)) and IsIdentStart(S[P]) then
      begin
        Q := P;
        W := ReadIdent(S, P);
        // only inside an object, not in a collection item (item classes
        // are not known here)
        if SameText(W, AOldName) and (Stack.Count > 0)
          and (Stack.Last.Kind = fkObject) and Matches(CurrentObjClass) then
          Hit(LineNo, Q, Length(W), dhPropertyName);
      end;

      // Right side.
      P := EqPos + 1;
      SkipSpaces(S, P);
      if P > Length(S) then Continue;
      case S[P] of
        '<':
          begin
            if Pos('>', Copy(S, P, MaxInt)) = 0 then
            begin
              F.Kind := fkCollection;
              F.ChildOwner := '';
              F.ObjClass := '';
              Stack.Add(F);
            end;
            Continue;
          end;
        '(':
          begin
            if UnquotedPos(')', S, P + 1) = 0 then InList := True;
            Continue;
          end;
        '{':
          begin
            if Pos('}', Copy(S, P, MaxInt)) = 0 then InBinary := True;
            Continue;
          end;
      end;
      if not IsIdentStart(S[P]) then Continue;   // string, number, set

      if (ATarget <> dtField) and (ATarget <> dtMethod) then Continue;

      // identifier or dotted identifier
      var Parts := TList<Integer>.Create;   // 1-based start columns
      try
        var Names := TStringList.Create;
        try
          repeat
            Parts.Add(P);
            Names.Add(ReadIdent(S, P));
            if (P <= Length(S)) and (S[P] = '.') and (P < Length(S))
              and IsIdentStart(S[P + 1]) then
              Inc(P)
            else
              Break;
          until False;

          if Names.Count = 1 then
          begin
            if SameText(Names[0], AOldName) then
            begin
              if ATarget = dtMethod then
              begin
                // event handlers resolve against the ROOT
                if Matches(RootClass) then
                  Hit(LineNo, Parts[0], Length(Names[0]), dhEventHandler);
              end
              else if Matches(CurrentOwner) or Matches(RootClass) then
                Hit(LineNo, Parts[0], Length(Names[0]), dhComponentRef);
            end;
          end
          else if (ATarget = dtField)
            and SameText(Names[Names.Count - 1], AOldName) then
          begin
            var Qual := Names[Names.Count - 2];
            var QClass: string;
            var Ok := False;
            if LocalClass.TryGetValue(UpperCase(Qual), QClass) then
              Ok := Matches(QClass)
            else if Assigned(AQualifierMatch) then
              Ok := AQualifierMatch(Qual);
            if Ok then
              Hit(LineNo, Parts[Parts.Count - 1],
                Length(Names[Names.Count - 1]), dhQualifiedRef);
          end;
        finally
          Names.Free;
        end;
      finally
        Parts.Free;
      end;
    end;
    Result := Hits.ToArray;
  finally
    LocalClass.Free;
    Hits.Free;
    Stack.Free;
  end;
end;

function ApplyDfmHits(const AText, AOldName, ANewName: string;
  const AHits: TArray<TDfmHit>; out ASkipped: Integer): string;
var
  Lines: TArray<string>;
  Sorted: TArray<TDfmHit>;
  I: Integer;
begin
  ASkipped := 0;
  Lines := AText.Split([#10]);
  Sorted := Copy(AHits);
  // bottom-up, right-to-left: earlier positions stay valid
  TArray.Sort<TDfmHit>(Sorted, TComparer<TDfmHit>.Construct(
    function(const A, B: TDfmHit): Integer
    begin
      Result := B.Line - A.Line;
      if Result = 0 then Result := B.Col - A.Col;
    end));
  for I := 0 to High(Sorted) do
  begin
    var H := Sorted[I];
    if (H.Line < 0) or (H.Line > High(Lines)) then begin Inc(ASkipped); Continue; end;
    var L := Lines[H.Line];
    if not SameText(Copy(L, H.Col + 1, H.Len), AOldName) then
    begin
      Inc(ASkipped);
      Continue;
    end;
    Lines[H.Line] := Copy(L, 1, H.Col) + ANewName + Copy(L, H.Col + H.Len + 1, MaxInt);
  end;
  Result := string.Join(#10, Lines);
end;

function IsTextFormFile(const AFileName: string): Boolean;
var
  FS: TFileStream;
  B: array[0..3] of Byte;
  N: Integer;
begin
  Result := False;
  if not TFile.Exists(AFileName) then Exit;
  try
    FS := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
    try
      N := FS.Read(B, 4);
    finally
      FS.Free;
    end;
  except
    Exit;
  end;
  if N < 4 then Exit;
  if (B[0] = $FF) and (B[1] = $0A) then Exit;                       // resource
  if (B[0] = Ord('T')) and (B[1] = Ord('P')) and (B[2] = Ord('F')) then Exit;  // TPF0
  Result := True;
end;

function FormFileOf(const APasFile: string): string;
begin
  Result := ChangeFileExt(APasFile, '.dfm');
  if TFile.Exists(Result) then Exit;
  Result := ChangeFileExt(APasFile, '.fmx');
  if TFile.Exists(Result) then Exit;
  Result := '';
end;

function DfmHitKindText(AKind: TDfmHitKind): string;
begin
  case AKind of
    dhComponentName: Result := 'Form: component name';
    dhComponentRef:  Result := 'Form: component reference';
    dhQualifiedRef:  Result := 'Form: qualified component reference';
    dhEventHandler:  Result := 'Form: event handler';
    dhPropertyName:  Result := 'Form: property';
  else
    Result := 'Form: class name';
  end;
end;

procedure CollectClassParents(const ALines: TArray<string>;
  AMap: TDictionary<string, string>);
var
  S: string;
  P, Q: Integer;
  Name, Word, Parent: string;
begin
  for S in ALines do
  begin
    P := 1;
    SkipSpaces(S, P);
    if (P > Length(S)) or not IsIdentStart(S[P]) then Continue;
    Name := ReadIdent(S, P);
    // generic parameters: TFoo<T> = class(...)
    if (P <= Length(S)) and (S[P] = '<') then
    begin
      Q := Pos('>', S, P);
      if Q = 0 then Continue;
      P := Q + 1;
    end;
    SkipSpaces(S, P);
    if (P > Length(S)) or (S[P] <> '=') then Continue;
    Inc(P);
    SkipSpaces(S, P);
    if (P > Length(S)) or not IsIdentStart(S[P]) then Continue;
    Word := LowerCase(ReadIdent(S, P));
    if Word = 'packed' then
    begin
      SkipSpaces(S, P);
      Word := LowerCase(ReadIdent(S, P));
    end;
    if Word <> 'class' then Continue;
    SkipSpaces(S, P);
    if (P > Length(S)) or (S[P] <> '(') then Continue;
    Inc(P);
    SkipSpaces(S, P);
    Parent := '';
    while (P <= Length(S)) and (IsIdentChar(S[P]) or (S[P] = '.')) do
    begin
      Parent := Parent + S[P];
      Inc(P);
    end;
    if Parent = '' then Continue;
    Q := LastDelimiter('.', Parent);
    if Q > 0 then Parent := Copy(Parent, Q + 1, MaxInt);
    AMap.AddOrSetValue(UpperCase(Name), UpperCase(Parent));
  end;
end;

function ClassMatchesType(AParents: TDictionary<string, string>;
  const AClass, AType: string): Boolean;
var
  C, Target: string;
  Guard: Integer;
begin
  C := UpperCase(AClass);
  Target := UpperCase(AType);
  Guard := 0;
  while (C <> '') and (Guard < 64) do
  begin
    if C = Target then Exit(True);
    if (AParents = nil) or not AParents.TryGetValue(C, C) then Break;
    Inc(Guard);
  end;
  Result := False;
end;

end.
