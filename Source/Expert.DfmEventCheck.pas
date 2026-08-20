(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.DfmEventCheck;

// Checks .dfm files for event references that have no matching handler
// method in the paired .pas file - the classic "Method 'BtnOkClick'
// not found" crash at form load time that the compiler never warns
// about. Also flags handlers whose parameter list does not match the
// event type's expected signature (for a table of well-known VCL
// events); a mismatched signature loads fine but corrupts the stack
// when the event fires.
//
// Scope / limitations (v1):
//   * Text DFMs only. Binary DFMs are skipped (reported as unparsed).
//   * The handler search covers the form class named in the DFM root
//     plus its ancestor classes as far as they can be resolved inside
//     the project files (visual form inheritance).
//   * Signature checking only applies to events in the built-in table;
//     unknown events are checked for handler existence only.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TDfmEventIssueKind = (eikMissingHandler, eikSignatureMismatch);

  TDfmEventIssue = record
    DfmFile: string;
    PasFile: string;
    ComponentName: string;
    ComponentType: string;
    EventName: string;
    HandlerName: string;
    /// <summary>1-based line of the event assignment in the .dfm.</summary>
    DfmLine: Integer;
    /// <summary>1-based line of the handler declaration in the .pas
    ///  (0 when the handler is missing).</summary>
    PasLine: Integer;
    Kind: TDfmEventIssueKind;
    /// <summary>Human-readable expected / actual parameter lists,
    ///  filled for eikSignatureMismatch.</summary>
    Expected: string;
    Actual: string;
  end;

  TDfmEventChecker = class
  private
    class function NormalizeParams(const AParamList: string): string;
    class function ExpectedSignature(const AEventName: string;
      out ASignature: string): Boolean;
    class procedure CollectClassMethods(const ALines: TArray<string>;
      const AClassName: string; AMethods: TDictionary<string, TPair<Integer, string>>;
      out AAncestor: string);
  public
    /// <summary>Checks one .dfm / .pas pair. APasLines may be pre-read
    ///  (live buffer); pass nil to read from disk.</summary>
    class function CheckPair(const ADfmFile, APasFile: string;
      const AProjectFiles: TArray<string>): TArray<TDfmEventIssue>;

    /// <summary>Runs CheckPair for every project .pas that has a
    ///  sibling .dfm.</summary>
    class function CheckProject(const AProjectFiles: TArray<string>): TArray<TDfmEventIssue>;
  end;

implementation

uses
  System.IOUtils, System.StrUtils,
  Expert.EditorHelperIntf, Delphi.FileEncoding;

function ReadFileLines(const AFile: string): TArray<string>;
var
  Content: string;
begin
  Result := nil;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, Content) then
    Exit(Content.Split([sLineBreak], TStringSplitOptions.None));
  if not TFile.Exists(AFile) then Exit;
  try
    Content := TDelphiFileEncoding.ReadAll(AFile);
    Result := Content.Split([sLineBreak], TStringSplitOptions.None);
  except
    Result := nil;
  end;
end;

function IsIdent(const S: string): Boolean;
var
  I: Integer;
begin
  Result := S <> '';
  for I := 1 to Length(S) do
    if not CharInSet(S[I], ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
      Exit(False);
  if CharInSet(S[1], ['0'..'9']) then Result := False;
end;

{ ---------------- Event signature table ----------------
  Signatures are normalized as pipe-separated entries, one per
  parameter, each "[var |out ]TypeName" (const dropped - it does not
  affect the call ABI for these types and keeps false positives down).
  Multi-name groups (X, Y: Integer) expand to one entry per name. }

type
  TKnownEvent = record
    Name: string;       // event property name, case-insensitive
    Signature: string;  // normalized expected parameter list
  end;

const
  // (Sender: TObject) - the TNotifyEvent family.
  NotifyEvents: array[0..21] of string = (
    'OnClick', 'OnDblClick', 'OnChange', 'OnChanged', 'OnEnter', 'OnExit',
    'OnCreate', 'OnDestroy', 'OnShow', 'OnHide', 'OnActivate',
    'OnDeactivate', 'OnResize', 'OnPopup', 'OnExecute', 'OnUpdate',
    'OnTimer', 'OnSelectionChange', 'OnStartDock', 'OnStartDrag',
    'OnPaint', 'OnCellClick');

  KnownEvents: array[0..11] of TKnownEvent = (
    (Name: 'OnKeyDown';    Signature: 'TObject|var Word|TShiftState'),
    (Name: 'OnKeyUp';      Signature: 'TObject|var Word|TShiftState'),
    (Name: 'OnKeyPress';   Signature: 'TObject|var Char'),
    (Name: 'OnMouseDown';  Signature: 'TObject|TMouseButton|TShiftState|Integer|Integer'),
    (Name: 'OnMouseUp';    Signature: 'TObject|TMouseButton|TShiftState|Integer|Integer'),
    (Name: 'OnMouseMove';  Signature: 'TObject|TShiftState|Integer|Integer'),
    (Name: 'OnClose';      Signature: 'TObject|var TCloseAction'),
    (Name: 'OnCloseQuery'; Signature: 'TObject|var Boolean'),
    (Name: 'OnMouseWheel'; Signature: 'TObject|TShiftState|Integer|TPoint|var Boolean'),
    (Name: 'OnDragOver';   Signature: 'TObject|TObject|Integer|Integer|TDragState|var Boolean'),
    (Name: 'OnDragDrop';   Signature: 'TObject|TObject|Integer|Integer'),
    (Name: 'OnCanResize';  Signature: 'TObject|var Integer|var Integer|var Boolean'));

class function TDfmEventChecker.ExpectedSignature(const AEventName: string;
  out ASignature: string): Boolean;
var
  I: Integer;
begin
  for I := Low(KnownEvents) to High(KnownEvents) do
    if SameText(KnownEvents[I].Name, AEventName) then
    begin
      ASignature := KnownEvents[I].Signature;
      Exit(True);
    end;
  for I := Low(NotifyEvents) to High(NotifyEvents) do
    if SameText(NotifyEvents[I], AEventName) then
    begin
      ASignature := 'TObject';
      Exit(True);
    end;
  ASignature := '';
  Result := False;
end;

class function TDfmEventChecker.NormalizeParams(const AParamList: string): string;
// "Sender: TObject; var Key: Word" -> "TObject|var Word"
// "X, Y: Integer" -> "Integer|Integer"
var
  Segs: TArray<string>;
  Seg, Prefix, Names, TypeName: string;
  ColonPos, NameCount, I: Integer;
  Parts: TArray<string>;
  Out_: TStringBuilder;
begin
  Out_ := TStringBuilder.Create;
  try
    Segs := AParamList.Split([';']);
    for Seg in Segs do
    begin
      var T := Trim(Seg);
      if T = '' then Continue;
      Prefix := '';
      if StartsText('var ', T) then begin Prefix := 'var '; Delete(T, 1, 4); end
      else if StartsText('out ', T) then begin Prefix := 'out '; Delete(T, 1, 4); end
      else if StartsText('const ', T) then Delete(T, 1, 6);  // const dropped
      ColonPos := Pos(':', T);
      if ColonPos = 0 then
      begin
        // Untyped parameter (var X) - keep the prefix, type '?'.
        TypeName := '?';
        Names := Trim(T);
      end
      else
      begin
        Names := Trim(Copy(T, 1, ColonPos - 1));
        TypeName := Trim(Copy(T, ColonPos + 1, MaxInt));
        // Strip default values ("= nil") and unit prefixes
        // ("Vcl.Controls.TMouseButton" -> "TMouseButton").
        var EqPos := Pos('=', TypeName);
        if EqPos > 0 then TypeName := Trim(Copy(TypeName, 1, EqPos - 1));
        var DotPos := LastDelimiter('.', TypeName);
        if DotPos > 0 then TypeName := Copy(TypeName, DotPos + 1, MaxInt);
      end;
      Parts := Names.Split([',']);
      NameCount := 0;
      for I := 0 to High(Parts) do
        if Trim(Parts[I]) <> '' then Inc(NameCount);
      if NameCount = 0 then NameCount := 1;
      for I := 1 to NameCount do
      begin
        if Out_.Length > 0 then Out_.Append('|');
        Out_.Append(Prefix).Append(TypeName);
      end;
    end;
    Result := Out_.ToString;
  finally
    Out_.Free;
  end;
end;

class procedure TDfmEventChecker.CollectClassMethods(
  const ALines: TArray<string>; const AClassName: string;
  AMethods: TDictionary<string, TPair<Integer, string>>;
  out AAncestor: string);
// Finds "AClassName = class(...)" and collects every method
// declaration up to the closing "end;". Values are (line, normalized
// param list). AAncestor returns the parenthesised ancestor name (''
// if none / TObject-implied).
var
  I, J, Depth: Integer;
  L, U, Decl, MName, Params: string;
  InClass: Boolean;
  P1, P2: Integer;
begin
  AAncestor := '';
  InClass := False;
  Depth := 0;
  I := 0;
  while I < Length(ALines) do
  begin
    L := ALines[I];
    U := UpperCase(Trim(L));
    if not InClass then
    begin
      // "TFORM1 = CLASS" with optional "(ANCESTOR)" - reject
      // "= CLASS;" (forward) and "= CLASS OF".
      var Key := UpperCase(AClassName) + ' = CLASS';
      var KP := Pos(Key, U);
      if (KP = 1) and not U.Contains('= CLASS;') and not U.Contains('= CLASS OF') then
      begin
        InClass := True;
        P1 := Pos('(', L);
        if P1 > 0 then
        begin
          P2 := Pos(')', L);
          if P2 > P1 then
          begin
            AAncestor := Trim(Copy(L, P1 + 1, P2 - P1 - 1));
            // Multi-interface lists: first entry is the class ancestor.
            var CommaPos := Pos(',', AAncestor);
            if CommaPos > 0 then AAncestor := Trim(Copy(AAncestor, 1, CommaPos - 1));
            var DotPos := LastDelimiter('.', AAncestor);
            if DotPos > 0 then AAncestor := Copy(AAncestor, DotPos + 1, MaxInt);
          end;
        end;
      end;
    end
    else
    begin
      if (U = 'END;') and (Depth = 0) then Exit;
      // Rudimentary nested-structure tracking (records/classes inside).
      if U.EndsWith('= RECORD') or U.EndsWith('= CLASS') then Inc(Depth)
      else if (U = 'END;') and (Depth > 0) then Dec(Depth);

      if (Depth = 0) and (U.StartsWith('PROCEDURE ') or U.StartsWith('FUNCTION ')) then
      begin
        // Concatenate lines until ';' closes the declaration.
        Decl := Trim(L);
        J := I;
        while (Pos(';', Decl) = 0) and (J + 1 < Length(ALines)) do
        begin
          Inc(J);
          Decl := Decl + ' ' + Trim(ALines[J]);
        end;
        // Extract name + params.
        var SP := Pos(' ', Decl);
        var Rest := Trim(Copy(Decl, SP + 1, MaxInt));
        P1 := Pos('(', Rest);
        var SemiPos := Pos(';', Rest);
        if (P1 > 0) and ((SemiPos = 0) or (P1 < SemiPos)) then
        begin
          MName := Trim(Copy(Rest, 1, P1 - 1));
          P2 := Pos(')', Rest);
          if P2 > P1 then
            Params := Copy(Rest, P1 + 1, P2 - P1 - 1)
          else
            Params := '';
        end
        else
        begin
          if SemiPos > 0 then MName := Trim(Copy(Rest, 1, SemiPos - 1))
          else MName := Rest;
          // Strip ": ReturnType" from parameterless functions.
          var CPos := Pos(':', MName);
          if CPos > 0 then MName := Trim(Copy(MName, 1, CPos - 1));
          Params := '';
        end;
        if IsIdent(MName) and not AMethods.ContainsKey(UpperCase(MName)) then
          AMethods.Add(UpperCase(MName),
            TPair<Integer, string>.Create(I + 1, NormalizeParams(Params)));
        I := J;
      end;
    end;
    Inc(I);
  end;
end;

class function TDfmEventChecker.CheckPair(const ADfmFile, APasFile: string;
  const AProjectFiles: TArray<string>): TArray<TDfmEventIssue>;
var
  DfmLines, PasLines: TArray<string>;
  Issues: TList<TDfmEventIssue>;
  Methods: TDictionary<string, TPair<Integer, string>>;
  CompStack: TStack<TPair<string, string>>;  // name, type
  FormClass, Ancestor, NextClass, ProbeFile: string;
  I, Hops: Integer;
  Issue: TDfmEventIssue;
begin
  Result := nil;
  DfmLines := ReadFileLines(ADfmFile);
  if Length(DfmLines) = 0 then Exit;
  // Binary DFM guard: text DFMs start with object/inherited/inline.
  var FirstWord := UpperCase(Trim(DfmLines[0]).Split([' '])[0]);
  if (FirstWord <> 'OBJECT') and (FirstWord <> 'INHERITED') and (FirstWord <> 'INLINE') then
    Exit;

  // Root line: "object Form1: TForm1"
  var RootParts := Trim(DfmLines[0]).Split([':']);
  if Length(RootParts) < 2 then Exit;
  FormClass := Trim(RootParts[1]);

  PasLines := ReadFileLines(APasFile);
  if Length(PasLines) = 0 then Exit;

  Methods := TDictionary<string, TPair<Integer, string>>.Create;
  Issues := TList<TDfmEventIssue>.Create;
  CompStack := TStack<TPair<string, string>>.Create;
  try
    // Collect methods of the form class + project-resolvable ancestors
    // (visual form inheritance keeps handlers in the base class).
    CollectClassMethods(PasLines, FormClass, Methods, Ancestor);
    Hops := 0;
    NextClass := Ancestor;
    while (NextClass <> '') and (Hops < 8) do
    begin
      Inc(Hops);
      var Found := False;
      for ProbeFile in AProjectFiles do
      begin
        if not SameText(ExtractFileExt(ProbeFile), '.pas') then Continue;
        var Probe := ReadFileLines(ProbeFile);
        if Length(Probe) = 0 then Continue;
        var NextAncestor: string;
        var CountBefore := Methods.Count;
        CollectClassMethods(Probe, NextClass, Methods, NextAncestor);
        if (Methods.Count > CountBefore) or (NextAncestor <> '') then
        begin
          NextClass := NextAncestor;
          Found := True;
          Break;
        end;
      end;
      if not Found then Break;
    end;

    // Walk the DFM: track the object stack, find "OnXxx = Handler".
    CompStack.Push(TPair<string, string>.Create(
      Trim(RootParts[0]).Split([' '])[High(Trim(RootParts[0]).Split([' ']))], FormClass));
    I := 1;
    while I < Length(DfmLines) do
    begin
      var L := Trim(DfmLines[I]);
      var U := UpperCase(L);
      if U.StartsWith('OBJECT ') or U.StartsWith('INHERITED ') or U.StartsWith('INLINE ') then
      begin
        var Parts := L.Split([' ', ':'], TStringSplitOptions.ExcludeEmpty);
        if Length(Parts) >= 3 then
          CompStack.Push(TPair<string, string>.Create(Parts[1], Parts[2]))
        else if Length(Parts) >= 2 then
          CompStack.Push(TPair<string, string>.Create(Parts[1], ''));
      end
      else if U = 'END' then
      begin
        if CompStack.Count > 1 then CompStack.Pop;
      end
      else
      begin
        // Multi-line property values: skip blocks opened by trailing
        // delimiters. Simplified but effective for real DFMs.
        if U.EndsWith('= {') or U.EndsWith('{') then
        begin
          while (I + 1 < Length(DfmLines)) and not Trim(DfmLines[I]).EndsWith('}') do Inc(I);
        end
        else if U.EndsWith('= (') then
        begin
          while (I + 1 < Length(DfmLines)) and not Trim(DfmLines[I]).EndsWith(')') do Inc(I);
        end
        else if U.EndsWith('= <') then
        begin
          // Collection property - items may contain events too, but
          // their handlers follow the same rules; keep scanning inside.
          // (No skip: item> lines parse as normal properties.)
        end
        else
        begin
          var EqPos := Pos(' = ', L);
          if EqPos > 0 then
          begin
            var PropName := Trim(Copy(L, 1, EqPos - 1));
            var PropValue := Trim(Copy(L, EqPos + 3, MaxInt));
            // Event heuristic: last dotted segment starts with 'On',
            // value is a bare identifier (no quotes, no digits-only,
            // not True/False/nil, no dots).
            var DotPos := LastDelimiter('.', PropName);
            if DotPos > 0 then PropName := Copy(PropName, DotPos + 1, MaxInt);
            if StartsText('On', PropName) and IsIdent(PropValue)
               and not SameText(PropValue, 'True') and not SameText(PropValue, 'False')
               and not SameText(PropValue, 'nil') then
            begin
              var MethodInfo: TPair<Integer, string>;
              if not Methods.TryGetValue(UpperCase(PropValue), MethodInfo) then
              begin
                Issue := Default(TDfmEventIssue);
                Issue.DfmFile := ADfmFile;
                Issue.PasFile := APasFile;
                if CompStack.Count > 0 then
                begin
                  Issue.ComponentName := CompStack.Peek.Key;
                  Issue.ComponentType := CompStack.Peek.Value;
                end;
                Issue.EventName := PropName;
                Issue.HandlerName := PropValue;
                Issue.DfmLine := I + 1;
                Issue.Kind := eikMissingHandler;
                Issues.Add(Issue);
              end
              else
              begin
                var Expected: string;
                if ExpectedSignature(PropName, Expected)
                   and not SameText(MethodInfo.Value, Expected) then
                begin
                  Issue := Default(TDfmEventIssue);
                  Issue.DfmFile := ADfmFile;
                  Issue.PasFile := APasFile;
                  if CompStack.Count > 0 then
                  begin
                    Issue.ComponentName := CompStack.Peek.Key;
                    Issue.ComponentType := CompStack.Peek.Value;
                  end;
                  Issue.EventName := PropName;
                  Issue.HandlerName := PropValue;
                  Issue.DfmLine := I + 1;
                  Issue.PasLine := MethodInfo.Key;
                  Issue.Kind := eikSignatureMismatch;
                  Issue.Expected := '(' + Expected.Replace('|', '; ') + ')';
                  Issue.Actual := '(' + MethodInfo.Value.Replace('|', '; ') + ')';
                  Issues.Add(Issue);
                end;
              end;
            end;
          end;
        end;
      end;
      Inc(I);
    end;

    Result := Issues.ToArray;
  finally
    CompStack.Free;
    Issues.Free;
    Methods.Free;
  end;
end;

class function TDfmEventChecker.CheckProject(
  const AProjectFiles: TArray<string>): TArray<TDfmEventIssue>;
var
  All: TList<TDfmEventIssue>;
  F, DfmFile: string;
begin
  All := TList<TDfmEventIssue>.Create;
  try
    for F in AProjectFiles do
    begin
      if not SameText(ExtractFileExt(F), '.pas') then Continue;
      DfmFile := ChangeFileExt(F, '.dfm');
      if not TFile.Exists(DfmFile) then Continue;
      All.AddRange(CheckPair(DfmFile, F, AProjectFiles));
    end;
    Result := All.ToArray;
  finally
    All.Free;
  end;
end;

end.
