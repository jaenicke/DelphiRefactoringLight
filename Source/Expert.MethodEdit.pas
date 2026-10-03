(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.MethodEdit;

// "Edit methods" - the pure half: which members of a class may be edited and
// moved, what a move writes into the two units, and what it deliberately does
// NOT do. No VCL, no ToolsAPI: lines in, lines out, so every rule here is
// testable without an IDE. The dialog and the MCP tool sit on top.

interface

uses
  System.SysUtils, System.Types, Expert.SignatureEdit;

// ===========================================================================
//  "Edit methods": a member moves into another class (user, 2026-10-03)
// ===========================================================================
//
// The decision behind this planner, and the reason it is this small: a full
// "move method" is a DESIGN question (where does a caller get the new
// instance from, does a member travel along, what about the hierarchy) and a
// tool that answers it produces code that compiles and behaves differently -
// which is why #11 dropped it. The user's answer is better: the tool does the
// MECHANICAL half - declaration out of the old class and into the chosen
// section of the new one, body moved and requalified - and NAMES everything
// it does not do. Then the division of labour is visible instead of implied.
//
// Everything here is pure: lines in, lines out. The IDE side (dialog, uses
// edits, the call list) sits on top of it.

type
  /// <summary>One member as the dialog's list shows it.</summary>
  TClassMemberInfo = record
    Name: string;
    DeclLine: Integer;      // 0-based
    Kind: string;           // procedure / function / constructor / destructor /
                            // property / field
    Directives: string;     // what follows the first ';' of the declaration
    Movable: Boolean;
    Why: string;            // why not, when Movable is False
  end;

  TMethodEditIssueKind = (
    meiVeto,        // the member cannot be moved at all
    meiOwnMember,   // the body uses something that stays in the old class
    meiNote);       // everything else the user has to look at

  TMethodEditIssue = record
    Kind: TMethodEditIssueKind;
    Member: string;
    Text: string;
  end;

  TMethodMovePlan = record
    Ok: Boolean;
    Error: string;
    SourceLines: TArray<string>;   // the source unit as it would be written
    TargetLines: TArray<string>;   // the target unit as it would be written
    Moved: TArray<string>;         // what really moves
    Issues: TArray<TMethodEditIssue>;
    // What the move had to add to a uses clause, '' when nothing was needed.
    SourceUsesAdded: string;       // the target unit, in the source
    TargetUsesAdded: string;       // the source unit, in the target
    TargetUsesInInterface: Boolean;// ... and there it can close a cycle
  end;

  /// <summary>What the dialog (or the tool) asks for. Everything is
  ///  optional: only members plus at least one action is required.</summary>
  TMethodEditRequest = record
    Members: TArray<string>;
    TargetFile: string;          // '' = stay in this unit
    TargetClass: string;         // '' = do not move at all
    Section: string;             // private / protected / public / published
    AddModifiers: TArray<string>;
    RemoveModifiers: TArray<string>;
    SigMember: string;           // '' = no signature change
    SigParams: TArray<TSigParam>;
    SigResultType: string;
    function WantsSomething: Boolean;
  end;

  TMethodEditResult = record
    Ok: Boolean;
    Error: string;
    SourceFile, TargetFile: string;   // TargetFile = '' when nothing moved
    SourceLines, TargetLines: TArray<string>;
    SourceUsesAdded, TargetUsesAdded: string;
    Moved: TArray<string>;
    Issues: TArray<TMethodEditIssue>;
    SameFile: Boolean;
    function Summary: string;
  end;

/// <summary>Pure: what AREQ writes into the one or two units. ATargetContent
///  is only read when the request moves something into another unit; for a
///  move inside one unit it is ignored and the result holds ONE text.</summary>
function PlanMethodEdit(const ASourceFile, ASourceContent, AOwnerType: string;
  const AReq: TMethodEditRequest; const ATargetContent: string): TMethodEditResult;

/// <summary>The members AType declares, in declaration order - the left-hand
///  list of the dialog. Movable is False for what cannot travel alone: an
///  overload, a virtual / dynamic / override / abstract / message member, a
///  published one, or anything whose declaration or body cannot be
///  delimited; Why says which.</summary>
function ClassMembersOf(const ALines: TArray<string>;
  const AType: string): TArray<TClassMemberInfo>;

/// <summary>Which of AMembers the line range AFrom..ATo (0-based, inclusive)
///  covers - what a SELECTION in the editor means for the member list. Two
///  shapes count, because the user may have marked either: a member whose
///  DECLARATION stands in the range, and a member whose IMPLEMENTATION the
///  range reaches (selecting three bodies names the same three members as
///  selecting their three declarations). The answer keeps the declaration
///  order of AMembers, and a member is named once however often the range
///  touches it. The two ends may arrive in either order; a range that covers
///  no member of AOwnerType answers nothing.</summary>
function MembersInLineRange(const ALines: TArray<string>;
  const AMembers: TArray<TClassMemberInfo>; const AOwnerType: string;
  AFrom, ATo: Integer): TArray<string>;

/// <summary>The TEXT of a planned file: ALINES joined with the line break
///  AOLDCONTENT uses. This is the exact inverse of SplitContentLines, which
///  keeps a file's final break as a TRAILING EMPTY ELEMENT - so an unchanged
///  plan comes back byte for byte. The result must be assigned to
///  TStringList.Text (whose setter drops that element again) and never added
///  element by element: TStringList.Text appends a break of its own, so the
///  trailing element turned into a new blank line and every apply grew both
///  units by one (reported by the user after a five-member move).</summary>
function JoinPlannedLines(const ALines: TArray<string>;
  const AOldContent: string): string;

/// <summary>Every whole-word occurrence of ANAMES in AMASKED (the masked
///  lines of one file) - ONE pass for all of them, which is what makes a
///  class with thirty methods affordable. AOut[i] belongs to
///  ANames[AWhich[i]]; X is the 0-based column, Y the 0-based line.</summary>
procedure CollectNameHits(const AMasked: TArray<string>;
  const ANames: TArray<string>; out AOut: TArray<TPoint>;
  out AWhich: TArray<Integer>);

type
  /// <summary>One occurrence a user still has to look at after an edit.</summary>
  TPostEditRef = record
    FilePath: string;
    Line: Integer;      // 0-based
    Col: Integer;       // 0-based
    Member: string;
    Kind: string;       // from Expert.ReferenceKind: Call / Write / ...
    Text: string;       // the source line, trimmed
  end;

/// <summary>What still needs hand work after "Edit methods" wrote: every
///  code occurrence of AMEMBERS in the files AFILES (AContents[i] is the
///  text of AFiles[i], as it is NOW - the edit shifted the lines of the two
///  units it touched), EXCEPT the declaration and the implementation header
///  in the TARGET, which are the edit's own result and nothing to adjust.
///  Sorted by file and line, the order someone works through them in. The
///  budget is the dialog's: at most AMaxPerMember rows per member and
///  AMaxTotal in all, and ATotal says how many there really are - a window
///  that silently renames "found" into "listed" is the answer this project
///  does not give.</summary>
function CollectPostEditRefs(const AFiles, AContents, AMembers: TArray<string>;
  const ATargetFile, ATargetClass, ATargetContent: string;
  AMaxPerMember, AMaxTotal: Integer; out ATotal: Integer;
  out ACapped: Boolean): TArray<TPostEditRef>;

/// <summary>Moves AMembers of AOwnerType into ATargetClass of the target
///  unit: declarations into ASection, bodies to the end of the target's
///  implementation, their headers requalified. Both files come back
///  rewritten; nothing is written here. Issues name what the move does NOT
///  do - the members of the old class a body still uses above all, because
///  that code will not compile in the new class.</summary>
function PlanMethodMove(const ASourceLines, ATargetLines: TArray<string>;
  const AOwnerType: string; const AMembers: TArray<string>;
  const ATargetClass, ASection: string): TMethodMovePlan;

/// <summary>ADeclLine with the directives in AAdd added and those in ARemove
///  taken out ('virtual', 'overload', 'inline', 'static', ...). The
///  declaration keeps its indentation, its name and the directives nobody
///  asked about; a directive already there is not added twice.</summary>
function ApplyModifiersToDecl(const ADeclLine: string;
  const AAdd, ARemove: TArray<string>): string;

/// <summary>The members of AOwnerType that the body ALines[AFirst..ALast]
///  uses - bare or through Self. They stay in the old class, so they are
///  exactly what a reader has to deal with after the move.</summary>
function OwnMembersUsedBy(const ALines: TArray<string>; AFirst, ALast: Integer;
  const AOwnerType: string): TArray<string>;

/// <summary>One issue record - the dialog and the wizard build them too.</summary>
function MethodEditIssue(AKind: TMethodEditIssueKind;
  const AMember, AText: string): TMethodEditIssue;

/// <summary>The class / record / object types the INTERFACE section declares
///  at top level - the candidates for a move target.</summary>
function ClassNamesOf(const ALines: TArray<string>): TArray<string>;

/// <summary>Is ATYPE declared with the 'class' keyword (not a record, not a
///  plain object)? A record IS a legitimate target, so this only picks the
///  default the dialog preselects.</summary>
function IsClassType(const ALines: TArray<string>; const AType: string): Boolean;

/// <summary>ALINES with AADD / AREMOVE applied to the declarations of
///  AMEMBERS inside ATYPE. Only those declaration lines change.</summary>
function ApplyModifiersInClass(const ALines: TArray<string>; const AType: string;
  const AMembers, AAdd, ARemove: TArray<string>): TArray<string>;

/// <summary>ALINES with the signature of ATYPE.AMEMBER rewritten - the
///  declaration inside the class AND its implementation header, so the two
///  cannot drift apart. Directives are kept. The CALL SITES are deliberately
///  not touched: text cannot tell a call from a method reference, and the
///  verified path for that is "Change signature...". The caller says so.
///  '' in AError and an unchanged result means there was nothing to do.</summary>
function ApplySignatureInClass(const ALines: TArray<string>;
  const AType, AMember: string; const AParams: TArray<TSigParam>;
  const AResultType: string; out AError: string): TArray<string>;

implementation

uses
  System.Classes, System.StrUtils, System.Math, System.Generics.Collections,
  Expert.ReferenceKind,
  Expert.PascalScanner, Expert.UnitIndex, Expert.AutoImport,
  Expert.SafeDeletePlan, Expert.SignatureCheck, Expert.UsesEditor;

{ ---- Edit methods: a member moves into another class ---------------------- }

// Every whole-word occurrence of ANAMES in AMASKED - one pass over the file
// for all of them, which is what makes a class with thirty methods
// affordable. AOut[i] belongs to ANames[AWhich[i]].
procedure CollectNameHits(const AMasked: TArray<string>;
  const ANames: TArray<string>; out AOut: TArray<TPoint>;
  out AWhich: TArray<Integer>);
var
  Hits: TList<TPoint>;
  Which: TList<Integer>;
  L, P, N: Integer;
  Line, Up: string;
  Ups: TArray<string>;
begin
  AOut := nil;
  AWhich := nil;
  if Length(ANames) = 0 then Exit;
  SetLength(Ups, Length(ANames));
  for N := 0 to High(ANames) do Ups[N] := UpperCase(ANames[N]);
  Hits := TList<TPoint>.Create;
  Which := TList<Integer>.Create;
  try
    for L := 0 to High(AMasked) do
    begin
      Line := AMasked[L];
      if Line = '' then Continue;
      Up := UpperCase(Line);
      for N := 0 to High(Ups) do
      begin
        P := 1;
        while True do
        begin
          P := PosEx(Ups[N], Up, P);
          if P = 0 then Break;
          if ((P = 1) or not IsIdentChar(Line[P - 1])) and
             ((P + Length(Ups[N]) > Length(Line)) or
              not IsIdentChar(Line[P + Length(Ups[N])])) then
          begin
            Hits.Add(Point(P - 1, L));
            Which.Add(N);
          end;
          Inc(P, Length(Ups[N]));
        end;
      end;
    end;
    AOut := Hits.ToArray;
    AWhich := Which.ToArray;
  finally
    Which.Free;
    Hits.Free;
  end;
end;

function CollectPostEditRefs(const AFiles, AContents, AMembers: TArray<string>;
  const ATargetFile, ATargetClass, ATargetContent: string;
  AMaxPerMember, AMaxTotal: Integer; out ATotal: Integer;
  out ACapped: Boolean): TArray<TPostEditRef>;
var
  Found, Kept: TArray<Integer>;
  TargetLines: TArray<string>;
begin
  Result := nil;
  ATotal := 0;
  ACapped := False;
  if (Length(AMembers) = 0) or (Length(AFiles) = 0) then Exit;
  SetLength(Found, Length(AMembers));
  SetLength(Kept, Length(AMembers));
  TargetLines := SplitContentLines(ATargetContent);

  for var FI := 0 to High(AFiles) do
  begin
    if FI > High(AContents) then Break;
    if AContents[FI] = '' then Continue;
    var FLines := SplitContentLines(AContents[FI]);
    var Masked := MaskCommentsAndStrings(FLines);
    var Hits: TArray<TPoint>;
    var Which: TArray<Integer>;
    CollectNameHits(Masked, AMembers, Hits, Which);
    if Length(Hits) = 0 then Continue;
    var IsTarget := (ATargetFile <> '') and
      SameText(ExpandFileName(AFiles[FI]), ExpandFileName(ATargetFile));
    // One classification pass per NAME, so the kinds belong to the symbol
    // whose declaration decided them.
    for var N := 0 to High(AMembers) do
    begin
      var Pos0: TArray<TPoint> := nil;
      for var H := 0 to High(Hits) do
        if Which[H] = N then Pos0 := Pos0 + [Hits[H]];
      if Length(Pos0) = 0 then Continue;
      // The symbol's kind comes from where the member lives NOW.
      var SK := rsUnknown;
      if ATargetClass <> '' then
      begin
        var DL := FindMemberDeclarationLine(ATargetContent, ATargetClass,
          AMembers[N]);
        if (DL >= 0) and (DL <= High(TargetLines)) then
          SK := SymbolKindFromDeclLine(TargetLines[DL], AMembers[N]);
      end;
      var RK := ClassifyReferences(AContents[FI], Pos0, Length(AMembers[N]), SK);
      for var H := 0 to High(Pos0) do
      begin
        var KindText := 'use';
        if H <= High(RK) then KindText := RefKindText(RK[H]);
        // The new declaration and the new body are the edit's own result.
        if IsTarget and (SameText(KindText, 'Declaration') or
           SameText(KindText, 'Implementation')) then
          Continue;
        Inc(Found[N]);
        Inc(ATotal);
        if ((AMaxTotal > 0) and (Length(Result) >= AMaxTotal)) or
           ((AMaxPerMember > 0) and (Kept[N] >= AMaxPerMember)) then
        begin
          ACapped := True;
          Continue;
        end;
        Inc(Kept[N]);
        var R := Default(TPostEditRef);
        R.FilePath := AFiles[FI];
        R.Line := Pos0[H].Y;
        R.Col := Pos0[H].X;
        R.Member := AMembers[N];
        R.Kind := KindText;
        if Pos0[H].Y <= High(FLines) then R.Text := Trim(FLines[Pos0[H].Y]);
        Result := Result + [R];
      end;
    end;
  end;

  // By file, then by line - the order someone works through them in. An
  // insertion sort: this is a handful of members' occurrences, bounded by
  // the budget above.
  for var I := 1 to High(Result) do
  begin
    var Cur := Result[I];
    var J := I - 1;
    while J >= 0 do
    begin
      var Cmp := CompareText(Result[J].FilePath, Cur.FilePath);
      if Cmp = 0 then Cmp := Result[J].Line - Cur.Line;
      if Cmp = 0 then Cmp := Result[J].Col - Cur.Col;
      if Cmp <= 0 then Break;
      Result[J + 1] := Result[J];
      Dec(J);
    end;
    Result[J + 1] := Cur;
  end;
end;

function JoinPlannedLines(const ALines: TArray<string>;
  const AOldContent: string): string;
var
  LB: string;
begin
  if Pos(#13#10, AOldContent) > 0 then LB := #13#10
  else if Pos(#10, AOldContent) > 0 then LB := #10
  else LB := sLineBreak;
  Result := string.Join(LB, ALines);
end;

function MembersInLineRange(const ALines: TArray<string>;
  const AMembers: TArray<TClassMemberInfo>; const AOwnerType: string;
  AFrom, ATo: Integer): TArray<string>;
var
  Impl: TArray<string>;   // member names whose BODY the range reaches
begin
  Result := nil;
  if (Length(AMembers) = 0) or (Length(ALines) = 0) then Exit;
  if AFrom > ATo then
  begin
    var Tmp := AFrom; AFrom := ATo; ATo := Tmp;
  end;
  if AFrom < 0 then AFrom := 0;
  if ATo > High(ALines) then ATo := High(ALines);
  if AFrom > ATo then Exit;

  // The bodies the range touches, collected in ONE pass: a routine found at
  // line L covers up to its own last line, so the walk continues behind it
  // instead of asking the same routine again for every line it spans.
  var L := AFrom;
  while L <= ATo do
  begin
    var HF, HL: Integer;
    if not FindEnclosingRoutineRangeIn(ALines, L, HF, HL) then
    begin
      Inc(L);
      Continue;
    end;
    if SameText(OwnerTypeOfImplHeader(ALines[HF]), AOwnerType) then
    begin
      var Kind, Qual, Params, Ret: string;
      var IsCM: Boolean;
      var HdrEnd: Integer;
      var Hdr := CollectHeader(ALines, HF, HdrEnd);
      if (Hdr <> '') and IsHeaderLine(Trim(Hdr), Kind, IsCM) and
         ParseHeader(Hdr, Kind, Qual, Params, Ret) then
      begin
        var D := LastDelimiter('.', Qual);
        if D > 0 then Qual := Copy(Qual, D + 1, MaxInt);
        if Qual <> '' then Impl := Impl + [Qual];
      end;
    end;
    L := Max(HL, L) + 1;
  end;

  for var M in AMembers do
  begin
    var Hit := (M.DeclLine >= AFrom) and (M.DeclLine <= ATo);
    if not Hit then
      for var N in Impl do
        if SameText(N, M.Name) then
        begin
          Hit := True;
          Break;
        end;
    if Hit then Result := Result + [M.Name];
  end;
end;

// The class body of AType: its header line and its own 'end'. Nested
// class / record declarations are counted, so an inner body cannot end the
// outer one (the trap FindMemberDeclarationLine fell into, 1.4.1).
function ClassBodyRange(const ALines: TArray<string>; const AType: string;
  out AFirst, ALast: Integer): Boolean;
var
  L, Depth: Integer;
  T: string;
begin
  Result := False;
  AFirst := -1;
  ALast := -1;
  if AType = '' then Exit;
  for L := 0 to High(ALines) do
  begin
    T := Trim(StripLineComment(ALines[L]));
    if SameText(ClassOpenerName(T), AType) and not T.EndsWith(';') then
    begin
      AFirst := L;
      Break;
    end;
  end;
  if AFirst < 0 then Exit;
  Depth := 1;
  for L := AFirst + 1 to High(ALines) do
  begin
    T := Trim(StripLineComment(ALines[L]));
    if ((ClassOpenerName(T) <> '') or
        (T.ToUpper.Contains('= RECORD') and not T.EndsWith(';'))) and
       not T.EndsWith(';') then
      Inc(Depth)
    else if SameText(T, 'end;') or SameText(T, 'end') then
    begin
      Dec(Depth);
      if Depth = 0 then
      begin
        ALast := L;
        Exit(True);
      end;
    end;
  end;
end;

// The signature up to its own ';' and the directives behind it. The first
// ';' INSIDE the parameter list is not the end of the signature.
procedure SplitSignatureAndDirectives(const AText: string;
  out ASignature, ADirectives: string);
var
  I, Depth: Integer;
begin
  ASignature := AText;
  ADirectives := '';
  Depth := 0;
  for I := 1 to Length(AText) do
  begin
    if CharInSet(AText[I], ['(', '[']) then Inc(Depth)
    else if CharInSet(AText[I], [')', ']']) then Dec(Depth)
    else if (AText[I] = ';') and (Depth <= 0) then
    begin
      ASignature := Copy(AText, 1, I);
      ADirectives := Trim(Copy(AText, I + 1, MaxInt));
      Exit;
    end;
  end;
end;

// The first identifier at or after AFrom (1-based).
function FirstIdentFrom(const S: string; AFrom: Integer): string;
var
  I, J: Integer;
begin
  Result := '';
  I := AFrom;
  while (I <= Length(S)) and not IsIdentChar(S[I]) do Inc(I);
  J := I;
  while (J <= Length(S)) and IsIdentChar(S[J]) do Inc(J);
  if J > I then Result := Copy(S, I, J - I);
end;

// Every name AType declares at its OWN level - methods, properties and
// fields. The warning half needs the fields too ("the body uses FColumns").
function ClassDeclaredNames(const ALines: TArray<string>;
  const AType: string): TArray<string>;
var
  First, Last, L, Depth, HdrEnd, P: Integer;
  T, Kind, Hdr, Qual, Params, Ret, Sig, Dirs: string;
  IsCM: Boolean;
begin
  Result := nil;
  if not ClassBodyRange(ALines, AType, First, Last) then Exit;
  Depth := 0;
  L := First + 1;
  while L < Last do
  begin
    T := Trim(StripLineComment(ALines[L]));
    if T = '' then begin Inc(L); Continue; end;
    if ((ClassOpenerName(T) <> '') or
        (T.ToUpper.Contains('= RECORD') and not T.EndsWith(';'))) and
       not T.EndsWith(';') then
    begin
      Inc(Depth);
      Inc(L);
      Continue;
    end;
    if SameText(T, 'end;') or SameText(T, 'end') then
    begin
      if Depth > 0 then Dec(Depth);
      Inc(L);
      Continue;
    end;
    if Depth > 0 then begin Inc(L); Continue; end;
    if IsHeaderLine(T, Kind, IsCM) then
    begin
      Hdr := CollectHeader(ALines, L, HdrEnd);
      if (Hdr <> '') and ParseHeader(Hdr, Kind, Qual, Params, Ret) then
        Result := Result + [Qual];
      if HdrEnd > L then L := HdrEnd;
    end
    else if T.ToUpper.StartsWith('PROPERTY ') or
            T.ToUpper.StartsWith('CLASS PROPERTY ') then
      Result := Result + [FirstIdentFrom(T, Pos('PROPERTY', T.ToUpper) + 8)]
    else
    begin
      SplitSignatureAndDirectives(T, Sig, Dirs);
      P := Pos(':', Sig);
      if (P > 1) and not T.ToUpper.StartsWith('CASE ') then
        for var N in Copy(Sig, 1, P - 1).Split([',']) do
          if IsIdentifier(Trim(N)) then Result := Result + [Trim(N)];
    end;
    Inc(L);
  end;
end;

function ClassMembersOf(const ALines: TArray<string>;
  const AType: string): TArray<TClassMemberInfo>;
var
  First, Last, L, Depth, HdrEnd: Integer;
  T, Kind, Hdr, Qual, Params, Ret, Sig, Dirs, Why: string;
  IsCM: Boolean;
  Info: TClassMemberInfo;
  Sym: TSafeDeleteSymbol;
begin
  Result := nil;
  if not ClassBodyRange(ALines, AType, First, Last) then Exit;
  Depth := 0;
  L := First + 1;
  while L < Last do
  begin
    T := Trim(StripLineComment(ALines[L]));
    if T = '' then begin Inc(L); Continue; end;
    if ((ClassOpenerName(T) <> '') or
        (T.ToUpper.Contains('= RECORD') and not T.EndsWith(';'))) and
       not T.EndsWith(';') then
    begin
      Inc(Depth);
      Inc(L);
      Continue;
    end;
    if SameText(T, 'end;') or SameText(T, 'end') then
    begin
      if Depth > 0 then Dec(Depth);
      Inc(L);
      Continue;
    end;
    if (Depth = 0) and IsHeaderLine(T, Kind, IsCM) then
    begin
      Hdr := CollectHeader(ALines, L, HdrEnd);
      if (Hdr <> '') and ParseHeader(Hdr, Kind, Qual, Params, Ret) then
      begin
        SplitSignatureAndDirectives(Hdr, Sig, Dirs);
        Info := Default(TClassMemberInfo);
        Info.Name := Qual;
        Info.DeclLine := L;
        Info.Kind := LowerCase(Kind);
        Info.Directives := Dirs;
        // THE SAME RULES SAFE DELETE USES: an overload, a virtual / override /
        // abstract / message member and a published one cannot travel alone.
        Info.Movable := PlanSafeDeleteSymbol(ALines, L, Qual, Sym, Why);
        if not Info.Movable then
          Info.Why := Why
        else if Length(Sym.Vetoes) > 0 then
        begin
          Info.Movable := False;
          Info.Why := string.Join('; ', Sym.Vetoes);
        end
        else if Sym.ImplLine < 0 then
        begin
          Info.Movable := False;
          Info.Why := 'it has no implementation in this unit';
        end;
        Result := Result + [Info];
      end;
      if HdrEnd > L then L := HdrEnd;
    end;
    Inc(L);
  end;
end;

function OwnMembersUsedBy(const ALines: TArray<string>; AFirst, ALast: Integer;
  const AOwnerType: string): TArray<string>;
var
  Names, Masked: TArray<string>;
  L, C, S: Integer;
  Line, Ident, Qual: string;
  Seen: TStringList;
begin
  Result := nil;
  if (AFirst < 0) or (ALast > High(ALines)) or (AFirst > ALast) then Exit;
  Names := ClassDeclaredNames(ALines, AOwnerType);
  if Length(Names) = 0 then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  Seen := TStringList.Create;
  try
    Seen.CaseSensitive := False;
    Seen.Duplicates := dupIgnore;
    Seen.Sorted := True;
    // The header line itself declares the parameters - the body starts below.
    for L := AFirst + 1 to ALast do
    begin
      Line := Masked[L];
      C := 1;
      while C <= Length(Line) do
      begin
        if IsIdentChar(Line[C]) and ((C = 1) or not IsIdentChar(Line[C - 1])) then
        begin
          S := C;
          while (C <= Length(Line)) and IsIdentChar(Line[C]) do Inc(C);
          Ident := Copy(Line, S, C - S);
          for var N in Names do
            if SameText(N, Ident) then
            begin
              // bare or Self-qualified is a member use; A.Name is someone
              // else's (QualifierBefore takes a 0-based column).
              Qual := QualifierBefore(ALines[L], S - 1);
              if (Qual = '') or SameText(Qual, 'Self') then
                if Seen.IndexOf(N) < 0 then
                begin
                  Seen.Add(N);
                  Result := Result + [N];
                end;
              Break;
            end;
        end
        else
          Inc(C);
      end;
    end;
  finally
    Seen.Free;
  end;
end;

function ApplyModifiersToDecl(const ADeclLine: string;
  const AAdd, ARemove: TArray<string>): string;
var
  Code, Sig, Tail, Indent: string;
  Dirs, Keep: TArray<string>;
  Has: Boolean;
begin
  Result := ADeclLine;
  Code := TrimRight(StripLineComment(ADeclLine));
  SplitSignatureAndDirectives(Code, Sig, Tail);
  // No ';' at all means this is not a declaration we may touch. When one was
  // found the signature ends with it - that is the honest test, because a
  // declaration whose ';' is the last character makes Sig = Code.
  if not TrimRight(Sig).EndsWith(';') then Exit;
  Dirs := nil;
  for var D in Tail.Split([';']) do
    if Trim(D) <> '' then Dirs := Dirs + [Trim(D)];
  for var R in ARemove do
  begin
    Keep := nil;
    for var D in Dirs do
      if not SameText(D, Trim(R)) then Keep := Keep + [D];
    Dirs := Keep;
  end;
  for var A in AAdd do
  begin
    if Trim(A) = '' then Continue;
    Has := False;
    for var D in Dirs do
      if SameText(D, Trim(A)) then Has := True;
    if not Has then Dirs := Dirs + [Trim(A)];
  end;
  Indent := Copy(ADeclLine, 1, Length(ADeclLine) - Length(TrimLeft(ADeclLine)));
  Result := Indent + TrimLeft(TrimRight(Sig));
  for var D in Dirs do
    Result := Result + ' ' + D + ';';
end;

// The implementation header of AOld.Member becomes ANew.Member.
function RequalifyHeader(const ALine, AOld, ANew: string): string;
var
  U, Needle: string;
  P: Integer;
begin
  Result := ALine;
  if (AOld = '') or (ANew = '') then Exit;
  U := UpperCase(ALine);
  Needle := UpperCase(AOld) + '.';
  P := Pos(Needle, U);
  while P > 0 do
  begin
    if (P = 1) or not IsIdentChar(ALine[P - 1]) then
      Exit(Copy(ALine, 1, P - 1) + ANew + Copy(ALine, P + Length(AOld), MaxInt));
    P := Pos(Needle, U, P + 1);
  end;
end;

function DeleteLineRange(const ALines: TArray<string>;
  AFirst, ALast: Integer): TArray<string>;
var
  L: Integer;
begin
  Result := nil;
  for L := 0 to High(ALines) do
    if (L < AFirst) or (L > ALast) then Result := Result + [ALines[L]];
end;

function InsertLinesAt(const ALines: TArray<string>; AAt: Integer;
  const ANew: TArray<string>): TArray<string>;
var
  L: Integer;
begin
  Result := nil;
  for L := 0 to High(ALines) do
  begin
    if L = AAt then
      for var N in ANew do Result := Result + [N];
    Result := Result + [ALines[L]];
  end;
  if AAt > High(ALines) then
    for var N in ANew do Result := Result + [N];
end;

// The unit's own name, from its first 'unit X;' line.
function UnitNameOf(const ALines: TArray<string>): string;
var
  T: string;
begin
  Result := '';
  for var L in ALines do
  begin
    T := Trim(StripLineComment(L));
    if T = '' then Continue;
    if T.ToUpper.StartsWith('UNIT ') then
    begin
      Result := Trim(Copy(T, 6, MaxInt));
      if Result.EndsWith(';') then SetLength(Result, Length(Result) - 1);
      Result := Trim(Result);
    end;
    Exit;
  end;
end;

// The type names the INTERFACE section declares at top level. Needed for one
// question only: does what moves still need the unit it came from?
function InterfaceTypeNames(const ALines: TArray<string>): TArray<string>;
var
  L, Depth, P: Integer;
  T, U, Nm: string;
  InType: Boolean;
begin
  Result := nil;
  InType := False;
  Depth := 0;
  for L := 0 to High(ALines) do
  begin
    T := Trim(StripLineComment(ALines[L]));
    U := T.ToUpper;
    if U = 'IMPLEMENTATION' then Break;
    if T = '' then Continue;
    if Depth = 0 then
    begin
      if (U = 'TYPE') or U.StartsWith('TYPE ') then InType := True
      else if (U = 'VAR') or (U = 'CONST') or (U = 'RESOURCESTRING') or
              (U = 'INTERFACE') or U.StartsWith('FUNCTION ') or
              U.StartsWith('PROCEDURE ') then InType := False;
    end;
    if ((ClassOpenerName(T) <> '') or
        (U.Contains('= RECORD') and not T.EndsWith(';'))) and
       not T.EndsWith(';') then
    begin
      Inc(Depth);
      if Depth = 1 then
      begin
        Nm := ClassOpenerName(T);
        if Nm = '' then Nm := Trim(Copy(T, 1, Pos('=', T) - 1));
        if IsIdentifier(Nm) then Result := Result + [Nm];
      end;
      Continue;
    end;
    if SameText(T, 'end;') or SameText(T, 'end') then
    begin
      if Depth > 0 then Dec(Depth);
      Continue;
    end;
    if (Depth = 0) and InType then
    begin
      P := Pos('=', T);
      if P > 1 then
      begin
        Nm := Trim(Copy(T, 1, P - 1));
        P := Pos('<', Nm);                 // a generic keeps its bare name
        if P > 1 then Nm := Trim(Copy(Nm, 1, P - 1));
        if IsIdentifier(Nm) then Result := Result + [Nm];
      end;
    end;
  end;
end;

// Does ANAMES occur as CODE in those lines (whole word, comments and strings
// blanked)? The question behind both uses decisions.
function AnyNameOccursIn(const ALines: TArray<string>;
  const ANames: TArray<string>): string;
var
  Masked: TArray<string>;
begin
  Result := '';
  if (Length(ALines) = 0) or (Length(ANames) = 0) then Exit;
  Masked := MaskCommentsAndStrings(ALines);
  for var L in Masked do
    for var N in ANames do
      if (N <> '') and HasWholeWordCI(L, N) then Exit(N);
end;

function PlanMethodMove(const ASourceLines, ATargetLines: TArray<string>;
  const AOwnerType: string; const AMembers: TArray<string>;
  const ATargetClass, ASection: string): TMethodMovePlan;
var
  Plan: TMethodMovePlan;

  procedure AddIssue(AKind: TMethodEditIssueKind; const AMember, AText: string);
  var
    Iss: TMethodEditIssue;
  begin
    Iss.Kind := AKind;
    Iss.Member := AMember;
    Iss.Text := AText;
    Plan.Issues := Plan.Issues + [Iss];
  end;

var
  SrcContent, Text, Body: string;
  TFirst, TLast, ImplIns, DeclLine, InsLine, FirstDecl, DF, DL, IF_, IL, L: Integer;
  Sym: TSafeDeleteSymbol;
  Why: string;
  Decls, BodyBlock, DeclText, Src, Tgt, SrcTypes: TArray<string>;
  Dels: TArray<TSafeDeleteEdit>;
  SrcUnit, TgtUnit: string;
  Probe: TArray<string>;
  SameUnit, NeedIntf, NeedImpl: Boolean;
  Sect: TUsesSection;
  Own: TArray<string>;
  E: TSafeDeleteEdit;
begin
  Plan := Default(TMethodMovePlan);
  Plan.SourceLines := ASourceLines;
  Plan.TargetLines := ATargetLines;
  if (AOwnerType = '') or (ATargetClass = '') or (Length(AMembers) = 0) then
  begin
    Plan.Error := 'nothing to move';
    Exit(Plan);
  end;
  // THE TARGET CLASS MAY LIVE IN THE SAME UNIT - the most ordinary case of
  // all, and the one where two independent rewrites of one file would lose
  // half the work. Everything then happens in ONE array: first the removal,
  // then the insertion into its RESULT, whose positions have moved.
  SameUnit := (UnitNameOf(ASourceLines) <> '') and
    SameText(UnitNameOf(ASourceLines), UnitNameOf(ATargetLines));
  if SameUnit then Probe := ASourceLines else Probe := ATargetLines;
  if not ClassBodyRange(Probe, ATargetClass, TFirst, TLast) then
  begin
    Plan.Error := Format('the target unit does not declare %s', [ATargetClass]);
    Exit(Plan);
  end;
  if SameUnit and SameText(ATargetClass, AOwnerType) then
  begin
    Plan.Error := 'source and target class are the same';
    Exit(Plan);
  end;
  ImplIns := ImplInsertLine(Probe);
  if ImplIns < 0 then
  begin
    Plan.Error := 'the target unit has no implementation section';
    Exit(Plan);
  end;
  SrcContent := string.Join(#13#10, ASourceLines);
  Decls := nil;
  BodyBlock := nil;
  Dels := nil;
  for var M in AMembers do
  begin
    DeclLine := FindMemberDeclarationLine(SrcContent, AOwnerType, M);
    if DeclLine < 0 then
    begin
      AddIssue(meiVeto, M, Format('%s does not declare it', [AOwnerType]));
      Continue;
    end;
    if not PlanSafeDeleteSymbol(ASourceLines, DeclLine, M, Sym, Why) then
    begin
      AddIssue(meiVeto, M, Why);
      Continue;
    end;
    if Length(Sym.Vetoes) > 0 then
    begin
      for var V in Sym.Vetoes do AddIssue(meiVeto, M, V);
      Continue;
    end;
    if Sym.ImplLine < 0 then
    begin
      AddIssue(meiVeto, M, 'it has no implementation in this unit');
      Continue;
    end;
    DF := -1; DL := -1; IF_ := -1; IL := -1;
    for E in Sym.Edits do
      if E.FirstLine = Sym.DeclLine then
      begin
        DF := E.FirstLine;
        DL := E.LastLine;
      end
      else if E.FirstLine = Sym.ImplLine then
      begin
        IF_ := E.FirstLine;
        IL := E.LastLine;
      end;
    if (DF < 0) or (IF_ < 0) then
    begin
      AddIssue(meiVeto, M, 'declaration or body could not be delimited');
      Continue;
    end;
    for L := DF to DL do Decls := Decls + [Trim(ASourceLines[L])];
    // ONE blank line between the bodies. The body's delete range INCLUDES
    // the blank line that follows it (safe delete takes it so the source
    // does not keep a double blank), so a copied block already ends with
    // one - adding a separator on top of that is what gave the user TWO
    // blank lines between five moved methods. So: trim what the range
    // brought and put exactly one blank between blocks; the blank the
    // insertion point needs is added there.
    var Blk: TArray<string> := nil;
    for L := IF_ to IL do
      if L = IF_ then
        Blk := Blk + [RequalifyHeader(ASourceLines[L], AOwnerType, ATargetClass)]
      else
        Blk := Blk + [ASourceLines[L]];
    while (Length(Blk) > 0) and (Trim(Blk[High(Blk)]) = '') do
      SetLength(Blk, Length(Blk) - 1);
    if Length(BodyBlock) > 0 then BodyBlock := BodyBlock + [''];
    BodyBlock := BodyBlock + Blk;
    // WHAT STAYS BEHIND - the honest half of this feature.
    Own := OwnMembersUsedBy(ASourceLines, IF_, IL, AOwnerType);
    if Length(Own) > 0 then
      AddIssue(meiOwnMember, M, Format('the body uses %s of %s - they stay there',
        [string.Join(', ', Own), AOwnerType]));
    Body := '';
    for L := IF_ to IL do Body := Body + ' ' + ASourceLines[L];
    if HasWholeWordCI(Body, 'inherited') then
      AddIssue(meiNote, M, 'the body calls "inherited" - in another class that ' +
        'reaches another ancestor');
    Dels := Dels + [Sym.Edits[0]];
    for E in Sym.Edits do
      if E.FirstLine <> Sym.Edits[0].FirstLine then Dels := Dels + [E];
    Plan.Moved := Plan.Moved + [M];
  end;
  if Length(Plan.Moved) = 0 then
  begin
    Plan.Error := 'none of the selected members can be moved';
    Exit(Plan);
  end;
  // SOURCE: the ranges go bottom-up, so every position stays valid.
  Src := ASourceLines;
  for var I := 0 to High(Dels) do
    for var J := 0 to High(Dels) - I - 1 do
      if Dels[J].FirstLine < Dels[J + 1].FirstLine then
      begin
        E := Dels[J];
        Dels[J] := Dels[J + 1];
        Dels[J + 1] := E;
      end;
  for E in Dels do
    if E.HasReplacement then
      Src[E.FirstLine] := E.Replacement
    else
      Src := DeleteLineRange(Src, E.FirstLine, E.LastLine);
  Plan.SourceLines := Src;
  // TARGET: the bodies sit below the class, so they go in first. In the same
  // unit the class and the implementation have MOVED - they are located again
  // in the already shortened text, never in the original.
  if SameUnit then
  begin
    Tgt := Src;
    if not ClassBodyRange(Tgt, ATargetClass, TFirst, TLast) then
    begin
      Plan.Error := Format('%s was lost while removing the members', [ATargetClass]);
      Exit(Plan);
    end;
    ImplIns := ImplInsertLine(Tgt);
  end
  else
    Tgt := ATargetLines;
  if not PlanMemberInsertion(Tgt, TFirst, ASection, Decls, InsLine,
    Text, FirstDecl) then
  begin
    Plan.Error := 'the target class body could not be delimited';
    Exit(Plan);
  end;
  // The insertion point is the line of the final 'end.' (or of whatever
  // follows the implementation), so the last body needs a blank line under
  // it - the per-member loop above deliberately trimmed the one the delete
  // range brought along.
  if (Length(BodyBlock) > 0) and (Trim(BodyBlock[High(BodyBlock)]) <> '') then
    BodyBlock := BodyBlock + [''];
  // One blank line separates the new body from the one above - but only
  // when there is not already one there, or every move leaves a double
  // blank behind (seen in the first live apply).
  if (ImplIns > 0) and (ImplIns - 1 <= High(Tgt)) and (Trim(Tgt[ImplIns - 1]) = '') then
    Tgt := InsertLinesAt(Tgt, ImplIns, BodyBlock)
  else
    Tgt := InsertLinesAt(Tgt, ImplIns, [''] + BodyBlock);
  DeclText := SplitContentLines(Text);
  while (Length(DeclText) > 0) and (DeclText[High(DeclText)] = '') do
    SetLength(DeclText, Length(DeclText) - 1);
  Tgt := InsertLinesAt(Tgt, InsLine, DeclText);
  Plan.TargetLines := Tgt;
  if SameUnit then
  begin
    // One file, one result - and no uses clause can be involved.
    Plan.SourceLines := Tgt;
    Plan.Ok := True;
    Exit(Plan);
  end;

  // ---- the two uses clauses ------------------------------------------------
  // Deliberately conservative in BOTH directions: a unit too many is a hint in
  // the preview, a unit too few does not compile.
  SrcUnit := UnitNameOf(ASourceLines);
  TgtUnit := UnitNameOf(ATargetLines);
  // (1) the source may still call what moved - then it needs the target unit.
  if (TgtUnit <> '') and
     ((AnyNameOccursIn(Plan.SourceLines, Plan.Moved) <> '') or
      (AnyNameOccursIn(Plan.SourceLines, [ATargetClass]) <> '')) and
     PlanAddUnitToUsesText(string.Join(#13#10, Plan.SourceLines), TgtUnit,
       usImplementation, Text) then
  begin
    Plan.SourceLines := SplitContentLines(Text);
    Plan.SourceUsesAdded := TgtUnit;
  end;
  // (2) what moved may still need the unit it came from. A type in the
  // DECLARATION means the target's INTERFACE needs it - and that is the only
  // direction that can close a circle, so it is said out loud.
  if SrcUnit <> '' then
  begin
    SrcTypes := InterfaceTypeNames(ASourceLines);
    NeedIntf := (AnyNameOccursIn(Decls, SrcTypes) <> '') or
                (AnyNameOccursIn(Decls, [AOwnerType]) <> '');
    NeedImpl := NeedIntf or (AnyNameOccursIn(BodyBlock, SrcTypes) <> '') or
                (AnyNameOccursIn(BodyBlock, [AOwnerType]) <> '');
    if NeedImpl then
    begin
      if NeedIntf then Sect := usInterface else Sect := usImplementation;
      if PlanAddUnitToUsesText(string.Join(#13#10, Plan.TargetLines), SrcUnit,
        Sect, Text) then
      begin
        Plan.TargetLines := SplitContentLines(Text);
        Plan.TargetUsesAdded := SrcUnit;
        Plan.TargetUsesInInterface := NeedIntf;
      end;
      if NeedIntf then
        AddIssue(meiNote, '', Format('%s now needs %s in its INTERFACE uses ' +
          '(a declaration names a type of it) - check that %s does not use %s ' +
          'in its own interface, or the units close a circle',
          [TgtUnit, SrcUnit, SrcUnit, TgtUnit]));
    end;
  end;
  Plan.Ok := True;
  Result := Plan;
end;

function MethodEditIssue(AKind: TMethodEditIssueKind;
  const AMember, AText: string): TMethodEditIssue;
begin
  Result.Kind := AKind;
  Result.Member := AMember;
  Result.Text := AText;
end;

function ClassNamesOf(const ALines: TArray<string>): TArray<string>;
var
  L, Depth: Integer;
  T, U, Nm: string;
begin
  Result := nil;
  Depth := 0;
  for L := 0 to High(ALines) do
  begin
    T := Trim(StripLineComment(ALines[L]));
    U := T.ToUpper;
    if U = 'IMPLEMENTATION' then Break;
    if T = '' then Continue;
    Nm := ClassOpenerName(T);
    if ((Nm <> '') or (U.Contains('= RECORD') and not T.EndsWith(';'))) and
       not T.EndsWith(';') then
    begin
      Inc(Depth);
      // Only a TOP-LEVEL type is a target: a nested one cannot be named
      // from outside without its outer qualifier, and the move writes a
      // bare 'TTarget.Member' header.
      if Depth = 1 then
      begin
        if Nm = '' then Nm := Trim(Copy(T, 1, Pos('=', T) - 1));
        if IsIdentifier(Nm) then Result := Result + [Nm];
      end;
      Continue;
    end;
    if (SameText(T, 'end;') or SameText(T, 'end')) and (Depth > 0) then
      Dec(Depth);
  end;
end;

function IsClassType(const ALines: TArray<string>; const AType: string): Boolean;
var
  First, Last, P: Integer;
  T: string;
begin
  Result := False;
  if not ClassBodyRange(ALines, AType, First, Last) then Exit;
  T := Trim(StripLineComment(ALines[First]));
  P := Pos('=', T);
  if P <= 0 then Exit;
  T := Trim(Copy(T, P + 1, MaxInt));
  Result := SameText(Copy(T, 1, 5), 'class') or
            SameText(Copy(T, 1, 9), 'interface');
end;

function ApplyModifiersInClass(const ALines: TArray<string>; const AType: string;
  const AMembers, AAdd, ARemove: TArray<string>): TArray<string>;
var
  Members: TArray<TClassMemberInfo>;
  New: string;
begin
  Result := ALines;
  if (Length(AAdd) = 0) and (Length(ARemove) = 0) then Exit;
  Members := ClassMembersOf(ALines, AType);
  for var M in Members do
  begin
    var Wanted := False;
    for var N in AMembers do
      if SameText(N, M.Name) then Wanted := True;
    if not Wanted then Continue;
    if (M.DeclLine < 0) or (M.DeclLine > High(Result)) then Continue;
    New := ApplyModifiersToDecl(Result[M.DeclLine], AAdd, ARemove);
    Result[M.DeclLine] := New;
  end;
end;

function ApplySignatureInClass(const ALines: TArray<string>;
  const AType, AMember: string; const AParams: TArray<TSigParam>;
  const AResultType: string; out AError: string): TArray<string>;
var
  Members: TArray<TClassMemberInfo>;
  Info: TClassMemberInfo;
  Sym: TSafeDeleteSymbol;
  Why, Sig, Dirs, Kind, Qual, Params, Ret, NewSig: string;
  Hdr: string;
  HdrEnd: Integer;
  IsCM: Boolean;
  Tmp: TArray<string>;
begin
  Result := ALines;
  AError := '';
  Members := ClassMembersOf(ALines, AType);
  Info := Default(TClassMemberInfo);
  Info.DeclLine := -1;
  for var M in Members do
    if SameText(M.Name, AMember) then Info := M;
  if Info.DeclLine < 0 then
  begin
    AError := Format('%s.%s was not found', [AType, AMember]);
    Exit;
  end;
  if (Info.Kind = 'property') or (Info.Kind = 'field') then
  begin
    AError := 'only a method has a parameter list';
    Exit;
  end;
  // The kind and the result type come from the DECLARATION, so a rewrite
  // cannot turn a procedure into a function by accident.
  Hdr := CollectHeader(ALines, Info.DeclLine, HdrEnd);
  // ParseHeader needs the keyword as INPUT - the line itself names it, and
  // taking it from there keeps its spelling.
  if (Hdr = '') or
     not IsHeaderLine(Trim(StripLineComment(ALines[Info.DeclLine])), Kind, IsCM) or
     not ParseHeader(Hdr, Kind, Qual, Params, Ret) then
  begin
    AError := 'the declaration could not be parsed';
    Exit;
  end;
  SplitSignatureAndDirectives(Hdr, Sig, Dirs);
  Ret := Trim(AResultType);
  if (Ret = '') and not SameText(Kind, 'procedure') and
     not SameText(Kind, 'constructor') and not SameText(Kind, 'destructor') then
  begin
    AError := 'a function needs a result type';
    Exit;
  end;
  NewSig := Kind + ' ' + AMember;
  if Length(AParams) > 0 then NewSig := NewSig + '(' + FormatParamList(AParams) + ')';
  if Ret <> '' then NewSig := NewSig + ': ' + Ret;

  // The implementation first: its line number is still the one the symbol
  // reported, and a declaration rewrite can move it.
  if not PlanSafeDeleteSymbol(ALines, Info.DeclLine, AMember, Sym, Why) then
    Sym.ImplLine := -1;
  if Sym.ImplLine >= 0 then
  begin
    if not TSignatureChecker.ReplaceSignature(Result, Sym.ImplLine,
      Kind + ' ' + AType + '.' + AMember +
      IfThen(Length(AParams) > 0, '(' + FormatParamList(AParams, False) + ')', '') +
      IfThen(Ret <> '', ': ' + Ret, ''), Tmp) then
    begin
      AError := Format('the implementation header of %s.%s could not be rewritten ' +
        '(a comment inside it would be lost)', [AType, AMember]);
      Exit;
    end;
    Result := Tmp;
  end
  else
    AError := Format('%s.%s has no implementation in this unit - only its ' +
      'declaration was changed', [AType, AMember]);

  if not TSignatureChecker.ReplaceSignature(Result, Info.DeclLine, NewSig, Tmp) then
  begin
    AError := 'the declaration could not be rewritten (a comment inside it ' +
      'would be lost)';
    Exit(ALines);
  end;
  Result := Tmp;
end;

{ TMethodEditRequest }

function TMethodEditRequest.WantsSomething: Boolean;
begin
  Result := (TargetClass <> '') or (Length(AddModifiers) > 0) or
    (Length(RemoveModifiers) > 0) or (SigMember <> '');
end;

{ TMethodEditResult }

function TMethodEditResult.Summary: string;
begin
  if not Ok then Exit('Nothing will be written: ' + Error);
  if Length(Moved) > 0 then
    Result := Format('%s moves into %s.', [string.Join(', ', Moved),
      IfThen(SameFile, 'the other class', ExtractFileName(TargetFile))])
  else
    Result := 'The declarations are changed where they are.';
  if SourceUsesAdded <> '' then
    Result := Result + sLineBreak + Format('%s gains "%s" in its implementation uses.',
      [ExtractFileName(SourceFile), SourceUsesAdded]);
  if TargetUsesAdded <> '' then
    Result := Result + sLineBreak + Format('%s gains "%s" in its uses.',
      [ExtractFileName(TargetFile), TargetUsesAdded]);
  for var I in Issues do
    Result := Result + sLineBreak + IfThen(I.Member <> '', I.Member + ': ', '') + I.Text;
end;

function PlanMethodEdit(const ASourceFile, ASourceContent, AOwnerType: string;
  const AReq: TMethodEditRequest; const ATargetContent: string): TMethodEditResult;
var
  Res: TMethodEditResult;
  Src, Tgt: TArray<string>;
  Plan: TMethodMovePlan;
  Err: string;
begin
  Res := Default(TMethodEditResult);
  Res.SourceFile := ASourceFile;
  if AOwnerType = '' then
  begin
    Res.Error := 'no class at the caret';
    Exit(Res);
  end;
  if Length(AReq.Members) = 0 then
  begin
    Res.Error := 'no member is selected';
    Exit(Res);
  end;
  if not AReq.WantsSomething then
  begin
    Res.Error := 'nothing is changed yet - pick a target class, a modifier ' +
      'or edit the parameter list';
    Exit(Res);
  end;

  Src := SplitContentLines(ASourceContent);

  // 1. THE SIGNATURE comes first: a moved member should travel with the list
  //    it is supposed to have, not with the old one.
  if AReq.SigMember <> '' then
  begin
    Src := ApplySignatureInClass(Src, AOwnerType, AReq.SigMember,
      AReq.SigParams, AReq.SigResultType, Err);
    if Err <> '' then
    begin
      // "no implementation in this unit" is a note, anything else stops.
      if not ContainsText(Err, 'only its declaration') then
      begin
        Res.Error := Err;
        Exit(Res);
      end;
      Res.Issues := Res.Issues + [MethodEditIssue(meiNote, AReq.SigMember, Err)];
    end;
    Res.Issues := Res.Issues + [MethodEditIssue(meiNote, AReq.SigMember,
      'the call sites are NOT changed - use "Change signature..." for the ' +
      'verified rewrite')];
  end;

  if Length(AReq.AddModifiers) > 0 then
    Res.Issues := Res.Issues + [MethodEditIssue(meiNote, '',
      Format('adds "%s" to %d declaration(s)',
        [string.Join(' ', AReq.AddModifiers), Length(AReq.Members)]))];
  if Length(AReq.RemoveModifiers) > 0 then
    Res.Issues := Res.Issues + [MethodEditIssue(meiNote, '',
      Format('takes "%s" out of %d declaration(s)',
        [string.Join(' ', AReq.RemoveModifiers), Length(AReq.Members)]))];

  // 2. THE MOVE.
  if AReq.TargetClass = '' then
  begin
    Res.SourceLines := ApplyModifiersInClass(Src, AOwnerType, AReq.Members,
      AReq.AddModifiers, AReq.RemoveModifiers);
    Res.Ok := True;
    Exit(Res);
  end;

  Res.SameFile := (AReq.TargetFile = '') or
    SameText(ExpandFileName(AReq.TargetFile), ExpandFileName(ASourceFile));
  if Res.SameFile then Tgt := Src else Tgt := SplitContentLines(ATargetContent);
  if Length(Tgt) = 0 then
  begin
    Res.Error := Format('%s could not be read', [ExtractFileName(AReq.TargetFile)]);
    Exit(Res);
  end;

  Plan := PlanMethodMove(Src, Tgt, AOwnerType, AReq.Members, AReq.TargetClass,
    IfThen(AReq.Section <> '', AReq.Section, 'private'));
  Res.Issues := Res.Issues + Plan.Issues;
  if not Plan.Ok then
  begin
    Res.Error := Plan.Error;
    Exit(Res);
  end;
  Res.Moved := Plan.Moved;
  Res.SourceUsesAdded := Plan.SourceUsesAdded;
  Res.TargetUsesAdded := Plan.TargetUsesAdded;

  // 3. THE MODIFIERS, on the declarations where they now LIVE - which after
  //    a move is the target class.
  Res.TargetLines := ApplyModifiersInClass(Plan.TargetLines, AReq.TargetClass,
    AReq.Members, AReq.AddModifiers, AReq.RemoveModifiers);
  if Res.SameFile then
  begin
    Res.SourceLines := Res.TargetLines;
    Res.TargetFile := '';
  end
  else
  begin
    Res.SourceLines := Plan.SourceLines;
    Res.TargetFile := AReq.TargetFile;
  end;
  Res.Ok := True;
  Result := Res;
end;

end.
