(*
 * Copyright (c) 2026 Sebastian Jaenicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.McpTools;

// The pure half of the MCP tools (no ToolsAPI): turning diagnostics and
// quick fixes into the JSON the bridge hands to the model, and the fix ids.
// Expert.McpServer does the IDE half (buffers, main thread, applying).
//
// FIX IDS carry the content hash of the buffer they were computed for:
// "<hash hex>-<index>". Applying checks the hash against the CURRENT
// buffer first, so a fix listed before an edit can never be applied to a
// buffer it no longer describes - the index alone would silently hit the
// wrong line.

interface

uses
  System.SysUtils, System.JSON, System.Generics.Collections,
  Lsp.Protocol, Expert.AutoImport;

type
  /// <summary>'in_project' / 'search_path' / 'dcu' / 'browsing_only' /
  ///  'unknown' for a candidate unit of an add-unit fix.</summary>
  TUnitAvailabilityFunc = reference to function(const AUnit: string): string;

function SeverityName(ASeverity: Integer): string;
function FixKindName(AKind: TQuickFixKind): string;
function FixDescription(const AFix: TQuickFix): string;

function MakeFixId(AHash: Cardinal; AIndex: Integer): string;
function ParseFixId(const AId: string; out AHash: Cardinal; out AIndex: Integer): Boolean;

type
  /// <summary>One changed line of a PREVIEW. Line is 1-based in the OLD
  ///  content; Before = '' marks a line the edit inserts there, After = ''
  ///  a line it deletes. Every writing MCP tool answers in this shape, so
  ///  a caller can judge (or reproduce) an edit without applying it
  ///  (user request 2026-09-24).</summary>
  TPreviewChange = record
    FilePath: string;
    Line: Integer;
    Before: string;
    After: string;
  end;

/// <summary>The line changes between AOld and ANew: common prefix and
///  suffix are skipped, the rest is paired up. At most AMax changes are
///  returned, ATotal says how many there are.</summary>
function DiffToChanges(const AFile, AOld, ANew: string; AMax: Integer;
  out ATotal: Integer): TArray<TPreviewChange>;
function ChangesToJson(const AChanges: TArray<TPreviewChange>): TJSONArray;

/// <summary>FNV-1a over the content - the revision a preview was computed
///  for (the same shape the fix ids use).</summary>
function PreviewContentHash(const AContent: string): Cardinal;

/// <summary>Is AName something that CANNOT be a Delphi unit name - empty, or
///  anything but a dotted identifier?
///  WHY A TOOL CHECKS ITS OWN ANSWER: analyze_uses reported
///  'Expert.PluginSettings {$IFNDEF STANDALONE_BUILD}' as a unit name for
///  two days (1.16.2 fixed the parser). The answer was wrong in a way that
///  is structurally impossible, I had read that very clause in the same
///  session without noticing, and it took an outside audit to find it. A
///  result like that is not data, it is a defect in whoever produced it - so
///  the tools say so instead of passing it on. The model on the other end
///  takes a tool answer for truth; this is the one place that can doubt
///  it.</summary>
function ImplausibleUnitName(const AName: string): Boolean;

/// <summary>'' when every name is plausible, otherwise the sentence a tool
///  puts into its answer as "selfCheck".</summary>
function UnitNameSelfCheck(const ANames: TArray<string>): string;

/// <summary>Remembers which files a preview describes and how they looked.
///  "apply" with that token refuses when a buffer changed meanwhile - the
///  caller then sees a stale preview instead of an edit it never saw.
///  AReadContent answers '' for a file that cannot be read.</summary>
/// <summary>A canonical fingerprint of a call's ARGUMENTS: every pair
///  except 'apply', 'token' and 'instance' (which say nothing about WHAT
///  is changed), keys sorted so the order a client sends them in does not
///  matter. The preview token carries it, so an apply with the same token
///  but OTHER arguments is refused (audit #36, H34) - a preview of
///  "add_unit Foo" used to hand out a token that "add_unit apply=true
///  Bar" accepted, and the answer then described Bar while the user had
///  reviewed Foo.</summary>
function PreviewArgsFingerprint(AArgs: TJSONObject): string;

/// <summary>1-based position of AEXPR in ALINE as CODE: not inside a
///  comment or a string literal, and - when AEXPR starts and ends with an
///  identifier character - on identifier boundaries, so "Idx" does not
///  match inside "MaxIdx". Searching from AFROMCOL when that is >= 1,
///  else from the start; 0 when there is no such occurrence. The MCP
///  extract_variable used a plain Pos() and happily picked the "Count"
///  inside 'Count: ' (audit #39, L7d).</summary>
function CodeOccurrenceOf(const ALine, AExpr: string; AFromCol: Integer): Integer;

function NewPreviewToken(const ATool: string; const AFiles: TArray<string>;
  const AContents: TArray<string>): string; overload;
/// <summary>With the call's ARGUMENTS, so the apply can be refused when
///  they differ from what was previewed (audit #36, H34).</summary>
function NewPreviewToken(const ATool: string; const AFiles: TArray<string>;
  const AContents: TArray<string>; AArgs: TJSONObject): string; overload;
/// <summary>Without AARGS: only the tool and the buffers are checked (the
///  older call shape).</summary>
function CheckPreviewToken(const AToken, ATool: string;
  const AReadContent: TFunc<string, string>;
  out AProblem: string): Boolean; overload;
/// <summary>With AARGS: the call's arguments must be the ones the preview
///  described as well (audit #36, H34).</summary>
function CheckPreviewToken(const AToken, ATool: string;
  const AReadContent: TFunc<string, string>; AArgs: TJSONObject;
  out AProblem: string): Boolean; overload;

function DiagnosticsToJson(const AFile: string; AHash: Cardinal;
  const ADiags: TArray<TLspErrorDiag>; const ASources: TArray<string>;
  const AUsed, AStale, ANote: string): TJSONObject;

/// <summary>ALine1 > 0 keeps only fixes anchored to that 1-based line.
///  AAvailability may be nil.</summary>
/// <summary>ACONTENT is the buffer the fixes were resolved for: every fix
///  then also carries its diagnostic, the affected line verbatim and the
///  CHANGES it would make (user request 2026-09-24 - a caller has to be
///  able to judge a fix before applying it). '' = the short form.</summary>
function QuickFixesToJson(const AFile: string; AHash: Cardinal;
  const AFixes: TArray<TQuickFix>; ALine1: Integer;
  const AAvailability: TUnitAvailabilityFunc;
  const ADiagCount: Integer; const AUsed, AStale, ANote: string;
  const AContent: string = ''): TJSONObject;

implementation

uses
  System.TypInfo, System.StrUtils, System.Classes,
  Expert.UsesEditor, Expert.PascalScanner;

function ImplausibleUnitName(const AName: string): Boolean;
var
  I: Integer;
begin
  Result := True;
  if AName = '' then Exit;
  if not IsIdentStart(AName[1]) then Exit;
  I := 2;
  while I <= Length(AName) do
  begin
    // a dotted name: every segment starts like an identifier
    if AName[I] = '.' then
    begin
      if (I = Length(AName)) or not IsIdentStart(AName[I + 1]) then Exit;
      Inc(I, 2);
      Continue;
    end;
    if not IsIdentChar(AName[I]) then Exit;
    Inc(I);
  end;
  Result := False;
end;

function UnitNameSelfCheck(const ANames: TArray<string>): string;
var
  Bad: string;
  N: Integer;
begin
  Result := '';
  Bad := '';
  N := 0;
  for var Name in ANames do
    if ImplausibleUnitName(Name) then
    begin
      Inc(N);
      if N <= 5 then
      begin
        if Bad <> '' then Bad := Bad + ' | ';
        Bad := Bad + '"' + Name + '"';
      end;
    end;
  if N = 0 then Exit;
  Result := Format('DEFECT IN THIS PLUGIN, please report: %d of the %d entries ' +
    'is not a valid unit name - %s. The uses clause was not parsed correctly, ' +
    'so every verdict and every lookup keyed on these names is unreliable.',
    [N, Length(ANames), Bad]);
end;

function PreviewContentHash(const AContent: string): Cardinal;
begin
  Result := 2166136261;
  for var I := 1 to Length(AContent) do
  begin
    Result := Result xor Ord(AContent[I]);
    Result := Cardinal((UInt64(Result) * 16777619) and $FFFFFFFF);
  end;
end;

function SplitLinesLocal(const AText: string): TArray<string>;
begin
  Result := AText.Replace(#13#10, #10).Replace(#13, #10).Split([#10]);
end;

function DiffToChanges(const AFile, AOld, ANew: string; AMax: Integer;
  out ATotal: Integer): TArray<TPreviewChange>;
var
  O, N: TArray<string>;
  P, SO, SN: Integer;

  procedure Add(ALine: Integer; const ABefore, AAfter: string);
  begin
    Inc(ATotal);
    if (AMax > 0) and (Length(Result) >= AMax) then Exit;
    var C: TPreviewChange;
    C.FilePath := AFile;
    C.Line := ALine;
    C.Before := ABefore;
    C.After := AAfter;
    Result := Result + [C];
  end;

begin
  Result := nil;
  ATotal := 0;
  O := SplitLinesLocal(AOld);
  N := SplitLinesLocal(ANew);
  // common prefix
  P := 0;
  while (P <= High(O)) and (P <= High(N)) and (O[P] = N[P]) do Inc(P);
  // common suffix (never back past the prefix)
  SO := High(O);
  SN := High(N);
  while (SO >= P) and (SN >= P) and (O[SO] = N[SN]) do
  begin
    Dec(SO);
    Dec(SN);
  end;
  var I := P;
  var J := P;
  while (I <= SO) or (J <= SN) do
  begin
    if (I <= SO) and (J <= SN) then
    begin
      Add(I + 1, O[I], N[J]);
      Inc(I);
      Inc(J);
    end
    else if I <= SO then
    begin
      Add(I + 1, O[I], '');      // deleted
      Inc(I);
    end
    else
    begin
      Add(I + 1, '', N[J]);      // inserted before this line
      Inc(J);
    end;
  end;
end;

function ChangesToJson(const AChanges: TArray<TPreviewChange>): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var C in AChanges do
  begin
    var O := TJSONObject.Create;
    O.AddPair('file', C.FilePath);
    O.AddPair('line', TJSONNumber.Create(C.Line));
    O.AddPair('before', C.Before);
    O.AddPair('after', C.After);
    Result.Add(O);
  end;
end;

type
  TPreviewEntry = record
    Tool: string;
    Files: TArray<string>;
    Hashes: TArray<Cardinal>;
    Args: string;          // PreviewArgsFingerprint of the preview call
    Stamp: TDateTime;
  end;

var
  GPreviews: TDictionary<string, TPreviewEntry> = nil;
  GPreviewLock: TObject = nil;
  GPreviewCounter: Integer = 0;

function CodeOccurrenceOf(const ALine, AExpr: string; AFromCol: Integer): Integer;
var
  Masked: TArray<string>;
  M: string;
  Start: Integer;
begin
  Result := 0;
  if (ALine = '') or (AExpr = '') then Exit;
  Masked := MaskCommentsAndStrings([ALine]);
  if Length(Masked) = 0 then Exit;
  M := Masked[0];
  Start := 1;
  var Restarted := False;
  if AFromCol >= 1 then Start := AFromCol;
  while Start <= Length(ALine) do
  begin
    var P := Pos(AExpr, ALine, Start);
    if P = 0 then
    begin
      // Nothing (more) from here on. The caller's column is only a hint,
      // so try the whole line - ONCE: restarting again after a rejected
      // match would loop forever (found by the suite hanging).
      if Restarted or (Start = 1) then Exit;
      Restarted := True;
      Start := 1;
      Continue;
    end;
    Start := P + 1;
    // Code, not a comment or a string: the masked copy keeps every
    // position, so an unchanged character is code.
    if (P + Length(AExpr) - 1 > Length(M)) or (M[P] <> ALine[P]) then Continue;
    // Whole-token when the expression itself begins / ends like an
    // identifier - "Idx" must not match inside "MaxIdx".
    if IsIdentChar(AExpr[1]) and (P > 1) and IsIdentChar(ALine[P - 1]) then Continue;
    var After := P + Length(AExpr);
    if IsIdentChar(AExpr[Length(AExpr)]) and (After <= Length(ALine))
      and IsIdentChar(ALine[After]) then Continue;
    Exit(P);
  end;
end;

function PreviewArgsFingerprint(AArgs: TJSONObject): string;
var
  Pairs: TStringList;
  Key, Val: string;
  P: TJSONPair;
begin
  Result := '';
  if AArgs = nil then Exit;
  Pairs := TStringList.Create;
  try
    Pairs.Sorted := True;
    Pairs.Duplicates := dupAccept;
    for P in AArgs do
    begin
      Key := LowerCase(P.JsonString.Value);
      // These three do not describe WHAT is changed: 'apply' is the
      // difference between preview and apply, 'token' is this very
      // mechanism, and 'instance' only picks the IDE.
      if (Key = 'apply') or (Key = 'token') or (Key = 'instance') then Continue;
      Val := '';
      if P.JsonValue <> nil then Val := P.JsonValue.ToJSON;
      Pairs.Add(Key + '=' + Val);
    end;
    for Key in Pairs do
      Result := Result + Key + '&';
  finally
    Pairs.Free;
  end;
end;

function NewPreviewToken(const ATool: string; const AFiles: TArray<string>;
  const AContents: TArray<string>): string; overload;
begin
  Result := NewPreviewToken(ATool, AFiles, AContents, nil);
end;

function NewPreviewToken(const ATool: string; const AFiles: TArray<string>;
  const AContents: TArray<string>; AArgs: TJSONObject): string; overload;
var
  E: TPreviewEntry;
begin
  E.Tool := ATool;
  E.Args := PreviewArgsFingerprint(AArgs);
  E.Files := AFiles;
  E.Hashes := nil;
  for var C in AContents do E.Hashes := E.Hashes + [PreviewContentHash(C)];
  E.Stamp := Now;
  TMonitor.Enter(GPreviewLock);
  try
    Inc(GPreviewCounter);
    Result := Format('%s-%d-%s', [ATool, GPreviewCounter,
      IntToHex(PreviewContentHash(ATool + DateTimeToStr(E.Stamp) +
        IntToStr(GPreviewCounter)), 8)]);
    // keep the store small - a preview nobody applied is dead weight
    if GPreviews.Count > 32 then GPreviews.Clear;
    GPreviews.AddOrSetValue(Result, E);
  finally
    TMonitor.Exit(GPreviewLock);
  end;
end;

function CheckPreviewToken(const AToken, ATool: string;
  const AReadContent: TFunc<string, string>;
  out AProblem: string): Boolean; overload;
begin
  Result := CheckPreviewToken(AToken, ATool, AReadContent, nil, AProblem);
end;

function CheckPreviewToken(const AToken, ATool: string;
  const AReadContent: TFunc<string, string>; AArgs: TJSONObject;
  out AProblem: string): Boolean; overload;
var
  E: TPreviewEntry;
begin
  Result := False;
  AProblem := '';
  TMonitor.Enter(GPreviewLock);
  try
    if not GPreviews.TryGetValue(AToken, E) then
    begin
      AProblem := 'unknown token "' + AToken + '" - preview again (a token is ' +
        'valid in this IDE session, and only until the buffer changes)';
      Exit;
    end;
  finally
    TMonitor.Exit(GPreviewLock);
  end;
  if not SameText(E.Tool, ATool) then
  begin
    AProblem := 'that token belongs to ' + E.Tool + ', not to ' + ATool;
    Exit;
  end;
  // The ARGUMENTS have to be the ones that were previewed (audit #36,
  // H34). Only checked when the caller passes them, so a call site that
  // does not yet is unchanged.
  if (AArgs <> nil) and (E.Args <> '')
    and (PreviewArgsFingerprint(AArgs) <> E.Args) then
  begin
    AProblem := 'that token belongs to a preview with OTHER arguments - ' +
      'preview again with the arguments you want to apply, and check the ' +
      'changes it reports';
    Exit;
  end;
  for var I := 0 to High(E.Files) do
    if PreviewContentHash(AReadContent(E.Files[I])) <> E.Hashes[I] then
    begin
      AProblem := ExtractFileName(E.Files[I]) + ' changed since the preview - ' +
        'preview again and check the changes';
      Exit;
    end;
  Result := True;
end;

function SeverityName(ASeverity: Integer): string;
begin
  case ASeverity of
    1: Result := 'error';
    2: Result := 'warning';
    3: Result := 'information';
    4: Result := 'hint';
  else
    Result := 'unknown';
  end;
end;

function FixKindName(AKind: TQuickFixKind): string;
var
  S: string;
  I: Integer;
begin
  // qfAddUnit -> add_unit
  S := GetEnumName(TypeInfo(TQuickFixKind), Ord(AKind));
  if S.StartsWith('qf') then S := Copy(S, 3, MaxInt);
  Result := '';
  for I := 1 to Length(S) do
  begin
    if (I > 1) and CharInSet(S[I], ['A'..'Z']) then Result := Result + '_';
    Result := Result + LowerCase(S[I]);
  end;
end;

function SectionName(ASection: TUsesSection): string;
begin
  if ASection = usImplementation then Result := 'implementation'
  else Result := 'interface';
end;

function FixDescription(const AFix: TQuickFix): string;
begin
  case AFix.Kind of
    qfAddUnit:
      Result := Format('Add a unit declaring %s to the %s uses clause',
        [AFix.Identifier, SectionName(AFix.Section)]);
    qfRenameIdent:
      Result := Format('Replace %s with %s', [AFix.Identifier, AFix.NewText]);
    qfFixUsesName:
      Result := Format('Correct the uses entry to %s', [AFix.NewText]);
    qfRemoveUses:
      Result := Format('Remove %s from the uses clause', [AFix.OldUnit]);
    qfAlignHeader:
      if AFix.AuxLine > 0 then
        Result := Format('Align the implementation (line %d) with this declaration',
          [AFix.AuxLine + 1])
      else
        Result := 'Align the implementation header with its declaration';
    qfAlignDeclToImpl:
      Result := Format('Align this declaration with its implementation (line %d)',
        [AFix.AuxLine + 1]);
    qfRemoveVar:
      Result := Format('Remove the unused variable %s', [AFix.Identifier]);
    qfInsertSemi:
      Result := 'Insert the missing semicolon';
    qfInitVar:
      Result := Format('Initialise %s at the start of the routine', [AFix.Identifier]);
    qfRemoveAssign:
      Result := 'Remove the assignment whose value is never used';
    qfAddReintroduce:
      Result := 'Add the reintroduce directive';
    qfImplStub:
      Result := 'Create an empty implementation for the declaration';
    qfClassStub:
      Result := Format('Create the declaration of the forward class %s', [AFix.Identifier]);
    qfRemoveToken:
      Result := Format('Remove the stray token %s', [AFix.Identifier]);
    qfRemovePrivate:
      Result := Format('Remove the unused private member %s (declaration and ' +
        'implementation)', [AFix.Identifier]);
    qfDeclareVar:
      Result := Format('Declare %s as a local variable%s', [AFix.Identifier,
        IfThen(AFix.NewText <> '', ': ' + AFix.NewText, '')]);
    qfDeclareInlineVar:
      Result := Format('Declare %s inline (var %s := ...)', [AFix.Identifier, AFix.Identifier]);
  else
    Result := AFix.Caption;
  end;
end;

function MakeFixId(AHash: Cardinal; AIndex: Integer): string;
begin
  Result := IntToHex(AHash, 8) + '-' + IntToStr(AIndex);
end;

function ParseFixId(const AId: string; out AHash: Cardinal; out AIndex: Integer): Boolean;
var
  P: Integer;
  H: Int64;
begin
  Result := False;
  AHash := 0;
  AIndex := -1;
  P := Pos('-', AId);
  if P <> 9 then Exit;
  if not TryStrToInt64('$' + Copy(AId, 1, 8), H) then Exit;
  if not TryStrToInt(Copy(AId, P + 1, MaxInt), AIndex) or (AIndex < 0) then Exit;
  AHash := Cardinal(H);
  Result := True;
end;

function DiagnosticsToJson(const AFile: string; AHash: Cardinal;
  const ADiags: TArray<TLspErrorDiag>; const ASources: TArray<string>;
  const AUsed, AStale, ANote: string): TJSONObject;
var
  Arr: TJSONArray;
  Counts: array[1..4] of Integer;
  I: Integer;
begin
  Result := TJSONObject.Create;
  Result.AddPair('file', AFile);
  Result.AddPair('revision', IntToHex(AHash, 8));
  FillChar(Counts, SizeOf(Counts), 0);
  Arr := TJSONArray.Create;
  for I := 0 to High(ADiags) do
  begin
    var D := ADiags[I];
    var O := TJSONObject.Create;
    O.AddPair('severity', SeverityName(D.Severity));
    O.AddPair('code', D.Code);
    O.AddPair('message', D.Message);
    O.AddPair('line', TJSONNumber.Create(D.Range.Start.Line + 1));
    O.AddPair('column', TJSONNumber.Create(D.Range.Start.Character + 1));
    if (D.Range.End_.Line <> D.Range.Start.Line) or
       (D.Range.End_.Character <> D.Range.Start.Character) then
    begin
      O.AddPair('endLine', TJSONNumber.Create(D.Range.End_.Line + 1));
      O.AddPair('endColumn', TJSONNumber.Create(D.Range.End_.Character + 1));
    end;
    if I <= High(ASources) then O.AddPair('sources', ASources[I]);
    Arr.Add(O);
    if (D.Severity >= 1) and (D.Severity <= 4) then Inc(Counts[D.Severity]);
  end;
  Result.AddPair('summary', Format('%d error(s), %d warning(s), %d hint(s)',
    [Counts[1], Counts[2], Counts[3] + Counts[4]]));
  Result.AddPair('diagnostics', Arr);
  Result.AddPair('sourcesUsed', AUsed);
  if AStale <> '' then
    Result.AddPair('sourcesStale', AStale + ' (computed for an older buffer ' +
      'state, ignored)');
  if ANote <> '' then Result.AddPair('note', ANote);
end;

function QuickFixesToJson(const AFile: string; AHash: Cardinal;
  const AFixes: TArray<TQuickFix>; ALine1: Integer;
  const AAvailability: TUnitAvailabilityFunc;
  const ADiagCount: Integer; const AUsed, AStale, ANote: string;
  const AContent: string): TJSONObject;
var
  Arr: TJSONArray;
  I: Integer;
  Lines: TArray<string>;

  function LineText(ALine0: Integer): string;
  begin
    if (ALine0 >= 0) and (ALine0 <= High(Lines)) then
      Result := Lines[ALine0]
    else
      Result := '';
  end;

begin
  Lines := nil;
  if AContent <> '' then Lines := SplitLinesLocal(AContent);
  Result := TJSONObject.Create;
  Result.AddPair('file', AFile);
  Result.AddPair('revision', IntToHex(AHash, 8));
  Arr := TJSONArray.Create;
  for I := 0 to High(AFixes) do
  begin
    var F := AFixes[I];
    if (ALine1 > 0) and (F.Line + 1 <> ALine1) then Continue;
    var O := TJSONObject.Create;
    O.AddPair('id', MakeFixId(AHash, I));
    O.AddPair('kind', FixKindName(F.Kind));
    O.AddPair('line', TJSONNumber.Create(F.Line + 1));
    if F.TokenLen > 0 then
      O.AddPair('column', TJSONNumber.Create(F.Col + 1));
    O.AddPair('description', FixDescription(F));
    if F.Caption <> '' then O.AddPair('caption', F.Caption);
    if F.Identifier <> '' then O.AddPair('identifier', F.Identifier);
    if F.Kind = qfAddUnit then
    begin
      var U := TJSONArray.Create;
      for var N in F.UnitNames do
      begin
        var UO := TJSONObject.Create;
        UO.AddPair('unit', N);
        if Assigned(AAvailability) then
          UO.AddPair('availability', AAvailability(N));
        U.Add(UO);
      end;
      O.AddPair('units', U);
      O.AddPair('section', SectionName(F.Section));
    end;
    if F.FollowUpUnit <> '' then
      O.AddPair('alsoAddsUnit', F.FollowUpUnit);
    if F.Kind = qfRemovePrivate then
      O.AddPair('note', 'A non-empty body is only removed after confirmation ' +
        'in the IDE; through this tool such a fix is refused.');
    if F.DiagCode <> '' then
    begin
      var D := TJSONObject.Create;
      D.AddPair('code', F.DiagCode);
      D.AddPair('message', F.DiagMessage);
      D.AddPair('line', TJSONNumber.Create(F.DiagLine + 1));
      D.AddPair('column', TJSONNumber.Create(F.DiagCol + 1));
      O.AddPair('diagnostic', D);
    end;
    if Lines <> nil then
    begin
      O.AddPair('before', LineText(F.Line));
      // the second line a fix works on: where a ';' goes / where the
      // implementation of an aligned declaration sits
      if F.AuxLine > 0 then
        case F.Kind of
          qfInsertSemi, qfInitVar:
            begin
              O.AddPair('insertAtLine', TJSONNumber.Create(F.AuxLine + 1));
              O.AddPair('beforeAux', LineText(F.AuxLine));
            end;
          qfAlignHeader, qfAlignDeclToImpl:
            begin
              O.AddPair('partnerLine', TJSONNumber.Create(F.AuxLine + 1));
              O.AddPair('beforeAux', LineText(F.AuxLine));
            end;
        end;
      if (F.NewText <> '') and not (F.Kind in [qfRemoveAssign, qfRemoveToken,
        qfRemovePrivate, qfDeclareVar, qfDeclareInlineVar]) then
        O.AddPair('newText', F.NewText);
      var Planned: string;
      if PlanQuickFixText(AContent, F, 0, Planned) then
      begin
        var Total := 0;
        var Ch := DiffToChanges(AFile, AContent, Planned, 20, Total);
        O.AddPair('changes', ChangesToJson(Ch));
        if Total > Length(Ch) then
          O.AddPair('changesTruncated', TJSONNumber.Create(Total));
      end
      else
        O.AddPair('changes', 'not computed for this kind - it generates code ' +
          'from a header elsewhere in the file (see newText and line)');
    end;
    Arr.Add(O);
  end;
  Result.AddPair('fixes', Arr);
  Result.AddPair('diagnosticsConsidered', TJSONNumber.Create(ADiagCount));
  Result.AddPair('sourcesUsed', AUsed);
  if AStale <> '' then
    Result.AddPair('sourcesStale', AStale + ' (computed for an older buffer ' +
      'state, ignored)');
  if ANote <> '' then Result.AddPair('note', ANote);
end;

initialization
  GPreviewLock := TObject.Create;
  GPreviews := TDictionary<string, TPreviewEntry>.Create;

finalization
  GPreviews.Free;
  GPreviewLock.Free;

end.
