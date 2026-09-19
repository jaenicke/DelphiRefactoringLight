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
  System.SysUtils, System.JSON, Lsp.Protocol, Expert.AutoImport;

type
  /// <summary>'in_project' / 'search_path' / 'dcu' / 'browsing_only' /
  ///  'unknown' for a candidate unit of an add-unit fix.</summary>
  TUnitAvailabilityFunc = reference to function(const AUnit: string): string;

function SeverityName(ASeverity: Integer): string;
function FixKindName(AKind: TQuickFixKind): string;
function FixDescription(const AFix: TQuickFix): string;

function MakeFixId(AHash: Cardinal; AIndex: Integer): string;
function ParseFixId(const AId: string; out AHash: Cardinal; out AIndex: Integer): Boolean;

function DiagnosticsToJson(const AFile: string; AHash: Cardinal;
  const ADiags: TArray<TLspErrorDiag>; const ASources: TArray<string>;
  const AUsed, AStale, ANote: string): TJSONObject;

/// <summary>ALine1 > 0 keeps only fixes anchored to that 1-based line.
///  AAvailability may be nil.</summary>
function QuickFixesToJson(const AFile: string; AHash: Cardinal;
  const AFixes: TArray<TQuickFix>; ALine1: Integer;
  const AAvailability: TUnitAvailabilityFunc;
  const ADiagCount: Integer; const AUsed, AStale, ANote: string): TJSONObject;

implementation

uses
  System.TypInfo, System.StrUtils, Expert.UsesEditor;

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
      Result := 'Align the implementation header with its declaration';
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
  const ADiagCount: Integer; const AUsed, AStale, ANote: string): TJSONObject;
var
  Arr: TJSONArray;
  I: Integer;
begin
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

end.
