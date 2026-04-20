(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.IdentifierCheck;

interface

uses
  System.SysUtils, System.Classes;

type
  TIdentifierCheckStatus = (icsEmpty, icsOk, icsInvalid, icsKeyword, icsUnchanged,
    icsInUnit, icsInProject);

  TIdentifierCheckResult = record
    Status: TIdentifierCheckStatus;
    Message: string;
    InUnitCount: Integer;
    InProjectFileCount: Integer;
  end;

  TIdentifierChecker = class
  public
    class function IsValidIdentifier(const AName: string): Boolean; static;
    class function IsPascalKeyword(const AName: string): Boolean; static;
    class function CountInText(const AText, AIdent: string): Integer; static;

    /// <summary>Full in-memory check. ACurrentFileText is the current
    ///  unit's full source (for in-unit collision detection).
    ///  AOtherContents are the pre-loaded contents of all OTHER project
    ///  files (current file already excluded). AOldName (may be empty) is
    ///  the identifier being renamed, used for "unchanged" detection.</summary>
    class function Check(const ANewName, AOldName, ACurrentFileText: string;
      const AOtherContents: TArray<string>): TIdentifierCheckResult; static;
  end;

  TProjectIndexBuildThread = class(TThread)
  private
    FProjectFiles: TArray<string>;
    FExcludeFile: string;
    FResultFiles: TArray<string>;
    FResultContents: TArray<string>;
  protected
    procedure Execute; override;
  public
    constructor Create(const AProjectFiles: TArray<string>; const AExcludeFile: string);
    property ResultFiles: TArray<string> read FResultFiles;
    property ResultContents: TArray<string> read FResultContents;
  end;

  /// <summary>Background-loaded cache of all project file contents. The
  ///  main thread calls Build once, then polls via PollReady on each
  ///  check attempt. Once ready, OtherContents is safe to read.</summary>
  TProjectTextIndex = class
  private
    FThread: TProjectIndexBuildThread;
    FOtherFiles: TArray<string>;
    FOtherContents: TArray<string>;
    FReady: Boolean;
    procedure Stop;
  public
    destructor Destroy; override;
    procedure Build(const AProjectFiles: TArray<string>; const ACurrentFile: string);
    /// <summary>Called from the main thread. Returns True once the
    ///  background load has finished; first such call moves results into
    ///  OtherFiles/OtherContents and frees the thread.</summary>
    function PollReady: Boolean;
    property IsReady: Boolean read FReady;
    property OtherFiles: TArray<string> read FOtherFiles;
    property OtherContents: TArray<string> read FOtherContents;
  end;

implementation

uses
  System.Character, System.StrUtils, System.IOUtils;

const
  KEYWORDS: array[0..65] of string = (
    'AND', 'ARRAY', 'AS', 'ASM', 'BEGIN', 'CASE', 'CLASS', 'CONST',
    'CONSTRUCTOR', 'DESTRUCTOR', 'DISPINTERFACE', 'DIV', 'DO', 'DOWNTO',
    'ELSE', 'END', 'EXCEPT', 'EXPORTS', 'FILE', 'FINALIZATION', 'FINALLY',
    'FOR', 'FUNCTION', 'GOTO', 'IF', 'IMPLEMENTATION', 'IN', 'INHERITED',
    'INITIALIZATION', 'INLINE', 'INTERFACE', 'IS', 'LABEL', 'LIBRARY',
    'MOD', 'NIL', 'NOT', 'OBJECT', 'OF', 'OR', 'PACKED', 'PROCEDURE',
    'PROGRAM', 'PROPERTY', 'RAISE', 'RECORD', 'REPEAT', 'RESOURCESTRING',
    'SET', 'SHL', 'SHR', 'STRING', 'THEN', 'THREADVAR', 'TO', 'TRY',
    'TYPE', 'UNIT', 'UNTIL', 'USES', 'VAR', 'WHILE', 'WITH', 'XOR',
    'FALSE', 'TRUE');

class function TIdentifierChecker.IsValidIdentifier(const AName: string): Boolean;
var
  I: Integer;
  Ch: Char;
begin
  if AName = '' then Exit(False);
  Ch := AName[1];
  if not (Ch.IsLetter or (Ch = '_')) then Exit(False);
  for I := 2 to Length(AName) do
  begin
    Ch := AName[I];
    if not (Ch.IsLetterOrDigit or (Ch = '_')) then Exit(False);
  end;
  Result := True;
end;

class function TIdentifierChecker.IsPascalKeyword(const AName: string): Boolean;
var
  U, K: string;
begin
  U := UpperCase(AName);
  for K in KEYWORDS do
    if K = U then Exit(True);
  Result := False;
end;

class function TIdentifierChecker.CountInText(const AText, AIdent: string): Integer;
var
  I, Len, ILen: Integer;
  UText, UIdent: string;
  Ch: Char;
  Matches: Boolean;
begin
  Result := 0;
  if (AText = '') or (AIdent = '') then Exit;
  UText := UpperCase(AText);
  UIdent := UpperCase(AIdent);
  ILen := Length(UIdent);
  Len := Length(UText);
  I := 1;
  while I <= Len - ILen + 1 do
  begin
    if UText[I] = UIdent[1] then
    begin
      Matches := (I = 1);
      if not Matches then
      begin
        Ch := UText[I - 1];
        Matches := not (Ch.IsLetterOrDigit or (Ch = '_'));
      end;
      if Matches then
      begin
        Matches := CompareMem(@UText[I], @UIdent[1], ILen * SizeOf(Char));
        if Matches and (I + ILen <= Len) then
        begin
          Ch := UText[I + ILen];
          Matches := not (Ch.IsLetterOrDigit or (Ch = '_'));
        end;
      end;
      if Matches then
      begin
        Inc(Result);
        Inc(I, ILen);
        Continue;
      end;
    end;
    Inc(I);
  end;
end;

class function TIdentifierChecker.Check(const ANewName, AOldName, ACurrentFileText: string;
  const AOtherContents: TArray<string>): TIdentifierCheckResult;
var
  Trimmed: string;
  InUnitTotal, C, I: Integer;
  ProjHits, ProjFilesWith: Integer;
begin
  Result := Default(TIdentifierCheckResult);
  Trimmed := Trim(ANewName);

  if Trimmed = '' then
  begin
    Result.Status := icsEmpty;
    Result.Message := 'Please enter an identifier.';
    Exit;
  end;

  if not IsValidIdentifier(Trimmed) then
  begin
    Result.Status := icsInvalid;
    Result.Message := 'Invalid identifier.';
    Exit;
  end;

  if IsPascalKeyword(Trimmed) then
  begin
    Result.Status := icsKeyword;
    Result.Message := '"' + Trimmed + '" is a reserved Pascal keyword.';
    Exit;
  end;

  if (AOldName <> '') and SameText(Trimmed, AOldName) then
  begin
    Result.Status := icsUnchanged;
    Result.Message := 'Name unchanged.';
    Exit;
  end;

  InUnitTotal := CountInText(ACurrentFileText, Trimmed);
  Result.InUnitCount := InUnitTotal;
  if InUnitTotal > 0 then
  begin
    Result.Status := icsInUnit;
    Result.Message := Format('"%s" already appears %d time(s) in this unit.',
      [Trimmed, InUnitTotal]);
  end;

  ProjHits := 0;
  ProjFilesWith := 0;
  for I := 0 to High(AOtherContents) do
  begin
    C := CountInText(AOtherContents[I], Trimmed);
    if C > 0 then
    begin
      Inc(ProjHits, C);
      Inc(ProjFilesWith);
    end;
  end;
  Result.InProjectFileCount := ProjFilesWith;

  if Result.Status = icsInUnit then Exit;

  if ProjHits > 0 then
  begin
    Result.Status := icsInProject;
    Result.Message := Format('"%s" exists in %d other project file(s) (%d occurrence(s)).',
      [Trimmed, ProjFilesWith, ProjHits]);
    Exit;
  end;

  Result.Status := icsOk;
  Result.Message := Format('"%s" is available.', [Trimmed]);
end;

{ TProjectIndexBuildThread }

constructor TProjectIndexBuildThread.Create(const AProjectFiles: TArray<string>;
  const AExcludeFile: string);
begin
  inherited Create(True);
  FProjectFiles := AProjectFiles;
  FExcludeFile := AExcludeFile;
  FreeOnTerminate := False;
end;

procedure TProjectIndexBuildThread.Execute;
var
  Paths, Contents: TArray<string>;
  N, K, I: Integer;
  Content: string;
begin
  N := Length(FProjectFiles);
  SetLength(Paths, N);
  SetLength(Contents, N);
  K := 0;
  for I := 0 to N - 1 do
  begin
    if Terminated then Break;
    if SameText(FProjectFiles[I], FExcludeFile) then Continue;
    try
      Content := TFile.ReadAllText(FProjectFiles[I]);
    except
      Content := '';
    end;
    Paths[K] := FProjectFiles[I];
    Contents[K] := Content;
    Inc(K);
  end;
  SetLength(Paths, K);
  SetLength(Contents, K);
  FResultFiles := Paths;
  FResultContents := Contents;
end;

{ TProjectTextIndex }

destructor TProjectTextIndex.Destroy;
begin
  Stop;
  inherited;
end;

procedure TProjectTextIndex.Stop;
begin
  if Assigned(FThread) then
  begin
    FThread.Terminate;
    FThread.WaitFor;
    FreeAndNil(FThread);
  end;
end;

procedure TProjectTextIndex.Build(const AProjectFiles: TArray<string>; const ACurrentFile: string);
begin
  Stop;
  FOtherFiles := nil;
  FOtherContents := nil;
  FReady := False;
  FThread := TProjectIndexBuildThread.Create(AProjectFiles, ACurrentFile);
  FThread.Start;
end;

function TProjectTextIndex.PollReady: Boolean;
begin
  if FReady then Exit(True);
  if not Assigned(FThread) then Exit(False);
  if FThread.Finished then
  begin
    FOtherFiles := FThread.ResultFiles;
    FOtherContents := FThread.ResultContents;
    FreeAndNil(FThread);
    FReady := True;
    Exit(True);
  end;
  Result := False;
end;

end.
