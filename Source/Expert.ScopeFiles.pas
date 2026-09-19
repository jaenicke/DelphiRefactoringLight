(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.ScopeFiles;

// The file set of a "whole project" scan - rename, find references, find
// implementations, find unit references.
//
// Forum report: "Rename changes the type in every unit of the project,
// but in the unit where I start the rename - which is not part of the
// project - nothing happens." The scans only ever looked at the project's
// own source list, so a unit that is merely OPEN (or reached through a
// search path) was invisible, even when it holds the declaration itself.
//
// Three layers, the first one unconditional:
//  1. the project sources PLUS the caret's own file - the place the user
//     starts from is always part of the answer;
//  2. optionally every unit OPEN in the editor (cheap, and exactly what
//     the user is looking at);
//  3. optionally every unit reachable through USES clauses (transitive),
//     resolved via the identifier index - which covers units that live on
//     the search/browsing path. Units below the RAD Studio installation
//     are never included: nobody renames inside the shipped RTL/VCL, and
//     scanning it would dwarf every project.
//  4. always: the files those units pull in via {$I file} / {$INCLUDE file}.
//     Several units are only "unit X; {$I X.inc}" - without this the scans
//     never saw their code (find references missed uses there, and a rename
//     would have left them unchanged). The scans verify positions inside an
//     include through Expert.IncludeExpansion.TLspIncludeContext.

interface

uses
  Expert.IncludeExpansion;

type
  TScopeExtra = record
    OpenUnits: Integer;    // added because they are open in the editor
    UsedUnits: Integer;    // added via uses clauses
    CaretFile: Boolean;    // the caret's file was not in the project
    Truncated: Boolean;    // uses walk hit MaxScopeFiles
    IncludeFiles: Integer; // added because a scanned unit includes them
  end;

const
  MaxScopeFiles = 5000;

/// <summary>Project sources + the caret's file + (optionally) open units
///  + (optionally) units reachable via uses. MAIN THREAD (reads editor
///  buffers). Duplicates are removed case-insensitively; the project's
///  own order comes first.</summary>
function ProjectScopeFiles(const ACaretFile: string; AIncludeOpen,
  AIncludeUsed: Boolean; out AExtra: TScopeExtra): TArray<string>; overload;

/// <summary>Same, with the two switches taken from TPluginSettings.</summary>
function ProjectScopeFiles(const ACaretFile: string): TArray<string>; overload;

/// <summary>"123 project file(s) + the current unit + 4 open unit(s) +
///  17 unit(s) via uses" - for status lines and diagnostics logs.
///  ATotal is the length of the returned file list.</summary>
function ScopeExtraText(ATotal: Integer; const AExtra: TScopeExtra): string;

/// <summary>Reader for Expert.IncludeExpansion: the editor buffer when the
///  file is open, else disk. MAIN THREAD only (ToolsAPI).</summary>
function EditorOrDiskReader: TIncludeReader;

implementation

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Generics.Collections,
  Expert.EditorHelperIntf, Expert.UnitIndex, Expert.UsesGraph,
  Expert.PluginSettings, Delphi.FileEncoding;

function IsSourceFile(const AFile: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(AFile));
  Result := (Ext = '.pas') or (Ext = '.dpr') or (Ext = '.dpk');
end;

function ReadContent(const AFile: string; out AContent: string): Boolean;
begin
  Result := True;
  if (Editor <> nil) and Editor.ReadEditorContent(AFile, AContent) then Exit;
  AContent := '';
  if not TFile.Exists(AFile) then Exit(False);
  try
    AContent := TDelphiFileEncoding.ReadAll(AFile);
  except
    Result := False;
  end;
end;

function ProjectScopeFiles(const ACaretFile: string; AIncludeOpen,
  AIncludeUsed: Boolean; out AExtra: TScopeExtra): TArray<string>;
var
  Seen: TDictionary<string, Boolean>;
  List: TList<string>;
  BdsRoot: string;

  function Add(const AFile: string): Boolean;
  var
    Key: string;
  begin
    Result := False;
    if (AFile = '') or not IsSourceFile(AFile) then Exit;
    Key := UpperCase(ExpandFileName(AFile));
    if Seen.ContainsKey(Key) then Exit;
    Seen.Add(Key, True);
    List.Add(AFile);
    Result := True;
  end;

  // an include file keeps whatever extension it has (.inc, .pas, .txt ...)
  function AddInclude(const AFile: string): Boolean;
  var
    Key: string;
  begin
    Key := UpperCase(ExpandFileName(AFile));
    Result := not Seen.ContainsKey(Key);
    if Result then
    begin
      Seen.Add(Key, True);
      List.Add(AFile);
    end;
  end;

  function BelowBds(const AFile: string): Boolean;
  begin
    Result := (BdsRoot <> '') and
      UpperCase(ExpandFileName(AFile)).StartsWith(BdsRoot);
  end;

var
  Queue: TQueue<string>;
  Snap: IUnitSnapshot;
  Content, Path: string;
begin
  AExtra := Default(TScopeExtra);
  Seen := TDictionary<string, Boolean>.Create;
  List := TList<string>.Create;
  try
    if Editor <> nil then
      for var F in Editor.GetProjectSourceFiles do
        Add(F);

    // 1. Always: where the user started.
    if Add(ACaretFile) then AExtra.CaretFile := True;

    // 2. Units open in the editor.
    if AIncludeOpen and (Editor <> nil) then
      for var F in Editor.GetOpenSourceFiles do
        if SameText(ExtractFileExt(F), '.pas') and Add(F) then
          Inc(AExtra.OpenUnits);

    // 3. Transitively through uses clauses.
    if AIncludeUsed then
    begin
      BdsRoot := FindBdsRoot;
      if BdsRoot <> '' then
        BdsRoot := UpperCase(IncludeTrailingPathDelimiter(ExpandFileName(BdsRoot)));
      Snap := TUnitIndex.Instance.Snapshot;
      if Snap <> nil then
      begin
        Queue := TQueue<string>.Create;
        try
          for var F in List do Queue.Enqueue(F);
          while Queue.Count > 0 do
          begin
            if List.Count >= MaxScopeFiles then
            begin
              AExtra.Truncated := True;
              Break;
            end;
            if not ReadContent(Queue.Dequeue, Content) then Continue;
            for var E in TUsesGraphAnalyzer.ParseUsesEntries(Content) do
            begin
              if not Snap.TryGetUnitPath(E.UnitName, Path) then Continue;
              if BelowBds(Path) then Continue;
              if Add(Path) then
              begin
                Inc(AExtra.UsedUnits);
                Queue.Enqueue(Path);
              end;
            end;
          end;
        finally
          Queue.Free;
        end;
      end;
    end;

    // 4. Include files of everything collected so far (nested ones too).
    if BdsRoot = '' then
    begin
      BdsRoot := FindBdsRoot;
      if BdsRoot <> '' then
        BdsRoot := UpperCase(IncludeTrailingPathDelimiter(ExpandFileName(BdsRoot)));
    end;
    var Units := List.ToArray;
    for var F in Units do
    begin
      if List.Count >= MaxScopeFiles then
      begin
        AExtra.Truncated := True;
        Break;
      end;
      if not ReadContent(F, Content) then Continue;
      for var IncF in CollectIncludeFiles(F, Content,
        function(const APath: string; out AContent: string): Boolean
        begin
          Result := ReadContent(APath, AContent);
        end) do
        if not BelowBds(IncF) and AddInclude(IncF) then
          Inc(AExtra.IncludeFiles);
    end;

    Result := List.ToArray;
  finally
    List.Free;
    Seen.Free;
  end;
end;

function ProjectScopeFiles(const ACaretFile: string): TArray<string>;
var
  Extra: TScopeExtra;
begin
  Result := ProjectScopeFiles(ACaretFile, TPluginSettings.ScopeIncludeOpenUnits,
    TPluginSettings.ScopeIncludeUsedUnits, Extra);
end;

function EditorOrDiskReader: TIncludeReader;
begin
  Result :=
    function(const APath: string; out AContent: string): Boolean
    begin
      Result := ReadContent(APath, AContent);
    end;
end;

function ScopeExtraText(ATotal: Integer; const AExtra: TScopeExtra): string;
begin
  Result := Format('%d project file(s)', [ATotal - AExtra.OpenUnits -
    AExtra.UsedUnits - AExtra.IncludeFiles - Ord(AExtra.CaretFile)]);
  if AExtra.CaretFile then Result := Result + ' + the current unit';
  if AExtra.OpenUnits > 0 then
    Result := Result + Format(' + %d open unit(s)', [AExtra.OpenUnits]);
  if AExtra.UsedUnits > 0 then
    Result := Result + Format(' + %d unit(s) via uses', [AExtra.UsedUnits]);
  if AExtra.IncludeFiles > 0 then
    Result := Result + Format(' + %d include file(s)', [AExtra.IncludeFiles]);
  if AExtra.Truncated then
    Result := Result + Format(' (uses walk stopped at %d files)', [MaxScopeFiles]);
end;

end.
