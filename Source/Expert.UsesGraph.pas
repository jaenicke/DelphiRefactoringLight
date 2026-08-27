(*
 * Copyright (c) 2026 Sebastian Jänicke (github.com/jaenicke)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *)
unit Expert.UsesGraph;

// Builds the unit-dependency graph from the uses clauses of all
// project sources and finds circular reference groups (strongly
// connected components with more than one member).
//
// Semantics worth knowing:
//   * interface-to-interface cycles are compile errors (F2047) - a
//     compiling project cannot contain them. What this analysis
//     surfaces are the LEGAL cycles that run through at least one
//     implementation-uses edge: they compile fine but are the
//     architecture knots the maintainer wants to see.
//   * Only project units participate. Units from the RTL/VCL or
//     library paths cannot form a cycle back into the project.
//   * {$IFDEF}-guarded uses entries: both branches are collected
//     (over-approximation) - conservative for cycle detection.

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type
  TUnitUsesEntry = record
    UnitName: string;
    InInterface: Boolean;
    /// <summary>1-based line of the entry in the source file.</summary>
    Line: Integer;
  end;

  /// <summary>One uses-clause edge that participates in a cycle.
  ///  FromUnit's uses clause names ToUnit, and both belong to the
  ///  same strongly connected component (cycle group).</summary>
  TCycleEdge = record
    FromUnit: string;
    FromFile: string;
    ToUnit: string;
    InInterface: Boolean;
    Line: Integer;
    /// <summary>1-based cycle group number. Groups are disjoint sets
    ///  of units that are all mutually reachable.</summary>
    Group: Integer;
    /// <summary>Number of units in this group (same for all edges of
    ///  a group; convenient for display).</summary>
    GroupSize: Integer;
  end;

  /// <summary>One hop of a concrete cycle path: FromUnit's uses clause
  ///  (in the given section, at the given line) names ToUnit.</summary>
  TCycleHop = record
    FromUnit: string;
    FromFile: string;
    ToUnit: string;
    InInterface: Boolean;
    Line: Integer;
  end;

  /// <summary>A unit ranked by how entangled it is - how many
  ///  cycle-internal uses edges touch it (incoming + outgoing). The
  ///  higher, the more cycles run through it, so the better a target
  ///  for breaking dependencies.</summary>
  TUnitHotspot = record
    UnitName: string;
    FileName: string;
    Group: Integer;
    GroupSize: Integer;
    EdgeCount: Integer;
  end;

  /// <summary>One concrete simple cycle as an ordered ring of unit
  ///  names; the last unit's uses closes back to the first.</summary>
  TCyclePath = record
    Units: TArray<string>;
  end;

  /// <summary>A dependency (FromUnit uses ToUnit) ranked by leverage:
  ///  how many units drop out of any cycle if this single uses entry
  ///  is removed. The biggest number = the biggest lever.</summary>
  TEdgeLever = record
    FromUnit: string;
    FromFile: string;
    ToUnit: string;
    Section: string;    // 'interface' | 'implementation' | 'both'
    Line: Integer;
    Group: Integer;
    UnitsFreed: Integer;
  end;

  /// <summary>Per-cycle-group metrics.</summary>
  TCycleGroupInfo = record
    Group: Integer;
    /// <summary>Number of units in the strongly connected component -
    ///  the largest a single cycle in this group could span (exact
    ///  longest is NP-hard, so this is the honest upper bound).</summary>
    UnitCount: Integer;
    /// <summary>Girth: length (in units) of the SHORTEST cycle in this
    ///  group. Exact.</summary>
    ShortestCycle: Integer;
    /// <summary>Number of cycle-internal uses edges in this group.</summary>
    EdgeCount: Integer;
  end;

  /// <summary>Result of a cycle analysis. Owns the graph so the dialog
  ///  can answer interactive queries (concrete cycle path for a
  ///  clicked edge, hotspot ranking). Free when done.</summary>
  TUsesCycleResult = class
  private
    FNames: TArray<string>;
    FFiles: TArray<string>;
    FComp: TArray<Integer>;         // node -> component id
    FCompSize: TArray<Integer>;
    FNameToIdx: TDictionary<string, Integer>;
    // SCC-internal out-adjacency with edge metadata (packed integers):
    // FAdjTo[u], FAdjIntf[u], FAdjLine[u] are parallel arrays.
    FAdjTo: TArray<TList<Integer>>;
    FAdjIntf: TArray<TList<Boolean>>;
    FAdjLine: TArray<TList<Integer>>;
    FEdges: TArray<TCycleEdge>;
    FCompToGroup: TDictionary<Integer, Integer>;
    function ComponentGirth(const AMembers: TArray<Integer>): Integer;
    /// <summary>Number of nodes in components of size >= 2 when every
    ///  edge from AExFrom to AExTo is removed. (-1,-1) = remove
    ///  nothing (baseline). Used by EdgeLevers.</summary>
    function CountCycleNodes(AExFrom, AExTo: Integer): Integer;
  public
    constructor Create;
    destructor Destroy; override;
    /// <summary>Flat list of every uses edge inside a cycle group -
    ///  the master list shown in the dialog.</summary>
    property Edges: TArray<TCycleEdge> read FEdges;
    /// <summary>Given a clicked edge, returns the shortest concrete
    ///  cycle that runs through it: hop 0 is the clicked edge itself,
    ///  the remaining hops form the shortest path back from ToUnit to
    ///  FromUnit. The last hop's ToUnit equals hop 0's FromUnit, so
    ///  the ring is closed.</summary>
    function ShortestCyclePath(const AStart: TCycleEdge): TArray<TCycleHop>;
    /// <summary>Units ranked (desc) by incident cycle-edge count.</summary>
    function Hotspots: TArray<TUnitHotspot>;
    /// <summary>Per-group metrics (shortest cycle length, component
    ///  size, edge count), ordered by group number.</summary>
    function GroupInfos: TArray<TCycleGroupInfo>;
    /// <summary>Enumerates simple cycles (each once, rooted at its
    ///  smallest node). Stops after AMax cycles; ATruncated then True.
    ///  The count of simple cycles can be exponential, hence the cap.</summary>
    function EnumerateCycles(AMax: Integer; out ATruncated: Boolean): TArray<TCyclePath>;
    /// <summary>Dependencies ranked (desc) by how many units they free
    ///  from cycles when removed - the biggest levers. Exact, no
    ///  enumeration.</summary>
    function EdgeLevers: TArray<TEdgeLever>;
  end;

  TUsesGraphAnalyzer = class
  public
    /// <summary>Extracts all uses entries (interface + implementation)
    ///  from a unit source. Comment/string aware; conditional blocks
    ///  contribute both branches.</summary>
    class function ParseUsesEntries(const AContent: string): TArray<TUnitUsesEntry>;

    /// <summary>Scans AFiles (.pas only), builds the dependency graph
    ///  restricted to project units, finds cycle groups and returns a
    ///  result object (nil-safe: always non-nil; Edges empty when no
    ///  cycles). Caller owns the result.</summary>
    class function Analyze(const AFiles: TArray<string>;
      const AProgress: TProc<Integer, Integer, string> = nil): TUsesCycleResult;
  end;

implementation

uses
  System.IOUtils, System.StrUtils, System.Generics.Defaults,
  Expert.EditorHelperIntf, Delphi.FileEncoding;

function ReadFileContent(const AFile: string; out AContent: string): Boolean;
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

function IsIdentChar(C: Char): Boolean; inline;
begin
  Result := CharInSet(C, ['A'..'Z', 'a'..'z', '0'..'9', '_', '.']);
end;

class function TUsesGraphAnalyzer.ParseUsesEntries(
  const AContent: string): TArray<TUnitUsesEntry>;
var
  Entries: TList<TUnitUsesEntry>;
  Lines: TArray<string>;
  InBrace, InParenStar, InImpl, Collecting: Boolean;
  LineNo: Integer;

  /// <summary>Blanks comments and string literals (stateful across
  ///  lines for { } and (* *) blocks) so the keyword scan below never
  ///  fires inside them. Blanking (instead of removing) keeps column
  ///  positions stable.</summary>
  function CleanLine(const L: string): string;
  var
    I: Integer;
    Res: TStringBuilder;
    InStr: Boolean;
  begin
    Res := TStringBuilder.Create;
    try
      InStr := False;
      I := 1;
      while I <= Length(L) do
      begin
        if InBrace then
        begin
          if L[I] = '}' then InBrace := False;
          Res.Append(' ');
        end
        else if InParenStar then
        begin
          if (L[I] = '*') and (I < Length(L)) and (L[I + 1] = ')') then
          begin
            InParenStar := False;
            Res.Append('  ');
            Inc(I, 2);
            Continue;
          end;
          Res.Append(' ');
        end
        else if InStr then
        begin
          if L[I] = '''' then InStr := False;
          Res.Append(' ');
        end
        else
          case L[I] of
            '''': begin InStr := True; Res.Append(' '); end;
            '{':  begin InBrace := True; Res.Append(' '); end;
            '(':
              if (I < Length(L)) and (L[I + 1] = '*') then
              begin
                InParenStar := True;
                Res.Append('  ');
                Inc(I, 2);
                Continue;
              end
              else
                Res.Append(L[I]);
            '/':
              if (I < Length(L)) and (L[I + 1] = '/') then
              begin
                // line comment - blank the rest
                while I <= Length(L) do begin Res.Append(' '); Inc(I); end;
                Break;
              end
              else
                Res.Append(L[I]);
          else
            Res.Append(L[I]);
          end;
        Inc(I);
      end;
      Result := Res.ToString;
    finally
      Res.Free;
    end;
  end;

  /// <summary>Consumes one comma-separated segment of a uses clause.
  ///  Handles the dpr-style "UnitName in 'file.pas'" form by cutting
  ///  at the ' in ' keyword.</summary>
  procedure AddEntry(const ARaw: string);
  var
    T: string;
    E: TUnitUsesEntry;
    InPos: Integer;
  begin
    T := Trim(ARaw);
    InPos := Pos(' IN ', UpperCase(' ' + T + ' '));
    if InPos > 0 then
      T := Trim(Copy(T, 1, InPos - 1));
    if T = '' then Exit;
    if not CharInSet(T[1], ['A'..'Z', 'a'..'z', '_']) then Exit;
    E.UnitName := T;
    E.InInterface := not InImpl;
    E.Line := LineNo;
    Entries.Add(E);
  end;

  /// <summary>Processes the (cleaned) remainder of a line while inside
  ///  a uses clause. Returns True when the terminating ';' was seen.</summary>
  function ConsumeUsesText(const S: string): Boolean;
  var
    SemiPos, P, Start: Integer;
    Segment: string;
  begin
    SemiPos := Pos(';', S);
    if SemiPos > 0 then
      Segment := Copy(S, 1, SemiPos - 1)
    else
      Segment := S;
    // split at commas
    Start := 1;
    for P := 1 to Length(Segment) + 1 do
      if (P > Length(Segment)) or (Segment[P] = ',') then
      begin
        AddEntry(Copy(Segment, Start, P - Start));
        Start := P + 1;
      end;
    Result := SemiPos > 0;
  end;

var
  RawLine, L, U, T: string;
  I, UsesPos: Integer;
begin
  Entries := TList<TUnitUsesEntry>.Create;
  try
    Lines := AContent.Replace(#13#10, #10).Replace(#13, #10)
      .Split([#10], TStringSplitOptions.None);
    InBrace := False;
    InParenStar := False;
    InImpl := False;
    Collecting := False;
    for I := 0 to High(Lines) do
    begin
      LineNo := I + 1;
      RawLine := Lines[I];
      L := CleanLine(RawLine);
      U := UpperCase(L);
      if Collecting then
      begin
        if ConsumeUsesText(L) then Collecting := False;
        Continue;
      end;
      T := Trim(U);
      if (T = 'IMPLEMENTATION') or T.StartsWith('IMPLEMENTATION ') then
        InImpl := True;
      // 'uses' with word boundaries, anywhere in the line
      // ("implementation uses X;" happens in the wild).
      UsesPos := Pos('USES', U);
      while UsesPos > 0 do
      begin
        var BeforeOk := (UsesPos = 1) or not IsIdentChar(U[UsesPos - 1]);
        var AfterIdx := UsesPos + 4;
        var AfterOk := (AfterIdx > Length(U)) or not IsIdentChar(U[AfterIdx]);
        if BeforeOk and AfterOk then
        begin
          Collecting := not ConsumeUsesText(Copy(L, AfterIdx, MaxInt));
          Break;
        end;
        UsesPos := PosEx('USES', U, UsesPos + 1);
      end;
    end;
    Result := Entries.ToArray;
  finally
    Entries.Free;
  end;
end;

{ TUsesCycleResult }

constructor TUsesCycleResult.Create;
begin
  inherited Create;
  FNameToIdx := TDictionary<string, Integer>.Create;
  FCompToGroup := TDictionary<Integer, Integer>.Create;
end;

destructor TUsesCycleResult.Destroy;
var
  I: Integer;
begin
  for I := 0 to High(FAdjTo) do
  begin
    FAdjTo[I].Free;
    FAdjIntf[I].Free;
    FAdjLine[I].Free;
  end;
  FCompToGroup.Free;
  FNameToIdx.Free;
  inherited;
end;

function TUsesCycleResult.ComponentGirth(const AMembers: TArray<Integer>): Integer;
// Girth = length (in edges = in units) of the shortest directed cycle
// among AMembers. For each start node u: BFS shortest distances, then a
// cycle through u is (shortest u->x) + edge x->u. Minimum over all u.
var
  Dist: TArray<Integer>;
  Queue: TQueue<Integer>;
  U, X, W, K: Integer;
begin
  Result := MaxInt;
  SetLength(Dist, Length(FNames));
  Queue := TQueue<Integer>.Create;
  try
    for U in AMembers do
    begin
      for K := 0 to High(Dist) do Dist[K] := -1;
      Dist[U] := 0;
      Queue.Clear;
      Queue.Enqueue(U);
      while Queue.Count > 0 do
      begin
        X := Queue.Dequeue;
        for K := 0 to FAdjTo[X].Count - 1 do
        begin
          W := FAdjTo[X][K];
          if Dist[W] < 0 then
          begin
            Dist[W] := Dist[X] + 1;
            Queue.Enqueue(W);
          end;
        end;
      end;
      // Any edge x -> U closes a cycle of length Dist[x] + 1.
      for X in AMembers do
        if (X <> U) and (Dist[X] >= 0) then
          for K := 0 to FAdjTo[X].Count - 1 do
            if FAdjTo[X][K] = U then
            begin
              if Dist[X] + 1 < Result then Result := Dist[X] + 1;
              Break;
            end;
    end;
  finally
    Queue.Free;
  end;
  if Result = MaxInt then Result := 0;
end;

function TUsesCycleResult.GroupInfos: TArray<TCycleGroupInfo>;
var
  MembersByComp: TDictionary<Integer, TList<Integer>>;
  I, Comp, Grp: Integer;
  List: TList<TCycleGroupInfo>;
  EdgeCountByGroup: TDictionary<Integer, Integer>;
  E: TCycleEdge;
  Pair: TPair<Integer, TList<Integer>>;
  Info: TCycleGroupInfo;
  C: Integer;
begin
  MembersByComp := TDictionary<Integer, TList<Integer>>.Create;
  EdgeCountByGroup := TDictionary<Integer, Integer>.Create;
  List := TList<TCycleGroupInfo>.Create;
  try
    // Group members by component (only components that map to a group).
    for I := 0 to High(FComp) do
    begin
      Comp := FComp[I];
      if not FCompToGroup.ContainsKey(Comp) then Continue;
      if not MembersByComp.ContainsKey(Comp) then
        MembersByComp.Add(Comp, TList<Integer>.Create);
      MembersByComp[Comp].Add(I);
    end;

    for E in FEdges do
    begin
      EdgeCountByGroup.TryGetValue(E.Group, C);
      EdgeCountByGroup.AddOrSetValue(E.Group, C + 1);
    end;

    for Pair in MembersByComp do
    begin
      Grp := FCompToGroup[Pair.Key];
      Info.Group := Grp;
      Info.UnitCount := FCompSize[Pair.Key];
      Info.ShortestCycle := ComponentGirth(Pair.Value.ToArray);
      EdgeCountByGroup.TryGetValue(Grp, C);
      Info.EdgeCount := C;
      List.Add(Info);
    end;

    List.Sort(TComparer<TCycleGroupInfo>.Construct(
      function(const L, R: TCycleGroupInfo): Integer
      begin
        Result := L.Group - R.Group;
      end));
    Result := List.ToArray;
  finally
    for Pair in MembersByComp do
      Pair.Value.Free;
    List.Free;
    EdgeCountByGroup.Free;
    MembersByComp.Free;
  end;
end;

function TUsesCycleResult.ShortestCyclePath(
  const AStart: TCycleEdge): TArray<TCycleHop>;
var
  FromIdx, ToIdx: Integer;
  Queue: TQueue<Integer>;
  Parent, ParentEdge: TArray<Integer>;
  Hops: TList<TCycleHop>;
  H: TCycleHop;
  Cur, K, EdgeIdx: Integer;
  RevPath: TList<Integer>;   // node sequence ToIdx..FromIdx
begin
  Result := nil;
  if not FNameToIdx.TryGetValue(UpperCase(AStart.FromUnit), FromIdx) then Exit;
  if not FNameToIdx.TryGetValue(UpperCase(AStart.ToUnit), ToIdx) then Exit;

  Hops := TList<TCycleHop>.Create;
  Queue := TQueue<Integer>.Create;
  RevPath := TList<Integer>.Create;
  try
    // Hop 0: the clicked edge itself.
    H.FromUnit := AStart.FromUnit; H.FromFile := AStart.FromFile;
    H.ToUnit := AStart.ToUnit; H.InInterface := AStart.InInterface;
    H.Line := AStart.Line;
    Hops.Add(H);

    // BFS ToIdx -> FromIdx over SCC-internal adjacency. Parent/ParentEdge
    // let us reconstruct the path and the exact edge used per hop.
    SetLength(Parent, Length(FNames));
    SetLength(ParentEdge, Length(FNames));
    for K := 0 to High(Parent) do begin Parent[K] := -2; ParentEdge[K] := -1; end;
    Parent[ToIdx] := -1;
    Queue.Enqueue(ToIdx);
    while Queue.Count > 0 do
    begin
      Cur := Queue.Dequeue;
      if Cur = FromIdx then Break;
      for K := 0 to FAdjTo[Cur].Count - 1 do
      begin
        var W := FAdjTo[Cur][K];
        if Parent[W] = -2 then
        begin
          Parent[W] := Cur;
          ParentEdge[W] := K;
          Queue.Enqueue(W);
        end;
      end;
    end;

    if Parent[FromIdx] <> -2 then
    begin
      // Walk parents FromIdx -> ToIdx, collect, then reverse to get
      // the forward ToIdx -> FromIdx order.
      Cur := FromIdx;
      while Cur <> ToIdx do
      begin
        RevPath.Add(Cur);
        Cur := Parent[Cur];
      end;
      RevPath.Add(ToIdx);
      for K := RevPath.Count - 1 downto 1 do
      begin
        var U := RevPath[K];
        var V := RevPath[K - 1];
        EdgeIdx := ParentEdge[V];  // edge U->V that BFS used to reach V
        H.FromUnit := FNames[U];
        H.FromFile := FFiles[U];
        H.ToUnit := FNames[V];
        if EdgeIdx >= 0 then
        begin
          H.InInterface := FAdjIntf[U][EdgeIdx];
          H.Line := FAdjLine[U][EdgeIdx];
        end
        else
        begin
          H.InInterface := False;
          H.Line := 0;
        end;
        Hops.Add(H);
      end;
    end;

    Result := Hops.ToArray;
  finally
    RevPath.Free;
    Queue.Free;
    Hops.Free;
  end;
end;

function TUsesCycleResult.Hotspots: TArray<TUnitHotspot>;
var
  Count: TDictionary<Integer, Integer>;
  E: TCycleEdge;
  FromIdx, ToIdx, C: Integer;
  List: TList<TUnitHotspot>;
  Pair: TPair<Integer, Integer>;
  H: TUnitHotspot;
begin
  Count := TDictionary<Integer, Integer>.Create;
  List := TList<TUnitHotspot>.Create;
  try
    // Every FEdges entry is one SCC-internal edge; count incidence at
    // both endpoints (in + out degree within cycles).
    for E in FEdges do
    begin
      if not FNameToIdx.TryGetValue(UpperCase(E.FromUnit), FromIdx) then Continue;
      if not FNameToIdx.TryGetValue(UpperCase(E.ToUnit), ToIdx) then Continue;
      Count.TryGetValue(FromIdx, C); Count.AddOrSetValue(FromIdx, C + 1);
      Count.TryGetValue(ToIdx, C);   Count.AddOrSetValue(ToIdx, C + 1);
    end;

    for Pair in Count do
    begin
      H.UnitName := FNames[Pair.Key];
      H.FileName := FFiles[Pair.Key];
      H.Group := 0;  // filled below from an edge of the same component
      H.GroupSize := FCompSize[FComp[Pair.Key]];
      H.EdgeCount := Pair.Value;
      List.Add(H);
    end;

    // Attach the display group number by matching component via edges.
    for var I := 0 to List.Count - 1 do
    begin
      H := List[I];
      if FNameToIdx.TryGetValue(UpperCase(H.UnitName), FromIdx) then
        for E in FEdges do
          if SameText(E.FromUnit, H.UnitName) or SameText(E.ToUnit, H.UnitName) then
          begin
            H.Group := E.Group;
            List[I] := H;
            Break;
          end;
    end;

    List.Sort(TComparer<TUnitHotspot>.Construct(
      function(const L, R: TUnitHotspot): Integer
      begin
        Result := R.EdgeCount - L.EdgeCount;               // desc by count
        if Result = 0 then Result := L.Group - R.Group;
        if Result = 0 then Result := CompareText(L.UnitName, R.UnitName);
      end));
    Result := List.ToArray;
  finally
    List.Free;
    Count.Free;
  end;
end;

function TUsesCycleResult.CountCycleNodes(AExFrom, AExTo: Integer): Integer;
var
  N: Integer;
  Idx, Lowlink, Comp: TArray<Integer>;
  OnStk: TArray<Boolean>;
  Stk: TStack<Integer>;
  Counter, CompCount, I: Integer;
  Sizes: TArray<Integer>;

  procedure SC(V: Integer);
  var K, W: Integer;
  begin
    Idx[V] := Counter; Lowlink[V] := Counter; Inc(Counter);
    Stk.Push(V); OnStk[V] := True;
    for K := 0 to FAdjTo[V].Count - 1 do
    begin
      W := FAdjTo[V][K];
      if (V = AExFrom) and (W = AExTo) then Continue;  // excluded dependency
      if Idx[W] < 0 then
      begin
        SC(W);
        if Lowlink[W] < Lowlink[V] then Lowlink[V] := Lowlink[W];
      end
      else if OnStk[W] then
        if Idx[W] < Lowlink[V] then Lowlink[V] := Idx[W];
    end;
    if Lowlink[V] = Idx[V] then
      repeat
        W := Stk.Pop; OnStk[W] := False; Comp[W] := CompCount;
      until W = V;
    if Lowlink[V] = Idx[V] then Inc(CompCount);
  end;

begin
  N := Length(FNames);
  SetLength(Idx, N); SetLength(Lowlink, N); SetLength(Comp, N); SetLength(OnStk, N);
  for I := 0 to N - 1 do Idx[I] := -1;
  Counter := 0; CompCount := 0;
  Stk := TStack<Integer>.Create;
  try
    for I := 0 to N - 1 do
      if Idx[I] < 0 then SC(I);
  finally
    Stk.Free;
  end;
  SetLength(Sizes, CompCount);
  for I := 0 to N - 1 do Inc(Sizes[Comp[I]]);
  Result := 0;
  for I := 0 to N - 1 do
    if Sizes[Comp[I]] >= 2 then Inc(Result);
end;

function TUsesCycleResult.EdgeLevers: TArray<TEdgeLever>;
var
  Baseline: Integer;
  Seen: TDictionary<string, TEdgeLever>;
  List: TList<TEdgeLever>;
  E: TCycleEdge;
  L: TEdgeLever;
  Key: string;
  FromIdx, ToIdx: Integer;
begin
  Baseline := CountCycleNodes(-1, -1);
  Seen := TDictionary<string, TEdgeLever>.Create;
  List := TList<TEdgeLever>.Create;
  try
    // One lever per unique (from -> to) dependency, aggregating the
    // interface / implementation sections.
    for E in FEdges do
    begin
      Key := UpperCase(E.FromUnit) + '>' + UpperCase(E.ToUnit);
      if Seen.TryGetValue(Key, L) then
      begin
        if E.InInterface and (L.Section = 'implementation') then L.Section := 'both'
        else if (not E.InInterface) and (L.Section = 'interface') then L.Section := 'both';
        Seen[Key] := L;
      end
      else
      begin
        L.FromUnit := E.FromUnit;
        L.FromFile := E.FromFile;
        L.ToUnit := E.ToUnit;
        if E.InInterface then L.Section := 'interface' else L.Section := 'implementation';
        L.Line := E.Line;
        L.Group := E.Group;
        L.UnitsFreed := 0;
        Seen.Add(Key, L);
      end;
    end;

    for L in Seen.Values do
    begin
      if not FNameToIdx.TryGetValue(UpperCase(L.FromUnit), FromIdx) then Continue;
      if not FNameToIdx.TryGetValue(UpperCase(L.ToUnit), ToIdx) then Continue;
      var Rec := L;
      Rec.UnitsFreed := Baseline - CountCycleNodes(FromIdx, ToIdx);
      List.Add(Rec);
    end;

    List.Sort(TComparer<TEdgeLever>.Construct(
      function(const A, B: TEdgeLever): Integer
      begin
        Result := B.UnitsFreed - A.UnitsFreed;              // desc
        if Result = 0 then Result := A.Group - B.Group;
        if Result = 0 then Result := CompareText(A.FromUnit, B.FromUnit);
      end));
    Result := List.ToArray;
  finally
    List.Free;
    Seen.Free;
  end;
end;

function TUsesCycleResult.EnumerateCycles(AMax: Integer;
  out ATruncated: Boolean): TArray<TCyclePath>;
var
  Cycles: TList<TCyclePath>;
  Path: TList<Integer>;
  OnPath: TArray<Boolean>;
  Start, N, I: Integer;
  Stop: Boolean;

  procedure DFS(V: Integer);
  var K, W, J: Integer;
  begin
    if Stop then Exit;
    for K := 0 to FAdjTo[V].Count - 1 do
    begin
      W := FAdjTo[V][K];
      if W < Start then Continue;             // enforce min-node rooting
      if W = Start then
      begin
        if Path.Count >= 2 then
        begin
          var CP: TCyclePath;
          SetLength(CP.Units, Path.Count);
          for J := 0 to Path.Count - 1 do CP.Units[J] := FNames[Path[J]];
          Cycles.Add(CP);
          if Cycles.Count >= AMax then
          begin
            Stop := True; ATruncated := True; Exit;
          end;
        end;
      end
      else if (W > Start) and (not OnPath[W]) then
      begin
        OnPath[W] := True; Path.Add(W);
        DFS(W);
        Path.Delete(Path.Count - 1); OnPath[W] := False;
        if Stop then Exit;
      end;
    end;
  end;

begin
  ATruncated := False;
  Stop := False;
  N := Length(FNames);
  Cycles := TList<TCyclePath>.Create;
  Path := TList<Integer>.Create;
  SetLength(OnPath, N);
  try
    for Start := 0 to N - 1 do
    begin
      if FAdjTo[Start].Count = 0 then Continue;
      for I := 0 to N - 1 do OnPath[I] := False;
      Path.Clear;
      Path.Add(Start); OnPath[Start] := True;
      DFS(Start);
      if Stop then Break;
    end;
    Result := Cycles.ToArray;
  finally
    Path.Free;
    Cycles.Free;
  end;
end;

class function TUsesGraphAnalyzer.Analyze(const AFiles: TArray<string>;
  const AProgress: TProc<Integer, Integer, string>): TUsesCycleResult;
type
  TRawEdge = record
    FromIdx, ToIdx: Integer;
    InInterface: Boolean;
    Line: Integer;
  end;
var
  R: TUsesCycleResult;
  Adj: TArray<TList<Integer>>;       // node idx -> adjacent node idxs (for SCC)
  RawEdges: TList<TRawEdge>;
  Content: string;
  Entries: TArray<TUnitUsesEntry>;
  I, J, N, FromIdx, ToIdx: Integer;
  // Tarjan state
  Index, Lowlink, Comp: TArray<Integer>;
  OnStack: TArray<Boolean>;
  Stack: TStack<Integer>;
  Counter, CompCount: Integer;
  CompSize: TArray<Integer>;

  procedure StrongConnect(V: Integer);
  var
    W: Integer;
  begin
    Index[V] := Counter;
    Lowlink[V] := Counter;
    Inc(Counter);
    Stack.Push(V);
    OnStack[V] := True;
    for W in Adj[V] do
    begin
      if Index[W] < 0 then
      begin
        StrongConnect(W);
        if Lowlink[W] < Lowlink[V] then Lowlink[V] := Lowlink[W];
      end
      else if OnStack[W] then
        if Index[W] < Lowlink[V] then Lowlink[V] := Index[W];
    end;
    if Lowlink[V] = Index[V] then
    begin
      repeat
        W := Stack.Pop;
        OnStack[W] := False;
        Comp[W] := CompCount;
      until W = V;
      Inc(CompCount);
    end;
  end;

var
  Edges: TList<TCycleEdge>;
  E: TCycleEdge;
  GroupRenumber: TDictionary<Integer, Integer>;
  NextGroup: Integer;
begin
  R := TUsesCycleResult.Create;

  // ---- collect nodes (project .pas units) ----
  N := 0;
  SetLength(R.FNames, Length(AFiles));
  SetLength(R.FFiles, Length(AFiles));
  for I := 0 to High(AFiles) do
  begin
    if not SameText(ExtractFileExt(AFiles[I]), '.pas') then Continue;
    var UName := ChangeFileExt(ExtractFileName(AFiles[I]), '');
    if R.FNameToIdx.ContainsKey(UpperCase(UName)) then Continue;
    R.FNames[N] := UName;
    R.FFiles[N] := AFiles[I];
    R.FNameToIdx.Add(UpperCase(UName), N);
    Inc(N);
  end;
  SetLength(R.FNames, N);
  SetLength(R.FFiles, N);

  SetLength(Adj, N);
  for I := 0 to N - 1 do
    Adj[I] := TList<Integer>.Create;
  RawEdges := TList<TRawEdge>.Create;
  Edges := TList<TCycleEdge>.Create;
  Stack := TStack<Integer>.Create;
  GroupRenumber := TDictionary<Integer, Integer>.Create;
  try
    // ---- collect edges ----
    for I := 0 to N - 1 do
    begin
      if Assigned(AProgress) then
        AProgress(I + 1, N, R.FFiles[I]);
      if not ReadFileContent(R.FFiles[I], Content) then Continue;
      Entries := ParseUsesEntries(Content);
      for J := 0 to High(Entries) do
        if R.FNameToIdx.TryGetValue(UpperCase(Entries[J].UnitName), ToIdx) then
        begin
          FromIdx := I;
          if FromIdx = ToIdx then Continue;
          Adj[FromIdx].Add(ToIdx);
          var RE: TRawEdge;
          RE.FromIdx := FromIdx;
          RE.ToIdx := ToIdx;
          RE.InInterface := Entries[J].InInterface;
          RE.Line := Entries[J].Line;
          RawEdges.Add(RE);
        end;
    end;

    // ---- Tarjan SCC ----
    SetLength(Index, N);
    SetLength(Lowlink, N);
    SetLength(Comp, N);
    SetLength(OnStack, N);
    for I := 0 to N - 1 do Index[I] := -1;
    Counter := 0;
    CompCount := 0;
    for I := 0 to N - 1 do
      if Index[I] < 0 then
        StrongConnect(I);

    SetLength(CompSize, CompCount);
    for I := 0 to N - 1 do
      Inc(CompSize[Comp[I]]);

    R.FComp := Comp;
    R.FCompSize := CompSize;

    // ---- SCC-internal adjacency with edge metadata (for path queries) ----
    SetLength(R.FAdjTo, N);
    SetLength(R.FAdjIntf, N);
    SetLength(R.FAdjLine, N);
    for I := 0 to N - 1 do
    begin
      R.FAdjTo[I] := TList<Integer>.Create;
      R.FAdjIntf[I] := TList<Boolean>.Create;
      R.FAdjLine[I] := TList<Integer>.Create;
    end;

    // ---- emit edges inside multi-node components ----
    NextGroup := 0;
    for I := 0 to RawEdges.Count - 1 do
    begin
      var RE := RawEdges[I];
      if Comp[RE.FromIdx] <> Comp[RE.ToIdx] then Continue;
      if CompSize[Comp[RE.FromIdx]] < 2 then Continue;
      var GroupId: Integer;
      if not GroupRenumber.TryGetValue(Comp[RE.FromIdx], GroupId) then
      begin
        Inc(NextGroup);
        GroupId := NextGroup;
        GroupRenumber.Add(Comp[RE.FromIdx], GroupId);
        R.FCompToGroup.Add(Comp[RE.FromIdx], GroupId);
      end;
      E.FromUnit := R.FNames[RE.FromIdx];
      E.FromFile := R.FFiles[RE.FromIdx];
      E.ToUnit := R.FNames[RE.ToIdx];
      E.InInterface := RE.InInterface;
      E.Line := RE.Line;
      E.Group := GroupId;
      E.GroupSize := CompSize[Comp[RE.FromIdx]];
      Edges.Add(E);

      R.FAdjTo[RE.FromIdx].Add(RE.ToIdx);
      R.FAdjIntf[RE.FromIdx].Add(RE.InInterface);
      R.FAdjLine[RE.FromIdx].Add(RE.Line);
    end;

    R.FEdges := Edges.ToArray;
    Result := R;
  finally
    GroupRenumber.Free;
    Stack.Free;
    Edges.Free;
    RawEdges.Free;
    for I := 0 to N - 1 do
      Adj[I].Free;
  end;
end;

end.
