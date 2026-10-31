unit Expert.IntfVisualizer;

(*
  Resolving an INTERFACE REFERENCE to the object behind it, and reading the
  class hierarchy with the interfaces each level implements - from the RTTI
  BYTES, with the names, and without calling anything in the inspected
  process. That is what a debug visualizer needs: it may read memory, it may
  not run code.

  EVERY LAYOUT BELOW IS READ OUT OF THE SHIPPED RTL (37.0), never recalled:

  * the way back from a reference to the object is the RTL's own, in
    TObject.GetInterface:
        Pointer(Obj) := Pointer(PByte(Self) + InterfaceEntry^.IOffset);
    so Obj = Reference - IOffset. The reference's target holds the entry's
    VTable pointer.
  * System.pas, the interface table:
        TInterfaceTable = packed record
          EntryCount: Integer;
          {$IF defined(CPU64BITS)} _Filler: UInt32; {$ENDIF}
          Entries: array[0..9999] of TInterfaceEntry;
         {Intfs: array[0..EntryCount - 1] of PPTypeInfo;}
        end;
    and the commented-out array is real - System.Rtti's own
    TRttiInstanceType.GetDeclaredImplementedInterfaces reads it:
        p := @pi^.Entries[pi^.EntryCount];
        for i := 0 to pi.EntryCount - 1 do ... DerefPointer(ReadPointer(p))
    THAT is where an interface's NAME comes from: one PPTypeInfo per entry,
    and a PTypeInfo starts with Kind + Name: TSymbolName (a UTF-8
    ShortString).
  * TTypeInfo = record Kind: TTypeKind; Name: TSymbolName; end, and for a
    class the type data is
        tkClass: (ClassType: TClass; ParentInfo: PPTypeInfo;
                  PropCount: SmallInt; UnitName: TSymbolName; ...)
    which is where the unit name comes from.
  * and the VMT offsets are NOT constants of this plugin's own bitness:
        vmtSelfPtr = -176 - vmtArcOffset - CPP_ABI_ADJUST;   // CPU64BITS
        vmtSelfPtr =  -88 - vmtArcOffset - CPP_ABI_ADJUST;   // 32 bit
    with vmtArcOffset = 0 off the ARC platforms and - the one that is easy
    to get wrong - CPP_ABI_ADJUST = 3 * SizeOf(Pointer) because
    CPP_ABI_SUPPORT is defined for WIN64. So a 64-bit debuggee has
    vmtSelfPtr = -200, not -176.
    Nothing here assumes which set applies: both are TRIED and vmtSelfPtr
    decides, because that slot must hold the VMT's own address. A layout
    this code guessed wrong can then not be mistaken for an answer.

  POLARITY, the rule this whole plugin follows: what the bytes do not prove
  is not reported. Every exit says WHY in Reason, and a name that does not
  read as an identifier is left empty rather than printed as garbage.
*)

interface

uses
  System.SysUtils;

type
  /// <summary>One read in the inspected process. False when the address is
  ///  not readable - which is the normal answer while searching, not an
  ///  error. ABuffer takes ACount bytes.</summary>
  TDebugMemReader = reference to function (AAddress: UInt64; ACount: Integer;
    ABuffer: Pointer): Boolean;

  /// <summary>One entry of one class's interface table.</summary>
  TIntfEntryInfo = record
    /// <summary>From the entry's own PPTypeInfo - empty when those bytes do
    ///  not read as an identifier.</summary>
    Name: string;
    IID: TGUID;
    IOffset: Integer;
    VTable: UInt64;
    /// <summary>True for the ONE entry the inspected reference points at.
    ///  That is what identifies the reference, and it is proven rather than
    ///  assumed: the entry's VTable and IOffset both had to match.</summary>
    IsTheReference: Boolean;
  end;

  /// <summary>One level of the hierarchy, most derived first.</summary>
  TIntfClassLevel = record
    ClassName: string;
    /// <summary>From the class's type data; empty when it is not readable.</summary>
    UnitName: string;
    Vmt: UInt64;
    InstanceSize: Integer;
    /// <summary>The interfaces THIS level declares - a class's interface
    ///  table holds its own entries, not its ancestors'.</summary>
    Interfaces: TArray<TIntfEntryInfo>;
  end;

  TIntfRefResolution = record
    Resolved: Boolean;
    /// <summary>Why not, when not - and also what was partial.</summary>
    Reason: string;
    Reference: UInt64;
    VTable: UInt64;
    ObjectAddr: UInt64;
    IOffset: Integer;
    /// <summary>The pointer size vmtSelfPtr confirmed, 4 or 8.</summary>
    PointerSize: Integer;
    Levels: TArray<TIntfClassLevel>;
    /// <summary>The name of the interface the reference points at, '' when
    ///  the entry carried no readable name.</summary>
    function ReferenceInterface: string;
  end;

const
  /// <summary>How far back the object is looked for. The reference points at
  ///  a FIELD of the object, so this is a field offset - 4 KB is far beyond
  ///  any real class and still one cheap read.</summary>
  MaxIntfFieldOffset = 4096;
  /// <summary>Guards against a cyclic or garbage parent chain.</summary>
  MaxClassLevels = 64;
  /// <summary>Guards against a garbage EntryCount.</summary>
  MaxIntfEntries = 512;

/// <summary>The vmt offsets for a given pointer size, as System.pas computes
///  them for a Windows target (vmtArcOffset = 0, CPP_ABI_ADJUST =
///  3 * SizeOf(Pointer) on 64 bit).</summary>
procedure VmtOffsets(APointerSize: Integer; out ASelfPtr, AIntfTable,
  ATypeInfo, AClassName, AInstanceSize, AParent: Integer);

/// <summary>Is AAddress a Delphi VMT? Proven by its own self-pointer slot,
///  which must hold AAddress.</summary>
function IsClassVmt(AAddress: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader): Boolean;

/// <summary>Reads a UTF-8 ShortString (one length byte, then the bytes) and
///  answers '' when it does not read as an identifier.</summary>
function ReadSymbolName(AAddress: UInt64; const ARead: TDebugMemReader): string;

/// <summary>The whole answer for one interface reference.</summary>
function ResolveInterfaceReference(AReference: UInt64;
  const ARead: TDebugMemReader): TIntfRefResolution;

implementation

procedure VmtOffsets(APointerSize: Integer; out ASelfPtr, AIntfTable,
  ATypeInfo, AClassName, AInstanceSize, AParent: Integer);
var
  Adjust: Integer;
begin
  if APointerSize = 8 then
  begin
    // CPP_ABI_SUPPORT is defined for WIN64, so CPP_ABI_ADJUST is 3 pointers
    Adjust := 3 * 8;
    ASelfPtr := -176 - Adjust;
    AIntfTable := -168 - Adjust;
    ATypeInfo := -144 - Adjust;
    AClassName := -112 - Adjust;
    AInstanceSize := -104 - Adjust;
    AParent := -96 - Adjust;
  end
  else
  begin
    ASelfPtr := -88;
    AIntfTable := -84;
    ATypeInfo := -72;
    AClassName := -56;
    AInstanceSize := -52;
    AParent := -48;
  end;
end;

// A vmt offset is NEGATIVE, so "Base + UInt64(Offset)" only works through
// sign extension and wrap-around - which compiles with a signed/unsigned
// warning on 64 bit and reads like a mistake. This says what it means.
function AtOffset(ABase: UInt64; AOffset: Integer): UInt64;
begin
  Result := UInt64(Int64(ABase) + AOffset);
end;

function TryReadPtr(AAddress: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader; out AValue: UInt64): Boolean;
var
  V32: UInt32;
  V64: UInt64;
begin
  AValue := 0;
  if AAddress = 0 then Exit(False);
  if APointerSize = 8 then
  begin
    Result := ARead(AAddress, SizeOf(V64), @V64);
    if Result then AValue := V64;
  end
  else
  begin
    V32 := 0;
    Result := ARead(AAddress, SizeOf(V32), @V32);
    if Result then AValue := V32;
  end;
end;

function TryReadInt32(AAddress: UInt64; const ARead: TDebugMemReader;
  out AValue: Integer): Boolean;
var
  V: Int32;
begin
  AValue := 0;
  V := 0;
  Result := (AAddress <> 0) and ARead(AAddress, SizeOf(V), @V);
  if Result then AValue := V;
end;

// The parent slot is a DOUBLE indirection, which is the RTL's own reading in
// TObject.ClassParent:
//     Result := PPointer(@PByte(Self)[vmtParent])^;
//     if Result <> nil then Result := PPointer(Result)^;
// Read it once and the answer is the intermediate pointer, which is no VMT -
// the walk then stops at the most derived class and every level above it is
// silently missing. That is what the first run of the fixture reported.
function TryReadParentVmt(AVmt: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader; out AParent: UInt64): Boolean;
var
  S1, S2, S3, S4, S5, ParentOfs: Integer;
  Slot: UInt64;
begin
  AParent := 0;
  VmtOffsets(APointerSize, S1, S2, S3, S4, S5, ParentOfs);
  Result := TryReadPtr(AtOffset(AVmt, ParentOfs), APointerSize, ARead, Slot);
  if not Result then Exit;
  if Slot = 0 then Exit(True);          // TObject: no parent, and that is no error
  Result := TryReadPtr(Slot, APointerSize, ARead, AParent);
end;

function IsClassVmt(AAddress: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader): Boolean;
var
  SelfPtr, I1, I2, I3, I4, I5: Integer;
  Slot: UInt64;
begin
  Result := False;
  if (AAddress = 0) or (AAddress mod UInt64(APointerSize) <> 0) then Exit;
  VmtOffsets(APointerSize, SelfPtr, I1, I2, I3, I4, I5);
  // the slot lies BEFORE the VMT, so the address must be large enough to
  // have one at all - a small value is no class and no readable memory
  if AAddress < UInt64(-SelfPtr) then Exit;
  if not TryReadPtr(AtOffset(AAddress, SelfPtr), APointerSize, ARead, Slot) then Exit;
  Result := Slot = AAddress;
end;

function ReadSymbolName(AAddress: UInt64; const ARead: TDebugMemReader): string;
var
  Len: Byte;
  Buf: TBytes;
  I: Integer;
  C: Byte;
begin
  Result := '';
  if AAddress = 0 then Exit;
  Len := 0;
  if not ARead(AAddress, 1, @Len) then Exit;
  if Len = 0 then Exit;
  SetLength(Buf, Len);
  if not ARead(AAddress + 1, Len, @Buf[0]) then Exit;
  // A name that is not an identifier means the layout assumption was wrong
  // somewhere - and then saying nothing is the only honest answer. '.' is
  // allowed because a unit name carries it; '<' and ',' because a generic
  // instantiation does ("TList<System.Integer>").
  for I := 0 to Len - 1 do
  begin
    C := Buf[I];
    if not (((C >= Ord('0')) and (C <= Ord('9'))) or
            ((C >= Ord('A')) and (C <= Ord('Z'))) or
            ((C >= Ord('a')) and (C <= Ord('z'))) or
            (C = Ord('_')) or (C = Ord('.')) or (C = Ord('<')) or
            (C = Ord('>')) or (C = Ord(',')) or (C = Ord(' ')) or
            (C >= $80)) then
      Exit('');
  end;
  Result := TEncoding.UTF8.GetString(Buf);
end;

{ the interface table of ONE class }

function ReadInterfaceTable(AVmt: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader; out AEntries: TArray<TIntfEntryInfo>): Boolean;
var
  SelfPtr, IntfTable, I2, I3, I4, I5: Integer;
  Table, NamePtrAddr, TypeInfoPtr, TypeInfoAddr: UInt64;
  Count, I, EntrySize, HeaderSize: Integer;
  Base: UInt64;
  G: TGUID;
  VT: UInt64;
  IOff: Integer;
begin
  AEntries := nil;
  Result := False;
  VmtOffsets(APointerSize, SelfPtr, IntfTable, I2, I3, I4, I5);
  if not TryReadPtr(AtOffset(AVmt, IntfTable), APointerSize, ARead, Table) then Exit;
  if Table = 0 then Exit(True);      // a class with no interfaces of its own
  if not TryReadInt32(Table, ARead, Count) then Exit;
  if (Count <= 0) or (Count > MaxIntfEntries) then Exit;

  // packed record, with an explicit filler on 64 bit:
  //   header  = EntryCount [+ _Filler]
  //   entry   = IID(16) + VTable(ptr) + IOffset(4) [+ _Filler(4)] + ImplGetter(ptr)
  if APointerSize = 8 then
  begin
    HeaderSize := 8;
    EntrySize := 16 + 8 + 4 + 4 + 8;
  end
  else
  begin
    HeaderSize := 4;
    EntrySize := 16 + 4 + 4 + 4;
  end;

  SetLength(AEntries, Count);
  Base := Table + UInt64(HeaderSize);
  for I := 0 to Count - 1 do
  begin
    if not ARead(Base + UInt64(I * EntrySize), SizeOf(G), @G) then Exit;
    if not TryReadPtr(Base + UInt64(I * EntrySize) + 16, APointerSize, ARead, VT) then Exit;
    if not TryReadInt32(Base + UInt64(I * EntrySize) + 16 + UInt64(APointerSize),
      ARead, IOff) then Exit;
    AEntries[I].IID := G;
    AEntries[I].VTable := VT;
    AEntries[I].IOffset := IOff;
  end;

  // ...and behind the entries, one PPTypeInfo per entry - the names
  NamePtrAddr := Base + UInt64(Count * EntrySize);
  for I := 0 to Count - 1 do
  begin
    if not TryReadPtr(NamePtrAddr + UInt64(I * APointerSize), APointerSize,
      ARead, TypeInfoPtr) then Continue;
    if TypeInfoPtr = 0 then Continue;
    if not TryReadPtr(TypeInfoPtr, APointerSize, ARead, TypeInfoAddr) then Continue;
    if TypeInfoAddr = 0 then Continue;
    // PTypeInfo: Kind (1 byte), then Name as a ShortString
    AEntries[I].Name := ReadSymbolName(TypeInfoAddr + 1, ARead);
  end;
  Result := True;
end;

function ReadUnitName(AVmt: UInt64; APointerSize: Integer;
  const ARead: TDebugMemReader): string;
var
  SelfPtr, I1, TypeInfoOfs, I3, I4, I5: Integer;
  Info, NameAddr: UInt64;
  Len: Byte;
begin
  Result := '';
  VmtOffsets(APointerSize, SelfPtr, I1, TypeInfoOfs, I3, I4, I5);
  if not TryReadPtr(AtOffset(AVmt, TypeInfoOfs), APointerSize, ARead, Info) then Exit;
  if Info = 0 then Exit;
  // PTypeInfo = Kind(1) + Name(ShortString) ; then TTypeData, whose tkClass
  // layout is ClassType(ptr) + ParentInfo(ptr) + PropCount(2) + UnitName
  Len := 0;
  if not ARead(Info + 1, 1, @Len) then Exit;
  NameAddr := Info + 1 + 1 + Len;                     // start of TTypeData
  NameAddr := NameAddr + UInt64(2 * APointerSize) + 2;
  Result := ReadSymbolName(NameAddr, ARead);
end;

function TIntfRefResolution.ReferenceInterface: string;
begin
  Result := '';
  for var L in Levels do
    for var E in L.Interfaces do
      if E.IsTheReference then Exit(E.Name);
end;

function ResolveInterfaceReference(AReference: UInt64;
  const ARead: TDebugMemReader): TIntfRefResolution;
var
  Res: TIntfRefResolution;

  // Walks the parent chain and fills the levels. The entry the reference
  // points at is marked on the way.
  function BuildLevels(AVmt: UInt64; APtrSize: Integer; AFoundOffset: Integer;
    AVTable: UInt64): Boolean;
  var
    SelfPtr, I1, I2, ClassNameOfs, InstSizeOfs, ParentOfs: Integer;
    Cur, NamePtr: UInt64;
    Level: TIntfClassLevel;
    Guard: Integer;
  begin
    Result := False;
    VmtOffsets(APtrSize, SelfPtr, I1, I2, ClassNameOfs, InstSizeOfs, ParentOfs);
    Cur := AVmt;
    Guard := 0;
    while (Cur <> 0) and (Guard < MaxClassLevels) do
    begin
      Inc(Guard);
      if not IsClassVmt(Cur, APtrSize, ARead) then
      begin
        // the chain left the classes - report what was read, and why it stopped
        Res.Reason := 'the parent chain stopped at an address that is no ' +
          'class (read ' + IntToStr(Length(Res.Levels)) + ' level(s))';
        Break;
      end;
      Level := Default(TIntfClassLevel);
      Level.Vmt := Cur;
      if TryReadPtr(AtOffset(Cur, ClassNameOfs), APtrSize, ARead, NamePtr) then
        Level.ClassName := ReadSymbolName(NamePtr, ARead);
      TryReadInt32(AtOffset(Cur, InstSizeOfs), ARead, Level.InstanceSize);
      Level.UnitName := ReadUnitName(Cur, APtrSize, ARead);
      if not ReadInterfaceTable(Cur, APtrSize, ARead, Level.Interfaces) then
        Res.Reason := Trim(Res.Reason + ' The interface table of ' +
          Level.ClassName + ' could not be read.');
      for var I := 0 to High(Level.Interfaces) do
        Level.Interfaces[I].IsTheReference :=
          (Level.Interfaces[I].VTable = AVTable) and
          (Level.Interfaces[I].IOffset = AFoundOffset);
      Res.Levels := Res.Levels + [Level];
      Result := True;
      if not TryReadParentVmt(Cur, APtrSize, ARead, Cur) then Break;
    end;
  end;

  // One bitness: find the object by walking the field offsets back from the
  // reference, and PROVE each candidate - the pointer there must be a VMT,
  // and that class (or one of its ancestors) must declare an entry whose
  // VTable and IOffset are exactly the ones we are looking at.
  function TryPointerSize(APtrSize: Integer): Boolean;
  var
    VTable, Cand, Cur: UInt64;
    Offset, Guard: Integer;
    Entries: TArray<TIntfEntryInfo>;
  begin
    Result := False;
    if not TryReadPtr(AReference, APtrSize, ARead, VTable) then Exit;
    if VTable = 0 then Exit;
    Offset := APtrSize;
    while Offset <= MaxIntfFieldOffset do
    begin
      if AReference > UInt64(Offset) then
      begin
        Cand := 0;
        if TryReadPtr(AReference - UInt64(Offset), APtrSize, ARead, Cand) and
           IsClassVmt(Cand, APtrSize, ARead) then
        begin
          Cur := Cand;
          Guard := 0;
          while (Cur <> 0) and (Guard < MaxClassLevels) and
                IsClassVmt(Cur, APtrSize, ARead) do
          begin
            Inc(Guard);
            if ReadInterfaceTable(Cur, APtrSize, ARead, Entries) then
              for var E in Entries do
                if (E.VTable = VTable) and (E.IOffset = Offset) then
                begin
                  Res.Reference := AReference;
                  Res.VTable := VTable;
                  Res.ObjectAddr := AReference - UInt64(Offset);
                  Res.IOffset := Offset;
                  Res.PointerSize := APtrSize;
                  Res.Levels := nil;
                  Res.Resolved := BuildLevels(Cand, APtrSize, Offset, VTable);
                  Exit(Res.Resolved);
                end;
            if not TryReadParentVmt(Cur, APtrSize, ARead, Cur) then Break;
          end;
        end;
      end;
      Inc(Offset, APtrSize);
    end;
  end;

begin
  Res := Default(TIntfRefResolution);
  Res.Reference := AReference;
  if not Assigned(ARead) then
  begin
    Res.Reason := 'no memory reader';
    Exit(Res);
  end;
  if AReference = 0 then
  begin
    Res.Reason := 'the reference is nil';
    Exit(Res);
  end;

  // A reference whose target is not even readable, or holds no VTable at
  // all, is answered before any search - that is the ordinary "this is not
  // a live interface reference" case and deserves its own sentence.
  var Probe: UInt64 := 0;
  if (not TryReadPtr(AReference, 8, ARead, Probe)) and
     (not TryReadPtr(AReference, 4, ARead, Probe)) then
  begin
    Res.Reason := 'the reference itself is not readable in the inspected ' +
      'process - the variable is out of scope, or it is not a pointer';
    Exit(Res);
  end;
  if Probe = 0 then
  begin
    Res.Reason := 'the reference points at a null VTable, so there is no ' +
      'interface behind it';
    Exit(Res);
  end;

  // Which bitness the inspected process has is not assumed: vmtSelfPtr
  // decides, because only the right offset set makes a VMT prove itself.
  // 8 first - a 32-bit address read as 64 bit would pull in a neighbouring
  // field, and that cannot satisfy the self-pointer test.
  if TryPointerSize(8) then Exit(Res);
  Res := Default(TIntfRefResolution);
  Res.Reference := AReference;
  if TryPointerSize(4) then Exit(Res);

  Res := Default(TIntfRefResolution);
  Res.Reference := AReference;
  Res.Reason := 'no object could be proven behind this reference: within ' +
    IntToStr(MaxIntfFieldOffset) + ' bytes before it there is no address ' +
    'that is a class and declares an entry with this reference''s VTable. ' +
    'That is what a delegating implementation looks like (the interface is ' +
    'provided by another object through an implements clause), and also ' +
    'what a reference that is not a Delphi class interface looks like - a ' +
    'COM object, or a stale pointer.';
  Result := Res;
end;

end.
