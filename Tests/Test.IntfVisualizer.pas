unit Test.IntfVisualizer;

{
  The resolver walks RTTI bytes, so the honest fixture is REAL RTTI: these
  tests resolve an interface reference IN THEIR OWN PROCESS, with a reader
  that is a guarded Move. Everything the assertions expect - the class
  names, which level declares which interface, the interface names - is
  what the compiler really emitted for the classes below.

  That is also what covers the one layout difference this code can get
  wrong: CPP_ABI_SUPPORT is defined for WIN64, so a 64-bit VMT's offsets
  differ from a 32-bit one's by three pointers. The suite runs on both
  platforms, so both offset sets are exercised by construction - and
  nothing in the resolver is told which one applies.
}

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TIntfVisualizerTests = class
  public
    /// <summary>The way back from a reference to its object, proven rather
    ///  than guessed: the candidate must be a class AND declare an entry
    ///  with this reference's VTable and offset.</summary>
    [Test] procedure AReferenceLeadsToItsObject;
    /// <summary>The hierarchy and the interfaces each level declares, with
    ///  the names read out of the PPTypeInfo array behind the entries.</summary>
    [Test] procedure EveryLevelNamesItsOwnInterfaces;
    /// <summary>What the bytes do not prove is not reported.</summary>
    [Test] procedure WhatCannotBeResolvedSaysWhy;
  end;

implementation

uses
  System.SysUtils, System.Classes, Winapi.Windows, Expert.IntfVisualizer;

type
  IOne = interface
    ['{7B1A0C64-6E2E-4B0A-9C31-3A5C9D1E0001}']
    procedure One;
  end;

  ITwo = interface(IOne)
    ['{7B1A0C64-6E2E-4B0A-9C31-3A5C9D1E0002}']
    procedure Two;
  end;

  IThree = interface
    ['{7B1A0C64-6E2E-4B0A-9C31-3A5C9D1E0003}']
    procedure Three;
  end;

  // Three levels on purpose, each declaring its own interfaces: the
  // resolver must attribute them per LEVEL, because a class's interface
  // table holds only its own entries.
  TVisBase = class(TInterfacedObject, IOne)
  public
    procedure One;
  end;

  TVisMiddle = class(TVisBase, ITwo)
  public
    procedure Two;
  end;

  TVisLeaf = class(TVisMiddle, IThree)
  public
    procedure Three;
  end;

procedure TVisBase.One; begin end;
procedure TVisMiddle.Two; begin end;
procedure TVisLeaf.Three; begin end;

{ A reader for this process. It must be GUARDED: the resolver walks
  addresses before the reference looking for the object, so it will ask for
  memory that is not ours. An unreadable address is a normal answer here,
  never an exception - which is exactly how the debugger's own
  ReadProcessMemory behaves. }
function OwnProcessReader: TDebugMemReader;
begin
  Result :=
    function (AAddress: UInt64; ACount: Integer; ABuffer: Pointer): Boolean
    var
      Info: TMemoryBasicInformation;
      Last: UInt64;
    begin
      Result := False;
      if (AAddress = 0) or (ACount <= 0) or (ABuffer = nil) then Exit;
      if AAddress > High(NativeUInt) - UInt64(ACount) then Exit;
      if VirtualQuery(Pointer(NativeUInt(AAddress)), Info,
        SizeOf(Info)) <> SizeOf(Info) then Exit;
      if Info.State <> MEM_COMMIT then Exit;
      if (Info.Protect and (PAGE_NOACCESS or PAGE_GUARD)) <> 0 then Exit;
      // the read must not leave the region VirtualQuery just described
      Last := UInt64(NativeUInt(Info.BaseAddress)) + Info.RegionSize;
      if AAddress + UInt64(ACount) > Last then Exit;
      Move(Pointer(NativeUInt(AAddress))^, ABuffer^, ACount);
      Result := True;
    end;
end;

function LevelNamed(const ARes: TIntfRefResolution;
  const AClass: string): Integer;
begin
  Result := -1;
  for var I := 0 to High(ARes.Levels) do
    if SameText(ARes.Levels[I].ClassName, AClass) then Exit(I);
end;

function LevelDeclares(const ARes: TIntfRefResolution;
  const AClass, AIntf: string): Boolean;
begin
  Result := False;
  var I := LevelNamed(ARes, AClass);
  if I < 0 then Exit;
  for var E in ARes.Levels[I].Interfaces do
    if SameText(E.Name, AIntf) then Exit(True);
end;

procedure TIntfVisualizerTests.AReferenceLeadsToItsObject;
var
  Obj: TVisLeaf;
  Ref: IThree;
  Res: TIntfRefResolution;
begin
  Obj := TVisLeaf.Create;
  Ref := Obj;                      // the reference we hand the resolver
  try
    Res := ResolveInterfaceReference(UInt64(NativeUInt(Pointer(Ref))),
      OwnProcessReader());

    Assert.IsTrue(Res.Resolved, 'the reference resolves: ' + Res.Reason);
    Assert.AreEqual(UInt64(NativeUInt(Pointer(Obj))), Res.ObjectAddr,
      'and it leads to the object itself - Obj = Reference - IOffset, ' +
      'which is the RTL''s own arithmetic in TObject.GetInterface');
    Assert.AreEqual(SizeOf(Pointer), Res.PointerSize,
      'the bitness was not assumed but proven by vmtSelfPtr, so it must be ' +
      'this process''s own');
    Assert.IsTrue(Res.IOffset > 0,
      'the offset is a real field offset in the object');
    Assert.AreEqual('TVisLeaf', Res.Levels[0].ClassName,
      'the most derived class comes first');
    Assert.AreEqual('IThree', Res.ReferenceInterface,
      'and the ONE entry the reference points at is named - it was ' +
      'identified by its VTable AND its offset, not by its position');
  finally
    Ref := nil;
  end;
end;

procedure TIntfVisualizerTests.EveryLevelNamesItsOwnInterfaces;
var
  Ref: ITwo;
  Res: TIntfRefResolution;
begin
  Ref := TVisLeaf.Create;          // asked through a MIDDLE interface
  try
    Res := ResolveInterfaceReference(UInt64(NativeUInt(Pointer(Ref))),
      OwnProcessReader());
    Assert.IsTrue(Res.Resolved, 'resolved: ' + Res.Reason);

    // the whole chain, in order, from the parent pointers
    Assert.IsTrue(Length(Res.Levels) >= 4,
      'TVisLeaf, TVisMiddle, TVisBase, TInterfacedObject and TObject - ' +
      'got ' + IntToStr(Length(Res.Levels)));
    Assert.IsTrue(LevelNamed(Res, 'TVisLeaf') = 0, 'TVisLeaf is level 0');
    Assert.IsTrue(LevelNamed(Res, 'TVisMiddle') = 1, 'TVisMiddle follows');
    Assert.IsTrue(LevelNamed(Res, 'TVisBase') = 2, 'then TVisBase');
    Assert.IsTrue(LevelNamed(Res, 'TInterfacedObject') > 2,
      'and the RTL class it descends from');
    Assert.IsTrue(LevelNamed(Res, 'TObject') > 0, 'TObject closes the chain');

    // PER LEVEL, which is the point: an interface table carries the entries
    // of its own class, so the attribution must not slide
    Assert.IsTrue(LevelDeclares(Res, 'TVisLeaf', 'IThree'),
      'TVisLeaf declares IThree');
    Assert.IsTrue(LevelDeclares(Res, 'TVisMiddle', 'ITwo'),
      'TVisMiddle declares ITwo');
    Assert.IsTrue(LevelDeclares(Res, 'TVisBase', 'IOne'),
      'TVisBase declares IOne');
    Assert.IsFalse(LevelDeclares(Res, 'TVisLeaf', 'IOne'),
      'and IOne is NOT attributed to TVisLeaf, although the object ' +
      'supports it - it is declared one level further up');
    Assert.IsTrue(LevelDeclares(Res, 'TInterfacedObject', 'IInterface'),
      'IInterface comes from TInterfacedObject');

    Assert.AreEqual('ITwo', Res.ReferenceInterface,
      'and asked through ITwo, that is the entry the reference points at');

    // the GUID is read from the same bytes and must be the declared one
    var G := StringToGUID('{7B1A0C64-6E2E-4B0A-9C31-3A5C9D1E0002}');
    var Found := False;
    for var E in Res.Levels[LevelNamed(Res, 'TVisMiddle')].Interfaces do
      if SameText(E.Name, 'ITwo') then
        Found := IsEqualGUID(E.IID, G);
    Assert.IsTrue(Found, 'ITwo''s entry carries its declared IID');
  finally
    Ref := nil;
  end;
end;

procedure TIntfVisualizerTests.WhatCannotBeResolvedSaysWhy;
var
  Res: TIntfRefResolution;
  Junk: UInt64;
begin
  Res := ResolveInterfaceReference(0, OwnProcessReader());
  Assert.IsFalse(Res.Resolved, 'nil is not an interface');
  Assert.IsTrue(Res.Reason.Contains('nil'), 'and it says so: ' + Res.Reason);

  Res := ResolveInterfaceReference(UInt64(NativeUInt(Pointer(Self))), nil);
  Assert.IsFalse(Res.Resolved, 'without a reader nothing is resolved');
  Assert.IsTrue(Res.Reason <> '', 'and that has a reason too');

  // a readable address that is no interface reference at all: a plain
  // string's data. Nothing may be invented for it.
  var S := 'not an interface';
  Junk := UInt64(NativeUInt(Pointer(S)));
  Res := ResolveInterfaceReference(Junk, OwnProcessReader());
  Assert.IsFalse(Res.Resolved,
    'a readable pointer that is no interface reference resolves to nothing');
  Assert.IsTrue(Res.Reason <> '', 'and says why: ' + Res.Reason);
  Assert.AreEqual(0, Integer(Length(Res.Levels)),
    'and it reports no hierarchy it could not prove');
end;

initialization
  TDUnitX.RegisterTestFixture(TIntfVisualizerTests);

end.
