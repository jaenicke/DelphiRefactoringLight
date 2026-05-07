unit Expert.Shortcuts;

interface

uses
  Winapi.Windows,
  System.UITypes,
  System.Classes;

type
  TExpertsShortCut = record
  const
    scRename      = TShortCut(vkR     or scAlt or scCtrl or scShift);
    scCompletion  = TShortCut(vkSpace or scAlt or scCtrl or scShift);
    scExtract     = TShortCut(vkM     or scAlt or scCtrl or scShift);
    scFindRef     = TShortCut(vkU     or scAlt or scCtrl or scShift);
    scFindImp     = TShortCut(vkI     or scAlt or scCtrl or scShift);
    scAlign       = TShortCut(vkA     or scAlt or scCtrl or scShift);
  end;

implementation

end.
