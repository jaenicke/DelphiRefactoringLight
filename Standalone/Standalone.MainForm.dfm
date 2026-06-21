object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'Refactoring Light - Standalone'
  ClientHeight = 800
  ClientWidth = 1200
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Menu = MainMenu1
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 15
  object Tree: TTreeView
    Left = 0
    Top = 0
    Width = 320
    Height = 781
    Align = alLeft
    Indent = 19
    ReadOnly = True
    ShowRoot = False
    TabOrder = 0
    OnClick = DoTreeClick
  end
  object Splitter1: TSplitter
    Left = 320
    Top = 0
    Width = 4
    Height = 781
    ExplicitLeft = 321
    ExplicitHeight = 781
  end
  object Pages: TPageControl
    Left = 324
    Top = 0
    Width = 876
    Height = 781
    ActivePage = TabEditor
    Align = alClient
    TabOrder = 1
    OnChange = DoPagesChange
    object TabEditor: TTabSheet
      Caption = '&Editor'
      object Memo: TMemo
        Left = 0
        Top = 0
        Width = 868
        Height = 753
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Consolas'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssBoth
        TabOrder = 0
        WordWrap = False
        OnChange = DoMemoChange
        OnClick = DoMemoClick
        OnKeyDown = DoMemoKeyDown
        OnKeyPress = DoMemoKeyPress
        OnKeyUp = DoMemoKeyUp
      end
    end
    object TabLspLog: TTabSheet
      Caption = '&LSP Diagnostics'
      ImageIndex = 1
      object LspLogToolbar: TPanel
        Left = 0
        Top = 0
        Width = 868
        Height = 28
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        object BtnLspLogClear: TButton
          Left = 4
          Top = 2
          Width = 75
          Height = 23
          Caption = '&Clear'
          TabOrder = 0
          OnClick = DoLspLogClear
        end
        object ChkLspLogAutoscroll: TCheckBox
          Left = 88
          Top = 4
          Width = 100
          Height = 19
          Caption = 'Auto&scroll'
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
        object LblLspLogHint: TLabel
          Left = 200
          Top = 6
          Width = 660
          Height = 15
          Caption = 'Live LSP traffic + status notes. Captured starting at project open.'
        end
      end
      object LspLog: TMemo
        Left = 0
        Top = 28
        Width = 868
        Height = 725
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Consolas'
        Font.Style = []
        HideSelection = False
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssBoth
        TabOrder = 1
        WordWrap = False
      end
    end
  end
  object ProgressPanel: TPanel
    Left = 0
    Top = 760
    Width = 1200
    Height = 21
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 4
    Visible = False
    object ProgressLabel: TLabel
      Left = 6
      Top = 3
      Width = 320
      Height = 15
      Caption = 'LSP indexing...'
    end
    object ProgressBar: TProgressBar
      Left = 336
      Top = 0
      Width = 864
      Height = 21
      Align = alRight
      TabOrder = 0
    end
  end
  object Status: TStatusBar
    Left = 0
    Top = 781
    Width = 1200
    Height = 19
    Panels = <
      item
        Width = 350
      end
      item
        Width = 320
      end
      item
        Width = 160
      end
      item
        Width = 370
      end>
  end
  object LspPoll: TTimer
    Enabled = False
    Interval = 500
    OnTimer = DoLspPollTick
    Left = 600
    Top = 100
  end
  object LspLogFlush: TTimer
    Enabled = False
    Interval = 250
    OnTimer = DoLspLogFlushTick
    Left = 600
    Top = 140
  end
  object MainMenu1: TMainMenu
    Left = 416
    Top = 96
    object MenuFile: TMenuItem
      Caption = '&File'
      object MenuFileOpen: TMenuItem
        Caption = '&Open project...'
        ShortCut = 16463
        OnClick = DoFileOpenProject
      end
      object MenuFileSave: TMenuItem
        Caption = '&Save active file'
        ShortCut = 16467
        OnClick = DoFileSave
      end
      object MenuFileSep: TMenuItem
        Caption = '-'
      end
      object MenuFileExit: TMenuItem
        Caption = 'E&xit'
        OnClick = DoFileExit
      end
    end
    object MenuRefactor: TMenuItem
      Caption = '&Refactor'
      object MenuRename: TMenuItem
        Caption = '&Rename identifier...'
        ShortCut = 24658
        OnClick = DoRefactorRename
      end
      object MenuFindRef: TMenuItem
        Caption = 'Find &references'
        ShortCut = 24646
        OnClick = DoRefactorFindReferences
      end
      object MenuFindImp: TMenuItem
        Caption = 'Find &implementations'
        ShortCut = 24649
        OnClick = DoRefactorFindImplementations
      end
      object MenuAlignSig: TMenuItem
        Caption = '&Align method signature...'
        OnClick = DoRefactorAlignSignature
      end
      object MenuExtractMethod: TMenuItem
        Caption = 'E&xtract method...'
        ShortCut = 57421
        OnClick = DoRefactorExtractMethod
      end
      object MenuCompletion: TMenuItem
        Caption = 'Code &completion'
        ShortCut = 16416
        OnClick = DoRefactorCompletion
      end
      object MenuSep1: TMenuItem
        Caption = '-'
      end
      object MenuRemoveWith: TMenuItem
        Caption = 'Remove &with'
        object MenuRwCursor: TMenuItem
          Caption = 'At &cursor only'
          OnClick = DoRefactorRemoveWithCursor
        end
        object MenuRwCurrent: TMenuItem
          Caption = 'In c&urrent unit'
          OnClick = DoRefactorRemoveWithCurrent
        end
        object MenuRwSelected: TMenuItem
          Caption = 'In &selected units...'
          OnClick = DoRefactorRemoveWithSelected
        end
        object MenuRwProject: TMenuItem
          Caption = 'In whole &project...'
          OnClick = DoRefactorRemoveWithProject
        end
      end
      object MenuMoveToUnit: TMenuItem
        Caption = '&Move to unit...'
        OnClick = DoRefactorMoveToUnit
      end
      object MenuUnitRefs: TMenuItem
        Caption = 'Find &unit references...'
        OnClick = DoRefactorUnitRefs
      end
      object MenuSep2: TMenuItem
        Caption = '-'
      end
      object MenuIface: TMenuItem
        Caption = '&Extract / extend interface'
        object MenuIfaceExtract: TMenuItem
          Caption = '&Extract new interface from class...'
          OnClick = DoRefactorExtractInterface
        end
        object MenuIfaceAdd: TMenuItem
          Caption = '&Add to existing interface...'
          OnClick = DoRefactorAddToInterface
        end
        object MenuIfaceImpl: TMenuItem
          Caption = 'Add &IInterface support to class...'
          OnClick = DoRefactorAddIInterface
        end
      end
      object MenuSep3: TMenuItem
        Caption = '-'
      end
      object MenuSemRep: TMenuItem
        Caption = '&Semantic replace'
        object MenuSemRepCurrent: TMenuItem
          Caption = 'In current unit'
          OnClick = DoRefactorSemanticCurrent
        end
        object MenuSemRepSelected: TMenuItem
          Caption = 'In selected units...'
          OnClick = DoRefactorSemanticSelected
        end
        object MenuSemRepProject: TMenuItem
          Caption = 'In whole project...'
          OnClick = DoRefactorSemanticProject
        end
        object MenuSemRepEditRules: TMenuItem
          Caption = 'Edit rules...'
          OnClick = DoRefactorSemanticEditRules
        end
      end
    end
  end
end
