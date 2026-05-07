object LspOptionsFrame: TLspOptionsFrame
  Left = 0
  Top = 0
  Width = 520
  Height = 320
  TabOrder = 0
  object grpShortcuts: TGroupBox
    Left = 8
    Top = 8
    Width = 504
    Height = 256
    Caption = ' Keyboard shortcuts '
    TabOrder = 0
    object lblRename: TLabel
      Left = 16
      Top = 28
      Width = 56
      Height = 13
      Caption = 'Rename:'
    end
    object lblCompletion: TLabel
      Left = 16
      Top = 60
      Width = 80
      Height = 13
      Caption = 'Code completion:'
    end
    object lblExtract: TLabel
      Left = 16
      Top = 92
      Width = 80
      Height = 13
      Caption = 'Extract method:'
    end
    object lblFindRef: TLabel
      Left = 16
      Top = 124
      Width = 80
      Height = 13
      Caption = 'Find references:'
    end
    object lblFindImp: TLabel
      Left = 16
      Top = 156
      Width = 100
      Height = 13
      Caption = 'Find implementations:'
    end
    object lblAlign: TLabel
      Left = 16
      Top = 188
      Width = 110
      Height = 13
      Caption = 'Align method signature:'
    end
    object lblHint: TLabel
      Left = 16
      Top = 224
      Width = 480
      Height = 13
      Caption =
        'Click into a field and press the desired key combination. Backsp' +
        'ace or Delete clears the shortcut.'
    end
    object edtRename: TEdit
      Left = 160
      Top = 24
      Width = 200
      Height = 21
      ReadOnly = False
      TabOrder = 0
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
    object edtCompletion: TEdit
      Left = 160
      Top = 56
      Width = 200
      Height = 21
      TabOrder = 1
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
    object edtExtract: TEdit
      Left = 160
      Top = 88
      Width = 200
      Height = 21
      TabOrder = 2
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
    object edtFindRef: TEdit
      Left = 160
      Top = 120
      Width = 200
      Height = 21
      TabOrder = 3
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
    object edtFindImp: TEdit
      Left = 160
      Top = 152
      Width = 200
      Height = 21
      TabOrder = 4
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
    object edtAlign: TEdit
      Left = 160
      Top = 184
      Width = 200
      Height = 21
      TabOrder = 5
      OnKeyDown = ShortcutEditKeyDown
      OnKeyPress = ShortcutEditKeyPress
    end
  end
  object btnDefaults: TButton
    Left = 8
    Top = 280
    Width = 145
    Height = 25
    Caption = 'Restore defaults'
    TabOrder = 1
    OnClick = btnDefaultsClick
  end
end
