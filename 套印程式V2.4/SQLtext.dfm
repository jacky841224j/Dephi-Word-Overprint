object SQLSetting: TSQLSetting
  Left = 0
  Top = 0
  Caption = 'SQLSetting.ini '#35373#23450
  ClientHeight = 564
  ClientWidth = 1145
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object lblSelect: TLabel
    Left = 8
    Top = 13
    Width = 56
    Height = 25
    Caption = 'Select'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -21
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object edtPath: TEdit
    Left = 432
    Top = 460
    Width = 247
    Height = 24
    BevelWidth = 2
    Color = clHighlightText
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWhite
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    TextHint = #35531#36984#25799#20786#23384#36335#24465
  end
  object btnDirpath2: TButton
    Left = 685
    Top = 460
    Width = 46
    Height = 24
    Caption = '...'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #24494#36575#27491#40657#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = btnDirpath2Click
  end
  object edtSQL: TEdit
    Left = 70
    Top = 11
    Width = 803
    Height = 27
    BevelWidth = 2
    Color = clHighlightText
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    TextHint = #35531#21247#36028#19978'"SELECT" '#23383#30524
    OnChange = edtSQLChange
  end
  object chk1: TCheckBox
    Left = 8
    Top = 44
    Width = 177
    Height = 25
    Caption = #26159#21542#21547'Where '#35486#27861
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
  end
  object pnl1: TPanel
    Left = 0
    Top = 143
    Width = 873
    Height = 271
    TabOrder = 4
    object lbl1: TLabel
      Left = 0
      Top = 0
      Width = 83
      Height = 31
      Caption = 'SQL'#38928#35261
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -21
      Font.Name = 'Noto Sans CJK TC Bold'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lbledtPreView: TLabeledEdit
      Left = 0
      Top = 72
      Width = 873
      Height = 27
      EditLabel.Width = 114
      EditLabel.Height = 28
      EditLabel.Caption = #22871#21360#38928#35261#35486#27861
      EditLabel.Font.Charset = ANSI_CHARSET
      EditLabel.Font.Color = clWindowText
      EditLabel.Font.Height = -19
      EditLabel.Font.Name = 'Noto Sans CJK TC Regular'
      EditLabel.Font.Style = []
      EditLabel.ParentFont = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
    end
    object lbledtTOP3: TLabeledEdit
      Left = 0
      Top = 145
      Width = 873
      Height = 27
      EditLabel.Width = 133
      EditLabel.Height = 28
      EditLabel.Caption = #36039#26009#24235#38928#35261#35486#27861
      EditLabel.Font.Charset = ANSI_CHARSET
      EditLabel.Font.Color = clWindowText
      EditLabel.Font.Height = -19
      EditLabel.Font.Name = 'Noto Sans CJK TC Regular'
      EditLabel.Font.Style = []
      EditLabel.ParentFont = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
    end
    object lbledtExport: TLabeledEdit
      Left = 0
      Top = 217
      Width = 873
      Height = 27
      EditLabel.Width = 76
      EditLabel.Height = 28
      EditLabel.Caption = #22871#21360#35486#27861
      EditLabel.Font.Charset = ANSI_CHARSET
      EditLabel.Font.Color = clWindowText
      EditLabel.Font.Height = -19
      EditLabel.Font.Name = 'Noto Sans CJK TC Regular'
      EditLabel.Font.Style = []
      EditLabel.ParentFont = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 2
    end
  end
  object edtExce: TEdit
    Left = 8
    Top = 92
    Width = 865
    Height = 27
    BevelWidth = 2
    Color = clHighlightText
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
    TextHint = #38928#23384#25110#20989#25976#35486#27861#35531#36028#27492
    OnChange = edtExceChange
  end
  object btnSave: TBitBtn
    Left = 768
    Top = 446
    Width = 105
    Height = 46
    Caption = #20786#23384
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Noto Sans CJK TC Regular'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = btnSaveClick
  end
end
