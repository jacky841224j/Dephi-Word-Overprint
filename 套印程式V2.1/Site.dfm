object Roster: TRoster
  Left = 0
  Top = 0
  Caption = 'Roster'
  ClientHeight = 634
  ClientWidth = 1109
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Gauge2: TGauge
    Left = -3
    Top = 0
    Width = 1112
    Height = 21
    Progress = 0
  end
  object dbgrd1: TDBGrid
    Left = 0
    Top = 21
    Width = 1109
    Height = 95
    DataSource = ds1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object strngrdCheckList: TStringGrid
    Left = 0
    Top = 116
    Width = 320
    Height = 518
    ColCount = 4
    RowCount = 2
    TabOrder = 1
  end
  object ds1: TDataSource
    DataSet = Form1.qry1
    Left = 1072
    Top = 37
  end
end
