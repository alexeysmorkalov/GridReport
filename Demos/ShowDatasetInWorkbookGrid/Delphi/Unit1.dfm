object Form1: TForm1
  Left = 192
  Top = 114
  Width = 696
  Height = 480
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 129
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    OnResize = Panel1Resize
    object Label1: TLabel
      Left = 8
      Top = 12
      Width = 59
      Height = 13
      Caption = 'Database:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 29
      Top = 32
      Width = 38
      Height = 13
      Caption = 'Query:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object EDDatabase: TEdit
      Left = 72
      Top = 8
      Width = 145
      Height = 21
      TabOrder = 0
      Text = 'DBDEMOS'
    end
    object MQuery: TMemo
      Left = 72
      Top = 32
      Width = 529
      Height = 89
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Courier New'
      Font.Style = []
      Lines.Strings = (
        'select * from customer')
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 1
    end
    object bExecute: TButton
      Left = 608
      Top = 32
      Width = 75
      Height = 25
      Caption = 'Execute'
      TabOrder = 2
      OnClick = bExecuteClick
    end
  end
  object vgrWorkbookGrid1: TvgrWorkbookGrid
    Left = 0
    Top = 129
    Width = 688
    Height = 317
    Align = alClient
    UseDockManager = False
    TabOrder = 1
    Workbook = Workbook
    OptionsCols.GridColor = clBtnShadow
    OptionsRows.GridColor = clBtnShadow
  end
  object Database: TDatabase
    AliasName = 'DBDEMOS'
    DatabaseName = 'DB'
    SessionName = 'Default'
    Left = 80
    Top = 72
  end
  object Query: TQuery
    DatabaseName = 'DB'
    Left = 112
    Top = 72
  end
  object Workbook: TvgrWorkbook
    Left = 152
    Top = 72
    DataStorageVersion = 1
    SystemInfo = (
      'OS: WIN32_NT 5.1.2600 '
      ''
      'PageSize: 4096'
      'ActiveProcessorMask: $1000'
      'NumberOfProcessors: 1'
      'ProcessorType: 586'
      ''
      'Compiler version: Delphi5'
      'DataStorage version: 1')
    RangeStylesData = {
      060000003E000000010000000464010000000A00000000000105417269616C00
      24750041247500419437000000000000A4978D01FCF88B011C08000080FFFFFF
      1F000000000000000000000000000100}
    BorderStylesData = {060000000900000001000000000001000000000000000000000000}
    StringsData = {010000000100000000000000}
  end
end
