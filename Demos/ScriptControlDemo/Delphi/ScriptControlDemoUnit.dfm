object Form1: TForm1
  Left = 373
  Top = 136
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'vgrScriptControl Demo'
  ClientHeight = 297
  ClientWidth = 290
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton6: TSpeedButton
    Left = 8
    Top = 264
    Width = 273
    Height = 25
    Caption = 'Show Script'
    OnClick = SpeedButton6Click
  end
  object Panel1: TPanel
    Left = 8
    Top = 8
    Width = 273
    Height = 249
    BevelInner = bvLowered
    TabOrder = 0
    object SpeedButton1: TSpeedButton
      Left = 208
      Top = 168
      Width = 57
      Height = 35
      Caption = '+'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton1Click
    end
    object SpeedButton2: TSpeedButton
      Left = 208
      Top = 128
      Width = 57
      Height = 35
      Caption = '-'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton2Click
    end
    object SpeedButton3: TSpeedButton
      Left = 208
      Top = 48
      Width = 57
      Height = 35
      Caption = '/'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton3Click
    end
    object SpeedButton4: TSpeedButton
      Left = 208
      Top = 88
      Width = 57
      Height = 35
      Caption = '*'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton4Click
    end
    object SpeedButton5: TSpeedButton
      Left = 11
      Top = 206
      Width = 94
      Height = 35
      Caption = 'C'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton5Click
    end
    object SpeedButton7: TSpeedButton
      Left = 74
      Top = 168
      Width = 57
      Height = 35
      Caption = '+/-'
      OnClick = SpeedButton7Click
    end
    object SpeedButton8: TSpeedButton
      Left = 10
      Top = 168
      Width = 57
      Height = 35
      Caption = '0'
      OnClick = SpeedButton8Click
    end
    object SpeedButton9: TSpeedButton
      Left = 10
      Top = 128
      Width = 57
      Height = 35
      Caption = '1'
      OnClick = SpeedButton9Click
    end
    object SpeedButton11: TSpeedButton
      Left = 74
      Top = 128
      Width = 57
      Height = 35
      Caption = '2'
      OnClick = SpeedButton11Click
    end
    object SpeedButton12: TSpeedButton
      Left = 138
      Top = 128
      Width = 57
      Height = 35
      Caption = '3'
      OnClick = SpeedButton12Click
    end
    object SpeedButton13: TSpeedButton
      Left = 74
      Top = 88
      Width = 57
      Height = 35
      Caption = '5'
      OnClick = SpeedButton13Click
    end
    object SpeedButton14: TSpeedButton
      Left = 138
      Top = 88
      Width = 57
      Height = 35
      Caption = '6'
      OnClick = SpeedButton14Click
    end
    object SpeedButton15: TSpeedButton
      Left = 10
      Top = 48
      Width = 57
      Height = 35
      Caption = '7'
      OnClick = SpeedButton15Click
    end
    object SpeedButton16: TSpeedButton
      Left = 74
      Top = 48
      Width = 57
      Height = 35
      Caption = '8'
      OnClick = SpeedButton16Click
    end
    object SpeedButton17: TSpeedButton
      Left = 138
      Top = 48
      Width = 57
      Height = 35
      Caption = '9'
      OnClick = SpeedButton17Click
    end
    object SpeedButton18: TSpeedButton
      Left = 10
      Top = 88
      Width = 57
      Height = 35
      Caption = '4'
      OnClick = SpeedButton18Click
    end
    object SpeedButton19: TSpeedButton
      Left = 138
      Top = 168
      Width = 57
      Height = 35
      Caption = ','
      OnClick = SpeedButton19Click
    end
    object SpeedButton20: TSpeedButton
      Left = 112
      Top = 206
      Width = 155
      Height = 35
      Caption = '='
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      OnClick = SpeedButton20Click
    end
    object ResultLabel: TLabel
      Left = 8
      Top = 8
      Width = 257
      Height = 25
      Alignment = taRightJustify
      AutoSize = False
      Caption = '0'
      Color = clBtnHighlight
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
  end
  object vgrScriptControl1: TvgrScriptControl
    Language = 'VBScript'
    Left = 16
    Top = 224
  end
end
