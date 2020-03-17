object ShowScriptForm: TShowScriptForm
  Left = 262
  Top = 238
  Width = 594
  Height = 390
  Caption = 'ShowScriptForm'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object vgrScriptEdit1: TvgrScriptEdit
    Left = 0
    Top = 57
    Width = 586
    Height = 306
    Align = alClient
    BriefCursorShapes = False
    GutterVisible = True
    ReadOnly = False
    ReplaceTabsWithSpaces = True
    Script.Language = 'VBScript'
    TabOrder = 0
    TextOptions.FontName = 'Courier New'
    TextOptions.FontSize = 10
    TextOptions.TextColors.AttrWhitespace.FontStyle = []
    TextOptions.TextColors.AttrWhitespace.Color = clWindowText
    TextOptions.TextColors.AttrWhitespace.DefaultColor = True
    TextOptions.TextColors.AttrKeyword.FontStyle = [fsBold]
    TextOptions.TextColors.AttrKeyword.Color = clBlack
    TextOptions.TextColors.AttrKeyword.DefaultColor = False
    TextOptions.TextColors.AttrComment.FontStyle = [fsItalic]
    TextOptions.TextColors.AttrComment.Color = clBlue
    TextOptions.TextColors.AttrComment.DefaultColor = False
    TextOptions.TextColors.AttrNonsource.FontStyle = []
    TextOptions.TextColors.AttrNonsource.Color = clWindowText
    TextOptions.TextColors.AttrNonsource.DefaultColor = True
    TextOptions.TextColors.AttrOperator.FontStyle = []
    TextOptions.TextColors.AttrOperator.Color = clGreen
    TextOptions.TextColors.AttrOperator.DefaultColor = False
    TextOptions.TextColors.AttrNumber.FontStyle = []
    TextOptions.TextColors.AttrNumber.Color = clMaroon
    TextOptions.TextColors.AttrNumber.DefaultColor = False
    TextOptions.TextColors.AttrString.FontStyle = []
    TextOptions.TextColors.AttrString.Color = clNavy
    TextOptions.TextColors.AttrString.DefaultColor = False
    TextOptions.TextColors.AttrFunction.FontStyle = []
    TextOptions.TextColors.AttrFunction.Color = clMaroon
    TextOptions.TextColors.AttrFunction.DefaultColor = False
    TextOptions.TextColors.AttrIdentifier.FontStyle = []
    TextOptions.TextColors.AttrIdentifier.Color = clBlack
    TextOptions.TextColors.AttrIdentifier.DefaultColor = False
    TextOptions.BkColor = clWindow
    TextOptions.ErrorBackColor = clMaroon
    TextOptions.ErrorForeColor = clWhite
    WantTabs = False
    WantReturns = False
    EditorMode = vmmInsert
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 586
    Height = 57
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    DesignSize = (
      586
      57)
    object RadioButton1: TRadioButton
      Left = 16
      Top = 24
      Width = 65
      Height = 17
      Caption = 'VBScript'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = RadioButton1Click
    end
    object RadioButton2: TRadioButton
      Left = 88
      Top = 24
      Width = 57
      Height = 17
      Caption = 'JScript'
      TabOrder = 1
      OnClick = RadioButton2Click
    end
    object Button2: TButton
      Left = 464
      Top = 16
      Width = 115
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Apply changes'
      TabOrder = 2
      OnClick = Button2Click
    end
  end
end
