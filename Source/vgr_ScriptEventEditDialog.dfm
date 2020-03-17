object vgrScriptEventEditForm: TvgrScriptEventEditForm
  Left = 89
  Top = 268
  ActiveControl = edProcName
  BorderStyle = bsDialog
  Caption = 'Define event procedure'
  ClientHeight = 95
  ClientWidth = 478
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 36
    Width = 111
    Height = 13
    Caption = 'Event &procedure name:'
  end
  object Label2: TLabel
    Left = 8
    Top = 8
    Width = 66
    Height = 13
    Caption = 'Event source:'
  end
  object Label3: TLabel
    Left = 88
    Top = 8
    Width = 39
    Height = 13
    Caption = 'Label3'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object edProcName: TComboBox
    Left = 128
    Top = 32
    Width = 345
    Height = 21
    ItemHeight = 13
    TabOrder = 0
  end
  object bOK: TButton
    Left = 320
    Top = 64
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
  object bCancel: TButton
    Left = 400
    Top = 64
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 2
  end
  object vgrFormLocalizer1: TvgrFormLocalizer
    Items = <
      item
        Component = Owner
        PropName = 'Caption'
        ResStringID = 216
      end
      item
        Component = Label1
        PropName = 'Caption'
        ResStringID = 217
      end
      item
        Component = Label2
        PropName = 'Caption'
        ResStringID = 218
      end
      item
        Component = bOK
        PropName = 'Caption'
        ResStringID = 219
      end
      item
        Component = bCancel
        PropName = 'Caption'
        ResStringID = 220
      end>
    Left = 8
    Top = 64
  end
end
