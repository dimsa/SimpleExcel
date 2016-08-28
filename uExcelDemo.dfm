object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'MainForm'
  ClientHeight = 577
  ClientWidth = 1126
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnDestroy = FormDestroy
  DesignSize = (
    1126
    577)
  PixelsPerInch = 96
  TextHeight = 13
  object MainLabel: TLabel
    Left = 8
    Top = 8
    Width = 89
    Height = 19
    Caption = 'Simple Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object OpenFileBtn: TButton
    Left = 9
    Top = 104
    Width = 193
    Height = 25
    Caption = 'Open File'
    TabOrder = 2
    OnClick = OpenFileBtnClick
  end
  object PageControl: TPageControl
    Left = 208
    Top = 0
    Width = 918
    Height = 577
    Align = alRight
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 4
  end
  object SaveFileBtn: TButton
    Left = 9
    Top = 135
    Width = 193
    Height = 25
    Caption = 'Save File'
    TabOrder = 3
    OnClick = SaveFileBtnClick
  end
  object CreateNewBtn: TButton
    Left = 8
    Top = 42
    Width = 193
    Height = 25
    Caption = 'Create File'
    TabOrder = 0
    OnClick = CreateNewBtnClick
  end
  object ExcelPanel: TPanel
    Left = 8
    Top = 464
    Width = 185
    Height = 105
    Anchors = [akLeft, akBottom]
    TabOrder = 5
    object AddSheetBtn: TButton
      Left = 8
      Top = 8
      Width = 169
      Height = 25
      Caption = 'AddSheetBtn'
      TabOrder = 0
      OnClick = AddSheetBtnClick
    end
    object AddRowBtn: TButton
      Left = 8
      Top = 39
      Width = 169
      Height = 25
      Caption = 'AddRowBtn'
      TabOrder = 1
      OnClick = AddRowBtnClick
    end
    object AddColBtn: TButton
      Left = 8
      Top = 70
      Width = 169
      Height = 25
      Caption = 'AddColBtn'
      TabOrder = 2
      OnClick = AddColBtnClick
    end
  end
  object Button2: TButton
    Left = 8
    Top = 73
    Width = 194
    Height = 25
    Caption = 'OpenDefaultFileBtn'
    TabOrder = 1
    OnClick = Button2Click
  end
  object SaveDlg: TSaveDialog
    Left = 56
    Top = 176
  end
  object OpenDlg: TOpenDialog
    Left = 16
    Top = 176
  end
end
