object frmDBSearch: TfrmDBSearch
  Left = 0
  Top = 0
  Caption = 'Database Search'
  ClientHeight = 673
  ClientWidth = 425
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    425
    673)
  PixelsPerInch = 96
  TextHeight = 13
  object lblSearchResults: TLabel
    Left = 8
    Top = 253
    Width = 71
    Height = 13
    Caption = 'Search Results'
  end
  object StaticText1: TStaticText
    Left = 24
    Top = 24
    Width = 323
    Height = 17
    Caption = 'Database will be searched for any appearance of the search term.'
    TabOrder = 0
  end
  object btnSearch: TButton
    Left = 312
    Top = 80
    Width = 105
    Height = 101
    Anchors = [akTop, akRight]
    Caption = 'Search'
    TabOrder = 1
    OnClick = btnSearchClick
  end
  object DBGridSearchResults: TDBGrid
    Left = 0
    Top = 272
    Width = 425
    Height = 393
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object StaticText2: TStaticText
    Left = 288
    Top = 249
    Width = 129
    Height = 17
    Anchors = [akTop, akRight]
    Caption = 'Double Click to see Details'
    TabOrder = 3
  end
  object RadioGroup1: TRadioGroup
    Left = 24
    Top = 72
    Width = 273
    Height = 57
    Caption = 'search context'
    Columns = 2
    Items.Strings = (
      'Users'
      'Projects'
      'Samples')
    TabOrder = 4
  end
  object edtSearchPhrase: TLabeledEdit
    Left = 24
    Top = 160
    Width = 273
    Height = 21
    EditLabel.Width = 69
    EditLabel.Height = 13
    EditLabel.Caption = 'Search Phrase'
    TabOrder = 5
  end
end
