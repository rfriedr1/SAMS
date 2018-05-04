object frmDBSearch: TfrmDBSearch
  Left = 0
  Top = 0
  Caption = 'Database Search'
  ClientHeight = 841
  ClientWidth = 531
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    531
    841)
  PixelsPerInch = 120
  TextHeight = 17
  object lblSearchResults: TLabel
    Left = 10
    Top = 316
    Width = 89
    Height = 17
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Search Results'
  end
  object StaticText1: TStaticText
    Left = 30
    Top = 30
    Width = 403
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Database will be searched for any appearance of the search term.'
    TabOrder = 0
  end
  object btnSearch: TButton
    Left = 390
    Top = 100
    Width = 131
    Height = 126
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akTop, akRight]
    Caption = 'Search'
    TabOrder = 1
    OnClick = btnSearchClick
  end
  object DBGridSearchResults: TDBGrid
    Left = 0
    Top = 340
    Width = 531
    Height = 491
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akTop, akRight, akBottom]
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -14
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDblClick = DBGridSearchResultsDblClick
  end
  object StaticText2: TStaticText
    Left = 360
    Top = 311
    Width = 162
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akTop, akRight]
    Caption = 'Double Click to see Details'
    TabOrder = 3
  end
  object RadioGroupSearchContext: TRadioGroup
    Left = 30
    Top = 90
    Width = 341
    Height = 71
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Search Context'
    Columns = 2
    Items.Strings = (
      'Users'
      'Projects'
      'Samples')
    TabOrder = 4
  end
  object edtSearchPhrase: TLabeledEdit
    Left = 30
    Top = 200
    Width = 341
    Height = 25
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    EditLabel.Width = 87
    EditLabel.Height = 17
    EditLabel.Margins.Left = 4
    EditLabel.Margins.Top = 4
    EditLabel.Margins.Right = 4
    EditLabel.Margins.Bottom = 4
    EditLabel.Caption = 'Search Phrase'
    TabOrder = 5
  end
end
