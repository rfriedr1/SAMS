object frmDBSearch: TfrmDBSearch
  Left = 0
  Top = 0
  Caption = 'Database Search'
  ClientHeight = 790
  ClientWidth = 749
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    749
    790)
  PixelsPerInch = 120
  TextHeight = 17
  object lblSearchResults: TLabel
    Left = 9
    Top = 291
    Width = 89
    Height = 17
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Search Results'
  end
  object btnSearch: TButton
    Left = 120
    Top = 232
    Width = 483
    Height = 42
    Hint = 'perform search'
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Search'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
    OnClick = btnSearchClick
    ExplicitWidth = 265
  end
  object DBGridSearchResults: TDBGrid
    Left = 0
    Top = 320
    Width = 749
    Height = 452
    Hint = 'double click to see details'
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akTop, akRight, akBottom]
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -14
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDblClick = DBGridSearchResultsDblClick
  end
  object StaticText2: TStaticText
    Left = 578
    Top = 291
    Width = 162
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akTop, akRight]
    Caption = 'Double Click to see Details'
    TabOrder = 2
    ExplicitLeft = 360
  end
  object PageControl: TPageControl
    Left = 0
    Top = 0
    Width = 749
    Height = 225
    ActivePage = Mask
    Align = alTop
    TabOrder = 3
    ExplicitWidth = 531
    object Mask: TTabSheet
      Caption = 'Mask'
      ExplicitWidth = 523
      ExplicitHeight = 180
      DesignSize = (
        741
        193)
      object edtSearchPhrase: TLabeledEdit
        Left = 20
        Top = 140
        Width = 685
        Height = 25
        Hint = 'phrase for searching database table'
        Margins.Left = 4
        Margins.Top = 4
        Margins.Right = 4
        Margins.Bottom = 4
        Anchors = [akLeft, akTop, akRight]
        EditLabel.Width = 87
        EditLabel.Height = 17
        EditLabel.Margins.Left = 4
        EditLabel.Margins.Top = 4
        EditLabel.Margins.Right = 4
        EditLabel.Margins.Bottom = 4
        EditLabel.Caption = 'Search Phrase'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        TextHint = 'search phrase'
        ExplicitWidth = 467
      end
      object RadioGroupSearchContext: TRadioGroup
        Left = 20
        Top = 29
        Width = 685
        Height = 84
        Hint = 'select table to be searched'
        Margins.Left = 4
        Margins.Top = 4
        Margins.Right = 4
        Margins.Bottom = 4
        Anchors = [akLeft, akTop, akRight]
        Caption = 'Search Context'
        Columns = 2
        Items.Strings = (
          'Users'
          'Projects'
          'Samples'
          'Preparation'
          'Target')
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        ExplicitWidth = 467
      end
      object StaticText1: TStaticText
        Left = 0
        Top = 0
        Width = 741
        Height = 21
        Margins.Left = 4
        Margins.Top = 4
        Margins.Right = 4
        Margins.Bottom = 4
        Align = alTop
        Caption = 
          'Database will be searched for any appearance of the search phras' +
          'e.'
        TabOrder = 2
        ExplicitLeft = 20
        ExplicitTop = 4
        ExplicitWidth = 523
      end
    end
    object Query: TTabSheet
      Caption = 'Query'
      ImageIndex = 1
      ExplicitWidth = 523
      ExplicitHeight = 180
      DesignSize = (
        741
        193)
      object MemoQuery: TMemo
        Left = 16
        Top = 40
        Width = 707
        Height = 121
        Anchors = [akLeft, akTop, akRight]
        Ctl3D = False
        Lines.Strings = (
          '')
        ParentCtl3D = False
        TabOrder = 0
        ExplicitWidth = 489
      end
      object StaticText3: TStaticText
        Left = 20
        Top = 13
        Width = 103
        Height = 21
        Caption = 'Database Query'
        TabOrder = 1
      end
    end
  end
end
