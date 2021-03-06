object frmDBSearch: TfrmDBSearch
  Left = 0
  Top = 0
  Caption = 'Database Search'
  ClientHeight = 713
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
    713)
  PixelsPerInch = 120
  TextHeight = 17
  object lblSearchResults: TLabel
    Left = 9
    Top = 275
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
    Top = 215
    Width = 265
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
  end
  object DBGridSearchResults: TDBGrid
    Left = 0
    Top = 304
    Width = 531
    Height = 391
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
    Left = 360
    Top = 275
    Width = 162
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akTop, akRight]
    Caption = 'Double Click to see Details'
    TabOrder = 2
  end
  object PageControl: TPageControl
    Left = 0
    Top = 0
    Width = 531
    Height = 212
    ActivePage = Mask
    Align = alTop
    TabOrder = 3
    object Mask: TTabSheet
      Caption = 'Mask'
      ExplicitHeight = 161
      DesignSize = (
        523
        180)
      object edtSearchPhrase: TLabeledEdit
        Left = 20
        Top = 132
        Width = 467
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
      end
      object RadioGroupSearchContext: TRadioGroup
        Left = 20
        Top = 29
        Width = 467
        Height = 71
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
          'Samples')
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
      end
      object StaticText1: TStaticText
        Left = 0
        Top = 0
        Width = 523
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
        ExplicitLeft = -134
        ExplicitTop = 22
        ExplicitWidth = 415
      end
    end
    object Query: TTabSheet
      Caption = 'Query'
      ImageIndex = 1
      ExplicitWidth = 281
      ExplicitHeight = 161
      DesignSize = (
        523
        180)
      object MemoQuery: TMemo
        Left = 16
        Top = 40
        Width = 489
        Height = 121
        Anchors = [akLeft, akTop, akRight]
        Ctl3D = False
        Lines.Strings = (
          '')
        ParentCtl3D = False
        TabOrder = 0
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
