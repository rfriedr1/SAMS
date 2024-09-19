object FormStorageLocations: TFormStorageLocations
  Left = 0
  Top = 0
  ActiveControl = btnClose
  Caption = 'FormStorageLocations'
  ClientHeight = 660
  ClientWidth = 721
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    721
    660)
  PixelsPerInch = 96
  TextHeight = 13
  object WarningLabel: TLabel
    Left = 600
    Top = 162
    Width = 88
    Height = 16
    Alignment = taRightJustify
    Caption = 'WarningLabel'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    Visible = False
    StyleElements = [seClient, seBorder]
  end
  object lblNumberOfShownSamples: TLabel
    Left = 16
    Top = 152
    Width = 113
    Height = 13
    Caption = '... (displayed in table): '
  end
  object lblNumberofNoLeftoverSamples: TLabel
    Left = 16
    Top = 171
    Width = 132
    Height = 13
    Caption = '... (not displayed in table): '
  end
  object btnClose: TButton
    Left = 391
    Top = 582
    Width = 105
    Height = 38
    Hint = 'close window without saving changes'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    Cancel = True
    Caption = 'Close (ESC)'
    ModalResult = 8
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
  end
  object btnSave: TButton
    Left = 72
    Top = 582
    Width = 291
    Height = 38
    Hint = 'save storage locations to database'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    Caption = 'Save'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    OnClick = btnSaveClick
  end
  object DBGrid1: TDBGrid
    Left = 1
    Top = 208
    Width = 720
    Height = 311
    Hint = 'list of found samples'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGrid1DrawColumnCell
  end
  object edtStatus: TEdit
    Left = 0
    Top = 639
    Width = 721
    Height = 21
    Align = alBottom
    BorderStyle = bsNone
    TabOrder = 3
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 721
    Height = 73
    Align = alTop
    Caption = 'Sample ID'#39's MAMS'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    object edtStartSampleID: TLabeledEdit
      Left = 48
      Top = 32
      Width = 121
      Height = 21
      Hint = 'start of the range of sample numbers'
      CustomHint = BalloonHint1
      EditLabel.Width = 24
      EditLabel.Height = 13
      EditLabel.Caption = 'Start'
      LabelPosition = lpLeft
      NumbersOnly = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnChange = edtStartSampleIDChange
    end
    object edtEndSampleID: TLabeledEdit
      Left = 242
      Top = 32
      Width = 121
      Height = 21
      Hint = 'end of the range of sample numbers'
      CustomHint = BalloonHint1
      EditLabel.Width = 18
      EditLabel.Height = 13
      EditLabel.Caption = 'End'
      LabelPosition = lpLeft
      NumbersOnly = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
    end
    object UpDownSampleIDStart: TUpDown
      Left = 170
      Top = 30
      Width = 17
      Height = 25
      TabOrder = 2
      OnClick = UpDownSampleIDStartClick
    end
    object UpDownSampleIDEnd: TUpDown
      Left = 365
      Top = 30
      Width = 17
      Height = 25
      TabOrder = 3
      OnClick = UpDownSampleIDEndClick
    end
  end
  object edtStorLoc: TLabeledEdit
    Left = 72
    Top = 547
    Width = 131
    Height = 21
    Hint = 'location of the remaining material'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    EditLabel.Width = 77
    EditLabel.Height = 13
    EditLabel.Caption = 'storage location'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
  end
  object RadioGroupLocations: TRadioGroup
    Left = 0
    Top = 79
    Width = 385
    Height = 57
    Caption = 'Set Storage Location of: '
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'sample'
      'prep')
    TabOrder = 6
  end
  object btnSearch: TButton
    Left = 391
    Top = 82
    Width = 82
    Height = 54
    Hint = 'search and display samples according to given IDs'
    CustomHint = BalloonHint1
    Caption = 'Search'
    Default = True
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    OnClick = btnSearchClick
  end
  object BalloonHint1: TBalloonHint
    Left = 528
    Top = 24
  end
end
