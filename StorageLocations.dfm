object FormStorageLocations: TFormStorageLocations
  Left = 0
  Top = 0
  ActiveControl = btnClose
  Caption = 'FormStorageLocations'
  ClientHeight = 574
  ClientWidth = 348
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    348
    574)
  PixelsPerInch = 96
  TextHeight = 13
  object edtSampleStorageLoc: TLabeledEdit
    Left = 32
    Top = 397
    Width = 281
    Height = 21
    Hint = 'location of the remaining sample material'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    EditLabel.Width = 118
    EditLabel.Height = 13
    EditLabel.Caption = 'Sample Storage Location'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
  end
  object edtPrepStorageLoc: TLabeledEdit
    Left = 32
    Top = 445
    Width = 281
    Height = 21
    Hint = 'location of the remaining prep'#39'd material'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    EditLabel.Width = 111
    EditLabel.Height = 13
    EditLabel.Caption = 'prep'#39'd material loaction'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
  end
  object btnClose: TButton
    Left = 208
    Top = 501
    Width = 105
    Height = 25
    Hint = 'close window without saving changes'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    Cancel = True
    Caption = 'Close (ESC)'
    ModalResult = 8
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
  end
  object btnSave: TButton
    Left = 32
    Top = 501
    Width = 113
    Height = 25
    Hint = 'save storage locations to database'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akBottom]
    Caption = 'Save'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    OnClick = btnSaveClick
  end
  object btnSearch: TButton
    Left = 64
    Top = 97
    Width = 209
    Height = 25
    Hint = 'search and display samples according to given IDs'
    CustomHint = BalloonHint1
    Caption = 'Search (Return)'
    Default = True
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    OnClick = btnSearchClick
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 136
    Width = 347
    Height = 225
    Hint = 'list of found samples'
    CustomHint = BalloonHint1
    Anchors = [akLeft, akTop, akRight, akBottom]
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object edtStatus: TEdit
    Left = 0
    Top = 553
    Width = 348
    Height = 21
    Align = alBottom
    BorderStyle = bsNone
    TabOrder = 6
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 336
    Height = 73
    Caption = 'Sample ID'#39's MAMS'
    TabOrder = 7
    object edtStartSampleID: TLabeledEdit
      Left = 40
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
      Left = 200
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
  end
  object BalloonHint1: TBalloonHint
    Left = 288
    Top = 96
  end
end
