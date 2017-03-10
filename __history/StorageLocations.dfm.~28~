object FormStorageLocations: TFormStorageLocations
  Left = 0
  Top = 0
  ActiveControl = btnClose
  Caption = 'FormStorageLocations'
  ClientHeight = 718
  ClientWidth = 435
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    435
    718)
  PixelsPerInch = 120
  TextHeight = 17
  object edtSampleStorageLoc: TLabeledEdit
    Left = 40
    Top = 496
    Width = 351
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akBottom]
    EditLabel.Width = 152
    EditLabel.Height = 17
    EditLabel.Margins.Left = 4
    EditLabel.Margins.Top = 4
    EditLabel.Margins.Right = 4
    EditLabel.Margins.Bottom = 4
    EditLabel.Caption = 'Sample Storage Location'
    TabOrder = 0
  end
  object edtPrepStorageLoc: TLabeledEdit
    Left = 40
    Top = 556
    Width = 351
    Height = 21
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akBottom]
    EditLabel.Width = 141
    EditLabel.Height = 17
    EditLabel.Margins.Left = 4
    EditLabel.Margins.Top = 4
    EditLabel.Margins.Right = 4
    EditLabel.Margins.Bottom = 4
    EditLabel.Caption = 'prep'#39'd material loaction'
    TabOrder = 1
  end
  object btnClose: TButton
    Left = 260
    Top = 626
    Width = 131
    Height = 32
    Hint = 'close window without saving changes'
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akBottom]
    Cancel = True
    Caption = 'Close (ESC)'
    ModalResult = 8
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
  end
  object btnSave: TButton
    Left = 40
    Top = 626
    Width = 141
    Height = 32
    Hint = 'save storage locations'
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akBottom]
    Caption = 'Save'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    OnClick = btnSaveClick
  end
  object btnSearch: TButton
    Left = 80
    Top = 121
    Width = 261
    Height = 32
    Hint = 'search and display samples according to given IDs'
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Search (Return)'
    Default = True
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    OnClick = btnSearchClick
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 170
    Width = 434
    Height = 281
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -14
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object edtStatus: TEdit
    Left = 0
    Top = 691
    Width = 435
    Height = 27
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Align = alBottom
    BorderStyle = bsNone
    TabOrder = 6
  end
  object GroupBox1: TGroupBox
    Left = 10
    Top = 10
    Width = 420
    Height = 91
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Sample ID'#39's MAMS'
    TabOrder = 7
    object edtStartSampleID: TLabeledEdit
      Left = 50
      Top = 40
      Width = 151
      Height = 21
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      EditLabel.Width = 30
      EditLabel.Height = 17
      EditLabel.Margins.Left = 4
      EditLabel.Margins.Top = 4
      EditLabel.Margins.Right = 4
      EditLabel.Margins.Bottom = 4
      EditLabel.Caption = 'Start'
      LabelPosition = lpLeft
      NumbersOnly = True
      TabOrder = 0
      OnChange = edtStartSampleIDChange
    end
    object edtEndSampleID: TLabeledEdit
      Left = 250
      Top = 40
      Width = 151
      Height = 21
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      EditLabel.Width = 24
      EditLabel.Height = 17
      EditLabel.Margins.Left = 4
      EditLabel.Margins.Top = 4
      EditLabel.Margins.Right = 4
      EditLabel.Margins.Bottom = 4
      EditLabel.Caption = 'End'
      LabelPosition = lpLeft
      NumbersOnly = True
      TabOrder = 1
    end
  end
end
