object MainForm: TMainForm
  Left = 347
  Top = 154
  Width = 609
  Height = 453
  Caption = 'Word Control Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    601
    425)
  PixelsPerInch = 120
  TextHeight = 16
  object Label1: TLabel
    Left = 8
    Top = 200
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = '&First name'
    FocusControl = EFirstName
    Layout = tlCenter
  end
  object Label2: TLabel
    Left = 232
    Top = 200
    Width = 73
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = '&Last name'
    FocusControl = ELastName
    Layout = tlCenter
  end
  object Label3: TLabel
    Left = 8
    Top = 232
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = 'T&itle'
    FocusControl = ETitle
    Layout = tlCenter
  end
  object Label4: TLabel
    Left = 8
    Top = 264
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = 'Address &1'
    FocusControl = EAddress1
    Layout = tlCenter
  end
  object Label5: TLabel
    Left = 8
    Top = 296
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = 'Address &2'
    FocusControl = EAddress2
    Layout = tlCenter
  end
  object Label6: TLabel
    Left = 8
    Top = 328
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = '&Zip'
    FocusControl = EZip
    Layout = tlCenter
  end
  object Label7: TLabel
    Left = 176
    Top = 328
    Width = 41
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = 'Cit&y'
    FocusControl = ECity
    Layout = tlCenter
  end
  object Label8: TLabel
    Left = 8
    Top = 360
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = '&Country'
    FocusControl = ECountry
    Layout = tlCenter
  end
  object Label9: TLabel
    Left = 8
    Top = 392
    Width = 72
    Height = 24
    Alignment = taRightJustify
    AutoSize = False
    Caption = '&Telephone'
    FocusControl = ETelephone
    Layout = tlCenter
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 40
    Width = 501
    Height = 153
    Anchors = [akLeft, akTop, akRight]
    DataSource = DSAddress
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -13
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
  end
  object DBNavigator1: TDBNavigator
    Left = 8
    Top = 8
    Width = 240
    Height = 25
    DataSource = DSAddress
    TabOrder = 2
  end
  object BClose: TButton
    Left = 516
    Top = 40
    Width = 75
    Height = 25
    Action = AClose
    Anchors = [akTop, akRight]
    TabOrder = 0
  end
  object BExport: TButton
    Left = 516
    Top = 72
    Width = 75
    Height = 25
    Action = AExport
    Anchors = [akTop, akRight]
    TabOrder = 1
  end
  object EFirstName: TDBEdit
    Left = 88
    Top = 200
    Width = 145
    Height = 24
    DataField = 'FirstName'
    DataSource = DSAddress
    TabOrder = 4
  end
  object ELastName: TDBEdit
    Left = 312
    Top = 200
    Width = 197
    Height = 24
    Anchors = [akLeft, akTop, akRight]
    DataField = 'LastName'
    DataSource = DSAddress
    TabOrder = 5
  end
  object ETitle: TDBEdit
    Left = 88
    Top = 232
    Width = 105
    Height = 24
    DataField = 'Title'
    DataSource = DSAddress
    TabOrder = 6
  end
  object EAddress1: TDBEdit
    Left = 88
    Top = 264
    Width = 421
    Height = 24
    Anchors = [akLeft, akTop, akRight]
    DataField = 'Address1'
    DataSource = DSAddress
    TabOrder = 7
  end
  object EAddress2: TDBEdit
    Left = 88
    Top = 296
    Width = 421
    Height = 24
    Anchors = [akLeft, akTop, akRight]
    DataField = 'Address2'
    DataSource = DSAddress
    TabOrder = 8
  end
  object ECity: TDBEdit
    Left = 216
    Top = 328
    Width = 293
    Height = 24
    Anchors = [akLeft, akTop, akRight]
    DataField = 'City'
    DataSource = DSAddress
    TabOrder = 10
  end
  object EZip: TDBEdit
    Left = 88
    Top = 328
    Width = 81
    Height = 24
    DataField = 'Zip'
    DataSource = DSAddress
    TabOrder = 9
  end
  object ECountry: TDBEdit
    Left = 88
    Top = 360
    Width = 421
    Height = 24
    Anchors = [akLeft, akTop, akRight]
    DataField = 'Country'
    DataSource = DSAddress
    TabOrder = 11
  end
  object ETelephone: TDBEdit
    Left = 88
    Top = 392
    Width = 121
    Height = 24
    DataField = 'Telephone'
    DataSource = DSAddress
    TabOrder = 12
  end
  object BSave: TButton
    Left = 516
    Top = 8
    Width = 75
    Height = 25
    Action = ASave
    TabOrder = 13
  end
  object CDSAddress: TClientDataSet
    Active = True
    Aggregates = <>
    FileName = 'addresses.xml'
    FieldDefs = <
      item
        Name = 'FirstName'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 30
      end
      item
        Name = 'LastName'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 30
      end
      item
        Name = 'Address1'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 100
      end
      item
        Name = 'Address2'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 100
      end
      item
        Name = 'Zip'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'City'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 50
      end
      item
        Name = 'Country'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 30
      end
      item
        Name = 'Telephone'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Title'
        Attributes = [faUnNamed]
        DataType = ftString
        Size = 10
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 16
    Top = 24
    Data = {
      AA0100009619E0BD0100000018000000090002000000030000003F0109466972
      73744E616D650100490010000100055749445448020002001E00084C6173744E
      616D650100490010000100055749445448020002001E00084164647265737331
      0100490010000100055749445448020002006400084164647265737332010049
      0010000100055749445448020002006400035A69700100490010000100055749
      445448020002000A000443697479010049001000010005574944544802000200
      320007436F756E7472790100490010000100055749445448020002001E000954
      656C6570686F6E65010049001000010005574944544802000200140005546974
      6C650100490010000100055749445448020002000A0001000A4348414E47455F
      4C4F470400820006000000010000000000000004000000020000000000000004
      00000004400000074D69636861656C0B56616E2043616E6E6579741247656D65
      656E7465737472616174203136330433303130094B657373656C2D4C6F074265
      6C6769756D0100034D722E0450410004526F736105526965626C074C75766967
      6E79064672616E6365044D72732E}
  end
  object DSAddress: TDataSource
    DataSet = CDSAddress
    Left = 48
    Top = 24
  end
  object ActionList1: TActionList
    Left = 384
    Top = 8
    object AClose: TAction
      Caption = 'Cl&ose'
      OnExecute = ACloseExecute
    end
    object ASave: TAction
      Caption = '&Save'
      OnExecute = ASaveExecute
    end
    object AExport: TAction
      Caption = 'E&xport'
      OnExecute = AExportExecute
      OnUpdate = AExportUpdate
    end
  end
end
