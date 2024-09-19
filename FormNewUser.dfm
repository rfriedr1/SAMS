object frmNewUser: TfrmNewUser
  Left = 0
  Top = 0
  Caption = 'Add New User'
  ClientHeight = 700
  ClientWidth = 517
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object EdtLastName: TLabeledEdit
    Left = 32
    Top = 40
    Width = 465
    Height = 21
    EditLabel.Width = 50
    EditLabel.Height = 13
    EditLabel.Caption = 'Last Name'
    TabOrder = 0
  end
  object EdtFirstName: TLabeledEdit
    Left = 32
    Top = 88
    Width = 465
    Height = 21
    EditLabel.Width = 51
    EditLabel.Height = 13
    EditLabel.Caption = 'First Name'
    TabOrder = 1
  end
  object btnOK: TButton
    Left = 264
    Top = 640
    Width = 105
    Height = 33
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 2
  end
  object btnCancel: TButton
    Left = 120
    Top = 640
    Width = 105
    Height = 33
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 3
  end
  object grpUserInfo: TGroupBox
    Left = 0
    Top = 160
    Width = 521
    Height = 457
    Caption = 'User Info'
    TabOrder = 4
    object edtAddress1: TLabeledEdit
      Left = 32
      Top = 159
      Width = 465
      Height = 21
      EditLabel.Width = 48
      EditLabel.Height = 13
      EditLabel.Caption = 'Address 1'
      TabOrder = 0
    end
    object edtAddress2: TLabeledEdit
      Left = 32
      Top = 204
      Width = 465
      Height = 21
      EditLabel.Width = 48
      EditLabel.Height = 13
      EditLabel.Caption = 'Address 2'
      TabOrder = 1
    end
    object edtCity: TLabeledEdit
      Left = 168
      Top = 260
      Width = 329
      Height = 21
      EditLabel.Width = 19
      EditLabel.Height = 13
      EditLabel.Caption = 'City'
      TabOrder = 2
    end
    object edtCountry: TLabeledEdit
      Left = 168
      Top = 303
      Width = 329
      Height = 21
      EditLabel.Width = 39
      EditLabel.Height = 13
      EditLabel.Caption = 'Country'
      TabOrder = 3
    end
    object edtEMail: TLabeledEdit
      Left = 32
      Top = 404
      Width = 465
      Height = 21
      EditLabel.Width = 24
      EditLabel.Height = 13
      EditLabel.Caption = 'eMail'
      TabOrder = 4
    end
    object EdtInstitution: TLabeledEdit
      Left = 32
      Top = 108
      Width = 465
      Height = 21
      EditLabel.Width = 49
      EditLabel.Height = 13
      EditLabel.Caption = 'Institution'
      TabOrder = 5
    end
    object edtOrganisation: TLabeledEdit
      Left = 32
      Top = 63
      Width = 465
      Height = 21
      EditLabel.Width = 61
      EditLabel.Height = 13
      EditLabel.Caption = 'Organisation'
      TabOrder = 6
    end
    object edtPhone1: TLabeledEdit
      Left = 32
      Top = 356
      Width = 217
      Height = 21
      EditLabel.Width = 39
      EditLabel.Height = 13
      EditLabel.Caption = 'Phone 1'
      TabOrder = 7
    end
    object edtPhone2: TLabeledEdit
      Left = 280
      Top = 356
      Width = 217
      Height = 21
      EditLabel.Width = 39
      EditLabel.Height = 13
      EditLabel.Caption = 'Phone 2'
      TabOrder = 8
    end
    object edtZip: TLabeledEdit
      Left = 32
      Top = 260
      Width = 121
      Height = 21
      EditLabel.Width = 16
      EditLabel.Height = 13
      EditLabel.Caption = 'ZIP'
      TabOrder = 9
    end
  end
  object StaticText1: TStaticText
    Left = 432
    Top = 128
    Width = 59
    Height = 17
    Caption = 'StaticText1'
    TabOrder = 5
  end
end
