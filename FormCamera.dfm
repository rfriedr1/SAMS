object CameraWindow: TCameraWindow
  Left = 392
  Top = 223
  Caption = 'Camera'
  ClientHeight = 541
  ClientWidth = 760
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 377
    Top = 0
    Width = 7
    Height = 541
    OnMoved = Splitter1Moved
    ExplicitHeight = 544
  end
  object Panel_Left: TPanel
    Left = 0
    Top = 0
    Width = 377
    Height = 541
    Align = alLeft
    TabOrder = 0
    inline Frame_Video1: TVideoFrame
      Left = 1
      Top = 1
      Width = 375
      Height = 539
      Align = alClient
      TabOrder = 0
      ExplicitLeft = 1
      ExplicitTop = 1
      ExplicitWidth = 375
      ExplicitHeight = 539
      inherited Panel_Top: TPanel
        Width = 375
        ExplicitWidth = 375
        DesignSize = (
          375
          104)
        inherited Label_Cameras: TLabel
          Width = 39
          Caption = 'Camera '
          ExplicitWidth = 39
        end
        inherited Label3: TLabel
          Width = 51
          ExplicitWidth = 51
        end
        inherited Label4: TLabel
          Width = 44
          ExplicitWidth = 44
        end
      end
      inherited Panel_Bottom: TPanel
        Width = 375
        Height = 435
        ExplicitWidth = 375
        ExplicitHeight = 435
        inherited PaintBox_Video: TPaintBox
          Width = 373
          Height = 393
        end
        inherited GroupBox1: TGroupBox
          Width = 373
          ExplicitWidth = 373
          inherited Label_fps: TLabel
            Width = 90
            ExplicitWidth = 90
          end
          inherited Label_VideoSize: TLabel
            Width = 50
            ExplicitWidth = 50
          end
          inherited Label2: TLabel
            Left = 322
            Width = 49
            ExplicitWidth = 49
          end
        end
      end
      inherited PopupMenu1: TPopupMenu
        Left = 128
        Top = 65528
      end
    end
  end
  object Panel_Right: TPanel
    Left = 384
    Top = 0
    Width = 376
    Height = 541
    Align = alClient
    TabOrder = 1
    object PaintBoxImage: TPaintBox
      Left = 1
      Top = 148
      Width = 374
      Height = 392
      Align = alBottom
      ExplicitLeft = 6
      ExplicitTop = 137
      ExplicitWidth = 363
    end
    object GroupBox1: TGroupBox
      Left = 1
      Top = 1
      Width = 374
      Height = 141
      Align = alTop
      Caption = 'GroupBox1'
      TabOrder = 0
      ExplicitLeft = 6
      ExplicitTop = 44
      ExplicitWidth = 339
      object Label1: TLabel
        Left = 32
        Top = 20
        Width = 32
        Height = 13
        Caption = 'MAMS'
      end
      object Label2: TLabel
        Left = 44
        Top = 47
        Width = 22
        Height = 13
        Caption = 'Path'
      end
      object lbl_SaveSuccessfull: TLabel
        Left = 208
        Top = 119
        Width = 37
        Height = 13
        Caption = 'Saved?'
      end
      object edtMAMS: TEdit
        Left = 80
        Top = 17
        Width = 233
        Height = 21
        TabOrder = 0
        OnChange = edtMAMSChange
      end
      object edtPathToImage: TJvDirectoryEdit
        Left = 80
        Top = 44
        Width = 233
        Height = 21
        DialogKind = dkWin32
        TabOrder = 1
        Text = 'PathToImage'
      end
      object btnSnapImage: TButton
        Left = 80
        Top = 73
        Width = 105
        Height = 40
        Caption = 'Snap'
        TabOrder = 2
        OnClick = btnSnapImageClick
      end
      object btnSave: TButton
        Left = 207
        Top = 73
        Width = 106
        Height = 40
        Caption = 'Save'
        TabOrder = 3
        OnClick = btnSaveClick
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 168
    Top = 65528
    object File1: TMenuItem
      Caption = '&File'
      object Quit1: TMenuItem
        Caption = '&Quit'
        OnClick = Quit1Click
      end
    end
  end
end
