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
  KeyPreview = True
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
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
  object PanelLeft: TPanel
    Left = 0
    Top = 0
    Width = 377
    Height = 541
    Align = alLeft
    TabOrder = 0
    object GroupBoxCameraSettings: TGroupBox
      Left = 1
      Top = 1
      Width = 375
      Height = 141
      Align = alTop
      Caption = 'Camera Setting'
      TabOrder = 0
      object ComboBoxCamera: TComboBox
        Left = 17
        Top = 29
        Width = 217
        Height = 21
        TabOrder = 0
        TabStop = False
        Text = 'select camera...'
        OnChange = ComboBoxCameraChange
        OnDropDown = ComboBoxCameraDropDown
      end
      object ComboBoxFormat: TComboBox
        Left = 240
        Top = 29
        Width = 121
        Height = 21
        TabOrder = 1
        TabStop = False
        OnChange = ComboBoxFormatChange
      end
    end
    object PanelPaintBox: TPanel
      Left = 1
      Top = 142
      Width = 375
      Height = 398
      Align = alClient
      TabOrder = 1
      OnResize = PanelPaintBoxResize
      object PaintBox: TPaintBox
        Left = 1
        Top = 1
        Width = 373
        Height = 396
        Align = alClient
        OnPaint = PaintBoxPaint
        ExplicitLeft = -47
        ExplicitTop = 105
        ExplicitWidth = 311
        ExplicitHeight = 255
      end
    end
  end
  object PanelRight: TPanel
    Left = 384
    Top = 0
    Width = 376
    Height = 541
    Align = alClient
    TabOrder = 1
    object PaintBoxImage: TPaintBox
      Left = 1
      Top = 142
      Width = 374
      Height = 398
      Align = alClient
      OnPaint = PaintBoxImagePaint
      ExplicitLeft = 6
      ExplicitTop = 137
      ExplicitWidth = 363
      ExplicitHeight = 392
    end
    object GroupBox1: TGroupBox
      Left = 1
      Top = 1
      Width = 374
      Height = 141
      Align = alTop
      Caption = 'Save Image'
      TabOrder = 0
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
        TabStop = False
        NumbersOnly = True
        TabOrder = 2
        OnChange = edtMAMSChange
      end
      object edtPathToImage: TJvDirectoryEdit
        Left = 80
        Top = 44
        Width = 233
        Height = 21
        TabStop = False
        DialogKind = dkWin32
        TabOrder = 0
        Text = 'PathToImage'
      end
      object btnSnapImage: TButton
        Left = 80
        Top = 73
        Width = 105
        Height = 40
        Hint = 'get image'
        Caption = 'Snap'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnSnapImageClick
      end
      object btnSave: TButton
        Left = 207
        Top = 73
        Width = 106
        Height = 40
        Hint = 'save image to given location'
        Caption = 'Save'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        TabStop = False
        OnClick = btnSaveClick
      end
      object SpinButton: TSpinButton
        Left = 315
        Top = 15
        Width = 20
        Height = 25
        DownGlyph.Data = {
          0E010000424D0E01000000000000360000002800000009000000060000000100
          200000000000D800000000000000000000000000000000000000008080000080
          8000008080000080800000808000008080000080800000808000008080000080
          8000008080000080800000808000000000000080800000808000008080000080
          8000008080000080800000808000000000000000000000000000008080000080
          8000008080000080800000808000000000000000000000000000000000000000
          0000008080000080800000808000000000000000000000000000000000000000
          0000000000000000000000808000008080000080800000808000008080000080
          800000808000008080000080800000808000}
        TabOrder = 4
        UpGlyph.Data = {
          0E010000424D0E01000000000000360000002800000009000000060000000100
          200000000000D800000000000000000000000000000000000000008080000080
          8000008080000080800000808000008080000080800000808000008080000080
          8000000000000000000000000000000000000000000000000000000000000080
          8000008080000080800000000000000000000000000000000000000000000080
          8000008080000080800000808000008080000000000000000000000000000080
          8000008080000080800000808000008080000080800000808000000000000080
          8000008080000080800000808000008080000080800000808000008080000080
          800000808000008080000080800000808000}
        OnDownClick = SpinButtonDownClick
        OnUpClick = SpinButtonUpClick
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 16
    Top = 104
    object File1: TMenuItem
      Caption = '&File'
      object Quit1: TMenuItem
        Caption = '&Quit'
        OnClick = Quit1Click
      end
    end
  end
  object WCamera: TWCamera
    AspectRatio = True
    BorderColor = clWhite
    PreviewControl = PanelPaintBox
    OnImageAvailable = WCameraImageAvailable
    Left = 16
    Top = 56
  end
  object SaveDialog: TSaveDialog
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 73
    Top = 105
  end
end
