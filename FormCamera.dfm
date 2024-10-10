object CameraWindow: TCameraWindow
  Left = 392
  Top = 223
  Caption = 'Camera'
  ClientHeight = 536
  ClientWidth = 991
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  Menu = MainMenu1
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 988
    Top = 0
    Height = 536
    Align = alRight
    OnMoved = Splitter1Moved
    ExplicitLeft = 168
    ExplicitTop = 392
    ExplicitHeight = 100
  end
  object PanelLeft: TPanel
    Left = 0
    Top = 0
    Width = 377
    Height = 536
    Align = alLeft
    TabOrder = 0
    ExplicitHeight = 541
    object GroupBoxCameraSettings: TGroupBox
      Left = 1
      Top = 1
      Width = 375
      Height = 141
      Align = alTop
      Caption = 'Camera Setting'
      TabOrder = 0
      object ComboBoxFormat: TComboBox
        Left = 198
        Top = 28
        Width = 121
        Height = 21
        TabOrder = 0
        TabStop = False
        OnChange = ComboBoxFormatChange
      end
    end
    object PanelPaintBox: TPanel
      Left = 1
      Top = 142
      Width = 375
      Height = 393
      Align = alClient
      TabOrder = 1
      OnResize = PanelPaintBoxResize
      ExplicitHeight = 398
      object PaintBox: TPaintBox
        Left = 1
        Top = 1
        Width = 373
        Height = 391
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
    Left = 613
    Top = 0
    Width = 375
    Height = 536
    Align = alClient
    TabOrder = 1
    ExplicitHeight = 541
    object PaintBoxImage: TPaintBox
      Left = 1
      Top = 142
      Width = 373
      Height = 393
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
      Width = 373
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
        TabOrder = 0
        Text = 'PathToImage'
        OnChange = edtPathToImageChange
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
  object ComboBoxCamera: TComboBox
    Left = 17
    Top = 29
    Width = 176
    Height = 21
    TabOrder = 2
    TabStop = False
    Text = 'select camera...'
    OnChange = ComboBoxCameraChange
    OnDropDown = ComboBoxCameraDropDown
  end
  object PanelCenter: TPanel
    Left = 377
    Top = 0
    Width = 236
    Height = 536
    Align = alLeft
    TabOrder = 3
    ExplicitHeight = 541
    object GroupBoxImageSettings: TGroupBox
      Left = 1
      Top = 1
      Width = 234
      Height = 534
      Align = alClient
      Anchors = [akLeft, akTop, akBottom]
      Caption = 'Image Settings'
      TabOrder = 0
      ExplicitHeight = 539
      object LabelWhiteBalance: TLabel
        Left = 10
        Top = 270
        Width = 69
        Height = 13
        Caption = 'White balance'
        FocusControl = TrackBarWhiteBalance
      end
      object LabelSharpness: TLabel
        Left = 10
        Top = 222
        Width = 50
        Height = 13
        Caption = 'Sharpness'
        FocusControl = TrackBarSharpness
      end
      object LabelFocus: TLabel
        Left = 10
        Top = 175
        Width = 29
        Height = 13
        Caption = 'Focus'
        FocusControl = TrackBarFocus
      end
      object LabelExposure: TLabel
        Left = 10
        Top = 127
        Width = 44
        Height = 13
        Caption = 'Exposure'
        FocusControl = TrackBarExposure
      end
      object LabelContrast: TLabel
        Left = 10
        Top = 79
        Width = 39
        Height = 13
        Caption = 'Contrast'
        FocusControl = TrackBarContrast
      end
      object LabelBrightness: TLabel
        Left = 10
        Top = 32
        Width = 49
        Height = 13
        Caption = 'Brightness'
        FocusControl = TrackBarBrightness
      end
      object TrackBarWhiteBalance: TTrackBar
        Left = 5
        Top = 285
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 0
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxWhiteBalance: TCheckBox
        Left = 151
        Top = 269
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 1
        OnClick = CheckBoxClick
      end
      object TrackBarSharpness: TTrackBar
        Left = 5
        Top = 237
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 2
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxSharpness: TCheckBox
        Left = 151
        Top = 221
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 3
        OnClick = CheckBoxClick
      end
      object TrackBarFocus: TTrackBar
        Left = 5
        Top = 190
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 4
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxFocus: TCheckBox
        Left = 149
        Top = 174
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 5
        OnClick = CheckBoxClick
      end
      object TrackBarExposure: TTrackBar
        Left = 5
        Top = 142
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 6
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxExposure: TCheckBox
        Left = 149
        Top = 126
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 7
        OnClick = CheckBoxClick
      end
      object TrackBarContrast: TTrackBar
        Left = 5
        Top = 94
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 8
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxContrast: TCheckBox
        Left = 149
        Top = 78
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 9
        OnClick = CheckBoxClick
      end
      object TrackBarBrightness: TTrackBar
        Left = 5
        Top = 47
        Width = 200
        Height = 26
        PageSize = 1
        TabOrder = 10
        ThumbLength = 7
        TickMarks = tmTopLeft
        OnChange = TrackBarChange
      end
      object CheckBoxBrightness: TCheckBox
        Left = 149
        Top = 31
        Width = 49
        Height = 17
        Caption = 'Auto'
        Enabled = False
        TabOrder = 11
        OnClick = CheckBoxClick
      end
      object ButtonDefault: TButton
        Left = 72
        Top = 488
        Width = 75
        Height = 25
        Caption = 'Default'
        TabOrder = 12
        OnClick = ButtonDefaultClick
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
  object wcamera: TWCamera
    AspectRatio = True
    BorderColor = clWhite
    PreviewControl = PanelPaintBox
    OnImageAvailable = wcameraImageAvailable
    Left = 16
    Top = 56
  end
  object SaveDialog: TSaveDialog
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 73
    Top = 105
  end
end
