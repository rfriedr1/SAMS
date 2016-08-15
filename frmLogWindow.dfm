object LogWindow: TLogWindow
  Left = 0
  Top = 0
  Width = 688
  Height = 328
  AutoScroll = True
  Caption = 'Log Window'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ListBox: TListBox
    Left = 0
    Top = 0
    Width = 672
    Height = 289
    Align = alClient
    ItemHeight = 13
    PopupMenu = PopupMenu1
    TabOrder = 0
  end
  object PopupMenu1: TPopupMenu
    Left = 584
    Top = 16
    object Clear1: TMenuItem
      Caption = 'Clear'
      OnClick = Clear1Click
    end
    object toclipboard1: TMenuItem
      Caption = 'To clipboard'
      OnClick = toclipboard1Click
    end
  end
end
