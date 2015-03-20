object frmLogSql: TfrmLogSql
  Left = 0
  Top = 0
  Caption = 'frmLogSQL'
  ClientHeight = 462
  ClientWidth = 945
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object JvNetscapeSplitter1: TJvNetscapeSplitter
    Left = 0
    Top = 353
    Width = 945
    Height = 11
    Cursor = crVSplit
    Align = alTop
    Maximized = False
    Minimized = False
    ButtonCursor = crDefault
    ExplicitTop = 185
    ExplicitWidth = 598
  end
  object ListViewSQL: TListView
    Left = 0
    Top = 0
    Width = 945
    Height = 353
    Align = alTop
    Columns = <
      item
        Caption = 'Date-Time'
        Width = 132
      end
      item
        Caption = 'Type'
        Width = 131
      end
      item
        Caption = 'Command'
        Width = 131
      end
      item
        Caption = 'Status'
      end
      item
        Caption = 'RecordsAffected'
      end
      item
        Caption = 'CursorType'
      end
      item
        Caption = 'LockType'
      end>
    GridLines = True
    TabOrder = 0
    ViewStyle = vsReport
    OnChange = ListViewSQLChange
  end
  object SynEdit1: TMemo
    Left = 0
    Top = 364
    Width = 945
    Height = 98
    Align = alClient
    Lines.Strings = (
      'SynEdit1')
    TabOrder = 1
    ExplicitHeight = 19
  end
end
