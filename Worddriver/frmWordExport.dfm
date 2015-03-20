object WordExportForm: TWordExportForm
  Left = 481
  Top = 283
  BorderStyle = bsSizeToolWin
  Caption = 'Export data to word'
  ClientHeight = 524
  ClientWidth = 710
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  Scaled = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Ptop: TPanel
    Left = 0
    Top = 0
    Width = 710
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      710
      33)
    object LFEDoc: TLabel
      Left = 8
      Top = 4
      Width = 113
      Height = 24
      Alignment = taRightJustify
      AutoSize = False
      Caption = 'Starting &document'
      Layout = tlCenter
    end
    object FEDoc: TJvFilenameEdit
      Left = 128
      Top = 4
      Width = 469
      Height = 24
      DefaultExt = '.doc'
      Filter = 'Word documents|*.doc;*.docx|All files|*.*'
      DialogOptions = [ofHideReadOnly, ofFileMustExist]
      DialogTitle = 'Selecteer het startdocument'
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 0
    end
    object Bstart: TButton
      Left = 605
      Top = 4
      Width = 97
      Height = 24
      Action = AStart
      Anchors = [akTop, akRight]
      TabOrder = 1
    end
  end
  object PCExport: TPageControl
    Left = 0
    Top = 33
    Width = 710
    Height = 458
    ActivePage = TSoptions
    Align = alClient
    TabOrder = 1
    object TSoptions: TTabSheet
      Caption = '&Options'
      ImageIndex = 1
      object GValues: TJvDrawGrid
        Left = 0
        Top = 241
        Width = 702
        Height = 186
        Align = alClient
        ColCount = 2
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing]
        TabOrder = 0
        OnDrawCell = GValuesDrawCell
        OnGetEditText = GValuesGetEditText
        OnSetEditText = GValuesSetEditText
        DrawButtons = False
        ColWidths = (
          224
          424)
      end
      object POptions: TPanel
        Left = 0
        Top = 0
        Width = 702
        Height = 241
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 1
        object PButtons: TPanel
          Left = 0
          Top = 208
          Width = 702
          Height = 33
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 0
          DesignSize = (
            702
            33)
          object SBadd: TSpeedButton
            Left = 0
            Top = 7
            Width = 23
            Height = 22
            Action = AAdd
            Flat = True
            Glyph.Data = {
              42020000424D4202000000000000420000002800000010000000100000000100
              1000030000000002000000000000000000000000000000000000007C0000E003
              00001F0000001F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C00401F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C004000401F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C10001F7C
              1F7C1F7C1F7C1F7C0040004000401F7C1F7C1F7C1F7C1F7C10001F7C10001F7C
              10001F7C1F7C1F7C00400040004000401F7C1F7C1F7C1F7C1F7C100010001000
              1F7C1F7C1F7C1F7C004000400040004000401F7C1F7C10001000100010001000
              100010001F7C1F7C00400040004000401F7C1F7C1F7C1F7C1F7C100010001000
              1F7C1F7C1F7C1F7C0040004000401F7C1F7C1F7C1F7C1F7C10001F7C10001F7C
              10001F7C1F7C1F7C004000401F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C10001F7C
              1F7C1F7C1F7C1F7C00401F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C}
            ParentShowHint = False
            ShowHint = True
          end
          object SBDelete: TSpeedButton
            Left = 24
            Top = 7
            Width = 23
            Height = 22
            Flat = True
            Glyph.Data = {
              42020000424D4202000000000000420000002800000010000000100000000100
              1000030000000002000000000000000000000000000000000000007C0000E003
              00001F0000001F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C0040104210421F7C1F7C1F7C1F7C1F7C1042
              10421F7C1F7C00401F7C1F7C1F7C00400040104210421F7C1F7C1F7C1F7C0040
              10421F7C1F7C004000401F7C1F7C10420040004010421F7C1F7C1F7C10420040
              1F7C1F7C1F7C0040004000401F7C1F7C00400040104210421F7C004000401042
              1F7C1F7C1F7C00400040004000401F7C10420040004010421042004000401F7C
              1F7C1F7C1F7C004000400040004000401F7C004000400040004000401F7C1F7C
              1F7C1F7C1F7C00400040004000401F7C1F7C104200400040004010421F7C1F7C
              1F7C1F7C1F7C0040004000401F7C104210420040004000400040104210421F7C
              1F7C1F7C1F7C004000401F7C1042004000401F7C1F7C10420040004010421042
              1F7C1F7C1F7C00401F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1042004000401042
              10421F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C00400040
              1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
              1F7C1F7C1F7C}
            ParentShowHint = False
            ShowHint = True
          end
          object ELExtraData: TLabel
            Left = 88
            Top = 4
            Width = 129
            Height = 29
            AutoSize = False
            Caption = 'E&xtra data'
            Layout = tlCenter
          end
          object BFetch: TButton
            Left = 541
            Top = 4
            Width = 155
            Height = 25
            Action = AFetchParams
            Anchors = [akTop, akRight]
            TabOrder = 0
          end
        end
        object GroupBox1: TGroupBox
          Left = 0
          Top = 0
          Width = 201
          Height = 208
          Align = alLeft
          Caption = 'Acties'
          TabOrder = 1
          object CBPrint: TCheckBox
            Left = 8
            Top = 16
            Width = 97
            Height = 24
            Caption = '&Print'
            TabOrder = 0
          end
          object CBSave: TCheckBox
            Left = 8
            Top = 40
            Width = 121
            Height = 24
            Caption = 'Sa&ve'
            TabOrder = 1
            OnClick = CBSaveClick
          end
          object CBKeepOpen: TCheckBox
            Left = 8
            Top = 62
            Width = 145
            Height = 24
            Caption = 'keep &Word open'
            TabOrder = 2
          end
          object CBkeepDocumentsOpen: TCheckBox
            Left = 8
            Top = 84
            Width = 171
            Height = 24
            Caption = 'keep &Documents open'
            TabOrder = 3
          end
          object CBMergeOnlyOpen: TCheckBox
            Left = 8
            Top = 106
            Width = 186
            Height = 24
            Caption = 'Only keep merge &doc open'
            TabOrder = 4
          end
        end
        object PTables: TPanel
          Left = 201
          Top = 0
          Width = 501
          Height = 208
          Align = alClient
          BevelOuter = bvNone
          Caption = 'PTables'
          TabOrder = 2
          DesignSize = (
            501
            208)
          object PCOptions: TPageControl
            Left = 0
            Top = 32
            Width = 501
            Height = 176
            ActivePage = TSSerial
            Align = alBottom
            Style = tsFlatButtons
            TabOrder = 0
            object TSSerial: TTabSheet
              Caption = 'TSSerial'
              TabVisible = False
              ExplicitLeft = 0
              ExplicitTop = 0
              ExplicitWidth = 0
              ExplicitHeight = 0
              DesignSize = (
                493
                166)
              object LDESave: TLabel
                Left = 8
                Top = 32
                Width = 83
                Height = 16
                Caption = '&Save to folder'
                Layout = tlCenter
              end
              object LCBTemplate: TLabel
                Left = 8
                Top = 60
                Width = 102
                Height = 16
                Caption = 'Naming &template'
                Layout = tlCenter
              end
              object EditLabel1: TLabel
                Left = 8
                Top = 113
                Width = 99
                Height = 16
                Caption = '&Name document'
                Layout = tlCenter
              end
              object CBCurrentRecordOnly: TCheckBox
                Left = 128
                Top = 8
                Width = 217
                Height = 24
                Caption = 'C&urrent record only'
                TabOrder = 0
              end
              object DESave: TJvDirectoryEdit
                Left = 128
                Top = 32
                Width = 353
                Height = 24
                DialogKind = dkWin32
                DialogText = 'Selecteer een map om de documenten op te slaan'
                DialogOptions = [sdAllowCreate, sdPerformCreate, sdPrompt]
                Anchors = [akLeft, akTop, akRight]
                TabOrder = 1
              end
              object CBTemplate: TComboBox
                Left = 128
                Top = 60
                Width = 321
                Height = 24
                ItemHeight = 0
                TabOrder = 2
              end
              object FEMergeDoc: TJvFilenameEdit
                Left = 128
                Top = 113
                Width = 353
                Height = 24
                DialogKind = dkSave
                DefaultExt = '.doc'
                Filter = 'Word documents|*.doc;*.docx|All files|*.*'
                Anchors = [akLeft, akTop, akRight]
                TabOrder = 3
              end
              object CBMergeDoc: TCheckBox
                Left = 128
                Top = 88
                Width = 241
                Height = 24
                Caption = 'Create &merge document'
                TabOrder = 4
                OnClick = CBMergeDocClick
              end
              object CBFillFields: TCheckBox
                Left = 128
                Top = 144
                Width = 97
                Height = 17
                Caption = 'Fill in fields'
                TabOrder = 5
              end
            end
            object TSNewTable: TTabSheet
              Caption = 'TSNewTable'
              ImageIndex = 1
              TabVisible = False
              ExplicitLeft = 0
              ExplicitTop = 0
              ExplicitWidth = 0
              ExplicitHeight = 0
              DesignSize = (
                493
                166)
              object Label1: TLabel
                Left = 4
                Top = 36
                Width = 117
                Height = 24
                Alignment = taRightJustify
                AutoSize = False
                Caption = 'Ta&ble location field'
                Layout = tlCenter
              end
              object LLBFields: TLabel
                Left = 4
                Top = 64
                Width = 117
                Height = 24
                Alignment = taRightJustify
                AutoSize = False
                Caption = 'Use Col&umns'
                Layout = tlCenter
              end
              object SBUp: TSpeedButton
                Left = 408
                Top = 64
                Width = 23
                Height = 22
                Action = AUp
                Flat = True
                Glyph.Data = {
                  76010000424D7601000000000000760000002800000020000000100000000100
                  04000000000000010000120B0000120B00001000000000000000000000000000
                  800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
                  FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000333
                  3333333333777F33333333333309033333333333337F7F333333333333090333
                  33333333337F7F33333333333309033333333333337F7F333333333333090333
                  33333333337F7F33333333333309033333333333FF7F7FFFF333333000090000
                  3333333777737777F333333099999990333333373F3333373333333309999903
                  333333337F33337F33333333099999033333333373F333733333333330999033
                  3333333337F337F3333333333099903333333333373F37333333333333090333
                  33333333337F7F33333333333309033333333333337373333333333333303333
                  333333333337F333333333333330333333333333333733333333}
                NumGlyphs = 2
              end
              object SBDown: TSpeedButton
                Left = 408
                Top = 88
                Width = 23
                Height = 22
                Action = ADown
                Flat = True
                Glyph.Data = {
                  76010000424D7601000000000000760000002800000020000000100000000100
                  04000000000000010000120B0000120B00001000000000000000000000000000
                  800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
                  FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333303333
                  333333333337F33333333333333033333333333333373F333333333333090333
                  33333333337F7F33333333333309033333333333337373F33333333330999033
                  3333333337F337F33333333330999033333333333733373F3333333309999903
                  333333337F33337F33333333099999033333333373333373F333333099999990
                  33333337FFFF3FF7F33333300009000033333337777F77773333333333090333
                  33333333337F7F33333333333309033333333333337F7F333333333333090333
                  33333333337F7F33333333333309033333333333337F7F333333333333090333
                  33333333337F7F33333333333300033333333333337773333333}
                NumGlyphs = 2
              end
              object Label2: TLabel
                Left = 4
                Top = 8
                Width = 117
                Height = 24
                Alignment = taRightJustify
                AutoSize = False
                Caption = 'File&name'
                Layout = tlCenter
              end
              object CBTableTag: TComboBox
                Left = 128
                Top = 36
                Width = 177
                Height = 24
                ItemHeight = 0
                TabOrder = 0
                OnChange = CBTableTagChange
              end
              object LBFields: TCheckListBox
                Left = 128
                Top = 64
                Width = 281
                Height = 100
                ItemHeight = 16
                TabOrder = 1
              end
              object FENewTableFileName: TJvFilenameEdit
                Left = 128
                Top = 8
                Width = 357
                Height = 24
                DialogKind = dkSave
                DefaultExt = '.doc'
                Filter = 'Word documenten|*.doc|Alle bestanden|*.*'
                DialogOptions = [ofHideReadOnly, ofPathMustExist]
                DialogTitle = 'Geef een naam op voor het nieuwe document'
                Anchors = [akLeft, akTop, akRight]
                TabOrder = 2
              end
            end
            object TSFillTable: TTabSheet
              Caption = 'TSFillTable'
              ImageIndex = 2
              TabVisible = False
              ExplicitLeft = 0
              ExplicitTop = 0
              ExplicitWidth = 0
              ExplicitHeight = 0
              DesignSize = (
                493
                166)
              object Bestandsnaam: TLabel
                Left = 0
                Top = 8
                Width = 105
                Height = 25
                Alignment = taRightJustify
                AutoSize = False
                Caption = 'File&name'
                Layout = tlCenter
              end
              object FEFillTableFileName: TJvFilenameEdit
                Left = 112
                Top = 8
                Width = 365
                Height = 24
                DialogKind = dkSave
                DefaultExt = '.doc'
                Filter = 'Word documenten|*.doc|Alle bestanden|*.*'
                DialogOptions = [ofHideReadOnly, ofPathMustExist]
                DialogTitle = 'Geef een naam op voor het nieuwe document'
                Anchors = [akLeft, akTop, akRight]
                TabOrder = 0
              end
            end
          end
          object RBFillTable: TRadioButton
            Left = 368
            Top = 0
            Width = 124
            Height = 25
            Anchors = [akTop, akRight]
            Caption = '&Fill existing table'
            TabOrder = 1
            OnClick = RBKindClick
          end
          object RBNewTable: TRadioButton
            Left = 197
            Top = 1
            Width = 113
            Height = 25
            Anchors = [akTop]
            Caption = 'New t&able'
            TabOrder = 2
            OnClick = RBKindClick
          end
          object RBSerial: TRadioButton
            Left = 16
            Top = 0
            Width = 113
            Height = 25
            Caption = 'Se&rial letter'
            Checked = True
            TabOrder = 3
            TabStop = True
            OnClick = RBKindClick
          end
        end
      end
    end
    object TSData: TTabSheet
      Caption = '&Export data'
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 702
        Height = 36
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        DesignSize = (
          702
          36)
        object DBNavigator1: TDBNavigator
          Left = 456
          Top = 4
          Width = 225
          Height = 25
          DataSource = DSExport
          VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete, nbEdit, nbPost, nbCancel]
          Anchors = [akTop, akRight]
          TabOrder = 0
        end
      end
      object GExportData: TDBGrid
        Left = 0
        Top = 36
        Width = 702
        Height = 391
        Align = alClient
        DataSource = DSExport
        Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -13
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
  end
  object PBottom: TPanel
    Left = 0
    Top = 491
    Width = 710
    Height = 33
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    DesignSize = (
      710
      33)
    object BClose: TButton
      Left = 605
      Top = 4
      Width = 99
      Height = 25
      Action = AClose
      Anchors = [akTop, akRight]
      Cancel = True
      TabOrder = 0
    end
  end
  object ALMain: TActionList
    Left = 336
    Top = 40
    object AStart: TAction
      Caption = '&Create'
      Hint = 'Start export'
      OnExecute = AStartExecute
      OnUpdate = AStartUpdate
    end
    object AFetchParams: TAction
      Caption = '&Fetch missing fields'
      Hint = 'Fetch fields from document not in dataset'
      OnExecute = AFetchParamsExecute
      OnUpdate = AFetchParamsUpdate
    end
    object ACancel: TAction
      Caption = 'C&ancel'
      Hint = 'Cancel'
      OnExecute = ACancelExecute
      OnUpdate = ACancelUpdate
    end
    object AAdd: TAction
      Hint = 'Add value'
      ShortCut = 16462
      OnExecute = AAddExecute
    end
    object ADelete: TAction
      Hint = 'Delete value'
      ShortCut = 16452
      OnExecute = ADeleteExecute
      OnUpdate = ADeleteUpdate
    end
    object AClose: TAction
      Caption = 'C&lose'
      OnExecute = ACloseExecute
      OnUpdate = ACloseUpdate
    end
    object AUp: TAction
      Hint = 'Move up'
      OnExecute = AUpExecute
      OnUpdate = AUpUpdate
    end
    object ADown: TAction
      Hint = 'Move down'
      OnExecute = ADownExecute
      OnUpdate = ADownUpdate
    end
  end
  object DSExport: TDataSource
    DataSet = CDSExport
    Left = 268
    Top = 36
  end
  object CDSExport: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 232
    Top = 40
  end
end
