unit SAMS_Main;

{
  parsenewsamplefile ist Entry in Parsing

  to do
     change project status : manually , auto ?
     transfer age from target to sample
     qry all samples with age, but no report_letter or out_date NULL
     OpenVPN für Riesling


 %USERPROFILE%\SAMS

 10-5-10 zu ergänzen
     foreign key in project_t falsch
     report_letter ergänzen

}

interface

uses Windows, Classes, Graphics, Forms, Controls, Menus,
  Dialogs, StdCtrls, Buttons, ExtCtrls, ComCtrls, ImgList, StdActns,
  ActnList, ToolWin, Grids, DBGrids, DB, ADODB, DBCtrls, JvAppStorage,
  JvAppIniStorage, JvComponentBase, JvFormPlacement, Mask, JvExExtCtrls,
  JvNetscapeSplitter, JvExComCtrls, JvDateTimePicker, JvDBDateTimePicker,
  JvDataEmbedded, JvDBCheckBox, JvExMask, JvSpin, JvDBSpinEdit, JvDBGridExport,
  JvExControls, JvLabel, JvDBControls, JvExStdCtrls, JvEdit, JvValidateEdit,
  DBCGrids, MdRView, jpeg, JvToolEdit, JvDBLookup, JvDBLookupComboEdit,
  IdComponent, IdMessage, IdBaseComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP,
  Datasnap.DBClient, System.Actions, Vcl.Samples.Spin, JvExDBGrids, JvDBGrid,
  JvImage, JvWizardRouteMapList, JvWizard, JvExGrids, JvStringGrid, JvMarkupLabel,
  WordDriver, (*DBClient, *) Provider, MidasLib, (*Spin,*)
  JvInspector, Vcl.OleServer, Word2000, VclTee.TeeGDIPlus,
  VCLTee.TeEngine, VCLTee.Series, VCLTee.TeeProcs, VCLTee.Chart, VCLTee.DBChart,
  System.ImageList, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL, IdUserPassProvider, IdSASL, IdSASLUserPass, IdSASLLogin, StrUtils, frmStartScreen,
  frmLogWindow, FormNewUser, Vcl.FileCtrl(*, frxDesgn*), System.IOUtils, System.Types,
  Vcl.ValEdit, Math, Vcl.WinXCtrls, FormCamera, vFrames, iniFiles, Vcl.ExtDlgs,
  Vcl.Touch.Keyboard, IPPeerServer, Vcl.OleCtrls, SHDocVw,
  Datasnap.DSCommonServer, Datasnap.DSHTTP, Datasnap.DSHTTPWebBroker, JvLogFile, JvLogClasses,
  Vcl.AppEvnts, SysUtils, JvSplitter;

const
  myVersion = '1.9.9 Built: Oct-21-2025';

type
  TDragSource = (drgMaterial, drgFraction, drgType, drgPrep);

type
  TUser = record
    FirstName, LastName, Organisation, Institute, Address1,
      Address2, Town, Postcode, Phone1, Phone2, Fax, Email,
      WWW, Account, Comment: string;
  end;

  TInvoice = record
    FirstName, LastName, Organisation, Institute, Address1,
      Address2, Town, Postcode, Phone1, Phone2, Fax, Email,
      WWW, Account, Comment: string;
  end;

type
  TfrmMAMS = class(TForm)
    LoadDialog: TOpenDialog;
    ActionList1: TActionList;
    FileNew1: TAction;
    FileOpen1: TAction;
    FileSave1: TAction;
    FileSaveAs1: TAction;
    FileSend1: TAction;
    FileExit1: TAction;
    EditCut1: TEditCut;
    EditCopy1: TEditCopy;
    EditPaste1: TEditPaste;
    HelpAbout1: TAction;
    StatusBar: TStatusBar;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    FileNewItem: TMenuItem;
    FileOpenItem: TMenuItem;
    FileSaveItem: TMenuItem;
    FileSaveAsItem: TMenuItem;
    N1: TMenuItem;
    FileSendItem: TMenuItem;
    N2: TMenuItem;
    FileExitItem: TMenuItem;
    Edit1: TMenuItem;
    CutItem: TMenuItem;
    CopyItem: TMenuItem;
    PasteItem: TMenuItem;
    Help1: TMenuItem;
    HelpAboutItem: TMenuItem;
    SaveDialog: TSaveDialog;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton3: TToolButton;
    ToolButton5: TToolButton;
    btnSubmitter: TToolButton;
    btnNewSample: TToolButton;
    btnEditSample: TToolButton;
    PopupMenu1: TPopupMenu;
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    btnShowPrior: TToolButton;
    btnShowPretreat: TToolButton;
    JvFormStorage1: TJvFormStorage;
    JvAppIniFileStorage1: TJvAppIniFileStorage;
    WordExport: TJvDBGridWordExport;
    HTMLExport: TJvDBGridHTMLExport;
    XMLExport: TJvDBGridXMLExport;
    ExcelExport: TJvDBGridExcelExport;
    actLabStats: TAction;
    ImageList2: TImageList;
    ImageList3: TImageList;
    actSampleInfo: TAction;
    actUserProjects: TAction;
    actLabTasks: TAction;
    actReport: TAction;
    actLabPlan: TAction;
    actOptions: TAction;
    actInsertSamples: TAction;
    actUserInfo: TAction;
    pgtMain: TPageControl;
    tbsInsertSamples: TTabSheet;
    wizInputSamples: TJvWizard;
    wizStartPage: TJvWizardInteriorPage;
    Panel4: TPanel;
    Label14: TLabel;
    edtNewSamplesFilename: TJvFilenameEdit;
    grdPreviewUser: TJvStringGrid;
    wizInputInvoiceAddress: TJvWizardInteriorPage;
    wizInputProject: TJvWizardInteriorPage;
    GroupBox4: TGroupBox;
    Label15: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    edtProjectName: TLabeledEdit;
    edtProjectComment: TLabeledEdit;
    rgpPriority: TRadioGroup;
    edtPrice: TLabeledEdit;
    edtInDate: TJvDateTimePicker;
    edtDesiredDate: TJvDateTimePicker;
    cmbProjectType: TDBLookupComboBox;
    cmbResearchType: TDBLookupComboBox;
    wizSelectMaterial: TJvWizardInteriorPage;
    grdPreviewSamples: TJvStringGrid;
    JvWizardRouteMapList1: TJvWizardRouteMapList;
    tbsProjectsOfUser: TTabSheet;
    gbxProjectsOfUser: TGroupBox;
    cmbSubmitterNameProject: TDBLookupComboBox;
    GroupBox30: TGroupBox;
    grdProjects: TDBGrid;
    grdSamplesOfProject: TDBGrid;
    Panel2: TPanel;
    pgtSample: TPageControl;
    tbsSampleOfProject: TTabSheet;
    gbxSampleRight: TGroupBox;
    GroupBox35: TGroupBox;
    grdTarget: TMdRecordView;
    grdShowPrepSteps: TMdRecordView;
    GroupBox36: TGroupBox;
    gbxSampleLeft: TGroupBox;
    gbxSelectPretreatment: TGroupBox;
    GroupBox29: TGroupBox;
    GroupBox32: TGroupBox;
    GroupBox34: TGroupBox;
    grdSample: TMdRecordView;
    tbsFoto: TTabSheet;
    SampleFoto: TImage;
    Panel27: TPanel;
    tbsCalibration: TTabSheet;
    tbsLabPlan: TTabSheet;
    tbsLabTasks: TTabSheet;
    tbsUserReport: TTabSheet;
    Panel13: TPanel;
    Label41: TLabel;
    Label46: TLabel;
    Jahr: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    cmbSubmNameReport: TDBLookupComboBox;
    edtStartYear: TJvValidateEdit;
    edtEndYear: TJvValidateEdit;
    edtStartAna: TJvValidateEdit;
    edtEndAna: TJvValidateEdit;
    radgrpStatus: TRadioGroup;
    btnQuerySubmitter: TBitBtn;
    btnExport: TButton;
    rgpExport: TRadioGroup;
    grdSamplesOfSubmitter: TDBGrid;
    tbsSampleInfo: TTabSheet;
    tbsOptions: TTabSheet;
    GroupBox28: TGroupBox;
    tbsTypesMat: TTabSheet;
    GroupBox7: TGroupBox;
    grdPretreatment: TDBGrid;
    DBNavigator3: TDBNavigator;
    GroupBox6: TGroupBox;
    grdTypes: TDBGrid;
    DBNavigator4: TDBNavigator;
    GroupBox12: TGroupBox;
    grdMaterial: TDBGrid;
    DBNavigator2: TDBNavigator;
    pnlSetMaterial: TPanel;
    JvMarkupLabel2: TJvMarkupLabel;
    JvMarkupLabel3: TJvMarkupLabel;
    Label1: TLabel;
    cmbReportType: TDBLookupComboBox;
    Panel7: TPanel;
    rgpTask: TRadioGroup;
    gbxBatch: TGroupBox;
    lbxBatch: TListBox;
    Panel8: TPanel;
    gbxSamplesAvailable: TGroupBox;
    grdSamplesAvailable: TJvDBGrid;
    btnRefreshSamplesAvailable: TButton;
    btnSaveBatch: TButton;
    grpbox_TableHeadings: TGroupBox;
    GroupBox13: TGroupBox;
    Panel9: TPanel;
    btnLoadBatch: TButton;
    rgpLabTask: TRadioGroup;
    pgtLabTasks: TPageControl;
    tbsPrepTask: TTabSheet;
    tbsGraphTask: TTabSheet;
    GroupBox14: TGroupBox;
    gbxPresentTask: TGroupBox;
    Panel10: TPanel;
    GroupBox16: TGroupBox;
    GroupBox17: TGroupBox;
    dtStartTaskDate: TDateTimePicker;
    dtStartTaskTime: TDateTimePicker;
    btnStartTask: TButton;
    lbxPrepSteps: TListBox;
    lbPrepTimesHelp: TJvMarkupLabel;
    dtEndTaskDate: TDateTimePicker;
    dtEndTime: TDateTimePicker;
    btnEndTask: TButton;
    dtEndTaskTime: TDateTimePicker;
    JvNetscapeSplitter2: TJvNetscapeSplitter;
    Panel6: TPanel;
    lbxMaterial: TListBox;
    wizSelectPretreatment: TJvWizardInteriorPage;
    wizFinalPage: TJvWizardInteriorPage;
    wizSelectType: TJvWizardInteriorPage;
    pnlSetType: TPanel;
    lbxTypes: TListBox;
    ToolButton2: TToolButton;
    tbsUserInfo: TTabSheet;
    btnSetSampleReadyForGraph: TBitBtn;
    GroupBox15: TGroupBox;
    btnReport: TButton;
    edtWordTemplate: TJvFilenameEdit;
    cdsExport: TClientDataSet;
    dsExport: TDataSource;
    ToolButton4: TToolButton;
    actTables: TAction;
    gbxSampleIdetification: TGroupBox;
    gbxTarget: TGroupBox;
    gbxAdministration: TGroupBox;
    gbxAnaResults: TGroupBox;
    gbxSampleInfo: TGroupBox;
    btnDoSampleQuery: TBitBtn;
    Label29: TLabel;
    edtSampleNr: TJvValidateEdit;
    Label37: TLabel;
    Label39: TLabel;
    gbxSamplePretreatment: TGroupBox;
    Label50: TLabel;
    lbxFraction: TListBox;
    btnIncSampleNr: TSpinButton;
    gbxPrepSteps: TGroupBox;
    lbxDefinePrepSteps: TListBox;
    edtSampleInfoInDate: TJvDBDateTimePicker;
    edtSampleInfoDesiredDate: TJvDBDateTimePicker;
    lbxSteps: TListBox;
    edtPrepEnd: TJvDBDateTimePicker;
    Label54: TLabel;
    Label55: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    chkPrepDiscarded: TDBCheckBox;
    gpbUserInfo: TGroupBox;
    tbsLabStats: TTabSheet;
    gbxInPrep: TGroupBox;
    pnlPrepChoice: TPanel;
    Label25: TLabel;
    cmbPretreatMethod: TDBLookupComboBox;
    edtUserDocuments: TJvDirectoryEdit;
    Label53: TLabel;
    cmbProjectStatus: TDBLookupComboBox;
    memPrepComments: TDBMemo;
    gbxLabStatPlanned: TGroupBox;
    btnSaveChangesAdmin: TBitBtn;
    btnSaveChangesUserSuppliedInfo: TBitBtn;
    btnSaveChangesPrep: TBitBtn;
    btnSaveChangesGraph: TBitBtn;
    GroupBox8: TGroupBox;
    Label30: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Ta: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    edtSampleName: TDBEdit;
    edtSampleLabelNr: TDBEdit;
    edtSampleDescr1: TDBEdit;
    edtSampleDescr2: TDBEdit;
    cmbSampleMaterial: TDBLookupComboBox;
    cmbSampleType: TDBLookupComboBox;
    cmbSampleFraction: TDBLookupComboBox;
    edtSampleUserComment: TDBEdit;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    Label12: TLabel;
    edtCO2init: TDBEdit;
    edtCO2final: TDBEdit;
    edtH2init: TDBEdit;
    edtH2final: TDBEdit;
    edtReactTime: TDBEdit;
    edtTargetPressed: TJvDBDateTimePicker;
    chkTargetDiscarded: TDBCheckBox;
    edtCatalyst: TDBEdit;
    edtSampPrep1: TEdit;
    edtSampPrep3: TEdit;
    edtSampPrep2: TEdit;
    edtSampPrep4: TEdit;
    edtSampPrep5: TEdit;
    Panel3: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    grdPlanned: TDBGrid;
    grdInPrep: TDBGrid;
    grdWaitingForGraph: TDBGrid;
    JvDBStatusLabel1: TJvDBStatusLabel;
    chkShowOnHold: TCheckBox;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    JvNetscapeSplitter3: TJvNetscapeSplitter;
    GroupBox2: TGroupBox;
    grdFraction: TDBGrid;
    DBNavigator1: TDBNavigator;
    chkNotToBeDated: TDBCheckBox;
    chkAllProjects: TCheckBox;
    cmbProjectOfReport: TDBLookupComboBox;
    Label17: TLabel;
    edtBatchName: TLabeledEdit;
    Panel1: TPanel;
    JvDBStatusLabel3: TJvDBStatusLabel;
    grdActiveBatches: TJvDBGrid;
    grdSamplesOfLabTask: TJvDBGrid;
    btnTransferC14Age: TButton;
    Label18: TLabel;
    Label19: TLabel;
    grdReportHeadings: TJvStringGrid;
    chkCalBP: TCheckBox;
    Panel23: TPanel;
    lbExportHeader: TJvMarkupLabel;
    Button5: TButton;
    tbsProject: TTabSheet;
    GroupBox3: TGroupBox;
    Label26: TLabel;
    DBEdit_ProjectName: TDBEdit;
    Label27: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    Label40: TLabel;
    DBEdit_InvoiceNumber: TDBEdit;
    Label45: TLabel;
    JvDBDateTimePicker1: TJvDBDateTimePicker;
    Label65: TLabel;
    JvDBDateTimePicker2: TJvDBDateTimePicker;
    JvDBDateTimePicker3: TJvDBDateTimePicker;
    Label66: TLabel;
    DBEdit_Letter: TDBEdit;
    Label67: TLabel;
    ToolButton6: TToolButton;
    tbsAccounting: TTabSheet;
    GroupBox5: TGroupBox;
    Panel24: TPanel;
    grdProjectsSinceYear: TDBGrid;
    edtProjectsSinceYear: TJvSpinEdit;
    btnRefreshProjectsSinceYear: TButton;
    tnstor: TButton;
    Label68: TLabel;
    DBEdit16: TDBEdit;
    Label69: TLabel;
    Label70: TLabel;
    edtWeightEnd: TDBEdit;
    edtWeightStart: TDBEdit;
    Label71: TLabel;
    edtWeightCombustion: TDBEdit;
    Button3: TButton;
    DBEdit12: TDBEdit;
    Label72: TLabel;
    Label73: TLabel;
    cmbUsernameUserInfo: TDBLookupComboBox;
    Label74: TLabel;
    DBEdit17: TDBEdit;
    Label75: TLabel;
    DBEdit18: TDBEdit;
    Label76: TLabel;
    DBEdit19: TDBEdit;
    Label77: TLabel;
    DBEdit20: TDBEdit;
    Label78: TLabel;
    DBEdit21: TDBEdit;
    DBEdit22: TDBEdit;
    Label80: TLabel;
    DBEdit23: TDBEdit;
    Label81: TLabel;
    DBEdit24: TDBEdit;
    Label82: TLabel;
    DBEdit25: TDBEdit;
    Label83: TLabel;
    DBEdit26: TDBEdit;
    Label84: TLabel;
    DBEdit27: TDBEdit;
    Label85: TLabel;
    DBEdit28: TDBEdit;
    Label86: TLabel;
    DBEdit29: TDBEdit;
    Label87: TLabel;
    DBMemo1: TDBMemo;
    ToolButton7: TToolButton;
    Button6: TButton;
    btnAdmin: TToolButton;
    actAdmin: TAction;
    tbsAdmin: TTabSheet;
    Label90: TLabel;
    edtGraphDate: TJvDBDateTimePicker;
    JvDBStatusLabel4: TJvDBStatusLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label11: TLabel;
    Label10: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    edtC13: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    memTargetComments: TDBMemo;
    DBMemo2: TDBMemo;
    btnSaveUserProfile: TBitBtn;
    JvDBStatusLabel5: TJvDBStatusLabel;
    GroupBoxCreateTargets: TGroupBox;
    btnCreateSamples: TButton;
    rgpSampleType: TRadioGroup;
    edtNumberOfTargets: TJvSpinEdit;
    Label91: TLabel;
    btnSendPlannedToExcel: TButton;
    GroupBoxModifyDB: TGroupBox;
    btnDeleteUser: TButton;
    GroupBoxMeasuredSamples: TGroupBox;
    Panel28: TPanel;
    stgMonthStat: TStringGrid;
    edtMonthStat: TJvSpinEdit;
    btnSetGraphDate: TButton;
    Panel29: TPanel;
    grdShowUsers: TDBGrid;
    btnInsertNewUser: TButton;
    btnInsertExistingUser: TButton;
    grdInvoiceAddress: TJvStringGrid;
    grdShowInvoice: TDBGrid;
    btnInsertNewInvoice: TButton;
    btnInsertSelectedInvoice: TButton;
    btnSendInPrepToExcel: TButton;
    btnSendWaitForGraphToExcel: TButton;
    JvDBStatusLabel6: TJvDBStatusLabel;
    edtGraphBatch: TDBEdit;
    Label92: TLabel;
    btnTransferAgeToSample: TButton;
    JvDBStatusLabel7: TJvDBStatusLabel;
    btnCheckProjectStatus: TButton;
    lblProject: TLabel;
    GroupBoxMADBTasks: TGroupBox;
    lblOxa: TLabel;
    lblBlank: TLabel;
    lblTotal: TLabel;
    DBGrid1: TDBGrid;
    Label94: TLabel;
    edtSamplingDate: TJvDBDateTimePicker;
    chkSamplingThisYear: TCheckBox;
    Button7: TButton;
    chkLECurrent: TCheckBox;
    DBGrid2: TDBGrid;
    Button8: TButton;
    btnTransfer: TToolButton;
    tbsTransferAge: TTabSheet;
    gbxMagazines: TGroupBox;
    Panel16: TPanel;
    Label88: TLabel;
    tnsto: TButton;
    edtMagazineLimit: TJvSpinEdit;
    Panel25: TPanel;
    grdMagazines: TDBGrid;
    Panel26: TPanel;
    grdMagazineData: TDBGrid;
    GroupBox25: TGroupBox;
    StrGrdTargetData: TJvStringGrid;
    btnCalculateMean: TButton;
    edtMeanAge: TLabeledEdit;
    edtSigmaAge: TLabeledEdit;
    edtC13Mean: TLabeledEdit;
    btnTransferTargetData: TButton;
    grdPendingReports: TDBGrid;
    tbsProjectDocs: TTabSheet;
    lblReport: TLabel;
    chkPreliminaryReport: TCheckBox;
    Label28: TLabel;
    edtSampleInfoOutDate: TJvDBDateTimePicker;
    edtSample1: TJvEdit;
    edtSample2: TJvEdit;
    Sample1: TLabel;
    Label96: TLabel;
    btnSwapResults: TButton;
    tbsSendMail: TTabSheet;
    Panel31: TPanel;
    Panel30: TPanel;
    btnSendMail: TButton;
    MailMemo: TMemo;
    edtMailFrom: TLabeledEdit;
    edtMailTo: TLabeledEdit;
    edtMailSubject: TLabeledEdit;
    edtSMTPServer: TLabeledEdit;
    smtpSendMail: TIdSMTP;
    mesgMessage: TIdMessage;
    btnSendMailEnglish: TButton;
    btnSendMailGerman: TButton;
    lboxStatus: TListBox;
    actSendMail: TAction;
    ToolBtnSendMail: TToolButton;
    memoOxaBlank: TMemo;
    edtGraphitized: TJvDBDateTimePicker;
    Label95: TLabel;
    gbxEAData: TGroupBox;
    Label43: TLabel;
    Label44: TLabel;
    Label49: TLabel;
    edtSampleCPercent: TDBEdit;
    edtSampleNPercent: TDBEdit;
    edtSampleCNRatio: TDBEdit;
    btnSaveCN: TBitBtn;
    Panel32: TPanel;
    GroupBoxPendingReports: TGroupBox;
    JvDBStatusLabel8: TJvDBStatusLabel;
    JvDBStatusLabel9: TJvDBStatusLabel;
    JvDBStatusLabel10: TJvDBStatusLabel;
    Panel_Buttons: TPanel;
    btn_report_page: TButton;
    btn_UserInfo: TButton;
    Panel33: TPanel;
    GroupBox26: TGroupBox;
    StaticText1: TStaticText;
    DBEdit_UserNr: TDBEdit;
    DBEdit_AssocUserNr: TDBEdit;
    Label16: TLabel;
    JvDirEdt_Server_Image_Path: TJvDirectoryEdit;
    Label79: TLabel;
    JvDirEdt_Server_Report_Path: TJvDirectoryEdit;
    Label98: TLabel;
    btnDBPlot: TToolButton;
    tbsDBPlot: TTabSheet;
    gbxPlotProperties: TGroupBox;
    MemoDBPlotQuery: TMemo;
    gbxDBPlotQueryData: TGroupBox;
    gbxDBPlot: TGroupBox;
    DBGridDBPlot: TDBGrid;
    SplitterDBPlot: TSplitter;
    DBChart: TChart;
    Label99: TLabel;
    Label100: TLabel;
    GroupBox31: TGroupBox;
    btnAddNewUser: TButton;
    btnAddNewProject: TButton;
    btnAddNewSamples: TButton;
    btnDBPlotLoadQuery: TButton;
    btnDBPlotSaveQuery: TButton;
    btnAddNewUser2: TButton;
    gbxTargetData: TGroupBox;
    gbxTargetComment: TGroupBox;
    DBMemoTargetComment: TDBMemo;
    lbl_ProjectComment: TLabel;
    DBMemo_ProjectComment: TDBMemo;
    Series1: TPointSeries;
    btnFillDateToday: TButton;
    Label101: TLabel;
    lbl_Project_NumberOfSamples: TLabel;
    Panel15: TGroupBox;
    Label42: TLabel;
    lbPrepNr: TLabel;
    edtSamplePrepNr: TJvSpinEdit;
    btnNewPrep: TBitBtn;
    Panel11: TGroupBox;
    lbTarget: TLabel;
    lbTargetnr: TLabel;
    edtSampleTargetNr: TJvSpinEdit;
    btnNewTarget: TBitBtn;
    GroupBox19: TGroupBox;
    Label32: TLabel;
    Label31: TLabel;
    Label24: TLabel;
    Label93: TLabel;
    Label97: TLabel;
    Label3: TLabel;
    edtSampleProjectName: TDBEdit;
    edtSampleUser: TDBEdit;
    edtInvoiceNr: TDBEdit;
    edtReport: TDBEdit;
    edtProjectNr: TDBEdit;
    edtUserNr: TDBEdit;
    btnChangeProject: TButton;
    Label102: TLabel;
    Label103: TLabel;
    SSLHandler: TIdSSLIOHandlerSocketOpenSSL;
    IdUserPassProvider1: TIdUserPassProvider;
    IdSASLLogin1: TIdSASLLogin;
    OptionsTree: TTreeView;
    TabOptionsPages: TPageControl;
    btnSaveOptions: TButton;
    btnSampleNrUpDown: TUpDown;
    TabGeneral: TTabSheet;
    TabDatabase: TTabSheet;
    TabPaths: TTabSheet;
    TabEmail: TTabSheet;
    chkSampleNoLeftover: TDBCheckBox;
    chkPrepNoLeftover: TDBCheckBox;
    chkSendCopyToSender: TCheckBox;
    edtMailCC: TLabeledEdit;
    edtSampleStorageLoc: TDBEdit;
    edtPrepSampleStorageLoc: TDBEdit;
    Label104: TLabel;
    Label105: TLabel;
    ools1: TMenuItem;
    AssignStorageLocations1: TMenuItem;
    actStorageLocation: TAction;
    btnUserExportClipboard: TButton;
    YieldLabel: TLabel;
    Label107: TLabel;
    ToolButton10: TToolButton;
    lbWizFinalPage: TMemo;
    tbsDBInfo: TTabSheet;
    SearchDatabase1: TMenuItem;
    actSearchDatabase: TAction;
    dbchkFreeOfCharge1: TDBCheckBox;
    dbchkFreeOfCharge2: TDBCheckBox;
    chkFreeOfCharge: TCheckBox;
    Panel5: TPanel;
    Label106: TLabel;
    actLogWindow: TAction;
    OpenLogWindow1: TMenuItem;
    DBMemo_LabComment: TDBMemo;
    Label13: TLabel;
    DBMemo_ProjComment: TDBMemo;
    Comment: TLabel;
    GroupBox1: TGroupBox;
    GroupBox9: TGroupBox;
    DBMemo3: TDBMemo;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    cmbFilterSampleMaterial: TDBLookupComboBox;
    lblFilterSampleType: TLabel;
    DBedtMagazine: TDBEdit;
    Label38: TLabel;
    btnPrintBatch: TButton;
    StaticText4: TStaticText;
    btnProjectLetters: TButton;
    btnSampleSpreadsheet: TButton;
    Button1: TButton;
    Label108: TLabel;
    edtReportFileName: TEdit;
    Label109: TLabel;
    edtSaveReportToFolder: TJvDirectoryEdit;
    btnGuessReportName: TButton;
    StaticText5: TStaticText;
    edtFilenamePrepDocTemplate: TJvFilenameEdit;
    StaticText6: TStaticText;
    gbxProjectTools: TGroupBox;
    btnSetSamplingDateTo1950: TButton;
    Label110: TLabel;
    DBEditProjectNr: TDBEdit;
    tbsHome: TTabSheet;
    HomeGridPanel: TGridPanel;
    gpxHomeOxasBlanks: TGroupBox;
    gbxHomeProjectsDue: TGroupBox;
    ToolbtnHome: TToolButton;
    pnlHomeNumberOfOxas: TPanel;
    pnlHomeNumberOfBlanks: TPanel;
    Label111: TLabel;
    btnMonthStat: TBitBtn;
    btnPendingReports: TBitBtn;
    btnUpdateNoOxBlanks: TBitBtn;
    btnPlotQuery: TBitBtn;
    pnlHomeNumberofUnprepdSamples: TPanel;
    pnlHomeNumberofReadyForGraph: TPanel;
    pnlNumberofSamplesReadyForMeas: TPanel;
    BalloonHint1: TBalloonHint;
    btnSaveInvoiceNr: TBitBtn;
    gpxHomeExpress: TGroupBox;
    pnlHomeNumberofExpress: TPanel;
    DBGrid3: TDBGrid;
    DBGrid4: TDBGrid;
    DBGridHomeExpressSamples: TDBGrid;
    lblSelectedPlotRow: TLabel;
    btn_RemoveProject: TButton;
    DBCheckBox1: TDBCheckBox;
    ActivityIndicator_processed_samples: TActivityIndicator;
    actCamera: TAction;
    Action11: TMenuItem;
    grpboxAppOptions: TGroupBox;
    StaticText8: TStaticText;
    c: TStaticText;
    StaticText9: TStaticText;
    JvAppIniFileStoragePrep: TJvAppIniFileStorage;
    btn_initest: TButton;
    ImageFilesListBox: TListBox;
    Splitter1: TSplitter;
    SpeedButton1: TSpeedButton;
    SavePictureDialog: TSavePictureDialog;
    Label112: TLabel;
    Label113: TLabel;
    Label114: TLabel;
    dbchkreturnToSender2: TDBCheckBox;
    dbchkCNIsotopA: TDBCheckBox;
    ToolButtonSampleExchange: TToolButton;
    SampleExchange: TTabSheet;
    DBGridSampleExchange: TDBGrid;
    RadioGroupSampleExchange: TRadioGroup;
    Panel12: TPanel;
    btnExchange: TSpeedButton;
    DBCheckBoxReturnedToSender: TDBCheckBox;
    DBCheckBoxReturnedToSender2: TDBCheckBox;
    dbchkCNIsotopAMoved: TDBCheckBox;
    DBEditMANr: TDBEdit;
    Label115: TLabel;
    chkInsertReturnToSender: TCheckBox;
    checkInsertAllCNIsotopA: TCheckBox;
    Label116: TLabel;
    DBEdit3: TDBEdit;
    Label117: TLabel;
    DBEdit13: TDBEdit;
    Label118: TLabel;
    Label119: TLabel;
    tbsTouch: TTabSheet;
    TouchKeyboardWeights: TTouchKeyboard;
    Label120: TLabel;
    edtTouchWeightsMAMS: TJvValidateEdit;
    edtTouchWeightsPrepNr: TJvValidateEdit;
    edtTouchWeightsTargetNr: TJvValidateEdit;
    Label121: TLabel;
    Label122: TLabel;
    DBedtTouchWeightsBeforePrep: TDBEdit;
    Label123: TLabel;
    DBedtTouchWeightsAfterPrep: TDBEdit;
    Label124: TLabel;
    DBedtTouchWeightsCombustion: TDBEdit;
    Label125: TLabel;
    ToolButtonTouch: TToolButton;
    btnTouchWeightsMAMSUp: TBitBtn;
    btnTouchWeightsPrepNrUp: TBitBtn;
    btnTouchWeightsTargetNrUp: TBitBtn;
    btnTouchWeightsMAMSDown: TBitBtn;
    btnTouchWeightsPrepNrDown: TBitBtn;
    btnTouchWeightsTargetNrDown: TBitBtn;
    lblTouchWeightsNPrep: TLabel;
    lblTouchWeightsNTargets: TLabel;
    btnTouchWeightsPrepSave: TBitBtn;
    DBedtTouchWeightsLastName: TDBEdit;
    DBedtTouchWeightsUserLabel: TDBEdit;
    lblTouchWeightsLastName: TLabel;
    Label126: TLabel;
    DBchkTouchWeightsSampleNoLeftover: TDBCheckBox;
    GroupBoxTouchWeightsPrep: TGroupBox;
    GroupBoxTouchWeightsGraph: TGroupBox;
    DBchkTouchWeightsPrepNoLeftover: TDBCheckBox;
    btnTouchWeightsPrepNeedsSaving: TSpeedButton;
    gbxTouchWeightsGraphBatch: TGroupBox;
    btnTouchWeightsAddToGraphBatch: TButton;
    edtTouchWeightsBatchName: TEdit;
    Label127: TLabel;
    btnTouchWeightsBatchNameMag: TButton;
    btnTouchWeightsBatchNameAge1: TButton;
    btnTouchWeightsBatchNameAge2: TButton;
    pnlTouchWeightsMAMS: TPanel;
    btnTouchWeightsGraphSave: TBitBtn;
    btnTouchWeightsGraphNeedsSaving: TSpeedButton;
    ListBoxTouchWeightsBatch: TListBox;
    btnTouchWeightsSaveBatch: TBitBtn;
    btnTouchWeightsClearGraphBatchList: TBitBtn;
    DBedtTouchWeightsUserLabelNumber: TDBEdit;
    Label128: TLabel;
    DBMemoTouchWeightsTargetComment: TDBMemo;
    Label129: TLabel;
    btnTouchWeightsGraphBatchNeedsSaving: TSpeedButton;
    DBMemoTouchWeightsPrepComment: TDBMemo;
    Label130: TLabel;
    DBDateTimeTouchPrepStart: TJvDBDateTimePicker;
    DBDateTimeTouchPrepEnd: TJvDBDateTimePicker;
    Label131: TLabel;
    btnTouchPrepDateStartToday: TButton;
    btnTouchPrepDateEndToday: TButton;
    edtPrepStart: TJvDBDateTimePicker;
    Label132: TLabel;
    Label133: TLabel;
    CheckBoxTouchPrepWeightsAutoConversion: TCheckBox;
    Label134: TLabel;
    Label135: TLabel;
    DSRESTWebDispatcher1: TDSRESTWebDispatcher;
    btnOpenOxcal: TButton;
    CheckBoxTouchGraphWeightsAutoConversion: TCheckBox;
    Panel14: TPanel;
    Panel17: TPanel;
    WebBrowserCalibration: TWebBrowser;
    btnSampleInfoShowAllSamplesOfProject: TButton;
    lblYTouchYieldLabel: TLabel;
    lblTouchYieldValue: TLabel;
    lblLoadingFoto: TLabel;
    btnPrepCard: TButton;
    JvDirEdt_PrepCards_Path: TJvDirectoryEdit;
    Label136: TLabel;
    cdsPrepBatch: TClientDataSet;
    dsPrepBatch: TDataSource;
    JvDirEdt_PrepBatch_Path: TJvDirectoryEdit;
    Label137: TLabel;
    DBedtTouchWeightsMediumPrep: TDBEdit;
    DBedtTouchWeightsMedium2Prep: TDBEdit;
    Label138: TLabel;
    Label139: TLabel;
    Label140: TLabel;
    Label141: TLabel;
    JvLogFile: TJvLogFile;
    ApplicationEvents: TApplicationEvents;
    DBEditWeightMedium: TDBEdit;
    DBEditWeightMedium2: TDBEdit;
    Label142: TLabel;
    Label143: TLabel;
    Label144: TLabel;
    Label145: TLabel;
    dbgrdLeHeCurrents: TDBGrid;
    DBedtLECurrent: TDBEdit;
    DBedtHECurrent: TDBEdit;
    Label146: TLabel;
    Label147: TLabel;
    pnlHomeNumberOfSamplesInPrep: TPanel;
    CSVExportClipBoard: TJvDBGridCSVExport;
    btnReloadTransferMagazines: TBitBtn;
    DBCheckBoxTouchPrepDiscarded: TDBCheckBox;
    DBCheckBoxTouchTargetDiscarded: TDBCheckBox;
    DBEditTouchPrep1: TDBEdit;
    Label89: TLabel;
    DBEditTouchPrep2: TDBEdit;
    DBEditTouchPrep3: TDBEdit;
    DBEditTouchMaterial: TDBEdit;
    Label148: TLabel;
    DBEditTouchPrep4: TDBEdit;
    btnTouchWeightsBatchNameAutosampler: TButton;
    btnTouchWeightsBatchNameAge641: TButton;
    btnTouchWeightsBatchNameAge642: TButton;
    edtLimitTransferMA: TEdit;
    Label149: TLabel;
    DBCheckBoxTouchCNAna: TDBCheckBox;
    DBCheckBoxTouchMovedToCN: TDBCheckBox;
    Button2: TButton;
    pnlHomeNumberofUnprepdSamplesBone: TPanel;
    pnlHomeNumberofUnprepdSamplesTeeth: TPanel;
    ToolButtonPress: TToolButton;
    tbsMagazine: TTabSheet;
    grdMagazinesUnpressed: TDBGrid;
    Label150: TLabel;
    DBGridSamplesOfUnpressedMagazine: TDBGrid;
    Button4: TButton;
    GroupBox10: TGroupBox;
    btnMagazinePressedDateToday: TButton;
    pnlHomeNumberOfC6: TPanel;
    pnlHomeNumberOfC7: TPanel;
    pnlHomeNumberOfC8: TPanel;
    pnlHomeNumberOfPferde: TPanel;
    dbedtsStorLoc: TDBEdit;
    chkTouchReturnToSender: TDBCheckBox;
    chkTouchSampleArchived: TCheckBox;
    btnTouchWarningStorageLoc: TSpeedButton;
    JvAppIniFileStorageOptions: TJvAppIniFileStorage;
    JvFormStorageOptions: TJvFormStorage;
    chkTouchPrepArchived: TCheckBox;
    dbedtprepStorLoc: TDBEdit;
    btnTouchWarningStorageLocPrep: TSpeedButton;
    dbchkprepreturnToSender2: TDBCheckBox;
    DBCheckBoxPrepReturnedToSender2: TDBCheckBox;
    dbchkprepreturnToSender: TDBCheckBox;
    DBCheckBoxPrepReturnedToSender: TDBCheckBox;
    chkInsertprepReturnToSender: TCheckBox;
    chkTouchprepReturnToSender: TDBCheckBox;
    ActivityIndicator1: TActivityIndicator;
    ActivityIndicatorDBInfo: TActivityIndicator;
    GroupBox_Transfer_Filter: TGroupBox;
    edtTransferMagazinFilter: TLabeledEdit;
    lblCAbsWeight: TLabel;
    edtOptionsTurnover: TLabeledEdit;
    edtOptionsMinCAmount: TLabeledEdit;
    procedure grdSamplesOfProjectMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure grdSamplesOfProjectKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure wizStartPageBackButtonClick(Sender: TObject; var Stop: Boolean);
    procedure edtSampleTargetNrChange(Sender: TObject);
    procedure btnNewTargetClick(Sender: TObject);
    procedure edtSamplePrepNrChange(Sender: TObject);
    procedure btnNewPrepClick(Sender: TObject);
    procedure edtSampPrep5DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure edtSampPrep4DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure edtSampPrep3DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure edtSampPrep2DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure edtSampPrep1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure edtSampPrep1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure lbxStepsMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Button6Click(Sender: TObject);
    procedure lbxDefinePrepStepsMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lbxTypesMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lbxFractionDblClick(Sender: TObject);
    procedure lbxMaterialDblClick(Sender: TObject);
    procedure btnIncSampleNrDownClick(Sender: TObject);
    procedure btnIncSampleNrUpClick(Sender: TObject);
    procedure lbxFractionMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure cmbSubmitterNameProjectKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnDoSampleQueryClick(Sender: TObject);
    procedure tbsTypesMatExit(Sender: TObject);
    procedure actTablesExecute(Sender: TObject);
    procedure cmbSubmNameReportKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cmbSubmNameReportCloseUp(Sender: TObject);
    procedure btnReportClick(Sender: TObject);
    procedure btnSetSampleReadyForGraphClick(Sender: TObject);
    procedure rgpTaskClick(Sender: TObject);
    procedure lbxTypesDblClick(Sender: TObject);
    procedure edtProjectNameChange(Sender: TObject);
    procedure actUserInfoExecute(Sender: TObject);
    procedure wizSelectTypeEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizSelectPretreatmentEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    ///  <summary>actually insert user, project and samples into the database</summary>
    procedure wizFinalPageFinishButtonClick(Sender: TObject; var Stop: Boolean);
    procedure wizFinalPageEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizSelectPretreatmentBackButtonClick(Sender: TObject;
      var Stop: Boolean);
    procedure wizSelectMaterialNextButtonClick(Sender: TObject;
      var Stop: Boolean);
    procedure lbxMaterialMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grdPreviewSamplesDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure grdPreviewSamplesDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure lbxLabTaskBatchClick(Sender: TObject);
    procedure btnNowEndClick(Sender: TObject);
    procedure btnNowStartClick(Sender: TObject);
    procedure btnEndTaskClick(Sender: TObject);
    procedure lbxPrepStepsClick(Sender: TObject);
    procedure btnStartTaskClick(Sender: TObject);
    procedure rgpLabTaskClick(Sender: TObject);
    procedure btnLoadBatchClick(Sender: TObject);
    procedure btnSaveBatchClick(Sender: TObject);
    procedure lbxBatchDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure lbxBatchDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure grdSamplesAvailableMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnRefreshSamplesAvailableClick(Sender: TObject);
    procedure grdSamplesOfProjectDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure wizInputSamplesFinishButtonClick(Sender: TObject);
    procedure wizStartPageEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizSelectMaterialEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizInputSamplesBackButtonClick(Sender: TObject);
    procedure btnInsertUserClick(Sender: TObject);
    procedure wizInputInvoiceAddressEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizInputSamplesCancelButtonClick(Sender: TObject);
    procedure wizInputProjectEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure edtNewSamplesFilenameAfterDialog(Sender: TObject;
      var AName: string; var AAction: Boolean);
    procedure grdPretreatmentStepsSelectCell(Sender: TObject; ACol,
      ARow: Integer; var CanSelect: Boolean);
    procedure cmbStep1CloseUp(Sender: TObject);
    procedure grdSamplesOfProjectCellClick(Column: TColumn);
    procedure grdProjectsCellClick(Column: TColumn);
    procedure cmbSubmitterNameProjectCloseUp(Sender: TObject);
    procedure actMagListExecute(Sender: TObject);
    procedure actOptionsExecute(Sender: TObject);
    procedure actLabPlanExecute(Sender: TObject);
    procedure actReportExecute(Sender: TObject);
    procedure actLabTasksExecute(Sender: TObject);
    procedure actUserProjectsExecute(Sender: TObject);
    procedure actSampleInfoExecute(Sender: TObject);
    procedure actInsertSamplesExecute(Sender: TObject);
    procedure btnQueryToPretreatClick(Sender: TObject);
    procedure actLabStatsExecute(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure grdSamplesOfSubmitterTitleClick(Column: TColumn);
    procedure btnQuerySubmitterClick(Sender: TObject);
    procedure grdSamplesOfSubmitterMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure pgtMainChange(Sender: TObject);
    procedure edtChangeDBClick(Sender: TObject);
    procedure edtSubmNameExit(Sender: TObject);
    procedure edtSubmNameEnter(Sender: TObject);
    procedure DBLookupComboBox2Exit(Sender: TObject);
    procedure DBLookupComboBox2Enter(Sender: TObject);
    procedure HelpAbout1Execute(Sender: TObject);
    procedure btnSaveChangesUserSuppliedInfoClick(Sender: TObject);
    procedure btnSaveChangesAdminClick(Sender: TObject);
    procedure btnSaveChangesPrepClick(Sender: TObject);
    procedure btnSaveChangesGraphClick(Sender: TObject);
    procedure grdPlannedTitleClick(Column: TColumn);
    procedure grdPlannedCellClick(Column: TColumn);
    procedure chkShowOnHoldClick(Sender: TObject);
    procedure grdTypesTitleClick(Column: TColumn);
    procedure grdMaterialTitleClick(Column: TColumn);
    procedure grdPretreatmentTitleClick(Column: TColumn);
    procedure grdFractionTitleClick(Column: TColumn);
    procedure chkAllProjectsClick(Sender: TObject);
    procedure edtBatchNameChange(Sender: TObject);
    procedure cmbPretreatMethodClick(Sender: TObject);
    procedure grdActiveBatchesCellClick(Column: TColumn);
    procedure grdSamplesOfLabTaskCellClick(Column: TColumn);
    procedure btnTransferC14AgeClick(Sender: TObject);
    procedure grdReportHeadingsExit(Sender: TObject);
    procedure pgtSampleChange(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure btnSaveInvoiceNrClick(Sender: TObject);
    procedure btnRefreshProjectsSinceYearClick(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure tnstorClick(Sender: TObject);
    procedure edtWeightEndChange(Sender: TObject);
    procedure edtWeightCombustionChange(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure edtWeightCombustionClick(Sender: TObject);
    procedure btnAdminClick(Sender: TObject);
    procedure tnstoClick(Sender: TObject);
    procedure grdMagazinesCellClick(Column: TColumn);
    procedure btnSaveUserProfileClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure cmbUsernameUserInfoEnter(Sender: TObject);
    procedure cmbUsernameUserInfoExit(Sender: TObject);
    procedure edtSampleNrKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnNewUserClick(Sender: TObject);
    procedure btnDeleteUserClick(Sender: TObject);
    procedure btnCreateSamplesClick(Sender: TObject);
    procedure btnSendPlannedToExcelClick(Sender: TObject);
    procedure btnMonthStatClick(Sender: TObject);
    procedure btnSetGraphDateClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure grdShowUsersCellClick(Column: TColumn);
    procedure btnInsertExistingUserClick(Sender: TObject);
    procedure btnInsertNewUserClick(Sender: TObject);
    procedure btnInsertNewInvoiceClick(Sender: TObject);
    procedure btnInsertSelectedInvoiceClick(Sender: TObject);
    procedure grdShowInvoiceCellClick(Column: TColumn);
    procedure btnSendInPrepToExcelClick(Sender: TObject);
    procedure btnSendWaitForGraphToExcelClick(Sender: TObject);
    procedure grdInPrepTitleClick(Column: TColumn);
    procedure grdWaitingForGraphTitleClick(Column: TColumn);
    procedure btnSetDiscardClick(Sender: TObject);
    procedure btnTransferAgeToSampleClick(Sender: TObject);
    procedure btnCheckProjectStatusClick(Sender: TObject);
    procedure chkSamplingThisYearClick(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure cmbSubmNameReportMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure cmbProjectOfReportCloseUp(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
    procedure grdMagazineDataDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure grdMagazineDataCellClick(Column: TColumn);
    procedure StrGrdTargetDataDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure StrGrdTargetDataClick(Sender: TObject);
    procedure btnCalculateMeanClick(Sender: TObject);
    procedure btnTransferTargetDataClick(Sender: TObject);
    procedure btnPendingReportsClick(Sender: TObject);
    procedure btnSaveCNClick(Sender: TObject);
    procedure btnSwapResultsClick(Sender: TObject);
    procedure btnSendMailClick(Sender: TObject);
    procedure btnSendMailGermanClick(Sender: TObject);
    procedure btnSendMailEnglishClick(Sender: TObject);
    procedure smtpSendMailStatus(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure actSendMailExecute(Sender: TObject);
    procedure grdSamplesOfProjectDblClick(Sender: TObject);
    procedure grdPendingReportsDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure grdPendingReportsDblClick(Sender: TObject);
    procedure btn_report_pageClick(Sender: TObject);
    procedure btn_UserInfoClick(Sender: TObject);
    procedure grdPlannedDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure grdWaitingForGraphDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure grdSamplesAvailableDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure grdSamplesAvailableDblClick(Sender: TObject);
    procedure grdSamplesOfSubmitterDblClick(Sender: TObject);
    procedure grdPlannedDblClick(Sender: TObject);
    procedure grdInPrepDblClick(Sender: TObject);
    procedure grdWaitingForGraphDblClick(Sender: TObject);
    procedure grdSamplesAvailableCellClick(Column: TColumn);
    procedure btnDBPlotClick(Sender: TObject);
    procedure btnPlotQueryClick(Sender: TObject);
    procedure DBChartSeriesGetMarkText(Sender:TChartSeries; ValueIndex:Integer; var MarkText:String);
    procedure GetAvailableOxasBlanks;
    procedure btnDBPlotSaveQueryClick(Sender: TObject);
    procedure btnDBPlotLoadQueryClick(Sender: TObject);
    procedure btnAddNewUser2Click(Sender: TObject);
    procedure btnAddNewUserClick(Sender: TObject);
    procedure edtWeightEndExit(Sender: TObject);
    procedure btnChangeProjectClick(Sender: TObject);
    procedure btnAddNewProjectClick(Sender: TObject);
    procedure btnAddNewSamplesClick(Sender: TObject);
    procedure btnFillDateTodayClick(Sender: TObject);
    procedure edtGraphDateChange(Sender: TObject);
    procedure edtTargetPressedChange(Sender: TObject);
    procedure edtGraphitizedChange(Sender: TObject);
    procedure edtCatalystChange(Sender: TObject);
    procedure edtCO2initChange(Sender: TObject);
    procedure edtCO2finalChange(Sender: TObject);
    procedure edtH2initChange(Sender: TObject);
    procedure edtH2finalChange(Sender: TObject);
    procedure edtReactTimeChange(Sender: TObject);
    procedure edtGraphBatchChange(Sender: TObject);
    procedure chkTargetDiscardedClick(Sender: TObject);
    procedure DBChartClickSeries(Sender: TCustomChart;
  Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
    procedure OptionsTreeClick(Sender: TObject);
    procedure btnSampleNrUpDownClick(Sender: TObject; Button: TUDBtnType);
    procedure actStorageLocationExecute(Sender: TObject);
    procedure btnUserExportClipboardClick(Sender: TObject);
    procedure ToolButton10Click(Sender: TObject);
    procedure actSearchDatabaseExecute(Sender: TObject);
    procedure dbchkFreeOfCharge1Click(Sender: TObject);
    procedure dbchkFreeOfCharge2Click(Sender: TObject);
    procedure grdProjectsDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure grdSamplesOfSubmitterDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actLogWindowExecute(Sender: TObject);
    procedure edtWeightStartChange(Sender: TObject);
    procedure grdInPrepDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnUpdateNoOxBlanksClick(Sender: TObject);
    procedure cmbFilterSampleMaterialClick(Sender: TObject);
    procedure btnProjectLettersClick(Sender: TObject);
    procedure edtReportFileNameChange(Sender: TObject);
    procedure btnGuessReportNameClick(Sender: TObject);
    procedure btnPrintBatchClick(Sender: TObject);
    procedure btnSetSamplingDateTo1950Click(Sender: TObject);
    procedure ToolBar1Click(Sender: TObject);
    procedure ToolbtnHomeClick(Sender: TObject);
    procedure DBGridHomeExpressSamplesDblClick(Sender: TObject);
    procedure DBGridHomeExpressSamplesDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGridDBPlotCellClick(Column: TColumn);
    procedure btn_RemoveProjectClick(Sender: TObject);
    procedure pgtSampleExit(Sender: TObject);
    procedure actCameraExecute(Sender: TObject);
    procedure grdPendingReportsTitleClick(Column: TColumn);
    procedure btnSaveOptionsClick(Sender: TObject);
    procedure btn_initestClick(Sender: TObject);
    procedure ImageFilesListBoxClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnExchangeClick(Sender: TObject);
    procedure ToolButtonSampleExchangeClick(Sender: TObject);
    procedure DBGridSampleExchangeDblClick(Sender: TObject);
    procedure DBGridSampleExchangeDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGridSampleExchangeCellClick(Column: TColumn);
    procedure btnTouchWeightsMAMSUpClick(Sender: TObject);
    procedure btnTouchWeightsMAMSDownClick(Sender: TObject);
    procedure btnTouchWeightsPrepNrUpClick(Sender: TObject);
    procedure btnTouchWeightsPrepNrDownClick(Sender: TObject);
    procedure btnTouchWeightsTargetNrUpClick(Sender: TObject);
    procedure btnTouchWeightsTargetNrDownClick(Sender: TObject);
    procedure edtSampleNrChange(Sender: TObject);
    procedure ToolButtonTouchClick(Sender: TObject);
    procedure edtTouchWeightsMAMSChange(Sender: TObject);
    procedure edtTouchWeightsTargetNrChange(Sender: TObject);
    procedure edtTouchWeightsPrepNrChange(Sender: TObject);
    procedure edtTouchWeightsMAMSKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnTouchWeightsPrepSaveClick(Sender: TObject);
    procedure DBedtTouchWeightsBeforePrepClick(Sender: TObject);
    procedure DBedtTouchWeightsAfterPrepClick(Sender: TObject);
    procedure DBedtTouchWeightsCombustionClick(Sender: TObject);
    procedure edtTouchWeightsMAMSClick(Sender: TObject);
    procedure edtTouchWeightsPrepNrClick(Sender: TObject);
    procedure edtTouchWeightsTargetNrClick(Sender: TObject);
    procedure DBedtTouchWeightsBeforePrepKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBedtTouchWeightsAfterPrepKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBedtTouchWeightsCombustionKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBchkTouchWeightsSampleNoLeftoverClick(Sender: TObject);
    procedure DBchkTouchWeightsPrepNoLeftoverClick(Sender: TObject);
    procedure btnTouchWeightsBatchNameMagClick(Sender: TObject);
    procedure btnTouchWeightsBatchNameAge1Click(Sender: TObject);
    procedure btnTouchWeightsBatchNameAge2Click(Sender: TObject);
    procedure btnTouchWeightsGraphSaveClick(Sender: TObject);
    procedure btnTouchWeightsAddToGraphBatchClick(Sender: TObject);
    procedure btnTouchWeightsClearGraphBatchListClick(Sender: TObject);
    procedure btnTouchWeightsSaveBatchClick(Sender: TObject);
    procedure DBMemoTouchWeightsTargetCommentChange(Sender: TObject);
    procedure DBMemoTouchWeightsPrepCommentKeyDown(Sender: TObject;
      var Key: Word; Shift: TShiftState);
    procedure btnTouchPrepDateStartTodayClick(Sender: TObject);
    procedure btnTouchPrepDateEndTodayClick(Sender: TObject);
    procedure DBDateTimeTouchPrepStartKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBDateTimeTouchPrepEndKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnOpenOxcalClick(Sender: TObject);
    procedure btnSampleInfoShowAllSamplesOfProjectClick(Sender: TObject);
    procedure DBedtTouchWeightsBeforePrepChange(Sender: TObject);
    procedure DBedtTouchWeightsAfterPrepChange(Sender: TObject);
    procedure btnPrepCardClick(Sender: TObject);
    procedure ApplicationEventsException(Sender: TObject; E: Exception);
    procedure DBedtTouchWeightsMediumPrepKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBedtTouchWeightsMedium2PrepKeyDown(Sender: TObject;
      var Key: Word; Shift: TShiftState);
    procedure DBedtTouchWeightsMediumPrepChange(Sender: TObject);
    procedure DBedtTouchWeightsMedium2PrepChange(Sender: TObject);
    procedure DBEditWeightMediumChange(Sender: TObject);
    procedure DBEditWeightMedium2Change(Sender: TObject);
    procedure DBDateTimeTouchPrepStartChange(Sender: TObject);
    procedure btnReloadTransferMagazinesClick(Sender: TObject);
    procedure DBCheckBoxTouchPrepDiscardedClick(Sender: TObject);
    procedure DBCheckBoxTouchTargetDiscardedClick(Sender: TObject);
    procedure btnTouchWeightsBatchNameAge641Click(Sender: TObject);
    procedure btnTouchWeightsBatchNameAge642Click(Sender: TObject);
    procedure btnTouchWeightsBatchNameAutosamplerClick(Sender: TObject);
    procedure DBCheckBoxTouchMovedToCNClick(Sender: TObject);
    procedure DBGridSampleExchangeColEnter(Sender: TObject);
    procedure DBGridSampleExchangeColExit(Sender: TObject);
    procedure ToolButtonPressClick(Sender: TObject);
    procedure grdMagazinesUnpressedCellClick(Column: TColumn);
    procedure DBGridSamplesOfUnpressedMagazineDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure Button4Click(Sender: TObject);
    procedure btnMagazinePressedDateTodayClick(Sender: TObject);
    procedure memTargetCommentsChange(Sender: TObject);
    procedure dbedtsStorLocClick(Sender: TObject);
    procedure dbedtsStorLocChange(Sender: TObject);
    procedure dbedtsStorLocMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbedtsStorLocExit(Sender: TObject);
    procedure chkTouchSampleArchivedMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure chkTouchSampleArchivedMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure chkTouchPrepArchivedMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure chkTouchPrepArchivedMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure dbedtprepStorLocChange(Sender: TObject);
    procedure dbedtprepStorLocClick(Sender: TObject);
    procedure dbedtprepStorLocExit(Sender: TObject);
    procedure dbedtprepStorLocMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grdPretreatmentCellClick(Column: TColumn);
    procedure lbxDefinePrepStepsDblClick(Sender: TObject);
    procedure edtSampleInfoInDateClick(Sender: TObject);
    procedure edtSampleInfoOutDateClick(Sender: TObject);
    procedure edtSampleInfoDesiredDateClick(Sender: TObject);
    procedure edtPrepStartClick(Sender: TObject);
    procedure edtPrepEndClick(Sender: TObject);
    procedure edtGraphitizedClick(Sender: TObject);
    procedure edtTargetPressedClick(Sender: TObject);
    procedure edtGraphDateClick(Sender: TObject);
    procedure DBDateTimeTouchPrepStartClick(Sender: TObject);
    procedure DBDateTimeTouchPrepEndClick(Sender: TObject);
    procedure JvDBDateTimePicker2Click(Sender: TObject);
    procedure JvDBDateTimePicker1Click(Sender: TObject);
    procedure JvDBDateTimePicker3Click(Sender: TObject);
    procedure DBedtHECurrentChange(Sender: TObject);
    procedure edtSampleCPercentChange(Sender: TObject);

  private
    AcceptCol: integer; //for drag drop
    DragSource: TDragSource;
    DSN: string;
    dbClosed: boolean;
    SampleTrans: array of Boolean;
    CurrentFraction, CurrentMethod, CurrentMaterial, CurrentType, CurrentPrep, Prep: string;
    DragCol, DragRow: Longint;
    FFileName: string;
    FCheck, FNoCheck: TBitmap;
    GlbSampleNr, GlbPrepNr: integer;
    ProjectName: string;
    MANummer: integer;
    OrderNr: integer;
    ProjectNameReport: string;
    ProjectInDate: string;
    OneDateInMean : boolean;
    TransferPrepNr, TransferTargetNr : integer;
    SetCurrentToAll: boolean;
    ProjectExists, UserExists, UserNeedsUpdate, InvoiceExists, SameAddressForInvoice: boolean;
    FUseWordVersion: TWordVersion;
    FWord: TWordDriver;
    FDataset: TDataset;
    FValues: TStrings;
    Invoice: TInvoice;
    glbUserNr, glbInvoiceNr: integer;
    User: TUser;
    KTLsupervisor: boolean;
    WeightsChanged, TargetDataChanged: boolean;
    ServerRoot: string;
    GridOriginalOptions : TDBGridOptions;  // this retainsthe GridOptions of the ExchangeDBGrid
    fVideoImage: TVideoImage; // image object that holds the webcam data
    sStorLocSuggestion: string; // holds the suggestion of the s_storage_loc value
    prepStorLocSuggestion: string; // holds the suggestion of the prep_storage_loc value
    procedure ToggleCheckbox(acol, arow: integer);
    procedure ClearInsertSampleGrids;
    procedure CreateCds(German: boolean);
    procedure CreateWordExport;
    procedure DoSampleInfo(SampleNr: Integer; PrepNr: Integer; TargetNr: Integer);
    procedure ShowSampleInfoPage(Grid: TObject);
    procedure ExportReport(Ext: string);
    procedure FillCds;
    function GetFileName: string;
    procedure GetProjects;
    procedure GetSamples;
    procedure SetFileName(const Value: string);
    procedure SetValues(const Value: TStrings);
    procedure InsertNewSamplesInDb;
    procedure FillFractionBox;
    procedure FillMaterialBox;
    procedure FillPrepBox;
    procedure FillPrepList;
    procedure FillTypeBox;
    procedure GetListOfRecentMagazines(MonthsBack: integer);
    procedure GetListOfUnpressedMagazines(MonthsBack: integer);
    procedure GetSampleofUserProject;
    procedure ParseNewSampleFile(const FName: string);
    procedure SendMail(German: boolean);
    procedure SetupInsertSamplesWizard;
    procedure SetupLabTaskPage;
    procedure SetFormat;
    procedure SetupCommonPretreatment;
    procedure SetupLabPlanPage;
    procedure SetupMaterial;
    procedure SetupNewSamplesGrid;
    procedure ReportsSetupOptions;
    procedure DisplayProgramOptions;
    procedure SetSampleGridForMaterial;
    procedure SetSampleGridForPrep;
    procedure SetSampleGridForType;
    procedure SetSampleGridWidth;
    procedure SetTaskStartDate;
    procedure SetTaskEndDate;
    procedure SetupProjectPage;
    procedure SetupPretreatmentStepsGrid;
    procedure Status(const AMsg: string);
    procedure UpdateUser(const LastName: string);
    procedure GetPlanned;
    procedure CheckProjectExists;
    ///  <summary>get all samples of the selected user and project from the DB
    ///  and display in the report format</summary>
    ///  <param name="DoExportHalle"> When set to true it changes the output of the
    ///  table to a different layout so that it can be used for exporting data
    ///  to the Halle database </param>
    procedure GetSamplesForReport(DoExportHalle: boolean);
    /// <summary>Alternates the row colors in Listboxes</summary>
    /// <comments></comments>
    procedure AlternateRowColors(Sender: TObject; State: TGridDrawState);
    /// <summary>Calculates the yield from mass before and after prep out of the SampleInfo tab</summary>
    procedure CalculateYield;
    /// <summary>Calculates the yield from mass before and after prep out of the touch tab</summary>
    procedure CalculateYieldTouch;
    /// <summary>Calculates the net weight of weight_end of the pre table using weight_medium and weight_medium_2. Used in SampleInfoTab</summary>
    procedure CalculateNetWeightEnd;
        /// <summary>Calculates the net weight of weight_end of the pre table using weight_medium and weight_medium_2. used in TouchTab</summary>
    procedure CalculateNetWeightEndTouch;
    /// <summary>keeps the button of the toolbar pressed</summary>
    procedure ToolBarButtonsState(Sender: TObject);
    procedure DBGridAutoSizeColumn(Grid: TDBGrid; Column: integer);
    procedure DBGridAutoSizeAllColumns(Grid: TDBGrid);
    procedure FixDBGridColumnsWidth(const DBGrid: TDBGrid);
    procedure JumpToEmptyWeightField;
    procedure SetCheckBoxTouchSampleArchived;
    procedure SetCheckBoxTouchPrepArchived;
    procedure DisplayArchiveWarning;
    procedure DisplayArchiveWarningPrep;
    procedure CalculateCAbsWeight;

  public
    SampleModified: boolean;
    property FileName: string read GetFileName write SetFileName;
    property Dataset: TDataset read FDataset write FDataset;
    property Values: TStrings read FValues write SetValues;
  end;

var
  frmMAMS: TfrmMAMS;
  myPrepIni : TIniFile;
  //LogWindow: TLogWindow;

implementation

uses
  about, SHFolder, Clipbrd, CommCtrl, DateUtils, JvJCLUtils,
  _dm, ShlObj, ActiveX, ShellApi, _LogSQL, ComObj, Variants,
  StorageLocations, FormDBSearch;

{$R *.dfm}

const
  TypeCol = 10;
  MaterialCol = 8;
  FractionCol = 9;
  SampleNameCol = 1;
  SampleNrCol = 2;
  SampleDescr1Col = 3;
  SampleDescr2Col = 4;
  SampleWeightCol = 5;
  SampleUserMaterialCol = 6;
  SampleCommentCol = 7;
  SamplePrep1Col = 11;
  SamplePrep2Col = 12;
  SamplePrep3Col = 13;
  SamplePrep4Col = 14;
  SamplePrep5Col = 15;
  CSIDL_PROFILE = $0028;  //for retreiving the sample docs and fotos when stored locally


  NoSelection: TGridRect = (Left: - 1; Top: - 1; Right: - 1; Bottom: - 1);

type
  TGridCracker = class(TStringgrid);

function DefaultSaveLocation: string;
var
  P: PChar;
begin
  {
    returns the location of 'My Documents' if it exists, otherwise it returns
    the current directory.
  }
  P := nil;
  try
    P := AllocMem(MAX_PATH);
    if SHGetFolderPath(0, CSIDL_PERSONAL, 0, 0, P) = S_OK then
      Result := P
    else
      Result := GetCurrentDir;
  finally
    FreeMem(P);
  end;
end;

procedure TfrmMAMS.btnQuerySubmitterClick(Sender: TObject);
begin
  GetSamplesForReport(rgpExport.ItemIndex = 2);
end;

procedure TfrmMAMS.btnQueryToPretreatClick(Sender: TObject);
var
  s: string;
begin
  if dm.adoConnKTL.Connected then
  Begin
  with dm.qryDb do
  begin
    SQL.Text := 'SELECT sample_nr, user_t.last_name, user_label, user_label_nr, in_date, desired_date, weight FROM sample_t '
      + ' INNER JOIN user_t ON  (sample_t.user_nr = user_t.user_nr) WHERE av_fm IS NULL ORDER BY desired_date ;';
    DisableControls;
    s := SQL.Text;
//    ClipBoard.SetTextBuf(PChar(s));
    try
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      Open;
    finally
      EnableControls;
    end;
  end;
  End;
end;

procedure TfrmMAMS.tnstoClick(Sender: TObject);
// transfer analytical results from the target table to the sample table
// also transfer LE and HE currents from the db_dc14 database to the target_table!!!!
var
  i, j, m: integer;
  sample_nr, prep_nr, target_nr, le_current, he_current: string;
begin
  i := 0;

  // count how many entrys have prep=1 and target=1
  with dm.qryMagazineData do
  begin
    First;
    repeat //wenn tar=1 und prep=1
      if (FieldByName('target_nr').AsInteger = 1) and (FieldByName('prep_nr').AsInteger = 1) then
        i := i + 1;
      Next;
    until EOF;
  end;

  if MessageDlg(IntToStr(i) + ' sample ages are ready for transfer. Do you want to continue ?' + #13 + #10 + #13 + #10 + 'Click on the yellow colored rows to transfer those samples individually.', mtInformation, [mbOK, mbCancel], 0) = mrOk then begin
    StrGrdTargetData.Clear;
    edtMeanAge.Clear;
    edtSigmaAge.Clear;
    edtC13Mean.Clear;
    with dm.qryMagazineData do
    begin
      First;
      repeat //wenn tar=1 und prep=1
        if (FieldByName('target_nr').AsInteger = 1) and (FieldByName('prep_nr').AsInteger = 1) then begin
          dm.TransferAgeFromTarget(FieldByName('sample_nr').AsInteger,
            FieldByName('prep_nr').AsInteger,
            FieldByName('target_nr').AsInteger);
          j := RecNo;
          m := Length(SampleTrans);
          SampleTrans[RecNo] := true;
        end;
        Next;
      until EOF;
      grdMagazineData.Repaint;
    end;

{  if MessageDlg(IntToStr(dm.qryMagazineData.RecordCount) + ' sample ages are ready for transfer. Do you want to continue ?', mtInformation, [mbOK, mbCancel], 0) = mrOk then begin
    with dm.qryMagazineData do begin
      First;
      repeat
        dm.TransferAgeFromTarget(FieldByName('sample_nr').AsInteger,
          FieldByName('prep_nr').AsInteger,
          FieldByName('target_nr').AsInteger);
        Next;
      until EOF;
      ShowMessage('C14 results transfered for ' + IntToStr(RecordCount) + ' samples');
    end;
  end;  }

  // transfer LE and HE Currents to target table
    dm.cdsLeHeCurrents.First;
    repeat
      sample_nr := dm.cdsLeHeCurrents.FieldByName('sample_nr').AsString;
      prep_nr := dm.cdsLeHeCurrents.FieldByName('prep_nr').AsString;
      target_nr := dm.cdsLeHeCurrents.FieldByName('target_nr').AsString;
      le_current := dm.cdsLeHeCurrents.FieldByName('LECurrent').AsString;
      le_current := ReplaceStr(le_current, ',', '.');
      he_current := dm.cdsLeHeCurrents.FieldByName('HECurrent').AsString;
      he_current := ReplaceStr(he_current, ',', '.');

      LogWindow.addLogEntry('DB -- transfer LE and HE Currents to target_t: ' + sample_nr+'.'+prep_nr+'.'+target_nr);
      JvLogFile.Add('DB',lesInformation,'transfer LE and HE Currents to target_t: ' + sample_nr+'.'+prep_nr+'.'+target_nr);

      if le_current.IsEmpty then le_current := 'NULL';
      if he_current.IsEmpty then he_current := 'NULL';

      dm.TransferLeHeCurrentToTarget(sample_nr, prep_nr, target_nr, le_current, he_current);
      dm.cdsLeHeCurrents.Next;

    until dm.cdsLeHeCurrents.Eof;
    grdMagazineData.Repaint;
    end;
end;

procedure TfrmMAMS.btnTransferAgeToSampleClick(Sender: TObject);
begin
  dm.TransferAgeFromTarget(Trunc(edtSampleNr.Value), Trunc(edtSamplePrepNr.Value),
    Trunc(edtSampleTargetNr.Value));
  ShowMessage('Age transfered from target to sample');
end;

procedure TfrmMAMS.btnRefreshProjectsSinceYearClick(Sender: TObject);
begin
  dm.GetProjectsSinceYear(round(edtProjectsSinceYear.Value));
end;

procedure TfrmMAMS.btnRefreshSamplesAvailableClick(Sender: TObject);
//refresh table with uptodate sample information
var
  method,material: string;
begin
// This was used when filtered by Method
//  if rgpTask.ItemIndex = 0 then
//  begin
//    method := cmbPretreatMethod.ListSource.DataSet.FieldByName('method').AsString;
//    if length(method) > 0 then
//    begin
//      dm.GetSamplesAvailableForPrep(method);
//      SetupLabPlanPage; // colwidth
//    end;
//  end
//  else
//  begin // get all samples which are pretreated, but with graph date is NULL
//    dm.GetTargetsAvailable;
//    SetupLabPlanPage; // colwidth
////      JvDBStatusLabel3.DataSource := dm.dsTargetsAvailable;
//  end;

  //clear list first
  //dm.dsSamplesAvailable.DataSet.close;

  if rgpTask.ItemIndex = 0 then
  // preparation batch
  begin
    material := cmbFilterSampleMaterial.ListSource.DataSet.FieldByName('material').AsString;
    method := cmbPretreatMethod.ListSource.DataSet.FieldByName('method').AsString;
    //if length(material) > 0 then
    //begin
    //  dm.GetSamplesAvailableForPrepByMaterial(material);
    //  SetupLabPlanPage; // adjust colwidth
    //end;
    if length(method) > 0 then
    begin
      dm.GetSamplesAvailableForPrepByMethod(method);
      SetupLabPlanPage; // adjust colwidth
    end;
  end
  else
  // graphitisation batch
  begin // get all samples which are pretreated, but with graph date is NULL
    dm.GetTargetsAvailable;
    SetupLabPlanPage; // adjust colwidth
  end;

end;

procedure TfrmMAMS.btnReloadTransferMagazinesClick(Sender: TObject);
begin
GetListOfRecentMagazines(3);
end;

procedure TfrmMAMS.btnReportClick(Sender: TObject);
var
  s: string;

  // SubRoutine that sets the database flag out_date to the current date
  // ####################################################################
  procedure InsertReport;
  var
    f: string;
  begin
  // this is the old way of doing it using worddriver
    if not chkPreliminaryReport.Checked and not chkAllProjects.Checked then
    begin // insert report name in project_t
    if dm.adoConnKTL.Connected then
    Begin
      with dm.adoCmd do
      begin
        //f := ExtractFileName(edtSaveReportAs.Text);
        //f := ReplaceString(f, #34, '');
        f:=edtReportFileName.Text;
        LogWindow.addLogEntry('Report Filename saved in DB... '+ f);
        CommandText := 'UPDATE project_t SET report=' + #34 + f + #34 + ', out_date=' +
          #34 + FormatDateTime('YYYY-MM-DD', Date) + #34 +
          ' WHERE project_nr=' + IntToStr(cmbProjectOfReport.KeyValue) + ';';
        s := CommandText;
        LogWindow.addLogEntry(s);
        Execute;
      end;
    end;
    end;
  end;
  // End of SubRoutine that sets the database flag out_date to the current date
  // ####################################################################


begin
// old way of doing it using worddriver
  CreateCds(true); //create in-memory-DB (Datasource: cdsExport) with the table headers from the current ReportGrid
  FillCds;  // enter the data óf the ReportGrid into the in-memory-DB (Datasource: cdsExport), do calculations etc etc

// ###### part 1 get and fill in user info
// query database for information such as last name etc etc
    with dm.qryDb do // get user_info from user_t
    begin
      Close;
      SQL.Text := 'SELECT first_name, last_name, organisation, institute, address_1, address_2,' +
        ' town, postcode, country, email FROM user_t ' +
        'WHERE user_nr=' + IntToStr(cmbSubmNameReport.KeyValue) + ';';
      s := SQL.Text;
  //    ClipBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
    Begin
      Try
        Open;
      Except
        ShowMessage('problem opening the database');
        JvLogFile.Add('DB',lesError,'Problem opening DB');
      End;
    End;
    end;


  // main routine for export to word from here on
  CreateWordExport;  // instantiate an FWord Object using Worddriver

  // fill in the information such as last name etc into the word file
  with FWord do
  begin
    //SaveFileName := edtSaveReportAs.FileName;
    SaveFileName := TPath.Combine(edtSaveReporttoFolder.Text,edtReportFileName.Text);
    LogWindow.addLogEntry('File name for Report = ' + SaveFileName);
//    FileName := 'C:\Temp\SAMS\C14AgeTableTemplate.doc';
    FileName := edtWordTemplate.Filename;    //used to be edtWordTemplate.Text
    Save := true;
    KeepWordOpen := true;
    OutPutDir := edtSaveReporttoFolder.Text;
    LogWindow.addLogEntry('output folder for report = ' + OutPutDir);
    DataSource := dsExport;
//    DataSource := dm.dsSampleOfSubmitter;
    WordVersion := wvDetect;
//    DocumentMode := dmNewTable;
    DocumentMode := dmFillTable;
    TableTag := 'Table';
//    for i := 0 to cdsExport.FieldCount - 1 do ColumnList.Add(cdsExport.FieldDefs[i].Name);
    FValues.Add('FirstName=' + dm.dsQryDb.DataSet.FieldByName('first_name').AsString);
    FValues.Add('LastName=' + dm.dsQryDb.DataSet.FieldByName('last_name').AsString);
    s := '';
    if Length(dm.dsQryDb.DataSet.FieldByName('organisation').AsString) > 0 then
      s := dm.dsQryDb.DataSet.FieldByName('organisation').AsString;
    if Length(dm.dsQryDb.DataSet.FieldByName('institute').AsString) > 0 then
      s := s + #$0D + dm.dsQryDb.DataSet.FieldByName('institute').AsString;
    FValues.Add('Organisation=' + s);
    s := '';
    if Length(dm.dsQryDb.DataSet.FieldByName('address_1').AsString) > 0 then
      s := dm.dsQryDb.DataSet.FieldByName('address_1').AsString;
    if Length(dm.dsQryDb.DataSet.FieldByName('address_2').AsString) > 0 then
      s := s + #$0D + dm.dsQryDb.DataSet.FieldByName('address_2').AsString;
    FValues.Add('Address1=' + s);
    FValues.Add('ZIP=' + dm.dsQryDb.DataSet.FieldByName('postcode').AsString);
    FValues.Add('City=' + dm.dsQryDb.DataSet.FieldByName('town').AsString);
    FValues.Add('email=' + dm.dsQryDb.DataSet.FieldByName('email').AsString);  // added 2020.02.13
    FValues.Add('ProjectName=' + ProjectNameReport);    // values come from another subroutine that pulls put the samples for each project
    FValues.Add('ProjectInDate=' + ProjectInDate);
    FValues.Add('OrderNr=' + IntToStr(OrderNr));
    FValues.Add('Country=' + dm.dsQryDb.DataSet.FieldByName('country').AsString);
    Values := FValues;
//    Values.SaveToFile('c:\temp\list.txt');
//    i := Values.Count;
    KeepDocumentsOpen := true;

    // now that all the parameters are set EXECUTE the word export
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;  //this executes the word export
          LogWindow.addLogEntry('Report generation -- report created');
        Except
          on E: Exception do
          begin
            ShowMessage('Problem creating report: ' + E.Message);
            LogWindow.addLogEntry('Report generation -- problem creating word file');
          End;
        End
      End;
    InsertReport; //set database flag out_date to current date
  end;

  btnSendMailEnglish.Enabled := true;
  btnSendMailGerman.Enabled := true;

  //also export the oxcal data right away just as one would push the export button
  btnExportClick(self);

  end;

procedure TfrmMAMS.btnSaveUserProfileClick(Sender: TObject);
begin
  IF dm.adoConnKTL.Connected THEN
  Begin
    dm.dsUserInfo.DataSet.Post;
  End;
end;

procedure TfrmMAMS.btnSendInPrepToExcelClick(Sender: TObject);
begin
  with ExcelExport do
  begin
    Grid := grdInPrep;
    ExportGrid;
  end;
end;

const olMailItem = 0;

procedure TfrmMAMS.btnSendMailClick(Sender: TObject);
var
  MailItem: OLEVariant;
  CCString: String;
  CCSplitted: TArray<String>;
  RecipientString: String;
  RecipientSplitted: TArray<String>;
  i: integer;

begin
  with mesgMessage do
  begin
    Clear;
    // get from-address
    From.Text := trim(edtMailFrom.Text);

    // CCList
     CCString:= trim(edtMailCC.Text);
     CCSplitted:= CCString.Split([';']);
     for i := 0 to length(CCSplitted)-1 do
      begin
        CCList.Add.Text :=  trim(CCSplitted[i]);
        LogWindow.addLogEntry('setting email CC to: ' + trim(CCSplitted[i]));
      end;
      LogWindow.addLogEntry('email CC are: ' + CCList.EMailAddresses);

    // send email to from-address also? Add to ccList
    if chkSendCopyToSender.Checked then
    begin
      CCList.Add.Text := trim(edtMailFrom.Text);
    end;

    // set recipient(s)
    RecipientString:= trim(edtMailTo.Text); // trim TO string
    RecipientSplitted:= RecipientString.Split([';']); // split TO string into recipients using ; as deliminter
    //LogWindow.addLogEntry('lenght of array email recipient splitted: ' + length(RecipientSplitted).ToString);
    for i := 0 to length(RecipientSplitted)-1 do
      begin
        Recipients.Add.Text :=  trim(RecipientSplitted[i]);
        LogWindow.addLogEntry('setting email recipient to: ' + trim(RecipientSplitted[i]));
      end;
      LogWindow.addLogEntry('email recipient are: ' + Recipients.EMailAddresses);

    // set message text
    Subject := trim(edtMailSubject.Text);
    Body.Assign(MailMemo.Lines);
  end;
  with smtpSendMail do
  begin
    Host := trim(edtSMTPServer.Text);
    Username :=trim(edtMailFrom.Text);
    Connect;
    try
      Send(mesgMessage);
    except on E: Exception do
        Status('ERROR: ' + E.Message);
    end;
    Disconnect;
  end;
end;

procedure TfrmMAMS.btnSendMailEnglishClick(Sender: TObject);
begin
  SendMail(false);
end;

procedure TfrmMAMS.btnSendMailGermanClick(Sender: TObject);
begin
  SendMail(true);
end;

procedure TfrmMAMS.btnSendPlannedToExcelClick(Sender: TObject);
begin
  with ExcelExport do begin
    Grid := grdPlanned;
    ExportGrid;
  end;
end;

procedure TfrmMAMS.btnSendWaitForGraphToExcelClick(Sender: TObject);
begin
  with ExcelExport do begin
    Grid := grdWaitingForGraph;
    ExportGrid;
  end;
end;

procedure TfrmMAMS.btnSetDiscardClick(Sender: TObject);
begin
  with dm.adoCmd do
  begin
    CommandText := 'UPDATE target_t SET  ';
  end;
end;

procedure TfrmMAMS.btnSetGraphDateClick(Sender: TObject);
begin
  dm.SetGraphitized;
  ShowMessage('Done');
end;

procedure TfrmMAMS.btnSetSampleReadyForGraphClick(Sender: TObject);
var
  sample_nr, prep_nr: integer;
  s: string;
  d, h: string;
begin
  prep_nr := 1;
  d := FormatDateTime('YYYY-MM-DD', DateOf(Date));
  h := FormatDateTime('hh:mm', Now);
  with dm.qrySamplesOfLabTask do
  begin
    First;
    while not EOF do
    begin
      sample_nr := FieldbyName('sample_nr').AsInteger;
      s := 'UPDATE preparation_t SET prep_end=' + #34 + d + ' ' + h + #34 +
        '  WHERE sample_nr=' + IntTostr(sample_nr) + ' AND prep_nr=' + IntTostr(prep_nr) + ';';
//      ClipBoard.SetTextBuf(PChar(s));
      dm.adoCmd.CommandText := s;
      LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        dm.adoCmd.Execute;
      End;
      Next;
    end;
    ShowMessage('End of preparation set');
  end;
end;

procedure TfrmMAMS.btnSetSamplingDateTo1950Click(Sender: TObject);
VAR s: STring;
begin

  if dm.adoConnKTL.Connected then
    Begin
    Screen.Cursor:=CrHourGlass;
    LogWindow.addLogEntry('Updating sampling date to 1950-01-02 for all samples of project ' + DBEditProjectNr.Text);
    s := 'UPDATE sample_t SET ' +
      'sampling_date="1950-01-01" '+
      ' WHERE project_nr=' + DBEditProjectNr.Text + ';';
    //   ClibBoard.SetTextBuf(PChar(s));
    dm.adoCmd.CommandText := s;
    LogWindow.addLogEntry(s);
    try
      dm.adoCmd.Execute;
      LogWindow.addLogEntry('Update performed.');
      Screen.Cursor:=CrDefault;
    Except
      LogWindow.addLogEntry('Update NOT performed.');
      Screen.Cursor:=CrDefault;
    end;
    end;

end;

procedure TfrmMAMS.btnStartTaskClick(Sender: TObject);
begin
  SetTaskStartDate;
end;

procedure TfrmMAMS.btnSwapResultsClick(Sender: TObject);
begin
  dm.SwapResults(StrToInt(edtSample1.Text), StrToInt(edtSample2.Text));
end;

procedure TfrmMAMS.btnTransferC14AgeClick(Sender: TObject);
begin
  with dm.qryDB do
  begin
    SQL.Text := 'SELECT sample_t.sample_nr, target_t.prep_nr, target_t.target_nr FROM sample_t ' +
      ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr ' +
      ' WHERE sample_t.c14_age IS NULL AND target_t.c14_age IS NOT NULL;';
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
    Begin
      Try
        Open;
      Except
        ShowMessage('problem opening the database');
        JvLogFile.Add('DB',lesError,'Problem opening DB');
      End;
    End;
    if RecordCount > 0 then
    begin
      First;
      repeat
        dm.TransferAgeFromTarget(dm.qryDB.FieldByName('sample_nr').AsInteger,
          dm.qryDB.FieldByName('prep_nr').AsInteger,
          dm.qryDB.FieldByName('target_nr').AsInteger);
        Next;
      until EOF;
    end;
    ShowMessage('C14 results transfered for ' + IntToStr(RecordCount) + ' samples');
  end;
end;

procedure TfrmMAMS.btnTransferClick(Sender: TObject);
begin
  pgtMain.ActivePage := tbsTransferAge;
  GetListOfRecentMagazines(1);

  // select the first entry already
  dm.qryMagazines.First;

  //adjust column width to match table width since there is only one column
  grdMagazines.Columns[0].Width:=grdMagazines.Width;

  // call the targets of the first selected magazin which us usually the one that needs to be transfered
  grdMagazinesCellClick(grdMagazines.Columns.Items[0]);

  end;


procedure TfrmMAMS.Button3Click(Sender: TObject);
var
  s: string;
  nr: integer;
begin
  with dm.qryDB do
  begin
    SQL.Text := 'SELECT user_label_nr, sample_nr FROM sample_t WHERE user_label_nr LIKE ' + #34 + 'MA-%' + #34 + ';';
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
    Begin
      Try
        Open;
      Except
        ShowMessage('problem opening the database');
        LogWindow.addLogEntry('not opened');
        JvLogFile.Add('DB',lesError,'Problem opening DB');
      End;
    End;
    First;
    while not EOF do
    begin
      s := Fields.Fields[0].AsString;
      s := ExtractWord(2, s, ['-']);
      if WordCount(s, [' ']) > 0 then s := ExtractWord(1, s, [' ']);
      nr := StrToInt(s);
      with dm.adoCmd do
      begin
        CommandText := 'UPDATE sample_t SET MA_nr=' + IntToStr(nr) + ' WHERE sample_nr=' +
          IntToStr(Fields.Fields[1].AsInteger) + ';';
        s := CommandText;
//        ClipBoard.SetTextBuf(PChar(s));
        LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      end;
      Next;
    end;
  end;
end;

procedure TfrmMAMS.Button4Click(Sender: TObject);
var
  s: string;
begin
  if dm.adoConnKTL.Connected then
    Begin
    Screen.Cursor:=CrHourGlass;
    LogWindow.addLogEntry('Updating pressed date date to today for all samples/targets of project ' + DBEditProjectNr.Text);
    s := 'UPDATE target_t ' +
          ' INNER JOIN sample_t ON target_t.sample_nr = sample_t.sample_nr ' +
          ' INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr ' +
          ' SET target_t.target_pressed = curdate() ' +
         ' WHERE project_t.project_nr=' + DBEditProjectNr.Text + ';';
    //   ClibBoard.SetTextBuf(PChar(s));
    dm.adoCmd.CommandText := s;
    LogWindow.addLogEntry(s);
      try
        dm.adoCmd.Execute;
        LogWindow.addLogEntry('Update performed.');
        Screen.Cursor:=CrDefault;
        ShowMessage('target_pressed for all samples set to today');
      Except
        LogWindow.addLogEntry('Update NOT performed.');
        Screen.Cursor:=CrDefault;
        ShowMessage('Problem updating date');
      end;
    end;
end;

procedure TfrmMAMS.Button5Click(Sender: TObject);
begin
  LoadDialog.InitialDir := '%USERPROFILE%\SAMS\SAMS BILDER\';
  if LoadDialog.Execute then begin

  end;
end;

procedure TfrmMAMS.btnNowStartClick(Sender: TObject);
begin
  dtStartTaskDate.Date := Date;
  dtStartTaskTime.Time := TimeOf(Now);
  SetTaskStartDate;
end;

procedure TfrmMAMS.btnOpenOxcalClick(Sender: TObject);
Var
  url: string;
begin
 url := 'http://192.168.123.30:1950';
 ShellExecute(HInstance, 'open', PChar(url), nil, nil, SW_NORMAL);
end;

procedure TfrmMAMS.btnPendingReportsClick(Sender: TObject);
// displays all reports that don't have an 'out_date'
// internal projects are not shown
// the result set is shown in grdPendingReports
var
  proj_number: string; // variable that stores the project number for further queries
begin
  with dm.qryPendingReports do
  begin
    (*SQL.Text := 'SELECT project, last_name, first_name, in_date, project_nr FROM project_t ' +
      'INNER JOIN user_t on project_t.user_nr=user_t.user_nr ' +
      'WHERE in_date > ' + #34 + FormatDateTime('YYYY-MM-DD', DateOf(Date - 200)) + #34 + // 200 Tage zurück
      ' AND (out_date IS NULL or out_date < "2010-01-01") AND last_name<>"intern" ORDER BY in_date;';    *)

    SQL.Text := 'SELECT project, last_name, first_name, in_date, desired_date, project_t.project_nr, user_t.user_nr, help2_t.Samples, help3_t.discPrep, help5_t.prepDone, help4_t.discTarget, help6_t.graphDone, help7_t.inMagazine, help_t.Measured FROM project_t ' +
      'INNER JOIN user_t on project_t.user_nr=user_t.user_nr ' +
      //create and join to a table to count the measured samples in all project
      'LEFT JOIN (select project_nr, count(sample_nr) AS Measured FROM sample_t WHERE NOT ISNULL(c14_age) GROUP BY project_nr) help_t ON project_t.project_nr=help_t.project_nr ' +
      //create and join to a table to count the number of received samples in all project
      'LEFT JOIN (select project_nr, count(sample_nr) AS Samples FROM sample_t WHERE not_tobedated=0 GROUP BY project_nr) help2_t ON project_t.project_nr=help2_t.project_nr ' +
      //create and join to a table to count the dicarded preps in all project
      'LEFT JOIN (select sample_t.project_nr, count(preparation_t.sample_nr) AS discPrep FROM preparation_t INNER JOIN sample_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
      'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr WHERE preparation_t.stop = "1" GROUP BY sample_t.project_nr) help3_t ON project_t.project_nr=help3_t.project_nr ' +
      //create and join to a table to count the dicarded preps in all project
      'LEFT JOIN (select sample_t.project_nr, count(target_t.sample_nr) AS discTarget FROM target_t INNER JOIN sample_t ON target_t.sample_nr=sample_t.sample_nr ' +
      'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr WHERE target_t.stop = "1" GROUP BY sample_t.project_nr) help4_t ON project_t.project_nr=help4_t.project_nr ' +
      //create and join a table to count all samples that already have a prep_end
      'LEFT JOIN (select sample_t.project_nr, count(preparation_t.sample_nr) AS prepDone FROM preparation_t INNER JOIN sample_t ON preparation_t.sample_nr = sample_t.sample_nr ' +
      'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr WHERE NOT ISNULL(prep_end) GROUP BY sample_t.project_nr) help5_t ON project_t.project_nr=help5_t.project_nr ' +
      //create and join a table to count all samples that already have a graphitized
      'LEFT JOIN (select sample_t.project_nr, count(target_t.sample_nr) AS graphDone FROM target_t INNER JOIN sample_t ON target_t.sample_nr = sample_t.sample_nr ' +
      'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr WHERE NOT ISNULL(graphitized) GROUP BY sample_t.project_nr) help6_t ON project_t.project_nr=help6_t.project_nr ' +
      // create and join a table that counts all samples that have a magazine assigned
      'LEFT JOIN (select sample_t.project_nr, count(target_t.sample_nr) AS inMagazine FROM target_t INNER JOIN sample_t ON target_t.sample_nr = sample_t.sample_nr ' +
      'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr WHERE NOT ISNULL(magazine) GROUP BY sample_t.project_nr) help7_t ON project_t.project_nr=help7_t.project_nr ' +
      // only list projects within a certain time frame and without out_date
      'WHERE in_date > ' + #34 + FormatDateTime('YYYY-MM-DD', DateOf(Date - 300)) + #34 + // 300 Tage zurück
      ' AND (out_date IS NULL or out_date < "2010-01-01") AND NOT (last_name="intern" AND first_name="intern") ORDER BY desired_date;';

    //LogWindow.addLogEntry(SQL.Text);
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports...');
    // JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    Application.ProcessMessages;
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('DB -- query executed');
        Except
          //ShowMessage('problem opening the database');
          // connections is closed, reconnect by setting connected to true and test again
            LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                       JvLogFile.Add('DB',lesError,'can not reconnect to DB');
                    End;
            Open;
        End;
      End;
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports... finished');
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports... ' + grdPendingReports.Columns.Count.ToString + ' columns to display in grid');
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports... setting column widths manually');
    LogWindow.addLogEntry('btnPendingReportsClick: DBGrid setting columns widths manually');

    // adjust column sizes of the data table
    with grdPendingReports do
      begin
        Columns[0].Width := 130; // project
        Columns[1].Width := 100; // last name
        Columns[2].Width := 90; // first name
        Columns[3].Width := 70;  // in date
        Columns[4].Width := 80;  // desired date
        Columns[5].Width := 60;  // project nr
        Columns[6].Width := 60;  // user nr
        Columns[7].Width := 60;  // samples
        Columns[8].Width := 60;  // discarded prep
        Columns[9].Width := 60;  // prepDone
        Columns[10].Width := 60;  // discraded target
        Columns[11].Width := 60;  // graphDone
        Columns[12].Width := 60;  // inMgazine
        Columns[13].Width := 70;  // measured
      end;

    LogWindow.addLogEntry('btnPendingReportsClick: DBGrid setting columns widths automatically');
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports... setting column widths automatically');

    FixDBGridColumnsWidth(grdPendingReports);
    LogWindow.addLogEntry('btnPendingReportsClick: DBGrid setting columns widths automatically done!');
    JvLogFile.Add('Pending Reports',lesInformation,'Query DB for pending reports... setting column widths automatically done!');

  // set the column that is neing used to order the grid by
  grdPendingReportsTitleClick(grdPendingReports.Columns[4]);
  //LogWindow.addLogEntry('btnPendingReportsClick: DBGrid TitleClick set');
  end;

end;

procedure TfrmMAMS.btnPlotQueryClick(Sender: TObject);
// send query to DB and return data for the DBPlot
VAR
  i: Integer;
  xlabel, ylabel: String;
  xvalue, yvalue: Extended;
begin
  with dm.qryDBPlot do
  begin
  //'SELECT magazine, fm from target_t WHERE magazine like "%HD%" order by magazine';
    SQL.Text := MemoDBPlotQuery.Lines.Text;
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
(*
      with DBgrdDBPlot do begin
        Columns[0].Width := 160;
        Columns[1].Width := 120;
      end;
*)
  end;

  JvFormStorage1.SaveFormPlacement; // perform a save of all form values of SAMS including the query so it will be retrieved the next time automatically

  //add line numbers to dbgrid
  // add column
    //TColumn(DBGridDBPlot.Columns.Insert(0)).Title.Caption := 'OK';
    //TStringGrid(DBGridDBPlot).Cells[1,1] := 'test';


  //Plot Data in DBChart
    DBChart.ClearChart;                     //clear the chart
    DBChart.AddSeries(TpointSeries);        // add new data series TPoint Style
    DBChart.Series[0].Depth:=0;             // Format the display style
    DBChart.Series[0].HasZValues:= false;
    DBChart.LeftWall.Visible:=false;
    DBChart.BottomWall.Visible:=false;
    DBChart.DepthAxis.Visible:=false;
    DBChart.BottomAxis.Labels:=true;
    // DBChart.Series[0].Marks.Visible :=true;
    DBChart.Legend.Visible:=false;
    DBChart.View3D:=false;

    xlabel:=DBGridDBPlot.Columns.Items[0].Title.Caption;  //get name of the first column
    ylabel:=DBGridDBPlot.Columns.Items[1].Title.Caption;  //get name of the second column
    DBChart.BottomAxis.Title.Caption:=xlabel; //set the x-axis title to the name of the first column
    DBChart.LeftAxis.Title.Caption:=ylabel; //set the left y-axis title to the name of the first column
    DBChart.LeftAxis.AxisValuesFormat := '0.0000';

    //add datapoints from the DBGrid, go through all the data in the DBGrid and send to the plot
    DBGridDBPlot.DataSource.DataSet.First;   // go to the first datapoint in the dataset
    for i := 1 to DBGridDBPlot.DataSource.DataSet.RecordCount do begin
      if NOT TryStrToFloat(DBGridDBPlot.DataSource.DataSet.FieldByName(xlabel).AsString, xvalue) then xvalue:=i; //try to convert x-value to a number, if it doesn't work use index i as value
      yvalue:= DBGridDBPlot.DataSource.DataSet.FieldByName(ylabel).AsExtended; //get y value from dataset
      DBChart.Series[0].AddXY(xvalue,yvalue,inttostr(DBGridDBPlot.DataSource.DataSet.RecNo),clTeeColor);  //add xy to plot, recordNo as label
      DBGridDBPlot.DataSource.DataSet.Next;  // move to the next datapoint in the dataset
    end;
     FixDBGridColumnsWidth(DBGridDBPlot);

    //DBChart.Series[0].Marks.Visible:=true;   //display the Text on the data points (marks)
    //DBChart.Series[0].OnGetMarkText:=DBChartSeriesGetMarkText;  // generate the text on the datapoints (marks)

    //s.XValues:= dm.qryDBPlot.FieldByName('magazine').Text;
    //s.YValues := dm.qryDBPlot.FieldByName('fm').Text;
    //s.XValues.ValueSource := dm.qryDBPlot.FieldByName('magazine').Text;
    //s.YValues.ValueSource := dm.qryDBPlot.FieldByName('fm').Text;
end;

procedure TfrmMAMS.btnPrepCardClick(Sender: TObject);
VAR
  Mams, LastName, UserLabel, InDate, PrepStart, WeightStart: string;
begin
  // create a PrepCard for the selected sample (sample in sampleInfo)
  //===============================================================
  // substantiate a new FWord object using Worddriver AddOn
    if Assigned(FWord) then FWord.Free;
    if assigned(FValues) then FValues.Free;
    FWord := TWordDriver.Create(Self);
    FUseWordVersion := wvDetect;
    //LogWindow.addLogEntry('PrepCard WordVersion Detected = ' + wvDetect);
    FValues := TStringList.Create;

   // fill in the information into the word template such as MAMS, last name, sample name...
  with FWord do
  begin
    // set filename of template file
    if Trim(JvDirEdt_PrepCards_Path.Text) <> '' then
      begin
        FileName := JvDirEdt_PrepCards_Path.Text + '\PrepCardBone.docx'; //filename of word file to use as template
        LogWindow.addLogEntry('PrepCard -- WordTemplate= ' + FileName);
      end
      else
      begin
        LogWindow.addLogEntry('PrepCard -- Path to WordTemplate not set');
        Exit;
      end;

    //SaveFileName := TPath.Combine(edtSaveReporttoFolder.Text,edtReportFileName.Text);
    OutPutDir := JvDirEdt_PrepCards_Path.Text;
    LogWindow.addLogEntry('PrepCard -- output folder = ' + OutPutDir);
    SaveFileName := 'temp_PrepCard.doc';  //use a temporary file only
    LogWindow.addLogEntry('PrepCard -- FileName for saving = ' + SaveFileName);
    WordVersion := wvDetect;
    LogWindow.addLogEntry('PrepCard -- WordVersion found = ' + GetWordVersion);
    Save := true;  //save the file when done
    KeepWordOpen := true; //keep word open after everything is finished
    DataSource := dm.dsSampleInfo;
    DocumentMode := dmSerialLetter; //dmFillTable
    FWord.OnlyCurrentRecord := True; //only uses the current record of the datasource (because we create a serial letter

    // query database for current sample (re-use SampleInfo button)
    btnDoSampleQueryClick(sender);

    //create the string list with the values to be replaced
    Mams:=edtSampleNr.Text;
    LastName:= dm.dsSampleInfo.DataSet.FieldByName('last_name').AsString;
    UserLabel:= dm.dsSampleInfo.DataSet.FieldByName('user_label').AsString;
    InDate:=dm.dsSampleInfo.DataSet.FieldByName('in_date').AsString;
    PrepStart:=dm.dsSampleInfo.DataSet.FieldByName('prep_start').AsString;
    WeightStart:=dm.dsWeights.DataSet.FieldByName('weight_start').AsString;

    FValues.Add('MAMS=' + mams);
        LogWindow.addLogEntry('PrepCard -- setting replacement value MAMS = ' + Mams);
    FValues.Add('LastName=' + LastName);
        LogWindow.addLogEntry('PrepCard -- setting replacement values LastName = ' + LastName);
    FValues.Add('UserLabel=' + UserLabel);
        LogWindow.addLogEntry('PrepCard -- setting replacement values UserLabel = ' + UserLabel);
    FValues.Add('InDate=' + InDate);
        LogWindow.addLogEntry('PrepCard -- setting replacement values InDate = ' + InDate);
    FValues.Add('PrepStart=' + PrepStart);
        LogWindow.addLogEntry('PrepCard -- setting replacement values PrepStart = ' + PrepStart);
    FValues.Add('WeightStart=' + WeightStart);
        LogWindow.addLogEntry('PrepCard -- setting replacement values WeightStart = ' + WeightStart);
    Values := FValues;  // "Values" is a list of identifiers/value pairs that will be replaced
    //Values.SaveToFile('c:\temp\list.txt');
    KeepDocumentsOpen := true;

    // now that all the parameters are set, EXECUTE the word export
    Try
      Execute;  // this executes the export
      LogWindow.addLogEntry('PrepCard -- PrepCard generated sucessfully');
    Except
      on E: Exception do
      begin
        ShowMessage('Problem creating Wordfile: ' + E.Message);
        LogWindow.addLogEntry('PrepCard -- problem creating report card');
      end;
    End;
  end;


end;

procedure TfrmMAMS.btnPrintBatchClick(Sender: TObject);
VAR i: integer;
    s, sample_nr : string;
begin
 // gather information that need to be written to word
 // use the infor from the current batch and create an in-memory-datasource
 // first create the table

 //create in-memory-dataset
 LogWindow.addLogEntry('PrepBatch Wordfile -- generate in memory DB');
    cdsPrepBatch.Close;  //cdsPrepBatches is a ClientDataSet that is being build in-memory and can be used like a database
  with cdsPrepBatch.FieldDefs do    // create FieldDefinitions  aka Tcolumn names
  begin
    Clear;
    with AddFieldDef do  //add new field to the dataset
    begin
      Name := 'sample_nr';
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := 'user_last_name';
      DataType := ftString;
      Size := 40;
    end;
    with AddFieldDef do
    begin
      Name := 'user_label';
      DataType := ftString;
      Size := 40;
    end;
    with AddFieldDef do
    begin
      Name := 'weight_start';
      DataType := ftString;
      Size := 10;
    end;
  end;
  cdsPrepBatch.CreateDataSet;  // now make an actual dataset out of it
  cdsPrepBatch.Close;
  LogWindow.addLogEntry('PrepBatch Wordfile -- generate in memory DB done!');

  // the corresponding datasource is called dsPrepBatches
  dsPrepBatch.DataSet := cdsPrepBatch;

   LogWindow.addLogEntry('PrepBatch Wordfile -- query sample DB');
  //fill cdsPrepBatch with data from the currently created prep batch
  //take info from lbxBatch listbox and extract MAMS
  cdsPrepBatch.Open;
  for i := 0 to lbxBatch.Count - 1 do
    begin
      cdsPrepBatch.Edit;
      cdsPrepBatch.Append;
      sample_nr := Trim(ExtractWord(1, lbxBatch.Items[i], ['|']));
      s := 'SELECT sample_t.sample_nr, user_label, last_name, weight_start' +
           ' FROM sample_t ' +
           ' INNER JOIN project_t ON sample_t.project_nr=project_t.project_nr' +
           ' INNER JOIN user_t ON project_t.user_nr=user_t.user_nr' +
           ' INNER JOIN preparation_t ON sample_t.sample_nr=preparation_t.sample_nr' +
           ' WHERE sample_t.sample_nr = ' + sample_nr + ';';
      dm.qryDB.SQL.Text := s;
      LogWindow.addLogEntry('PrepBatch Wordfile -- query: ' + s);
          IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('PrepBatch Wordfile -- query executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      cdsPrepBatch.Fields.Fields[0].Value := dm.qryDB.FieldByName('sample_nr').AsString;  //Sample_nr
      cdsPrepBatch.Fields.Fields[1].Value := dm.qryDB.FieldByName('last_name').AsString;  //last_name
      cdsPrepBatch.Fields.Fields[2].Value := dm.qryDB.FieldByName('user_label').AsString;  //user_label
      cdsPrepBatch.Fields.Fields[3].Value := dm.qryDB.FieldByName('weight_start').AsString;  //weight_start
      dm.qryDB.Close;
    end;
    cdsPrepBatch.Post;
    cdsPrepBatch.Close;

    // connect to WordApplication
    CreateWordExport;

 // fill in the information into the word file
  cdsPrepBatch.Open;
  with FWord do
  begin
    if Trim(edtFilenamePrepDocTemplate.Text) <> '' then
      begin
        OutPutDir:= JvDirEdt_PrepBatch_Path.Text;
        LogWindow.addLogEntry('PrepBatch Wordfile -- OutputDir = ' + OutPutDir);
        FileName := edtFilenamePrepDocTemplate.Filename;  //template
        LogWindow.addLogEntry('PrepBatch Wordfile -- Template Filename = ' + FileName);
        SaveFileName := TPath.Combine(OutPutDir,edtBatchName.Text+'.doc');
        LogWindow.addLogEntry('PrepBatch Wordfile -- target filename = ' + SaveFileName);
      end
      else
      begin
        LogWindow.addLogEntry('PrepBatch Wordfile -- no template selected');
        Exit;    // no template file is selected exit here
      end;
    Save := true;
    KeepWordOpen := true;
    DataSource := dsPrepBatch;
//    DataSource := dm.dsSampleOfSubmitter;
    WordVersion := wvDetect;
    LogWindow.addLogEntry('PrepBatch Wordfile -- WordVersion found = ' + GetWordVersion);
    DocumentMode := dmFillTable;
    TableTag := 'Table';
    //replace Text in template (this is not for the table)
    FValues.Add('batch_name=' + edtBatchName.Text);
    Values := FValues;
    KeepDocumentsOpen := true;

    // now that all the parameters are set, EXECUTE the word export
    Try
      Execute;  //this executes the export
      LogWindow.addLogEntry('PrepBatch Wordfile -- Wordfile sucessfully');
    Except
      on E: Exception do
      begin
        ShowMessage('Problem creating Wordfile: ' + E.Message);
        LogWindow.addLogEntry('PrepBatch Wordfile --  problem creating word file');
      end;
    End;
  end;


end;

procedure TfrmMAMS.btnProjectLettersClick(Sender: TObject);
// display the documents that where send by the user
Var
  s, fname: string;
  Rec: TSearchRec;
begin
    s := IntToStr(dm.dsProjectsOfUser.DataSet.FieldByName('project_nr').AsInteger) + '?.pdf';
    fname := ServerRoot + 'SAMS Documents\';
    Statusbar.Panels[2].Text := fname;
    if FindFirst(fname + s, faAnyFile - faDirectory, Rec) = 0 then
    try
      repeat
        ShellExecute(Handle, 'open', PChar(fname + Rec.Name), nil, nil, sw_show);
      until FindNext(Rec) <> 0;
    finally
      FindClose(Rec);
//     if FileExists(s) then ShellExecute(Handle, 'open',PChar(s),nil,nil, sw_show)
    end;
end;

procedure TfrmMAMS.DBChartClickSeries(Sender: TCustomChart;
  Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
// when clicking on a datapoint the x and y data of the
// datapoint is displayed next to the cursor
VAR i: integer;
begin
  with (Sender AS TChart) do
  begin
    Repaint;
    Canvas.Brush.Style := bsSolid;
    Canvas.Pen.Color := clBlack;
    Canvas.Brush.Color := clWhite;
    //Canvas.TextOut(X+10,Y,Sender.Series[0].GetHorizAxis.LabelValue(ValueIndex)+','+Sender.Series[0].GetVertAxis.LabelValue(ValueIndex));
    Canvas.TextOut(X+10,Y,Sender.Series[0].XValue[ValueIndex].ToString()+', '+Sender.Series[0].YValue[ValueIndex].ToString());

    DBGridDBPlot.SetFocus;
    //locate the clicked data point in the table by locating the x and y values
    if DBGridDBPlot.DataSource.DataSet.Locate(DBGridDBPlot.Columns.Items[0].FieldName+';'+DBGridDBPlot.Columns.Items[1].FieldName,VarArrayOf([Sender.Series[0].XValue[ValueIndex],Sender.Series[0].YValue[ValueIndex]]),[loCaseInsensitive]) = false Then
      begin
        //ShowMessage('TableIndex =' + Sender.Series[0].Labels[ValueIndex] + #13 + DBGridDBPlot.Columns.Items[0].FieldName + ' = ' + Sender.Series[0].XValue[ValueIndex].ToString + #13 + DBGridDBPlot.Columns.Items[1].FieldName + ' = '+ Sender.Series[0].YValue[ValueIndex].ToString);
        DBGridDBPlot.SelectedRows.CurrentRowSelected := true;
        // the label of the selected datapoint holds the row-number inside the DBGrid
        DBGridDBPlot.DataSource.DataSet.RecNo := Sender.Series[0].Labels[ValueIndex].ToInteger;
      end;

    //DBGridDBPlot.SetFocus;
    //DBGridDBPlot.DataSource.DataSet.First;
    //DBGridDBPlot.DataSource.DataSet.Locate();
    //DBGridDBPlot.DataSource.DataSet.MoveBy(ValueIndex); // go to index of selected data point
    //DBGridDBPlot.SelectedRows.CurrentRowSelected:=true;

  end;
end;

procedure TfrmMAMS.DBChartSeriesGetMarkText(Sender:TChartSeries; ValueIndex:Integer; var MarkText:String);
// this procedure generates the MarksText, the text that
// is displayed on the datapoint
begin
  DBGridDBPlot.DataSource.DataSet.First;  //go to first record
  DBGridDBPlot.DataSource.DataSet.MoveBy(ValueIndex); //move to the reecord that is is equal to the datapoint
  MarkText:=DBGridDBPlot.Fields[0].Text;  //get value from database
  //if (ValueIndex<>Sender.Count-1) then MarkText:='';
end;


procedure TfrmMAMS.DBCheckBoxTouchMovedToCNClick(Sender: TObject);
begin
  btnTouchWeightsGraphNeedsSaving.Visible := True;

  SetCheckBoxTouchPrepArchived;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.DBCheckBoxTouchPrepDiscardedClick(Sender: TObject);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;
  DBDateTimeTouchPrepEnd.Date := Now;
end;

procedure TfrmMAMS.DBCheckBoxTouchTargetDiscardedClick(Sender: TObject);
begin
btnTouchWeightsGraphNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.dbchkFreeOfCharge1Click(Sender: TObject);
begin
 With (Sender As TDBCheckBox) Do
 Begin
    if checked then
    Begin
      Font.Style:=[fsBold];
    End
    else
    begin
     Font.Style:=[];
    end;
 End;
end;

procedure TfrmMAMS.dbchkFreeOfCharge2Click(Sender: TObject);
begin
 With (Sender As TDBCheckBox) Do
 Begin
    if checked then
    Begin
      Font.Style:=[fsBold];
    End
    else
    begin
     Font.Style:=[];
    end;
 End;
end;

procedure TfrmMAMS.btnNowEndClick(Sender: TObject);
begin
  dtEndTaskDate.Date := Date;
  dtEndTaskTime.Time := TimeOf(Now);
  SetTaskEndDate;
end;

procedure TfrmMAMS.Button6Click(Sender: TObject);
var
  SampNr, ProjectNr, ProjectNr100: integer;
begin
  if (MessageDlg('all project numbers will be changed ! Are you sure ? (Delete foreign key in sample_t)', mtWarning, [mbYes, mbNo], 0) = mrYes) then begin
    with dm.qryDB do
    begin
      SQL.Text := ' SELECT * FROM sample_t';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
          IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      First;
      while not EOF do
      begin
        SampNr := FieldByName('sample_nr').AsInteger;
        ProjectNr := FieldByName('project_nr').AsInteger;
        ProjectNr100 := ProjectNr + 100;
        dm.adoCmd.CommandText := 'UPDATE sample_t  SET project_nr=' + IntToStr(ProjectNr100) +
          ' WHERE sample_nr=' + IntToStr(SampNr) + ';';
                IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
        Next;
      end;
      Close;
    end;
    with dm.qryDB do
    begin
      SQL.Text := ' SELECT * FROM project_t';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
          IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      Last;
      while not BOF do
      begin
        ProjectNr := FieldByName('project_nr').AsInteger;
        ProjectNr100 := ProjectNr + 100;
        dm.adoCmd.CommandText := 'UPDATE project_t  SET project_nr=' + IntToStr(ProjectNr100) +
          ' WHERE project_nr=' + IntToStr(projectnr) + ';';
        LogWindow.addLogEntry(dm.adoCmd.CommandText);
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
        Prior;
      end;
      Close;
    end;
    showmessage('done');
  end;
end;

procedure TfrmMAMS.Button7Click(Sender: TObject);
var
  Month: integer;
  s: string;
begin
  with dm.adoCmd do
  begin
    dm.qryDb.SQL.Text := 'Select * FROM sample_t WHERE sample_nr=' + IntToStr(YearOf(Date) - 2000) + ';';
    LogWindow.addLogEntry(dm.qryDb.SQL.Text);
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    if dm.qryDb.RecordCount = 0 then
    begin
      CommandText := ' INSERT INTO sample_t (sample_nr, project_nr) VALUES(' +
        IntToStr(YearOf(Date) - 2000) + ',1);';
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      ShowMessage('Current year inserted in sample_t');
    end
    else
      ShowMessage('Current year already exists in sample_t');
    for Month := 1 to 12 do begin
      dm.qryDb.SQL.Text := 'Select * FROM preparation_t WHERE sample_nr='
        + IntToStr(YearOf(Date) - 2000) + ' AND prep_nr='
        + IntToStr(Month) + ';';
      LogWindow.addLogEntry(dm.qryDb.SQL.Text);
          IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      if dm.qryDb.RecordCount = 0 then begin
        CommandText := 'INSERT INTO preparation_t (sample_nr, prep_nr) VALUES(' +
          IntToStr(YearOf(Date) - 2000) + ',' + IntToStr(Month) + ');';
        Execute;
      end;
      dm.qryDb.SQL.Text := 'Select * FROM target_t WHERE sample_nr='
        + IntToStr(YearOf(Date) - 2000) + ' AND prep_nr='
        + IntToStr(Month) + ';';
      LogWindow.addLogEntry(dm.qryDb.SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDb.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      if dm.qryDb.RecordCount = 0 then
      begin
        CommandText := 'INSERT INTO target_t (sample_nr, prep_nr, target_nr) VALUES(' +
          IntToStr(YearOf(Date) - 2000) + ',' + IntToStr(Month) + ',1);';
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
        s := s + IntToStr(Month) + ',';
      end;
    end;
    ShowMessage('Months ' + s + ' inserted into preparation_t and target_t');
  end;
end;

procedure TfrmMAMS.Button8Click(Sender: TObject);
begin
  dm.TransferMA_Nr_To_MAMS(edtLimitTransferMA.Text);
end;


procedure TfrmMAMS.btn_initestClick(Sender: TObject);
var test: Tstrings;
begin
  myPrepIni.ReadSections(test);
  showmessage(test.GetText);
end;

procedure TfrmMAMS.btn_RemoveProjectClick(Sender: TObject);
// remove the selected project from the database
VAR
  project_nr, number_of_samples, result, buttonSelected: integer;
begin

  // get project number from the table
  project_nr := grdProjects.DataSource.DataSet.FieldByName('project_nr').AsInteger;
  // showmessage(grdProjects.DataSource.DataSet.FieldByName('project_nr').AsString);

  buttonSelected := messagedlg('Delete selected Project (' + inttostr(project_nr) + ')?',mtError, mbOKCancel, 0);
  if buttonSelected = mrOK then
    Begin
         // check whether or not samples still exist in that project
        number_of_samples := dm.GetSamplesByProjectNr(project_nr);
         //showmessage(inttostr(number_of_samples));

        // if number_of_samples is zero then remove project
        if number_of_samples = 0 then
          begin
              result := dm.RemoveProjectByProjectNr(project_nr);
              // showmessage(inttostr(result));
              cmbSubmitterNameProjectCloseUp(self);
          end
          else
          begin
            showmessage('there are still ' + inttostr(number_of_samples) + ' samples associated to the project. Delete samples first.');
          end;

    End;

end;

procedure TfrmMAMS.btn_report_pageClick(Sender: TObject);
begin
  pgtMain.ActivePage:=tbsUserReport; //go to report page

  cmbSubmNameReport.KeyValue:=cmbSubmitterNameProject.ListSource.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list to correct user_nr
  // ShowMessage(cmbSubmitterNameProject.ListSource.DataSet.FieldByName('user_nr').AsString);
  cmbSubmNameReportCloseUp(self); //correct user is selected now show their projects
  cmbProjectOfReport.ListSource.DataSet.Locate('project_nr',grdProjects.DataSource.DataSet.FieldByName('project_nr').AsString,[loPartialKey]);//now the projects are being displayed, select the correct project
  cmbProjectOfReport.KeyValue:=  grdProjects.DataSource.DataSet.FieldByName('project_nr').AsString;  //not sure why I have to do this
  cmbProjectOfReport.CloseUp(false);  //close the dropdown, which also triggers the OnCloseUp event

end;


procedure TfrmMAMS.btn_UserInfoClick(Sender: TObject);
// switch page to user info and display info of current user
begin
  pgtMain.ActivePage:=tbsUserInfo; //go to report page
  cmbUsernameUserInfo.KeyValue:=cmbSubmitterNameProject.ListSource.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list to correct user_nr
end;

procedure TfrmMAMS.btnTransferTargetDataClick(Sender: TObject);
var
  snr: string;
  s: string;
  c: char;
  c14age: string;
begin
  snr := StrGrdTargetData.Cells[1, 1];
  if MessageDlg('Sample ' + snr + ' ready for transfer. Do you want to continue ?' + #13 + #10 + #13 + #10 + 'Calculated data will be used.', mtInformation, [mbOK, mbCancel], 0) = mrOk then begin
    with dm.adoCmd do
    begin
      c14age := edtMeanAge.Text;
      if Length(c14age) = 0 then c14age := '99999';
      CommandText := 'UPDATE sample_t' +
        ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr' +
        ' SET sample_t.c14_age=' + c14age + ',' +
        '  sample_t.c14_age_sig=' + edtSigmaAge.Text + ',' +
        '  sample_t.av_dc13=' + ReplaceStr(edtC13Mean.Text, ',', '.') +
        ' WHERE sample_t.sample_nr=' + StrGrdTargetData.Cells[0, 1] +
        ';';
      s := CommandText;
//    ClipBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      if OneDateInMean then
          dm.TransferAgeFromTarget(StrToInt(StrGrdTargetData.Cells[0, 1]),
             TransferPrepNr, TransferTargetNr );
    end;
  end;
  btnTransferTargetData.enabled := false;
end;

procedure TfrmMAMS.btnUpdateNoOxBlanksClick(Sender: TObject);
begin
  GetAvailableOxasBlanks;
end;

procedure TfrmMAMS.btnUserExportClipboardClick(Sender: TObject);
//export user data to clipboard
Var
  s: string;
begin
  s :=  DBEdit12.Text + ' ' + cmbUsernameUserInfo.Text + #13+#10 +
        DBEdit17.Text + #13+#10 +
        DBEdit18.Text + #13+#10 +
        DBEdit19.Text + #13+#10 +
        DBEdit20.Text + #13+#10 +
        DBEdit21.Text + ' ' + DBEdit22.Text + #13+#10;
  Clipboard.AsText := s;
end;

procedure TfrmMAMS.btnSampleInfoShowAllSamplesOfProjectClick(Sender: TObject);
Var
    Column: TColumn; //this is used because the 'OnClick' method needs this parameter
begin
// switch to tab "User Projects" and display all samples that belong that Project

  // switch tab to user projects
  pgtMain.ActivePage:=tbsProjectsOfUser;
  // get project number and select the correct project
  cmbSubmitterNameProject.KeyValue:=dm.dsSampleInfo.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list to correct user_nr
  cmbSubmitterNameProjectCloseUp(self); //correct user is selected, now show their projects
  grdProjects.DataSource.DataSet.Locate('project_nr',dm.dsSampleInfo.DataSet.FieldByName('project_nr').AsInteger,[loPartialKey]);//now the projects are being displayed, select the correct project
  grdProjectsCellClick(Column); //now display the samples that belong to the project
  grdSamplesOfProject.DataSource.DataSet.Locate('sample_nr',dm.dsSampleInfo.DataSet.FieldByName('sample_nr').AsInteger, [loPartialKey]); // select the correct sample_nr
  grdSamplesOfProject.OnCellClick(Column);

end;

procedure TfrmMAMS.btnSampleNrUpDownClick(Sender: TObject; Button: TUDBtnType);
Var SampleNr, NPreps, NTargets: Integer;
begin
  // change sample nr according to the action either one up or one down
  case Button of
    btNext:
    begin
      SampleNr := round(edtSampleNr.Value) + 1;
      edtSampleNr.Value := SampleNr;
    end;
    btPrev:
    Begin
      SampleNr := round(edtSampleNr.Value) - 1;
      edtSampleNr.Value := SampleNr;
    End;
  end;
  // also update TouchWeightsPanel
  edtTouchWeightsMAMS.Value := SampleNr;

  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtSamplePrepNr.MaxValue := NPreps;
  edtSamplePrepNr.Value := NPreps;
  edtSampleTargetNr.MaxValue := NTargets;
  edtSampleTargetNr.Value := NTargets;

  DoSampleInfo(SampleNr, NPreps, NTargets);
  //DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));

  edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
  edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
end;

procedure TfrmMAMS.btnSaveBatchClick(Sender: TObject);
var
  i: integer;
  s, SampNr, PrepNr, TargetNr: string;
begin
{
  if rgpTask.ItemIndex = 0 then begin
    SaveDialog.FilterIndex := 1;
    SaveDialog.InitialDir := edtPretreatmentBatches.Text;
  end
  else
    SaveDialog.FilterIndex := 2;
  if SaveDialog.Execute then begin
    lbxBatch.Items.SaveToFile(SaveDialog.FileName);
    ShowMessage(SaveDialog.FileName + ' gespeichert');
  end;
}
  if rgpTask.ItemIndex = 0 then   // for prep batches
  begin
    for i := 0 to lbxBatch.Count - 1 do
    begin
      SampNr := ExtractWord(1, lbxBatch.Items[i], ['|']);
      PrepNr := ExtractWord(4, lbxBatch.Items[i], ['|']);
      s := 'UPDATE preparation_t SET batch=' + #34 + edtBatchName.Text + #34 +
        ' WHERE preparation_t.sample_nr = ' + SampNr +
        ' AND preparation_t.prep_nr=' + PrepNr +
        ';';
//      ClipBoard.SetTextBuf(PChar(s));
      dm.adoCmd.CommandText := s;
      LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
        Begin
          Try
            dm.adoCmd.Execute;
            LogWindow.addLogEntry('executed');
          Except
            ShowMessage('problem opening the database');
            JvLogFile.Add('DB',lesError,'Problem opening DB');
          End;
        End;
    end;
  end
  else     // for graph batches
  begin
    for i := 0 to lbxBatch.Count - 1 do
    begin
      SampNr := ExtractWord(1, lbxBatch.Items[i], ['|']);
      PrepNr := ExtractWord(4, lbxBatch.Items[i], ['|']);
      TargetNr := ExtractWord(5, lbxBatch.Items[i], ['|']);
      s := 'UPDATE target_t SET graph_batch=' + #34 + edtBatchName.Text + #34 +
        ' WHERE target_t.sample_nr = ' + SampNr +
        ' AND target_t.prep_nr=' + PrepNr +
        ' AND target_t.target_nr=' + TargetNr +
        ';';
//      ClipBoard.SetTextBuf(PChar(s));
      dm.adoCmd.CommandText := s;
      LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;
  end;
  ShowMessage(' Batch data inserted into database');
end;

procedure TfrmMAMS.btnSaveChangesAdminClick(Sender: TObject);
//saves all data in the ProjectInfo pane
var
  sample_nr, project_nr: integer;
  s, in_date_str, desired_date_str, Out_Date_str, freeofcharge, returnToSender, returnedToSender, prep_returnToSender, prep_returnedToSender: string;
begin
  sample_nr := round(edtSampleNr.Value);
  project_nr := dm.GetProjectNrBySampleNr(sample_nr);

  if edtSampleInfoDesiredDate.Date > 1 then
    begin
      desired_date_str := #34 + FormatDateTime('YYYY-MM-DD', edtSampleInfoDesiredDate.Date) + #34;
    end
    else
    begin
      desired_date_str := 'NULL';
    end;

  if edtSampleInfoInDate.Date > 1 then
    begin
      in_date_str := #34 + FormatDateTime('YYYY-MM-DD', edtSampleInfoInDate.Date) + #34;
    end
    else
    begin
      in_date_str := 'NULL';
    end;

  if edtSampleInfoOutDate.Date > 1 then
    begin
      Out_date_str := #34 + FormatDateTime('YYYY-MM-DD', edtSampleInfoOutDate.Date) + #34;
    end
    else
    begin
      Out_date_str := 'NULL';
    end;
  //showmessage('old = ' + FormatDateTime('YYYY-MM-DD', edtSampleInfoOutDate.Date) + ', New= ' + Out_date_str);

  if dbchkFreeOfCharge2.Checked then freeofcharge:='1'
  else freeofcharge:='0';

  if dbchkreturnToSender2.Checked then returnToSender:='1'
  else returnToSender:='0';

  if DBCheckBoxReturnedToSender2.Checked then returnedToSender:='1'
  else returnedToSender:='0';

  if dbchkprepreturnToSender2.Checked then prep_returnToSender:='1'
  else prep_returnToSender:='0';

  if DBCheckBoxPrepReturnedToSender2.Checked then prep_returnedToSender:='1'
  else prep_returnedToSender:='0';

  // update the project dates and status
  s := 'UPDATE project_t SET ' +
    'desired_date= ' + desired_date_str + ',' +
    'in_date=' + in_date_str + ',' +
    'out_date=' + out_date_str + ',' +
    'FreeOfCharge=' + #34 +freeofcharge + #34 + ',' +
    'return_to_sender=' + #34 +returnToSender + #34 + ',' +
    'returned_to_sender=' + #34 +returnedToSender + #34 + ',' +
    'prep_return_to_sender=' + #34 +prep_returnToSender + #34 + ',' +
    'prep_returned_to_sender=' + #34 +prep_returnedToSender + #34 + ',' +
    'project_comment=' + #34 +DBMemo_ProjComment.Text + #34 + ',' +
    'status=' + #34 + cmbProjectStatus.Text + #34 +
    ' WHERE project_nr=' + IntToStr(project_nr) + ';';
  //ShowMessage(s);
  //ClipBoard.SetTextBuf(PChar(s));
  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

// update project_nr that is associated with this sample, project_t
// update user_nr that is associated with this sample, user_t
end;


procedure TfrmMAMS.btnSaveChangesGraphClick(Sender: TObject);
// updates the database with the current data in the fields
var
  s, cmd: string;
  date: TDate;
begin
//  if TargetDataChanged then
    with dm.adoCmd do
    begin
      cmd := 'UPDATE target_t SET ';
      s := edtWeightCombustion.Text;
      s := ReplaceStr(s, ',', '.');
      if length(s) = 0 then s := 'NULL';
      cmd := cmd + 'weight_combustion=' + s;
      s := edtCatalyst.Text;
      s := ReplaceStr(s, ',', '.');
      if length(s) = 0 then s := '0';
      cmd := cmd + ',catalyst=' + s;
      s := edtCO2init.Text;
      s := ReplaceStr(s, ',', '.');
      if Length(s) > 0 then cmd := cmd + ',co2_init=' + s;
      s := edtCO2final.Text;
      s := ReplaceStr(s, ',', '.');
      if Length(s) > 0 then cmd := cmd + ',co2_final=' + s;
      s := edtH2init.Text;
      s := ReplaceStr(s, ',', '.');
      if Length(s) > 0 then cmd := cmd + ',hydro_init=' + s;
      s := edtH2final.Text;
      s := ReplaceStr(s, ',', '.');
      if Length(s) > 0 then
        cmd := cmd + ',hydro_final=' + s
      else
        cmd := cmd + ',hydro_final=NULL ';
      s := edtReactTime.Text;
      s := ReplaceStr(s, ',', '.');
      if Length(s) > 0 then cmd := cmd + ',react_time=' + s;

      // saves graph dates
      // don't check the value since "0" is allowed = NULL
      LogWindow.addLogEntry('updating dates in target table');

      date := edtGraphDate.Date;
      LogWindow.addLogEntry('GraphDate = ' + FloatToStr(date));
      if date > 1 then
      begin
        s := FormatDateTime('YYYY-MM-DD', DateOf(date));
        cmd := cmd + ',graph_date=' + #34 + s + #34;
      end
      else
      begin
        LogWindow.addLogEntry('GraphDate = NULL');
        cmd := cmd + ',graph_date= NULL';
      end;

      date := edtTargetPressed.Date;
      LogWindow.addLogEntry('target pressed date = ' + FloatToStr(date));
      if date > 1 then
      begin
        s := FormatDateTime('YYYY-MM-DD', DateOf(date));
        cmd := cmd + ',target_pressed=' + #34 + s + #34;
      end
      else
      begin
        LogWindow.addLogEntry('target pressed date = NULL');
        cmd := cmd + ',target_pressed=NULL';
      end;

      date := edtGraphitized.Date;
      LogWindow.addLogEntry('graphitized date = ' + FloatToStr(date));
      if date > 1 then
      begin
        s := FormatDateTime('YYYY-MM-DD', DateOf(date));
        cmd := cmd + ',graphitized=' + #34 + s + #34;
      end
      else
      begin
        LogWindow.addLogEntry('graphitized date = NULL');
        cmd := cmd + ',graphitized=NULL'
      end;

      s := edtGraphBatch.Text;
      if Length(s) > 0 then
        cmd := cmd + ',graph_batch=' + #34 + s + #34
      else
        cmd := cmd + ',graph_batch=NULL ';
      if chkTargetDiscarded.Checked then
        cmd := cmd + ',stop=1'
      else
        cmd := cmd + ',stop=0';
      cmd := cmd + ' WHERE sample_nr=' + IntToStr(round(edtSampleNr.Value)) +
        ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) +
        ' AND target_nr=' + IntToStr(round(edtSampleTargetNr.Value)) + ';';
      CommandText := cmd;
      LogWindow.addLogEntry(cmd);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('date update: executed');
        Except
          ShowMessage('date update: problem executing query');
          JvLogFile.Add('DB',lesError,'date update: problem executing query');
        End;
      End;

      // save the comment field
      with dm.qryTest do
      begin
        SQL.Text := 'UPDATE target_t SET target_comment=:TargetMemo' +
          ' WHERE sample_nr=' + IntToStr(round(edtSampleNr.Value)) +
          ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) +
          ' AND target_nr=' + IntToStr(round(edtSampleTargetNr.Value)) + ';';

        // replace query Parameter with actual text from the memo field
        LogWindow.addLogEntry('replace query parameter with memo text = ' + memTargetComments.Lines.Text);
        if Trim(memTargetComments.Lines.Text) = '' then
        begin
          Parameters.ParamByName('TargetMemo').DataType:=ftString;
          Parameters.ParamByName('TargetMemo').Value := NULL;
        end
        else
        begin
          Parameters.ParamByName('TargetMemo').DataType:=ftString;
          Parameters.ParamByName('TargetMemo').Value := memTargetComments.Lines.Text;
        end;
          LogWindow.addLogEntry(SQL.Text);
          JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
          //ClipBoard.SetTextBuf(PChar(SQL.Text));
        Try
          ExecSQL;
          LogWindow.addLogEntry('comment field: change executed');
        Except
          ShowMessage('comment field: problem executing query');
          JvLogFile.Add('DB',lesError,'comment field: problem executing query');
        End;
      end;
    end;

    // now that all is saved set the flag back to 0
    TargetDataChanged := false;
end;

procedure TfrmMAMS.btnSaveChangesPrepClick(Sender: TObject);
var
  sample_nr: integer;
  s, prep_end_str, prep_start_str, PrepStop, swstart, swend, swmedium, swmedium2, scnRatio: string;
  s1, s2, s3, s4, s5: string;
  dt, dt2: TDate;
  CN_Ratio: double;
  leftover: integer;
begin
  LogWindow.addLogEntry('SaveChangesPrepClick');
  sample_nr := round(edtSampleNr.Value);

  // dt := edtPrepEnd.DataSource.DataSet.FieldByName('prep_end').AsDateTime;  //this was done before, don't know why
  // format prep_end Date
  dt := edtPrepEnd.Date;
  if dt > 1 then
  begin
    prep_end_str := #34 + FormatDateTime('YYYY-MM-DD', dt) + #34;
          //ShowMessage('true' + prep_end_str);
  end
  else
  begin
    prep_end_str := 'NULL';
          //ShowMessage('else' + prep_end_str);
  end;

  // format prep_start Date
  dt2 := edtPrepStart.Date;
  if dt2 > 1 then
  begin
    prep_start_str := #34 + FormatDateTime('YYYY-MM-DD', dt2) + #34;
  end
  else
  begin
    prep_start_str := 'NULL';
  end;

  if chkPrepDiscarded.Checked then
    PrepStop := '1'
  else
    PrepStop := '0';
  if chkPrepNoLeftover.Checked then
    leftover := 1
  else
    leftover := 0;
  if Length(edtSampPrep1.Text) = 0 then
    s1 := 'NULL'
  else
    s1 := #34 + edtSampPrep1.Text + #34;
  if Length(edtSampPrep2.Text) = 0 then
    s2 := 'NULL'
  else
    s2 := #34 + edtSampPrep2.Text + #34;
  if Length(edtSampPrep3.Text) = 0 then
    s3 := 'NULL'
  else
    s3 := #34 + edtSampPrep3.Text + #34;
  if Length(edtSampPrep4.Text) = 0 then
    s4 := 'NULL'
  else
    s4 := #34 + edtSampPrep4.Text + #34;
  if Length(edtSampPrep5.Text) = 0 then
    s5 := 'NULL'
  else
    s5 := #34 + edtSampPrep5.Text + #34;
  s := 'UPDATE preparation_t SET ' +
    'step1_method=' + s1 + ',' +
    'step2_method=' + s2 + ',' +
    'step3_method=' + s3 + ',' +
    'step4_method=' + s4 + ',' +
    'step5_method=' + s5 + ',' +
    'prep_start=' + prep_start_str + ',' +
    'prep_end=' + prep_end_str + ',' +
    'p_no_leftover=' + inttostr(leftover) + ',' +
    'stop=' + PrepStop +
    ' WHERE sample_nr=' + IntToStr(sample_nr) +
    ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) + ';';
  //ClipBoard.SetTextBuf(PChar(s));
  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
    Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

  // save weights if they have changed
  if WeightsChanged then
    with dm.adoCmd do
    begin
      swstart := edtWeightStart.Text;
      swstart := ReplaceStr(swstart, ',', '.');
      swend := edtWeightEnd.Text;
      swend := ReplaceStr(swend, ',', '.');
      swmedium := DBEditWeightMedium.Text;        // weightEmpty
      swmedium := ReplaceStr(swmedium, ',', '.');
      swmedium2 := DBEditWeightMedium2.Text;      //weightFull
      swmedium2 := ReplaceStr(swmedium2, ',', '.');
      if Length(swstart) = 0 then
        begin
        swstart := 'NULL';
        end;
      if Length(swend) = 0 then
        begin
        swend := 'NULL';
        end;
      if Length(swmedium) = 0 then
        begin
        swmedium := 'NULL';
        end;
      if Length(swmedium2) = 0 then
        begin
        swmedium2 := 'NULL';
        end;

      CommandText :=  'UPDATE preparation_t SET weight_start=' + swstart + ',' +
                      'weight_end=' + swend + ',' +
                      'weight_medium=' + swmedium + ',' +
                      'weight_medium_2=' + swmedium2 +
                      ' WHERE sample_nr=' + IntToStr(Sample_Nr) + ' AND prep_nr=' +
                      IntToStr(round(edtSamplePrepNr.Value)) + ';';
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;

   // save prep comment field
  with dm.qryTest do
  begin
    if memPrepComments.Lines.Count > 0 then
    begin
      SQL.Text := 'UPDATE preparation_t SET prep_comment=:PrepMemo ' +
        ' WHERE sample_nr=' + IntToStr(sample_nr) +
        ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) + ';';
      Parameters.ParamByName('PrepMemo').DataType:=ftString;
      Parameters.ParamByName('PrepMemo').Value := memPrepComments.Lines.Text;
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      ExecSQL;
    end
    else
    Begin
      SQL.Text := 'UPDATE preparation_t SET prep_comment=NULL' +
        ' WHERE sample_nr=' + IntToStr(sample_nr) +
        ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) + ';';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      ExecSQL;
    End;
  end;
end;

procedure TfrmMAMS.btnSaveChangesUserSuppliedInfoClick(Sender: TObject);
var
  sample_nr,b: integer;
  s, CNIsotopA, CNIsotopAMoved: string;
begin
  sample_nr := round(edtSampleNr.Value);
  // update sample_t name,label,desc1, desc2, material, fraction, type, user prep, comment, old info
  if dbchkCNIsotopA.Checked then CNIsotopA := '1'
  else CNIsotopA := '0';
  if dbchkCNIsotopAMoved.Checked then CNIsotopAMoved := '1'
  else CNIsotopAMoved := '0';

  s := 'UPDATE sample_t SET ' +
    'user_label=' + #34 + edtSampleName.Text + #34 + ',' +
    'user_label_nr=' + #34 + edtSampleLabelNr.Text + #34 + ',' +
    'user_desc1=' + #34 + edtSampleDescr1.Text + #34 + ',' +
    'user_desc2=' + #34 + edtSampleDescr2.Text + #34 + ',' +
    'material=' + #34 + cmbSampleMaterial.Text + #34 + ',' +
    'fraction=' + #34 + cmbSampleFraction.Text + #34 + ',' +
    'type=' + #34 + cmbSampleType.Text + #34 + ',' +
    'user_comment=' + #34 + edtSampleUserComment.Text + #34 + ',' +
    'lab_comment=' + #34 + DBMemo_LabComment.Text + #34 + ',' +
    's_storage_loc=' + #34 + edtSampleStorageLoc.Text + #34 + ',' +
    'prep_storage_loc=' + #34 + edtPrepSampleStorageLoc.Text + #34 + ',' +
    'CNIsotopA=' + #34 + CNIsotopA + #34 + ',' +
    'CNIsotopAMoved=' + #34 + CNIsotopAMoved + #34 + ',' +
    'sampling_date=' + #34 + FormatDateTime('YYYY-MM-DD', edtSamplingDate.Date) + #34 +
    '  WHERE sample_nr=' + IntTostr(sample_nr) + ';';
  //   ClibBoard.SetTextBuf(PChar(s));
  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
    Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
    End;

  //update not to be dated flag
  if chkNotToBeDated.Checked then
    b := 1
  else
    b := 0;

  s := 'UPDATE sample_t SET ' +
    'not_tobedated= ' + IntToStr(b) +
    ' WHERE sample_nr=' + IntToStr(sample_nr) + ';';

  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

  //update no leftover flag
  if chkSampleNoLeftover.Checked then
    b := 1
  else
    b := 0;
  s := 'UPDATE sample_t SET ' +
    's_no_leftover= ' + IntToStr(b) +
    ' WHERE sample_nr=' + IntToStr(sample_nr) + ';';
  //ClipBoard.SetTextBuf(PChar(s));
  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

  end;

procedure TfrmMAMS.btnSaveCNClick(Sender: TObject);
var
  sample_nr: integer;
  s, prep_end_str, PrepStop, swstart, swend: string;
  s1, s2, s3, s4, s5: string;
  dt: TDate;
begin
  sample_nr := round(edtSampleNr.Value);
  with dm.adoCmd do
  begin
    swstart := edtSampleCPercent.Text;
    swstart := ReplaceStr(swstart, ',', '.');
    swend := edtSampleNPercent.Text;
    swend := ReplaceStr(swend, ',', '.');
    CommandText := 'UPDATE target_t SET conc_c=' + swstart + ', ' +
      ' conc_n=' + swend + ' ' +
      ' WHERE sample_nr=' + IntToStr(Sample_Nr) +
      ' AND prep_nr=' + IntToStr(round(edtSamplePrepNr.Value)) +
      ' AND target_nr=' + IntToStr(round(edtSampleTargetNr.Value)) + ';';
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  end;
end;

procedure TfrmMAMS.btnSaveInvoiceNrClick(Sender: TObject);
// if allow changes is sellected the database is updated with the current values of the data fields
begin
    dm.qryProject.Post; //Post has the same result as the SQL command update
end;

procedure TfrmMAMS.btnSaveOptionsClick(Sender: TObject);
begin
// save settings to JvFormStorage
JvFormStorageOptions.SaveFormPlacement;
end;

procedure TfrmMAMS.actTablesExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsTypesMat;
end;

procedure TfrmMAMS.actCameraExecute(Sender: TObject);
begin
  CameraWindow.Show;

  // send active sample number to dialog
  CameraWindow.edtMAMS.Text := edtSampleNr.Text;
end;

procedure TfrmMAMS.actInsertSamplesExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsInsertSamples;
  wizInputSamples.ActivePage := wizStartPage;
  btnInsertNewUser.Visible := false;
  btnInsertExistingUser.Visible := false;
end;

procedure TfrmMAMS.actLabTasksExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsLabTasks;
  SetupLabTaskPage;
end;

procedure TfrmMAMS.actLogWindowExecute(Sender: TObject);
begin
  LogWindow.Show;
  // also show the JvLogFile log
  JvLogFile.ShowLog('logfile.log');
end;

procedure TfrmMAMS.actMagListExecute(Sender: TObject);
begin
//
end;

procedure TfrmMAMS.actOptionsExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsOptions;
  ReportsSetupOptions;
  DisplayProgramOptions;
end;

procedure TfrmMAMS.actSampleInfoExecute(Sender: TObject);
Var SampleNr, NPreps, NTargets: Integer;
begin
  FillPrepList;
  pgtMain.ActivePage := tbsSampleInfo;

  SampleNr := edtSampleNr.Value;
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtSamplePrepNr.MaxValue := NPreps;
  edtSamplePrepNr.Value := NPreps;
  edtSampleTargetNr.MaxValue := NTargets;
  edtSampleTargetNr.Value := NTargets;
  //DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  DoSampleInfo(SampleNr, NPreps, NTargets);
//  btnDoSampleQuery.SetFocus;
  edtSampleNr.SetFocus;
end;

procedure TfrmMAMS.actSearchDatabaseExecute(Sender: TObject);
begin
  frmDBSearch.Show;
end;

procedure TfrmMAMS.actSendMailExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsSendMail;
end;

procedure TfrmMAMS.actStorageLocationExecute(Sender: TObject);
// shows a window (FormStorageLocations in StorageLocations.pas
// that allows to set the storage locations of multiple sampleID's
begin
  FormStorageLocations.ShowModal;
end;

procedure TfrmMAMS.actLabStatsExecute(Sender: TObject);
begin
  GetPlanned;
  pgtMain.ActivePage := tbsLabStats;
end;

procedure TfrmMAMS.actUserInfoExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsUserInfo;
  cmbUsernameUserInfo.ListSource.DataSet.Active := false;
  cmbUsernameUserInfo.ListSource.DataSet.Active := true;

  //ToolButton2.Style.tbsCheck:=true;
  //ToolButton2.Down:=True;
end;

procedure TfrmMAMS.actUserProjectsExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsProjectsOfUser;
  //cmbSubmitterNameProject.SetFocus;
  //ToolButton3.Style.tbsCheck:=true;
  //ToolButton3.Down:=True;
  //cmbSubmitterNameProject.KeyValue := 1;
end;

procedure TfrmMAMS.actReportExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsUserReport;
//  cmbProjectOfReport.Enabled := false;
  cmbSubmNameReport.SetFocus;
  btnReport.Enabled := true;
  btnSendMailEnglish.Enabled := false;
  btnSendMailGerman.Enabled := false;
//  btnQuerySubmitter.Enabled := false;
//  cmbSubmNameReport.DropDown;
end;

procedure TfrmMAMS.actLabPlanExecute(Sender: TObject);
begin
  pgtMain.ActivePage := tbsLabPlan;
  edtBatchName.Text := '';
  btnSaveBatch.Enabled := false;
  btnRefreshSamplesAvailable.Enabled := false;
  grdSamplesAvailable.Visible := false;
  JvDBStatusLabel3.Visible := false;
  cmbPretreatMethod.KeyValue := 0;
  rgpTask.ItemIndex := 1; // set selection to graphitisation already (for no selection use -1)
  rgpTaskClick(self);
//  gbxBatch.Enabled := false;

  GetAvailableOxasBlanks; //get Oxas and Blanks that are currently available for analysis (graphitized)

end;

procedure TfrmMAMS.btnAddNewProjectClick(Sender: TObject);
// create new project for the selected user
Var
  user_nr, project_name, project_nr: string;
  Proj_exists: boolean;
  buttonSelected : Integer;
begin
  user_nr:= cmbSubmitterNameProject.KeyValue;   // read the user_nr of the currently selected user
  project_name := InputBox('Add New Project', 'Project Name (Case Sensitive)', '');  // display InputBox (Case sensitive!!)

  if project_name<>'' then
  begin //check if a project_name was actually given
      //check if project for this user already exists
      Proj_exists:= dm.CheckProjectOfUser_nr(project_name, user_nr.ToInteger());
      if Proj_exists then
      begin
        ShowMessage('Projekt exists already!');
      end
      else
      begin
        // add new project
        buttonSelected:=MessageDlg('Projekt does not exist yet. Proceed?',mtError, mbOKCancel, 0); // show another confirmation dialog
        if buttonSelected=mrOK then
        begin
              //ShowMessage('UserNr: ' + user_nr + ', project_name: ' + project_name);
              project_nr:=dm.AddNewProjectByUserNr(user_nr, project_name);  //add new project and return project number
              GetProjects; // reload projects list
        end;
      end;
  end;

end;

procedure TfrmMAMS.btnAddNewSamplesClick(Sender: TObject);
Var
  project_nr, sample_nr, sample_name: string;
  Sample_exists: boolean;
  buttonSelected : Integer;
begin
  project_nr:= grdProjects.DataSource.DataSet.FieldByName('project_nr').AsString;   // read the project_nr of the currently selected project
  sample_name := InputBox('Add New Sample', 'Sample Name (Case Sensitive)', '');  // display InputBox (Case sensitive!!)

  if sample_name<>'' then begin //check if a sample name was actually given
      //check if sample for this project already exists
      Sample_exists:= dm.CheckSampleOfProjectNr(sample_name, project_nr);
      if Sample_exists then begin
        ShowMessage('Sample exists already!');
      end
      else begin
        // add new sample
        buttonSelected:=MessageDlg('SAmple does not exist yet. Proceed?',mtError, mbOKCancel, 0); // show another confirmation dialog
        if buttonSelected=mrOK then begin
              sample_nr:= dm.AddNewSampleByProjectNr(project_nr, sample_name);  //add new sample and return sample number
              //showmessage(sample_nr);
              dm.CreateBlankPrepRecord(sample_nr.ToInteger, 1); // add new prep to this samples
              dm.CreateBlankTargetRecord(sample_nr.ToInteger, 1, 1); // add new target to this samples
              GetSamples; // reload Sample list
        end;
      end;
  end;

end;

procedure TfrmMAMS.btnAddNewUser2Click(Sender: TObject);
// create a new user in the user table
// only adds a new last name and user_nr
VAR
  s: String;
  form: TfrmNewUser;
begin
   //open multi input dialog
   //form:=frmNewUser.Create(nil);
//   try
//    if frmNewUser.ShowModal = mrOK then        //Add Info into database
//    begin
//    end;
//  finally
//    frmNewUser.Free;
//  end;

   s:='none';
   s:=dm.AddNewUser;
   //ShowMessage('test');                         // add new user; returns user_nr of the newly created user
   if s<>'none' then begin
      StatusBar.Panels[2].Text:='reloading user data...';
      dm.tblUser.Close; dm.tblUser.Open;       //update list of users from the DB by closing na dopening the query
      pgtMain.ActivePage:=tbsUserInfo;    //switch to user info page in order to enter the new user info
      StatusBar.Panels[2].Text:='';
      cmbUsernameUserInfo.KeyValue:= strtoint(s); // dropdown list to correct user_nr
   end;
end;

procedure TfrmMAMS.btnAddNewUserClick(Sender: TObject);
// adds new user to user_t
VAR
  s:string;
begin
  s:='none';
  s:= dm.AddNewUser;    //insert new user into DB
  if s<>'none' then begin
      StatusBar.Panels[2].Text:='reloading user data...';
      dm.tblUser.Close; dm.tblUser.Open;  //get list of users from the DB again
      StatusBar.Panels[2].Text:='';
      cmbUsernameUserInfo.KeyValue:= strtoint(s); // dropdown list to correct user_nr
  end;
end;

procedure TfrmMAMS.btnAdminClick(Sender: TObject);
begin
  pgtMain.ActivePage := tbsAdmin;
  GetListOfRecentMagazines(1);
end;

procedure TfrmMAMS.GetAvailableOxasBlanks;
begin
  memoOxaBlank.Lines.Clear;
  // get number of available oxas
  dm.GetNOxaBlank(true);
  lblOxa.Caption := 'Oxas = ' + IntToStr(dm.qryDB.RecordCount);

    with dm.qryDB do begin
      memoOxaBlank.Lines.Add('Oxas');
      First;
      while not EOF do begin
        memoOxaBlank.Lines.Add(Fields.FieldByName('sample_nr').AsString);
        Next;
      end;
    end;
  // get number of available Pthalic Acids (blanks)
  dm.GetNOxaBlank(false);
  lblBlank.Caption := 'Blanks = ' + IntToStr(dm.qryDB.RecordCount);
    with dm.qryDB do
    begin
      memoOxaBlank.Lines.Add(' ');
      memoOxaBlank.Lines.Add('Blanks');
      First;
      while not EOF do
      begin
        memoOxaBlank.Lines.Add(Fields.FieldByName('sample_nr').AsString);
        Next;
      end;
    end;
end;

procedure TfrmMAMS.btnCalculateMeanClick(Sender: TObject);
var
  n, sum, m_c14_age, dm_c14_age, m_c13: double;
  i, j: integer;
  temp: array of double;
begin
  SetLength(temp, 0); //c14_mean
  sum := 0;
  j := 0;
  with StrGrdTargetData do begin
    for i := 1 to RowCount - 1 do begin
      if (Assigned(Objects[7, i])) then begin
        SetLength(temp, Length(temp) + 1);
        temp[j] := StrToFloat(Cells[4, i]); //column ändern
        TransferPrepNr := StrToInt(cells[1,i]);
        TransferTargetNr := StrToInt(cells[2,i]);
        j := j + 1;
      end;
    end;
    if Length(temp) > 0 then begin
      for i := 0 to length(temp) - 1 do begin
        sum := sum + temp[i];
      end;
      m_c14_age := (sum / length(temp));
      edtMeanAge.Text := FloatToStr(m_c14_age);
    end;
  end;
   //c14_sig
  SetLength(temp, 0);
  sum := 0;
  j := 0;
  with StrGrdTargetData do
  begin
    for i := 1 to RowCount - 1 do
    begin
      if (Assigned(Objects[7, i])) then
      begin
        SetLength(temp, Length(temp) + 1);
        temp[j] := StrToFloat(Cells[5, i]); //column ändern
        j := j + 1;
      end;
    end;
    if Length(temp) > 0 then
    begin
      for i := 0 to length(temp) - 1 do
      begin
        sum := sum + temp[i];
      end;
      dm_c14_age := (sum / length(temp));
      edtSigmaAge.Text := FloatToStr(dm_c14_age);
    end;
  end;
  //c13
  SetLength(temp, 0);
  sum := 0;
  j := 0;
  with StrGrdTargetData do begin
    for i := 1 to RowCount - 1 do begin
      if (Assigned(Objects[7, i])) then begin
        SetLength(temp, Length(temp) + 1);
        temp[j] := StrToFloat(Cells[6, i]); //column ändern
        j := j + 1;
      end;
    end;
    if Length(temp) > 0 then begin
      for i := 0 to length(temp) - 1 do begin
        sum := sum + temp[i];
      end;
      m_c13 := (sum / length(temp));
      edtC13Mean.Text := FloatToStr(m_c13);
    end;
  end;
  if (edtMeanAge.Text <> '') and (edtSigmaAge.Text <> '') and (edtC13Mean.Text <> '') then
    btnTransferTargetData.Enabled := true;
  if length(temp)=1 then OneDateInMean := true;
  
end;

procedure TfrmMAMS.btnChangeProjectClick(Sender: TObject);
//changes the project_nr the sample is associated with
VAR
  new_project, new_project_nr,s,last_name: string;
  sample_nr: integer;
begin
  sample_nr := round(edtSampleNr.Value);
  new_project_nr := InputBox('New Project Nr', 'new project_nr', '');  // display InputBox (Case sensitive!!)
  // check if a valid string was given
 IF new_project_nr <> '' THEN
 BEGIN
    //check if a project exists having the new_project number
    new_project:= dm.GetProjectByProjectNr(new_project_nr.ToInteger());  // look for project number and retunr name
    last_name:= dm.GetUserByProjectNr(new_project_nr.ToInteger()); // get users last name of this project
    if new_project<>'' then
    begin
      // project with new number exists -> re-assign sample to this number
      case MessageDlg('Assign Sample to:' + #13#10 + 'Project: '+ new_project + #13#10 + 'User: ' + last_name + #13#10 + #13#10 + '"OK" will save to Database!!!', mtConfirmation, [mbOK, mbCancel], 0, mbCancel) of
        mrOK:
          begin
            //perform the query that re-assigns the preoject nr
            s:='UPDATE sample_t SET project_nr=' + #34 + new_project_nr + #34 + ' WHERE sample_nr=' + IntToStr(sample_nr) + ';';
            //ShowMessage(s);
            dm.adoCmd.CommandText := s;
            LogWindow.addLogEntry(s);
            IF dm.adoConnKTL.Connected THEN
            Begin
              Try
                dm.adoCmd.Execute;
                LogWindow.addLogEntry('executed');
              Except
                ShowMessage('problem opening the database');
                JvLogFile.Add('DB',lesError,'Problem opening DB');
              End;
            End;
            DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text)); // read sample data again
          end;
        mrCancel:
          begin
            // do nothing
          end;
      end;
    end
    else begin
      ShowMessage('No existing project found using project_nr '+new_project_nr);
    end;

 END;
end;

procedure TfrmMAMS.btnCheckProjectStatusClick(Sender: TObject);
begin
  dm.CheckProjectStatus;
end;

procedure TfrmMAMS.btnCreateSamplesClick(Sender: TObject);
// creates a new sample in sample_t and selected number of targets in target_t
// being used to create entries for Oxa's, blanks etc etc
var
  ProjectName: string;
  i, InternNr, project_nr, sample_nr: integer;
  s, desired_date_str, in_date_str, SampleType: string;
begin
{
 create sample in sample_t, user intern, typ nach rgp, create prep_t, create n entries in target_t
}
    with dm.qryDB do
      begin
      SQL.Text := 'SELECT user_nr from user_t WHERE last_name=' + #34 + 'intern' + #34;
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
        Begin
          Try
            Open;
            LogWindow.addLogEntry('executed');
          Except
            ShowMessage('problem opening the database');
            JvLogFile.Add('DB',lesError,'Problem opening DB');
          End;
        End;
      InternNr := Fields.Fields[0].AsInteger;
      end;

    // generate ProjectName
  case rgpSampleType.ItemIndex of
    0: ProjectName := 'Oxa2 AGE ';
    1: ProjectName := 'Oxa2 MAG ';
    2: ProjectName := 'Oxa2 AGE64 ';
    3: ProjectName := 'Oxa1 ';
    4: ProjectName := 'Blank CO2gas ';
    5: ProjectName := 'Phthalic acid AGE ';
    6: ProjectName := 'Phthalic acid MAG ';
    7: ProjectName := 'Phthalic acid AGE64 ';
    8: ProjectName := 'Blank marble ';
    9: ProjectName := 'Blank bone ';
  end;

   // generate SampleType
  case rgpSampleType.ItemIndex of
    0: SampleType := 'oxa2';
    1: SampleType := 'oxa2';
    2: SampleType := 'oxa2';
    3: SampleType := 'oxa1';
    4: SampleType := 'blank_CO2';
    5: SampleType := 'blank';
    6: SampleType := 'blank';
    7: SampleType := 'blank';
    8: SampleType := 'blank_carb';
    9: SampleType := 'blank_bone';
  end;

   // generate Strings that hold the ProjectName and Daten etc
  ProjectName := ProjectName + FormatDateTime('DDMMYY', Date);
  desired_date_str := FormatDateTime('YYYY-MM-DD', Date);
  in_date_str := FormatDateTime('YYYY-MM-DD', Date);
   // insert new project into database
  dm.adoCmd.CommandText := 'INSERT INTO project_t (user_nr,project,desired_date,'
    + 'in_date,status)'
    + ' VALUES(' + IntToStr(InternNr) + ',' + #34 + ProjectName + #34 + ','
    + #34 + desired_date_str + #34 + ','
    + #34 + in_date_str + #34 + ','
    + #34 + 'closed' + #34
    + ');';
  s := dm.adoCmd.CommandText;
   //  ClibBoard.SetTextBuf(PChar(s));
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  dm.tblProjects.Close;
  dm.tblProjects.Open;
  dm.qryDB.Close; // get project_nr des gerade eingefügten Samples

   // return project number for the project that was just created
  dm.qryDb.SQL.Text := 'SELECT project_nr FROM project_t WHERE project = ' + #34 + ProjectName + #34 + ';';
  LogWindow.addLogEntry(dm.qryDb.SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  if dm.qryDB.RecordCount > 0 then project_nr := dm.qryDb.Fields[0].AsInteger;

  // insert the new sample into the database to the project
  // project name and sample name are the same
  s := 'INSERT INTO sample_t (project_nr,user_label,type)' +
    '  VALUES(' +
    IntToStr(project_nr) + ',' +
    #34 + ProjectName + #34 + ',' +
    #34 + SampleType + #34;
  s := s + ');';
  dm.adoCmd.CommandText := s;
  //   ClibBoard.SetTextBuf(PChar(s));
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  dm.qryDB.Close;

   // get the maximum/latest sample nummer
   // this would be the sample number that was just created
  dm.qryDb.SQL.Text := 'SELECT Max(sample_nr) FROM sample_t;';
  LogWindow.addLogEntry(dm.qryDB.SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  sample_nr := 0;
  if dm.qryDB.RecordCount > 0 then sample_nr := dm.qryDb.Fields[0].AsInteger;

  // insert a new prep for the sample that was just created
  s := 'INSERT INTO preparation_t (prep_nr, sample_nr, prep_end) VALUES(1,' +
    IntToStr(sample_nr) + ',' +
    #34 + in_date_str + #34 +
    ');';
  dm.adoCmd.CommandText := s;
  //   ClibBoard.SetTextBuf(PChar(s));
  LogWindow.addLogEntry(s);
  IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

   // insert new targets for the sample and prep that was just created
  for i := 1 to trunc(edtNumberOfTargets.Value) do
  begin
    dm.adoCmd.CommandText := 'INSERT INTO target_t  (target_nr, prep_nr, sample_nr)'
      + ' VALUES (' + IntToStr(i) + ',1,' + IntTostr(sample_nr) + ');';
    s := dm.adoCmd.CommandText;
     //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(s);
    dm.adoCmd.Execute;
  end;

  // refresh table
  btnRefreshSamplesAvailableClick(self);
end;

procedure TfrmMAMS.btnDeleteUserClick(Sender: TObject);
// remove users without projects
var
  i, b: integer;
begin
  i := 0;
  with dm.tblUser do begin
    First;
    while not EOF do begin
      with dm.qryDB do begin
        SQL.Text := 'SELECT user_nr FROM project_t WHERE user_nr=' +
          dm.tblUser.FieldbyName('user_nr').AsString + ' OR invoice_nr=' +
          dm.tblUser.FieldbyName('user_nr').AsString;
        LogWindow.addLogEntry(SQL.Text);
        JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
        if RecordCount = 0 then begin
//      i := dm.tblUser.FieldbyName('user_nr').AsInteger;
          dm.adoCmd.CommandText := 'DELETE FROM user_t WHERE user_nr=' + dm.tblUser.FieldbyName('user_nr').AsString;
          LogWindow.addLogEntry(dm.adoCmd.CommandText);
        IF dm.adoConnKTL.Connected THEN
        Begin
          Try
            dm.adoCmd.Execute;
            LogWindow.addLogEntry('executed');
          Except
            ShowMessage('problem opening the database');
            JvLogFile.Add('DB',lesError,'Problem opening DB');
          End;
        End;
          inc(i);
        end;
      end;
      Next;
    end;
  end;
  ShowMessage(IntToStr(i) + ' users deleted');
end;

procedure TfrmMAMS.btnDoSampleQueryClick(Sender: TObject);
Var SampleNr, NPreps, NTargets: Integer;
begin
  SampleNr := edtSampleNr.Value;
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtSamplePrepNr.MaxValue := NPreps;
  edtSamplePrepNr.Value := NPreps;
  edtSampleTargetNr.MaxValue := NTargets;
  edtSampleTargetNr.Value := NTargets;

  //DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  DoSampleInfo(SampleNr, NPreps, NTargets);

  //DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  // also update TouchWeightsPanel
  edtTouchWeightsMAMS.Value := edtSampleNr.Value;
  edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
  edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
end;

procedure TfrmMAMS.btnEndTaskClick(Sender: TObject);
begin
  SetTaskEndDate;
end;

procedure TfrmMAMS.btnExportClick(Sender: TObject);
begin
  case rgpExport.ItemIndex of
    0: ExportReport('oxcal');  // OxCal export
    1: with ExcelExport do   // Excel Export
      begin
        Grid := grdSamplesOfSubmitter; //Grid where the data come from
        FileName:= TPath.Combine(edtSaveReporttoFolder.Text, edtReportFileName.Text); //save in the same directory as the report
        LogWindow.addLogEntry('Filename for export = ' + Filename);
        ExportGrid;
      end;
    2: with CSVExportClipBoard do   // Clipboard export
      begin
        Grid := grdSamplesOfSubmitter;
        ExportGrid;
      end;
    3: with HTMLExport do
      begin
        Grid := grdSamplesOfSubmitter;
        ExportGrid;
      end;
    4: with XMLExport do
      begin
        Grid := grdSamplesOfSubmitter;
        ExportGrid;
      end;
  end;
end;

procedure TfrmMAMS.btnFillDateTodayClick(Sender: TObject);
// sets the date fiedls to todays date

begin
  edtGraphDate.Date := Now;
  edtTargetPressed.Date := Now;
  edtGraphitized.Date := Now;

  //set flag to true
  TargetDataChanged := true;

end;

procedure TfrmMAMS.btnGuessReportNameClick(Sender: TObject);
//try to guess the report name using the user name and report number
begin
  edtReportFilename.Text := dm.ReplaceUmlaute(cmbSubmNameReport.Text) + '_' + lblReport.Caption; //replace Umlaute
end;

procedure TfrmMAMS.tnstorClick(Sender: TObject);
begin
  with ExcelExport do begin
    FileName := 'c:\temp\ProjectsSinceYear.xls';
    Grid := grdProjectsSinceYear;
    ExportGrid;
  end;
end;

procedure TfrmMAMS.btnIncSampleNrDownClick(Sender: TObject);
begin
  if edtSampleNr.Value > 1 then begin
    edtSampleNr.Value := edtSampleNr.Value - 1;
    DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  end;
end;

procedure TfrmMAMS.btnIncSampleNrUpClick(Sender: TObject);
var
  MaxSampleNr: integer;
begin
  MaxSampleNr := dm.GetMaxSampleNr;
  if edtSampleNr.Value <= MaxSampleNr then begin
    edtSampleNr.Value := edtSampleNr.Value + 1;
    DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  end;
end;

procedure TfrmMAMS.btnInsertExistingUserClick(Sender: TObject);
begin
  if (MessageDlg('User details of user ' + IntToStr(glbUserNr) + ' will be used; do you want to continue ?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
    wizInputInvoiceAddress.EnableButton(bkNext, true);
    UserExists := true;
    wizInputSamples.SelectNextPage;
  end
  else
    UserExists := false;
  if SameAddressForInvoice then wizInputSamples.SelectNextPage;
end;

procedure TfrmMAMS.btnInsertNewInvoiceClick(Sender: TObject);
begin
  InvoiceExists := false;
  glbInvoiceNr := 0;
  wizInputSamples.SelectNextPage;
end;

procedure TfrmMAMS.btnInsertNewUserClick(Sender: TObject);
begin
  wizInputInvoiceAddress.EnableButton(bkNext, true);
  UserExists := false;
  wizInputSamples.SelectNextPage;
  if SameAddressForInvoice then wizInputSamples.SelectNextPage;
end;

procedure TfrmMAMS.btnInsertSelectedInvoiceClick(Sender: TObject);
begin
  if (MessageDlg('Invoice details of user ' + IntToStr(glbInvoiceNr) + ' will be used; do you want to continue ?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes) then begin
    InvoiceExists := true;
    wizInputSamples.SelectNextPage;
  end;
end;

procedure TfrmMAMS.btnInsertUserClick(Sender: TObject);
var
  s: string;
begin
  with dm.adoCmd, grdPreviewUser do begin
    CommandText := 'INSERT INTO user_t (first_name, last_name,'
      + 'organisation, institute, address_1, '
      + 'address_2, town, postcode,country,'
      + 'phone_1,phone_2, fax,'
      + 'email,www)'
      + ' VALUES (' + #34
      + cells[1, 1] + #34 + ',' + #34
      + cells[1, 2] + #34 + ',' + #34
      + cells[1, 3] + #34 + ',' + #34
      + cells[1, 4] + #34 + ',' + #34
      + cells[1, 5] + #34 + ',' + #34
      + cells[1, 6] + #34 + ',' + #34
      + cells[1, 7] + #34 + ',' + #34
      + cells[1, 8] + #34 + ',' + #34
      + cells[1, 9] + #34 + ',' + #34
      + cells[1, 10] + #34 + ',' + #34
      + cells[1, 11] + #34 + ',' + #34
      + cells[1, 12] + #34 + ',' + #34
      + cells[1, 13] + #34 + ',' + #34
      + cells[1, 14]
      + #34 + ')';
    s := CommandText;
    //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(s);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
  end;
end;

procedure TfrmMAMS.btnLoadBatchClick(Sender: TObject);
begin
  if rgpLabTask.ItemIndex = 0 then
  begin
    with dm.qryActiveBatches do
    begin
      Close;
      SQL.Text := 'SELECT DISTINCT batch FROM preparation_t WHERE batch IS NOT NULL AND ' +
        'prep_end IS NULL ORDER BY batch;';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;
    grdActiveBatches.Visible := true;
  end
  else
  begin
    with dm.qryActiveBatches do
    begin
      Close;
      //SQL.Text := 'SELECT DISTINCT graph_batch FROM target_t WHERE graph_batch IS NOT NULL AND ' +
      //  'graph_date IS NULL AND graphitized IS NULL ORDER BY graph_batch asc;';
      SQL.Text := 'SELECT DISTINCT graph_batch FROM target_t WHERE graph_batch IS NOT NULL ' +
        'ORDER BY graph_date desc;';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      grdActiveBatches.Visible := true;
    end;
  end;
end;

procedure TfrmMAMS.btnTouchPrepDateEndTodayClick(Sender: TObject);
begin
  DBDateTimeTouchPrepEnd.Date := Now;
  btnTouchWeightsPrepNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.btnTouchPrepDateStartTodayClick(Sender: TObject);
begin
  DBDateTimeTouchPrepStart.Date := Now;
  DBDateTimeTouchPrepStart.Update;

  DBchkTouchWeightsSampleNoLeftover.SetFocus;

  //showmessage(FormatDateTime('YYYY-MM-DD',DBDateTimeTouchPrepStart.Date));
  btnTouchWeightsPrepNeedsSaving.Visible := True;

end;

procedure TfrmMAMS.btnTouchWeightsAddToGraphBatchClick(Sender: TObject);
Var s, SampleNr,PrepNr,TargetNr,UserLabel,UserLabelNr: string;
        SampleNrPresent,PrepNrPresent,TargetNrPresent: string;
        i, presentFlag: integer;
begin
  // create new entry in the ListBox
  // get info from currebtly selected sample out of the fields
  SampleNr := edtTouchWeightsMAMS.Text;
  PrepNr := edtTouchWeightsPrepNr.Text;
  TargetNr := edtTouchWeightsTargetNr.Text;
  UserLabel := DBedtTouchWeightsUserLabel.Text;
  UserLabelNr := DBedtTouchWeightsUserLabelNumber.Text;

  // check that sample has no batch name assigned yet
  presentFlag := 0;
  if edtGraphBatch.Text <> '' then
    Begin
     presentFlag := 1;
     showmessage('Sample has a Graph Batch already assigned!');
    End;

    // check that sample has not been measured yet
  if DBedtMagazine.Text <> '' then
    Begin
     presentFlag := 1;
     showmessage('Sample has already been measured - magazine is present!');
    End;

  //check that this sample is not in the list already
  // go through list and check
  if presentFlag = 0 then
    Begin
      for i := 0 to ListBoxTouchWeightsBatch.Count - 1 do
        Begin
          SampleNrPresent := Trim(ExtractWord(1, ListBoxTouchWeightsBatch.Items[i], ['|']));
          PrepNrPresent := Trim(ExtractWord(4, ListBoxTouchWeightsBatch.Items[i], ['|']));
          TargetNrPresent := trim(ExtractWord(5, ListBoxTouchWeightsBatch.Items[i], ['|']));

          if (SampleNr = SampleNrPresent) AND (TargetNr = TargetNrPresent) AND (PrepNr = PrepNrPresent) then
          Begin
            presentFlag := 1;
            showmessage('Sample ' + SampleNr+'.'+PrepNr+'.'+TargetNr+ ' is already listed!');
          End;

        End;
    End;

  // add sample when flag is fine
  if presentFlag = 0 then
    Begin
      // create string for the list
      s := SampleNr + ' | ' +
            UserLabel + ' | ' +
            UserLabelNr  + ' | ' +
            PrepNr + ' | ' +
            TargetNr;
      // add sample to the list
      ListBoxTouchWeightsBatch.Items.Add(s);
      btnTouchWeightsGraphBatchNeedsSaving.Visible := True;
    End;

end;

procedure TfrmMAMS.btnTouchWeightsBatchNameAge1Click(Sender: TObject);
begin
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_age.1';
end;

procedure TfrmMAMS.btnTouchWeightsBatchNameAge2Click(Sender: TObject);
begin
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_age.2';
end;

procedure TfrmMAMS.btnTouchWeightsBatchNameAge641Click(Sender: TObject);
begin
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_age64.1';
end;

procedure TfrmMAMS.btnTouchWeightsBatchNameAge642Click(Sender: TObject);
begin
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_age64.2';
end;

procedure TfrmMAMS.btnTouchWeightsBatchNameAutosamplerClick(Sender: TObject);
begin
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_autosampler';
end;

procedure TfrmMAMS.btnTouchWeightsBatchNameMagClick(Sender: TObject);
begin
  // create a batch name
  edtTouchWeightsBatchName.Text := 'graph_' + FormatDateTime('yymmdd',today) + '_mag';
end;

procedure TfrmMAMS.btnTouchWeightsClearGraphBatchListClick(Sender: TObject);
begin
  ListBoxTouchWeightsBatch.Clear;
  btnTouchWeightsGraphBatchNeedsSaving.Visible := False;
end;

procedure TfrmMAMS.btnTouchWeightsGraphSaveClick(Sender: TObject);
VAR weightCombustion, SampleNr, PrepNr, TargetNr, prepStorLoc, cmd: string;
  flag: integer;
begin
  // get data from Textedits
  //IDs
  SampleNr := edtTouchWeightsMAMS.Text;
  PrepNr := edtTouchWeightsPrepNr.Text;
  TargetNr := edtTouchWeightsTargetNr.Text;
  //weights
  weightCombustion := DBedtTouchWeightsCombustion.Text;
  weightCombustion := ReplaceStr(weightCombustion, ',', '.');

  // save target weight
    with dm.adoCmd do
    begin
      if Length(weightCombustion) = 0 then
        begin
        weightCombustion := 'NULL';
        end;
      CommandText := 'UPDATE target_t SET weight_combustion=' + weightCombustion +
                ' WHERE sample_nr=' + SampleNr +
                ' AND prep_nr=' + PrepNr +
                ' AND target_nr=' + TargetNr + ';';
      LogWindow.addLogEntry('saving combustion weight with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;

      // prep_storage_location
  if dbedtprepStorLoc.Text = '' then
    Begin
       prepStorLoc := '';
    End
    Else
    Begin
       prepStorLoc := dbedtprepStorLoc.Text;
    End;

  // save no leftover flag of prep
    with dm.adoCmd do
    Begin
      if DBchkTouchWeightsPrepNoLeftover.Checked then
        flag := 1
        else
        flag := 0;

    CommandText := 'UPDATE preparation_t SET' +
                ' p_no_leftover=' + inttostr(flag) +
                ' WHERE sample_nr=' + SampleNr +
                ' AND prep_nr=' + PrepNr + ';';
    LogWindow.addLogEntry('saving NoLeftover of Prep flag with query:');
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;

  // save stop flag (discarded) of this target
    with dm.adoCmd do
    Begin
      if DBCheckBoxTouchTargetDiscarded.Checked then
        flag := 1
        else
        flag := 0;

    CommandText := 'UPDATE target_t SET' +
                ' stop=' + inttostr(flag) +
                ' WHERE sample_nr=' + SampleNr +
                ' AND prep_nr=' + PrepNr +
                ' AND target_nr=' + TargetNr + ';';
    LogWindow.addLogEntry('saving stop/discarded flag of target with query:');
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;

      // save "moved to CN" flag of this sample
    with dm.adoCmd do
    Begin
      if DBCheckBoxTouchMovedToCN.Checked then
        flag := 1
        else
        flag := 0;

    CommandText := 'UPDATE sample_t SET' +
                ' CNIsotopAMoved=' + inttostr(flag) +
                ' WHERE sample_nr=' + SampleNr + ';';
    LogWindow.addLogEntry('saving CNIsotopAMoved flag of sample with query:');
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;

      // save target comment
    with dm.adoCmd do
      Begin
          CommandText := 'UPDATE target_t SET target_comment=:TargetMemo' +
            ' WHERE sample_nr=' + SampleNr +
            ' AND prep_nr=' + PrepNr +
            ' AND target_nr=' + TargetNr + ';';
          if Trim(DBMemoTouchWeightsTargetComment.Lines.Text) = '' then
            begin
              Parameters.ParamByName('TargetMemo').DataType:=ftString;
              Parameters.ParamByName('TargetMemo').Value := NULL;
            end
            else
            begin
              Parameters.ParamByName('TargetMemo').Value := DBMemoTouchWeightsTargetComment.Lines.Text;
            end;

      LogWindow.addLogEntry('saving TargetComment with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
        Begin
          Try
            Execute;
            LogWindow.addLogEntry('executed');
            btnTouchWeightsPrepNeedsSaving.Visible := False;
          Except
            btnTouchWeightsPrepNeedsSaving.Visible := True;
            ShowMessage('problem opening the database');
            JvLogFile.Add('DB',lesError,'Problem opening DB');
          End;
      End;
    End;

    // save prep_storage_location
    with dm.adoCmd do
    begin
      CommandText :=  'UPDATE sample_t SET prep_storage_loc=' + #34 + prepStorLoc + #34 +
                      ' WHERE sample_nr=' + SampleNr + ';' ;

      //showMessage(CommandText);

      LogWindow.addLogEntry('saving prep_storage_location with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;

    // set focus to MAMS Nr
    edtTouchWeightsMAMS.SetFocus;
    edtTouchWeightsMAMS.SelectAll;
end;

procedure TfrmMAMS.btnTouchWeightsMAMSDownClick(Sender: TObject);
Var SampleNr, NPreps, NTargets: Integer;
begin
  SampleNr := (edtTouchWeightsMAMS.Value-1);
  edtTouchWeightsMAMS.Value := SampleNr;
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtTouchWeightsPrepNr.MaxValue := NPreps;
  edtTouchWeightsPrepNr.Value := NPreps;
  lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
  edtTouchWeightsTargetNr.MaxValue := NTargets;
  edtTouchWeightsTargetNr.Value := NTargets;
  lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value := edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  // get all sample info from DB
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));

  // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
  // jump to the first weight that is empty
  JumpToEmptyWeightField;

  // update checkbox Sample Archived
  //SetCheckBoxTouchSampleArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  //DisplayArchiveWarning;

end;

procedure TfrmMAMS.btnTouchWeightsMAMSUpClick(Sender: TObject);
Var SampleNr, NPreps, NTargets: Integer;
begin
  SampleNr := (edtTouchWeightsMAMS.Value+1);
  edtTouchWeightsMAMS.Value := SampleNr;
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtTouchWeightsPrepNr.MaxValue := NPreps;
  edtTouchWeightsPrepNr.Value := NPreps;
  lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
  edtTouchWeightsTargetNr.MaxValue := NTargets;
  edtTouchWeightsTargetNr.Value := NTargets;
  lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value := edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  // get all sample info from DB
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));

  // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
  // jump to the first weight that is empty
  JumpToEmptyWeightField;

  // update checkbox Sample Archived
  //SetCheckBoxTouchSampleArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  //DisplayArchiveWarning;
end;

procedure TfrmMAMS.btnTouchWeightsPrepNrDownClick(Sender: TObject);
Var SampleNr, PrepNr, NTargets: Integer;
begin
  edtTouchWeightsPrepNr.Value :=  edtTouchWeightsPrepNr.Value - 1;
  SampleNr := edtTouchWeightsMAMS.Value;
  PrepNr := edtTouchWeightsPrepNr.Value;
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, PrepNr);
  edtTouchWeightsTargetNr.MaxValue := NTargets;
  edtTouchWeightsTargetNr.Value := NTargets;
  lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
  // jump to the first weight that is empty
  JumpToEmptyWeightField;
    // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
end;

procedure TfrmMAMS.btnTouchWeightsPrepNrUpClick(Sender: TObject);
Var SampleNr, PrepNr, NTargets: Integer;
begin
  edtTouchWeightsPrepNr.Value :=  edtTouchWeightsPrepNr.Value + 1;
  SampleNr := edtTouchWeightsMAMS.Value;
  PrepNr := edtTouchWeightsPrepNr.Value;
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, PrepNr);
  edtTouchWeightsTargetNr.MaxValue := NTargets;
  edtTouchWeightsTargetNr.Value := NTargets;
  lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
  // jump to the first weight that is empty
  JumpToEmptyWeightField;
    // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
end;

procedure TfrmMAMS.btnTouchWeightsPrepSaveClick(Sender: TObject);
Var
  weightStart, weightEnd, weightMedium, weightMedium2, SampleNr, PrepNr, TargetNr, cmd, prep_start_str, prep_end_str, sStorLoc: string;
  prep_start_date, prep_end_date : TDate;
  flag: integer;
begin
  // get data from Textedits and reformat
  //IDs
  SampleNr := edtTouchWeightsMAMS.Text;
  PrepNr := edtTouchWeightsPrepNr.Text;
  TargetNr := edtTouchWeightsTargetNr.Text;
  //weights
  weightStart := DBedtTouchWeightsBeforePrep.Text;
  weightStart := ReplaceStr(weightStart, ',', '.');
  weightEnd := DBedtTouchWeightsAfterPrep.Text;
  weightEnd := ReplaceStr(weightEnd, ',', '.');
  weightMedium := DBedtTouchWeightsMediumPrep.Text;
  weightMedium := ReplaceStr(weightMedium, ',', '.');
  weightMedium2 := DBedtTouchWeightsMedium2Prep.Text;
  weightMedium2 := ReplaceStr(weightMedium2, ',', '.');

  // dates
  prep_start_date := DBDateTimeTouchPrepStart.Date;
  prep_end_date := DBDateTimeTouchPrepEnd.Date;
  if prep_start_date > 1 then
  begin
    prep_start_str := #34 + FormatDateTime('YYYY-MM-DD', prep_start_date) + #34;
  end
  else
  begin
    prep_start_str := 'NULL';
  end;
  if prep_end_date > 1 then
  begin
    prep_end_str := #34 + FormatDateTime('YYYY-MM-DD', prep_end_date) + #34;
  end
  else
  begin
    prep_end_str := 'NULL';
  end;

  // s_storage_location
  if dbedtsStorLoc.Text = '' then
    Begin
       sStorLoc := '';
    End
    Else
    Begin
       sStorLoc := dbedtsStorLoc.Text;
    End;

  // save into the Database
  // save prep weights
    with dm.adoCmd do
    begin
      if Length(weightStart) = 0 then
        begin
        weightStart := 'NULL';
        end;
      if Length(weightEnd) = 0 then
        begin
        weightEnd := 'NULL';
        end;
      if Length(weightMedium) = 0 then
        begin
        weightMedium := 'NULL';
        end;
      if Length(weightMedium2) = 0 then
        begin
        weightMedium2 := 'NULL';
        end;
      CommandText :=  'UPDATE preparation_t SET weight_start=' + weightStart + ',' +
                      ' weight_end=' + weightEnd + ',' +
                      ' weight_medium=' + weightMedium + ',' +
                      ' weight_medium_2=' + weightMedium2 +
                      ' WHERE sample_nr=' + SampleNr + ' AND prep_nr=' + PrepNr + ';';
      LogWindow.addLogEntry('saving prep weights with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;

    // save stop flag (discarded) of this prep
    with dm.adoCmd do
    Begin
      if DBCheckBoxTouchPrepDiscarded.Checked then
        flag := 1
        else
        flag := 0;

    CommandText := 'UPDATE preparation_t SET' +
                ' stop=' + inttostr(flag) +
                ' WHERE sample_nr=' + SampleNr +
                ' AND prep_nr=' + PrepNr +';';
    LogWindow.addLogEntry('saving stop/discarded flag of prep with query:');
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;

    // save prep comment
    with dm.adoCmd do
      Begin
          CommandText := 'UPDATE preparation_t SET prep_comment=:PrepMemo ' +
        ' WHERE sample_nr=' + SampleNr +
        ' AND prep_nr=' + PrepNr + ';';
          if Trim(DBMemoTouchWeightsPrepComment.Lines.Text) = '' then
            begin
              Parameters.ParamByName('PrepMemo').DataType:=ftString;
              Parameters.ParamByName('PrepMemo').Value := NULL;
            end
            else
            begin
              Parameters.ParamByName('PrepMemo').Value := DBMemoTouchWeightsPrepComment.Lines.Text;
            end;

      LogWindow.addLogEntry('saving PrepComment with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
        Begin
          Try
            Execute;
            LogWindow.addLogEntry('executed');
            btnTouchWeightsPrepNeedsSaving.Visible := False;
          Except
            btnTouchWeightsPrepNeedsSaving.Visible := True;
            ShowMessage('problem opening the database');
            JvLogFile.Add('DB',lesError,'Problem opening DB');
          End;
      End;
    End;

    // save prep_start and prep_end dates
    with dm.adoCmd do
    begin
      CommandText :=  'UPDATE preparation_t SET prep_start= ' + prep_start_str+ ',' +
                      ' prep_end= ' + prep_end_str +
                      ' WHERE sample_nr=' + SampleNr +
                      ' AND prep_nr=' + PrepNr + ';';

      //Parameters.ParamByName('PrepStart').DataType:=ftString;
      //Parameters.ParamByName('PrepStart').Value := prep_start_str;
      //Parameters.ParamByName('PrepEnd').DataType:=ftString;
      //Parameters.ParamByName('PrepEnd').Value := prep_end_str;

      LogWindow.addLogEntry('saving prep_start and prep_end dates with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;


    // save no leftover flag of sample
    with dm.adoCmd do
    Begin
      if DBchkTouchWeightsSampleNoLeftover.Checked then
        flag := 1
        else
        flag := 0;
    CommandText :=  'UPDATE sample_t SET ' +
                    's_no_leftover= ' + IntToStr(flag) +
                    ' WHERE sample_nr=' + SampleNr + ';';
    LogWindow.addLogEntry('saving NoLeftover of sample flag with query:');
    LogWindow.addLogEntry(CommandText);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;

    // save sample_storage_location
    with dm.adoCmd do
    begin
      CommandText :=  'UPDATE sample_t SET s_storage_loc=' + #34 + sStorLoc + #34 +
                      ' WHERE sample_nr=' + SampleNr + ';' ;

      //showMessage(CommandText);

      LogWindow.addLogEntry('saving s_storage_location with query:');
      LogWindow.addLogEntry(CommandText);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsPrepNeedsSaving.Visible := False;
        Except
          btnTouchWeightsPrepNeedsSaving.Visible := True;
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;

    // set focus to MAMS Nr
    edtTouchWeightsMAMS.SetFocus;
    edtTouchWeightsMAMS.SelectAll;
end;

procedure TfrmMAMS.btnTouchWeightsSaveBatchClick(Sender: TObject);
VAR
  SampleNr, PrepNr, TargetNr, s: string;
  i: integer;
begin
  // extract numbers out of strings in the BatchList
  for i := 0 to ListBoxTouchWeightsBatch.Count - 1 do
    Begin
      SampleNr := Trim(ExtractWord(1, ListBoxTouchWeightsBatch.Items[i], ['|']));
      PrepNr := Trim(ExtractWord(4, ListBoxTouchWeightsBatch.Items[i], ['|']));
      TargetNr := trim(ExtractWord(5, ListBoxTouchWeightsBatch.Items[i], ['|']));
      // save batch name to DB
    s := 'UPDATE target_t SET graph_batch=' + #34 + edtTouchWeightsBatchName.Text + #34 +
        ' WHERE target_t.sample_nr = ' + SampleNr +
        ' AND target_t.prep_nr=' + PrepNr +
        ' AND target_t.target_nr=' + TargetNr +
        ';';
    dm.adoCmd.CommandText := s;
    LogWindow.addLogEntry('### write graph batch to DB using query:');
    LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
          btnTouchWeightsGraphBatchNeedsSaving.Visible := False;
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    End;



end;

procedure TfrmMAMS.btnTouchWeightsTargetNrDownClick(Sender: TObject);
begin
  edtTouchWeightsTargetNr.Value := edtTouchWeightsTargetNr.Value - 1;
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Value));
  // jump to the first weight that is empty
  JumpToEmptyWeightField;
    // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
end;

procedure TfrmMAMS.btnTouchWeightsTargetNrUpClick(Sender: TObject);
begin
  edtTouchWeightsTargetNr.Value := edtTouchWeightsTargetNr.Value + 1;
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Value));
  // jump to the first weight that is empty
  JumpToEmptyWeightField;
    // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
end;

procedure TfrmMAMS.btnMagazinePressedDateTodayClick(Sender: TObject);
var
  s, magazine: string;
begin
  magazine := grdMagazinesUnpressed.DataSource.DataSet.FieldByName('magazine').AsString;
 if magazine = '' then
     begin
      ShowMessage('no magazine selected');
      end
 else
     begin
        if dm.adoConnKTL.Connected then
          Begin
          Screen.Cursor:=CrHourGlass;
          LogWindow.addLogEntry('Updating pressed date date to today for all targets of the magazine: ' + magazine);
          s := 'UPDATE target_t ' +
                ' SET target_t.target_pressed = curdate() ' +
               ' WHERE target_t.magazine = ' + #34 + magazine + #34 + ';';
          ShowMessage(s);
          dm.adoCmd.CommandText := s;
          LogWindow.addLogEntry(s);
            try
              dm.adoCmd.Execute;
              LogWindow.addLogEntry('Update performed.');
              Screen.Cursor:=CrDefault;
              ShowMessage('target_pressed for all targets set to today');
            Except
              LogWindow.addLogEntry('Update NOT performed.');
              Screen.Cursor:=CrDefault;
              ShowMessage('Problem updating date');
            end;
          end;
     end;
    end;

procedure TfrmMAMS.btnMonthStatClick(Sender: TObject);
var
  i, n_graph, n_meas, sum_graph, sum_meas, n_recvd, sum_recvd: integer;
begin
  sum_graph := 0;
  sum_meas := 0;
  sum_recvd := 0;
  // create headers and sizes
  stgMonthStat.Cells[0, 0] := 'month';
  stgMonthStat.ColWidths[0] := 40;

  stgMonthStat.Cells[1, 0] := 'received';
  stgMonthStat.ColWidths[1] := 80;

  stgMonthStat.Cells[2, 0] := 'graphitized';
  stgMonthStat.ColWidths[2] := 80;

  stgMonthStat.Cells[3, 0] := 'measured';
  stgMonthStat.ColWidths[3] := 80;

  ActivityIndicator_processed_samples.BringToFront;
  ActivityIndicator_processed_samples.StartAnimation;
  ActivityIndicator_processed_samples.Animate:= True;

  // go through every month of that year and query DB
  // fill in data into the string grid
  for i := 1 to 12 do
  begin
    //n := dm.GetArchSamplesOfMonth(trunc(edtMonthStat.Value), i);
    n_recvd := dm.GetAllSamplesOfMonth(trunc(edtMonthStat.Value), i, 0);
    n_graph := dm.GetAllSamplesOfMonth(trunc(edtMonthStat.Value), i, 1);
    n_meas := dm.GetAllSamplesOfMonth(trunc(edtMonthStat.Value), i, 2);
    stgMonthStat.Cells[0, i] := IntToStr(i);
    stgMonthStat.Cells[1, i] := IntToStr(n_recvd);
    stgMonthStat.Cells[2, i] := IntToStr(n_graph);
    stgMonthStat.Cells[3, i] := IntToStr(n_meas);
    sum_recvd := sum_recvd + n_recvd;
    sum_graph := sum_graph + n_graph;
    sum_meas := sum_meas + n_meas;
    Application.ProcessMessages;
  end;

  // fill in sum of all counts
  stgMonthStat.Cells[0, 13] := 'total';
  stgMonthStat.Cells[1, 13] := IntToStr(sum_recvd);
  stgMonthStat.Cells[2, 13] := IntToStr(sum_graph);
  stgMonthStat.Cells[3, 13] := IntToStr(sum_meas);

  ActivityIndicator_processed_samples.StopAnimation;
  ActivityIndicator_processed_samples.SendToBack;
end;

procedure TfrmMAMS.btnNewPrepClick(Sender: TObject);
var
  MaxPrepNr, MaxTargetNr: integer;
begin
  if (MessageDlg('Creating one more preparation nr.; ' +
    'do you want to continue ?', mtConfirmation, [mbYes, mbNo], 0) in [mrYes]) then
  begin
    MaxPrepNr := dm.GetMaxPrepNrBySampleNr(round(edtSampleNr.Value));
    dm.CreateBlankPrepRecord(round(edtSampleNr.Value), succ(MaxPrepNr));
    dm.CreateBlankTargetRecord(round(edtSampleNr.Value), succ(MaxPrepNr), 1);
  end;
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
end;

procedure TfrmMAMS.btnNewTargetClick(Sender: TObject);
var
  MaxTargetNr: integer;
begin
  if (MessageDlg('Creating one more target nr.; ' +
    'do you want to continue ?', mtConfirmation, [mbYes, mbNo], 0) in [mrYes]) then begin

    MaxTargetNr := dm.GetMaxTargetNrBySampleNr(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text));
    dm.CreateBlankTargetRecord(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), succ(MaxTargetNr));
  end;
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
end;

procedure TfrmMAMS.btnNewUserClick(Sender: TObject);
begin
  UserExists := false;
end;

procedure TfrmMAMS.edtBatchNameChange(Sender: TObject);
begin
  //enable save button
  btnSaveBatch.Enabled := (Length(edtBatchName.Text) > 0) and (lbxBatch.Items.Count > 0);
  //show print button only if prep batch is selected
  btnPrintBatch.Visible := (Length(edtBatchName.Text) > 0) and (rgpTask.ItemIndex = 0);
  edtFilenamePrepDocTemplate.Visible := (Length(edtBatchName.Text) > 0) and (rgpTask.ItemIndex = 0);
end;

procedure TfrmMAMS.edtCatalystChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtChangeDBClick(Sender: TObject);
begin
  if dm.adoConnKTL.Connected then
    DbClosed := true;
  try
    dm.adoConnKTL.Close;
  except
    dbClosed := false;
  end;
  if not dbClosed then begin
    Exit;
  end;
//  if EditConnectionString(adoConnKTL) then
  DSN := dm.adoConnKTL.ConnectionString
//  else
end;

procedure TfrmMAMS.edtCO2finalChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtCO2initChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtGraphBatchChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtGraphDateChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtGraphDateClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtGraphitizedChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtGraphitizedClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtH2finalChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtH2initChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtNewSamplesFilenameAfterDialog(Sender: TObject;
  var AName: string; var AAction: Boolean);
begin
  ParseNewSampleFile(AName);
end;

procedure TfrmMAMS.edtPrepEndClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtPrepStartClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtProjectNameChange(Sender: TObject);
begin
  ProjectName := edtProjectName.Text;
  ProjectName := copy(ProjectName, 1, 99);
  if Length(edtProjectName.Text) > 0 then wizInputProject.EnableButton(bkNext, true);
end;


procedure TfrmMAMS.edtReactTimeChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtReportFileNameChange(Sender: TObject);
begin
btnReport.Enabled:=true;
end;

procedure TfrmMAMS.edtSampleCPercentChange(Sender: TObject);
begin
  CalculateCAbsWeight;
end;

procedure TfrmMAMS.edtSampleInfoDesiredDateClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtSampleInfoInDateClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;

end;

procedure TfrmMAMS.edtSampleInfoOutDateClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;

end;

procedure TfrmMAMS.edtSampleNrChange(Sender: TObject);
begin
  // also update the editButton in TouchWeights to reflect the same value
    //DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
  //edtTouchWeightsMAMS.Value := edtSampleNr.Value;
  //edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
  //edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
end;

procedure TfrmMAMS.edtSampleNrKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var SampleNr, NPreps, NTargets: Integer;
begin
  if Key = VK_RETURN then
  Begin
    SampleNr := edtSampleNr.Value;
    NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
    NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);

    // btnDoSampleQueryClick(self);
    DoSampleInfo(SampleNr, NPreps, NTargets);

    // // also update the editButton in TouchWeights to reflect the same value
    edtTouchWeightsMAMS.Value := SampleNr;
    edtTouchWeightsPrepNr.Value := NPreps;
    edtTouchWeightsTargetNr.Value := NTargets;
  End;
  if Key = VK_UP then
  Begin
    edtSampleNr.Value:= round(edtSampleNr.Value)+1;
    DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    // also update the editButton in TouchWeights to reflect the same value
    edtTouchWeightsMAMS.Value := edtSampleNr.Value;
    edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
    edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
  End;
  if Key = VK_DOWN then
  begin
    edtSampleNr.Value:= round(edtSampleNr.Value)-1;
    DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    // also update the editButton in TouchWeights to reflect the same value
    edtTouchWeightsMAMS.Value := edtSampleNr.Value;
    edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
    edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
  end;
end;

procedure TfrmMAMS.edtSamplePrepNrChange(Sender: TObject);
begin
  LogWindow.addLogEntry('edtSampleTargetNrChange');
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    // also update the editButton in TouchWeights to reflect the same value
  edtTouchWeightsMAMS.Value := edtSampleNr.Value;
  edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
  edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
end;

procedure TfrmMAMS.edtSampleTargetNrChange(Sender: TObject);
begin
  LogWindow.addLogEntry('edtSampleTargetNrChange');
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    // also update the editButton in TouchWeights to reflect the same value
  //edtTouchWeightsMAMS.Value := edtSampleNr.Value;
  //edtTouchWeightsPrepNr.Value := edtSamplePrepNr.Value;
  //edtTouchWeightsTargetNr.Value := edtSampleTargetNr.Value;
end;

procedure TfrmMAMS.edtSampPrep1DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  edtSampPrep1.ReadOnly := false;
  edtSampPrep1.Text := Prep;
  edtSampPrep1.ReadOnly := true;
end;

procedure TfrmMAMS.edtSampPrep1DragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  Accept := true;
end;

procedure TfrmMAMS.edtSampPrep2DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  edtSampPrep2.ReadOnly := false;
  edtSampPrep2.Text := Prep;
  edtSampPrep2.ReadOnly := true;
end;

procedure TfrmMAMS.edtSampPrep3DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  edtSampPrep3.ReadOnly := false;
  edtSampPrep3.Text := Prep;
  edtSampPrep3.ReadOnly := true;
end;

procedure TfrmMAMS.edtSampPrep4DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  edtSampPrep4.ReadOnly := false;
  edtSampPrep4.Text := Prep;
  edtSampPrep4.ReadOnly := true;
end;

procedure TfrmMAMS.edtSampPrep5DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  edtSampPrep5.ReadOnly := false;
  edtSampPrep5.Text := Prep;
  edtSampPrep5.ReadOnly := true;
end;


procedure TfrmMAMS.edtSubmNameEnter(Sender: TObject);
begin
  dm.tblUser.IndexFieldNames := 'last_name';
end;

procedure TfrmMAMS.edtSubmNameExit(Sender: TObject);
begin
//  dm.tblUser.IndexFieldNames := 'user_nr';
  if pgtMain.ActivePage = tbsUserReport then btnQuerySubmitter.SetFocus;
end;

procedure TfrmMAMS.edtTargetPressedChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.edtTargetPressedClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.edtTouchWeightsMAMSChange(Sender: TObject);
begin
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  LogWindow.addLogEntry('edtTouchWeightsMAMSChange');

  edtSampleNr.Value:= edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;

  // get all sample info from DB
  //DoSampleInfo(StrToInt(edtTouchWeightsMAMS.Text), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
  dm.GetSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Value), StrToInt(edtTouchWeightsTargetNr.Value));

  //showmessage('start start start' + edtTouchWeightsMAMS.Text);

  sleep(100);
  SetCheckBoxTouchSampleArchived;
  SetCheckBoxTouchPrepArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarning;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.edtTouchWeightsMAMSClick(Sender: TObject);
begin
  (Sender AS TJvValidateEdit).SelectAll;
end;

procedure TfrmMAMS.edtTouchWeightsMAMSKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var SampleNr, NPreps, NTargets: Integer;
  begin
  if Key = VK_RETURN then
    Begin
      SampleNr := edtTouchWeightsMAMS.Value;
      // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
      NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
      NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
      edtTouchWeightsPrepNr.MaxValue := NPreps;
      edtTouchWeightsPrepNr.Value := NPreps;
      lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
      edtTouchWeightsTargetNr.MaxValue := NTargets;
      edtTouchWeightsTargetNr.Value := NTargets;
      lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
      DoSampleInfo(SampleNr, StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
      // jump to the first weight that is empty
      JumpToEmptyWeightField;
      // hide some buttons
      btnTouchWeightsPrepNeedsSaving.Visible := False;
      btnTouchWeightsGraphNeedsSaving.Visible := False;
    End;
  if Key = VK_UP then
    Begin
      SampleNr := (edtTouchWeightsMAMS.Value+1);
      edtTouchWeightsMAMS.Value := SampleNr;
      // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
      NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
      NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
      edtTouchWeightsPrepNr.MaxValue := NPreps;
      edtTouchWeightsPrepNr.Value := NPreps;
      lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
      edtTouchWeightsTargetNr.MaxValue := NTargets;
      edtTouchWeightsTargetNr.Value := NTargets;
      lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
      DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
      // jump to the first weight that is empty
      JumpToEmptyWeightField;
      // hide some buttons
      btnTouchWeightsPrepNeedsSaving.Visible := False;
      btnTouchWeightsGraphNeedsSaving.Visible := False;
    End;
  if Key = VK_DOWN then
    begin
      SampleNr := (edtTouchWeightsMAMS.Value-1);
      edtTouchWeightsMAMS.Value := SampleNr;
      // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
      NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
      NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
      edtTouchWeightsPrepNr.MaxValue := NPreps;
      edtTouchWeightsPrepNr.Value := NPreps;
      lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
      edtTouchWeightsTargetNr.MaxValue := NTargets;
      edtTouchWeightsTargetNr.Value := NTargets;
      lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
      DoSampleInfo(round(edtTouchWeightsMAMS.Value), StrToInt(edtTouchWeightsPrepNr.Text), StrToInt(edtTouchWeightsTargetNr.Text));
      // jump to the first weight that is empty
      JumpToEmptyWeightField;
      // hide some buttons
      btnTouchWeightsPrepNeedsSaving.Visible := False;
      btnTouchWeightsGraphNeedsSaving.Visible := False;
    end;
      // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value := edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;
end;

procedure TfrmMAMS.edtTouchWeightsPrepNrChange(Sender: TObject);
begin
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value:= edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;

  SetCheckBoxTouchSampleArchived;
  SetCheckBoxTouchPrepArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarning;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.edtTouchWeightsPrepNrClick(Sender: TObject);
begin
  (Sender AS TJvValidateEdit).SelectAll;
end;

procedure TfrmMAMS.edtTouchWeightsTargetNrChange(Sender: TObject);
begin
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value:= edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;

  SetCheckBoxTouchSampleArchived;
  SetCheckBoxTouchPrepArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarning;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.edtTouchWeightsTargetNrClick(Sender: TObject);
begin
  (Sender AS TJvValidateEdit).SelectAll;
end;

procedure TfrmMAMS.edtWeightEndChange(Sender: TObject);
begin
  LogWindow.addLogEntry('edtWeightEndChange');
  WeightsChanged := true;
  CalculateYield; // calculate the yield from the weights
end;

procedure TfrmMAMS.edtWeightEndExit(Sender: TObject);
begin
// set PrepEnd Date to today automatically if not set already and if weight_end is set
  if (Sender AS TDBEdit).Text <> '' then begin
      if edtPrepEnd.Date = 0 then begin
          edtPrepEnd.DataSource.Edit;
          edtPrepEnd.Date:= DATE;
          //edtPrepEnd.Update;
    //edtPrepEnd.Refresh;
      end;
  end;
end;

procedure TfrmMAMS.edtWeightStartChange(Sender: TObject);
begin
  WeightsChanged := true;
  CalculateYield; // calculate the yield from the weights
end;

procedure TfrmMAMS.ExportReport(Ext: string);
var
//  OutF : TextFile;
  Line: string;
  Age, Err: integer;
  StringList: TStringList;
begin
  StringList := TStringList.Create;
//  AssignFile(OutF,'c:\temp\cal.'+Ext);
//  Rewrite(OutF);
  ClipBoard.Open;
  with dm.qrySampleOfSubmitter do begin
    First;
    while not EOF do begin
      Age := Fields.FieldByName('C14_age').AsInteger;
      Err := Fields.FieldByName('C14_age_sig').AsInteger;
//      Line := 'MAMS ' + Fields.FieldByName('sample_nr').AsString + ' ' + Fields.FieldByName('user_label').AsString;
      //Line := Fields.FieldByName('user_label').AsString;
      Line := Fields.FieldByName('sample_nr').AsString;
      //Line := Line + #9 + IntToStr(Age) + #9 + IntToStr(Err);  //old way of doing it. values with tabs in between
      Line := 'R_date(' + '"' + Line + '",' + IntToStr(Age) + ',' + IntToStr(Err)+');';
      //      WriteLn(OutF,Line);
      StringList.Add(Line);
      Next;
    end;
    ClipBoard.AsText := StringList.Text;
  end;
//  CloseFile(OutF);
  ClipBoard.Close;
  StringList.Free;
end;

procedure TfrmMAMS.ImageFilesListBoxClick(Sender: TObject);
VAR
  fname: string;
  ListBox: TListBox;
begin
// load image according to the path that is selected in the ListBox
   ListBox := Sender as TListBox;
   fname := ListBox.Items[ListBox.ItemIndex];
   SampleFoto.Visible:=True;
   LogWindow.addLogEntry('selected image: ' + fname);
   if FileExists(fname) then
     begin
       SampleFoto.Picture.LoadFromFile(fname);
       LogWindow.addLogEntry('Image found: '+fname);
     end
     else
     begin
       ShowMessage('Unable to find image. Drive is connected but image is missing');
       LogWindow.addLogEntry('Image file does not exist!');
     end;
end;

procedure TfrmMAMS.InsertNewSamplesInDb;
// used for sample import using a wizzard
// Insert SAmples into Database
var
  Sample_nr, user_nr, Acol, Arow, PrepCol, FreeOfCharge, i: integer;
  InsertReturnToSender, InsertprepReturnToSender, InsertAllCNIsotopA: integer;
  project_nr: integer;
  s, desired_date_str, in_date_str, strWeight: string;
  w: double;
  SaveDecimalSep: char;
  buttonDlg: integer;
begin
//  insert user, if new
//  get user_nr
//  insert Sample
//  insert new samples

  lbWizFinalPage.Clear;
// insert User into DB if user is new
  if not UserExists then
  begin
    with grdPreviewUser do
    begin
      User.FirstName := cells[1, 1];
      User.LastName := cells[1, 2];
      User.Organisation := cells[1, 3];
    end;
    s := 'INSERT INTO user_t (first_name, last_name,'
      + 'organisation, institute, address_1, '
      + 'address_2, town,postcode,'
      + 'country,phone_1, phone_2, '
      + 'fax,'
      + 'email,www,invoice, correspondance)'
      + ' VALUES (';
    for Arow := 1 to 13 do
    begin // insert values for first_name to email into query
      if Length(grdPreviewUser.cells[1, Arow]) > 0 then
        s := s + #34 + grdPreviewUser.cells[1, Arow] + #34 + ','
      else
        s := s + 'NULL,'
    end;
    if Length(grdPreviewUser.cells[1, 14]) > 0 then // final entry
      s := s + #34 + grdPreviewUser.cells[1, Arow] + #34 + ','
    else
      s := s + 'NULL,';
    if SameAddressForInvoice then
      s := s + '1,1);'
    else
      s := s + '0,1);'; //query ends here
    dm.adoCmd.CommandText := s;
    //ClipBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry('### IMPORT: Import User...');
    LogWindow.addLogEntry(s);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
              LogWindow.addLogEntry('### IMPORT: Import User...Done');
        Except
          ShowMessage('problem opening the database during import: user');
          JvLogFile.Add('DB',lesError,'Problem opening DB during import: user');
        End;
      End;
    lbWizFinalPage.lines.add('New user created');
    //reload user table in order to update new user
    dm.tblUser.Close;
    dm.tblUser.Open;

    // return user_nr of the newly created user
    with dm.qryDb do
    begin // get user_nr from user_t
      Close;
      SQL.Text := 'SELECT last_insert_id()';
      //SQL.Text := 'SELECT user_nr FROM user_t WHERE last_name=' + #34 + User.LastName + #34 +
      //  ' AND first_name=' + #34 + User.FirstName + #34;
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database reading user number');
          JvLogFile.Add('DB',lesError,'Problem opening DB reading user number');
        End;
      End;
      user_nr := 2;  // currently there is no user with user_nr=2 in the database
      //if RecordCount > 0 then user_nr := FieldByName('user_nr').AsInteger;
      if RecordCount > 0 then
        begin
          if FieldByName('last_insert_id()').AsInteger > 0 then
            user_nr := FieldByName('last_insert_id()').AsInteger;
        end;
      lbWizFinalPage.lines.add('New User Nr = ' + inttostr(user_nr));
    end;
  end
  else

  // user already exists
  begin
    lbWizFinalPage.Lines.Add('User already known');
    user_nr := glbUserNr;
  end;
  lbWizFinalPage.Lines.Add('User Nr used for import: ' + inttostr(user_nr));

  //insert Invoice Information
  if not SameAddressForInvoice then
   begin // insert invoice user
    if not InvoiceExists then
    begin
      with grdInvoiceAddress do
      begin
        Invoice.LastName := cells[1, 2];
        InVoice.Institute := cells[1, 4];
        Invoice.Organisation := cells[1, 3];
      end;
      s := 'INSERT INTO user_t ( first_name, last_name,'
        + 'organisation, institute, address_1, '
        + 'address_2, town, postcode,'
        + 'country,phone_1, phone_2, '
        + 'fax,'
        + 'email,www, invoice, correspondance)'
        + ' VALUES (';
      for Arow := 1 to 13 do
      begin // first_name to email
        if Length(grdInvoiceAddress.cells[1, Arow]) > 0 then
          s := s + #34 + grdInvoiceAddress.cells[1, Arow] + #34 + ','
        else
          s := s + 'NULL,'
      end;
      if Length(grdInvoiceAddress.cells[1, 14]) > 0 then // final entry
        s := s + #34 + grdInvoiceAddress.cells[1, Arow] + #34 + ','
      else
        s := s + 'NULL,';
      s := s + '1,0);';
      dm.adoCmd.CommandText := s;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry('### IMPORT: Import Invoice...');
      LogWindow.addLogEntry(s);
          IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
          LogWindow.addLogEntry('### IMPORT: Import Invoice...Done');
        Except
          ShowMessage('problem opening the database during import: invoice');
          JvLogFile.Add('DB',lesError,'Problem opening DB during import: invoice');
        End;
      End;
      lbWizFinalPage.lines.add('New invoice user created');
      with dm.qryDb do
      begin // get user_nr from user_t
        Close;
        SQL.Text := 'SELECT user_nr FROM user_t WHERE last_name=' + #34 + Invoice.LastName + #34 +
          ' AND organisation=' + #34 + Invoice.Organisation + #34;
        LogWindow.addLogEntry(SQL.Text);
        JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database reading user number');
          JvLogFile.Add('DB',lesError,'Problem opening DB user number');
        End;
      End;
        glbInvoiceNr := 1;
        if RecordCount > 0 then glbInvoiceNr := FieldByName('user_nr').AsInteger;
      end;
    end
    else
    begin
      lbWizFinalPage.lines.add('<b> Invoice user already known');
    end;
  end;

//Import Project Data
  //check if this user already has a project with the same name
  CheckProjectExists;
  // project doesn't exists yet
  if not ProjectExists then
  begin
    // prepare a few data types
    desired_date_str := FormatDateTime('YYYY-MM-DD', edtDesiredDate.Date);
    in_date_str := FormatDateTime('YYYY-MM-DD', edtInDate.Date);
    if chkFreeOfCharge.Checked = true then FreeOfCharge:=1
      else FreeOfCharge:=0;
    if chkInsertReturnToSender.Checked = true then InsertReturnToSender:=1
      else InsertReturnToSender:=0;
    if chkInsertprepReturnToSender.Checked = true then InsertprepReturnToSender:=1
      else InsertprepReturnToSender:=0;

    if glbInvoiceNr > 0 then
      s := IntTostr(glbInvoiceNr)
    else
      s := 'NULL';
    dm.adoCmd.CommandText := 'INSERT INTO project_t (user_nr,invoice_nr,project,desired_date,'
      + 'in_date,priority,report_type,letter,status,price,project_type,research, project_comment, invoice, FreeOfCharge,return_to_sender,prep_return_to_sender)'
      + ' VALUES(' + IntToStr(user_nr) + ',' + s + ','
      + #34 + ProjectName + #34 + ','
      + #34 + desired_date_str + #34 + ','
      + #34 + in_date_str + #34 + ','
      + IntToStr(rgpPriority.ItemIndex) + ','
      + #34 + cmbReportType.Text + #34 + ','
      + #34 + edtUserDocuments.Directory + #34 + ','
      + #34 + 'planned' + #34 + ','
      + #34 + edtPrice.Text + #34 + ','
      + #34 + cmbProjectType.Text + #34 + ','
      + #34 + cmbResearchType.Text + #34 + ','
      + #34 + edtProjectComment.Text + #34 + ','
      + IntToStr(MANummer) + ','
      + IntToStr(FreeOfCharge) + ','
      + IntToStr(InsertReturnToSender) + ','
      + IntToStr(InsertprepReturnToSender) + ')'
      + ';';
    s := dm.adoCmd.CommandText;
    //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry('### IMPORT: Import Project data...');
    LogWindow.addLogEntry(s);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          //showmessage(s);
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
          LogWindow.addLogEntry('### IMPORT: Import Project data...Done');
        Except
          ShowMessage('problem opening the database during import: project data');
          JvLogFile.Add('DB',lesError,'Problem opening DB during import: project data');
        End;
      End;
    lbWizFinalPage.lines.add('New project created');
    //update list of projects
    dm.tblProjects.Close;
    dm.tblProjects.Open;
  end
  //project exists already
  else
  begin
    lbWizFinalPage.lines.add('Project already known');
  end;

  // get project number in order to insert the samples
  // PROBLEM PROBLEM PROBLEM PROBLEM
  dm.qryDB.Close; // get project_nr des gerade eingefügten Samples  !!!!!!!!!! project name is NOT unique PROBLEM!!!!
  //dm.qryDb.SQL.Text := 'SELECT project_nr FROM project_t WHERE project = ' + #34 + ProjectName + #34 + ';';// add WHERE project = ' + #34 + ProjectName + #34 + "AND user_nr= '+ #34 + user_nr + #34

  // new query hopefully solves the problem with non-unique projet names
  dm.qryDb.SQL.Text := 'SELECT project_nr FROM project_t WHERE project = ' + #34 + ProjectName + #34 + 'AND user_nr= '+ #34 + inttostr(user_nr) + #34 + ';';
  LogWindow.addLogEntry(dm.qryDB.SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.qryDB.Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database reading project number');
          JvLogFile.Add('DB',lesError,'Problem opening DB reading project number');
        End;
      End;
  if dm.qryDB.RecordCount > 0 then
  begin
  project_nr := dm.qryDb.Fields[0].AsInteger;
  lbWizFinalPage.lines.add('Project Nr: ' + inttostr(project_nr));
  end;

// insert samples into database
  // generate flag from the checkbox whether or not all samples should be set to CNIsotopA
  if checkInsertAllCNIsotopA.Checked = true then InsertAllCNIsotopA:=1
  else InsertAllCNIsotopA:=0;

  // go through the sample grid and generate sqlquery string from the grid entries for each sample
  with grdPreviewSamples do
  begin // insert samples one by one at each loop
    for Arow := 1 to RowCount - 1 do
    begin
      s := 'INSERT INTO sample_t (project_nr,user_label,user_label_nr,user_desc1, user_desc2,' +
        'sample_t.weight,user_comment,material,fraction, type, sampling_date, CNIsotopA) VALUES(' +
        IntToStr(project_nr);
      for Acol := 1 to 4 do  //add user_lable, user_label_nr, user_desc1, user_desc2 to the query
        if length(cells[ACol, Arow]) > 0 then s := s + ',' + #34 + ReplaceStr(cells[ACol, Arow],'?','') + #34 //remove '?' from string
        else s := s + ',' + 'NULL';
      if length(cells[5, ARow]) > 0 then
          begin // add weight to the query
            strWeight := cells[5, Arow];
            if FormatSettings.DecimalSeparator = ',' then // deutsches Windows
              if Pos('.', strWeight) > 0 then strWeight := ReplaceStr(strWeight, '.', ',');
            w := StrToFloat(strWeight);
            SaveDecimalSep := FormatSettings.DecimalSeparator;
            FormatSettings.DecimalSeparator := '.';
            s := s + ',' + FloatToStr(w);
            FormatSettings.DecimalSeparator := SaveDecimalSep;
          end
      else
      // no weight given
          begin
            s := s + ',' + 'NULL';
          end;
      s := s + ',' + #34 + cells[7, Arow] + #34; //user_comment
      s := s + ',' + #34 + cells[MaterialCol, Arow] + #34;  //material
      s := s + ',' + #34 + cells[FractionCol, Arow] + #34;  //fraction
      s := s + ',' + #34 + cells[TypeCol, Arow] + #34;  //type
      //add sampling date to the query
      with grdPreviewSamples do
          begin
            if Pos('co2atm', cells[TypeCol, Arow]) > 0 then
              s := s + ',' + #34 + FormatDateTime('YYYY-MM-DD', StartOfTheYear(Date)) + #34
            else
              s := s + ',' + 'NULL';
          end;
      // add CNIsotopA flag to the query
      s := s + ', ' + inttostr(InsertAllCNIsotopA);
      s := s + ');'; //query ends here +++++++++++++++++++++++++++++
      // copy query string to query object
      dm.adoCmd.CommandText := s;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry('### IMPORT: Import Sample data...');
      LogWindow.addLogEntry(s);
      // excute query
      IF dm.adoConnKTL.Connected THEN
          Begin
            Try
              //showmessage(s);
              dm.adoCmd.Execute;
              LogWindow.addLogEntry('executed');
              LogWindow.addLogEntry('### IMPORT: Import Sample data...Done');
            Except
              ShowMessage('problem opening the database during import: sample data');
              JvLogFile.Add('DB',lesError,'Problem opening DB during import: sample data');
            End;
          End;
      dm.qryDB.Close;

      // +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      //get the sample_nr of the just inserted sample = is max(sample_nr)
      dm.qryDb.SQL.Text := 'SELECT Max(sample_nr) FROM sample_t;';
      LogWindow.addLogEntry(dm.qryDB.SQL.Text);
      IF dm.adoConnKTL.Connected THEN
          Begin
            Try
              dm.qryDb.Open;
              LogWindow.addLogEntry('executed');
            Except
              ShowMessage('problem opening the database reading max sample number');
              JvLogFile.Add('DB',lesError,'Problem opening DB reading max sample number');
            End;
          End;
      sample_nr := 0;

      // +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      // insert new prep for this samples (qryDB stores the max(sample_nr)
      // assign new sample_nr to the variable from the database
      if dm.qryDB.RecordCount > 0 then sample_nr := dm.qryDb.Fields[0].AsInteger;
      with grdPreviewSamples do
      begin
        if (cells[SamplePrep1Col, ARow] = 'none') or (Pos('collagen', cells[MaterialCol, Arow]) > 0) or
          (Pos('graphite', cells[MaterialCol, Arow]) > 0) then
          begin
          LogWindow.addLogEntry('new prep for MAMS-' + IntToStr(sample_nr) + 'doesnt need prep_method');
          s := 'INSERT INTO preparation_t (prep_nr, sample_nr,step1_method, prep_end) VALUES(1,' +
            IntToStr(sample_nr) + ',' + #34 + 'none' + #34 + ',' +
            #34 + FormatDateTime('YYYY-MM-DD', DateOf(date)) + #34;
          end
        else
          begin // set prep_end date since sample type doesn't need any prep steps
          LogWindow.addLogEntry('new prep for MAMS-' + IntToStr(sample_nr) + 'needs prep_methods');
            s := 'INSERT INTO preparation_t (prep_nr, sample_nr,step1_method,step2_method,' +
              'step3_method,step4_method,step5_method) VALUES(1,' + IntToStr(sample_nr);
            for PrepCol := SamplePrep1Col to SamplePrep5Col do
                begin
                  if Length(cells[PrepCol, aRow]) > 0 then
                    s := s + ',' + #34 + cells[PrepCol, aRow] + #34
                  else
                    s := s + ',NULL';
                end;
          end;
        dm.adoCmd.CommandText := s + ');';
//        ClipBoard.SetTextBuf(PChar(s));
        LogWindow.addLogEntry('### IMPORT: create new prep for MAMS-' + IntToStr(sample_nr) + ' ...');
        LogWindow.addLogEntry(dm.adoCmd.CommandText);
        IF dm.adoConnKTL.Connected THEN
            Begin
              Try
                dm.adoCmd.Execute;
                LogWindow.addLogEntry('executed');
                LogWindow.addLogEntry('### IMPORT: create new prep for MAMS-' + IntToStr(sample_nr) + ' ... Done');
              Except
                ShowMessage('problem opening the database creating new prep for MAMS-' + IntToStr(sample_nr));
                JvLogFile.Add('DB',lesError,'Problem opening DB creating new prep for MAMS-' + IntToStr(sample_nr));
              End;
            End;

        // +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // insert new target for this sample
        s := 'INSERT INTO target_t  (target_nr, prep_nr, sample_nr, graphitized)'
          + ' VALUES (1,1,' + IntTostr(sample_nr);
        with grdPreviewSamples do
            begin
              if Pos('graphite', cells[MaterialCol, Arow]) > 0 then
                s := s + ',' + #34 + FormatDateTime('YYYY-MM-DD', DateOf(Date)) + #34
              else
                s := s + ',' + 'NULL';
            end;
        LogWindow.addLogEntry('### IMPORT: create new target for MAMS-' + IntToStr(sample_nr) + ' ...');
        dm.adoCmd.CommandText := s + ');';
        s := dm.adoCmd.CommandText;
        //   ClibBoard.SetTextBuf(PChar(s));
        LogWindow.addLogEntry(dm.adoCmd.CommandText);
        IF dm.adoConnKTL.Connected THEN
            Begin
              Try
                dm.adoCmd.Execute;
                LogWindow.addLogEntry('executed');
                LogWindow.addLogEntry('### IMPORT: create new target for MAMS-' + IntToStr(sample_nr) + ' ... Done');
              Except
                ShowMessage('problem opening the database creating new target for MAMS-' + IntToStr(sample_nr));
                JvLogFile.Add('DB',lesError,'Problem opening DB creating new target for MAMS-' + IntToStr(sample_nr));
              End;
            End;
      end;
      // display sample number in dialog for this sample
      lbWizFinalPage.lines.Add('new sample added to database - MAMS: ' + inttostr(sample_nr));
    end;
  wizFinalPage.EnableButton(bkFinish, false);
  end;

  // query database again for user information
  // this is necessay in case a user is already known
    with dm.qryDB do
      begin
      SQL.Text := 'SELECT last_name, first_name, email FROM user_t where user_nr=' + IntTostr(user_nr);
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database reading user info');
          JvLogFile.Add('DB',lesError,'Problem opening DB reading user info');
        End;
      End;
      end;

  lbWizFinalPage.lines.add(IntToStr(grdPreviewSamples.RowCount - 1) + ' new samples added to database');
  wizFinalPage.EnableButton(bkFinish, false);

//ask whether to send a confirmation email
//if so the email page will be opened
  edtMailSubject.Clear;
  MailMemo.Lines.Clear;
  edtMailTo.Clear;

  buttonDlg := messagedlg('Send confirmation email?',mtError, mbOKCancel, 0);
  if buttonDlg = mrOK then
  begin
    edtMailSubject.Text := 'C14 sample receipt';
    edtMailTo.Text := dm.qrydb.FieldByName('email').AsString;  // used to be grdPreviewUser.Cells[1, 13];  only works for a new user
  with MailMemo.Lines do
    begin
      Add('Sehr geehrter Herr XXX');
      Add('hiermit bestätigen wir den Erhalt ihrer C14 Proben.');
      Add(' ');
      Add('C14 Sample receipt:');
      Add('We confirm to have received the samples listed below.');
      Add(' ');
      Add('Best regards');
      Add('The Team of the CEZA C14-Lab');
      Add('This is an auto-generated email.');
      Add(' ');
      Add(' ');
      Add('Received Samples:');
    end;
  with dm.qryDB do
    begin
    SQL.Text := 'SELECT sample_nr, user_label, user_label_nr, user_desc1, user_desc2 from sample_t ' +
      ' WHERE project_nr=' + IntToStr(project_nr);
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database reading sample info');
          JvLogFile.Add('DB',lesError,'Problem opening DB reading sample info');
        End;
      End;
    if RecordCount > 0 then
      begin
      First;
      while not EOF do
        begin
        s := 'MAMS ' + Fields.Fields[0].AsString + sLineBreak; // sample_nr
        for i := 1 to 4 do s := s + '  ' + Fields.Fields[i].AsString;  //add user_label etc etc
        s := s + sLineBreak;  // add a line break after each sample
        MailMemo.Lines.Add(s);
        Next;
        end;
      end;
    end;
  //open tabsheet SendMail
  pgtMain.ActivePage := tbsSendMail;
  end;
  //  actSendMailExecute(self);

end;

procedure TfrmMAMS.FillCds;
// file dindices may have to be adjusted
var
  i, a, b: integer;
  f: double;
  s: string;
  Rounding: boolean;
begin
  cdsExport.Open;
  with dm.qrySampleOfSubmitter do
  begin
    First;
    while not EOF do
    begin
      with cdsExport do
      begin
        DisableControls;
        cdsExport.Edit;
        cdsExport.Append;
{                   0           1           2             3
      s := 'SELECT sample_nr, user_label,user_label_nr,user_desc1,' +
             4          5         6           7
        'user_desc2, C14_age, C14_age_sig, av_dc13, ' +
           8           9     10       11           12              13       14
      ' Cal1sMin, Cal1sMax,Cal2sMin,Cal2sMax, project_t.project, av_fm, av_fm_sig ' + 'FROM sample_t ' +
}
        cdsExport.Fields.Fields[0].Value := dm.qrySampleOfSubmitter.Fields.Fields[0].AsInteger; // sample_nr
        for i := 1 to 4 do cdsExport.Fields.Fields[i].Value := dm.qrySampleOfSubmitter.Fields.Fields[i].AsString;
        a := dm.qrySampleOfSubmitter.Fields.Fields[5].AsInteger;      //was 5
        b := dm.qrySampleOfSubmitter.Fields.Fields[6].AsInteger;      //was 6
        if a > 12000 then
          Rounding := true
        else
          Rounding := false;
        if Rounding then
        begin
          cdsExport.Fields.Fields[5].Value := (a div 10) * 10;     //was 5
          cdsExport.Fields.Fields[6].Value := ((b div 10) + 1) * 10;   //was 6
        end
        else
        begin
          cdsExport.Fields.Fields[5].Value := a;    //was 5
          cdsExport.Fields.Fields[6].Value := b;    //was 6
        end;
        for i := 10 to 12 do
        begin       //was 10 to 12
          s := Format('%4.1f', [dm.qrySampleOfSubmitter.Fields.Fields[i + 5].AsFloat]);
          cdsExport.Fields.Fields[i].Value := s;
        end;
        for i := 13 to 14 do  //for av_fm and av_fm_sig
        begin
          s := Format('%5.3f', [dm.qrySampleOfSubmitter.Fields.Fields[i].AsFloat]);
          cdsExport.Fields.Fields[i].Value := s;
        end;

        //this is for material
        s := dm.qrySampleOfSubmitter.Fields.Fields[19].AsString; //column 19 in this table
        cdsExport.Fields.Fields[15].Value := s;                  //column 15 in the in-memory-table

        //for pmC and pmC_err (both are calculated fields and not in the database)
        cdsExport.Fields.Fields[16].Value := Format('%5.2f', [dm.qrySampleOfSubmitter.Fields.Fields[28].AsFloat]);
        cdsExport.Fields.Fields[17].Value := Format('%5.2f', [dm.qrySampleOfSubmitter.Fields.Fields[29].AsFloat]);

//        clData.Fields.Fields[0] as TDateTimeField).DisplayFormat := 'tt';
//        f := dm.qrySampleOfSubmitter.Fields.Fields[7].AsFloat * 1000;
        f := dm.qrySampleOfSubmitter.Fields.Fields[7].AsFloat;
//        cdsExport.Fields.Fields[7].Value := dm.qrySampleOfSubmitter.Fields.Fields[7].AsFloat;   //dC13
        cdsExport.Fields.Fields[7].Value := Format('%4.1f', [f]);
//        dm.qrySampleOfSubmitter.Fields.Fields[7].Disp;   //dC13
//        cdsExport.Fields.Fields[8].Value := dm.qrySampleOfSubmitter.Fields.Fields[8].AsFloat*1000;
        s := '';
        a := dm.qrySampleOfSubmitter.Fields.Fields[8].AsInteger;   //was 8
        b := dm.qrySampleOfSubmitter.Fields.Fields[9].AsInteger;   //was 9
        if Rounding then
        begin
          a := (a div 10) * 10; // rounding
          b := ((b div 10) + 1) * 10;
        end;
        if chkCalBP.Checked then
        begin
          s := 'cal BP ' + IntToStr(a + 1950);
          s := s + '-' + IntToStr(b + 1950);
          cdsExport.Fields.Fields[8].Value := s;
          s := '';
          a := dm.qrySampleOfSubmitter.Fields.Fields[10].AsInteger;   //was 10
          b := dm.qrySampleOfSubmitter.Fields.Fields[11].AsInteger;   //was 11
          if Rounding then
          begin
            a := (a div 10) * 10; // rounding
            b := ((b div 10) + 1) * 10;
          end;
          s := 'cal BP ' + IntToStr(a + 1950);
          s := s + '-' + IntToStr(b + 1950);
          cdsExport.Fields.Fields[9].Value := s;
        end
        else
        begin
          if a < 0 then
            s := 'cal AD ' + IntToStr(-a)
          else
            s := 'cal BC ' + IntToStr(a);
          i := 9;
          a := dm.qrySampleOfSubmitter.Fields.Fields[i].AsInteger;
          if a < 0 then
            if Pos('BC', s) > 0 then
              s := s + '-cal AD ' + IntToStr(-a)
            else
              s := s + '-' + IntToStr(-a)
          else
            s := s + '-' + IntToStr(a);
          cdsExport.Fields.Fields[8].Value := s;
          i := 10;
          s := '';
          a := dm.qrySampleOfSubmitter.Fields.Fields[i].AsInteger;
          if a < 0 then
            s := 'cal AD ' + IntToStr(-a)
          else
            s := 'cal BC' + IntToStr(a);
          i := 11;
          a := dm.qrySampleOfSubmitter.Fields.Fields[i].AsInteger;
          if a < 0 then
            if Pos('BC', s) > 0 then
              s := s + '-cal AD ' + IntToStr(-a)
            else
              s := s + '-' + IntToStr(-a)
          else
            s := s + '-' + IntToStr(a);
          cdsExport.Fields.Fields[9].Value := s;
        end;
      end;
      cdsExport.Post;
      dm.qrySampleOfSubmitter.Next;
    end;
    cdsExport.EnableControls;
    cdsExport.Close;
    cdsExport.Open;
  end;
end;

procedure TfrmMAMS.FillFractionBox;
begin
  lbxFraction.Clear;
  lbxFraction.Items.Add('Fraction :');
  with dm.tblFraction do begin
    First;
    while not EOF do begin
      lbxFraction.Items.Add(Fields.FieldByName('fraction').AsString);
      Next;
    end;
  end;
end;

procedure TfrmMAMS.FillMaterialBox;
begin
  lbxMaterial.Clear;
  lbxMaterial.Items.Add('Material :');
  with dm.tblMaterial do begin
    First;
    while not EOF do begin
      lbxMaterial.Items.Add(Fields.FieldByName('material').AsString);
      Next;
    end;
  end;
end;

procedure TfrmMAMS.FillPrepBox;
begin
  lbxDefinePrepSteps.Clear;
  with dm.tblMethod do begin
    First;
    while not EOF do begin
      lbxDefinePrepSteps.Items.Add(Fields.FieldByName('method').AsString);
      Next;
    end;
  end;
end;

procedure TfrmMAMS.FillPrepList;
begin
  lbxSteps.Clear;
  with dm.tblMethod do begin
    First;
    while not EOF do begin
      lbxSteps.Items.Add(Fields.FieldByName('method').AsString);
      Next;
    end;
  end;
end;

procedure TfrmMAMS.FillTypeBox;
begin
  lbxTypes.Clear;
  with dm.tblTypes do
  begin
    First;
    while not EOF do
    begin
      lbxTypes.Items.Add(Fields.FieldByName('type').AsString);
      Next;
    end;
  end;
end;

procedure TfrmMAMS.FormCreate(Sender: TObject);
var
  bmp: TBitmap;
  pathToIni: string;
  txtFile: textFile;
begin
  FCheck := TBitmap.Create;
  FNoCheck := TBitmap.Create;
  bmp := TBitmap.create;
  try
    bmp.handle := LoadBitmap(0, PChar(OBM_CHECKBOXES));
    // bmp now has a 4x3 bitmap of divers state images
    // used by checkboxes and radiobuttons
    with FNoCheck do
    begin
      // the first subimage is the unchecked box
      width := bmp.width div 4;
      height := bmp.height div 3;
      canvas.copyrect(canvas.cliprect, bmp.canvas, canvas.cliprect);
    end;
    with FCheck do
    begin
      // the second subimage is the checked box
      width := bmp.width div 4;
      height := bmp.height div 3;
      canvas.copyrect(
        canvas.cliprect,
        bmp.canvas,
        rect(width, 0, 2 * width, height));
    end;
  finally
    bmp.free
  end;

  // create iniFile Object for AutoPrep (assign prep according to material)
  // and connect to the ini file that SAMS uses
  pathToIni := extractFilePath(paramstr(0)) + 'autoprep.ini';
  myPrepIni := TiniFile.Create(pathToIni);
  if not fileexists(pathToIni) then
    Begin
      // create new empty text file
      AssignFile(txtFile,pathToini);
      Rewrite(txtFile);
      CloseFile(txtFile);
    End;

 // display the program Version in the name of the main aplication window
 frmMAMS.Caption:= 'SAMS' + myVersion;
end;

procedure TfrmMAMS.FormDestroy(Sender: TObject);
begin
   JvLogFile.Add('Closeup',lesInformation,'SAMS closed by user.');
  FNoCheck.Free;
  FCheck.Free;
end;

procedure TfrmMAMS.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
// delas with key strokes that happen in the whole program
begin
  if Shift = [ssShift, ssCtrl] then
    if Key = 75 then //^K
      if (MessageDlg('Sample will be deleted. Are you sure?', mtWarning, [mbNo, mbOK], 0) = mrOk) then
        dm.DeleteFromDb(round(edtSampleNr.Value));

  // if in tbsSAmpleInfo and ESC set focus to edtSampleNr
  if pgtMain.ActivePage = tbsSampleInfo then
  begin
    if Key = vk_ESCAPE then edtSampleNr.SetFocus;
  end;

    // if in tbsSAmpleInfo and + increase SampleNr
  if pgtMain.ActivePage = tbsSampleInfo then
  begin
    if Key = vk_ADD then
    begin
          edtSampleNr.Value:= round(edtSampleNr.Value)+1;
          DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    end;
  end;

    // if in tbsSAmpleInfo and - degrease SampleNr
  if pgtMain.ActivePage = tbsSampleInfo then
  begin
    if Key = vk_SUBTRACT then
    begin
          edtSampleNr.Value:= round(edtSampleNr.Value)-1;
          DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    end;
  end;

  // if in tbsSAmpleInfo and CTRL S = save Weights???????

end;

procedure TfrmMAMS.grdActiveBatchesCellClick(Column: TColumn);
var
  s: string;
  sample_nr, steps: integer;
  method: string;
begin
  if rgpLabTask.ItemIndex = 0 then
  begin
    lbxPrepSteps.Clear;
    with dm.qrySamplesOfLabTask do
    begin
      SQL.Text := 'SELECT sample_t.sample_nr, user_label, user_label_nr  '
        + 'FROM sample_t '
        + ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr '
        + 'WHERE batch=' + #34 + dm.qryActiveBatches.FieldByName('batch').AsString + #34 + ';'; ;
      s := SQL.Text;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;
    grdSamplesOfLabTask.Visible := true;
//  grdSamplesOfLabTask.Top := 130;
    with dm.qryDb1 do
    begin
      sample_nr := dm.qrySamplesOfLabTask.FieldByName('sample_nr').AsInteger;
      SQL.Text := 'SELECT step1_method, step2_method, step3_method,step4_method,' +
        'step5_method FROM preparation_t WHERE sample_nr=' + IntToStr(sample_nr);
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End; // muss noch um mehrere preps erweitert werden
      for steps := 0 to 4 do
      begin
        method := dm.qryDb1.Fields.Fields[steps].AsString;
        if length(method) > 0 then lbxPrepSteps.Items.Add(method);
      end;
    end;
    if lbxPrepSteps.Items.Count = 1 then lbxPrepSteps.Selected[0] := true;
{    grdSamplesOfProject.Columns[0].Width := 60;
    grdSamplesOfProject.Columns[1].Width := 100;
    grdSamplesOfProject.Columns[2].Width := 80;
    grdSamplesOfProject.Columns[3].Width := 80;
    grdSamplesOfProject.Columns[4].Width := 80;
  end;
  grdSamplesOfProject.Visible := true;

}
  end
  else
  begin
    lbxPrepSteps.Clear;
    with dm.qrySamplesOfLabTask do
    begin
      SQL.Text := 'SELECT sample_t.sample_nr, user_label, user_label_nr  '
        + 'FROM sample_t '
        + ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr '
        + 'WHERE graph_batch=' + #34 + dm.qryActiveBatches.FieldByName('graph_batch').AsString + #34 + ';'; ;
      s := SQL.Text;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;
    grdSamplesOfLabTask.Visible := true;
//  grdSamplesOfLabTask.Top := 130;
{    with dm.qryDb1 do begin
      sample_nr := dm.qrySamplesOfLabTask.FieldByName('sample_nr').AsInteger;
      SQL.Text := 'SELECT step1_methode, step2_methode, step3_methode,step4_methode,' +
        'step5_methode FROM preparation_t WHERE sample_nr=' + IntToStr(sample_nr);
      Open; // muss noch um mehrere preps erweitert werden
      for steps := 0 to 4 do begin
        method := dm.qryDb1.Fields.Fields[steps].AsString;
        if length(method) > 0 then lbxPrepSteps.Items.Add(method);
      end;
    end;
    if lbxPrepSteps.Items.Count = 1 then lbxPrepSteps.Selected[0] := true;
    }
  end;
end;

procedure TfrmMAMS.grdFractionTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdFraction.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdFraction.DataSource.DataSet) do
    begin
      try
        grdFraction.Columns[PreviousColumnIndex].title.Font.Style :=
          grdFraction.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)  //test
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdInPrepDblClick(Sender: TObject);
begin
 ShowSampleInfoPage(Sender);
end;

procedure TfrmMAMS.grdInPrepDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);

  CONST CtrlState: array[0..1] of integer = (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED) ;
begin
  AlternateRowColors(Sender, State);

  if dm.qryInPrep.FieldByName('desired_date').AsDateTime<=Date()+21 Then     //if desired_date approcahes 2 weeks
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clFuchsia;
     End;
  if dm.qryInPrep.FieldByName('desired_date').AsDateTime<=Date() Then     //if desired_date has run out color text red
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clRed;
     End;

  TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);

    // display boolean fileds as checkboxes in the infreeze column
  if (Column.FieldName='InFreeze') then
  begin
  //LogWindow.addLogEntry('inside if');
    TDBGrid(Sender).Canvas.FillRect(Rect) ;
    if VarIsNull(Column.Field.Value) then
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, DFCS_BUTTONCHECK or DFCS_INACTIVE) {grayed}
    else
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, CtrlState[Column.Field.AsInteger]) ; {checked or unchecked}
  end;
end;

procedure TfrmMAMS.grdInPrepTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdPlanned.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdInPrep.DataSource.DataSet) do
    begin
      try
        grdPlanned.Columns[PreviousColumnIndex].title.Font.Style :=
          grdPlanned.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdMagazineDataCellClick(Column: TColumn);
var
  i, j, m, n: integer;
  s, snr: string;
  f: double;
  c: char;
begin
  edtMeanAge.Clear;
  edtSigmaAge.Clear;
  edtC13Mean.Clear;
  btnTransferTargetData.Enabled := false;
  i := grdMagazineData.DataSource.DataSet.FieldByName('target_nr').AsInteger;
  n := grdMagazineData.DataSource.DataSet.FieldByName('prep_nr').AsInteger;
  if (i > 1) or (n > 1) then
  begin
    with dm.qryTargetData do
    begin // get user_nr from user_t
      Close;
      snr := IntToStr(grdMagazineData.DataSource.DataSet.FieldByName('sample_nr').AsInteger);
      SQL.Text := 'SELECT sample_nr, prep_nr, target_nr, magazine, c14_age, c14_age_sig, dc13' +
        ' FROM target_t' +
        ' WHERE sample_nr=' + snr + ';';
      s := SQL.Text;
//      ClipBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      if RecordCount > 0 then
      begin
        j := RecordCount;
        StrGrdTargetData.Clear;
        with StrGrdTargetData {, grdTargetData} do
        begin
          ColCount := 8;
          RowCount := j + 1;
          ColWidths[0] := 60; //sample
          ColWidths[1] := 40; //prep
          ColWidths[2] := 50; //tar
          ColWidths[3] := 60; //mag
          ColWidths[4] := 100; //c14age
          ColWidths[5] := 100; //c14age_sig
          ColWidths[6] := 90; //d13c
          ColWidths[7] := 30;
          for m := 0 to FieldCount - 1 do
          begin
            Cells[m, 0] := Fields[m].DisplayName;
          end;
          First;
          while not Eof do
          begin
            for m := 0 to 3 do
            begin
              s := Fields[m].AsString;
              Cells[m, RecNo] := s;
            end;
            s := cells[0, 1];
            for m := 4 to 5 do
            begin
              s := Fields[m].AsString;
              if Length(s) > 0 then
              begin
                f := StrToFloat(s);
                Cells[m, RecNo] := Format('%8.0f', [f]);
              end;
            end;
            for m := 6 to 6 do
            begin
              s := Fields[m].AsString;
              if Length(s) > 0 then
              begin
                f := StrToFloat(s);
                Cells[m, RecNo] := Format('%6.4f', [f]);
              end;
            end;
            Next;
          end;
          for m := 1 to RowCount - 1 do
          begin
            if (Length(Cells[4, m]) <> 0) then //hat c14age
              ToggleCheckbox(7, m);
          end;
        end;
      end
      else begin
        ShowMessage('RecordCount < 0. Exit. Sample_nr=' + snr + '.');
        Exit;
      end;
    end;
  end
  else begin
    StrGrdTargetData.Clear;
  end;
end;

procedure TfrmMAMS.grdMagazineDataDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  i, j, m: integer;
begin
  with (Sender as TDBGrid) do
  begin
    j := (DataSource.DataSet.FieldByName('target_nr').AsInteger);
    m := (DataSource.DataSet.FieldByName('prep_nr').AsInteger);
    // prüfen ob in der Wert im Feld Farbcode <>NULL ist
    // Wenn ja, dann Farbe wählen. Wenn nicht, dann weiß.
    // Die Prüfung ist bei dem o.a. SQL Befehl nicht unbedingt notwendig, da Farbcode nie NULL werden kann
    // aber falls man mal auf reale Spalten zurückgreift muss er rein ...

    if (j > 1) or (m > 1) then
    begin
      Canvas.Brush.Color := clYellow; // << hier an deine Query anpassen ...
    end
    else begin
      Canvas.Brush.Color := clWhite;
    end;

    for i := 0 to grdMagazineData.FieldCount - 1 do
    begin
      if (SampleTrans[i] = true) and not ((j > 1) or (m > 1)) then
      begin
        Canvas.Brush.Color := clLime; // << hier an deine Query anpassen ...
      end;
    end;

    // Bei selektierten Feldern die Farbe nochmal ändern
    if (gdSelected in state) then
    begin
      Canvas.Brush.Color := clHighlight; //Farbe für die Zelle mit dem Focus
    end;
    // Zeichnen
    Canvas.FillRect(Rect); //Hintergrundfarbe zeichnen
    Canvas.TextOut(Rect.Left + 2, Rect.Top + 1, Column.Field.AsString); //Den Text in der Zelle ausgeben
  end;
end;

procedure TfrmMAMS.grdMagazinesCellClick(Column: TColumn);
var
  Mag, sample_nr, prep_nr, target_nr, le_current, he_current, runIDprefix, dbstring: string;
  i, w, j: integer;
  DataSrcLeHeCurrent, DataSrcRunIDPrefix: TDataSource;
  ADOQueryLeHeCurrent, ADOQueryRunIDPrefix: TADOQuery;
begin
  edtMeanAge.Clear;
  edtSigmaAge.Clear;
  edtC13Mean.Clear;
  btnTransferTargetData.Enabled := false;

  // get magazine name and display the name in the transfer button
  Mag := dm.qryMagazines.FieldByName('magazine').AsString;
  tnsto.Caption := 'Transfer ' + Mag;

  // close dataset
  dm.cdsLeHeCurrents.Close;

  // query db_dmams for target data of the selected magazine
  JvLogFile.Add('DB',lesInformation,'Get data of all targets of selected magazine');
  with dm.qryDB do
  begin
    Close;
    // not sure why this is done
    SQL.Text := 'SELECT sample_t.sample_nr, target_t.prep_nr, target_t.target_nr FROM sample_t ' +
      ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr ' +
      ' WHERE target_t.magazine = ' + #34 + Mag + #34;
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;

    // Read data from magazine by magazine name and fill in data into the table
    // also display the RunIDPrefix in order to distinguish between AMS 1 or AMS2
    dm.GetMagazineData(Mag);

    SetLength(SampleTrans, RecordCount + 1);
    for j := 0 to Length(SampleTrans) - 1 do
      SampleTrans[j] := false;
    with grdMagazineData do
    begin
      Columns[0].Width := 40;
      Columns[1].Width := 140;
      Columns[2].Width := 120;
      Columns[3].Width := 55;
      Columns[4].Width := 55;
      Columns[5].Width := 55;
      Columns[6].Width := 60;
      // compute the size of the groupbox by adding up the coulmns and otehr stuff
      // w := 0;
      // for i := 0 to 6 do w := w + Columns[i].Width + 1; // 1 = gridline
      // gbxMagazines.Width := grdMagazines.Width + 6 + w + 6 + 20 + 20; // adjust size of the groupbox, 20=scrollbar
    end;
  end;

  // query db_dc14 for LE and HE Currents of the samples of the selected magazine
  // use qryDB for list of sample, prep and target numbers of the selected magazin since db_dc14 doesn't seem to store the magazine names
  if dm.qryMagazineData.RecordCount>0 then
  begin
      LogWindow.addLogEntry('query raw DB (db_dc14 or db_ac14) for LE and HE Currents');
      JvLogFile.Add('DB',lesInformation,'query raw DB (db_dc14 or db_ac14) for LE and HE Currents');

      // create a query object for querying the db_dc14 or db_ac14 Database
      ADOQueryLeHeCurrent := TADOQuery.Create(Self);
      ADOQueryLeHeCurrent.Connection := dm.ADOConnKTL;
      LogWindow.addLogEntry('QueryObject created');
      //ADOQueryLeHeCurrent.SQL.Add(s);
      //ADOQueryLeHeCurrent.Open;
      // Create the data source.
      DataSrcLeHeCurrent := TDataSource.Create(Self);
      DataSrcLeHeCurrent.DataSet := ADOQueryLeHeCurrent;
      DataSrcLeHeCurrent.Enabled := true;
      LogWindow.addLogEntry('DataSource created');

      //dm.cdsLeHeCurrents.Open;
      dm.cdsLeHeCurrents.CreateDataSet;
      LogWindow.addLogEntry('ClientDataSet created');

      dm.qryMagazineData.First;  // set pointer to first entry
      i := 0;
      repeat           // go through qryDB and retrieve sample, prep and target numbers
          LogWindow.addLogEntry('reading data from magazine from position ' + inttostr(i));

          sample_nr := dm.qryMagazineData.FieldByName('sample_nr').AsString;
          prep_nr := dm.qryMagazineData.FieldByName('prep_nr').AsString;
          target_nr := dm.qryMagazineData.FieldByName('target_nr').AsString;
          runIDprefix := dm.qryMagazineData.FieldByName('RunIDPrefix').AsString; // this shows where the data were measured AMS1 = db_dc14, AMS2 = db_ac14

          LogWindow.addLogEntry('sample_nr = ' + sample_nr);
          LogWindow.addLogEntry('prep_nr = ' + prep_nr);
          LogWindow.addLogEntry('target_nr = ' + target_nr);
          LogWindow.addLogEntry('RunIDPrefix = ' + runIDprefix);

          // using the RunIDPrefix, decided what database to query.
          // DC = db_dc14, AC = db_ac14
          dbstring := '';
          if runIDPrefix = 'AC' then dbstring := 'db_ac14';
          if runIDPrefix = 'DC' then dbstring := 'db_dc14';
          LogWindow.addLogEntry('raw database to use = ' + dbstring);
          JvLogFile.Add('DB',lesInformation,'raw database to use = ' + dbstring);

          if not (dbstring = '') then
            Begin
              // query db_dc14 for a single sample
              ADOQueryLeHeCurrent.SQL.Text := 'select sample_nr, prep_nr, target_nr, round(ANA,2) AS ANA, round(A,2) AS A from ' + dbstring + '.workfinal ' +
                   'where sample_nr = ' + sample_nr + ' ' +
                   'AND prep_nr = ' + prep_nr + ' ' +
                   'AND target_nr = ' + target_nr + ';';
              LogWindow.addLogEntry(ADOQueryLeHeCurrent.SQL.Text);

              // query
              Try
                ADOQueryLeHeCurrent.Open;
              Except
                LogWindow.addLogEntry('Problem querying ' + dbstring);
                JvLogFile.Add('DB',lesError,'Problem querying '  + dbstring);
              End;

              //read LE and HE currents from query
              le_current := ADOQueryLeHeCurrent.FieldByName('ANA').AsString;
              LogWindow.addLogEntry('LECurrent= ' + le_current);
              he_current := ADOQueryLeHeCurrent.FieldByName('A').AsString;
              ADOQueryLeHeCurrent.Close;

              // write data into clientdataset, clientdataset is connected to a grid in the UI
              LogWindow.addLogEntry('fill string grid with data from query');
              JvLogFile.Add('DB',lesInformation,'fill string grid with data from query');
            End
            Else
              Begin
               LogWindow.addLogEntry('RunIDPrefix is empty. No raw database assigned. DB not queried.');
              JvLogFile.Add('DB',lesInformation,'RunIDPrefix is empty. No raw database assigned. DB not queried.');
              End;

          // write the data into the table
          dm.cdsLeHeCurrents.Append;
          dm.cdsLeHeCurrents.FieldByName('sample_nr').AsString := sample_nr;
          dm.cdsLeHeCurrents.FieldByName('prep_nr').AsString := prep_nr;
          dm.cdsLeHeCurrents.FieldByName('target_nr').AsString := target_nr;
          dm.cdsLeHeCurrents.FieldByName('LECurrent').AsString := le_current;
          dm.cdsLeHeCurrents.FieldByName('HECurrent').AsString := he_current;
          dm.cdsLeHeCurrents.Post;

          // go to next target in magazine
          dm.qryMagazineData.Next;
      until dm.qryMagazineData.Eof;

      dbgrdLeHeCurrents.Columns[0].Width := 60;
      dbgrdLeHeCurrents.Columns[1].Width := 40;
      dbgrdLeHeCurrents.Columns[2].Width := 40;
      dbgrdLeHeCurrents.Columns[3].Width := 50;
      dbgrdLeHeCurrents.Columns[4].Width := 50;
  end;

end;

procedure TfrmMAMS.grdMagazinesUnpressedCellClick(Column: TColumn);
var
  mag : string;
begin
  // get magazine name from the selected cell
  mag := dm.qryMagazinesUnpressed.FieldByName('magazine').AsString;

  // get targets of this magazine and fill grid
  if not (mag = '') then  dm.GetMagazineDataUnpressed(mag);

  // set columns
  //DBGridAutoSizeAllColumns(DBGridSamplesOfUnpressedMagazine);
   with DBGridSamplesOfUnpressedMagazine do
    begin
      Columns[0].Width := 40; // position
      Columns[1].Width := 60; // target_id
      Columns[2].Width := 65; // type
      Columns[3].Width := 150; // user_label
      Columns[4].Width := 70; // last_name
      Columns[5].Width := 50; // co2_final
    end;
end;

procedure TfrmMAMS.grdMaterialTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdMaterial.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdMaterial.DataSource.DataSet) do
    begin
      try
        grdMaterial.Columns[PreviousColumnIndex].title.Font.Style :=
          grdMaterial.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdPendingReportsDblClick(Sender: TObject);
var
  Column: TColumn; //this is used because the 'OnClick' method needs this parameter
begin
  pgtMain.ActivePage:=tbsProjectsOfUser; //change to project info page
  cmbSubmitterNameProject.KeyValue:=grdPendingReports.DataSource.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list to correct user_nr
  cmbSubmitterNameProjectCloseUp(self); //correct user is selected now show their projects
  grdProjects.DataSource.DataSet.Locate('project_nr',grdPendingReports.DataSource.DataSet.FieldByName('project_nr').AsInteger,[loPartialKey]);//now the projects are being displayed, select the correct project
  grdProjectsCellClick(Column); //now display the samples that belong to the project
end;

procedure TfrmMAMS.grdPendingReportsDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
// when tables is drawn, change the color of the cells depending on the cell content
// fill cell 'sample count' red, when value >0
begin
  AlternateRowColors(Sender, State);
  if dm.qryPendingReports.FieldByName('measured').AsInteger>0 Then     //if samplecount is bigger than 0 change the color of the row
     Begin
     TDBGrid(Sender).Canvas.Brush.Color:=clSkyBlue;
     End;
  if dm.qryPendingReports.FieldByName('desired_date').AsDateTime<=Date()+14 Then     //if desired_date approaches 2 weeks
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clFuchsia;
     End;
  if dm.qryPendingReports.FieldByName('desired_date').AsDateTime<=Date() Then     //if desired_date has run out color text red
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clRed;
     End;


  TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdPendingReportsTitleClick(Column: TColumn);
//sort by selected column
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdTypes.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdPendingReports.DataSource.DataSet) do
    begin
      try
        grdTypes.Columns[PreviousColumnIndex].title.Font.Style :=
          grdTypes.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdPlannedCellClick(Column: TColumn);
var
  sample_nr: integer;
begin
  sample_nr := dm.dsPlanned.DataSet.FieldByName('sample_nr').AsInteger;
  edtSampleNr.Text := IntToStr(sample_nr);
end;

procedure TfrmMAMS.grdPlannedDblClick(Sender: TObject);
begin
 ShowSampleInfoPage(Sender);
end;

procedure TfrmMAMS.grdPlannedDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

  AlternateRowColors(Sender, State);
  if (POS('Eil',dm.qryPlanned.FieldByName('user_label').AsString)>0)
      OR (POS('EIL',dm.qryPlanned.FieldByName('user_label').AsString)>0)
      OR (POS('Blitz',dm.qryPlanned.FieldByName('user_label').AsString)>0) Then
      Begin
      TDBGrid(Sender).Canvas.Brush.Color:=clRed;
      End;      //if a certain substring is in user_label then change row color

  if dm.qryPlanned.FieldByName('desired_date').AsDateTime<=Date()+14 Then     //if desired_date approaches 2 weeks
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clOlive;
     End;
  if dm.qryPlanned.FieldByName('desired_date').AsDateTime<=Date() Then     //if desired_date has run out color text red
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clRed;
     End;

      TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdPlannedTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdPlanned.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdPlanned.DataSource.DataSet) do
    begin
      try
        grdPlanned.Columns[PreviousColumnIndex].title.Font.Style :=
          grdPlanned.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdPretreatmentCellClick(Column: TColumn);
var value: string;
begin
  // if descr is clicked show a window in order to edit the field (which is a memo field)
  if grdPretreatment.SelectedField = dm.tblMethoddescr then
  begin
    try
      value := InputBox('Test program', 'Please type your town', dm.tblMethoddescr.AsString);
      dm.tblMethod.Edit;
      dm.tblMethoddescr.AsString := value;
    finally

    end;
  end;
//   with TMemoEditorForm.Create(nil) do
//   try
//   DBMemoEditor.Text := dm.tblMethoddescr.AsString;
//   ShowModal;
//   dm.tblMethod.Edit;
//   dm.tblMethoddescr.AsString := DBMemoEditor.Text;
//   finally
//   Free;
//   end;
end;

procedure TfrmMAMS.grdPretreatmentStepsSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
  s: string;
begin
//  grdPretreatmentSteps.Cells[1, ARow] := cmbSelectMethod.ListSource.DataSet.FieldByName(cmbSelectMethod.ListField).AsString;
end;

procedure TfrmMAMS.grdPretreatmentTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdPretreatment.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdPretreatment.DataSource.DataSet) do
    begin
      try
        grdPretreatment.Columns[PreviousColumnIndex].title.Font.Style :=
          grdPretreatment.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdPreviewSamplesDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  c, r, i: Integer;
  MaterialFillDone, TypeFillDone: boolean;
begin
  case DragSource of
    drgMaterial:
    begin
        with grdPreviewSamples do
        begin
          MouseToCell(x, y, c, r);
          if r > 0 then
          begin
            if SetCurrentToAll then
              for i := 1 to RowCount - 1 do Cells[c, i] := CurrentMaterial
            else
              Cells[c, r] := CurrentMaterial;
            MaterialFillDone := true;
            for i := 0 to RowCount - 1 do
              if (length(cells[MaterialCol, i]) = 0) or (length(cells[FractionCol, i]) = 0)
                then MaterialFillDone := false;
          end;
        end;
        if MaterialFillDone then
          wizSelectMaterial.EnableButton(bkNext, true)
        else
          wizSelectMaterial.EnableButton(bkNext, false);
      end;
    drgFraction:
    begin
        with grdPreviewSamples do
        begin
          MouseToCell(x, y, c, r);
          if r > 0 then
          begin
            if SetCurrentToAll then
              for i := 1 to RowCount - 1 do Cells[c, i] := CurrentFraction
            else
              Cells[c, r] := CurrentFraction;
            MaterialFillDone := true;
            for i := 0 to RowCount - 1 do
              if (length(cells[MaterialCol, i]) = 0) or (length(cells[FractionCol, i]) = 0)
                then MaterialFillDone := false;
          end;
        end;
        if MaterialFillDone then
          wizSelectMaterial.EnableButton(bkNext, true)
        else
          wizSelectMaterial.EnableButton(bkNext, false);
      end;
    drgType:
    begin
        with grdPreviewSamples do
        begin
          MouseToCell(x, y, c, r);
          if r > 0 then
          begin
            if SetCurrentToAll then
              for i := 1 to RowCount - 1 do Cells[c, i] := CurrentType
            else
              Cells[c, r] := CurrentType;
            TypeFillDone := true;
            for i := 0 to RowCount - 1 do
              if (length(cells[MaterialCol, i]) = 0) or (length(cells[TypeCol, i]) = 0)
                then TypeFillDone := false;
          end;
        end;
        if TypeFillDone then
          wizSelectType.EnableButton(bkNext, true)
        else
          wizSelectType.EnableButton(bkNext, false);
      end;
    drgPrep:
    begin
        with grdPreviewSamples do
        begin
          MouseToCell(x, y, c, r);
          if r > 0 then
          begin
            if SetCurrentToAll then
              for i := 1 to RowCount - 1 do Cells[c, i] := CurrentPrep
            else
              Cells[c, r] := CurrentPrep;
          end;
        end;
        wizSelectType.EnableButton(bkNext, true)
      end;
  end;
end;

procedure TfrmMAMS.grdPreviewSamplesDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
var
  c, r: Integer;
begin
  grdPreviewSamples.MouseToCell(x, y, c, r);
  case DragSource of
    drgFraction: Accept := (c = FractionCol);
    drgMaterial: Accept := (c = MaterialCol);
    drgType: Accept := (c = TypeCol);
    drgPrep: Accept := (c >= SamplePrep1Col) and (c <= SamplePrep5Col);
  end;
end;

procedure TfrmMAMS.GetSamples;
Var
  s:string;
// performs a query of the database for all samples
// of the selected project in the project list
// query results is displayed in the SampleOfPorject Grid
begin
  with dm.qrySamplesOfProject do
  begin
    SQL.Text := 'SELECT sample_nr, user_label, user_label_nr, user_desc1, user_desc2, ROUND(c14_age) AS age  '
      + 'FROM sample_t '
      + 'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr '
      + 'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr '
      + 'WHERE project_t.project_nr=' + #34 + dm.dsProjectsOfUser.DataSet.FieldByName('project_nr').AsString + #34 + ';'; ;
    s := SQL.Text;
    //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    grdSamplesOfProject.Columns[0].Width := 60;
    grdSamplesOfProject.Columns[1].Width := 100;
    grdSamplesOfProject.Columns[2].Width := 80;
    grdSamplesOfProject.Columns[3].Width := 80;
    grdSamplesOfProject.Columns[4].Width := 80;
    lbl_Project_NumberOfSamples.Caption:='# Samples: ' + grdSamplesOfProject.DataSource.DataSet.RecordCount.ToString;
  end;
end;

procedure TfrmMAMS.grdProjectsCellClick(Column: TColumn);
// clicking on the project grid
// things that happen when selecting a project
var
  s: string;

begin
  grdSamplesOfProject.Visible := true;
  dm.QueryProject(dm.dsProjectsOfUser.DataSet.FieldByName('project').AsString);  // query database for all project info
  GetSamples; // queries the database for all samples of this project and displays them in the SampleOfProject Grid

  (*with dm.qrySamplesOfProject do begin
    SQL.Text := 'SELECT sample_nr, user_label, user_label_nr, user_desc1, user_desc2, ROUND(c14_age) AS age  '
      + 'FROM sample_t '
      + 'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr '
      + 'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr '
      + 'WHERE project_t.project_nr=' + #34 + dm.dsProjectsOfUser.DataSet.FieldByName('project_nr').AsString + #34 + ';'; ;
    s := SQL.Text;
    //   ClibBoard.SetTextBuf(PChar(s));
    Open;
    grdSamplesOfProject.Columns[0].Width := 60;
    grdSamplesOfProject.Columns[1].Width := 100;
    grdSamplesOfProject.Columns[2].Width := 80;
    grdSamplesOfProject.Columns[3].Width := 80;
    grdSamplesOfProject.Columns[4].Width := 80;
  end; *)

  pgtSample.ActivePageIndex:=0;  //show the project_data page (Index=0)
  btnAddNewSamples.Visible:=true;   //show add samples button
  lbl_Project_NumberOfSamples.Visible:=true; // show number of samples

//  spltSampleSamples.Visible := true;
end;

procedure TfrmMAMS.grdProjectsDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var grid: TDBGrid;
begin
  AlternateRowColors(Sender, State);
  (Sender as TDBGrid).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdReportHeadingsExit(Sender: TObject);
begin
  grdReportHeadings.SaveToFile(ExtractFilePath(Application.ExeName) + '\ReportHeadings.txt');
end;


procedure TfrmMAMS.chkAllProjectsClick(Sender: TObject);
begin
//  cmbProjectOfReport.Enabled := not chkAllProjects.Checked;
  if chkAllProjects.Checked then begin
    btnQuerySubmitter.Enabled := true;
    cmbProjectOfReport.Enabled := false;
    cmbProjectOfReport.CloseUp(false);
  end
  else begin
    btnQuerySubmitter.Enabled := true;
    cmbProjectOfReport.Enabled := true;
  end;
end;

procedure TfrmMAMS.chkSamplingThisYearClick(Sender: TObject);
begin
  if chkSamplingThisYear.Checked then
    edtSamplingDate.Date := EncodeDate(YearOf(Date), 1, 1)
  else
    edtSamplingDate.Date := EncodeDate(1950, 1, 1);
end;

procedure TfrmMAMS.chkShowOnHoldClick(Sender: TObject);
begin
  GetPlanned;
end;

procedure TfrmMAMS.chkTargetDiscardedClick(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.chkTouchPrepArchivedMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  dlgvalue: string;
  dlgOK: boolean;
begin
  btnTouchWeightsGraphNeedsSaving.Visible := True;

  // if button was unchecked then the user wants to assign a storage location that wasn't assigned before
  // ask for the value
  // also safe that value to the variable sStorLocSuggestion to keep it in memory for later use
  // if that variable was already set do not ask for the value and use the variable value
  if chkTouchPrepArchived.Checked = False then
    begin
      if prepStorLocSuggestion = '' then
        begin
          //there is no sStorLocSuggestion value given, ask for the value
          dlgOK:= InputQuery('set new storage location for prepped material', 'value', dlgvalue);
                if dlgOK = True then
                  begin
                    prepStorLocSuggestion:= dlgvalue;
                    dbedtprepStorLoc.Text:= dlgvalue;
                  end;
        end
        else
        begin
          // a sStorLocSuggestion value was already given, thenuse this value without asking anymore
          dbedtprepStorLoc.Text:=prepStorLocSuggestion;
        end;
        //chkTouchSampleArchived.Checked := True;
        //if chkTouchSampleArchived.Checked then showmessage('was false, is checked now');
        //if not chkTouchSampleArchived.Checked then showmessage('was false, is NOT checked now');
    end
    else
    begin
      dbedtprepStorLoc.DataSource.DataSet.Edit;
      dbedtprepStorLoc.SetFocus;
      dbedtprepStorLoc.Clear;
      chkTouchPrepArchived.SetFocus;
      chkTouchPrepArchived.Checked := False;
    end;
    chkTouchPrepArchived.SetFocus;

  //SetCheckBoxTouchPrepArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  //DisplayArchiveWarning;
  //ShowMessage(dlgvalue);

end;

procedure TfrmMAMS.chkTouchPrepArchivedMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
     SetCheckBoxTouchPrepArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.chkTouchSampleArchivedMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  dlgvalue: string;
  dlgOK: boolean;
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

  // if button was unchecked then the user wants to assign a storage location that wasn't assigned before
  // ask for the value
  // also safe that value to the variable sStorLocSuggestion to keep it in memory for later use
  // if that variable was already set do not ask for the value and use the variable value
  if chkTouchSampleArchived.Checked = False then
    begin
      if sStorLocSuggestion = '' then
        begin
          //there is no sStorLocSuggestion value given, ask for the value
          dlgOK:= InputQuery('set new storage location for a sample', 'value', dlgvalue);
                if dlgOK = True then
                  begin
                    sStorLocSuggestion:= dlgvalue;
                    dbedtsStorLoc.Text:= dlgvalue;
                  end;
        end
        else
        begin
          // a sStorLocSuggestion value was already given, thenuse this value without asking anymore
          dbedtsStorLoc.Text:=sStorLocSuggestion;
        end;
        //chkTouchSampleArchived.Checked := True;
        //if chkTouchSampleArchived.Checked then showmessage('was false, is checked now');
        //if not chkTouchSampleArchived.Checked then showmessage('was false, is NOT checked now');
    end
    else
    begin
      dbedtsStorLoc.DataSource.DataSet.Edit;
      dbedtsStorLoc.SetFocus;
      dbedtsStorLoc.Clear;
      chkTouchSampleArchived.SetFocus;
      chkTouchSampleArchived.Checked := False;
    end;
    chkTouchSampleArchived.SetFocus;

  //SetCheckBoxTouchSampleArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  //DisplayArchiveWarning;
  //ShowMessage(dlgvalue);

end;

procedure TfrmMAMS.chkTouchSampleArchivedMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
     SetCheckBoxTouchSampleArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarning;
end;

procedure TfrmMAMS.ClearInsertSampleGrids;
var
  aCol, aRow, i: integer;
begin
  with grdPreviewUser do
  begin
    grdPreviewUser.ColCount := 3; // could have been rest by previous list
    for aRow := 1 to RowCount - 1 do
      for aCol := 1 to Colcount - 1 do cells[aCol, aRow] := '';
  end;
  with grdPreviewSamples do
  begin
    for aRow := 0 to RowCount - 1 do
      for aCol := 0 to Colcount - 1 do cells[aCol, aRow] := '';
  end;
  ProjectName := '';
  //clear invoice table
  for i := 0 to grdInvoiceAddress.RowCount - 1 do
  Begin
    grdInvoiceAddress.Rows[i].Clear();
  End;
end;

procedure TfrmMAMS.cmbFilterSampleMaterialClick(Sender: TObject);
begin
  btnRefreshSamplesAvailable.Enabled := (cmbFilterSampleMaterial.KeyValue > 0);
  gbxBatch.Enabled := (cmbFilterSampleMaterial.KeyValue > 0);
end;

procedure TfrmMAMS.cmbPretreatMethodClick(Sender: TObject);
begin
  btnRefreshSamplesAvailable.Enabled := (cmbPretreatMethod.KeyValue > 0);
  gbxBatch.Enabled := (cmbPretreatMethod.KeyValue > 0);
end;

procedure TfrmMAMS.cmbProjectOfReportCloseUp(Sender: TObject);
begin
  btnQuerySubmitter.Enabled := true;
  GetSamplesForReport(rgpExport.ItemIndex = 2);
  btnGuessReportNameClick(self);
  btnReport.Enabled := true;
end;

procedure TfrmMAMS.cmbStep1CloseUp(Sender: TObject);
begin
  SampleModified := true;
end;

procedure TfrmMAMS.cmbSubmitterNameProjectCloseUp(Sender: TObject);
// when dropdown list closes, show the projects associated to this user
begin
  if cmbSubmitterNameProject.KeyValue = Null then exit;
  GetProjects;
  cmbUsernameUserInfo.KeyValue:=cmbSubmitterNameProject.ListSource.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list in user info to correct user_nr in case user switches to this page
  btnAddNewProject.Visible:=true;
end;

procedure TfrmMAMS.cmbSubmitterNameProjectKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then begin
    cmbSubmNameReport.KeyValue := cmbSubmitterNameProject.KeyValue;
    cmbUsernameUserInfo.KeyValue := cmbSubmitterNameProject.KeyValue;
    GetProjects;
    grdSamplesOfProject.Visible := true;
  end;
end;

procedure TfrmMAMS.cmbSubmNameReportCloseUp(Sender: TObject);
begin
  if cmbSubmNameReport.KeyValue = Null then exit;
  if not chkAllProjects.Checked then
  begin
    with dm.qryProjectsOfReport do
    begin
      Close;
      SQL.Text := 'SELECT project_nr, project, in_date FROM project_t WHERE user_nr='
        + IntToStr(cmbSubmNameReport.KeyValue) + ' ORDER BY in_date DESC;';
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      cmbProjectOfReport.DropDown;
    end;
  end
  else
  begin
    btnQuerySubmitter.Enabled := true;
    btnQuerySubmitter.SetFocus;
    GetSamplesForReport(rgpExport.ItemIndex = 2);
  end;
end;

procedure TfrmMAMS.cmbSubmNameReportKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then begin
    cmbSubmNameReportCloseUp(Sender);
  end;
  if chkAllProjects.Checked then begin
    btnQuerySubmitter.Enabled := true;
    GetSamplesForReport(rgpExport.ItemIndex = 2);
  end;
{
  cmbProjectOfReport.Enabled := not chkAllProjects.Checked;
    else begin
      cmbSubmNameReportCloseUp(Sender);
    end;
  end;
//  if Key = VK_RETURN then btnQuerySubmitter.SetFocus;
}
end;


procedure TfrmMAMS.cmbSubmNameReportMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
{
  cmbProjectOfReport.Enabled := not chkAllProjects.Checked;
    if chkAllProjects.Checked then begin
       btnQuerySubmitter.Enabled := true;
       GetSamplesForReport;
    end
    else begin
      cmbSubmNameReportCloseUp(Sender);
      GetSamplesForReport;
    end;
    }
end;

procedure TfrmMAMS.cmbUsernameUserInfoEnter(Sender: TObject);
begin
  dm.tblUser.IndexFieldNames := 'last_name';
end;

procedure TfrmMAMS.cmbUsernameUserInfoExit(Sender: TObject);
begin
//  dm.tblUser.IndexFieldNames := 'user_nr';
end;

procedure TfrmMAMS.CreateCds(German: boolean);
//creates an in memory database with the colums that need to be transfered to the report table
begin
  cdsExport.Close;
  with cdsExport.FieldDefs do    // go through the list of ReportHeadings and create FieldDefinitions
  begin
    Clear;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 1];
      DataType := ftInteger;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 2];
      DataType := ftString;
      Size := 100;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 3];
      DataType := ftString;
      Size := 60;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 4];
      DataType := ftString;
      Size := 60;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 5];
      DataType := ftString;
      Size := 60;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 6];
      DataType := ftInteger;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 7];
      DataType := ftInteger;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 8];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := 'Cal 1 sigma';
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := 'Cal 2 sigma';
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 9];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 10];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 11];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 12];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 13];
      DataType := ftString;
    end;
    with AddFieldDef do
    begin
      Name := grdReportHeadings.cells[2, 14];
      DataType := ftString;
    end;
    with AddFieldDef do // pmC
    begin
      Name := grdReportHeadings.cells[2, 16];
      //ShowMessage(Name);
      DataType := ftString;
    end;
    with AddFieldDef do   // pmC_err
    begin
      Name := grdReportHeadings.cells[2, 17];
      //ShowMessage(Name);
      DataType := ftString;
    end;
  end;
  cdsExport.CreateDataSet;
end;

procedure TfrmMAMS.CreateWordExport;
  var
  FileNameTemplate, FileNameSave:OleVariant;
  vWhat, vBookmark:OleVariant;
begin
  (*FilenameTemplate:=edtWordTemplate.Text;   // path to word template
  FilenameSave:=edtSaveReportAs.Text;       // save loaction of new word file
  WordApplication1.Connect;    //connect to word aplication
  WordApplication1.Documents.OpenOld(FileNameTemplate, EmptyParam, EmptyParam,
  EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
  EmptyParam, EmptyParam);          //open word document
  WordDocument1.ConnectTo(WordApplication1.ActiveDocument);  //connect word document to document object
  WordApplication1.Visible:=true;  // show word window

  // replace bookmarks in word document with text
  vWhat:=wdGoToBookmark;   //define type of object (bookmark)
  vBookmark:='OrderNr';    //define name of object
  WordApplication1.Selection.GoTo_(vWhat,emptyParam,emptyParam,vBookmark);  //jump to object an select text
  WordApplication1.Selection.TypeText('22-07-1976'); //replace bookmark with new text

  //WordDocument1.Close;
  //WordDocument1.Disconnect;
  WordDocument1.ConnectTo(WordApplication1.ActiveDocument);
  WordDocument1.SaveAs(FileNameSave);
  WordDocument1.Disconnect;

  WordApplication1.Disconnect;     // disconnect word application *)


  (* with WordExport do begin
        Grid := grdSamplesOfSubmitter; //Grid where the data come from
        FileName:= edtSaveReportAs.FileName; //save in the same directory as the report
        ExportGrid;
  end; *)

// old way of doing it using worddriver
 if Assigned(FWord) then FWord.Free;
  if assigned(FValues) then FValues.Free;
  FWord := TWordDriver.Create(Self);
  FUseWordVersion := wvDetect;
  FValues := TStringList.Create;
end;

procedure TfrmMAMS.edtWeightCombustionChange(Sender: TObject);
begin
  WeightsChanged := true;
  CalculateCAbsWeight;
end;

procedure TfrmMAMS.edtWeightCombustionClick(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.grdShowInvoiceCellClick(Column: TColumn);
begin
  glbInvoiceNr := dm.tblInvoice.FieldByName('user_nr').AsInteger;
  btnInsertSelectedInvoice.Enabled := true;
  InvoiceExists := true;
end;

procedure TfrmMAMS.DBGridDBPlotCellClick(Column: TColumn);
// when clicking on a cell, display the record number in a lable
// select datapoint in plot
Var list: TlabelsList;
begin
  lblSelectedPlotRow.Caption := 'selected: ' + inttostr(DBGridDBPlot.DataSource.DataSet.RecNo);
  // using the recNo, find the datapoint where label==recNo
  list := DBChart.Series[0].Labels;
  DBChart.Series[0].Selected;
  //showmessage(list.Labels[DBGridDBPlot.DataSource.DataSet.RecNo]);


  //DBChart.Series[0].Selected
end;

procedure TfrmMAMS.DBGridHomeExpressSamplesDblClick(Sender: TObject);
// jump to the sample info page of the selected sample
Var samplenr :string;
begin
  // get MAMS of samples of the query
  samplenr := dm.dsWaitingExpress.DataSet.FieldByName('sample_nr').AsString;
  // switch to sample info age
  pgtMain.ActivePage:=tbsSampleInfo; //display sample info page and fill in information
  // set MAMS in sample info page
  edtSampleNr.Text:=samplenr;
  // get sample info
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
end;

procedure TfrmMAMS.DBGridHomeExpressSamplesDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  AlternateRowColors(Sender, State);
end;

procedure TfrmMAMS.DBGridSampleExchangeCellClick(Column: TColumn);
begin
// toggle checkboxes
LogWindow.addLogEntry('Exchange Table: column clicked = ' + Column.FieldName);
  if ((Column.FieldName='return_to_sender') OR
  (Column.FieldName='returned_to_sender') OR
  (Column.FieldName='CNIsotopA') OR
  (Column.FieldName='CNIsotopAMoved')) then
      begin
      LogWindow.addLogEntry('Exchange Table: trying to chenge the boolean value');
      LogWindow.addLogEntry('Exchange Table: old value is ' + Column.Field.AsString);

      Column.Grid.DataSource.DataSet.Edit;
      if Column.Field.AsInteger=0 then Column.Field.Value :=1;
      if Column.Field.AsInteger=1 then Column.Field.Value :=0;

      LogWindow.addLogEntry('Exchange Table: new  value is ' + Column.Field.AsString);
        //Column.Field.Value:= not Column.Field.AsBoolean;
        //Column.Grid.DataSource.DataSet.Post;


        // toggle True and False
        //Column.Grid.DataSource.DataSet.Edit;
        //if Column.Field.AsInteger=0 then Column.Field.Value:=1;
        //if Column.Field.AsInteger=1 then Column.Field.Value:=0;
          //immediate post - see for yourself whether you want this
          //Column.Grid.DataSource.DataSet.Post;
          //Column.Grid.DataSource.DataSet.Cancel;
      end;

end;

procedure TfrmMAMS.DBGridSampleExchangeColEnter(Sender: TObject);
begin
if Self.DBGridSampleExchange.SelectedField.DataType = ftBoolean then
  begin
    Self.GridOriginalOptions := Self.DBGridSampleExchange.Options;
    Self.DBGrid1.Options := Self.DBGridSampleExchange.Options - [dgEditing];
  end;
end;

procedure TfrmMAMS.DBGridSampleExchangeColExit(Sender: TObject);
begin
 if Self.DBGridSampleExchange.SelectedField.DataType = ftBoolean then
    Self.DBGridSampleExchange.Options := Self.GridOriginalOptions;
end;

procedure TfrmMAMS.DBGridSampleExchangeDblClick(Sender: TObject);
var
 Column: TColumn;
begin
  if RadioGroupSampleExchange.ItemIndex = 0 then
    begin
      pgtMain.ActivePage:=tbsProjectsOfUser; //change to project info page
      cmbSubmitterNameProject.KeyValue:=DBGridSampleExchange.DataSource.DataSet.FieldByName('user_nr').AsInteger; // get user_nr from dataset and set dropdown list to correct user_nr
      cmbSubmitterNameProjectCloseUp(self); //correct user is selected now show their projects
      grdProjects.DataSource.DataSet.Locate('project_nr',DBGridSampleExchange.DataSource.DataSet.FieldByName('project_nr').AsInteger,[loPartialKey]);//now the projects are being displayed, select the correct project
      grdProjectsCellClick(Column); //now display the samples that belong to the project
    end
  else if RadioGroupSampleExchange.ItemIndex = 1 then
    begin
      ShowSampleInfoPage(Sender);
    end;

end;

procedure TfrmMAMS.DBGridSampleExchangeDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
CONST
  CtrlState: array[0..1] of integer = (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED) ;
begin
  // AlternateRowColors(Sender, State);
  // TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);

  // display boolean fields as checkboxes
  if (Column.FieldName='return_to_sender') OR
  (Column.FieldName='returned_to_sender') OR
  (Column.FieldName='prep_return_to_sender') OR
  (Column.FieldName='prep_returned_to_sender') OR
  (Column.FieldName='CNIsotopA') OR
  (Column.FieldName='CNIsotopAMoved') then
  begin
    TDBGrid(Sender).Canvas.FillRect(Rect) ;
    if (VarIsNull(Column.Field.Value)) then
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, DFCS_BUTTONCHECK or DFCS_INACTIVE) {grayed}
    else
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, CtrlState[Column.Field.AsInteger]) ; {checked or unchecked}
  end;
end;

procedure TfrmMAMS.DBGridSamplesOfUnpressedMagazineDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  //if co2_final is below 300 color in red since this might be too little of a sample
  if dm.qrySamplesOfUnpressedMagazine.FieldByName('co2_final').AsInteger < 300 Then
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clRed;
     End;
  TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.DBLookupComboBox2Enter(Sender: TObject);
begin
  dm.tblUser.IndexFieldNames := 'last_name';
end;

procedure TfrmMAMS.DBLookupComboBox2Exit(Sender: TObject);
begin
  dm.tblUser.IndexFieldNames := 'user_nr';
end;

procedure TfrmMAMS.DBMemoTouchWeightsPrepCommentKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
    btnTouchWeightsPrepNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.DBMemoTouchWeightsTargetCommentChange(Sender: TObject);
begin
    btnTouchWeightsGraphNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.DoSampleInfo(SampleNr: Integer; PrepNr: Integer; TargetNr: Integer);
var
  NPreps, NTargets: integer;
begin
  LogWindow.addLogEntry('Run DoSampleInfo');
  JvLogFile.Add('DB',lesError,'Run DoSampleInfo');
  //get and display number of available sample preps
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  edtSamplePrepNr.MaxValue := NPreps;
  if NPreps > 1 then
  begin     // display number of preps in a static label
    lbPrepNr.Caption := '1...' + IntToStr(NPreps);
    edtSamplePrepNr.Enabled := true;
    edtSamplePrepNr.MaxValue := NPreps;
  end
  else
  begin
    lbPrepNr.Caption := '';
    // if there is only one prep, set the prep number to one.
    // This sometimes necessary when swithcing from a sample with many preps
    // to a sample with one prep
    PrepNr := NPreps;
  end;
  // diplay number of available targets for the selected sampleNr and PrepNr.
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, round(edtSamplePrepNr.Value));
  edtSampleTargetNr.MaxValue := NTargets;
  if NTargets > 1 then
  begin     // display number of targets
    lbTargetNr.Caption := '1...' + IntToStr(NTargets);
    edtSampleTargetNr.Enabled := true;
  end
  else
  begin
    lbTargetNr.Caption := '';
    // if there is only one target, set the target number to one.
    // This sometimes necessary when switching from a sample with many taregts
    // to a sample with one target
    TargetNr := Ntargets;
  end;

  // get the pretreatment steps from the database and populate list
  SetupPretreatmentStepsGrid;
  //display preparation steps
  dm.QueryPrepStepsBySampleNr(SampleNr, PrepNr);

  // display the sample info
  with dm.qrySampleInfo do
  begin
    cmbSampleMaterial.KeyValue := 0;
    cmbSampleFraction.KeyValue := 0;
    cmbSampleType.KeyValue := 0;

    // ########################
    // query the database and get all sample info
    // dm.GetSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
    dm.GetSampleInfo(SampleNr, PrepNr, TargetNr);

    LogWindow.addLogEntry('reformat number of digits');
    JvLogFile.Add('DB',lesError,'reformat number of digits');

    TFloatField(FieldByName('c14_age')).DisplayFormat := '0';
    TFloatField(FieldByName('c14_age_sig')).DisplayFormat := '0';
    TFloatField(FieldByName('fm')).DisplayFormat := '0.0000';
    TFloatField(FieldByName('fm_sig')).DisplayFormat := '0.0000';
    TFloatField(FieldByName('dc13')).DisplayFormat := '0.0000';
    TFloatField(FieldByName('dc13_sig')).DisplayFormat := '0.0';
    TFloatField(FieldByName('cn_ratio')).DisplayFormat := '0.0';
    TFloatField(FieldByName('conc_c')).DisplayFormat := '0.0';
    TFloatField(FieldByName('conc_n')).DisplayFormat := '0.0';

    dm.tblMaterial.Locate('material', FieldByName('material').AsString, [loPartialKey]);
    cmbSampleMaterial.KeyValue := dm.tblMaterial.FieldByName('indexnr').AsInteger;
    dm.tblFraction.Locate('fraction', FieldByName('fraction').AsString, [loPartialKey]);
    cmbSampleFraction.KeyValue := dm.tblFraction.FieldByName('indexnr').AsInteger;
    dm.tblTypes.Locate('type', FieldByName('type').AsString, [loPartialKey]);
    cmbSampleType.KeyValue := dm.tblTypes.FieldByName('indexnr').AsInteger;
    dm.tblProjectStatus.Locate('status', FieldByName('status').AsString, [loPartialKey]);
    cmbProjectStatus.KeyValue := dm.tblProjectStatus.FieldByName('indexnr').AsInteger;
    edtSampPrep1.Text := FieldByName('step1_method').AsString;
    edtSampPrep2.Text := FieldByName('step2_method').AsString;
    edtSampPrep3.Text := FieldByName('step3_method').AsString;
    edtSampPrep4.Text := FieldByName('step4_method').AsString;
    edtSampPrep5.Text := FieldByName('step5_method').AsString;
  end;
  dm.GetGraphWeight(SampleNr, round(edtSamplePrepNr.Value), round(edtSampleTargetNr.Value));
  dm.GetWeights(SampleNr, round(edtSamplePrepNr.Value));

  //CalculateYield; // calculate the yield from the weights

  TargetDataChanged := false;
end;

procedure TfrmMAMS.HelpAbout1Execute(Sender: TObject);
begin
  AboutBox.Version.Caption := 'Version: ' + myVersion;
  AboutBox.ShowModal;
end;

procedure TfrmMAMS.lbxBatchDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  s: string;
begin
  if rgpTask.ItemIndex = 0 then
    with (Source as TDBGrid) do begin
      s := DataSource.DataSet.Fields.FieldByName('sample_nr').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('user_label').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('user_label_nr').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('prep_nr').AsString;
    end
  else
    with (Source as TDBGrid) do begin
      s := DataSource.DataSet.Fields.FieldByName('sample_nr').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('user_label').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('user_label_nr').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('prep_nr').AsString + ' | ' +
        DataSource.DataSet.Fields.FieldByName('target_nr').AsString;
    end;
  lbxBatch.Items.Add(s);
  btnSaveBatch.Enabled := (Length(edtBatchName.Text) > 0) and (lbxBatch.Items.Count > 0);
end;

procedure TfrmMAMS.lbxBatchDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  i, nr, lbnr, targetnr, lbtargetnr: integer;
  found: boolean;
begin
  if rgpTask.ItemIndex = 1 then begin
    with (Source as TDBGrid) do begin
      nr := DataSource.DataSet.Fields.FieldByName('sample_nr').AsInteger;
      targetnr := DataSource.DataSet.Fields.FieldByName('target_nr').AsInteger;
    end;
    found := false;
    for i := 0 to lbxBatch.Count - 1 do begin
      lbnr := StrToInt(trim(ExtractWord(1, lbxBatch.Items[i], ['|'])));
      lbtargetnr := StrToInt(trim(ExtractWord(5, lbxBatch.Items[i], ['|'])));
      if (lbnr = nr) and (targetnr = lbtargetnr) then found := true;
    end;
  end
  else
    with (Source as TDBGrid) do begin
      nr := DataSource.DataSet.Fields.FieldByName('sample_nr').AsInteger;
    end;
  found := false;
  for i := 0 to lbxBatch.Count - 1 do begin
    lbnr := StrToInt(trim(ExtractWord(1, lbxBatch.Items[i], ['|'])));
    lbtargetnr := StrToInt(trim(ExtractWord(4, lbxBatch.Items[i], ['|'])));
    if (lbnr = nr) and (targetnr = lbtargetnr) then found := true;
//    if (lbnr = nr) then found := true;
  end;
  Accept := not found;
end;

procedure TfrmMAMS.lbxDefinePrepStepsDblClick(Sender: TObject);
var i: integer;
begin
  for i := 0 to lbxDefinePrepSteps.Count - 1 do
    if lbxDefinePrepSteps.Selected[i] then CurrentPrep := lbxDefinePrepSteps.Items[i];
  with grdPreviewSamples do begin
    for i := 1 to RowCount - 1 do Cells[SamplePrep1Col, i] := CurrentPrep;
  end;
end;

procedure TfrmMAMS.lbxDefinePrepStepsMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  SetCurrentToAll := false;
  for i := 0 to lbxDefinePrepSteps.Count - 1 do
    if lbxDefinePrepSteps.Selected[i] then CurrentPrep := lbxDefinePrepSteps.Items[i];
  DragSource := drgPrep;
  if ssCtrl in Shift then SetCurrentToAll := true;
  lbxDefinePrepSteps.BeginDrag(false, 5);
end;

procedure TfrmMAMS.lbxFractionDblClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lbxFraction.Count - 1 do
    if lbxFraction.Selected[i] then CurrentFraction := lbxFraction.Items[i];
  with grdPreviewSamples do begin
    for i := 1 to RowCount - 1 do Cells[FractionCol, i] := CurrentFraction;
  end;
  wizSelectMaterial.EnableButton(bkNext, true);
end;

procedure TfrmMAMS.lbxFractionMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  SetCurrentToAll := false;
  for i := 0 to lbxFraction.Count - 1 do
    if lbxFraction.Selected[i] then CurrentFraction := lbxFraction.Items[i];
  DragSource := drgFraction;
  if ssCtrl in Shift then SetCurrentToAll := true;
  lbxFraction.BeginDrag(false, 5);
end;

procedure TfrmMAMS.lbxLabTaskBatchClick(Sender: TObject);
begin
//  if rgpLabTask.ItemIndex = 1 then gbxGraphParams.Visible := true;
end;

procedure TfrmMAMS.lbxMaterialDblClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lbxMaterial.Count - 1 do
    if lbxMaterial.Selected[i] then CurrentMaterial := lbxMaterial.Items[i];
  with grdPreviewSamples do begin
    for i := 1 to RowCount - 1 do Cells[MaterialCol, i] := CurrentMaterial;
  end;
end;

procedure TfrmMAMS.lbxMaterialMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  SetCurrentToAll := false;
  for i := 0 to lbxMaterial.Count - 1 do
    if lbxMaterial.Selected[i] then CurrentMaterial := lbxMaterial.Items[i];
  DragSource := drgMaterial;
  if ssCtrl in Shift then SetCurrentToAll := true;
  lbxMaterial.BeginDrag(false, 5);
end;

procedure TfrmMAMS.lbxPrepStepsClick(Sender: TObject);
begin
  btnStartTask.Enabled := true;
  btnEndTask.Enabled := true;
end;

procedure TfrmMAMS.lbxStepsMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  for i := 0 to lbxSteps.Count - 1 do
    if lbxSteps.Selected[i] then Prep := lbxSteps.Items[i];
//  DragSource := drgPrep;
  lbxSteps.BeginDrag(false, 5);
end;

procedure TfrmMAMS.lbxTypesDblClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lbxTypes.Count - 1 do
    if lbxTypes.Selected[i] then CurrentType := lbxTypes.Items[i];
  with grdPreviewSamples do begin
    for i := 1 to RowCount - 1 do Cells[TypeCol, i] := CurrentType;
  end;
  wizSelectType.EnableButton(bkNext, true);
end;

procedure TfrmMAMS.lbxTypesMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  SetCurrentToAll := false;
  for i := 0 to lbxTypes.Count - 1 do
    if lbxTypes.Selected[i] then CurrentType := lbxTypes.Items[i];
  Dragsource := drgType;
  if ssCtrl in Shift then SetCurrentToAll := true;
  lbxTypes.BeginDrag(false, 5);
end;

procedure TfrmMAMS.memTargetCommentsChange(Sender: TObject);
begin
  TargetDataChanged := true;
end;

procedure TfrmMAMS.OptionsTreeClick(Sender: TObject);
//open the correct Tab in TabOptionsPages
begin
  //showmessage((Sender AS TTreeView).Selected.Text)
   // case structure does not allow strings, enter the strings into the IndexText function in order to get an integer value for CASE
   CASE IndexText( (Sender AS TTreeView).Selected.Text, ['General','Database','Paths','Email']) of
   0: TabOptionsPages.ActivePage := TabGeneral;  //General node selected
   1: TabOptionsPages.ActivePage := TabDatabase; //Database node selected
   2: TabOptionsPages.ActivePage := TabPaths;    //Paths node selected
   3: TabOptionsPages.ActivePage := TabEmail;    //Email node selected
   end;
end;

procedure TfrmMAMS.ParseNewSampleFile(const FName: string);
//performs the wizzard, asking for new samples to be importet into the DB
var
  InF: TextFile;
  s, s1, s2, surname,first_name: string;
  i, k, Row: integer;
  end_found, Entry_empty: boolean;
begin
  if Length(FName) > 0 then
  begin
    grdPreviewUser.Visible := true;
    grdPreviewSamples.Visible := true;
    AssignFile(InF, FName);
    Reset(InF);
    Row := 1;
    SameAddressForInvoice := false;

    // build headers for table that holds the new user information
    with grdPreviewUser do
    begin
//      Selection := NoSelection;
      ColWidths[0] := 180;
      cells[1, 0] := 'user;  * required';
      ColWidths[1] := 250;
      cells[0, 1] := 'first name *';
      cells[0, 2] := 'last name *';
      cells[0, 3] := 'organisation';
      cells[0, 4] := 'institution';
      cells[0, 5] := 'adress 1';
      cells[0, 6] := 'adress 2';
      cells[0, 7] := 'city';
      cells[0, 8] := 'zip';
      cells[0, 9] := 'country';
      cells[0, 10] := 'phone 1';
      cells[0, 11] := 'phone 2';
      cells[0, 12] := 'fax';
      cells[0, 13] := 'email';
      cells[0, 14] := 'www';
      cells[0, 15] := 'invoice';
      ColCount := 2;
    end;

    with grdInvoiceAddress do
    begin
//      Selection := NoSelection;
      ColWidths[0] := 180;
      ColWidths[1] := 250;
      cells[1, 0] := 'invoice; * required';
      cells[0, 1] := 'first name *';
      cells[0, 2] := 'last name *';
      cells[0, 3] := 'organisation';
      cells[0, 4] := 'institution';
      cells[0, 5] := 'adress 1';
      cells[0, 6] := 'adress 2';
      cells[0, 7] := 'city';
      cells[0, 8] := 'zip';
      cells[0, 9] := 'country';
      cells[0, 10] := 'phone 1';
      cells[0, 11] := 'phone 2';
      cells[0, 12] := 'fax';
      cells[0, 13] := 'email';
      cells[0, 14] := 'www';
      cells[0, 15] := 'invoice';
      ColCount := 2;
    end;

    SetupNewSamplesGrid;
    btnInsertNewUser.Visible := true;
    btnInsertExistingUser.Visible := true;

    // read data from import file

    // Project Name ######################
    // read lines until "Project Name" is discovered
    repeat // skip header
      Readln(InF, s);
      s1 := ExtractWord(1, s, [';']);
      s1 := dm.ReplaceBadCharacters(s1);
    until (Pos('Project name', s1) > 0) or EOF(InF);
    if EOF(InF) then ShowMessage('Text Project name in cell A1 is missing!');
    ProjectName := ExtractWord(2, s, [';']);
    ProjectName := copy(ProjectName, 1, 99);   // project name limited to 99 characters
    ProjectName := dm.ReplaceBadCharacters(ProjectName);

    Readln(InF, s); // skip 'required'
    Readln(InF, s);
    s1 := ExtractWord(1, s, [';']);
    s1 := dm.ReplaceBadCharacters(s1);
    if Length(s1) > 0 then
    MANummer := StrToInt(copy(s1, 1, Length(s1))); //MA-113333

    // First Name ######################
    repeat // now parse user, skip header
      Readln(InF, s);
      s1 := ExtractWord(1, s, [';']);
      s1 := dm.ReplaceBadCharacters(s1);
    until Pos('first name', s1) > 0;

    // get all other user data ######################
    while not EOF(InF) and not end_found do
    begin
      end_found := false;
      repeat // read user data
        while Pos(';', s) = 0 do Readln(InF, s); // skip german headers (CR seperated from English header)
        s1 := ExtractWord(1, s, [';']); // read 1st col
        s1 := dm.ReplaceBadCharacters(s1);
        if pos('accounting', s1) > 0 then
        begin
          end_found := true;
          s2 := trim(ExtractWord(2, s, [';'])); // get user data
          s2 := UpperCase(s2);
          if (Pos('JA', s2) > 0) or (Pos('YES', s2) > 0) then  //check whether user address is also invoice adress
          begin
            SameAddressForInvoice := true;
          end
          else
          Begin
            SameAddressForInvoice := false;
          End;
        end;
        grdPreviewUser.ColWidths[0] := 80;
        dm.tblUser.IndexFieldNames := 'last_name';
        s2 := trim(ExtractWord(2, s, [';'])); // get user data
        s2 := dm.ReplaceBadCharacters(s2);
        s2 := copy(s2, 1, 59);
        if Length(s2) > 0 then
        begin  //write info into cells
          if Pos('first name', s1) > 0 then grdPreviewUser.Cells[1, 1] := s2;
          if Pos('last name', s1) > 0 then grdPreviewUser.Cells[1, 2] := s2;
          if Pos('organisation', s1) > 0 then grdPreviewUser.Cells[1, 3] := s2;
          if Pos('institution', s1) > 0 then grdPreviewUser.Cells[1, 4] := s2;
          if Pos('address1', s1) > 0 then grdPreviewUser.Cells[1, 5] := s2;
          if Pos('address2', s1) > 0 then grdPreviewUser.Cells[1, 6] := s2;
          if Pos('city', s1) > 0 then grdPreviewUser.Cells[1, 7] := s2;
          if Pos('zip', s1) > 0 then grdPreviewUser.Cells[1, 8] := s2;
          if Pos('country', s1) > 0 then grdPreviewUser.Cells[1, 9] := s2;
          if Pos('phone1', s1) > 0 then grdPreviewUser.Cells[1, 10] := s2;
          if Pos('phone2', s1) > 0 then grdPreviewUser.Cells[1, 11] := s2;
          if Pos('Fax', s1) > 0 then grdPreviewUser.Cells[1, 12] := s2;
          if Pos('email', s1) > 0 then grdPreviewUser.Cells[1, 13] := s2;
          if Pos('www', s1) > 0 then grdPreviewUser.Cells[1, 14] := s2;
          if Pos('accounting', s1) > 0 then grdPreviewUser.Cells[1, 15] := s2;
        end;
        Readln(InF, s);
      until end_found;
    end;

    UserExists := false;
    with wizStartPage do
    begin
      Title.Text := '';
      EnableButton(bkNext, false);
    end;
    grdShowUsers.Visible := true;
    btnInsertExistingUser.Enabled := false;
    glbUserNr := 0;
    // get first und last name from user given in file
    // and try to find that user in the database
    surname := grdPreviewUser.Cells[1, 2];
    first_name := grdPreviewUser.Cells[1, 1];
    //UserExists := dm.tblUser.Locate('last_name; first_name', VarArrayOf([surname, first_name]), [loPartialKey]); //locate user in database
    UserExists := dm.tblUser.Locate('last_name',surname, [loPartialKey]); //locate users lastname in database
    if UserExists then
    begin
      btnInsertExistingUser.Enabled := true;
      glbUserNr := dm.tblUser.FieldByName('user_nr').AsInteger;
    end;

    // Invoice Adress ######################
    if not SameAddressForInvoice then
    begin
      // now parse invoice data
      repeat //skip header
        Readln(InF, s);
        s1 := ExtractWord(1, s, [';']);
        s1 := dm.ReplaceBadCharacters(s1);
      until Pos('institution', s1) > 0;
      end_found := false;

      // read all invoice data from file
      while not EOF(InF) and not end_found do
      begin
        end_found := false;
        repeat // read user data
          while Pos(';', s) = 0 do Readln(InF, s); // skip german headers (CR seperated from English header)
          s1 := ExtractWord(1, s, [';']); // read 1st col
          s1 := dm.ReplaceBadCharacters(s1);
          if pos('www', s1) > 0 then  //www is the last line of the invoice data
          begin
            end_found := true;
          end;
          s2 := trim(ExtractWord(2, s, [';'])); // get invoice data
          s2 := dm.ReplaceBadCharacters(s2);
          if Length(s2) > 0 then
          begin
            if Pos('first name', s1) > 0 then grdInvoiceAddress.Cells[1, 1] := s2;
            if Pos('last name', s1) > 0 then grdInvoiceAddress.Cells[1, 2] := s2;
            if Pos('organisation', s1) > 0 then grdInvoiceAddress.Cells[1, 3] := s2;
            if Pos('institution', s1) > 0 then grdInvoiceAddress.Cells[1, 4] := s2;
            if Pos('address1', s1) > 0 then grdInvoiceAddress.Cells[1, 5] := s2;
            if Pos('adress2', s1) > 0 then grdInvoiceAddress.Cells[1, 6] := s2;
            if Pos('city', s1) > 0 then grdInvoiceAddress.Cells[1, 7] := s2;
            if Pos('zip', s1) > 0 then grdInvoiceAddress.Cells[1, 8] := s2;
            if Pos('country', s1) > 0 then grdInvoiceAddress.Cells[1, 9] := s2;
            if Pos('phone1', s1) > 0 then grdInvoiceAddress.Cells[1, 10] := s2;
            if Pos('phone2', s1) > 0 then grdInvoiceAddress.Cells[1, 11] := s2;
            if Pos('fax', s1) > 0 then grdInvoiceAddress.Cells[1, 12] := s2;
            if Pos('email', s1) > 0 then grdInvoiceAddress.Cells[1, 13] := s2;
            if Pos('www', s1) > 0 then grdInvoiceAddress.Cells[1, 14] := s2;
            if Pos('accounting', s1) > 0 then grdInvoiceAddress.Cells[1, 15] := s2;
          end;
          Readln(InF, s);
        until end_found;
      end;
    end
    else // clear all invoice fields when SameAddressForInvoice, to avoid carrying over old data
    begin
         for i := 0 to grdInvoiceAddress.RowCount - 1 do grdInvoiceAddress.Rows[i].Clear();
    end;

    surname := grdPreviewUser.Cells[1, 3];
    if dm.tblInvoice.Locate('last_name', surname, [loPartialKey]) then
    begin
      btnInsertSelectedInvoice.Enabled := true;
      glbInvoiceNr := dm.tblInvoice.FieldByName('user_nr').AsInteger;
    end
    else
    begin
      btnInsertSelectedInvoice.Enabled := false;
      glbInvoiceNr := 0;
    end;

    if Length(ProjectName) = 0 then
    begin
      ShowMessage('Project name cannot be empty; you may add a project name later on in the project page of this wizard');
    end;
    Row := 1;

    repeat // skip header
      Readln(InF, s);
      s1 := ExtractWord(1, s, [';']);
      s1 := dm.ReplaceBadCharacters(s1);
    until Pos('required', s1) > 0;

    // get all sample information #############################
    while not EOF(InF) do
    begin // now parse samples
      ReadLn(InF, s);
      s1 := ExtractWord(1, s, [';']); // read sample name (first word in the delimted string)
      s1 := copy(s1, 1, 100);   // limit sample name to 100 characters
      if Length(s1) > 0 then
      begin
      //showmessage(s1);
        s := StringReplace(s, ';;', '; ;', [rfReplaceAll]);
        s := StringReplace(s, ';;', '; ;', [rfReplaceAll]); // once does not do all
        s := StringReplace(s, ';;', '; ;', [rfReplaceAll]);
        k := WordCount(s, [';']);  // number of sample information (columns in file) given

        // get the individual sample information of one specific sample here, e.g. descr, weight...
        // starting with "sample_nr"
        // i = 2: sample_nr
        // i = 3: descr_1
        // i = 4: desc_2
        // i = 5: weight
        // i = 6: material
        // i = 7: comment

        for i := 2 to k do
        begin
          s2 := ExtractWord(i, s, [';']);
          //showmessage('i=' + inttostr(i) + ', // s='+ s + ', // s2=' + s2);

          // limit length of string according to the database field size
          if i =2  then s2 := copy(s2, 1, 40); // sample_nr, 40 characters
          if i =3  then s2 := copy(s2, 1, 100); // desc_1, 100 characters
          if i =4  then s2 := copy(s2, 1, 40); // descr_2, 40 characters
          if i =5  then s2 := dm.ExtractNumber(s2);  // weight, clean the string and allow only a number
          if i =6  then s2 := copy(s2, 1, 20); // material, 20 characters
          if i =7  then s2 := copy(s2, 1, 100); // comment, 100 characters

//          if i < 7 then
//            s2 := copy(s2, 1, 40)
//          else
//            s2 := copy(s2, 1, 100);

          if s2 <> '' then
          begin
            s2 := dm.ReplaceBadCharacters(s2);
            grdPreviewSamples.Cells[i, Row] := s2;
          end;
        end;

        if not Entry_Empty then
        begin
          s1 := dm.ReplaceBadCharacters(s1);
          grdPreviewSamples.Cells[0, Row] := IntToStr(Row);
          grdPreviewSamples.Cells[1, Row] := s1;
          inc(Row);
          grdPreviewSamples.RowCount := Row;
        end;
      end;
    end;

    CloseFile(InF);
  end;
end;

procedure TfrmMAMS.pgtMainChange(Sender: TObject);
var
  Year, Month, Day: word;
begin
  SetFormat;
  if pgtMain.ActivePage = tbsUserReport then        // User Report
    begin
      DecodeDate(Date, Year, Month, Day);
      edtStartYear.Value := Year - 2;
      edtEndYear.Value := Year;
      edtStartAna.Value := 1;
      edtEndAna.Value := 1000000;
    end;
  if pgtMain.ActivePage = tbsLabTasks then         // Lab Tasks
    begin
      btnQueryToPretreatClick(Sender);
    end;
  if pgtMain.ActivePage = tbsProjectsOfUser then   // Projects of User
    Begin
    SetupProjectPage;
    End;
  if pgtMain.ActivePage = tbsSampleInfo then       // Sample Info
    begin
      btnDoSampleQuery.SetFocus;
    end;
  if pgtMain.ActivePage = tbsUserInfo then         // User Info
    begin
      cmbUsernameUserInfo.SetFocus;
    end;
  if pgtMain.ActivePage = tbsAdmin then            // Admin
    begin
      dm.GetNOxaBlank(true);
      lblOxa.Caption := ' Oxa = ' + IntToStr(dm.qryDB.RecordCount);
      with dm.qryDB do
        begin
          First;
          while not EOF do memoOxaBlank.Lines.Add(Fields.FieldByName('sample_nr').AsString);
          Next;
        end;
      dm.GetNOxaBlank(false);
      lblBlank.Caption := ' Oxa = ' + IntToStr(dm.qryDB.RecordCount);
    end;
end;

function GetSpecialFolderLocation(HWND: THandle; nFolder: integer): string;
//this is being used if the document data and fotos are stored locally
//it retrieves the location of the users profile folder
var
  pidl: PItemIDList;
  Path: array[0..MAX_PATH] of Char;
begin
  Result := '';
  if Succeeded(ShGetSpecialFolderLocation(HWND, nFolder, pidl)) then
  begin
    if ShGetPathfromIDList(pidl, Path) then Result := StrPas(Path);
    CoTaskMemFree(pidl);
  end;
end;


procedure TfrmMAMS.pgtSampleChange(Sender: TObject);
var
  fname, fnamePattern, fotoDir: string;
  DeviceList, VideoSizesList: TStringList;
  FilesList: TStringDynArray;
  i,j: Integer;

begin

  if pgtSample.ActivePage <> tbsProject then
  begin
    //Prepare folder information for images, documents etc etc
    if DirectoryExists('\\Riesling\KTA') then
    begin
      ServerRoot := '\\Riesling\KTA\';  // this is the root folder which is located on a server
      gbxProjectsOfUser.Caption := ' Projects of user (server)';
      LogWindow.addLogEntry('Server Root found: '+ ServerRoot);
      //enable buttons in order to load the docs if the correct folders can be found
      fname := ServerRoot + 'SAMS Documents\';
      if DirectoryExists(fname) then
        begin
          btnProjectLetters.Enabled:=true;
          LogWindow.addLogEntry('Document path found: '+fname);
        end
        else
        begin
          LogWindow.addLogEntry('Document folder does not exist!');
        end;
    end
    else
    begin
      ServerRoot := IncludeTrailingBackslash(GetSpecialFolderLocation(Handle, CSIDL_PROFILE));
      gbxProjectsOfUser.Caption := ' Projects of user (local)';
    end;
  end;
  Statusbar.Panels[2].Text := fname;

  // =======================================================================
  // display Images, Fotos
  if pgtSample.ActivePage = tbsFoto then
  lblLoadingFoto.Visible:=True;
  SampleFoto.Visible:=False;
  begin
    ImageFilesListBox.Clear;
    fotoDir := ServerRoot + 'SAMS Images\';    // this is the folder where the images are stored
    fname := fotoDir + edtSampleNr.Text + '.jpg';
    fnamePattern := edtSampleNr.Text + '*.jpg';
    LogWindow.addLogEntry('Image path should be: '+fname);
    Statusbar.Panels[2].Text := 'looking for images:' + fname;

    // check whether Directory exists
    // load found files into a ListBox in case more than one file exists
    // load the first file of the file list into the image control

    LogWindow.addLogEntry('check image directory :'+fotoDir);
    if DirectoryExists(fotoDir, False) then
      begin
        LogWindow.addLogEntry('Directory exists :'+fotoDir);
        // look for all images that fit the pattern, multiple images are xxb.jpg, xxc.jpg
        // save found files in FilesList
        // search will also be performed in SubDirectories
        LogWindow.addLogEntry('looking for multiple files matching:'+fnamePattern);
        FilesList := TDirectory.GetFiles(fotoDir, fnamePattern, TSearchOption.soAllDirectories);
        LogWindow.addLogEntry('number of matching files found:' + inttostr(Length(FilesList)));
        LogWindow.addLogEntry('displaying files in ListBox');
        for i := 0 to Length(FilesList)-1 do
          Begin
             ImageFilesListBox.Items.Add(FilesList[i]);
          End;

        // load first image
        if FileExists(FilesList[0]) then
          begin
            ImageFilesListBox.ItemIndex:= 0; //select first image in the list
            SampleFoto.Picture.LoadFromFile(FilesList[0]);
            LogWindow.addLogEntry('Image found: '+FilesList[0]);
            lblLoadingFoto.Visible:=False;
            SampleFoto.Visible:=True;
          end
          else
          begin
            ShowMessage('Unable to find image. Drive is connected but image is missing');
            LogWindow.addLogEntry('Image file does not exist!');
          end;
      end
      else
      begin
        LogWindow.addLogEntry('Directory does not exist :'+fotoDir);
        ShowMessage('Unable to find directory "' + fotoDir + '". Make sure network drive is connected.');
      end;
  end;

  // =======================================================================
  // display Calibration Graph
  if pgtSample.ActivePage = tbsCalibration then
  begin
    LogWindow.addLogEntry('creating calibration graph using OxCal at: '+ServerRoot);
    Statusbar.Panels[2].Text := 'looking for OxCalServer:' + ServerRoot;

    // looking for OxCal Server
    WebBrowserCalibration.Navigate('http://192.168.123.30:1950');
  end;



end;

procedure TfrmMAMS.pgtSampleExit(Sender: TObject);
begin
  LogWindow.addLogEntry('stopping video');
  fVideoImage.VideoStop;  // stop webcam video
end;

procedure TfrmMAMS.rgpLabTaskClick(Sender: TObject);
begin
  SetupLabTaskPage;
end;

procedure TfrmMAMS.rgpTaskClick(Sender: TObject);
begin
  if rgpTask.ItemIndex = 0 then
  // preparation batch
  begin
    pnlPrepChoice.Visible := true;
  end
  else
  // graphitisation batch
  begin
    pnlPrepChoice.Visible := false;
    edtBatchName.Text := '';
    btnRefreshSamplesAvailable.Enabled := true;
    btnRefreshSamplesAvailableClick(self);
  end;
end;

procedure TfrmMAMS.SendMail(German: boolean);
var
  i: integer;
  s: string;
begin
  edtMailTo.Text := dm.dsQryDb.DataSet.FieldByName('email').AsString;
  edtMailSubject.Text := cmbProjectOfReport.Text;
  pgtMain.ActivePage := tbsSendMail;
  with MailMemo.Lines do
  begin
    if German then
    begin
      Add('Sehr geehrter Herr/Frau ' + dm.dsQryDb.DataSet.FieldByName('last_name').AsString);
      Add(' ');
      Add('Wir bestätigen den Eingang der nachfolgend aufgeführten Proben.');
      Add('Dies ist eine automatisch generierte Mail.');
      Add(' ');
      Add('mit freundlichen Grüssen');
      Add(' Das Team des Klaus-Tschira-Labors');
      Add(' ');
    end
    else begin
      Add('Dear ' + dm.dsQryDb.DataSet.FieldByName('first_name').AsString + ' ' +
        dm.dsQryDb.DataSet.FieldByName('last_name').AsString);
      Add(' ');
      Add('We confirm to have received the samples listed below.');
      Add('This is an auto-generated mail.');
      Add(' ');
      Add('best regards');
      Add(' The Team of the Klaus-Tschira-Laboratory');
      Add(' ');
    end;
    with dm.qrySampleOfSubmitter do
    begin
      First;
      while not EOF do
      begin
        s := 'MAMS ' + Fields.Fields[0].AsString; // sample_nr
        for i := 1 to 4 do s := s + '  ' + Fields.Fields[i].AsString;
        Add(s);
        Next;
      end;
    end;
  end;
end;

procedure TfrmMAMS.SetFileName(const Value: string);
begin

end;

procedure TfrmMAMS.SetFormat;
begin
{
  dtStartTaskDate.Perform(DTM_SETFORMAT, 0, integer(PChar('dd.MM.yy')));
  dtStartTaskTime.Perform(DTM_SETFORMAT, 0, integer(PChar('HH:mm')));
  dtEndTaskDate.Perform(DTM_SETFORMAT, 0, integer(PChar('dd.MM.yy')));
  dtEndTaskTime.Perform(DTM_SETFORMAT, 0, integer(PChar('HH:mm')));
  }
end;

procedure TfrmMAMS.SetSampleGridForMaterial;
var
  i, w: integer;
begin
  with grdPreviewSamples do begin
    ColWidths[0] := 30;
    ColWidths[SampleNameCol] := 120;
    ColWidths[SampleNrCol] := 40;
    ColWidths[SampleDescr1Col] := 0;
    ColWidths[SampleDescr2Col] := 0;
    ColWidths[SampleWeightCol] := 0;
    ColWidths[SampleUserMaterialCol] := 80;
    ColWidths[SampleCommentCol] := 100;
    ColWidths[MaterialCol] := 70;
    ColWidths[FractionCol] := 70;
    ColWidths[TypeCol] := 0;
    ColWidths[SamplePrep1Col] := 0;
    ColWidths[SamplePrep2Col] := 0;
    ColWidths[SamplePrep3Col] := 0;
    ColWidths[SamplePrep4Col] := 0;
    ColWidths[SamplePrep5Col] := 0;
  end;
  SetSampleGridWidth;
end;

procedure TfrmMAMS.SetSampleGridForPrep;
begin
  SetSampleGridForMaterial;
  with grdPreviewSamples do begin
    ColWidths[SamplePrep1Col] := 60;
    ColWidths[SamplePrep2Col] := 60;
    ColWidths[SamplePrep3Col] := 60;
    ColWidths[SamplePrep4Col] := 60;
    ColWidths[SamplePrep5Col] := 60;
  end;
  SetSampleGridWidth;
end;

procedure TfrmMAMS.SetSampleGridForType;
begin
  SetSampleGridForMaterial;
  grdPreviewSamples.ColWidths[TypeCol] := 60;
  SetSampleGridWidth;
end;

procedure TfrmMAMS.SetSampleGridWidth;
var
  i, w: integer;
begin
  with grdPreviewSamples do begin
    ColWidths[TypeCol] := 60;
    w := 0;
    for I := 0 to ColCount - 1 do w := w + ColWidths[i] + 1; // 1 = gridline
    Width := w + 6;
  end;
end;

procedure TfrmMAMS.SetTaskEndDate;
var
  i, sample_nr, StepNr, prep_nr: integer;
  Method, s: string;
  d, h: string;
begin
  d := FormatDateTime('YYYY-MM-DD', DateOf(dtENDTaskDate.Date));
  h := FormatDateTime('hh:mm', dtEndTaskTime.Time);
//   SampleNrList := '';
//   SampleNrList := SampleNrList + trim(ExtractWord(1,lbxLabTaskBatch.Items[i],['|'])) + ',';
   // create list of sample_nr of batch
   // insert in prep_t, vorläufig bei prep=1
  prep_nr := 1;
  d := FormatDateTime('YYYY-MM-DD', DateOf(dtEndTaskDate.Date));
  h := FormatDateTime('hh:mm', dtEndTaskTime.Time);
  for i := 0 to lbxPrepSteps.Count - 1 do
    if lbxPrepSteps.Selected[i] then method := lbxPrepSteps.Items[i];
  with dm.qrySamplesOfLabTask do
  begin
    First;
    while not EOF do
    begin
      sample_nr := FieldByName('sample_nr').AsInteger;
      StepNr := dm.GetPretreatStepNrBySampleNr(sample_nr, 1, method); // später auch höhere prepnr
      if (StepNr > 0) and (StepNr < 6) then
      begin
        s := 'UPDATE preparation_t SET step' + IntToStr(StepNr) + '_end=' + #34 + d + ' ' + h + #34 +
          '  WHERE sample_nr=' + IntTostr(sample_nr) + ' AND prep_nr=' + IntTostr(prep_nr) + ';';
//        //   ClibBoard.SetTextBuf(PChar(s));
        dm.adoCmd.CommandText := s;
        LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      end;
      Next;
    end;
  end;
end;

procedure TfrmMAMS.SetTaskStartDate;
var
  i, sample_nr, StepNr, prep_nr: integer;
  Method, s: string;
  d, h: string;
begin
  prep_nr := 1; // anpassen !
  d := FormatDateTime('YYYY-MM-DD', DateOf(dtStartTaskDate.Date));
  h := FormatDateTime('hh:mm', dtStartTaskTime.Time);
  for i := 0 to lbxPrepSteps.Count - 1 do
    if lbxPrepSteps.Selected[i] then method := lbxPrepSteps.Items[i];
  with dm.qrySamplesOfLabTask do begin
    First;
    while not EOF do begin
      sample_nr := FieldByName('sample_nr').AsInteger;
      StepNr := dm.GetPretreatStepNrBySampleNr(sample_nr, 1, method); // später auch höhere prepnr
      if (StepNr > 0) and (StepNr < 6) then begin
        s := 'UPDATE preparation_t SET step' + IntToStr(StepNr) + '_start=' + #34 + d + ' ' + h + #34 +
          '  WHERE sample_nr=' + IntTostr(sample_nr) + ' AND prep_nr=' + IntTostr(prep_nr) + ';';
        //   ClibBoard.SetTextBuf(PChar(s));
        dm.adoCmd.CommandText := s;
        LogWindow.addLogEntry(s);
        IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
        dm.SetProjectStatusRunning(sample_nr);
      end;
      Next;
    end;
  end;
end;

procedure TfrmMAMS.SetupPretreatmentStepsGrid;
begin
end;

procedure TfrmMAMS.SetupProjectPage;
begin
//  dm.qryAllUser.Close;
//  dm.qryAllUser.Open;
  grdProjects.Visible := false;
  gbxSampleLeft.Visible := false;
  gbxSampleRight.Visible := false;
  grdSamplesOfProject.Visible := false;
//  btnSaveModifySample.Enabled := false;
  SetupPretreatmentStepsGrid;
//     spltSampleSamples.Visible := false;
//     spltSamples.Visible  := false;
//     spltSample.Visible := false;
end;

procedure TfrmMAMS.SetValues(const Value: TStrings);
begin
  FValues := Value;
end;

procedure TfrmMAMS.smtpSendMailStatus(ASender: TObject;
  const AStatus: TIdStatus; const AStatusText: string);
begin
  Status(AStatusText);
end;

procedure TfrmMAMS.SpeedButton1Click(Sender: TObject);
// download selected image of the listbox to the local machine
var
  fname: string;
begin
  // get path of the selected image
  fname := ImageFilesListBox.Items[ImageFilesListBox.ItemIndex];
  LogWindow.addLogEntry('download image from: ' + fname);
  // download
  if FileExists(fname) then
    begin
      LogWindow.addLogEntry('source file exists');
      SavePictureDialog.FileName := extractfilename(fname);
      if SavePictureDialog.Execute then
        begin
          //actually copy the file to disk
          LogWindow.addLogEntry('copy file to: ' + SavePictureDialog.FileName);
          Try
            LogWindow.addLogEntry('start copying');
            TFile.Copy(fname, SavePictureDialog.FileName);
            LogWindow.addLogEntry('copying finished');
          Except;
            LogWindow.addLogEntry('problem saving image');
            ShowMessage('Problems saving file.');
          End;
        end;
    end;


end;

procedure TfrmMAMS.btnExchangeClick(Sender: TObject);
Var
  ADOQueryListing: TADOQuery;
  DataSrcListing: TDataSource;
  s: string;
begin
  if RadioGroupSampleExchange.ItemIndex = 0 then // return to sender
    begin
//      s := 'SELECT sample_nr, sample_t.project_nr, project_t. project, project_t.in_date, project_t.out_date, last_name, user_label, MA_nr, AuftragsNr, return_to_sender, returned_to_sender ' +
//           'FROM sample_t ' +
//           'INNER JOIN project_t on sample_t.project_nr = project_t.project_nr ' +
//           'INNER JOIN user_t ON project_t.user_nr = user_t.user_nr ' +
//           'WHERE return_to_sender = "1" ' +
//           'AND NOT returned_to_sender = "1" ' +
//           'Order by sample_nr';

//        s := 'SELECT project, project_nr, last_name, user_t.user_nr, in_date, out_date, AuftragsNr, return_to_sender, returned_to_sender, prep_return_to_sender, prep_returned_to_sender ' +
//             'FROM project_t ' +
//             'INNER JOIN user_t ON project_t.user_nr = user_t.user_nr ' +
//             'WHERE return_to_sender = "1" ' +
//             'AND NOT returned_to_sender = "1" ' +
//             'Order by project_nr';

        s := 'SELECT project, project_nr, last_name, user_t.user_nr, in_date, out_date, AuftragsNr, return_to_sender, returned_to_sender, prep_return_to_sender, prep_returned_to_sender ' +
             'FROM project_t ' +
             'INNER JOIN user_t ON project_t.user_nr = user_t.user_nr ' +
             'WHERE ' +
             '(return_to_sender = 1 AND returned_to_sender = 0) ' +
             'OR ' +
             '(prep_return_to_sender = 1 AND prep_returned_to_sender = 0) ' +
             'Order by project_nr';
    end
  else if RadioGroupSampleExchange.ItemIndex = 1 then //samples to CN
    begin
      s := 'SELECT sample_t.sample_nr, sample_t.project_nr, project_t. project, project_t.in_date, preparation_t.prep_end, project_t.out_date, last_name, user_label, MA_nr, ' +
           'AuftragsNr, CNIsotopA, CNIsotopAMoved ' +
           'FROM sample_t ' +
           'INNER JOIN project_t on sample_t.project_nr = project_t.project_nr ' +
           'INNER JOIN user_t ON project_t.user_nr = user_t.user_nr ' +
           'INNER JOIN preparation_t ON sample_t.sample_nr = preparation_t.sample_nr ' +
           'WHERE CNIsotopA = "1" ' +
           'AND NOT CNIsotopAMoved = "1" ' +
           'Order by sample_nr';;
    end;


  // create a query object
  ADOQueryListing := TADOQuery.Create(Self);
  ADOQueryListing.Connection := dm.ADOConnKTL;
  ADOQueryListing.SQL.Add(s);
  ADOQueryListing.Open;
  // Create the data source.
  DataSrcListing := TDataSource.Create(Self);
  DataSrcListing.DataSet := ADOQueryListing;
  DataSrcListing.Enabled := true;
  //Finally, initialize the grid.
  DBGridSAmpleExchange.DataSource := DataSrcListing;
  // auto scale the column width's
  // DBGridAutoSizeAllColumns(DBGridSAmpleExchange);
  FixDBGridColumnsWidth(DBGridSAmpleExchange);
  // DBGridSampleExchange.EditorMode := True;

end;

procedure TfrmMAMS.DBGridAutoSizeColumn(Grid: TDBGrid; Column: integer);
// determines the longest string in Column of the DBGrid
// and adjusts the Column Width accordingly
var
  W, WMax, WTitle: integer;
begin
  WMax := 0;
  // set datasource to teh first record
  Grid.DataSource.DataSet.First;
  // loop through all records and get longest string of all records of this column
  while not Grid.DataSource.DataSet.EOF do
  begin
    W := Grid.Canvas.TextWidth(Grid.Columns[Column].Field.AsString);
    if W > WMax then WMax := W;
    Grid.DataSource.DataSet.Next;
  end;
  // Column width has to be at least the width of the title
  WTitle := Grid.Canvas.TextWidth(Grid.Columns[Column].Title.Caption);
  //ShowMessage(Grid.Columns[Column].Title.Caption);
  if (WMax + 10) < WTitle then WMax := WTitle;
  Grid.Columns[Column].Width:= WMax + 10;
end;

procedure TfrmMAMS.FixDBGridColumnsWidth(const DBGrid: TDBGrid);
var
i,W,WMax, WTitle, TotWidth, VarWidth, ResizableColumnCount, colCount, MaxTextLength, TextLength : integer;
AColumn : TColumn;
percentChange: Single;
begin
 i := 0;
 DBGrid.DataSource.DataSet.DisableControls;
 //total width of all columns before resize
 TotWidth := 10;
 //how to divide any extra space in the grid
 VarWidth := 5;
 //how many columns need to be auto-resized
 colCount := DBGrid.Columns.Count;
 //calculate total with of all colums according to the text width within the columns
 ResizableColumnCount := 1;
 // maximum length of the text in a column as number of characters - avoid overly long text
 MaxTextLength := 70;

 for i := 0 to colCount-1 do
   begin
       // find max text-width of this column
        LogWindow.addLogEntry('FixDBGridColumnsWidth: find total width, operate on column: ' + inttostr(i));
        JvLogFile.Add('FixDBGridColumnsWidth',lesInformation,'FixDBGridColumnsWidth: find total width, operate on column: ' + inttostr(i));
        WMax := 0;
        DBGrid.DataSource.DataSet.First;
        while not DBGrid.DataSource.DataSet.EOF do
          begin
            W := DBGrid.Canvas.TextWidth(DBGrid.Columns[i].Field.AsString);
            TextLength := length(DBGrid.Columns[i].Field.AsString);
            if W > WMax then
              if TextLength < MaxTextLength then
                begin
                  WMax := W;
                end;
            DBGrid.DataSource.DataSet.Next;
          end;
          LogWindow.addLogEntry('found wmax = ' + inttostr(WMax));
          DBGrid.DataSource.DataSet.First;
       // find with of the column title
          WTitle := DBGrid.Canvas.TextWidth(DBGrid.Columns[i].Title.Caption);
          LogWindow.addLogEntry('found WTitle = ' + inttostr(WTitle));
       // if column title is larger than column content then use width of the column title
          if WMax < WTitle+10 then WMax := WTitle+10;
       // set with of the column to max width of the text or title caption
          DBGrid.Columns[i].Width:= WMax;
      // count total with of all columns
      TotWidth := TotWidth + DBGrid.Columns[i].Width;
     if DBGrid.Columns[i].Field.Tag=0 then Inc(ResizableColumnCount);// tag could be used to define a minimum width
   end;
   LogWindow.addLogEntry('FixDBGridColumnsWidth: Colums to Resize = ' + inttostr(ResizableColumnCount));
 //add 1px for the column separator lineif dgColLines in DBGrid.Options then
 TotWidth := TotWidth + colCount;
 LogWindow.addLogEntry('FixDBGridColumnsWidth: TotalWidth = ' + inttostr(TotWidth));
 //add indicator column width if dgIndicator in DBGrid.Options then
 TotWidth := TotWidth + IndicatorWidth;
  LogWindow.addLogEntry('FixDBGridColumnsWidth: TotalWidth with Indicators = ' + inttostr(TotWidth));
 //width value "left"
 VarWidth := DBGrid.ClientWidth - TotWidth;
 //Equally distribute percentChange
 //to all auto-resizable columnsif ResizableColumnCount > 0 then
 percentChange := (TotWidth/DBGrid.ClientWidth);
 LogWindow.addLogEntry('FixDBGridColumnsWidth: Percentage change of column widths = ' + floattostr(percentChange));
 // showmessage('percentChange = '+ floattostr(percentChange));
 VarWidth := varWidth div ResizableColumnCount;

 // resize all colums according to its percentage change
 LogWindow.addLogEntry('FixDBGridColumnsWidth: start resizing columns');
 for i := 0 to colCount -1 do
   begin
   LogWindow.addLogEntry('FixDBGridColumnsWidth: resize, operate on column: ' + inttostr(i));
   JvLogFile.Add('FixDBGridColumnsWidth',lesInformation,'FixDBGridColumnsWidth: resize, operate on column: ' + inttostr(i));
   AColumn := DBGrid.Columns[i];
   if AColumn.Field.Tag=0 then
     begin
     AColumn.Width := Round(AColumn.Width/percentChange);
     //AColumn.Width := AColumn.Width + VarWidth;
     if AColumn.Width=0 then
     AColumn.Width := AColumn.Field.Tag;
     end;
   end;

 DBGrid.DataSource.DataSet.EnableControls;
end;

procedure TfrmMAMS.JumpToEmptyWeightField;
begin
  // jump to the first weight that is empty
  // check the weights in revers so that eventually the first one is in focus if all of them are empty
  if DBedtTouchWeightsCombustion.Text = '' then
    DBedtTouchWeightsCombustion.SetFocus;
  if DBedtTouchWeightsAfterPrep.Text = '' then
    DBedtTouchWeightsAfterPrep.SetFocus;
  if DBedtTouchWeightsMedium2Prep.Text = '' then
    DBedtTouchWeightsMedium2Prep.SetFocus;
  if DBedtTouchWeightsMediumPrep.Text = '' then
    DBedtTouchWeightsMediumPrep.SetFocus;
  if DBedtTouchWeightsBeforePrep.Text = '' then
    DBedtTouchWeightsBeforePrep.SetFocus;
end;

procedure TfrmMAMS.JvDBDateTimePicker1Click(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.JvDBDateTimePicker2Click(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.JvDBDateTimePicker3Click(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.SetCheckBoxTouchSampleArchived;
Var
test: string;
begin
  LogWindow.addLogEntry('running: SetCheckBoxTouchSampleArchived');
  // set the checkbox "sample archived" depending on the value of the s_storage_location
  test := dbedtsStorLoc.Text;
  //ShowMessage(test);
  if test = '' then
  begin
    chkTouchSampleArchived.Checked := False;
  end
  else
  begin
    chkTouchSampleArchived.Checked := True;
  end;
  // if return_to_sender is true then grey out the archive CheckBox und Archive
  if chkTouchReturnToSender.Checked = True then
  begin
  if test='' then
    begin
      chkTouchSampleArchived.Enabled:=False;
      dbedtsStorLoc.Enabled:=False;
    end
    else begin
      chkTouchSampleArchived.Enabled:=True;
      dbedtsStorLoc.Enabled:=True;
    end;
  end;

end;

procedure TfrmMAMS.SetCheckBoxTouchPrepArchived;
Var
test: string;
begin
  LogWindow.addLogEntry('running: SetCheckBoxTouchPrepArchived');
  // set the checkbox "prep archived" depending on the value of the prep_storage_location
  test := dbedtprepStorLoc.Text;
  //ShowMessage(test);
  if test = '' then
  begin
    chkTouchPrepArchived.Checked := False;
  end
  else
  begin
    chkTouchPrepArchived.Checked := True;
  end;

    // if prep_return_to_sender is true then grey out the archive CheckBox und Archive
  if chkTouchprepReturnToSender.Checked = True then
  begin
  if test='' then
    begin
      chkTouchPrepArchived.Enabled:=False;
      dbedtprepStorLoc.Enabled:=False;
    end
    else begin
      chkTouchPrepArchived.Enabled:=True;
      dbedtprepStorLoc.Enabled:=True;
    end;
  end;

end;

procedure TfrmMAMS.DisplayArchiveWarning;
VAR
  CheckedCount: integer;
begin
  // this meant to make sure that samples are archived if they are not returned or no leftover
  // check if at least one of the checkboxes "return to sender" or " s no leftover" or "s_storage_loc"/"SampleArchived" are checked
  // if not show a SpeedButton displaing a warning symbol
  LogWindow.addLogEntry('running: DisplayArchiveWarning');
  // only one checkmark MUST be present
  CheckedCount:=0;
  if DBchkTouchWeightsSampleNoLeftover.Checked then Inc(CheckedCount);
  if chkTouchReturnToSender.Checked then Inc(CheckedCount);
  if chkTouchSampleArchived.Checked then Inc(CheckedCount);
  // ç
  if (CheckedCount = 0) or (CheckedCount > 1) then
  begin
    btnTouchWarningStorageLoc.Visible := True;
  end
  else
  begin
    btnTouchWarningStorageLoc.Visible := False;
  end;
end;

procedure TfrmMAMS.DisplayArchiveWarningPrep;
VAR
  CheckedCount: integer;
begin
  // this meant to make sure that preped materials are archieved if they are not moved to CN or no leftover
  // check if at least one of the checkboxes "moved to CN" or " s no leftover" or "prep_storage_loc"/"SampleArchived" or "prep_return_to_sender" are checked
  // if not show a SpeedButton displaing a warning symbol
  LogWindow.addLogEntry('running: DisplayArchiveWarningPrep');
  // only one checkmark MUST be present

    CheckedCount := 0;

  // Check how many checkboxes are checked
  if chkTouchprepReturnToSender.Checked then Inc(CheckedCount);
  if DBCheckBoxTouchMovedToCN.Checked then Inc(CheckedCount);
  if chkTouchPrepArchived.Checked then Inc(CheckedCount);
  if DBchkTouchWeightsPrepNoLeftover.Checked then Inc(CheckedCount);

  // showMessage(CheckedCount.ToString);

    // Logic for the visibility of the speed button
  if (CheckedCount = 0) or (CheckedCount > 1) then
    btnTouchWarningStorageLocPrep.Visible := True
  else
    btnTouchWarningStorageLocPrep.Visible := False;

end;

procedure TfrmMAMS.DBchkTouchWeightsPrepNoLeftoverClick(Sender: TObject);
begin
    btnTouchWeightsGraphNeedsSaving.Visible := True;

    SetCheckBoxTouchPrepArchived;
    DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.DBchkTouchWeightsSampleNoLeftoverClick(Sender: TObject);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

  SetCheckBoxTouchSampleArchived;
  // check if necessary number of checkboxes is checked, if not display the ArchiveWarningButton
  DisplayArchiveWarning;
end;

procedure TfrmMAMS.DBDateTimeTouchPrepEndClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.DBDateTimeTouchPrepEndKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.DBDateTimeTouchPrepStartChange(Sender: TObject);
begin
  DBDateTimeTouchPrepStart.Update;
end;

procedure TfrmMAMS.DBDateTimeTouchPrepStartClick(Sender: TObject);
begin
  // if no date is set, set it to today
  if (Sender AS TJvDBDateTimePicker).Date = 0 then (Sender AS TJvDBDateTimePicker).Date := Now;
end;

procedure TfrmMAMS.DBDateTimeTouchPrepStartKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.DBEditWeightMedium2Change(Sender: TObject);
begin
  WeightsChanged := true;
  CalculateNetWeightEnd;
end;

procedure TfrmMAMS.DBEditWeightMediumChange(Sender: TObject);
begin
  WeightsChanged := true;
  CalculateNetWeightEnd;
end;

procedure TfrmMAMS.DBedtHECurrentChange(Sender: TObject);
begin
// change background color when value below threshold
  //if (Sender as TDBEdit).DataSource.DataSet.FieldByName('he_curr').Value < 20 then (Sender as TDBEdit).Color := clRed;
end;

procedure TfrmMAMS.dbedtprepStorLocChange(Sender: TObject);
begin
  // update checkbox Sample Archived
  SetCheckBoxTouchPrepArchived;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.dbedtprepStorLocClick(Sender: TObject);
begin
  btnTouchWeightsGraphNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.dbedtprepStorLocExit(Sender: TObject);
begin
  if dbedtprepStorLoc.Text <> '' then
  begin
     prepStorLocSuggestion := dbedtprepStorLoc.Text;
  end;

  //showmessage(dbedtprepStorLoc.Text);
end;

procedure TfrmMAMS.dbedtprepStorLocMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if dbedtprepStorLoc.Text = '' then
    begin
       if prepStorLocSuggestion <> '' then
       begin
        dbedtprepStorLoc.DataSource.DataSet.Edit;
        dbedtprepStorLoc.Text := prepStorLocSuggestion;
       end;
    end;
end;

procedure TfrmMAMS.dbedtsStorLocChange(Sender: TObject);
begin
  // update checkbox Sample Archived
  SetCheckBoxTouchSampleArchived;
  DisplayArchiveWarning;
end;

procedure TfrmMAMS.dbedtsStorLocClick(Sender: TObject);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;
end;

procedure TfrmMAMS.dbedtsStorLocExit(Sender: TObject);
begin

  if dbedtsStorLoc.Text <> '' then
  begin
     sStorLocSuggestion := dbedtsStorLoc.Text;
  end;

  //showmessage(dbedtsStorLoc.Text);
end;

procedure TfrmMAMS.dbedtsStorLocMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if dbedtsStorLoc.Text = '' then
    begin
       if sStorLocSuggestion <> '' then
       begin
        dbedtsStorLoc.DataSource.DataSet.Edit;
        dbedtsStorLoc.Text := sStorLocSuggestion;
       end;
    end;
end;

procedure TfrmMAMS.DBedtTouchWeightsAfterPrepChange(Sender: TObject);
begin
  CalculateYieldTouch; // calculate the yield from the weights
end;

procedure TfrmMAMS.DBedtTouchWeightsAfterPrepClick(Sender: TObject);
begin
  DBedtTouchWeightsAfterPrep.SelectAll;
end;

procedure TfrmMAMS.DBedtTouchWeightsAfterPrepKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

  // if date_end is not set yet, set the current date as soon
  // as the weight_end is entered
  // if desired (bollean switch) convert the entered number from g to mg
  if Key = vk_Return then
    begin
      if DBDateTimeTouchPrepEnd.IsNull then
        begin
          DBDateTimeTouchPrepEnd.Date := Now;
        end;
      // convert from g to mg
      if CheckBoxTouchPrepWeightsAutoConversion.Checked then
        Begin
          (Sender AS TDBEdit).Field.Text := floattostr( strtofloat((Sender AS TDBEdit).Text) * 1000 );
        End;
    end;
end;

  procedure TfrmMAMS.DBedtTouchWeightsBeforePrepChange(Sender: TObject);
begin
  CalculateYieldTouch; // calculate the yield from the weights
end;

procedure TfrmMAMS.DBedtTouchWeightsBeforePrepClick(Sender: TObject);
begin
  DBedtTouchWeightsBeforePrep.SelectAll;
end;

procedure TfrmMAMS.DBedtTouchWeightsBeforePrepKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

  // if date_start is not set yet, set the current date as soon
  // as the weight_start is entered
  // if desired (bollean switch) convert the entered number from g to mg
  if Key = vk_Return then
    begin
      if DBDateTimeTouchPrepStart.IsNull then
        begin
          DBDateTimeTouchPrepStart.Date := Now;
        end;
      // convert from g to mg
      if CheckBoxTouchPrepWeightsAutoConversion.Checked then
        Begin
          (Sender AS TDBEdit).Field.Text := floattostr( strtofloat((Sender AS TDBEdit).Text) * 1000 );
        End;
    end;

end;

procedure TfrmMAMS.DBedtTouchWeightsCombustionClick(Sender: TObject);
begin
  DBedtTouchWeightsCombustion.SelectAll;
end;

procedure TfrmMAMS.DBedtTouchWeightsCombustionKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsGraphNeedsSaving.Visible := True;

  // if desired (bollean switch) convert the entered number from g to mg
  if Key = vk_Return then
    begin
      // convert from g to mg
      if CheckBoxTouchGraphWeightsAutoConversion.Checked then
        Begin
          (Sender AS TDBEdit).Field.Text := floattostr( strtofloat((Sender AS TDBEdit).Text) * 1000 );
        End;
    end;
end;

procedure TfrmMAMS.DBedtTouchWeightsMedium2PrepChange(Sender: TObject);
begin
  // calulate net weight of weight_end
  CalculateNetWeightEndTouch;
end;

procedure TfrmMAMS.DBedtTouchWeightsMedium2PrepKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

  if Key = vk_Return then
    begin
      // convert from g to mg
      if CheckBoxTouchPrepWeightsAutoConversion.Checked then
        Begin
          (Sender AS TDBEdit).Field.Text := floattostr( strtofloat((Sender AS TDBEdit).Text) * 1000 );
        End;
    end;

end;

procedure TfrmMAMS.DBedtTouchWeightsMediumPrepChange(Sender: TObject);
begin
  // calulate net weight of weight_end
  CalculateNetWeightEndTouch;
end;

procedure TfrmMAMS.DBedtTouchWeightsMediumPrepKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  btnTouchWeightsPrepNeedsSaving.Visible := True;

    if Key = vk_Return then
    begin
      // convert from g to mg
      if CheckBoxTouchPrepWeightsAutoConversion.Checked then
        Begin
          (Sender AS TDBEdit).Field.Text := floattostr( strtofloat((Sender AS TDBEdit).Text) * 1000 );
        End;
    end;

end;

procedure TfrmMAMS.DBGridAutoSizeAllColumns(Grid: TDBGrid);
// autosizes all columns of the DBGrid
var
  Column, W, WMax: integer;
begin
  if Grid.Columns.Count = 0 then Exit;
  // disable controls in order to speed up screen update
  Grid.DataSource.DataSet.DisableControls;
  // loop through all colums and adkust width
  for Column:=0 to (Grid.Columns.Count-1) do
    begin
      DBGridAutoSizeColumn(Grid, Column);
    end;
  // enable controls again
  Grid.DataSource.DataSet.EnableControls;
end;

procedure TfrmMAMS.Status(const AMsg: string);
begin
  lboxStatus.ItemIndex := lboxStatus.Items.Add(AMsg);
end;

procedure TfrmMAMS.StrGrdTargetDataClick(Sender: TObject);
var
  aCol, aRow: Integer;
  grid: TJvStringGrid;
  pt, mouse: TPoint;
begin
  edtMeanAge.Clear;
  edtSigmaAge.Clear;
  edtC13Mean.Clear;
  btnTransferTargetData.Enabled := false;
  GetCursorPos(pt);
  grid := Sender as TJvStringgrid;
  pt := grid.ScreenToClient(pt);
  grid.MouseToCell(pt.x, pt.y, aCol, aRow);
  if (aCol = 7) and (aRow >= grid.fixedRows) then
  // click landed in a checkbox cell
    ToggleCheckbox(aCol, aRow);
end;

procedure TfrmMAMS.StrGrdTargetDataDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  s: string;
begin
  if not (gdFixed in State) and (aCol = 7) then begin
    with (Sender as TJvStringgrid).Canvas do begin
      brush.color := $E0E0E0;
      // checkboxes look better on a non-white background
      Fillrect(rect);
      if Assigned(StrGrdTargetData.Objects[aCol, aRow]) then
        Draw((rect.right + rect.left - FCheck.width) div 2,
          (rect.bottom + rect.top - FCheck.height) div 2,
          FCheck)
      else
        Draw((rect.right + rect.left - FNoCheck.width) div 2,
          (rect.bottom + rect.top - FNoCheck.height) div 2,
          FNoCheck)
    end;
  end;
end;

procedure TfrmMAMS.tbsTypesMatExit(Sender: TObject);
begin
  dm.tblMaterial.Close;
  dm.tblMaterial.Open;
  dm.tblTypes.Close;
  dm.tblTypes.Open;
  dm.tblMethod.Close;
  dm.tblMethod.Open;
end;

procedure TfrmMAMS.ToggleCheckbox(acol, arow: integer);
begin
  if aCol = 7 then
    with TGridCracker(StrGrdTargetData) do begin
      if Assigned(Objects[aCol, aRow]) then
        Objects[aCol, aRow] := nil
      else
        Objects[aCol, aRow] := Pointer(1);
      InvalidateCell(aCol, aRow);
    end;
end;

procedure TfrmMAMS.btnDBPlotClick(Sender: TObject);
// switch to the Plot tab of the main tab
begin
  pgtMain.ActivePage := tbsDBPlot;
end;

procedure TfrmMAMS.btnDBPlotLoadQueryClick(Sender: TObject);
// load query from file to query memo
var
  OpenDialog: TOpenDialog;
begin
  OpenDialog := TOpenDialog.Create(self);
  OpenDialog.Title:='Open Query File';
  OpenDialog.InitialDir := GetCurrentDir;
  OpenDialog.Filter := 'Text files (*.txt)|*.TXT|Any file (*.*)|*.*';
  OpenDialog.DefaultExt := 'txt';
  OpenDialog.FilterIndex := 1;

  // Execute an open file dialog.
  if OpenDialog.Execute then
    //First check if the file exists.
    if FileExists(OpenDialog.FileName) then
      //If it exists, load the data into the memo box. }
      MemoDBPlotQuery.Lines.LoadFromFile(OpenDialog.FileName)
    else
      { Otherwise, raise an exception. }
      raise Exception.Create('File does not exist.');
end;

procedure TfrmMAMS.btnDBPlotSaveQueryClick(Sender: TObject);
// save query to text file
var
  SaveDialog: TSaveDialog;
begin
  SaveDialog := TSaveDialog.Create(self);
  SaveDialog.Title:='Save Query to File';
  saveDialog.InitialDir := GetCurrentDir;
  SaveDialog.Filter := 'Text files (*.txt)|*.TXT|Any file (*.*)|*.*';
  saveDialog.DefaultExt := 'txt';
  saveDialog.FilterIndex := 1;
  saveDialog.Options := [ofOverwritePrompt];

  // Execute an open file dialog.
  if SaveDialog.Execute then begin
      MemoDBPlotQuery.Lines.SaveToFile(SaveDialog.FileName);
  end;
  saveDialog.Free;
end;

procedure TfrmMAMS.ToolButton10Click(Sender: TObject);
// switch to the DBInfo tab of the main tab
begin
  pgtMain.ActivePage := tbsDBInfo;
//  btnPendingReports.Click; // perform click event

 //set the year for calculating numbers of processed samples to this-year
 edtMonthStat.Value := Yearof(today);
end;

procedure TfrmMAMS.ToolButton6Click(Sender: TObject);
begin
  pgtMain.ActivePage := tbsAccounting;
end;

procedure TfrmMAMS.ToolButtonPressClick(Sender: TObject);
begin
  pgtMain.ActivePage := tbsMagazine;
    GetListOfUnpressedMagazines(1);
    grdMagazinesUnpressed.Columns[0].Width:=grdMagazinesUnpressed.Width; //adjust column width to match table width
end;

procedure TfrmMAMS.ToolButtonSampleExchangeClick(Sender: TObject);
begin
  pgtMain.ActivePage := SampleExchange;
  //ShowMessage(inttostr(RadioGroupSampleExchange.ItemIndex));
  if RadioGroupSampleExchange.ItemIndex = -1 then RadioGroupSampleExchange.ItemIndex := 1;
  btnExchangeClick(self);
end;

procedure TfrmMAMS.ToolButtonTouchClick(Sender: TObject);
Var SampleNr, NPreps, NTargets: Integer;
begin
  pgtMain.ActivePage := tbsTouch;

  edtTouchWeightsMAMS.SetFocus;
  SampleNr := edtTouchWeightsMAMS.Value;
  // Retrieve all SampleInfo from the DB (same function as in tab:SampleInfo)
  NPreps := dm.GetMaxPrepNrBySampleNr(SampleNr);
  NTargets := dm.GetMaxTargetNrBySampleNr(SampleNr, NPreps);
  edtTouchWeightsPrepNr.MaxValue := NPreps;
  //edtTouchWeightsPrepNr.Value := NPreps;
  lblTouchWeightsNPrep.Caption := '1...' + IntToStr(NPreps);
  edtTouchWeightsTargetNr.MaxValue := NTargets;
  //edtTouchWeightsTargetNr.Value := NTargets;
  lblTouchWeightsNTargets.Caption := '1...' + IntToStr(NTargets);
  // also update the editButtons in tab:SampleInfo to reflect the same values
  // so that DoSampleInfo can be called
  edtSampleNr.Value := edtTouchWeightsMAMS.Value;
  edtSamplePrepNr.Value := edtTouchWeightsPrepNr.Value;
  edtSampleTargetNr.Value := edtTouchWeightsTargetNr.Value;

  // get all sample info from DB
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));

  // hide some buttons
  btnTouchWeightsPrepNeedsSaving.Visible := False;
  btnTouchWeightsGraphNeedsSaving.Visible := False;
  // btnTouchWeightsGraphBatchNeedsSaving.Visible := False;

  // jump to the first weight that is empty
  //JumpToEmptyWeightField;

  // set focus on MAMS number in order to be ready to scan barcode
    edtTouchWeightsMAMS.SetFocus;
    edtTouchWeightsMAMS.SelectAll;

  // calculate the yield from the weights
    CalculateYieldTouch;

    btnTouchWarningStorageLoc.Visible:=False;
  // update checkbox Sample Archived
  sleep(200);
    SetCheckBoxTouchSampleArchived;
    SetCheckBoxTouchPrepArchived;
  sleep(200);
  DisplayArchiveWarning;
  DisplayArchiveWarningPrep;
end;

procedure TfrmMAMS.UpdateUser(const LastName: string);
var
  i: integer;
  s: string;
begin
  with grdPreviewUser do begin
    s := 'UPDATE user_t SET ' +
      ' first_name=' + #34 + cells[1, 1] + #34 + ',' +
      ' organisation=' + #34 + cells[1, 3] + #34 + ',' +
      ' institute=' + #34 + cells[1, 4] + #34 + ',' +
      ' address_1=' + #34 + cells[1, 5] + #34 + ',' +
      ' address_2=' + #34 + cells[1, 6] + #34 + ',' +
      ' town=' + #34 + cells[1, 7] + #34 + ',' +
      ' postcode=' + #34 + cells[1, 8] + #34 + ',' +
      ' country=' + #34 + cells[1, 9] + #34 + ',' +
      ' phone_1=' + #34 + cells[1, 10] + #34 + ',' +
      ' phone_2=' + #34 + cells[1, 11] + #34 + ',' +
      ' fax=' + #34 + cells[1, 12] + #34 + ',' +
      ' email=' + #34 + cells[1, 13] + #34 + ',' +
      ' www=' + #34 + cells[1, 14] + #34 + ',' +
      'invoice=0, correspondance=1 ' +
      ' WHERE last_name=' + #34 + LastName + #34 + ';';
  end;
  for i := 1 to 10 do s := ReplaceStr(s, '""', 'NULL');
  //   ClibBoard.SetTextBuf(PChar(s));
  dm.adoCmd.CommandText := s;
  LogWindow.addLogEntry(s);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          dm.adoCmd.Execute;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
end;

procedure TfrmMAMS.GetPlanned;
//display alle the planned samples:
// samples before prep and before graphitisation
var
  n: integer;
begin

  //The width of the grid less the space ocupied by the scrollbar is
  //DBGrid.ClientWidth.
  //The width of each column is DBGrid.Columns[n].Width.

  n := dm.GetAllPlanned(chkShowOnHold.Checked, 'none'); // get all samples that are not prep'd yet
  with grdPlanned do
    begin
      Columns[0].Width := 40; // sample_nr
      Columns[1].Width := 90; // project
      Columns[2].Width := 90; // user_label
      Columns[3].Width := 90; // last_name
      Columns[4].Width := 60; // desired_date
    end;
  n := n + dm.GetAllInPrep;                      // get all samples that are in prep
  with grdInPrep do
    begin
      Columns[0].Width := 45; // sample_nr
      Columns[1].Width := 100; // user_label
      Columns[2].Width := 100; // project
      Columns[3].Width := 50; // material
      Columns[4].Width := 80; // lastname
      Columns[5].Width := 90; // desired date
      Columns[6].Width := 30; // days in prep
      Columns[7].Width := 50; // sample in freezer?
    end;
  n := n + dm.GetAllWaitingForGraph;            // get all samples that are prep'd but not graphitized
  with grdWaitingForGraph do
    begin
      Columns[0].Width := 50; // sample_nr
      Columns[1].Width := 110; // user_label
      Columns[2].Width := 110; // project
      Columns[3].Width := 90; // last_name
      Columns[4].Width := 70; // first_name
      Columns[5].Width := 70; // date
    end;
  lblTotal.Caption := 'Total = ' + IntToStr(n) + ' samples (' + IntToStr((n div 80) + 1) + ' weeks)'; //calculate total time in weeks till all is finished, 80 samples per week?
end;

procedure TfrmMAMS.CheckBox1Click(Sender: TObject);
begin
  dm.tblUser.Edit;
end;

procedure TfrmMAMS.CheckProjectExists;
var
  s: string;
begin
  with dm.qryDB do
  begin
    Close;
    SQL.Text := 'Select project FROM project_t ' + ' INNER JOIN user_t ON' +
      '(project_t.user_nr = user_t.user_nr) ' + ' WHERE project=' + '"' +
      ProjectName + '"' + ' AND user_t.user_nr=' + IntToStr(glbUserNr) + ';';
    s := SQL.Text;
    //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    ProjectExists := false;
    if RecordCount > 0 then
    begin
      //      ShowMessage(' Sample already exists ');
      ProjectExists := true;
    end;
  end;
  //      ShowMessage('New Sample');
end;

procedure TfrmMAMS.GetSamplesForReport(DoExportHalle: boolean);
//query the database in order to retrieve all samples of the selected user and project
var
  s1: string;
  ProjectNrList: string;
  s: string;
  s2: string;
begin
  dm.qrySampleOfSubmitter.Close;
  s1 := FormatDateTime('YYYY-MM-DD', EncodeDate(round(edtStartYear.Value), 1, 1));
  s2 := FormatDateTime('YYYY-MM-DD', EncodeDate(round(edtEndYear.Value), 12, 31));
  with dm.qryProjectsOfUser do
  begin     //get all projects of the selected user
    if chkAllProjects.Checked then
    begin
      SQL.Text := 'SELECT project_nr, in_date, desired_date, project,invoice FROM project_t ' +
        'WHERE user_nr=' + IntToStr(cmbSubmNameReport.KeyValue) +
        ' AND (in_date BETWEEN ' + '"' + s1 + '"' + ' AND ' + '"' + s2 + '"' + ');';
      s := SQL.Text;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
      if RecordCount > 0 then
      begin
        ProjectNrList := '(';
        First;
        while not EOF do
        begin
          ProjectNrList := ProjectNrList + dm.qryProjectsOfUser.Fields.Fields[0].AsString + ',';
          Next;
        end;
        ProjectNrList := copy(ProjectNrList, 1, length(ProjectNrList) - 1);
      // trim final comma
        ProjectNrList := ProjectNrList + ')';
      end;
    end
    else
    begin
      if Length(cmbProjectOfReport.Text) > 0 then
      begin
        ProjectNrList := '(' + IntToStr(cmbProjectOfReport.KeyValue) + ')';
        lblReport.Caption := IntToStr(cmbProjectOfReport.KeyValue);
        dm.qryProjectsOfUser.SQL.Text := 'SELECT AuftragsNr, project_nr, project, in_date, desired_date FROM project_t ' +
          'WHERE project_nr=' + IntToStr(cmbProjectOfReport.KeyValue) + ';';
        s := SQL.Text;
        //ClipBoard.SetTextBuf(PChar(s));
        dm.qryProjectsOfUser.Open;
        // get and save order number etc etc
        if dm.qryProjectsOfUser.RecordCount > 0 then
        begin
          OrderNr := dm.qryProjectsOfUser.Fields.Fields[0].AsInteger;
          ProjectNameReport := dm.qryProjectsOfUser.Fields.Fields[2].AsString;
          ProjectInDate := dm.qryProjectsOfUser.Fields.Fields[3].AsString;
          //ShowMessage(IntToStr(OrderNr) + ProjectNameReport + ProjectInDate);
        end;
      end
      else
      begin
        ShowMessage('No user project selected');
        Exit;
      end;
    end;
  end;

  // query database so get the samples for the selected project
  // in the old version ages and fm values where pulled from the the samples table
  // Thus, multiple targets or preps only showed up as one result
  //
  // Now the results are pulled from the target table so that every prep and every
  // target is shown.
  // Halle Export is still using the old way, displaying only the samples_t
  if dm.qryProjectsOfUser.RecordCount > 0 then
  begin
    with dm.qrySampleOfSubmitter do
    begin
      Close;
      SQL.Clear;
      if not DoExportHalle then
      begin
        s := 'SELECT DISTINCT sample_t.sample_nr, user_label,user_label_nr,user_desc1,' +
          'user_desc2, target_t.C14_age, target_t.C14_age_sig, target_t.dc13*1000 as dc13, ' +
          ' target_t.Cal1sMin, target_t.Cal1sMax, target_t.Cal2sMin, target_t.Cal2sMax,' +
          ' project_t.project,target_t.fm AS av_fm, target_t.fm_sig AS av_fm_sig,' +
          ' conc_c/conc_n*14/12 as cn_ratio, conc_c, weight_end/weight_start*100 AS collpc,' +
          ' target_t.magazine, sample_t.material, sample_t.fraction, preparation_t.prep_nr, target_t.target_nr, '+
          ' preparation_t.weight_start, preparation_t.weight_end, round(preparation_t.weight_end/preparation_t.weight_start*100,0) AS yield, '+
          ' preparation_t.stop AS Prep_Discarded, target_t.stop AS Graph_Discarded, round((av_fm*100),2) as pmC, round((av_fm_sig * 100),2) AS pmC_err';
        if (radgrpStatus.ItemIndex = 0) and chkLECurrent.Checked then s := s + ',ANA';
        s := s +
          ' FROM sample_t  ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
          'INNER JOIN target_t ON (target_t.sample_nr=preparation_t.sample_nr AND target_t.prep_nr=preparation_t.prep_nr) ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ';
        if (radgrpStatus.ItemIndex = 0) and chkLECurrent.Checked  then
          s := s + 'INNER JOIN db_dc14.workfinal ON db_dc14.workfinal.sample_nr=sample_t.sample_nr ';
        s := s +
          ' WHERE sample_t.project_nr IN ' + ProjectNrList + // ' AND target_t.stop=0 ' +
          ' AND ( sample_t.sample_nr >= ' +
          IntToStr(round(edtStartAna.Value)) + ')' + ' AND ( sample_t.sample_nr <= ' +
          IntToStr(round(edtEndAna.Value)) + ')';
        if radgrpStatus.ItemIndex < 3 then
        begin
          if radgrpStatus.ItemIndex = 0 then
          begin
            s := s + ' AND  ( target_t.C14_age IS NOT NULL)';
          end
          else
          begin
            if radgrpStatus.ItemIndex = 2 then
              s := s + ' AND  ( target_t.C14_age IS NULL) AND NOT (not_tobedated) '
            else
              s := s + ' AND  ( target_t.C14_age IS NULL) ';
          end;
        end;
        SQL.Add(s + ' ORDER BY  sample_t.sample_nr');
        LogWindow.addLogEntry(s);
      end
      else
      begin
      // Halle export:
      // Probennr. Labor,	BP-Ergebnis,	+/-,	delta13C, Datierungsverfahren, Bearbeiter, Calibrierungsprogram,
      // cal max 1sigma,	cal min 1sigma, 	cal max 2sigma,	cal min 2sigma,
      // Kollagengehalt(%),	Kohlenstoffmenge, C:N ratio,	F14C, Fraktion, Zuverlaessigkeit,	Tag der Datierung
        s := 'SELECT DISTINCT sample_t.sample_nr, ' +
          ' target_t.C14_age, target_t.C14_age_sig, av_dc13*1000 as dc13, ' +
          #34 + '14C AMS' + #34 + ' AS DatVerfahren, ' +
          #34 + 'n.a.' + #34 + ' AS Bearbeiter, ' +
          #34 + 'SwissCal' + #34 + ' AS CalProgr, ' +
          ' target_t.Cal1sMax,  target_t.Cal1sMin, target_t.Cal2sMax, target_t.Cal2sMin,' +
          ' weight_end/weight_start*100 AS collpc,' +
          ' weight_combustion * conc_c/100 AS c, ' +
          ' conc_c/conc_n*14/12 as cn_ratio, av_fm, ' +
          #34 + 'n.a.' + #34 + ' AS Fraktion, ' +
          #34 + 'n.a.' + #34 + ' AS Zuverl, ' +
          'concat(substring(target_t.magazine, 7, 2),".",substring(target_t.magazine, 5, 2),".",substring(target_t.magazine, 3, 2)) AS MessDatum ';
        s := s +
          ' FROM sample_t  ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
          'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ';
        s := s +
          ' WHERE sample_t.project_nr IN ' + ProjectNrList + // ' AND target_t.stop=0 ' +
          ' AND ( sample_t.sample_nr >= ' +
          IntToStr(round(edtStartAna.Value)) + ')' + ' AND ( sample_t.sample_nr <= ' +
          IntToStr(round(edtEndAna.Value)) + ')';
        SQL.Add(s + ' ORDER BY  sample_t.sample_nr');
        LogWindow.addLogEntry(s);
      end;

    //    SortOrderStr := 'SampNr';
 //     ClipBoard.SetTextBuf(PChar(s));
      DisableControls;
      try
        Open;
      finally
        EnableControls;
      end;
      TFloatField(FieldByName('c14_age')).DisplayFormat := '0';
      TFloatField(FieldByName('c14_age')).DisplayWidth := 5;
      TFloatField(FieldByName('c14_age_sig')).DisplayFormat := '0';
      TFloatField(FieldByName('av_fm')).DisplayFormat := '0.0000';
      if rgpExport.ItemIndex <> 2 then TFloatField(FieldByName('av_fm_sig')).DisplayFormat := '0.0000'; // not for LD Halle
      TFloatField(FieldByName('dc13')).DisplayFormat := '0.0';
      if rgpExport.ItemIndex <> 2 then TFloatField(FieldByName('conc_c')).DisplayFormat := '0.0';
      //if radgrpStatus.ItemIndex = 2 then TFloatField(FieldByName('conc_n')).DisplayFormat := '0.0'; // if LDA Halle is selected as export, don't use conc_n (doesn't excist)
      if rgpExport.ItemIndex = 2 then TFloatField(FieldByName('c')).DisplayFormat := '0.0'; // if LDA Halle is selected as export
      TFloatField(FieldByName('cn_ratio')).DisplayFormat := '0.0';
      TFloatField(FieldByName('collpc')).DisplayFormat := '0.0';
      if (radgrpStatus.ItemIndex = 0) and chkLECurrent.Checked then TFloatField(FieldByName('ANA')).DisplayFormat := '0.0';
    end;
    with grdSamplesOfSubmitter do
    begin
      Columns[0].Width := 55;
      //Columns[1].Width := 55;
      Columns[1].Width := 140;
      Columns[2].Width := 100;
      Columns[3].Width := 80;
      Columns[4].Width := 50;
      Columns[5].Width := 50;
      Columns[6].Width := 50;
      Columns[7].Width := 50;
      Columns[8].Width := 50;
      Columns[9].Width := 50;
      Columns[10].Width := 50;
      Columns[11].Width := 50;
      Columns[12].Width := 150;
      Columns[13].Width := 40;
      if not DoExportHalle then begin
        Columns[14].Width := 40;
        Columns[15].Width := 45;
        Columns[16].Width := 40;
        Columns[17].Width := 40;
        Columns[18].Width := 80;
        Columns[19].Width := 60; // material
        Columns[20].Width := 60; // fraction
        Columns[21].Width := 40; // prep_nr
        Columns[22].Width := 50; // target_nr
      end;
      if (radgrpStatus.ItemIndex = 0) and chkLECurrent.Checked then Columns[19].Width := 40;
    end;
  end;
  // FixDBGridColumnsWidth(grdSamplesOfSubmitter);
end;

procedure TfrmMAMS.AlternateRowColors(Sender: TObject; State: TGridDrawState);
var
  row: Integer;
begin
  // alternate row colors
  if Sender is TDBGrid then
  Begin
    row := (Sender as TDBGrid).DataSource.DataSet.RecNo;
    if Odd(row) then
    begin
      (Sender as TDBGrid).canvas.Brush.Color := (Sender as TDBGrid).GradientStartColor;
    end
    else
    begin
      (Sender as TDBGrid).Canvas.Brush.Color := (Sender as TDBGrid).FixedColor;
    end;
    // now that we have changed the colors of the rows
    //we need to deal with the appearance of selected rows
    if (gdSelected in State) then
       begin
          with (Sender as TDBGrid).Canvas do
          begin
            Brush.Color := clHighlight;
            Font.Style := Font.Style + [fsBold];
            Font.Color := clHighlightText;
          end;
       end;
  End;

end;

procedure TfrmMAMS.ApplicationEventsException(Sender: TObject; E: Exception);
begin
 // use this to automatically catch excrptions and log them away
 //
 // write Exception into logfile as Errors (lesError)
 JvLogFile.Add('Exception',lesError,'UnitName: '+ E.UnitName + ' -- ClassName: ' + E.ClassName + ' -- StackTrace: ' + E.StackTrace + ' -- Message: ' + E.Message);
 // since Exceptions are now caught atomatically they are not shown anymore by the app
 // tell App to show excption
 Application.ShowException(E);

end;

procedure TfrmMAMS.CalculateCAbsWeight;
// calulcate and display the absolute amount of carbon when C% and weight_start is known
VAR CAbsWeight, MinCAllowed: double;
begin
  LogWindow.addLogEntry('CalculateCAbsWeight');

  // read from the OptionsTab (iniFile)
  MinCAllowed := StrToIntDef(edtOptionsMinCAmount.Text, 200)/1000;   // convert string to int, if it fails use 200 as default
  LogWindow.addLogEntry('CalculateCAbsWeight -- MinCAllowed = ' + edtOptionsMinCAmount.Text + ' ug');
  //showmessage(MinCAllowed.ToString);

  if not (edtWeightCombustion.Text = '') AND not (edtSampleCPercent.Text = '') then
  begin
    CAbsWeight := SimpleRoundTo((strtofloat(edtWeightCombustion.Text) * strtofloat(edtSampleCPercent.text)/100),-2);
    lblCAbsWeight.Caption := floattostr(CAbsWeight) + ' mg';

    if CAbsWeight < MinCAllowed then
      begin
        // set fontsytle to bold
        lblCAbsWeight.Font.Style := lblCAbsWeight.Font.Style + [fsBold];
        // remove the themed font color first
        lblCAbsWeight.StyleElements := lblCAbsWeight.StyleElements - [seFont];
        // set the new font color
        lblCAbsWeight.Font.Color := clRed;
      end
      else
      begin
        // remove the fontsytle bold
        lblCAbsWeight.Font.Style := lblCAbsWeight.Font.Style - [fsBold];
        // reset the default themed font color
        lblCAbsWeight.StyleElements := lblCAbsWeight.StyleElements + [seFont];
      end;
  end
  else
  //display yield in percent from weights and round
  begin
    lblCAbsWeight.Caption := '... mg';

    // remove the fontsytle bold
    lblCAbsWeight.Font.Style := lblCAbsWeight.Font.Style - [fsBold];
    // reset the default themed font color
    lblCAbsWeight.StyleElements := lblCAbsWeight.StyleElements + [seFont];
  end;
end;

procedure TfrmMAMS.CalculateYield;
// calulcates the yield from the values used in the SampleInfo Tab
VAR yield: double;
begin
  LogWindow.addLogEntry('CalculateYield');
  if not (edtWeightEnd.Text = '') AND not (edtWeightStart.Text = '') then
  begin
    yield := SimpleRoundTo(100 * (strtofloat(edtWeightEnd.Text) / strtofloat(edtWeightStart.text)),-2);
    YieldLabel.Caption := floattostr(yield) + ' %';
  end
  else
  //display yield in percent from weights and round
  begin
    YieldLabel.Caption := '';
  end;
end;

procedure TfrmMAMS.CalculateYieldTouch;
// calulcates the yield from the values used in the touch tab
VAR yield: double;
begin
  LogWindow.addLogEntry('CalculateYieldTouch');
  if not (DBedtTouchWeightsAfterPrep.Text = '') AND not (DBedtTouchWeightsBeforePrep.Text = '') then
  begin
    yield := SimpleRoundTo(100 * (strtofloat(DBedtTouchWeightsAfterPrep.Text) / strtofloat(DBedtTouchWeightsBeforePrep.text)),-2);
    lblTouchYieldValue.Caption := floattostr(yield) + ' %';
  end
  else
  //display yield in percent from weights and round
  begin
    YieldLabel.Caption := 'no value';
  end;
end;

procedure TfrmMAMS.CalculateNetWeightEnd;
// calulcates the net weight of weight_end of the prep table by
// weight_medium_2 - weight_medium
// this is used for weighing empty and full containers in order to calucate the weight of the sample therein
VAR netweight: double;
begin
  if not (DBEditWeightMedium.Text = '') AND not (DBEditWeightMedium2.Text = '') then
  begin
    netweight := strtofloat(DBEditWeightMedium2.Text) - strtofloat(DBEditWeightMedium.text);
    edtWeightEnd.Text := floattostr(netweight);
    //edtWeightEnd.Field.AsString := floattostr(netweight);
  end
end;

procedure TfrmMAMS.CalculateNetWeightEndTouch;
// calulcates the net weight of weight_end of the prep table by
// weight_medium_2 - weight_medium
// this is used for weighing empty and full containers in order to calucate the weight of the sample therein
VAR netweight: double;
begin
  if not (DBedtTouchWeightsMediumPrep.Text = '') AND not (DBedtTouchWeightsMedium2Prep.Text = '') then
  begin
    netweight := strtofloat(DBedtTouchWeightsMedium2Prep.Text) - strtofloat(DBedtTouchWeightsMediumPrep.Text);
    DBedtTouchWeightsAfterPrep.Text := floattostr(netweight);
    //DBedtTouchWeightsAfterPrep.Field.AsString := floattostr(netweight);
  end
end;

procedure TfrmMAMS.ToolBar1Click(Sender: TObject);
begin
 //keep the current button pressed
 ToolBarButtonsState(Sender);
end;

procedure TfrmMAMS.ToolBarButtonsState(Sender: TObject);
var
  i: integer;
  name: string;
begin
  //ToolButton1.Style.tbsCheck:=true;
  //  ToolButton1.Down:=False;
  //  ToolButton2.Down:=False;
  //  ToolButton3.Down:=False;
  //  btnSubmitter.Down:=False;
  //  btnEditSample.Down:=False;
  //  btnTransfer.Down:=False;
  //  btnAdmin.Down:=False;
  //  ToolButton10.Down:=False;
  //  btnDBPlot.Down:=False;
  //  btnShowPrior.Down:=False;
  //  ToolButton5.Down:=False;
  //  btnShowPretreat.Down:=False;
  //  btnNewSample.Down:=False;
  //  ToolButton4.Down:=False;
  //  ToolButton6.Down:=False;
  //  ToolBtnSendMail.Down:=False;
  name := (Sender AS TToolButton).Name;
  LogWindow.addLogEntry('found unit name '+ name);
  for i := 0 to ToolBar1.ButtonCount - 1 do
  begin
    if ToolBar1.Buttons[i].Name = name then
    begin
      ToolBar1.Buttons[i].Down := true;
    end;
  end;
end;


procedure TfrmMAMS.ToolbtnHomeClick(Sender: TObject);
// show home tab and reload all the data in the tab
VAR express: integer;
begin
  // show Home tab
  pgtMain.ActivePage := tbsHome;

  // #### update available Oxas, Blanks, etc
  // load number of available oxas
  dm.GetNOxaBlank(true);
  pnlHomeNumberOfOxas.Caption := '# Oxas: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 9 then pnlHomeNumberOfOxas.Color:=clRed;
  if (dm.qryDB.RecordCount in [9..14]) then pnlHomeNumberOfOxas.Color:=clYellow;
  if dm.qryDB.RecordCount > 14 then pnlHomeNumberOfOxas.Color:=clGreen;
  // load number of available Blanks
  dm.GetNOxaBlank(false);
  pnlHomeNumberOfBlanks.Caption := '# Blanks: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 5 then pnlHomeNumberOfBlanks.Color:=clRed;
  if (dm.qryDB.RecordCount in [5..8]) then pnlHomeNumberOfBlanks.Color:=clYellow;
  if dm.qryDB.RecordCount > 8 then pnlHomeNumberOfBlanks.Color:=clGreen;
    // load number of available Pferde
  dm.GetNPferde();
  pnlHomeNumberOfPferde.Caption := '# Pferde: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 1 then pnlHomeNumberOfPferde.Color:=clRed;
  if (dm.qryDB.RecordCount in [2..3]) then pnlHomeNumberOfPferde.Color:=clYellow;
  if dm.qryDB.RecordCount > 3 then pnlHomeNumberOfPferde.Color:=clGreen;
    // load number of available C6
  dm.GetNIAEA('C6');
  pnlHomeNumberOfC6.Caption := '# C6: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 1 then pnlHomeNumberOfC6.Color:=clRed;
  if (dm.qryDB.RecordCount in [2..3]) then pnlHomeNumberOfC6.Color:=clYellow;
  if dm.qryDB.RecordCount > 3 then pnlHomeNumberOfC6.Color:=clGreen;
    // load number of available C7
  dm.GetNIAEA('C7');
  pnlHomeNumberOfC7.Caption := '# C7: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 1 then pnlHomeNumberOfC7.Color:=clRed;
  if (dm.qryDB.RecordCount in [2..3]) then pnlHomeNumberOfC7.Color:=clYellow;
  if dm.qryDB.RecordCount > 3 then pnlHomeNumberOfC7.Color:=clGreen;
    // load number of available C8
  dm.GetNIAEA('C8');
  pnlHomeNumberOfC8.Caption := '# C8: ' + IntToStr(dm.qryDB.RecordCount);
  if dm.qryDB.RecordCount < 1 then pnlHomeNumberOfC8.Color:=clRed;
  if (dm.qryDB.RecordCount in [2..3]) then pnlHomeNumberOfC8.Color:=clYellow;
  if dm.qryDB.RecordCount > 3 then pnlHomeNumberOfC8.Color:=clGreen;

  // ### update number of unprep'd samples and samples ready for graph and express
  // unpre'd samples
  pnlHomeNumberofUnprepdSamples.Caption := '# to-prep all: ' + IntToStr(dm.GetAllPlanned(False, 'none'));
    // unpre'd bone samples
  pnlHomeNumberofUnprepdSamplesBone.Caption := '# to-prep bones: ' + IntToStr(dm.GetAllPlanned(False, 'bone'));
    // unpre'd teeth samples
  pnlHomeNumberofUnprepdSamplesTeeth.Caption := '# to-prep teeth: ' + IntToStr(dm.GetAllPlanned(False, 'tooth'));
  // samples inPrep
  pnlHomeNumberofSamplesInPrep.Caption := '# in-prep: ' + IntToStr(dm.GetAllInPrep);
  // ready for graph
  pnlHomeNumberofReadyForGraph.Caption := '# to-graph: ' + IntToStr(dm.GetAllWaitingForGraph);
  // ready for meas
  pnlNumberofSamplesReadyForMeas.Caption := '# to-meas: ' + IntToStr(dm.GetAllWaitingForMeas);

  // number of express samples
  express:=dm.GetWaitingExpress;
  pnlHomeNumberofExpress.Caption := '# express: ' + IntToStr(express);
  if express>0 then pnlHomeNumberofExpress.Color:=clRed
  else pnlHomeNumberofExpress.Color:=clGreen;

  DbGridHomeExpressSamples.Columns[0].Width := 70;
  DbGridHomeExpressSamples.Columns[1].Width := 80;
  DbGridHomeExpressSamples.Columns[2].Width := 100;
  DbGridHomeExpressSamples.Columns[3].Width := 80;
  DbGridHomeExpressSamples.Columns[4].Width := 80;



end;

procedure TfrmMAMS.SetupCommonPretreatment;
begin
//  gbxSelectPretreatment.Parent := pnlSetCommonPretreatment;
  grdPreviewSamples.Parent := wizSelectPretreatment;
end;

procedure TfrmMAMS.SetupInsertSamplesWizard;
begin
  wizInputSamples.ActivePage := wizStartPage;
  with edtNewSamplesFilename do
  begin
    if Length(Filename) > 0 then
    begin // strip previously used filename
      InitialDir := ExtractFilePath(Filename);
      Filename := 'Select *.csv';
    end;
  end;
  grdPreviewUser.Visible := false;
  grdPreviewSamples.Visible := false;
  wizStartPage.Subtitle.Text := 'Excel list supplied by user must have been converted and stored as *.csv file' +
    #10 + 'Check input in grids and click next';
end;

procedure TfrmMAMS.SetupLabPlanPage;
var
  i, w: integer;
begin
   if dm.dsSamplesAvailable.DataSet.RecordCount > 0 then
     begin
      with grdSamplesAvailable do begin
        Visible := true;
        Columns[0].Width := 55;   //sample_nr
        Columns[1].Width := 170;  //project
        Columns[2].Width := 170;  //user_label
        Columns[3].Width := 150;  //user_label_nr
        Columns[4].Width := 40;   //weight
        Columns[5].Width := 50;   //prep_nr
        Columns[6].Width := 60;  //target_nr
        w := 0;
        for i := 0 to Columns.Count - 1 do w := w + Columns[i].Width;
      end;
      gbxSamplesAvailable.Width := w + 50; // scroll
      lbxBatch.Clear;
      JvDBStatusLabel3.Visible := true;
     end;
end;

procedure TfrmMAMS.SetupLabTaskPage;
begin
  dtStartTaskDate.Date := Date;
  dtStartTaskTime.Time := TimeOf(Now);
  dtEndTaskDate.Date := Date;
  dtEndTaskTime.Time := TimeOf(Now);
  if rgpLabTask.ItemIndex = 0 then begin
    pgtLabTasks.ActivePage := tbsPrepTask;
    btnLoadBatchClick(self);
  end;
  lbPrepTimesHelp.Text := '<b>How to enter date/time of lab tasks :</b>' +
    '<br> <br>' +
    '1. Load desired batch <br>' +
    '2. Select desired method <br>' +
    '3. Click NOW button to set start time for all samples in batch<br>' +
    '4. Do same procedure for end of task<br><br>';
  btnStartTask.Enabled := false;
  btnEndTask.Enabled := false;
  btnSetSampleReadyForGraph.Caption := 'Set all samples ' + #13 + 'ready for' + #13 + 'graphitization';
  grdActiveBatches.Visible := false;
  grdSamplesOfLabTask.Visible := false;
end;

procedure TfrmMAMS.SetupMaterial;
begin
  grdPreviewSamples.Parent := wizSelectMaterial;
end;

procedure TfrmMAMS.SetupNewSamplesGrid;
begin
  with grdPreviewSamples do begin
    ColCount := 16;
    Selection := NoSelection;
    ColWidths[0] := 30;
    cells[1, 0] := 'sample name';
    ColWidths[1] := 100;
    cells[2, 0] := 'sample nr.';
    ColWidths[2] := 40;
    cells[3, 0] := 'sample descr.1';
    ColWidths[3] := 70;
    cells[4, 0] := 'sample descr.2';
    ColWidths[4] := 70;
    cells[5, 0] := 'weight';
    ColWidths[5] := 50;
    cells[6, 0] := 'material (user)';
    ColWidths[6] := 70;
    cells[7, 0] := 'comment';
    ColWidths[7] := 100;
    cells[8, 0] := 'material (our)';
    ColWidths[8] := 70;
    cells[9, 0] := 'fraction';
    ColWidths[9] := 70;
    cells[10, 0] := 'type';
    ColWidths[9] := 60;
    cells[11, 0] := 'pretreat';
    ColWidths[10] := 40;
    cells[SamplePrep1Col, 0] := 'PrepStep 1';
    ColWidths[SamplePrep1Col] := 0;
    cells[SamplePrep2Col, 0] := 'PrepStep 2';
    ColWidths[SamplePrep2Col] := 0;
    cells[SamplePrep3Col, 0] := 'PrepStep 3';
    ColWidths[SamplePrep3Col] := 0;
    cells[SamplePrep4Col, 0] := 'PrepStep 4';
    ColWidths[SamplePrep4Col] := 0;
    cells[SamplePrep5Col, 0] := 'PrepStep 5';
    ColWidths[SamplePrep5Col] := 0;
  end;
end;

procedure TfrmMAMS.DisplayProgramOptions;
// setup the AppOptions Elements
// most of the values are created at design time
Var page: integer;
begin

  // don't show tab names
  for page := 0 to TabOptionsPages.PageCount - 1 do
   begin
     TabOptionsPages.Pages[page].TabVisible := false;
   end;
  //dispaly General Tab
  TabOptionsPages.ActivePage := TabGeneral;
  // set treeView to general
  OptionsTree.Items[0].Selected := True;
  OptionsTree.Refresh;
  //values are automatically loaded using JvFormStorage1


end;


procedure TfrmMAMS.ReportsSetupOptions;
//loads the setup option of the reports
// and displays them in the optiosn tab
var
  i: integer;
begin
  with grdReportHeadings do
  begin
    for i := 1 to RowCount - 1 do cells[0, i] := IntToStr(i);
    ColWidths[0] := 20;
    ColWidths[1] := 80;
    ColWidths[2] := 120;
    ColWidths[3] := 120;
    cells[1, 0] := 'Field database';
    cells[2, 0] := 'Column in Report';
    //cells[3, 0] := 'Column German';
  end;
  with dm.qryDB, grdReportHeadings do
  begin
    SQL.Text := 'SELECT sample_t.sample_nr, user_label,user_label_nr,user_desc1,' +
      'user_desc2, target_t.C14_age, target_t.C14_age_sig, av_dc13 as dc13, conc_c, conc_c/conc_n as cn_ratio, weight_end/weight_start as collpc, av_fm, av_fm_sig, material ' +
      'FROM sample_t ' +
      'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
      'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
      'WHERE sample_t.sample_nr=1' + ';';
    LogWindow.addLogEntry(SQL.Text);
    JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
    IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    for i := 1 to RowCount - 1 do // label row number
//    begin
//      cells[1, i] := Fields.Fields[i - 1].FieldName;
//    end;
  end;
  lbExportHeader.Text := '<b>How to create table headings for reports using MS Word, coresponding to field names in database :</b>' +
    '<br> <br>' +
    ' 1. Create table and headings in MS Word (two rows, identical headings) <br> <br>' +
    ' 2. Copy/paste headings from MS Word into cells of grid above (double click in cells), according to field in database and language <br> <br>' +
    ' 3. Grid contents are saved automatically on exit <br> <br>' +
    ' The table in the MS word document should have two more columns, labeled Cal 1 sigma and Cal 2 sigma  <br> <br>' +
    ' These columns are filled with calibrated ranges during data export to MS Word <br> <br>';
end;

procedure TfrmMAMS.wizFinalPageEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
begin
  lbWizFinalPage.Text := 'waiting for your action ...';
  wizFinalPage.EnableButton(bkFinish, true);
end;



procedure TfrmMAMS.wizFinalPageFinishButtonClick(Sender: TObject;  var Stop: Boolean);
// sample import is done, insert sample into database
begin
  wizFinalPage.EnableButton(bkFinish, false);
  lbWizFinalPage.Text := 'Inserting data into database, please wait';
  //insert the samples into the database
  InsertNewSamplesInDb;
  //clear the list of samples
  ClearInsertSampleGrids;
  //pgtMain.ActivePage := tbsInsertSamples;
  chkFreeOfCharge.Checked := false;
  chkInsertReturnToSender.Checked := false;
  checkInsertAllCNIsotopA.Checked := false;

end;

procedure TfrmMAMS.wizInputProjectEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
begin
  wizInputProject.EnableButton(bkNext, false);
  edtProjectName.Text := ProjectName;
  edtInDate.Date := Date;
  edtDesiredDate.Date := Date + strtointdef(edtOptionsTurnover.Text, 90); // this value is taken from the options tab, if conversion fails use 90 days as default
  if Length(edtProjectName.Text) > 0 then wizInputProject.EnableButton(bkNext, true);
  CheckProjectExists;
  if ProjectExists then ShowMessage('Project already exists! Samples will be added to the project! ');
end;

procedure TfrmMAMS.wizInputSamplesBackButtonClick(Sender: TObject);
begin
//  if (MessageDlg('Are you sure to leave the sample input wizard ?', mtConfirmation, [mbNo, mbOK], 0) in [mrOk]) then
//    frmInputNewSamples.Close;
end;

procedure TfrmMAMS.wizInputSamplesCancelButtonClick(Sender: TObject);
begin
  if (MessageDlg('Are you sure to leave the sample input wizard ?', mtConfirmation, [mbNo, mbOK], 0) in [mrOk]) then begin
    ClearInsertSampleGrids;
    wizInputSamples.SelectFirstPage;
  end;
end;

procedure TfrmMAMS.wizInputSamplesFinishButtonClick(Sender: TObject);
begin
// now insert everything (Sample, user, samples)
  gbxSelectPretreatment.Parent := gbxSampleLeft;
end;

procedure TfrmMAMS.wizInputInvoiceAddressEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
var
  s: string;

  procedure PrepareUserGrid;
  begin
    with dm.qryDB do
    begin
      SQL.Text := 'Select first_name, last_name, organisation, institute, address_1, user_nr FROM user_t WHERE ' +
        ' last_name=' + #34 + grdPreviewUser.Cells[1, 2] + #34;
      s := SQL.Text;
      //   ClibBoard.SetTextBuf(PChar(s));
      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      IF dm.adoConnKTL.Connected THEN
      Begin
        Try
          Open;
          LogWindow.addLogEntry('executed');
        Except
          ShowMessage('problem opening the database');
          JvLogFile.Add('DB',lesError,'Problem opening DB');
        End;
      End;
    end;
    with grdShowUsers do
    begin
      Columns[0].Width := 100; // name
      Columns[1].Width := 150; // surname
      Columns[2].Width := 200; // org
      Columns[3].Width := 200; // inst
      Columns[4].Width := 200; // addr
      Columns[5].Width := 60; // nr
    end;
    grdShowUsers.Visible := true;
  end;

begin
  with wizInputInvoiceAddress do
  begin
    Title.Text := '';
    EnableButton(bkNext, false);
  end;
  with wizInputInvoiceAddress do
  begin
    EnableButton(bkNext, false);
    Title.Text := 'Select invoice address';
  end;
end;

procedure TfrmMAMS.wizSelectMaterialEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
var
  i, w: integer;
  MaterialFillDone: boolean;
begin
  grdPreviewSamples.Parent := wizSelectMaterial;
  with wizSelectMaterial do
  begin
    SubTitle.Text := 'Select material from the list on the right and drag to the "our material" column ' + #13 +
      'press Ctrl during drag or double-click to assign to all samples; ' +
      'repeat for fraction';
    EnableButton(bkNext, false);
  end;
  with grdPreviewSamples do
  begin
    MaterialFillDone := true;
    for i := 0 to RowCount - 1 do
      if length(cells[MaterialCol, i]) = 0 then MaterialFillDone := false;
  end;
  if MaterialFillDone then wizSelectMaterial.EnableButton(bkNext, true);
  FillMaterialBox;
  FillFractionBox;
  SetSampleGridForMaterial;
  w := 0;
  with grdPreviewSamples do
  begin
    for I := 0 to ColCount - 1 do w := w + ColWidths[i] + 1; // 1 = gridline
    Width := w + 6;
  end;
end;

procedure TfrmMAMS.wizSelectMaterialNextButtonClick(Sender: TObject;
  var Stop: Boolean);
begin
  SetupCommonPretreatment;
end;

procedure TfrmMAMS.wizSelectPretreatmentBackButtonClick(Sender: TObject;
  var Stop: Boolean);
begin
  SetupMaterial;
end;

procedure TfrmMAMS.wizSelectPretreatmentEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
begin
  grdPreviewSamples.Parent := wizSelectPretreatment;
  SetSampleGridForPrep;
  FillPrepBox;
  lbxDefinePrepSteps.Align := alClient;
  wizSelectPretreatment.Subtitle.Text := 'Select prep step from the list and drag to one of the the prep columns;' + #13 +
    'press Ctrl during drag to assign this step to all samples. Double click on a prep-step to assign to all samples for prep_step_1.';
end;

procedure TfrmMAMS.wizSelectTypeEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
begin
  grdPreviewSamples.Parent := wizSelectType;
  SetSampleGridForType;
  FillTypeBox;
  pnlSetType.Align := alClient;
  with wizSelectType do begin
    EnableButton(bkNext, false);
    Subtitle.Text := ' Double-click on the type in the list to set it for all samples ' +
      ' or drag the typ to the each sample individually';
  end;
  with grdPreviewSamples do begin
    if length(Cells[TypeCol, 1]) > 0 then wizSelectType.EnableButton(bkNext, true);
  end;
end;

procedure TfrmMAMS.wizStartPageBackButtonClick(Sender: TObject;
  var Stop: Boolean);
begin
  wizStartPage.EnableButton(bkNext, true);
end;

procedure TfrmMAMS.wizStartPageEnterPage(Sender: TObject;
  const FromPage: TJvWizardCustomPage);
begin
  if length(grdPreviewUser.Cells[1, 1]) = 0 then wizStartPage.EnableButton(bkNext, false); //last name empty
  grdShowUsers.Visible := false;
end;

procedure TfrmMAMS.FormShow(Sender: TObject);
//initial setup when form of main program shows
var
  Year, Month, Day: word;
  i, j: integer;
  s: string;
  Schema, username, password: string;
begin

// some initial setup
  //  frmLogSql.Show;
  SetFormat;
  KTLsupervisor := false;
  DecodeDate(Date, Year, Month, Day);
  edtStartYear.Value := Year - 2;
  edtEndYear.Value := Year;
  edtStartAna.Value := 1;
  edtEndAna.Value := 1000000;

  //LogWindow:=LogWindow.Create(self);
  JvLogFile.Add('Startup',lesInformation,'starting Application');
  frmStart.MemoStartScreenMessages.Lines.Clear;
  StatusBar.Panels[2].Text:='creating Database objects...';      //display a status in the status bar

  // define MySQL connection to main database from the parameters
// that need to be given when starting SAMS
// at this time those are parameters to the ODBC connection as set up in windows
// params are : Data Source=...(ODBC in control panel) User ID = ... Password=....
// e.g. SAMS_v1.exe db_dmams_KTL rfriedrich pswd
  s := 'Provider=MSDASQL.1';  // beginning of the connection string
  dm.adoConnKTL.LoginPrompt := true;

  // try building the connection string from the paramters give during startup
  JvLogFile.Add('Startup',lesInformation,'DB - Param count = ' + ParamCount.ToString);
  frmStart.MemoStartScreenMessages.Lines.Add('DB - Param count = ' + ParamCount.ToString);
  with dm.adoConnKTL do    //this specifies the ADOConnection to the database
  begin
    if Connected then
    begin
    dm.adoConnKTL.Close;
    //JvLogFile.Add('Startup',lesInformation,'test');
      if ParamCount > 0 then   //some parameters are given
      begin
        JvLogFile.Add('Startup',lesInformation,'DB - ODBC-DataSource= ' + ParamStr(1));
        frmStart.MemoStartScreenMessages.Lines.Add('DB - found connection parameters (' + ParamCount.ToString+ ')');
     //Provider=MSDASQL.1;Password=Micadas;Persist Security Info=True;User ID=root;Data Source=DMYSQL_KTL
        s := s + ';Data Source =' + ParamStr(1);     // one parameter given, this must be the name of the ODBC connection
        frmStart.MemoStartScreenMessages.Lines.Add('DB - Data Source= ' + ParamStr(1));
        if ParamCount > 1 then
          begin
            s := s + ';User ID=' + ParamStr(2);    // two parameters given, must be the ODBC connection and user name
            frmStart.MemoStartScreenMessages.Lines.Add('DB - User ID= ' + ParamStr(2));
            JvLogFile.Add('Startup',lesInformation,'DB - UserID= ' + ParamStr(2));
          end;
        if ParamCount > 2 then                       // three parameters give, must be ODBC connection, user name and password
        begin
          s := s + '; Password :=' + ParamStr(3);
          frmStart.MemoStartScreenMessages.Lines.Add('DB - found connection password');
          JvLogFile.Add('Startup',lesInformation,'DB - Connection Password is given');
          dm.adoConnKTL.LoginPrompt := false;    //no logon prompt is needed since password and user name is supplied
        end;
        if ParamCount > 3 then      // if four parameters are give and the last one is "KTLsupervisor" some special settings are applied
        begin
          // special setting for super users
          if ParamStr(4) = 'KTLsupervisor' then
          begin
            btnTransfer.Visible := true;
            KTLsupervisor := true;
            JvLogFile.Add('Startup',lesInformation,'starting Application as KTLSupervisor');
          end;
        end;
        //JvLogFile.Add('Startup',lesInformation,'DB - Connection String 1 = ' + s);
        dm.adoConnKTL.ConnectionString := s; // send the connection string to the ADO connection
        dm.adoConnKTL.Open;
        //JvLogFile.Add('Startup',lesInformation,'DB - Connection String 2 = ' + s);
        frmMAMS.Caption := ' SAMS v.' + myVersion + ' -- ODBC: ' + ParamStr(1) + ' -- User: ' + ParamStr(2); //display the connection info in the header of the form
      end
      else
      begin
      // no ODBC parameters are give in order to connect to the database
      StatusBar.Panels[2].Text:='DB - not enough parameters to connect to DB using ODBC';
      JvLogFile.Add('Startup',lesError,'Not enough parameters to connect to DB using ODBC');
      //    adoConnKTL.
      end;
    end;

    // try to connect to the database
    Try
      frmStart.MemoStartScreenMessages.Lines.Add('DB - try connecting to database');
      JvLogFile.Add('Startup',lesInformation,'connecting to DB at startup');
      Open;   // open ADOConnection
      StatusBar.Panels[2].Text:='Open Database connection...';
      frmStart.MemoStartScreenMessages.Lines.Add('DB - connected');
    Except
        frmStart.MemoStartScreenMessages.Lines.Add('DB - ERROR: not able to connect!');
        JvLogFile.Add('Startup',lesError,'Not able to connect to DB');
        showmessage('Unable to connect to the database. Make sure an ODBC driver is selected and you are connected to the network.');
    End;
      end;
  s := dm.adoConnKTL.ConnectionString;
  //   ClibBoard.SetTextBuf(PChar(s));

  // load some of the tables from the Database if the ADO connection worked
  if dm.adoConnKTL.Connected then
  Begin
    with dm do
    begin
    JvLogFile.Add('Startup',lesInformation,'Connected to DB.');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading user data...');
      tblUser.Open;
        StatusBar.Panels[2].Text:='user data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - user data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading invoice data...');
      tblInvoice.Open;
        StatusBar.Panels[2].Text:='invoice data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - invoice data loaded');
  //    qryAllUser.Open;
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading project data...');
      tblProjectTypes.Open;
        StatusBar.Panels[2].Text:='project data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - project data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading project status data...');
      tblProjectStatus.Open;
        StatusBar.Panels[2].Text:='project status data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - project status data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading method data...');
      tblMethod.Open;
        StatusBar.Panels[2].Text:='method data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - method data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading type data...');
      tblResearchType.Open;
        StatusBar.Panels[2].Text:='research type data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - research type data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading report type data...');
      tblReportType.Open;
        StatusBar.Panels[2].Text:='report type data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - report type data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading material data...');
      tblMaterial.Open;
        StatusBar.Panels[2].Text:='material data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - material data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading type data...');
      tblTypes.Open;
        StatusBar.Panels[2].Text:='type data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - type data loaded');
      frmStart.MemoStartScreenMessages.Lines.Add('DB - loading fraction data...');
      tblFraction.Open;
        StatusBar.Panels[2].Text:='fraction data loaded';
        frmStart.MemoStartScreenMessages.Lines.Add('DB - fraction data loaded');
    end;
  End;

  // set some initial values
  cmbReportType.KeyValue := 0;
  cmbProjectType.KeyValue := 0;
  cmbResearchType.KeyValue := 0;

  SetupProjectPage;

  //  pgtMain.ActivePage := tbsSamplesOfUser;
  for i := 0 to pgtMain.PageCount - 1 do pgtMain.Pages[i].TabVisible := False;
  for i := 0 to pgtLabTasks.PageCount - 1 do pgtLabTasks.Pages[i].TabVisible := False;

  SetupInsertSamplesWizard;

  frmStart.MemoStartScreenMessages.Lines.Add('checking report headings');
  if FileExists(ExtractFilePath(Application.ExeName) + '\ReportHeadings.txt') then begin
    grdReportHeadings.LoadFromFile(ExtractFilePath(Application.ExeName) + '\ReportHeadings.txt');
  end;
  frmStart.MemoStartScreenMessages.Lines.Add('Ready');
  StatusBar.Panels[2].Text:='Ready';

  //close the StartScreen window Unit: frmStartScreen.pas
  frmStart.Hide; frmStart.Free;

  // set some default values for the date pickers
  JvDBDateTimePicker1.DropDownDate := Today;
  JvDBDateTimePicker2.DropDownDate := Today;
  JvDBDateTimePicker3.DropDownDate := Today;
  edtSampleInfoInDate.DropDownDate := Today;
  edtSampleInfoDesiredDate.DropDownDate := Today;
  edtSampleInfoOutDate.DropDownDate := Today;
  edtSamplingDate.DropDownDate := Today;
  edtPrepStart.DropDownDate := Today;
  edtPrepEnd.DropDownDate := Today;
  edtGraphDate.DropDownDate := Today;
  edtTargetPressed.DropDownDate := Today;
  edtGraphitized.DropDownDate := Today;
  DBDateTimeTouchPrepStart.DropDownDate := Today;
  DBDateTimeTouchPrepEnd.DropDownDate := Today;

  edtMonthStat.Value := YearOf(Date());

  // switch pgtMain to SampleInfo and show sample info of the currently selected Sample_nr
  //actSampleInfoExecute(self);

  // switch to home tab  and update data
  pgtMain.ActivePage := tbsHome;
  ToolbtnHomeClick(Self);

end;

function TfrmMAMS.GetFileName: string;
begin

end;

procedure TfrmMAMS.GetListOfRecentMagazines(MonthsBack: integer);
begin
  if dm.adoConnKTL.Connected then
  Begin
    with dm.qryMagazines do
    begin
      Close;
      // this is pre-2023 when only one AMS was in operation and all Magazines came from one source
      //SQL.Text := 'SELECT DISTINCT magazine from target_t  ' +
      //  ' WHERE magazine not like "set%" ' +
      //  ' order by magazine desc ' +
      //  ' LIMIT ' + FloatToStr(edtMagazineLimit.Value) +';';

        // this is new since 2023 when operating two AMS, Magazines can come
        // from two machines and just sorting by name wont work anymore since
        // the magazines have different names depending on the AMS
      SQL.Text := 'SELECT name as magazine from magazine_t  ' +
        ' WHERE name LIKE '+ #34 + '%' + edtTransferMagazinFilter.Text + '%' + #34 +
        ' order by last_changed desc ' +
        ' LIMIT ' + FloatToStr(edtMagazineLimit.Value) +';';

      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      Open;
    end;
  End;
end;

procedure TfrmMAMS.GetListOfUnpressedMagazines(MonthsBack: integer);
begin
  if dm.adoConnKTL.Connected then
  Begin
    with dm.qryMagazinesUnpressed do
    begin
      Close;
      // only show magazines with targets that have not been measured yet
      SQL.Text := 'SELECT DISTINCT magazine from target_t  ' +
        ' inner join sample_t on target_t.sample_nr = sample_t.sample_nr ' +
        ' WHERE magazine not like "set%" ' +
        ' and fm IS NULL ' +
        ' and type not like "blank" ' +
        ' and type not like "oxa%" ' +
        ' and target_t.stop ="0" ' +
        ' order by magazine desc ';

      LogWindow.addLogEntry(SQL.Text);
      JvLogFile.Add('SQLQuery',lesInformation,SQL.Text);
      Open;
    end;
  End;
end;

procedure TfrmMAMS.GetProjects;
begin
  SetupProjectPage;
  dm.qryProjectsOfUser.Close;
  dm.qryProjectsOfUser.SQL.Text := 'SELECT project, in_date, desired_date, project_nr, AuftragsNr ' +
    'FROM project_t WHERE user_nr=' +
    IntToStr(cmbSubmitterNameProject.KeyValue) + ' ORDER BY in_date DESC;';
  dm.qryProjectsOfUser.Open;
//  dm.qryProjectsOfUser.
  grdProjects.Visible := true;
  grdProjects.Columns[0].Width := 200;
  //grdProjects.Columns[1].Width := 80;
end;

procedure TfrmMAMS.GetSampleofUserProject;
var
  i, sample_nr: integer;
begin
  sample_nr := dm.dsSamplesOfProject.DataSet.FieldByName('sample_nr').AsInteger;
  // Query Sample Info by SampleNr
  dm.QuerySampleBySampleNr(sample_nr);
  SetupPretreatmentStepsGrid;
  //Query the Prep Information by SampleNr
  dm.QueryPrepStepsBySampleNr(sample_nr, 1); // use prep_nr = 1
  with grdShowPrepSteps do
  begin
    RowHeights[8] := 3;
  end;
  // query Target info by SampleNr
  dm.QueryTargetInfoBySampleNr(sample_nr);
  // set tab with sample Info to active
  pgtSample.ActivePage := tbsSampleOfProject;
  gbxSampleRight.Visible := true;
  gbxSampleLeft.Visible := true;
  // set sample_nr in the SampleDetail page to the current selected sample
  edtSampleNr.Text := IntToStr(sample_nr);
end;

procedure TfrmMAMS.grdSamplesAvailableCellClick(Column: TColumn);
// set the selected sample number in the SAmpleINof Page
// so that the suer can switch to the SampleINfo page
// and sees all the information
begin
  edtSampleNr.Value := dm.dsSamplesAvailable.DataSet.FieldByName('sample_nr').AsInteger;
end;

procedure TfrmMAMS.grdSamplesAvailableDblClick(Sender: TObject);
//Drag and Drop doesn't seem to work, use double click instead
begin
  grdSamplesAvailable.MouseCoord(DragCol,DragRow);
  //grdSamplesAvailable.mousetocell(x, y, DragCol, DragRow);
  lbxBatchDragDrop(grdSamplesAvailable,grdSamplesAvailable,DragCol,DragRow);
end;

procedure TfrmMAMS.grdSamplesAvailableDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if (POS('Eil',dm.qrySamplesAvailable.FieldByName('user_label').AsString)>0) OR (POS('EIL',dm.qrySamplesAvailable.FieldByName('user_label').AsString)>0) OR (POS('Blitz',dm.qrySamplesAvailable.FieldByName('user_label').AsString)>0)      //if a certain substring is in user_label then change row color
    then TDBGrid(Sender).Canvas.Brush.Color:=clRed;
      TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdSamplesAvailableMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  // ShowMessage('test');
  grdSamplesAvailable.mousetocell(x, y, DragCol, DragRow);
  grdSamplesAvailable.BeginDrag(false, 5);
end;

procedure TfrmMAMS.grdSamplesOfLabTaskCellClick(Column: TColumn);
begin
  edtSampleNr.Text := IntToStr(dm.qrySamplesOfLabTask.FieldByName('sample_nr').AsInteger);
end;

procedure TfrmMAMS.grdSamplesOfProjectCellClick(Column: TColumn);
begin
  GetSampleOfUserProject;
end;

procedure TfrmMAMS.grdSamplesOfProjectDblClick(Sender: TObject);
// double clicking a sample shows the sample info page
begin
   pgtMain.ActivePage:=tbsSampleInfo; //display sample info page and fill in information
   GetSampleOfUserProject;
   FillPrepList;
   DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text));
   edtSampleNr.SetFocus;
end;

procedure TfrmMAMS.grdSamplesOfProjectDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  AlternateRowColors(Sender, State);
  (Sender as TDBGrid).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdSamplesOfProjectKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  GetSampleofUserProject;
end;

procedure TfrmMAMS.grdSamplesOfProjectMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
begin
  GetSampleofUserProject;
end;

procedure TfrmMAMS.grdSamplesOfSubmitterDblClick(Sender: TObject);
// Display SampleInfo and switch to SampleInfo View
begin
  ShowSampleInfoPage(Sender);
end;

procedure TfrmMAMS.grdSamplesOfSubmitterDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
CONST
  CtrlState: array[0..1] of integer = (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED) ; //DFCS_BUTTONRADIO
begin
  AlternateRowColors(Sender, State);

  // display boolean fields as checkboxes
  if (Column.FieldName='Prep_Discarded') OR
  (Column.FieldName='Graph_Discarded') then
  begin
    TDBGrid(Sender).Canvas.FillRect(Rect) ;
    if VarIsNull(Column.Field.Value) then
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, DFCS_BUTTONCHECK or DFCS_INACTIVE) {grayed}
    else
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, CtrlState[Column.Field.AsInteger]) ; {checked or unchecked}
  end;

  // (Sender as TDBGrid).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.ShowSampleInfoPage(Grid: TObject);
// Grid must be a TDBGrid with at least sample_nr in it
// sample_nr is written into the SampleInfo page and all other
// information are populated in order to display the full sample info
// is being used for DoubleClick Events
begin
  edtSampleNr.Value:=(Grid as TDBGrid).DataSource.DataSet.FieldByName('sample_nr').AsInteger; // enter sample_nr into editfield of sampleInfo tab
  pgtMain.ActivePage:=tbsSampleInfo; //display sample info page and fill in information
  FillPrepList; // fill list with possible prep steps
  DoSampleInfo(round(edtSampleNr.Value), StrToInt(edtSamplePrepNr.Text), StrToInt(edtSampleTargetNr.Text)); //get sample info
  edtSampleNr.SetFocus;
end;

procedure TfrmMAMS.grdSamplesOfSubmitterMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  pt: TGridcoord;
begin
  pt := grdSamplesOfSubmitter.MouseCoord(x, y);
  if pt.y = 0 then
    grdSamplesOfSubmitter.Cursor := crHandPoint
  else
    grdSamplesOfSubmitter.Cursor := crDefault;
end;

procedure TfrmMAMS.grdSamplesOfSubmitterTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdSamplesOfSubmitter.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdSamplesOfSubmitter.DataSource.DataSet) do
    begin
      try
        grdSamplesOfSubmitter.Columns[PreviousColumnIndex].title.Font.Style :=
          grdSamplesOfSubmitter.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdShowUsersCellClick(Column: TColumn);
begin
  glbUserNr := dm.tblUser.FieldByName('user_nr').AsInteger;
  btnInsertExistingUser.Enabled := true;
  UserExists := true;
end;

procedure TfrmMAMS.grdTypesTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdTypes.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdTypes.DataSource.DataSet) do
    begin
      try
        grdTypes.Columns[PreviousColumnIndex].title.Font.Style :=
          grdTypes.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

procedure TfrmMAMS.grdWaitingForGraphDblClick(Sender: TObject);
begin
 ShowSampleInfoPage(Sender);
end;

procedure TfrmMAMS.grdWaitingForGraphDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

  AlternateRowColors(Sender, State);

  if (POS('Eil',dm.qryWaitingForGraph.FieldByName('user_label').AsString)>0)            //if a certain substring is in user_label then change row color
    OR (POS('EIL',dm.qryWaitingForGraph.FieldByName('user_label').AsString)>0)
    OR (POS('Blitz',dm.qryWaitingForGraph.FieldByName('user_label').AsString)>0) then
    begin
      TDBGrid(Sender).Canvas.Brush.Color:=clRed;
    End;

  if dm.qryWaitingForGraph.FieldByName('desired_date').AsDateTime<=Date()+21 Then     //if desired_date approcahes 3 weeks
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clFuchsia;
     End;
  if dm.qryWaitingForGraph.FieldByName('desired_date').AsDateTime<=Date() Then     //if desired_date has run out color text red
     Begin
     TDBGrid(Sender).Canvas.Font.Color := clRed;
     End;

  TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmMAMS.grdWaitingForGraphTitleClick(Column: TColumn);
{$J+}
const PreviousColumnIndex: integer = -1;
{$J-}
begin
  if grdTypes.DataSource.DataSet is TCustomADODataSet then
    with TCustomADODataSet(grdWaitingForGraph.DataSource.DataSet) do
    begin
      try
        grdTypes.Columns[PreviousColumnIndex].title.Font.Style :=
          grdTypes.Columns[PreviousColumnIndex].title.Font.Style - [fsBold];
      except
      end;
      Column.title.Font.Style := Column.title.Font.Style + [fsBold];
      PreviousColumnIndex := Column.Index;
      if (Pos(WideString(Column.Field.FieldName), Sort) = 1)
        and (Pos(WideString(' DESC'), Sort) = 0) then
        Sort := Column.Field.FieldName + ' DESC'
      else
        Sort := Column.Field.FieldName + ' ASC';
    end;
end;

end.

