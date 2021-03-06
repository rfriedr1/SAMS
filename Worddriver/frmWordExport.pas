unit frmWordExport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ActnList, JvToolEdit, Grids, Jvgrids, DBGrids, JvDBControls,
  DBCtrls, Db, DBClient, ComCtrls, Mask, ExtCtrls, worddriver, Buttons, CheckLst,
  JvExGrids, JvExMask;

type
  TWordExportForm = class(TForm)
    Ptop: TPanel;
    FEDoc: TJvFilenameEdit;
    LFEDoc: TLabel;
    PCExport: TPageControl;
    Bstart: TButton;
    ALMain: TActionList;
    TSData: TTabSheet;
    TSoptions: TTabSheet;
    DSExport: TDatasource;
    Panel1: TPanel;
    DBNavigator1: TDBNavigator;
    GExportData: TDBgrid;
    GValues: TJvDrawGrid;
    POptions: TPanel;
    AStart: TAction;
    AFetchParams: TAction;
    ACancel: TAction;
    AAdd: TAction;
    ADelete: TAction;
    CDSExport: TClientDataSet;
    PBottom: TPanel;
    BClose: TButton;
    AClose: TAction;
    PButtons: TPanel;
    SBadd: TSpeedButton;
    SBDelete: TSpeedButton;
    ELExtraData: TLabel;
    BFetch: TButton;
    GroupBox1: TGroupBox;
    CBPrint: TCheckBox;
    CBSave: TCheckBox;
    CBKeepOpen: TCheckBox;
    PTables: TPanel;
    PCOptions: TPageControl;
    TSSerial: TTabSheet;
    LDESave: TLabel;
    LCBTemplate: TLabel;
    CBCurrentRecordOnly: TCheckBox;
    DESave: TJvDirectoryEdit;
    CBTemplate: TComboBox;
    TSNewTable: TTabSheet;
    TSFillTable: TTabSheet;
    RBFillTable: TRadioButton;
    RBNewTable: TRadioButton;
    RBSerial: TRadioButton;
    CBTableTag: TComboBox;
    Label1: TLabel;
    LBFields: TCheckListBox;
    LLBFields: TLabel;
    SBUp: TSpeedButton;
    SBDown: TSpeedButton;
    FEFillTableFileName: TJvFilenameEdit;
    Bestandsnaam: TLabel;
    Label2: TLabel;
    FENewTableFileName: TJvFilenameEdit;
    AUp: TAction;
    ADown: TAction;
    FEMergeDoc: TJvFilenameEdit;
    CBMergeDoc: TCheckBox;
    EditLabel1: TLabel;
    CBkeepDocumentsOpen: TCheckBox;
    CBMergeOnlyOpen: TCheckBox;
    CBFillFields: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure AStartExecute(Sender: TObject);
    procedure AStartUpdate(Sender: TObject);
    procedure AFetchParamsUpdate(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure ACancelUpdate(Sender: TObject);
    procedure AFetchParamsExecute(Sender: TObject);
    procedure GValuesDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure AAddExecute(Sender: TObject);
    procedure ADeleteUpdate(Sender: TObject);
    procedure ADeleteExecute(Sender: TObject);
    procedure GValuesGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure GValuesSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure FormShow(Sender: TObject);
    procedure ACloseExecute(Sender: TObject);
    procedure ACloseUpdate(Sender: TObject);
    procedure CBSaveClick(Sender: TObject);
    procedure RBKindClick(Sender: TObject);
    procedure AUpUpdate(Sender: TObject);
    procedure AUpExecute(Sender: TObject);
    procedure ADownExecute(Sender: TObject);
    procedure ADownUpdate(Sender: TObject);
    procedure CBTableTagChange(Sender: TObject);
    procedure CBMergeDocClick(Sender: TObject);
  private
    { Private declarations }
    FLastFileCheck : String;
    FUseWordVersion : TWordVersion;
    FDataSet : TDataset;
    FValues : TStrings;
    FWord : TWordDriver;
    FDocExists : Boolean;
    function GetFileName: String;
    procedure SetFileName(const Value: String);
    procedure SetValues(const Value: TStrings);
    procedure ExportData;
    Function DocFileExists: Boolean;
    procedure InitGrid;
    procedure CheckFileEnabled;
  public
    { Public declarations }
    Property FileName : String Read GetFileName Write SetFileName;
    Property Dataset : TDataset Read FDataset Write FDataset;
    PRoperty Values : TStrings Read FValues Write SetValues;
  end;

var
  WordExportForm: TWordExportForm;

implementation

uses provider;

{$R *.DFM}

resourcestring
  SField = 'FieldName';
  SValue = 'Value';
  SNewValue = 'New field=';
  SSucces1 = '1 document was created.';
  SSuccesN = '%d documents were created.';

procedure TWordExportForm.FormCreate(Sender: TObject);
begin
  FWord:=TWordDriver.Create(Self);
  FUseWordVersion:=wvDetect;
  FValues:=TStringList.Create;
end;

function TWordExportForm.GetFileName: String;
begin
  Result:=FEDoc.FileName
end;

procedure TWordExportForm.SetFileName(const Value: String);
begin
  FEDoc.FileName:=Value;
end;

procedure TWordExportForm.SetValues(const Value: TStrings);
begin
  FValues.Assign(Value);
end;

procedure TWordExportForm.FormDestroy(Sender: TObject);

begin
  FWord.Free;
  FValues.Free;
end;

procedure TWordExportForm.AStartExecute(Sender: TObject);

begin
  BStart.Action:=ACancel;
  Try
    ExportData;
  Finally
    BStart.Action:=AStart;
  end;
end;

procedure TWordExportForm.ExportData;

Var
  S : String;
  I : Integer;

begin
  With FWord do
    begin
    If RBSerial.Checked then
      DocumentMode:=dmSerialLetter;
    If RBNewTable.Checked then
      DocumentMode:=dmNewTable;
    If RBFillTable.Checked then
      DocumentMode:=dmFillTable;
    FileName:=FEDoc.FileName;
    Save:=CBSave.Checked;
    Print:=CBPrint.Checked;
    KeepWordOpen:=CBKeepOpen.Checked;
    KeepDocumentsOpen:=CBkeepDocumentsOpen.Checked;
    FWord.KeepOnlyMergeOpen:=CBMergeOnlyOpen.Checked;
    UseFormFields:=CBFillFields.Checked;
    With CDSExport do
      If Active and not (EOF and BOF) then
        FWord.DataSource:=DSExport;
    if Save then
      Case DocumentMode of
         dmNewTable:
           SaveFileName:=FENewTableFileName.FileName;
         dmFillTable :
           SaveFileName:=FEFillTableFileName.FileName;
         dmSerialLetter :
           begin
           OutputDir:=DESave.Text;
           SaveTemplate:=CBTemplate.Text;
           If CBMergeDoc.Checked then
             begin
             CreateMergeDocument:=True;
             MergeDocumentFileName:=FEMErgeDoc.FileName;
             end;
           end;
      end;
    WordVersion:=FUseWordVersion;
    If (DocumentMode=dmNewTable) then begin
      i:=FValues.IndexOfName(CBTableTag.Text);
      If I<>-1 then
        FValues.Delete(i);
      end;
    Values:=FValues;
    OnlyCurrentRecord:=CBCurrentRecordOnly.Checked;
    if DocumentMode=dmNewTable then
      begin
      TableTag:=CBTableTag.Text;
      For I:=0 to LBFields.Items.Count-1 do
        If LBFields.Checked[i] then
          begin
          S:=LBFields.Items[I];
          If (TableTag<>S) then
            ColumnList.Add(LBFields.Items[I]);
          end;
      end;
    Screen.Cursor:=crHourGlass;
    Try
      Execute;
    Finally
      Screen.Cursor:=crDefault;
    end;
    If (DocumentCount=1) then
      S:=SSucces1
    else
      S:=Format(SSuccesN,[DocumentCount]);
    MessageDlg(S, mtInformation, [mbOK], 0);
    end;
end;

Function TWordExportForm.DocFileExists : Boolean;

begin
  If FEDoc.FileName<>FlastFileCheck then
    begin
    FlastFileCheck:=FEDoc.FileName;
    FDocExists:=(FLAstFileCheck<>'') and FileExists(FlastFileCheck);
    end;
  Result:=FDocExists;
end;

procedure TWordExportForm.AStartUpdate(Sender: TObject);

Var
  B : Boolean;

begin
  B:=DocFileExists;
  If (B and CBSave.Checked) then
    begin
    If RBNewTable.Checked then
      B:=(FENEwTableFileName.FileName<>'')
    else If RBFillTable.Checked then
      B:=(FEFillTableFileName.FileName<>'')
    end;
  If (B and CBMergeDoc.Checked) then
    B:=(FEMergeDoc.FileName<>'');
  (Sender as TAction).Enabled:=B;
end;

procedure TWordExportForm.AFetchParamsUpdate(Sender: TObject);
begin
  (Sender as TAction).Enabled:=DocFileExists;
end;

procedure TWordExportForm.ACancelExecute(Sender: TObject);
begin
  FWord.Cancel;
end;

procedure TWordExportForm.ACancelUpdate(Sender: TObject);
begin
  (Sender as TAction).Enabled:=FWord.Busy;
end;

procedure TWordExportForm.AFetchParamsExecute(Sender: TObject);

Var
  S : TStrings;
  VN : String;
  I : Integer;
  A,B : boolean;

begin
  S:=TStringList.Create;
  Try
    With FWord do
      begin
      FileName:=FEDoc.FileName;
      WordVersion:=FUseWordVersion;
      UseFormFields:=RBSerial.Checked and CBFillFields.Checked;
      GetMergeFieldNames(S);
      end;
    B:=False;
    For I:=0 to S.Count-1 do
      begin
      A:=True;
      VN:=S[i];
      If CDSExport.Active then
        A:=CDSExport.Fields.FindField(VN)=Nil;
      A:=A and (FValues.IndexOfName(VN)=-1) and (CBTableTag.Text<>VN);
      If A then
        begin
        FValues.Add(VN+'=');
        CBTableTag.Items.Add(VN);
        B:=True;
        end;
      end;
    If B then
      InitGrid;
  Finally
    S.Free;
  end;
end;

procedure TWordExportForm.InitGrid;

Var
  ACount : Integer;


begin
  With GValues do
    begin
    ACount:=FValues.Count+1;
    If (ACount<2) then
      ACount:=2;
    RowCount:=ACount;
    Invalidate;
    end;
end;

procedure TWordExportForm.GValuesDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);

Var
  P : Integer;
  S : String;
  A : TAlignment;

begin
  S:='';
  If ARow=0 then
    Case ACol of
      0 : S:=SField;
      1 : S:=SValue;
    end
  else
    begin
    Dec(Arow);
    If ARow<FValues.Count then
      begin
      S:=FValues[ARow];
      P:=Pos('=',S);
      Case ACol of
        0 : S:=Copy(S,1,P-1);
        1 : Delete(S,1,P);
      end;
      end;
    end;
  Case ACol of
    0 : A:=taRightJustify;
    1 : A:=taLeftJustify;
  else
    A:=taCenter;
  end;
  GValues.DrawStr(Rect,S,A);
end;

procedure TWordExportForm.AAddExecute(Sender: TObject);
begin
  With FValues do
    Add(SNewValue);
  With GValues do
    If FValues.Count>1 then
      RowCount:=RowCount+1;
end;

procedure TWordExportForm.ADeleteUpdate(Sender: TObject);
begin
  (Sender as TAction).Enabled:=(FValues.Count>0) and
                               (GValues.Row>0) and
                               (GValues.Row<FValues.Count+1);
end;

procedure TWordExportForm.ADeleteExecute(Sender: TObject);

Var
  R : Integer;

begin
  R:=GValues.Row;
  Dec(R);
  If (R>=0) and (R<FValues.Count) then
    begin
    FValues.Delete(R);
    With GValues do
      begin
      If RowCount>2 then
        RowCount:=RowCount-1;
      Invalidate;
      end;
    end;
end;

procedure TWordExportForm.GValuesGetEditText(Sender: TObject; ACol,
  ARow: Integer; var Value: String);

begin
  Dec(ARow);
  If (ARow>0) and (ARow<FValues.Count) then
    begin
    Value:=FValues[Arow];
    Arow:=Pos('=',Value);
    Case ACol of
      0 : Value:=Copy(Value,1,ARow-1);
      1 : Delete(Value,1,ARow);
    end;
    end;
end;

procedure TWordExportForm.GValuesSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: String);

Var
  S : String;
  P : Integer;
begin
  Dec(ARow);
  If (ARow>=0) and (ARow<FValues.Count) then
    begin
    S:=FValues[Arow];
    P:=Pos('=',S);
    Case ACol of
      0 : begin
          Delete(S,1,P-1);
          S:=Value+S;
          end;
      1 : begin
          S:=Copy(S,1,P);
          S:=S+Value;
          end;
    end;
    FValues[ARow]:=S;
    end;
end;

procedure TWordExportForm.FormShow(Sender: TObject);

Var
  D : OleVariant;
  P : TDatasetProvider;
  B : TBookMarkStr;
  F : String;
  H : Boolean;
  I : Integer;

begin
  F:='';
  H:=Assigned(FDataset);
  RBNewtable.Enabled:=H;
  RBNewtable.Enabled:=H;
  If H then
    begin
    If FDataset is TClientDataset then
      With TClientDataset(FDataset) do
        begin
        D:=Data;
        If Filtered then
          F:=Filter;
        end
    else
      begin
      P:=TDatasetProvider.Create(Self);
      try
        P.DataSet:=FDataset;
        Fdataset.DisableControls;
        Try
          B:=FDataset.Bookmark;
          D:=P.Data;
        Finally
          FDataset.BookMark:=B;
          FDataset.EnableControls;
        end;
      Finally
        P.Free;
      end;
      end;
    CDSExport.Data:=D;
    If (F<>'') then
      begin
      CDSExport.Filter:=F;
      CDSExport.Filtered:=True;
      end;
    CDSExport.Open;
    With CDSExport.Fields do
      For I:=0 to Count-1 do
        begin
        LBFields.Items.Add(Fields[i].FieldName);
        LBFields.Checked[i]:=True;
        end;
    end;
end;


procedure TWordExportForm.ACloseExecute(Sender: TObject);
begin
  Close;
end;

procedure TWordExportForm.ACloseUpdate(Sender: TObject);
begin
  (Sender as TAction).Enabled:=Not FWord.Busy;
end;

procedure TWordExportForm.CBSaveClick(Sender: TObject);
begin
  CheckFileEnabled;
end;

procedure TWordExportForm.CheckFileEnabled;

Var
  B : Boolean;

begin
  B:=CBSave.Checked;
  FENewTableFileName.Enabled:=B;
  FEFillTableFileName.Enabled:=B;
  DESave.Enabled:=B;
  CBTemplate.Enabled:=B;
  CBKeepOpen.Enabled:=B;
  If Not CBKeepOpen.Enabled then
    CBKeepOpen.Checked:=False;
  CBMergedoc.Enabled:=True;
  If Not CBMergeDoc.Enabled then
    CBMergeDoc.Checked:=False;
end;

procedure TWordExportForm.RBKindClick(Sender: TObject);
begin
  If RBSerial.Checked then
    PCOptions.ActivePage:=TSSerial
  else if RBNewTable.Checked then
    PCOptions.ActivePage:=TSNewTable
  else if RBFillTable.Checked then
    PCOptions.ActivePage:=TSFillTable;
end;

procedure TWordExportForm.AUpUpdate(Sender: TObject);
begin
  (Sender as Taction).Enabled:=LBFields.ItemIndex>0;
end;

procedure TWordExportForm.AUpExecute(Sender: TObject);

begin
  With LBFields do
    Items.Exchange(ItemIndex,ItemIndex-1);
end;

procedure TWordExportForm.ADownExecute(Sender: TObject);
begin
  With LBFields do
    Items.Exchange(ItemIndex,ItemIndex+1);
end;

procedure TWordExportForm.ADownUpdate(Sender: TObject);
begin
  With LBFields do
    (Sender as TAction).Enabled:=ItemIndex<Items.Count-1;
end;

procedure TWordExportForm.CBTableTagChange(Sender: TObject);

Var
  I : Integer;

begin
  I:=FValues.IndexOf(CBTableTag.Text);
  If (I<>-1) then
    begin
    FValues.Delete(I);
    InitGrid;
    end;
end;

procedure TWordExportForm.CBMergeDocClick(Sender: TObject);
begin
  FEMergeDoc.Enabled:=CBMergedoc.Checked;
end;

end.
