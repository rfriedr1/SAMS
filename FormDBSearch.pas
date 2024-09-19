unit FormDBSearch;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Vcl.Grids,
  Vcl.DBGrids, Vcl.ExtCtrls, System.UITypes, _dm, ADODB, Sams_Main, Vcl.ComCtrls;

type
  TfrmDBSearch = class(TForm)
    StaticText1: TStaticText;
    btnSearch: TButton;
    DBGridSearchResults: TDBGrid;
    StaticText2: TStaticText;
    lblSearchResults: TLabel;
    RadioGroupSearchContext: TRadioGroup;
    edtSearchPhrase: TLabeledEdit;
    PageControl: TPageControl;
    Mask: TTabSheet;
    Query: TTabSheet;
    MemoQuery: TMemo;
    StaticText3: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure DBGridSearchResultsDblClick(Sender: TObject);
    procedure DBGridAutoSizeColumn(Grid: TDBGrid; Column: integer);
    procedure DBGridAutoSizeAllColumns(Grid: TDBGrid);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDBSearch: TfrmDBSearch;
  ADOQuerySearch : TADOQuery;
  DataSrcSearch  : TDataSource;

implementation

{$R *.dfm}

procedure TfrmDBSearch.btnSearchClick(Sender: TObject);
// perfomr the database search
var s: string;
begin
  // use the same database connection as the main Program
  // Create the query during runtime
  ADOQuerySearch := TADOQuery.Create(Self);
  ADOQuerySearch.Connection := dm.ADOConnKTL;

  // check whether the mask or the query is being used for searching the DB
  if PageControl.ActivePageIndex = 0 then
    Begin
      // depending on the selected radio button
      // get all column names of that table
      // concatenate the columns and use the concatenated field in WHERE field LIKE
      // so all fields can be searched an ones
      if edtSearchPhrase.Text = '' then Exit;
      case RadioGroupSearchContext.ItemIndex of
        0: // Users
            s:='SELECT * from user_t WHERE CONCAT_WS(";",user_nr,last_name,first_name,organisation,address_1,address_2,town,country,institute,postcode,phone_1,phone_2,email,account,user_comment) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
        1: // Projects
            s:='SELECT * from project_t WHERE CONCAT_WS(";",project_nr,project,user_nr,in_date,out_date,invoice,AuftragsNr,order_nr,invoice_nr,letter,project_comment,report,sample_storage_loc) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
        2: // samples
            // concatenate all fields of the table and search the resulting string
            s := 'SELECT * from sample_t WHERE CONCAT_WS(";",sample_nr,project_nr,type,material,fraction,weight,sampling_date,user_label,user_label_nr,user_desc1,user_desc2,MA_nr,lab_comment,user_comment,prep_storage_loc,lab_comment,storage) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
        3: // prep
            // concatenate all fields of the table and search the resulting string
            s := 'SELECT * from preparation_t WHERE CONCAT_WS(";",sample_nr,pre_nr,batch,cn_ratio,c_percent,n_percent,prep_end,prep_start,prep_comment) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
        4: // target
            // concatenate all fields of the table and search the resulting string
            s := 'SELECT * from target_t WHERE CONCAT_WS(";",sample_nr,target_nr,prep_nr,magazine,position,target_comment,meas_comment,graph_batch,weight,conc_c,conc_c,target_id) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
      end;
    End
  else
    Begin
      if MemoQuery.Lines.Text = '' then Exit;
      s := memoQuery.Lines.Text;
    End;
  //showmessage(s);
  // concatenate all fields of the table and search the resulting string
  //s := 'SELECT * from sample_t WHERE CONCAT_WS('+#34+';'+#34+',sample_nr,project_nr,photo,type,material,user_label) LIKE ' + #34 + '%' + edtSearchPhrase.Text + '%' + #34 +';';
  ADOQuerySearch.SQL.Add(s);
  ADOQuerySearch.Open;

// Create the data source.
  DataSrcSearch := TDataSource.Create(Self);
  DataSrcSearch.DataSet := ADOQuerySearch;
  DataSrcSearch.Enabled := true;

//Finally, initialize the grid.
  DBGridSearchResults.DataSource := DataSrcSearch;

//  with DBGridSearchResults do
//  begin
//    DBGrid1.Columns.Items[0].Width := 60;
//    DBGrid1.Columns.Items[1].Width := 120;
//    DBGrid1.Columns.Items[2].Width := 60;
//    DBGrid1.Columns.Items[3].Width := 60;
//  end;

// auto scale the column width's
DBGridAutoSizeAllColumns(DBGridSearchResults);

end;

procedure TfrmDBSearch.DBGridSearchResultsDblClick(Sender: TObject);
begin
// display info of the selected project or sample in main SAMS window
// depending on the radio group selection, perform different actions
    case RadioGroupSearchContext.ItemIndex of // Oder eben "case TRadioGroupN.ItemIndex of"
      0: //Users
       begin
        // display user in SAMS
        frmMAMS.cmbSubmitterNameProject.KeyValue:=DBGridSearchResults.DataSource.DataSet.FieldByName('user_nr').AsString;
        frmMAMS.cmbSubmitterNameProjectCloseUp(Self);
        frmMAMS.actUserProjectsExecute(Self);
       end;
      1: //Projects
       begin
       // Display all projects of the selected user in SAMS
        frmMAMS.cmbSubmitterNameProject.KeyValue:=DBGridSearchResults.DataSource.DataSet.FieldByName('user_nr').AsString;
        frmMAMS.cmbSubmitterNameProjectCloseUp(Self);
        frmMAMS.actUserProjectsExecute(Self);
        frmMAMS.grdProjects.DataSource.DataSet.Locate('project_nr',DBGridSearchResults.DataSource.DataSet.FieldByName('project_nr').AsString,[loPartialKey]);
        frmMAMS.grdProjectsCellClick(Nil);
       end;
      2: // Samples
       begin
         // set sample number in SAMS
          frmMAMS.edtSampleNr.Value:=DBGridSearchResults.DataSource.DataSet.FieldByName('sample_nr').AsString;
        // simulate a click on the sample info button in SAMS
          frmMAMS.actSampleInfoExecute(Self);
       end;
      3: // Prep
       begin
         // set sample number in SAMS
          frmMAMS.edtSampleNr.Value:=DBGridSearchResults.DataSource.DataSet.FieldByName('sample_nr').AsString;
        // simulate a click on the sample info button in SAMS
          frmMAMS.actSampleInfoExecute(Self);
       end;
      4: // target
       begin
         // set sample number in SAMS
          frmMAMS.edtSampleNr.Value:=DBGridSearchResults.DataSource.DataSet.FieldByName('sample_nr').AsString;
        // simulate a click on the sample info button in SAMS
          frmMAMS.actSampleInfoExecute(Self);
       end;
    end;
end;

procedure TfrmDBSearch.FormShow(Sender: TObject);
begin
  PageControl.ActivePage := Mask;
// check whether DB connection is active or not
  try
    dm.ADOConnKTL.Connected := True;
  except
    on e: EADOError do
    begin
      MessageDlg('Error while connecting to DB', mtError,
                  [mbOK], 0);

      Exit;
    end;
  end;
end;

procedure TfrmDBSearch.DBGridAutoSizeColumn(Grid: TDBGrid; Column: integer);
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

procedure TfrmDBSearch.DBGridAutoSizeAllColumns(Grid: TDBGrid);
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

end.
