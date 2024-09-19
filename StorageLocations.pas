unit StorageLocations;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Vcl.Grids, Vcl.DBGrids,System.UITypes, frmLogWindow, _dm,ADODB, Vcl.ComCtrls;

type
  TFormStorageLocations = class(TForm)
    edtStartSampleID: TLabeledEdit;
    edtEndSampleID: TLabeledEdit;
    btnClose: TButton;
    btnSave: TButton;
    btnSearch: TButton;
    DBGrid1: TDBGrid;
    edtStatus: TEdit;
    GroupBox1: TGroupBox;
    BalloonHint1: TBalloonHint;
    WarningLabel: TLabel;
    UpDownSampleIDStart: TUpDown;
    UpDownSampleIDEnd: TUpDown;
    lblNumberOfShownSamples: TLabel;
    lblNumberofNoLeftoverSamples: TLabel;
    edtStorLoc: TLabeledEdit;
    RadioGroupLocations: TRadioGroup;
    procedure btnSaveClick(Sender: TObject);
    procedure edtStartSampleIDChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure UpDownSampleIDStartClick(Sender: TObject; Button: TUDBtnType);
    procedure UpDownSampleIDEndClick(Sender: TObject; Button: TUDBtnType);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormStorageLocations: TFormStorageLocations;
  ADOQueryIDs : TADOQuery;
  ADOQueryNoLeftover : TADOQuery;
  DataSrc  : TDataSource;

implementation

{$R *.dfm}

procedure TFormStorageLocations.btnSearchClick(Sender: TObject);
Var
  s: string;
  ReturnToSender, Returned, CNAna, CNAnaMoved: integer;

begin
//create datasource and query, ADOConnection has been checked at FormShow, uses the ADOConnection from _dm

// Create the query for all samples with leftover
  ADOQueryIDs := TADOQuery.Create(Self);
  ADOQueryIDs.Connection := dm.ADOConnKTL;

  // generate different queries depending on the storagle location that should be assigned.
  if RadioGroupLocations.ItemIndex = 0 then   // sample storage location
  begin
    s := 'SELECT sample_nr, user_label, material, s_storage_loc AS Sample_Loc, ' +
    'prep_storage_loc AS Prep_Loc, return_to_sender AS Return_Sample, ' +
    'returned_to_sender as Returned, ' +
    'CNIsotopA as CNAna, ' +
    'CNIsotopAMoved as CNAna_Moved ' +
    'FROM sample_t ' +
    'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr ' +
    'WHERE s_no_leftover = 0 AND sample_t.sample_nr BETWEEN ' + edtStartSampleID.Text + ' AND ' + edtEndSampleID.Text + ';';
  end;
  if RadioGroupLocations.ItemIndex = 1 then   // prep storage location
  begin
    s := 'SELECT sample_t.sample_nr, user_label, material, s_storage_loc AS Sample_Loc, ' +
    'prep_storage_loc AS Prep_Loc, return_to_sender AS Return_Sample, ' +
    'returned_to_sender as Returned, ' +
    'CNIsotopA as CNAna, ' +
    'CNIsotopAMoved as CNAna_Moved ' +
    'FROM sample_t ' +
    'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr ' +
    'INNER JOIN preparation_t ON sample_t.sample_nr = preparation_t.sample_nr ' +
    'WHERE p_no_leftover = 0 AND sample_t.sample_nr BETWEEN ' + edtStartSampleID.Text + ' AND ' + edtEndSampleID.Text + ';';
  end;
  // ShowMessage(s);
  ADOQueryIDs.SQL.Add(s);
  ADOQueryIDs.Open;
// display number of shown samples
  if RadioGroupLocations.ItemIndex = 0 then   // sample storage location
  begin
       lblNumberOfShownSamples.Caption := 'Samples with leftover (displayed in table): ' + ADOQueryIDs.RecordCount.ToString;
  end;
  if RadioGroupLocations.ItemIndex = 1 then   // prep storage location
  begin
       lblNumberOfShownSamples.Caption := 'Prep with leftover (displayed in table): ' + ADOQueryIDs.RecordCount.ToString;
  end;
// Create the data source.
  DataSrc := TDataSource.Create(Self);
  DataSrc.DataSet := ADOQueryIDs;
  DataSrc.Enabled := true;
//Finally, initialize the grid.
  DBGrid1.DataSource := DataSrc;
  with DBGrid1 do
  begin
    DBGrid1.Columns.Items[0].Width := 60;
    DBGrid1.Columns.Items[1].Width := 120;
    DBGrid1.Columns.Items[2].Width := 70;
    DBGrid1.Columns.Items[3].Width := 70;
    DBGrid1.Columns.Items[4].Width := 70;
    DBGrid1.Columns.Items[5].Width := 77;
    DBGrid1.Columns.Items[6].Width := 60;
    DBGrid1.Columns.Items[7].Width := 55;
    DBGrid1.Columns.Items[8].Width := 75;
  end;

  // count number of samples oder prep that have x_no_leftover = 1
  ADOQueryNoLeftover := TADOQuery.Create(Self);
  ADOQueryNoLeftover.Connection := dm.ADOConnKTL;
  if RadioGroupLocations.ItemIndex = 0 then   // sample storage location
  begin
  s := 'SELECT sample_nr, user_label, material, s_storage_loc AS Sample_Loc, ' +
  'prep_storage_loc AS Prep_Loc, return_to_sender AS Return_Sample, ' +
  'returned_to_sender as Returned, ' +
  'CNIsotopA as CNAna, ' +
  'CNIsotopAMoved as CNAna_Moved ' +
  'FROM sample_t ' +
  'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr ' +
  'WHERE s_no_leftover = 1 AND sample_t.sample_nr BETWEEN ' + edtStartSampleID.Text + ' AND ' + edtEndSampleID.Text + ';';
  end;
  if RadioGroupLocations.ItemIndex = 1 then   // prep storage location
  begin
  s := 'SELECT sample_t.sample_nr, user_label, material, s_storage_loc AS Sample_Loc, ' +
  'prep_storage_loc AS Prep_Loc, return_to_sender AS Return_Sample, ' +
  'returned_to_sender as Returned, ' +
  'CNIsotopA as CNAna, ' +
  'CNIsotopAMoved as CNAna_Moved ' +
  'FROM sample_t ' +
  'INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr ' +
  'INNER JOIN preparation_t ON sample_t.sample_nr = preparation_t.sample_nr ' +
  'WHERE p_no_leftover = 1 AND sample_t.sample_nr BETWEEN ' + edtStartSampleID.Text + ' AND ' + edtEndSampleID.Text + ';';
  end;
  // ShowMessage(s);
  ADOQueryNoLeftover.SQL.Add(s);
  ADOQueryNoLeftover.Open;

  // display number of shown samples
  if RadioGroupLocations.ItemIndex = 0 then   // sample storage location
  begin
    lblNumberofNoLeftoverSamples.Caption := 'Samples with no leftover (not displayed in table): ' + ADOQueryNoLeftover.RecordCount.ToString;
  end;
  if RadioGroupLocations.ItemIndex = 1 then   // prep storage location
  begin
    lblNumberofNoLeftoverSamples.Caption := 'Prep with no leftover (not displayed in table): ' + ADOQueryNoLeftover.RecordCount.ToString;
  end;

  // set status of some buttons
  btnSave.Enabled:=true;
  WarningLabel.Visible := False;

  // check whether or not to display the warning label that samples
  // exist that need to be returned to to CN
  // go through all records
  ADOQueryIDs.DisableControls;
  ADOQueryIDs.First;
  while not ADOQueryIDs.Eof do
    begin
      ReturnToSender := ADOQueryIDs.FieldByName('Return_Sample').AsInteger;
      Returned := ADOQueryIDs.FieldByName('Returned').AsInteger;
      CNAna := ADOQueryIDs.FieldByName('CNAna').AsInteger;
      CNAnaMoved := ADOQueryIDs.FieldByName('CNAna_Moved').AsInteger;
      if (ReturnToSender-Returned = 1) OR (CNAna-CNAnaMoved = 1) Then
        begin
          WarningLabel.Visible := True;
          WarningLabel.Caption := 'Attention: Return to sender / CN Analysis';
        end;
      ADOQueryIDs.Next;
    end;
    ADOQueryIDs.First;
    ADOQueryIDs.EnableControls;

  //ADOQuery.Close;
end;

procedure TFormStorageLocations.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
VAR ReturnToSender, Returned, CNAna, CNAnaMoved: integer;
CONST
  CtrlState: array[0..1] of integer = (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED) ;
begin
  ReturnToSender := DBGrid1.DataSource.DataSet.FieldByName('Return_Sample').AsInteger;
  Returned := DBGrid1.DataSource.DataSet.FieldByName('Returned').AsInteger;
  CNAna := DBGrid1.DataSource.DataSet.FieldByName('CNAna').AsInteger;
  CNAnaMoved := DBGrid1.DataSource.DataSet.FieldByName('CNAna_Moved').AsInteger;

   // display boolean fields as checkboxes
  if (Column.FieldName='Return_Sample') OR
  (Column.FieldName='Returned') OR
  (Column.FieldName='CNAna') OR
  (Column.FieldName='CNAna_Moved') then
  begin
    TDBGrid(Sender).Canvas.FillRect(Rect) ;
    if (VarIsNull(Column.Field.Value)) then
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, DFCS_BUTTONCHECK or DFCS_INACTIVE) {grayed}
    else
      DrawFrameControl(TDBGrid(Sender).Canvas.Handle,Rect, DFC_BUTTON, CtrlState[Column.Field.AsInteger]) ; {checked or unchecked}
  end;

  // change cell color to red when return to sender is true
  if Column.FieldName='Return_Sample' then
  begin
      if ReturnToSender-Returned = 1 Then
          Begin
          TDBGrid(Sender).Canvas.Brush.Color:=clRed;
          End;
  end
  else TDBGrid(Sender).Canvas.Brush.Color:=clWindow;

  // change row color to red when return to sender is true
  if Column.FieldName='CNAna' then
  begin
      if CNAna-CNAnaMoved = 1 Then
          Begin
          TDBGrid(Sender).Canvas.Brush.Color:=clRed;
          End;
  end
  else TDBGrid(Sender).Canvas.Brush.Color:=clWindow;

  //TDBGrid(Sender).DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TFormStorageLocations.btnSaveClick(Sender: TObject);
// performs a database query in order to set the storage locations
// of teh selected samples
Var
  ADOCmdUpdate : TADOCommand;
  s, s2: string;
  i, button: integer;
  can_update_flag1: boolean; // flag that determines whether the record can be updated for Sample_Loc
  can_update_flag2: boolean; // flag that determines whether the record can be updated for Prep_loc
  can_update_flag3: boolean; // flag that determines whether the record can be updated for combined sample and pre loc

begin
//check if the storage location field contains anything
if (trim(edtStorLoc.Text) <> '') then
  begin
    // Create the query
  ADOCmdUpdate := TADOCommand.Create(Self);
  ADOCmdUpdate.Connection := dm.ADOConnKTL;

  // set query with sampleID#s to first record
  ADOQueryIDs.First;

  // go through data that are displayed in DBGrid
  // create the UPDATE query depending on what data are given
  // then update records in DB
  while NOT ADOQueryIDs.Eof do
    begin
      // check first if there are already values stored in the database
      // or certain values such as 0
      // then set flag accordingly
      // for sample location
      LogWindow.addLogEntry('storrage locations -- check if record update is allowed');
      can_update_flag1:= true;
      if (ADOQueryIDs.FieldByName('Sample_Loc').AsString <> '') then can_update_flag1:= false
      else if (ADOQueryIDs.FieldByName('Sample_Loc').AsString = '0') then can_update_flag1:= true
      else if (ADOQueryIDs.FieldByName('Sample_Loc').AsString = ' ') then can_update_flag1:= true;
      LogWindow.addLogEntry('storrage locations -- sample_loc_flag=' + booltostr(can_update_flag1));

      // for prep location
      can_update_flag2:= true;
      if (ADOQueryIDs.FieldByName('Prep_Loc').AsString <> '') then can_update_flag2:= false
      else if (ADOQueryIDs.FieldByName('Prep_Loc').AsString = '0') then can_update_flag2:= true
      else if (ADOQueryIDs.FieldByName('Prep_Loc').AsString = ' ') then can_update_flag2:= true;
      LogWindow.addLogEntry('storrage locations -- pre_loc_flag=' + booltostr(can_update_flag2));

      // see if one of the above flags is triggered
      // then show message and let user decide what to do
//      can_update_flag3:= false;
//      if ((can_update_flag1 = true) AND (can_update_flag2 = true)) then can_update_flag3:= true
//      else
//        begin
//          button:= messagedlg('sample nr: ' + ADOQueryIDs.FieldByName('sample_nr').AsString + ' location already exists - OVERWRITE?', mtWarning, [mbYes, mbNo], 0, mbNo);
//          if button = mrYes then can_update_flag3:= true;
//        end;
//      LogWindow.addLogEntry('storrage locations -- allow update=' + booltostr(can_update_flag3));

      // if user wants to update sample_storrage location although there is already a value given
      // show error and  let user decide
    if RadioGroupLocations.ItemIndex = 0 then   // sample storage location
    begin
        LogWindow.addLogEntry('storrage locations -- check flags and user input');
        if ((trim(edtStorLoc.Text) <> '') AND (can_update_flag1 = false)) then
          begin
            button:= messagedlg('sample nr: ' + ADOQueryIDs.FieldByName('sample_nr').AsString + ' SAMPLE location already exists - OVERWRITE?', mtWarning, [mbYes, mbNo], 0, mbNo);
            if button = mrYes then can_update_flag1:= true;
            LogWindow.addLogEntry('storrage locations -- sample_loc_flag=' + booltostr(can_update_flag3));
          end;
    end;
    if RadioGroupLocations.ItemIndex = 0 then   // prep storage location
    begin
      // if user wants to update prep_storrage location although there is already a value given
      // show error and  let user decide
      if ((trim(edtStorLoc.Text) <> '') AND (can_update_flag2 = false)) then
        begin
          button:= messagedlg('sample nr: ' + ADOQueryIDs.FieldByName('sample_nr').AsString + ' PREP location already exists - OVERWRITE?', mtWarning, [mbYes, mbNo], 0, mbNo);
          if button = mrYes then can_update_flag2:= true;
          LogWindow.addLogEntry('storrage locations -- sample_loc_flag=' + booltostr(can_update_flag2));
        end;
    end;

      // if flag is true then perfrom update for that sample
      if ((can_update_flag1 = true) OR (can_update_flag2 = true)) then
        begin
          s := 'UPDATE sample_t SET ';
          if ((RadioGroupLocations.ItemIndex = 0) AND (can_update_flag1 = true)) then //SampleStorageLocation is given
            begin
            s := s + 's_storage_loc = ' + #34 + trim(edtStorLoc.Text) + #34;
            end;
          if ((RadioGroupLocations.ItemIndex = 1) AND (can_update_flag2 = true)) then  //PrepStorageLocation is given
            begin
            s := s + 'prep_storage_loc = ' + #34 + trim(edtStorLoc.Text) + #34;
            end;
          s := s+ ' WHERE sample_nr = ' + ADOQueryIDs.FieldByName('sample_nr').AsString + ';';

          LogWindow.addLogEntry('storrage locations -- Query: ' + s);

          s2 := 'updating Sample_nr = ' +  ADOQueryIDs.FieldByName('sample_nr').AsString;
          edtStatus.Text := s2;
          LogWindow.addLogEntry('storrage locations -- ' + s2);
          //ShowMessage(s);
          ADOCmdUpdate.CommandText := s;
          ADOCmdUpdate.Execute;
        end;

        ADOQueryIDs.Next;
    end;
    //show an updated DBGrid by executing the SearchClick
    btnSearchClick(self);
  end;

end;

procedure TFormStorageLocations.edtStartSampleIDChange(Sender: TObject);
//display the same string in the EndSampleID edit field
begin
 edtEndSampleID.Text := (Sender AS TLabeledEdit).Text;
end;

procedure TFormStorageLocations.FormShow(Sender: TObject);
begin

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

// Some settings
  WarningLabel.Visible := False;
end;


procedure TFormStorageLocations.UpDownSampleIDEndClick(Sender: TObject;
  Button: TUDBtnType);
Var
SampleNr: integer;
begin
  begin
    // change sample nr according to the action either one up or one down
    case Button of
      btNext:
            begin
              SampleNr := StrToInt(edtEndSampleID.Text) + 1;
              edtEndSampleID.Text := IntToStr(SampleNr);
            end;
      btPrev:
            Begin
              SampleNr := StrToInt(edtEndSampleID.Text) - 1;
              edtEndSampleID.Text := IntToStr(SampleNr);
            End;
    end;
  end;
end;

procedure TFormStorageLocations.UpDownSampleIDStartClick(Sender: TObject;
  Button: TUDBtnType);
Var
SampleNr: integer;
  begin
    // change sample nr according to the action either one up or one down
    case Button of
      btNext:
            begin
              SampleNr := StrToInt(edtStartSampleID.Text) + 1;
              edtStartSampleID.Text := IntToStr(SampleNr);
            end;
      btPrev:
            Begin
              SampleNr := StrToInt(edtStartSampleID.Text) - 1;
              edtStartSampleID.Text := IntToStr(SampleNr);
            End;
    end;
  end;

end.
