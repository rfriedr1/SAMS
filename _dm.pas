unit _dm;

interface

uses
  SysUtils, Classes, DB, ADODB, TypInfo(*, frxClass, frxDBSet*), Dialogs, Controls,
  Datasnap.DBClient;

{
Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=DMYSQL_KTL;Extended Properties="DSN=DMYSQL_KTL;UID=root;"

for qrySampleInfo at design time :

 SELECT sample_t.sample_nr, type, material, fraction, pre_sub_treat, weight, preparation,
 sampling_date, editable, not_tobedated, user_label, sample_t.user_label_nr, user_desc1,
 user_desc2, residue, sample_t.user_comment,project_t.project, invoice_nr, in_date,
 desired_date, priority, status, price,   user_t.last_name, preparation_t.prep_nr, batch,
 prep_comment, weight_start, weight_medium, weight_end, prep_end, step1_methode,
 step2_methode,step3_methode,step4_methode,step5_methode , target_t.target_nr, cn_ratio,
 c_percent, n_percent, preparation_t.stop, magazine, position, precis, cycle_min, cycle_max,
 catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_init, hydro_final, react_time,
 target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcset, editallowed,
 target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, target_t.cal1sMax, target_t.cal2sMin,
 target_t.cal2sMax  FROM sample_t
 INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr
 INNER JOIN user_t ON user_t.user_nr=project_t.user_nr
 INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr
 INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr
 WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND target_nr=1;
 }
 
type
  Tdm = class(TDataModule)
    qrySubmit: TADOQuery;
    qryDB: TADOQuery;
    qrySampleOfSubmitter: TADOQuery;
    dsSampleOfSubmitter: TDataSource;
//    qryAllUser: TADOQuery;
    dsQryDb: TDataSource;
    qryProjectsOfUser: TADOQuery;
    dsProjectsOfUser: TDataSource;
    qrySamplesOfProject: TADOQuery;
    dsSamplesOfProject: TDataSource;
    qrySample: TADOQuery;
    qryTarget: TADOQuery;
    dsTarget: TDataSource;
    dsSample: TDataSource;
    adoConnKTL: TADOConnection;
    tblUser: TADOTable;
    tblUseruser_nr: TAutoIncField;
    tblUserfirst_name: TStringField;
    tblUserlast_name: TStringField;
    tblUserorganisation: TStringField;
    tblUserinstitute: TStringField;
    tblUseraddress_1: TStringField;
    tblUseraddress_2: TStringField;
    tblUsertown: TStringField;
    tblUserpostcode: TStringField;
    tblUsercountry: TStringField;
    tblUserphone_1: TStringField;
    tblUserphone_2: TStringField;
    tblUserfax: TStringField;
    tblUseremail: TStringField;
    tblUserwww: TStringField;
    tblUseraccount: TStringField;
    tblUserinvoice: TSmallintField;
    tblUsercorrespondance: TSmallintField;
    tblUseruser_comment: TMemoField;
    tblProjects: TADOTable;
    dsProjects: TDataSource;
    tblTypes: TADOTable;
    dsTypes: TDataSource;
    tblMaterial: TADOTable;
    dsMaterial: TDataSource;
    tblProjectTypes: TADOTable;
    dsProjectTypes: TDataSource;
    tblMethod: TADOTable;
    dsMethod: TDataSource;
    tblResearchType: TADOTable;
    dsResearchType: TDataSource;
    tblReportType: TADOTable;
    dsReportType: TDataSource;
    qrySamplesAvailable: TADOQuery;
    dsSamplesAvailable: TDataSource;
    adoCmd: TADOCommand;
    qryPrepSteps: TADOQuery;
    dsPrepSteps: TDataSource;
    qryTargetsAvailable: TADOQuery;
    dsTargetsAvailable: TDataSource;
    tblFraction: TADOTable;
    dsFraction: TDataSource;
    qryTest: TADOQuery;
    dsTest: TDataSource;
    qrySampleInfo: TADOQuery;
    dsSampleInfo: TDataSource;
    tblProjectStatus: TADOTable;
    dsProjectStatus: TDataSource;
    qryPlanned: TADOQuery;
    dsPlanned: TDataSource;
    qryInPrep: TADOQuery;
    dsInPrep: TDataSource;
    qrySamplesOfLabTask: TADOQuery;
    dsSamplesOfLabtask: TDataSource;
    qryDb1: TADOQuery;
    dsQryDb1: TDataSource;
    qryActiveBatches: TADOQuery;
    dsActiveBatches: TDataSource;
    qryProject: TADOQuery;
    dsProject: TDataSource;
    qryProjectsSinceYear: TADOQuery;
    dsProjectsSinceYear: TDataSource;
    qryWeights: TADOQuery;
    dsWeights: TDataSource;
    qryGraphWeight: TADOQuery;
    dsGraphWeight: TDataSource;
    dsUserInfo: TDataSource;
    dsMagazines: TDataSource;
    qryMagazines: TADOQuery;
    qryMagazineData: TADOQuery;
    dsMagazineData: TDataSource;
    qryWaitingForGraph: TADOQuery;
    dsWaitingForGraph: TDataSource;
    tblInvoice: TADOTable;
    AutoIncField1: TAutoIncField;
    StringField1: TStringField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    StringField5: TStringField;
    StringField6: TStringField;
    StringField7: TStringField;
    StringField8: TStringField;
    StringField9: TStringField;
    StringField10: TStringField;
    StringField11: TStringField;
    StringField12: TStringField;
    StringField13: TStringField;
    StringField14: TStringField;
    StringField15: TStringField;
    SmallintField1: TSmallintField;
    SmallintField2: TSmallintField;
    MemoField1: TMemoField;
    dsInvoice: TDataSource;
    adoConnectionCEZA: TADOConnection;
    adoCmdCEZA: TADOCommand;
    qryCEZA: TADOQuery;
    qryProjectsOfReport: TADOQuery;
    dsProjectsOfReport: TDataSource;
    dsCEZA: TDataSource;
    qryTargetData: TADOQuery;
    dsTargetData: TDataSource;
    qryPendingReports: TADOQuery;
    dsPendingReports: TDataSource;
    qrySample1: TADOQuery;
    qrySample2: TADOQuery;
    ADOConnection1: TADOConnection;
    qryProjectHasResults: TADOQuery;
    qryDBPlot: TADOQuery;
    dsDBPlot: TDataSource;
    qryWaitingForMeas: TADOQuery;
    dsWaitingForMeas: TDataSource;
    qryWaitingExpress: TADOQuery;
    dsWaitingExpress: TDataSource;
    tblUsersalutation: TStringField;
    tblUserlanguage: TStringField;
    dsLeHeCurrents: TDataSource;
    cdsLeHeCurrents: TClientDataSet;
    cdsLeHeCurrentsSample_nr: TStringField;
    cdsLeHeCurrentsPrep_nr: TStringField;
    cdsLeHeCurrentsTarget_nr: TStringField;
    cdsLeHeCurrentsLECurrent: TStringField;
    cdsLeHeCurrentsHECurrent: TStringField;
    qryMagazinesUnpressed: TADOQuery;
    dsMagazinesUnpressed: TDataSource;
    qrySamplesOfUnpressedMagazine: TADOQuery;
    dsSamplesOfUnpressedMagazine: TDataSource;
    tblMethodmethod: TStringField;
    tblMethoddescr: TMemoField;
    tblMethodindexnr: TIntegerField;
    qryWaitingForMeasAll: TADOQuery;
    dsWaitingForMeasAll: TDataSource;
    procedure adoConnKTLExecuteComplete(Connection: TADOConnection;
      RecordsAffected: Integer; const Error: Error;
      var EventStatus: TEventStatus; const Command: _Command;
      const Recordset: _Recordset);
    procedure tblMethoddescrGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
  private
   //
  public
    function AddNewProjectByUserNr(user_nr, project_name: string): String;
    function AddNewSampleByProjectNr(project_nr, sample_name: string): String;
    function AddNewUser: String;
    function CheckExistingDBValue(table, column, value:string) : boolean;
    function CheckProjectOfUser_nr(project_name: string; user_nr: integer): boolean;
    function CheckSampleOfProjectNr(sample_name, project_nr: string): boolean;
    function InsertIntoDB(table, column, value: string): boolean;
    procedure CheckProjectStatus;
    procedure CreateBlankPrepRecord(SampleNr, PrepNr : integer);
    procedure CreateBlankTargetRecord(SampleNr, PrepNr, TargetNr : integer);
    procedure DeleteFromDb(SampleNr : integer);
    function GetAllInPrep : integer;
    function GetAllPlanned(ShowOnHold : boolean; material: string = 'none') : integer;
    function GetAllWaitingForGraph : integer;
    function GetAllWaitingForMeas : integer;
    function GetAllWaitingForMeasAll : integer;
    function GetWaitingExpress : integer;
    function GetArchSamplesOfMonth(Year, Month : word) : integer;
    function GetAllSamplesOfMonth(Year, Month : word; flag: integer) : integer;
    procedure GetMagazineData(Magazine : string);
    procedure GetMagazineDataUnpressed(Magazine : string);
    function GetMaxPrepNrBySampleNr(sample_nr : integer) : integer;
    function GetMaxSampleNr : integer;
    function GetMaxTargetNrBySampleNr(sample_nr, prep_nr : integer) : integer;
    procedure GetNOxaBlank(Oxa : boolean);
    procedure GetNIAEA(stdname: string);
    procedure GetNPferde();
    function GetPretreatStepNrBySampleNr(sample_nr, prep_nr : integer; const method : string) : integer;
    function GetProjectNrBySampleNr(sample_nr : integer) : integer;
    function GetProjectNrByProjectAndUserNr(project_name, user_nr: string): string;
    function GetProjectByProjectNr(project_nr: integer): string;
    procedure GetProjectsSinceYear(StartYear : integer);
    function GetSampleNrBySampleNameAndProjectNr(sample_name, project_nr: string): string;
    function GetSamplesByProjectNr(project_nr: integer): integer;
    function RemoveProjectByProjectNr(project_nr: integer): integer;
    function GetUserByProjectNr(project_nr: integer): string;
    function GetUserNrByLastName(last_name: string): String;
    procedure GetSamplesAvailableForPrep(const method : string);
    procedure GetSamplesAvailableForPrepByMaterial(const material : string);
    procedure GetSamplesAvailableForPrepByMethod(const method : string);
    procedure GetSampleInfo(SampleNr, PrepNr, TargetNr : integer);
    procedure GetTargetsAvailable;
    procedure GetWeights(SampleNr, PrepNr : integer);
    procedure GetGraphWeight(SampleNr, PrepNr, TargetNr : integer);
    function QueryInvoiceUser(const LastName, Institute, Organisation : string) : integer;
    procedure QueryPrepStepsBySampleNr(sample_nr, prep_nr : integer);
    procedure QueryTargetInfoBySampleNr(sample_nr: Integer);
    procedure QueryProject(const ProjectName : string);
    procedure QuerySampleBySampleNr(sample_nr : integer);
    procedure SetGraphDate;
    procedure SetGraphitized;
    procedure SetProjectStatusRunning(sample_nr : integer);
    procedure StrToClipBoard(const s: string; msDelay : integer);
    procedure SwapResults(Nr1, Nr2 : integer);
    function TargetRecordExistsForSample(SampleNr : integer) : boolean;
    procedure TransferAgeFromTarget(SampleNr, PrepNr, TargetNr : integer);
    procedure TransferMA_Nr_To_MAMS(limit: string);
    procedure TransferLeHeCurrentToTarget(SampleNr, PrepNr, TargetNr, le_current, he_current : string);
    //procedure SendToLog(s:string);
    function ReplaceUmlaute(s: string): string;
    function ReplaceBadCharacters(s: string): string;
    procedure DBreconnect;
    function ExtractNumber(Input: string): string;
  end;

var
  dm: Tdm;

implementation

{$R *.dfm}

{ Tdm }

uses
  Clipbrd, DateUtils, SAMS_Main, _LogSQL, frmLogWindow;

//procedure Tdm.SendToLog(s:string);
////sends the string to the log window if the window is open
//begin
//  if LogWindow.Showing then
//  begin
//    LogWindow.ListBox.Items.Add(s);
//  end;
//end;

procedure Tdm.adoConnKTLExecuteComplete(Connection: TADOConnection;
  RecordsAffected: Integer; const Error: Error; var EventStatus: TEventStatus;
  const Command: _Command; const Recordset: _Recordset);
begin
{  FrmLogSql.AddLogSQL(
  Command.CommandText,
  'After '+GetEnumName(TypeInfo(TCommandType),Integer(Command.CommandType)),
  GetEnumName(TypeInfo(TEventStatus),Integer(EventStatus)),
  GetEnumName(TypeInfo(TCursorType),Integer(Recordset.CursorType)),
  GetEnumName(TypeInfo(TADOLockType),Integer(Recordset.LockType)),
  RecordsAffected);
  }
end;


function Tdm.AddNewProjectByUserNr(user_nr, project_name: string): String;
// add a new project to the project_t (only project_name)
// that is associated to the user_nr
// function returns the project_nr of the newly created project
// or 'none' if no project was created
Var
  ADOCmd: TADOCommand;
begin
 // insert project into project_t
 if adoConnKTL.Connected then
 begin           // only perform command if connection is established (use ADOCommand because it can also handle queries that don't return datasets)
      ADOCmd:= TADOCommand.Create(nil);       // create new command object
      try
        ADOCmd.Connection:=adoConnKTL;        // set DB connection
        ADOCmd.Parameters.Clear;              // clear all Parameters
        ADOCmd.CommandType:=cmdtext;          // set command type to text
        ADOCmd.CommandText:='INSERT INTO project_t (project, user_nr, in_date, desired_date, status, out_date) VALUES ( '  + #34 + project_name + #34 + ',' + #34 + user_nr + #34 + ',' + #34 + FormatDateTime('YYYY-MM-DD', Date) + #34 + ',' + #34 + FormatDateTime('YYYY-MM-DD', IncMonth(Date,+3)) + #34 + ',' + #34 + 'planned' + #34 + ',' + 'NULL' + ');';    // set query text
        //ShowMessage(ADOCmd.CommandText);
        //ADOCmd.Parameters.ParamByName('param1').Value := last_name;      //insert parameter value into query
        LogWindow.addLogEntry(ADOCmd.CommandText);
        ADOCmd.Execute; // issue command (no result set must be returned)
      except
        ShowMessage('Unknown error encountered while inserting into DB.');   //somehting went wrong
        result:='none';
      end;
      ADOCmd.Free;                            // cleanup command object
      result:= GetProjectNrByProjectAndUserNr(project_name, user_nr);
    end
  else
  begin
        ShowMessage('Database not connected.');   //Database is not connected
        result:= 'none';
      end;

end;


function Tdm.AddNewSampleByProjectNr(project_nr, sample_name: string): String;
// add a new sample to the sample_t (only user_label)
// that is associated to the preoject_nr
// function returns the sample_nr of the newly created sample
// or 'none' if no sample was created
Var
  ADOCmd: TADOCommand;
begin
 // insert sample into sample_t
 if adoConnKTL.Connected then
 begin           // only perform command if connection is established (use ADOCommand because it can also handle queries that don't return datasets)
      ADOCmd:= TADOCommand.Create(nil);       // create new command object
      try
        ADOCmd.Connection:=adoConnKTL;        // set DB connection
        ADOCmd.Parameters.Clear;              // clear all Parameters
        ADOCmd.CommandType:=cmdtext;          // set command type to text
        ADOCmd.CommandText:='INSERT INTO sample_t (user_label, project_nr) VALUES ( '  + #34 + sample_name + #34 + ',' + #34 + project_nr + #34 + ');';    // set query text
        //ShowMessage(ADOCmd.CommandText);
        //ADOCmd.Parameters.ParamByName('param1').Value := last_name;      //insert parameter value into query
        ADOCmd.Execute; // issue command (no result set must be returned)
      except
        ShowMessage('Unknown error encountered while inserting into DB.');   //somehting went wrong
        result:='none';
      end;
      ADOCmd.Free;                            // cleanup command object
      result:= GetSampleNrBySampleNameAndProjectNr(sample_name, project_nr);
    end
  else begin
        ShowMessage('Database not connected.');   //Database is not connected
        result:= 'none';
      end;

end;



function Tdm.AddNewUser: String;
// add a new user to the user_t (only last_name and user_nr)
// open a dialog that asks for the users last_name
// add the new user to the user table if user doesn't exist yet
// function returns the user_nr of the newly created user
// or 'none' if no user was created
Var
  last_name: string;
  choice: integer;
begin
 last_name := InputBox('Add New User', 'Users Last Name (Case Sensitive)', '');  // display InputBox (Case sensitive!!)
  // check if a valid string was given
 IF last_name <> '' THEN
  BEGIN
    //check if the name already exists
    IF CheckExistingDBValue('user_t', 'last_name', last_name) THEN
      BEGIN
        choice := MessageDlg('Last Name already exists! Create new user anyway?', mtWarning, [mbNo, mbYes],0); // last_name already exists
        IF choice = mrYes THEN   // if new user should be created any even if exists already
          BEGIN
          IF InsertIntoDB('user_t', 'last_name', last_name) THEN
            BEGIN
              // some error orrured while inserting new data
              ShowMessage('Error inserting new user.');
              Result:='none'
            END
            ELSE
            BEGIN
              // data inserted without any error
              ShowMessage('New user created.');  // show message
              Result:= dm.GetUserNrByLastName(last_name);// return userID of the newly created user
            END;
        END;
      END
      ELSE  // user doesn't exists yet -> insert into DB
      BEGIN
        IF InsertIntoDB('user_t', 'last_name', last_name) THEN
          BEGIN
            // some error orrured while inserting new data
            ShowMessage('Error inserting new user.');
            Result:='none'
          END
          ELSE
          BEGIN
            // data inserted without any error
            ShowMessage('New user created.');  // show message
            Result:= dm.GetUserNrByLastName(last_name);// return userID of the newly created user
          END;
      END;
   END;
end;


function Tdm.CheckExistingDBValue(table, column, value: string): boolean;
// check "table" if the "value" exists in the specified "column"
// returns true = "value exists" or false = "value doesn't exist"
Var
  ADOCmd: TADOCommand;
  ADODataSet: TADODataSet;
begin
if adoConnKTL.Connected then
begin            // only perform command if connection is established (use ADOCommand because it can also handle queries that don't return datasets)
      ADOCmd:= TADOCommand.Create(nil);       // create new command object
      ADODataSet:= TADODataSet.Create(nil);   // create dataset object - this time we get a result from the command so we need a dataset
      try
        ADOCmd.Connection:=adoConnKTL;        // set DB connection
        ADOCmd.Parameters.Clear;              // clear all Parameters
        ADOCmd.CommandType:=cmdtext;          // set command type to text
        ADOCmd.CommandText:='SELECT DISTINCT ' + column + ' FROM ' + table + ' WHERE ' + column + ' = ' + #34 + value + #34 + ';';    // set query text
        //ShowMessage(ADOCmd.CommandText);
        ADODataSet.Recordset:=ADOCmd.Execute; // issue command query and assign result to the dataset (use this onl if there is a result set returned)
      except
        ShowMessage('Unknown error encountered while checking DB for existing value.');   //somehting went wrong
        result:=true;                         // return true (value exists) which is a saver answer than "does not exist"
      end;

      //check if result of the DB query is the same as value
      if ADODataSet.FieldByName(column).Text=value then
        result:=true                          // value exists already
      else
        result:=false;                        // value doesn't exist yet

      ADOCmd.Free;                            // cleanup command object
      ADODataSet.Free;                        // cleanup dataset onject
    end
else
  result:= true;
end;

function Tdm.CheckProjectOfUser_nr(project_name: string; user_nr: integer): boolean;
// checks if the project_name exists within the selected user_nr
var
  s : string;
begin
   with qryDb do
   begin
     Close;
     SQL.Text := 'SELECT project FROM project_t INNER JOIN user_t ON user_t.user_nr=project_t.user_nr WHERE project=' + #34 + project_name + #34 + ' AND project_t.user_nr=' + IntToStr(user_nr) + ';';
     s := SQL.Text;
     //ShowMessage(s);
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     if Fields.Fields[0].AsString=project_name then
     begin
        Result:=true;
        LogWindow.addLogEntry('DB -- Projects exists already');
     end
     else
     begin
       Result:=false;
       LogWindow.addLogEntry('DB -- Projects does not yet exist.');
     end;
     //Result := Fields.Fields[0].AsString;
   end;
end;

function Tdm.CheckSampleOfProjectNr(sample_name, project_nr: string): boolean;
// checks if the sample_name (user_label) exists within the selected project_nr
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT user_label FROM sample_t WHERE user_label=' + #34 + sample_name + #34 + ' AND project_nr=' + #34 + project_nr + #34 +';';
     s := SQL.Text;
     //ShowMessage(s);
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     if Fields.Fields[0].AsString=sample_name then begin
        Result:=true;
     end
     else begin
       Result:=false;
     end;
     //Result := Fields.Fields[0].AsString;
   end;
end;


function Tdm.InsertIntoDB(table, column, value: string): boolean;
// insert "value" of "column" into "table"
// returns false (no error) if everything worked out
Var
  ADOCmd: TADOCommand;
begin
result:= false;
if adoConnKTL.Connected then
begin            // only perform command if connection is established (use ADOCommand because it can also handle queries that don't return datasets)
      ADOCmd:= TADOCommand.Create(nil);       // create new command object
      try
        ADOCmd.Connection:=adoConnKTL;        // set DB connection
        ADOCmd.Parameters.Clear;              // clear all Parameters
        ADOCmd.CommandType:=cmdtext;          // set command type to text
        ADOCmd.CommandText:='INSERT INTO ' + table + ' ( ' + column + ' ) VALUES ( '  + #34 + value + #34 + ');';    // set query text
        //ShowMessage(ADOCmd.CommandText);
        //ADOCmd.Parameters.ParamByName('param1').Value := last_name;      //insert parameter value into query
        ADOCmd.Execute; // issue command (no result set must be returned)
      except
        ShowMessage('Unknown error encountered while inserting into DB.');   //somehting went wrong
        result:=true;
      end;
      ADOCmd.Free;                            // cleanup command object
      result:= false;
    end
else
      begin
        ShowMessage('Database not connected.');   //Database is not connected
        result:= true;
      end;
end;



procedure Tdm.CheckProjectStatus;
var
  project_nr : integer;
begin
  with qryDb do
  begin
    SQL.Text := 'SELECT project_nr, project, status FROM project_t WHERE status<>"closed";';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    First;
    while not EOF do
    begin
       project_nr := Fields.FieldByName('project_nr').AsInteger;
       frmMAMS.lblProject.Caption :=  Fields.Fields[1].AsString;
       with qryDb1 do
       begin  // suche offene samples
          SQL.Text :=
            'SELECT sample_t.sample_nr FROM sample_t ' +
            'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
            'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
            'WHERE project_nr=' + IntToStr(project_nr) +
            ' AND preparation_t.stop=0 AND target_t.stop=0 AND sample_t.c14_age IS NULL ' +
            ';';
          LogWindow.addLogEntry(SQL.Text);
              IF dm.adoConnKTL.Connected THEN
              Begin
                Try
                  Open;
                  LogWindow.addLogEntry('DB -- query executed...');
                Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                End;
              End;
          if qryDb1.RecordCount=0 then
          begin   // alles datiert oder nottobedated
             with adoCmd do
             begin
               CommandText := 'UPDATE project_t SET status="closed" WHERE project_nr=' +
                            IntToStr(project_nr) + ';';
               LogWindow.addLogEntry(CommandText);
               IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
             end;
          end;
       end;
       Next;
    end;
  end;
end;

procedure Tdm.CreateBlankPrepRecord(SampleNr, PrepNr: integer);
//creates a blank prep record in the preparation_t for this sample_nr
begin
  adoCmd.CommandText :=
       'INSERT INTO preparation_t  (prep_nr, sample_nr) VALUES (' +
        IntTostr(PrepNr) + ',' +
        IntTostr(SampleNr) +
        ');';
  IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      adoCmd.Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
end;

procedure Tdm.CreateBlankTargetRecord(SampleNr, PrepNr, TargetNr: integer);
begin
   adoCmd.CommandText :=
         'INSERT INTO target_t  (target_nr, prep_nr, sample_nr) VALUES (' +
          IntTostr(TargetNr) + ',' +
          IntTostr(PrepNr) + ',' +
          IntTostr(SampleNr)  +
          ');';
   IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed...');;
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      adoCmd.Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
end;

procedure Tdm.DeleteFromDb(SampleNr: integer);
begin
   with adoCmd do begin
     CommandText := 'DELETE FROM target_t WHERE sample_nr=' + IntToStr(SampleNr);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
   with adoCmd do begin
     CommandText := 'DELETE FROM  preparation_t WHERE sample_nr=' + IntToStr(SampleNr);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
   with adoCmd do begin
     CommandText := 'DELETE FROM sample_t WHERE sample_nr=' + IntToStr(SampleNr);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
end;

function Tdm.GetAllInPrep : integer;
//get all samples that are in prep
var
  s : string;
begin
  with qryInPrep do
    begin
       Close;
//       SQL.Text :=  'SELECT sample_t.sample_nr, user_label, project_t.project, user_t.last_name ' +
//                    ' FROM sample_t ' +
//                    'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
//                    'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
//                    'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
//                    ' WHERE preparation_t.step1_start IS NOT NULL AND preparation_t.prep_end IS NULL;';
     SQL.Text :=  'SELECT sample_t.sample_nr, user_label, project_t.project, sample_t.material, user_t.last_name, ' +
                  'project_t.desired_date, (DATEDIFF(curdate(),preparation_t.prep_start)) AS Days, NOT ISNULL(weight_medium) AS InFreeze ' +
                  'FROM sample_t ' +
                  'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
                  'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
                  'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
                  'WHERE preparation_t.prep_start IS NOT NULL AND preparation_t.prep_end IS NULL AND preparation_t.stop=0;';
       s := SQL.Text;
       //ClipBoard.SetTextBuf(PChar(s));
       LogWindow.addLogEntry(SQL.Text);
       IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    end;
  Result := qryInPrep.RecordCount;
end;

function Tdm.GetAllPlanned(ShowOnHold : boolean; material: string = 'none') : integer;
// get all samples where prep has not started yet
// parameter "material" can be used to filter for specific materials, 'None' = no filter
var
  s : string;
begin
  with qryPlanned do
  begin
     Close;
     SQL.Text :=  'SELECT sample_t.sample_nr, user_label, project_t.project, sample_t.material, user_t.last_name,' +
                  ' project_t.desired_date ' +
                  ' FROM sample_t ' +
                  'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
                  'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
                  'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
                  ' WHERE sample_t.sample_nr>9999 AND preparation_t.step1_start IS NULL AND preparation_t.prep_end IS NULL ' +
                  ' AND c14_age IS NULL ' +
                  ' AND material <> "graphite" ' +
                  ' AND  preparation_t.stop=0' +
                  ' AND prep_start IS NULL';
      if ShowOnHold then
        SQL.Text := SQL.Text +  ' AND sample_t.not_tobedated=1'
      else
        SQL.Text := SQL.Text + ' AND sample_t.not_tobedated=0';
      // add a filter for the material if the filter string is not None
      if (material = 'none') then
        SQL.Text := SQL.Text
      else
        SQL.Text := SQL.Text +  ' AND material = ' + #34 + material + #34 + ';';

      s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));

     LogWindow.addLogEntry('getting samples planned for material: ' + material);
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    Open;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  Result := qryPlanned.RecordCount;
end;

function Tdm.GetAllWaitingForGraph : integer;
// get all samples that are not graphitized yet but prep'd
var
  s : string;
begin
  with qryWaitingForGraph do
  begin
     Close;
     SQL.Text :=   'SELECT DISTINCT sample_t.sample_nr, user_label, project_t.project, sample_t.material, user_t.last_name, user_t.first_name, project_t.desired_date ' +
                   ' FROM sample_t ' +
                   'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
                   'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
                   'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
                   'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
                   ' WHERE preparation_t.prep_end IS NOT NULL and target_t.graphitized IS NULL ' +
                   ' and target_t.target_pressed IS NULL and target_t.calcset is NULL and sample_t.not_tobedated=0' +
                   ' and preparation_t.stop=0 and  target_t.stop=0 ' +
                   ' and project_t.out_date IS NULL ' +
    //               ' and sample_t.type NOT LIKE ' + #34 + 'blank%' + #34 +
                   ' and sample_t.type NOT LIKE ' + #34 + 'oxa%' + #34 +
                   ' and NOT (user_t.last_name = "intern" AND user_t.first_name ="intern")' +
                   ' ORDER BY sample_t.sample_nr;' ;
     s := SQL.Text;
     LogWindow.addLogEntry('getting samples WaitingForGraph');
     LogWindow.addLogEntry(SQL.Text);
     // ClipBoard.SetTextBuf(PChar(s));
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    Open;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  Result := qryWaitingForGraph.RecordCount;
end;

function Tdm.GetAllWaitingForMeas : integer;
// get all samples that are graphitized but not measured
// but exlude internal samples and standards
var
  s : string;
begin
  with qryWaitingForMeas do
  begin
     Close;
//     SQL.Text :=   'SELECT DISTINCT sample_t.sample_nr, user_label, project_t.project, user_t.last_name, project_t.desired_date ' +
//                   ' FROM sample_t ' +
//                   'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
//                   'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
//                   'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
//                   'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
//                   ' WHERE preparation_t.prep_end IS NOT NULL and target_t.graphitized IS NOT NULL ' +
//                   ' and target_t.target_pressed IS NOT NULL and target_t.calcset is NULL and sample_t.not_tobedated=0' +
//                   ' and preparation_t.stop=0 and  target_t.stop=0 ' +
//                   ' and project_t.out_date IS NULL ' +
//                   ' and target_t.fm is NULL' +
//                   ' and sample_t.type NOT LIKE ' + #34 + 'blank%' + #34 +
//                   ' and sample_t.type NOT LIKE ' + #34 + 'oxa%' + #34 +
//                   ' and user_label NOT LIKE ' + #34 + 'HEI_%' + #34 +
//                   ' ORDER BY sample_t.sample_nr;' ;

     SQL.Text := 'SELECT target_id, type, user_label, user_label_nr, last_name, target_comment, co2_final, desired_date, datediff(curdate(), desired_date) as "days to deadline" ' +
     'FROM target_p ' +
     'WHERE target_p.sample_nr >= 10000 AND graphitized IS NOT NULL AND magazine IS NULL AND target_stop = 0 AND not_tobedated = 0 ' +
     'AND project_nr NOT IN ("798", "129", "767", "795", "1032", "1083", "1174","1658","1759","2295") ' +
     'AND type not in ("oxa2", "blank", "C1", "C2", "C3", "C6", "C7", "C8") ' ;

     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    Open;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  Result := qryWaitingForMeas.RecordCount;
end;

function Tdm.GetAllWaitingForMeasAll : integer;
// get all samples that are graphitized but not measured
// inlcusing internal samples and standards
//but not ICOS
var
  s : string;
begin
  with qryWaitingForMeasAll do
  begin
     //Close;

     SQL.Text := 'SELECT target_id, type, user_label, user_label_nr, last_name, target_comment, co2_final, desired_date, datediff(curdate(), desired_date) as "days to deadline" ' +
     'FROM target_p ' +
     'WHERE target_p.sample_nr >= 10000 AND graphitized IS NOT NULL AND magazine IS NULL AND target_stop = 0 AND not_tobedated = 0 ' +
     'AND project_nr not in ("798")' ;

     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    Open;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  Result := qryWaitingForMeasAll.RecordCount;
end;

function Tdm.GetWaitingExpress : integer;
// get all express samples not measured
var
  s : string;
begin
  with qryWaitingExpress do
  begin
     Close;
     SQL.Text :=   'SELECT DISTINCT sample_t.sample_nr, user_t.last_name, user_label, target_t.graphitized, project_t.desired_date, target_t.magazine ' +
                   ' FROM sample_t ' +
                   'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
                   'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
                   'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
                   'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
                   ' WHERE sample_t.user_label LIKE "%eil%" ' +
                   ' and target_t.calcset is NULL and sample_t.not_tobedated=0' +
                   ' and preparation_t.stop=0 and  target_t.stop=0 ' +
                   ' and (project_t.out_date < "1900-01-01" or project_t.out_date IS NULL) ' +
                   ' ORDER BY sample_t.sample_nr;' ;
     s := SQL.Text;
     LogWindow.addLogEntry(s);
//     ClipBoard.SetTextBuf(PChar(s));
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    Open;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  Result := qryWaitingExpress.RecordCount;
end;

function Tdm.GetArchSamplesOfMonth(Year, Month: word): integer;
// get number of archaeologic samples per month
var
  MoStart, MoEnd : TDateTime;
  s1, s2 : string;
begin
  MoStart := EncodeDate(Year, Month,1);
  MoEnd := EndOfAMonth(Year, Month);
  s1 := FormatDateTime('YYYY-MM-DD', DateOf(MoStart));
  s2 := FormatDateTime('YYYY-MM-DD', DateOf(MoEnd));
  with qryDb do begin
    SQL.Text := 'SELECT DISTINCT sample_t.sample_nr FROM sample_t ' +
                ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr ' +
                ' WHERE sample_t.c14_age IS NOT NULL ' +
                ' AND (target_pressed BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34  +
                ' OR graph_date BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34 +
                ' OR graphitized BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34 + ')' +
                ' AND sample_t.type="arch" ';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed...');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
//    ClipBoard.AsText := SQL.Text;
    Result := RecordCount;
  end;
end;

function Tdm.GetAllSamplesOfMonth(Year, Month: word; flag: integer): integer;
// returns the number of samples per month (not blanks or oxas)
// if flag = 1 -> graphitized samples
// if flag = 2 -> measured samples
var
  MoStart, MoEnd : TDateTime;
  s1, s2 : string;
begin
  MoStart := EncodeDate(Year, Month,1);
  MoEnd := EndOfAMonth(Year, Month);
  s1 := FormatDateTime('YYYY-MM-DD', DateOf(MoStart));
  s2 := FormatDateTime('YYYY-MM-DD', DateOf(MoEnd));

  // return number of received samples
  if flag = 0 then
  begin
  LogWindow.addLogEntry('getting number of received samples');
  // return number of received  samples
        qryDb.Close;
        qryDb.SQL.Text := 'select distinct sample_t.sample_nr from sample_t' +
                 ' INNER JOIN project_t ON sample_t.project_nr = project_t.project_nr' +
                 ' INNER JOIN user_t ON project_t.user_nr = user_t.user_nr' +
                 ' WHERE sample_t.type NOT IN ("oxa2", "oxa1", "blank")' +
                 ' AND sample_t.user_label NOT LIKE "%IAEA%"' +
                 ' AND year(in_date) = ' + #34 + inttostr(Year) + #34 +
                 ' AND month(in_date) = ' + #34 + inttostr(Month) + #34;
        LogWindow.addLogEntry(qryDb.SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                    Begin
                    LogWindow.addLogEntry('DB -- adoConnKTL is connected');
                      Try
                        qryDb.Open;
                        LogWindow.addLogEntry('DB -- query executed');
                        Result := qryDb.RecordCount;
                        LogWindow.addLogEntry('counted: ' + inttostr(Result));
                      Except
                        // connections is closed, reconnect by setting connected to true and test again
                        LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                          Try
                            dm.DBreconnect;
                            LogWindow.addLogEntry('DB -- execute query again');
                            qryDb.Open;
                          Except
                             LogWindow.addLogEntry('DB -- can not reconnect');
                          End;
                        //ShowMessage('Database connection is closed.');
                      End;
                    End
            else
                    begin
                     LogWindow.addLogEntry('DB -- adoConnKTL is not connected');
                    end;

  end;

  // return number of graphitized samples
  if flag = 1 then
  begin
  LogWindow.addLogEntry('getting number of graphitized samples');
  // return number of graphitized samples
        qryDb.Close;
        qryDb.SQL.Text := 'SELECT DISTINCT sample_t.sample_nr FROM sample_t ' +
                    ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr ' +
                    ' WHERE target_t.graphitized IS NOT NULL ' +
                    ' AND (target_pressed BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34  +
                    ' OR graph_date BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34 +
                    ' OR graphitized BETWEEN ' + #34 + s1 + #34 + ' AND ' + #34 + s2 + #34 + ')' +
                    ' AND sample_t.type NOT IN("blank","oxa1","oxa","oxa2") ';
        LogWindow.addLogEntry(qryDb.SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                    Begin
                    LogWindow.addLogEntry('DB -- adoConnKTL is connected');
                      Try
                        qryDb.Open;
                        LogWindow.addLogEntry('DB -- query executed');
                        Result := qryDb.RecordCount;
                        LogWindow.addLogEntry('counted: ' + inttostr(Result));
                      Except
                        // connections is closed, reconnect by setting connected to true and test again
                        LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                        Try
                          dm.DBreconnect;
                          LogWindow.addLogEntry('DB -- execute query again');
                          qryDb.Open;
                        Except
                           LogWindow.addLogEntry('DB -- can not reconnect');
                        End;
                        //ShowMessage('Database connection is closed.');
                      End;
                    End
            else
                    begin
                     LogWindow.addLogEntry('DB -- adoConnKTL is not connected');
                    end;

  end;

  // return number of measured samples
  if flag = 2 then
      begin
      LogWindow.addLogEntry('getting number of measured samples');
      // return number of measured samples
          qryDb.Close;
            qryDb.SQL.Text := 'SELECT DISTINCT CAST(SUBSTRING(magazine, 3, 2) AS UNSIGNED), sample_t.sample_nr FROM sample_t ' +
                        ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr' +
                        ' WHERE sample_t.c14_age IS NOT NULL' +
                        ' AND magazine LIKE "MA%"' +
                        ' AND CAST(SUBSTRING(magazine, 3, 2) AS UNSIGNED) = ' + #34 + Copy(inttostr(Year),3,2) + #34 +
                        ' AND CAST(SUBSTRING(magazine, 5, 2) AS UNSIGNED) = ' + #34 + inttostr(Month) + #34 +
                        ' AND sample_t.type NOT IN("blank","oxa1","oxa","oxa2") ';
            LogWindow.addLogEntry(qryDb.SQL.Text);
            IF dm.adoConnKTL.Connected THEN
                        Begin
                        LogWindow.addLogEntry('DB -- adoConnKTL is connected');
                          Try
                            qryDb.Open;
                            LogWindow.addLogEntry('DB -- query executed');
                            Result := qryDb.RecordCount;
                            LogWindow.addLogEntry('counted: ' + inttostr(Result));
                          Except
                            // connections is closed, reconnect by setting connected to true and test again
                            LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                            Try
                              dm.DBreconnect;
                              LogWindow.addLogEntry('DB -- execute query again');
                              qryDb.Open;
                            Except
                               LogWindow.addLogEntry('DB -- can not reconnect');
                            End;
                            //ShowMessage('Database connection is closed.');
                          End;
                        End
                else
                        begin
                         LogWindow.addLogEntry('DB -- adoConnKTL is not connected');
                        end;

      end;

end;

procedure Tdm.GetGraphWeight(SampleNr, PrepNr, TargetNr: integer);
begin
   with qryGraphWeight do begin
     SQL.Text :=  'SELECT weight_combustion FROM target_t WHERE sample_nr=' +
                  IntToStr(SampleNr) + ' AND prep_nr=' + IntToStr(PrepNr) +
                  ' AND target_nr=' + IntToStr(TargetNr)+';';
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
end;

procedure Tdm.GetMagazineData(Magazine : string);
// read the target data from the magazine
// used for transfering the target data into the sample data
var
  s : string;
begin
 with qryMagazineData do begin
    SQL.Text := 'SELECT position, project_t.project, user_t.last_name, sample_t.sample_nr, ' +
                ' target_t.prep_nr, target_t.target_nr, target_t.c14_age, target_t.calcset, LEFT(calc_set_t.first_run,2) AS RunIDPrefix  ' +
                ' FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr' +
                ' INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr' +
                ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr' +
                ' INNER JOIN calc_set_t ON target_t.calcset=calc_set_t.calcset' +
                ' WHERE target_t.magazine='+ #34 + Magazine + #34 + ' AND target_t.fm IS NOT NULL;';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
   LogWindow.addLogEntry(SQL.Text);
   IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
 end;
end;

procedure Tdm.GetMagazineDataUnpressed(Magazine : string);
// read the target data from the magazine with the unpressed samples
// meant for setting the pressed date of a whole magazine
var
  s : string;
begin
 with qrySamplesOfUnpressedMagazine do begin
    SQL.Text := 'SELECT position, concat(target_t.sample_nr, ".", target_t.prep_nr, ".", target_t.target_nr) as targetID, type, user_label, user_t.last_name, co2_final ' +
                ' FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr' +
                ' INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr' +
                ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr' +
                ' WHERE target_t.magazine='+ #34 + Magazine + #34 +
                ' ORDER BY position;';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
   LogWindow.addLogEntry(SQL.Text);
   IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
 end;
end;

function Tdm.GetMaxPrepNrBySampleNr(sample_nr: integer): integer;
var
  s : string;
begin
   with qryDb do
   begin
     Close;
     SQL.Text := 'SELECT Max(prep_nr) FROM preparation_t WHERE sample_nr='+IntToStr(sample_nr) + ';';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsInteger;
     Close;
   end;
end;

function Tdm.GetMaxSampleNr: integer;
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT Max(sample_nr) FROM sample_t;';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsInteger;
   end;
end;

function Tdm.GetMaxTargetNrBySampleNr(sample_nr, prep_nr: integer): integer;
var
  s : string;
begin
   with qryDb do
   begin
     Close;
     SQL.Text := 'SELECT Max(target_nr) FROM target_t WHERE sample_nr='+IntToStr(sample_nr) +
                 ' AND prep_nr=' + IntToStr(prep_nr)+ ';';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsInteger;
     Close;
   end;
end;

procedure Tdm.GetNOxaBlank(Oxa: boolean);
// get oxas or blanks that are graphiyized but not yet in a magazine
// those are available for analysis
var
  s, s2 : string;
begin
  if Oxa then begin
    s := '"oxa%"';
    s2:= '"oxa%"';
  end
  else begin
    s :=  '"blank%"';
    s2:=  '"Phthalic%"';
  end;

  with qryDB do begin
    SQL.Text := ' SELECT target_t.sample_nr, sample_t.user_label FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr ' +
                ' INNER JOIN project_t ON sample_t.project_nr=project_t.project_nr ' +
                ' INNER JOIN user_t ON project_t.user_nr=user_t.user_nr ' +
                ' WHERE magazine IS NULL AND graphitized IS NOT NULL and stop=0 ' +
                ' and sample_t.type like ' + s +' and sample_t.user_label like ' + s2 +
                ' AND user_t.user_nr = "129"' + ';';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

procedure Tdm.GetNIAEA(stdname: string);
// get the number of available IAEA standards defined by "stdname"
begin
  with qryDB do begin
    SQL.Text := ' SELECT target_t.sample_nr, sample_t.user_label FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr ' +
                ' INNER JOIN project_t ON sample_t.project_nr=project_t.project_nr ' +
                ' INNER JOIN user_t ON project_t.user_nr=user_t.user_nr ' +
                ' WHERE magazine IS NULL AND graphitized IS NOT NULL and stop=0 ' +
                ' and sample_t.type like ' + #34 + stdname + #34 + ';';
    //ShowMessage(SQL.Text);
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

procedure Tdm.GetNPferde();
// get the number of available Pferde standards defined by its project number
begin
  with qryDB do begin
    SQL.Text := ' SELECT target_t.sample_nr, sample_t.user_label FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr ' +
                ' INNER JOIN project_t ON sample_t.project_nr=project_t.project_nr ' +
                ' INNER JOIN user_t ON project_t.user_nr=user_t.user_nr ' +
                ' WHERE magazine IS NULL AND graphitized IS NOT NULL and stop=0 ' +
                ' and project_t.project_nr='+ '748' +';';
    //ShowMessage(SQL.Text);
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

function Tdm.GetPretreatStepNrBySampleNr(sample_nr, prep_nr: integer; const method : string): integer;
var
  StepFound : boolean;
  StepNr : integer;
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT step1_method, step2_method,step3_method,step4_method,' +
                 'step5_method FROM preparation_t WHERE sample_nr=' +
                 IntTostr(sample_nr) +
                 ' AND prep_nr=' + IntTostr(prep_nr) + ';';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     First;
     StepNr := 0;
     StepFound := false;
     while not EOF and not StepFound do begin
       Inc(StepNr);
       if Pos(method,Fields.Fields[0].AsString)>0 then StepFound := true;
       Next;
     end;
     if not StepFound then StepNr := 0;
   end;
   Result := StepNr;
end;

function Tdm.GetProjectNrBySampleNr(sample_nr: integer): integer;
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT project_nr FROM sample_t WHERE sample_nr='+IntToStr(sample_nr) + ';';
     s := SQL.Text;
//     ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsInteger;
   end;
end;

function Tdm.GetProjectByProjectNr(project_nr: integer): string;
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT project FROM project_t WHERE project_nr='+IntToStr(project_nr) + ';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsString;
   end;
end;

function Tdm.GetSamplesByProjectNr(project_nr: integer): integer;
//returns number of samples by project number
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT count(sample_nr) FROM sample_t WHERE project_nr='+IntToStr(project_nr) + ';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsInteger;
   end;
end;

function Tdm.RemoveProjectByProjectNr(project_nr: integer): integer;
//removes the project from the database
// return 0 of no error
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'DELETE FROM project_t WHERE project_nr='+IntToStr(project_nr) + ';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     Result := 1;
     IF dm.adoConnKTL.Connected THEN
                Begin
                Result := 0;
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                    Result := 1;
                  End;
                End;
   end;
end;

function Tdm.GetProjectNrByProjectAndUserNr(project_name, user_nr: string): string;
// returns the project_nr of the project_name
// user_nr needs to be give because project_name is not unique
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT project_nr FROM project_t WHERE project='+ #34 + project_name + #34 + ' AND user_nr=' + #34 + user_nr + #34 +';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsString;
   end;
end;

function Tdm.GetSampleNrBySampleNameAndProjectNr(sample_name, project_nr: string): string;
// returns the sample_nr of the sample (user_label)
// project_nr needs to be give because sample_name (user_label) is not unique
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT sample_nr FROM sample_t WHERE user_label='+ #34 + sample_name + #34 + ' AND project_nr=' + #34 + project_nr + #34 +';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsString;
   end;
end;

function Tdm.GetUserByProjectNr(project_nr: integer): string;
var
  s : string;
begin
   with qryDb do begin
     Close;
     SQL.Text := 'SELECT user_t.last_name FROM project_t INNER JOIN user_t ON user_t.user_nr=project_t.user_nr WHERE project_nr='+IntToStr(project_nr) + ';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsString;
   end;
end;

function Tdm.GetUserNrByLastName(last_name: string): String;
var
  s : string;
begin
  with qryDb do begin
     Close;
     SQL.Text := 'SELECT user_nr FROM user_t WHERE last_name=' + #34 + last_name + #34 +';';
     s := SQL.Text;
     //ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
     Result := Fields.Fields[0].AsString;
   end;
end;

procedure Tdm.GetProjectsSinceYear(StartYear: integer);
var
  s : string;
begin
   with qryProjectsSinceYear do begin
     Close;
     SQL.Text := 'SELECT project, in_date, last_name, town ' +
                 ' FROM sample_t ' +
                 ' INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
                 ' INNER JOIN user_t ON project_t.user_nr=user_t.user_nr ' +
                 ' WHERE in_date > ' + #34+IntToStr(StartYear)+'-01-01' +#34 +
                 ' AND user_t.last_name<>"intern" ' +  //user intern
                 ';';
     s := SQL.Text;
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
end;

procedure Tdm.GetSampleInfo(SampleNr, PrepNr, TargetNr: integer);
// get all the sample info from the database, from various tables
var
  s : string;
  HasTarget : boolean;
begin
  HasTarget := TargetRecordExistsForSample(SampleNr);
  with qrySampleInfo do
  begin
    Close;
    s :=
        ' SELECT sample_t.sample_nr, project_t.project_comment, sample_t.lab_comment, freeofcharge, s_no_leftover, s_storage_loc, prep_storage_loc, type, material, fraction, pre_sub_treat, sample_t.weight, preparation, sampling_date,' +
        ' editable, not_tobedated, user_label, sample_t.user_label_nr, user_desc1, user_desc2, residue,' +
        ' sample_t.user_comment,sample_t.old_info, project_t.project, project_t.project_nr, report, invoice_nr, AuftragsNr, MA_Nr, in_date, desired_date, out_date, priority, status,' +
        ' price, user_t.last_name, user_t.user_nr, preparation_t.prep_nr, preparation_t.batch, p_no_leftover,' +
        ' prep_comment, weight_start, weight_medium, weight_end, weight_medium_2, ' +
        ' prep_start, prep_end, step1_method, step2_method,step3_method,step4_method,step5_method, preparation_t.old_info ';
    s := s + ', target_t.target_nr,' +
        ' conc_c/conc_n*14/12 as cn_ratio, conc_c, conc_n, preparation_t.stop,' +
        ' magazine, position, precis, cycle_min, cycle_max, catalyst, cathode_nr, reactor_nr,' +
        ' co2_init, co2_final, hydro_init, hydro_final, react_time, target_pressed, target_t.stop,' +
        ' le_curr, he_curr, fm, fm_sig, dc13, dc13_sig,calcset, editallowed, target_t.c14_age, target_t.c14_age_sig,' +
        ' target_t.cal1sMin, target_t.cal1sMax, target_t.cal2sMin, target_t.cal2sMax, target_t.graph_date, ' +
        ' target_t.target_comment, target_t.graph_batch, target_t.graphitized, ' +
        ' return_to_sender, returned_to_sender, prep_return_to_sender, prep_returned_to_sender, ' +
        ' CNIsotopA, CNIsotopAMoved ';
    s := s + ' FROM sample_t ' +
         'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
         'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
         'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ';
    s := s + 'INNER JOIN (select * from target_t WHERE sample_nr='+IntToStr(SampleNr)+' AND prep_nr='+ IntToStr(PrepNr) +' AND target_nr='+ IntToStr(TargetNr) +') target_t ON target_t.sample_nr=sample_t.sample_nr ';
    s := s + 'WHERE sample_t.sample_nr=' + IntToStr(SampleNr) + ' AND preparation_t.prep_nr=' + IntToStr(PrepNr) + ' ';

    if HasTarget then s := s + ' AND target_nr=' + IntToStr(TargetNr);
    s := s + ';';
    //StrToClipBoard(s,300);
    SQL.Text := s;
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
        begin
          try
            Open;
            LogWindow.addLogEntry('DB -- query executed');
          except
            // connections is closed, reconnect by setting connected to true and test again
            LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
            Open;
            // dm.adoConnKTL.Open();
            //ShowMessage('Database connection is closed.');
          end;
        end;
  end;
end;

procedure Tdm.DBreconnect;
// reconnect to the database
begin
  dm.adoConnKTL.Close;
  dm.adoConnKTL.Open;
end;

procedure Tdm.GetSamplesAvailableForPrep(const method: string);
// get all samples that are not prep'd yet
var
  s, SampleNrList: string;
begin
      with qryDB do
      begin
        Close;
        s := 'SELECT sample_nr FROM preparation_t WHERE ' +
          ' (step1_method=' + #34 + method + #34 + ' AND step1_start IS NULL) OR  ' +
          ' (step2_method=' + #34 + method + #34 + ' AND step2_start IS NULL) OR  ' +
          ' (step3_method=' + #34 + method + #34 + ' AND step3_start IS NULL) OR  ' +
          ' (step4_method=' + #34 + method + #34 + ' AND step4_start IS NULL) OR  ' +
          ' (step5_method=' + #34 + method + #34 + ' AND step5_start IS NULL) AND ' +
          ' batch IS NULL;';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;
    if qryDb.RecordCount > 0 then
    begin
      with qryDb do
      begin
        SampleNrList := '(';
        First;
        while not EOF do
        begin
          SampleNrList := SampleNrList + qryDb.Fields.Fields[0].AsString + ',';
          Next;
        end;
        SampleNrList := copy(SampleNrList, 1, length(SampleNrList) - 1); // trim final comma
        SampleNrList := SampleNrList + ')';
      end;
      with qrySamplesAvailable do
      begin
        s := 'SELECT sample_t.sample_nr, project_t.project, user_label, user_label_nr, sample_t.weight,' +
             ' preparation_t.prep_nr, user_t.last_name ' +
          'FROM sample_t ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' + // to get prep_nr
          'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
          'WHERE sample_t.sample_nr IN ' + SampleNrList +
          ' AND preparation_t.batch IS NULL AND sample_t.not_tobedated=0 ' +
          ';';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;
    end;
end;

procedure Tdm.GetSamplesAvailableForPrepByMaterial(const material: string);
// get all samples that are not prep'd yet
var
  s, SampleNrList: string;
begin
      // get all sample_nr that have the correct sample_type
      with qryDB do
      begin
        Close;
        s := 'SELECT sample_t.sample_nr FROM sample_t ' +
          ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
          ' WHERE material =' + #34 + material + #34 + ' AND' +
          ' batch IS NULL AND prep_end IS NULL;';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                   on E: Exception do
                      begin
                       LogWindow.addLogEntry('DB -- problems with DB: ' + E.Message);
                       showmessage('Problem with DB: '+E.Message);
                      // connections is closed, reconnect by setting connected to true and test again
                      Try
                        dm.DBreconnect;
                        LogWindow.addLogEntry('DB -- execute query again');
                        Open;
                      Except
                         LogWindow.addLogEntry('DB -- can not reconnect');
                      End;
                    end;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;

    if qryDb.RecordCount > 0 then
    begin
      with qryDb do // using the query above create a list with sample numbers
      begin
        SampleNrList := '(';
        First;
        while not EOF do
        begin
          SampleNrList := SampleNrList + qryDb.Fields.Fields[0].AsString + ',';
          Next;
        end;
        SampleNrList := copy(SampleNrList, 1, length(SampleNrList) - 1); // trim final comma
        SampleNrList := SampleNrList + ')';
      end;
      // get all sample information from all the samples in the list
      LogWindow.addLogEntry('DB -- generated SampleNrList: ' + SampleNrList);
      with qrySamplesAvailable do
      begin
        s := 'SELECT sample_t.sample_nr, project_t.project, user_label, user_label_nr, sample_t.weight,' +
             ' preparation_t.prep_nr, user_t.last_name, sample_t.material, ' +
             'preparation_t.step1_method, preparation_t.step2_method, preparation_t.step3_method, preparation_t.step4_method ' +
          'FROM sample_t ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' + // to get prep_nr
          'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
          'WHERE sample_t.sample_nr IN ' + SampleNrList +
          ' AND preparation_t.batch IS NULL AND sample_t.not_tobedated=0 ' +
          ';';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                   on E: Exception do
                   begin
                       LogWindow.addLogEntry('DB -- problems with DB: ' + E.Message);
                       showmessage('Problem with DB: '+E.Message);
                      // connections is closed, reconnect by setting connected to true and test again
                      Try
                        dm.DBreconnect;
                        LogWindow.addLogEntry('DB -- execute query again');
                        Open;
                      Except
                         LogWindow.addLogEntry('DB -- can not reconnect');
                      End;
                    end;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;
    end;
end;

procedure Tdm.GetSamplesAvailableForPrepByMethod(const method: string);
// get all samples that are not prep'd yet
var
  s, SampleNrList: string;
begin
      // get all sample_nr that have the correct sample_type
      with qryDB do
      begin
        Close;
        s := 'SELECT sample_t.sample_nr FROM sample_t' +
          ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr' +
          ' INNER JOIN project_t ON sample_t.project_nr=project_t.project_nr' +
          ' WHERE step1_method =' + #34 + method + #34 +
          ' AND batch IS NULL' +
          ' AND prep_end IS NULL' +
          ' AND av_fm IS NULL' +
          ' AND project_t.out_date IS NULL;';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    on E: Exception do
                    begin
                       LogWindow.addLogEntry('DB -- problems with DB: ' + E.Message);
                       showmessage('Problem with DB: '+E.Message);
                      // connections is closed, reconnect by setting connected to true and test again
                      Try
                        dm.DBreconnect;
                        LogWindow.addLogEntry('DB -- execute query again');
                        Open;
                      Except
                         LogWindow.addLogEntry('DB -- can not reconnect');
                      End;
                    end;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;

    if qryDb.RecordCount > 0 then
    begin
      with qryDb do // using the query above create a list with sample numbers
      begin
        SampleNrList := '(';
        First;
        while not EOF do
        begin
          SampleNrList := SampleNrList + qryDb.Fields.Fields[0].AsString + ',';
          Next;
        end;
        SampleNrList := copy(SampleNrList, 1, length(SampleNrList) - 1); // trim final comma
        SampleNrList := SampleNrList + ')';
      end;
      // get all sample information from all the samples in the list
      LogWindow.addLogEntry('DB -- generated SampleNrList: ' + SampleNrList);

      with qrySamplesAvailable do
      begin
        s := 'SELECT sample_t.sample_nr, project_t.project, user_label, user_label_nr, sample_t.weight,' +
             ' preparation_t.prep_nr, user_t.last_name, sample_t.material, ' +
             'preparation_t.step1_method, preparation_t.step2_method, preparation_t.step3_method, preparation_t.step4_method ' +
          'FROM sample_t ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' + // to get prep_nr
          'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
          'WHERE sample_t.sample_nr IN ' + SampleNrList +
          ' AND preparation_t.batch IS NULL AND sample_t.not_tobedated=0 ' +
          ';';
//        ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    on E: Exception do
                    begin
                       LogWindow.addLogEntry('DB -- problems with DB: ' + E.Message);
                       showmessage('Problem with DB: '+E.Message);
                      // connections is closed, reconnect by setting connected to true and test again
                      Try
                        dm.DBreconnect;
                        LogWindow.addLogEntry('DB -- execute query again');
                        Open;
                      Except
                         LogWindow.addLogEntry('DB -- can not reconnect');
                      End;
                    end;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;
    end;
end;

procedure Tdm.GetTargetsAvailable;
// get all samples that are prep'd but are not
// graphitized yet
var
  s : string;
begin
      with qrySamplesAvailable do
      begin
        Close;
        s := 'SELECT DISTINCT sample_t.sample_nr, project_t.project, user_label, user_label_nr, sample_t.weight,' +
             ' target_t.prep_nr, target_t.target_nr ' +
             ' FROM sample_t ' +
          'INNER JOIN project_t ON project_t.project_nr=sample_t.project_nr ' +
          'INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr ' +
          'INNER JOIN user_t ON user_t.user_nr=project_t.user_nr ' +
          'INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sample_nr ' +
          'WHERE ' +
          ' target_t.graphitized IS NULL AND target_t.graph_date IS NULL AND target_t.target_pressed IS NULL AND target_t.graph_batch IS NULL' +
//        ' AND preparation_t.prep_nr=1 AND preparation_t.prep_end IS NOT NULL AND target_t.c14_age IS NULL' +
          ' AND preparation_t.prep_end IS NOT NULL AND target_t.c14_age IS NULL' +
          ' AND sample_t.not_tobedated=0 AND target_t.stop=0 and preparation_t.stop=0 order by sample_nr;';
        //ClipBoard.SetTextBuf(PChar(s));
        SQL.Text := s;
        LogWindow.addLogEntry(SQL.Text);
        IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
      end;
end;

procedure Tdm.GetWeights(SampleNr, PrepNr: integer);
var
  s : string;
begin
   LogWindow.addLogEntry('Query prep weights');
   with qryWeights do begin
     SQL.Text :=  'SELECT weight_start, weight_end, weight_medium, weight_medium_2 FROM preparation_t WHERE sample_nr=' +
                  IntToStr(SampleNr) + ' AND prep_nr=' + IntToStr(PrepNr) +';';
    s := SQL.Text;
//    ClipBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
end;

function Tdm.QueryInvoiceUser(const LastName, Institute,
  Organisation: string): integer;
var
  s : string;
begin
  with qryDB do begin
    Close;
    if Pos('n.a.', Organisation) = 0 then
      SQL.Text := 'Select * FROM user_t WHERE ' +
        'institute=' + #34 + Institute + #34 +
        ' AND last_name=' + #34 + LastName + #34 +
        ' AND organisation=' + #34 + Organisation + #34
    else
      SQL.Text := 'Select * FROM user_t WHERE ' +
        'institute=' + #34 + Institute + #34 +
        ' AND last_name=' + #34 + LastName + #34;
    s := SQL.Text;
//    ClipBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    Result := RecordCount;
  end;
end;

procedure Tdm.QueryPrepStepsBySampleNr(sample_nr, prep_nr: integer);
begin
  with qryPrepSteps do begin
    Close;
    SQL.Text := 'SELECT prep_nr, step1_method, step2_method, step3_method, step4_method,' +
      ' step5_method, prep_end, old_info, prep_comment FROM preparation_t WHERE sample_nr=' + IntToStr(sample_nr) +
      ' AND prep_nr=' + IntToStr(prep_nr) + ';';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

procedure Tdm.QueryTargetInfoBySampleNr(sample_nr: Integer);
var
  s: string;
begin
  // query the target information
  with dm.qryTarget do
  begin
    SQL.Text := 'SELECT target_nr, magazine, position, catalyst, co2_init, co2_final, ' +
    'hydro_init, hydro_final, react_time, target_pressed, target_t.stop, round(conc_c,1) AS conc_c,' +
    'round(conc_c/conc_n*14/12,1) as cn_ratio, round(weight_end/weight_start*100,1) AS collpc, target_comment ' +
    'FROM target_t ' + 'INNER JOIN preparation_t ON preparation_t.sample_nr=target_t.sample_nr ' +
    'WHERE target_t.sample_nr=' + IntToStr(sample_nr) + ' AND target_t.prep_nr=1 and target_t.target_nr=1;';
    s := SQL.Text;
    //   ClibBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;


procedure Tdm.QueryProject(const ProjectName: string);
// returns all info about this project
begin
  with qryProject do begin
    SQL.Text := 'SELECT * FROM project_t WHERE project='+#34+ProjectName+#34+';';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

procedure Tdm.QuerySampleBySampleNr(sample_nr: integer);
begin
  with qrySample do begin
    SQL.Text := 'SELECT sample_nr, user_label, user_label_nr, user_desc1, user_desc2, user_comment, ' +
      'material, weight, residue, not_tobedated, lab_comment, ' +
      'round(c14_age,0) AS C14_age, round(c14_age_sig,0) AS C14_age_sig  ' +
      'FROM sample_t ' +
      'WHERE sample_t.sample_nr=' + IntToStr(sample_nr) + ';';
//    strClipBoard := SQL.Text;
//    ClipBoard.SetTextBuf(PChar(strClipboard));
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
end;

procedure Tdm.SetGraphDate;
var
  Y,Mo,D : word;
  Magazine : string;
  GraphDate : TDateTime;
  SampleNr : integer;
begin
  with qryDb do
  begin
    SQL.Text := 'SELECT magazine, sample_nr FROM target_t WHERE graph_date IS NULL AND c14_age IS NOT NULL';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    if RecordCount>0 then
    begin
      First;
      while Not EOF do
      begin
        Magazine := Fields.Fields[0].AsString;
        SampleNr := Fields.Fields[1].AsInteger;
        if Pos('MA',Magazine)>0 then
        begin
          Y := StrToInt(copy(Magazine,3,2)) + 2000;
          Mo := StrToInt(copy(Magazine,5,2));
          D := StrToInt(copy(Magazine,7,2));
          GraphDate := EncodeDate(Y,Mo,D);
          with adoCmd do
          begin
            CommandText := 'UPDATE target_t SET graph_date=' +
            #34 + FormatDateTime('YYYY-MM-DD', DateOf(GraphDate)) + #34 +
            ' WHERE sample_nr=' + IntToStr(SampleNr) +
            ';';
            IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
          end;
        end;
        Next;
      end;
    end;

  end;
end;

procedure Tdm.SetGraphitized;
begin
   with adoCmd do
   begin
      CommandText := 'UPDATE target_t SET graphitized=graph_date ' +
            ' WHERE graph_date is not null and graphitized is null and sample_nr>19000' +
            ';';
      IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
   end;
end;

procedure Tdm.SetProjectStatusRunning(sample_nr: integer);
begin
  with qryDB do
  begin
    SQL.Text := 'SELECT project_nr FROM sample_t WHERE sample_nr=' + IntToStr(sample_nr);
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    if RecordCount=1 then
    begin
       adoCmd.CommandText := 'UPDATE project_t SET status="running" WHERE project_nr=' +
                              qryDB.Fields.Fields[0].AsString;
       IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      adoCmd.Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    end;
  end;
end;

procedure Tdm.StrToClipBoard(const s: string; msDelay: integer);
CONST
   MaxRetries= 5;
VAR
  RetryCount: Integer;
begin
 for RetryCount:= 1 to MaxRetries DO
  TRY
    Clipboard.AsText:= s;
    Break;
  EXCEPT
    on Exception DO
      if RetryCount = MaxRetries then
         RAISE Exception.Create('Cannot set clipboard')
      else
         Sleep(msDelay*RetryCount)
  END;
end;

procedure Tdm.SwapResults(Nr1, Nr2: integer);
var
  s : string;
  c : char;
begin
  with qrySample1 do begin
    SQL.Text := 'SELECT magazine,position,fm,fm_sig,dc13,calcset,c14_age,' +
        'c14_age_sig,cal1sMin,cal1sMax,cal2sMin,cal2sMax FROM target_t' +
        ' WHERE sample_nr=' + IntToStr(Nr1) + ';';
    s := SQL.Text;
//  ClipBoard.SetTextBuf(PChar(s));
     LogWindow.addLogEntry(SQL.Text);
     IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  with qrySample2 do begin
    SQL.Text := 'SELECT magazine,position,fm,fm_sig,dc13,calcset,c14_age,' +
        'c14_age_sig,cal1sMin,cal1sMax,cal2sMin,cal2sMax FROM target_t' +
        ' WHERE sample_nr=' + IntToStr(Nr2) + ';';
      LogWindow.addLogEntry(SQL.Text);
      IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
    c := FormatSettings.DecimalSeparator;
    FormatSettings.DecimalSeparator := '.';
  s := 'UPDATE target_t SET ';
  s := s + 'magazine=' + #34 + qrySample1.FieldByName('magazine').AsString + #34 + ',';
  s := s + 'position=' + qrySample1.FieldByName('position').AsString+ ',';
  s := s + 'fm=' + qrySample1.FieldByName('fm').AsString+ ',';
  s := s + 'fm_sig=' + qrySample1.FieldByName('fm_sig').AsString+ ',';
  s := s + 'dc13=' + qrySample1.FieldByName('dc13').AsString+ ',';
  s := s + 'calcset=' + #34 + qrySample1.FieldByName('calcset').AsString+ #34 + ',';
  s := s + 'c14_age=' + qrySample1.FieldByName('c14_age').AsString+ ',';
  s := s + 'c14_age_sig=' + qrySample1.FieldByName('c14_age_sig').AsString+ ',';
  s := s + 'cal1sMin=' + qrySample1.FieldByName('cal1sMin').AsString+ ',';
  s := s + 'cal1sMax=' + qrySample1.FieldByName('cal1sMax').AsString+ ',';
  s := s + 'cal2sMin=' + qrySample1.FieldByName('cal2sMin').AsString+ ',';
  s := s + 'cal2sMax=' + qrySample1.FieldByName('cal2sMax').AsString;
  s := s + ' WHERE sample_nr=' + IntToStr(Nr2) + ';';
//  ClipBoard.SetTextBuf(PChar(s));
  adoCmd.CommandText := s;
  IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      adoCmd.Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  s := 'UPDATE target_t SET ';
  s := s + 'magazine=' + #34 + qrySample2.FieldByName('magazine').AsString + #34 + ',';
  s := s + 'position=' + qrySample2.FieldByName('position').AsString+ ',';
  s := s + 'fm=' + qrySample2.FieldByName('fm').AsString+ ',';
  s := s + 'fm_sig=' + qrySample2.FieldByName('fm_sig').AsString+ ',';
  s := s + 'dc13=' + qrySample2.FieldByName('dc13').AsString+ ',';
  s := s + 'calcset=' + #34 + qrySample1.FieldByName('calcset').AsString+ #34 + ',';
  s := s + 'c14_age=' + qrySample2.FieldByName('c14_age').AsString+ ',';
  s := s + 'c14_age_sig=' + qrySample2.FieldByName('c14_age_sig').AsString+ ',';
  s := s + 'cal1sMin=' + qrySample2.FieldByName('cal1sMin').AsString+ ',';
  s := s + 'cal1sMax=' + qrySample2.FieldByName('cal1sMax').AsString+ ',';
  s := s + 'cal2sMin=' + qrySample2.FieldByName('cal2sMin').AsString+ ',';
  s := s + 'cal2sMax=' + qrySample2.FieldByName('cal2sMax').AsString;
  s := s + ' WHERE sample_nr=' + IntToStr(Nr1) + ';';
//  ClipBoard.SetTextBuf(PChar(s));
  adoCmd.CommandText := s;
  IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    adoCmd.Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      adoCmd.Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  FormatSettings.DecimalSeparator := c;
  qrySample1.Close;
  qrySample2.Close;
end;

function Tdm.TargetRecordExistsForSample(SampleNr: integer): boolean;
var
  s : string;
begin
  with qryDB do begin
    Close;
    SQL.Text := 'Select * FROM target_t WHERE sample_nr=' + IntToStr(SampleNr) + ';';
    s := SQL.Text;
//    ClipBoard.SetTextBuf(PChar(s));
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
    Result := (RecordCount>0);
  end;
end;

procedure Tdm.tblMethoddescrGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
 // actually show the descr column as text in the dbgrid
 // if this is not done the dbgrid would only show "Memo" since the descr field is of type "text"
   Text := Copy(tblMethoddescr.AsString, 1, 50);
end;

procedure Tdm.TransferAgeFromTarget(SampleNr, PrepNr, TargetNr: integer);
var
  s : string;
  c : char;
  c14age, c14age_sig : string;
  fm, fm_sig, fC14age : extended;
begin
  with qryDB1 do
  begin
    SQL.Text := 'SELECT c14_age, c14_age_sig, fm, fm_sig, Cal1sMin, Cal1sMax,' +
                ' Cal2sMin, Cal2sMax, dc13, dc13_sig ' +
                ' FROM target_t WHERE sample_nr='+IntToStr(SampleNr) +
                ' AND prep_nr=' + IntToStr(PrepNr) +
                ' AND target_nr=' + IntToStr(TargetNr) +';';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;
  with adoCmd do
  begin
    c := FormatSettings.DecimalSeparator;
    FormatSettings.DecimalSeparator := '.';
    c14age := qryDb1.FieldByName('c14_age').AsString;
    c14age_sig := qryDb1.FieldByName('c14_age_sig').AsString;
    if Length(c14age)=0 then
    begin
      c14age:='99999';
      c14age_sig := '99999';
    end
    else
    begin
      fm := qryDb1.FieldByName('fm').AsFloat;
      fm_sig := qryDb1.FieldByName('fm_sig').AsFloat;
      if fm_sig<0.0005 then fm_sig := 0.0005;
      if fm<2*fm_sig then
      begin
        fC14Age := -Ln(fm+2*fm_sig)*8033;
        fC14Age := round(fC14Age/100)*100;
        C14Age :=  Format('%5.0f',[fC14Age]);
        c14age_sig := '99999';
      end;
    end;
    CommandText := 'UPDATE sample_t' +
                   ' INNER JOIN target_t ON sample_t.sample_nr=target_t.sample_nr' +
                   ' SET sample_t.c14_age=' + c14age + ',' +
                   '  sample_t.c14_age_sig=' + c14age_sig + ',' +
                   '  sample_t.av_fm=' + qryDb1.FieldByName('fm').AsString + ',' +
                   '  sample_t.av_fm_sig=' + qryDb1.FieldByName('fm_sig').AsString +',' +
                   '  sample_t.av_dc13=' + qryDb1.FieldByName('dc13').AsString +',' +
//                   '  sample_t.av_dc13_sig=' + qryDb1.FieldByName('dc13_sig').AsString + ',' +
                   '  sample_t.Cal1sMin=' + qryDb1.FieldByName('Cal1sMin').AsString + ',' +
                   '  sample_t.Cal1sMax=' + qryDb1.FieldByName('Cal1sMax').AsString + ',' +
                   '  sample_t.Cal2sMin=' + qryDb1.FieldByName('Cal2sMin').AsString + ',' +
                   '  sample_t.Cal2sMax=' + qryDb1.FieldByName('Cal2sMax').AsString +
                   ' WHERE sample_t.sample_nr=' + IntToStr(SampleNr) +
                   ';';
    s := CommandText;
    FormatSettings.DecimalSeparator := c;
//    ClipBoard.SetTextBuf(PChar(s));
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;

end;

procedure Tdm.TransferMA_Nr_To_MAMS(limit: string);
// transfer MA Invoice numbers from CEZA Database to MAMS database
var
  MAMSProjectNr, MA_Nr, MA_Auftrags_Nr, MAMSAuftragsNr, ExitingMAMSAuftragsNr : integer;
begin
  with qryCEZA do
  begin
    LogWindow.addLogEntry('getting data from MA-DB. Limit is set to ' + limit);
    // query Main MA-Database for all Aufträge that have a "mams project nr"
    SQL.Text := 'SELECT Auftrags_Nr_,mams_project_nr FROM tab_auftraege t where Auftrags_Nr_ IS NOT NULL and mams_project_nr IS NOT NULL order by mams_project_nr desc LIMIT ' + limit +';';
    LogWindow.addLogEntry(SQL.Text);
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Open;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Open;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;

    // got to first entry
    First;
    LogWindow.addLogEntry('----------------------------');
    while not EOF do
      begin
        MAMSProjectNr := qryCEZA.Fields.Fields[1].AsInteger;
        MA_Auftrags_Nr := qryCEZA.Fields.Fields[0].AsInteger;
        LogWindow.addLogEntry('MA-DB says: MAMS Project = ' + inttostr(MAMSProjectNr) + ' has MA_Auftrags_Nr = ' + inttostr(MA_Auftrags_Nr));
         with qryDb do
         begin
           // display AuftragsNr of the Project in the MAMS DB, just for logging
           SQL.Text := 'SELECT AuftragsNr FROM project_t WHERE Project_nr=' + IntToStr(MAMSProjectNr) + ';';
           LogWindow.addLogEntry(SQL.Text);
           IF dm.adoConnKTL.Connected THEN
                  Begin
                    Try
                      Open;
                      LogWindow.addLogEntry('DB -- query executed');
                    Except
                      // connections is closed, reconnect by setting connected to true and test again
                      LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                      Try
                        dm.DBreconnect;
                        LogWindow.addLogEntry('DB -- execute query again');
                        Open;
                      Except
                         LogWindow.addLogEntry('DB -- can not reconnect');
                      End;
                      //ShowMessage('Database connection is closed.');
                    End;
                  End;
           ExitingMAMSAuftragsNr := Fields.Fields[0].AsInteger; // this is the AuftragsNr that is already in MAMS DB, maybe empty, then it needs updating
           LogWindow.addLogEntry('MAMS-DB says: MAMS Project = ' + inttostr(MAMSProjectNr) + ' already has AuftragsNr = ' + inttostr(ExitingMAMSAuftragsNr));

           if  ExitingMAMSAuftragsNr > 0 then
           Begin
             LogWindow.addLogEntry('no update needed');
           End
           Else
             BEgin
             LogWindow.addLogEntry('update needed');
               with adoCmd do
               begin
                 CommandText := 'UPDATE project_t SET AuftragsNr=' + IntToStr(MA_Auftrags_Nr) +
                                ' WHERE Project_nr=' + IntToStr(MAMSProjectNr) + ' AND AuftragsNr IS NULL;';
                                //' SELECT ROW_COUNT() AS numbersUpdated;';
                 LogWindow.addLogEntry(CommandText);
                 IF dm.adoConnKTL.Connected THEN
                    Begin
                      Try
                        Execute;
                        LogWindow.addLogEntry('DB -- query executed');
                      Except
                        // connections is closed, reconnect by setting connected to true and test again
                        LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                        Try
                          dm.DBreconnect;
                          LogWindow.addLogEntry('DB -- execute query again');
                          Open;
                        Except
                           LogWindow.addLogEntry('DB -- can not reconnect');
                        End;
                        //ShowMessage('Database connection is closed.');
                      End;
                    End;
               end;
             End
         end;
         Next;
      end;
          LogWindow.addLogEntry('----------------------------');
    end;
end;

procedure Tdm.TransferLeHeCurrentToTarget(SampleNr, PrepNr, TargetNr, le_current, he_current: string);
var
  s : string;
  c : char;
begin
  with adoCmd do
  begin
    c := FormatSettings.DecimalSeparator;
    FormatSettings.DecimalSeparator := '.';
    CommandText := 'UPDATE target_t ' +
                   'SET le_curr=' + le_current + ', ' +
                   'he_curr=' + he_current + ' ' +
                   'WHERE target_t.sample_nr=' + SampleNr + ' ' +
                   'AND prep_nr=' + PrepNr + ' ' +
                   'AND target_nr=' + TargetNr +
                   ';';
    s := CommandText;
    LogWindow.addLogEntry(s);
    FormatSettings.DecimalSeparator := c;
//    ClipBoard.SetTextBuf(PChar(s));
    IF dm.adoConnKTL.Connected THEN
                Begin
                  Try
                    Execute;
                    LogWindow.addLogEntry('DB -- query executed');
                  Except
                    // connections is closed, reconnect by setting connected to true and test again
                    LogWindow.addLogEntry('DB -- connection is closed, reconnecting...');
                    Try
                      dm.DBreconnect;
                      LogWindow.addLogEntry('DB -- execute query again');
                      Execute;
                    Except
                       LogWindow.addLogEntry('DB -- can not reconnect');
                    End;
                    //ShowMessage('Database connection is closed.');
                  End;
                End;
  end;

end;

function Tdm.ExtractNumber(Input: string): string;
// extract the fractional number from a string
// remove als other characters other than numbers and decimal points
var
  i: Integer;
  ResultBuilder: string;
begin
  ResultBuilder := '';
  for i := 1 to Length(Input) do
  begin
    // Allow digits, the decimal point, and the comma
    if CharInSet(Input[i], ['0'..'9', '.', ',']) then
      ResultBuilder := ResultBuilder + Input[i];
  end;
  // Replace the comma with a decimal point if necessary
  Result := StringReplace(ResultBuilder, ',', '.', [rfReplaceAll]);
end;

function Tdm.ReplaceBadCharacters(s: string): string;
Var
  i: integer;
begin
 // replaces 'bad' characters
   //result := '';
  s := StringReplace(s, ';', ',', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, '&', '_', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, '%', '_', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, '$', '_', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, '?', '', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, '"', '', [rfReplaceAll, rfIgnoreCase]);

  result := s;

//  for i := 1 to length(s) do
//  begin
//    Case s[i] of
//    '"': result := result + '';
//    ';': result := result + ',';
//    '?': result := result + '';
//    else result := result+s[i];
//    end;
//  end;
end;

function Tdm.ReplaceUmlaute(s: string): string;
var i: integer;
begin
  result := '';
  for i := 1 to length(s) do
  begin
    Case s[i] of
    'ä': result := result+'ae';
    'ü': result := result+'ue';
    'ö': result := result+'oe';
    'ß': result := result+'ss';
    else result := result+s[i];
    end;
  end;
end;


end.
