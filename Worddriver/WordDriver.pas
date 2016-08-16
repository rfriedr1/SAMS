unit WordDriver;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,DB,
  word2000, word97, WordXP, Word2003, Word2007 {$ifndef ver130},variants{$endif};

type
  TWordVersion = (wvDetect,wvGeneric,wvWord97,wvWord2000,wvWord2003,
                  wvWord2007,wvWordxp,wvOpenOffice);
  TDocumentMode = (dmSerialLetter,dmNewtable,dmFillTable);

  TOLEWord = class(TComponent)
  Private
    FHaveDocument : Boolean;
    FCancelled : Boolean;
    procedure DisconnectWord(QuitWord : Boolean); Virtual; Abstract;
    procedure OpenWord(WithVisible : Boolean); Virtual; Abstract;
    procedure OpenDocument(FN : String); Virtual; Abstract;
    procedure SaveDocument(FN : String); Virtual; Abstract;
    procedure DisconnectDocument(KeepOpen : Boolean); Virtual; Abstract;
    procedure PrintDocument; Virtual; Abstract;
    Procedure ReplaceFields(Value : TStrings); virtual;Abstract;
    Procedure ReplaceValues(Value : TStrings); virtual;Abstract;
    procedure GetMergeFieldNames(List: TStrings);virtual;abstract;
    procedure GetFormFieldNames(List: TStrings);virtual;abstract;
    Property  HaveDocument : boolean read FHaveDocument;
    procedure GetTableFieldNames(List : TStrings; Dataset : TDataset); virtual; abstract;
    procedure InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings); virtual; abstract;
    procedure InsertTableFromTemplate(Dataset : TDataset); virtual; abstract;
    Procedure AddMergeDoc (FN : String; AddPageBreak : Boolean); virtual;abstract;
  end;

  TWordDriver = class(TComponent)
  private
    FBusy,
    FSave: Boolean;
    FPrint: Boolean;
    FDocumentCount : Integer;
    FOutputDir: String;
    FSaveTemplate: String;
    FFileName: String;
    FDatasource: TDataSource;
    FValues: TStrings;
    FCancelled: boolean;
    FKeepWordOpen: Boolean;
    FKeepDocumentsOpen: Boolean;
    FOnlyKeepMerge: Boolean;
    FWordVersion: TWordVersion;
    FDriver : TOLeWord;
    FOnlyCurrentRecord: Boolean;
    FOnProgress: TNotifyEvent;
    FDocumentMode: TDocumentMode;
    FTableTag: String;
    FDocumentParams,
    FDocumentNames,
    FColumnList: TStrings;
    FSaveFileName: String;
    FCreateMergeDoc: Boolean;
    FMergeDocFileName: String;
    FUseFormFields: Boolean;
    procedure SetDataSource(const Value: TDataSource);
    procedure SetFileName(const Value: String);
    procedure SetPrint(const Value: Boolean);
    procedure SetSave(const Value: Boolean);
    procedure SetValues(const Value: TStrings);
    procedure DriverError(MSg: String);
    procedure GetDatasetValues(L: TStrings);
    function GetSaveFileName(NR: Integer; Values: TStrings): String;
    procedure GetValues(TheValues: TStrings);
    procedure ResetCancel;
    function GetHaveDocument: Boolean;
    procedure SetColumnList(const Value: TStrings);
    procedure SetCreateMergeDoc(const Value: Boolean);
    procedure SetKeepDocumentsOpen(const Value: Boolean);
    procedure SetKeepWordOpen(const Value: Boolean);
    procedure SetOnlyKeepMergeOpen(const Value: Boolean);
    procedure ReplaceFields(Value: TStrings);
    { Private declarations }
  protected
    { Protected declarations }
    procedure DoDocument(Nr : Integer); virtual; // No need to override;
    procedure DisconnectWord(QuitWord: Boolean); Virtual;
    procedure OpenWord; Virtual;
    procedure OpenDocument(FN : String); Virtual;
    procedure SaveDocument(FN : String); Virtual;
    procedure DisconnectDocument(KeepOpen : Boolean); Virtual;
    procedure PrintDocument; Virtual;
    Procedure ReplaceValues(Value : TStrings); virtual;
    Procedure DoProgress; virtual;
    procedure InsertTable(Dataset: TDataset; NameTag: String;
      FieldList: TStrings);virtual;
    procedure InsertTableFromTemplate(Dataset: TDataset);virtual;
    procedure CreateNewTable;virtual;
    procedure CreateSerialDoc;virtual;
    procedure FillTable;virtual;
    Procedure CreateMergeDoc;
    Procedure AddMergeDoc (FN : String; AddPageBreak : Boolean);
  public
    { Public declarations }
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    Procedure Execute;
    Procedure Cancel;
    Property Cancelled : boolean Read FCancelled;
    function GetWordVersion: String;
    Procedure CreateDriver;
    procedure GetTableFieldNames(List : TStrings; Dataset : TDataset); virtual;
    procedure GetMergeFieldNames(List: TStrings);virtual;
    Property DocumentParams : TStrings Read FDocumentParams;
    Property HaveDocument : Boolean Read GetHaveDocument;
  published
    { Published declarations }
    Property FileName : String Read FFileName Write SetFileName;
    Property SaveFileName : String Read FSaveFileName Write FSaveFileName;
    Property Print : Boolean Read FPrint Write SetPrint;
    Property Save : Boolean Read FSave Write SetSave;
    Property OutputDir : String Read FOutputDir Write FOutputDir;
    Property SaveTemplate : String Read FSaveTemplate Write FSaveTemplate;
    Property Datasource : TDataSource Read FDatasource Write SetDataSource;
    Property Values : TStrings Read FValues Write SetValues;
    Property WordVersion : TWordVersion Read FWordVersion Write FWordVersion;
    Property OnlyCurrentRecord : Boolean Read FOnlyCurrentRecord Write FOnlyCurrentRecord;
    Property Busy : Boolean Read Fbusy;
    Property DocumentCount : Integer Read FDocumentCount;
    Property OnProgress : TNotifyEvent Read FOnProgress Write FOnProgress;
    Property DocumentMode : TDocumentMode Read FDocumentMode Write FDocumentMode;
    Property TableTag : String read FTableTag Write FTableTag;
    Property ColumnList : TStrings Read FColumnList write SetColumnList;
    Property CreateMergeDocument : Boolean Read FCreateMergeDoc Write SetCreateMergeDoc;
    Property KeepWordOpen : Boolean Read FKeepWordOpen Write SetKeepWordOpen;
    Property KeepDocumentsOpen : Boolean Read FKeepDocumentsOpen Write SetKeepDocumentsOpen;
    Property KeepOnlyMergeOpen : Boolean Read FOnlyKeepMerge Write SetOnlyKeepMergeOpen;
    Property MergeDocumentFileName : String Read FMergeDocFileName Write FMergeDocFileName;
    property UseFormFields : Boolean Read FUseFormFields Write FUseFormFields;
  end;

  TWord97Driver = Class(TOleWord)
  private
    FWord : Word97.TWordApplication;
    FDocument : Word97.TWordDocument;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings); override;
    procedure InsertTableFromTemplate(Dataset : TDataset); override;
    procedure GetTableFieldNames(List : TStrings; Dataset : TDataset); override;
    Procedure AddMergeDoc (FN : String; AddPageBreak : Boolean); override;
  end;

  TWord2000Driver = Class(TOleWord)
  private
    FWord : Word2000.TWordApplication;
    FDocument : Word2000.TWordDocument;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean); override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset); override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset); override;
  end;

  TWord2003Driver = Class(TOleWord)
  private
    FWord : Word2003.TWordApplication;
    FDocument : Word2003.TWordDocument;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean); override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset); override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset); override;
  end;

  TWord2007Driver = Class(TOleWord)
  private
    FWord : Word2007.TWordApplication;
    FDocument : Word2007.TWordDocument;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean); override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset); override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset); override;
  end;

  TWordXPDriver = Class(TOleWord)
  private
    FWord : WordXP.TWordApplication;
    FDocument : WordXP.TWordDocument;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean); override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset); override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset);override;
  end;

  EWordDriver = CLass(Exception);

  TGenericWordDriver = Class(TOleWord)
  private
    FWord : Variant;
    FDocument : Variant;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean);override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset);override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset);override;
  end;

  TOfficeWriterDriver = Class(TOleWord)
  private
    FOffice : Variant;
    FCoreReflection : Variant;
    FDesktop : Variant;
    FDocument : Variant;
  Public
    Constructor Create(AOwner : TComponent);override;
    Destructor Destroy; override;
    Procedure CreatePropertyValue(Var PropertyValue : Variant; APropertyName : String);
    procedure DisconnectWord(QuitWord : Boolean); override;
    procedure OpenWord(WithVisible : Boolean); override;
    procedure OpenDocument(FN : String); Override;
    procedure SaveDocument(FN : String); Override;
    procedure DisconnectDocument(KeepOpen : Boolean); Override;
    procedure PrintDocument; override;
    Procedure ReplaceFields(Value : TStrings); override;
    Procedure ReplaceValues(Value : TStrings); override;
    procedure GetMergeFieldNames(List: TStrings);override;
    procedure GetFormFieldNames(List: TStrings);override;
    procedure AddMergeDoc(FN: String; AddPageBreak: Boolean);override;
    procedure GetTableFieldNames(List: TStrings; Dataset: TDataset);override;
    procedure InsertTable(Dataset: TDataset; NameTag: String;FieldList: TStrings);override;
    procedure InsertTableFromTemplate(Dataset: TDataset);override;
  end;

// procedure Register;

implementation

uses comobj;

resourcestring
  SErrKeepOpenNeedsSave = 'Documents must be saved if Word must remain open.';
  SErrNoSuchFile = 'Bestand "%s" bestaat niet.';
  SErrGettingVersion = 'Kon de versie van MS Word niet detecteren.';
  SErrUnSupportedWordVersion = 'Versie "%s" van MS Word wordt niet ondersteund.';
  SErrNoSuchDriver = 'The requested driver is not implemented.';
  SErrNoSuchTable = 'Veld voor tabel "%s" niet gevonden.';
  SErrNoTablesInDocument = 'Er zijn geen tabellen gevonden in dit document.';
  SErrNoDatasetAvailableForTable = 'Geen gegevens beschikbaar voor aanmaken tabel';

{
procedure Register;
begin
  RegisterComponents('Extra', [TWordDriver]);
end;
}
{ TWordDriver }

procedure TWordDriver.Cancel;
begin
  FCancelled:=True;
  If Assigned(FDriver) then
    FDriver.FCancelled:=True;
end;

procedure TWordDriver.ResetCancel;

begin
  If Assigned(FDriver) then
    FDriver.FCancelled:=False;
  FCancelled:=False;
end;

constructor TWordDriver.Create(AOwner: TComponent);
begin
  Inherited;
  FValues:=TStringList.Create;
  FColumnList:=TStringList.Create;
  FDocumentParams:=TStringList.Create;
  FDocumentNames:=TStringList.Create;
  TStringList(FDocumentParams).Sorted:=True;
end;

destructor TWordDriver.Destroy;
begin
  FValues.Free;
  FColumnList.Free;
  FDocumentParams.Free;
  FDocumentNames.Free;
  inherited;
end;

procedure TWordDriver.GetDatasetValues(L : TStrings);

Var
  I : Integer;

begin
  If (DataSource=Nil) or (Datasource.Dataset=Nil) then
    exit;
  With Datasource.Dataset,Fields do
    begin
    If (Not Active) or (EOF and BOF) then
      Exit;
    For i:=0 to Count-1 do
      With Fields[i] do
        if (FDocumentParams.Count=0) or
           (FDocumentParams.IndexOf(FieldName)<>-1) then
          L.Add(FieldName+'='+AsString);
    end;
end;

Function TWordDriver.GetSaveFileName(NR : Integer; Values : TStrings) : String;

Var
  S,R,N : String;
  I,J : Integer;

begin
  S:=SaveTemplate;
  If S='' then
    S:='%N%';
  S:=StringReplace(S,'%N%',IntToStr(Nr),[rfReplaceAll]);
  R:=FormatDateTime('YYYYMMDD',date);
  S:=StringReplace(S,'%D%',R,[rfReplaceAll]);
  For I:=0 to Values.Count-1 do
    begin
    R:=Values[I];
    J:=Pos('=',R);
    N:='%'+Copy(R,1,J-1)+'%';
    System.Delete(R,1,J);
    S:=StringReplace(S,N,R,[rfReplaceAll,rfIgnoreCase]);
    end;
  Result:=IncludeTrailingBackSlash(OutputDir)+S;
end;

procedure TWordDriver.GetValues(TheValues : TStrings);

Var
  S : String;
  I,J : Integer;

begin
  If FDocumentParams.Count=0 then
    Thevalues.Assign(Values)
  else begin
    For I:=0 to Values.Count-1 do begin
      S:=Values[i];
      J:=Pos('=',S);
      If (J<>0) and (FDocumentParams.IndexOf(Copy(S,1,J-1))<>-1) then
        TheValues.Add(S);
    end;
  end;
//  TheValues.SaveToFile('c:\temp\thevalues.txt');
end;

procedure TWordDriver.DoDocument(NR : Integer);

Var
  TheValues : TStringList;

begin
  OpenDocument(FFileName);
  Inc(FDocumentCount);
  If (NR=1) then
    GetMergeFieldNames(FDocumentParams);
  Try
    TheValues:=TStringList.Create;
    Try
      GetValues(TheValues);
      GetDatasetValues(TheValues);
      ReplaceValues(TheValues);
      If FuseFormFields then
        ReplaceFields(TheValues);
      If Not FCancelled then
        begin
        If Print then
          PrintDocument;
        If Save then
          SaveDocument(GetSaveFileName(Nr,TheValues));
        end;
    Finally
      TheValues.Free;
    end;
  finally
    DisconnectDocument(FKeepDocumentsOpen and (not FOnlyKeepMerge));
  end;
end;


procedure TWordDriver.Execute;

begin
  FBusy:=True;
  Try
    FDocumentNames.Clear;
    FDocumentCount:=0;
    If Not FileExists(FFileName) then
      DriverError(Format(SErrNoSuchFile,[FFileName]));
    OpenWord;
    Try
      ResetCancel;
      Case DocumentMode of
        dmSerialLetter : CreateSerialDoc;
        dmNewTable     : CreateNewTable;
        dmFillTable    : FillTable;
      end;
    Finally
      DisconnectWord(Not FKeepWordOpen);
    end;
  Finally
    FBusy:=False;
  end;
end;

Procedure TWordDriver.CreateSerialDoc;
Var
  I : Integer;

begin
  If (Datasource=Nil) or (DataSource.Dataset=Nil) or (FOnlyCurrentRecord) then
    DoDocument(1)
  else
    begin
    With DataSource.Dataset do
      begin
      First;
      I:=0;
      While Not EOF or FCancelled do
        begin
        Inc(I);
        DoDocument(i);
        Next;
        end;
      end;
    If CreateMergeDocument then
      CreateMergeDoc;
    end;
end;

Procedure TWordDriver.CreateNewTable;

Var
  TheValues : TStrings;

begin
  If Not (Assigned(FDataSource) and Assigned(DataSource.Dataset)) then
    DriverError(SErrNoDatasetAvailableForTable);
  OpenDocument(FFileName);
  Try
    Inc(FDocumentCount);
    TheValues:=TStringList.Create;
    Try
      GetValues(TheValues);
      ReplaceValues(TheValues);
      If Not FCancelled then
        InsertTable(Datasource.Dataset,TableTag,FColumnList);
      If Not FCancelled then
        begin
        If Print then
          PrintDocument;
        If Save then
          SaveDocument(SaveFileName);
        end;
    finally
      TheValues.Free;
    end;
  finally
    DisconnectDocument(FKeepDocumentsOpen);
  end;
end;

Procedure TWordDriver.FillTable;

Var
  TheValues : TStrings;

begin
  If Not (Assigned(FDataSource) and Assigned(DataSource.Dataset)) then
    DriverError(SErrNoDatasetAvailableForTable);
  OpenDocument(FFileName);
  Try
    Inc(FDocumentCount);
    TheValues:=TStringList.Create;
    Try
      GetValues(TheValues);
      ReplaceValues(TheValues);
      If Not FCancelled then
        InsertTableFromTemplate(DataSource.Dataset);
      If Not FCancelled then
        begin
        If Print then
          PrintDocument;
        If Save then
          SaveDocument(SaveFileName);
        end;
    finally
      TheValues.Free;
    end;
  finally
    DisconnectDocument(FKeepDocumentsOpen);
  end;
end;


function TWordDriver.GetWordVersion : String;

Var
  M : Variant;

begin
  Try
    try
      M:=CreateOleObject('Word.Application');
      Result:=M.Application.Version;
    Except
      DriverError(SErrGettingVersion);
    end;
  Finally
    VarClear(M);
  end;
end;

procedure TWordDriver.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited;
   If (Operation=opRemove) and (AComponent=FDataSource) then
     FDatasource:=Nil;
end;

procedure TWordDriver.SetDataSource(const Value: TDataSource);
begin
  FDatasource := Value;
  FDataSource.FreeNotification(Self);
end;

procedure TWordDriver.SetFileName(const Value: String);
begin
  FFileName := Value;
end;

procedure TWordDriver.DriverError(MSg : String);

begin
  Raise EWordDriver.Create(Msg);
end;

procedure TWordDriver.SetPrint(const Value: Boolean);
begin
  FPrint := Value;
end;

procedure TWordDriver.SetSave(const Value: Boolean);
begin
  FSave := Value;
end;

procedure TWordDriver.SetValues(const Value: TStrings);
begin
  FValues.Assign(Value);
end;


procedure TWordDriver.DisconnectDocument(KeepOpen : Boolean);
begin
  FDriver.DisconnectDocument(KeepOpen);
end;

procedure TWordDriver.DisconnectWord(QuitWord : Boolean);
begin
  FDriver.DisconnectWord(QuitWord);
end;

procedure TWordDriver.GetMergeFieldNames(List: TStrings);

Var
  W,F : Boolean;

begin
  W:=False;
  F:=not HaveDocument;
  If F then
    begin
    W:=not Assigned(FDriver);
    If W then
      OpenWord;
    OpenDocument(FileName);
    end;
  Try
    FDriver.GetMergeFieldNames(List);
    If UseFormFields then
      FDriver.GetFormFieldNames(List);
  finally
    If F then
      DisconnectDocument(False);
    if W then
      DisconnectWord(True);
  end;
end;

procedure TWordDriver.OpenDocument(FN: String);
begin
  FDriver.OpenDocument(FN);
end;

procedure TWordDriver.OpenWord;
begin
  If FDriver=Nil then
    CreateDriver;
  FDriver.OpenWord(FKeepWordOpen);
end;

procedure TWordDriver.PrintDocument;
begin
  FDriver.PrintDocument;
end;

procedure TWordDriver.ReplaceValues(Value: TStrings);
begin
  FDriver.ReplaceValues(Value);
end;

procedure TWordDriver.ReplaceFields(Value: TStrings);
begin
  FDriver.ReplaceFields(Value);
end;

procedure TWordDriver.SaveDocument(FN: String);
begin
  FDriver.SaveDocument(FN);
  FDocumentNames.Add(FN);
end;

procedure TWordDriver.CreateDriver;

Var
  V : TWordVersion;
  RV : String;

begin
  V:=FWordVersion;
  If (V=wvDetect) then
    begin
    RV:=GetWordVersion;
    If Pos('8.',RV)=1 then
      V:=wvWord97
    else if pos('9.',RV)=1 then
      V:=wvWord2000
    else if pos('10.',RV)=1 then
      V:=wvWordxp
    else if pos('11.',RV)=1 then
      V:=wvWord2003
    else if pos('12.',RV)=1 then
      V:=wvWord2007
    else // try generic driver.
      V:=wvGeneric;
      //DriverError(Format(SErrUnSupportedWordVersion,[RV]));
    end;
  Case V of
    wvWord97 : FDriver:=TWord97Driver.Create(Self);
    wvWord2000 : FDriver:=TWord2000Driver.Create(Self);
    wvWord2003 : FDriver:=TWord2003Driver.Create(Self);
    wvWord2007 : FDriver:=TWord2007Driver.Create(Self);
    wvWordXP : FDriver:=TWordXPDriver.Create(Self);
    wvGeneric : FDriver:=TGenericWordDriver.Create(Self);
    wvOpenOffice : FDriver:=TOfficeWriterDriver.Create(Self);
  else
    DriverError(SErrNoSuchDriver);
  end;
end;

function TWordDriver.GetHaveDocument: Boolean;
begin
  Result:=Assigned(FDriver) and (FDriver.HaveDocument);
end;

procedure TWordDriver.DoProgress;
begin
  If Assigned(FOnPRogress) then
    FOnProgress(Self);
end;

procedure TWordDriver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

begin
  FDriver.InsertTable(Dataset,NameTag,FieldList);
end;

procedure TWordDriver.InsertTableFromTemplate(Dataset : TDataset);

begin
  FDriver.InsertTableFromTemplate(Dataset)
end;


procedure TWordDriver.GetTableFieldNames(List: TStrings; Dataset : TDataset);
begin
  FDriver.GetTableFieldNames(List,Dataset);
end;


procedure TWordDriver.SetColumnList(const Value: TStrings);
begin
  FColumnList.Assign(Value);
end;

procedure TWordDriver.AddMergeDoc(FN: String; AddPageBreak: Boolean);
begin
  FDriver.AddMergeDoc(FN,AddPageBreak);
end;

procedure TWordDriver.CreateMergeDoc;

Var
  I : Integer;

begin
  If FDocumentNames.Count=0 then
    Exit;
  If FileExists(MergeDocumentFileName) then
    OpenDocument(MergeDocumentFileName)
  else
    OpenDocument('');
  try
    With FDocumentNames do
      For I:=0 to Count-1 do
        AddMergeDoc(Strings[i],(i<(Count-1)));
    SaveDocument(MergeDocumentFileName);
  Finally
    DisconnectDocument(FKeepDocumentsOpen);
  end;
end;

procedure TWordDriver.SetCreateMergeDoc(const Value: Boolean);
begin
  FCreateMergeDoc := Value;
  If Not Value then
    FOnlyKeepMerge:=False;
end;

procedure TWordDriver.SetKeepDocumentsOpen(const Value: Boolean);
begin
  If Value and not FSave Then
    DriverError(SErrKeepOpenNeedsSave);
  FKeepDocumentsOpen := Value;
  If Not Value then
    FOnlyKeepMerge := False;
end;

procedure TWordDriver.SetKeepWordOpen(const Value: Boolean);
begin
  FKeepWordOpen := Value;
  If Not Value then
    KeepDocumentsOpen:=False;
end;

procedure TWordDriver.SetOnlyKeepMergeOpen(const Value: Boolean);
begin
  If FCreateMergeDoc then
    FOnlyKeepMerge:=Value;
end;

{ TWord97Driver }

procedure TWord97Driver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  FDocument.Disconnect;
  FHaveDocument:=False;
end;

procedure TWord97Driver.DisconnectWord(QuitWord : Boolean);

Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  FWord.Disconnect;
end;

constructor TWord97Driver.Create(AOwner: TComponent);
begin
  inherited;
  FWord:=Word97.TWordApplication.Create(Self);
  FDocument:=Word97.TWordDocument.Create(Self);
end;

destructor TWord97Driver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TWord97Driver.OpenDocument(FN: String);

Var
  D : word97._Document;
  O : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  FHaveDocument:=False;
  If (FN<>'') then
    begin
    O:=FN;
    D:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
       EmptyParam,EmptyParam,EmptyParam,EmptyParam);
    end
  else
    D:=FWord.Documents.Add(Emptyparam,EmptyParam);
  FDocument.ConnectTo(D as Word97._Document);
  FHaveDocument:=True;
end;

procedure TWord97Driver.OpenWord(WithVisible : Boolean);
begin
  FWord.Connect;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TWord97Driver.GetFormFieldNames(List: TStrings);

Var
  F : Word97.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;
end;

procedure TWord97Driver.GetMergeFieldNames(List : TStrings);

Var
  R : Word97.Range;
  S : String;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='\{[!\{]@\}';
    ClearFormatting;
    MatchWildcards:=True;
    MatchCase:=False;
    Repeat
      If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
                 EmptyParam,EmptyParam,EmptyParam,
                 EmptyParam,EmptyParam,EmptyParam) then
        begin
        S:=R.Text;
        Delete(S,1,1);
        SetLength(S,Length(S)-1);
        If (Pos(' ',S)=0) and (length(S)<32) then
          List.Add(S);
        end;
    Until not Found;
    end;
end;


procedure TWord97Driver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TWord97Driver.ReplaceFields(Value: TStrings);

Var
  F : Word97.FormField;
  N : String;
  I : Integer;
  O : olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TWord97Driver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      With FDocument.Content.Find do
        begin
        Text:=S;
        ClearFormatting;
        Replacement.Text:=R;
        ReplaceMent.ClearFormatting;
        D1:=True;
        D2:=wdfindcontinue;
        D3:=wdReplaceAll;
        Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,D1,D2,EMptyParam,emptyParam,D3);
        end;
      end;
    Inc(I);
    end;
end;

procedure TWord97Driver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TWord97Driver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  R : Word97.Range;
  T : Word97.Table;
  TR : Word97.Row;
  I : Integer;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='{'+NameTag+'}';
    ClearFormatting;
    MatchWildcards:=False;
    MatchCase:=False;
    If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,EmptyParam,EmptyParam) then
      begin
      R.Text:='';
      If FieldList=Nil then
        begin
        FieldList:=TStringList.Create;
        For I:=0 to Dataset.Fields.Count-1 do
          FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
        end
      else
        For I:=0 to FieldList.Count-1 do
          FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
      T:=R.Tables.Add(R,1,FieldList.Count);
      TR:=T.Rows.Item(1);
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          begin
          Bold:=1;
          Text:=Tfield(FieldList.Objects[i]).DisplayName;
          end;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        TR:=T.Rows.Add(EmptyParam);
        For I:=0 to FieldList.Count-1 do
          With TR.Cells.Item(I+1).Range do
            begin
            Text:=Tfield(FieldList.Objects[i]).AsString;
            Bold:=0;
            end;
        Dataset.Next;
        end;
      end
    else
      Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
    end;
end;

procedure TWord97Driver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R : Word97.Range;
  TR : Word97.Row;
  T : Word97.Table;
  I : Integer;
  FieldList : TStrings;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          If Assigned(FieldList.Objects[i]) then
            Text:=Tfield(FieldList.Objects[i]).AsString
          else
            //Text:='';
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TWord97Driver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R : Word97.Range;
  TR : Word97.Row;
  T : Word97.Table;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TWord97Driver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : Word97.Range;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;
{ TWord2000Driver }

procedure TWord2000Driver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  FDocument.Disconnect;
  FHaveDocument:=False;
end;

procedure TWord2000Driver.DisconnectWord(QuitWord : Boolean);
Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  FWord.Disconnect;
end;

constructor TWord2000Driver.Create(AOwner: TComponent);
begin
  inherited;
  FWord:=Word2000.TWordApplication.Create(Self);
  FDocument:=Word2000.TWordDocument.Create(Self);
end;

destructor TWord2000Driver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TWord2000Driver.OpenDocument(FN: String);

Var
  D : word2000._Document;
  O : OleVariant;

begin
  FHaveDocument:=False;
  If (FN<>'') then
    begin
    O:=FN;
    D:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EMptyParam,
                            EmptyParam) as Word2000._Document;
    end
  else
    D:=FWord.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
  FDocument.ConnectTo(D as Word2000._Document);
  FHaveDocument:=True;
end;

procedure TWord2000Driver.OpenWord(WithVisible : Boolean);
begin
  FWord.Connect;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TWord2000Driver.GetFormFieldNames(List: TStrings);

Var
  F : Word2000.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;
end;

procedure TWord2000Driver.GetMergeFieldNames(List : TStrings);

Var
  R : Word2000.WordRange;  //used to be Word2000.Range but caused error RF
  S : String;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='\{[!\{]@\}';
    ClearFormatting;
    MatchWildcards:=True;
    MatchCase:=False;
    Repeat
      If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
                 EmptyParam,EmptyParam,EmptyParam,
                 EmptyParam,EmptyParam,EmptyParam,EMptyParam,EMptyParam,
                 EMptyParam,EmptyParam) then
        begin
        S:=R.Text;
        Delete(S,1,1);
        SetLength(S,Length(S)-1);
        If (Pos(' ',S)=0) and (length(S)<32) then
          List.Add(S);
        end;
    Until not Found;
    end;
end;


procedure TWord2000Driver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TWord2000Driver.ReplaceFields(Value: TStrings);

Var
  F : Word2000.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TWord2000Driver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;

begin
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      With FDocument.Content.Find do
        begin
        Text:=S;
        ClearFormatting;
        Replacement.Text:=R;
        ReplaceMent.ClearFormatting;
        D1:=True;
        D2:=wdfindcontinue;
        D3:=wdReplaceAll;
        Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                EmptyParam,D1,D2,EMptyParam,emptyParam,D3,EMptyParam,
                EmptyParam,EmptyParam,EmptyParam);
        end;
      end;
    Inc(I);
    end;
end;

procedure TWord2000Driver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TWord2000Driver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  R : Word2000.WordRange;
  T : Word2000.Table;
  TR : Word2000.Row;
  I : Integer;

begin
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='{'+NameTag+'}';
    ClearFormatting;
    MatchWildcards:=False;
    MatchCase:=False;
    If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,emptyparam,emptyParam,EmptyParam,EmptyParam) then
      begin
      R.Text:='';
      If FieldList=Nil then
        begin
        FieldList:=TStringList.Create;
        For I:=0 to Dataset.Fields.Count-1 do
          FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
        end
      else
        For I:=0 to FieldList.Count-1 do
          FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
      T:=R.Tables.Add(R,1,FieldList.Count,EmptyParam,EmptyParam);
      TR:=T.Rows.Item(1);
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          begin
          Bold:=1;
          Text:=Tfield(FieldList.Objects[i]).DisplayName;
          end;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        TR:=T.Rows.Add(EmptyParam);
        For I:=0 to FieldList.Count-1 do
          With TR.Cells.Item(I+1).Range do
            begin
            Text:=Tfield(FieldList.Objects[i]).AsString;
            Bold:=0;
            end;
        Dataset.Next;
        end;
      end
    else
      Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
    end;
end;

procedure TWord2000Driver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R : Word2000.WordRange;
  TR : Word2000.Row;
  T : Word2000.Table;
  I : Integer;
  FieldList : TStrings;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          If Assigned(FieldList.Objects[i]) then
            Text:=Tfield(FieldList.Objects[i]).AsString
          else
            //Text:='';
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TWord2000Driver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R : Word2000.WordRange;
  TR : Word2000.Row;
  T : Word2000.Table;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TWord2000Driver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : Word2000.WordRange;

begin
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;

{ TWord2003Driver }

procedure TWord2003Driver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  FDocument.Disconnect;
  FHaveDocument:=False;
end;

procedure TWord2003Driver.DisconnectWord(QuitWord : Boolean);
Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  FWord.Disconnect;
end;

constructor TWord2003Driver.Create(AOwner: TComponent);
begin
  inherited;
  FWord:=Word2003.TWordApplication.Create(Self);
  FDocument:=Word2003.TWordDocument.Create(Self);
end;

destructor TWord2003Driver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TWord2003Driver.OpenDocument(FN: String);

Var
  D : word2003._Document;
  O : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  FHaveDocument:=False;
  If (FN<>'') then
    begin
    O:=FN;
    D:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EMptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam) as Word2003._Document;
    end
  else
    D:=FWord.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
  FDocument.ConnectTo(D as Word2003._Document);
  FHaveDocument:=True;
end;

procedure TWord2003Driver.OpenWord(WithVisible : Boolean);
begin
  FWord.Connect;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TWord2003Driver.GetFormFieldNames(List: TStrings);

Var
  F : Word2003.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;
end;

procedure TWord2003Driver.GetMergeFieldNames(List : TStrings);

Var
  R : Word2003.WordRange;
  S : String;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='\{[!\{]@\}';
    ClearFormatting;
    MatchWildcards:=True;
    MatchCase:=False;
    Repeat
      If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
                 EmptyParam,EmptyParam,EmptyParam,
                 EmptyParam,EmptyParam,EmptyParam,EMptyParam,EMptyParam,
                 EMptyParam,EmptyParam) then
        begin
        S:=R.Text;
        Delete(S,1,1);
        SetLength(S,Length(S)-1);
        If (Pos(' ',S)=0) and (length(S)<32) then
          List.Add(S);
        end;
    Until not Found;
    end;
end;


procedure TWord2003Driver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TWord2003Driver.ReplaceFields(Value: TStrings);

Var
  F : Word2003.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TWord2003Driver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      With FDocument.Content.Find do
        begin
        Text:=S;
        ClearFormatting;
        Replacement.Text:=R;
        ReplaceMent.ClearFormatting;
        D1:=True;
        D2:=wdfindcontinue;
        D3:=wdReplaceAll;
        Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                EmptyParam,D1,D2,EMptyParam,emptyParam,D3,EMptyParam,
                EmptyParam,EmptyParam,EmptyParam);
        end;
      end;
    Inc(I);
    end;
end;

procedure TWord2003Driver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TWord2003Driver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  R : Word2003.WordRange;
  T : Word2003.Table;
  TR : Word2003.Row;
  I : Integer;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='{'+NameTag+'}';
    ClearFormatting;
    MatchWildcards:=False;
    MatchCase:=False;
    If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,emptyparam,emptyParam,EmptyParam,EmptyParam) then
      begin
      R.Text:='';
      If FieldList=Nil then
        begin
        FieldList:=TStringList.Create;
        For I:=0 to Dataset.Fields.Count-1 do
          FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
        end
      else
        For I:=0 to FieldList.Count-1 do
          FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
      T:=R.Tables.Add(R,1,FieldList.Count,EmptyParam,EmptyParam);
      TR:=T.Rows.Item(1);
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          begin
          Bold:=1;
          Text:=Tfield(FieldList.Objects[i]).DisplayName;
          end;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        TR:=T.Rows.Add(EmptyParam);
        For I:=0 to FieldList.Count-1 do
          With TR.Cells.Item(I+1).Range do
            begin
            Text:=Tfield(FieldList.Objects[i]).AsString;
            Bold:=0;
            end;
        Dataset.Next;
        end;
      end
    else
      Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
    end;
end;

procedure TWord2003Driver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R : Word2003.WordRange;
  TR : Word2003.Row;
  T : Word2003.Table;
  I : Integer;
  FieldList : TStrings;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
//    FieldList.SaveToFile('c:\temp\fieldlist.txt');
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          If Assigned(FieldList.Objects[i]) then
            Text:=Tfield(FieldList.Objects[i]).AsString
          else
            //Text:='';
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TWord2003Driver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R : Word2003.WordRange;
  TR : Word2003.Row;
  T : Word2003.Table;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TWord2003Driver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : Word2003.WordRange;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;

{ TWord2007Driver }

procedure TWord2007Driver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  FDocument.Disconnect;
  FHaveDocument:=False;
end;

procedure TWord2007Driver.DisconnectWord(QuitWord : Boolean);
Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  FWord.Disconnect;
end;

constructor TWord2007Driver.Create(AOwner: TComponent);
begin
  inherited;
  FWord:=Word2007.TWordApplication.Create(Self);
  FDocument:=Word2007.TWordDocument.Create(Self);
end;

destructor TWord2007Driver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TWord2007Driver.OpenDocument(FN: String);

Var
  D : word2007._Document;
  O : OleVariant;
EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  FHaveDocument:=False;
  If (FN<>'') then
    begin
    O:=FN;
    D:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EMptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam) as Word2007._Document;
    end
  else
    D:=FWord.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
  FDocument.ConnectTo(D as Word2007._Document);
  FHaveDocument:=True;
end;

procedure TWord2007Driver.OpenWord(WithVisible : Boolean);
begin
  FWord.Connect;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TWord2007Driver.GetFormFieldNames(List: TStrings);

Var
  F : Word2007.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;
end;

procedure TWord2007Driver.GetMergeFieldNames(List : TStrings);

Var
  R : Word2007.WordRange;
  S : String;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='\{[!\{]@\}';
    ClearFormatting;
    MatchWildcards:=True;
    MatchCase:=False;
    Repeat
      If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
                 EmptyParam,EmptyParam,EmptyParam,
                 EmptyParam,EmptyParam,EmptyParam,EMptyParam,EMptyParam,
                 EMptyParam,EmptyParam) then
        begin
        S:=R.Text;
        Delete(S,1,1);
        SetLength(S,Length(S)-1);
        If (Pos(' ',S)=0) and (length(S)<32) then
          List.Add(S);
        end;
    Until not Found;
    end;
end;


procedure TWord2007Driver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TWord2007Driver.ReplaceFields(Value: TStrings);

Var
  F : Word2007.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TWord2007Driver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      With FDocument.Content.Find do
        begin
        Text:=S;
        ClearFormatting;
        Replacement.Text:=R;
        ReplaceMent.ClearFormatting;
        D1:=True;
        D2:=wdfindcontinue;
        D3:=wdReplaceAll;
        Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                EmptyParam,D1,D2,EMptyParam,emptyParam,D3,EMptyParam,
                EmptyParam,EmptyParam,EmptyParam);
        end;
      end;
    Inc(I);
    end;
end;

procedure TWord2007Driver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TWord2007Driver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  R : Word2007.WordRange;
  T : Word2007.Table;
  TR : Word2007.Row;
  I : Integer;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='{'+NameTag+'}';
    ClearFormatting;
    MatchWildcards:=False;
    MatchCase:=False;
    If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,emptyparam,emptyParam,EmptyParam,EmptyParam) then
      begin
      R.Text:='';
      If FieldList=Nil then
        begin
        FieldList:=TStringList.Create;
        For I:=0 to Dataset.Fields.Count-1 do
          FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
        end
      else
        For I:=0 to FieldList.Count-1 do
          FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
      T:=R.Tables.Add(R,1,FieldList.Count,EmptyParam,EmptyParam);
      TR:=T.Rows.Item(1);
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          begin
          Bold:=1;
          Text:=Tfield(FieldList.Objects[i]).DisplayName;
          end;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        TR:=T.Rows.Add(EmptyParam);
        For I:=0 to FieldList.Count-1 do
          With TR.Cells.Item(I+1).Range do
            begin
            Text:=Tfield(FieldList.Objects[i]).AsString;
            Bold:=0;
            end;
        Dataset.Next;
        end;
      end
    else
      Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
    end;
end;

procedure TWord2007Driver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R : Word2007.WordRange;
  TR : Word2007.Row;
  T : Word2007.Table;
  I : Integer;
  FieldList : TStrings;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          If Assigned(FieldList.Objects[i]) then
            Text:=Tfield(FieldList.Objects[i]).AsString
          else
            //Text:='';
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TWord2007Driver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R : Word2007.WordRange;
  TR : Word2007.Row;
  T : Word2007.Table;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TWord2007Driver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : Word2007.WordRange;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;


{ TWordXPDriver }

procedure TWordXPDriver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  FDocument.Disconnect;
  FHaveDocument:=False;
end;

procedure TWordXPDriver.DisconnectWord(QuitWord : Boolean);
Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  FWord.Disconnect;
end;

constructor TWordXPDriver.Create(AOwner: TComponent);
begin
  inherited;
  FWord:=WordXP.TWordApplication.Create(Self);
  FDocument:=WordXP.TWordDocument.Create(Self);
end;

destructor TWordXPDriver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TWordXPDriver.OpenDocument(FN: String);

Var
  D : wordXP._Document;
  O : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  FHaveDocument:=False;
  If (FN<>'') then
    begin
    O:=FN;
    D:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,EMptyParam,
                            EmptyParam,EmptyParam,EmptyParam,EmptyParam) as WordXP._Document;
    end
  else
    D:=FWord.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
  FDocument.ConnectTo(D as WordXP._Document);
  FHaveDocument:=True;
end;

procedure TWordXPDriver.OpenWord(WithVisible : Boolean);
begin
  FWord.Connect;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TWordXPDriver.GetFormFieldNames(List: TStrings);

Var
  F : WordXP.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;

end;

procedure TWordXPDriver.GetMergeFieldNames(List : TStrings);

Var
  R : WordXP.Range;
  S : String;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='\{[!\{]@\}';
    ClearFormatting;
    MatchWildcards:=True;
    MatchCase:=False;
    Repeat
      If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
                 EmptyParam,EmptyParam,EmptyParam,
                 EmptyParam,EmptyParam,EmptyParam,EMptyParam,EMptyParam,
                 EMptyParam,EmptyParam) then
        begin
        S:=R.Text;
        Delete(S,1,1);
        SetLength(S,Length(S)-1);
        If (Pos(' ',S)=0) and (length(S)<32) then
          List.Add(S);
        end;
    Until not Found;
    end;
end;


procedure TWordXPDriver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TWordXPDriver.ReplaceFields(Value: TStrings);

Var
  F : WordXP.FormField;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TWordXPDriver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;
  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      With FDocument.Content.Find do
        begin
        Text:=S;
        ClearFormatting;
        Replacement.Text:=R;
        ReplaceMent.ClearFormatting;
        D1:=True;
        D2:=wdfindcontinue;
        D3:=wdReplaceAll;
        Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                EmptyParam,D1,D2,EMptyParam,emptyParam,D3,EMptyParam,
                EmptyParam,EmptyParam,EmptyParam);
        end;
      end;
    Inc(I);
    end;
end;

procedure TWordXPDriver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TWordXPDriver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  R : WordXP.Range;
  T : WordXP.Table;
  TR : WordXP.Row;
  I : Integer;

  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  With R.Find do
    begin
    Text:='{'+NameTag+'}';
    ClearFormatting;
    MatchWildcards:=False;
    MatchCase:=False;
    If Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,emptyparam,emptyParam,EmptyParam,EmptyParam) then
      begin
      R.Text:='';
      If FieldList=Nil then
        begin
        FieldList:=TStringList.Create;
        For I:=0 to Dataset.Fields.Count-1 do
          FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
        end
      else
        For I:=0 to FieldList.Count-1 do
          FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
      T:=R.Tables.Add(R,1,FieldList.Count,EmptyParam,EmptyParam);
      TR:=T.Rows.Item(1);
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          begin
          Bold:=1;
          Text:=Tfield(FieldList.Objects[i]).DisplayName;
          end;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        TR:=T.Rows.Add(EmptyParam);
        For I:=0 to FieldList.Count-1 do
          With TR.Cells.Item(I+1).Range do
            begin
            Text:=Tfield(FieldList.Objects[i]).AsString;
            Bold:=0;
            end;
        Dataset.Next;
        end;
      end
    else
      Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
    end;
end;

procedure TWordXPDriver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R : WordXP.Range;
  TR : WordXP.Row;
  T : WordXP.Table;
  I : Integer;
  FieldList : TStrings;

  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        With TR.Cells.Item(I+1).Range do
          If Assigned(FieldList.Objects[i]) then
            Text:=Tfield(FieldList.Objects[i]).AsString
          else
            //Text:='';
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TWordXPDriver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R : WordXP.Range;
  TR : WordXP.Row;
  T : WordXP.Table;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TWordXPDriver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : WordXP.Range;

  EmptyParam: OleVariant;
begin
  EmptyParam := Variants.EmptyParam;
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;


{ TGenericWordDriver }

procedure TGenericWordDriver.DisconnectDocument(KeepOpen : Boolean);

Var
  O : OleVariant;

begin
  If Not KeepOpen then
    begin
    O:=False;
    FDocument.Close(O);
    end;
  VarClear(FDocument);
  FHaveDocument:=False;
end;

procedure TGenericWordDriver.DisconnectWord(QuitWord : Boolean);

Var
  S : OleVariant;

begin
  If QuitWord then
    begin
    S:=False;
    FWord.Quit(S);
    end;
  VarClear(FWord)
end;

constructor TGenericWordDriver.Create(AOwner: TComponent);

begin
  VarClear(FWord);
  VarClear(FDocument)
end;

destructor TGenericWordDriver.Destroy;
begin
  FDocument.Free;
  FWord.Free;
  inherited;
end;

procedure TGenericWordDriver.OpenDocument(FN: String);

Var
  O : OleVariant;

begin
  FHaveDocument:=False;
  O:=FN;
  FDocument:=FWord.Documents.Open(O,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,
  EmptyParam,EmptyParam,EmptyParam,EmptyParam);
  FHaveDocument:=True;
end;

procedure TGenericWordDriver.OpenWord(WithVisible : Boolean);

Var
  M : Variant;

begin
  Try
    M:=CreateOleObject('Word.Application');
    FWord:=M.Application;
  Finally
    VarClear(M);
  end;
  If WithVisible then
    FWord.Visible:=True;
end;

procedure TGenericWordDriver.GetFormFieldNames(List: TStrings);

Var
  F : Olevariant;
  I : Integer;
  N : String;
  O : Olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    List.Add(N);
    end;
end;

procedure TGenericWordDriver.GetMergeFieldNames(List : TStrings);

Var
  R,F : Variant;

  S : String;

begin
  R:=FDocument.Content;
  F:=R.Find;
  F.Text:='\{[!\{]@\}';
  F.ClearFormatting;
  F.MatchWildcards:=True;
  F.MatchCase:=False;
  Repeat
    If F.Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,EmptyParam,EmptyParam) then
      begin
      S:=R.Text;
      Delete(S,1,1);
      SetLength(S,Length(S)-1);
      If (Pos(' ',S)=0) and (length(S)<32) then
        List.Add(S);
      end;
  Until not F.Found;
end;


procedure TGenericWordDriver.PrintDocument;

Var
  B,O : OleVariant;
begin
  B:=True;
  O:=wdPrintAllDocument;
  FDocument.PrintOut;//(B,EmptyParam,O);
end;

procedure TGenericWordDriver.ReplaceFields(Value: TStrings);

Var
  F : Olevariant;
  N : String;
  I : Integer;
  O : olevariant;

begin
  For I:=1 to FDocument.FormFields.Count do
    begin
    O:=I;
    F:=FDocument.FormFields.Item(O);
    N:=F.Name;
    If (Value.IndexOfName(N)<>-1) then
      F.Result:=Value.Values[N];
    end;
end;

procedure TGenericWordDriver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  D1,D2,D3 : OleVariant;
  F : Variant;

begin
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      F:=FDocument.Content.Find;
      F.Text:=S;
      F.ClearFormatting;
      F.Replacement.Text:=R;
      F.ReplaceMent.ClearFormatting;
      D1:=True;
      D2:=wdfindcontinue;
      D3:=wdReplaceAll;
      F.Execute(EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,D1,D2,EMptyParam,emptyParam,D3);
      end;
    Inc(I);
    end;
end;

procedure TGenericWordDriver.SaveDocument(FN: String);

Var
  O : OleVariant;

begin
  O:=Fn;
  FDocument.SaveAs(O);
end;

procedure TGenericWordDriver.InsertTable(Dataset : TDataset; NameTag : String; FieldList : TStrings);

Var
  F,R,T,TR,R2 : OleVariant;
  I : Integer;

begin
  R:=FDocument.Content;
  F:=R.Find;
  F.Text:='{'+NameTag+'}';
  F.ClearFormatting;
  F.MatchWildcards:=False;
  F.MatchCase:=False;
  If F.Execute(emptyParam,emptyParam,emptyParam,emptyParam,emptyparam,
               EmptyParam,EmptyParam,EmptyParam,
               EmptyParam,EmptyParam,EmptyParam) then
    begin
    R.Text:='';
    If FieldList=Nil then
      begin
      FieldList:=TStringList.Create;
      For I:=0 to Dataset.Fields.Count-1 do
        FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
      end
    else
      For I:=0 to FieldList.Count-1 do
        FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);
    T:=R.Tables.Add(R,1,FieldList.Count);
    TR:=T.Rows.Item(1);
    For I:=0 to FieldList.Count-1 do
      begin
      R2:=TR.Cells.Item(I+1).Range;
      R2.Bold:=1;
      R2.Text:=Tfield(FieldList.Objects[i]).DisplayName;
      end;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      TR:=T.Rows.Add(EmptyParam);
      For I:=0 to FieldList.Count-1 do
        begin
        R:=TR.Cells.Item(I+1).Range;
        R.Text:=Tfield(FieldList.Objects[i]).AsString;
        R.Bold:=0;
        end;
      Dataset.Next;
      end;
    end
  else
    Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag])
end;

procedure TGenericWordDriver.InsertTableFromTemplate(Dataset : TDataset);

Var
  R,TR,T,R2 : OleVariant;
  I : Integer;
  FieldList : TStrings;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    FieldList:=TStringList.Create;
    GetTableFieldNames(FieldList,Dataset);
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    Dataset.First;
    While Not Dataset.Eof do
      begin
      For I:=0 to FieldList.Count-1 do
        begin
        R2:=TR.Cells.Item(I+1).Range;
        If Assigned(FieldList.Objects[i]) then
          R2.Text:=Tfield(FieldList.Objects[i]).AsString
        else
          //R2.Text:='';
        end;
      Dataset.Next;
      If Not Dataset.EOF then
        TR:=T.Rows.Add(EmptyParam);
      end;
    end;
end;

procedure TGenericWordDriver.GetTableFieldNames(List: TStrings; Dataset : TDataset);

Var
  R,TR,T : OleVariant;
  I : Integer;
  S : String;
  F : TField;

begin
  R:=FDocument.Content;
  If R.Tables.Count>0 then
    begin
    T:=R.Tables.Item(1);
    TR:=T.Rows.Last;
    For I:=1 to TR.Cells.Count do
      begin
      S:=Trim(TR.Cells.item(i).Range.Text);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

Procedure TGenericWordDriver.AddMergeDoc (FN : String; AddPageBreak : Boolean);

Var
  O1,O2 : OleVAriant;
  R : OleVariant;

begin
  Fn:=StringReplace(FN,'\','\\',[rfReplaceAll]);
  R:=FDocument.Content;
  O1:=wdCollapseEnd;
  R.Collapse(O1);
  O1:=wdFieldIncludeText;
  O2:='"'+FN+'"';
  FDocument.Content.Fields.Add(R,O1,O2,EmptyParam);
  if AddPageBreak then
    begin
    R:=FDocument.Content;
    O1:=wdCollapseEnd;
    R.Collapse(O1);
    O1:=wdPageBreak;
    R.InsertBreak(O1);
    end;
end;


{ TOfficeWriterDriver }

procedure TOfficeWriterDriver.AddMergeDoc(FN: String;
  AddPageBreak: Boolean);

Var
  Txt,TS,Curs : OleVAriant;
  LC,LO : OleVariant;
  S : String;

begin
  Fn:=StringReplace(FN,'\','/',[rfReplaceAll]);
  FN:='file:///'+FN;
  TS:=FDocument.createInstance('com.sun.star.text.TextSection');
  If VarIsNull(FCoreReflection) or VarIsEmpty(FCoreReflection) then
    FCoreReflection:=FOffice.createInstance('com.sun.star.reflection.CoreReflection');
  LC:=FCoreReflection.forName('com.sun.star.text.SectionFileLink');
  LC.createObject(LO);
  LO.FileURL:=FN;
  TS.setPropertyValue('FileLink',LO);
  TS.setName(FN);
  Txt:=FDocument.getText;
  Curs:=Txt.createTextCursor;
  Curs.gotoEnd(False);
  Txt.insertTextContent(Curs,TS,true);
  if AddPageBreak then
    begin
    Curs.gotoEnd(False);
    S:=#13;
    Txt.insertString(Curs,S,false);
    Curs.setPropertyValue('BreakType',4);
    end;
end;

procedure TOfficeWriterDriver.DisconnectDocument(KeepOpen: Boolean);

begin
  If Not KeepOpen then
    FDocument.close(True);
  VarClear(FDocument);
  FHaveDocument:=False;
end;

procedure TOfficeWriterDriver.DisconnectWord(QuitWord: Boolean);

begin
  VarClear(FDesktop);
  VarClear(FOffice)
end;

constructor TOfficeWriterDriver.Create(AOwner: TComponent);
begin
  // Nothing to do
end;

procedure TOfficeWriterDriver.CreatePropertyValue(
  var PropertyValue: Variant; APropertyName : String);
begin
  If VarIsNull(FCoreReflection) or VarIsEmpty(FCoreReflection) then
    FCoreReflection:=FOffice.createInstance('com.sun.star.reflection.CoreReflection');
  FCoreReflection.forName('com.sun.star.beans.PropertyValue').CreateObject(PropertyValue);
  PropertyValue.Name:=APropertyName;
end;

destructor TOfficeWriterDriver.Destroy;
begin

  inherited;
end;

procedure TOfficeWriterDriver.GetFormFieldNames(List: TStrings);
begin
  // Do nothing
end;

procedure TOfficeWriterDriver.GetMergeFieldNames(List: TStrings);

Var
  FindDescriptor : Variant;
  Found : Variant;

  S : String;

begin
  FindDescriptor:=FDocument.createSearchDescriptor;
  FindDescriptor.setSearchString('\{[^\{]*\}');
  FindDescriptor.SearchRegularExpression:=True;
  Found:=FDocument.FindFirst(FindDescriptor);
  While Not (VarIsNull(Found) or VarIsEmpty(Found) or (VarType(Found)=varUnknown)) do
    begin
    S:=Found.getString;
    Delete(S,1,1);
    SetLength(S,Length(S)-1);
    If (Pos(' ',S)=0) and (length(S)<32) then
      List.Add(S);
    Found:=FDocument.FindNext(Found.End,FindDescriptor);
    end;
end;

Function ColName(I : Integer) : String;

begin
  Result:='';
  Repeat
    Result:=Result+Char(Ord('A')+(I mod 26));
    I:=I div 26;
  Until (I=0);
end;


procedure TOfficeWriterDriver.GetTableFieldNames(List: TStrings;
  Dataset: TDataset);

Var
  TC,T : OleVariant;
  I,RCount,CCount : Integer;
  S : String;
  F : TField;

begin
  TC:=FDocument.GetTextTables;
  If TC.getCount>0 then
    begin
    // Get first table.
    T:=TC.getByIndex(0);
    // Get number of rows, columns
    RCount:=T.getRows.getCount;
    CCount:=T.getColumns.getCount;
    // Retrieve names from cells.
    For I:=0 to CCount-1 do
      begin
      S:=Trim(T.getCellByName(ColName(I)+IntToStr(RCount)).getString);
      If Length(S)>0 then
        begin
        If S[1]='{' then
          Delete(S,1,1);
        If Length(S)>0 then
          If S[Length(S)]='}' then
            S:=Copy(S,1,Length(S)-1);
        end;
      If Assigned(Dataset) and (S<>'') then
        F:=Dataset.Fields.FindField(S)
      else
        F:=Nil;
      List.AddObject(S,F);
      end;
    end
  else
    Raise EWordDriver.Create(SErrNoTablesInDocument);
end;

procedure TOfficeWriterDriver.InsertTable(Dataset: TDataset;
  NameTag: String; FieldList: TStrings);

Var
  F,Found : Variant;
  FreeList : Boolean;
  T  : Variant;
  TR : Variant;
  Cell : Variant;
  CurrRow,I : Integer;


begin
  F:=FDocument.createSearchDescriptor;
  F.setSearchString('{'+NameTag+'}');
  Found:=FDocument.FindFirst(F);
  If (VarIsNull(Found) or VarIsEmpty(Found)) then
    Raise EWordDriver.CreateFmt(SErrNoSuchTable,[NameTag]);
//  FDocument.createCursorFromRange(
  // Curs:= initialize with location of tag.
  FreeList:=(FieldList=Nil);
  Try
    If FreeList then
      begin
      FieldList:=TStringList.Create;
      For I:=0 to Dataset.Fields.Count-1 do
        FieldList.AddObject(Dataset.Fields[i].FieldName,Dataset.Fields[i]);
      end
    else
      For I:=0 to FieldList.Count-1 do
        FieldList.Objects[i]:=Dataset.Fields.FindField(Fieldlist[i]);

    T:=FDocument.createInstance( 'com.sun.star.text.TextTable' );
    T.setName(NameTag);
    T.initialize(1,FieldList.Count);
    FDocument.Text.InsertTextContent(Found,T,True);
    For I:=0 to FieldList.Count-1 do
       begin
       Cell:=T.getCellByName(ColName(I)+IntToStr(1));
       // Set bold too.
       Cell.setString(TField(FieldList.Objects[i]).DisplayName);
       end;
     TR:=T.getRows;
     CurrRow:=1;
     Dataset.First;
     While Not Dataset.Eof do
       begin
       Inc(CurrRow);
       TR.insertByIndex(CurrRow,1);
       For I:=0 to FieldList.Count-1 do
         begin
         Cell:=T.getCellByName(ColName(I)+IntToStr(CurrRow));
         Cell.setString(Tfield(FieldList.Objects[i]).AsString);
         end;
       Dataset.Next;
       end;
  Finally
    If FreeList then
      FieldList.Free;
  end;
end;

procedure TOfficeWriterDriver.InsertTableFromTemplate(Dataset: TDataset);

Var
  TC,T,C : OleVariant;
  I,RCount : Integer;
  FieldList : TStrings;

begin
  TC:=FDocument.GetTextTables;
  If TC.getCount>0 then
    begin
    FieldList:=TStringList.Create;
    Try
      GetTableFieldNames(FieldList,Dataset);
      // Get first table.
      T:=TC.getByIndex(0);
      // Get number of rows, columns
      RCount:=T.getRows.getCount;
      Dataset.First;
      While Not Dataset.Eof do
        begin
        For I:=0 to FieldList.Count-1 do
          begin
          C:=T.getCellByName(ColName(I)+IntToStr(RCount));
          If Assigned(FieldList.Objects[i]) then
            C.setString(Tfield(FieldList.Objects[i]).AsString)
          else
            //C.setString('');
          end;
        Dataset.Next;
        If Not Dataset.EOF then
          begin
          Inc(RCount);
          T.getRows.insertByIndex(RCount,1);
          end;
        end;
    finally
      FieldList.Free;
    end;
    end;
end;

procedure TOfficeWriterDriver.OpenDocument(FN: String);

Var
  OpenParams : Variant;

begin
  OpenParams:=VarArrayCreate([0, -1], varVariant);
  If (fn<>'') then
    FN:='file:///'+StringReplace(FN,'\','/',[rfReplaceAll])
  else
    FN:='private:factory/swriter';
  FDocument:=FDesktop.LoadComponentFromURL(FN,'_default',0,OpenParams);
  If (VarIsEmpty(FDocument) or VarIsNull(FDocument)) then
    Raise EWordDriver.Create('Could not open OpenOffice');
  FHaveDocument:=True;
end;

procedure TOfficeWriterDriver.OpenWord(WithVisible: Boolean);
begin
  if  VarIsEmpty(FOffice) then
     FOffice := CreateOleObject('com.sun.star.ServiceManager');
  If (VarIsEmpty(FOffice) or VarIsNull(FOffice)) then
    Raise EWordDriver.Create('Could not open OpenOffice');
  if VarIsEmpty(FDesktop) then
    FDesktop := FOffice.createInstance('com.sun.star.frame.Desktop');
end;

procedure TOfficeWriterDriver.PrintDocument;

Var
  PrintOpts : Variant;

begin
  PrintOpts:=VarArrayCreate([0, -1], varVariant);
  FDocument.Print(PrintOpts);
end;

procedure TOfficeWriterDriver.ReplaceFields(Value: TStrings);
begin
  Raise Exception.Create('Cannot replace fields in OpenOffice');
end;

procedure TOfficeWriterDriver.ReplaceValues(Value: TStrings);

Var
  I,J : Integer;
  R,S : String;
  ReplaceDescriptor : Variant;

begin
  I:=0;
  While (Not FCancelled) and (I<Value.Count) do
    begin
    R:=Value[i];
    J:=Pos('=',R);
    If (J>0) then
      begin
      S:='{'+Copy(R,1,J-1)+'}';
      System.Delete(R,1,J);
      ReplaceDescriptor:=FDocument.createReplaceDescriptor;
      ReplaceDescriptor.setSearchString(S);
      ReplaceDescriptor.setReplaceString(R);
//      ReplaceDescriptor.  // Maybe set case sensitivity.
      FDocument.replaceAll(ReplaceDescriptor);
      end;
    Inc(I);
    end;

end;

procedure TOfficeWriterDriver.SaveDocument(FN: String);

Var
  SaveParams : Variant;

begin
  SaveParams:=VarArrayCreate([0, -1], varVariant);
  FN:='file:///'+StringReplace(FN,'\','/',[rfReplaceAll]);
  FDocument.StoreAsUrl(FN,SaveParams);
end;

end.

