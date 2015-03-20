unit frmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, StdCtrls, Mask, DBCtrls, ExtCtrls, Grids, DBGrids, DB,
  DBClient;

type
  TMainForm = class(TForm)
    CDSAddress: TClientDataSet;
    DSAddress: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    BClose: TButton;
    BExport: TButton;
    ActionList1: TActionList;
    EFirstName: TDBEdit;
    ELastName: TDBEdit;
    ETitle: TDBEdit;
    EAddress1: TDBEdit;
    EAddress2: TDBEdit;
    ECity: TDBEdit;
    EZip: TDBEdit;
    ECountry: TDBEdit;
    ETelephone: TDBEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    BSave: TButton;
    AClose: TAction;
    ASave: TAction;
    AExport: TAction;
    procedure ASaveExecute(Sender: TObject);
    procedure ACloseExecute(Sender: TObject);
    procedure AExportExecute(Sender: TObject);
    procedure AExportUpdate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses frmWordExport;

{$R *.dfm}

procedure TMainForm.ASaveExecute(Sender: TObject);
begin
  CDSAddress.SaveToFile;
end;

procedure TMainForm.ACloseExecute(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.AExportExecute(Sender: TObject);
begin
  With TWordExportForm.Create(Self) do
    try
      Dataset:=CDSAddress;
      ShowModal;
    Finally
      Free;
    end;
end;

procedure TMainForm.AExportUpdate(Sender: TObject);
begin
  (Sender as Taction).Enabled:=CDSAddress.Active;
end;

end.
