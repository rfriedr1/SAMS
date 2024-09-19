unit _InputSamples;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, JvWizard, JvWizardRouteMapList, JvExControls, StdCtrls,
  JvComponentBase, JvFormPlacement, Mask, JvExMask, JvToolEdit, Grids,
  JvExGrids, JvStringGrid, ExtCtrls, JvExExtCtrls, JvNetscapeSplitter, dblookup,
  ComCtrls, JvExComCtrls, JvDateTimePicker, DBCtrls, DBGrids, JvAppStorage,
  JvAppIniStorage;

type
  TfrmInputNewSamples = class(TForm)
    JvFormStorage1: TJvFormStorage;
    procedure wizListSamplesEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure wizInputProjectEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure btnInsertUserClick(Sender: TObject);
    procedure wizInputUserEnterPage(Sender: TObject;
      const FromPage: TJvWizardCustomPage);
    procedure edtNewSamplesFilenameAfterDialog(Sender: TObject;
      var AName: string; var AAction: Boolean);
    procedure FormShow(Sender: TObject);
    procedure wizInputSamplesCancelButtonClick(Sender: TObject);
    procedure wizListSamplesCancelButtonClick(Sender: TObject;
      var Stop: Boolean);
    procedure wizInsertSamplesBackButtonClick(Sender: TObject;
      var Stop: Boolean);
  private
    ProjectName: string;
  public
    { Public declarations }
  end;

var
  frmInputNewSamples: TfrmInputNewSamples;

implementation

{$R *.dfm}

uses
  SAMS_Main, JvJCLUtils, Clipbrd, ADODB; // wegen FormStorage


end.

