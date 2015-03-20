program WordDemo;

uses
  Forms,
  frmMain in 'frmMain.pas' {MainForm},
  frmWordExport in '..\frmWordExport.pas' {WordExportForm},
  WordDriver in '..\WordDriver.pas',
  WordXP in '..\wordxp\WordXP.pas',
  VBIDEXP in '..\wordxp\VBIDEXP.pas',
  OfficeXP in '..\wordxp\OfficeXP.pas',
  Word97 in '..\Word97\Word97.pas',
  VBIDE97 in '..\Word97\VBIDE97.pas',
  Office97 in '..\Word97\Office97.pas';

{$R *.res}

begin
  Forms.Application.Initialize;
  Forms.Application.CreateForm(TMainForm, MainForm);
  Forms.Application.CreateForm(TWordExportForm, WordExportForm);
  Forms.Application.Run;
end.
