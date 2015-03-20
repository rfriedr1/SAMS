program GetOxasBlanks;

uses
  Forms,
  GetOxasBlanksMain in 'GetOxasBlanksMain.pas' {frmGetOxasBlanks};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmGetOxasBlanks, frmGetOxasBlanks);
  Application.Run;
end.
