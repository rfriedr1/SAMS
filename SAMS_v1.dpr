program SAMS_v1;

uses
  Forms,
  SAMS_Main in 'SAMS_Main.pas' {frmMAMS},
  About in 'About.pas' {AboutBox},
  _dm in '_dm.pas' {dm: TDataModule},
  MdRView in 'MdRView.pas',
  _LogSQL in '_LogSQL.pas' {frmLogSql},
  WordDriver in 'Worddriver\WordDriver.pas',
  Vcl.Themes,
  Vcl.Styles,
  StorageLocations in 'StorageLocations.pas' {FormStorageLocations},
  frmStartScreen in 'frmStartScreen.pas' {frmStart},
  FormDBSearch in 'FormDBSearch.pas' {frmDBSearch},
  frmLogWindow in 'frmLogWindow.pas' {LogWindow},
  FormNewUser in 'FormNewUser.pas' {frmNewUser},
  Frame_Video in 'Frame_Video.pas' {VideoFrame: TFrame},
  FormCamera in 'FormCamera.pas' {CameraWindow},
  VFrames in 'WebCam\GRBLize\VFrames.pas';

{$R *.RES}

begin
  Application.Initialize;
  //create and display StartScreen
  frmStart := TfrmStart.Create(nil); frmStart.Show; frmStart.Update;
  TStyleManager.TrySetStyle('Turquoise Gray');
  Application.Title := 'SAMS';
  Application.CreateForm(TfrmMAMS, frmMAMS);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(Tdm, dm);
  Application.CreateForm(TfrmLogSql, frmLogSql);
  Application.CreateForm(TFormStorageLocations, FormStorageLocations);
  Application.CreateForm(TfrmDBSearch, frmDBSearch);
  Application.CreateForm(TLogWindow, LogWindow);
  Application.CreateForm(TfrmNewUser, frmNewUser);
  Application.CreateForm(TCameraWindow, CameraWindow);
  //  Application.CreateForm(TfrmInputNewSamples, frmInputNewSamples);
  // closing the StartScreen happens in OnShow of the frmMAMS as soon as everything has been loaded
   // frmStart.Hide; frmStart.Free;
  Application.Run;
end.

