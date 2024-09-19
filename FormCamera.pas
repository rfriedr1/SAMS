// this form takes it's content partially from the file "Frame_Video"
// and serves as a container for the frame
// historically it was intented to use multiple frames at the same time
// in order to see multiple webcams at the same time.

unit FormCamera;

interface

uses
  Windows, Classes, Controls, Forms, ExtCtrls, Frame_Video, Menus,
  StdCtrls, Vcl.Mask, JvExMask, JvToolEdit, JvComponentBase, IniFiles, SysUtils,
  Dialogs, jpeg, Graphics, WCamera, Vcl.Samples.Spin, Vcl.ComCtrls;


type
  TCameraWindow = class(TForm)
    PanelLeft: TPanel;
    PanelRight: TPanel;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Quit1: TMenuItem;
    edtMAMS: TEdit;
    Label1: TLabel;
    edtPathToImage: TJvDirectoryEdit;
    Label2: TLabel;
    PaintBoxImage: TPaintBox;
    btnSnapImage: TButton;
    btnSave: TButton;
    GroupBox1: TGroupBox;
    lbl_SaveSuccessfull: TLabel;
    WCamera: TWCamera;
    ComboBoxCamera: TComboBox;
    PaintBox: TPaintBox;
    GroupBoxCameraSettings: TGroupBox;
    ComboBoxFormat: TComboBox;
    PanelPaintBox: TPanel;
    SpinButton: TSpinButton;
    SaveDialog: TSaveDialog;
    PanelCenter: TPanel;
    GroupBoxImageSettings: TGroupBox;
    LabelWhiteBalance: TLabel;
    TrackBarWhiteBalance: TTrackBar;
    CheckBoxWhiteBalance: TCheckBox;
    LabelSharpness: TLabel;
    TrackBarSharpness: TTrackBar;
    CheckBoxSharpness: TCheckBox;
    TrackBarFocus: TTrackBar;
    CheckBoxFocus: TCheckBox;
    LabelFocus: TLabel;
    LabelExposure: TLabel;
    TrackBarExposure: TTrackBar;
    CheckBoxExposure: TCheckBox;
    TrackBarContrast: TTrackBar;
    CheckBoxContrast: TCheckBox;
    LabelContrast: TLabel;
    LabelBrightness: TLabel;
    TrackBarBrightness: TTrackBar;
    CheckBoxBrightness: TCheckBox;
    ButtonDefault: TButton;
    Splitter1: TSplitter;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Quit1Click(Sender: TObject);
    procedure Splitter1Moved(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure btnSnapImageClick(Sender: TObject);
    procedure edtMAMSChange(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure ComboBoxCameraDropDown(Sender: TObject);
    procedure ComboBoxCameraChange(Sender: TObject);
    procedure ComboBoxFormatChange(Sender: TObject);
    procedure PaintBoxPaint(Sender: TObject);
    procedure WCameraImageAvailable(Sender: TObject; SampleTime: Double);
    procedure PanelPaintBoxResize(Sender: TObject);
    procedure PaintBoxImagePaint(Sender: TObject);
    procedure SpinButtonDownClick(Sender: TObject);
    procedure SpinButtonUpClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CheckBoxClick(Sender: TObject);
    procedure TrackBarChange(Sender: TObject);
    procedure ButtonDefaultClick(Sender: TObject);
    procedure edtPathToImageChange(Sender: TObject);
  private
    { Private declarations }
    SplitterRatio : double;
    MemoryStream: TMemoryStream;
    Bitmap: TBitmap;
    BitmapSnappedImage: TBitmap;
    Format: TWCameraFormat;
    procedure SetDevices;
    procedure SetFormats;
    //procedure SetTrackBars;
    procedure StartDevice;
    procedure UpdateVideoPosition;
    procedure SetTrackBars;
  public
    { Public declarations }
  end;


var
  CameraWindow: TCameraWindow;
  myIni : TIniFile;


implementation

{$R *.dfm}



procedure TCameraWindow.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  Screen.Cursor := crHourGlass;
  //ComboBoxCamera.Clear;
  //ComboBoxFormat.Clear;
  //WCamera.Stop;
  //WCamera.Active := False;
  //WCamera.Free;
  //Bitmap.Free;
  //BitmapSnappedImage.Free;
  //MemoryStream.Free;
  Screen.Cursor := crdefault;
end;



procedure TCameraWindow.btnSaveClick(Sender: TObject);
Var
  Bmp: TBitmap;
  jpgImage: TJPEGImage;
begin
 // create bmp from PaintBox
 //Bmp := TBitmap.Create();
 //Bmp.SetSize(PaintBoxImage.Canvas.ClipRect.Right, PaintBoxImage.Canvas.ClipRect.Bottom);
 //BitBlt(Bmp.Canvas.Handle, 0, 0, PaintBoxImage.Width, PaintBoxImage.Height, PaintBoxImage.Canvas.Handle, 0, 0, SRCCOPY);
 //BitBlt(Bmp.Canvas.Handle, 0, 0, Width, Height, Canvas.Handle, 0, 0, SRCCOPY);

 // create SaveDialog
 saveDialog.FileName := edtPathToImage.Text;

 // create IMage first
 jpgImage := TJPEGImage.Create();
 jpgImage.Assign(BitmapSnappedImage);

 // save jpeg
 // showmessage(extractfilepath(edtPathToImage.Text));
 // check that directory really exists
 if directoryexists(extractfilepath(edtPathToImage.Text)) then
  begin
   //ShowMessage('Directory =' + extractfilepath(edtPathToImage.Text));
   if fileexists(saveDialog.FileName) then
    begin
        saveDialog.Title := 'File exists already';
        if saveDialog.Execute then
          begin
          jpgImage.SaveToFile(edtPathToImage.Text);
          lbl_SaveSuccessfull.Caption := 'Saved';
          lbl_SaveSuccessfull.Visible := True;
          end
          else
          begin
          lbl_SaveSuccessfull.Caption := 'not Saved';
          lbl_SaveSuccessfull.Visible := False;
          end;
    end
    else
    begin
        jpgImage.SaveToFile(edtPathToImage.Text);
        lbl_SaveSuccessfull.Caption := 'Saved';
        lbl_SaveSuccessfull.Visible := True;
    end;
  end
 else
   begin
   lbl_SaveSuccessfull.Caption := 'Not Saved';
   lbl_SaveSuccessfull.Visible := True;
   showmessage('error saving file. NetworkDrive connected?');
   end;


end;

procedure TCameraWindow.btnSnapImageClick(Sender: TObject);
Var
  AspectRatio: Double;
  Height, Width: Integer;
  Rect: TRect;
begin
  // save image from camera to Bitmap in Memory
  Try
    BitmapSnappedImage.FreeImage;
    BitmapSnappedImage:=WCamera.CurrentImageToBitmap();
    if BitmapSnappedImage.Width / BitmapSnappedImage.Height >= PaintBoxImage.Width / PaintBoxImage.Height then
      begin
        AspectRatio := PaintBoxImage.Width / BitmapSnappedImage.Width;
        Height := Round(AspectRatio * BitmapSnappedImage.Height);
        Rect.Left := 0;
        Rect.Top := (PaintBoxImage.Height - Height) div 2;
        Rect.Right := PaintBoxImage.Width;
        Rect.Bottom := Rect.Top + Height;
    end
    else
      begin
        AspectRatio := PaintBoxImage.Height / BitmapSnappedImage.Height;
        Width := Round(AspectRatio * BitmapSnappedImage.Width);
        Rect.Left := (PaintBoxImage.Width - Width) div 2;
        Rect.Top := 0;
        Rect.Right := Rect.Left + Width;
        Rect.Bottom := PaintBoxImage.Height;
      end;
    PaintBoxImage.Canvas.StretchDraw(Rect,BitmapSnappedImage);
    btnSave.Enabled := True;
  Except
      btnSave.Enabled := False;
      ShowMessage('Error occured capturing the image from camera.');
  End;

  lbl_SaveSuccessfull.Caption := 'Not Saved';
  lbl_SaveSuccessfull.Visible := True;
end;

procedure TCameraWindow.ButtonDefaultClick(Sender: TObject);
begin
  try
    if WCamera.BacklightCompensationSupported then
      WCamera.BacklightCompensation := WCamera.BacklightCompensationRange.Default;

    if WCamera.BrightnessSupported then
      WCamera.Brightness := WCamera.BrightnessRange.Default;

    if WCamera.ColorEnableSupported then
      WCamera.ColorEnable := WCamera.ColorEnableRange.Default;

    if WCamera.ContrastSupported then
      WCamera.Contrast := WCamera.ContrastRange.Default;

    if WCamera.ExposureSupported then
      WCamera.Exposure := WCamera.ExposureRange.Default;

    if WCamera.FocusSupported then
      WCamera.Focus := WCamera.FocusRange.Default;

    if WCamera.GainSupported then
      WCamera.Gain := WCamera.GainRange.Default;

    if WCamera.GammaSupported then
      WCamera.Gamma := WCamera.GammaRange.Default;

    if WCamera.HueSupported then
      WCamera.Hue := WCamera.HueRange.Default;

    if WCamera.IrisSupported then
      WCamera.Iris := WCamera.IrisRange.Default;

    if WCamera.PanSupported then
      WCamera.Pan := WCamera.PanRange.Default;

    if WCamera.RollSupported then
      WCamera.Roll := WCamera.RollRange.Default;

    if WCamera.SaturationSupported then
      WCamera.Saturation := WCamera.SaturationRange.Default;

    if WCamera.SharpnessSupported then
      WCamera.Sharpness := WCamera.SharpnessRange.Default;

    if WCamera.TiltSupported then
      WCamera.Tilt := WCamera.TiltRange.Default;

    if WCamera.WhiteBalanceSupported then
      WCamera.WhiteBalance := WCamera.WhiteBalanceRange.Default;

    if WCamera.ZoomSupported then
      WCamera.Zoom := WCamera.ZoomRange.Default;
  finally
    SetTrackBars;
  end
end;

procedure TCameraWindow.ComboBoxCameraChange(Sender: TObject);
begin
  SetFormats;
end;

procedure TCameraWindow.SetFormats;
var
  Formats: TWCameraFormats;
  I: Integer;
begin
  Formats := nil;
  ComboBoxFormat.Items.BeginUpdate;
  try
    btnSnapImage.Enabled := False;
    WCamera.Active := False;
    WCamera.DeviceName := ComboBoxCamera.Text;
    //ShowMessage('SelectedDevice = ' + ComboBoxCamera.Text);
    ComboBoxFormat.Items.Clear;
    ComboBoxFormat.Enabled := WCamera.DeviceName <> '';
    if ComboBoxFormat.Enabled then
    begin
      Formats := WCamera.SupportedFormats;
      for I := 0 to Length(Formats) - 1 do
        if Formats[I].AvgTimePerFrame = 0 then
          ComboBoxFormat.Items.Add(IntToStr(Formats[I].Width) + 'x' + IntToStr(Formats[I].Height) + ' ' + IntToStr(Formats[I].BitsPerPixel) + 'bpp')
        else
          ComboBoxFormat.Items.Add(IntToStr(Formats[I].Width) + 'x' + IntToStr(Formats[I].Height) + ' ' + IntToStr(10000000 div Formats[I].AvgTimePerFrame) + 'Hz ' + IntToStr(Formats[I].BitsPerPixel) + 'bpp');
    end
  finally
    ComboBoxFormat.Items.EndUpdate;
  end;
end;

procedure TCameraWindow.ComboBoxCameraDropDown(Sender: TObject);
begin
 SetDevices;
end;

procedure TCameraWindow.ComboBoxFormatChange(Sender: TObject);
begin
  StartDevice;
end;

procedure TCameraWindow.StartDevice;
begin
  //SetTrackBars;
  //set all to Auto (no manual adjustment so far. that would be done using the trackbars)
  // checkout demo program of winsofts camera add on
     // WCamera.BacklightCompensationAuto := True;
     // WCamera.BrightnessAuto := True;
     // WCamera.ColorEnableAuto := True;
     // WCamera.ContrastAuto := True;
     // WCamera.ExposureAuto := True;
     // WCamera.FocusAuto := True;
     // WCamera.GainAuto := True;
     // WCamera.GammaAuto := True;
     // WCamera.HueAuto := True;
     // WCamera.IrisAuto := True;
     // WCamera.PanAuto := True;
     // WCamera.RollAuto := True;
     // WCamera.SaturationAuto := True;
     // WCamera.SharpnessAuto := True;
     // WCamera.TiltAuto := True;
     // WCamera.WhiteBalanceAuto := True;
     // WCamera.ZoomAuto := True;

  if ComboBoxFormat.ItemIndex <> -1 then
  begin
    WCamera.Active := False;
    Format := WCamera.SupportedFormats[ComboBoxFormat.ItemIndex];
    WCamera.Format := Format;
    WCamera.Active := True;
    UpdateVideoPosition;
    WCamera.Run;
    PaintBox.Visible := WCamera.CaptureType = ctGrabber;
    btnSnapImage.Enabled := True;
    btnSnapImage.SetFocus;
  end;
end;

procedure TCameraWindow.UpdateVideoPosition;
var Rect: TRect;
begin
  if WCamera.Active and (WCamera.CaptureType = ctVmr9) then
  begin
    if true then       // old condition: CheckBoxStretch.Checked
    begin
      WCamera.AspectRatio := True;  //old value: CheckBoxAspectRatio.Checked;
      WCamera.VideoPositionDest := PanelPaintBox.ClientRect;
    end
    else
    begin
      Rect.Left := (PanelPaintBox.Width - Format.Width) div 2;
      Rect.Top := (PanelPaintBox.Height - Format.Height) div 2;
      Rect.Right := Rect.Left + Format.Width;
      Rect.Bottom := Rect.Top + Format.Height;
      WCamera.VideoPositionDest := Rect;
    end;
    PanelPaintBox.Invalidate;
  end;
end;

procedure TCameraWindow.WCameraImageAvailable(Sender: TObject; SampleTime: Double);
begin
  PaintBox.Invalidate;
end;

procedure TCameraWindow.SetDevices;
var
  CurrentDeviceName: string;
  Devices: TWVideoCaptureDevices;
  I, Index: Integer;
begin
  ComboBoxCamera.Items.BeginUpdate;
  try
    CurrentDeviceName := ComboBoxCamera.Text;
    Devices := WCamera.Devices;
    ComboBoxCamera.Items.Clear;
    for I := 0 to Length(Devices) - 1 do
      ComboBoxCamera.Items.Add(Devices[I].Name);

    Index := ComboBoxCamera.Items.IndexOf(CurrentDeviceName);
    if Index <> -1 then
      ComboBoxCamera.ItemIndex := Index
  finally
    ComboBoxCamera.Items.EndUpdate;
  end;
end;

procedure TCameraWindow.edtMAMSChange(Sender: TObject);
begin
 // update the path
 edtPathToImage.Text := myIni.ReadString('Win2kAppForm', 'JvDirEdt_Server_Image_Path_Text', '') + '\'  + edtMAMS.Text + '.jpg'
end;

procedure TCameraWindow.edtPathToImageChange(Sender: TObject);
begin
 // save new path to ini file
  //showmessage(edtPathToImage.Directory);
end;

procedure TCameraWindow.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //ComboBoxCamera.Clear;
  //ComboBoxFormat.Clear;
  if WCamera.State = csRunning then WCamera.Stop;
  WCamera.Active := False;
  //WCamera.Free;
  //Bitmap.Free;
  //BitmapSnappedImage.Free;
  //MemoryStream.Free;
  //PaintBox.Free;
  //PaintBoxImage.Free;
  //CameraWindow.Release;
end;

procedure TCameraWindow.FormShow(Sender: TObject);
Var
pathToIni : string;
begin
  // some settings
  lbl_SaveSuccessfull.Visible := False;
  btnSnapImage.Enabled := False;
  btnSave.Enabled := False;

  // create iniFile Object and connect to the ini file that SAMS uses
  // in order to read the path settings to the Server where images are stored
  pathToIni := extractFilePath(paramstr(0)) + 'Persistent.ini';
  myIni := TiniFile.Create(pathToIni);
  if fileexists(pathToIni) then
    Begin
      if myIni.ValueExists('Win2kAppForm', 'JvDirEdt_Server_Image_Path_Text') then
      Begin
        //showmessage(pathToIni);
        //showmessage(myIni.ReadString('Win2kAppForm', 'JvDirEdt_Server_Image_Path_Text', 'test'));
        edtPathToImage.Text := myIni.ReadString('Win2kAppForm', 'JvDirEdt_Server_Image_Path_Text', '');
        // alreaday set the path to a MAMS
        edtPathToImage.Text := edtPathToImage.Text + '\' + edtMAMS.Text;
      End
      Else
      Begin
       edtPathToImage.Text := 'path not set in SAMS!!';
       btnSave.Enabled := False;
      End;
    End;

  //check if camera has already been selected.
  if not (WCamera.DeviceName = '')  then
  begin
    if WCamera.Format.Width <>0 then
    begin
    StartDevice;
    end;
  end;


end;

procedure TCameraWindow.PaintBoxImagePaint(Sender: TObject);
Var
  AspectRatio: Double;
  Height, Width: Integer;
  Rect: TRect;
begin

 if BitmapSnappedImage.Width = 0 then
      Exit;

  if BitmapSnappedImage.Width / BitmapSnappedImage.Height >= PaintBoxImage.Width / PaintBoxImage.Height then
  begin
    AspectRatio := PaintBoxImage.Width / BitmapSnappedImage.Width;
    Height := Round(AspectRatio * BitmapSnappedImage.Height);
    Rect.Left := 0;
    Rect.Top := (PaintBoxImage.Height - Height) div 2;
    Rect.Right := PaintBoxImage.Width;
    Rect.Bottom := Rect.Top + Height;
  end
  else
  begin
    AspectRatio := PaintBoxImage.Height / BitmapSnappedImage.Height;
    Width := Round(AspectRatio * BitmapSnappedImage.Width);
    Rect.Left := (PaintBoxImage.Width - Width) div 2;
    Rect.Top := 0;
    Rect.Right := Rect.Left + Width;
    Rect.Bottom := PaintBoxImage.Height;
  end;
  PaintBoxImage.Canvas.StretchDraw(Rect, BitmapSnappedImage);

end;

procedure TCameraWindow.PaintBoxPaint(Sender: TObject);
var
  AspectRatio: Double;
  Height, Width: Integer;
  Rect: TRect;
begin
  if WCamera.Active and (WCamera.CaptureType = ctGrabber) then
  begin
    MemoryStream.Position := 0;
    WCamera.CurrentImageToStream(MemoryStream);
    MemoryStream.Size := MemoryStream.Position;
    MemoryStream.Position := 0;
    Bitmap.LoadFromStream(MemoryStream);

    if Bitmap.Width = 0 then
      Exit;

    if True then  // old condition: CheckBoxStretch.Checked
    begin
      if True then   // old condition: CheckBoxAspectRatio.Checked
      begin
        if Bitmap.Width / Bitmap.Height >= PaintBox.Width / PaintBox.Height then
        begin
          AspectRatio := PaintBox.Width / Bitmap.Width;
          Height := Round(AspectRatio * Bitmap.Height);
          Rect.Left := 0;
          Rect.Top := (PaintBox.Height - Height) div 2;
          Rect.Right := PaintBox.Width;
          Rect.Bottom := Rect.Top + Height;
        end
        else
        begin
          AspectRatio := PaintBox.Height / Bitmap.Height;
          Width := Round(AspectRatio * Bitmap.Width);
          Rect.Left := (PaintBox.Width - Width) div 2;
          Rect.Top := 0;
          Rect.Right := Rect.Left + Width;
          Rect.Bottom := PaintBox.Height;
        end;

        PaintBox.Canvas.StretchDraw(Rect,  Bitmap);
      end
      else
        Rect := PaintBox.ClientRect;

      PaintBox.Canvas.StretchDraw(Rect, Bitmap);
    end
    else
    begin
      PaintBox.Canvas.Draw(
        (PaintBox.Width - Bitmap.Width) div 2,
        (PaintBox.Height - Bitmap.Height) div 2,
        Bitmap);
    end;
  end;
end;


procedure TCameraWindow.PanelPaintBoxResize(Sender: TObject);
begin
  UpdateVideoPosition;
end;

procedure TCameraWindow.Quit1Click(Sender: TObject);
begin
  close;
end;

procedure TCameraWindow.SpinButtonDownClick(Sender: TObject);
begin
 edtMAMS.Text := inttostr(strtoint(edtMAMS.Text)-1);
end;

procedure TCameraWindow.SpinButtonUpClick(Sender: TObject);
begin
 edtMAMS.Text := inttostr(strtoint(edtMAMS.Text)+1);
end;

procedure TCameraWindow.Splitter1Moved(Sender: TObject);
begin
  SplitterRatio := (PanelLeft.Width+Splitter1.Width div 2) / Width;
end;

procedure TCameraWindow.FormCreate(Sender: TObject);
begin
  PanelLeft.DoubleBuffered := True;
  MemoryStream := TMemoryStream.Create;
  Bitmap := TBitmap.Create;
  BitmapSnappedImage := TBitmap.Create;
  SplitterRatio := 0.4;
  SetTrackBars;
end;

procedure TCameraWindow.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

// wathc up down key to change mams number
// make sure that KeyPreview = true

 if Key = vk_DOWN then
  begin
    // showmessage('down');
    edtMams.Text := intToStr(StrToInt(edtMams.Text)-1);
    btnSnapImage.SetFocus;
  end;
 if Key = vk_UP then
  begin
    //showmessage('up');
    edtMams.Text := intToStr(StrToInt(edtMams.Text)+1);
    btnSnapImage.SetFocus;
  end;

end;

procedure TCameraWindow.FormResize(Sender: TObject);
begin
  PanelLeft.Width := round(SplitterRatio * (Width-Splitter1.Width div 2));
end;

procedure InitTrackBar(TrackbarLabel: TLabel; TrackBar: TTrackBar; CheckBoxAuto: TCheckBox);
begin
  TrackbarLabel.Enabled := False;
  TrackBar.Enabled := False;
  CheckBoxAuto.Enabled := False;
end;

procedure SetTrackBar(TrackbarLabel: TLabel; TrackBar: TTrackBar; CheckBoxAuto: TCheckBox; Position: Integer; Range: TWRange; Auto: Boolean);
begin
  CheckBoxAuto.Checked := Auto;
  TrackBar.Min := Range.Min;
  TrackBar.Max := Range.Max;
  TrackBar.PageSize := Range.Delta;
  if (Range.Max - Range.Min) div Range.Delta > 25 then
    TrackBar.Frequency := (Range.Max - Range.Min) div 25
  else
    TrackBar.Frequency := 1;
  TrackBar.Position := Position;

  TrackbarLabel.Enabled := True;
  CheckBoxAuto.Enabled := Range.Auto and Range.Manual;
  TrackBar.Enabled := not CheckBoxAuto.Checked;
end;

var TrackBarsSetting: Boolean;

procedure TCameraWindow.SetTrackBars;
begin
  if not TrackBarsSetting then
  try
    TrackBarsSetting := True;
    //InitTrackBar(LabelBacklightCompensation, TrackBarBacklightCompensation, CheckBoxBacklightCompensation);
    InitTrackBar(LabelBrightness, TrackBarBrightness, CheckBoxBrightness);
    //InitTrackBar(LabelColorEnable, TrackBarColorEnable, CheckBoxColorEnable);
    InitTrackBar(LabelContrast, TrackBarContrast, CheckBoxContrast);
    InitTrackBar(LabelExposure, TrackBarExposure, CheckBoxExposure);
    InitTrackBar(LabelFocus, TrackBarFocus, CheckBoxFocus);
    //InitTrackBar(LabelGain, TrackBarGain, CheckBoxGain);
    //InitTrackBar(LabelGamma, TrackBarGamma, CheckBoxGamma);
    //InitTrackBar(LabelHue, TrackBarHue, CheckBoxHue);
    //InitTrackBar(LabelIris, TrackBarIris, CheckBoxIris);
    //InitTrackBar(LabelPan, TrackBarPan, CheckBoxPan);
    //InitTrackBar(LabelRoll, TrackBarRoll, CheckBoxRoll);
    //InitTrackBar(LabelSaturation, TrackBarSaturation, CheckBoxSaturation);
    InitTrackBar(LabelSharpness, TrackBarSharpness, CheckBoxSharpness);
    //InitTrackBar(LabelTilt, TrackBarTilt, CheckBoxTilt);
    InitTrackBar(LabelWhiteBalance, TrackBarWhiteBalance, CheckBoxWhiteBalance);
    //InitTrackBar(LabelZoom, TrackBarZoom, CheckBoxZoom);

    if WCamera.DeviceName <> '' then
    begin
      //if WCamera.BacklightCompensationSupported then
      //  SetTrackBar(LabelBacklightCompensation, TrackBarBacklightCompensation, CheckBoxBacklightCompensation, WCamera.BacklightCompensation, WCamera.BacklightCompensationRange, WCamera.BacklightCompensationAuto);

      if WCamera.BrightnessSupported then
        SetTrackBar(LabelBrightness, TrackBarBrightness, CheckBoxBrightness, WCamera.Brightness, WCamera.BrightnessRange, WCamera.BrightnessAuto);

      //if WCamera.ColorEnableSupported then
      //  SetTrackBar(LabelColorEnable, TrackBarColorEnable, CheckBoxColorEnable, WCamera.ColorEnable, WCamera.ColorEnableRange, WCamera.ColorEnableAuto);

      if WCamera.ContrastSupported then
        SetTrackBar(LabelContrast, TrackBarContrast, CheckBoxContrast, WCamera.Contrast, WCamera.ContrastRange, WCamera.ContrastAuto);

      if WCamera.ExposureSupported then
        SetTrackBar(LabelExposure, TrackBarExposure, CheckBoxExposure, WCamera.Exposure, WCamera.ExposureRange, WCamera.ExposureAuto);

      if WCamera.FocusSupported then
        SetTrackBar(LabelFocus, TrackBarFocus, CheckBoxFocus, WCamera.Focus, WCamera.FocusRange, WCamera.FocusAuto);

      //if WCamera.GainSupported then
      //  SetTrackBar(LabelGain, TrackBarGain, CheckBoxGain, WCamera.Gain, WCamera.GainRange, WCamera.GainAuto);

      //if WCamera.GammaSupported then
      //  SetTrackBar(LabelGamma, TrackBarGamma, CheckBoxGamma, WCamera.Gamma, WCamera.GammaRange, WCamera.GammaAuto);

      //if WCamera.HueSupported then
      //  SetTrackBar(LabelHue, TrackBarHue, CheckBoxHue, WCamera.Hue, WCamera.HueRange, WCamera.HueAuto);

      //if WCamera.IrisSupported then
      //  SetTrackBar(LabelIris, TrackBarIris, CheckBoxIris, WCamera.Iris, WCamera.IrisRange, WCamera.IrisAuto);

      //if WCamera.PanSupported then
      //  SetTrackBar(LabelPan, TrackBarPan, CheckBoxPan, WCamera.Pan, WCamera.PanRange, WCamera.PanAuto);

      //if WCamera.RollSupported then
      //  SetTrackBar(LabelRoll, TrackBarRoll, CheckBoxRoll, WCamera.Roll, WCamera.RollRange, WCamera.RollAuto);

      //if WCamera.SaturationSupported then
      //  SetTrackBar(LabelSaturation, TrackBarSaturation, CheckBoxSaturation, WCamera.Saturation, WCamera.SaturationRange, WCamera.SaturationAuto);

      if WCamera.SharpnessSupported then
        SetTrackBar(LabelSharpness, TrackBarSharpness, CheckBoxSharpness, WCamera.Sharpness, WCamera.SharpnessRange, WCamera.SharpnessAuto);

      //if WCamera.TiltSupported then
      //  SetTrackBar(LabelTilt, TrackBarTilt, CheckBoxTilt, WCamera.Tilt, WCamera.TiltRange, WCamera.TiltAuto);

      if WCamera.WhiteBalanceSupported then
        SetTrackBar(LabelWhiteBalance, TrackBarWhiteBalance, CheckBoxWhiteBalance, WCamera.WhiteBalance, WCamera.WhiteBalanceRange, WCamera.WhiteBalanceAuto);

      //if WCamera.ZoomSupported then
      //  SetTrackBar(LabelZoom, TrackBarZoom, CheckBoxZoom, WCamera.Zoom, WCamera.ZoomRange, WCamera.ZoomAuto);
    end
  finally
    TrackBarsSetting := False;
  end;
end;

procedure TCameraWindow.CheckBoxClick(Sender: TObject);
begin
  if not TrackBarsSetting then
  begin
    //if Sender = CheckBoxBacklightCompensation then
    //  WCamera.BacklightCompensationAuto := CheckBoxBacklightCompensation.Checked
    if Sender = CheckBoxBrightness then
      WCamera.BrightnessAuto := CheckBoxBrightness.Checked
    //else if Sender = CheckBoxColorEnable then
    //  WCamera.ColorEnableAuto := CheckBoxColorEnable.Checked
    else if Sender = CheckBoxContrast then
      WCamera.ContrastAuto := CheckBoxContrast.Checked
    else if Sender = CheckBoxExposure then
      WCamera.ExposureAuto := CheckBoxExposure.Checked
    else if Sender = CheckBoxFocus then
      WCamera.FocusAuto := CheckBoxFocus.Checked
    //else if Sender = CheckBoxGain then
    //  WCamera.GainAuto := CheckBoxGain.Checked
    //else if Sender = CheckBoxGamma then
    //  WCamera.GammaAuto := CheckBoxGamma.Checked
    //else if Sender = CheckBoxHue then
    //  WCamera.HueAuto := CheckBoxHue.Checked
    //else if Sender = CheckBoxIris then
    //  WCamera.IrisAuto := CheckBoxIris.Checked
    //else if Sender = CheckBoxPan then
    //  WCamera.PanAuto := CheckBoxPan.Checked
    //else if Sender = CheckBoxRoll then
    //  WCamera.RollAuto := CheckBoxRoll.Checked
    //else if Sender = CheckBoxSaturation then
    //  WCamera.SaturationAuto := CheckBoxSaturation.Checked
    else if Sender = CheckBoxSharpness then
      WCamera.SharpnessAuto := CheckBoxSharpness.Checked
    //else if Sender = CheckBoxTilt then
    //  WCamera.TiltAuto := CheckBoxTilt.Checked
    else if Sender = CheckBoxWhiteBalance then
      WCamera.WhiteBalanceAuto := CheckBoxWhiteBalance.Checked;
    //else if Sender = CheckBoxZoom then
    //  WCamera.ZoomAuto := CheckBoxZoom.Checked;
    SetTrackBars;
  end
end;

procedure TCameraWindow.TrackBarChange(Sender: TObject);
begin
  if not TrackBarsSetting then
    //if Sender = TrackBarBacklightCompensation then
    //  WCamera.BacklightCompensation := TrackBarBacklightCompensation.Position
    if Sender = TrackBarBrightness then
      WCamera.Brightness := TrackBarBrightness.Position
    //else if Sender = TrackBarColorEnable then
    //  WCamera.ColorEnable := TrackBarColorEnable.Position
    else if Sender = TrackBarContrast then
      WCamera.Contrast := TrackBarContrast.Position
    else if Sender = TrackBarExposure then
      WCamera.Exposure := TrackBarExposure.Position
    else if Sender = TrackBarFocus then
      WCamera.Focus := TrackBarFocus.Position
    //else if Sender = TrackBarGain then
    //  WCamera.Gain := TrackBarGain.Position
    //else if Sender = TrackBarGamma then
    //  WCamera.Gamma := TrackBarGamma.Position
    //else if Sender = TrackBarHue then
    //  WCamera.Hue := TrackBarHue.Position
    //else if Sender = TrackBarIris then
    //  WCamera.Iris := TrackBarIris.Position
    //else if Sender = TrackBarPan then
    //  WCamera.Pan := TrackBarPan.Position
    //else if Sender = TrackBarRoll then
    //  WCamera.Roll := TrackBarRoll.Position
    //else if Sender = TrackBarSaturation then
    //  WCamera.Saturation := TrackBarSaturation.Position
    else if Sender = TrackBarSharpness then
      WCamera.Sharpness := TrackBarSharpness.Position
    //else if Sender = TrackBarTilt then
    //  WCamera.Tilt := TrackBarTilt.Position
    else if Sender = TrackBarWhiteBalance then
      WCamera.WhiteBalance := TrackBarWhiteBalance.Position;
    //else if Sender = TrackBarZoom then
    //  WCamera.Zoom := TrackBarZoom.Position;
end;

end.
