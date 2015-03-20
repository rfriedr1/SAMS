unit GetOxasBlanksMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ADODB, ExtCtrls;

type
  TfrmGetOxasBlanks = class(TForm)
    ADOConnKTL: TADOConnection;
    qryDb: TADOQuery;
    lblBlank: TLabel;
    lblOxas: TLabel;
    Button1: TButton;
    Label3: TLabel;
    Timer: TTimer;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
  private
    procedure GetOxasBlanks(Oxa : boolean);
  public
    { Public declarations }
  end;

var
  frmGetOxasBlanks: TfrmGetOxasBlanks;

implementation

{$R *.dfm}

{ TForm5 }

procedure TfrmGetOxasBlanks.Button1Click(Sender: TObject);
begin
  GetOxasBlanks(true);
  if qryDB.RecordCount<6 then  lblOxas.Color := clYellow;
  if qryDB.RecordCount<3 then  lblOxas.Color := clRed;
  lblOxas.Caption := ' Oxas=' + IntToStr(qryDB.RecordCount);
  if qryDB.RecordCount<3 then  frmGetOxasBlanks.Color := clYellow;
  GetOxasBlanks(false);
  if qryDB.RecordCount<6 then  lblBlank.Color := clYellow;
  if qryDB.RecordCount<3 then  lblBlank.Color := clRed;
  lblBlank.Caption := ' Blanks=' + IntToStr(qryDB.RecordCount);
end;

procedure TfrmGetOxasBlanks.FormShow(Sender: TObject);
var
  s : string;
begin
  s := 'Provider=MSDASQL.1';
  // params are : Data Source=...(ODBC in control panel); User ID = ...; Password=....
  adoConnKTL.LoginPrompt := false;
  with adoConnKTL do begin
    if not Connected then begin
      if ParamCount > 0 then begin
     //Provider=MSDASQL.1;Password=Micadas;Persist Security Info=True;User ID=root;Data Source=DMYSQL_KTL
        s := s + ';Data Source =' + ParamStr(1);
        if ParamCount > 1 then s := s + ';User ID=' + ParamStr(2);
        if ParamCount > 2 then begin
          s := s + '; Password :=' + ParamStr(3);
          LoginPrompt := false;
        end;
        ConnectionString := s;
//        frmGetOxasBlanks.Caption := ' GetOxasBlanks';
        Label3.Caption := ' ODBC:' + ParamStr(1) + ' User:' + ParamStr(2);
      end
      else begin
      end;
    end;
    Open;
  end;
  Button1Click(self);
end;

procedure TfrmGetOxasBlanks.GetOxasBlanks(Oxa: boolean);
var
  s : string;
begin
  if Oxa then
    s := '"oxa%"'
  else
    s :=  '"blank%"';
  with qryDB do begin
    SQL.Text := ' SELECT target_t.sample_nr, sample_t.user_label FROM target_t ' +
                ' INNER JOIN sample_t ON sample_t.sample_nr=target_t.sample_nr ' +
                ' WHERE magazine IS NULL AND target_pressed IS NOT NULL and stop=0 ' +
                ' and sample_t.type like ' + s +';';
    Open;
  end;

end;

procedure TfrmGetOxasBlanks.TimerTimer(Sender: TObject);
begin
  Button1Click(self);
end;

end.
