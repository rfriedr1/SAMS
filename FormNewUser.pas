unit FormNewUser;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmNewUser = class(TForm)
    EdtLastName: TLabeledEdit;
    EdtFirstName: TLabeledEdit;
    btnOK: TButton;
    btnCancel: TButton;
    EdtInstitution: TLabeledEdit;
    edtAddress1: TLabeledEdit;
    edtAddress2: TLabeledEdit;
    edtZip: TLabeledEdit;
    edtCity: TLabeledEdit;
    edtOrganisation: TLabeledEdit;
    edtCountry: TLabeledEdit;
    edtPhone1: TLabeledEdit;
    edtPhone2: TLabeledEdit;
    edtEMail: TLabeledEdit;
    grpUserInfo: TGroupBox;
    StaticText1: TStaticText;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmNewUser: TfrmNewUser;

implementation

{$R *.dfm}

end.
