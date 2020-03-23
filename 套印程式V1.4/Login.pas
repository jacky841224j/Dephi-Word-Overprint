unit Login;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ADODB, GIFImg, ExtCtrls;

type
  TLoginCheck = class(TForm)
    edtAd: TEdit;
    edtPw: TEdit;
    btnOK: TButton;
    edtDomain: TEdit;
    img1: TImage;
    procedure btnOKClick(Sender: TObject);
    procedure edtPwKeyPress(Sender: TObject; var Key: Char);
    procedure edtDomainKeyPress(Sender: TObject; var Key: Char);
    procedure edtAdKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    function checkAD(Username, Password: String): Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  LoginCheck: TLoginCheck;

implementation

 uses natl,Grade ;
{$R *.dfm}

procedure TLoginCheck.btnOKClick(Sender: TObject);
var
phToken: cardinal;
begin

//if  checkAD(edtad.Text,edtpw.Text) then

if LogonUser(pChar(edtad.Text), pChar(edtDomain.Text), pChar(edtpw.Text),
LOGON32_LOGON_NETWORK,LOGON32_PROVIDER_DEFAULT,phToken) then
begin
Form1.Show();
//hide();
end
else ShowMessage('登入失敗，請重試');

end;


procedure TLoginCheck.edtAdKeyPress(Sender: TObject; var Key: Char);
begin
if (Key = #13) then
begin
Key := #0;
Perform(Wm_NextDlgCtl,0,0);
end;
end;

procedure TLoginCheck.edtDomainKeyPress(Sender: TObject; var Key: Char);
begin
if (Key = #13) then
begin
Key := #0;
Perform(Wm_NextDlgCtl,0,0);
end;
end;

procedure TLoginCheck.edtPwKeyPress(Sender: TObject; var Key: Char);
begin
if Key=#13 then btnokClick(Sender);
end;

procedure TLoginCheck.FormShow(Sender: TObject);
begin
Self.Left := (Screen.Width - Self.Width) div 2 ;
Self.Top := (Screen.Height - Self.Height) div 2 ;
end;

function TLoginCheck.checkAD(Username, Password: String): Boolean;
const
  ProviderStr ='Provider=ADsDSOObject;User ID=%s;Password=%s;Encrypt Password=False; Mode=Read;Bind Flags=0;ADSI Flag=-2147483648';
  ChkADSQL='SELECT cn FROM '#39'%s'#39' WHERE objectClass='#39'user'#39;
var
ChkAD : TADOQuery;
ADPath :string;
begin
  ChkAD := TADOQuery.Create(nil);
  try
    ADPath := 'seat.local';
    Format('LDAP://%s',[ADPath]);
    ShowMessage(Format('LDAP://%s',[ADPath]));
    ChkAD.ConnectionString := Format(ProviderStr,[Username,Password]);
    ChkAD.SQL.Text := Format(ChkADSQL,[ADPath]);
  try
    ChkAD.Open;
    Result := (ChkAD.RecordCount>0);
  except
   Result := False;
  end;
  finally
    FreeAndNil(ChkAD);
  end;
end;


end.
