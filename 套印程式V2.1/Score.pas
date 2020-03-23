unit Score;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Grids, DBGrids;

type
  TScoreLogin = class(TForm)
    dlgOpen1: TOpenDialog;
    con1: TADOConnection;
    qryExport: TADOQuery;
    ds1: TDataSource;
    dbgrd1: TDBGrid;
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    lbl5: TLabel;
    edtID: TEdit;
    btnSing: TButton;
    edtNo: TEdit;
    btnOUT: TButton;
    edtItem4: TEdit;
    edtItem5: TEdit;
    edtItem8: TEdit;
    edtItem9: TEdit;
    edtItem11: TEdit;
    edtItem12: TEdit;
    edtItem13: TEdit;
    edtItem14: TEdit;
    edtItem15: TEdit;
    btnCalculation: TButton;
    btnCheck: TButton;
    btnOutExcel: TButton;
    btnAllSt: TButton;
    edtItem1: TEdit;
    edtItem2: TEdit;
    edtItem3: TEdit;
    edtItem6: TEdit;
    edtItem7: TEdit;
    edtItem10: TEdit;
    btnLoginOff: TButton;
    btnLoginOn: TButton;
    edtItemTotal: TEdit;
    grpreferee: TGroupBox;
    rb1: TRadioButton;
    rb2: TRadioButton;
    rb3: TRadioButton;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ScoreLogin: TScoreLogin;

implementation

{$R *.dfm}

procedure TScoreLogin.FormShow(Sender: TObject);
begin
if dlgOpen1.Execute then
  try
    con1.Connected := false;
    con1.ConnectionString := 'FILE NAME='+dlgOpen1.FileName+'\'+ 'db.udl';
    con1.Provider := dlgOpen1.FileName+'\'+'db.udl';
    con1.Connected := true;
  except
    showmessage('連線錯誤,請檢查.udl設定是否正確');
    EXIT;
  end;
end;

end.
