unit SQLtext;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,FileCtrl,StrUtils,
  Vcl.Buttons,IniFiles, Vcl.Imaging.jpeg,ComCtrls;

type
  TSQLSetting = class(TForm)
    edtPath: TEdit;
    btnDirpath2: TButton;
    edtSQL: TEdit;
    lblSelect: TLabel;
    chk1: TCheckBox;
    pnl1: TPanel;
    lbl1: TLabel;
    lbledtPreView: TLabeledEdit;
    lbledtTOP3: TLabeledEdit;
    lbledtExport: TLabeledEdit;
    edtExce: TEdit;
    btnSave: TBitBtn;
    lbl2: TLabel;
    btnClose: TBitBtn;
    procedure btnDirpath2Click(Sender: TObject);
    procedure edtSQLChange(Sender: TObject);
    procedure edtExceChange(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SQLSetting: TSQLSetting;
  DirPath : string;

implementation
 uses NATL;
{$R *.dfm}

procedure TSQLSetting.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('SQLSetting.ini_設定');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
end;

procedure TSQLSetting.btnDirpath2Click(Sender: TObject);
begin
  repeat
      SelectDirectory('請選擇存檔目錄', '', DirPath); //選擇目錄
      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
        Exit;
  until DirPath <> '';
    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
  edtPath.Text := DirPath;
end;

procedure TSQLSetting.btnSaveClick(Sender: TObject);
var
Myinifile:Tinifile;
Filename : string;
i : Integer;
begin
//  AssignFile(textfile, SqlSetting);
//  ReWrite(textfile);
//  Writeln(textfile,lbledtExport.text + #13 + '@'+ #13 + lbledtPreView.text + #13 + '%'+ #13 + lbledtTOP3.Text + #13 + '#'+ #13 );
//  Closefile(textfile);

  {$REGION '準備ini檔案資料'}
  Filename:=DirPath+'\'+'SqlSetting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  try
    //寫入
    myinifile.writestring('SQL','Save',lbledtExport.Text);
    myinifile.writestring('SQL','TOP1',lbledtPreView.text);
    myinifile.writestring('SQL','TOP3',lbledtTOP3.Text);
  finally
    myinifile.Free;
    ShowMessage('寫入完成');
  end;
End;

procedure TSQLSetting.chk1Click(Sender: TObject);
begin
  if chk1.Checked then
    begin
      lbledtPreView.text := 'select top 1 ' + edtSQL.text ;
      lbledtTOP3.Text := 'select top 3 ' + edtSQL.text;
      lbledtExport.Text := 'select ' + edtSQL.text ;
    end
  else
    begin
      lbledtPreView.text := 'select top 1 ' + edtSQL.text+ ' where 1 = 1 ' ;
      lbledtTOP3.Text := 'select top 3 ' + edtSQL.text+ ' where 1 = 1 ' ;
      lbledtExport.Text := 'select ' + edtSQL.text+ ' where 1 = 1 ' ;
    end;
end;

procedure TSQLSetting.edtExceChange(Sender: TObject);
begin
  if Length ( edtExce.text ) > 0 then
    begin
      edtSQL.Enabled := False ;
      chk1.Enabled := False;
      lbledtPreView.text :=  edtExce.text ;
      lbledtTOP3.Text :=  edtExce.text ;
      lbledtExport.Text :=  edtExce.text ;
    end
  else
    begin
      lbledtPreView.text := '';
      lbledtTOP3.Text := '';
      lbledtExport.Text := '';
      edtSQL.Enabled := True;
      chk1.Enabled := true;
      edtExce.text := '';
    end;
end;

procedure TSQLSetting.edtSQLChange(Sender: TObject);
begin
  if Length ( edtSQL.text ) > 0 then
    begin
      chk1Click(nil);
      edtExce.Enabled := False ;
      chk1.Enabled := True;
    end
  else
    begin
      lbledtPreView.text := '';
      lbledtTOP3.Text := '';
      lbledtExport.Text := '';
      edtExce.Enabled := True;
      chk1.Enabled := False;
    end;
end;

end.
