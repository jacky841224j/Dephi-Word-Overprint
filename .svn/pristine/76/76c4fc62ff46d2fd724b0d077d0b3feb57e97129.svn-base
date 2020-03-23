unit SQLtext;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,FileCtrl,StrUtils,
  Vcl.Buttons,IniFiles, Vcl.Imaging.jpeg,ComCtrls,ShellAPI;

type
  TSQLSetting = class(TForm)
    lblSelect: TLabel;
    chk1: TCheckBox;
    pnl1: TPanel;
    lbledtPreView: TLabeledEdit;
    lbledtTOP3: TLabeledEdit;
    lbledtExport: TLabeledEdit;
    edtTest: TEdit;
    img1: TImage;
    pnl2: TPanel;
    btnFolder: TSpeedButton;
    pnl3: TPanel;
    btnSave: TSpeedButton;
    lbl2: TLabel;
    lbl1: TLabel;
    edtIP: TEdit;
    edtID: TEdit;
    edtPw: TEdit;
    pnl4: TPanel;
    btnudl: TSpeedButton;
    edtTest1: TEdit;
    pnl5: TPanel;
    btnOpen: TSpeedButton;
    btnClose: TBitBtn;
    mmoSQL: TMemo;
    mmoExec: TMemo;
    pnl6: TPanel;
    btnChange: TSpeedButton;
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnFolderClick(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure mmoSQLChange(Sender: TObject);
    procedure mmoExecChange(Sender: TObject);
    procedure mmoExecClick(Sender: TObject);
    procedure btnChangeClick(Sender: TObject);
    procedure lbledtPreViewChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DirPath : string;

implementation
 uses NATL;
{$R *.dfm}

procedure TSQLSetting.btn1Click(Sender: TObject);
var
temp : TStringlist ;
//textfile:TextFile;
begin
temp := TStringlist.Create;
//  AssignFile(textfile, DirPath);
//  ReWrite(textfile);
  temp.Add( '[oledb]');
  temp.Add( '; Everything after this line is an OLE DB initstring');
  temp.add( 'Provider=SQLOLEDB.1'+  //;Integrated Security=SSPI
            ';Password='+ edtPw.Text +
            ';Persist Security Info=true'+
            ';User ID=' + edtID.Text +
            ';Initial Catalog=' +edtTest1.Text +
            ';Data Source=' + edtIP.Text) ;
  temp.SaveToFile(DirPath+'\'+edtTest.Text+'\db.udl',TEncoding.Unicode);
//  Writeln(textfile,temp);
//  Closefile(textfile);
ShowMessage('db.udl建立成功');
end;

procedure TSQLSetting.btnChangeClick(Sender: TObject);
begin
  if SelectDirectory('請選擇設定檔目錄', '', DirPath) then
  else Exit;
  mmoSQL.Clear;
  mmoExec.Clear;
end;

procedure TSQLSetting.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('建立設定檔');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
end;

procedure TSQLSetting.btnDirpath2Click(Sender: TObject);
begin
//  repeat
//      SelectDirectory('請選擇存檔目錄', '', DirPath); //選擇目錄
//      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
//        Exit;
//  until DirPath <> '';
//    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
//  edtPath.Text := DirPath;
end;

procedure TSQLSetting.btnFolderClick(Sender: TObject);
begin
if Length(edtTest.text) < 1  then
  begin
    ShowMessage('請輸入考試名稱');
    Exit;
  end
else
  begin
    if SelectDirectory('請選擇設定檔目錄', '', DirPath) then
    else Exit;

    if not directoryExists(DirPath+'\'+edtTest.Text) then  //判斷此資料夾是否存在
    CreateDir(DirPath+'\'+edtTest.Text);                 //建立資料夾
    btnSave.Enabled := True;
  end;
  mmoSQL.Clear;
  mmoExec.Clear;
ShowMessage('資料夾建立成功');
end;

procedure TSQLSetting.btnOpenClick(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar(DirPath+'\'+edtTest.Text), nil, nil, SW_SHOW);
end;

procedure TSQLSetting.btnSaveClick(Sender: TObject);
var
Myinifile:Tinifile;
Filename : string;
i : Integer;
begin
  if lbledtPreView.Text = '' then
    begin
      ShowMessage('請輸入SQL語法');
      Exit;
    end
  else
    begin
      {$REGION '準備ini檔案資料'}
      Filename:= DirPath+'\'+edtTest.Text+'\SqlSetting.ini' ;
      myinifile:=Tinifile.Create(filename);
      {$ENDREGION}

      try
        //寫入
        myinifile.writestring('SQL','Save',StringReplace(lbledtExport.Text, #13#10,' ',[rfReplaceAll]));
        myinifile.writestring('SQL','TOP1',StringReplace(lbledtPreView.Text, #13#10,' ',[rfReplaceAll]));
        myinifile.writestring('SQL','TOP3',StringReplace(lbledtTOP3.Text, #13#10,' ',[rfReplaceAll]));
      finally
        myinifile.Free;
//        ShowMessage('寫入完成');
        if MessageBox(0,'設定檔建立成功,是否開啟資料夾?','OPEN',
                    MB_OKCANCEL + MB_ICONASTERISK + MB_DEFBUTTON2 ) = 1 then
        ShellExecute(Handle, 'open', PChar(DirPath+'\'+edtTest.Text), nil, nil, SW_SHOW);
      end;
    end;
End;

procedure TSQLSetting.chk1Click(Sender: TObject);
begin
  if chk1.Checked then
    begin
      lbledtPreView.text := 'select top 1 ' + mmoSQL.text ;
      lbledtTOP3.Text := 'select top 3 ' + mmoSQL.text;
      lbledtExport.Text := 'select ' + mmoSQL.text ;
    end
  else
    begin
      lbledtPreView.text := 'select top 1 ' + mmoSQL.text+ ' where 1 = 1 ' ;
      lbledtTOP3.Text := 'select top 3 ' + mmoSQL.text+ ' where 1 = 1 ' ;
      lbledtExport.Text := 'select ' + mmoSQL.text+ ' where 1 = 1 ' ;
    end;
end;

procedure TSQLSetting.lbledtPreViewChange(Sender: TObject);
begin
  btnSave.Enabled := True;
end;

procedure TSQLSetting.mmoExecChange(Sender: TObject);
begin
  if Length ( mmoExec.text ) > 0 then
    begin
      mmoSQL.Enabled := False ;
      chk1.Enabled := False;
      lbledtPreView.text :=  mmoExec.text ;
      lbledtTOP3.Text :=  mmoExec.text ;
      lbledtExport.Text :=  mmoExec.text ;
    end
  else
    begin
      lbledtPreView.text := '';
      lbledtTOP3.Text := '';
      lbledtExport.Text := '';
      mmoSQL.Enabled := True;
      chk1.Enabled := true;
      mmoExec.text := '';
    end;
end;

procedure TSQLSetting.mmoExecClick(Sender: TObject);
begin
  mmoSQL.Clear;
  mmoExec.Clear;
end;

procedure TSQLSetting.mmoSQLChange(Sender: TObject);
begin
  if Length ( mmoSQL.text ) > 0 then
    begin
      chk1Click(nil);
      mmoExec.Enabled := False ;
      chk1.Enabled := True;
    end
  else
    begin
      lbledtPreView.text := '';
      lbledtTOP3.Text := '';
      lbledtExport.Text := '';
      mmoExec.Enabled := True;
      chk1.Enabled := False;
    end;
end;

end.
