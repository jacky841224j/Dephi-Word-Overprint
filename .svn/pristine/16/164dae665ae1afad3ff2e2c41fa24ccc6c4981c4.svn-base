unit Grade;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, CheckLst, Gauges, Mask, DBCtrls,IniFiles ,
  DBClient, ExtCtrls,TlHelp32,Excel2000,ComObj,ADODB,FileCtrl,StrUtils,ComCtrls,Math,sockets,
  GIFImg, Buttons, jpeg,shellapi, pngimage,System.Win.ScktComp;



type

//  TMyThread = class(TThread)
//  protected
//  procedure Execute; override;
//  end;

  TForm4 = class(TForm)
    Gauge2: TGauge;
    dbgrd1: TDBGrid;
    ds1: TDataSource;
    strngrdCheckList: TStringGrid;
    lblSave: TLabel;
    edtPath: TEdit;
    btnDirpath2: TButton;
    edtFile: TEdit;
    btnFile: TButton;
    lblFile: TLabel;
    tmr1: TTimer;
    qryExport: TADOQuery;
    chkPic: TCheckBox;
    edtRows: TEdit;
    edtCols: TEdit;
    btnPic: TButton;
    edtPic: TEdit;
    chkWord: TCheckBox;
    lblCNumber: TLabel;
    chkExit: TCheckBox;
    lblTime: TLabel;
    mmoLog: TMemo;
    mmoQuickSetting: TMemo;
    btnExport: TBitBtn;
    btnPreview: TBitBtn;
    imgBackGroup: TImage;
    imgWord: TImage;
    pnlWord: TPanel;
    pnlPic: TPanel;
    pnlExit: TPanel;
    imgStop: TImage;
    btnKillTask: TSpeedButton;
    btnChgSet: TBitBtn;
    btnOpenWord: TBitBtn;
    btnOpenFile: TBitBtn;
    blnhnt1: TBalloonHint;
    mmoTemp: TMemo;
    btnThead: TBitBtn;
    pnlReplace: TPanel;
    imgReplace: TImage;
    chkReplace: TCheckBox;
    con1: TADOConnection;
    qry1: TADOQuery;
    dlgOpen1: TOpenDialog;
    dlgOpendlgudl: TOpenDialog;
    dlgSave1: TSaveDialog;
    btnClose: TBitBtn;
    ClientSocket1: TClientSocket;
    procedure ChoiceTest(Sender: TObject);
    procedure FormShow(Sender: TObject) ;
    procedure btnKillTaskClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure chkPicClick(Sender: TObject);
    procedure btnPicClick(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure btnChgSetClick(Sender: TObject);
    procedure btnOpenWordClick(Sender: TObject);
    procedure btnOpenFileClick(Sender: TObject);
    procedure btnTheadClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);

  private
    { Private declarations }
  public
//     procedure GetRepExFile(const ExName, SavePath: String);
    { Public declarations }
  end;

var
  Form4: TForm4;
  Filename,dirpath,folder :string;
  Myinifile:Tinifile;
  t,m,h,Cnt,ttTime,LoginInCheck : Integer ;
  times : longint = 0 ;
  timeh : longint = 0 ;
  timem : longint = 0 ;
  SqlFS,SqlPV,SqlEP : string ;
implementation
uses NATL, Login,UnitMyThread;
{$R *.dfm}


procedure TForm4.btnTheadClick(Sender: TObject);
var
  WordFileName,find,LogTemp: string;
  SavePath,PicPath,Cols,Rows : string;
  WordApp, WordDoc, myRange, vSaveNone : Variant;
  i,p,chk,x,y: Integer;
  sqlstr : TStringList;
  SaveTxt:TextFile;
begin
 t:= 0;
 m:= 0;
 h:=0;
 Cols := edtCols.Text ;
 Rows := edtRows.Text;
 PicPath := edtPic.Text;
 SavePath := edtPath.Text;
 edtPath.Enabled := false;
 edtFile.Enabled := false;
 edtPic.Enabled := false;
 mmoTemp.Clear;

 sqlstr := TStringList.Create;
 ShowMessage('資料內含個資，請妥善保存！');

  try
  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION 'SQL語法'}
  Cnt := 0;
  chk := 0;
  qryExport.Close;

  for p := 1 to strngrdCheckList.RowCount  do     //儲存有搜尋條件的欄位
  if strngrdCheckList.Cells[3,p] <> '' then
    begin
      find := find + #32'And'#32 + strngrdCheckList.Cells[0,p] + #32 + strngrdCheckList.Cells[3,p] + #32 + strngrdCheckList.Cells[4,p];
      chk:=1;
    end;
  if chk = 0 then qryExport.SQL.Text := SqlEP
  else if chk = 1 then qryExport.SQL.Text := SqlEP + find ;
  qryExport.Open;
  qryExport.First;
  ttTime :=  qryExport.RecordCount;

  if qryExport.RecordCount < 1  then
    begin
      ShowMessage('未找到符合條件的資料');
      exit;
    end;
  {$ENDREGION}

  {$REGION '進度條'}
  Gauge2.MinValue:= 0;
  Gauge2.MaxValue:= qryExport.RecordCount+1;
  Gauge2.Progress:= 0;
  {$ENDREGION}


{$REGION '判斷Word是否安裝'}
  WordFileName := dlgOpen1.FileName;
  WordApp := CreateOleObject('Word.Application');
  if WordApp.Version < 12 then
    begin
    ShowMessage('此電腦未正確安裝Word 2007或以上的版本');
    Exit;
    end;
{$ENDREGION}
  tmr1.Enabled:=true;
  with qryExport do
  begin

  {$REGION '判斷有值欄位'}
  for p := 1 to FieldCount  do
   if strngrdCheckList.Cells[2,p] <> '' then
    begin
      mmoTemp.lines.add( IntToStr(p) );    //將有值的欄位存在mmoLog
      LogTemp := LogTemp+'['+strngrdCheckList.Cells[0,p]+'/'+strngrdCheckList.Cells[2,p]  +']'+',';
    end;
  {$ENDREGION}

  {$REGION '套印'}
  for i:= 1 to RecordCount  do
   begin

    if chkExit.Checked then  Break;
    if chkReplace.Checked = False then
      if FileExists(DirPath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString+'.pdf') then
        begin
          Inc(Cnt);
          Gauge2.Progress:= i;
          lblCNumber.Caption := '筆數：'+ IntToStr(qryExport.RecordCount)+'/'+ IntToStr(Cnt);
          Next;
          continue;
        end;
    WordApp.Visible := false;
    WordApp.Application.DisplayAlerts := False;
    WordDoc := WordApp.Documents.Open(WordFileName);
    myRange := WordDoc.Content;

    for x := 0 to  mmoTemp.lines.Count-1 do   //套印資料

    myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,StrToInt(mmoTemp.lines[x])], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,StrToInt(mmoTemp.lines[x])] ).AsString, Replace:=2);

    if chkPic.Checked then   //套印圖片
      WordDoc.Tables.Item(1).Cell(Cols,Rows).range.InlineShapes.AddPicture((PicPath + FieldByName(myinifile.readstring('Photo','pic','')).AsString),false,true);

    //判斷是否儲存Word檔
    if chkWord.Checked  then  WordDoc.SaveAs(SavePath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString+'.docx');
    WordDoc.ExportAsFixedFormat(SavePath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString + '.pdf',17);
    WordApp.Documents.close(vSaveNone);

    Inc(Cnt);
    Gauge2.Progress:= i;
    lblCNumber.Caption := '筆數：'+ IntToStr(qryExport.RecordCount)+'/'+ IntToStr(Cnt);
    Next;
   end;
  {$ENDREGION}
    tmr1.Enabled:=false;
    mmoLog.Lines.Add(SavePath) ;
  end;

  finally
    WordApp.Quit;
    WordApp:=Unassigned;
    edtPath.Enabled := true;
    edtFile.Enabled := true;
    edtPic.Enabled := true;
  end;

   Gauge2.Progress:= Gauge2.MaxValue;
   if MessageDlg('匯出完成',mtInformation,[mbYes],0)=mrYes then lblTime.Caption := '' ;

  {$REGION '寫入LOG'}
  //判斷資料庫是否存在
  sqlstr.Add  (' if not EXISTS(SELECT * FROM sysobjects  WHERE name= ''Log_CheckIn'' ) ' );
  sqlstr.Add  ('    begin');
  sqlstr.Add  ('    create table Log_CheckIn ( ');
  sqlstr.Add  ('    No  int  PRIMARY KEY IDENTITY(1,1), ');
  sqlstr.Add  ('    Time            nvarchar(16)    ,');
  sqlstr.Add  ('    AD_Name         nvarchar(50)    ,');
  sqlstr.Add  ('    Local_Name      nvarchar(50)    ,');
  sqlstr.Add  ('    DB_IP           nvarchar(15)    ,');
  sqlstr.Add  ('    DB_dbo          nvarchar(MAX)   ,');
  sqlstr.Add  ('    SQL             nvarchar(MAX)   ,');
  sqlstr.Add  ('    Setting         nvarchar(MAX)   ,');
  sqlstr.Add  ('    SavePath        nvarchar(MAX))  ');
  sqlstr.Add  ('    INSERT INTO '+ myinifile.readstring('SQL','SQLLOG','') + 'VALUES (' + QuotedStr(mmoLog.Lines[0]) +','+ QuotedStr(mmoLog.Lines[1]) +','+QuotedStr(mmoLog.Lines[2]) +','+QuotedStr(mmoLog.Lines[3]) +',' + QuotedStr(mmoLog.Lines[4]) + ',' + QuotedStr(LogTemp) +',' + QuotedStr(SqlEP) +',' + QuotedStr(mmoLog.Lines[5])+ ') end'  ) ;
  sqlstr.Add  (' else ');
  sqlstr.Add  ('    INSERT INTO '+ myinifile.readstring('SQL','SQLLOG','') + 'VALUES (' + QuotedStr(mmoLog.Lines[0]) +','+ QuotedStr(mmoLog.Lines[1]) +','+QuotedStr(mmoLog.Lines[2]) +','+QuotedStr(mmoLog.Lines[3]) +',' + QuotedStr(mmoLog.Lines[4]) + ',' + QuotedStr(LogTemp) +',' + QuotedStr(SqlEP) +',' + QuotedStr(mmoLog.Lines[5]) + ')'  ) ;
  qryExport.Close;
  qryExport.SQL.Text := StringReplace(sqlstr.text ,#$D#$A,' ',[rfReplaceAll]);
  qryExport.ExecSQL;
  {$ENDREGION}

  {$REGION '詢問是否儲存設定檔'}
  if not FileExists(folder+'\'+'QuickSetting.txt') then
    begin
      AssignFile(SaveTxt,folder+'\'+'QuickSetting.txt');
      Rewrite(SaveTxt);
      for y := 0 to strngrdCheckList.RowCount-1 do   writeln(SaveTxt,strngrdCheckList.Cells[2,y+1]);
      CloseFile(SaveTxt);
    end
  else if MessageDlg('是否覆蓋設定檔',mtInformation,[mbYes,mbNo],0)=mrYes then
    begin
      AssignFile(SaveTxt,folder+'\'+'QuickSetting.txt');
      Rewrite(SaveTxt);
      for y := 0 to strngrdCheckList.RowCount-1 do   writeln(SaveTxt,strngrdCheckList.Cells[2,y+1]);
      CloseFile(SaveTxt);
      ShowMessage('設定檔已儲存！');
    end;

  {$ENDREGION}

  if chkExit.Checked = true then chkExit.Checked := false;
  myinifile.Free;
  sqlstr.Free;
  Form1.KillExcelTask;

end;

procedure TForm4.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('資料套印');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
//  form1.pgc1.Visible := false ;
end;

procedure TForm4.btnChgSetClick(Sender: TObject);
begin
  ChoiceTest(nil);
end;

procedure TForm4.btnOpenFileClick(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar(edtPath.Text), nil, nil, SW_SHOW);
end;

procedure TForm4.btnOpenWordClick(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar(edtFile.Text), nil, nil, SW_SHOW);
end;

procedure TForm4.btnDirpath2Click(Sender: TObject);
begin
  repeat
      SelectDirectory('請選擇存檔目錄', '', DirPath); //選擇目錄
      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
        Exit;
    until DirPath <> '';
    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
  edtPath.Text := DirPath;
end;

procedure TForm4.btnExportClick(Sender: TObject);
begin
MyThread(btnTheadClick,nil);
end;

procedure TForm4.btnFileClick(Sender: TObject);
begin
  if dlgOpen1.Execute then   edtFile.Text := dlgOpen1.FileName ;
end;

procedure TForm4.btnKillTaskClick(Sender: TObject);
begin
  Form1.KillWordTask;
  ShowMessage('已關閉Word');
end;

procedure TForm4.btnPicClick(Sender: TObject);
begin
  repeat
      SelectDirectory('請選擇圖片目錄', '', DirPath); //選擇目錄
      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
        Exit;
    until DirPath <> '';
    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
  edtPic.Text := DirPath;
end;

procedure TForm4.btnPreviewClick(Sender: TObject);
var
  WordFileName: string;
  WordApp, WordDoc, myRange : Variant;
  i : Integer;
begin
  form1.KillExcelTask;

  {$REGION 'SQL語法'}
  qryExport.Close;
  qryExport.SQL.Text := SqlPV;
  qryExport.Open;
  qryExport.First;
  {$ENDREGION}

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION '預覽'}

  WordFileName := dlgOpen1.FileName;
  WordApp := CreateOleObject('Word.Application');
  if WordApp.Version < 12 then
    begin
    ShowMessage('此電腦未正確安裝Word 2007或以上的版本');
    Exit;
    end;

  with  qryExport do
  begin
   WordDoc := WordApp.Documents.Open(WordFileName);
   myRange := WordDoc.Content;
   WordApp.Visible := true;
   WordApp.Application.DisplayAlerts := False;
  for i:= 1 to qryExport.FieldCount  do
  if strngrdCheckList.Cells[2,i] <> '' then
  myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,i], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,i]).AsString, Replace:=2);

  if chkPic.Checked then
    WordDoc.Tables.Item(1).Cell(edtCols.Text,edtRows.Text).range.InlineShapes.AddPicture((edtPic.Text + FieldByName(myinifile.readstring('Photo','pic','')).AsString),false,true);
  end;
   {$ENDREGION}
end;

procedure TForm4.chkPicClick(Sender: TObject);
begin
if chkPic.Checked then
  begin
    edtRows.Visible := true ;
    edtCols.Visible := true ;
    edtPic.Visible := true ;
    btnPic.Visible := true ;
  end
else
  begin
    edtRows.Visible := false ;
    edtCols.Visible := false ;
    edtPic.Visible := false ;
    btnPic.Visible := false ;
  end;
end;

procedure TForm4.FormShow(Sender: TObject);
begin
ChoiceTest(nil);
end;

procedure TForm4.ChoiceTest(Sender: TObject);
var
i,temp1,temp2 : Integer;
strtemp,temp,IP,DB : string ;
sock : TClientSocket;
begin
  SqlEP := '';
  SqlPV := '';
  SqlFS := '';
  sock := TClientSocket.Create(self);  //獲取本機名稱
  if SelectDirectory('請選擇設定檔目錄', '', folder) then
  mmoTemp.Lines.LoadFromFile(folder+'\SqlSetting.ini');

  {$REGION '設定DB連接'}
  try
    con1.Connected := false;
    con1.ConnectionString := 'FILE NAME='+folder+'\'+ 'db.udl';
    con1.Provider := folder+'\'+'db.udl';
    con1.Connected := true;
  except
    showmessage('連線失敗,請檢查路徑或是.udl設定是否正確');
    EXIT;
  end;
  {$ENDREGION}

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION '擷取SQL 字串'}
  //匯出SQL
  temp := mmoTemp.Lines.Text;
  temp1:= POS('@', temp);
  strtemp:= copy(temp, 0,temp1-1);       //擷取字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
  if not (strtemp[i] = '@') then SqlEP := SqlEP + strtemp[i]
  else break;
  strtemp := '';
  //預覽SQL
  temp1:= POS('@', temp);
  temp2:= POS('%', temp);
  strtemp:= copy(temp, temp1+1,temp2-1);       //擷取字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
  if not (strtemp[i] = '%') then SqlPV := SqlPV + strtemp[i]
  else break;
  strtemp := '';
  //顯示前三筆SQL
  temp1:= POS('%', temp);
  temp2:= POS('#', temp);
  strtemp:= copy(temp, temp1+1,temp2-1);       //擷取字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
  if not (strtemp[i] = '#') then SqlFS := SqlFS + strtemp[i]
  else break;
  strtemp := '';
  {$ENDREGION}

  {$REGION '設定欄位'}
  qry1.Close;
  qry1.SQL.Text := SqlFS;
  qry1.Open;
  qry1.First;
  with qry1 do
  for I := 0 to  qry1.FieldCount -1 do
  begin
    dbgrd1.Columns[i].Width := 60;
    strngrdCheckList.Cells[0,i+1]:=dbgrd1.Columns[i].FieldName;
    strngrdCheckList.Cells[1,i+1]:=myinifile.readstring('Help',('help'+ IntToStr(i+1)),'');
  end;
  {$ENDREGION}

  {$REGION 'StringGrid'}
  strngrdCheckList.Cells[0,0] := #32#32#32#32'---【資料欄位】---' ;
  strngrdCheckList.Cells[1,0] := #32#32#32#32'---【欄位說明】---' ;
  strngrdCheckList.Cells[2,0] := #32#32#32#32'---【取代文字】---' ;
  strngrdCheckList.Cells[3,0] := #32#32#32#32'=' ;
  strngrdCheckList.Cells[4,0] := #32#32#32#32'---【搜尋條件】---' ;
  strngrdCheckList.ColWidths[0] := 200;
  strngrdCheckList.ColWidths[1] := 180;
  strngrdCheckList.ColWidths[2] := 180;
  strngrdCheckList.ColWidths[3] := 50;
  strngrdCheckList.ColWidths[4] := 180;
  strngrdCheckList.RowCount:=qry1.FieldCount+1 ;
  {$ENDREGION}

  {$REGION '設定預設路徑'}
  if not directoryExists(folder+'\'+'套印資料夾') then  //判斷此資料夾是否存在
  CreateDir(folder+'\'+'套印資料夾');                 //建立資料夾
  edtPath.Text:= folder+'\'+'套印資料夾'+'\';
  dirpath:= edtPath.Text   ;
  edtFile.Text :=folder+'\'+'套印範本.docx' ;
  edtPic.Text := folder+'\' + 'Photo\';
  {$ENDREGION}

  {$REGION '擷取LOG 字串'}
  temp := con1.ConnectionString;
  //擷取IP字串
  temp1:= POS('e=', temp);
  strtemp:= copy(temp, (temp1)+2,15);       //擷取IP 最大長度字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
  if ( (strtemp[i] = '.') or (strtemp[i]  in ['0'..'9'])) then  IP := IP+ strtemp[i] ;
  strtemp := '' ;
  //擷取使用DB資料表字串
  temp1:= POS('g=', temp);
  strtemp:= copy(temp, (temp1)+2,20);       //擷取db 最大長度字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
  if not (strtemp[i] = ';') then DB := DB + strtemp[i]
  else break;
  {$ENDREGION}

  {$REGION 'Log紀錄'}
  mmoLog.Lines.Add(FormatDateTime('yyyy/mm/dd'+#32+'hh:mm', Now)) ;      //取得當前時間
  mmoLog.Lines.Add(LoginCheck.edtad.Text) ;                              //AD 使用者名稱
  mmoLog.Lines.Add(sock.host) ;                                 //本機名稱
  mmoLog.Lines.Add(IP) ;                                                 //DB IP
  mmoLog.Lines.Add(DB) ;                                                 //DB 資料表
  {$ENDREGION}

  {$REGION '判斷是否有設定檔'}
  if  FileExists(folder+'\'+'QuickSetting.txt') then
    begin
      if MessageDlg('是否套用上次設定？',mtInformation,[mbYes,mbNo],0)=mrYes then
         mmoQuickSetting.Lines.LoadFromFile(folder+'\'+'QuickSetting.txt');
      for i := 0 to mmoQuickSetting.Lines.Count-1 do strngrdCheckList.Cells[2,i+1] := mmoQuickSetting.Lines[i];
    end;
  {$ENDREGION}

  myinifile.Free;

  end;

procedure TForm4.tmr1Timer(Sender: TObject);
var
time : string ;
begin
  t := t+2;
  if t = 60 then
  begin
    t := t-60 ;
    m := m+1;
  end;

  if h = 60 then
  begin
    m := m-60 ;
    h := h+1;
  end;

  if (h <1) and (m<1) and (t = 10) then
  begin
    time := IntToStr( ceil((ttTime/(cnt/10)))) ; // 預估產完時間
    if StrToFloat (time) > 60 then
    begin
    times := ( StrToInt(time) mod 60);
    timeh := StrToInt(time) div 3600;
    timem := ((StrToInt(time) - times) - timeh*3600 ) div 60 ;
    end;
  end;

  if  (h <1) and (m<1)      then lblTime.Caption :='花費時間：'+IntToStr(t) +'  (預估時間' + IntToStr (timeh) +'時' + IntToStr (timem) +'分)'
  else if  (h <1) and (m>0) then lblTime.Caption :='花費時間：'+IntToStr(m)+':'+IntToStr(t) +'  (預估時間' + IntToStr (timeh) +'時' + IntToStr (timem) +'分)'
  else                           lblTime.Caption :='花費時間：'+IntToStr(h)+':'+IntToStr(m)+':'+IntToStr(t) +'  (預估時間' + IntToStr (timeh) +'時' + IntToStr (timem) +'分)';
  Application.ProcessMessages;
end;

end.

