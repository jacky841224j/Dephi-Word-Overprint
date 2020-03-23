unit Grade;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, CheckLst, Gauges, Mask, DBCtrls,IniFiles ,
  DBClient, ExtCtrls,TlHelp32,Excel2000,ComObj,ADODB,FileCtrl,StrUtils,ComCtrls,Math,sockets;



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
    btnPreview: TButton;
    lbl19: TLabel;
    edtPath: TEdit;
    btnDirpath2: TButton;
    edtFile: TEdit;
    btnFile: TButton;
    lbl1: TLabel;
    btnExport: TButton;
    tmr1: TTimer;
    btnKillTask: TButton;
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
    mmoSql: TMemo;
    mmoLog: TMemo;
    mmoQuickSetting: TMemo;
    lblRows: TLabel;
    procedure FormShow(Sender: TObject) ;
    procedure btnKillTaskClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure chkPicClick(Sender: TObject);
    procedure btnPicClick(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure strngrdCheckListClick(Sender: TObject);
    procedure btnDelTextClick(Sender: TObject);

  private
    { Private declarations }
  public

    { Public declarations }
  end;

var
  Form4: TForm4;
  Filename,dirpath,folder :string;
  Myinifile:Tinifile;
  t,m,h,Cnt,ttTime : Integer ;
  times : longint = 0 ;
  timeh : longint = 0 ;
  timem : longint = 0 ;
implementation
uses NATL, Login;
{$R *.dfm}


procedure TForm4.btnDelTextClick(Sender: TObject);
begin
  if MessageDlg('是否儲存此次設定',mtInformation,[mbYes,mbNo],0)=mrYes then
    begin
      DeleteFile(folder+'\'+'QuickSetting.txt');
      ShowMessage('設定檔已刪除');
    end
  else ShowMessage('已取消動作');

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
var
  S_Photo,WordFileName,find,temp,LogTemp: string;
  WordApp, WordDoc, myRange, vSaveNone : Variant;
  i,p,chk,x,y: Integer;
  sqlstr : TStringList;
  SaveTxt:TextFile;
begin
 t:= 0;
 m:= 0;
 h:=0;
 tmr1.Enabled:=true;
 sqlstr := TStringList.Create;
 ShowMessage('資料內含個資，請妥善保存！');

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION 'SQL語法'}
  Cnt := 0;
  chk := 0;
  qryExport.Close;
  for p := 1 to strngrdCheckList.RowCount  do
     if strngrdCheckList.Cells[3,p] <> '' then
      begin
       find := find + #32'And'#32+'[' + strngrdCheckList.Cells[0,p] +']'+ #32 + strngrdCheckList.Cells[3,p] + #32 + strngrdCheckList.Cells[4,p];
       chk:=1;
      end;
  if chk = 0 then qryExport.SQL.Text := myinifile.readstring('SQL','SQLExport','')
  else if chk = 1 then qryExport.SQL.Text := myinifile.readstring('SQL','SQLExport','') + find ;
  qryExport.Open;
  qryExport.First;
  ttTime :=  qryExport.RecordCount;
  {$ENDREGION}

  {$REGION '進度條'}
  Gauge2.MinValue:= 0;
  Gauge2.MaxValue:= qryExport.RecordCount+1;
  Gauge2.Progress:= 0;
  {$ENDREGION}
try
  {$REGION '判斷Word是否安裝'}
  WordFileName := Form1.dlgOpen1.FileName;
  WordApp := CreateOleObject('Word.Application');
  if WordApp.Version < 12 then
    begin
    ShowMessage('此電腦未正確安裝Word 2007或以上的版本');
    Exit;
    end;
   {$ENDREGION}
  with qryExport do
  begin
  {$REGION '判斷有值欄位'}
  for p := 1 to FieldCount  do
   if strngrdCheckList.Cells[2,p] <> '' then
    begin
    mmoSql.lines.add( IntToStr(p) );    //將有值的欄位存在mmoSql
    LogTemp := LogTemp+'['+strngrdCheckList.Cells[0,p]+'-'+strngrdCheckList.Cells[2,p]  +']'+',';
    end;
    mmoLog.Lines.Add(LogTemp) ;
  {$ENDREGION}

  {$REGION '套印'}
  for i:= 1 to RecordCount  do
   begin
    if chkExit.Checked then  Break;

//    if FileExists(DirPath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString+'.pdf') then
//        if  (MessageDlg('檔案已存在，是否覆蓋',mtcustom,[mbYes]+[mbNo],0) = 7) then break ;
    WordApp.Visible := false;
    WordApp.Application.DisplayAlerts := False;
    WordDoc := WordApp.Documents.Open(WordFileName);
    myRange := WordDoc.Content;
    for x := 0 to  mmoSql.lines.Count-1 do   //套印資料
    begin
    temp := mmoSql.lines[x];
    myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,StrToInt(temp)], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,StrToInt(temp)] ).AsString, Replace:=2);
    end;

    if chkPic.Checked then   //套印圖片
    begin
    S_Photo := edtPic.Text + FieldByName(myinifile.readstring('Photo','pic','')).AsString ;
    WordDoc.Tables.Item(1).Cell(edtCols.Text,edtRows.Text).range.InlineShapes.AddPicture(S_Photo,false,true);
    end;

    //判斷是否儲存Word檔
    if chkWord.Checked  then  WordDoc.SaveAs(DirPath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString+'.docx');
    WordDoc.ExportAsFixedFormat(DirPath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString + '.pdf',17);
    WordApp.Documents.close(vSaveNone);

    Inc(Cnt);
    Gauge2.Progress:= i;
    lblCNumber.Caption := '筆數：'+ IntToStr(qryExport.RecordCount)+'/'+ IntToStr(Cnt);
    Application.ProcessMessages;
    Next;
   end;
  {$ENDREGION}
    tmr1.Enabled:=false;
    mmoLog.Lines.Add(DirPath) ;
  end;
finally
  WordApp.Quit;
  WordApp:=Unassigned;
end;

  if qryExport.RecordCount < 1  then
    begin
       ShowMessage('未找到符合條件的資料');
       exit;
    end;
   Gauge2.Progress:= Gauge2.MaxValue;
   if MessageDlg('匯出完成',mtInformation,[mbYes],0)=mrYes then lblTime.Caption := '' ;

  {$REGION '寫入LOG'}
  //判斷資料庫是否存在
  sqlstr.Add  (' if not EXISTS(SELECT * FROM sysobjects  WHERE name= ''CheckInLog'' ) ' );
  sqlstr.Add  ('    begin');
  sqlstr.Add  ('    create table CheckInLog ( ');
  sqlstr.Add  ('    No  int  PRIMARY KEY IDENTITY(1,1), ');
  sqlstr.Add  ('    Time            nvarchar(16)    ,');
  sqlstr.Add  ('    AD_Name         nvarchar(50)    ,');
  sqlstr.Add  ('    Local_Name      nvarchar(50)    ,');
  sqlstr.Add  ('    DB_IP           nvarchar(15)    ,');
  sqlstr.Add  ('    DB_dbo          nvarchar(MAX)   ,');
  sqlstr.Add  ('    SQL             nvarchar(MAX)   ,');
  sqlstr.Add  ('    Setting         nvarchar(MAX)   ,');
  sqlstr.Add  ('    SavePath        nvarchar(MAX))  ');
  sqlstr.Add  ('    INSERT INTO '+ myinifile.readstring('SQL','SQLLOG','') + 'VALUES (' + QuotedStr(mmoLog.Lines[0]) +','+ QuotedStr(mmoLog.Lines[1]) +','+QuotedStr(mmoLog.Lines[2]) +','+QuotedStr(mmoLog.Lines[3]) +',' + QuotedStr(mmoLog.Lines[4]) + ',' + QuotedStr(mmoLog.Lines[5]) +',' + QuotedStr(mmoLog.Lines[6]) +',' + QuotedStr(mmoLog.Lines[7])+ ') end'  ) ;
  sqlstr.Add  (' else ');
  sqlstr.Add  ('    INSERT INTO '+ myinifile.readstring('SQL','SQLLOG','') + 'VALUES (' + QuotedStr(mmoLog.Lines[0]) +','+ QuotedStr(mmoLog.Lines[1]) +','+QuotedStr(mmoLog.Lines[2]) +','+QuotedStr(mmoLog.Lines[3]) +',' + QuotedStr(mmoLog.Lines[4]) + ',' + QuotedStr(mmoLog.Lines[5]) +',' + QuotedStr(mmoLog.Lines[6]) +',' + QuotedStr(mmoLog.Lines[7]) + ')'  ) ;
  qryExport.Close;
  qryExport.SQL.Text := sqlstr.text ;
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

  myinifile.Free;
  Form1.KillExcelTask;

end;

procedure TForm4.btnFileClick(Sender: TObject);
begin
  if Form1.dlgOpen1.Execute then   edtFile.Text := Form1.dlgOpen1.FileName ;
end;

procedure TForm4.btnKillTaskClick(Sender: TObject);
begin
  Form1.KillExcelTask;
  ShowMessage('已關閉Excel');
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
  WordFileName,S_Photo,col,row: string;
  WordApp, WordDoc, myRange : Variant;
  i,chk : Integer;
begin
  chk:=0;
  Form1.KillExcelTask;

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION 'SQL語法'}
//  STR := sqlstr.Text;
  qryExport.Close;
  qryExport.SQL.Text := myinifile.readstring('SQL','SQLPreview','');
  qryExport.Open;
  qryExport.First;
  {$ENDREGION}

  {$REGION '預覽'}

  WordFileName := Form1.dlgOpen1.FileName;
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
   begin
    if strngrdCheckList.Cells[2,i] <> '' then
    begin
      myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,i], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,i]).AsString, Replace:=2);
      Application.ProcessMessages;
    end;
   end;

   if chkPic.Checked then
    begin
      S_Photo := edtPic.Text + FieldByName(myinifile.readstring('Photo','pic','')).AsString ;
      col :=   edtCols.Text;
      row :=   edtRows.Text ;
      StrToInt(col);
      StrToInt(row);
      WordDoc.Tables.Item(1).Cell(col,row).range.InlineShapes.AddPicture(S_Photo,false,true);
    end;
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
var
i,temp1 : Integer;
temp,IP,IPtemp,DB,DBtemp : string ;
sock : TIpSocket;
begin
sock := TIpSocket.Create(self);  //獲取本機名稱
if SelectDirectory('請選擇設定檔目錄', '', folder) then

{$REGION '設定DB連接'}
  try
    Form1.con1.Connected := false;
    Form1.con1.ConnectionString := 'FILE NAME='+folder+'\'+ 'db.udl';
    Form1.con1.Provider := folder+'\'+'db.udl';
    Form1.con1.Connected := true;
  except
    showmessage('連線錯誤,請檢查.udl設定是否正確');
    EXIT;
  end;
{$ENDREGION}

{$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
{$ENDREGION}

{$REGION '設定欄位'}
  Form1.qry1.Close;
  form1.qry1.SQL.Text := myinifile.readstring('SQL','SQLFormShow ','');
  form1.qry1.Open;
  form1.qry1.First;
  with form1.qry1 do
  for I := 0 to  form1.qry1.FieldCount -1 do
  begin
    dbgrd1.Columns[i].Width := 50;
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
  strngrdCheckList.RowCount:=form1.qry1.FieldCount+1 ;
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

 temp := Form1.con1.ConnectionString;
 //擷取IP字串
 temp1:= POS('e=', temp);
 IPtemp:= copy(temp, (temp1)+2,15);       //擷取IP 最大長度字串
 for i:=1 to Length(IPtemp) do            //去除多餘字串
 if ( (IPtemp[i] = '.') or (IPtemp[i]  in ['0'..'9'])) then  IP := IP+ IPtemp[i] ;

 //擷取使用DB資料表字串
 temp1:= POS('g=', temp);
 DBtemp:= copy(temp, (temp1)+2,20);       //擷取IP 最大長度字串
 for i:=1 to Length(DBtemp) do            //去除多餘字串
 if not (DBtemp[i] = ';') then DB := DB + DBtemp[i]
 else break;

{$ENDREGION}

{$REGION 'Log紀錄'}
mmoLog.Lines.Add(FormatDateTime('yyyy/mm/dd'+#32+'hh:mm', Now)) ;      //取得當前時間
mmoLog.Lines.Add(LoginCheck.edtad.Text) ;                              //AD 使用者名稱
mmoLog.Lines.Add(sock.localHostName) ;                                 //本機名稱
mmoLog.Lines.Add(IP) ;                                                 //DB IP
mmoLog.Lines.Add(DB) ;                                                 //DB 資料表
mmoLog.Lines.Add(myinifile.readstring('SQL','SQLExport ','')) ;        //使用語法
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

procedure TForm4.strngrdCheckListClick(Sender: TObject);
begin
lblRows.Caption := '當前行數：'+ IntToStr(strngrdCheckList.Selection.top) ;
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

