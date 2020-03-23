unit Grade;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, CheckLst, Gauges, Mask, DBCtrls,IniFiles ,
  DBClient, ExtCtrls,TlHelp32,Excel2000,ComObj,ADODB,FileCtrl,StrUtils,ComCtrls,Math  ;



type

  TMyThread = class(TThread)
  protected
  procedure Execute; override;
  end;

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
    procedure FormShow(Sender: TObject) ;
    procedure btnKillTaskClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure chkPicClick(Sender: TObject);
    procedure btnPicClick(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);


  private
    { Private declarations }
  public

    { Public declarations }
  end;

var
  Form4: TForm4;
  Filename,dirpath :string;
  Myinifile:Tinifile;
  t,m,h,Cnt,ttTime : Integer ;
  times : longint = 0 ;
  timeh : longint = 0 ;
  timem : longint = 0 ;
implementation
uses NATL;

{$R *.dfm}


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
  S_Photo,WordFileName,find,temp: string;
  WordApp, WordDoc, myRange, vSaveNone : Variant;
  i,p,chk,x: Integer;
begin
 t:= 0;
 m:= 0;
 h:=0;

  {$REGION '準備ini檔案資料'}
  Filename:=ExtractFilePath(Paramstr(0))+'Setting.ini';
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
    end;
  {$ENDREGION}

  {$REGION '套印'}
  for i:= 1 to RecordCount  do
   begin
    tmr1.Enabled:=true;
    if chkExit.Checked then Break  ;
//    if FileExists(DirPath + FieldByName(myinifile.readstring('SQL','SaveField','')).AsString+'.pdf') then
//        if  (MessageDlg('檔案已存在，是否覆蓋',mtcustom,[mbYes]+[mbNo],0) = 7) then break ;
    WordApp.Visible := false;
    WordApp.Application.DisplayAlerts := False;

    for x := 0 to  mmoSql.lines.Count-1 do   //套印資料
    begin
    WordDoc := WordApp.Documents.Open(WordFileName);
    myRange := WordDoc.Content;
    temp := mmoSql.lines[x];
    myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,StrToInt(temp)], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,StrToInt(temp)] ).AsString, Replace:=2);
    Application.ProcessMessages;
    end;

   if chkPic.Checked then   //套印圖片
    begin
    WordDoc := WordApp.Documents.Open(WordFileName);
    myRange := WordDoc.Content;
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
  end;
finally
  WordApp.Quit;
  WordApp:=Unassigned;
  tmr1.Enabled:=false;
end;

  if qryExport.RecordCount < 1  then
    begin
      ShowMessage('未找到符合條件的資料');
      exit;
    end;
   Gauge2.Progress:= Gauge2.MaxValue;
   if MessageDlg('匯出完成',mtInformation,[mbYes],0)=mrYes then lblTime.Caption := '' ;

end;

procedure TForm4.btnFileClick(Sender: TObject);
begin
  if Form1.dlgOpen1.Execute then   edtFile.Text := Form1.dlgOpen1.InitialDir ;
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
  WordFileName,STR,S_Photo,col,row: string;
  WordApp, WordDoc, myRange : Variant;
  i : Integer;
begin

  {$REGION '準備ini檔案資料'}
  Filename:=ExtractFilePath(Paramstr(0))+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION 'SQL語法'}
//  STR := sqlstr.Text;
  qryExport.Close;
  qryExport.SQL.Text := myinifile.readstring('SQL','SQLPreview','');
  qryExport.Open;
  qryExport.First;
  {$ENDREGION}

  {$REGION '匯出成績單'}

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
i : Integer;
begin

{$REGION '準備ini檔案資料'}
  Filename:=ExtractFilePath(Paramstr(0))+'Setting.ini';
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
    Application.ProcessMessages;
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
  if not directoryExists(ExtractFilePath(Application.ExeName)+'\'+'套印資料夾') then  //判斷此資料夾是否存在
  CreateDir(ExtractFilePath(Application.ExeName)+'\'+'套印資料夾');                 //建立資料夾

  edtPath.Text := ExtractFilePath(Application.Exename);
  dirpath :=  edtPath.Text+'\'+'套印資料夾'+'\' ;
  Form1.dlgOpen1.InitialDir := ExtractFilePath(Application.ExeName);
  Form1.dlgOpen1.FileName := ExtractFilePath(Application.ExeName)+'\'+'套印範本.docx';
  edtFile.Text := Form1.dlgOpen1.InitialDir ;
  edtPic.Text := ExtractFilePath(Application.Exename) + 'Photo\';
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


procedure TMyThread.Execute;
begin

end;

end.

