unit Rep;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, CheckLst, Gauges, Mask, DBCtrls,IniFiles ,
  DBClient, ExtCtrls,TlHelp32,Excel2000,ComObj,ADODB,FileCtrl,StrUtils,ComCtrls,Math,
  GIFImg, Buttons, jpeg,shellapi, pngimage,System.Win.ScktComp;



type

//  TMyThread = class(TThread)
//  protected
//  procedure Execute; override;
//  end;

  TRepExport = class(TForm)
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
    pnlExit: TPanel;
    imgStop: TImage;
    btnKillTask: TSpeedButton;
    btnChgSet: TBitBtn;
    btnOpenWord: TBitBtn;
    btnOpenFile: TBitBtn;
    blnhnt1: TBalloonHint;
    mmoTemp: TMemo;
    btnThead: TBitBtn;
    con1: TADOConnection;
    qry1: TADOQuery;
    dlgOpen1: TOpenDialog;
    dlgOpendlgudl: TOpenDialog;
    dlgSave1: TSaveDialog;
    btnClose: TBitBtn;
    ClientSocket1: TClientSocket;
    cbbTitle1: TComboBox;
    cbbTitle2: TComboBox;
    qryREP: TADOQuery;
    chklstTitle: TCheckListBox;
    lstchk: TListBox;
    btnup: TButton;
    btndown: TButton;
    cbbTitle3: TComboBox;
    edtCount: TEdit;
    edtRow: TEdit;
    edtTotal: TEdit;
    lstOrderBy: TListBox;
    btnup2: TButton;
    btndowm2: TButton;
    btnDel: TBitBtn;
    btnAdd: TBitBtn;
    cbbTitle4: TComboBox;
    procedure ChoiceTest(Sender: TObject);
    procedure FormShow(Sender: TObject) ;
    procedure btnKillTaskClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure btnPicClick(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure btnChgSetClick(Sender: TObject);
    procedure btnOpenWordClick(Sender: TObject);
    procedure btnOpenFileClick(Sender: TObject);
    procedure btnTheadClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure cbbTitle1Change(Sender: TObject);
    procedure btnupClick(Sender: TObject);
    procedure btndownClick(Sender: TObject);
    procedure cbbTitle2Change(Sender: TObject);
    procedure cbbTitle3Change(Sender: TObject);
    procedure edtRowChange(Sender: TObject);
    procedure btnup2Click(Sender: TObject);
    procedure btndowm2Click(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure chklstTitleClickCheck(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure edtCountKeyPress(Sender: TObject; var Key: Char);
    procedure edtRowKeyPress(Sender: TObject; var Key: Char);
    procedure cbbTitle4Change(Sender: TObject);
  private
    { Private declarations }
  public
//     procedure GetRepExFile(const ExName, SavePath: String);
    { Public declarations }
  end;

var
  RepExport: TRepExport;
  Filename,dirpath,folder :string;
  Myinifile:Tinifile;
  t,m,h,Cnt,ttTime,LoginInCheck,SelectItem : Integer ;
  times : longint = 0 ;
  timeh : longint = 0 ;
  timem : longint = 0 ;
  SqlFS,SqlEP : string ;
  clearchk : Boolean;
implementation
uses UnitMyThread, NATL, Login;
{$R *.dfm}


procedure TRepExport.btnTheadClick(Sender: TObject);
var
  WordFileName,LogTemp,find,temp: string;
  SavePath,chkstr,chkstr2,chkstr3,chkstr4,strtemp,SQLExport,orderby : string;
  WordApp, WordDoc, myRange, vSaveNone : Variant;
  i,p,x,y,count,row,total,chk,temp1: Integer;
  sqlstr : TStringList;
  SaveTxt:TextFile;
begin
 y:= 1;
 total := 0;
 chk := 0 ;
 row := StrToInt(edtRow.Text);
 SavePath := edtPath.Text;
 sqlstr := TStringList.Create;
 count := StrToInt( edtCount.Text  );
 ShowMessage('資料內含個資，請妥善保存！');

  try
    {$REGION '準備ini檔案資料'}
    Filename:=folder+'\'+'Setting.ini';
    myinifile:=Tinifile.Create(filename);
    {$ENDREGION}

    {$REGION 'SQL語法'}

    //加入Order By 語法
    orderby := 'order by ';
    if cbbTitle4.ItemIndex <> -1 then
      orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+ ', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] + ', ' + cbbTitle3.Items[cbbTitle3.ItemIndex]+', ' + cbbTitle4.Items[cbbTitle4.ItemIndex]+', '
    else if cbbTitle3.ItemIndex <> -1 then
      orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+ ', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] + ', ' + cbbTitle3.Items[cbbTitle3.ItemIndex]+', '
    else if cbbTitle2.ItemIndex <> -1 then
      orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] +', '
    else
      orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+', ';
    for I := 0 to lstOrderBy.Count -1 do
      orderby :=  orderby + lstOrderBy.Items[i] + ', ';
    orderby := Copy(orderby,0,Length(orderby)-2);

    qryExport.Close;
    qryExport.SQL.CLEAR;
    for p := 1 to strngrdCheckList.RowCount  do     //儲存有搜尋條件的欄位
    if strngrdCheckList.Cells[3,p] <> '' then
      begin
        find := find + #32'And'#32 + strngrdCheckList.Cells[0,p] + #32 + strngrdCheckList.Cells[3,p] + #32 + strngrdCheckList.Cells[4,p];
        chk:=1;
      end;
    if chk = 0 then qryExport.SQL.Text := SqlEP + orderby
    else if (chk = 1) and ((LeftStr(UpperCase(SqlEP),6) = 'SELECT') ) then qryExport.SQL.Text := SqlEP + find  + orderby
    else
      begin
        if LeftStr(UpperCase(SqlEP),6) <> 'SELECT' then
        temp := SqlEP;
        temp1:= POS(',', temp);
        strtemp:= copy(temp,0,temp1-1);       //擷取字串
        for i:=1 to Length(strtemp) do            //去除多餘字串
        if not (strtemp[i] = ',') then SQLExport := SQLExport + strtemp[i]
        else break;
        strtemp := '';
        qryExport.SQL.Text := SQLExport +','+QuotedStr('where 1=1 '+find);
      end;
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

    {$REGION '套印'}
    with qryExport do
    begin
      tmr1.Enabled:=true;  //開始計時
      //判斷並儲存有值欄位
      mmoTemp.Clear;
      for p := 1 to FieldCount  do
       if strngrdCheckList.Cells[2,p] <> '' then
        begin
          mmoTemp.lines.add( IntToStr(p) );    //將有值的欄位存在mmoLog
          LogTemp := LogTemp+'['+strngrdCheckList.Cells[0,p]+'/'+strngrdCheckList.Cells[2,p]  +']'+','; //寫入LOG
        end;

      //套印
      for i:= 1 to RecordCount  do
       begin
          if chkExit.Checked then  Break; //判斷是否離開
          WordApp.Visible := False;
          WordApp.Application.DisplayAlerts := False;
          WordDoc := WordApp.Documents.Open(WordFileName);
          myRange := WordDoc.Content;
          chk := 0; //判斷是否存檔
          //判斷選取幾個欄位
          if cbbTitle4.ItemIndex <> -1 then
            begin
              if (( chkstr2 <> qryExport.FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString ) or ( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString ) or ( chkstr3 <> FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString )or ( chkstr4 <> FieldByName(cbbTitle4.Items[cbbTitle4.ItemIndex]).AsString ) )
                and ( i > 1 ) then chk := 4;
              chkstr4 := FieldByName(cbbTitle4.Items[cbbTitle4.ItemIndex]).AsString;
              chkstr3 := FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString;
              chkstr2 := FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString;
            end
          else if cbbTitle3.ItemIndex <> -1 then
            begin
              if (( chkstr2 <> qryExport.FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString ) or ( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString ) or ( chkstr3 <> FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString ) )
                and ( i > 1 ) then chk := 3;
              chkstr3 := FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString;
              chkstr2 := FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString;
            end
          else if cbbTitle2.ItemIndex <> -1 then
            begin
              if (( chkstr2 <> qryExport.FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString ) or (( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString )) ) and ( i > 1 ) then
                chk := 2;
              chkstr2 := FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString;
            end
          else if cbbTitle1.ItemIndex <> -1 then
            if  ( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString ) and ( i > 1 ) then
            chk := 1;

         //判斷是否存檔過
         if chk = 0 then
            begin
              //插入資料
              for x := 0 to lstchk.Count-1 do
                WordDoc.Tables.Item(y).Cell(row,x+1).Range.Text := FieldByName(lstchk.Items[x]).AsString;
              //判斷是否換行
              if i = count then
                begin
                  Inc(y);
                  count := i + StrToInt( edtCount.Text) ;
                end
              else if i > 1 then  WordDoc.Tables.Item(y).Range.Rows.Add(EmptyParam);
              Inc(row);   //當前行數
              Inc(total); //記錄總筆數
            end
         else
            begin
              //套印資料
              for x := 0 to  mmoTemp.lines.Count-1 do
                myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,StrToInt(mmoTemp.lines[x])], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,StrToInt(mmoTemp.lines[x])] ).AsString, Replace:=2);
              myRange.Find.Execute(FindText:=edtTotal.Text,ReplaceWith:=total , Replace:=2);
              //存檔
              if chk = 1 then
                begin
                  if chkWord.Checked  then  WordDoc.SaveAs(SavePath + chkstr +'.docx');
                  WordDoc.ExportAsFixedFormat(SavePath + chkstr + '.pdf',17);
                end
              else if chk = 2 then
                begin
                  if chkWord.Checked  then  WordDoc.SaveAs(SavePath + chkstr +'-'+chkstr2 +'.docx');
                  WordDoc.ExportAsFixedFormat(SavePath + chkstr +'-'+chkstr2+ '.pdf',17);
                end
              else if chk = 3 then
                begin
                  if chkWord.Checked  then  WordDoc.SaveAs(SavePath + chkstr +'-'+chkstr2 +'-'+chkstr3+'.docx');
                  WordDoc.ExportAsFixedFormat(SavePath + chkstr +'-'+chkstr2 +'-'+chkstr3+ '.pdf',17);
                end
              else if chk = 4 then
                begin
                  if chkWord.Checked  then  WordDoc.SaveAs(SavePath + chkstr +'-'+chkstr2+'-'+chkstr3+'-'+chkstr4 +'.docx');
                  WordDoc.ExportAsFixedFormat(SavePath + chkstr +'-'+chkstr2 +'-'+chkstr3+'-'+chkstr4 +'.pdf',17);
                end ;
              WordApp.Documents.close(vSaveNone);
              //存檔後參數重置
              total := 0 ;
              Y := 1;
              count := i + StrToInt( edtCount.Text) ;
              row := StrToInt(edtRow.Text);
            end;
            chkstr  := FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString;
            Inc(Cnt);
            Gauge2.Progress:= i;
            lblCNumber.Caption := '筆數：'+ IntToStr(qryExport.RecordCount)+'/'+ IntToStr(Cnt);
            Next;
       end;
    end;
    {$ENDREGION}

  finally
    tmr1.Enabled:=false;
    mmoLog.Lines.Add(SavePath) ;
    WordApp.Quit;
    WordApp:=Unassigned;
    edtPath.Enabled := true;
    edtFile.Enabled := true;
    edtCount.Enabled := True ;
    edtRow.Enabled := true ;
    Gauge2.Progress:= Gauge2.MaxValue;
   if MessageDlg('匯出完成',mtInformation,[mbYes],0)=mrYes then lblTime.Caption := '' ;
  end;

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
end;

procedure TRepExport.btnup2Click(Sender: TObject);
var
i : Integer;
begin
  for I := 0 to lstOrderBy.Count -1 do
   begin
    if lstOrderBy.Selected[i] and (i >= 1) then
      begin
        lstOrderBy.Items.Move(i,i-1);
        Break;
      end;
   end;
  if  lstOrderBy.Selected[0] = True then Exit;
  lstOrderBy.Selected[i-1] := True;
end;

procedure TRepExport.btnupClick(Sender: TObject);
var
i: Integer;
begin
  for I := 0 to lstchk.Count -1 do
   begin
    if lstchk.Selected[i] and (i >= 1) then
      begin
        lstchk.Items.Move(i,i-1);
        Break;
      end;
   end;
  if  lstchk.Selected[0] = True then Exit;
  lstchk.Selected[i-1] := True;
end;

procedure TRepExport.cbbTitle1Change(Sender: TObject);
var
i : Integer;
begin
cbbTitle2.Enabled := True;
  for i := 0 to cbbTitle1.Items.Count -1 do
    cbbTitle2.Items.Add(cbbTitle1.Items[i]);
edtCount.Enabled := True;
edtRow.Enabled := True;
end;

procedure TRepExport.cbbTitle2Change(Sender: TObject);
var
i : Integer;
begin
cbbTitle3.Enabled := True;
  for i := 0 to cbbTitle1.Items.Count -1 do
    cbbTitle3.Items.Add(cbbTitle1.Items[i]);
  if cbbTitle2.ItemIndex = 0 then
    begin
      cbbTitle2.Enabled := false;
      cbbTitle2.ItemIndex := -1;
      cbbTitle3.Enabled := false;
      cbbTitle3.ItemIndex := -1;
    end;
end;

procedure TRepExport.cbbTitle3Change(Sender: TObject);
var
i : Integer;
begin
cbbTitle4.Enabled := True;
  for i := 0 to cbbTitle1.Items.Count -1 do
    cbbTitle4.Items.Add(cbbTitle1.Items[i]);
  if cbbTitle3.ItemIndex = 0 then
    begin
      cbbTitle3.ItemIndex := -1 ;
      cbbTitle3.Enabled := false;
      cbbTitle4.ItemIndex := -1 ;
      cbbTitle4.Enabled := false;
    end;
end;


procedure TRepExport.cbbTitle4Change(Sender: TObject);
begin
  if cbbTitle4.ItemIndex = 0 then
    begin
      cbbTitle4.ItemIndex := -1 ;
      cbbTitle4.Enabled := false;
    end;
end;

procedure TRepExport.chklstTitleClickCheck(Sender: TObject);
var
i,y : Integer;
chk : Boolean;
begin
chk := False ;
if clearchk = true then
  begin
    lstchk.Clear;
    lstOrderBy.Clear;
    clearchk := False ;
  end;

  for I := 0 to chklstTitle.Count -1 do
    if chklstTitle.Checked[i] then
      begin
        if lstchk.Items.Count = 0 then lstchk.Items.Add(chklstTitle.Items[i])
        else
          begin
            for y := 0 to lstchk.Count -1  do
              if chklstTitle.Items[i] = lstchk.Items[y] then
              begin
                chk := True;
                Break;
              end;
            if chk = False then
              lstchk.Items.Add(chklstTitle.Items[i]);
          end;
        chk := false;
      end
    else
      for y := 0 to lstchk.Count -1  do
        if chklstTitle.Items[i] = lstchk.Items[y] then
              begin
                lstchk.Selected[y] := True;
                lstchk.DeleteSelected;
                Break;
              end;

  if lstchk.Items.Count > 0 then
    lstchk.Enabled := True
  else
    lstchk.Enabled := False ;

  for I := 0 to chklstTitle.Count -1 do
    if chklstTitle.Checked[i] then
      begin
        if lstOrderBy.Items.Count = 0 then lstOrderBy.Items.Add(chklstTitle.Items[i])
        else
          begin
            for y := 0 to lstOrderBy.Count -1  do
              if chklstTitle.Items[i] = lstOrderBy.Items[y] then
              begin
                chk := True;
                Break;
              end;
            if chk = False then
              lstOrderBy.Items.Add(chklstTitle.Items[i]);
          end;
        chk := false;
      end
    else
      for y := 0 to lstOrderBy.Count -1  do
        if chklstTitle.Items[i] = lstOrderBy.Items[y] then
              begin
                lstOrderBy.Selected[y] := True;
                lstOrderBy.DeleteSelected;
                Break;
              end;
  if lstOrderBy.Items.Count > 0 then
    lstOrderBy.Enabled := True
  else
    lstOrderBy.Enabled := False ;

btnExport.Enabled := True;
end;

procedure TRepExport.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('名冊套印');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
end;

procedure TRepExport.btnAddClick(Sender: TObject);
begin
lstOrderBy.Items.Add(chklstTitle.Items[chklstTitle.ItemIndex]);
end;

procedure TRepExport.btnChgSetClick(Sender: TObject);
begin
  ChoiceTest(nil);
end;

procedure TRepExport.btnOpenFileClick(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar(edtPath.Text), nil, nil, SW_SHOW);
end;

procedure TRepExport.btnOpenWordClick(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar(edtFile.Text), nil, nil, SW_SHOW);
end;

procedure TRepExport.btnDelClick(Sender: TObject);
begin
lstOrderBy.DeleteSelected;
end;

procedure TRepExport.btnDirpath2Click(Sender: TObject);
begin
  repeat
      SelectDirectory('請選擇存檔目錄', '', DirPath); //選擇目錄
      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
        Exit;
    until DirPath <> '';
    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
  edtPath.Text := DirPath;
end;

procedure TRepExport.btndowm2Click(Sender: TObject);
var
i : Integer;
begin
  for I := 0 to lstOrderBy.Count -2 do
   begin
    if lstOrderBy.Selected[i]  then
      begin
        lstOrderBy.Items.Move(i,i+1);
        Break;
      end;
   end;
   if i <= lstOrderBy.Count -2 then
     lstOrderBy.Selected[i+1] := True;
end;

procedure TRepExport.btndownClick(Sender: TObject);
var
i : Integer;
begin
  for I := 0 to lstchk.Count -2 do
   begin
    if lstchk.Selected[i]  then
      begin
        lstchk.Items.Move(i,i+1);
        Break;
      end;
   end;
   if i <= lstchk.Count -2 then
     lstchk.Selected[i+1] := True;
end;

procedure TRepExport.btnExportClick(Sender: TObject);
begin
if edtRow.text = '' then
  begin
    ShowMessage('請輸入起始欄位');
    Exit;
  end;
if edtCount.Text = '' then
  begin
    ShowMessage('請輸入換頁數');
    Exit;
  end;

t:= 0;
m:= 0;
h:=0;
edtPath.Enabled := false;
edtFile.Enabled := false;
edtCount.Enabled := False ;
edtRow.Enabled := False ;
MyThread(btnTheadClick,nil);
end;

procedure TRepExport.btnFileClick(Sender: TObject);
begin
  if dlgOpen1.Execute then   edtFile.Text := dlgOpen1.FileName ;
end;

procedure TRepExport.btnKillTaskClick(Sender: TObject);
begin
  Form1.KillWordTask;
  ShowMessage('已關閉Word');
end;

procedure TRepExport.btnPicClick(Sender: TObject);
begin
  repeat
      SelectDirectory('請選擇圖片目錄', '', DirPath); //選擇目錄
      if (DirPath = '') and (MessageDlg('確定取消嗎?',mtcustom,[mbYes]+[mbNo],0) = 6) then
        Exit;
    until DirPath <> '';
    if RightStr(DirPath,1) <> '\' then	DirPath := DirPath + '\';  //檢查字尾是否為'/'符號
end;

procedure TRepExport.btnPreviewClick(Sender: TObject);
var
  WordFileName,LogTemp,temp,strtemp: string;
  chkstr,chkstr2,chkstr3,chkstr4,orderby,find,SQLExport : string;
  WordApp, WordDoc, myRange : Variant;
  i,p,x,y,row,total,temp1,chk: Integer;
begin
  form1.KillExcelTask;
  y:= 1;
  total := 0;
  mmoTemp.Clear;
  row := StrToInt(edtRow.Text);
  {$REGION 'SQL語法'}
  //加入Order By 語法
  orderby := 'order by ';
  if cbbTitle4.ItemIndex <> -1 then
    orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+ ', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] + ', ' + cbbTitle3.Items[cbbTitle3.ItemIndex]+', ' + cbbTitle4.Items[cbbTitle4.ItemIndex]+', '
  else if cbbTitle3.ItemIndex <> -1 then
    orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+ ', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] + ', ' + cbbTitle3.Items[cbbTitle3.ItemIndex]+', '
  else if cbbTitle2.ItemIndex <> -1 then
    orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+', '+ cbbTitle2.Items[cbbTitle2.ItemIndex] +', '
  else
    orderby := orderby + cbbTitle1.Items[cbbTitle1.ItemIndex]+', ';
  for I := 0 to lstOrderBy.Count -1 do
    orderby :=  orderby + lstOrderBy.Items[i] + ', ';
  orderby := Copy(orderby,0,Length(orderby)-2);
  for p := 1 to strngrdCheckList.RowCount  do     //儲存有搜尋條件的欄位
    if strngrdCheckList.Cells[3,p] <> '' then
      begin
        find := find + #32'And'#32 + strngrdCheckList.Cells[0,p] + #32 + strngrdCheckList.Cells[3,p] + #32 + strngrdCheckList.Cells[4,p];
        chk:=1;
      end;
    if chk = 0 then qryExport.SQL.Text := SqlEP + orderby
    else if (chk = 1) and ((LeftStr(UpperCase(SqlEP),6) = 'SELECT') ) then qryExport.SQL.Text := SqlEP + find  + orderby
    else
      begin
        if LeftStr(UpperCase(SqlEP),6) <> 'SELECT' then
        temp := SqlEP;
        temp1:= POS(',', temp);
        strtemp:= copy(temp,0,temp1-1);       //擷取字串
        for i:=1 to Length(strtemp) do            //去除多餘字串
        if not (strtemp[i] = ',') then SQLExport := SQLExport + strtemp[i]
        else break;
        strtemp := '';
        qryExport.SQL.Text := SQLExport +','+QuotedStr('where 1=1 '+find);
      end;

  qryExport.Close;
  qryExport.SQL.Clear;
  qryExport.SQL.Text := SqlEP;
  qryExport.Open;
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

  with qryExport do
    begin
      //判斷有值欄位
      for p := 1 to FieldCount  do
       if strngrdCheckList.Cells[2,p] <> '' then
        begin
          mmoTemp.lines.add( IntToStr(p) );    //將有值的欄位存在mmoLog
          LogTemp := LogTemp+'['+strngrdCheckList.Cells[0,p]+'/'+strngrdCheckList.Cells[2,p]  +']'+',';
        end;

      //套印
      for i:= 1 to RecordCount  do
       begin
        if chkExit.Checked then  Break;
        WordApp.Visible := true;
        WordApp.Application.DisplayAlerts := true;
        WordDoc := WordApp.Documents.Open(WordFileName);
        myRange := WordDoc.Content;
        //判斷是否存檔
        if cbbTitle3.ItemIndex <> -1 then
          begin
            if (( chkstr2 <> qryExport.FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString ) or ( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString ) or ( chkstr3 <> FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString ) )
            and ( i > 1 ) then
             begin
                myRange.Find.Execute(FindText:=edtTotal.Text,ReplaceWith:=total , Replace:=2);
                Exit;
              end;
            chkstr3 := FieldByName(cbbTitle3.Items[cbbTitle3.ItemIndex]).AsString;
            chkstr2 := FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString;
            if i > 1 then
              WordDoc.Tables.Item(y).Range.Rows.Add(EmptyParam);
          end
        else if cbbTitle2.ItemIndex <> -1 then
          begin
            if (( chkstr2 <> qryExport.FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString ) or (( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString )) ) and ( i > 1 ) then
              begin
                myRange.Find.Execute(FindText:=edtTotal.Text,ReplaceWith:=total , Replace:=2);
                Exit;
              end;
            chkstr2 := FieldByName(cbbTitle2.Items[cbbTitle2.ItemIndex]).AsString;
            if i > 1 then
             WordDoc.Tables.Item(y).Range.Rows.Add(EmptyParam);
          end
        else if cbbTitle1.ItemIndex <> -1 then
          begin
            if( chkstr <> FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString ) and ( i > 1 )  then
              begin
                myRange.Find.Execute(FindText:=edtTotal.Text,ReplaceWith:=total , Replace:=2);
                Exit;
              end;
            if i > 1 then
              WordDoc.Tables.Item(y).Range.Rows.Add(EmptyParam);
          end;

          //套印資料
          if i = 1 then
          for x := 0 to  mmoTemp.lines.Count-1 do
            myRange.Find.Execute(FindText:=strngrdCheckList.Cells[2,StrToInt(mmoTemp.lines[x])], ReplaceWith:=FieldByName(strngrdCheckList.Cells[0,StrToInt(mmoTemp.lines[x])] ).AsString, Replace:=2);

          //插入資料
          for x := 0 to lstchk.Count-1 do
            WordDoc.Tables.Item(y).Cell(row,x+1).Range.Text := FieldByName(lstchk.Items[x]).AsString;

          chkstr  := FieldByName(cbbTitle1.Items[cbbTitle1.ItemIndex]).AsString;
          Inc(total);
          Inc(row);
          Next;
        end;

    end;
   {$ENDREGION}
end;


procedure TRepExport.FormShow(Sender: TObject);
begin
ChoiceTest(nil);
end;

procedure TRepExport.ChoiceTest(Sender: TObject);
var
i,temp1,temp2 : Integer;
strtemp,temp,IP,DB : string ;
sock : TClientSocket;
begin
  clearchk := True ;
  SqlEP := '';
  SqlFS := '';
  cbbTitle1.Clear;
  cbbTitle2.Clear;
  cbbTitle3.Clear;
  cbbTitle4.Clear;
  cbbTitle2.Items.Add('');
  cbbTitle3.Items.Add('');
  cbbTitle4.Items.Add('');
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
    cbbTitle1.Items.Add(dbgrd1.Columns[i].FieldName);
    chklstTitle.Items.Add(dbgrd1.Columns[i].FieldName);
    strngrdCheckList.Cells[1,i+1]:=myinifile.readstring('Help',('help'+ IntToStr(i+1)),'');
  end;
  {$ENDREGION}

  {$REGION 'StringGrid'}
  strngrdCheckList.Cells[0,0] := #32#32#32#32'---【資料欄位】---' ;
  strngrdCheckList.Cells[1,0] := #32#32#32#32'---【欄位說明】---' ;
  strngrdCheckList.Cells[2,0] := #32#32#32#32'---【取代文字】---' ;
  strngrdCheckList.ColWidths[0] := 200;
  strngrdCheckList.ColWidths[1] := 180;
  strngrdCheckList.ColWidths[2] := 180;
  strngrdCheckList.RowCount:=qry1.FieldCount+1 ;
  {$ENDREGION}

  {$REGION '設定預設路徑'}
  if not directoryExists(folder+'\'+'套印資料夾') then  //判斷此資料夾是否存在
  CreateDir(folder+'\'+'套印資料夾');                 //建立資料夾
  edtPath.Text:= folder+'\'+'套印資料夾'+'\';
  dirpath:= edtPath.Text   ;
  edtFile.Text :=folder+'\'+'套印範本.docx' ;
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
procedure TRepExport.edtCountKeyPress(Sender: TObject; var Key: Char);
begin
if Key = #13 then edtRow.SetFocus;
end;

procedure TRepExport.edtRowChange(Sender: TObject);
begin
chklstTitle.Enabled := True;
end;

procedure TRepExport.edtRowKeyPress(Sender: TObject; var Key: Char);
begin
if Key = #13 then edtTotal.SetFocus;
end;

procedure TRepExport.tmr1Timer(Sender: TObject);
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

