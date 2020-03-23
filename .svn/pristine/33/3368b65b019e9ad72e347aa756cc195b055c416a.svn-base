unit Table;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids,FileCtrl,IniFiles, DB, Buttons,ComCtrls,
  ADODB, ExtCtrls,StrUtils,ClipBrd, Excel2000,ComObj, Gauges, jpeg;

type
  TFormTable = class(TForm)
    dbgrd1: TDBGrid;
    mmoTemp: TMemo;
    ds1: TDataSource;
    con1: TADOConnection;
    qry1: TADOQuery;
    btnChgSet: TBitBtn;
    btnSearch: TBitBtn;
    TimeTitle: TComboBox;
    btnNext: TBitBtn;
    btnPrevious: TBitBtn;
    lblCurrent: TLabel;
    lblAll: TLabel;
    btnExport: TBitBtn;
    dlgSave1: TSaveDialog;
    Gauge2: TGauge;
    btnClose: TBitBtn;
    edtSearch: TEdit;
    pnl1: TPanel;
    lbl1: TLabel;
    lbl2: TLabel;
    imgBackGroup: TImage;
    procedure FormShow(Sender: TObject);
    procedure dbgrd1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnChgSetClick(Sender: TObject);
    procedure ChoiceTest(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnPreviousClick(Sender: TObject);
    procedure dbgrd1ColumnMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure dbgrd1TitleClick(Column: TColumn);
    procedure btnExportClick(Sender: TObject);
    procedure SaveToExcel(Sender: TObject);
    procedure dbgrd1CellClick(Column: TColumn);
    procedure btnCloseClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormTable: TFormTable;
  SqlEP : string;
  Reci,RecChk  : Integer;

implementation
uses NATL,UnitMyThread;
{$R *.dfm}

procedure TFormTable.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('資料核對');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
end;

procedure TFormTable.btnChgSetClick(Sender: TObject);
begin
  ChoiceTest(nil);
end;

procedure TFormTable.btnExportClick(Sender: TObject);
begin
  Form1.KillExcelTask;
  MyThread(SaveToExcel,nil);
end;

procedure TFormTable.btnNextClick(Sender: TObject);
var
  j : Integer;
begin
  if (TimeTitle.Text = '') or (edtSearch.Text = '') then ShowMessage('請選擇欄位及輸入搜尋條件')
  else
    begin
      dbgrd1.DataSource.DataSet.Next();
      for j:= dbgrd1.DataSource.DataSet.RecNo to dbgrd1.DataSource.DataSet.RecordCount do
        begin
          if qry1.FieldByName(dbgrd1.Fields[TimeTitle.ItemIndex].FieldName).AsString =  edtSearch.Text then
            begin
              dbgrd1.Fields[TimeTitle.ItemIndex].FocusControl;
              Reci := dbgrd1.DataSource.DataSet.RecNo ;
              RecChk := 1;
              break;
            end
          else dbgrd1.DataSource.DataSet.Next();
        end;
      if  j = dbgrd1.DataSource.DataSet.RecordCount+1  then
        begin
          if RecChk = 1  then
            begin
              dbgrd1.DataSource.DataSet.First;
              dbgrd1.DataSource.DataSet.MoveBy(Reci-1);
              RecChk := 0;
              ShowMessage('此為最後一筆');
            end
          else ShowMessage('無相符結果');
        end;
    end;
end;


procedure TFormTable.btnPreviousClick(Sender: TObject);
var
  j : Integer;
begin
  if (TimeTitle.Text = '') or (edtSearch.Text = '') then ShowMessage('請選擇欄位及輸入搜尋條件')
  else
    begin
      dbgrd1.DataSource.DataSet.Prior();
      for j:= dbgrd1.DataSource.DataSet.RecordCount-dbgrd1.DataSource.DataSet.RecNo to dbgrd1.DataSource.DataSet.RecordCount do
        begin
          if qry1.FieldByName(dbgrd1.Fields[TimeTitle.ItemIndex].FieldName).AsString = edtSearch.Text then
            begin
              dbgrd1.Fields[TimeTitle.ItemIndex].FocusControl ;
              Reci := dbgrd1.DataSource.DataSet.RecNo ;
              RecChk := 1;
              break;
            end
          else dbgrd1.DataSource.DataSet.Prior();
        end;
      if  dbgrd1.DataSource.DataSet.RecNo = 1  then
        begin
          if RecChk = 1  then
            begin
              dbgrd1.DataSource.DataSet.MoveBy(Reci-1);
              RecChk := 0;
              ShowMessage('此為最後一筆');
            end
          else ShowMessage('無相符結果');
        end;
    end;
end;


procedure TFormTable.btnSearchClick(Sender: TObject);
var
  j : Integer;
begin
  if (TimeTitle.Text = '') or (edtSearch.Text = '') then ShowMessage('請選擇欄位及輸入搜尋條件')
  else
    begin
      dbgrd1.DataSource.DataSet.First;
      for j:=0 to dbgrd1.DataSource.DataSet.RecordCount-1 do
        begin
          if qry1.FieldByName(dbgrd1.Fields[TimeTitle.ItemIndex].FieldName).AsString = edtSearch.Text then
            begin
              dbgrd1.Fields[TimeTitle.ItemIndex].FocusControl ;
              exit;
            end;
          dbgrd1.DataSource.DataSet.Next();
        end;
    end;
end;


procedure TFormTable.dbgrd1CellClick(Column: TColumn);
begin
lblCurrent.Caption := '當前行數：' + IntToStr( dbgrd1.DataSource.DataSet.RecNo );
end;

procedure TFormTable.dbgrd1ColumnMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
begin
  RecChk := 0;
end;

procedure TFormTable.dbgrd1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if gdSelected in State then Exit;  //隔行改變網格背景色：
    if qry1.RecNo mod 2 = 0 then
      (Sender as TDBGrid).Canvas.Brush.Color := clinfobk //定義背景顏色
    else
    (Sender as TDBGrid).Canvas.Brush.Color := RGB(189, 230, 255);  //定義背景顏色
//  定義網格線的顏色：
  dbgrd1.DefaultDrawColumnCell(Rect,DataCol,Column,State);
  with (Sender as TDBGrid).Canvas do //畫 cell 的邊框
  begin
    Pen.Color := $545454; //定義畫筆顏色(藍色)
    MoveTo(Rect.Left, Rect.Bottom); //畫筆定位
    LineTo(Rect.Right, Rect.Bottom); //畫藍色的橫線
    Pen.Color := $545454; //定義畫筆顏色(蘭色)
    MoveTo(Rect.Right, Rect.Top); //畫筆定位
    LineTo(Rect.Right, Rect.Bottom); //畫綠色
  end;
  dbgrd1.Font.Color :=  RGB(31,31,31);
end;

procedure TFormTable.dbgrd1TitleClick(Column: TColumn);
begin
  if qry1.Sort<>(Column.FieldName+' ASC') then //判断原排序方式
      qry1.Sort := Column.FieldName+' ASC'
  else
      qry1.Sort := Column.FieldName+' DESC';
end;

procedure TFormTable.FormShow(Sender: TObject);
var
  Filename,folder :string;
  Myinifile:Tinifile;
  i,temp1 : Integer;
  strtemp,temp: string ;
begin
  SqlEP := '';
  if SelectDirectory('請選擇設定檔目錄', '', folder) then
  mmoTemp.Lines.LoadFromFile(folder+'\SqlSetting.ini');

  {$REGION '設定DB連接'}
  try
    con1.Connected := false;
    con1.ConnectionString := 'FILE NAME='+folder+'\'+ 'db.udl';
    con1.Provider := folder+'\'+'db.udl';
    con1.Connected := true;
  except
    showmessage('連線錯誤,請檢查.udl設定是否正確');
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
  {$ENDREGION}

  qry1.Close;
  qry1.SQL.Text := SqlEP;
  qry1.Open;
  qry1.First;
  with qry1 do
  for I := 0 to  qry1.FieldCount -1 do
    begin
      TimeTitle.Items.Add(dbgrd1.Columns[i].FieldName);
      dbgrd1.Columns[i].Width := 85;
    end;
  lblAll.Caption := '';
  lblAll.Caption := '全部筆數：' + IntToStr( dbgrd1.DataSource.DataSet.RecordCount );
end;

procedure TFormTable.ChoiceTest(Sender: TObject);
var
  Filename,folder :string;
  Myinifile:Tinifile;
  i,temp1 : Integer;
  strtemp,temp: string ;
begin
  SqlEP := '';
  if SelectDirectory('請選擇設定檔目錄', '', folder) then
  mmoTemp.Lines.LoadFromFile(folder+'\SqlSetting.ini');

  {$REGION '設定DB連接'}
  try
    con1.Connected := false;
    con1.ConnectionString := 'FILE NAME='+folder+'\'+ 'db.udl';
    con1.Provider := folder+'\'+'db.udl';
    con1.Connected := true;
  except
    showmessage('連線錯誤,請檢查.udl設定是否正確');
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
  {$ENDREGION}

  qry1.Close;
  qry1.SQL.Text := SqlEP;
  qry1.Open;
  qry1.First;
  with qry1 do
  for I := 0 to  qry1.FieldCount -1 do
    begin
      TimeTitle.Items.Add(dbgrd1.Columns[i].FieldName);
      dbgrd1.Columns[i].Width := 80;
    end;
end;

procedure TFormTable.SaveToExcel(Sender: TObject);
var
Excel: Variant;
Sheet,xlQuery: Variant;
snocount,i : Integer;
xlsFileName,rowcell : string;
begin
  {$REGION '進度條'}
  Gauge2.MinValue:= 0;
  Gauge2.MaxValue:= 5;
  Gauge2.Progress:= 0;
  {$ENDREGION}


  dlgSave1.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave1.Filter := 'Excel files(*.xlsx)|*.xlsx|Excel files(*.xls)|*.xls|Excel files(*.csv)|*.csv|Excel files(*.txt)|*.txt';
  dlgSave1.Title := '選擇儲存位置';
  dlgSave1.FileName := '資料庫資料核對';
if dlgSave1.Execute then
  begin
   try
     xlsFileName := dlgSave1.FileName;
     Excel := CreateOLEObject('Excel.Application');
     Excel.Visible := false;
     Excel.WorkBooks.Add;
     Excel.WorkSheets[1].Activate;
     Excel.Workbooks[1].Worksheets[1].Name := '大表';
     Sheet := Excel.Workbooks[1].Worksheets['大表'];
   {$REGION '內容'}
    xlQuery := Excel.WorkSheets[1].QueryTables.Add(qry1.Recordset,Excel.Range['A1']);
    xlQuery.FieldNames := True;
    xlQuery.RowNumbers := False;
    xlQuery.FillAdjacentFormulas := False;
    xlQuery.PreserveFormatting := True;
    xlQuery.RefreshOnFileOpen := False;
    xlQuery.BackgroundQuery := True;
    xlQuery.SavePassword := True;
    xlQuery.SaveData := True;
    xlQuery.AdjustColumnWidth := True;
    xlQuery.RefreshPeriod := 0;
    xlQuery.PreserveColumnInfo := True;
    xlQuery.FieldNames := True;
    xlQuery.Refresh;
   {$ENDREGION}
   Gauge2.Progress:= 1;
   {$REGION '內容樣式設定'}
    //抓取欄位值
   snocount:=(Excel.ActiveSheet.UsedRange.Columns.Count);
     if (snocount) > 26 then
       if (snocount mod 26 ) = 0 then
          rowcell := char((snocount div 26)+64)+'Z'
       else
         begin
            rowcell := char((snocount div 26)+64)+ char((snocount mod 26)+64);
         end
     else
      begin
        rowcell := char(snocount+65);
      end;
   Gauge2.Progress:=2;
     //邊線
     Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Borders[1].Weight := xlThin   ;
     Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Borders[2].Weight := xlThin   ;
     Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Borders[3].Weight := xlThin   ;
     Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Borders[4].Weight := xlThin   ;
   Gauge2.Progress:=3;
    //標題
    Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Font.Name:='微軟正黑體';
    Excel.Range['A1',rowcell+IntToStr(Excel.ActiveSheet.UsedRange.Rows.Count)].Font.Color:=RGB(31,31,31);
    Excel.ActiveSheet.Rows[1].Font.Color := RGB(255,255,255);
    Excel.ActiveSheet.Rows[1].Font.Bold := True;
    Sheet.Rows[1].Interior.Color := RGB(150,150,150);
    Excel.ActiveWindow.SplitColumn := 0;
    Excel.ActiveWindow.SplitRow := 1;
    Excel.ActiveWindow.FreezePanes := True;
    //內容
    for i := 2 to Excel.ActiveSheet.UsedRange.Rows.Count do
      if i mod 2 = 0 then Sheet.Rows[i].Interior.Color := RGB(193,226,247)
      else Sheet.Rows[i].Interior.Color := RGB(248,222,200);
   Gauge2.Progress:=4;
   {$ENDREGION}

   finally
    Excel.ActiveWorkBook.Saved := False; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
    Excel.WorkBooks[1].SaveAs(xlsFileName);
    Excel.WorkBooks.close;  //關閉Excel
    Excel.Quit;             //離開Excel
    Excel:=Unassigned;      //釋放ExcelApp;
   end;
   Gauge2.Progress:= Gauge2.MaxValue;
   ShowMessage('匯出完成');
  end;
end;

end.
