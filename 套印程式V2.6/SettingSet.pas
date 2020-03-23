unit SettingSet;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Grids,
  Data.DB, Data.Win.ADODB,FileCtrl,IniFiles, Vcl.DBGrids,ComCtrls,
  Vcl.Imaging.jpeg, Vcl.ExtCtrls,ShellAPI;

type
  TSSet = class(TForm)
    con1: TADOConnection;
    ds1: TDataSource;
    qry1: TADOQuery;
    strngrdCheckList: TStringGrid;
    btnChgSet: TBitBtn;
    btnClose: TBitBtn;
    mmoTemp: TMemo;
    dbgrd1: TDBGrid;
    btnExport: TBitBtn;
    cbbSaveField: TComboBox;
    cbbPhoto: TComboBox;
    imgBackGroup: TImage;
    procedure btnChgSetClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure dbgrd1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SSet: TSSet;
  folder : string;

implementation
uses NATL;

{$R *.dfm}

procedure TSSet.btnChgSetClick(Sender: TObject);
var
i,temp1 : Integer;
strtemp,temp,SqlEP : string ;
Myinifile:Tinifile;
Filename : string;
begin
  if SelectDirectory('請選擇設定檔目錄', '', folder) then

  {$REGION '設定DB連接'}
  try
    con1.Connected := false;
    con1.ConnectionString := 'FILE NAME='+folder+'\'+ 'db.udl';
    con1.Provider := folder+'\'+'db.udl';
    con1.Connected := true;
  except
    showmessage('路徑選擇錯誤，請重新選擇(無法連線至SQL)');
    EXIT;
  end;
  {$ENDREGION}

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'SqlSetting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION '擷取SQL 字串'}
  //匯出SQL
  SqlEP := myinifile.readstring('SQL','Save','');
  {$ENDREGION}

  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  {$REGION '設定欄位'}
  qry1.Close;
  qry1.SQL.Text := SqlEP;
  qry1.Open;
  qry1.First;
  cbbPhoto.Items.Add('');
  with qry1 do
  for I := 0 to  qry1.FieldCount -1 do
    begin
      dbgrd1.Columns[i].Width := 70;
      strngrdCheckList.Cells[0,i+1]:=dbgrd1.Columns[i].FieldName;
      cbbSaveField.Items.Add(dbgrd1.Columns[i].FieldName);
      cbbPhoto.Items.Add(dbgrd1.Columns[i].FieldName);
    end;
  {$ENDREGION}

  {$REGION 'StringGrid'}
  strngrdCheckList.Cells[0,0] := #32#32#32#32'---【資料欄位】---' ;
  strngrdCheckList.Cells[1,0] := '---【欄位說明】---' ;
  strngrdCheckList.ColWidths[0] := 200;
  strngrdCheckList.ColWidths[1] := 140;
  strngrdCheckList.RowCount:=qry1.FieldCount+1 ;
  {$ENDREGION}

  {$REGION '判斷是否有設定檔'}
  if  FileExists(folder+'\'+'Setting.ini') then
    begin
//      if MessageDlg('是否讀取上次設定？',mtInformation,[mbYes,mbNo],0)=mrYes then
         mmoTemp.Lines.LoadFromFile(folder+'\'+'Setting.ini');
      for i := 0 to qry1.FieldCount -1  do
        strngrdCheckList.Cells[1,i+1]:= myinifile.Readstring('Help',dbgrd1.Columns[i].FieldName,'');
    end;
  {$ENDREGION}
  end;

procedure TSSet.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('Setting.ini_設定');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
//  form1.pgc1.Visible := false ;
end;

procedure TSSet.btnExportClick(Sender: TObject);
var
Myinifile:Tinifile;
Filename : string;
i : Integer;
begin
  {$REGION '準備ini檔案資料'}
  Filename:=folder+'\'+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
  {$ENDREGION}

  try
    //寫入檔名欄位
    myinifile.writestring('SQL','SaveField',cbbSaveField.Items[cbbSaveField.ItemIndex]);
    //寫入Log欄位
    myinifile.writestring('SQL','SQLLOG','[Log_CheckIn]');
    //寫入相片欄位
    myinifile.writestring('Photo','pic',cbbPhoto.Items[cbbPhoto.ItemIndex]);
    //寫入說明欄位
    for i := 0 to qry1.FieldCount -1 do
      myinifile.writestring('Help',dbgrd1.Columns[i].FieldName,strngrdCheckList.Cells[1,i+1]);
  finally
    myinifile.Free;
    if MessageBox(0,'Setting.ini 建立完成,是否開啟資料夾?','OPEN',
                    MB_OKCANCEL + MB_ICONASTERISK + MB_DEFBUTTON2 ) = 1 then
      ShellExecute(Handle, 'open',PWideChar(ExtractFileDir(fileName)), nil, nil, SW_SHOW);
  end;
end;
procedure TSSet.dbgrd1DrawColumnCell(Sender: TObject; const Rect: TRect;
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

end.
