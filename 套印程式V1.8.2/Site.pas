unit Site;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, CheckLst, Gauges, Mask, DBCtrls,IniFiles ,
  DBClient, ExtCtrls,TlHelp32,Excel2000,ComObj,ADODB,FileCtrl,StrUtils,ComCtrls  ;

type
  TRoster = class(TForm)
    dbgrd1: TDBGrid;
    ds1: TDataSource;
    strngrdCheckList: TStringGrid;
    Gauge2: TGauge;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Roster: TRoster;
  Filename,dirpath :string;
  Myinifile:Tinifile;
implementation
uses natl;
{$R *.dfm}





procedure TRoster.FormShow(Sender: TObject);
var
i : Integer;
begin
{$REGION '準備ini檔案資料'}
  Filename:=ExtractFilePath(Paramstr(0))+'Setting.ini';
  myinifile:=Tinifile.Create(filename);
{$ENDREGION}

{$REGION '設定DBGRID'}
  Form1.qry1.Close;
  form1.qry1.SQL.Text := myinifile.readstring('SiteSql','SQLFormShow ','');
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
//  strngrdCheckList.Cells[2,0] := #32#32#32#32'---【取代文字】---' ;
//  strngrdCheckList.Cells[3,0] := #32#32#32#32'=' ;
//  strngrdCheckList.Cells[4,0] := #32#32#32#32'---【搜尋條件】---' ;
  strngrdCheckList.ColWidths[0] := 200;
  strngrdCheckList.ColWidths[1] := 180;
//  strngrdCheckList.ColWidths[2] := 180;
//  strngrdCheckList.ColWidths[3] := 50;
//  strngrdCheckList.ColWidths[4] := 180;
  strngrdCheckList.RowCount:=form1.qry1.FieldCount+1 ;
  {$ENDREGION}
end;

end.
