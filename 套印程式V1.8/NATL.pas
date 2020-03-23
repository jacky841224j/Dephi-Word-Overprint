unit NATL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ComObj, DB, ADODB,StrUtils, ExtCtrls, Grids,
  DBGrids, ComCtrls, Gauges,TlHelp32,Excel2000, Menus, pngimage;

type
  TForm1 = class(TForm)
    pgc1: TPageControl;
    chkPw: TCheckBox;
    mm1: TMainMenu;
    mniN1: TMenuItem;
    mniCheck: TMenuItem;
    mniExport: TMenuItem;
    img1: TImage;
    function Qstr(str:String):String;
    function DBTableExists(aTableName: string;aADOConn:TADOConnection): Boolean;
//    function  WordReplace(str:String ; OldText:String ; NewText:String):String ;
    function KillExcelTask : integer;
    procedure FormShow(Sender: TObject);
    procedure mniCheckClick(Sender: TObject);
    procedure mniExportClick(Sender: TObject);
    procedure pgc1Change(Sender: TObject);
    function KillWordTask : integer;
  type
//  ThreadForm = class(TThread)
//    private
//      str :string;
//      pg1 :TGauge;
//    public
//      constructor Create(str :String);
//      destructor Destroy;override;
//    protected
//      procedure Execute; override;
//      procedure sendmes;
//      procedure progress;
//  end;

  private
    { Private declarations }
  public
    t,m:integer;
    Tablist:Tstringlist;
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

uses GRADE,login,UnitMyThread,Table;

{$R *.dfm}
var
  DirPath: string;
//function  TForm1.WordReplace(str:String ; OldText:String ; NewText:String):String ;
//begin
//str:= myRange.Find.Execute(FindText:=QuotedStr(OldText), ReplaceWith:=FieldByName(QuotedStr(OldText)).AsString, Replace:=2);
//end;

function TForm1.Qstr(str: string):string;
begin
Result:=QuotedStr(Str);
end;

function TForm1.DBTableExists(aTableName: string;aADOConn:TADOConnection): Boolean;
var
vTableNames : TStringList;
begin
Result:=False;
vTableNames := TStringList.Create;
  try
    aADOConn.GetTableNames(vTableNames);//取得所有表名
    if vTableNames.IndexOf(aTableName)>= 0 then //判断是否存在
    Result:=True;
  finally
    vTableNames.Free;
  end;
end;

procedure TForm1.FormShow(Sender: TObject);

begin
Self.Left := (Screen.Width - Self.Width) div 2 ;
Self.Top := (Screen.Height - Self.Height) div 2 ;

Application.CreateForm(TLoginCheck, LoginCheck);
LoginCheck.ShowModal;
if chkPw.Checked = False  then Close;
LoginCheck.Close;
Tablist := Tstringlist.Create;
end;


function TForm1.KillWordTask : integer;
const
  PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  result := 0;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = 'WINWORD.EXE') or
       (UpperCase(FProcessEntry32.szExeFile) = 'WINWORD.EXE')) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID), 0));
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;

function TForm1.KillExcelTask : integer;
const
  PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  result := 0;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = 'EXCEL.EXE') or
       (UpperCase(FProcessEntry32.szExeFile) = 'EXCEL.EXE')) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID), 0));
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;


procedure TForm1.mniExportClick(Sender: TObject);
var
Form : TForm4;
TabSheet : TTabSheet;
I: Integer;
b :boolean;
begin
  pgc1.Visible := True;
  b := false;
  try
  for I := 0 to pgc1.PageCount - 1 do
    if  pgc1.Pages[i].Caption  = '資料套印' then
      begin
          pgc1.ActivePageIndex := i;
          b:=true;
      end;
  if not b then
    begin
      with pgc1 do
      TabSheet := TTabSheet.Create(self);
      TabSheet.Name := 'tsExport';
      TabSheet.Caption := '資料套印';
      TabSheet.PageControl := pgc1;
      TabSheet.Align := alClient;
      Tablist.AddObject('資料套印',TObject(TabSheet));
      pgc1.ActivePageIndex := pgc1.PageCount-1;
    end
  except
    FreeAndNil(TabSheet);
    Exit;
  end;

  try
  //创建窗口
    Form := TForm4.Create(self);
    Form.Parent := TabSheet;
    Form.BorderStyle := bsNone;
    Form.Top := 0;
    Form.Left := 0;
    Form.Width := TabSheet.Width;
    Form.Height := TabSheet.Height;
    Form.Align := alClient;
    Form.Show;
  except
    FreeAndNil(Form);
    Abort;
  end;
end;

procedure TForm1.mniCheckClick(Sender: TObject);
var
Form : TFormTable;
TabSheet : TTabSheet;
I: Integer;
b :boolean;
begin
  pgc1.Visible := True;
  b := false;
  try
    for I := 0 to pgc1.PageCount - 1 do
      if  pgc1.Pages[i].Caption  = '資料核對' then
      begin
          pgc1.ActivePageIndex := i;
          b:=true;
      end;
    if not b then
    begin
    with pgc1 do
    TabSheet := TTabSheet.Create(self);
    TabSheet.Name := 'tsTable';
    TabSheet.Caption := '資料核對';
    TabSheet.PageControl := pgc1;
    TabSheet.Align := alClient;
    Tablist.AddObject('資料核對',TObject(TabSheet));
    pgc1.ActivePageIndex := pgc1.PageCount-1;
    end
  except
    FreeAndNil(TabSheet);
    Exit;
  end;

  try
  //创建窗口
    Form := TFormTable.Create(self);
    Form.Parent := TabSheet;
    Form.BorderStyle := bsNone;
    Form.Top := 0;
    Form.Left := 0;
    Form.Width := TabSheet.Width;
    Form.Height := TabSheet.Height;
    Form.Align := alClient;
    Form.Show;
  except
    FreeAndNil(Form);
    Abort;
  end;
end;

procedure TForm1.pgc1Change(Sender: TObject);
begin
//    ShowMessage(TPageControl(Sender).Pages[TPageControl(Sender).ActivePageIndex].Caption);
end;

end.
