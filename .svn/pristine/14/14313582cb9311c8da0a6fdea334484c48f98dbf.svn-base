unit NATL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ComObj, DB, ADODB,StrUtils, ExtCtrls, Grids,
  DBGrids, ComCtrls, Gauges,TlHelp32,Excel2000;

type
  TForm1 = class(TForm)
    dlgOpen1: TOpenDialog;
    dlgSave1: TSaveDialog;
    con1: TADOConnection;
    qry1: TADOQuery;
    pgc1: TPageControl;
    tsExport: TTabSheet;
    dlgOpendlgudl: TOpenDialog;
    chkPw: TCheckBox;

    function Qstr(str:String):String;
    function DBTableExists(aTableName: string;aADOConn:TADOConnection): Boolean;
//    function  WordReplace(str:String ; OldText:String ; NewText:String):String ;
    function KillExcelTask : integer;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsExportExit(Sender: TObject);
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
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

uses GRADE,login;

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

procedure TForm1.tsExportExit(Sender: TObject);
var
  Form : TForm4;
  TabSheet : TTabSheet;
begin
   try
    TabSheet := tsExport;
    TabSheet.PageControl := pgc1;
    TabSheet.Tag := 0;
    TabSheet.Align := alClient;
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

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LoginCheck.close;
end;

procedure TForm1.FormShow(Sender: TObject);
var
Form : TForm4;
TabSheet : TTabSheet;
begin
Self.Left := (Screen.Width - Self.Width) div 2 ;
Self.Top := (Screen.Height - Self.Height) div 2 ;
Application.CreateForm(TLoginCheck, LoginCheck);
LoginCheck.ShowModal;

  if chkPw.Checked = true then
    begin
     try
        TabSheet := tsExport;
        TabSheet.PageControl := pgc1;
        TabSheet.Tag := 0;
        TabSheet.Align := alClient;
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
    end
  else Close;

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
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = 'WINWORD.EXE') or
       (UpperCase(FProcessEntry32.szExeFile) = 'WINWORD.EXE')) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID), 0));
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;
end.
