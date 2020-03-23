unit Rep;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,FileCtrl, StdCtrls, DB, ADODB, Gauges, Grids, DBGrids, Buttons,IniFiles,ComCtrls,
  CheckLst, ExtCtrls,StrUtils, jpeg;

type
  TRepExport = class(TForm)
    dbgrd1: TDBGrid;
    con1: TADOConnection;
    Gauge2: TGauge;
    qry1: TADOQuery;
    ds1: TDataSource;
    mmoTemp: TMemo;
    btnChgSet: TBitBtn;
    btnClose: TBitBtn;
    chklstTitle: TCheckListBox;
    pnl1: TPanel;
    qry2: TADOQuery;
    lblSave: TLabel;
    edtPath: TEdit;
    btnDirpath2: TButton;
    lblFile: TLabel;
    edtFile: TEdit;
    btnFile: TButton;
    dlgOpen1: TOpenDialog;
    btnPreview: TBitBtn;
    btnExport: TBitBtn;
    btnKillTask: TSpeedButton;
    mmoSql: TMemo;
    procedure FormShow(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure chklstTitleClickCheck(Sender: TObject);
    procedure btnDirpath2Click(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure CLBOnClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RepExport: TRepExport;
  folder : string ;
  Filename,dirpath :string;
  Myinifile:Tinifile;
  MemoList :Tstringlist;
  SqlFS,SqlEP,SqlTitle : string ;
implementation
uses NATL,Login,UnitMyThread;
{$R *.dfm}

procedure TRepExport.btnCloseClick(Sender: TObject);
var i :integer;
begin
  close;
  i:=Form1.Tablist.IndexOf('報表套印');
  TTabSheet(Form1.Tablist.Objects[i]).Destroy;
  Form1.Tablist.Delete(i);
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

procedure TRepExport.btnExportClick(Sender: TObject);
var i,p:integer;
    Memo :Tmemo;
    ColName,Con :String;
begin
  for I := 0 to MemoList.Count - 1 do
    begin
      Memo := TMemo(Memolist.Objects[i]);
      ColName :=  StringReplace( Memolist.Strings[i],'chk_','',[rfReplaceAll]);  //MEMO 名稱
      for p:= 0 to  Memo.Lines.Count -1 do
        begin
          Con :=  Con+Memo.Lines[p]+',';
        end;
    end;
    Con := ColName +' in '+'('+ QuotedStr(Copy(Con,1,Length(Con)-1) )  + ')';
    ShowMessage(con);
end;

procedure TRepExport.btnFileClick(Sender: TObject);
begin
  if dlgOpen1.Execute then   edtFile.Text := dlgOpen1.FileName ;
end;

procedure TRepExport.chklstTitleClickCheck(Sender: TObject);
var
i,J,list,q,q1,chkCount : Integer ;
CLB : TCheckListBox ;
CLBtext : TLabel;
CLBMemo : TMemo;
same :boolean;
begin
  chkCount := 1;
  for J :=ComponentCount-1 downto 0 do
    if (Components[J] is TCheckListBox) and
      (TCheckListBox(Components[J]).Parent = Pnl1) and
      (TCheckListBox(Components[J]).Name <> 'chklstTitle' )  or
      ( (Components[J] is TLabel) and (TLabel(Components[J]).Parent = Pnl1)  ) or
      ( (Components[J] is TMemo) and (TMemo(Components[J]).Parent = Pnl1)  )

    then Components[J].Free;

  list := 0 ;
  for I := 0 to chklstTitle.Count - 1 do
  if chklstTitle.Checked[i] and (chkCount <= 4) then
    begin
      Inc(list);
      inc(chkCount);
      //動態產生CheckListBox
      CLB:=TCheckListBox.Create(Self);
      CLB.Left:= list*194 + chklstTitle.Left ;
      CLB.Top:= chklstTitle.Top+33;
      CLB.Width:=chklstTitle.Width;
      CLB.Height:=chklstTitle.Height-33;
      CLB.Name:='chk_'+chklstTitle.Items[i];
      CLB.Font.Name:='Times New Roman';
      CLB.Font.Size:=14;
      CLB.Parent:=Pnl1;
      CLB.Enabled:=True;
      CLB.Visible:=True;
      CLB.OnClickCheck := CLBOnClick ;

      //動態產生Label
      CLBtext:=TLabel.Create(Self);
      CLBtext.Left:= list*194 + chklstTitle.Left ;
      CLBtext.Top:= 8;
      CLBtext.Name:= 'lbl_'+chklstTitle.Items[i] ;
      CLBtext.Caption := chklstTitle.Items[i] ;
      CLBtext.Font.Name:='新細明體';
      CLBtext.Font.Size:=14;
      CLBtext.Font.Style := CLBtext.Font.Style + [fsBold] ;
      CLBtext.Parent:=Pnl1;
      CLBtext.Enabled:=True;
      CLBtext.Visible:=True;

      //動態產生Memo
      CLBMemo:=TMemo.Create(Self);
      CLBtext.Name:= 'memo_'+chklstTitle.Items[i] ;
      CLBMemo.Parent:=Pnl1;
      CLBMemo.WordWrap := false ;
      CLBMemo.Enabled:=True;
      CLBMemo.Visible:=false;
      MemoList.AddObject('chk_'+chklstTitle.Items[i],CLBMemo);

      qry2.Close;
      qry2.SQL.Text :=  StringReplace(SqlTitle,#$D#$A,' ',[rfReplaceAll]);
      qry2.Open;
      qry2.First;
      with qry1 do
      for q := 0 to  qry2.RecordCount  -1 do
      begin
        same := false;
        if  CLB.Items.Count = 0 then CLB.Items.Add(qry2.FieldByName(CLBtext.caption).AsString) ;
        for q1 := 0 to  CLB.Items.Count -1 do
          if   CLB.Items.Strings[q1] = qry2.FieldByName(CLBtext.caption).AsString   then
            begin
              same := true;
              Break;
            end
          else continue;
          if not same then CLB.Items.Add(qry2.FieldByName(CLBtext.caption).AsString);
        qry2.next;
      end;

    end

  else if chklstTitle.Checked[i] and (chkCount > 4) then
    begin
      ShowMessage('最多只能選取四個，請先取消其他勾選');
      chklstTitle.Checked[i] := false;
      break;
    end;
end;


procedure TRepExport.FormShow(Sender: TObject);
var
i,temp1,temp2,temp3 : Integer;
strtemp,temp : string ;
begin

  SqlEP := '';
  SqlFS := '';
  SqlTitle := '';
  if SelectDirectory('請選擇設定檔目錄', '', folder) then
  mmoTemp.Lines.LoadFromFile(folder+'\SqlSetting.ini');
  MemoList := TStringList.Create;
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

  {$REGION '設定預設路徑'}
  if not directoryExists(folder+'\'+'套印資料夾') then  //判斷此資料夾是否存在
  CreateDir(folder+'\'+'套印資料夾');                 //建立資料夾
  edtPath.Text:= folder+'\'+'套印資料夾'+'\';
  dirpath:= edtPath.Text   ;
  edtFile.Text :=folder+'\'+'套印範本.docx' ;
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

  //欄位不重複值
  temp2:= POS('#', temp);
  temp3:= POS('$', temp);
  strtemp:= copy(temp, temp2+1,temp3-1);       //擷取字串
  for i:=1 to Length(strtemp) do            //去除多餘字串
    if not (strtemp[i] = '$') then SqlTitle := SqlTitle + strtemp[i]
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
      chklstTitle.Items.Add(dbgrd1.Columns[i].FieldName);
    end;
  {$ENDREGION}

end;

procedure TRepExport.CLBOnClick (Sender: TObject);
var
CheckClick: TCheckListBox;
Memo :TMemo;
begin
  Memo := TMemo(Memolist.Objects[MemoList.IndexOf(CheckClick.Name)]);
  CheckClick := TCheckListBox(Sender);
  if Memo.Lines.IndexOf(CheckClick.items[CheckClick.ItemIndex]) >=0 then
      Memo.Lines.Delete(Memo.Lines.IndexOf(CheckClick.items[CheckClick.ItemIndex]))
  else
    Memo.Lines.Add(CheckClick.items[CheckClick.ItemIndex]);
end;
end.
