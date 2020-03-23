unit FirstExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DB, ADODB;
type
  TstringlistEx = class(Tstringlist)
    public
      function IndexOf(const S: string): Integer; override;
  end;
type
  TForm1 = class(TForm)
    btn1: TButton;
    rg1: TRadioGroup;
    btn2: TButton;
    btn3: TButton;
    con1: TADOConnection;
    qry1: TADOQuery;
    dlgSave1: TSaveDialog;
    qry2: TADOQuery;
    procedure SQLOpen(SQL :string);
    procedure SQLExec(SQL :string);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure GetRepExFile(const ExName, SavePath: String);
    procedure btn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure StdData;
  end;


var
  Form1: TForm1;
  StdDataList :TstringlistEx;
implementation
 USES ExcelUnit;
{$R *.dfm}
procedure TForm1.StdData;
var SQL,Data,Rk :String;
    RMK :Tstringlist;
    i,Cnt:integer;
begin
    SQLOpen('select ''1''+Ex_Year from Exam');
    SQL := 'select '''+qry1.Fields[0].AsString+'''+S.Sch_Code+S.Class_No+right(Seat_No,2) 學生辨識碼,'''+qry1.Fields[0].AsString+''' 年度,''ID'' 名冊序號,Right(RTRIM(Remark),1) 群組代碼,' + #13#10 +
          'Zip,Tel,Address,S.sch_Code,Sch_name,RTrim(Scl.Class_name),S.Class_no,Right(Seat_no,2),student_name,CASE sex WHEN 1 then ''男'' ELSE ''女'' END 性別,Remark1' + #13#10 +
          'from Student S Left JOIN School Sch on S.Sch_code = Sch.Sch_code LEFT JOIN Sch_Class Scl on S.Sch_code = Scl.Sch_code and S.Grade = Scl.Grade and S.Class_no = Scl.Class_no and Remark1 <>''一般生''';
    SQLOpen(SQL);
    RMK:= TStringList.Create;
    while not qry1.Eof do
    begin
        Cnt := 0;
        Rk := qry1.FieldByName('Remark1').AsString;
        for I := 0 to Length(Rk) - 1 do
        begin
            if Rk[i] = '-' then
               inc(Cnt);
        end;
        if Cnt <6 then
           for I := Cnt to 6 do
              Rk := Rk+'-';
        RMK.Clear;
        RMK.Delimiter := '-';
        RMK.DelimitedText := Rk;
        Data := '';
        for I := 0 to qry1.FieldCount - 2 do
            Data := Data + qry1.fields[i].AsString + #9;
        Data := Data +RMK[1]+#9+RMK[2]+#9+RMK[3]+#9+#9+RMK[0]+#9+#9;
        StdDataList.Add(Data);
        qry1.Next;
    end;
end;
procedure TForm1.btn1Click(Sender: TObject);
var SQL,Data,Path,io :String;
    Excel:ExcelLab;
    Cnt,i :integer;
    DataList:tstringlist;
begin
    dlgSave1.FileName := 'Defalt';
    dlgSave1.Filter := 'Excel 2007|*.xlsx';
    DataList := TStringList.Create;
    if not dlgSave1.Execute then
       exit;
    Path := ExtractFilePath(dlgSave1.FileName);
    GetRepExFile('TPTEST',Path+'TPTEST.xlsx');
    if StdDataList.Count <=0 then
      StdData;
    SQL :='select ''108''+S.Sch_Code+S.Class_No+right(Seat_No,2),ISNULL([01],'','') [01],ISNULL([03],'','') [03],ISNULL([02],'','') [02] from ' + #13#10 +
          '(select Student_No,Sub_no,施測情形+'',''+備註 施測情形 FROM' + #13#10 +
          '(select Student_No,Right(BarCode,2) Sub_no,' + #13#10 +
          'Case WHEN A1 = 1 Then ''A 缺席（請假）''' + #13#10 +
          ' WHEN A2 = 1 Then ''B 缺席（轉出）''' + #13#10 +
          ' WHEN A3 = 1 Then ''C 更正座號''' + #13#10 +
          ' WHEN A4 = 1 Then ''D 轉入''' + #13#10 +
          ' WHEN A5 = 1 Then ''E 其他'' ' + #13#10 +
          'END 施測情形,' + #13#10 +
          'Case WHEN RTrim(A3_Rem) <> '''' Then Replace(RTrim(REPLACE(A3_Rem, '','', ''，'')),CHAR(13)+CHAR(10),'''')' + #13#10 +
          '	    WHEN RTrim(A5_Rem) <> '''' Then Replace(REPLACE(A5_Rem, '','', ''，''),REPLACE(A5_Rem, '','', ''，''),''"''+REPLACE(A5_Rem, '','', ''，'')+''"'') ' + #13#10 +
          '     WHEN RTrim(A5_Rem) <> '''' and RTrim(A3_Rem) <> '''' Then Replace(RTrim(REPLACE(A3_Rem,'','', ''，'')),CHAR(13)+CHAR(10),'''')+''-''+ Replace(REPLACE(A5_Rem, '','', ''，''),CHAR(13)+CHAR(10),'''')' + #13#10 +
          'END 備註' + #13#10 +
          ' from ReadCardRecTab where Flag = 1 and student_no <> ''''  ) a ) b' + #13#10 +
          'PIVOT' + #13#10 +
          '(' + #13#10 +
          ' MAX(施測情形)' + #13#10 +
          ' FOR Sub_no IN ([01],[02],[03])' + #13#10 +
          ') AS PivotTable LEFT JOIN Student S on S.Student_no = PivotTable.Student_no' + #13#10 +
          'ORDER BY S.student_no';
    SQLOpen(SQL);

    while not qry1.Eof do
    begin
        Cnt := 0;
        Data := '';

        for I := 1 to qry1.FieldCount - 1 do
            Data := Data + qry1.Fields[i].AsString +',';
        Data := StringReplace(Data,',',#9,[rfReplaceAll, rfIgnoreCase]);
        if Trim(qry1.Fields[0].AsString) = '1080070228' then
        begin
           io := '123';
        end;
        Cnt := StdDataList.IndexOf(Trim(qry1.Fields[0].AsString));
        DataList.Add(StdDataList[Cnt]+Data);
        qry1.Next;
    end;
    Excel := ExcelLab.create(Path+'TPTEST.xlsx');
    Excel.Paste('A9',DataList.Text);
    Excel.AutoReplace := true;
    Excel.Save;
    Excel.destory;
end;

procedure TForm1.btn2Click(Sender: TObject);
var Path,SQL,Xls,Data:string;
    Excel:ExcelLab;
    SubList,Datalist :Tstringlist;
    i,cnt:integer;
begin
    if rg1.ItemIndex = -1 then
    begin
        ShowMessage('請選擇科目');
        exit;
    end;
    dlgSave1.FileName := 'Defalt';
    dlgSave1.Filter := 'Excel 2007|*.xlsx';
    if not dlgSave1.Execute then
       exit;
    SubList := TStringList.Create;
    Datalist := TStringList.Create;
    Path := ExtractFilePath(dlgSave1.FileName);
    case rg1.ItemIndex of
        0:Xls:= '01';
        1:Xls:= '02';
        2:Xls:= '03';
        3:Xls:= '01-02-03';
    end;
    SubList.Delimiter := '-';
    SubList.DelimitedText := Xls;
    for I := 0 to SubList.Count - 1 do
    begin
        GetRepExFile('TPCARD'+SubList[i],Path+'TPCARD'+SubList[i]+'.xlsx');
        if SubList[i] = '01' then
        begin
            SQL :='select BarCode+'','',isnull( dbo.AnsD(''01'',left(Ans,dbo.AnsPos(28,Ans))),'',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'') Ans,Barcode+'','',isnull(substring(Ans,dbo.AnsPos(28,Ans)+1,len(Ans)),'',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'') Ans2' + #13#10 +
                  'from Sub_Score SS INNER JOIN student S on SS.Student_No = S.Student_No ' + #13#10 +
                  'INNER JOIN ReadCard1 RC1 on S.Student_no = LEFT(RC1.X_No,9) and Sub_no = Right(RC1.X_No,2) where Sub_No = ''01'' ORDER BY BarCode';
            SQLOpen(SQL);
            Excel := ExcelLab.create(Path+'TPCARD'+SubList[i]+'.xlsx');

        end
        else
        begin
            SQl :='select BarCode+'','',Ans' + #13#10 +
                  'from Sub_Score SS INNER JOIN student S on SS.Student_No = S.Student_No ' + #13#10 +
                  'INNER JOIN ReadCard1 RC1 on S.Student_no = LEFT(RC1.X_No,9) and Sub_no = Right(RC1.X_No,2) where Sub_No = '''+SubList[i]+''' ORDER BY BarCode';
            SQLOpen(SQL);
            Excel := ExcelLab.create(Path+'TPCARD'+SubList[i]+'.xlsx');
        end;
        while not qry1.Eof do
        begin
            Data := '';
            for cnt := 0 to qry1.FieldCount - 1 do
                Data := Data + qry1.Fields[cnt].AsString;
            Data := StringReplace(Data,',',#9,[rfReplaceAll, rfIgnoreCase]);
            Datalist.Add(Data);
            qry1.Next;
        end;
        Excel.Paste('A8',Datalist.Text);
        Excel.AutoReplace := true;
        Excel.Save;
        Excel.destory;
    end;
end;
procedure TForm1.btn3Click(Sender: TObject);
var SQL,Data,Path,Xls,io,SubCnt :String;
    Excel:ExcelLab;
    Cnt,i,idx :integer;
    DataList,SubList:tstringlist;
begin
    if rg1.ItemIndex = -1 then
    begin
        ShowMessage('請選擇科目');
        exit;
    end;
    dlgSave1.FileName := 'Defalt';
    dlgSave1.Filter := 'Excel 2007|*.xlsx';
    if not dlgSave1.Execute then
       exit;
    SubList := TStringList.Create;
    Datalist := TStringList.Create;
    Path := ExtractFilePath(dlgSave1.FileName);
    case rg1.ItemIndex of
        0:Xls:= '01';
        1:Xls:= '02';
        2:Xls:= '03';
        3:Xls:= '01-02-03';
    end;
    SubList.Delimiter := '-';
    SubList.DelimitedText := Xls;
    if StdDataList.Count <=0 then
      StdData;
    for I := 0 to SubList.Count - 1 do
    begin
        SQL := 'SELECT count(*) c FROM Sub_Ans WHERE Q_Ans <> ''@'' AND Sub_No = '''+SubList[i]+'''';
        SQLOpen(SQL);
        SubCnt := qry1.Fields[0].AsString;
        GetRepExFile('TP'+SubList[i],Path+'TP'+SubList[i]+'.xlsx');
        if SubList[i] = '01' then
        begin
            io := '';
            SQL :='select ''108''+S.Sch_Code+S.Class_No+right(Seat_No,2) Std,RC1.BarCode+'','',' + #13#10 +
                  'isnull( dbo.AnsD(''01'',left(Ans,dbo.AnsPos(28,Ans))),'',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'') Ans,' + #13#10 +
                  'Case WHEN A1 = 1 Then ''A 缺席（請假）,'' WHEN A2 = 1 Then ''B 缺席（轉出）,'' WHEN A3 = 1 Then ''C 更正座號,'' WHEN A4 = 1 Then ''D 轉入,'' WHEN A5 = 1 Then ''E 其他,''  Else '','' END 施測情形,' + #13#10 +
                  'Case WHEN RTrim(A3_Rem) <> '''' Then RTrim(REPLACE(A3_Rem, '','', ''，''))+'','' ' + #13#10 +
                  '	   WHEN RTrim(A5_Rem) <> '''' Then Replace(REPLACE(A5_Rem, '','', ''，''),REPLACE(A5_Rem, '','', ''，''),''"''+REPLACE(A5_Rem, '','', ''，'')+''"'')+'','' ' + #13#10 +
                  '     WHEN RTrim(A5_Rem) <> '''' and RTrim(A3_Rem) <> '''' Then RTrim(REPLACE(A3_Rem,'','', ''，''))+''-''+ Replace(REPLACE(A5_Rem, '','', ''，''),REPLACE(A5_Rem, '','', ''，''),''"''+REPLACE(A5_Rem, '','', ''，'')+''"'') +'',''' + #13#10 +
                  'ELSE '','' END 備註,' + #13#10 +
                  'Case absent_c WHEN 1 Then ''缺考,'' ELSE ''到考,'' END 出席情形,RC1.BarCode+'',,'',' + #13#10 +
                  'isnull(substring(Ans,dbo.AnsPos(28,Ans)+1,len(Ans)),'',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'') Ans2,dbo.AnsAddCom(Right(Ans_Status2,'+SubCnt+')) Ans3,' + #13#10 +
                  'CASE WHEN Ans_Status2 is Null THEN '','' ELSE CAST((Right_Count-29) as VARCHAR)+'','' END Right_Count ,' + #13#10 +
                  'CASE WHEN Ans_Status2 is Null THEN '','' ELSE CAST(ROUND((Cast(Right_Count as Float)-29)/'+SubCnt+'*100,2) as VARCHAR)+'','' END Rate' + #13#10 +
                  'from Sub_Score SS INNER JOIN student S on SS.Student_No = S.Student_No ' + #13#10 +
                  'INNER JOIN ReadCard1 RC1 on S.Student_no = LEFT(RC1.X_No,9) and Sub_no = Right(RC1.X_No,2) ' + #13#10 +
                  'INNER JOIN ReadCardRecTab RCT on S.student_no = RCT.student_no and RIGHT(RCT.Barcode,2) = '''+SubList[i]+'''' + #13#10 +
                  'where Sub_No = '''+SubList[i]+'''' + #13#10 +
                  'ORDER BY Std';
            SQLOpen(SQL);
            while not qry1.Eof do
            begin
                Data := '';
                for cnt := 1 to qry1.FieldCount - 1 do
                    Data := Data+ qry1.Fields[cnt].AsString;
                qry1.Next;
                Data := StringReplace(Data,',',#9,[rfReplaceAll, rfIgnoreCase]);
                idx := StdDataList.IndexOf(Trim(qry1.Fields[0].AsString));
                DataList.Add(StdDataList[idx]+Data);
            end;
            Excel := ExcelLab.create(Path+'TP'+SubList[i]+'.xlsx');

        end
        else
        begin
            //書寫提
            if SubList[i] = '02' then
               io := ',,,,,,,,,,,,,,,,,,'
            else
            begin
               io := ',,'
            end;
            SQL :='select ''108''+S.Sch_Code+S.Class_No+right(Seat_No,2) Std,' + #13#10 +
                  'Case WHEN A1 = 1 Then ''A 缺席（請假）,'' WHEN A2 = 1 Then ''B 缺席（轉出）,'' WHEN A3 = 1 Then ''C 更正座號,'' WHEN A4 = 1 Then ''D 轉入,'' WHEN A5 = 1 Then ''E 其他,'' ELSE '',''' + #13#10 +
                  'END 施測情形,' + #13#10 +
                  'Case WHEN RTrim(A3_Rem) <> '''' Then RTrim(REPLACE(A3_Rem, '','', ''，''))+'',''' + #13#10 +
                  '	   WHEN RTrim(A5_Rem) <> '''' Then Replace(REPLACE(A5_Rem, '','', ''，''),REPLACE(A5_Rem, '','', ''，''),''"''+REPLACE(A5_Rem, '','', ''，'')+''"'') +'',''' + #13#10 +
                  '     WHEN RTrim(A5_Rem) <> '''' and RTrim(A3_Rem) <> '''' Then RTrim(REPLACE(A3_Rem,'','', ''，''))+''-''+ Replace(REPLACE(A5_Rem, '','', ''，''),REPLACE(A5_Rem, '','', ''，''),''"''+REPLACE(A5_Rem, '','', ''，'')+''"'') +'','' ' + #13#10 +
                  'ELSE '','' END 備註,' + #13#10 +
                  'Case absent_c WHEN 1 Then ''缺考,'' ELSE ''到考,'' END 出席情形,' + #13#10 +
                  'RC1.BarCode+'','',Ans+'''+io+''' Ans1,' + #13#10 +
                  'CASE WHEN Ans_Status2 is NULL THEN Ans+'''+io+''' ELSE dbo.AnsAddCom(Right(Ans_Status2,'+SubCnt+'))+'''+io+''' END Ans2,' + #13#10 +
                  'CASE WHEN Ans_Status2 is Null THEN ''0,'' ELSE CAST(Right_Count as VARCHAR)+'','' END Right_Count ,' + #13#10 +
                  'CASE WHEN Ans_Status2 is Null THEN ''0,'' ELSE CAST(ROUND((Cast(Right_Count as Float))/'+SubCnt+'*100,2) as VARCHAR)+'','' END Rate' + #13#10 +
                  'from Sub_Score SS INNER JOIN student S on SS.Student_No = S.Student_No ' + #13#10 +
                  'INNER JOIN ReadCard1 RC1 on S.Student_no = LEFT(RC1.X_No,9) and Sub_no = Right(RC1.X_No,2) ' + #13#10 +
                  'INNER JOIN ReadCardRecTab RCT on S.student_no = RCT.student_no and RIGHT(RCT.Barcode,2) = '''+SubList[i]+'''' + #13#10 +
                  'where Sub_No = '''+SubList[i]+''' order by Std';
            SQLOpen(SQL);
            while not qry1.Eof do
            begin
                Data := '';
                for cnt := 1 to qry1.FieldCount - 1 do
                    Data := Data+ qry1.Fields[cnt].AsString;
                qry1.Next;
                Data := StringReplace(Data,',',#9,[rfReplaceAll, rfIgnoreCase]);
                DataList.Add(Data);
            end;
        end;
        Excel.Paste('A9',Datalist.Text);
        Excel.AutoReplace := true;
        Excel.Save;
        Excel.destory;
        DataList.SaveToFile('D:\'+SubList[i]+'.txt');
    end;
end;

procedure Tform1.GetRepExFile(const ExName, SavePath: String);
Var
  RCS : TResourceStream;
begin
  RCS := TResourceStream.Create(HInstance, ExName, 'EXCELFILE');
  RCS.SaveToFile(SavePath); //另存檔案
  RCS.Free;
end;
procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    StdDataList.Free;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    StdDataList := TstringlistEx.Create;
end;
function TstringlistEx.IndexOf(const S: string):Integer;
begin
  if not Sorted then
  begin
      for Result := 0 to Count - 1 do
          if pos(S,Strings[Result]) >0 then
              exit;
      Result := -1;
  end
  else if not Find(S, Result) then
    Result := -1;
end;
procedure TForm1.SQLOpen(SQL :string);
begin
   qry1.Close;
   qry1.SQL.Text := SQL;
   qry1.Open;
end;
procedure TForm1.SQLExec(SQL :string);
begin
   qry1.Close;
   qry1.SQL.Text := SQL;
   qry1.ExecSQL;
end;
end.



