unit ExcelUnit;
interface
uses Windows, Messages, SysUtils, Variants, Classes,Excel2000,ComObj,ClipBrd,Dialogs,StrUtils,DB ,ADODB;
type
    ExcelLab = class
      private
        Excelapp :variant;
        F :Tstringlist;
        ExPath :String;
        ExAutoReplace :Boolean;
        ExAStyle :olevariant;
        function getCell(i,j :integer):string;overload;
        procedure SetCell(i,j :integer;S :string);overload;
        function getCell(R :String):string;overload;
        procedure SetCell(R :String;S :string);overload;
        function getCell(R,C :String):string;overload;
        procedure SetCell(R,C :String;S :string);overload;
        function GetRowCount:integer;
        function GetColCount:integer;
        procedure ExPage(Page :integer);
        procedure SetFontName(R,C,Name :string);
        procedure SetFontColor(R,C :string;Index :integer);
        procedure SetFontSize(R,C :string;Value :integer);
        procedure SetFontBold(R,C :string;Value :Boolean);
        procedure SetFontItalic(R,C :string;Value :Boolean);
        procedure SetFontUnderline(R,C :string;Value :Boolean);
        procedure SetHorizontalAlignment(R,C :string;Value :integer);
        procedure SetVerticalAlignment(R,C :string;Value :integer);
        procedure SetBorders(R,C :String;B,Value :integer);
      public
        constructor create(Page :integer);overload;
        constructor create(Path :string);overload;
        destructor destory;
        function RowList(row :string) :Tstringlist;
        procedure Paste(R,C,S :String);overload;
        procedure Paste(R,S :String);overload;
        procedure Clear(R,C :String);
        function Copy(R,C :string):string;
        procedure CutAndPaste(R,C,I :string);
        procedure Save;overload;
        procedure Save(Path :string);overload;
        procedure QueryDataAdd(Recordset :_Recordset;R :string);
        procedure CellReplace(FindStr,ReplaceStr :string;LookAt,SearchOrder :integer);
        property Excell :variant read Excelapp write Excelapp;
        property Page :integer write Expage;
        property RowCount :integer read GetRowCount;
        property ColCount :integer read GetColCount;
        property Path :string read ExPath write Expath;
        property Cell[i,j :integer] :string read getCell write SetCell;
        property RangeCell[R :String] :string read getCell write SetCell;
        property RangeCellRC[R,C :String] :string read getCell write SetCell;
        property AutoReplace :Boolean read ExAutoReplace write ExAutoReplace;
        property FontName[R,C :String] :String write SetFontName;
        property FontColor[R,C :String] :integer write SetFontColor;
        property FontBold[R,C :String] :Boolean write SetFontBold;
        property FontItalic[R,C :String] :Boolean write SetFontItalic;
        property FontSize[R,C :String] :integer write SetFontSize;
        procedure Font(R,C,Name :String;Color,Size :integer;Bold,Italic :Boolean);
        property FontUnderLine[R,C :String] :Boolean write SetFontUnderline;
        property VerticalAlignment[R,C :String] :Integer write SetVerticalAlignment;
        property HorizontalAlignment[R,C :String] :Integer write SetHorizontalAlignment;
        property Borders[R,C :String;B:integer] :integer write SetBorders;
        property AStyle :olevariant read ExAStyle write ExAstyle;
        procedure AutoFit;
        procedure cellselect(R,C :String);
    end;
implementation
constructor ExcelLab.create(Page :integer);
//var AStyle : OleVariant;
begin
    Self.Path := GetCurrentDir + '\default.xlsx';
    Excelapp := CreateOleObject('Excel.Application');
    ExcelApp.WorkBooks.Add(Page);
    ExcelApp.Visible := False;
    ExcelApp.WorkSheets[1].Activate;
    AStyle := '@';
    ExcelApp.Cells.NumberFormatLocal := AStyle;
end;
constructor ExcelLab.create(Path :string);
begin
    self.Path := Path;
    Excelapp := CreateOleObject('Excel.Application');
    Excelapp.WorkBooks.Open(path);
    Excelapp.WorkBooks[1].Activate;
    ExcelApp.Visible := False;
    F := TStringList.Create;
end;
procedure ExcelLab.autofit;
begin
  ExcelApp.Selection.Columns.AutoFit;
end;
procedure ExcelLab.ExPage(Page :integer);
begin
    if Excelapp.Worksheets.Count < Page then
    begin
        ShowMessage('Wordsheets error');
        exit;
    end;
    ExcelApp.WorkSheets[Page].Activate;
end;
function ExcelLab.GetRowCount:integer;
begin
    Result :=  ExcelApp.ActiveSheet.UsedRange.Rows.Count
end;
function ExcelLab.GetColCount:integer;
begin
    Result :=  ExcelApp.ActiveSheet.UsedRange.Columns.Count
end;
procedure ExcelLab.SetCell(i,j :integer;S :string);
begin
    Excelapp.Cells[i,j].value := S;
end;
function ExcelLab.GetCell(i,j :integer):string;
begin
    result := Excelapp.Cells[i,j].value;
end;
function ExcelLab.GetCell(R :string):string;
begin
    result := ExcelApp.Range[R].Value;
end;
function ExcelLab.GetCell(R,C :string):string;
begin
    result := ExcelApp.Range[R+':'+C].Value;
end;
procedure ExcelLab.SetCell(R,S :string);
begin
    ExcelApp.Range[R].Value := S;
end;
procedure ExcelLab.SetCell(R,C,S :string);
begin
    ExcelApp.Range[R+':'+C].Value := S;
end;
procedure ExcelLab.SetFontName(R,C,Name :string);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.name := Name;
end;
procedure ExcelLab.SetFontSize(R,C :string;Value :integer);
begin
  ExcelApp.Range[R+':'+C].Select;
  ExcelApp.Selection.Font.Size := Value;
end;
procedure ExcelLab.SetFontColor(R,C :string;Index :integer);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.ColorIndex := Index;
end;
procedure ExcelLab.SetFontBold(R,C :string;Value :Boolean);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.Bold := Value;
end;
procedure ExcelLab.SetFontItalic(R,C :string;Value :Boolean);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.Italic := Value;
end;
procedure ExcelLab.Font(R,C,Name :String;Color,Size :integer;Bold,Italic :Boolean);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.name := Name;
    ExcelApp.Selection.Font.Size := Size;
    ExcelApp.Selection.Font.ColorIndex := Color;
    ExcelApp.Selection.Font.Bold := Bold;
    ExcelApp.Selection.Font.Italic := Italic;
end;
procedure ExcelLab.SetFontUnderline(R,C :string;Value :Boolean);
begin
    ExcelApp.Range[R+':'+C].Select;
    ExcelApp.Selection.Font.UnderLine := Value;
end;
procedure ExcelLab.SetHorizontalAlignment(R,C :string;Value :integer);
begin
    ExcelApp.Range[R+':'+C].Select;
    case Value of
        0:begin
            ExcelApp.Selection.HorizontalAlignment := xlHAlignLeft;
        end;
        1:begin
            ExcelApp.Selection.HorizontalAlignment := xlHAlignCenter;
        end;
        2:begin
            ExcelApp.Selection.HorizontalAlignment := xlHAlignRight;
        end;
    end;
end;
procedure ExcelLab.SetVerticalAlignment(R,C :string;Value :integer);
begin
    ExcelApp.Range[R+':'+C].Select;
    case Value of
        0:begin
            ExcelApp.Selection.VerticalAlignment := xlHAlignLeft;
        end;
        1:begin
            ExcelApp.Selection.VerticalAlignment := xlHAlignCenter;
        end;
        2:begin
            ExcelApp.Selection.VerticalAlignment := xlHAlignRight;
        end;
    end;
end;
procedure ExcelLab.Paste(R,C,S :string);
begin
    Clipboard.Clear;
    Clipboard.AsText := S;
    ExcelApp.Range[R+':'+C].PasteSpecial;
end;
procedure ExcelLab.Paste(R,S :string);
begin
    Clipboard.Clear;
    Clipboard.AsText := S;
    ExcelApp.Range[R].PasteSpecial;
end;
procedure ExcelLab.Clear(R,C :string);
begin
    ExcelApp.Range[R+':'+C].ClearContents;
end;
procedure ExcelLab.cellselect(R,C :String);
begin
    ExcelApp.Range[R+':'+C].Select;
end;
procedure ExcelLab.CutAndPaste(R,C,I :string);
begin
    ExcelApp.Range[R+':'+C].Cut;
    ExcelApp.Range[I].Insert(xlDown);
end;
procedure ExcelLab.SetBorders(R,C :String;B,Value :integer);
begin
     ExcelApp.Range[R,C].Borders[B].Weight := Value; //先畫第一條
end;
Function ExcelLab.Copy(R,C :string):string;
begin
    ExcelApp.Range[ R+':'+C ].Copy;
    Result := Clipboard.AsText;
    Clipboard.Clear;
end;
procedure ExcelLab.Save;
begin
    if ExAutoReplace then
       ExcelApp.DisplayAlerts := false;
    if pos('.pdf',path) <= 0 then
        ExcelApp.WorkBooks[1].SaveAs(ExPath)
    else
        ExcelApp.WorkBooks[1].ExportAsFixedFormat(0,ExPath);
    if ExAutoReplace then
       ExcelApp.DisplayAlerts := True;
end;

procedure ExcelLab.Save(Path :String);
begin
    if ExAutoReplace then
       ExcelApp.DisplayAlerts := false;
    if pos('.pdf',path) <= 0 then
        ExcelApp.WorkBooks[1].SaveAs(Path)
    else
        ExcelApp.WorkBooks[1].ExportAsFixedFormat(0,Path);
    if ExAutoReplace then
       ExcelApp.DisplayAlerts := True;
end;
function ExcelLab.RowList(row :string):Tstringlist;
begin
    Clipboard.Clear;
    ExcelApp.Range[ row+'2:'+row+''+inttostr(RowCount)+'' ].Copy;
    F.text := Clipboard.AsText;
    Result := F;
end;
procedure ExcelLab.CellReplace(FindStr,ReplaceStr :string;LookAt,SearchOrder :integer);
begin
    ExcelApp.Cells.Replace(FindStr, ReplaceStr, LookAt, SearchOrder,false,false,False, False);
end;
procedure ExcelLab.QueryDataAdd(Recordset :_Recordset;R :string);
var xlQuery :variant;
begin
    xlQuery := ExcelApp.WorkSheets[1].QueryTables.Add(Recordset, ExcelApp.Range[R]);
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
end;
destructor ExcelLab.destory;
begin
    Clipboard.Clear;
    ExcelApp.WorkBooks.close;  //Ãö³¬Excel
    ExcelApp.Quit;             //Â÷¶}Excel
    ExcelApp:=Unassigned;
    F.Free;
end;
end.

