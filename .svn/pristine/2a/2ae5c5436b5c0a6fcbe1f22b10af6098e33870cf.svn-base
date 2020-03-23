program BatchPrint;

uses
  Forms,
  NATL in 'NATL.pas' {Form1},
  Login in 'Login.pas' {LoginCheck},
  UnitMyThread in 'UnitMyThread.pas',
  Table in 'Table.pas' {FormTable},
  ExcelUnit in 'ExcelUnit.pas',
  Grade in 'Grade.pas' {Print},
  Rep in 'Rep.pas' {RepExport},
  SettingSet in 'SettingSet.pas' {SSet},
  SQLtext in 'SQLtext.pas' {SQLSetting};

{Form3}


{$R *.res}
{$R RepEx.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TPrint, Print);
  Application.CreateForm(TRepExport, RepExport);
  Application.CreateForm(TSSet, SSet);
  Application.CreateForm(TSQLSetting, SQLSetting);
  //  Application.CreateForm(TFormTable, FormTable);
  Application.Run;
end.
