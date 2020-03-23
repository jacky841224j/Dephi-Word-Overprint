program BatchPrint;

uses
  Forms,
  NATL in 'NATL.pas' {Form1},
  Grade in 'Grade.pas' {Form4},
  Login in 'Login.pas' {LoginCheck},
  UnitMyThread in 'UnitMyThread.pas',
  Table in 'Table.pas' {FormTable},
  ExcelUnit in 'ExcelUnit.pas',
  Rep in 'Rep.pas' {RepExport};

{Form3}


{$R *.res}
{$R RepEx.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TRepExport, RepExport);
  //  Application.CreateForm(TFormTable, FormTable);
  Application.Run;
end.
