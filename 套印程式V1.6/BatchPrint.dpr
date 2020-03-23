program BatchPrint;

uses
  Forms,
  NATL in 'NATL.pas' {Form1},
  Grade in 'Grade.pas' {Form4},
  Login in 'Login.pas' {LoginCheck};

{Form3}


{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;

  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
