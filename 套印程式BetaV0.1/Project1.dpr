program Project1;

uses
  Forms,
  NATL in 'NATL.pas' {Form1},
  Grade in 'Grade.pas' {Form4},
  Site in 'Site.pas' {Roster};

{Form3}


{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TRoster, Roster);
  //  Application.CreateForm(TForm4, Form4);
  //  Application.CreateForm(TForm2, Form2);
//  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
