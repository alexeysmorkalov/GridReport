program GroupsTut;

uses
  Forms,
  GroupsTutUnit1 in 'GroupsTutUnit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
