program MasterDetailTut;

uses
  Forms,
  MasterDetailTutUnit1 in 'MasterDetailTutUnit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
