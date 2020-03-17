program simplecrosstab;

uses
  Forms,
  simplecrosstab_unit in 'simplecrosstab_unit.pas' {Form1};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
