program ScriptControlDemo;

uses
  Forms,
  ScriptControlDemoUnit in 'ScriptControlDemoUnit.pas' {Form1},
  ShowScriptFormUnit in 'ShowScriptFormUnit.pas' {ShowScriptForm};

{$R *.res}

begin
  Application.Initialize;
  ShowScriptForm := TShowScriptForm.Create(nil);
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
