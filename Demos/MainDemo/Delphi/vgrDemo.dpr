program vgrDemo;

uses
  Forms,
  Main_Form in 'Main_Form.pas' {MainForm},
  Global_DM in 'Global_DM.pas' {GlobalDM: TDataModule},
  WriteXXExamples_Form in 'WriteXXExamples_Form.pas' {WriteXXExamplesForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '#Report demo';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TGlobalDM, GlobalDM);
  Application.CreateForm(TWriteXXExamplesForm, WriteXXExamplesForm);
  Application.Run;
end.
