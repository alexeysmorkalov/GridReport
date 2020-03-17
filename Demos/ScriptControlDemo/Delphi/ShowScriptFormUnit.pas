unit ShowScriptFormUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, vgr_ScriptEdit, StdCtrls;

type
  TShowScriptForm = class(TForm)
    vgrScriptEdit1: TvgrScriptEdit;
    Panel1: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    Button2: TButton;
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ShowScriptForm: TShowScriptForm;

implementation

uses ScriptControlDemoUnit;

{$R *.dfm}

procedure TShowScriptForm.RadioButton1Click(Sender: TObject);
begin
  Form1.InitScriptContol;
end;

procedure TShowScriptForm.RadioButton2Click(Sender: TObject);
begin
  Form1.InitScriptContol;
end;

procedure TShowScriptForm.Button2Click(Sender: TObject);
begin
  if RadioButton1.Checked = True then
    vgrScriptEdit1.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + cVBScriptFile);
  if RadioButton2.Checked = True then
    vgrScriptEdit1.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + cJScriptFile);
  Form1.InitScriptContol;    
end;

end.
