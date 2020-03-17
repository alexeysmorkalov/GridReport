unit ScriptControlDemoUnit;

interface

{$I vtk.inc}

uses
  Windows, Messages, SysUtils, {$IFDEF VTK_D6_OR_D7} Variants, {$ENDIF}Classes, Graphics, Controls, Forms,
  Dialogs, vgr_ScriptEdit, vgr_CommonClasses, vgr_ScriptControl, StdCtrls,
  Buttons, ExtCtrls;

type
  TForm1 = class(TForm)
    vgrScriptControl1: TvgrScriptControl;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    ResultLabel: TLabel;
    SpeedButton6: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton20Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton19Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton12Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure SpeedButton13Click(Sender: TObject);
    procedure SpeedButton18Click(Sender: TObject);
    procedure SpeedButton15Click(Sender: TObject);
    procedure SpeedButton16Click(Sender: TObject);
    procedure SpeedButton17Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SpeedButton6Click(Sender: TObject);
  private
    { Private declarations }
    procedure ExecuteCommandInScript(Sender: TObject);
  public
    { Public declarations }
    procedure InitScriptContol;    
  end;

var
  Form1: TForm1;

const
  cVBScriptFile = 'script.vbs';
  cJScriptFile = 'script.js';
  
implementation

uses ShowScriptFormUnit;

{$R *.dfm}



procedure TForm1.ExecuteCommandInScript(Sender: TObject);
var
  ASpeedButton: TSpeedButton;
  AProperties: TvgrVariantDynArray;
begin
  ASpeedButton := Sender as TSpeedButton;
  SetLength(AProperties, 1);
  AProperties[0] := ASpeedButton.Caption;
  ResultLabel.Caption := vgrScriptControl1.RunProcedure('Calc',AProperties);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  InitScriptContol;
end;

procedure TForm1.SpeedButton20Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //End executing script
  vgrScriptControl1.EndScriptExecuting;
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton19Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton7Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton8Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton12Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton11Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton9Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton4Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton14Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton13Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton18Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton15Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton16Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton17Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton3Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.SpeedButton5Click(Sender: TObject);
begin
  ExecuteCommandInScript(Sender);
end;

procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  AProperties: TvgrVariantDynArray;
begin
  if Key = VK_DELETE then
    Key := Ord('C');
  if Key = VK_ADD then
    Key := Ord('+');
  if Key = VK_SUBTRACT then
    Key := Ord('-');
  if Key = VK_MULTIPLY then
    Key := Ord('*');
  if Key = VK_DIVIDE then
    Key := Ord('/');
  if Key = VK_NUMPAD0 then
    Key := Ord('0');
  if Key = VK_NUMPAD1 then
    Key := Ord('1');
  if Key = VK_NUMPAD2 then
    Key := Ord('2');
  if Key = VK_NUMPAD3 then
    Key := Ord('3');
  if Key = VK_NUMPAD4 then
    Key := Ord('4');
  if Key = VK_NUMPAD5 then
    Key := Ord('5');
  if Key = VK_NUMPAD6 then
    Key := Ord('6');
  if Key = VK_NUMPAD7 then
    Key := Ord('7');
  if Key = VK_NUMPAD8 then
    Key := Ord('8');
  if Key = VK_NUMPAD9 then
    Key := Ord('9');
  if Key = VK_DECIMAL then
    Key := Ord(',');
  if Key = VK_RETURN then
    Key := Ord('=');
  SetLength(AProperties, 1);
  if (Char(Key) in ['0'..'9'])or(Char(Key)='C')or(Char(Key)='+')or(Char(Key)='-')or(Char(Key)='/')or(Char(Key)='*')or(Char(Key)='=') then
  begin
    AProperties[0] := Chr(Key);
    ResultLabel.Caption := vgrScriptControl1.RunProcedure('Calc',AProperties);
  end;
end;

//procedure initialize script
procedure TForm1.InitScriptContol;
begin
  vgrScriptControl1.EndScriptExecuting;
  //For VBScript:
  if ShowScriptForm.RadioButton1.Checked = True then
  begin
    vgrScriptControl1.Language := 'VBScript';  
    ShowScriptForm.vgrScriptEdit1.Script.Language := 'VBScript';
    ShowScriptForm.vgrScriptEdit1.Lines.LoadFromFile(ExtractFilePath(Application.ExeName) + cVBScriptFile);
  end;
  //For JScript:
  if ShowScriptForm.RadioButton2.Checked = True then
  begin
    vgrScriptControl1.Language := 'JScript';
    ShowScriptForm.vgrScriptEdit1.Script.Language := 'JScript';
    ShowScriptForm.vgrScriptEdit1.Lines.LoadFromFile(ExtractFilePath(Application.ExeName) + cJScriptFile);
  end;
  vgrScriptControl1.Script.Text := ShowScriptForm.vgrScriptEdit1.Lines.Text;
  //Begin executing script
  vgrScriptControl1.BeginScriptExecuting;
  //Run procedure in script to initialize script variables
  ResultLabel.Caption := vgrScriptControl1.RunProcedure('Init');
end;

procedure TForm1.SpeedButton6Click(Sender: TObject);
begin
  ShowScriptForm.Show;
end;

end.
