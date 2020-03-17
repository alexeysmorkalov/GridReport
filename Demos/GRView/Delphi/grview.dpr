program grview;

uses
  Forms,
  vgr_WorkbookDesigner,
  vgr_Report,
  vgr_ReportDesigner,
  vgr_DataStorage,
  Sysutils,
  registry,
  classes,
  Windows;

{$R *.res}

var
  AInputFile: string;
  AInputFileExt:string;
  Reg: TRegistry;
  rt: TvgrReportTemplate;
  rtd: TvgrReportDesigner;
  rtdf: TvgrReportDesignerForm;
  wb: TvgrWorkbook;
  wbd: TvgrWorkbookDesigner;
  wbdf: TvgrWorkbookDesignerForm;
begin
  Application.Initialize;
  Application.Title := 'GRView';


    Reg := nil;
    try
      Reg := TRegistry.Create;
      Reg.RootKey := HKEY_CLASSES_ROOT;
      if (Reg.OpenKey('\.grw', false) = false) or (ParamStr(1) = 'reg') then
      begin
        if Application.MessageBox(
          'Do you want to register GRView by default viewer for .grt and .grw files?',
          'Message',
          MB_YESNO) = IDYES then
         begin
           Reg.OpenKey('\.grw', true);
           Reg.WriteString('', 'GRView');
           Reg.CloseKey;
           Reg.OpenKey('\.grt', true);
           Reg.WriteString('', 'GRView');
           Reg.CloseKey;
           Reg.OpenKey('\GRView', true);
           Reg.WriteString('', 'Grid Report Viewer');
           Reg.CloseKey;
           Reg.OpenKey('\GRView\Shell\Open\Command', true);
           Reg.WriteString('', ParamStr(0) + ' "%1"');
           Reg.CloseKey;
           Reg.OpenKey('\GRView\DefaultIcon', true);
           Reg.WriteString('', ParamStr(0) + ',0');
           Reg.CloseKey;
         end;
      end;
    finally
      if Assigned(Reg) then Reg.Destroy;
    end;
  AInputFile := ParamStr(1);
  AInputFileExt := ExtractFileExt(AInputFile);
  if not FileExists(AInputFile) or ( (AInputFileExt <> '.grw') and (AInputFileExt <> '.grt')) then
  begin
    Exit;
  end
  else
  begin
    if(AInputFileExt = '.grw') then
    begin
      wb := TvgrWorkbook.Create(Application);
      wb.LoadFromFile(AInputFile);
      wbd := TvgrWorkbookDesigner.Create(Application);
      wbd.Workbook := wb;
      Application.CreateForm(TvgrWorkbookDesignerForm, wbdf);
      wbdf.Workbook := wb;
      wbdf.DocCaption := AInputFile;
    end;
    if(AInputFileExt = '.grt') then
    begin
      rt := TvgrReportTemplate.Create(Application);
      rt.LoadFromFile(AInputFile);
      rtd := TvgrReportDesigner.Create(Application);
      rtd.Template := rt;
      Application.CreateForm(TvgrReportDesignerForm, rtdf);
      rtdf.Workbook := rt;
      rtdf.DocCaption := AInputFile;
    end;
  end;
  Application.Run;
end.
