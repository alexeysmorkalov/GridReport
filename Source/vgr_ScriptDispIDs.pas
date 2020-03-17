{******************************************}
{                                          }
{           vtk GridReport library         }
{                                          }
{   Copyright (c) 2003-2004 by vtkTools    }
{                                          }
{******************************************}

unit vgr_ScriptDispIDs;

interface

type
{Describes the properties of script method or property.
Syntax:
  rvgrObjectScriptDesc = record
    DispId: Integer;
    GetMin: Integer;
    GetMax: Integer;
    SetMin: Integer;
    SetMax: Integer;
  end;}
  rvgrObjectScriptDesc = record
    DispId: Integer;
    GetMin: Integer; // if contans -2 then GET is not allowed
    GetMax: Integer; // if
    SetMin: Integer;
    SetMax: Integer;
  end;

{Array of rvgrObjectScriptDesc structures.}
  rvgrScriptInfo = array [0..16384] of rvgrObjectScriptDesc;

{Pointer to an array of rvgrObjectScriptDesc structures.}
  pvgrScriptInfo = ^rvgrScriptInfo;
  
const
  GetDisabled = -2;
  SetDisabled = -2;
  MethodDisabled = -2;
  AnyParams = -1;

  cs_PersistentMax = 10000;
  cs_PersistentRTTIPropsStart = 20000;

  cs_DatasetBof = 10000;
  cs_DatasetEof = 10001;
  cs_DatasetFirst = 10002;
  cs_DatasetLast = 10003;
  cs_DatasetNext = 10004;
  cs_DatasetPrior = 10005;
  cs_DatasetField = 10006;
  cs_DatasetLocate = 10007;
  cs_DatasetIifLocate = 10008;
  cs_DatasetOpen = 10009;
  cs_DatasetFieldsStart = 10010;

  siDatasetLength = 10;
  siDataset: array [0..siDatasetLength-1] of rvgrObjectScriptDesc =
   (
    (DispId: cs_DatasetBof; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetEof; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetFirst; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetLast; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetNext; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetPrior; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetField; GetMin: 1; GetMax: 1; SetMin: 2; SetMax: 2),
    (DispId: cs_DatasetLocate; GetMin: 4; GetMax: AnyParams; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetIifLocate; GetMin: 6; GetMax: AnyParams; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_DatasetOpen; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled)
    );

  cs_TvgrWorksheet_Cols = 10000;
  cs_TvgrWorksheet_Rows = 10001;
  cs_TvgrWorksheet_Borders = 10002;
  cs_TvgrWorksheet_Ranges = 10003;
  cs_TvgrWorksheet_HorzSections = 10004;
  cs_TvgrWorksheet_VertSections = 10005;
  cs_TvgrWorksheet_Max = cs_TvgrWorksheet_VertSections;

  siTvgrWorksheetLength = 6;
  siTvgrWorksheet: array [0..siTvgrWorksheetLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrWorksheet_Cols; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorksheet_Rows; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorksheet_Borders; GetMin: 3; GetMax: 3; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorksheet_Ranges; GetMin: 4; GetMax: 4; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorksheet_HorzSections; GetMin: 2; GetMax: 2; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorksheet_VertSections; GetMin: 2; GetMax: 2; SetMin: SetDisabled; SetMax: SetDisabled)
  );

  cs_TvgrVector_Number = 1;
  cs_TvgrVector_Size = 2;
  cs_TvgrVector_Visible = 3;
  cs_TvgrVector_Max = 3;

  siTvgrVectorLength = 3;
  siTvgrVector: array [0..siTvgrVectorLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrVector_Number; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrVector_Size; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrVector_Visible; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );

  cs_TvgrPageVector_PageBreak = cs_TvgrVector_Max + 1;
  cs_TvgrPageVector_Max = cs_TvgrPageVector_PageBreak;

  siTvgrPageVectorLength = 1;
  siTvgrPageVector: array [0..siTvgrPageVectorLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrPageVector_PageBreak; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );


  cs_TvgrCol_Width = cs_TvgrPageVector_Max + 1;
  cs_TvgrCol_Max = cs_TvgrCol_Width;

  siTvgrColLength = 1;
  siTvgrCol: array [0..siTvgrColLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrCol_Width; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );

  cs_TvgrRow_Height = cs_TvgrPageVector_Max + 1;
  cs_TvgrRow_Max = cs_TvgrRow_Height;

  siTvgrRowLength = 1;
  siTvgrRow: array [0..siTvgrRowLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrRow_Height; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );

  cs_TvgrRange_Left = 1;
  cs_TvgrRange_Top = 2;
  cs_TvgrRange_Right = 3;
  cs_TvgrRange_Bottom = 4;
  cs_TvgrRange_Value = 5;
  cs_TvgrRange_Font = 6;
  cs_TvgrRange_FillForeColor = 7;
  cs_TvgrRange_FillBackColor = 8;
  cs_TvgrRange_FillPattern = 9;
  cs_TvgrRange_DisplayFormat = 10;
  cs_TvgrRange_HorzAlign = 11;
  cs_TvgrRange_VertAlign = 12;
  cs_TvgrRange_Angle = 13;
  cs_TvgrRange_WordWrap = 14;
  cs_TvgrRange_Name = 15;
  cs_TvgrRange_Style = 16;
  cs_TvgrRange_Color = 17;
  cs_TvgrRange_Size = 18;
  cs_TvgrRange_Formula = 19;
  cs_TvgrRange_Charset = 20;
  cs_TvgrRange_Max = 20;

  siTvgrRangeLength = 20;
  siTvgrRange: array [0..siTvgrRangeLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrRange_Left; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrRange_Top; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrRange_Right; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrRange_Bottom; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrRange_Value; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Font; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrRange_FillForeColor; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_FillBackColor; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_FillPattern; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_DisplayFormat; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_HorzAlign; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_VertAlign; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Angle; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_WordWrap; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Name; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Style; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Color; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Size; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Formula; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrRange_Charset; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );


  cs_TvgrBorder_Left = 1;
  cs_TvgrBorder_Top = 2;
  cs_TvgrBorder_Orientation = 3;
  cs_TvgrBorder_Width = 4;
  cs_TvgrBorder_Color = 5;
  cs_TvgrBorder_Pattern = 6;

  siTvgrBorderLength = 6;
  siTvgrBorder: array [0..siTvgrBorderLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrBorder_Left; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrBorder_Top; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrBorder_Orientation; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrBorder_Width; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrBorder_Color; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrBorder_Pattern; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );


  cs_TvgrSection_StartPos = 1;
  cs_TvgrSection_EndPos = 2;
  cs_TvgrSection_Level = 3;
  cs_TvgrSection_RepeatOnPageTop = 4;
  cs_TvgrSection_RepeatOnPageBottom = 5;
  cs_TvgrSection_PrintWithNextSection = 6;
  cs_TvgrSection_PrintWithPreviosSection = 7;

  siTvgrSectionLength = 7;
  siTvgrSection: array [0..siTvgrSectionLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrSection_StartPos; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrSection_EndPos; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrSection_Level; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrSection_RepeatOnPageTop; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrSection_RepeatOnPageBottom; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrSection_PrintWithNextSection; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1),
    (DispId: cs_TvgrSection_PrintWithPreviosSection; GetMin: 0; GetMax: 0; SetMin: 1; SetMax: 1)
  );


  cs_TvgrWorkbook_Worksheets = cs_PersistentMax;
  cs_TvgrWorkbook_WorksheetsCount = cs_PersistentMax + 1;
  cs_TvgrWorkbook_Clear = cs_PersistentMax + 2;
  cs_TvgrWorkbook_AddWorksheet = cs_PersistentMax + 3;
  cs_TvgrWorkbook_SaveToFile = cs_PersistentMax + 4;
  cs_TvgrWorkbook_LoadFromFile = cs_PersistentMax + 5;
  cs_TvgrWorkbook_Max = cs_TvgrWorkbook_LoadFromFile;

  siTvgrWorkbookLength = 6;
  siTvgrWorkbook: array [0..siTvgrWorkbookLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrWorkbook_Worksheets; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorkbook_WorksheetsCount; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorkbook_Clear; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorkbook_AddWorksheet; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorkbook_SaveToFile; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrWorkbook_LoadFromFile; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled)
  );

  cs_TvgrBand_GenBegin = cs_PersistentMax;
  cs_TvgrBand_GenEnd = cs_PersistentMax + 1;
  cs_TvgrBand_Max = cs_TvgrBand_GenEnd;

  siTvgrBandLength = 2;
  siTvgrBand: array [0..siTvgrBandLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrBand_GenBegin; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrBand_GenEnd; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled)
  );


  cs_TvgrReportTemplateWorksheet_CurrentRow = cs_TvgrWorksheet_Max + 1;
  cs_TvgrReportTemplateWorksheet_CurrentColumn = cs_TvgrWorksheet_Max + 2;
  cs_TvgrReportTemplateWorksheet_CurrentColumnCaption = cs_TvgrWorksheet_Max + 3;
  cs_TvgrReportTemplateWorksheet_WriteLnBand = cs_TvgrWorksheet_Max + 4;
  cs_TvgrReportTemplateWorksheet_WriteLnCell = cs_TvgrWorksheet_Max + 5;
  cs_TvgrReportTemplateWorksheet_WriteCell = cs_TvgrWorksheet_Max + 6;
  cs_TvgrReportTemplateWorksheet_WriteLn = cs_TvgrWorksheet_Max + 7;
  cs_TvgrReportTemplateWorksheet_WorkbookRow = cs_TvgrWorksheet_Max + 8;
  cs_TvgrReportTemplateWorksheet_WorkbookColumn = cs_TvgrWorksheet_Max + 9;
  cs_TvgrReportTemplateWorksheet_WorkbookColumnCaption = cs_TvgrWorksheet_Max + 10;
  cs_TvgrReportTemplateWorksheet_Max = cs_TvgrReportTemplateWorksheet_WorkbookColumnCaption;

  siTvgrReportTemplateWorksheetLength = 10;
  siTvgrReportTemplateWorksheet: array [0..siTvgrReportTemplateWorksheetLength - 1] of rvgrObjectScriptDesc =
  (
    (DispId: cs_TvgrReportTemplateWorksheet_CurrentRow; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_CurrentColumn; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_CurrentColumnCaption; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WriteLnBand; GetMin: 1; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WriteLnCell; GetMin: 2; GetMax: 2; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WriteCell; GetMin: 2; GetMax: 2; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WriteLn; GetMin: 0; GetMax: 1; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WorkbookRow; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WorkbookColumn; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled),
    (DispId: cs_TvgrReportTemplateWorksheet_WorkbookColumnCaption; GetMin: 0; GetMax: 0; SetMin: SetDisabled; SetMax: SetDisabled)
  );

implementation

end.
