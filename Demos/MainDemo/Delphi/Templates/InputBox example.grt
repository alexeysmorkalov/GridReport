TPF0TvgrReportTemplateReportTemplateScript.LanguageVBScriptScript.Script.StringsPublic CurRow, InputCountryName 4sub vgrDetailBand1_Events_BeforeGenerate(ByVal Band)  CurRow = 2&  'Input name of country to highlight:U  InputCountryName = InputBox("Please enter name of country to highlight in report.")end sub 1sub vgrDataBand1_Events_AfterGenerate(ByVal Band)  Dim ARange  ' Fill background of row:%  if CurRow/2 <> Round(CurRow/2) thenA    Highlightrow Workbook.Worksheets(0), CurRow, RGB(220,255,255)  elseA    Highlightrow Workbook.Worksheets(0), CurRow, RGB(250,255,255)  end if   ' get value of "Country" rangeC  Set ARange = Workbook.Worksheets(0).Ranges(2,CurRow,2,CurRow)		  )  ' Highlighting InputCountryName country)  If ARange.Value = InputCountryName then9    Highlightrow Workbook.Worksheets(0), CurRow, clYellow  end if  CurRow = CurRow + 1end sub W' Procedure highlights the five horizontal cells in the row specified by ARow parameter' All merges are deleted' Parameters:-'  pWorksheet - reference to Worksheet object'  ARow - row to highlight'  AColor - color*Sub Highlightrow(pWorksheet, ARow, AColor)  with pWorksheet1    .Ranges(0,ARow,0,ARow).FillBackColor = AColor1    .Ranges(1,ARow,1,ARow).FillBackColor = AColor1    .Ranges(2,ARow,2,ARow).FillBackColor = AColor1    .Ranges(3,ARow,3,ARow).FillBackColor = AColor1    .Ranges(4,ARow,4,ARow).FillBackColor = AColor
  End withEnd Sub Script.AliasManagerMainForm.AliasManagerLeft`Top(DataStorageVersion
SystemInfoOS: WIN32_NT 5.1.2600  PageSize: 4096ActiveProcessorMask: $1000NumberOfProcessors: 1ProcessorType: 586 Compiler version: Delphi7DataStorage version: 1 RangeStylesData
     >      m�   
     Arialtools\commoncontrols\sourc  ����              M�    
     Arialtools\commoncontrols\sourc  ����               ~�   
    Arialtools\commoncontrols\sourc  � ��             ^�   
    Arialtools\commoncontrols\sourc  � ��              n�   
     Arialtools\commoncontrols\sourc  � ��               O@   
    Arialtools\commoncontrols\sourc  � ��              O    
    Arialtools\commoncontrols\sourc  � ��              N       Arialtools\commoncontrols\sourc��� ���            N�   
    Arialtools\commoncontrols\sourc  � ��               N�   
    Arialtools\commoncontrols\sourc  � ��               N�   
    Arialtools\commoncontrols\sourc  � ��               ]�   
     Arialtools\commoncontrols\sourc  ����             }�   
     Arialtools\commoncontrols\sourc  ����            ]�   
     Arialtools\commoncontrols\sourc  ����             M�    
     Arialtools\commoncontrols\sourc  ����              BorderStylesData
      	                    StringsData
�               7   =COUNT(A[vgrDataBand1.GenBegin]:A[vgrDataBand1.GenEnd])      All Customers Count:       [GlobalDM.Customers.CustNo]      [GlobalDM.Customers.Company]      [GlobalDM.Customers.Country]   6   [GlobalDM.Customers.Addr1]
[GlobalDM.Customers.Addr2]      [GlobalDM.Customers.Phone]      Company      Address      Country      Country             Address             Phone      CustNo      InputBox exampleFormulasData
b  
             &                                   '                                  �                                                                                                          �F          M Y�                  M Y�                                         ����������������                                TvgrReportTemplateWorksheet-vgrReportTemplate1vgrReportTemplateWorksheet1TitleSheet1PageProperties.Height�APageProperties.Width�. PageProperties.MeasurementSystemvgrmsMetricColsData
4            �       R       �           RowsData
)             v       �       �    HorzSectionsData
    
RangesData
   $                           f�  ������              
      f�  ������            	      f�  ������                  f�  ������                  f�  ������                  f�  ������                    f�  ������                  f�  ������                  f�  ������                  f�  ������                  f�  ������                   f�  ������                  f�  ������                    \������� TvgrBandvgrBand1StartPos EndPos RepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSectionVertical  TvgrDetailBandvgrDetailBand1StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSection#Events.BeforeGenerateScriptProcName$vgrDetailBand1_Events_BeforeGenerateDataSetGlobalDM.CustomersVertical  TvgrBandvgrBand2StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSectionVertical  TvgrDataBandvgrDataBand1StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSection"Events.AfterGenerateScriptProcName!vgrDataBand1_Events_AfterGenerateVertical    