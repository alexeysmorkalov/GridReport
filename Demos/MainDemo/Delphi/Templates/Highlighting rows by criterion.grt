TPF0TvgrReportTemplateReportTemplateScript.LanguageVBScriptScript.Script.Strings#' Define variable for counting rowsPublic CurRow @' This procedure is called when generating of detail band begins5Sub vgrDetailBand1_Events_BeforeGenerate (ByVal Band)  ' First row of data band  CurRow = 2End Sub =' This procedure is called when row of data band is generated2Sub vgrDataBand1_Events_AfterGenerate (ByVal Band)  Dim ARange  ' Fill background of row:%  if CurRow/2 <> Round(CurRow/2) then    ' even rowA    Highlightrow Workbook.Worksheets(0), CurRow, RGB(220,255,255)  else    ' odd rowA    Highlightrow Workbook.Worksheets(0), CurRow, RGB(250,255,255)  end if   ' get value of "Country" rangeC  Set ARange = Workbook.Worksheets(0).Ranges(2,CurRow,2,CurRow)		    ' Highlighting "US" country  If ARange.Value = "US" then9    Highlightrow Workbook.Worksheets(0), CurRow, clYellow"    ARange.Value = "United States"  end if  CurRow = CurRow + 1End Sub W' Procedure highlights the five horizontal cells in the row specified by ARow parameter' All merges are deleted' Parameters:-'  pWorksheet - reference to Worksheet object'  ARow - row to highlight'  AColor - color*Sub Highlightrow(pWorksheet, ARow, AColor)  with pWorksheet1    .Ranges(0,ARow,0,ARow).FillBackColor = AColor1    .Ranges(1,ARow,1,ARow).FillBackColor = AColor1    .Ranges(2,ARow,2,ARow).FillBackColor = AColor1    .Ranges(3,ARow,3,ARow).FillBackColor = AColor1    .Ranges(4,ARow,4,ARow).FillBackColor = AColor
  End withEnd Sub Script.AliasManagerMainForm.AliasManagerLeft� TopxDataStorageVersion
SystemInfo$OS: WIN32_NT 5.1.2600 Service Pack 1 PageSize: 4096ActiveProcessorMask: $1000NumberOfProcessors: 1ProcessorType: 586 Compiler version: Delphi7DataStorage version: 1 RangeStylesData
�     >      ��,   
     Arial ���         �:� �:� $      �  �  �          �|       Tahoma���         �:� �:� $      � ��   �           �}   
     Tahoma���         �:� �:� $      ����  �         �y   
     Tahoma���         �:� �:� $      ����  �         �   
     Tahoma���         �:� �:� $      ����  �         ��       Tahoma���         �:� �:� $      � ��   �          ��       Arial ���         �:� �:� $    ��� ���  �          �)       Arial ���         �:� �:� $      � ��   �         �+       Arial ���         �:� �:� $      � ��   �          @{        Tahoma���         �:� �:� $      ����   �           �}   
     Tahoma���         �:� �:� $      ����  �        @{        Tahoma���         �:� �:� $      ����   �           ��        Arial ���         �:� �:� $      ����   �          ��        Arial ���         �:� �:� $      ����   �         ��        Arial ���         �:� �:� $    ��� ���  �         ��        Arial ���         �:� �:� $    ��� ���  �         ��        Arial ���         �:� �:� $    ��� ���  �         ��        Arial ���         �:� �:� $    ��� ���  �         BorderStylesData
      	      �           �StringsData
�                  CustNo      Country      Address      Phone              [GlobalDM.Customers.CustNo]      [GlobalDM.Customers.Country]   6   [GlobalDM.Customers.Addr1]
[GlobalDM.Customers.Addr2]      [GlobalDM.Customers.Phone]                      All Customers Count:    7   =COUNT(A[vgrDataBand1.GenBegin]:A[vgrDataBand1.GenEnd])   4   Highlighting rows in Simple customer list by country                    Company      [GlobalDM.Customers.Company]FormulasData
b  
             &                                   '                                  �                                                                                                          �F          M Y�                  M Y�                                         ����������������                                TvgrReportTemplateWorksheetvgrReportTemplateWorksheet1TitleSheet1PageProperties.Height�APageProperties.Width|.PageProperties.Margins.Left�PageProperties.Margins.TopPageProperties.Margins.Right PageProperties.MeasurementSystemvgrmsMetricColsData
4            ^         ~�   =       �    RowsData
4             �  ri   J  $"   �  *�   �    HorzSectionsData
    
RangesData
   $                              Cus����                       Cus����                        ����                        ����                        ����                        ����                          ����            
         y] ����                     y] ����                        ����               	         ����                        ����                        ����                          ���� TvgrBandvgrBand1StartPos EndPos RepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSectionVertical  TvgrDetailBandvgrDetailBand1StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSection#Events.BeforeGenerateScriptProcName$vgrDetailBand1_Events_BeforeGenerateDataSetGlobalDM.CustomersVertical  TvgrBandvgrBand2StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSectionVertical  TvgrDataBandvgrDataBand1StartPosEndPosRepeatOnPageTopRepeatOnPageBottomPrintWithNextSectionPrintWithPreviosSection"Events.AfterGenerateScriptProcName!vgrDataBand1_Events_AfterGenerateVertical    