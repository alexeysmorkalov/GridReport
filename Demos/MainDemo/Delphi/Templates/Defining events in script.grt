TPF0TvgrReportTemplateReportTemplateScript.LanguageVBScriptScript.Script.StringsPublic CurRow ]sub ReportTemplatevgrReportTemplateWorksheet1_Events_BeforeGenerate (ByVal TemplateWorksheet)  CurRow = 1$  'Defining new AfterGenerate event:K  TemplateWorksheet.Events.AfterGenerateScriptProcName = "AfterGenerateSub"(  'Defining new AfterAddWorksheet event:V  TemplateWorksheet.Events.AddWorkbookWorksheetScriptProcName = "AfterAddWorksheetSub"end sub usub ReportTemplatevgrReportTemplateWorksheet1_Events_AfterGenerate (ByVal TemplateWorksheet, ByVal WorkbookWorksheet)  with WorkbookWorksheetW    .Ranges(0, CurRow, 5, CurRow).Value = "Old event AfterGenerate of worksheet occurs"    CurRow = CurRow + 1  
  end withend sub xsub ReportTemplatevgrReportTemplateWorksheet1_Events_AfterAddWorksheet(ByVal TemplateWorksheet, ByVal WorkbookWorksheet)  with WorkbookWorksheet[    .Ranges(0, CurRow, 5, CurRow).Value = "Old event AfterAddWorksheet of worksheet occurs"    CurRow = CurRow + 1  
  end withend sub "'Procedure for AfterGenerate eventFsub AfterGenerateSub(ByVal TemplateWorksheet, ByVal WorkbookWorksheet)  with WorkbookWorksheeta    .Ranges(0, CurRow, 5, CurRow).Value = "Newly defined event AfterGenerate of worksheet occurs"    CurRow = CurRow + 1  
  end withend sub &'Procedure for AfterAddWorksheet eventJsub AfterAddWorksheetSub(ByVal TemplateWorksheet, ByVal WorkbookWorksheet)  with WorkbookWorksheete    .Ranges(0, CurRow, 5, CurRow).Value = "Newly defined event AfterAddWorksheet of worksheet occurs"    CurRow = CurRow + 1  
  end withend sub Script.AliasManagerMainForm.AliasManagerLeft� TopxDataStorageVersion
SystemInfo$OS: WIN32_NT 5.1.2600 Service Pack 1 PageSize: 4096ActiveProcessorMask: $1000NumberOfProcessors: 1ProcessorType: 586 Compiler version: Delphi7DataStorage version: 1 RangeStylesData
8     >      ��X   
     Arial ���         �:� �:� $      �  �  �          �|       Tahoma���         �:� �:� $      � ��   �         ��        Tahoma���         �:� �:� $      �  �  �        ��        Arial ���         �:� �:� $      �  �  �         �}        Tahoma���         �:� �:� $      � ��   �          ��   
     Tahoma���         �:� �:� $      �  �  �          �|       Tahoma���         �:� �:� $      � ��   �         x       Tahoma���         �:� �:� $      ����   �        x       Tahoma���         �:� �:� $      ����   �        x       Tahoma���         �:� �:� $      ����   �        x       Tahoma���         �:� �:� $      ����   �        x       Tahoma���         �:� �:� $      ����   �        x       Tahoma���         �:� �:� $      ����   �        �}        Tahoma���         �:� �:� $      � ��   �          �|       Tahoma���         �:� �:� $      � ��   �         �}        Tahoma���         �:� �:� $      � ��   �         �}        Tahoma���         �:� �:� $      � ��   �          �}        Tahoma���         �:� �:� $      � ��   �         �Y        Arial ���         �:� �:� $    ��� ���   �        �Y        Arial ���         �:� �:� $    ��� ���   �        �Y        Arial ���         �:� �:� $    ��� ���   �        �Y        Arial ���         �:� �:� $    ��� ���   �        �Y        Arial ���         �:� �:� $    ��� ���   �        �Y        Arial ���         �:� �:� $    ��� ���   �        �_       Arial ���         �:� �:� $    ��� ���   �        �+    
    Arial ���         �:� �:� $      � ��   �         �.    
     Arial ���         �:� �:� $      � ��   �          BorderStylesData
      	      �           �StringsData
�                  OrderNo   	   Sale date   	   Ship date      Terms      Payment
method      Total   	   Customer:      Phone:      Fax:   0   [GlobalDM.Customer.Company] ([Customer.Country])      [Customer.Phone]      [Customer.Fax]      Simple Master - detail report      (orders linked to customers)      [Orders.OrderNo]      [Orders.SaleDate]      [Orders.ShipDate]      [Orders.Terms]      [Orders.PaymentMethod]      [Orders.ItemsTotal]                   �     Example of Defining events in script
  This example demonstrates 
  how to define events for GridReport's objects.
  Please generate report and see result workbook.                                        FormulasData
b  
             &                                   '                                  �                                                                                                          �F          M Y�                  M Y�                                         ����������������                                TvgrReportTemplateWorksheet)ReportTemplatevgrReportTemplateWorksheet1TitleSheet1PageProperties.Height�APageProperties.Width|. PageProperties.MeasurementSystemvgrmsMetric#Events.BeforeGenerateScriptProcName?ReportTemplatevgrReportTemplateWorksheet1_Events_BeforeGenerate"Events.AfterGenerateScriptProcName>ReportTemplatevgrReportTemplateWorksheet1_Events_AfterGenerateColsData
�                    �       Q	       
       	       �       �       �       �    	   �    
   �       �       w    RowsData
             '    
RangesData
,   $                                 ����   