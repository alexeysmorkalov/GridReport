rem *********************************************
rem Usage:
rem   CopyDelphiDCU.bat <��⠫���㤠�����஢���> <�����Delphi���⮖��ன>
rem *********************************************

set dest=%1
set dver=%2

copy %prcat%\Source\*.dcu %dest%
copy %prcat%\Source\*.dfm %dest%
copy %prcat%\Source\*.inc %dest%
copy %prcat%\Source\vgr_*GridReportD%dver%.dpk %dest%
copy %prcat%\Source\vgr_StringIDs.pas %dest%

copy %prcat%\Source\*.res %dest%
del %dest%\vgr_*GridReport*.res
copy %prcat%\Source\vgr_*GridReportD%dver%.res %dest%
