rem *********************************************
rem Usage:
rem   CopyBuilderDCU.bat <��⠫���㤠�����஢���> <�����Builder���⮖��ன>
rem *********************************************

set dest=%1
set dver=%2

copy %prcat%\Source\*.dcu %dest%
copy %prcat%\Source\*.dfm %dest%
copy %prcat%\Source\*.hpp %dest%
copy %prcat%\Source\*.obj %dest%
copy %prcat%\Source\*.inc %dest%
copy %prcat%\Source\vgr_*GridReportBCB%dver%.bpk %dest%
copy %prcat%\Source\vgr_*GridReportBCB%dver%.cpp %dest%
copy %prcat%\Source\vgr_StringIDs.pas %dest%

copy %prcat%\Source\*.res %dest%
del %dest%\vgr_*GridReport*.res
copy %prcat%\Source\vgr_*GridReportBCB%dver%.res %dest%
