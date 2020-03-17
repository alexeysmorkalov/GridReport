rem *********************************************
rem Usage:
rem   CopyRes.bat
rem *********************************************

call %prcat%\RES\ENG\compile.bat
copy %prcat%\RES\*.res %dppr%\GridReport\RES

call batch\CopyResDir.bat %prcat%\RES\ENG %dppr%\GridReport\RES\ENG
call batch\CopyResDir.bat %prcat%\RES\RUS %dppr%\GridReport\RES\RUS
