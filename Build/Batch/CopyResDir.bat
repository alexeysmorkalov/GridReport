rem *********************************************
rem Usage:
rem   CopyResDir.bat <��४��� ���筨�>, <��४��� �ਥ����>
rem *********************************************

set ressrc=%1
set resdst=%2

mkdir %resdst%

copy %ressrc%\*.rc %resdst%
copy %ressrc%\*.dpr %resdst%
copy %ressrc%\*.bat %resdst%
