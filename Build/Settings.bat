@echo off
rem ������ ��� ���� 䠩��� ����ਡ�⨢�
rem ��ࠬ��� - ��姭��� ����� ���ᨨ (����� ᮮ⢥����� ��� �� CVS)
set d4=GridReport%prversion%D4
set d5=GridReport%prversion%D5
set d6=GridReport%prversion%D6
set d7=GridReport%prversion%D7
set bcb5=GridReport%prversion%BCB5
set bcb6=GridReport%prversion%BCB6
set d4std=GridReport%prversion%D4Std
set d5std=GridReport%prversion%D5Std
set d6std=GridReport%prversion%D6Std
set d7std=GridReport%prversion%D7Std
set bcb5std=GridReport%prversion%BCB5Std
set bcb6std=GridReport%prversion%BCB6Std
set doc=GridReport%prversion%eng
set src=GridReport%prversion%SRC
set dcu=GridReport%prversion%

set prcat=%CD%\..

rem ��⠫��, � htmp ���� ��� GridReport
set prhtml=%prcat%\..\GridReport_documentation\Draft_Html

rem ���� � ������� �ணࠬ�� ��� �������樨 ���
set dipas=%prcat%\..\Bin\dipas_console_m.exe

rem ��⠫��, �㤠 ������� ����ਡ�⨢
set dpup=%prcat%\..\..\upload

rem �६���� ��⠫��� ��� ᡮન ����ਡ�⨢�
set dppr=%dpup%\Temp
set dplib=%dpup%\TempLib

rem ������騪 
set wz=%CD%\..\..\bin\pkzip25.exe

set prrescompile=%prcat%\RES\ENG

set commoncontrols=%prcat%\..\CommonControls\Source

set d4cat=D:\Program Files\Borland\Delphi4
set d5cat=D:\Program Files\Borland\Delphi5
set d6cat=D:\Program Files\Borland\Delphi6
set d7cat=D:\Program Files\Borland\Delphi7
set bcb5cat=D:\Program Files\Borland\CBuilder5
set bcb6cat=D:\Program Files\Borland\CBuilder6
