@echo off
rem Шаблон для имен файлов дистрибутива
rem параметр - трехзначный номер версии (который соответствует тэгу на CVS)
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

rem Каталог, с htmp хэлпом для GridReport
set prhtml=%prcat%\..\GridReport_documentation\Draft_Html

rem Путь к Диминой программе для компиляции хэлпа
set dipas=%prcat%\..\Bin\dipas_console_m.exe

rem Каталог, куда помещать дистрибутив
set dpup=%prcat%\..\..\upload

rem Временные каталоги для сборки дистрибутива
set dppr=%dpup%\Temp
set dplib=%dpup%\TempLib

rem Упаковщик 
set wz=%CD%\..\..\bin\pkzip25.exe

set prrescompile=%prcat%\RES\ENG

set commoncontrols=%prcat%\..\CommonControls\Source

set d4cat=D:\Program Files\Borland\Delphi4
set d5cat=D:\Program Files\Borland\Delphi5
set d6cat=D:\Program Files\Borland\Delphi6
set d7cat=D:\Program Files\Borland\Delphi7
set bcb5cat=D:\Program Files\Borland\CBuilder5
set bcb6cat=D:\Program Files\Borland\CBuilder6
