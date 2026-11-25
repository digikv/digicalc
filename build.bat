@echo off
rem DigiCalc build (Windows):  build [hbmk2 options, e.g. -rebuild]
rem HBMK2  hbmk2 to use (default: Harbour 3.2.1 in D:\tools\harbour-core)
rem Needs ..\Harbour-CUI (the hbtui library built) and, for .xlsx, ..\xlsx.
if not defined HBMK2 set HBMK2=D:\tools\harbour-core\bin\win\mingw\hbmk2.exe
"%HBMK2%" digicalc.hbp %1 %2 %3 %4
