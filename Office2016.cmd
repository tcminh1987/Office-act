@echo off
title Activate Microsoft Office 2016 ALL versions for FREE - https://github.com/BsNgChiThanh
cls
echo =====================================================================================
echo #Project: Activating Microsoft software products for FREE without additional software
echo =====================================================================================
echo.
echo #Supported products:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.
echo.

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)

echo.
echo ============================================================================
echo Activating your Office...


cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul
cscript //nologo ospp.vbs /unpkey:CPQVG >nul


set i=1
cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul||goto notsupported

:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=107.175.77.7
if %i% GTR 2 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul


:ato
cscript //nologo ospp.vbs /act | find /i "successful" 
:busy
:halt
pause >nul
