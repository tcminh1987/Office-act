@echo off
title  Activate Office 365 Mondo for FREE - https://github.com/BsNgChiThanh 
cls
echo Activating your Office Office 365 Mondo...
 
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ppd.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ul-oob.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ul.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-pl.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ppd.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ul-oob.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ul-phn.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)




echo.




cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul
cscript //nologo ospp.vbs /inpkey:HFTND-W9MK4-8B7MJ-B6C4G-XQBR2 >nul



set i=1
:server
if %i%==1 set KMS=kms7.MSGuides.com
if %i%==2 set KMS=kms8.MSGuides.com
if %i%==3 set KMS=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS% >nul

cscript //nologo ospp.vbs /act | find /i "successful" 
echo OK, Done!
echo Dang kiem tra trang thai kich hoat...
 


:office2016
IF EXIST %systemroot%\SysWOW64\cmd.exe (SET bit=64&SET wow=1) ELSE (SET bit=32&SET wow=0)
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2016 %bit%-bit          ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2013
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2016 32-bit             ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2013
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2013 %bit%-bit          ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2010
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2013 32-bit             ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2010
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2010 %bit%-bit             ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)
IF %wow%==0 GOTO :office2016C2R
SET office=
FOR /F "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>NUL') DO (SET "office=%%b")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2010 32-bit             ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2016C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :office2013C2R
SET office=
for /f "tokens=2*" %%a IN ('"REG QUERY HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>NUL') do (set "office=%%b\Office16")
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***            Trang thai Office 2016/2019 C2R           ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2013C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :office2010C2R
SET office=
IF EXIST "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office15"
) else IF EXIST "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office15"
)
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Trang thai Office 2013 C2R              ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:office2010C2R
REG QUERY HKLM\SOFTWARE\Microsoft\Office\14.0\ClickToRun /v InstallPath >NUL 2>&1 || GOTO :End
SET office=
IF EXIST "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles%\Microsoft Office\Office14"
) else IF EXIST "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office14"
)
IF EXIST "%office%\OSPP.VBS" (
ECHO.
ECHO ************************************************************
ECHO ***              Trang thai Office 2010 C2R              ***
ECHO ************************************************************
cscript //nologo "%office%\OSPP.VBS" /dstatus
)

:end
:notsupported
:halt
pause >nul
