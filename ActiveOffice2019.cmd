@echo off
title Kich hoat Microsoft Office 2019 ALL versions - https://github.com/BsNgChiThanh
cls

echo Activating your Office...
 
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
echo.
 


cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul
cscript //nologo ospp.vbs /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act


set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=s8.uk.to
if %i%==3 set KMS_Sev=s9.uk.to
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul


:end
:notsupported
:halt
pause >nul
