@echo off
title Activate Microsoft Visio 2016 (ALL versions) for FREE - https://github.com/BsNgChiThanh
cls

echo Activating your Visio 2016...


(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"  >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-bridge-office.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root-bridge-test.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul

 


cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:W8GF4 >nul
cscript //nologo ospp.vbs /unpkey:RJRJK >nul

set i=1
cscript //nologo ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul||goto notsupported

:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=e8.us.to
if %i% EQU 3 set KMS=e9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul


:ato
cscript //nologo ospp.vbs /act | find /i "successful" 


:halt
pause >nul