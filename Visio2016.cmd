@echo off
title Activate Microsoft Visio 2016 (ALL versions) for FREE - https://github.com/BsNgChiThanh
cls

echo Activating your Visio 2016...


(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance*.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\visioprovl_kms*.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\visiopro2019vl_kms*.xrm-ms"


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act


:end
:notsupported
:halt
pause >nul
