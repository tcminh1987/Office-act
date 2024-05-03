@echo off
title Activate Microsoft Visio 2021 for FREE - https://github.com/BsNgChiThanh 
cls

echo Activating Microsoft Visio 2021...

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance*.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\visioprovl_kms*.xrm-ms"
cscript ospp.vbs /inslic:"..\Licenses16\visiopro2021vl_kms*.xrm-ms"



cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
cscript ospp.vbs /inpkey:MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act


timeout 10
:end
:notsupported
:halt
pause >nul
