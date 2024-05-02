@echo off
title Activate Microsoft Visio 2019 (ALL versions) for FREE - https://github.com/BsNgChiThanh 
cls

echo Activating Microsoft Visio 2019...
 

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance*.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\visioprovl_kms*.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\visiopro2019vl_kms*.xrm-ms"


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
cscript ospp.vbs /inpkey:7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act


:end
:notsupported
:halt
pause >nul

