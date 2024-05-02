@echo off
title Activate Microsoft Office Prolus 2013 (ALL versions) for FREE - https://github.com/BsNgChiThanh 


echo Activating Microsoft Office Prolus 2013...


(if exist “%ProgramFiles%\Microsoft Office\Office15\ospp.vbs” cd /d “%ProgramFiles%\Microsoft Office\Office15”)
(if exist “%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs” cd /d “%ProgramFiles(x86)%\Microsoft Office\Office15”)


cscript ospp.vbs /setprt:1688
cscript OSPP.VBS /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript OSPP.VBS /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act



:end
:notsupported
:halt
pause >nul