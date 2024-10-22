
@echo off
title Activate Microsoft Project 2021 for FREE - https://github.com/BsNgChiThanh 
cls

color F0
mode con cols=98 lines=30

echo Activating Microsoft Project 2021...


cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16



cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
cscript ospp.vbs /inpkey:J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Nhan phim bat ky de thoat.

:end
:notsupported
:halt
pause >nul
