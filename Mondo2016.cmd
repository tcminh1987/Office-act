@echo off
color F0
mode con cols=98 lines=30
title Activate Office 365 Mondo, 2021, 2019, 2016 Free - https://github.com/BsNgChiThanh 
setlocal EnableExtensions EnableDelayedExpansion

:======================================================================================================================================================
:MAINMENU
cls
mode con cols=98 lines=30

echo.                                        
echo.                                       BSCK1. NGUYEN CHI THANH
echo.                     _________________________________________________________
echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Kich hoat Office 365 Mondo.                       ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Kich hoat Office 2021.                            ^|  
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Kich hoat Office 2019                             ^|
echo.                    ^|                                                         ^|
echo.                    ^|   [4] Kich hoat Office 2016                             ^| 
echo.                    ^|                                                         ^| 
echo.                    ^|   [5] Xem trang thai kich hoat Office                   ^| 
echo.                    ^|                                                         ^| 
Echo.                    ^|   [6] Thoat.                                            ^|
Echo.                    ^|_________________________________________________________^|
ECHO.            
choice /C:123456 /N /M ".                      Nhap lua chon cua ban [1,2,3,4,5,6] : "
if errorlevel 6 goto:Exit
if errorlevel 5 goto:TrangthaiOffice
if errorlevel 4 goto:Office2016
if errorlevel 3 goto:Office2019
if errorlevel 2 goto:Office2021
if errorlevel 1 goto:365Mondo

:======================================================================================================================================================
:TrangthaiOffice
cls

color F0
mode con cols=200 lines=200
echo Xem trang thai Office cua ban...

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

cscript ospp.vbs /dstatus

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:======================================================================================================================================================

:365Mondo
cls

color F0
mode con cols=98 lines=30
echo Activating your Microsft Office 365 Mondo...
 
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ppd.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ul-oob.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_KMS_Client-ul.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-pl.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ppd.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ul-oob.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\MondoVL_MAK-ul-phn.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:DMTCJ-KNRKX-26982-JYCKT-P7KB6
cscript ospp.vbs /inpkey:HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU



 ===============================================================================================
:Office2021
 ===========================================================
cls
mode con cols=98 lines=30


echo.                                       BSCK1. NGUYEN CHI THANH
echo.                      _________________________________________________________
echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Office 2021 Prolus.                               ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Project 2021.                                     ^|  
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Visio 2021                                        ^|
echo.                    ^|                                                         ^|
echo.                    ^|   [4] Ve Menu chinh                                     ^| 
echo.                    ^|                                                         ^| 
Echo.                    ^|_________________________________________________________^|
ECHO.            
choice /C:1234 /N /M ".                       Nhap lua chon cua ban [1,2,3,4] : "

if errorlevel 4 goto:MAINMENU
if errorlevel 3 goto:Visio2021
if errorlevel 2 goto:Project2021
if errorlevel 1 goto:Office2021Prolus

 ===========================================================
:Office2021Prolus

cls
color F0
mode con cols=98 lines=30

echo Activating your Microsoft Office 2021...

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


script ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /inpkey:KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act


echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 

:Project 2021
cls
color F0
mode con cols=98 lines=30

echo Activating Microsoft Project 2021...


cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
(for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\projectpro2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)

cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
cscript ospp.vbs /inpkey:J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act 

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:===========================================================  

:Visio2021
cls
color F0
mode con cols=98 lines=30

echo Activating Microsoft Visio 2021...
 

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


(cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled)&cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
(for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\visioprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\visiopro2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
cscript ospp.vbs /inpkey:MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU

 ===============================================================================================

:=========================================================== 
:Office2019
 ===========================================================
cls
mode con cols=98 lines=30


echo.                                       BSCK1. NGUYEN CHI THANH
echo.                     _________________________________________________________
echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Office 2019 Prolus.                               ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Project 2019.                                     ^|  
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Visio 2019                                        ^|
echo.                    ^|                                                         ^|
echo.                    ^|   [4] Ve Menu chinh                                     ^| 
echo.                    ^|                                                         ^| 
Echo.                    ^|_________________________________________________________^|
ECHO.            
choice /C:1234 /N /M ".                       Nhap lua chon cua ban [1,2,3,4] : "

if errorlevel 4 goto:MAINMENU
if errorlevel 3 goto:Visio2019
if errorlevel 2 goto:Project2019
if errorlevel 1 goto:Office2019Prolus
 =========================================================== 

:Office2019Prolus
cls
color F0
mode con cols=98 lines=30

echo Activating your Microsoft Office 2019...
 
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
echo.

cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 

:Project2019
cls
color F0
mode con cols=98 lines=30

echo Activating Microsoft Project Professional Plus 2019...


(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
(for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
cscript ospp.vbs /inpkey:C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act


echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 

:Visio2019
cls
color F0
mode con cols=98 lines=30

echo Activating Microsoft Visio 2019...
 

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


(cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled)&cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
(for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\visioprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\visiopro2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
cscript ospp.vbs /inpkey:7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 

:Office2016
 ===========================================================
cls
mode con cols=98 lines=30


echo.                                       BSCK1. NGUYEN CHI THANH
echo.                     _________________________________________________________
echo.                    ^|                                                         ^|
Echo.                    ^|   [1] Office 2016 Prolus.                               ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] Project 2016.                                     ^|  
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Visio 2016                                        ^|
echo.                    ^|                                                         ^|
echo.                    ^|   [4] Ve Menu chinh                                     ^| 
echo.                    ^|                                                         ^| 
Echo.                    ^|_________________________________________________________^|
ECHO.            
choice /C:1234 /N /M ".                       Nhap lua chon cua ban [1,2,3,4] : "

if errorlevel 4 goto:MAINMENU
if errorlevel 3 goto:Visio2016
if errorlevel 2 goto:Project2016
if errorlevel 1 goto:Office2016Prolus
 =========================================================== 
:Office2016Prolus
cls
color F0
mode con cols=98 lines=30

echo Activating your Microsoft Office prolus 2016...
 
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)


cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 
:Project2016
cls
color F0
mode con cols=98 lines=30

echo Activating your Project 2016...

(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")


cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-bridge-office.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root-bridge-test.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul.xrm-ms" >nul
cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms" >nul



cscript ospp.vbs /setprt:1688
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript ospp.vbs /inpkey:GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 

:Visio2016
cls
color F0
mode con cols=98 lines=30

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
cscript ospp.vbs /sethst:172.16.0.2
cscript ospp.vbs /act

echo.
echo.
echo Bam phim bat ki de ve menu chinh

:notsupported
:halt
pause >nul

goto:MAINMENU
:=========================================================== 
 ===============================================================================================

:======================================================================================================================================================
:Exit
cls
echo.
echo.
echo.
echo.
echo Kich hoat moi windows va Office vinh vien bang MAS Tools 
start https://github.com/BsNgChiThanh/MAS-TOOL
pause > nul
exit
:======================================================================================================================================================
