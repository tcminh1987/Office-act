@echo off
title Bs Nguyen Chi Thanh, Kich hoat Microsoft Office 2019 ALL versions!&cls&echo ==========================Bs Nguyen Chi Thanh======================================&echo # Bs Nguyen Chi Thanh, Khoa CC_HSTC_CD BV Dam Doi Kich hoat Microsoft Office 2019&echo ==========================Bs Nguyen Chi Thanh======================================&echo.&echo #San pham ho tro:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ==========================Bs Nguyen Chi Thanh======================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ==========================Bs Nguyen Chi Thanh======================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ==========================Bs Nguyen Chi Thanh======================================&choice /n /c YN /m "Truy cap trang Web Bs Nguyen Chi Thanh: [Y,N]?" & if errorlevel 2 exit) || (echo Ket noi voi may chu KMS khong thanh cong! Dang ket noi lai... & echo Vui long cho... & echo. & echo. & set /a i+=1 & goto server)
explorer "https://phong-kham-bsck1-nguyen-chi-thanh.business.site/?m=true"&goto halt
:notsupported
echo.&echo ==========================Bs Nguyen Chi Thanh======================================&echo Phien ban Office cua ban khong duoc ho tro.&echo Download phien ban moi nhat tai day: http://topthuthuat.vn/:halt
pause >nul