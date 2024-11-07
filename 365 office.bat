@echo on
title Activate Microsoft Office 365!
cls
echo ============================================================================
echo #Project: Activating Microsoft software products
echo ============================================================================
echo.
echo #Supported products:
echo - Microsoft Office 365
echo - Microsoft Office Standard 2019
echo - Microsoft Office Professional Plus 2019
echo.
echo.

call :CheckOSPP
call :ActivateProduct
call :SetKMS
call :ActivateOffice

goto :EOF

:CheckOSPP
echo Checking for ospp.vbs...
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    echo ospp.vbs not found!
    exit /b 1
)

:ActivateProduct
echo Activating product...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)

:SetKMS
echo Setting KMS server...
set KMS_SERVER=kms.03k.org
cscript //nologo ospp.vbs /sethst:%KMS_SERVER% >nul

:ActivateOffice
echo Activating Office...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul

echo ============================================================================
echo.
echo.
cscript //nologo ospp.vbs /act | find /i "successful"

if %errorlevel% neq 0 (
    echo The connection to the KMS server failed! Trying to connect to another one...
    echo Please wait...
    echo.
    echo.
    set /a i+=1
    goto SetKMS
)

explorer "viphacker.100"
goto :EOF

:notsupported
echo.
echo ============================================================================
echo Sorry! Your version is not supported.
echo.
goto :EOF

:halt
pause >nul