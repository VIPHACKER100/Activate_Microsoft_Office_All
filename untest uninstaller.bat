@echo off
title Check and Uninstall Microsoft Office Product Key
title created by viphacker.100
cls
echo ============================================================================
echo #Project: Check and Uninstall Microsoft Office Product Key
echo ============================================================================
echo.

:: Check if ospp.vbs exists
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    set OSPP_PATH=%ProgramFiles%\Microsoft Office\Office16\ospp.vbs
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    set OSPP_PATH=%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs
) else if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" (
    set OSPP_PATH=%ProgramFiles%\Microsoft Office\Office15\ospp.vbs
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" (
    set OSPP_PATH=%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs
) else (
    echo ospp.vbs not found!
    exit /b 1
)

:: Check the status of the Office installation
echo Checking the status of the Office installation...
cscript //nologo %OSPP_PATH% /dstatus | find /i "Product key"

:: Uninstall the product key
echo Uninstalling the product key...
cscript //nologo %OSPP_PATH% /unpkey:YOUR_PRODUCT_KEY

echo ============================================================================
echo Uninstallation complete.
echo ============================================================================

pause >nul