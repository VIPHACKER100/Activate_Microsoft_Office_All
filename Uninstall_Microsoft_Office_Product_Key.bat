@echo off
title Uninstall Microsoft Office Product Key by viphacker100
cls
echo ============================================================================
echo #Project: Uninstall Microsoft Office Product Key
echo ============================================================================
echo.

:: Set the path to the MS Office installation directory
set OFFICE_PATH="C:\Program Files\Microsoft Office\Office16"

:: Check if the path exists
if exist %OFFICE_PATH% (
    echo Navigating to the Office installation directory...
    cd %OFFICE_PATH%
) else (
    echo Office installation directory not found!
    exit /b 1
)

:: Check the status of the Office installation
echo Checking the status of the Office installation...
cscript ospp.vbs /dstatus

:: Uninstall the product key
echo Uninstalling the product key...
set /p LAST5=Enter the last 5 characters of the product key:
cscript ospp.vbs /unpkey:%LAST5%

echo ============================================================================
echo Uninstallation complete.
echo ============================================================================

pause >nul