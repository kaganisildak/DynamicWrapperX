@echo off
echo =======================================
echo DynamicWrapperX Uninstaller
echo =======================================
echo.

echo Uninstalling DynamicWrapperX...
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This uninstaller must be run as Administrator.
    echo Please right-click and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo Unregistering DynamicWrapperX from system...

REM Try to unregister both versions
if exist "x64\dynwrapx.dll" (
    echo Unregistering x64 version...
    regsvr32 /u "x64\dynwrapx.dll"
)

if exist "x86\dynwrapx.dll" (
    echo Unregistering x86 version...
    %WINDIR%\SysWOW64\regsvr32.exe /u "x86\dynwrapx.dll"
)

echo.
echo DynamicWrapperX has been uninstalled.
echo.
pause
