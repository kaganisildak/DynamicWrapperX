@echo off
echo =======================================
echo DynamicWrapperX Auto-Installer
echo =======================================
echo.

echo Installing DynamicWrapperX for your system...
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This installer must be run as Administrator.
    echo Please right-click and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

REM Detect system architecture
set ARCH=unknown
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" set ARCH=x64
if "%PROCESSOR_ARCHITECTURE%"=="x86" set ARCH=x86
if "%PROCESSOR_ARCHITEW6432%"=="AMD64" set ARCH=x64

echo Detected system architecture: %ARCH%
echo.

REM Install appropriate version
if "%ARCH%"=="x64" (
    echo Installing x64 version for 64-bit system...
    if exist "x64\dynwrapx.dll" (
        cd x64
        regsvr32 dynwrapx.dll
        if %errorlevel% equ 0 (
            echo x64 version installed successfully!
        ) else (
            echo Failed to install x64 version.
        )
        cd ..
    ) else (
        echo ERROR: x64\dynwrapx.dll not found!
    )
    
    echo.
    echo Also installing x86 version for compatibility...
    if exist "x86\dynwrapx.dll" (
        cd x86
        %WINDIR%\SysWOW64\regsvr32.exe dynwrapx.dll
        if %errorlevel% equ 0 (
            echo x86 version installed successfully!
        ) else (
            echo Failed to install x86 version.
        )
        cd ..
    ) else (
        echo ERROR: x86\dynwrapx.dll not found!
    )
    
) else if "%ARCH%"=="x86" (
    echo Installing x86 version for 32-bit system...
    if exist "x86\dynwrapx.dll" (
        cd x86
        regsvr32 dynwrapx.dll
        if %errorlevel% equ 0 (
            echo x86 version installed successfully!
        ) else (
            echo Failed to install x86 version.
        )
        cd ..
    ) else (
        echo ERROR: x86\dynwrapx.dll not found!
    )
    
) else (
    echo ERROR: Could not detect system architecture.
    echo Please manually run the appropriate installer:
    echo   For 64-bit: x64\install.bat
    echo   For 32-bit: x86\install.bat
)

echo.
echo Installation completed!
echo.
echo You can now test DynamicWrapperX by running:
echo   cd tests
echo   run_tests.bat
echo.
pause
