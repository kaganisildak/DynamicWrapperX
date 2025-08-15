@echo off
REM Build script for DynamicWrapperX - both x64 and x86 architectures
REM Execute this in Visual Studio Developer Command Prompt

setlocal enabledelayedexpansion

REM Parse command line arguments
set DEBUG_BUILD=false
set RELEASE_BUILD=false

:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="-debug" (
    set DEBUG_BUILD=true
    echo Debug build enabled - will include debug logging
    shift
    goto parse_args
)
if /i "%~1"=="-release" (
    set RELEASE_BUILD=true
    echo Release build enabled - optimized for size, no debug logging
    shift
    goto parse_args
)
echo Unknown option: %~1
echo Usage: %0 [-debug^|-release]
echo   -debug   : Enable debug logging ^(larger DLL^)
echo   -release : Optimize for size, disable debug logging ^(smaller DLL^)
exit /b 1

:args_done
REM Default to debug if neither specified
if "%DEBUG_BUILD%"=="false" if "%RELEASE_BUILD%"=="false" (
    set DEBUG_BUILD=true
    echo No build type specified, defaulting to debug build
)

if "%DEBUG_BUILD%"=="true" (
    echo Building DynamicWrapperX for Windows ^(DEBUG build^)...
) else (
    echo Building DynamicWrapperX for Windows ^(RELEASE build^)...
)

REM Check if we're in a Visual Studio Developer Command Prompt
where cl >nul 2>&1
if errorlevel 1 (
    echo Error: Visual Studio C++ compiler ^(cl.exe^) not found.
    echo Please run this script from a Visual Studio Developer Command Prompt.
    echo You can find it in Start Menu under Visual Studio tools.
    exit /b 1
)


REM Set compiler flags based on build type
if "%DEBUG_BUILD%"=="true" (
    set CFLAGS=/Wall /Od /Zi /MDd /std:c11 /D_WIN32_WINNT=0x0601 /DDEBUG_BUILD
    set LDFLAGS=/DEBUG /DLL
    echo Using debug flags: !CFLAGS!
) else (
    set CFLAGS=/Wall /O1 /MD /std:c11 /D_WIN32_WINNT=0x0601 /DNDEBUG /GL
    set LDFLAGS=/LTCG /DLL
    echo Using release flags: !CFLAGS!
)

set LIBS=ole32.lib oleaut32.lib uuid.lib advapi32.lib shlwapi.lib kernel32.lib user32.lib

REM Create build directories
if not exist "build" mkdir build
if not exist "build\x64" mkdir build\x64
if not exist "build\x86" mkdir build\x86

REM Create module definition file if it doesn't exist
if not exist "build\dynwrapx.def" (
    echo Creating module definition file...
    echo EXPORTS > build\dynwrapx.def
    echo DllCanUnloadNow >> build\dynwrapx.def
    echo DllGetClassObject >> build\dynwrapx.def
    echo DllRegisterServer >> build\dynwrapx.def
    echo DllUnregisterServer >> build\dynwrapx.def
    echo DllInstall >> build\dynwrapx.def
)

REM Build x64 version
echo.
echo Building x64 version...
pushd build\x64



REM Compile source files for x64
echo Compiling x64 source files...
cl !CFLAGS! /c ..\..\main.c /Fo:main.obj
if errorlevel 1 (
    echo Failed to compile main.c for x64
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\dynwrapx.c /Fo:dynwrapx.obj
if errorlevel 1 (
    echo Failed to compile dynwrapx.c for x64
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\methods.c /Fo:methods.obj
if errorlevel 1 (
    echo Failed to compile methods.c for x64
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\factory.c /Fo:factory.obj
if errorlevel 1 (
    echo Failed to compile factory.c for x64
    popd
    exit /b 1
)

REM Link the x64 DLL
echo Linking x64 DLL...
link !LDFLAGS! /OUT:dynwrapx.dll /DEF:..\dynwrapx.def main.obj dynwrapx.obj methods.obj factory.obj !LIBS!
if errorlevel 1 (
    echo Failed to link x64 DLL
    popd
    exit /b 1
)

echo x64 build completed: build\x64\dynwrapx.dll

REM Clean up intermediate files
echo Cleaning up x64 intermediate files...
del *.obj *.res *.exp *.lib >nul 2>&1
if "%DEBUG_BUILD%"=="false" del *.pdb >nul 2>&1

popd

REM Build x86 version
echo.
echo Building x86 version...
pushd build\x86


REM Compile source files for x86
echo Compiling x86 source files...
cl !CFLAGS! /c ..\..\main.c /Fo:main.obj
if errorlevel 1 (
    echo Failed to compile main.c for x86
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\dynwrapx.c /Fo:dynwrapx.obj
if errorlevel 1 (
    echo Failed to compile dynwrapx.c for x86
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\methods.c /Fo:methods.obj
if errorlevel 1 (
    echo Failed to compile methods.c for x86
    popd
    exit /b 1
)

cl !CFLAGS! /c ..\..\factory.c /Fo:factory.obj
if errorlevel 1 (
    echo Failed to compile factory.c for x86
    popd
    exit /b 1
)

REM Link the x86 DLL
echo Linking x86 DLL...
link !LDFLAGS! /OUT:dynwrapx.dll /DEF:..\dynwrapx.def main.obj dynwrapx.obj methods.obj factory.obj !LIBS!
if errorlevel 1 (
    echo Failed to link x86 DLL
    popd
    exit /b 1
)

echo x86 build completed: build\x86\dynwrapx.dll

REM Clean up intermediate files
echo Cleaning up x86 intermediate files...
del *.obj *.res *.exp *.lib >nul 2>&1
if "%DEBUG_BUILD%"=="false" del *.pdb >nul 2>&1

popd

echo.
echo Build completed successfully!
echo.
echo Files created:
echo   build\x64\dynwrapx.dll ^(64-bit^)
echo   build\x86\dynwrapx.dll ^(32-bit^)
echo.
echo To register the DLLs on Windows:
echo   For x64: regsvr32 build\x64\dynwrapx.dll
echo   For x86: regsvr32 build\x86\dynwrapx.dll
echo.
echo For current user only, add /i flag:
echo   regsvr32 /i build\x64\dynwrapx.dll
echo.

endlocal
