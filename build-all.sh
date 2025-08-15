#!/bin/bash
# Cross-compilation script for DynamicWrapperX - both x64 and x86 architectures

# Parse command line arguments
DEBUG_BUILD=false
RELEASE_BUILD=false

while [[ $# -gt 0 ]]; do
    case $1 in
        -debug)
            DEBUG_BUILD=true
            echo "Debug build enabled - will include debug logging"
            shift
            ;;
        -release)
            RELEASE_BUILD=true
            echo "Release build enabled - optimized for size, no debug logging"
            shift
            ;;
        *)
            echo "Unknown option: $1"
            echo "Usage: $0 [-debug|-release]"
            echo "  -debug   : Enable debug logging (larger DLL)"
            echo "  -release : Optimize for size, disable debug logging (smaller DLL)"
            exit 1
            ;;
    esac
done

# Default to debug if neither specified
if [ "$DEBUG_BUILD" = false ] && [ "$RELEASE_BUILD" = false ]; then
    DEBUG_BUILD=true
    echo "No build type specified, defaulting to debug build"
fi

if [ "$DEBUG_BUILD" = true ]; then
    echo "Cross-compiling DynamicWrapperX for Windows (DEBUG build)..."
else
    echo "Cross-compiling DynamicWrapperX for Windows (RELEASE build)..."
fi

# Check if MinGW-w64 is available
if ! command -v x86_64-w64-mingw32-gcc &> /dev/null; then
    echo "Error: MinGW-w64 cross-compiler not found."
    echo "On Ubuntu/Debian: sudo apt-get install mingw-w64"
    echo "On macOS: brew install mingw-w64"
    exit 1
fi

# Check if 32-bit MinGW is available
if ! command -v i686-w64-mingw32-gcc &> /dev/null; then
    echo "Warning: 32-bit MinGW cross-compiler not found."
    echo "Only x64 build will be created."
    BUILD_X86=false
else
    BUILD_X86=true
fi

# Set compiler flags based on build type
if [ "$DEBUG_BUILD" = true ]; then
    CFLAGS="-Wall -Og -g -std=c99 -D_WIN32_WINNT=0x0601 -DDEBUG_BUILD"
    LDFLAGS="-shared -Wl,--kill-at"
    echo "Using debug flags: $CFLAGS"
else
    CFLAGS="-Wall -Os -std=c99 -D_WIN32_WINNT=0x0601 -DNDEBUG \
        -ffunction-sections -fdata-sections -fno-ident"
    LDFLAGS="-shared -Wl,--kill-at -Wl,--gc-sections -s \
         -Wl,--disable-auto-import \
         -Wl,--disable-runtime-pseudo-reloc \
         -Wl,--enable-stdcall-fixup"
    echo "Using release flags: $CFLAGS"
fi
LIBS="-lole32 -loleaut32 -luuid -ladvapi32 -lshlwapi -lkernel32 -luser32"

# Create build directories
mkdir -p build/x64
mkdir -p build/x86

# Create module definition file
cat > build/dynwrapx.def << EOF
EXPORTS
DllCanUnloadNow
DllGetClassObject
DllRegisterServer
DllUnregisterServer
DllInstall
EOF

echo "Building x64 version..."
cd build/x64

# Set x64 compiler
CC=x86_64-w64-mingw32-gcc
WINDRES=x86_64-w64-mingw32-windres

# Compile resource file
echo "Compiling x64 resources..."

# Compile object files
echo "Compiling x64 source files..."
$CC $CFLAGS -c ../../main.c -o main.o || exit 1
$CC $CFLAGS -c ../../dynwrapx.c -o dynwrapx.o || exit 1
$CC $CFLAGS -c ../../methods.c -o methods.o || exit 1
$CC $CFLAGS -c ../../factory.c -o factory.o || exit 1

# Link the x64 DLL
echo "Linking x64 DLL..."
$CC $LDFLAGS -o dynwrapx.dll main.o dynwrapx.o methods.o factory.o ../dynwrapx.def $LIBS || exit 1

echo "x64 build completed: build/x64/dynwrapx.dll"

# Clean up object files
echo "Cleaning up x64 object files..."
rm -f *.o

# Build x86 version if available
if [ "$BUILD_X86" = true ]; then
    echo ""
    echo "Building x86 version..."
    cd ../x86
    
    # Set x86 compiler
    CC=i686-w64-mingw32-gcc
    WINDRES=i686-w64-mingw32-windres
    
    # Compile resource file
    echo "Compiling x86 resources..."
    
    # Compile object files
    echo "Compiling x86 source files..."
    $CC $CFLAGS -c ../../main.c -o main.o || exit 1
    $CC $CFLAGS -c ../../dynwrapx.c -o dynwrapx.o || exit 1
    $CC $CFLAGS -c ../../methods.c -o methods.o || exit 1
    $CC $CFLAGS -c ../../factory.c -o factory.o || exit 1
    
    # Link the x86 DLL
    echo "Linking x86 DLL..."
    $CC $LDFLAGS -o dynwrapx.dll main.o dynwrapx.o methods.o factory.o ../dynwrapx.def $LIBS || exit 1
    
    echo "x86 build completed: build/x86/dynwrapx.dll"
    
    # Clean up object files
    echo "Cleaning up x86 object files..."
    rm -f *.o
    
    cd ..
else
    cd ..
fi

echo ""
echo "Cross-compilation completed successfully!"
echo ""
echo "Files created:"
echo "  build/x64/dynwrapx.dll (64-bit)"
if [ "$BUILD_X86" = true ]; then
    echo "  build/x86/dynwrapx.dll (32-bit)"
fi
echo ""
echo "To register the DLLs on Windows:"
echo "  For x64: regsvr32 build\\x64\\dynwrapx.dll"
if [ "$BUILD_X86" = true ]; then
    echo "  For x86: regsvr32 build\\x86\\dynwrapx.dll"
fi
echo ""
echo "For current user only, add /i flag:"
echo "  regsvr32 /i build\\x64\\dynwrapx.dll"

cd ..
