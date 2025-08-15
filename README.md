# DynamicWrapperX

A complete reimplementation of the DynamicWrapperX COM component for calling Windows APIs from scripting languages. This project provides a drop-in replacement for the original DynamicWrapperX with enhanced features, modern architecture, and comprehensive debugging capabilities.

## Overview

DynamicWrapperX is a COM (Component Object Model) wrapper that enables scripting languages like JScript, VBScript, and Windows Script Host to call native Windows API functions dynamically. This reimplementation is inspired by the original DynamicWrapperX from [script-coding.com](https://www.script-coding.com/dynwrapx_eng.html) 

## Features

### Core Functionality
- **Dynamic API Loading**: Load any Windows DLL and call its exported functions
- **Multiple Parameter Types**: Support for integers, pointers, strings (ANSI/Unicode), and structures
- **64-bit & 32-bit Architecture**: Full support for both x64 and x86 platforms with proper pointer handling
- **COM Interface**: Standard COM implementation
- **Memory Management**: Automatic memory allocation and cleanup for parameters and return values

### Architecture-Specific Optimizations
- **x64 Support**: Proper 64-bit pointer handling with VT_I8 VARIANT types
- **Calling Conventions**: Support for different calling conventions across architectures

## Build Options

The project supports two build modes through command-line flags:

### Debug Build (Default)
```bash
./build-all.sh -debug
```
- Includes debug logging
- Writes detailed logs to `debugdyn.log`
- Larger DLL size with full debugging capabilities
- Optimized for development and troubleshooting

### Release Build
```bash
./build-all.sh -release
```

## Installation & Usage

### Building from Source

#### Prerequisites
- MinGW-w64 cross-compiler
- On Ubuntu/Debian: `sudo apt-get install mingw-w64`
- On macOS: `brew install mingw-w64`

#### Compilation
```bash
# Debug build (includes logging)
./build-all.sh -debug

# Release build (optimized)
./build-all.sh -release
```

### DLL Registration

After building, register the DLL on Windows:

```cmd
# For x64 systems
regsvr32 build\x64\dynwrapx.dll

# For x86 systems
regsvr32 build\x86\dynwrapx.dll

# For current user only
regsvr32 /i build\x64\dynwrapx.dll
```

### Basic Usage Example

```javascript
// Create DynamicWrapperX instance
var DX = new ActiveXObject("DynamicWrapperX");

// Register a Windows API function
DX.Register("kernel32.dll", "GetCurrentProcessId", "r=", "", "r=l");

// Call the function
var processId = DX.GetCurrentProcessId();
WScript.Echo("Current Process ID: " + processId);

// Process enumeration example
DX.Register("kernel32.dll", "CreateToolhelp32Snapshot", "l=ll", "", "r=l");
DX.Register("kernel32.dll", "Process32FirstW", "l=lp", "", "r=l");

var snapshot = DX.CreateToolhelp32Snapshot(2, 0); // TH32CS_SNAPPROCESS
if (snapshot != -1) {
    var processEntry = DX.Space(568); // PROCESSENTRY32W structure
    DX.NumPut(568, processEntry, 0, "u"); // Structure size
    
    if (DX.Process32FirstW(snapshot, processEntry)) {
        var processName = DX.StrGet(processEntry, 44, "w");
        WScript.Echo("First Process: " + processName);
    }
}
```

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.


## Security Considerations

This tool provides direct access to Windows API functions and memory manipulation capabilities. Use responsibly and only in controlled environments. The debug build includes comprehensive logging that may contain sensitive information - use release builds in production environments.

## Acknowledgments

- Original DynamicWrapperX implementation from script-coding.com
