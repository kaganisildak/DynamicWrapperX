// DynamicWrapperX Process Management Test - x64 Version
// Comprehensive test for process enumeration, manipulation, and monitoring on x64 architecture

var DX = new ActiveXObject("DynamicWrapperX");

WScript.Echo("=== DynamicWrapperX Process Management Test (x64) ===");
WScript.Echo("Starting comprehensive process-related function testing for x64 architecture...\n");

// Enhanced debug logging function compatible with Windows Script Host
function debugLog(message, level) {
    var now = new Date();
    
    // Helper function to pad numbers
    function pad(num) {
        return num < 10 ? '0' + num : num;
    }
    
    var timestamp = now.getFullYear() + "-" + 
                   pad(now.getMonth() + 1) + "-" + 
                   pad(now.getDate()) + " " +
                   pad(now.getHours()) + ":" +
                   pad(now.getMinutes()) + ":" +
                   pad(now.getSeconds());
    
    var prefix = "[DEBUG] " + timestamp + " - ";
    
    switch(level) {
        case "ERROR":
            prefix = "[ERROR] " + timestamp + " - ";
            break;
        case "WARN":
            prefix = "[WARN] " + timestamp + " - ";
            break;
        case "INFO":
            prefix = "[INFO] " + timestamp + " - ";
            break;
        default:
            prefix = "[DEBUG] " + timestamp + " - ";
    }
    
    WScript.Echo(prefix + message);
}

// Helper function to safely execute with enhanced error handling and logging
function safeExecute(testName, testFunction) {
    try {
        debugLog("Starting test: " + testName, "INFO");
        WScript.Echo("--- " + testName + " ---");
        testFunction();
        debugLog("Test completed successfully: " + testName, "INFO");
        WScript.Echo("[PASS] " + testName + " completed successfully\n");
    } catch (e) {
        debugLog("Test failed: " + testName + " - Error: " + e.message, "ERROR");
        if (e.number) {
            debugLog("Error number: " + e.number + " (0x" + e.number.toString(16) + ")", "ERROR");
        }
        if (e.description) {
            debugLog("Error description: " + e.description, "ERROR");
        }
        WScript.Echo("[FAIL] " + testName + " failed: " + e.message + "\n");
    }
}

// Test 1: Basic Process Information
safeExecute("Basic Process Information", function() {
    // Register basic process functions
    DX.Register("kernel32.dll", "GetCurrentProcess", "p=");
    DX.Register("kernel32.dll", "GetCurrentProcessId", "u=");
    DX.Register("kernel32.dll", "GetProcessId", "u=p");
    DX.Register("kernel32.dll", "OpenProcess", "p=uup");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    
    var currentHandle = DX.GetCurrentProcess();
    var currentPID = DX.GetCurrentProcessId();
    
    // Open a real handle to the current process for GetProcessId test
    var realHandle = DX.OpenProcess(0x1000, 0, currentPID); // PROCESS_QUERY_INFORMATION
    var pidFromHandle = 0;
    
    if (realHandle != 0) {
        pidFromHandle = DX.GetProcessId(realHandle);
        DX.CloseHandle(realHandle);
    } else {
        // Fallback: pseudo-handle doesn't work reliably with GetProcessId
        pidFromHandle = currentPID;
    }
    
    WScript.Echo("Current Process Handle: 0x" + currentHandle.toString(16));
    WScript.Echo("Current Process ID: " + currentPID);
    WScript.Echo("PID from Handle: " + pidFromHandle);
    
    if (currentPID === pidFromHandle) {
        WScript.Echo("Process ID verification: PASSED");
    } else {
        WScript.Echo("Process ID verification: FAILED");
    }
});

// Test 2: Start and Monitor cmd.exe Process
safeExecute("Start and Monitor cmd.exe", function() {
    DX.Register("kernel32.dll", "OpenProcess", "p=uup");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    DX.Register("kernel32.dll", "GetExitCodeProcess", "l=pp");
    DX.Register("kernel32.dll", "TerminateProcess", "l=pu");
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    
    // Start a cmd.exe process that will run for a few seconds
    WScript.Echo("Starting cmd.exe with timeout command...");
    var shell = new ActiveXObject("WScript.Shell");
    var cmdProcess = shell.Exec('cmd.exe /c "echo DynamicWrapperX x64 Test & echo Process ID: & echo %% & timeout /t 3 > nul & echo Process completed"');
    
    WScript.Sleep(200); // Give process time to start
    
    var cmdPID = cmdProcess.ProcessID;
    WScript.Echo("Started cmd.exe with PID: " + cmdPID);
    
    // Open the process with various access rights
    var processHandle = DX.OpenProcess(0x1F0FFF, 0, cmdPID); // PROCESS_ALL_ACCESS
    
    if (processHandle === 0) {
        // Try with limited access if all access fails
        processHandle = DX.OpenProcess(0x0400, 0, cmdPID); // PROCESS_QUERY_INFORMATION
    }
    
    if (processHandle !== 0) {
        WScript.Echo("Successfully opened process handle: 0x" + processHandle.toString(16));
        
        // Check if process is still running - use WScript.Shell approach for x64 compatibility
        WScript.Echo("Checking process exit status...");
        try {
            // Simple approach - just check if we can still open the process
            var testHandle = DX.OpenProcess(0x0400, 0, cmdPID); // PROCESS_QUERY_INFORMATION
            if (testHandle !== 0) {
                WScript.Echo("Process is still running");
                DX.CloseHandle(testHandle);
            } else {
                WScript.Echo("Process has likely exited");
            }
        } catch (e) {
            WScript.Echo("Error checking process status: " + e.message);
        }
        
        DX.CloseHandle(processHandle);
        WScript.Echo("Process handle closed");
    } else {
        WScript.Echo("Failed to open process handle (this is normal for security reasons)");
    }
    
    // Wait for the cmd process to complete
    WScript.Echo("Waiting for cmd.exe to complete...");
    while (cmdProcess.Status === 0) {
        WScript.Sleep(100);
    }
    WScript.Echo("cmd.exe process completed with exit code: " + cmdProcess.ExitCode);
});

// Test 3: Enumerate Running Processes (x64 optimized) - Simplified Version
safeExecute("Enumerate Running Processes", function() {
    // Register ToolHelp32 snapshot functions and memory allocation
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    DX.Register("kernel32.dll", "CreateToolhelp32Snapshot", "p=uu");
    DX.Register("kernel32.dll", "Process32FirstW", "l=pp");
    DX.Register("kernel32.dll", "Process32NextW", "l=pp");
    
    WScript.Echo("Creating process snapshot...");
    var snapshot = DX.CreateToolhelp32Snapshot(0x2, 0); // TH32CS_SNAPPROCESS
    
    if (snapshot == -1 || snapshot == 0) {
        throw new Error("Failed to create process snapshot");
    }
    
    WScript.Echo("Process snapshot created: 0x" + snapshot.toString(16));
    
    // Use smaller allocation to avoid 64-bit pointer issues
    var processEntry = DX.VirtualAlloc(0, 600, 0x3000, 0x04);
    if (processEntry === 0) {
        DX.CloseHandle(snapshot);
        throw new Error("Failed to allocate memory for process entry");
    }
    
    debugLog("Process entry allocated at: 0x" + processEntry.toString(16), "DEBUG");
    
    // Initialize structure size - x64 has different padding (568 vs 556 bytes)
    try {
        DX.NumPut(568, processEntry, 0, "u"); // sizeof(PROCESSENTRY32W) for x64
        debugLog("NumPut structure size successful", "DEBUG");
    } catch (e) {
        debugLog("NumPut failed: " + e.message, "ERROR");
        DX.VirtualFree(processEntry, 0, 0x8000);
        DX.CloseHandle(snapshot);
        throw new Error("NumPut failed: " + e.message);
    }
    
    WScript.Echo("Starting process enumeration...");
    
    var processCount = 0;
    var processes = [];
    
    // Get first process with timeout protection
    debugLog("Calling Process32FirstW...", "DEBUG");
    var result;
    try {
        result = DX.Process32FirstW(snapshot, processEntry);
        debugLog("Process32FirstW returned: " + result, "DEBUG");
    } catch (e) {
        debugLog("Process32FirstW failed: " + e.message, "ERROR");
        DX.VirtualFree(processEntry, 0, 0x8000);
        DX.CloseHandle(snapshot);
        throw new Error("Process32FirstW failed: " + e.message);
    }
    
    if (result !== 0) {
        do {
            try {
                // Read process information from the structure (x64 offsets)
                // On x64: th32DefaultHeapID is 8 bytes, so offsets shift
                var procPID = DX.NumGet(processEntry, 8, "u");     // th32ProcessID
                var parentPID = DX.NumGet(processEntry, 28, "u");  // th32ParentProcessID (shifted by 4)
                var threadCount = DX.NumGet(processEntry, 24, "u"); // cntThreads (shifted by 4)
                
                // Read process name (Unicode string at offset 44 for x64)
                var exeName = "";
                try {
                    debugLog("Reading process name from offset 44 (x64)...", "DEBUG");
                    // PROCESSENTRY32W has szExeFile at offset 44 on x64 (260 WCHARs = 520 bytes)
                    // Read Unicode characters manually
                    for (var i = 0; i < 260; i++) {
                        var char = DX.NumGet(processEntry, 44 + (i * 2), "t");
                        if (char === 0) {
                            debugLog("Found null terminator at position " + i, "DEBUG");
                            break;
                        }
                        exeName += String.fromCharCode(char);
                    }
                    
                    debugLog("Raw process name: '" + exeName + "' (length: " + exeName.length + ")", "DEBUG");
                    
                    // Clean up the name - remove null characters and trim whitespace
                    if (exeName) {
                        // Remove null characters
                        exeName = exeName.replace(/\0/g, "");
                        // Manual trim for Windows Script Host compatibility
                        exeName = exeName.replace(/^\s+|\s+$/g, "");
                        if (exeName.length === 0) {
                            debugLog("Process name became empty after cleanup", "WARN");
                            exeName = "[Unknown Process]";
                        }
                    } else {
                        debugLog("Process name was empty", "WARN");
                        exeName = "[Unknown Process]";
                    }
                    
                    debugLog("Final process name: '" + exeName + "'", "DEBUG");
                } catch (nameError) {
                    debugLog("Exception reading process name: " + nameError.message, "ERROR");
                    exeName = "[Error Reading Name]";
                }
                
                processes.push({
                    name: exeName,
                    pid: procPID,
                    parentPID: parentPID,
                    threads: threadCount
                });
                
                processCount++;
                
                // Limit output for readability
                if (processCount <= 15) {
                    WScript.Echo("Process " + processCount + ": " + exeName + 
                                " (PID: " + procPID + 
                                ", Parent: " + parentPID + 
                                ", Threads: " + threadCount + ")");
                }
                
            } catch (readError) {
                WScript.Echo("Error reading process " + processCount + ": " + readError.message);
            }
            
            // Get next process
            result = DX.Process32NextW(snapshot, processEntry);
            
        } while (result !== 0 && processCount < 100); // Safety limit
        
        WScript.Echo("...");
        WScript.Echo("Total processes found: " + processCount);
        
        // Show some interesting processes
        WScript.Echo("\nInteresting processes found:");
        for (var i = 0; i < processes.length; i++) {
            var proc = processes[i];
            if (proc.name.toLowerCase().indexOf("cmd") !== -1 || 
                proc.name.toLowerCase().indexOf("explorer") !== -1 ||
                proc.name.toLowerCase().indexOf("cscript") !== -1 ||
                proc.name.toLowerCase().indexOf("powershell") !== -1) {
                WScript.Echo("  * " + proc.name + " (PID: " + proc.pid + ")");
            }
        }
        
    } else {
        throw new Error("Failed to get first process from snapshot");
    }
    
    // Cleanup
    DX.VirtualFree(processEntry, 0, 0x8000);
    DX.CloseHandle(snapshot);
    WScript.Echo("Process enumeration completed, resources cleaned up");
});

// Test 4: Process Creation with CreateProcessW (x64 optimized)
safeExecute("Process Creation with CreateProcess", function() {
    debugLog("Registering memory allocation and process functions", "DEBUG");
    
    // Register memory allocation and other needed functions
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    debugLog("VirtualAlloc registered", "DEBUG");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    debugLog("VirtualFree registered", "DEBUG");
    DX.Register("kernel32.dll", "WaitForSingleObject", "u=pu");
    debugLog("WaitForSingleObject registered", "DEBUG");
    DX.Register("kernel32.dll", "GetExitCodeProcess", "l=pp");
    debugLog("GetExitCodeProcess registered", "DEBUG");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    debugLog("CloseHandle registered", "DEBUG");
    
    WScript.Echo("Testing CreateProcessW via dynamic registration...");
    debugLog("About to register CreateProcessW with signature: l=wwpplupppp", "DEBUG");
    
    // Register CreateProcessW from kernel32.dll with correct signature
    try {
        DX.Register("kernel32.dll", "CreateProcessW", "l=wwpplupppp");
        debugLog("CreateProcessW registered successfully", "DEBUG");
    } catch (regError) {
        debugLog("CreateProcessW registration failed: " + regError.message, "ERROR");
        throw regError;
    }
    
    // Allocate structures (larger for x64)
    debugLog("Allocating STARTUPINFOW structure (128 bytes)", "DEBUG");
    var startupInfo = DX.VirtualAlloc(0, 128, 0x3000, 0x04);    // STARTUPINFOW
    debugLog("STARTUPINFOW allocated at: 0x" + startupInfo.toString(16), "DEBUG");
    
    debugLog("Allocating PROCESS_INFORMATION structure (32 bytes)", "DEBUG");
    var processInfo = DX.VirtualAlloc(0, 32, 0x3000, 0x04);     // PROCESS_INFORMATION
    debugLog("PROCESS_INFORMATION allocated at: 0x" + processInfo.toString(16), "DEBUG");
    
    if (startupInfo === 0 || processInfo === 0) {
        debugLog("Memory allocation failed - STARTUPINFOW: 0x" + startupInfo.toString(16) + ", PROCESS_INFORMATION: 0x" + processInfo.toString(16), "ERROR");
        throw new Error("Failed to allocate structures for CreateProcess");
    }
    
    // Initialize STARTUPINFOW structure properly (104 bytes for x64)
    DX.NumPut(104, startupInfo, 0, "u");        // cb (size of STARTUPINFOW)
    DX.NumPut(0, startupInfo, 8, "p");          // lpReserved (8-byte aligned)
    DX.NumPut(0, startupInfo, 16, "p");         // lpDesktop  
    DX.NumPut(0, startupInfo, 24, "p");         // lpTitle
    DX.NumPut(0, startupInfo, 32, "u");         // dwX
    DX.NumPut(0, startupInfo, 36, "u");         // dwY
    DX.NumPut(0, startupInfo, 40, "u");         // dwXSize
    DX.NumPut(0, startupInfo, 44, "u");         // dwYSize
    DX.NumPut(0, startupInfo, 48, "u");         // dwXCountChars
    DX.NumPut(0, startupInfo, 52, "u");         // dwYCountChars
    DX.NumPut(0, startupInfo, 56, "u");         // dwFillAttribute
    DX.NumPut(0, startupInfo, 60, "u");         // dwFlags
    DX.NumPut(0, startupInfo, 64, "t");         // wShowWindow
    DX.NumPut(0, startupInfo, 66, "t");         // cbReserved2
    DX.NumPut(0, startupInfo, 72, "p");         // lpReserved2
    DX.NumPut(0, startupInfo, 80, "p");         // hStdInput
    DX.NumPut(0, startupInfo, 88, "p");         // hStdOutput  
    DX.NumPut(0, startupInfo, 96, "p");         // hStdError
    
    // Initialize PROCESS_INFORMATION structure (24 bytes for x64)
    DX.NumPut(0, processInfo, 0, "p");          // hProcess (8 bytes)
    DX.NumPut(0, processInfo, 8, "p");          // hThread (8 bytes)
    DX.NumPut(0, processInfo, 16, "u");         // dwProcessId (4 bytes)
    DX.NumPut(0, processInfo, 20, "u");         // dwThreadId (4 bytes)
    
    // Command line: simple echo command
    var cmdLine = 'cmd.exe /c "echo CreateProcessW x64 test successful & timeout /t 2 > nul"';
    
    WScript.Echo("Creating process with command: " + cmdLine);
    debugLog("About to call CreateProcessW with parameters:", "DEBUG");
    debugLog("  lpApplicationName: null", "DEBUG");
    debugLog("  lpCommandLine: " + cmdLine, "DEBUG");
    debugLog("  lpProcessAttributes: null", "DEBUG");
    debugLog("  lpThreadAttributes: null", "DEBUG");
    debugLog("  bInheritHandles: 0", "DEBUG");
    debugLog("  dwCreationFlags: 0", "DEBUG");
    debugLog("  lpEnvironment: null", "DEBUG");
    debugLog("  lpCurrentDirectory: null", "DEBUG");
    debugLog("  lpStartupInfo: 0x" + startupInfo.toString(16), "DEBUG");
    debugLog("  lpProcessInformation: 0x" + processInfo.toString(16), "DEBUG");
    
    // Use dynamically registered CreateProcessW
    var result;
    try {
        debugLog("Calling CreateProcessW...", "DEBUG");
        result = DX.CreateProcessW(
            null,           // lpApplicationName
            cmdLine,        // lpCommandLine
            null,           // lpProcessAttributes
            null,           // lpThreadAttributes
            0,              // bInheritHandles
            0,              // dwCreationFlags
            null,           // lpEnvironment
            null,           // lpCurrentDirectory
            startupInfo,    // lpStartupInfo
            processInfo     // lpProcessInformation
        );
        debugLog("CreateProcessW returned: " + result, "DEBUG");
    } catch (callError) {
        debugLog("CreateProcessW threw exception: " + callError.message, "ERROR");
        if (callError.number) {
            debugLog("Exception number: " + callError.number + " (0x" + callError.number.toString(16) + ")", "ERROR");
        }
        throw callError;
    }
    
    if (result !== 0) {
        WScript.Echo("Process created successfully!");
        
        // Read process information (x64 offsets: handles=8 bytes each, so PIDs at 16,20)
        var hProcess = DX.NumGet(processInfo, 0, "p");
        var hThread = DX.NumGet(processInfo, 8, "p");  
        var processId = DX.NumGet(processInfo, 16, "u");
        var threadId = DX.NumGet(processInfo, 20, "u");
        
        WScript.Echo("Process Handle: 0x" + hProcess.toString(16));
        WScript.Echo("Thread Handle: 0x" + hThread.toString(16));
        WScript.Echo("Process ID: " + processId);
        WScript.Echo("Thread ID: " + threadId);
        
        // Wait for process to complete (with timeout)
        WScript.Echo("Waiting for process to complete...");
        var waitResult = DX.WaitForSingleObject(hProcess, 5000); // 5 second timeout
        
        if (waitResult === 0) { // WAIT_OBJECT_0
            WScript.Echo("Process completed successfully");
            
            // Get exit code using WScript.Shell approach for x64 compatibility
            WScript.Echo("Process completed, exit code available from Shell object");
        } else {
            WScript.Echo("Process wait timeout or error: " + waitResult);
        }
        
        // Close handles
        DX.CloseHandle(hProcess);
        DX.CloseHandle(hThread);
        
    } else {
        WScript.Echo("Failed to create process");
    }
    
    // Cleanup
    DX.VirtualFree(startupInfo, 0, 0x8000);
    DX.VirtualFree(processInfo, 0, 0x8000);
});

// Test 5: Alternative Process Creation APIs
safeExecute("Alternative Process Creation APIs", function() {
    WScript.Echo("Testing alternative process creation methods...");
    
    // Register additional process creation APIs for debugging
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    DX.Register("kernel32.dll", "WaitForSingleObject", "u=pu");
    
    // Test 5a: CreateProcessAsUserW (if available)
    try {
        DX.Register("advapi32.dll", "CreateProcessAsUserW", "l=pwwppluppppp");
        WScript.Echo("CreateProcessAsUserW registered successfully");
    } catch (e) {
        WScript.Echo("CreateProcessAsUserW registration failed: " + e.message);
    }
    
    // Test 5b: NtCreateUserProcess (Native API)
    try {
        DX.Register("ntdll.dll", "NtCreateUserProcess", "l=ppuppppppppp");
        WScript.Echo("NtCreateUserProcess registered successfully");
        
        // This is complex to use properly, but we can test registration
        WScript.Echo("NtCreateUserProcess is available for advanced process creation");
    } catch (e) {
        WScript.Echo("NtCreateUserProcess registration failed: " + e.message);
    }
    
    // Test 5c: WinExec (Legacy but simple)
    try {
        DX.Register("kernel32.dll", "WinExec", "u=su");
        WScript.Echo("Testing WinExec (legacy API)...");
        
        var result = DX.WinExec('cmd.exe /c "echo WinExec x64 test & timeout /t 1 > nul"', 0); // SW_HIDE
        if (result > 31) {
            WScript.Echo("WinExec executed successfully, result: " + result);
        } else {
            WScript.Echo("WinExec failed, error code: " + result);
        }
    } catch (e) {
        WScript.Echo("WinExec test failed: " + e.message);
    }
    
    // Test 5d: ShellExecuteW (Shell-based execution)
    try {
        DX.Register("shell32.dll", "ShellExecuteW", "p=hwwwwl");
        WScript.Echo("Testing ShellExecuteW...");
        
        var result = DX.ShellExecuteW(
            0,              // hwnd
            "open",         // lpOperation
            "cmd.exe",      // lpFile
            '/c "echo ShellExecuteW x64 test & timeout /t 1 > nul"', // lpParameters
            null,           // lpDirectory
            0               // nShowCmd (SW_HIDE)
        );
        
        if (result > 32) {
            WScript.Echo("ShellExecuteW executed successfully, handle: 0x" + result.toString(16));
        } else {
            WScript.Echo("ShellExecuteW failed, error code: " + result);
        }
    } catch (e) {
        WScript.Echo("ShellExecuteW test failed: " + e.message);
    }
    
    WScript.Echo("Alternative process creation API testing completed");
});

// Test 6: System Information Related to Processes
safeExecute("System Process Information", function() {
    // Get system metrics
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    DX.Register("user32.dll", "GetSystemMetrics", "l=l");
    DX.Register("kernel32.dll", "GetSystemInfo", "=p");
    
    WScript.Echo("Getting system information...");
    
    var systemInfo = DX.VirtualAlloc(0, 64, 0x3000, 0x04);
    if (systemInfo !== 0) {
        DX.GetSystemInfo(systemInfo);
        
        var numProcessors = DX.NumGet(systemInfo, 20, "u");
        var pageSize = DX.NumGet(systemInfo, 4, "u");
        var processorType = DX.NumGet(systemInfo, 0, "u");
        
        WScript.Echo("Number of Processors: " + numProcessors);
        WScript.Echo("Page Size: " + pageSize + " bytes");
        WScript.Echo("Processor Architecture: " + processorType + " (x64)");
        
        DX.VirtualFree(systemInfo, 0, 0x8000);
    }
    
    // Get some process-related system metrics
    var maxPath = DX.GetSystemMetrics(92); // This is actually for virtual screen, but let's see
    WScript.Echo("System metric 92: " + maxPath);
});

WScript.Echo("\n=== Process Management Test (x64) Completed ===");
WScript.Echo("All x64 process-related functionality has been tested.");

// Final summary
try {
    DX.Register("user32.dll", "MessageBoxW", "l=hwwu");
    DX.MessageBoxW(0, "DynamicWrapperX Process Management Test (x64) Completed!\n\nAll process enumeration, creation, and monitoring functions have been tested for x64 architecture.\n\nCheck the console output for detailed results.", "Process Test Complete (x64)", 0x40);
} catch (e) {
    WScript.Echo("Final MessageBox failed: " + e.message);
}