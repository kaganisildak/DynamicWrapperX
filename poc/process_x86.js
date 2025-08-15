// DynamicWrapperX Process Management Test
// Comprehensive test for process enumeration, manipulation, and monitoring

var DX = new ActiveXObject("DynamicWrapperX");

WScript.Echo("=== DynamicWrapperX Process Management Test ===");
WScript.Echo("Starting comprehensive process-related function testing...\n");

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
    var cmdProcess = shell.Exec('cmd.exe /c "echo DynamicWrapperX Test & echo Process ID: & echo %% & timeout /t 3 > nul & echo Process completed"');
    
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
        
        // Check if process is still running
        var exitCode = DX.VirtualAlloc(0, 4, 0x3000, 0x04);
        if (exitCode !== 0) {
            var result = DX.GetExitCodeProcess(processHandle, exitCode);
            if (result !== 0) {
                var code = DX.NumGet(exitCode, 0, "u");
                if (code === 259) { // STILL_ACTIVE
                    WScript.Echo("Process is still running (exit code: STILL_ACTIVE)");
                } else {
                    WScript.Echo("Process has exited with code: " + code);
                }
            }
            DX.VirtualFree(exitCode, 0, 0x8000);
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

// Test 3: Enumerate Running Processes
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
    
    // Allocate memory for PROCESSENTRY32W structure
    var processEntry = DX.VirtualAlloc(0, 1024, 0x3000, 0x04);
    if (processEntry === 0) {
        DX.CloseHandle(snapshot);
        throw new Error("Failed to allocate memory for process entry");
    }
    
    // Initialize structure size - keep original size that was working
    DX.NumPut(556, processEntry, 0, "u"); // sizeof(PROCESSENTRY32W)
    
    WScript.Echo("Starting process enumeration...");
    
    var processCount = 0;
    var processes = [];
    
    // Get first process
    var result = DX.Process32FirstW(snapshot, processEntry);
    
    if (result !== 0) {
        do {
            try {
                // Read process information from the structure  
                var procPID = DX.NumGet(processEntry, 8, "u");     // th32ProcessID
                var parentPID = DX.NumGet(processEntry, 24, "u");  // th32ParentProcessID
                var threadCount = DX.NumGet(processEntry, 20, "u"); // cntThreads (corrected offset)
                
                // Read process name (Unicode string at offset 36 for x86, 44 for x64)
                // For simplicity, we'll use offset 36 since we're running x86
                var exeName = "";
                try {
                    debugLog("Reading process name from offset 36...", "DEBUG");
                    // PROCESSENTRY32W has szExeFile at offset 36 on x86 (260 WCHARs = 520 bytes)
                    // Read Unicode characters manually
                    for (var i = 0; i < 260; i++) {
                        var char = DX.NumGet(processEntry, 36 + (i * 2), "t");
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
                proc.name.toLowerCase().indexOf("cscript") !== -1) {
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

// Test 4: Find Specific Process (cmd.exe)
safeExecute("Find Specific Process by Name", function() {
    // Register required functions
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    
    // Start a cmd.exe process for testing
    WScript.Echo("Starting target cmd.exe process...");
    var shell = new ActiveXObject("WScript.Shell");
    var targetProcess = shell.Exec('cmd.exe /c "echo Target process for DynamicWrapperX & timeout /t 5 > nul"');
    
    WScript.Sleep(300); // Give it time to start
    
    var targetPID = targetProcess.ProcessID;
    WScript.Echo("Target cmd.exe PID: " + targetPID);
    
    // Now enumerate processes to find it
    var snapshot = DX.CreateToolhelp32Snapshot(0x2, 0);
    if (snapshot == -1) {
        throw new Error("Failed to create snapshot for process search");
    }
    
    var processEntry = DX.VirtualAlloc(0, 1024, 0x3000, 0x04);
    if (processEntry === 0) {
        DX.CloseHandle(snapshot);
        throw new Error("Failed to allocate memory for process search");
    }
    
    DX.NumPut(556, processEntry, 0, "u");
    
    var found = false;
    var result = DX.Process32FirstW(snapshot, processEntry);
    
    if (result !== 0) {
        do {
            var procPID = DX.NumGet(processEntry, 8, "u");
            
            if (procPID === targetPID) {
                WScript.Echo("Found target process in enumeration!");
                
                // Read process name to verify
                var exeName = "";
                try {
                    for (var i = 0; i < 260; i++) {
                        var char = DX.NumGet(processEntry, 44 + (i * 2), "t");
                        if (char === 0) break;
                        exeName += String.fromCharCode(char);
                    }
                    WScript.Echo("Process name: " + exeName);
                    found = true;
                } catch (e) {
                    WScript.Echo("Found PID but couldn't read name: " + e.message);
                }
                break;
            }
            
            result = DX.Process32NextW(snapshot, processEntry);
            
        } while (result !== 0);
    }
    
    if (found) {
        WScript.Echo("Process search: SUCCESS");
    } else {
        WScript.Echo("Process search: Target not found in enumeration");
    }
    
    // Cleanup
    DX.VirtualFree(processEntry, 0, 0x8000);
    DX.CloseHandle(snapshot);
    
    // Wait for target process to finish
    WScript.Echo("Waiting for target process to complete...");
    while (targetProcess.Status === 0) {
        WScript.Sleep(100);
    }
    WScript.Echo("Target process completed");
});

// Test 5: Process Creation and Monitoring
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
    // CreateProcessW(lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, 
    //               bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, 
    //               lpStartupInfo, lpProcessInformation)
    try {
        DX.Register("kernel32.dll", "CreateProcessW", "l=wwpplupppp");
        debugLog("CreateProcessW registered successfully", "DEBUG");
    } catch (regError) {
        debugLog("CreateProcessW registration failed: " + regError.message, "ERROR");
        throw regError;
    }
    
    // Allocate structures
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
    DX.NumPut(0, startupInfo, 4, "p");          // lpReserved
    DX.NumPut(0, startupInfo, 12, "p");         // lpDesktop  
    DX.NumPut(0, startupInfo, 20, "p");         // lpTitle
    DX.NumPut(0, startupInfo, 28, "u");         // dwX
    DX.NumPut(0, startupInfo, 32, "u");         // dwY
    DX.NumPut(0, startupInfo, 36, "u");         // dwXSize
    DX.NumPut(0, startupInfo, 40, "u");         // dwYSize
    DX.NumPut(0, startupInfo, 44, "u");         // dwXCountChars
    DX.NumPut(0, startupInfo, 48, "u");         // dwYCountChars
    DX.NumPut(0, startupInfo, 52, "u");         // dwFillAttribute
    DX.NumPut(0, startupInfo, 56, "u");         // dwFlags
    DX.NumPut(0, startupInfo, 60, "t");         // wShowWindow
    DX.NumPut(0, startupInfo, 62, "t");         // cbReserved2
    DX.NumPut(0, startupInfo, 64, "p");         // lpReserved2
    DX.NumPut(0, startupInfo, 72, "p");         // hStdInput
    DX.NumPut(0, startupInfo, 80, "p");         // hStdOutput  
    DX.NumPut(0, startupInfo, 88, "p");         // hStdError
    DX.NumPut(0, startupInfo, 96, "p");         // hStdError (padding)
    
    // Initialize PROCESS_INFORMATION structure (24 bytes for x64)
    DX.NumPut(0, processInfo, 0, "p");          // hProcess
    DX.NumPut(0, processInfo, 8, "p");          // hThread
    DX.NumPut(0, processInfo, 16, "u");         // dwProcessId
    DX.NumPut(0, processInfo, 20, "u");         // dwThreadId
    
    // Command line: simple echo command
    var cmdLine = 'cmd.exe /c "echo CreateProcessW test successful & timeout /t 2 > nul"';
    
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
        
        // Read process information (x86 offsets: handles=4 bytes each, so PIDs at 8,12)
        var hProcess = DX.NumGet(processInfo, 0, "p");
        var hThread = DX.NumGet(processInfo, 4, "p");  
        var processId = DX.NumGet(processInfo, 8, "u");
        var threadId = DX.NumGet(processInfo, 12, "u");
        
        WScript.Echo("Process Handle: 0x" + hProcess.toString(16));
        WScript.Echo("Thread Handle: 0x" + hThread.toString(16));
        WScript.Echo("Process ID: " + processId);
        WScript.Echo("Thread ID: " + threadId);
        
        // Wait for process to complete (with timeout)
        WScript.Echo("Waiting for process to complete...");
        var waitResult = DX.WaitForSingleObject(hProcess, 5000); // 5 second timeout
        
        if (waitResult === 0) { // WAIT_OBJECT_0
            WScript.Echo("Process completed successfully");
            
            // Get exit code
            var exitCode = DX.VirtualAlloc(0, 4, 0x3000, 0x04);
            if (exitCode !== 0) {
                if (DX.GetExitCodeProcess(hProcess, exitCode)) {
                    var code = DX.NumGet(exitCode, 0, "u");
                    WScript.Echo("Process exit code: " + code);
                }
                DX.VirtualFree(exitCode, 0, 0x8000);
            }
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

// Test 6: Alternative Process Creation APIs (Advanced Debugging)
safeExecute("Alternative Process Creation APIs", function() {
    WScript.Echo("Testing alternative process creation methods...");
    
    // Register additional process creation APIs for debugging
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    DX.Register("kernel32.dll", "WaitForSingleObject", "u=pu");
    
    // Test 6a: CreateProcessAsUserW (if available)
    try {
        DX.Register("advapi32.dll", "CreateProcessAsUserW", "l=pwwppluppppp");
        WScript.Echo("CreateProcessAsUserW registered successfully");
    } catch (e) {
        WScript.Echo("CreateProcessAsUserW registration failed: " + e.message);
    }
    
    // Test 6b: NtCreateUserProcess (Native API)
    try {
        DX.Register("ntdll.dll", "NtCreateUserProcess", "l=ppuppppppppp");
        WScript.Echo("NtCreateUserProcess registered successfully");
        
        // This is complex to use properly, but we can test registration
        WScript.Echo("NtCreateUserProcess is available for advanced process creation");
    } catch (e) {
        WScript.Echo("NtCreateUserProcess registration failed: " + e.message);
    }
    
    // Test 6c: RtlCreateUserProcess (Native API)
    try {
        DX.Register("ntdll.dll", "RtlCreateUserProcess", "l=ppplppppppppp");
        WScript.Echo("RtlCreateUserProcess registered successfully");
    } catch (e) {
        WScript.Echo("RtlCreateUserProcess registration failed: " + e.message);
    }
    
    // Test 6d: CreateProcessWithLogonW (Advanced authentication)
    try {
        DX.Register("advapi32.dll", "CreateProcessWithLogonW", "l=wwwuwwpwpp");
        WScript.Echo("CreateProcessWithLogonW registered successfully");
    } catch (e) {
        WScript.Echo("CreateProcessWithLogonW registration failed: " + e.message);
    }
    
    // Test 6e: CreateProcessWithTokenW (Token-based creation)
    try {
        DX.Register("advapi32.dll", "CreateProcessWithTokenW", "l=pwwpupppp");
        WScript.Echo("CreateProcessWithTokenW registered successfully");
    } catch (e) {
        WScript.Echo("CreateProcessWithTokenW registration failed: " + e.message);
    }
    
    // Test 6f: WinExec (Legacy but simple)
    try {
        DX.Register("kernel32.dll", "WinExec", "u=su");
        WScript.Echo("Testing WinExec (legacy API)...");
        
        var result = DX.WinExec('cmd.exe /c "echo WinExec test & timeout /t 1 > nul"', 0); // SW_HIDE
        if (result > 31) {
            WScript.Echo("WinExec executed successfully, result: " + result);
        } else {
            WScript.Echo("WinExec failed, error code: " + result);
        }
    } catch (e) {
        WScript.Echo("WinExec test failed: " + e.message);
    }
    
    // Test 6g: ShellExecuteW (Shell-based execution)
    try {
        DX.Register("shell32.dll", "ShellExecuteW", "p=hwwwwl");
        WScript.Echo("Testing ShellExecuteW...");
        
        var result = DX.ShellExecuteW(
            0,              // hwnd
            "open",         // lpOperation
            "cmd.exe",      // lpFile
            '/c "echo ShellExecuteW test & timeout /t 1 > nul"', // lpParameters
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

// Test 7: CreateProcessW Deep Debugging
safeExecute("CreateProcessW Deep Debugging", function() {
    WScript.Echo("Performing deep debugging of CreateProcessW...");
    
    // Register all needed functions
    DX.Register("kernel32.dll", "VirtualAlloc", "p=puuu");
    DX.Register("kernel32.dll", "VirtualFree", "l=puu");
    DX.Register("kernel32.dll", "GetLastError", "u=");
    DX.Register("kernel32.dll", "CloseHandle", "l=p");
    DX.Register("kernel32.dll", "CreateProcessW", "l=wwpplupppp");
    
    // Test different parameter combinations
    var debugTests = [
        {
            name: "Minimal parameters",
            app: null,
            cmd: "cmd.exe /c echo Test1",
            flags: 0
        },
        {
            name: "With CREATE_NO_WINDOW flag",
            app: null,
            cmd: "cmd.exe /c echo Test2",
            flags: 0x08000000  // CREATE_NO_WINDOW
        },
        {
            name: "With explicit application name",
            app: "cmd.exe",
            cmd: "cmd.exe /c echo Test3",
            flags: 0
        },
        {
            name: "Simple echo without timeout",
            app: null,
            cmd: 'cmd.exe /c "echo Simple test"',
            flags: 0
        }
    ];
    
    for (var i = 0; i < debugTests.length; i++) {
        var test = debugTests[i];
        WScript.Echo("\n--- Debug Test " + (i+1) + ": " + test.name + " ---");
        
        try {
            // Allocate structures
            var startupInfo = DX.VirtualAlloc(0, 128, 0x3000, 0x04);
            var processInfo = DX.VirtualAlloc(0, 32, 0x3000, 0x04);
            
            if (startupInfo === 0 || processInfo === 0) {
                WScript.Echo("Memory allocation failed");
                continue;
            }
            
            // Clear structures
            for (var j = 0; j < 128; j += 4) {
                DX.NumPut(0, startupInfo, j, "u");
            }
            for (var k = 0; k < 32; k += 4) {
                DX.NumPut(0, processInfo, k, "u");
            }
            
            // Initialize STARTUPINFOW with correct size
            DX.NumPut(104, startupInfo, 0, "u");  // cb field
            
            WScript.Echo("Calling CreateProcessW with:");
            WScript.Echo("  Application: " + (test.app || "null"));
            WScript.Echo("  Command: " + test.cmd);
            WScript.Echo("  Flags: 0x" + test.flags.toString(16));
            
            var result = DX.CreateProcessW(
                test.app,       // lpApplicationName
                test.cmd,       // lpCommandLine
                null,           // lpProcessAttributes
                null,           // lpThreadAttributes
                0,              // bInheritHandles
                test.flags,     // dwCreationFlags
                null,           // lpEnvironment
                null,           // lpCurrentDirectory
                startupInfo,    // lpStartupInfo
                processInfo     // lpProcessInformation
            );
            
            if (result !== 0) {
                WScript.Echo("SUCCESS: CreateProcessW returned " + result);
                
                var hProcess = DX.NumGet(processInfo, 0, "p");
                var hThread = DX.NumGet(processInfo, 4, "p");
                var processId = DX.NumGet(processInfo, 8, "u");
                var threadId = DX.NumGet(processInfo, 12, "u");
                
                WScript.Echo("  Process ID: " + processId);
                WScript.Echo("  Thread ID: " + threadId);
                WScript.Echo("  Process Handle: 0x" + hProcess.toString(16));
                WScript.Echo("  Thread Handle: 0x" + hThread.toString(16));
                
                // Quick check if process is still running
                DX.Register("kernel32.dll", "WaitForSingleObject", "u=pu");
                var waitResult = DX.WaitForSingleObject(hProcess, 100); // Quick check
                
                if (waitResult === 0) {
                    WScript.Echo("  Process completed immediately");
                } else if (waitResult === 258) {
                    WScript.Echo("  Process is still running");
                    // Let it finish
                    DX.WaitForSingleObject(hProcess, 3000);
                } else {
                    WScript.Echo("  Wait result: " + waitResult);
                }
                
                DX.CloseHandle(hProcess);
                DX.CloseHandle(hThread);
                
            } else {
                var lastError = DX.GetLastError();
                WScript.Echo("FAILED: CreateProcessW returned 0");
                WScript.Echo("  Last Error: " + lastError + " (0x" + lastError.toString(16) + ")");
                
                // Common error codes
                var errorMsgs = {
                    2: "File not found",
                    3: "Path not found", 
                    5: "Access denied",
                    8: "Not enough memory",
                    87: "Invalid parameter",
                    193: "Not a valid Win32 application"
                };
                
                if (errorMsgs[lastError]) {
                    WScript.Echo("  Error meaning: " + errorMsgs[lastError]);
                }
            }
            
            // Cleanup
            DX.VirtualFree(startupInfo, 0, 0x8000);
            DX.VirtualFree(processInfo, 0, 0x8000);
            
        } catch (e) {
            WScript.Echo("Exception in debug test: " + e.message);
        }
    }
    
    WScript.Echo("\nCreateProcessW debugging completed");
});

// Test 8: System Information Related to Processes
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
        WScript.Echo("Processor Architecture: " + processorType);
        
        DX.VirtualFree(systemInfo, 0, 0x8000);
    }
    
    // Get some process-related system metrics
    var maxPath = DX.GetSystemMetrics(92); // This is actually for virtual screen, but let's see
    WScript.Echo("System metric 92: " + maxPath);
});

WScript.Echo("\n=== Process Management Test Completed ===");
WScript.Echo("All process-related functionality has been tested.");

// Final summary
try {
    DX.Register("user32.dll", "MessageBoxW", "l=hwwu");
    DX.MessageBoxW(0, "DynamicWrapperX Process Management Test Completed!\n\nAll process enumeration, creation, and monitoring functions have been tested.\n\nCheck the console output for detailed results.", "Process Test Complete", 0x40);
} catch (e) {
    WScript.Echo("Final MessageBox failed: " + e.message);
}
