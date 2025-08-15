// Simple DynamicWrapperX Test - MessageBox and IsDebuggerPresent
// Run this script with: cscript test_simple.js

// Create DynamicWrapperX object
var DX = new ActiveXObject("DynamicWrapperX");

WScript.Echo("Testing DynamicWrapperX DLL...");

try {
    // Test 1: MessageBox
    WScript.Echo("1. Testing MessageBox...");
    DX.Register("user32.dll", "MessageBoxW", "l=hwwu");
    
    var mbResult = DX.MessageBoxW(
        0,  // No parent window
        "DynamicWrapperX is working!\n\nClick OK to continue with IsDebuggerPresent test.",  // Message
        "Success!",  // Title
        0x40 | 0x1  // MB_ICONINFORMATION | MB_OKCANCEL
    );
    
    WScript.Echo("MessageBox result: " + mbResult + " (1=OK, 2=Cancel)");
    
    // Test 2: IsDebuggerPresent
    WScript.Echo("2. Testing IsDebuggerPresent...");
    DX.Register("kernel32.dll", "IsDebuggerPresent", "l=");
    
    var debuggerResult = DX.IsDebuggerPresent();
    
    WScript.Echo("IsDebuggerPresent result: " + debuggerResult);
    
    // Show result in a message box
    var debugMsg = debuggerResult ? 
        "DEBUGGER DETECTED!\n\nIsDebuggerPresent() returned: " + debuggerResult :
        "No debugger detected.\n\nIsDebuggerPresent() returned: " + debuggerResult;
    
    DX.MessageBoxW(
        0, 
        debugMsg, 
        "Debugger Check Result", 
        debuggerResult ? 0x30 : 0x40  // Warning icon if debugger, Info icon if not
    );
    
    WScript.Echo("Test completed successfully!");
    
} catch (error) {
    WScript.Echo("ERROR: " + error.message);
    WScript.Echo("Make sure DynamicWrapperX is registered:");
    WScript.Echo("  regsvr32 dynwrapx.dll");
}
