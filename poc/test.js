// Test script for DynamicWrapperX (JScript)
// Save as test.js and run with: cscript test.js

try {
    WScript.Echo("Testing DynamicWrapperX...");
    
    // Create DynamicWrapperX object
    var DX = new ActiveXObject("DynamicWrapperX");
    WScript.Echo("✓ DynamicWrapperX object created successfully");
    
    // Test 1: Simple API call - MessageBox
    WScript.Echo("\n=== Test 1: MessageBox ===");
    DX.Register("user32.dll", "MessageBoxW", "i=hwwu", "r=l");
    var result = DX.MessageBoxW(0, "Hello from DynamicWrapperX!", "Test", 0);
    WScript.Echo("MessageBox result: " + result);
    
    // Test 2: Get Windows directory
    WScript.Echo("\n=== Test 2: Get Windows Directory ===");
    DX.Register("kernel32.dll", "GetWindowsDirectoryW", "i=pl", "r=l");
    var buffer = DX.Space(520); // 260 chars * 2 bytes for Unicode
    var bufferPtr = DX.StrPtr(buffer);
    var length = DX.GetWindowsDirectoryW(bufferPtr, 260);
    if (length > 0) {
        var winDir = DX.StrGet(bufferPtr, length, "w");
        WScript.Echo("Windows directory: " + winDir);
    } else {
        WScript.Echo("Failed to get Windows directory");
    }
    
    // Test 3: Memory operations with NumGet/NumPut
    WScript.Echo("\n=== Test 3: Memory Operations ===");
    var testStr = "Hello, World!";
    
    // Read character codes
    var codes = "";
    for (var i = 0; i < Math.min(testStr.length, 10); i++) {
        var code = DX.NumGet(testStr, i * 2, "t");
        codes += code + " ";
    }
    WScript.Echo("Character codes: " + codes);
    
    // Create buffer and write data
    var writeBuffer = DX.Space(20);
    var addr = DX.StrPtr(writeBuffer);
    DX.NumPut(0x41, addr, 0, "t");  // Write 'A'
    DX.NumPut(0x42, addr, 2, "t");  // Write 'B'
    DX.NumPut(0x43, addr, 4, "t");  // Write 'C'
    DX.NumPut(0x00, addr, 6, "t");  // Null terminator
    WScript.Echo("Written to buffer: " + writeBuffer.substring(0, 3));
    
    // Test 4: String operations
    WScript.Echo("\n=== Test 4: String Operations ===");
    var originalStr = "Test String";
    var strPtr = DX.StrPtr(originalStr);
    WScript.Echo("String pointer: 0x" + strPtr.toString(16));
    
    var retrievedStr = DX.StrGet(strPtr);
    WScript.Echo("Retrieved string: " + retrievedStr);
    
    // Test ANSI conversion
    var ansiPtr = DX.StrPtr(originalStr, "s");
    var ansiStr = DX.StrGet(ansiPtr, "s");
    WScript.Echo("ANSI round-trip: " + ansiStr);
    
    // Test 5: Create spaces with custom character
    WScript.Echo("\n=== Test 5: Space Function ===");
    var spaces = DX.Space(10);
    WScript.Echo("10 spaces: '" + spaces + "' (length: " + spaces.length + ")");
    
    var stars = DX.Space(5, "*");
    WScript.Echo("5 stars: '" + stars + "'");
    
    var nulls = DX.Space(3, "");
    WScript.Echo("3 nulls: length = " + nulls.length);
    
    // Test 6: Get system info
    WScript.Echo("\n=== Test 6: System Information ===");
    DX.Register("kernel32.dll", "GetComputerNameW", "i=pP", "r=l");
    var nameBuffer = DX.Space(128); // 64 chars * 2 bytes for Unicode
    var nameBufferPtr = DX.StrPtr(nameBuffer);
    var lengthBuffer = DX.Space(4); // DWORD size buffer
    DX.NumPut(64, lengthBuffer, 0, "u"); // Set initial buffer size
    var lengthPtr = DX.StrPtr(lengthBuffer);
    
    if (DX.GetComputerNameW(nameBufferPtr, lengthPtr)) {
        var actualLength = DX.NumGet(lengthPtr, 0, "u");
        var computerName = DX.StrGet(nameBufferPtr, actualLength, "w");
        WScript.Echo("Computer name: " + computerName);
    } else {
        WScript.Echo("Failed to get computer name");
    }
    
    // Test 7: File operations
    WScript.Echo("\n=== Test 7: File Operations ===");
    DX.Register("kernel32.dll", "GetTempPathW", "i=lp", "r=l");
    var tempPathBuffer = DX.Space(520); // 260 chars * 2 bytes for Unicode
    var tempPathPtr = DX.StrPtr(tempPathBuffer);
    var tempPathLength = DX.GetTempPathW(260, tempPathPtr);
    if (tempPathLength > 0) {
        var tempPath = DX.StrGet(tempPathPtr, tempPathLength, "w");
        WScript.Echo("Temp path: " + tempPath);
    } else {
        WScript.Echo("Failed to get temp path");
    }
    
    // Test 8: Mathematical operations
    WScript.Echo("\n=== Test 8: Mathematical Operations ===");
    
    // Test floating point
    var floatBuffer = DX.Space(4);
    var floatPtr = DX.StrPtr(floatBuffer);
    DX.NumPut(3.14159, floatPtr, 0, "f");
    var retrievedFloat = DX.NumGet(floatPtr, 0, "f");
    WScript.Echo("Float round-trip: " + retrievedFloat);
    
    // Test double
    var doubleBuffer = DX.Space(8);
    var doublePtr = DX.StrPtr(doubleBuffer);
    DX.NumPut(2.71828182845904, doublePtr, 0, "d");
    var retrievedDouble = DX.NumGet(doublePtr, 0, "d");
    WScript.Echo("Double round-trip: " + retrievedDouble);
    
    // Test 9: Handle operations
    WScript.Echo("\n=== Test 9: Handle Operations ===");
    DX.Register("kernel32.dll", "GetCurrentProcess", "r=h", "");
    var processHandle = DX.GetCurrentProcess();
    WScript.Echo("Current process handle: 0x" + processHandle.toString(16));
    
    DX.Register("kernel32.dll", "GetCurrentProcessId", "r=u", "");
    var processId = DX.GetCurrentProcessId();
    WScript.Echo("Current process ID: " + processId);
    
    WScript.Echo("\n=== All Tests Completed Successfully! ===");
    
} catch (e) {
    WScript.Echo("Error: " + e.message);
    WScript.Echo("Error code: 0x" + (e.number >>> 0).toString(16));
}
