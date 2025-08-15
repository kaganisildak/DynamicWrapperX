#include "dynwrapx.h"
#include <stdio.h>
#include <time.h>

// Debug logging system
#ifdef DEBUG_BUILD
static FILE* g_debugFile = NULL;
static CRITICAL_SECTION g_debugCS;
static BOOL g_debugInitialized = FALSE;

void InitializeDebugLogging(void)
{
    if (g_debugInitialized) return;
    
    InitializeCriticalSection(&g_debugCS);
    g_debugFile = fopen("C:\\Users\\kaganisildak\\Desktop\\test2\\debugdyn.log", "w");
    if (g_debugFile) {
        fprintf(g_debugFile, "=== DynamicWrapperX Debug Log Started ===\n");
        fflush(g_debugFile);
    }
    g_debugInitialized = TRUE;
}
#else
void InitializeDebugLogging(void) { /* No-op in release build */ }
#endif

#ifdef DEBUG_BUILD
void CleanupDebugLogging(void)
{
    if (!g_debugInitialized) return;
    
    EnterCriticalSection(&g_debugCS);
    if (g_debugFile) {
        fprintf(g_debugFile, "=== DynamicWrapperX Debug Log Ended ===\n");
        fclose(g_debugFile);
        g_debugFile = NULL;
    }
    LeaveCriticalSection(&g_debugCS);
    DeleteCriticalSection(&g_debugCS);
    g_debugInitialized = FALSE;
}
#else
void CleanupDebugLogging(void) { /* No-op in release build */ }
#endif

#ifdef DEBUG_BUILD
void DebugLog(const char* level, const char* function, const char* format, ...)
{
    if (!g_debugInitialized || !g_debugFile) return;
    
    EnterCriticalSection(&g_debugCS);
    
    time_t rawtime;
    struct tm* timeinfo;
    char timestamp[64];
    
    time(&rawtime);
    timeinfo = localtime(&rawtime);
    strftime(timestamp, sizeof(timestamp), "%Y-%m-%d %H:%M:%S", timeinfo);
    
    fprintf(g_debugFile, "[%s] %s - %s: ", level, timestamp, function);
    
    va_list args;
    va_start(args, format);
    vfprintf(g_debugFile, format, args);
    va_end(args);
    
    fprintf(g_debugFile, "\n");
    fflush(g_debugFile);
    
    LeaveCriticalSection(&g_debugCS);
}
#endif
// Note: Release builds use the macro definition in dynwrapx.h

// Missing constants
#ifndef SELFREG_E_CLASS
#include "dynwrapx.h"
#include <objbase.h>
#include <shlwapi.h>

// Define GUIDs
const GUID CLSID_DynamicWrapperX = {0x89565275, 0xA714, 0x4a43, {0x91, 0x2E, 0x97, 0x8B, 0x93, 0x5E, 0xDC, 0xCC}};
const GUID IID_IDynamicWrapperX = {0x89565276, 0xA714, 0x4a43, {0x91, 0x2E, 0x97, 0x8B, 0x93, 0x5E, 0xDC, 0xCC}};

// Constants for COM registration
#define SELFREG_E_CLASS 0x80040201L
#endif

// Global variables
HMODULE g_hModule = NULL;
LONG g_cObjects = 0;
LONG g_cLocks = 0;
DynamicWrapperX* g_callbackObjects[16] = {0};

// DLL Entry Point
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hModule = hModule;
        DisableThreadLibraryCalls(hModule);
        InitializeDebugLogging();
        DebugLog("INFO", "DllMain", "DLL_PROCESS_ATTACH - DLL loaded, hModule=0x%p", hModule);
        break;
    case DLL_PROCESS_DETACH:
        DebugLog("INFO", "DllMain", "DLL_PROCESS_DETACH - DLL unloading");
        CleanupDebugLogging();
        break;
    }
    return TRUE;
}

// DLL can unload check
STDAPI DllCanUnloadNow(void)
{
    return (g_cObjects == 0 && g_cLocks == 0) ? S_OK : S_FALSE;
}

// Registry functions for COM registration
STDAPI DllRegisterServer(void)
{
    HKEY hKey;
    LONG result;
    WCHAR szModule[MAX_PATH];
    WCHAR szCLSID[] = L"Software\\Classes\\CLSID\\{89565275-A714-4a43-912E-978B935EDCCC}\\InProcServer32";
    WCHAR szProgID[] = L"Software\\Classes\\DynamicWrapperX\\CLSID";
    WCHAR szGUID[] = L"{89565275-A714-4a43-912E-978B935EDCCC}";
    
    // Get module file name
    GetModuleFileNameW(g_hModule, szModule, MAX_PATH);
    
    // Register CLSID
    result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, szCLSID, 0, NULL, 
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (result != ERROR_SUCCESS)
        return SELFREG_E_CLASS;
        
    result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)szModule, 
                           (wcslen(szModule) + 1) * sizeof(WCHAR));
    RegCloseKey(hKey);
    
    if (result != ERROR_SUCCESS)
        return SELFREG_E_CLASS;
    
    // Register ProgID
    result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, szProgID, 0, NULL,
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (result != ERROR_SUCCESS)
        return SELFREG_E_CLASS;
        
    result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)szGUID,
                           (wcslen(szGUID) + 1) * sizeof(WCHAR));
    RegCloseKey(hKey);
    
    return (result == ERROR_SUCCESS) ? S_OK : SELFREG_E_CLASS;
}

STDAPI DllUnregisterServer(void)
{
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"Software\\Classes\\CLSID\\{89565275-A714-4a43-912E-978B935EDCCC}\\InProcServer32");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"Software\\Classes\\CLSID\\{89565275-A714-4a43-912E-978B935EDCCC}");
    RegDeleteKeyW(HKEY_LOCAL_MACHINE, L"Software\\Classes\\DynamicWrapperX");
    return S_OK;
}

STDAPI DllInstall(BOOL bInstall, PCWSTR pszCmdLine)
{
    if (bInstall)
        return DllRegisterServer();
    else
        return DllUnregisterServer();
}

// Parameter type parsing
ParameterType ParseParameterType(WCHAR c)
{
    switch (c)
    {
    case L'l': return TYPE_LONG;
    case L'u': return TYPE_ULONG;
    case L'h': return TYPE_HANDLE;
    case L'p': return TYPE_POINTER;
    case L'n': return TYPE_SHORT;
    case L't': return TYPE_USHORT;
    case L'c': return TYPE_CHAR;
    case L'b': return TYPE_UCHAR;
    case L'f': return TYPE_FLOAT;
    case L'd': return TYPE_DOUBLE;
    case L'q': return TYPE_LONGLONG;
    case L'w': return TYPE_WSTRING;
    case L's': return TYPE_ASTRING;
    case L'z': return TYPE_OSTRING;
    
    // Output parameters (uppercase)
    case L'L': return TYPE_OUT_LONG;
    case L'U': return TYPE_OUT_ULONG;
    case L'H': return TYPE_OUT_HANDLE;
    case L'P': return TYPE_OUT_POINTER;
    case L'N': return TYPE_OUT_SHORT;
    case L'T': return TYPE_OUT_USHORT;
    case L'C': return TYPE_OUT_CHAR;
    case L'B': return TYPE_OUT_UCHAR;
    case L'F': return TYPE_OUT_FLOAT;
    case L'D': return TYPE_OUT_DOUBLE;
    case L'W': return TYPE_OUT_WSTRING;
    case L'S': return TYPE_OUT_ASTRING;
    case L'Z': return TYPE_OUT_OSTRING;
    case L'v': return TYPE_VOID;
    
    default: return TYPE_LONG; // Default fallback
    }
}

// Load function from DLL
FARPROC LoadFunction(LPCWSTR libraryName, LPCWSTR functionName)
{
    HMODULE hLib;
    FARPROC proc;
    char szFuncName[256];
    char szLibName[MAX_PATH];
    
    // Convert to ANSI for logging
    WideCharToMultiByte(CP_ACP, 0, libraryName, -1, szLibName, MAX_PATH, NULL, NULL);
    WideCharToMultiByte(CP_ACP, 0, functionName, -1, szFuncName, 256, NULL, NULL);
    
    DebugLog("DEBUG", "LoadFunction", "Loading function '%s' from library '%s'", szFuncName, szLibName);
    
    // Try loading as Unicode first
    hLib = LoadLibraryW(libraryName);
    if (!hLib)
    {
        DebugLog("DEBUG", "LoadFunction", "LoadLibraryW failed, trying LoadLibraryA");
        // Try again with ANSI
        hLib = LoadLibraryA(szLibName);
        if (!hLib)
        {
            DebugLog("ERROR", "LoadFunction", "Failed to load library '%s', GetLastError=%lu", szLibName, GetLastError());
            return NULL;
        }
    }
    
    DebugLog("DEBUG", "LoadFunction", "Library loaded successfully, hLib=0x%p", hLib);
    
    // Try to get the function
    proc = GetProcAddress(hLib, szFuncName);
    if (!proc)
    {
        DebugLog("DEBUG", "LoadFunction", "GetProcAddress failed for '%s', trying with 'A' suffix", szFuncName);
        // Try with 'A' suffix for ANSI version
        strcat(szFuncName, "A");
        proc = GetProcAddress(hLib, szFuncName);
    }
    
    if (proc) {
        DebugLog("DEBUG", "LoadFunction", "Function '%s' loaded successfully, proc=0x%p", szFuncName, proc);
    } else {
        DebugLog("ERROR", "LoadFunction", "Failed to load function '%s', GetLastError=%lu", szFuncName, GetLastError());
    }
    
    return proc;
}

// Convert VARIANT to specific type
HRESULT ConvertVariantToType(VARIANT* pVar, ParameterType type, void* pOut, DynamicWrapperX* pObj)
{
    HRESULT hr = S_OK;
    VARIANT vTemp;
    
    DebugLog("DEBUG", "ConvertVariantToType", "Converting VARIANT vt=%d to type=%d", pVar ? pVar->vt : -1, type);
    
    VariantInit(&vTemp);
    
    switch (type)
    {
    case TYPE_LONG:
    case TYPE_OUT_LONG:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_I4);
        if (SUCCEEDED(hr))
            *(LONG*)pOut = V_I4(&vTemp);
        break;
        
    case TYPE_ULONG:
    case TYPE_OUT_ULONG:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_UI4);
        if (SUCCEEDED(hr))
            *(ULONG*)pOut = V_UI4(&vTemp);
        break;
        
    case TYPE_HANDLE:
    case TYPE_POINTER:
    case TYPE_OUT_HANDLE:
    case TYPE_OUT_POINTER:
        DebugLog("DEBUG", "ConvertVariantToType", "Converting POINTER/HANDLE, vt=%d", pVar->vt);
        // Handle various VARIANT types that might represent pointers
        switch (pVar->vt) {
            case VT_NULL:
            case VT_EMPTY:
                DebugLog("DEBUG", "ConvertVariantToType", "Setting pointer to NULL (VT_NULL/VT_EMPTY)");
                *(void**)pOut = NULL;
                hr = S_OK;
                break;
            case VT_I4:
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_I4 %ld to pointer", pVar->lVal);
                *(void**)pOut = (void*)(LONG_PTR)pVar->lVal;
                hr = S_OK;
                break;
            case VT_UI4:
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_UI4 %lu to pointer", pVar->ulVal);
                *(void**)pOut = (void*)(ULONG_PTR)pVar->ulVal;
                hr = S_OK;
                break;
            case VT_I8:
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_I8 %lld to pointer", pVar->llVal);
                *(void**)pOut = (void*)(LONG_PTR)pVar->llVal;
                hr = S_OK;
                break;
            case VT_UI8:
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_UI8 %llu to pointer", pVar->ullVal);
                *(void**)pOut = (void*)(ULONG_PTR)pVar->ullVal;
                hr = S_OK;
                break;
            case VT_R8:
                // Handle JavaScript numbers passed as doubles
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_R8 %f to pointer", pVar->dblVal);
                *(void**)pOut = (void*)(ULONG_PTR)pVar->dblVal;
                hr = S_OK;
                break;
            case VT_BSTR:
                DebugLog("DEBUG", "ConvertVariantToType", "Converting VT_BSTR to pointer (string address)");
                *(void**)pOut = (void*)pVar->bstrVal;
                hr = S_OK;
                break;
            default:
                DebugLog("DEBUG", "ConvertVariantToType", "Trying VariantChangeType for vt=%d", pVar->vt);
                // Try to convert to appropriate integer type
#ifdef _WIN64
                hr = VariantChangeType(&vTemp, pVar, 0, VT_I8);
                if (SUCCEEDED(hr)) {
                    DebugLog("DEBUG", "ConvertVariantToType", "Successfully converted to VT_I8: %lld", V_I8(&vTemp));
                    *(void**)pOut = (void*)V_I8(&vTemp);
                } else {
                    DebugLog("DEBUG", "ConvertVariantToType", "VT_I8 conversion failed (hr=0x%08x), trying VT_UI4", hr);
                    hr = VariantChangeType(&vTemp, pVar, 0, VT_UI4);
                    if (SUCCEEDED(hr)) {
                        DebugLog("DEBUG", "ConvertVariantToType", "Successfully converted to VT_UI4: %lu", V_UI4(&vTemp));
                        *(void**)pOut = (void*)(ULONG_PTR)V_UI4(&vTemp);
                    } else {
                        DebugLog("ERROR", "ConvertVariantToType", "Both VT_I8 and VT_UI4 conversions failed, hr=0x%08x", hr);
                    }
                }
#else
                hr = VariantChangeType(&vTemp, pVar, 0, VT_UI4);
                if (SUCCEEDED(hr))
                    *(void**)pOut = (void*)V_UI4(&vTemp);
                else {
                    hr = VariantChangeType(&vTemp, pVar, 0, VT_I4);
                    if (SUCCEEDED(hr))
                        *(void**)pOut = (void*)(LONG_PTR)V_I4(&vTemp);
                }
#endif
                break;
        }
        break;
        
    case TYPE_LONGLONG:
        // Handle various VARIANT types that might represent 64-bit values
        switch (pVar->vt) {
            case VT_I8:
                *(LONGLONG*)pOut = pVar->llVal;
                hr = S_OK;
                break;
            case VT_UI8:
                *(LONGLONG*)pOut = (LONGLONG)pVar->ullVal;
                hr = S_OK;
                break;
            case VT_I4:
                *(LONGLONG*)pOut = (LONGLONG)pVar->lVal;
                hr = S_OK;
                break;
            case VT_UI4:
                *(LONGLONG*)pOut = (LONGLONG)pVar->ulVal;
                hr = S_OK;
                break;
            case VT_R8:
                *(LONGLONG*)pOut = (LONGLONG)pVar->dblVal;
                hr = S_OK;
                break;
            default:
                // Try to convert to VT_I8 first, then fallback
                hr = VariantChangeType(&vTemp, pVar, 0, VT_I8);
                if (SUCCEEDED(hr))
                    *(LONGLONG*)pOut = V_I8(&vTemp);
                else {
                    hr = VariantChangeType(&vTemp, pVar, 0, VT_I4);
                    if (SUCCEEDED(hr))
                        *(LONGLONG*)pOut = (LONGLONG)V_I4(&vTemp);
                }
                break;
        }
        break;
        
    case TYPE_SHORT:
    case TYPE_OUT_SHORT:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_I2);
        if (SUCCEEDED(hr))
            *(SHORT*)pOut = V_I2(&vTemp);
        break;
        
    case TYPE_USHORT:
    case TYPE_OUT_USHORT:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_UI2);
        if (SUCCEEDED(hr))
            *(USHORT*)pOut = V_UI2(&vTemp);
        break;
        
    case TYPE_CHAR:
    case TYPE_OUT_CHAR:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_I1);
        if (SUCCEEDED(hr))
            *(CHAR*)pOut = V_I1(&vTemp);
        break;
        
    case TYPE_UCHAR:
    case TYPE_OUT_UCHAR:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_UI1);
        if (SUCCEEDED(hr))
            *(UCHAR*)pOut = V_UI1(&vTemp);
        break;
        
    case TYPE_FLOAT:
    case TYPE_OUT_FLOAT:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_R4);
        if (SUCCEEDED(hr))
            *(FLOAT*)pOut = V_R4(&vTemp);
        break;
        
    case TYPE_DOUBLE:
    case TYPE_OUT_DOUBLE:
        hr = VariantChangeType(&vTemp, pVar, 0, VT_R8);
        if (SUCCEEDED(hr))
            *(DOUBLE*)pOut = V_R8(&vTemp);
        break;
        
    case TYPE_WSTRING:
    case TYPE_ASTRING:
    case TYPE_OSTRING:
    case TYPE_OUT_WSTRING:
    case TYPE_OUT_ASTRING:
    case TYPE_OUT_OSTRING:
        DebugLog("DEBUG", "ConvertVariantToType", "Converting STRING type, vt=%d to string type %d", pVar->vt, type);
        
        // Handle NULL strings explicitly
        if (pVar->vt == VT_NULL || pVar->vt == VT_EMPTY) {
            DebugLog("DEBUG", "ConvertVariantToType", "NULL string parameter, setting to NULL");
            *(void**)pOut = NULL;
            hr = S_OK;
            break;
        }
        
        hr = VariantChangeType(&vTemp, pVar, 0, VT_BSTR);
        if (SUCCEEDED(hr))
        {
            DebugLog("DEBUG", "ConvertVariantToType", "Successfully converted to BSTR");
            if (type == TYPE_WSTRING || type == TYPE_OUT_WSTRING)
            {
                *(BSTR*)pOut = V_BSTR(&vTemp);
                V_BSTR(&vTemp) = NULL; // Transfer ownership
                DebugLog("DEBUG", "ConvertVariantToType", "Stored as Unicode string");
            }
            else
            {
                // Convert to ANSI/OEM
                int len = SysStringLen(V_BSTR(&vTemp));
                char* pszAnsi = (char*)GlobalAlloc(GMEM_FIXED, len + 1);
                if (pszAnsi)
                {
                    UINT codePage = (type == TYPE_OSTRING || type == TYPE_OUT_OSTRING) ? CP_OEMCP : CP_ACP;
                    WideCharToMultiByte(codePage, 0, V_BSTR(&vTemp), len, pszAnsi, len + 1, NULL, NULL);
                    pszAnsi[len] = '\0';
                    *(char**)pOut = pszAnsi;
                    
                    DebugLog("DEBUG", "ConvertVariantToType", "Converted to ANSI string, len=%d", len);
                    
                    // Track memory allocation
                    MemoryBlock* pBlock = (MemoryBlock*)GlobalAlloc(GMEM_FIXED, sizeof(MemoryBlock));
                    if (pBlock)
                    {
                        pBlock->ptr = pszAnsi;
                        pBlock->isGlobal = TRUE;
                        pBlock->next = pObj->memoryBlocks;
                        pObj->memoryBlocks = pBlock;
                    }
                }
                else
                {
                    DebugLog("ERROR", "ConvertVariantToType", "Failed to allocate memory for ANSI string");
                    hr = E_OUTOFMEMORY;
                }
            }
        }
        else
        {
            DebugLog("ERROR", "ConvertVariantToType", "Failed to convert to BSTR, hr=0x%08x", hr);
        }
        break;
        
    default:
        DebugLog("ERROR", "ConvertVariantToType", "Unhandled parameter type: %d", type);
        hr = E_INVALIDARG;
        break;
    }
    
    VariantClear(&vTemp);
    if (SUCCEEDED(hr)) {
        DebugLog("DEBUG", "ConvertVariantToType", "Conversion successful, hr=0x%08x", hr);
    } else {
        DebugLog("ERROR", "ConvertVariantToType", "Conversion failed, hr=0x%08x", hr);
    }
    return hr;
}

// Convert type to VARIANT
HRESULT ConvertTypeToVariant(void* pData, ParameterType type, VARIANT* pVar)
{
    VariantInit(pVar);
    
    switch (type)
    {
    case TYPE_LONG:
        V_VT(pVar) = VT_I4;
        V_I4(pVar) = *(LONG*)pData;
        break;
        
    case TYPE_ULONG:
        V_VT(pVar) = VT_UI4;
        V_UI4(pVar) = *(ULONG*)pData;
        break;
        
    case TYPE_HANDLE:
    case TYPE_POINTER:
#ifdef _WIN64
        {
            ULONG_PTR ptr = (ULONG_PTR)*(void**)pData;
            DebugLog("DEBUG", "ConvertTypeToVariant", "Converting pointer: 0x%p", (void*)ptr);
            if (ptr > 0xFFFFFFFF) {
                V_VT(pVar) = VT_I8;
                V_I8(pVar) = (LONGLONG)ptr;
                DebugLog("DEBUG", "ConvertTypeToVariant", "Large pointer as VT_I8: %lld (0x%llx)", V_I8(pVar), (UINT64)V_I8(pVar));
            } else {
                V_VT(pVar) = VT_UI4;
                V_UI4(pVar) = (ULONG)ptr;
                DebugLog("DEBUG", "ConvertTypeToVariant", "Small pointer as VT_UI4: %u (0x%08x)", V_UI4(pVar), V_UI4(pVar));
            }
        }
#else
        V_VT(pVar) = VT_UI4;
        V_UI4(pVar) = (ULONG)*(void**)pData;
#endif
        break;
        
    case TYPE_LONGLONG:
        V_VT(pVar) = VT_I8;
        V_I8(pVar) = *(LONGLONG*)pData;
        break;
        
    case TYPE_SHORT:
        V_VT(pVar) = VT_I2;
        V_I2(pVar) = *(SHORT*)pData;
        break;
        
    case TYPE_USHORT:
        V_VT(pVar) = VT_UI2;
        V_UI2(pVar) = *(USHORT*)pData;
        break;
        
    case TYPE_CHAR:
        V_VT(pVar) = VT_I1;
        V_I1(pVar) = *(CHAR*)pData;
        break;
        
    case TYPE_UCHAR:
        V_VT(pVar) = VT_UI1;
        V_UI1(pVar) = *(UCHAR*)pData;
        break;
        
    case TYPE_FLOAT:
        V_VT(pVar) = VT_R4;
        V_R4(pVar) = *(FLOAT*)pData;
        break;
        
    case TYPE_DOUBLE:
        V_VT(pVar) = VT_R8;
        V_R8(pVar) = *(DOUBLE*)pData;
        break;
        
    case TYPE_WSTRING:
        V_VT(pVar) = VT_BSTR;
        V_BSTR(pVar) = SysAllocString((BSTR)pData);
        break;
        
    case TYPE_ASTRING:
    case TYPE_OSTRING:
        {
            V_VT(pVar) = VT_BSTR;
            int len = strlen((char*)pData);
            BSTR bstr = SysAllocStringLen(NULL, len);
            if (bstr)
            {
                UINT codePage = (type == TYPE_OSTRING) ? CP_OEMCP : CP_ACP;
                MultiByteToWideChar(codePage, 0, (char*)pData, len, bstr, len);
                V_BSTR(pVar) = bstr;
            }
            else
            {
                return E_OUTOFMEMORY;
            }
        }
        break;
        
    default:
        return E_INVALIDARG;
    }
    
    return S_OK;
}

// Memory cleanup
void CleanupMemoryBlocks(DynamicWrapperX* pObj)
{
    MemoryBlock* pBlock = pObj->memoryBlocks;
    while (pBlock)
    {
        MemoryBlock* pNext = pBlock->next;
        
        if (pBlock->isGlobal)
            GlobalFree(pBlock->ptr);
        else
            SysFreeString((BSTR)pBlock->ptr);
            
        GlobalFree(pBlock);
        pBlock = pNext;
    }
    pObj->memoryBlocks = NULL;
}
