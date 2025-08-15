#include "dynwrapx.h"

// Helper function to parse single type from variant
ParameterType ParseSingleType(VARIANT* pVarType)
{
    if (!pVarType || pVarType->vt == VT_EMPTY) return TYPE_LONG;
    
    if (pVarType->vt == VT_BSTR) {
        BSTR typeStr = V_BSTR(pVarType);
        if (!typeStr) return TYPE_LONG;
        
        // Support both short and full type names
        if (wcsicmp(typeStr, L"l") == 0 || wcsicmp(typeStr, L"long") == 0) return TYPE_LONG;
        if (wcsicmp(typeStr, L"u") == 0 || wcsicmp(typeStr, L"ulong") == 0 || wcsicmp(typeStr, L"uint") == 0) return TYPE_ULONG;
        if (wcsicmp(typeStr, L"q") == 0 || wcsicmp(typeStr, L"longlong") == 0 || wcsicmp(typeStr, L"int64") == 0) return TYPE_LONGLONG;
        if (wcsicmp(typeStr, L"p") == 0 || wcsicmp(typeStr, L"ptr") == 0 || wcsicmp(typeStr, L"pointer") == 0) return TYPE_POINTER;
        if (wcsicmp(typeStr, L"h") == 0 || wcsicmp(typeStr, L"handle") == 0) return TYPE_HANDLE;
        if (wcsicmp(typeStr, L"n") == 0 || wcsicmp(typeStr, L"short") == 0) return TYPE_SHORT;
        if (wcsicmp(typeStr, L"t") == 0 || wcsicmp(typeStr, L"ushort") == 0) return TYPE_USHORT;
        if (wcsicmp(typeStr, L"c") == 0 || wcsicmp(typeStr, L"char") == 0) return TYPE_CHAR;
        if (wcsicmp(typeStr, L"b") == 0 || wcsicmp(typeStr, L"uchar") == 0 || wcsicmp(typeStr, L"byte") == 0) return TYPE_UCHAR;
        if (wcsicmp(typeStr, L"f") == 0 || wcsicmp(typeStr, L"float") == 0) return TYPE_FLOAT;
        if (wcsicmp(typeStr, L"d") == 0 || wcsicmp(typeStr, L"double") == 0) return TYPE_DOUBLE;
    }
    
    return TYPE_LONG;  // Default
}

// Parse parameter string (e.g., "i=luu" or "r=l")
HRESULT ParseParameterString(LPCWSTR paramStr, ParameterType** types, int* count)
{
    if (!paramStr || !types || !count)
        return E_INVALIDARG;
    
    *types = NULL;
    *count = 0;
    
    // Find the '=' character
    LPCWSTR equalPos = wcschr(paramStr, L'=');
    if (!equalPos)
        return E_INVALIDARG;
    
    // Skip past the '='
    LPCWSTR typeStr = equalPos + 1;
    int len = wcslen(typeStr);
    
    if (len == 0)
        return S_OK; // Empty parameter list is valid
    
    *types = (ParameterType*)GlobalAlloc(GMEM_FIXED, len * sizeof(ParameterType));
    if (!*types)
        return E_OUTOFMEMORY;
    
    for (int i = 0; i < len; i++)
    {
        (*types)[i] = ParseParameterType(typeStr[i]);
    }
    
    *count = len;
    return S_OK;
}

// Register method implementation
HRESULT DynWrap_Register(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    BSTR libraryName = NULL;
    BSTR functionName = NULL;
    BSTR paramTypes = NULL;
    BSTR returnType = NULL;
    FunctionInfo* pFunc = NULL;
    
    // Validate parameters (minimum 2: library name and function name)
    if (pDispParams->cArgs < 2)
        return DISP_E_BADPARAMCOUNT;
    
    // Get library name (parameter 0, but arguments are in reverse order)
    VARIANT* pLibArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    if (V_VT(pLibArg) != VT_BSTR)
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        hr = VariantChangeType(&vTemp, pLibArg, 0, VT_BSTR);
        if (FAILED(hr))
            return hr;
        libraryName = V_BSTR(&vTemp);
        V_BSTR(&vTemp) = NULL;
    }
    else
    {
        libraryName = SysAllocString(V_BSTR(pLibArg));
    }
    
    if (!libraryName)
        return E_OUTOFMEMORY;
    
    // Get function name (parameter 1)
    VARIANT* pFuncArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
    if (V_VT(pFuncArg) != VT_BSTR)
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        hr = VariantChangeType(&vTemp, pFuncArg, 0, VT_BSTR);
        if (FAILED(hr))
            goto cleanup;
        functionName = V_BSTR(&vTemp);
        V_BSTR(&vTemp) = NULL;
    }
    else
    {
        functionName = SysAllocString(V_BSTR(pFuncArg));
    }
    
    if (!functionName)
    {
        hr = E_OUTOFMEMORY;
        goto cleanup;
    }
    
    // Get parameter types (parameter 2, optional)
    if (pDispParams->cArgs >= 3)
    {
        VARIANT* pParamArg = &pDispParams->rgvarg[pDispParams->cArgs - 3];
        if (V_VT(pParamArg) == VT_BSTR)
        {
            paramTypes = SysAllocString(V_BSTR(pParamArg));
        }
        else if (V_VT(pParamArg) != VT_EMPTY && V_VT(pParamArg) != VT_NULL)
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pParamArg, 0, VT_BSTR);
            if (SUCCEEDED(hr))
            {
                paramTypes = V_BSTR(&vTemp);
                V_BSTR(&vTemp) = NULL;
            }
        }
    }
    
    // Get return type (parameter 3, optional)
    if (pDispParams->cArgs >= 4)
    {
        VARIANT* pRetArg = &pDispParams->rgvarg[pDispParams->cArgs - 4];
        if (V_VT(pRetArg) == VT_BSTR)
        {
            returnType = SysAllocString(V_BSTR(pRetArg));
        }
        else if (V_VT(pRetArg) != VT_EMPTY && V_VT(pRetArg) != VT_NULL)
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pRetArg, 0, VT_BSTR);
            if (SUCCEEDED(hr))
            {
                returnType = V_BSTR(&vTemp);
                V_BSTR(&vTemp) = NULL;
            }
        }
    }
    
    // Load the function
    FARPROC proc = LoadFunction(libraryName, functionName);
    if (!proc)
    {
        hr = E_FAIL;
        goto cleanup;
    }
    
    // Create function info structure
    pFunc = (FunctionInfo*)GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof(FunctionInfo));
    if (!pFunc)
    {
        hr = E_OUTOFMEMORY;
        goto cleanup;
    }
    
    pFunc->functionName = SysAllocString(functionName);
    pFunc->functionPtr = proc;
    pFunc->returnType = TYPE_LONG; // Default return type
    
    // Parse parameter types and return type
    if (paramTypes)
    {
        LPCWSTR equalPos = wcschr(paramTypes, L'=');
        if (equalPos && equalPos > paramTypes)
        {
            // Combined format like "p=puuu" - return type before '=', parameters after
            pFunc->returnType = ParseParameterType(paramTypes[0]);
            DebugLog("DEBUG", "DynWrap_Register", "Parsed return type '%c' as %d from combined signature", paramTypes[0], pFunc->returnType);
            
            // Parse parameters after '=' sign
            LPCWSTR paramStr = equalPos + 1;
            if (wcslen(paramStr) > 0) {
                int len = wcslen(paramStr);
                pFunc->paramTypes = (ParameterType*)GlobalAlloc(GMEM_FIXED, len * sizeof(ParameterType));
                if (!pFunc->paramTypes) {
                    hr = E_OUTOFMEMORY;
                    goto cleanup;
                }
                
                for (int i = 0; i < len; i++) {
                    pFunc->paramTypes[i] = ParseParameterType(paramStr[i]);
                }
                pFunc->paramCount = len;
            }
        }
        else
        {
            // Just parameter types, no return type specified
            hr = ParseParameterString(paramTypes, &pFunc->paramTypes, &pFunc->paramCount);
            if (FAILED(hr))
                goto cleanup;
        }
    }
    
    // Parse separate return type (from the 4th argument, if provided)
    if (returnType)
    {
        LPCWSTR equalPos = wcschr(returnType, L'=');
        if (equalPos && equalPos > returnType)
        {
            // Return type is the character before the '=' sign
            pFunc->returnType = ParseParameterType(returnType[0]);
            DebugLog("DEBUG", "DynWrap_Register", "Parsed return type '%c' as %d from separate argument", returnType[0], pFunc->returnType);
        }
        else if (wcslen(returnType) > 0)
        {
            // No '=' sign, entire string is return type
            pFunc->returnType = ParseParameterType(returnType[0]);
            DebugLog("DEBUG", "DynWrap_Register", "Parsed return type '%c' as %d from separate argument", returnType[0], pFunc->returnType);
        }
    }
    
    // Assign function ID
    EnterCriticalSection(&pObj->cs);
    
    // Find a unique function ID
    DWORD funcId = 2000; // Start from 2000 to avoid conflicts with built-in methods
    FunctionInfo* pExisting = pObj->functionList;
    while (pExisting)
    {
        if (pExisting->functionId >= funcId)
            funcId = pExisting->functionId + 1;
        pExisting = pExisting->next;
    }
    
    pFunc->functionId = funcId;
    
    // Add to function list
    pFunc->next = pObj->functionList;
    pObj->functionList = pFunc;
    
    LeaveCriticalSection(&pObj->cs);
    
    // Return success
    if (pVarResult)
    {
        VariantInit(pVarResult);
        V_VT(pVarResult) = VT_I4;
        V_I4(pVarResult) = 0; // Success
    }
    
    pFunc = NULL; // Don't cleanup on success
    
cleanup:
    SAFE_SYSFREE(libraryName);
    SAFE_SYSFREE(functionName);
    SAFE_SYSFREE(paramTypes);
    SAFE_SYSFREE(returnType);
    
    if (pFunc)
    {
        SAFE_SYSFREE(pFunc->functionName);
        if (pFunc->paramTypes)
            GlobalFree(pFunc->paramTypes);
        GlobalFree(pFunc);
    }
    
    return hr;
}

// RegisterCallback method implementation
HRESULT DynWrap_RegisterCallback(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    IDispatch* pCallback = NULL;
    BSTR paramTypes = NULL;
    BSTR returnType = NULL;
    int callbackIndex = -1;
    
    // Validate parameters (minimum 1: callback function)
    if (pDispParams->cArgs < 1)
        return DISP_E_BADPARAMCOUNT;
    
    // Get callback function (parameter 0)
    VARIANT* pCallbackArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    if (V_VT(pCallbackArg) != VT_DISPATCH)
        return E_INVALIDARG;
    
    pCallback = V_DISPATCH(pCallbackArg);
    if (!pCallback)
        return E_INVALIDARG;
    
    // Get parameter types (parameter 1, optional)
    if (pDispParams->cArgs >= 2)
    {
        VARIANT* pParamArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
        if (V_VT(pParamArg) == VT_BSTR)
        {
            paramTypes = SysAllocString(V_BSTR(pParamArg));
        }
    }
    
    // Get return type (parameter 2, optional)
    if (pDispParams->cArgs >= 3)
    {
        VARIANT* pRetArg = &pDispParams->rgvarg[pDispParams->cArgs - 3];
        if (V_VT(pRetArg) == VT_BSTR)
        {
            returnType = SysAllocString(V_BSTR(pRetArg));
        }
    }
    
    // Find an available callback slot
    EnterCriticalSection(&pObj->cs);
    
    for (int i = 0; i < 16; i++)
    {
        if (pObj->callbacks[i].scriptFunction == NULL)
        {
            callbackIndex = i;
            break;
        }
    }
    
    if (callbackIndex == -1)
    {
        LeaveCriticalSection(&pObj->cs);
        hr = E_FAIL; // No available callback slots
        goto cleanup;
    }
    
    // Initialize callback info
    CallbackInfo* pInfo = &pObj->callbacks[callbackIndex];
    pInfo->scriptFunction = pCallback;
    pCallback->lpVtbl->AddRef(pCallback);
    pInfo->callbackIndex = callbackIndex;
    pInfo->returnType = TYPE_LONG; // Default return type
    
    // Parse parameter types
    if (paramTypes)
    {
        hr = ParseParameterString(paramTypes, &pInfo->paramTypes, &pInfo->paramCount);
        if (FAILED(hr))
        {
            SAFE_RELEASE(pInfo->scriptFunction);
            pInfo->callbackIndex = -1;
            LeaveCriticalSection(&pObj->cs);
            goto cleanup;
        }
    }
    
    // Parse return type (from the part BEFORE the '=' sign)
    if (returnType)
    {
        LPCWSTR equalPos = wcschr(returnType, L'=');
        if (equalPos && equalPos > returnType)
        {
            // Return type is the character before the '=' sign
            pInfo->returnType = ParseParameterType(returnType[0]);
            DebugLog("DEBUG", "DynWrap_RegisterCallback", "Parsed return type '%c' as %d", returnType[0], pInfo->returnType);
        }
        else if (wcslen(returnType) > 0)
        {
            // No '=' sign, entire string is return type
            pInfo->returnType = ParseParameterType(returnType[0]);
            DebugLog("DEBUG", "DynWrap_RegisterCallback", "Parsed return type '%c' as %d", returnType[0], pInfo->returnType);
        }
    }
    
    // Store object reference for callback
    g_callbackObjects[callbackIndex] = pObj;
    
    LeaveCriticalSection(&pObj->cs);
    
    // Return callback pointer
    if (pVarResult)
    {
        VariantInit(pVarResult);
        V_VT(pVarResult) = VT_UI4;
        
        // Return pointer to appropriate callback stub
        switch (callbackIndex)
        {
        case 0: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub0; break;
        case 1: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub1; break;
        case 2: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub2; break;
        case 3: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub3; break;
        case 4: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub4; break;
        case 5: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub5; break;
        case 6: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub6; break;
        case 7: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub7; break;
        case 8: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub8; break;
        case 9: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub9; break;
        case 10: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub10; break;
        case 11: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub11; break;
        case 12: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub12; break;
        case 13: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub13; break;
        case 14: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub14; break;
        case 15: V_UI4(pVarResult) = (ULONG)(ULONG_PTR)CallbackStub15; break;
        }
    }
    
cleanup:
    SAFE_SYSFREE(paramTypes);
    SAFE_SYSFREE(returnType);
    
    return hr;
}

// NumGet method implementation
HRESULT DynWrap_NumGet(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    void* address = NULL;
    LONG offset = 0;
    ParameterType type = TYPE_LONG;
    
    DebugLog("DEBUG", "DynWrap_NumGet", "NumGet called with %d parameters", pDispParams->cArgs);
    
    // Validate parameters (minimum 1: address)
    if (pDispParams->cArgs < 1)
        return DISP_E_BADPARAMCOUNT;
    
    // Get address (parameter 0)
    VARIANT* pAddrArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    if (V_VT(pAddrArg) == VT_BSTR)
    {
        // String pointer
        address = V_BSTR(pAddrArg);
    }
    else
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        
        // Handle both 32-bit and 64-bit addresses properly
#ifdef _WIN64
        // On 64-bit, try multiple conversion methods (same as NumPut)
        DebugLog("DEBUG", "DynWrap_NumGet", "Original VT: %d", V_VT(pAddrArg));
        
        // Try VT_R8 (double) first for large JavaScript numbers
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_R8);
        if (SUCCEEDED(hr)) {
            // Convert double to ULONG_PTR, avoiding sign extension
            double dblValue = V_R8(&vTemp);
            address = (void*)(ULONG_PTR)(UINT64)dblValue;
            DebugLog("DEBUG", "DynWrap_NumGet", "VT_R8 conversion successful, double=%.0f, address=0x%p", dblValue, address);
        }
        else 
        {
            DebugLog("DEBUG", "DynWrap_NumGet", "VT_R8 failed, trying VT_I8...");
            hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_I8);
            if (SUCCEEDED(hr)) {
                // Treat as unsigned to avoid sign extension issues
                UINT64 unsignedAddr = (UINT64)V_I8(&vTemp);
                address = (void*)unsignedAddr;
                DebugLog("DEBUG", "DynWrap_NumGet", "VT_I8 conversion successful, signed=%lld, unsigned=0x%llx, address=0x%p", V_I8(&vTemp), unsignedAddr, address);
            }
            else
            {
                DebugLog("DEBUG", "DynWrap_NumGet", "VT_I8 failed, trying VT_UI4...");
                hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
                if (SUCCEEDED(hr)) {
                    address = (void*)(ULONG_PTR)V_UI4(&vTemp);
                    DebugLog("DEBUG", "DynWrap_NumGet", "VT_UI4 conversion successful, address=0x%p", address);
                }
            }
        }
#else
        // On 32-bit, use VT_UI4
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
        if (SUCCEEDED(hr))
            address = (void*)(ULONG_PTR)V_UI4(&vTemp);
#endif
        
        if (FAILED(hr))
            return hr;
    }
    
    // Get offset (parameter 1, optional)
    if (pDispParams->cArgs >= 2)
    {
        VARIANT* pOffsetArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
        VARIANT vTemp;
        VariantInit(&vTemp);
        hr = VariantChangeType(&vTemp, pOffsetArg, 0, VT_I4);
        if (SUCCEEDED(hr))
            offset = V_I4(&vTemp);
    }
    
    // Get type (parameter 2, optional)
    if (pDispParams->cArgs >= 3)
    {
        VARIANT* pTypeArg = &pDispParams->rgvarg[pDispParams->cArgs - 3];
        if (V_VT(pTypeArg) == VT_BSTR && SysStringLen(V_BSTR(pTypeArg)) > 0)
        {
            BSTR typeStr = V_BSTR(pTypeArg);
            int len = SysStringLen(typeStr);
            
            // Handle full type names (case-insensitive)
            if (len > 1)
            {
                if (_wcsicmp(typeStr, L"UInt") == 0 || _wcsicmp(typeStr, L"ULong") == 0)
                    type = TYPE_ULONG;
                else if (_wcsicmp(typeStr, L"Int") == 0 || _wcsicmp(typeStr, L"Long") == 0)
                    type = TYPE_LONG;
                else if (_wcsicmp(typeStr, L"UShort") == 0)
                    type = TYPE_USHORT;
                else if (_wcsicmp(typeStr, L"Short") == 0)
                    type = TYPE_SHORT;
                else if (_wcsicmp(typeStr, L"UChar") == 0)
                    type = TYPE_UCHAR;
                else if (_wcsicmp(typeStr, L"Char") == 0)
                    type = TYPE_CHAR;
                else if (_wcsicmp(typeStr, L"Float") == 0)
                    type = TYPE_FLOAT;
                else if (_wcsicmp(typeStr, L"Double") == 0)
                    type = TYPE_DOUBLE;
                else if (_wcsicmp(typeStr, L"Ptr") == 0 || _wcsicmp(typeStr, L"Pointer") == 0)
                    type = TYPE_POINTER;
                else
                    type = ParseParameterType(typeStr[0]);
            }
            else
            {
                type = ParseParameterType(typeStr[0]);
            }
        }
    }
    
    // Calculate final address
    void* finalAddr = (BYTE*)address + offset;
    
    DebugLog("DEBUG", "DynWrap_NumGet", "Reading from address=0x%p, offset=%ld, type=%d, finalAddr=0x%p", address, offset, type, finalAddr);
    
    // Basic null pointer check
    if (!finalAddr)
    {
        DebugLog("ERROR", "DynWrap_NumGet", "Final address is NULL");
        return E_POINTER;
    }
    
    // Read value and return
    if (pVarResult)
    {
        VariantInit(pVarResult);
        
        switch (type)
        {
            case TYPE_CHAR:
                V_VT(pVarResult) = VT_I1;
                V_I1(pVarResult) = *(CHAR*)finalAddr;
                break;
            case TYPE_UCHAR:
                V_VT(pVarResult) = VT_UI1;
                V_UI1(pVarResult) = *(UCHAR*)finalAddr;
                break;
            case TYPE_SHORT:
                V_VT(pVarResult) = VT_I2;
                V_I2(pVarResult) = *(SHORT*)finalAddr;
                break;
            case TYPE_USHORT:
                {
                    USHORT value = *(USHORT*)finalAddr;
                    DebugLog("DEBUG", "DynWrap_NumGet", "Reading USHORT from 0x%p, value=%u (0x%04x)", finalAddr, value, value);
                    V_VT(pVarResult) = VT_UI2;
                    V_UI2(pVarResult) = value;
                }
                break;
            case TYPE_LONG:
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = *(LONG*)finalAddr;
                break;
            case TYPE_HANDLE:
            case TYPE_POINTER:
                {
#ifdef _WIN64
                    ULONG_PTR ptrValue = *(ULONG_PTR*)finalAddr;
                    if (ptrValue > 0xFFFFFFFFULL) {
                        V_VT(pVarResult) = VT_I8;
                        V_I8(pVarResult) = (LONGLONG)ptrValue;
                    } else {
                        V_VT(pVarResult) = VT_I4;
                        V_I4(pVarResult) = (LONG)ptrValue;
                    }
#else
                    V_VT(pVarResult) = VT_I4;
                    V_I4(pVarResult) = *(LONG*)finalAddr;
#endif
                }
                break;
            case TYPE_ULONG:
                V_VT(pVarResult) = VT_UI4;
                V_UI4(pVarResult) = *(ULONG*)finalAddr;
                break;
            case TYPE_LONGLONG:
                V_VT(pVarResult) = VT_I8;
                V_I8(pVarResult) = *(LONGLONG*)finalAddr;
                break;
            case TYPE_FLOAT:
                V_VT(pVarResult) = VT_R4;
                V_R4(pVarResult) = *(FLOAT*)finalAddr;
                break;
            case TYPE_DOUBLE:
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = *(DOUBLE*)finalAddr;
                break;
            default:
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = *(LONG*)finalAddr;
                break;
        }
    }
    
    return S_OK;
}


// NumPut method implementation
HRESULT DynWrap_NumPut(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    void* address = NULL;
    LONG offset = 0;
    ParameterType type = TYPE_LONG;
    
    DebugLog("DEBUG", "DynWrap_NumPut", "NumPut called with %d arguments", pDispParams->cArgs);
    
    // Validate parameters (minimum 2: value and address)
    if (pDispParams->cArgs < 2) {
        DebugLog("ERROR", "DynWrap_NumPut", "Insufficient parameters: %d", pDispParams->cArgs);
        return DISP_E_BADPARAMCOUNT;
    }
    
    // Get value (parameter 0)
    VARIANT* pValueArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    
    // Get address (parameter 1)
    VARIANT* pAddrArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
    DebugLog("DEBUG", "DynWrap_NumPut", "Address parameter VT: %d", V_VT(pAddrArg));
    
    if (V_VT(pAddrArg) == VT_BSTR)
    {
        address = V_BSTR(pAddrArg);
        DebugLog("DEBUG", "DynWrap_NumPut", "Using BSTR address");
    }
    else
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        
        DebugLog("DEBUG", "DynWrap_NumPut", "Converting address parameter...");
        
        // Handle both 32-bit and 64-bit addresses properly
#ifdef _WIN64
        // On 64-bit, try multiple conversion methods
        DebugLog("DEBUG", "DynWrap_NumPut", "Original VT: %d", V_VT(pAddrArg));
        
        // Try VT_R8 (double) first for large JavaScript numbers
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_R8);
        if (SUCCEEDED(hr)) {
            // Convert double to ULONG_PTR, avoiding sign extension
            double dblValue = V_R8(&vTemp);
            address = (void*)(ULONG_PTR)(UINT64)dblValue;
            DebugLog("DEBUG", "DynWrap_NumPut", "VT_R8 conversion successful, double=%.0f, address=0x%p", dblValue, address);
        }
        else 
        {
            DebugLog("DEBUG", "DynWrap_NumPut", "VT_R8 failed, trying VT_I8...");
            hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_I8);
            if (SUCCEEDED(hr)) {
                // Treat as unsigned to avoid sign extension issues
                UINT64 unsignedAddr = (UINT64)V_I8(&vTemp);
                address = (void*)unsignedAddr;
                DebugLog("DEBUG", "DynWrap_NumPut", "VT_I8 conversion successful, signed=%lld, unsigned=0x%llx, address=0x%p", V_I8(&vTemp), unsignedAddr, address);
            }
            else
            {
                DebugLog("DEBUG", "DynWrap_NumPut", "VT_I8 failed, trying VT_UI4...");
                hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
                if (SUCCEEDED(hr)) {
                    address = (void*)(ULONG_PTR)V_UI4(&vTemp);
                    DebugLog("DEBUG", "DynWrap_NumPut", "VT_UI4 conversion successful, address=0x%p", address);
                }
            }
        }
#else
        // On 32-bit, use VT_UI4
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
        if (SUCCEEDED(hr))
            address = (void*)(ULONG_PTR)V_UI4(&vTemp);
#endif
        
        if (FAILED(hr)) {
            DebugLog("ERROR", "DynWrap_NumPut", "Address conversion failed, hr=0x%08x", hr);
            return hr;
        }
    }
    
    // Get offset (parameter 2, optional)
    if (pDispParams->cArgs >= 3)
    {
        VARIANT* pOffsetArg = &pDispParams->rgvarg[pDispParams->cArgs - 3];
        VARIANT vTemp;
        VariantInit(&vTemp);
        hr = VariantChangeType(&vTemp, pOffsetArg, 0, VT_I4);
        if (SUCCEEDED(hr))
            offset = V_I4(&vTemp);
    }
    
    // Get type (parameter 3, optional)
    if (pDispParams->cArgs >= 4)
    {
        VARIANT* pTypeArg = &pDispParams->rgvarg[pDispParams->cArgs - 4];
        if (V_VT(pTypeArg) == VT_BSTR && SysStringLen(V_BSTR(pTypeArg)) > 0)
        {
            BSTR typeStr = V_BSTR(pTypeArg);
            int len = SysStringLen(typeStr);
            
            // Handle full type names (case-insensitive)
            if (len > 1)
            {
                if (_wcsicmp(typeStr, L"UInt") == 0 || _wcsicmp(typeStr, L"ULong") == 0)
                    type = TYPE_ULONG;
                else if (_wcsicmp(typeStr, L"Int") == 0 || _wcsicmp(typeStr, L"Long") == 0)
                    type = TYPE_LONG;
                else if (_wcsicmp(typeStr, L"UShort") == 0)
                    type = TYPE_USHORT;
                else if (_wcsicmp(typeStr, L"Short") == 0)
                    type = TYPE_SHORT;
                else if (_wcsicmp(typeStr, L"UChar") == 0)
                    type = TYPE_UCHAR;
                else if (_wcsicmp(typeStr, L"Char") == 0)
                    type = TYPE_CHAR;
                else if (_wcsicmp(typeStr, L"Float") == 0)
                    type = TYPE_FLOAT;
                else if (_wcsicmp(typeStr, L"Double") == 0)
                    type = TYPE_DOUBLE;
                else if (_wcsicmp(typeStr, L"Ptr") == 0 || _wcsicmp(typeStr, L"Pointer") == 0)
                    type = TYPE_POINTER;
                else
                    type = ParseParameterType(typeStr[0]);
            }
            else
            {
                type = ParseParameterType(typeStr[0]);
            }
        }
    }
    
    // Calculate final address
    void* finalAddr = (BYTE*)address + offset;
    
    DebugLog("DEBUG", "DynWrap_NumPut", "Final address calculated: base=0x%p, offset=%d, final=0x%p", address, offset, finalAddr);
    
    // Basic null pointer check
    if (!finalAddr)
    {
        DebugLog("ERROR", "DynWrap_NumPut", "Null final address");
        return E_POINTER;
    }
    
    DebugLog("DEBUG", "DynWrap_NumPut", "About to write value of type %d", type);
    
    // Write value with type conversion
    switch (type)
    {
        case TYPE_CHAR:
            {
                VARIANT vTemp;
                VariantInit(&vTemp);
                hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I1);
                if (SUCCEEDED(hr))
                    *(CHAR*)finalAddr = V_I1(&vTemp);
            }
            break;
    case TYPE_UCHAR:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_UI1);
            if (SUCCEEDED(hr))
                *(UCHAR*)finalAddr = V_UI1(&vTemp);
        }
        break;
    case TYPE_SHORT:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I2);
            if (SUCCEEDED(hr))
                *(SHORT*)finalAddr = V_I2(&vTemp);
        }
        break;
    case TYPE_USHORT:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_UI2);
            if (SUCCEEDED(hr))
                *(USHORT*)finalAddr = V_UI2(&vTemp);
        }
        break;
    case TYPE_LONG:
    case TYPE_HANDLE:
    case TYPE_POINTER:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            
#ifdef _WIN64
            // On 64-bit, try VT_I8 first for large values, fallback to VT_I4
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I8);
            if (SUCCEEDED(hr))
            {
                *(ULONG_PTR*)finalAddr = (ULONG_PTR)V_I8(&vTemp);
            }
            else
            {
                hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I4);
                if (SUCCEEDED(hr))
                    *(ULONG_PTR*)finalAddr = (ULONG_PTR)V_I4(&vTemp);
            }
#else
            // On 32-bit, use VT_I4
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I4);
            if (SUCCEEDED(hr))
                *(LONG*)finalAddr = V_I4(&vTemp);
#endif
        }
        break;
    case TYPE_ULONG:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            DebugLog("DEBUG", "DynWrap_NumPut", "Writing ULONG value...");
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_UI4);
            if (SUCCEEDED(hr)) {
                ULONG valueToWrite = V_UI4(&vTemp);
                DebugLog("DEBUG", "DynWrap_NumPut", "About to write ULONG %u to address 0x%p", valueToWrite, finalAddr);
                
                // Check if memory is writable using VirtualQuery
                MEMORY_BASIC_INFORMATION mbi;
                SIZE_T result = VirtualQuery(finalAddr, &mbi, sizeof(mbi));
                if (result == sizeof(mbi)) {
                    DebugLog("DEBUG", "DynWrap_NumPut", "Memory state: 0x%08x, protect: 0x%08x, type: 0x%08x", 
                             mbi.State, mbi.Protect, mbi.Type);
                    
                    if (mbi.State == MEM_COMMIT && (mbi.Protect & PAGE_READWRITE || mbi.Protect & PAGE_EXECUTE_READWRITE)) {
                        *(ULONG*)finalAddr = valueToWrite;
                        DebugLog("DEBUG", "DynWrap_NumPut", "ULONG value written successfully: %u", valueToWrite);
                    } else {
                        DebugLog("ERROR", "DynWrap_NumPut", "Memory not writable: state=0x%08x, protect=0x%08x", mbi.State, mbi.Protect);
                        hr = E_ACCESSDENIED;
                    }
                } else {
                    DebugLog("ERROR", "DynWrap_NumPut", "VirtualQuery failed for address 0x%p", finalAddr);
                    // Try the write anyway
                    *(ULONG*)finalAddr = valueToWrite;
                    DebugLog("DEBUG", "DynWrap_NumPut", "ULONG value written (VirtualQuery failed): %u", valueToWrite);
                }
            } else {
                DebugLog("ERROR", "DynWrap_NumPut", "ULONG conversion failed, hr=0x%08x", hr);
            }
        }
        break;
    case TYPE_FLOAT:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_R4);
            if (SUCCEEDED(hr))
                *(FLOAT*)finalAddr = V_R4(&vTemp);
        }
        break;
    case TYPE_DOUBLE:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_R8);
            if (SUCCEEDED(hr))
                *(DOUBLE*)finalAddr = V_R8(&vTemp);
        }
        break;
    default:
        {
            VARIANT vTemp;
            VariantInit(&vTemp);
            hr = VariantChangeType(&vTemp, pValueArg, 0, VT_I4);
            if (SUCCEEDED(hr))
                *(LONG*)finalAddr = V_I4(&vTemp);
        }
        break;
    }
    
    // Return address after written data
    if (pVarResult)
    {
        VariantInit(pVarResult);
        ULONG_PTR resultAddr = (ULONG_PTR)((BYTE*)finalAddr + GetTypeSize(type));
        
#ifdef _WIN64
        // On 64-bit, return as VT_I8 if the address is large
        if (resultAddr > 0xFFFFFFFF)
        {
            V_VT(pVarResult) = VT_I8;
            V_I8(pVarResult) = (LONGLONG)resultAddr;
        }
        else
        {
            V_VT(pVarResult) = VT_UI4;
            V_UI4(pVarResult) = (ULONG)resultAddr;
        }
#else
        // On 32-bit, always use VT_UI4
        V_VT(pVarResult) = VT_UI4;
        V_UI4(pVarResult) = (ULONG)resultAddr;
#endif
    }
    
    DebugLog("DEBUG", "DynWrap_NumPut", "NumPut completed successfully, hr=0x%08x", hr);
    return hr;
}

// StrPtr method implementation  
HRESULT DynWrap_StrPtr(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    BSTR inputStr = NULL;
    ParameterType type = TYPE_WSTRING;
    
    // Validate parameters (minimum 1: string)
    if (pDispParams->cArgs < 1)
        return DISP_E_BADPARAMCOUNT;
    
    // Get string (parameter 0)
    VARIANT* pStrArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    if (V_VT(pStrArg) != VT_BSTR)
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        hr = VariantChangeType(&vTemp, pStrArg, 0, VT_BSTR);
        if (FAILED(hr))
            return hr;
        inputStr = V_BSTR(&vTemp);
        V_BSTR(&vTemp) = NULL;
    }
    else
    {
        inputStr = SysAllocString(V_BSTR(pStrArg));
    }
    
    if (!inputStr)
        return E_OUTOFMEMORY;
    
    // Get type (parameter 1, optional)
    if (pDispParams->cArgs >= 2)
    {
        VARIANT* pTypeArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
        if (V_VT(pTypeArg) == VT_BSTR && SysStringLen(V_BSTR(pTypeArg)) > 0)
        {
            WCHAR typeChar = V_BSTR(pTypeArg)[0];
            if (typeChar == L'w')
                type = TYPE_WSTRING;
            else if (typeChar == L's')
                type = TYPE_ASTRING;
            else if (typeChar == L'z')
                type = TYPE_OSTRING;
        }
    }
    
    void* resultPtr = NULL;
    
    if (type == TYPE_WSTRING)
    {
        // Return pointer to Unicode string
        resultPtr = inputStr;
    }
    else
    {
        // Convert to ANSI or OEM
        int len = SysStringLen(inputStr);
        char* pszConverted = (char*)GlobalAlloc(GMEM_FIXED, len + 1);
        if (pszConverted)
        {
            UINT codePage = (type == TYPE_OSTRING) ? CP_OEMCP : CP_ACP;
            WideCharToMultiByte(codePage, 0, inputStr, len, pszConverted, len + 1, NULL, NULL);
            pszConverted[len] = '\0';
            resultPtr = pszConverted;
            
            // Track memory allocation
            MemoryBlock* pBlock = (MemoryBlock*)GlobalAlloc(GMEM_FIXED, sizeof(MemoryBlock));
            if (pBlock)
            {
                pBlock->ptr = pszConverted;
                pBlock->isGlobal = TRUE;
                pBlock->next = pObj->memoryBlocks;
                pObj->memoryBlocks = pBlock;
            }
        }
        else
        {
            SAFE_SYSFREE(inputStr);
            return E_OUTOFMEMORY;
        }
    }
    
    // Return pointer
    if (pVarResult)
    {
        VariantInit(pVarResult);
        ULONG_PTR ptrValue = (ULONG_PTR)resultPtr;
        
#ifdef _WIN64
        // On 64-bit, return large pointers as VT_I8
        if (ptrValue > 0xFFFFFFFF)
        {
            V_VT(pVarResult) = VT_I8;
            V_I8(pVarResult) = (LONGLONG)ptrValue;
        }
        else
        {
            V_VT(pVarResult) = VT_UI4;
            V_UI4(pVarResult) = (ULONG)ptrValue;
        }
#else
        // On 32-bit, always use VT_UI4
        V_VT(pVarResult) = VT_UI4;
        V_UI4(pVarResult) = (ULONG)ptrValue;
#endif
    }
    
    if (type != TYPE_WSTRING)
    {
        SAFE_SYSFREE(inputStr);
    }
    
    return S_OK;
}

// StrGet method implementation
HRESULT DynWrap_StrGet(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    void* address = NULL;
    ParameterType type = TYPE_WSTRING;
    
    // Validate parameters (minimum 1: address)
    if (pDispParams->cArgs < 1)
        return DISP_E_BADPARAMCOUNT;
    
    // Get address (parameter 0)
    VARIANT* pAddrArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    if (V_VT(pAddrArg) == VT_BSTR)
    {
        address = V_BSTR(pAddrArg);
    }
    else
    {
        VARIANT vTemp;
        VariantInit(&vTemp);
        
        // Handle both 32-bit and 64-bit addresses properly
#ifdef _WIN64
        // On 64-bit, try VT_I8 first, fallback to VT_UI4
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_I8);
        if (SUCCEEDED(hr))
            address = (void*)(ULONG_PTR)V_I8(&vTemp);
        else
        {
            hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
            if (SUCCEEDED(hr))
                address = (void*)(ULONG_PTR)V_UI4(&vTemp);
        }
#else
        // On 32-bit, use VT_UI4
        hr = VariantChangeType(&vTemp, pAddrArg, 0, VT_UI4);
        if (SUCCEEDED(hr))
            address = (void*)(ULONG_PTR)V_UI4(&vTemp);
#endif
        
        if (FAILED(hr))
            return hr;
    }
    
    // Get type (parameter 1, optional)
    if (pDispParams->cArgs >= 2)
    {
        VARIANT* pTypeArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
        if (V_VT(pTypeArg) == VT_BSTR && SysStringLen(V_BSTR(pTypeArg)) > 0)
        {
            WCHAR typeChar = V_BSTR(pTypeArg)[0];
            if (typeChar == L'w')
                type = TYPE_WSTRING;
            else if (typeChar == L's')
                type = TYPE_ASTRING;
            else if (typeChar == L'z')
                type = TYPE_OSTRING;
        }
    }
    
    BSTR resultStr = NULL;
    
    if (type == TYPE_WSTRING)
    {
        // Read Unicode string
        WCHAR* pwszStr = (WCHAR*)address;
        if (pwszStr)
        {
            resultStr = SysAllocString(pwszStr);
        }
    }
    else
    {
        // Read ANSI/OEM string and convert to Unicode
        char* pszStr = (char*)address;
        if (pszStr)
        {
            int len = strlen(pszStr);
            resultStr = SysAllocStringLen(NULL, len);
            if (resultStr)
            {
                UINT codePage = (type == TYPE_OSTRING) ? CP_OEMCP : CP_ACP;
                MultiByteToWideChar(codePage, 0, pszStr, len, resultStr, len);
            }
        }
    }
    
    // Return string
    if (pVarResult)
    {
        VariantInit(pVarResult);
        V_VT(pVarResult) = VT_BSTR;
        V_BSTR(pVarResult) = resultStr ? resultStr : SysAllocString(L"");
    }
    else if (resultStr)
    {
        SysFreeString(resultStr);
    }
    
    return S_OK;
}

// Space method implementation
HRESULT DynWrap_Space(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    UINT count = 0;
    WCHAR fillChar = L' ';
    
    // Validate parameters (minimum 1: count)
    if (pDispParams->cArgs < 1)
        return DISP_E_BADPARAMCOUNT;
    
    // Get count (parameter 0)
    VARIANT* pCountArg = &pDispParams->rgvarg[pDispParams->cArgs - 1];
    VARIANT vTemp;
    VariantInit(&vTemp);
    hr = VariantChangeType(&vTemp, pCountArg, 0, VT_UI4);
    if (FAILED(hr))
        return hr;
    count = V_UI4(&vTemp);
    
    // Get fill character (parameter 1, optional)
    if (pDispParams->cArgs >= 2)
    {
        VARIANT* pCharArg = &pDispParams->rgvarg[pDispParams->cArgs - 2];
        if (V_VT(pCharArg) == VT_BSTR)
        {
            BSTR charStr = V_BSTR(pCharArg);
            if (SysStringLen(charStr) > 0)
            {
                fillChar = charStr[0];
            }
            else
            {
                fillChar = L'\0'; // Empty string means null character
            }
        }
        else if (V_VT(pCharArg) != VT_EMPTY && V_VT(pCharArg) != VT_NULL)
        {
            VARIANT vCharTemp;
            VariantInit(&vCharTemp);
            hr = VariantChangeType(&vCharTemp, pCharArg, 0, VT_BSTR);
            if (SUCCEEDED(hr) && SysStringLen(V_BSTR(&vCharTemp)) > 0)
            {
                fillChar = V_BSTR(&vCharTemp)[0];
            }
            VariantClear(&vCharTemp);
        }
    }
    
    // Create string
    BSTR resultStr = SysAllocStringLen(NULL, count);
    if (!resultStr)
        return E_OUTOFMEMORY;
    
    // Fill with character
    for (UINT i = 0; i < count; i++)
    {
        resultStr[i] = fillChar;
    }
    
    // Return string
    if (pVarResult)
    {
        VariantInit(pVarResult);
        V_VT(pVarResult) = VT_BSTR;
        V_BSTR(pVarResult) = resultStr;
    }
    else
    {
        SysFreeString(resultStr);
    }
    
    return S_OK;
}

// Callback implementation helper
HRESULT CallScriptFunction(DynamicWrapperX* pObj, int callbackIndex, void** args, void* returnValue)
{
    if (callbackIndex < 0 || callbackIndex >= 16)
        return E_INVALIDARG;
    
    CallbackInfo* pInfo = &pObj->callbacks[callbackIndex];
    if (!pInfo->scriptFunction)
        return E_FAIL;
    
    HRESULT hr = S_OK;
    DISPPARAMS dispParams = {0};
    VARIANT varResult;
    VARIANT* pArgs = NULL;
    
    VariantInit(&varResult);
    
    // Prepare arguments
    if (pInfo->paramCount > 0 && args)
    {
        pArgs = (VARIANT*)GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, pInfo->paramCount * sizeof(VARIANT));
        if (!pArgs)
            return E_OUTOFMEMORY;
        
        for (int i = 0; i < pInfo->paramCount; i++)
        {
            VariantInit(&pArgs[i]);
            hr = ConvertTypeToVariant(args[i], pInfo->paramTypes[i], &pArgs[i]);
            if (FAILED(hr))
                goto cleanup;
        }
        
        dispParams.rgvarg = pArgs;
        dispParams.cArgs = pInfo->paramCount;
    }
    
    // Call script function
    hr = pInfo->scriptFunction->lpVtbl->Invoke(
        pInfo->scriptFunction,
        DISPID_VALUE,
        &IID_NULL,
        LOCALE_USER_DEFAULT,
        DISPATCH_METHOD,
        &dispParams,
        &varResult,
        NULL,
        NULL
    );
    
    // Convert return value
    if (SUCCEEDED(hr) && returnValue)
    {
        hr = ConvertVariantToType(&varResult, pInfo->returnType, returnValue, pObj);
    }
    
cleanup:
    if (pArgs)
    {
        for (int i = 0; i < pInfo->paramCount; i++)
        {
            VariantClear(&pArgs[i]);
        }
        GlobalFree(pArgs);
    }
    
    VariantClear(&varResult);
    return hr;
}

// Generic callback stub implementation
static LRESULT CALLBACK GenericCallbackStub(int callbackIndex, LPARAM lParam1, LPARAM lParam2, LPARAM lParam3, LPARAM lParam4)
{
    DynamicWrapperX* pObj = g_callbackObjects[callbackIndex];
    if (!pObj)
        return 0;
    
    // Prepare arguments - this is simplified, real implementation would need
    // to properly handle the actual parameters based on the calling convention
    void* args[4];
    args[0] = &lParam1;
    args[1] = &lParam2;
    args[2] = &lParam3;
    args[3] = &lParam4;
    
    LRESULT result = 0;
    CallScriptFunction(pObj, callbackIndex, args, &result);
    
    return result;
}

// Callback stubs - these would normally be implemented in assembly
// for proper parameter handling, but simplified here
LRESULT CALLBACK CallbackStub0(void) { return GenericCallbackStub(0, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub1(void) { return GenericCallbackStub(1, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub2(void) { return GenericCallbackStub(2, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub3(void) { return GenericCallbackStub(3, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub4(void) { return GenericCallbackStub(4, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub5(void) { return GenericCallbackStub(5, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub6(void) { return GenericCallbackStub(6, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub7(void) { return GenericCallbackStub(7, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub8(void) { return GenericCallbackStub(8, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub9(void) { return GenericCallbackStub(9, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub10(void) { return GenericCallbackStub(10, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub11(void) { return GenericCallbackStub(11, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub12(void) { return GenericCallbackStub(12, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub13(void) { return GenericCallbackStub(13, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub14(void) { return GenericCallbackStub(14, 0, 0, 0, 0); }
LRESULT CALLBACK CallbackStub15(void) { return GenericCallbackStub(15, 0, 0, 0, 0); }
