#include "dynwrapx.h"

// Forward declarations for vtable functions
static HRESULT STDMETHODCALLTYPE DynWrap_QueryInterface(IDynamicWrapperX* This, REFIID riid, void** ppv);
static ULONG STDMETHODCALLTYPE DynWrap_AddRef(IDynamicWrapperX* This);
static ULONG STDMETHODCALLTYPE DynWrap_Release(IDynamicWrapperX* This);
static HRESULT STDMETHODCALLTYPE DynWrap_GetTypeInfoCount(IDynamicWrapperX* This, UINT* pctinfo);
static HRESULT STDMETHODCALLTYPE DynWrap_GetTypeInfo(IDynamicWrapperX* This, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
static HRESULT STDMETHODCALLTYPE DynWrap_GetIDsOfNames(IDynamicWrapperX* This, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
static HRESULT STDMETHODCALLTYPE DynWrap_Invoke(IDynamicWrapperX* This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

// Virtual function table for IDynamicWrapperX
static IDynamicWrapperXVtbl DynWrapVtbl = {
    DynWrap_QueryInterface,
    DynWrap_AddRef,
    DynWrap_Release,
    DynWrap_GetTypeInfoCount,
    DynWrap_GetTypeInfo,
    DynWrap_GetIDsOfNames,
    DynWrap_Invoke
};

// Create DynamicWrapperX instance
HRESULT CreateDynamicWrapperX(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    DynamicWrapperX* pObj;
    HRESULT hr;
    
    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION;
        
    pObj = (DynamicWrapperX*)GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof(DynamicWrapperX));
    if (!pObj)
        return E_OUTOFMEMORY;
        
    pObj->vtbl.lpVtbl = &DynWrapVtbl;
    pObj->refCount = 1;
    pObj->functionList = NULL;
    pObj->memoryBlocks = NULL;
    
    // Initialize callbacks array
    for (int i = 0; i < 16; i++)
    {
        pObj->callbacks[i].scriptFunction = NULL;
        pObj->callbacks[i].paramTypes = NULL;
        pObj->callbacks[i].callbackIndex = -1;
    }
    
    InitializeCriticalSection(&pObj->cs);
    
    InterlockedIncrement(&g_cObjects);
    
    hr = DynWrap_QueryInterface(&pObj->vtbl, riid, ppv);
    DynWrap_Release(&pObj->vtbl);
    
    return hr;
}

// IUnknown implementation
static HRESULT STDMETHODCALLTYPE DynWrap_QueryInterface(IDynamicWrapperX* This, REFIID riid, void** ppv)
{
    
    if (!ppv)
        return E_POINTER;
        
    *ppv = NULL;
    
    if (IsEqualIID(riid, &IID_IUnknown) || 
        IsEqualIID(riid, &IID_IDispatch) ||
        IsEqualIID(riid, &IID_IDynamicWrapperX))
    {
        *ppv = This;
        DynWrap_AddRef(This);
        return S_OK;
    }
    
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE DynWrap_AddRef(IDynamicWrapperX* This)
{
    DynamicWrapperX* pObj = (DynamicWrapperX*)This;
    return InterlockedIncrement(&pObj->refCount);
}

static ULONG STDMETHODCALLTYPE DynWrap_Release(IDynamicWrapperX* This)
{
    DynamicWrapperX* pObj = (DynamicWrapperX*)This;
    LONG refs = InterlockedDecrement(&pObj->refCount);
    
    if (refs == 0)
    {
        // Cleanup function list
        FunctionInfo* pFunc = pObj->functionList;
        while (pFunc)
        {
            FunctionInfo* pNext = pFunc->next;
            SAFE_SYSFREE(pFunc->functionName);
            if (pFunc->paramTypes)
                GlobalFree(pFunc->paramTypes);
            GlobalFree(pFunc);
            pFunc = pNext;
        }
        
        // Cleanup callbacks
        for (int i = 0; i < 16; i++)
        {
            if (pObj->callbacks[i].scriptFunction)
            {
                SAFE_RELEASE(pObj->callbacks[i].scriptFunction);
                if (pObj->callbacks[i].paramTypes)
                    GlobalFree(pObj->callbacks[i].paramTypes);
                g_callbackObjects[i] = NULL;
            }
        }
        
        // Cleanup memory blocks
        CleanupMemoryBlocks(pObj);
        
        DeleteCriticalSection(&pObj->cs);
        GlobalFree(pObj);
        
        InterlockedDecrement(&g_cObjects);
    }
    
    return refs;
}

// IDispatch implementation
static HRESULT STDMETHODCALLTYPE DynWrap_GetTypeInfoCount(IDynamicWrapperX* This, UINT* pctinfo)
{
    if (!pctinfo)
        return E_POINTER;
    *pctinfo = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE DynWrap_GetTypeInfo(IDynamicWrapperX* This, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    if (!ppTInfo)
        return E_POINTER;
    *ppTInfo = NULL;
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE DynWrap_GetIDsOfNames(IDynamicWrapperX* This, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    DynamicWrapperX* pObj = (DynamicWrapperX*)This;
    HRESULT hr = S_OK;
    
    if (!rgszNames || !rgDispId || cNames == 0)
        return E_INVALIDARG;
        
    for (UINT i = 0; i < cNames; i++)
    {
        rgDispId[i] = DISPID_UNKNOWN;
        
        // Check built-in methods
        if (wcscmp(rgszNames[i], L"Register") == 0)
            rgDispId[i] = DISPID_REGISTER;
        else if (wcscmp(rgszNames[i], L"RegisterCallback") == 0)
            rgDispId[i] = DISPID_REGISTERCALLBACK;
        else if (wcscmp(rgszNames[i], L"NumGet") == 0)
            rgDispId[i] = DISPID_NUMGET;
        else if (wcscmp(rgszNames[i], L"NumPut") == 0)
            rgDispId[i] = DISPID_NUMPUT;
        else if (wcscmp(rgszNames[i], L"StrPtr") == 0)
            rgDispId[i] = DISPID_STRPTR;
        else if (wcscmp(rgszNames[i], L"StrGet") == 0)
            rgDispId[i] = DISPID_STRGET;
        else if (wcscmp(rgszNames[i], L"Space") == 0)
            rgDispId[i] = DISPID_SPACE;
        else
        {
            // Check registered functions
            EnterCriticalSection(&pObj->cs);
            FunctionInfo* pFunc = pObj->functionList;
            while (pFunc)
            {
                if (wcscmp(rgszNames[i], pFunc->functionName) == 0)
                {
                    rgDispId[i] = pFunc->functionId;
                    break;
                }
                pFunc = pFunc->next;
            }
            LeaveCriticalSection(&pObj->cs);
            
            if (rgDispId[i] == DISPID_UNKNOWN)
                hr = DISP_E_UNKNOWNNAME;
        }
    }
    
    return hr;
}

static HRESULT STDMETHODCALLTYPE DynWrap_Invoke(IDynamicWrapperX* This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    DynamicWrapperX* pObj = (DynamicWrapperX*)This;
    HRESULT hr = S_OK;
    
    DebugLog("DEBUG", "DynWrap_Invoke", "Method called: dispIdMember=%ld, wFlags=0x%x, paramCount=%d", dispIdMember, wFlags, pDispParams ? pDispParams->cArgs : 0);
    
    if (wFlags & DISPATCH_METHOD)
    {
        switch (dispIdMember)
        {
        case DISPID_REGISTER:
            hr = DynWrap_Register(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_REGISTERCALLBACK:
            hr = DynWrap_RegisterCallback(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_NUMGET:
            hr = DynWrap_NumGet(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_NUMPUT:
            hr = DynWrap_NumPut(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_STRPTR:
            hr = DynWrap_StrPtr(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_STRGET:
            hr = DynWrap_StrGet(pObj, pDispParams, pVarResult);
            break;
            
        case DISPID_SPACE:
            hr = DynWrap_Space(pObj, pDispParams, pVarResult);
            break;
            
        default:
            // Check if it's a registered function
            DebugLog("DEBUG", "DynWrap_Invoke", "Looking for registered function with ID %ld", dispIdMember);
            EnterCriticalSection(&pObj->cs);
            FunctionInfo* pFunc = pObj->functionList;
            while (pFunc)
            {
                if (pFunc->functionId == dispIdMember)
                {
                    DebugLog("DEBUG", "DynWrap_Invoke", "Found registered function: %S", pFunc->functionName ? pFunc->functionName : L"<unknown>");
                    hr = CallRegisteredFunction(pObj, pFunc, pDispParams, pVarResult);
                    break;
                }
                pFunc = pFunc->next;
            }
            LeaveCriticalSection(&pObj->cs);
            
            if (!pFunc)
                hr = DISP_E_MEMBERNOTFOUND;
            break;
        }
    }
    else
    {
        hr = DISP_E_MEMBERNOTFOUND;
    }
    
    return hr;
}

// Call a registered function
HRESULT CallRegisteredFunction(DynamicWrapperX* pObj, FunctionInfo* pFunc, DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
    HRESULT hr = S_OK;
    void** args = NULL;
    void* returnValue = NULL;
    
    if (!pFunc->functionPtr)
        return E_FAIL;
    
    // Allocate memory for arguments
    if (pFunc->paramCount > 0)
    {
        args = (void**)GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, pFunc->paramCount * sizeof(void*));
        if (!args)
            return E_OUTOFMEMORY;
        
        // Convert parameters
        for (int i = 0; i < pFunc->paramCount && i < (int)pDispParams->cArgs; i++)
        {
            VARIANT* pArg = &pDispParams->rgvarg[pDispParams->cArgs - 1 - i]; // Parameters are in reverse order
            
            // Allocate space for the parameter
            size_t argSize = GetTypeSize(pFunc->paramTypes[i]);
            args[i] = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, argSize);
            if (!args[i])
            {
                hr = E_OUTOFMEMORY;
                goto cleanup;
            }
            
            hr = ConvertVariantToType(pArg, pFunc->paramTypes[i], args[i], pObj);
            if (FAILED(hr))
                goto cleanup;
        }
    }
    
    // Allocate space for return value
    if (pFunc->returnType != TYPE_VOID)
    {
        size_t retSize = GetTypeSize(pFunc->returnType);
        returnValue = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, retSize);
        if (!returnValue)
        {
            hr = E_OUTOFMEMORY;
            goto cleanup;
        }
    }
    
    // Call the function using inline assembly or a generic caller
    hr = CallFunction(pFunc->functionPtr, args, pFunc->paramCount, returnValue);
    
    // Convert return value
    if (SUCCEEDED(hr) && pVarResult && returnValue)
    {
        DebugLog("DEBUG", "DynWrap_Invoke", "Converting return value, returnType=%d", pFunc->returnType);
        hr = ConvertTypeToVariant(returnValue, pFunc->returnType, pVarResult);
        if (SUCCEEDED(hr) && pFunc->returnType == TYPE_POINTER) {
            DebugLog("DEBUG", "DynWrap_Invoke", "Pointer return converted to VT=%d", V_VT(pVarResult));
            if (V_VT(pVarResult) == VT_I8) {
                DebugLog("DEBUG", "DynWrap_Invoke", "VT_I8 value: %lld (0x%llx)", V_I8(pVarResult), (UINT64)V_I8(pVarResult));
            } else if (V_VT(pVarResult) == VT_UI4) {
                DebugLog("DEBUG", "DynWrap_Invoke", "VT_UI4 value: %u (0x%08x)", V_UI4(pVarResult), V_UI4(pVarResult));
            }
        }
    }
    
cleanup:
    // Cleanup arguments
    if (args)
    {
        for (int i = 0; i < pFunc->paramCount; i++)
        {
            if (args[i])
                GlobalFree(args[i]);
        }
        GlobalFree(args);
    }
    
    if (returnValue)
        GlobalFree(returnValue);
    
    return hr;
}

// Get size of a type
size_t GetTypeSize(ParameterType type)
{
    switch (type)
    {
    case TYPE_CHAR:
    case TYPE_UCHAR:
    case TYPE_OUT_CHAR:
    case TYPE_OUT_UCHAR:
        return sizeof(CHAR);
        
    case TYPE_SHORT:
    case TYPE_USHORT:
    case TYPE_OUT_SHORT:
    case TYPE_OUT_USHORT:
        return sizeof(SHORT);
        
    case TYPE_LONG:
    case TYPE_ULONG:
    case TYPE_OUT_LONG:
    case TYPE_OUT_ULONG:
        return sizeof(LONG);
        
    case TYPE_HANDLE:
    case TYPE_POINTER:
    case TYPE_OUT_HANDLE:
    case TYPE_OUT_POINTER:
        return sizeof(void*);
        
    case TYPE_LONGLONG:
        return sizeof(LONGLONG);
        
    case TYPE_FLOAT:
    case TYPE_OUT_FLOAT:
        return sizeof(FLOAT);
        
    case TYPE_DOUBLE:
    case TYPE_OUT_DOUBLE:
        return sizeof(DOUBLE);
        
    case TYPE_WSTRING:
    case TYPE_ASTRING:
    case TYPE_OSTRING:
    case TYPE_OUT_WSTRING:
    case TYPE_OUT_ASTRING:
    case TYPE_OUT_OSTRING:
        return sizeof(void*);
        
    case TYPE_VOID:
        return 0;
        
    default:
        return sizeof(LONG);
    }
}

// Enhanced generic function caller supporting more parameters
HRESULT CallFunction(FARPROC proc, void** args, int argCount, void* returnValue)
{

    DebugLog("DEBUG", "CallFunction", "Starting function call: proc=0x%p, argCount=%d, returnValue=0x%p", proc, argCount, returnValue);
    
    typedef LONG_PTR (WINAPI *GenericFunc0)(void);
    typedef LONG_PTR (WINAPI *GenericFunc1)(LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc2)(LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc3)(LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc4)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc5)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc6)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc7)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc8)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc9)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    typedef LONG_PTR (WINAPI *GenericFunc10)(LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR, LONG_PTR);
    
    LONG_PTR result = 0;
    
    // Validate parameters
    if (!proc) {
        DebugLog("ERROR", "CallFunction", "proc parameter is NULL");
        return E_INVALIDARG;
    }
    
    // Convert args to LONG_PTR array for easier handling with validation
    LONG_PTR convertedArgs[10] = {0};
    for (int i = 0; i < argCount && i < 10; i++) {
        if (args[i]) {
            convertedArgs[i] = *(LONG_PTR*)args[i];
            DebugLog("DEBUG", "CallFunction", "arg[%d]: 0x%p -> 0x%llx", i, args[i], (unsigned long long)convertedArgs[i]);
        } else {
            DebugLog("DEBUG", "CallFunction", "arg[%d]: NULL -> 0x0", i);
        }
    }
    
    // Call function - simplified without exception handling for MinGW compatibility
    DebugLog("DEBUG", "CallFunction", "About to call function with %d parameters", argCount);
    
    switch (argCount)
    {
    case 0:
        DebugLog("DEBUG", "CallFunction", "Calling 0-parameter function");
        result = ((GenericFunc0)proc)();
        break;
    case 1:
        DebugLog("DEBUG", "CallFunction", "Calling 1-parameter function with args[0]=0x%llx", (unsigned long long)convertedArgs[0]);
        result = ((GenericFunc1)proc)(convertedArgs[0]);
        break;
    case 2:
        DebugLog("DEBUG", "CallFunction", "Calling 2-parameter function with args[0]=0x%llx, args[1]=0x%llx", (unsigned long long)convertedArgs[0], (unsigned long long)convertedArgs[1]);
        result = ((GenericFunc2)proc)(convertedArgs[0], convertedArgs[1]);
        break;
    case 3:
        DebugLog("DEBUG", "CallFunction", "Calling 3-parameter function");
        result = ((GenericFunc3)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2]);
        break;
    case 4:
        DebugLog("DEBUG", "CallFunction", "Calling 4-parameter function");
        result = ((GenericFunc4)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3]);
        break;
    case 5:
        DebugLog("DEBUG", "CallFunction", "Calling 5-parameter function");
        result = ((GenericFunc5)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4]);
        break;
    case 6:
        DebugLog("DEBUG", "CallFunction", "Calling 6-parameter function");
        result = ((GenericFunc6)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4], convertedArgs[5]);
        break;
    case 7:
        DebugLog("DEBUG", "CallFunction", "Calling 7-parameter function");
        result = ((GenericFunc7)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4], convertedArgs[5], convertedArgs[6]);
        break;
    case 8:
        DebugLog("DEBUG", "CallFunction", "Calling 8-parameter function");
        result = ((GenericFunc8)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4], convertedArgs[5], convertedArgs[6], convertedArgs[7]);
        break;
    case 9:
        DebugLog("DEBUG", "CallFunction", "Calling 9-parameter function");
        result = ((GenericFunc9)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4], convertedArgs[5], convertedArgs[6], convertedArgs[7], convertedArgs[8]);
        break;
    case 10:
        DebugLog("DEBUG", "CallFunction", "Calling 10-parameter function (CreateProcessW)");
        result = ((GenericFunc10)proc)(convertedArgs[0], convertedArgs[1], convertedArgs[2], convertedArgs[3], convertedArgs[4], convertedArgs[5], convertedArgs[6], convertedArgs[7], convertedArgs[8], convertedArgs[9]);
        break;
    default:
        DebugLog("ERROR", "CallFunction", "Too many parameters: %d (maximum 10)", argCount);
        return E_NOTIMPL; // More than 10 parameters - very rare for Windows APIs
    }
    
    DebugLog("DEBUG", "CallFunction", "Function call completed, result=0x%llx (%lld)", (unsigned long long)result, (long long)result);
    
    // Store result
    if (returnValue) {
        *(LONG_PTR*)returnValue = result;
        DebugLog("DEBUG", "CallFunction", "Result stored at 0x%p", returnValue);
    }
    
    DebugLog("DEBUG", "CallFunction", "CallFunction completed successfully");
    return S_OK;
}
