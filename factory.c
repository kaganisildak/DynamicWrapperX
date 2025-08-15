#include "dynwrapx.h"

// Forward declarations for class factory
static HRESULT STDMETHODCALLTYPE ClassFactory_QueryInterface(IClassFactory* This, REFIID riid, void** ppv);
static ULONG STDMETHODCALLTYPE ClassFactory_AddRef(IClassFactory* This);
static ULONG STDMETHODCALLTYPE ClassFactory_Release(IClassFactory* This);
static HRESULT STDMETHODCALLTYPE ClassFactory_CreateInstance(IClassFactory* This, IUnknown* pUnkOuter, REFIID riid, void** ppv);
static HRESULT STDMETHODCALLTYPE ClassFactory_LockServer(IClassFactory* This, BOOL fLock);

// Class factory vtable
static IClassFactoryVtbl ClassFactoryVtbl = {
    ClassFactory_QueryInterface,
    ClassFactory_AddRef,
    ClassFactory_Release,
    ClassFactory_CreateInstance,
    ClassFactory_LockServer
};

// DllGetClassObject implementation
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    ClassFactory* pFactory;
    HRESULT hr;
    
    if (!ppv)
        return E_POINTER;
    
    *ppv = NULL;
    
    // Check if requesting our CLSID
    if (!IsEqualCLSID(rclsid, &CLSID_DynamicWrapperX))
        return CLASS_E_CLASSNOTAVAILABLE;
    
    // Create class factory
    pFactory = (ClassFactory*)GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof(ClassFactory));
    if (!pFactory)
        return E_OUTOFMEMORY;
    
    pFactory->vtbl.lpVtbl = &ClassFactoryVtbl;
    pFactory->refCount = 1;
    
    hr = ClassFactory_QueryInterface(&pFactory->vtbl, riid, ppv);
    ClassFactory_Release(&pFactory->vtbl);
    
    return hr;
}

// Class factory implementation
static HRESULT STDMETHODCALLTYPE ClassFactory_QueryInterface(IClassFactory* This, REFIID riid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    
    *ppv = NULL;
    
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
    {
        *ppv = This;
        ClassFactory_AddRef(This);
        return S_OK;
    }
    
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE ClassFactory_AddRef(IClassFactory* This)
{
    ClassFactory* pFactory = (ClassFactory*)This;
    return InterlockedIncrement(&pFactory->refCount);
}

static ULONG STDMETHODCALLTYPE ClassFactory_Release(IClassFactory* This)
{
    ClassFactory* pFactory = (ClassFactory*)This;
    LONG refs = InterlockedDecrement(&pFactory->refCount);
    
    if (refs == 0)
    {
        GlobalFree(pFactory);
    }
    
    return refs;
}

static HRESULT STDMETHODCALLTYPE ClassFactory_CreateInstance(IClassFactory* This, IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    return CreateDynamicWrapperX(pUnkOuter, riid, ppv);
}

static HRESULT STDMETHODCALLTYPE ClassFactory_LockServer(IClassFactory* This, BOOL fLock)
{
    if (fLock)
        InterlockedIncrement(&g_cLocks);
    else
        InterlockedDecrement(&g_cLocks);
    
    return S_OK;
}

HRESULT CallRegisteredFunction(DynamicWrapperX* pObj, FunctionInfo* pFunc, DISPPARAMS* pDispParams, VARIANT* pVarResult);
size_t GetTypeSize(ParameterType type);
HRESULT CallFunction(FARPROC proc, void** args, int argCount, void* returnValue);

