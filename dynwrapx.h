#ifndef DYNWRAPX_H
#define DYNWRAPX_H

#include <windows.h>
#include <oleauto.h>
#include <objbase.h>
#include <comcat.h>
#include <stdarg.h>

// Debug logging functions
void InitializeDebugLogging(void);
void CleanupDebugLogging(void);
#ifdef DEBUG_BUILD
void DebugLog(const char* level, const char* function, const char* format, ...);
#else
// No-op macro for release builds - completely removes debug calls
#define DebugLog(...) ((void)0)
#endif

// Architecture detection and feature definitions
#ifdef _WIN64
    #define DYNWRAPX_TARGET_X64 1
    #define DYNWRAPX_TARGET_X86 0
    #define DYNWRAPX_POINTER_SIZE 8
    #define DYNWRAPX_ADDRESS_TYPE LONGLONG
    #define DYNWRAPX_VARIANT_ADDR_TYPE VT_I8
    #define DYNWRAPX_ARCH_SUFFIX "x64"
#else
    #define DYNWRAPX_TARGET_X64 0
    #define DYNWRAPX_TARGET_X86 1
    #define DYNWRAPX_POINTER_SIZE 4
    #define DYNWRAPX_ADDRESS_TYPE LONG
    #define DYNWRAPX_VARIANT_ADDR_TYPE VT_I4
    #define DYNWRAPX_ARCH_SUFFIX "x86"
#endif

// Feature toggles based on architecture
#if DYNWRAPX_TARGET_X64
    #define DYNWRAPX_ENABLE_64BIT_ADDRESSES 1
    #define DYNWRAPX_ENABLE_LARGE_MEMORY 1
    #define DYNWRAPX_ENABLE_EXTENDED_VALIDATION 1
#else
    #define DYNWRAPX_ENABLE_64BIT_ADDRESSES 0
    #define DYNWRAPX_ENABLE_LARGE_MEMORY 0
    #define DYNWRAPX_ENABLE_EXTENDED_VALIDATION 0
#endif

// Architecture-specific type definitions
#if DYNWRAPX_TARGET_X64
    typedef ULONG_PTR DYNWRAPX_ADDR_T;
    #define DYNWRAPX_INVALID_ADDR ((DYNWRAPX_ADDR_T)-1)
    #define DYNWRAPX_MAX_SAFE_ADDR 0x7FFFFFFFFFFFFFFFULL
#else
    typedef ULONG DYNWRAPX_ADDR_T;
    #define DYNWRAPX_INVALID_ADDR ((DYNWRAPX_ADDR_T)-1)
    #define DYNWRAPX_MAX_SAFE_ADDR 0xFFFFFFFFUL
#endif

// Simple validation macros
#define DYNWRAPX_VALIDATE_ADDRESS(addr) (addr != 0)
#define DYNWRAPX_CONVERT_VARIANT_TO_ADDR(var, addr) E_NOTIMPL
#define DYNWRAPX_CONVERT_ADDR_TO_VARIANT(addr, var) E_NOTIMPL

// CLSID and IID definitions
// {89565275-A714-4a43-912E-978B935EDCCC}
extern const GUID CLSID_DynamicWrapperX;

// {89565276-A714-4a43-912E-978B935EDCCC} 
extern const GUID IID_IDynamicWrapperX;

// Type mappings for parameter parsing
typedef enum {
    TYPE_LONG = 0,          // l - signed 32-bit integer
    TYPE_ULONG,             // u - unsigned 32-bit integer  
    TYPE_HANDLE,            // h - handle
    TYPE_POINTER,           // p - pointer
    TYPE_SHORT,             // n - signed 16-bit integer
    TYPE_USHORT,            // t - unsigned 16-bit integer
    TYPE_CHAR,              // c - signed 8-bit integer
    TYPE_UCHAR,             // b - unsigned 8-bit integer
    TYPE_FLOAT,             // f - single-precision float
    TYPE_DOUBLE,            // d - double-precision float
    TYPE_LONGLONG,          // q - signed 64-bit integer (x64 only)
    TYPE_WSTRING,           // w - Unicode string
    TYPE_ASTRING,           // s - ANSI string
    TYPE_OSTRING,           // z - OEM string
    
    // Output parameter types (uppercase)
    TYPE_OUT_LONG,          // L - pointer to LONG
    TYPE_OUT_ULONG,         // U - pointer to ULONG
    TYPE_OUT_HANDLE,        // H - pointer to HANDLE
    TYPE_OUT_POINTER,       // P - pointer to pointer
    TYPE_OUT_SHORT,         // N - pointer to SHORT
    TYPE_OUT_USHORT,        // T - pointer to USHORT
    TYPE_OUT_CHAR,          // C - pointer to CHAR
    TYPE_OUT_UCHAR,         // B - pointer to UCHAR
    TYPE_OUT_FLOAT,         // F - pointer to FLOAT
    TYPE_OUT_DOUBLE,        // D - pointer to DOUBLE
    TYPE_OUT_WSTRING,       // W - output Unicode string
    TYPE_OUT_ASTRING,       // S - output ANSI string
    TYPE_OUT_OSTRING,       // Z - output OEM string
    TYPE_VOID               // v - no return value
} ParameterType;

// Function registration structure
typedef struct _FunctionInfo {
    BSTR functionName;
    DWORD functionId;
    int paramCount;
    ParameterType* paramTypes;
    ParameterType returnType;
    FARPROC functionPtr;
    struct _FunctionInfo* next;
} FunctionInfo;

// Callback registration structure
typedef struct _CallbackInfo {
    IDispatch* scriptFunction;
    int paramCount;
    ParameterType* paramTypes;
    ParameterType returnType;
    int callbackIndex;
} CallbackInfo;

// Memory allocation tracking
typedef struct _MemoryBlock {
    void* ptr;
    BOOL isGlobal;      // TRUE for GlobalAlloc, FALSE for SysAllocString
    struct _MemoryBlock* next;
} MemoryBlock;

// Main interface declaration
#undef INTERFACE
#define INTERFACE IDynamicWrapperX

DECLARE_INTERFACE_(IDynamicWrapperX, IDispatch)
{
    // IUnknown methods
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppv) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    
    // IDispatch methods
    STDMETHOD(GetTypeInfoCount)(THIS_ UINT* pctinfo) PURE;
    STDMETHOD(GetTypeInfo)(THIS_ UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) PURE;
    STDMETHOD(GetIDsOfNames)(THIS_ REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                            LCID lcid, DISPID* rgDispId) PURE;
    STDMETHOD(Invoke)(THIS_ DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                     DISPPARAMS* pDispParams, VARIANT* pVarResult,
                     EXCEPINFO* pExcepInfo, UINT* puArgErr) PURE;
};

// DynamicWrapperX implementation class
typedef struct _DynamicWrapperX {
    IDynamicWrapperX vtbl;
    LONG refCount;
    FunctionInfo* functionList;
    CallbackInfo callbacks[16];  // Maximum 16 callbacks
    MemoryBlock* memoryBlocks;
    CRITICAL_SECTION cs;
} DynamicWrapperX;

// COM factory
typedef struct _ClassFactory {
    IClassFactory vtbl;
    LONG refCount;
} ClassFactory;

// Exported DLL functions
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved);
STDAPI DllCanUnloadNow(void);
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);
STDAPI DllInstall(BOOL bInstall, PCWSTR pszCmdLine);

// Internal functions
// Forward declarations
typedef struct _DynamicWrapperX IDynWrapX;
typedef struct _ClassFactory IDynWrapXFactory;

// Common function declarations (architecture-agnostic)

// Function declarations
HRESULT CreateDynamicWrapperX(IUnknown* pUnkOuter, REFIID riid, void** ppv);
ParameterType ParseParameterType(WCHAR c);
HRESULT ConvertVariantToType(VARIANT* pVar, ParameterType type, void* pOut, DynamicWrapperX* pObj);
HRESULT ConvertTypeToVariant(void* pData, ParameterType type, VARIANT* pVar);
FARPROC LoadFunction(LPCWSTR libraryName, LPCWSTR functionName);
void CleanupMemoryBlocks(DynamicWrapperX* pObj);

// Additional function declarations
HRESULT CallRegisteredFunction(DynamicWrapperX* pObj, FunctionInfo* pFunc, DISPPARAMS* pDispParams, VARIANT* pVarResult);
size_t GetTypeSize(ParameterType type);
HRESULT CallFunction(FARPROC proc, void** args, int argCount, void* returnValue);
HRESULT ParseParameterString(LPCWSTR paramStr, ParameterType** types, int* count);
HRESULT CallScriptFunction(DynamicWrapperX* pObj, int callbackIndex, void** args, void* returnValue);

// Built-in method implementations
HRESULT DynWrap_Register(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_RegisterCallback(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_NumGet(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_NumPut(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_StrPtr(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_StrGet(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);
HRESULT DynWrap_Space(DynamicWrapperX* pObj, DISPPARAMS* pDispParams, VARIANT* pVarResult);

// Callback stubs (16 maximum)
LRESULT CALLBACK CallbackStub0(void);
LRESULT CALLBACK CallbackStub1(void);
LRESULT CALLBACK CallbackStub2(void);
LRESULT CALLBACK CallbackStub3(void);
LRESULT CALLBACK CallbackStub4(void);
LRESULT CALLBACK CallbackStub5(void);
LRESULT CALLBACK CallbackStub6(void);
LRESULT CALLBACK CallbackStub7(void);
LRESULT CALLBACK CallbackStub8(void);
LRESULT CALLBACK CallbackStub9(void);
LRESULT CALLBACK CallbackStub10(void);
LRESULT CALLBACK CallbackStub11(void);
LRESULT CALLBACK CallbackStub12(void);
LRESULT CALLBACK CallbackStub13(void);
LRESULT CALLBACK CallbackStub14(void);
LRESULT CALLBACK CallbackStub15(void);

// Global variables
extern HMODULE g_hModule;
extern LONG g_cObjects;
extern LONG g_cLocks;
extern DynamicWrapperX* g_callbackObjects[16];

// Method name constants
#define DISPID_REGISTER         1000
#define DISPID_REGISTERCALLBACK 1001
#define DISPID_NUMGET          1002
#define DISPID_NUMPUT          1003
#define DISPID_STRPTR          1004
#define DISPID_STRGET          1005
#define DISPID_SPACE           1006

// Helper macros
#define SAFE_RELEASE(p) if(p) { (p)->lpVtbl->Release(p); (p) = NULL; }
#define SAFE_SYSFREE(p) if(p) { SysFreeString(p); (p) = NULL; }

#endif // DYNWRAPX_H
