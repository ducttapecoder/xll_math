// export.h - export known functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

// Used to export undecorated function name from a dll.
// Put '#pragma XLLEXPORT' in every add-in function body.
#ifdef _MSC_VER
    #define XLLEXPORT comment(linker, "/export:" __FUNCDNAME__ "=" __FUNCTION__)
#else
    // GCC/Clang: Use __attribute__((used)) to prevent optimization from removing the function
    // Functions must use extern "C" and be listed in .def file
    #define XLLEXPORT __attribute__((used))
#endif

// For GCC/Clang, declare Excel callback functions with C linkage
#if defined(__GNUC__) || defined(__clang__)
    #define XLL_EXTERN_C extern "C"
    // Helper macro for function definitions
    #define XLL_BEGIN_EXTERN_C extern "C" {
    #define XLL_END_EXTERN_C }
#else
    #define XLL_EXTERN_C
    #define XLL_BEGIN_EXTERN_C
    #define XLL_END_EXTERN_C
#endif

#ifdef _MSC_VER
//#pragma comment(linker, "/include:" "_DllMain@12")
//#pragma comment(linker, "/export:" "XLCallVer@0")
#pragma comment(linker, "/export:xlAutoOpen@0=xlAutoOpen")
#pragma comment(linker, "/export:xlAutoClose@0=xlAutoClose")
#pragma comment(linker, "/export:xlAutoAdd@0=xlAutoAdd")
#pragma comment(linker, "/export:xlAutoRemove@0=xlAutoRemove")
#pragma comment(linker, "/export:xlAutoFree12@4=xlAutoFree12")
#pragma comment(linker, "/export:xlAutoRegister12@4=xlAutoRegister12")
#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")
//#pragma comment(linker, "/export:LPenHelper")
//#pragma comment(linker, "/export:xlAddInManagerInfo@4=xlAddInManagerInfo")
#endif