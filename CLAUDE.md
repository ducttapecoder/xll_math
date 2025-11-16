# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a modern Excel add-in (XLL) project built with C++23 using the xll24 framework. The project creates native Excel functions callable from worksheets, providing high-performance code.

### Requirements

**Build Tools:**
- CMake 3.28+
- Ninja build system (recommended)

**Compilers (choose one):**
- Visual Studio 2026 with MSVC 19+ (v145 toolset)
- GCC 15+ (MSYS2 UCRT64 for x64, MINGW32 for x86)
- Clang 21+ (MSYS2 CLANG64 for x64, CLANG32 for x86)

**Platform:**
- Windows 10 or Windows 11
- Both x64 and x86 (32-bit) architectures supported

**Language Standard:**
- C++23 (`/std:c++latest` or `-std=c++23`)
- C23 for C files

### Key Components

**Project Structure:**
- `xll_template.cpp` - Main add-in with Excel function definitions
- `xll_template.h` - Header for add-in declarations
- `xll24/` - Static library providing XLL framework
  - `include/xll.h` - Main header (includes all components)
  - `include/addin.h` - AddIn class for registration
  - `include/args.h` - Argument specifications
  - `include/oper.h` - OPER class for Excel data types
  - `src/` - Implementation files
  - `test/` - Test implementations

**CMake Configuration:**
- `CMakeLists.txt` - Root project configuration
- `xll24/CMakeLists.txt` - xll24 library configuration
- `cmake/CompilerSettings.cmake` - Core compiler settings
- `cmake/CompilerWarnings.cmake` - Warning configuration
- `cmake/FindXLCall32.cmake` - Excel SDK library detection
- `CMakePresets.json` - Build presets for all compilers

### Architecture

**Two-Target Build:**
1. **xll24** - Static library (`.lib`) providing XLL framework
2. **xll_template** - Shared library (`.xll`) - the Excel add-in

**Function Registration:**
- Create global `AddIn` object with metadata using `Function()` or `Macro()` helpers
- Implement function with `WINAPI` calling convention
- Add `#pragma XLLEXPORT` in function body
- AddIn object auto-registers when Excel loads the XLL

**Data Flow:**
- Excel ↔ XLL via C SDK interface (XLOPER12 structures)
- xll24 provides `OPER` wrapper for easier manipulation
- Functions receive C++ types or OPER objects
- Return values auto-convert to Excel types

## Build System

### Quick Start

```bash
# Configure
cmake --preset release              # MSVC x64 Release
cmake --preset ucrt64-release       # GCC/UCRT64 Release
cmake --preset clang64-release      # Clang64 Release

# Build
cmake --build --preset release
cmake --build --preset ucrt64-release
cmake --build --preset clang64-release

# Output locations
build/msvc-x64-release/bin/xll_template.xll
build/ucrt64-release/bin/libxll_template.xll
build/clang64-release/bin/libxll_template.xll
```

### Available Presets

**MSVC (Visual Studio 2026):**
- `debug` - MSVC x64 Debug
- `release` - MSVC x64 Release
- `multi` - MSVC x64 Multi-Config (both Debug and Release)
- `debug-x86` - MSVC x86 Debug
- `release-x86` - MSVC x86 Release

**GCC (MSYS2):**
- `ucrt64-debug` - GCC UCRT64 (x64) Debug
- `ucrt64-release` - GCC UCRT64 (x64) Release
- `mingw32-debug` - GCC MinGW32 (x86) Debug
- `mingw32-release` - GCC MinGW32 (x86) Release

**Clang (MSYS2):**
- `clang64-debug` - Clang64 (x64) Debug
- `clang64-release` - Clang64 (x64) Release
- `clang32-debug` - Clang32 (x86) Debug
- `clang32-release` - Clang32 (x86) Release

### Debug Configuration

Visual Studio automatically launches Excel with the XLL loaded when you press F5:
- Launch path: Retrieved from registry `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`
- Command arguments: `/x "$(TargetPath)"`
- Configured in CMakeLists.txt via `VS_DEBUGGER_COMMAND_ARGUMENTS`

### Build Configuration

**Modern CMake Features:**
- Target-based configuration (all settings via `target_*` commands)
- Generator expressions for config-specific settings
- Imported targets (e.g., `XLCall32::XLCall32`)
- Proper package configuration files
- Presets for reproducible builds

**Compiler Settings:**
- C++23 standard (`CMAKE_CXX_STANDARD 23`)
- Unicode support (`UNICODE`, `_UNICODE`)
- UTF-8 source encoding (`/utf-8` or `-finput-charset=UTF-8`)
- Unsigned char default (`/J` or `-funsigned-char`)
- Static runtime linking
- LTO/IPO optimization (Release builds)
- Debug info generation (`/Zi`, `/DEBUG`)
- SDL security checks (`/sdl`)

**Warning Levels:**
- xll24 library: Warnings as errors (MSVC only, practical for GCC/Clang)
- Application code: Standard warnings enabled
- Comprehensive warning coverage via CompilerWarnings.cmake

## Development Guide

### Adding Excel Functions

1. **Define the function:**

```cpp
AddIn xai_myfunction(
    Function(XLL_DOUBLE, "xll_myfunction", "MY.FUNCTION")
    .Arguments({
        Arg(XLL_DOUBLE, "x", "is the input value."),
        Arg(XLL_DOUBLE, "y", "is another input value.")
    })
    .FunctionHelp("Computes result from x and y.")
    .Category("MATH")
    .HelpTopic("https://docs.example.com/myfunction")
);
```

2. **Implement the function:**

```cpp
double WINAPI xll_myfunction(double x, double y)
{
#pragma XLLEXPORT
    return x * y;  // Your implementation
}
```

**Important:**
- Use `WINAPI` calling convention (required by Excel)
- Include `#pragma XLLEXPORT` in function body
- Function name convention: `xll_` prefix for implementation
- Excel name convention: `MY.FUNCTION` or similar

### Adding Excel Macros

Macros perform actions (side effects) and take no arguments:

```cpp
AddIn xai_mymacro(
    Macro("xll_mymacro", "MY.MACRO")
    .FunctionHelp("Performs an action in Excel.")
);

int WINAPI xll_mymacro(void)
{
#pragma XLLEXPORT
    Excel(xlcAlert, OPER("Macro executed!"));
    return TRUE;
}
```

Call macros from Excel: Alt+F8, type macro name, click Run.

### Excel Data Types Reference

This section documents how to work with different Excel data types based on patterns from real-world XLL projects. All types are defined in `xll24/include/type.h`.

#### Return Type Patterns

When you specify a return type in `Function()`, you must match it with the corresponding C++ type in your implementation:

| XLL Type | C++ Return Type | Use Case | Example |
|----------|----------------|----------|---------|
| `XLL_DOUBLE` | `double` | Single numeric value | Simple calculations |
| `XLL_FP` | `_FP12*` or `FP*` | 2D array of doubles | Matrices, ranges |
| `XLL_LPXLOPER` | `LPXLOPER12` | Any Excel value | Generic return, mixed types |
| `XLL_HANDLE` | `HANDLEX` | Handle to C++ object | Object persistence |
| `XLL_BOOL` | `BOOL` | True/false value | Flags, conditions |
| `XLL_CSTRING` | `const wchar_t*` | Unicode text | String results |
| `XLL_WORD` | `WORD` | 16-bit unsigned integer | Small integers |
| `XLL_DWORD` | `DWORD` | 32-bit unsigned integer | Large integers |
| `XLL_LONG` | `LONG` | 32-bit signed integer | Signed integers |

**Examples:**

```cpp
// Return double
AddIn xai_square(
    Function(XLL_DOUBLE, "xll_square", "SQUARE")
    .Arguments({Arg(XLL_DOUBLE, "x", "is a number.")})
);
double WINAPI xll_square(double x) {
#pragma XLLEXPORT
    return x * x;
}

// Return array (FP12)
AddIn xai_identity(
    Function(XLL_FP, "xll_identity", "IDENTITY")
    .Arguments({Arg(XLL_WORD, "n", "is matrix dimension.")})
);
_FP12* WINAPI xll_identity(WORD n) {
#pragma XLLEXPORT
    static FP result(n, n);
    for (WORD i = 0; i < n; ++i)
        for (WORD j = 0; j < n; ++j)
            result(i, j) = (i == j) ? 1.0 : 0.0;
    return result.get();
}

// Return generic value (LPXLOPER12)
AddIn xai_info(
    Function(XLL_LPXLOPER, "xll_info", "INFO")
);
LPXLOPER12 WINAPI xll_info() {
#pragma XLLEXPORT
    static OPER result;
    result = OPER("Version 1.0");
    return &result;
}

// Return handle to C++ object
AddIn xai_array_set(
    Function(XLL_HANDLE, "xll_array_set", "ARRAY.SET")
    .Arguments({Arg(XLL_FP, "array", "is an array.")})
);
HANDLEX WINAPI xll_array_set(_FP12* pa) {
#pragma XLLEXPORT
    handle<FP12> h(new FP12(*pa));
    return h.get();
}
```

#### Argument Type Patterns

When you specify argument types in `.Arguments()`, your function parameters must match:

| XLL Type | C++ Parameter Type | Use Case | Notes |
|----------|-------------------|----------|-------|
| `XLL_DOUBLE` | `double` | Numeric input | Single cell or value |
| `XLL_FP` | `_FP12*` or `const _FP12*` | Array input | Range of numbers only |
| `XLL_LPXLOPER` | `LPXLOPER12` | Any Excel value | Generic input, check type |
| `XLL_HANDLE` | `HANDLEX` | Object handle | From previous HANDLE return |
| `XLL_BOOL` | `BOOL` | Boolean flag | TRUE/FALSE, 1/0 |
| `XLL_CSTRING` | `const wchar_t*` | Unicode string | Text input |
| `XLL_WORD` | `WORD` | Small integer | 0 to 65,535 |
| `XLL_DWORD` | `DWORD` | Large unsigned | 0 to 4,294,967,295 |
| `XLL_LONG` | `LONG` | Signed integer | -2,147,483,648 to 2,147,483,647 |

**Examples:**

```cpp
// Multiple numeric arguments
AddIn xai_multiply(
    Function(XLL_DOUBLE, "xll_multiply", "MULTIPLY")
    .Arguments({
        Arg(XLL_DOUBLE, "x", "is first number."),
        Arg(XLL_DOUBLE, "y", "is second number.")
    })
);
double WINAPI xll_multiply(double x, double y) {
#pragma XLLEXPORT
    return x * y;
}

// Array argument
AddIn xai_array_sum(
    Function(XLL_DOUBLE, "xll_array_sum", "ARRAY.SUM")
    .Arguments({Arg(XLL_FP, "array", "is array of numbers.")})
);
double WINAPI xll_array_sum(const _FP12* pa) {
#pragma XLLEXPORT
    double sum = 0;
    for (int i = 0; i < pa->rows * pa->columns; ++i)
        sum += pa->array[i];
    return sum;
}

// Generic OPER argument
AddIn xai_type_name(
    Function(XLL_LPXLOPER, "xll_type_name", "TYPE.NAME")
    .Arguments({Arg(XLL_LPXLOPER, "value", "is any value.")})
);
LPXLOPER12 WINAPI xll_type_name(LPXLOPER12 px) {
#pragma XLLEXPORT
    static OPER result;
    switch (px->xltype) {
        case xltypeNum: result = OPER("Number"); break;
        case xltypeStr: result = OPER("String"); break;
        case xltypeBool: result = OPER("Boolean"); break;
        case xltypeErr: result = OPER("Error"); break;
        case xltypeMulti: result = OPER("Array"); break;
        default: result = OPER("Other"); break;
    }
    return &result;
}

// Handle argument - retrieve stored object
AddIn xai_array_get(
    Function(XLL_FP, "xll_array_get", "ARRAY.GET")
    .Arguments({Arg(XLL_HANDLE, "handle", "is array handle.")})
);
_FP12* WINAPI xll_array_get(HANDLEX h) {
#pragma XLLEXPORT
    handle<FP12> ha(h);
    return ha.ptr();
}

// Boolean and integer arguments
AddIn xai_random_array(
    Function(XLL_FP, "xll_random_array", "RANDOM.ARRAY")
    .Arguments({
        Arg(XLL_WORD, "rows", "is number of rows."),
        Arg(XLL_WORD, "cols", "is number of columns."),
        Arg(XLL_BOOL, "integers", "is flag for integer values.")
    })
);
_FP12* WINAPI xll_random_array(WORD rows, WORD cols, BOOL integers) {
#pragma XLLEXPORT
    static FP result;
    result.resize(rows, cols);
    for (int i = 0; i < rows * cols; ++i) {
        double val = (double)rand() / RAND_MAX;
        result.array()[i] = integers ? floor(val * 100) : val;
    }
    return result.get();
}

// String argument
AddIn xai_string_length(
    Function(XLL_DOUBLE, "xll_string_length", "STRING.LENGTH")
    .Arguments({Arg(XLL_CSTRING, "text", "is a string.")})
);
double WINAPI xll_string_length(const wchar_t* text) {
#pragma XLLEXPORT
    return static_cast<double>(wcslen(text));
}
```

#### Working with FP12 Arrays

The `_FP12` structure represents a 2D array of doubles:

```cpp
// FP12 structure (from Excel SDK)
typedef struct _FP12 {
    int rows;       // Number of rows
    int columns;    // Number of columns
    double array[1]; // First element of array (variable length)
} FP12;

// Access elements
double value = pa->array[row * pa->columns + col];

// Using FP wrapper class (recommended)
FP matrix(3, 3);  // Create 3x3 matrix
matrix(0, 0) = 1.0;  // Set element at row 0, col 0
matrix(1, 2) = 5.0;  // Set element at row 1, col 2
double val = matrix(2, 1);  // Get element at row 2, col 1
return matrix.get();  // Returns _FP12*
```

#### Working with OPER/XLOPER12

The `OPER` class wraps Excel's `XLOPER12` structure for type-safe handling:

```cpp
// Check type
if (px->xltype == xltypeNum) {
    double value = px->val.num;
}
else if (px->xltype == xltypeStr) {
    const wchar_t* text = px->val.str;
}
else if (px->xltype == xltypeMulti) {
    int rows = px->val.array.rows;
    int cols = px->val.array.columns;
}

// XLOPER12 types (from XLCALL.H)
// xltypeNum    - Number
// xltypeStr    - String
// xltypeBool   - Boolean
// xltypeErr    - Error value
// xltypeMulti  - Array
// xltypeMissing - Missing argument
// xltypeNil    - Empty cell
// xltypeRef    - Reference
```

#### Optional Arguments

Use `XLL_LPXLOPER` for optional arguments and check for `xltypeMissing`:

```cpp
AddIn xai_optional_example(
    Function(XLL_DOUBLE, "xll_optional_example", "OPTIONAL.EXAMPLE")
    .Arguments({
        Arg(XLL_DOUBLE, "x", "is required value."),
        Arg(XLL_LPXLOPER, "_y", "is optional value (default 1.0).")
    })
);
double WINAPI xll_optional_example(double x, LPXLOPER12 py) {
#pragma XLLEXPORT
    double y = 1.0;  // Default value
    if (py->xltype != xltypeMissing && py->xltype == xltypeNum) {
        y = py->val.num;
    }
    return x * y;
}
```

**Note:** Optional arguments conventionally have names starting with `_` (underscore).

#### Handles for Object Persistence

Use handles to store C++ objects between Excel function calls:

```cpp
#include "xll24/include/handle.h"

// Store object
HANDLEX WINAPI xll_create_object(double value) {
#pragma XLLEXPORT
    handle<MyClass> h(new MyClass(value));
    return h.get();  // Returns handle ID
}

// Retrieve object
double WINAPI xll_use_object(HANDLEX h) {
#pragma XLLEXPORT
    handle<MyClass> obj(h);
    return obj->getValue();  // Use object
}

// Delete object (optional - handles auto-cleanup)
BOOL WINAPI xll_delete_object(HANDLEX h) {
#pragma XLLEXPORT
    handle<MyClass> obj(h);
    obj.~handle();  // Explicit cleanup
    return TRUE;
}
```

#### GCC/Clang Name Mangling

For GCC and Clang compilers, you must prevent C++ name mangling for exported functions:

```cpp
// Add extern "C" for GCC/Clang
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_myfunction(double x) {
#pragma XLLEXPORT
    return x * x;
}
```

This is required because Excel expects C-style function names, not C++ mangled names.

#### Summary: Type Matching Rules

**Critical:** The return type in `Function()` MUST match your implementation:

```cpp
// ✓ CORRECT
Function(XLL_DOUBLE, ...)
double WINAPI xll_func(...) { ... }

// ✗ WRONG - Mismatch!
Function(XLL_DOUBLE, ...)
LPXLOPER12 WINAPI xll_func(...) { ... }  // Runtime error!
```

**Critical:** Argument types in `.Arguments()` MUST match parameters:

```cpp
// ✓ CORRECT
.Arguments({Arg(XLL_DOUBLE, "x", ...)})
double WINAPI xll_func(double x) { ... }

// ✗ WRONG - Mismatch!
.Arguments({Arg(XLL_DOUBLE, "x", ...)})
double WINAPI xll_func(LPXLOPER12 x) { ... }  // Runtime error!
```

### Calling Excel Functions

Use the `Excel()` function with Excel function codes from XLCALL.H:

```cpp
OPER sum = Excel(xlfSum, range);           // Worksheet function
OPER cell = Excel(xlfActiveCell);          // Get active cell
Excel(xlcAlert, OPER("Message"));          // Show alert (macro)
```

**Function Codes:**
- `xlf*` - Worksheet functions (pure, no side effects)
- `xlc*` - Macro functions (side effects allowed)

Reference: https://xlladdins.github.io/Excel4Macros/

### Mathematical Documentation

Use HTML with KaTeX for formulas:

```cpp
.Documentation(R"(
<p>The Black-Scholes formula:</p>
<p>\[C = S_0 N(d_1) - K e^{-rT} N(d_2)\]</p>
<p>where \(d_1 = \frac{\ln(S_0/K) + (r + \sigma^2/2)T}{\sigma\sqrt{T}}\)</p>
)")
```

## Code Style

**General:**
- Namespace: `namespace xll { ... }`
- Modern C++23 features encouraged
- Header-only where possible
- UTF-8 strings supported

**Naming:**
- Implementation functions: `xll_function_name`
- Excel names: `MY.FUNCTION` or `MYFUNCTION`
- AddIn objects: `xai_function_name`

**Error Handling:**
- Exceptions in registration code
- Return Excel error values in functions (`xlerr::Num`, `xlerr::Value`, etc.)

**Important:**
- All Excel-facing functions MUST use `WINAPI`
- All Excel-facing functions MUST have `#pragma XLLEXPORT` in body
- Target file extension is `.xll` (not `.dll`)

## Testing

Currently manual testing:
1. Build the XLL
2. Load in Excel (File → Options → Add-ins → Go → Browse)
3. Verify functions work in worksheets
4. Verify macros work (Alt+F8)

Automated testing can be added using:
- CTest for C++ unit tests
- Python with xlwings for integration tests

## CMake Helper Functions

Available from CompilerSettings.cmake:

```cmake
# Apply LTO/IPO and optimizations
apply_compiler_settings(target_name)

# Link library with whole-archive (needed for XLL auto-functions)
link_whole_archive(xll_target static_library)

# Set Windows subsystem
set_windows_subsystem(target WINDOWS)  # or CONSOLE
```

Available from CompilerWarnings.cmake:

```cmake
# Apply warning configuration
set_project_warnings(target_name)                # Standard warnings
set_project_warnings(target_name WARNINGS_AS_ERRORS)  # Strict mode
```

## Git Workflow

The xll24 library is included as a Git submodule:

```bash
# Initial clone
git clone --recurse-submodules <repo-url>

# Update submodule
git submodule update --init --recursive

# Update to latest xll24
cd xll24
git pull origin main
cd ..
git add xll24
git commit -m "Update xll24 submodule"
```

## Common Issues

**Issue:** XLL doesn't load in Excel
- Check that file extension is `.xll`
- Verify Excel bitness matches build (x64)
- Check for missing dependencies (XLCALL32.DLL)

**Issue:** Functions not appearing in Excel
- Verify `#pragma XLLEXPORT` in function body
- Check `WINAPI` calling convention
- Ensure AddIn object is global (not local)

**Issue:** Build fails with linker errors
- Verify XLCALL32.LIB is in `xll24/lib/` directory
- Check that whole-archive linking is enabled
- For MSVC: Ensure library and linker can find xll.lib

**Issue:** Compiler version too old
- CMake 3.28+ required
- MSVC 19+ (Visual Studio 2026) required
- GCC 15+ required
- Clang 21+ required
- C++23 support required

## Additional Resources

- Excel SDK Documentation: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk
- Excel4 Macro Reference: https://xlladdins.github.io/Excel4Macros/
- xll24 Repository: https://github.com/keithalewis/xll
- Modern CMake Guide: https://cliutils.gitlab.io/modern-cmake/

## Important Notes for Claude

- **NEVER** create markdown documentation files unless explicitly requested
- **ALWAYS** use the Read tool before editing existing files
- **ALWAYS** prefer editing existing files over creating new ones
- **ALWAYS** use CMake for building (never use MSBuild directly)
- **ALWAYS** verify function signature matches Excel expectations (WINAPI, #pragma XLLEXPORT)
- When adding functions, follow the pattern shown in xll_template.cpp exactly
- The xll.props file configures Visual Studio debugging, but CMake handles actual compilation
