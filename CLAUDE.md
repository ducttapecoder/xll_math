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

### Excel Data Types

**XLL Types** (defined in `xll24/include/type.h`):
- `XLL_DOUBLE` - double precision floating point
- `XLL_LPXLOPER` - OPER object (arrays, strings, ranges, etc.)
- `XLL_CSTRING` - const wchar_t* (null-terminated Unicode string)
- `XLL_WORD` - 16-bit unsigned integer
- `XLL_DWORD` - 32-bit unsigned integer
- `XLL_FP` - FP12 array structure (2D array of doubles)
- `XLL_BOOL` - boolean value

### Working with OPER

OPER wraps Excel's XLOPER12 structure:

```cpp
// Construction
OPER num(42.0);                    // Number
OPER str(L"Hello");                // String
OPER arr(2, 3);                    // 2x3 array
OPER range = Excel(xlfActiveCell); // Get active cell

// Access
double value = oper.val.num;       // Number value
const wchar_t* text = oper.val.str;// String value

// Type checking
if (oper.xltype == xltypeNum) { /* ... */ }
if (oper.xltype == xltypeStr) { /* ... */ }

// Helper functions
double Num(const OPER& x);         // Get number value
const wchar_t* Str(const OPER& x); // Get string value
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
