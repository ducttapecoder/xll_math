# xll_math - Excel Add-in Project

Modern C++23 Excel XLL add-in using the xll24 framework and CMake build system.

## Requirements

**Platform:**
- Windows 10 or Windows 11 (x64 or x86)

**Build Tools:**
- CMake 3.28+
- Ninja (recommended) or Visual Studio

**Compiler** (choose one):
- Visual Studio 2026 (MSVC 19+, v145 toolset)
- GCC 15+ (MSYS2 UCRT64 for x64, MINGW32 for x86)
- Clang 21+ (MSYS2 CLANG64 for x64, CLANG32 for x86)

**Language:**
- C++23
- C23

## Quick Start

### Build with CMake

``` Powershell
# MSVC (default)
clear ; cmake --preset release ; cmake --build --preset release
```

```bash
# GCC (MSYS2 UCRT64)
clear && cmake --preset ucrt64-release && cmake --build --preset ucrt64-release

# Clang (MSYS2 CLANG64)
clear && cmake --preset clang64-release && cmake --build --preset clang64-release
```

### Load in Excel

1. Build the project
2. Open Excel
3. File → Options → Add-ins → Manage Excel Add-ins → Go
4. Browse to `build/*/bin/*/xll_template.xll`
5. Check the add-in and click OK

Your custom functions are now available in Excel!

## Available Build Presets

### MSVC (Visual Studio 2026)
- `debug` - x64 Debug
- `release` - x64 Release
- `multi` - x64 Multi-configuration (both Debug and Release)
- `debug-x86` - x86 Debug
- `release-x86` - x86 Release

### GCC (MSYS2)
- `ucrt64-debug` - UCRT64 (x64) Debug
- `ucrt64-release` - UCRT64 (x64) Release
- `mingw32-debug` - MinGW32 (x86) Debug
- `mingw32-release` - MinGW32 (x86) Release

### Clang (MSYS2)
- `clang64-debug` - Clang64 (x64) Debug
- `clang64-release` - Clang64 (x64) Release
- `clang32-debug` - Clang32 (x86) Debug
- `clang32-release` - Clang32 (x86) Release

## Development

### Visual Studio

Open the project folder in Visual Studio 2022:
1. File → Open → Folder → select xll_math directory
2. Visual Studio will detect CMake configuration
3. Select a preset from the dropdown
4. Press F5 to build and debug (automatically launches Excel with XLL loaded)

### Adding Excel Functions

#### Quick Example

1. Edit `xll_template.cpp`
2. Add function registration and implementation:

```cpp
AddIn xai_myfunction(
    Function(XLL_DOUBLE, "xll_myfunction", "MY.FUNCTION")
    .Arguments({
        Arg(XLL_DOUBLE, "x", "is the input.")
    })
    .FunctionHelp("Returns x squared.")
    .Category("MATH")
);

// For GCC/Clang compatibility
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_myfunction(double x)
{
#pragma XLLEXPORT
    return x * x;
}
```

3. Rebuild and reload in Excel
4. Use `=MY.FUNCTION(5)` in any cell

#### Supported Data Types

The xll24 framework supports various Excel data types:

**Common Types:**
- `XLL_DOUBLE` - Double precision numbers
- `XLL_FP` - 2D arrays of doubles (ranges)
- `XLL_LPXLOPER` - Generic Excel values (any type)
- `XLL_BOOL` - Boolean values
- `XLL_CSTRING` - Unicode strings
- `XLL_HANDLE` - Handles to C++ objects

**Example with Arrays:**

```cpp
AddIn xai_array_sum(
    Function(XLL_DOUBLE, "xll_array_sum", "ARRAY.SUM")
    .Arguments({Arg(XLL_FP, "array", "is array of numbers.")})
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_array_sum(const _FP12* pa)
{
#pragma XLLEXPORT
    double sum = 0;
    for (int i = 0; i < pa->rows * pa->columns; ++i)
        sum += pa->array[i];
    return sum;
}
```

**See [CLAUDE.md](CLAUDE.md) for:**
- Complete type reference with examples
- Working with arrays and matrices
- Optional arguments
- Object handles
- Advanced patterns

## Project Structure

```
xll_math/
├── CMakeLists.txt              # Root CMake configuration
├── xll_template.cpp            # Your Excel functions
├── xll_template.h              # Function declarations
├── cmake/
│   ├── CompilerSettings.cmake  # Compiler configuration
│   ├── CompilerWarnings.cmake  # Warning settings
│   └── FindXLCall32.cmake      # Excel SDK library finder
├── xll24/                      # XLL framework (submodule)
│   ├── include/                # Framework headers
│   ├── src/                    # Framework implementation
│   └── lib/                    # Excel SDK library
└── build/                      # Build output (generated)
```

## Configuration

All modern compiler settings are configured via CMake:
- C++23 standard
- Static runtime linking
- Link-Time Optimization (Release)
- Unicode support
- UTF-8 encoding
- Comprehensive warnings
- SDL security checks

## Documentation

- [CLAUDE.md](CLAUDE.md) - Comprehensive development guide
- Excel SDK: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk
- xll24 Framework: https://github.com/keithalewis/xll

## License

See LICENSE file for details.
