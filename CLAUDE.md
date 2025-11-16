# CLAUDE.md

Excel add-in (XLL) project with C++23 using xll24 framework and Eigen 5.0+ for matrix operations.

## Requirements

- CMake 3.28+, Ninja
- MSVC 19+ / GCC 15+ / Clang 21+
- Windows 10/11 (x64/x86)
- C++23 standard
- Eigen 5.0+ (header-only, for linear algebra)

## Project Structure

```
xll_math/
├── include/          # core.h, linalg.h
├── src/              # core.cpp, linalg.cpp, exports.def
├── cmake/            # CompilerSettings, Warnings, Dependencies, FindXLCall32
├── xll24/            # Framework submodule
└── CMakeLists.txt
```

**Targets:** xll24 (static .lib) + xll_math (shared .xll)

## Quick Start

```bash
# Install Eigen
pacman -S mingw-w64-x86_64-eigen3  # MSYS2
vcpkg install eigen3:x64-windows   # vcpkg

# Build
cmake --preset release              # MSVC x64
cmake --preset ucrt64-release       # GCC
cmake --preset clang64-release      # Clang
cmake --build --preset release

# Output: build/msvc-x64-release/bin/Release/xll_math.xll
```

**Presets:** debug, release, multi, debug-x86, release-x86 (MSVC) | ucrt64-*, mingw32-* (GCC) | clang64-*, clang32-* (Clang)

## Adding Excel Functions

**1. Define function:**
```cpp
AddIn xai_myfunction(
    Function(XLL_DOUBLE, "xll_myfunction", "MY.FUNCTION")
    .Arguments({Arg(XLL_DOUBLE, "x", "is the input value.")})
    .FunctionHelp("Computes result from x.")
    .Category("MATH")
);
```

**2. Implement:**
```cpp
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_myfunction(double x) {
#pragma XLLEXPORT
    return x * x;
}
```

**3. For GCC/Clang:** Add `xll_myfunction` to `src/exports.def`

**4. Update CMakeLists.txt** if adding new source files

## Linear Algebra Functions (22 total)

**Installation:** `pacman -S mingw-w64-x86_64-eigen3` or `vcpkg install eigen3:x64-windows`

**Basic (8):** ADD, SUB, MUL, TRANSPOSE, TRACE, NORM, DETERMINANT, RANK

**Decompositions (8):** INVERSE, LU, QR, CHOLESKY, SVD, SVD_FULL, EIGENVALUES, EIGENVECTORS

**Solvers (3):** SOLVE, LSTSQ, PSEUDO_INV

**Utilities (3):** IDENTITY, ZEROS, DIAG

**Examples:**
```excel
=MATRIX.MUL(A1:B2, C1:D2)
=MATRIX.EIGENVALUES(A1:D4)
=MATRIX.SVD_FULL(A1:C4)
```

## Excel Data Types

| XLL Type | C++ Type | Use |
|----------|----------|-----|
| XLL_DOUBLE | double | Single number |
| XLL_FP | _FP12* | Array of doubles |
| XLL_LPXLOPER | LPXLOPER12 | Any Excel value |
| XLL_HANDLE | HANDLEX | Object handle |
| XLL_CSTRING | const wchar_t* | Unicode string |

## Project Organization

**Core functions:** `src/core.cpp`, `include/core.h` (TGAMMA, ABOUT, macros)

**Linear algebra:** `src/linalg.cpp`, `include/linalg.h` (22 matrix operations, requires Eigen)

**New modules:** Create `src/modulename.cpp`, `include/modulename.h`, update CMakeLists.txt

**Steps:**
1. Add declaration to `include/*.h`
2. Implement in `src/*.cpp`
3. For GCC/Clang: Add to `src/exports.def`
4. Update CMakeLists.txt if adding new files
5. Rebuild

## Code Style

- Namespace: `namespace xll { ... }`
- Functions: `xll_function_name`, Excel: `MY.FUNCTION`, AddIn: `xai_function_name`
- Files: lowercase (core.h, linalg.h, statistics.h)
- All Excel functions: `WINAPI` + `#pragma XLLEXPORT`
- Include: `#include "linalg.h"` (quotes, not brackets)

## Common Issues

**XLL not loading:** Check .xll extension, Excel bitness (x64), dependencies

**Functions missing:** Verify `#pragma XLLEXPORT`, `WINAPI`, global AddIn object

**Build errors:** Check XLCALL32.LIB in xll24/lib/, whole-archive linking

**Eigen not found:** Install via MSYS2/vcpkg, set CMAKE_PREFIX_PATH

## Resources

- Excel SDK: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk
- Excel4 Macros: https://xlladdins.github.io/Excel4Macros/
- xll24: https://github.com/keithalewis/xll
- Eigen: https://eigen.tuxfamily.org/
- Eigen 5.0: https://gitlab.com/libeigen/eigen/-/releases/5.0.0

## Notes for Claude

- Use Read tool before editing files
- Prefer editing existing files over creating new ones
- Use CMake for building (never MSBuild directly)
- Follow patterns in `src/core.cpp` or `src/linalg.cpp`
- Add new functions to appropriate modules (core/linalg) or create new modules
- For GCC/Clang: Remember `src/exports.def`
- NEVER create markdown docs unless requested
