# ==============================================================================
# xll24 Static Library
# ==============================================================================
# Framework library for building Excel XLL add-ins
#
# Requirements:
#   - CMake 3.28+
#   - C++23 compiler
#   - Windows 10/11 (x64 or x86)
#
# This library provides:
#   - Excel C API wrapper (OPER class)
#   - Automatic function registration
#   - Type conversions
#   - Error handling
#   - Excel macro functions
# ==============================================================================

cmake_minimum_required(VERSION 3.28 FATAL_ERROR)

# ==============================================================================
# Source Files
# ==============================================================================

# Header files (public API)
set(XLL24_PUBLIC_HEADERS
    include/addin.h
    include/alert.h
    include/args.h
    include/auto.h
    include/defines.h
    include/ensure.h
    include/enum.h
    include/excel.h
    include/excel_time.h
    include/export.h
    include/fp.h
    include/fpx.h
    include/handle.h
    include/macrofun.h
    include/on.h
    include/oper.h
    include/ref.h
    include/register.h
    include/type.h
    include/utf8.h
    include/win_mem_view.h
    include/XLCALL.H
    include/xll.h
    include/xll_mem_oper.h
    include/xloper.h
)

# Implementation files
set(XLL24_SOURCES
    src/addin.cpp
    src/alert.cpp
    src/debug.cpp
    src/depends.cpp
    src/dllmain.cpp
    src/doevents.cpp
    src/evaluate.cpp
    src/fpx.c
    src/paste.cpp
    src/py.cpp
    src/range.cpp
    src/xlauto.cpp
    src/XLCALL.CPP
)

# ==============================================================================
# Library Target
# ==============================================================================

# Create static library
add_library(xll24 STATIC
    ${XLL24_PUBLIC_HEADERS}
    ${XLL24_SOURCES}
)

# Create alias for namespaced usage
add_library(xll::xll24 ALIAS xll24)

# ==============================================================================
# Include Directories
# ==============================================================================

target_include_directories(xll24
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include/xll24>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# ==============================================================================
# Dependencies
# ==============================================================================

# Find and link Excel SDK library
find_package(XLCall32 REQUIRED)
target_link_libraries(xll24 PUBLIC XLCall32::XLCall32)

# ==============================================================================
# Compiler Settings
# ==============================================================================

# Apply cross-compiler settings
apply_compiler_settings(xll24)

# Set Windows subsystem
set_windows_subsystem(xll24 WINDOWS)

# Apply compiler warnings
# Note: WARNINGS_AS_ERRORS only for MSVC due to Excel C API legacy code patterns
if(CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
    set_project_warnings(xll24 WARNINGS_AS_ERRORS)
else()
    set_project_warnings(xll24)
endif()

# GCC/Clang-specific settings
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    # Disable specific warnings that conflict with Excel C API and legacy code patterns
    target_compile_options(xll24 PRIVATE
        -Wno-unknown-pragmas                             # Ignore MSVC-specific pragmas
        -Wno-missing-field-initializers                  # Allow partial initializers
        -Wno-shadow                                      # Constructor params often shadow members
        -Wno-old-style-cast                              # Excel C API requires C-style casts
        -Wno-sign-conversion                             # Excel API uses mixed signed/unsigned
        -Wno-sign-compare                                # Excel API uses mixed signed/unsigned
    )
endif()

# Clang-specific warnings
if(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    target_compile_options(xll24 PRIVATE
        -Wno-macro-redefined                             # Handle NOMINMAX and other macro redefs
    )
endif()

# GCC-specific warnings
if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    target_compile_options(xll24 PRIVATE
        -Wno-useless-cast                                # GCC is overly strict about casts
    )
endif()

# Preprocessor definitions
target_compile_definitions(xll24 PRIVATE
    $<$<CONFIG:Debug>:_DEBUG>
    $<$<CONFIG:Release>:NDEBUG>
)

# ==============================================================================
# IDE Organization
# ==============================================================================

# Group files in Visual Studio/IDEs
source_group("Header Files" FILES ${XLL24_PUBLIC_HEADERS})
source_group("Source Files" FILES ${XLL24_SOURCES})

# ==============================================================================
# Output Configuration
# ==============================================================================

# Set library output name (xll.lib instead of xll24.lib for backward compatibility)
set_target_properties(xll24 PROPERTIES
    OUTPUT_NAME "xll"
    ARCHIVE_OUTPUT_NAME "xll"
)

# ==============================================================================
# Installation
# ==============================================================================

# Install library
install(TARGETS xll24
    EXPORT xll24Targets
    ARCHIVE DESTINATION lib
    INCLUDES DESTINATION include/xll24
)

# Install headers
install(FILES ${XLL24_PUBLIC_HEADERS}
    DESTINATION include/xll24
)

# Export targets for downstream projects
install(EXPORT xll24Targets
    FILE xll24Targets.cmake
    NAMESPACE xll::
    DESTINATION lib/cmake/xll24
)

# ==============================================================================
# Package Config
# ==============================================================================

include(CMakePackageConfigHelpers)

# Create version file
write_basic_package_version_file(
    "${CMAKE_CURRENT_BINARY_DIR}/xll24ConfigVersion.cmake"
    VERSION 1.0.0
    COMPATIBILITY SameMajorVersion
)

# Create config file
configure_package_config_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake/xll24Config.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/xll24Config.cmake"
    INSTALL_DESTINATION lib/cmake/xll24
)

# Install package files
install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/xll24Config.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/xll24ConfigVersion.cmake"
    DESTINATION lib/cmake/xll24
)
