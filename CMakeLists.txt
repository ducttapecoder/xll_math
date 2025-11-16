# ==============================================================================
# xll_math - Excel Add-in (XLL) Project
# ==============================================================================
# Modern CMake configuration for building Excel XLL add-ins using xll24 framework
#
# Requirements:
#   - CMake 3.28+
#   - C++23 compiler
#   - Visual Studio 2026 (MSVC 19+) with v145 toolset
#   - GCC 15+ (MSYS2 UCRT64/MINGW32)
#   - Clang 21+ (MSYS2 CLANG64/CLANG32)
#   - Windows 10/11 (x64 or x86)
#
# Quick Start:
#   cmake --preset <preset-name>
#   cmake --build --preset <preset-name>
#
# Available presets:
#   MSVC: debug, release, multi, debug-x86, release-x86
#   GCC:  ucrt64-debug, ucrt64-release, mingw32-debug, mingw32-release
#   Clang: clang64-debug, clang64-release, clang32-debug, clang32-release
# ==============================================================================

cmake_minimum_required(VERSION 3.28 FATAL_ERROR)

# ==============================================================================
# Project Definition
# ==============================================================================

project(xll_math
    VERSION 1.0.0
    DESCRIPTION "Excel Add-in (XLL) project using xll24 framework"
    HOMEPAGE_URL "https://github.com/keithalewis/xll"
    LANGUAGES C CXX
)

# ==============================================================================
# Platform Requirements
# ==============================================================================

# Windows 10/11 only (x64 or x86)
if(NOT WIN32)
    message(FATAL_ERROR "This project requires Windows 10 or Windows 11")
endif()

# Detect architecture
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    message(STATUS "Building for Windows x64")
    set(ARCH_64BIT TRUE)
elseif(CMAKE_SIZEOF_VOID_P EQUAL 4)
    message(STATUS "Building for Windows x86 (32-bit)")
    set(ARCH_32BIT TRUE)
else()
    message(FATAL_ERROR "Unsupported architecture")
endif()

# Compiler version requirements
if(CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS "19.0")
        message(FATAL_ERROR "MSVC 19+ required (Visual Studio 2026). Found: ${CMAKE_CXX_COMPILER_VERSION}")
    endif()
    message(STATUS "MSVC version: ${CMAKE_CXX_COMPILER_VERSION}")
elseif(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS "15.0")
        message(FATAL_ERROR "GCC 15+ required. Found: ${CMAKE_CXX_COMPILER_VERSION}")
    endif()
    message(STATUS "GCC version: ${CMAKE_CXX_COMPILER_VERSION}")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS "21.0")
        message(FATAL_ERROR "Clang 21+ required. Found: ${CMAKE_CXX_COMPILER_VERSION}")
    endif()
    message(STATUS "Clang version: ${CMAKE_CXX_COMPILER_VERSION}")
else()
    message(FATAL_ERROR "Unsupported compiler: ${CMAKE_CXX_COMPILER_ID}")
endif()

# ==============================================================================
# CMake Policies and Settings
# ==============================================================================

# Use modern CMake policies
if(POLICY CMP0091)
    cmake_policy(SET CMP0091 NEW) # MSVC runtime library selection
endif()
if(POLICY CMP0092)
    cmake_policy(SET CMP0092 NEW) # Don't add /W3 by default for MSVC
endif()

# Language standards
set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

set(CMAKE_C_STANDARD 17)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_C_EXTENSIONS OFF)

# Module path for custom CMake modules
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

# ==============================================================================
# Compiler Configuration
# ==============================================================================

# Include cross-compiler settings and warning configuration
include(CompilerSettings)
include(CompilerWarnings)

# ==============================================================================
# Output Directories
# ==============================================================================

# Single-configuration generators (Ninja, Makefiles)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin")
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib")
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib")

# Multi-configuration generators (Visual Studio)
foreach(CONFIG IN LISTS CMAKE_CONFIGURATION_TYPES)
    string(TOUPPER "${CONFIG}" CONFIG_UPPER)
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_${CONFIG_UPPER} "${CMAKE_BINARY_DIR}/bin/${CONFIG}")
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY_${CONFIG_UPPER} "${CMAKE_BINARY_DIR}/lib/${CONFIG}")
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY_${CONFIG_UPPER} "${CMAKE_BINARY_DIR}/lib/${CONFIG}")
endforeach()

# ==============================================================================
# Dependencies
# ==============================================================================

# xll24 static library (framework)
add_subdirectory(xll24)

# ==============================================================================
# xll_template Target - Main Excel Add-in
# ==============================================================================

# Source files
set(XLL_TEMPLATE_SOURCES
    xll_template.cpp
    xll_template.h
)

# For GCC/Clang, include .def file for function exports
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    list(APPEND XLL_TEMPLATE_SOURCES xll_template.def)
endif()

# Create shared library (.xll)
add_library(xll_template SHARED ${XLL_TEMPLATE_SOURCES})

# Set target properties
set_target_properties(xll_template PROPERTIES
    OUTPUT_NAME "xll_template"
    SUFFIX ".xll"
)

# Apply compiler settings
apply_compiler_settings(xll_template)

# Apply compiler warnings (without WARNINGS_AS_ERRORS for application code)
set_project_warnings(xll_template)

# Link against xll24 with whole-archive to include all static library symbols
# This ensures xlAuto* functions are included even if not directly referenced
link_whole_archive(xll_template xll24)

# Compile definitions
target_compile_definitions(xll_template PRIVATE
    $<$<CONFIG:Debug>:_DEBUG>
    $<$<CONFIG:Release>:NDEBUG>
    _WINDOWS
    _USRDLL
)

# Windows subsystem
set_windows_subsystem(xll_template WINDOWS)

# ==============================================================================
# Visual Studio Debugger Configuration
# ==============================================================================

if(MSVC_IDE)
    # Configure debugger to launch Excel with the XLL loaded
    # Excel path can be set in CMakeUserPresets.json with CMAKE_VS_DEBUGGER_COMMAND
    set_target_properties(xll_template PROPERTIES
        VS_DEBUGGER_COMMAND_ARGUMENTS "/x \"$<TARGET_FILE:xll_template>\""
    )
endif()

# ==============================================================================
# Installation Rules
# ==============================================================================

# Install the XLL add-in
install(TARGETS xll_template
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

# Install development headers
install(DIRECTORY xll24/include/
    DESTINATION include/xll24
    FILES_MATCHING
        PATTERN "*.h"
        PATTERN "*.H"
)

# ==============================================================================
# Package Configuration
# ==============================================================================

# Allow the project to be used via find_package()
include(CMakePackageConfigHelpers)

# Create package version file
write_basic_package_version_file(
    "${CMAKE_CURRENT_BINARY_DIR}/xll_mathConfigVersion.cmake"
    VERSION ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)

# Install package files
install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/xll_mathConfigVersion.cmake"
    DESTINATION lib/cmake/xll_math
)

# ==============================================================================
# Status Messages
# ==============================================================================

message(STATUS "")
message(STATUS "=================================")
message(STATUS "xll_math Configuration Summary")
message(STATUS "=================================")
message(STATUS "Version: ${PROJECT_VERSION}")
message(STATUS "Compiler: ${CMAKE_CXX_COMPILER_ID} ${CMAKE_CXX_COMPILER_VERSION}")
message(STATUS "Build Type: ${CMAKE_BUILD_TYPE}")
message(STATUS "C++ Standard: ${CMAKE_CXX_STANDARD}")
message(STATUS "Install Prefix: ${CMAKE_INSTALL_PREFIX}")
message(STATUS "=================================")
message(STATUS "")
