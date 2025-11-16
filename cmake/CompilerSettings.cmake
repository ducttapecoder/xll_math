# ==============================================================================
# CompilerSettings.cmake - Core Compiler Configuration
# ==============================================================================
# Modern compiler configuration for C++23 and latest toolchains
#
# Requirements:
#   - CMake 3.28+
#   - MSVC 19+ (Visual Studio 2026)
#   - GCC 15+ (MSYS2 UCRT64/MINGW32)
#   - Clang 21+ (MSYS2 CLANG64/CLANG32)
#   - Windows 10/11 (x64 or x86)
#
# Features:
#   - C++23 standard
#   - Static runtime linking
#   - Link-Time Optimization (LTO/IPO)
#   - Modern optimizations
#   - Security features (SDL, CFG)
#   - Helper functions
#
# Note: Warnings configured in CompilerWarnings.cmake
# ==============================================================================

# ==============================================================================
# Compiler Detection
# ==============================================================================

# Detect compiler type
if(CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
    set(COMPILER_MSVC TRUE)
    message(STATUS "Detected MSVC compiler")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
    set(COMPILER_GCC TRUE)
    message(STATUS "Detected GCC compiler")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    set(COMPILER_CLANG TRUE)
    message(STATUS "Detected Clang compiler")
else()
    message(WARNING "Unknown compiler: ${CMAKE_CXX_COMPILER_ID}")
endif()

# Detect MSYS2 environment
if(DEFINED ENV{MSYSTEM})
    set(MSYS2_ENV $ENV{MSYSTEM})

    if(MSYS2_ENV STREQUAL "UCRT64")
        set(MSYS2_UCRT64 TRUE)
        message(STATUS "MSYS2 UCRT64 environment detected")
    elseif(MSYS2_ENV STREQUAL "MINGW32")
        set(MSYS2_MINGW32 TRUE)
        message(STATUS "MSYS2 MINGW32 environment detected")
    elseif(MSYS2_ENV STREQUAL "CLANG64")
        set(MSYS2_CLANG64 TRUE)
        message(STATUS "MSYS2 CLANG64 environment detected")
    elseif(MSYS2_ENV STREQUAL "CLANG32")
        set(MSYS2_CLANG32 TRUE)
        message(STATUS "MSYS2 CLANG32 environment detected")
    elseif(MSYS2_ENV STREQUAL "MSYS")
        # MSYS shell, likely using MSVC - that's OK
        message(STATUS "MSYS shell environment detected")
    else()
        message(WARNING "Unknown MSYS2 environment: ${MSYS2_ENV}")
    endif()
endif()

# ==============================================================================
# Global Compiler Flags
# ==============================================================================

# Unicode Support (Windows wide character set)
add_compile_definitions(UNICODE _UNICODE)

# UTF-8 source and execution encoding
if(COMPILER_MSVC)
    add_compile_options(/utf-8)
elseif(COMPILER_GCC OR COMPILER_CLANG)
    add_compile_options(-finput-charset=UTF-8 -fexec-charset=UTF-8)
endif()

# Make 'char' unsigned by default (MSVC /J flag behavior)
if(COMPILER_MSVC)
    add_compile_options(/J)
elseif(COMPILER_GCC OR COMPILER_CLANG)
    add_compile_options(-funsigned-char)
endif()

# ==============================================================================
# Standards Conformance
# ==============================================================================

if(COMPILER_MSVC)
    # Conformance mode (permissive- for strict standards compliance)
    # Note: Use permissive- in Release for maximum compatibility
    # Use permissive (without -) in Debug if needed for faster compilation
    add_compile_options(
        $<$<CONFIG:Release>:/permissive->   # Strict conformance in release
    )
endif()

# ==============================================================================
# Runtime Library Linkage
# ==============================================================================

if(COMPILER_MSVC)
    # Static MSVC runtime library (/MT for Release, /MTd for Debug)
    set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

elseif(COMPILER_GCC OR COMPILER_CLANG)
    # Static link GCC/Clang runtime libraries
    add_link_options(-static-libgcc -static-libstdc++)

    # Note: UCRT64/CLANG64 environments use static UCRT by default
    if(MSYS2_UCRT64 OR MSYS2_CLANG64)
        message(STATUS "Using static UCRT runtime")
    endif()
endif()

# ==============================================================================
# Optimization Flags
# ==============================================================================

if(COMPILER_MSVC)
    # MSVC optimization
    add_compile_options(
        $<$<CONFIG:Release>:/O2>   # Maximize speed
        $<$<CONFIG:Release>:/Oi>   # Enable intrinsics
        $<$<CONFIG:Release>:/GL>   # Whole program optimization
    )
    add_link_options(
        $<$<CONFIG:Release>:/LTCG>     # Link-time code generation
        $<$<CONFIG:Release>:/OPT:REF>  # Remove unreferenced functions
        $<$<CONFIG:Release>:/OPT:ICF>  # COMDAT folding
    )

elseif(COMPILER_GCC OR COMPILER_CLANG)
    # GCC/Clang optimization
    add_compile_options(
        $<$<CONFIG:Release>:-O3>      # Maximum optimization
        $<$<CONFIG:Release>:-DNDEBUG> # Disable assertions
    )

    # Link-Time Optimization (LTO/IPO)
    if(NOT CMAKE_BUILD_TYPE STREQUAL "Debug")
        include(CheckIPOSupported)
        check_ipo_supported(RESULT IPO_SUPPORTED OUTPUT IPO_ERROR)

        if(IPO_SUPPORTED)
            message(STATUS "LTO/IPO is supported and will be enabled for Release builds")
        else()
            message(STATUS "LTO/IPO not supported: ${IPO_ERROR}")
        endif()
    endif()
endif()

# ==============================================================================
# Debug Information
# ==============================================================================

if(COMPILER_MSVC)
    # Always generate debug information (even in Release for better diagnostics)
    add_compile_options(/Zi)        # Generate debug info
    add_link_options(/DEBUG)        # Include debug info in binaries

elseif(COMPILER_GCC OR COMPILER_CLANG)
    add_compile_options(
        $<$<CONFIG:Debug>:-g>       # Debug symbols
        $<$<CONFIG:Debug>:-O0>      # No optimization in debug
    )
endif()

# ==============================================================================
# Security Features
# ==============================================================================

if(COMPILER_MSVC)
    # Security Development Lifecycle (SDL) checks
    add_compile_options(/sdl)

    # Note: Additional security features available:
    # - /guard:cf (Control Flow Guard) - can be enabled per-project
    # - /GS (Buffer Security Check) - enabled by default
endif()

# ==============================================================================
# Windows Platform Settings
# ==============================================================================

# Windows 10/11 (x64 or x86)
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    add_compile_definitions(WIN32 _WIN64)
else()
    add_compile_definitions(WIN32)
endif()

# Target Windows 10+
add_compile_definitions(
    WINVER=0x0A00          # Windows 10
    _WIN32_WINNT=0x0A00    # Windows 10
)

# ==============================================================================
# Helper Functions
# ==============================================================================

#[========================================[.rst:
apply_compiler_settings
-----------------------

Applies common compiler settings to a target.

.. code-block:: cmake

  apply_compiler_settings(<target>)

Configures:
- Link-Time Optimization (if supported and not Debug build)
- Note: LTO disabled for GCC due to .def file compatibility issues

]========================================]
function(apply_compiler_settings TARGET_NAME)
    # Enable LTO for Release builds if supported
    # Note: Disabled for GCC due to linker issues with .def file exports
    if(IPO_SUPPORTED AND NOT CMAKE_BUILD_TYPE STREQUAL "Debug" AND NOT COMPILER_GCC)
        set_target_properties(${TARGET_NAME} PROPERTIES
            INTERPROCEDURAL_OPTIMIZATION_RELEASE TRUE
        )
        message(STATUS "LTO enabled for ${TARGET_NAME} (Release builds)")
    endif()
endfunction()

#[========================================[.rst:
link_whole_archive
------------------

Links a static library with whole-archive semantics to ensure all symbols are included.

.. code-block:: cmake

  link_whole_archive(<target> <library>)

This is essential for Excel XLL add-ins to ensure xlAuto* functions from the
static library are included even if not directly referenced.

]========================================]
function(link_whole_archive TARGET_NAME LIBRARY_NAME)
    if(COMPILER_MSVC)
        # MSVC: Link library first, then apply /WHOLEARCHIVE
        # We need to link the library so the linker knows where to find it,
        # then use /WHOLEARCHIVE to include all symbols
        target_link_libraries(${TARGET_NAME} PRIVATE ${LIBRARY_NAME})

        get_target_property(LIB_OUTPUT_NAME ${LIBRARY_NAME} OUTPUT_NAME)
        if(NOT LIB_OUTPUT_NAME)
            set(LIB_OUTPUT_NAME ${LIBRARY_NAME})
        endif()

        # Apply /WHOLEARCHIVE to this specific library
        target_link_options(${TARGET_NAME} PRIVATE "/WHOLEARCHIVE:${LIB_OUTPUT_NAME}")

    elseif(COMPILER_GCC OR COMPILER_CLANG)
        # GCC/Clang: Use --whole-archive wrapper
        target_link_libraries(${TARGET_NAME} PRIVATE
            -Wl,--whole-archive
            ${LIBRARY_NAME}
            -Wl,--no-whole-archive
        )
    endif()
endfunction()

#[========================================[.rst:
set_windows_subsystem
---------------------

Sets the Windows subsystem for an executable or DLL.

.. code-block:: cmake

  set_windows_subsystem(<target> CONSOLE|WINDOWS)

WINDOWS subsystem: GUI application (no console)
CONSOLE subsystem: Console application

]========================================]
function(set_windows_subsystem TARGET_NAME SUBSYSTEM)
    if(COMPILER_MSVC)
        target_link_options(${TARGET_NAME} PRIVATE /SUBSYSTEM:${SUBSYSTEM})
    elseif(COMPILER_GCC OR COMPILER_CLANG)
        if(SUBSYSTEM STREQUAL "WINDOWS")
            target_link_options(${TARGET_NAME} PRIVATE -mwindows)
        elseif(SUBSYSTEM STREQUAL "CONSOLE")
            target_link_options(${TARGET_NAME} PRIVATE -mconsole)
        endif()
    endif()
endfunction()

# ==============================================================================
# Compiler-Specific Diagnostics
# ==============================================================================

if(COMPILER_MSVC)
    message(STATUS "MSVC-specific optimizations enabled")
    # Security Development Lifecycle checks enabled via /sdl

elseif(COMPILER_GCC)
    message(STATUS "GCC-specific settings applied")
    # LTO disabled for GCC due to .def file compatibility

elseif(COMPILER_CLANG)
    message(STATUS "Clang-specific settings applied")
endif()
