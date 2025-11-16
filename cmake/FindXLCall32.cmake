# FindXLCall32.cmake - Locate Excel SDK XLCall32 library
#
# This module defines:
#   XLCall32_FOUND - System has XLCall32
#   XLCall32_LIBRARY - The XLCall32 library
#
# The module handles different library formats for different compilers:
#   - MSVC/Clang: XLCALL32.LIB (MSVC import library format)
#   - GCC: libxlcall32.a (GCC import library format)
#
# Usage:
#   find_package(XLCall32 REQUIRED)
#   target_link_libraries(mytarget PRIVATE XLCall32::XLCall32)

# Determine library name based on compiler
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
    # GCC requires GCC-compatible import library
    set(_XLCALL32_NAMES xlcall32 libxlcall32.a)
else()
    # MSVC and Clang can use MSVC library format
    set(_XLCALL32_NAMES XLCALL32 xlcall32)
endif()

# Search for the library
find_library(XLCall32_LIBRARY
    NAMES ${_XLCALL32_NAMES}
    PATHS
        ${CMAKE_CURRENT_SOURCE_DIR}/xll24/lib
        ${CMAKE_CURRENT_SOURCE_DIR}/lib
        ${CMAKE_SOURCE_DIR}/xll24/lib
        ${CMAKE_SOURCE_DIR}/lib
    DOC "Excel SDK XLCall32 library"
    NO_DEFAULT_PATH
)

# Handle standard find_package arguments
include(FindPackageHandleStandardArgs)
find_package_handle_standard_args(XLCall32
    REQUIRED_VARS XLCall32_LIBRARY
    FAIL_MESSAGE "XLCall32 library not found. Place XLCALL32.LIB (or libxlcall32.a for GCC) in xll24/lib directory."
)

# Create imported target
if(XLCall32_FOUND AND NOT TARGET XLCall32::XLCall32)
    add_library(XLCall32::XLCall32 UNKNOWN IMPORTED)
    set_target_properties(XLCall32::XLCall32 PROPERTIES
        IMPORTED_LOCATION "${XLCall32_LIBRARY}"
    )

    mark_as_advanced(XLCall32_LIBRARY)
endif()
