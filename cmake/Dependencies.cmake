# ==============================================================================
# Dependencies.cmake - External Dependencies Configuration
# ==============================================================================
# Manages external library dependencies for xll_math project
#
# Current dependencies:
#   - Eigen3 5.0+ : Header-only C++ linear algebra library
# ==============================================================================

# ==============================================================================
# Eigen3 - Linear Algebra Library
# ==============================================================================

# Try to find Eigen3 5.0 or later
find_package(Eigen3 5.0 QUIET NO_MODULE)

if(Eigen3_FOUND)
    message(STATUS "Eigen3 found: ${Eigen3_VERSION}")
    message(STATUS "Eigen3 include dir: ${EIGEN3_INCLUDE_DIR}")
else()
    message(STATUS "Eigen3 5.0+ not found via find_package")

    # Try to find via vcpkg
    if(DEFINED ENV{VCPKG_ROOT})
        message(STATUS "Checking vcpkg installation...")

        # Add vcpkg paths based on architecture
        if(CMAKE_SIZEOF_VOID_P EQUAL 8)
            set(VCPKG_ARCH "x64-windows")
        else()
            set(VCPKG_ARCH "x86-windows")
        endif()

        list(APPEND CMAKE_PREFIX_PATH "$ENV{VCPKG_ROOT}/installed/${VCPKG_ARCH}")

        # Try again with vcpkg path
        find_package(Eigen3 5.0 QUIET NO_MODULE)

        if(Eigen3_FOUND)
            message(STATUS "Eigen3 found via vcpkg: ${Eigen3_VERSION}")
        endif()
    endif()

    # If still not found, provide installation instructions
    if(NOT Eigen3_FOUND)
        message(FATAL_ERROR
            "Eigen3 5.0+ is required but not found.\n"
            "\n"
            "Installation options:\n"
            "  1. Using vcpkg (recommended):\n"
            "       vcpkg install eigen3:x64-windows\n"
            "       vcpkg install eigen3:x86-windows\n"
            "\n"
            "  2. Manual download:\n"
            "       Download from: https://gitlab.com/libeigen/eigen/-/releases/5.0.0\n"
            "       Extract and set CMAKE_PREFIX_PATH or Eigen3_DIR\n"
            "\n"
            "  3. System package manager:\n"
            "       MSYS2: pacman -S mingw-w64-x86_64-eigen3\n"
        )
    endif()
endif()

# Verify Eigen3::Eigen target is available
if(NOT TARGET Eigen3::Eigen)
    message(FATAL_ERROR "Eigen3 package found but Eigen3::Eigen target not available")
endif()

message(STATUS "Eigen3 configuration complete")
