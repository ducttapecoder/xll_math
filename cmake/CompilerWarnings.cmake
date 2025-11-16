# ==============================================================================
# CompilerWarnings.cmake - Compiler Warning Configuration
# ==============================================================================
# Provides standardized warning settings for MSVC, GCC, and Clang compilers
#
# Based on cpp-best-practices:
# https://github.com/cpp-best-practices/cmake_template
#
# Features:
#   - Comprehensive warning coverage for all supported compilers
#   - Optional warnings-as-errors enforcement
#   - Customizable warning sets per target
#   - Separate warning configurations for C and C++
# ==============================================================================

#[========================================[.rst:
set_project_warnings
--------------------

Applies compiler warnings to a target with optional customization.

.. code-block:: cmake

  set_project_warnings(<target>
    [WARNINGS_AS_ERRORS]
    [MSVC_WARNINGS warning1 warning2 ...]
    [CLANG_WARNINGS warning1 warning2 ...]
    [GCC_WARNINGS warning1 warning2 ...]
  )

Arguments:
  target              - Target to apply warnings to
  WARNINGS_AS_ERRORS  - Treat all warnings as compilation errors
  MSVC_WARNINGS       - Custom MSVC warning list (overrides defaults)
  CLANG_WARNINGS      - Custom Clang warning list (overrides defaults)
  GCC_WARNINGS        - Custom GCC warning list (overrides defaults)

Example:
  set_project_warnings(mylib WARNINGS_AS_ERRORS)

]========================================]
function(set_project_warnings project_name)
    # Parse function arguments
    set(options WARNINGS_AS_ERRORS)
    set(multi_value_args MSVC_WARNINGS CLANG_WARNINGS GCC_WARNINGS)
    cmake_parse_arguments(ARG "${options}" "" "${multi_value_args}" ${ARGN})

    # ==============================================================================
    # MSVC Warning Configuration
    # ==============================================================================
    if(NOT ARG_MSVC_WARNINGS)
        set(MSVC_WARNINGS
            /W4           # High warning level (equivalent to -Wall -Wextra)

            # Specific warnings promoted to level 4
            /w14242       # Conversion from 'type1' to 'type2', possible loss of data
            /w14254       # Conversion from 'type1:field_bits' to 'type2:field_bits', possible loss of data
            /w14263       # Member function does not override any base class virtual member function
            /w14265       # Class has virtual functions, but destructor is not virtual
            /w14287       # Unsigned/negative constant mismatch
            /we4289       # Nonstandard extension used: loop control variable used outside for-loop scope
            /w14296       # Expression is always 'boolean_value'
            /w14311       # Pointer truncation from 'type1' to 'type2'
            /w14545       # Expression before comma evaluates to a function which is missing an argument list
            /w14546       # Function call before comma missing argument list
            /w14547       # Operator before comma has no effect; expected operator with side-effect
            /w14549       # Operator before comma has no effect; did you intend 'operator'?
            /w14555       # Expression has no effect; expected expression with side-effect
            /w14619       # pragma warning: there is no warning number 'number'
            /w14640       # Thread un-safe static member initialization
            /w14826       # Conversion from 'type1' to 'type2' is sign-extended
            /w14905       # Wide string literal cast to 'LPSTR'
            /w14906       # String literal cast to 'LPWSTR'
            /w14928       # Illegal copy-initialization; more than one user-defined conversion applied

            # Conformance and standards
            /permissive-  # Standards conformance mode (disable non-standard extensions)
        )
    else()
        set(MSVC_WARNINGS ${ARG_MSVC_WARNINGS})
    endif()

    # ==============================================================================
    # Clang Warning Configuration
    # ==============================================================================
    if(NOT ARG_CLANG_WARNINGS)
        set(CLANG_WARNINGS
            -Wall                    # Enable most warnings
            -Wextra                  # Enable extra warnings
            -Wpedantic               # Warn if non-standard C++ is used

            # Memory and safety
            -Wshadow                 # Warn if variable declaration shadows one from parent context
            -Wnon-virtual-dtor       # Warn if class with virtual functions has non-virtual destructor
            -Wnull-dereference       # Warn if null dereference is detected

            # Type safety
            -Wold-style-cast         # Warn for c-style casts
            -Wcast-align             # Warn for potential performance problem casts
            -Wconversion             # Warn on type conversions that may lose data
            -Wsign-conversion        # Warn on sign conversions
            -Wdouble-promotion       # Warn if float is implicitly promoted to double

            # Code quality
            -Wunused                 # Warn on anything being unused
            -Woverloaded-virtual     # Warn if you overload (not override) a virtual function
            -Wformat=2               # Warn on security issues around functions that format output
            -Wimplicit-fallthrough   # Warn on statements that fallthrough without explicit annotation
        )
    else()
        set(CLANG_WARNINGS ${ARG_CLANG_WARNINGS})
    endif()

    # ==============================================================================
    # GCC Warning Configuration
    # ==============================================================================
    if(NOT ARG_GCC_WARNINGS)
        set(GCC_WARNINGS
            ${CLANG_WARNINGS}        # Start with Clang warnings (GCC supports most of them)

            # GCC-specific warnings
            -Wmisleading-indentation # Warn if indentation implies blocks where blocks do not exist
            -Wduplicated-cond        # Warn if if/else chain has duplicated conditions
            -Wduplicated-branches    # Warn if if/else branches have duplicated code
            -Wlogical-op             # Warn about logical operations being used where bitwise were probably wanted
            -Wuseless-cast           # Warn if you perform a cast to the same type
        )
    else()
        set(GCC_WARNINGS ${ARG_GCC_WARNINGS})
    endif()

    # ==============================================================================
    # Warnings as Errors
    # ==============================================================================
    if(ARG_WARNINGS_AS_ERRORS)
        message(STATUS "Warnings are treated as errors for target: ${project_name}")
        list(APPEND CLANG_WARNINGS -Werror)
        list(APPEND GCC_WARNINGS -Werror)
        list(APPEND MSVC_WARNINGS /WX)
    endif()

    # ==============================================================================
    # Apply Warnings Based on Compiler
    # ==============================================================================
    if(CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
        set(PROJECT_WARNINGS ${MSVC_WARNINGS})
    elseif(CMAKE_CXX_COMPILER_ID MATCHES ".*Clang")
        set(PROJECT_WARNINGS ${CLANG_WARNINGS})
    elseif(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
        set(PROJECT_WARNINGS ${GCC_WARNINGS})
    else()
        message(AUTHOR_WARNING "No compiler warnings set for compiler: '${CMAKE_CXX_COMPILER_ID}'")
    endif()

    # Apply warnings to target
    target_compile_options(${project_name} PRIVATE ${PROJECT_WARNINGS})

    message(VERBOSE "Applied ${CMAKE_CXX_COMPILER_ID} warnings to target: ${project_name}")
endfunction()
