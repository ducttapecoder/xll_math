#include "xll_template.h"

#include <cmath> // for double tgamma(double)
#include <format>

using namespace xll;

AddIn xai_tgamma(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Array of function arguments.
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
		})
	// Function Wizard help.
	.FunctionHelp("Return the Gamma function value.")
	// Function Wizard category.
	.Category("MATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx")
);
// WINAPI calling convention must be specified
// For GCC/Clang, use extern "C" to prevent name mangling
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT // must be specified to export function

	return tgamma(x);
}

// Compile-time compiler information
namespace {
	// Compiler name detection
	constexpr std::string_view COMPILER_NAME =
#if defined(_MSC_VER)
		"MSVC"
#elif defined(__clang__)
		"Clang"
#elif defined(__GNUC__)
		"GCC"
#else
		"Unknown"
#endif
		;

	// Compiler version as string
	constexpr std::string_view COMPILER_VERSION =
#if defined(_MSC_VER)
		// MSVC version: _MSC_VER is defined as major*100 + minor
#define STRINGIFY(x) #x
#define TOSTRING(x) STRINGIFY(x)
		"v" TOSTRING(_MSC_VER)
#elif defined(__clang__)
		__clang_version__
#elif defined(__GNUC__)
		__VERSION__
#else
		"Unknown"
#endif
		;

	// C++ standard version
	constexpr std::string_view CPP_STANDARD =
#if __cplusplus >= 202600L
		"C++26"
#elif __cplusplus >= 202302L
		"C++23"
#elif __cplusplus >= 202002L
		"C++20"
#elif __cplusplus >= 201703L
		"C++17"
#elif __cplusplus >= 201402L
		"C++14"
#elif __cplusplus >= 201103L
		"C++11"				
#else
		"C++ grandma"
#endif
		;
}

AddIn xai_about(
	Function(XLL_CSTRING4, "xll_about", "XLL.ABOUT")
	.FunctionHelp("Returns compiler and build information for this XLL add-in.")
	.Category("XLL")
	.Documentation(R"(
Returns information about the compiler used to build this Excel add-in, including:
<ul>
<li>Compiler name and version</li>
<li>C++ standard version</li>
<li>Build date and time</li>
</ul>
All values are locked in at compile time using <code>constexpr</code>.
)")
);
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
LPCSTR WINAPI xll_about()
{
#pragma XLLEXPORT
	static std::string result = std::format("Compiler: {} {}\n C++ Standard: {}\n Built: {} {}", COMPILER_NAME,
		COMPILER_VERSION, CPP_STANDARD, __DATE__, __TIME__);

	return LPCSTR(result.data());
}

// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
AddIn xai_macro(
	// C++ function, Excel name of macro
	Macro("xll_macro", "XLL.MACRO")
);
// Macros must have `int WINAPI (*)(void)` signature.
// For GCC/Clang, use extern "C" to prevent name mangling
#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	// https://xlladdins.github.io/Excel4Macros/reftext.html
	// A1 style instead of default R1C1.
	OPER reftext = Excel(xlfReftext, Excel(xlfActiveCell), OPER(true));
	// UTF-8 strings can be used.
	Excel(xlcAlert, OPER("XLL.MACRO called with активный 细胞: ") & reftext);

	return TRUE;
}