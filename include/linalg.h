#pragma once
// ==============================================================================
// linalg.h - Linear Algebra Excel Functions using Eigen
// ==============================================================================
// Provides matrix operations, decompositions, and solvers for Excel
// Uses Eigen 5.0+ for high-performance linear algebra
// ==============================================================================

#include "xll24/include/xll.h"

// Suppress warnings from Eigen library headers (external code)
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wnull-dereference"
#elif defined(__clang__)
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wnull-dereference"
#elif defined(_MSC_VER)
#pragma warning(push)
#pragma warning(disable: 4996)
#endif

#include <Eigen/Dense>

#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic pop
#elif defined(__clang__)
#pragma clang diagnostic pop
#elif defined(_MSC_VER)
#pragma warning(pop)
#endif

namespace xll {

// ==============================================================================
// Helper Functions - Excel ↔ Eigen Conversions
// ==============================================================================

// Convert Excel FP12 array to Eigen MatrixXd
Eigen::MatrixXd fp_to_eigen(const _FP12* fp);

// Convert Eigen MatrixXd to Excel FP12 array (static storage)
_FP12* eigen_to_fp(const Eigen::MatrixXd& mat);

// Convert Eigen VectorXd to Excel FP12 array as column vector
_FP12* vector_to_fp(const Eigen::VectorXd& vec);

// Convert Eigen VectorXd to Excel FP12 array as row vector
_FP12* row_vector_to_fp(const Eigen::VectorXd& vec);

} // namespace xll
