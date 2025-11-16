// ==============================================================================
// linalg.cpp - Linear Algebra Excel Functions using Eigen
// ==============================================================================
// Implementation of matrix operations, decompositions, and solvers for Excel
// Uses Eigen 5.0+ for high-performance linear algebra computations
// ==============================================================================

#include "linalg.h"

// Suppress warnings from Eigen library headers (external code)
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wnull-dereference"
#elif defined(__clang__)
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wnull-dereference"
#elif defined(_MSC_VER)
#pragma warning(push)
#pragma warning(disable: 4996) // deprecated warnings
#endif

#include <Eigen/Dense>
#include <Eigen/SVD>
#include <Eigen/QR>
#include <Eigen/Cholesky>
#include <Eigen/Eigenvalues>

#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic pop
#elif defined(__clang__)
#pragma clang diagnostic pop
#elif defined(_MSC_VER)
#pragma warning(pop)
#endif

using namespace xll;
using namespace Eigen;

// ==============================================================================
// Helper Function Implementations
// ==============================================================================

MatrixXd xll::fp_to_eigen(const _FP12* fp)
{
    // Excel stores in row-major order, Eigen defaults to column-major
    // Use Map with RowMajor to avoid copy and transpose
    return Map<const Matrix<double, Dynamic, Dynamic, RowMajor>>(
        fp->array, fp->rows, fp->columns);
}

_FP12* xll::eigen_to_fp(const MatrixXd& mat)
{
    static FPX result;
    result.resize(static_cast<int>(mat.rows()), static_cast<int>(mat.cols()));
    Map<Matrix<double, Dynamic, Dynamic, RowMajor>>(
        result.array(), mat.rows(), mat.cols()) = mat;
    return result.get();
}

_FP12* xll::vector_to_fp(const VectorXd& vec)
{
    static FPX result;
    result.resize(static_cast<int>(vec.size()), 1);
    for (int i = 0; i < vec.size(); ++i) {
        result(i, 0) = vec(i);
    }
    return result.get();
}

_FP12* xll::row_vector_to_fp(const VectorXd& vec)
{
    static FPX result;
    result.resize(1, static_cast<int>(vec.size()));
    for (int i = 0; i < vec.size(); ++i) {
        result(0, i) = vec(i);
    }
    return result.get();
}

// ==============================================================================
// Basic Operations (8 functions)
// ==============================================================================

// -----------------------------------------------------------------------------
// MATRIX.ADD - Element-wise matrix addition
// -----------------------------------------------------------------------------
AddIn xai_matrix_add(
    Function(XLL_FP, "xll_matrix_add", "MATRIX.ADD")
    .Arguments({
        Arg(XLL_FP, "A", "is the first matrix."),
        Arg(XLL_FP, "B", "is the second matrix.")
    })
    .FunctionHelp("Add two matrices element-wise.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes element-wise matrix addition: C = A + B</p>
<p><b>Input:</b> A and B must have the same dimensions (m×n)</p>
<p><b>Output:</b> Matrix C (m×n) with C[i,j] = A[i,j] + B[i,j]</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_add(_FP12* pa, _FP12* pb)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd B = fp_to_eigen(pb);

        if (A.rows() != B.rows() || A.cols() != B.cols()) {
            return nullptr; // Dimension mismatch
        }

        MatrixXd C = A + B;
        return eigen_to_fp(C);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.SUB - Element-wise matrix subtraction
// -----------------------------------------------------------------------------
AddIn xai_matrix_sub(
    Function(XLL_FP, "xll_matrix_sub", "MATRIX.SUB")
    .Arguments({
        Arg(XLL_FP, "A", "is the first matrix."),
        Arg(XLL_FP, "B", "is the second matrix.")
    })
    .FunctionHelp("Subtract matrix B from A element-wise.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes element-wise matrix subtraction: C = A - B</p>
<p><b>Input:</b> A and B must have the same dimensions (m×n)</p>
<p><b>Output:</b> Matrix C (m×n) with C[i,j] = A[i,j] - B[i,j]</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_sub(_FP12* pa, _FP12* pb)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd B = fp_to_eigen(pb);

        if (A.rows() != B.rows() || A.cols() != B.cols()) {
            return nullptr;
        }

        MatrixXd C = A - B;
        return eigen_to_fp(C);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.MUL - Matrix multiplication
// -----------------------------------------------------------------------------
AddIn xai_matrix_mul(
    Function(XLL_FP, "xll_matrix_mul", "MATRIX.MUL")
    .Arguments({
        Arg(XLL_FP, "A", "is the first matrix."),
        Arg(XLL_FP, "B", "is the second matrix.")
    })
    .FunctionHelp("Multiply matrices A and B.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes matrix multiplication: C = A × B</p>
<p>Mathematical formula: \[C = AB\]</p>
<p><b>Input:</b> A(m×n) and B(n×p) - columns of A must equal rows of B</p>
<p><b>Output:</b> Matrix C(m×p)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_mul(_FP12* pa, _FP12* pb)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd B = fp_to_eigen(pb);

        if (A.cols() != B.rows()) {
            return nullptr; // Dimension mismatch
        }

        MatrixXd C = A * B;
        return eigen_to_fp(C);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.TRANSPOSE - Matrix transpose
// -----------------------------------------------------------------------------
AddIn xai_matrix_transpose(
    Function(XLL_FP, "xll_matrix_transpose", "MATRIX.TRANSPOSE")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix to transpose.")
    })
    .FunctionHelp("Transpose a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes matrix transpose: \[A^T\]</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Matrix A^T(n×m) where (A^T)[i,j] = A[j,i]</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_transpose(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd At = A.transpose();
        return eigen_to_fp(At);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.TRACE - Sum of diagonal elements
// -----------------------------------------------------------------------------
AddIn xai_matrix_trace(
    Function(XLL_DOUBLE, "xll_matrix_trace", "MATRIX.TRACE")
    .Arguments({
        Arg(XLL_FP, "A", "is a square matrix.")
    })
    .FunctionHelp("Compute the trace of a square matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes the trace (sum of diagonal elements): \[\text{tr}(A) = \sum_{i=1}^{n} a_{ii}\]</p>
<p><b>Input:</b> Square matrix A(n×n)</p>
<p><b>Output:</b> Scalar value equal to sum of diagonal elements</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_matrix_trace(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return std::numeric_limits<double>::quiet_NaN(); // Not square
        }

        return A.trace();
    }
    catch (...) {
        return std::numeric_limits<double>::quiet_NaN();
    }
}

// -----------------------------------------------------------------------------
// MATRIX.NORM - Frobenius norm
// -----------------------------------------------------------------------------
AddIn xai_matrix_norm(
    Function(XLL_DOUBLE, "xll_matrix_norm", "MATRIX.NORM")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute the Frobenius norm of a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes the Frobenius norm: \[\|A\|_F = \sqrt{\sum_{i,j} a_{ij}^2}\]</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Scalar value representing the Frobenius norm</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_matrix_norm(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        return A.norm();
    }
    catch (...) {
        return std::numeric_limits<double>::quiet_NaN();
    }
}

// -----------------------------------------------------------------------------
// MATRIX.DETERMINANT - Matrix determinant
// -----------------------------------------------------------------------------
AddIn xai_matrix_det(
    Function(XLL_DOUBLE, "xll_matrix_det", "MATRIX.DETERMINANT")
    .Arguments({
        Arg(XLL_FP, "A", "is a square matrix.")
    })
    .FunctionHelp("Compute the determinant of a square matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes the determinant: \[\det(A)\]</p>
<p><b>Input:</b> Square matrix A(n×n)</p>
<p><b>Output:</b> Scalar determinant value</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_matrix_det(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return std::numeric_limits<double>::quiet_NaN(); // Not square
        }

        return A.determinant();
    }
    catch (...) {
        return std::numeric_limits<double>::quiet_NaN();
    }
}

// -----------------------------------------------------------------------------
// MATRIX.RANK - Matrix rank
// -----------------------------------------------------------------------------
AddIn xai_matrix_rank(
    Function(XLL_DOUBLE, "xll_matrix_rank", "MATRIX.RANK")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute the rank of a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes the rank (number of linearly independent rows/columns).</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Integer rank value (0 to min(m,n))</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
double WINAPI xll_matrix_rank(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        FullPivLU<MatrixXd> lu(A);
        return static_cast<double>(lu.rank());
    }
    catch (...) {
        return std::numeric_limits<double>::quiet_NaN();
    }
}

// ==============================================================================
// Matrix Decompositions (6 functions)
// ==============================================================================

// -----------------------------------------------------------------------------
// MATRIX.INVERSE - Matrix inverse
// -----------------------------------------------------------------------------
AddIn xai_matrix_inv(
    Function(XLL_FP, "xll_matrix_inv", "MATRIX.INVERSE")
    .Arguments({
        Arg(XLL_FP, "A", "is a square invertible matrix.")
    })
    .FunctionHelp("Compute the inverse of a square matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes matrix inverse: \[A^{-1}\] such that \[AA^{-1} = I\]</p>
<p><b>Input:</b> Invertible square matrix A(n×n)</p>
<p><b>Output:</b> Inverse matrix A^-1(n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_inv(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return nullptr; // Not square
        }

        MatrixXd Ainv = A.inverse();
        return eigen_to_fp(Ainv);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.LU - LU decomposition (returns combined L+U)
// -----------------------------------------------------------------------------
AddIn xai_matrix_lu(
    Function(XLL_FP, "xll_matrix_lu", "MATRIX.LU")
    .Arguments({
        Arg(XLL_FP, "A", "is a square matrix.")
    })
    .FunctionHelp("Compute LU decomposition of a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes LU decomposition: \[A = PLU\]</p>
<p>Returns combined matrix with L (lower) and U (upper) triangular parts.</p>
<p><b>Input:</b> Square matrix A(n×n)</p>
<p><b>Output:</b> Combined LU matrix (n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_lu(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return nullptr;
        }

        PartialPivLU<MatrixXd> lu(A);
        MatrixXd L = MatrixXd::Identity(A.rows(), A.cols());
        L.triangularView<StrictlyLower>() = lu.matrixLU();
        MatrixXd U = lu.matrixLU().triangularView<Upper>();

        // Return combined (L below diagonal + U on and above diagonal)
        MatrixXd LU = lu.matrixLU();
        return eigen_to_fp(LU);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.QR - QR decomposition (returns R matrix)
// -----------------------------------------------------------------------------
AddIn xai_matrix_qr(
    Function(XLL_FP, "xll_matrix_qr", "MATRIX.QR")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute QR decomposition of a matrix (returns R).")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes QR decomposition: \[A = QR\]</p>
<p>Returns upper triangular matrix R.</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Upper triangular matrix R(n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_qr(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        HouseholderQR<MatrixXd> qr(A);
        MatrixXd R = qr.matrixQR().triangularView<Upper>();
        return eigen_to_fp(R);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.CHOLESKY - Cholesky decomposition
// -----------------------------------------------------------------------------
AddIn xai_matrix_cholesky(
    Function(XLL_FP, "xll_matrix_cholesky", "MATRIX.CHOLESKY")
    .Arguments({
        Arg(XLL_FP, "A", "is a symmetric positive-definite matrix.")
    })
    .FunctionHelp("Compute Cholesky decomposition of a SPD matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes Cholesky decomposition: \[A = LL^T\]</p>
<p>Returns lower triangular matrix L.</p>
<p><b>Input:</b> Symmetric positive-definite matrix A(n×n)</p>
<p><b>Output:</b> Lower triangular matrix L(n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_cholesky(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return nullptr;
        }

        LLT<MatrixXd> llt(A);
        if (llt.info() != Success) {
            return nullptr; // Not positive definite
        }

        MatrixXd L = llt.matrixL();
        return eigen_to_fp(L);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.SVD - Singular values
// -----------------------------------------------------------------------------
AddIn xai_matrix_svd(
    Function(XLL_FP, "xll_matrix_svd", "MATRIX.SVD")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute singular values of a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes singular value decomposition: \[A = U\Sigma V^T\]</p>
<p>Returns vector of singular values (diagonal of Σ).</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Column vector of singular values (min(m,n)×1)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_svd(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        JacobiSVD<MatrixXd> svd(A);
        VectorXd sigma = svd.singularValues();
        return vector_to_fp(sigma);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.SVD_FULL - Full SVD with U, Σ, V^T stacked vertically
// -----------------------------------------------------------------------------
AddIn xai_matrix_svd_full(
    Function(XLL_FP, "xll_matrix_svd_full", "MATRIX.SVD_FULL")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute full SVD decomposition with all components stacked.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes singular value decomposition: \[A = U\Sigma V^T\]</p>
<p>Returns U, Σ (as diagonal matrix), and V^T stacked vertically in one matrix.</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Stacked matrix with U (m×k), Σ (k×k), V^T (k×n) where k=min(m,n)</p>
<p>Total dimensions: (m+k+k)×max(m,n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_svd_full(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        // Compute thin SVD
        JacobiSVD<MatrixXd, ComputeThinU | ComputeThinV> svd(A);

        MatrixXd U = svd.matrixU();          // m × k
        VectorXd sigma = svd.singularValues(); // k × 1
        MatrixXd V = svd.matrixV();          // n × k
        MatrixXd Vt = V.transpose();         // k × n

        int m = static_cast<int>(A.rows());
        int n = static_cast<int>(A.cols());
        int k = static_cast<int>(sigma.size()); // min(m,n)

        // Create Σ as diagonal matrix (k×k)
        MatrixXd Sigma = sigma.asDiagonal();

        // Determine max width needed
        int max_cols = (std::max)(m, n);
        max_cols = (std::max)(max_cols, k);

        // Total rows: m (U) + k (Σ) + k (V^T)
        int total_rows = m + k + k;

        // Create result matrix with zero padding
        static FPX result;
        result.resize(total_rows, max_cols);

        // Initialize to zero
        for (int i = 0; i < total_rows * max_cols; ++i) {
            result[i] = 0.0;
        }

        // Copy U (m×k) to top section
        for (int i = 0; i < m; ++i) {
            for (int j = 0; j < k; ++j) {
                result(i, j) = U(i, j);
            }
        }

        // Copy Σ (k×k) to middle section
        for (int i = 0; i < k; ++i) {
            for (int j = 0; j < k; ++j) {
                result(m + i, j) = Sigma(i, j);
            }
        }

        // Copy V^T (k×n) to bottom section
        for (int i = 0; i < k; ++i) {
            for (int j = 0; j < n; ++j) {
                result(m + k + i, j) = Vt(i, j);
            }
        }

        return result.get();
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.EIGENVALUES - Eigenvalues
// -----------------------------------------------------------------------------
AddIn xai_matrix_eigenvalues(
    Function(XLL_FP, "xll_matrix_eigenvalues", "MATRIX.EIGENVALUES")
    .Arguments({
        Arg(XLL_FP, "A", "is a square matrix.")
    })
    .FunctionHelp("Compute eigenvalues of a square matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes eigenvalues satisfying: \[Av = \lambda v\]</p>
<p>Returns real parts of eigenvalues (imaginary parts ignored).</p>
<p><b>Input:</b> Square matrix A(n×n)</p>
<p><b>Output:</b> Column vector of eigenvalues (n×1)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_eigenvalues(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return nullptr;
        }

        EigenSolver<MatrixXd> es(A);
        VectorXd eigenvalues = es.eigenvalues().real();
        return vector_to_fp(eigenvalues);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.EIGENVECTORS - Eigenvectors
// -----------------------------------------------------------------------------
AddIn xai_matrix_eigenvectors(
    Function(XLL_FP, "xll_matrix_eigenvectors", "MATRIX.EIGENVECTORS")
    .Arguments({
        Arg(XLL_FP, "A", "is a square matrix.")
    })
    .FunctionHelp("Compute eigenvectors of a square matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes eigenvectors satisfying: \[Av = \lambda v\]</p>
<p>Returns matrix where each column is an eigenvector (real parts only).</p>
<p><b>Input:</b> Square matrix A(n×n)</p>
<p><b>Output:</b> Eigenvector matrix V(n×n) where column i corresponds to eigenvalue i</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_eigenvectors(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        if (A.rows() != A.cols()) {
            return nullptr;
        }

        EigenSolver<MatrixXd> es(A);
        MatrixXd eigenvectors = es.eigenvectors().real();
        return eigen_to_fp(eigenvectors);
    }
    catch (...) {
        return nullptr;
    }
}

// ==============================================================================
// Linear Solvers (3 functions)
// ==============================================================================

// -----------------------------------------------------------------------------
// MATRIX.SOLVE - Solve linear system Ax = b
// -----------------------------------------------------------------------------
AddIn xai_matrix_solve(
    Function(XLL_FP, "xll_matrix_solve", "MATRIX.SOLVE")
    .Arguments({
        Arg(XLL_FP, "A", "is a square coefficient matrix."),
        Arg(XLL_FP, "b", "is the right-hand side vector or matrix.")
    })
    .FunctionHelp("Solve linear system Ax = b.")
    .Category("LINALG")
    .Documentation(R"(
<p>Solves linear system: \[Ax = b\]</p>
<p><b>Input:</b> A(n×n) coefficient matrix, b(n×1) or b(n×m) right-hand side</p>
<p><b>Output:</b> Solution x(n×1) or x(n×m)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_solve(_FP12* pa, _FP12* pb)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd b = fp_to_eigen(pb);

        if (A.rows() != A.cols() || A.rows() != b.rows()) {
            return nullptr;
        }

        MatrixXd x = A.lu().solve(b);
        return eigen_to_fp(x);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.LSTSQ - Least squares solution
// -----------------------------------------------------------------------------
AddIn xai_matrix_lstsq(
    Function(XLL_FP, "xll_matrix_lstsq", "MATRIX.LSTSQ")
    .Arguments({
        Arg(XLL_FP, "A", "is the coefficient matrix."),
        Arg(XLL_FP, "b", "is the right-hand side vector.")
    })
    .FunctionHelp("Solve overdetermined system in least squares sense.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes least squares solution: \[\min_x \|Ax - b\|_2\]</p>
<p><b>Input:</b> A(m×n) coefficient matrix, b(m×1) right-hand side</p>
<p><b>Output:</b> Solution x(n×1) minimizing residual</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_lstsq(_FP12* pa, _FP12* pb)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);
        MatrixXd b = fp_to_eigen(pb);

        if (A.rows() != b.rows()) {
            return nullptr;
        }

        // Use SVD for robust least squares
        BDCSVD<MatrixXd, ComputeThinU | ComputeThinV> svd(A);
        MatrixXd x = svd.solve(b);
        return eigen_to_fp(x);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.PSEUDO_INV - Moore-Penrose pseudoinverse
// -----------------------------------------------------------------------------
AddIn xai_matrix_pinv(
    Function(XLL_FP, "xll_matrix_pinv", "MATRIX.PSEUDO_INV")
    .Arguments({
        Arg(XLL_FP, "A", "is the matrix.")
    })
    .FunctionHelp("Compute Moore-Penrose pseudoinverse of a matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Computes pseudoinverse: \[A^+\]</p>
<p>For full rank matrices: A^+ = (A^T A)^-1 A^T (left) or A^T(AA^T)^-1 (right)</p>
<p><b>Input:</b> Matrix A(m×n)</p>
<p><b>Output:</b> Pseudoinverse A^+(n×m)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_pinv(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        // Compute pseudoinverse using SVD
        JacobiSVD<MatrixXd, ComputeThinU | ComputeThinV> svd(A);

        // Tolerance for singular values
        auto max_dim = (std::max)(A.rows(), A.cols());
        double tolerance = std::numeric_limits<double>::epsilon() *
                          max_dim *
                          svd.singularValues().array().abs().maxCoeff();

        // Compute pseudoinverse
        VectorXd singularValuesInv(svd.singularValues().size());
        for (int i = 0; i < svd.singularValues().size(); ++i) {
            if (svd.singularValues()(i) > tolerance) {
                singularValuesInv(i) = 1.0 / svd.singularValues()(i);
            } else {
                singularValuesInv(i) = 0.0;
            }
        }

        MatrixXd Ainv = svd.matrixV() *
                       singularValuesInv.asDiagonal() *
                       svd.matrixU().adjoint();

        return eigen_to_fp(Ainv);
    }
    catch (...) {
        return nullptr;
    }
}

// ==============================================================================
// Utility Functions (3 functions)
// ==============================================================================

// -----------------------------------------------------------------------------
// MATRIX.IDENTITY - Create identity matrix
// -----------------------------------------------------------------------------
AddIn xai_matrix_identity(
    Function(XLL_FP, "xll_matrix_identity", "MATRIX.IDENTITY")
    .Arguments({
        Arg(XLL_DOUBLE, "n", "is the dimension of the identity matrix.")
    })
    .FunctionHelp("Create an n×n identity matrix.")
    .Category("LINALG")
    .Documentation(R"(
<p>Creates identity matrix: \[I_n\]</p>
<p>I[i,j] = 1 if i=j, 0 otherwise</p>
<p><b>Input:</b> Dimension n</p>
<p><b>Output:</b> Identity matrix I(n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_identity(double n)
{
#pragma XLLEXPORT
    try {
        int dim = static_cast<int>(n);
        if (dim <= 0 || dim > 10000) {
            return nullptr; // Invalid dimension
        }

        MatrixXd I = MatrixXd::Identity(dim, dim);
        return eigen_to_fp(I);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.ZEROS - Create zero matrix
// -----------------------------------------------------------------------------
AddIn xai_matrix_zeros(
    Function(XLL_FP, "xll_matrix_zeros", "MATRIX.ZEROS")
    .Arguments({
        Arg(XLL_DOUBLE, "m", "is the number of rows."),
        Arg(XLL_DOUBLE, "n", "is the number of columns.")
    })
    .FunctionHelp("Create an m×n matrix of zeros.")
    .Category("LINALG")
    .Documentation(R"(
<p>Creates zero matrix with all elements equal to 0.</p>
<p><b>Input:</b> Dimensions m (rows) and n (columns)</p>
<p><b>Output:</b> Zero matrix (m×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_zeros(double m, double n)
{
#pragma XLLEXPORT
    try {
        int rows = static_cast<int>(m);
        int cols = static_cast<int>(n);

        if (rows <= 0 || cols <= 0 || rows > 10000 || cols > 10000) {
            return nullptr;
        }

        MatrixXd Z = MatrixXd::Zero(rows, cols);
        return eigen_to_fp(Z);
    }
    catch (...) {
        return nullptr;
    }
}

// -----------------------------------------------------------------------------
// MATRIX.DIAG - Extract or create diagonal matrix
// -----------------------------------------------------------------------------
AddIn xai_matrix_diag(
    Function(XLL_FP, "xll_matrix_diag", "MATRIX.DIAG")
    .Arguments({
        Arg(XLL_FP, "A", "is a matrix or vector.")
    })
    .FunctionHelp("Extract diagonal from matrix or create diagonal matrix from vector.")
    .Category("LINALG")
    .Documentation(R"(
<p>If A is a matrix: extracts diagonal elements as column vector</p>
<p>If A is a vector: creates diagonal matrix with vector elements on diagonal</p>
<p><b>Input:</b> Matrix A(n×n) or vector v(n×1)</p>
<p><b>Output:</b> Vector (n×1) or diagonal matrix (n×n)</p>
)")
);

#if defined(__GNUC__) || defined(__clang__)
extern "C"
#endif
_FP12* WINAPI xll_matrix_diag(_FP12* pa)
{
#pragma XLLEXPORT
    try {
        MatrixXd A = fp_to_eigen(pa);

        // If vector (1 column), create diagonal matrix
        if (A.cols() == 1) {
            VectorXd v = A.col(0);
            MatrixXd D = v.asDiagonal();
            return eigen_to_fp(D);
        }
        // If matrix, extract diagonal
        else {
            VectorXd d = A.diagonal();
            return vector_to_fp(d);
        }
    }
    catch (...) {
        return nullptr;
    }
}
