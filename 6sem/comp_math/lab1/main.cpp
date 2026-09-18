#include <iostream>
#include <vector>
#include <cmath>        // Для std::abs, std::max
#include <numeric>      // Для std::accumulate (если понадобится)
#include <iomanip>      // Для std::setprecision, std::fixed
#include <string>
#include <stdexcept>    // Для std::runtime_error
#include <limits>       // Для std::numeric_limits
#include <algorithm>    // Для std::swap, std::max_element


typedef std::vector<std::vector<double>> Matrix;
const double EPSILON = 1e-9;

void print_vector(const std::vector<double>& v, const std::string& name = "", int precision = 15) {
    if (!name.empty()) {
        std::cout<< name << ": ";
    }
    std::cout << std::fixed << std::setprecision(precision);
    std::cout << "[";
    for (size_t i = 0; i < v.size(); ++i) {
        std::cout<< std::fixed << std::setprecision(precision) << v[i] << (i == v.size() - 1 ? "" : ", ");
    }
    std::cout << "]" << std::endl;
}

void print_matrix(const Matrix& matrix, const std::string& name = "", int precision = 6) {
    if (!name.empty()) {
        std::cout << name << ":" << std::endl;
    }
    std::cout << std::fixed << std::setprecision(precision);
    for (size_t i = 0; i < matrix.size(); ++i) {
        std::cout << "  "; // Отступ для строки матрицы
        print_vector(matrix[i], "", precision);
    }
}

// Вычисление 1-нормы вектора: sum(|v_i|)
double compute_vector_norm_1(const std::vector<double>& v) {
    double norm = 0.0;
    for (double val : v) {
        norm += std::abs(val);
    }
    return norm;
}

// Вычисление infinity-нормы вектора: max(|v_i|)
double compute_vector_norm_inf(const std::vector<double>& v) {
    double norm = 0.0;
    for (double val : v) {
        norm = std::max(norm, std::abs(val));
    }
    return norm;
}

// Вычисление разности векторов: A - B
std::vector<double> diff_vector(const std::vector<double>& A, const std::vector<double>& B) {
    if (A.size() != B.size()) {
        throw std::invalid_argument("Векторы должны иметь одинаковый размер для вычитания.");
    }
    const size_t n = A.size();
    std::vector<double> result(n);
    for (size_t i = 0; i < n; ++i) {
        result[i] = A[i] - B[i];
    }
    return result;
}

// Вычисление 1-нормы матрицы: max по столбцам (sum(|a_ij|))
double compute_matrix_norm_1(const Matrix& matrix) {
    if (matrix.empty() || matrix[0].empty()) return 0.0;
    const size_t N = matrix.size();
    const size_t M = matrix[0].size();
    double max_col_sum = 0.0;

    for (size_t j = 0; j < M; ++j) {
        double current_col_sum = 0.0;
        for (size_t i = 0; i < N; ++i) {
            current_col_sum += std::abs(matrix[i][j]);
        }
        max_col_sum = std::max(max_col_sum, current_col_sum);
    }
    return max_col_sum;
}

// Вычисление infinity-нормы матрицы: max по строкам (sum(|a_ij|))
double compute_matrix_norm_inf(const Matrix& matrix) {
    if (matrix.empty()) return 0.0;
    const size_t N = matrix.size();
    const size_t M = matrix[0].size(); // Предполагаем, что матрица не пустая
    double max_row_sum = 0.0;

    for (size_t i = 0; i < N; ++i) {
        double current_row_sum = 0.0;
        for (size_t j = 0; j < M; ++j) {
             if (j < matrix[i].size()) {
                 current_row_sum += std::abs(matrix[i][j]);
             }
        }
        max_row_sum = std::max(max_row_sum, current_row_sum);
    }
    return max_row_sum;
}


// Умножение матрицы на вектор: y = A * x
std::vector<double> multiply_matrix_vector(const Matrix& A, const std::vector<double>& x) {
    const size_t rowsA = A.size();
    if (rowsA == 0) return {};
    const size_t colsA = A[0].size();
    const size_t sizeX = x.size();

    if (colsA != sizeX) {
        throw std::invalid_argument("Количество столбцов матрицы должно быть равно размеру вектора для умножения.");
    }

    std::vector<double> y(rowsA, 0.0);
    for (size_t i = 0; i < rowsA; ++i) {
        for (size_t j = 0; j < colsA; ++j) {
            y[i] += A[i][j] * x[j];
        }
    }
    return y;
}

// Умножение матриц: C = A * B
Matrix multiply(const Matrix& A, const Matrix& B) {
    const size_t rowsA = A.size();
    if (rowsA == 0) return {};
    const size_t colsA = A[0].size();
    const size_t rowsB = B.size();
    if (rowsB == 0) return {};
    const size_t colsB = B[0].size();

    if (colsA != rowsB) {
         throw std::invalid_argument("Несовместимые размеры матриц для умножения.");
    }

    Matrix result(rowsA, std::vector<double>(colsB, 0.0));
    for (size_t i = 0; i < rowsA; ++i) {
        for (size_t j = 0; j < colsB; ++j) {
            for (size_t k = 0; k < colsA; ++k) {
                result[i][j] += A[i][k] * B[k][j];
            }
        }
    }
    return result;
}

// Создание единичной матрицы
Matrix create_identity_matrix(size_t n) {
    Matrix I(n, std::vector<double>(n, 0.0));
    for (size_t i = 0; i < n; ++i) {
        I[i][i] = 1.0;
    }

    return I;
}

// Сравнение двух матриц с допуском epsilon
bool are_matrices_close(const Matrix& A, const Matrix& B, double tolerance = EPSILON) {
    if (A.size() != B.size() || (A.empty() && !B.empty()) || (!A.empty() && B.empty())) {
        return false;
    }
    if (A.empty()) return true; // Обе пустые
    if (A[0].size() != B[0].size()) {
        return false;
    }

    const size_t rows = A.size();
    const size_t cols = A[0].size();

    for (size_t i = 0; i < rows; ++i) {
        for (size_t j = 0; j < cols; ++j) {
            std::cout << std::abs(A[i][j] - B[i][j]) << ' ';

            if (std::abs(A[i][j] - B[i][j]) > tolerance) {
                return false;
            }
        }

        std::cout << std::endl;
    }
    return true;
}


// Решение СЛАУ методом Гаусса с выбором главного элемента по столбцу
std::vector<double> gauss(Matrix A, std::vector<double> b) { // Принимаем копии, т.к. будем их изменять
    const size_t n = A.size();
    if (n == 0 || A[0].size() != n || b.size() != n) {
        throw std::invalid_argument("Некорректные размеры матрицы или вектора для метода Гаусса.");
    }

    // Прямой ход
    for (size_t i = 0; i < n; ++i) {
        // Выбор главного элемента в текущем столбце (начиная с i-й строки)
        size_t maxRow = i;
        for (size_t k = i + 1; k < n; ++k) {
            if (std::abs(A[k][i]) > std::abs(A[maxRow][i])) {
                maxRow = k;
            }
        }

        // Проверка на вырожденность (с учетом погрешности)
        if (std::abs(A[maxRow][i]) < EPSILON) {
             bool found_pivot = false;
             for(size_t check_col = i + 1; check_col < n; ++check_col) {
                 if (std::abs(A[maxRow][check_col]) >= EPSILON) {
                     found_pivot = true; // Нашли потенциальный опорный, но в другом столбце
                     break;
                 }
             }
             if (!found_pivot) {
                 throw std::runtime_error("Матрица вырождена или близка к вырожденной (метод Гаусса).");
             }
             throw std::runtime_error("Матрица вырождена или требует перестановки столбцов (метод Гаусса).");
        }


        // Перестановка строк (матрицы A и вектора b)
        std::swap(A[i], A[maxRow]);
        std::swap(b[i], b[maxRow]);

        // Обнуление элементов под главным элементом
        double pivot = A[i][i]; // Обновляем pivot после возможной перестановки
        for (size_t k = i + 1; k < n; ++k) {
            const double factor = A[k][i] / pivot;
            if (std::abs(factor) < EPSILON) continue; // Пропускаем, если множитель почти нулевой
            for (size_t j = i; j < n; ++j) { // Начинаем с j=i, т.к. A[k][i] обнулится
                A[k][j] -= factor * A[i][j];
            }
            b[k] -= factor * b[i];
            // Явно обнуляем для точности (хотя математически уже должно быть 0)
            A[k][i] = 0.0;
        }
    }

    // Обратный ход
    std::vector<double> x(n);
    for (int i = static_cast<int>(n) - 1; i >= 0; --i) {
        if (std::abs(A[i][i]) < EPSILON) {
             throw std::runtime_error("Нулевой элемент на диагонали после прямого хода (метод Гаусса).");
        }
        double sum = 0.0;
        for (size_t j = i + 1; j < n; ++j) {
            sum += A[i][j] * x[j];
        }
        x[i] = (b[i] - sum) / A[i][i];
    }

    return x;
}


// LU-разложение (Doolittle's method: L имеет 1 на диагонали)
// Возвращает пару {L, U}
std::pair<Matrix, Matrix> LU_dec(const Matrix& matrix) {
    const size_t n = matrix.size();
    if (n == 0 || matrix[0].size() != n) {
        throw std::invalid_argument("Матрица должна быть квадратной для LU-разложения.");
    }

    Matrix L(n, std::vector<double>(n, 0.0));
    Matrix U(n, std::vector<double>(n, 0.0));

    for (size_t i = 0; i < n; ++i) {
        // Расчет U
        for (size_t j = i; j < n; ++j) {
            double sum = 0.0;
            for (size_t k = 0; k < i; ++k) {
                sum += L[i][k] * U[k][j];
            }
            U[i][j] = matrix[i][j] - sum;
        }

        // Проверка на ноль на диагонали U
        if (std::abs(U[i][i]) < EPSILON) {
            // LU-разложение без перестановок не существует или матрица вырождена
            throw std::runtime_error("Нулевой диагональный элемент в U при LU-разложении без перестановок.");
        }

        // Расчет L
        L[i][i] = 1.0; // Диагональ L равна 1 для Doolittle
        for (size_t j = i + 1; j < n; ++j) {
            double sum = 0.0;
            for (size_t k = 0; k < i; ++k) {
                sum += L[j][k] * U[k][i];
            }
            L[j][i] = (matrix[j][i] - sum) / U[i][i];
        }
    }

    return {L, U};
}

// Прямая подстановка для решения Ly = b (L - нижнетреугольная)
std::vector<double> solve_forward_L(const Matrix& L, const std::vector<double>& b) {
    const size_t n = L.size();
    if (n == 0 || L[0].size() != n || b.size() != n) {
         throw std::invalid_argument("Некорректные размеры для прямой подстановки.");
    }
    std::vector<double> y(n);

    for (size_t i = 0; i < n; ++i) {
        double sum = 0.0;
        for (size_t j = 0; j < i; ++j) {
            sum += L[i][j] * y[j];
        }
        // Проверка на ноль на диагонали L (хотя в Doolittle L[i][i]=1)
        if (std::abs(L[i][i]) < EPSILON) {
            throw std::runtime_error("Нулевой диагональный элемент в L при прямой подстановке.");
        }
        y[i] = (b[i] - sum) / L[i][i];
    }
    return y;
}

// Обратная подстановка для решения Ux = y (U - верхнетреугольная)
std::vector<double> solve_backward_U(const Matrix& U, const std::vector<double>& y) {
    const size_t n = U.size();
     if (n == 0 || U[0].size() != n || y.size() != n) {
         throw std::invalid_argument("Некорректные размеры для обратной подстановки.");
    }
    std::vector<double> x(n);

    for (int i = static_cast<int>(n) - 1; i >= 0; --i) {
        double sum = 0.0;
        for (size_t j = i + 1; j < n; ++j) {
            sum += U[i][j] * x[j];
        }
         if (std::abs(U[i][i]) < EPSILON) {
            throw std::runtime_error("Нулевой диагональный элемент в U при обратной подстановке.");
        }
        x[i] = (y[i] - sum) / U[i][i];
    }
    return x;
}

// Решение СЛАУ Ax=b с использованием LU-разложения
std::vector<double> solve_lu(const Matrix& A, const std::vector<double>& b) {
    try {
        auto [L, U] = LU_dec(A);
        std::vector<double> y = solve_forward_L(L, b);
        std::vector<double> x = solve_backward_U(U, y);
        return x;
    } catch (const std::runtime_error& e) {
        // Перебрасываем исключение с добавлением информации
        throw std::runtime_error(std::string("Ошибка при решении методом LU: ") + e.what());
    } catch (const std::invalid_argument& e) {
         throw std::invalid_argument(std::string("Ошибка при решении методом LU: ") + e.what());
    }
}

// Нахождение обратной матрицы методом Гаусса-Жордана
Matrix inverse_matrix(Matrix matrix) { // Принимаем копию
    const size_t n = matrix.size();
     if (n == 0 || matrix[0].size() != n) {
        throw std::invalid_argument("Матрица должна быть квадратной для нахождения обратной.");
    }

    // Создаем расширенную матрицу [A | E]
    Matrix augmented(n, std::vector<double>(2 * n));
    for (size_t i = 0; i < n; ++i) {
        for (size_t j = 0; j < n; ++j) {
            augmented[i][j] = matrix[i][j];
        }
        augmented[i][i + n] = 1.0; // Формируем единичную матрицу справа
    }

    // Прямой ход (приведение левой части к верхнетреугольной)
    for (size_t i = 0; i < n; ++i) {
        // Выбор главного элемента
        size_t maxRow = i;
        for (size_t k = i + 1; k < n; ++k) {
            if (std::abs(augmented[k][i]) > std::abs(augmented[maxRow][i])) {
                maxRow = k;
            }
        }

        if (std::abs(augmented[maxRow][i]) < EPSILON) {
            throw std::runtime_error("Матрица вырождена, обратной не существует.");
        }

        // Перестановка строк
        std::swap(augmented[i], augmented[maxRow]);

        // Нормализация i-й строки (делим на ведущий элемент)
        double pivot = augmented[i][i];
        for (size_t j = i; j < 2 * n; ++j) { // Делим всю строку
            augmented[i][j] /= pivot;
        }
         augmented[i][i] = 1.0; // Убедимся, что диагональный элемент ровно 1

        // Обнуление элементов под главным элементом
        for (size_t k = 0; k < n; ++k) {
            if (k != i) { // Для всех строк, кроме текущей
                double factor = augmented[k][i];
                if (std::abs(factor) < EPSILON) continue; // Пропускаем, если множитель мал
                for (size_t j = i; j < 2 * n; ++j) { // Начинаем с j=i
                    augmented[k][j] -= factor * augmented[i][j];
                }
                augmented[k][i] = 0.0; // Явно обнуляем для точности
            }
        }
    }

    // Извлечение обратной матрицы (правая часть расширенной матрицы)
    Matrix inv(n, std::vector<double>(n));
    for (size_t i = 0; i < n; ++i) {
        for (size_t j = 0; j < n; ++j) {
            inv[i][j] = augmented[i][j + n];
        }
    }

    return inv;
}

// Вычисление вектора невязки r = Ax - b
std::vector<double> compute_residual(const Matrix& A, const std::vector<double>& x, const std::vector<double>& b) {
    std::vector<double> Ax = multiply_matrix_vector(A, x);
    return diff_vector(Ax, b);
}


// --- Метод прогонки для трехдиагональной матрицы ---
std::vector<double> solve_tridiagonal(
    const std::vector<double>& a, // под-диагональ (индексы 1..n-1 в матрице)
    const std::vector<double>& b, // главная диагональ (индексы 0..n-1)
    const std::vector<double>& c, // над-диагональ (индексы 0..n-2)
    const std::vector<double>& d)
{
    const size_t n = b.size();
    if (a.size() != n - 1 || c.size() != n - 1 || d.size() != n || n == 0) {
        throw std::invalid_argument("Некорректные размеры векторов для метода прогонки.");
    }

    std::vector<double> P(n); // Прогоночные коэффициенты P_i
    std::vector<double> Q(n); // Прогоночные коэффициенты Q_i
    std::vector<double> x(n); // Решение

    // Прямой ход (вычисление P и Q)
    if (std::abs(b[0]) < EPSILON) {
         throw std::runtime_error("Нулевой элемент b[0] в методе прогонки.");
    }
    P[0] = -c[0] / b[0];
    Q[0] = d[0] / b[0];

    for (size_t i = 1; i < n; ++i) {
        double denominator;
        // Индексы для a и c: a[i-1] соответствует a_i в формулах, c[i-1] соответствует c_{i-1}
        if (i < n - 1) { // Для всех, кроме последней строки
             denominator = b[i] + a[i-1] * P[i-1];
             if (std::abs(denominator) < EPSILON) {
                 throw std::runtime_error("Нулевой знаменатель при вычислении прогоночных коэффициентов.");
             }
             P[i] = -c[i] / denominator; // c[i] соответствует c_i в формулах
             Q[i] = (d[i] - a[i-1] * Q[i-1]) / denominator;
        } else { // Последняя строка (i = n - 1)
             denominator = b[i] + a[i-1] * P[i-1];
             if (std::abs(denominator) < EPSILON) {
                 throw std::runtime_error("Нулевой знаменатель при вычислении последнего прогоночного коэффициента Q.");
             }
             // P[n-1] не нужен для обратного хода (или можно считать его 0)
             P[n-1] = 0.0; // По определению или по отсутствию c[n-1]
             Q[i] = (d[i] - a[i-1] * Q[i-1]) / denominator;
        }
    }

    // Обратный ход (вычисление x)
    x[n - 1] = Q[n - 1];
    for (int i = static_cast<int>(n) - 2; i >= 0; --i) {
        x[i] = P[i] * x[i + 1] + Q[i];
    }

    return x;
}


// Построение полной матрицы из трех диагоналей для проверки невязки
Matrix build_tridiagonal_matrix(
    const std::vector<double>& a, // под-диагональ (a_1 .. a_{n-1}), размер n-1
    const std::vector<double>& b, // главная диагональ (b_0 .. b_{n-1}), размер n
    const std::vector<double>& c) // над-диагональ (c_0 .. c_{n-2}), размер n-1
{
    const size_t n = b.size();
    if (a.size() != n - 1 || c.size() != n - 1 || n == 0) {
         throw std::invalid_argument("Некорректные размеры векторов для построения трехдиагональной матрицы.");
    }
    Matrix T(n, std::vector<double>(n, 0.0));
    for (size_t i = 0; i < n; ++i) {
        T[i][i] = b[i];
        if (i > 0) {
            T[i][i - 1] = a[i - 1]; // a[k] соответствует элементу T[k+1][k]
        }
        if (i < n - 1) {
            T[i][i + 1] = c[i]; // c[k] соответствует элементу T[k][k+1]
        }
    }
    return T;
}


// --- Основная программа ---

int main() {
    // Устанавливаем точность вывода по умолчанию
    const int precision = 15;
    std::cout << std::fixed << std::setprecision(precision);

    // --- ЗАДАЧА 1: Хорошо обусловленная матрица ---
    std::cout << "======================================================" << std::endl;
    std::cout << "   ЗАДАЧА 1.1: Хорошо обусловленная матрица" << std::endl;
    std::cout << "======================================================" << std::endl;

    Matrix A_good = {
        {-38.4000,   4.0300,   8.3800,   3.5300},
        { -8.3200, -81.2000,  -8.0900,  -3.6700},
        {  4.3300,   7.2100, -110.8000,  -2.6300},
        {  4.2200,   4.2200,   6.1500,  73.8000}
    };
    std::vector<double> b_good = {292.5900, -504.3200, -185.6600, -430.5000};
    std::vector<double> x_exact_good = {-7.0, 7.0, 2.0, -6.0};

    print_matrix(A_good, "Матрица A_good");
    print_vector(b_good, "Вектор b_good");
    print_vector(x_exact_good, "Точное решение x_exact_good");
    std::cout << std::endl;

    try {
        // 1. Решение методом Гаусса
        std::cout << "--- Метод Гаусса ---" << std::endl;
        std::vector<double> x_gauss_good = gauss(A_good, b_good);
        print_vector(x_gauss_good, "Решение x_gauss_good");

        // 2. Вычисление невязки для Гаусса
        std::vector<double> r_gauss_good = compute_residual(A_good, x_gauss_good, b_good);
        print_vector(r_gauss_good, "Невязка r_gauss_good", precision); // Повышенная точность для невязки
        double norm1_r_gauss_good = compute_vector_norm_1(r_gauss_good);
        double normInf_r_gauss_good = compute_vector_norm_inf(r_gauss_good);
        std::cout << "||r_gauss_good||_1   = " << norm1_r_gauss_good << std::endl;
        std::cout << "||r_gauss_good||_inf = " << normInf_r_gauss_good << std::endl;

        // 3. Вычисление погрешности для Гаусса
        std::vector<double> err_gauss_good = diff_vector(x_gauss_good, x_exact_good);
        print_vector(err_gauss_good, "Абс. погрешность err_gauss_good", precision);
        double norm1_err_gauss_good = compute_vector_norm_1(err_gauss_good);
        double normInf_err_gauss_good = compute_vector_norm_inf(err_gauss_good);
        std::cout << "||err_gauss_good||_1   = " << norm1_err_gauss_good << std::endl;
        std::cout << "||err_gauss_good||_inf = " << normInf_err_gauss_good << std::endl;

        // 4. Вычисление относительной погрешности для Гаусса
        double norm1_exact_good = compute_vector_norm_1(x_exact_good);
        double normInf_exact_good = compute_vector_norm_inf(x_exact_good);
        std::cout << "||x_exact_good||_1   = " << norm1_exact_good << std::endl;
        std::cout << "||x_exact_good||_inf = " << normInf_exact_good << std::endl;
        if (norm1_exact_good > EPSILON)
             std::cout << "Отн. погрешность (1-норма)   = " << norm1_err_gauss_good / norm1_exact_good << std::endl;
        if (normInf_exact_good > EPSILON)
             std::cout << "Отн. погрешность (inf-норма) = " << normInf_err_gauss_good / normInf_exact_good << std::endl;
        std::cout << std::endl;


        // 5. Решение методом LU
        std::cout << "--- Метод LU ---" << std::endl;
        std::vector<double> x_lu_good = solve_lu(A_good, b_good);
        print_vector(x_lu_good, "Решение x_lu_good");

        // 6. Вычисление невязки для LU
        std::vector<double> r_lu_good = compute_residual(A_good, x_lu_good, b_good);
        print_vector(r_lu_good, "Невязка r_lu_good", precision);
        double norm1_r_lu_good = compute_vector_norm_1(r_lu_good);
        double normInf_r_lu_good = compute_vector_norm_inf(r_lu_good);
        std::cout << "||r_lu_good||_1   = " << norm1_r_lu_good << std::endl;
        std::cout << "||r_lu_good||_inf = " << normInf_r_lu_good << std::endl;

        // 7. Вычисление погрешности для LU
        std::vector<double> err_lu_good = diff_vector(x_lu_good, x_exact_good);
        print_vector(err_lu_good, "Абс. погрешность err_lu_good", precision);
        double norm1_err_lu_good = compute_vector_norm_1(err_lu_good);
        double normInf_err_lu_good = compute_vector_norm_inf(err_lu_good);
        std::cout << "||err_lu_good||_1   = " << norm1_err_lu_good << std::endl;
        std::cout << "||err_lu_good||_inf = " << normInf_err_lu_good << std::endl;

        // 8. Вычисление относительной погрешности для LU
         if (norm1_exact_good > EPSILON)
             std::cout << "Отн. погрешность (1-норма)   = " << norm1_err_lu_good / norm1_exact_good << std::endl;
        if (normInf_exact_good > EPSILON)
             std::cout << "Отн. погрешность (inf-норма) = " << normInf_err_lu_good / normInf_exact_good << std::endl;
        std::cout << std::endl;

        // 9. Нахождение обратной матрицы
        std::cout << "--- Обратная матрица и число обусловленности ---" << std::endl;
        Matrix A_inv_good = inverse_matrix(A_good);
        print_matrix(A_inv_good, "Обратная матрица A_inv_good", precision);

        // 10. Проверка A * A_inv = E
        Matrix Product_good = multiply(A_good, A_inv_good);
        Matrix E_good = create_identity_matrix(A_good.size());
        std::cout << "Проверка A_good * A_inv_good:" << std::endl;
        // print_matrix(Product_good, "Результат A*A_inv", 15); // Можно раскомментировать для детального просмотра
        if (are_matrices_close(Product_good, E_good)) {
            std::cout << "A_good * A_inv_good близка к единичной матрице." << std::endl;
        } else {
            std::cout << "ПРЕДУПРЕЖДЕНИЕ: A_good * A_inv_good НЕ близка к единичной матрице!" << std::endl;
             print_matrix(Product_good, "Результат A*A_inv", precision); // Показать результат, если не сошлось
        }

        // 11. Вычисление числа обусловленности
        double norm1_A_good = compute_matrix_norm_1(A_good);
        double normInf_A_good = compute_matrix_norm_inf(A_good);
        double norm1_A_inv_good = compute_matrix_norm_1(A_inv_good);
        double normInf_A_inv_good = compute_matrix_norm_inf(A_inv_good);

        double cond1_good = norm1_A_good * norm1_A_inv_good;
        double condInf_good = normInf_A_good * normInf_A_inv_good;

        std::cout << "||A_good||_1       = " << norm1_A_good << std::endl;
        std::cout << "||A_good||_inf     = " << normInf_A_good << std::endl;
        std::cout << "||A_inv_good||_1   = " << norm1_A_inv_good << std::endl;
        std::cout << "||A_inv_good||_inf = " << normInf_A_inv_good << std::endl;
        std::cout << "cond_1(A_good)   = " << cond1_good << std::endl;
        std::cout << "cond_inf(A_good) = " << condInf_good << std::endl;

    } catch (const std::exception& e) {
        std::cerr << "Ошибка при обработке хорошо обусловленной матрицы: " << e.what() << std::endl;
    }
    std::cout << std::endl;


    // --- ЗАДАЧА 2: Плохо обусловленная матрица ---
    std::cout << "======================================================" << std::endl;
    std::cout << "   ЗАДАЧА 1.2: Плохо обусловленная матрица" << std::endl;
    std::cout << "======================================================" << std::endl;

    Matrix A_bad = {
        {-190.3270,  189.7600,  -18.0160,   72.0640},
        {-194.5200,  193.9530,  -18.4320,   73.7280},
        {-919.4800,  919.4800,  -99.2470,  398.6800},
        {-219.3900,  219.3900,  -23.7720,   95.5110}
    };
    std::vector<double> b_bad = {1414.2810, 1446.7620, 7705.5000, 1845.1920};
    std::vector<double> x_exact_bad = {1.0, 2.0, 20.0, 22.0};

    print_matrix(A_bad, "Матрица A_bad");
    print_vector(b_bad, "Вектор b_bad");
    print_vector(x_exact_bad, "Точное решение x_exact_bad");
    std::cout << std::endl;

     try {
        // 1. Решение методом Гаусса
        std::cout << "--- Метод Гаусса ---" << std::endl;
        std::vector<double> x_gauss_bad = gauss(A_bad, b_bad);
        print_vector(x_gauss_bad, "Решение x_gauss_bad");

        // 2. Вычисление невязки для Гаусса
        std::vector<double> r_gauss_bad = compute_residual(A_bad, x_gauss_bad, b_bad);
        print_vector(r_gauss_bad, "Невязка r_gauss_bad", precision); // Повышенная точность
        double norm1_r_gauss_bad = compute_vector_norm_1(r_gauss_bad);
        double normInf_r_gauss_bad = compute_vector_norm_inf(r_gauss_bad);
        std::cout << "||r_gauss_bad||_1   = " << norm1_r_gauss_bad << std::endl;
        std::cout << "||r_gauss_bad||_inf = " << normInf_r_gauss_bad << std::endl;

        // 3. Вычисление погрешности для Гаусса
        std::vector<double> err_gauss_bad = diff_vector(x_gauss_bad, x_exact_bad);
        print_vector(err_gauss_bad, "Абс. погрешность err_gauss_bad", precision);
        double norm1_err_gauss_bad = compute_vector_norm_1(err_gauss_bad);
        double normInf_err_gauss_bad = compute_vector_norm_inf(err_gauss_bad);
        std::cout << "||err_gauss_bad||_1   = " << norm1_err_gauss_bad << std::endl;
        std::cout << "||err_gauss_bad||_inf = " << normInf_err_gauss_bad << std::endl;

        // 4. Вычисление относительной погрешности для Гаусса
        double norm1_exact_bad = compute_vector_norm_1(x_exact_bad);
        double normInf_exact_bad = compute_vector_norm_inf(x_exact_bad);
        std::cout << "||x_exact_bad||_1   = " << norm1_exact_bad << std::endl;
        std::cout << "||x_exact_bad||_inf = " << normInf_exact_bad << std::endl;
         if (norm1_exact_bad > EPSILON)
             std::cout << "Отн. погрешность (1-норма)   = " << norm1_err_gauss_bad / norm1_exact_bad << std::endl;
        if (normInf_exact_bad > EPSILON)
             std::cout << "Отн. погрешность (inf-норма) = " << normInf_err_gauss_bad / normInf_exact_bad << std::endl;
        std::cout << std::endl;

        // 5. Решение методом LU
        std::cout << "--- Метод LU ---" << std::endl;
         // Для плохо обусловленной матрицы LU без перестановок может не сработать или дать плохой результат
         // Попробуем, но ожидаем возможных проблем
         try {
            std::vector<double> x_lu_bad = solve_lu(A_bad, b_bad);
            print_vector(x_lu_bad, "Решение x_lu_bad");

             // 6. Вычисление невязки для LU
            std::vector<double> r_lu_bad = compute_residual(A_bad, x_lu_bad, b_bad);
            print_vector(r_lu_bad, "Невязка r_lu_bad", precision);
            double norm1_r_lu_bad = compute_vector_norm_1(r_lu_bad);
            double normInf_r_lu_bad = compute_vector_norm_inf(r_lu_bad);
            std::cout << "||r_lu_bad||_1   = " << norm1_r_lu_bad << std::endl;
            std::cout << "||r_lu_bad||_inf = " << normInf_r_lu_bad << std::endl;

            // 7. Вычисление погрешности для LU
            std::vector<double> err_lu_bad = diff_vector(x_lu_bad, x_exact_bad);
            print_vector(err_lu_bad, "Абс. погрешность err_lu_bad", precision);
            double norm1_err_lu_bad = compute_vector_norm_1(err_lu_bad);
            double normInf_err_lu_bad = compute_vector_norm_inf(err_lu_bad);
            std::cout << "||err_lu_bad||_1   = " << norm1_err_lu_bad << std::endl;
            std::cout << "||err_lu_bad||_inf = " << normInf_err_lu_bad << std::endl;

            // 8. Вычисление относительной погрешности для LU
             if (norm1_exact_bad > EPSILON)
                 std::cout << "Отн. погрешность (1-норма)   = " << norm1_err_lu_bad / norm1_exact_bad << std::endl;
            if (normInf_exact_bad > EPSILON)
                 std::cout << "Отн. погрешность (inf-норма) = " << normInf_err_lu_bad / normInf_exact_bad << std::endl;

         } catch (const std::runtime_error& e) {
             std::cerr << "Ошибка при решении методом LU для плохо обусловленной матрицы: " << e.what() << std::endl;
             std::cerr << "Это ожидаемо для LU без перестановок для таких матриц." << std::endl;
         }
        std::cout << std::endl;


        // 9. Нахождение обратной матрицы
        std::cout << "--- Обратная матрица и число обусловленности ---" << std::endl;
        Matrix A_inv_bad = inverse_matrix(A_bad);
        print_matrix(A_inv_bad, "Обратная матрица A_inv_bad", precision);

        // 10. Проверка A * A_inv = E
        Matrix Product_bad = multiply(A_bad, A_inv_bad);
        Matrix E_bad = create_identity_matrix(A_bad.size());
        std::cout << "Проверка A_bad * A_inv_bad:" << std::endl;
        // print_matrix(Product_bad, "Результат A*A_inv", 15);
        if (are_matrices_close(Product_bad, E_bad, 1e-5)) { // Увеличим допуск для плохой матрицы
            std::cout << "A_bad * A_inv_bad близка к единичной матрице (с допуском 1e-5)." << std::endl;
        } else {
            std::cout << "ПРЕДУПРЕЖДЕНИЕ: A_bad * A_inv_bad НЕ близка к единичной матрице!" << std::endl;
            print_matrix(Product_bad, "Результат A*A_inv", precision); // Показать результат, если не сошлось
        }

        // 11. Вычисление числа обусловленности
        double norm1_A_bad = compute_matrix_norm_1(A_bad);
        double normInf_A_bad = compute_matrix_norm_inf(A_bad);
        double norm1_A_inv_bad = compute_matrix_norm_1(A_inv_bad);
        double normInf_A_inv_bad = compute_matrix_norm_inf(A_inv_bad);

        // Используем double для хранения больших чисел обусловленности
        double cond1_bad = norm1_A_bad * norm1_A_inv_bad;
        double condInf_bad = normInf_A_bad * normInf_A_inv_bad;

        std::cout << std::scientific; // Переключимся на научную нотацию для больших чисел
        std::cout << "||A_bad||_1       = " << norm1_A_bad << std::endl;
        std::cout << "||A_bad||_inf     = " << normInf_A_bad << std::endl;
        std::cout << "||A_inv_bad||_1   = " << norm1_A_inv_bad << std::endl;
        std::cout << "||A_inv_bad||_inf = " << normInf_A_inv_bad << std::endl;
        std::cout << "cond_1(A_bad)   = " << cond1_bad << std::endl;
        std::cout << "cond_inf(A_bad) = " << condInf_bad << std::endl;
        std::cout << std::fixed << std::setprecision(precision); // Вернем обычную нотацию


    } catch (const std::exception& e) {
        std::cerr << "Ошибка при обработке плохо обусловленной матрицы: " << e.what() << std::endl;
    }
    std::cout << std::endl;


     // --- ЗАДАЧА 3: Трехдиагональная матрица ---
    std::cout << "======================================================" << std::endl;
    std::cout << "   ЗАДАЧА 2: Трехдиагональная система (Метод прогонки)" << std::endl;
    std::cout << "======================================================" << std::endl;

    // Вариант 13
    // a = (a_1, ..., a_{n-1}) - поддиагональ (size n-1)
    // b = (b_0, ..., b_{n-1}) - диагональ (size n)
    // c = (c_0, ..., c_{n-2}) - наддиагональ (size n-1)
    // d = (d_0, ..., d_{n-1}) - правая часть (size n)
    // n = 6
    
    std::vector<double> a_tri = {1.0, -1.0,  0.0, 1.0, 1.0}; // size 5 (n-1)
    std::vector<double> b_tri = {65.0, 91.0, 147.0, 78.0, 99.0, 132.0}; // size 6 (n)
    std::vector<double> c_tri = {1.0, -1.0,  1.0, 0.0, 1.0}; // size 5 (n-1)
    std::vector<double> d_tri = {6.0, 10.0,  14.0, 8.0, 10.0, 13.0}; // size 6 (n)

    std::cout << "Трехдиагональная система:" << std::endl;
    print_vector(a_tri, "Поддиагональ a (a1..a5)");
    print_vector(b_tri, "Диагональ b (b0..b5)");
    print_vector(c_tri, "Наддиагональ c (c0..c4)");
    print_vector(d_tri, "Правая часть d (d0..d5)");
    std::cout << std::endl;

    try {
        // 1. Решение методом прогонки
        std::vector<double> x_tridiagonal = solve_tridiagonal(a_tri, b_tri, c_tri, d_tri);
        print_vector(x_tridiagonal, "Решение x_tridiagonal");

        // 2. Построение полной матрицы T для проверки невязки
        Matrix T_tridiagonal = build_tridiagonal_matrix(a_tri, b_tri, c_tri);
        // print_matrix(T_tridiagonal, "Полная матрица T_tridiagonal"); // Раскомментировать для просмотра

        // 3. Вычисление невязки r = Tx - d
        std::vector<double> r_tridiagonal = compute_residual(T_tridiagonal, x_tridiagonal, d_tri);
        print_vector(r_tridiagonal, "Невязка r_tridiagonal", 15);
        double norm1_r_tridiagonal = compute_vector_norm_1(r_tridiagonal);
        double normInf_r_tridiagonal = compute_vector_norm_inf(r_tridiagonal);
        std::cout << "||r_tridiagonal||_1   = " << norm1_r_tridiagonal << std::endl;
        std::cout << "||r_tridiagonal||_inf = " << normInf_r_tridiagonal << std::endl;


    } catch (const std::exception& e) {
         std::cerr << "Ошибка при решении трехдиагональной системы: " << e.what() << std::endl;
    }

    return 0;
}