#include <iostream>
#include <vector>
#include <cmath>
#include <iomanip>
#include <numeric>   // Для std::iota
#include <stdexcept> // Для обработки ошибок
#include <algorithm>


// Вычисляет значение полинома Лагранжа в точке xp
double lagrange_interpolation(const std::vector<double>& x_nodes,
                               const std::vector<double>& y_nodes,
                               double xp) {
    if (x_nodes.size() != y_nodes.size() || x_nodes.empty()) {
        throw std::invalid_argument("Некорректные узлы для интерполяции Лагранжа.");
    }
    size_t n = x_nodes.size();
    double interpolated_value = 0.0;

    for (size_t i = 0; i < n; ++i) {
        double basis_polynomial = 1.0;
        for (size_t j = 0; j < n; ++j) {
            if (i != j) {
                // Проверка на совпадение узлов
                if (std::abs(x_nodes[i] - x_nodes[j]) < 1e-9) {
                    throw std::runtime_error(
                        "Обнаружены дублирующиеся x‑узлы, для интерполяции Лагранжа требуются различные узлы.");
                }
                basis_polynomial *= (xp - x_nodes[j]) / (x_nodes[i] - x_nodes[j]);
            }
        }
        interpolated_value += y_nodes[i] * basis_polynomial;
    }
    return interpolated_value;
}

void lagrange_method() {
    std::cout << "============================================" << std::endl;
    std::cout << "       Пример интерполяции Лагранжа         " << std::endl;
    std::cout << "============================================" << std::endl;

    // Данные из Варианта 12, Таблица 1
    std::vector<double> x_nodes = {0.0, 1.0, 2.0, 3.0, 4.0,
                                   5.0, 6.0, 7.0, 8.0, 9.0};
    std::vector<double> y_nodes = {9.0, 8.2, 7.4, 6.6, 5.8,
                                   4.9, 4.1, 3.3, 2.5, 1.8};

    std::cout << "Входные точки:\n";
    std::cout << std::fixed << std::setprecision(1);
    std::cout << "xi\t| yi\n";
    std::cout << "--------|-------\n";
    for (size_t i = 0; i < x_nodes.size(); ++i) {
        std::cout << x_nodes[i] << "\t " << y_nodes[i] << std::endl;
    }
    std::cout << std::endl;

    // Точки, в которых оцениваем полином Лагранжа
    std::vector<double> x_eval = {0.1, 0.2, 0.5, 1.5, 2.5, 3.5,
                                  4.5, 5.5, 6.5, 7.5, 8.5, 8.6, 8.8};

    std::cout << "Значения полинома Лагранжа:\n";
    std::cout << std::fixed << std::setprecision(6);
    std::cout << "xp\t| P(xp)\n";
    std::cout << "--------|-------------\n";
    try {
        for (double xp : x_eval) {
            double yp = lagrange_interpolation(x_nodes, y_nodes, xp);
            std::cout << xp << "\t " << yp << std::endl;
        }
    } catch (const std::exception& e) {
        std::cerr << "Ошибка интерполяции Лагранжа: " << e.what() << std::endl;
    }
    std::cout << std::endl;
}

// Функция для интерполяции
double func(double x) {
    // f(x) = sin(x^2) * e^(−x^2)
    return sin(x * x) * exp(-x * x);
}

// Структура для хранения данных сплайна
struct SplineData {
    std::vector<double> x; // Узлы xi
    std::vector<double> y; // Значения yi = f(xi)
    std::vector<double> M; // Вторые производные Mi в узлах
    double h;              // Шаг (равномерная сетка)
};

// Решение трёхдиагональной СЛАУ методом Томаса
void solve_tridiagonal(const std::vector<double>& a,
                      const std::vector<double>& b,
                      const std::vector<double>& c,
                      const std::vector<double>& d,
                      std::vector<double>& x,
                      int n) {
    if (n <= 0) return;

    std::vector<double> c_prime(n);
    std::vector<double> d_prime(n);

    // Прямой ход
    c_prime[0] = c[0] / b[0];
    d_prime[0] = d[0] / b[0];
    for (int i = 1; i < n; ++i) {
        double m = 1.0 / (b[i] - a[i] * c_prime[i - 1]);
        c_prime[i] = (i < n - 1) ? (c[i] * m) : 0.0;
        d_prime[i] = (d[i] - a[i] * d_prime[i - 1]) * m;
    }

    // Обратный ход
    x[n - 1] = d_prime[n - 1];
    for (int i = n - 2; i >= 0; --i) {
        x[i] = d_prime[i] - c_prime[i] * x[i + 1];
    }
}

// Построение естественного кубического сплайна
SplineData build_natural_cubic_spline(double a, double b, int n) {
    if (n < 3) {
        throw std::invalid_argument("Для кубического сплайна нужно минимум 3 точки.");
    }

    SplineData spline;
    spline.x.resize(n);
    spline.y.resize(n);
    spline.M.resize(n);
    spline.h = (b - a) / (n - 1);

    // Генерация узлов и значений функции
    for (int i = 0; i < n; ++i) {
        spline.x[i] = a + i * spline.h;
        spline.y[i] = func(spline.x[i]);
    }

    // Настройка трёхдиагональной системы для M1 … M_{n-2}
    int system_size = n - 2; // Решаем для M1..M_{n-2}
    if (system_size <= 0) {
        spline.M[0] = 0.0;
        if (n > 1) spline.M[n - 1] = 0.0;
        if (n == 3) { // Отдельный случай n=3
            double d1 = (6.0 / spline.h) *
                        ((spline.y[2] - spline.y[1]) / spline.h -
                         (spline.y[1] - spline.y[0]) / spline.h);
            double b1 = 4.0 * spline.h;
            spline.M[1] = d1 / b1;
        }
        return spline;
    }

    // Коэффициенты системы
    std::vector<double> sub_diag(system_size);
    std::vector<double> main_diag(system_size);
    std::vector<double> super_diag(system_size);
    std::vector<double> rhs(system_size);

    // Формулы для естественного сплайна
    for (int k = 0; k < system_size; ++k) {
        int i = k + 1;
        main_diag[k] = 4.0 * spline.h;
        rhs[k] = (6.0 / spline.h) *
                 (spline.y[i + 1] - 2.0 * spline.y[i] + spline.y[i - 1]);
        if (k > 0) sub_diag[k] = spline.h;
        if (k < system_size - 1) super_diag[k] = spline.h;
    }

    // Подготовка к TDMA
    std::vector<double> tdma_a(system_size), tdma_b(system_size),
        tdma_c(system_size), tdma_d(system_size), solution_M(system_size);

    for (int k = 0; k < system_size; ++k) {
        tdma_b[k] = main_diag[k];
        tdma_d[k] = rhs[k];
        tdma_a[k] = (k > 0) ? sub_diag[k] : 0;
        tdma_c[k] = (k < system_size - 1) ? super_diag[k] : 0;
    }

    solve_tridiagonal(tdma_a, tdma_b, tdma_c, tdma_d, solution_M, system_size);

    // Запись полученных Mi
    spline.M[0] = 0.0;
    for (int k = 0; k < system_size; ++k) {
        spline.M[k + 1] = solution_M[k];
    }
    spline.M[n - 1] = 0.0;

    return spline;
}

// Оценка значения сплайна S(xp) в точке xp
double evaluate_spline(const SplineData& spline, double xp) {
    size_t n = spline.x.size();
    if (n < 2) {
        throw std::runtime_error("Сплайн не построен или содержит слишком мало точек.");
    }

    // Поиск интервала [xi, x_{i+1}], в котором находится xp
    auto it = std::lower_bound(spline.x.begin(), spline.x.end(), xp);
    int i = std::distance(spline.x.begin(), it);

    if (i == n) i = n - 1;
    if (i > 0 && (xp < spline.x[i] || std::abs(xp - spline.x[i]) < 1e-9)) --i;

    i = std::max(0, std::min((int)n - 2, i));

    double xi = spline.x[i];
    double xi1 = spline.x[i + 1];
    double yi = spline.y[i];
    double yi1 = spline.y[i + 1];
    double Mi = spline.M[i];
    double Mi1 = spline.M[i + 1];
    double h = spline.h;

    if (std::abs(h) < 1e-9) {
        throw std::runtime_error("Нулевая ширина интервала сплайна.");
    }

    // Формула кубического сплайна на [xi, xi+1]
    double term1 = Mi * pow(xi1 - xp, 3) / (6.0 * h);
    double term2 = Mi1 * pow(xp - xi, 3) / (6.0 * h);
    double term3 = (yi / h - Mi * h / 6.0) * (xi1 - xp);
    double term4 = (yi1 / h - Mi1 * h / 6.0) * (xp - xi);

    return term1 + term2 + term3 + term4;
}

void сubic_spline_method() {
    std::cout << "============================================" << std::endl;
    std::cout << "     Пример кубической сплайн‑интерполяции   " << std::endl;
    std::cout << "============================================" << std::endl;

    double a = 0.0, b = 3.0;
    int n = 20;
    std::cout << "Функция: f(x) = sin(x^2) * exp(-x^2)\n";
    std::cout << "Интервал: [" << a << ", " << b << "]\n";
    std::cout << "Число узлов n: " << n << "\n\n";

    SplineData spline;
    try {
        spline = build_natural_cubic_spline(a, b, n);
    } catch (const std::exception& e) {
        std::cerr << "Ошибка построения сплайна: " << e.what() << std::endl;
        return;
    }

    std::cout << "Узлы (xi, yi=f(xi)) и вторые производные (Mi):\n";
    std::cout << std::fixed << std::setprecision(8);
    std::cout << " xi\t\t| yi=f(xi)\t| Mi\n";
    std::cout << "--------|---------------|---------------\n";
    for (int i = 0; i < n; ++i) {
        std::cout << spline.x[i] << "\t " << spline.y[i] << "\t " << spline.M[i]
                  << std::endl;
    }
    std::cout << std::endl;

    std::cout << "Сравнение в серединах интервалов:\n";
    std::cout << " x_mid\t\t| S(x_mid)\t| f(x_mid)\t| |S‑f|\n";
    std::cout << "---------------|---------------|---------------|---------------\n";

    for (int i = 0; i < n - 1; ++i) {
        double x_mid = (spline.x[i] + spline.x[i + 1]) / 2.0;
        double spline_val = 0.0;
        double exact_val = func(x_mid);
        try {
            spline_val = evaluate_spline(spline, x_mid);
        } catch (const std::exception& e) {
            std::cerr << "Ошибка вычисления сплайна в x=" << x_mid << ": "
                      << e.what() << std::endl;
            spline_val = NAN;
        }
        double error = std::abs(spline_val - exact_val);

        std::cout << x_mid << "\t " << spline_val << "\t " << exact_val << "\t "
                  << error << std::endl;
    }
    std::cout << std::endl;
}

int main() {
    lagrange_method();

    std::cout << "\n\n";

    сubic_spline_method();

    return 0;
}
