#include <iostream>
#include <cmath>
#include <iomanip>
#include <limits> // Для numeric_limits

// --- Функции для Задачи 2 ---
const double LN10 = std::log(10.0); // Натуральный логарифм 10

// f(x) = 2x - log10(x) - 2
double f2(double x) {
    if (x <= 0)
        return std::numeric_limits<double>::quiet_NaN(); // Не определено

    return 2.0 * x - std::log10(x) - 2.0;
}

// f'(x) = 2 - 1 / (x * ln(10))
double f2_prime(double x) {
     if (x <= 0)
         return std::numeric_limits<double>::quiet_NaN();

     double denom = x * LN10;

     if (std::fabs(denom) < 1e-15)
         return std::numeric_limits<double>::infinity(); // Избегаем деления на 0

     return 2.0 - 1.0 / denom;
}

// g1(x) = (log10(x) + 2) / 2  (Для корня x=1)
double g1_task2(double x) {
    if (x <= 0) return std::numeric_limits<double>::quiet_NaN();
    return (std::log10(x) + 2.0) / 2.0;
}

// g2(x) = 10^(2x - 2) (Для корня около 0)
double g2_task2(double x) {
    // Основание 10 в степени (2x-2)
    return std::pow(10.0, 2.0 * x - 2.0);
}

// --- Метод Простой Итерации ---
void simple_iteration_method(double x0, double epsilon, double (*g)(double), int max_iterations = 1000) {
    std::cout << "Метод Простой Итерации\n";
    std::cout << "Начальное приближение x0 = " << x0 << ", Точность: " << epsilon << "\n";
    std::cout << "---------------------------------------------------\n";
    std::cout << std::setw(5) << "Iter" << std::setw(20) << "x_k" << std::setw(20) << "x_{k+1}=g(x_k)" << std::setw(20) << "|x_{k+1}-x_k|" << "\n";
    std::cout << "---------------------------------------------------\n";

    double x_prev = x0;
    double x_next;
    int iterations = 0;
    bool converged = false;

    for (iterations = 1; iterations <= max_iterations; ++iterations) {
        x_next = g(x_prev);

        // Проверка на NaN или бесконечность
        if (std::isnan(x_next) || std::isinf(x_next)) {
            std::cout << "Ошибка: Получено NaN или бесконечность на итерации " << iterations << std::endl;
            break;
        }

        double diff = std::fabs(x_next - x_prev);

        std::cout << std::fixed << std::setprecision(7);
        std::cout << std::setw(5) << iterations << std::setw(20) << x_prev << std::setw(20) << x_next << std::setw(20) << diff << "\n";

        if (diff < epsilon) {
            converged = true;
            break;
        }
        x_prev = x_next;
    }

    std::cout << "---------------------------------------------------\n";
    if (converged) {
        std::cout << "Результат:\n";
        std::cout << "Приближенный корень: " << std::fixed << std::setprecision(7) << x_next << std::endl;
        std::cout << "Количество итераций: " << iterations << std::endl;
        std::cout << "f(корень) = " << f2(x_next) << std::endl;
    } else {
        std::cout << "Метод не сошелся за " << max_iterations << " итераций." << std::endl;
        std::cout << "Последнее приближение: " << x_next << std::endl;
    }
}

// --- Метод Ньютона ---
void newton_method(double x0, double epsilon, double (*f)(double), double (*f_prime)(double), int max_iterations = 1000) {
    std::cout << "Метод Ньютона\n";
    std::cout << "Начальное приближение x0 = " << x0 << ", Точность: " << epsilon << "\n";
    std::cout << "-------------------------------------------------------------\n";
    std::cout << std::setw(5) << "Iter" << std::setw(18) << "x_k" << std::setw(18) << "f(x_k)" << std::setw(18) << "f'(x_k)" << std::setw(18) << "|delta_x|" << "\n";
    std::cout << "-------------------------------------------------------------\n";

    double x_prev = x0;
    double x_next = 0;
    int iterations = 0;
    bool converged = false;
    double diff =0;

    for (iterations = 1; iterations <= max_iterations; ++iterations) {
        double fx = f(x_prev);
        double fpx = f_prime(x_prev);

        if (std::isnan(fx) || std::isnan(fpx) || std::isinf(fx) || std::isinf(fpx)) {
            std::cout << "Ошибка: Получено NaN или бесконечность (f или f') на итерации " << iterations << std::endl;
            break;
        }

         if (std::fabs(fpx) < 1e-15) {
            std::cout << "Ошибка: Производная близка к нулю на итерации " << iterations << ". f'(" << x_prev << ") = " << fpx << std::endl;
            break;
        }

        double delta_x = fx / fpx;
        x_next = x_prev - delta_x;
        diff = std::fabs(delta_x);

        std::cout << std::fixed << std::setprecision(7);
        std::cout << std::setw(5) << iterations << std::setw(18) << x_prev << std::setw(18) << fx << std::setw(18) << fpx << std::setw(18) << diff << "\n";


        if (f2(x_next) < epsilon) {
             converged = true;
            break;
        }

        x_prev = x_next;
    }

    std::cout << "-------------------------------------------------------------\n";
     if (converged) {
        std::cout << "Результат:\n";
        std::cout << "Приближенный корень: " << std::fixed << std::setprecision(7) << x_next << std::endl;
        std::cout << "Количество итераций: " << iterations << std::endl;
        std::cout << "f(корень) = " << f(x_next) << std::endl;
         std::cout << "diff = " << diff << std::endl;
     } else {
         std::cout << "Метод не сошелся за " << max_iterations << " итераций." << std::endl;
         std::cout << "Последнее приближение: " << x_next << std::endl;
     }
}


int main() {
    double eps2 = 0.001;

    std::cout << "ЗАДАЧА 2 (2x - log10(x) - 2 = 0) \n";
    double x0_root1 = 1e-5;
    double x0_root2 = 10;

    std::cout << "\n--- Метод Простой Итерации (g(x)=10^(2x-2)) ---\n";
    std::cout << "*** Поиск корня в интервале  [0.01, 0.1] ***\n";
    simple_iteration_method(x0_root1, eps2, g2_task2);

    std::cout << "\n\n*** Поиск корня в интервале (корень x=1) ***\n";
    simple_iteration_method(x0_root2, eps2, g1_task2);

    std::cout << "\n--- Метод Ньютона ---\n";
    std::cout << "*** Поиск корня в интервале [0.01, 0.1] ***\n";
    newton_method(x0_root1, eps2, f2, f2_prime);

    std::cout << "\n\n*** Поиск корня в интервале [0.5, 1.5] (корень x=1) ***\n";
    newton_method(x0_root2, eps2, f2, f2_prime);

    return 0;
}