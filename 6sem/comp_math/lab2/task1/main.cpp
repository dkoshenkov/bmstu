#include <iostream>

/*
 *Отделить корни уравнения и найти его методом половинного деления с точностью
 *
 * x^3 -2x + 4 = 0
 *
 * eps = 0.001
 */


#include <iostream>
#include <cmath>
#include <iomanip>

double f1(double x) {
    return x * x * x - 2 * x + 4;
}

// Метод Половинного Деления
void bisection_method(double a, double b, double epsilon) {
    if (f1(a) * f1(b) >= 0) {
        std::cout << "Ошибка: f(a) и f(b) должны иметь разные знаки." << std::endl;
        // Дополнительная проверка на случай точного корня на границе
        if (std::fabs(f1(a)) < 1e-9) {
             std::cout << "Точный корень найден на границе a: " << a << std::endl;
             return;
        }

        if (std::fabs(f1(b)) < 1e-9) {
             std::cout << "Точный корень найден на границе b: " << b << std::endl;
             return;
        }

        return;
    }

    double c = a;
    int iterations = 0;

    std::cout << "Метод Половинного Деления для f(x) = x^3 - 2x + 4\n";
    std::cout << "Интервал: [" << a << ", " << b << "], Точность: " << epsilon << "\n";
    std::cout << "---------------------------------------------------\n";
    std::cout << std::setw(5) << "Iter" << std::setw(15) << "a" << std::setw(15) << "b" << std::setw(15) << "c=(a+b)/2" << std::setw(15) << "f(c)" << std::setw(15) << "|b-a|/2" << "\n";
    std::cout << "---------------------------------------------------\n";


    while ((b - a) > epsilon) {
        ++iterations;
        c = (a + b) / 2.0;

        std::cout << std::fixed << std::setprecision(7);
        std::cout << std::setw(5) << iterations << std::setw(15) << a << std::setw(15) << b << std::setw(15) << c << std::setw(15) << f1(c) << std::setw(15) << (b-a)/2.0 << "\n";

        // Проверка, если корень найден точно (маловероятно с double)
        if (std::fabs(f1(c)) < 1e-15) {
            std::cout << "Точный корень найден: " << c << std::endl;
            break;
        }

        if (f1(c) * f1(a) < 0)
            b = c;
        else
            a = c;

    }

     c = (a + b) / 2.0;

    std::cout << "---------------------------------------------------\n";
    std::cout << "Результат:\n";
    std::cout << "Приближенный корень: " << std::fixed << std::setprecision(7) << c << std::endl;
    std::cout << "Количество итераций: " << iterations << std::endl;
    std::cout << "f(корень) = " << f1(c) << std::endl;
    std::cout << "корень в диапазоне: {" << a << " " << b<<"}" << std::endl;
    std::cout << "Достигнутая точность |b-a|/2 = " << (b-a)/2.0 << std::endl;
}

int main() {
    double a1 = -15, b1 = 15, eps1 = 0.001;

    bisection_method(a1, b1, eps1);

    return 0;
}