#include <iostream>
#include <vector>
#include <cmath>
#include <iomanip>
#include <functional> // Для std::function
#include <algorithm>  // Для std::min/max
#include <string>     // Для std::string

// --- Константы и параметры задачи ---
constexpr double X0 = 0.0;
constexpr double Y0 = 1.0;
constexpr double XN = 1.0;
constexpr double EPSILON = 1e-3;      // Заданная точность контроля локальной погрешности
constexpr double H_INITIAL = 0.1;   // Начальный шаг

// Коэффициенты для адаптации шага
constexpr double SAFETY_FACTOR = 0.2;       // Фактор безопасности для нового шага
constexpr double GROW_LIMIT_FACTOR = 1.2;   // Максимальное относительное увеличение шага
constexpr double SHRINK_LIMIT_FACTOR = 0.2; // Минимальное относительное уменьшение шага
constexpr double H_MIN_ALLOWED = 1e-7;      // Минимально допустимый абсолютный шаг
constexpr double H_MAX_ALLOWED = (XN - X0) / 20.0; // Максимально допустимый абсолютный шаг (например, 1/5 интервала)


// --- Функция правой части ДУ ---
// y' = f(x, y) = e^x / ((1 + e^x)y)
double derivative(double x, double y) {
    if (std::abs(y) < 1e-12) { // Избегаем деления на очень малое число
        // Эта ситуация не должна возникать для данного y(0)=1 и интервала [0,1]
        // так как y(x) = sqrt(2*ln((1+e^x)/2)+1) всегда > 0.
        // y(0)=1, y(1) ~ 1.496739
        // Если бы возникла, возвращаем большое значение, чтобы сигнализировать о проблеме
        // или чтобы алгоритм уменьшил шаг.
        return (y >= 0 ? 1.0 : -1.0) * 1e12;
    }
    return std::exp(x) / ((1.0 + std::exp(x)) * y);
}

// --- Функция точного решения ---
// y(x) = sqrt(2 * ln((1 + e^x)/2) + 1)
double exact_solution(double x) {
    double term_inside_ln = (1.0 + std::exp(x)) / 2.0;
    // Для x в [0,1], 1 <= e^x <= e. (1+1)/2=1, (1+e)/2 ~ 1.859. term_inside_ln > 0.
    double val_ln = std::log(term_inside_ln); // ln(1)=0, ln(1.859) ~ 0.619
    // 2*val_ln+1: 2*0+1=1, 2*0.619+1 ~ 2.238. term_inside_sqrt > 0.
    double term_inside_sqrt = 2.0 * val_ln + 1.0;
    return std::sqrt(term_inside_sqrt);
}

// --- Вспомогательная структура для возврата результата одного шага метода ---
struct StepResult {
    double y_next; // Значение y на следующем шаге
    int f_evals;   // Количество вычислений функции f(x,y) для этого шага
};

// --- Реализация одного шага метода Эйлера-Коши (модифицированный Эйлер/Хойна) ---
// Порядок точности p = 2
StepResult euler_cauchy_step(double x, double y, double h, const std::function<double(double, double)>& f) {
    double f_xy = f(x, y);
    double y_predictor = y + h * f_xy;
    double f_xph_ypred = f(x + h, y_predictor);
    double y_next = y + (h / 2.0) * (f_xy + f_xph_ypred);
    return {y_next, 2}; // 2 вызова функции f(x,y)
}

// --- Реализация одного шага метода Рунге-Кутты 4-го порядка ---
// Порядок точности p = 4
StepResult runge_kutta_4_step(double x, double y, double h, const std::function<double(double, double)>& f) {
    double k1 = f(x, y);
    double k2 = f(x + h / 2.0, y + h * k1 / 2.0);
    double k3 = f(x + h / 2.0, y + h * k2 / 2.0);
    double k4 = f(x + h, y + h * k3);
    double y_next = y + (h / 6.0) * (k1 + 2.0 * k2 + 2.0 * k3 + k4);
    return {y_next, 4}; // 4 вызова функции f(x,y)
}

// --- Общая функция для решения ОДУ с автоматическим выбором шага ---
void solve_ode_auto_step(
    const std::string& method_name_debug, // Для отладочных сообщений
    const std::function<StepResult(double, double, double, const std::function<double(double, double)>&)>& method_step_func,
    int method_order_p, // Порядок точности метода
    const std::function<double(double, double)>& deriv_func,
    double x_start, double y_start, double x_end,
    double initial_h, double target_epsilon,
    std::vector<double>& out_x_values, // Вектор для сохранения значений x
    std::vector<double>& out_y_values, // Вектор для сохранения значений y
    long long& total_f_evaluations)    // Общее число вычислений f(x,y)
{
    out_x_values.clear();
    out_y_values.clear();
    total_f_evaluations = 0;

    double x_current = x_start;
    double y_current = y_start;
    double h_current = initial_h;

    out_x_values.push_back(x_current);
    out_y_values.push_back(y_current);

    int max_iterations_safety = 200000; // Защита от бесконечного цикла
    int iterations_count = 0;

    while (x_current < x_end && iterations_count < max_iterations_safety) {
        iterations_count++;

        if (x_current + h_current > x_end) {
            h_current = x_end - x_current; // Корректируем последний шаг, чтобы точно попасть в x_end
        }
        // Предотвращаем слишком маленький шаг, который может вызвать проблемы
        if (h_current < H_MIN_ALLOWED / 10.0 && x_current < x_end) { // Добавлено /10.0 для большей устойчивости
             std::cerr << "Warning (" << method_name_debug << "): Step size h_current (" << h_current
                       << ") became extremely small at x = " << x_current << ". Stopping to prevent infinite loop." << std::endl;
             break;
        }


        // 1. Делаем один "большой" шаг h_current
        //    y1 - результат после одного шага h
        StepResult res_h = method_step_func(x_current, y_current, h_current, deriv_func);
        double y1 = res_h.y_next;
        long long f_evals_for_y1 = res_h.f_evals;

        // 2. Делаем два "маленьких" шага h_current / 2.0
        //    y2 - результат после двух шагов h/2
        StepResult res_h_half1 = method_step_func(x_current, y_current, h_current / 2.0, deriv_func);
        double y_temp_half = res_h_half1.y_next;
        StepResult res_h_half2 = method_step_func(x_current + h_current / 2.0, y_temp_half, h_current / 2.0, deriv_func);
        double y2 = res_h_half2.y_next;
        long long f_evals_for_y2 = res_h_half1.f_evals + res_h_half2.f_evals;

        // Все эти вычисления функции f(x,y) были произведены
        total_f_evaluations += f_evals_for_y1 + f_evals_for_y2;

        // Оценка локальной погрешности по правилу Рунге
        // R = |y1 - y2| / (2^p - 1)
        double error_estimate_R = std::abs(y1 - y2) / (std::pow(2.0, method_order_p) - 1.0);

        // Принимаем шаг, если оценка погрешности в норме ИЛИ если шаг уже предельно мал
        if (error_estimate_R <= target_epsilon || h_current < H_MIN_ALLOWED) {
            x_current += h_current;
            y_current = y2; // y2 (результат с меньшим шагом) обычно точнее

            out_x_values.push_back(x_current);
            out_y_values.push_back(y_current);

            // Адаптация шага для следующей итерации: пытаемся увеличить шаг
            if (error_estimate_R == 0.0) { // Если ошибка нулевая (редко, но возможно)
                h_current = std::min(h_current * GROW_LIMIT_FACTOR, H_MAX_ALLOWED);
            } else if (error_estimate_R < target_epsilon) {
                // Формула для оптимального шага, p - порядок метода, p+1 - порядок локальной погрешности O(h^(p+1))
                double optimal_factor = SAFETY_FACTOR * std::pow(target_epsilon / error_estimate_R, 1.0 / (method_order_p + 1.0));
                h_current *= std::min(optimal_factor, GROW_LIMIT_FACTOR); // Ограничиваем рост
            }
            // Если error_estimate_R был близок к target_epsilon, шаг может не измениться сильно, или даже уменьшится немного из-за SAFETY_FACTOR.
            // Это нормально. Если error_estimate_R == target_epsilon, optimal_factor = SAFETY_FACTOR.
            // Если h_current был H_MIN_ALLOWED, то он и останется H_MIN_ALLOWED после этой логики, если не увеличится.

        } else { // Отвергаем шаг: оценка погрешности слишком велика. Уменьшаем h_current.
            double shrink_factor = SAFETY_FACTOR * std::pow(target_epsilon / error_estimate_R, 1.0 / (method_order_p + 1.0));
            // shrink_factor будет < SAFETY_FACTOR, так как (target_epsilon / error_estimate_R) < 1.
            h_current *= std::max(shrink_factor, SHRINK_LIMIT_FACTOR); // Ограничиваем уменьшение
            // x_current и y_current не меняются, цикл повторится с новым, уменьшенным h_current
        }

        // Применяем глобальные ограничения на шаг
        h_current = std::min(h_current, H_MAX_ALLOWED);
        h_current = std::max(h_current, H_MIN_ALLOWED);

        if (x_current >= x_end) break; // Выход, если достигли или немного пересекли конец промежутка
    }

     if (iterations_count >= max_iterations_safety && x_current < x_end) {
        std::cerr << "Warning (" << method_name_debug << "): Max iterations (" << max_iterations_safety
                  << ") reached. Solution might be incomplete. x_current = " << x_current << std::endl;
    }
}


// --- Функция для вывода результатов в виде таблицы ---
void print_results_table(const std::string& method_name,
                         const std::vector<double>& x_vals,
                         const std::vector<double>& y_vals,
                         long long f_evals,
                         double y_xn_exact_val,
                         double target_eps)
{
    std::cout << "\n--- " << method_name << " ---\n";
    std::cout << std::fixed << std::setprecision(7); // Увеличим точность вывода
    std::cout << "   x        y_computed   y_exact(x)   |error(x)|\n";
    std::cout << "------------------------------------------------------\n";

    if (x_vals.empty()) {
        std::cout << "Нет данных для вывода.\n";
        return;
    }

    // Вывод некоторых промежуточных точек и последней точки
    const int points_to_show = 10;
    size_t step = x_vals.size() / points_to_show;
    if (step == 0) step = 1;

    for (size_t i = 0; i < x_vals.size(); ++i) {
        if (i % step == 0 || i == x_vals.size() - 1) { // Показываем каждую step-ую точку и последнюю
            double x = x_vals[i];
            double y_comp = y_vals[i];
            double y_ex = exact_solution(x);
            double error_local = std::abs(y_comp - y_ex);
            std::cout << std::setw(9) << x << "   "
                      << std::setw(12) << y_comp << "   "
                      << std::setw(12) << y_ex << "   "
                      << std::setw(12) << error_local << std::endl;
        }
    }
    if (x_vals.size() > 1 && (x_vals.size()-1) % step != 0) { // Если последняя точка не была показана
         double x = x_vals.back();
         double y_comp = y_vals.back();
         double y_ex = exact_solution(x);
         double error_local = std::abs(y_comp - y_ex);
         std::cout << std::setw(9) << x << "   "
                   << std::setw(12) << y_comp << "   "
                   << std::setw(12) << y_ex << "   "
                   << std::setw(12) << error_local << "(Last actual point)\n";
    }


    std::cout << "------------------------------------------------------\n";
    double y_xn_computed = y_vals.back();
    double final_global_error = std::abs(y_xn_computed - y_xn_exact_val);

    std::cout << "Конечное значение x:                         " << x_vals.back() << std::endl;
    std::cout << "Конечное значение y(x_n) (вычисленное):    " << y_xn_computed << std::endl;
    std::cout << "Конечное значение y(x_n) (точное):         " << y_xn_exact_val << std::endl;
    std::cout << "Абсолютная погрешность в конечной точке:   " << final_global_error << std::endl;
    std::cout << "Заданная точность контроля лок. погр. eps: " << target_eps << std::endl;
    std::cout << "Число вычислений правой части f(x,y):      " << f_evals << std::endl;
    std::cout << "Число успешных шагов (выч. значений y):  " << (x_vals.empty() ? 0 : x_vals.size() - 1) << std::endl;

    if (final_global_error <= target_eps) {
        std::cout << "Точность в конечной точке по отношению к epsilon ДОСТИГНУТА." << std::endl;
    } else {
        std::cout << "Точность в конечной точке по отношению к epsilon НЕ ДОСТИГНУТА (глобальная ошибка > eps_локальный)." << std::endl;
        std::cout << " (Это ожидаемо, т.к. epsilon контролирует локальную погрешность на шаге, а не итоговую глобальную)." << std::endl;
    }
     std::cout << " (Обычно глобальная ошибка для метода порядка p при контроле локальной погрешности eps ~ O(eps) или чуть хуже)" << std::endl;
}

int main() {
    std::vector<double> x_euler_cauchy, y_euler_cauchy;
    long long f_evals_ec = 0;

    std::vector<double> x_rk4, y_rk4;
    long long f_evals_rk4 = 0;

    std::cout << "Решение ОДУ y' = e^x / ((1 + e^x)y) на [" << X0 << ", " << XN << "]" << std::endl;
    std::cout << "y(" << X0 << ") = " << Y0 << ", epsilon (для контроля локальной погрешности) = " << EPSILON
              << ", H_initial = " << H_INITIAL << std::endl;

    // Решение методом Эйлера-Коши
    solve_ode_auto_step("Euler-Cauchy", euler_cauchy_step, 2, derivative, X0, Y0, XN, H_INITIAL, EPSILON,
                        x_euler_cauchy, y_euler_cauchy, f_evals_ec);

    // Решение методом Рунге-Кутты 4
    solve_ode_auto_step("Runge-Kutta 4", runge_kutta_4_step, 4, derivative, X0, Y0, XN, H_INITIAL, EPSILON,
                        x_rk4, y_rk4, f_evals_rk4);

    // Точное значение в конечной точке
    double y_xn_exact_val = exact_solution(XN);

    // Пункт 2 и 3: Вывод результатов и сравнение
    print_results_table("Метод Эйлера-Коши (с автовыбором шага)", x_euler_cauchy, y_euler_cauchy, f_evals_ec, y_xn_exact_val, EPSILON);
    print_results_table("Метод Рунге-Кутты 4 (с автовыбором шага)", x_rk4, y_rk4, f_evals_rk4, y_xn_exact_val, EPSILON);

    std::cout << "\n--- Сравнение общего числа вычислений f(x,y) (Пункт 2) ---" << std::endl;
    std::cout << "Метод Эйлера-Коши: " << f_evals_ec << " вызовов f(x,y)" << std::endl;
    std::cout << "Метод Рунге-Кутты 4: " << f_evals_rk4 << " вызовов f(x,y)" << std::endl;

    std::cout << "\n--- Сравнение числа успешных шагов (Пункт 2) ---" << std::endl;
    std::cout << "Метод Эйлера-Коши: " << (x_euler_cauchy.empty() ? 0 : x_euler_cauchy.size() - 1) << " шагов" << std::endl;
    std::cout << "Метод Рунге-Кутты 4: " << (x_rk4.empty() ? 0 : x_rk4.size() - 1) << " шагов" << std::endl;


    // Пункт 4: Данные для построения графиков
    std::cout << "\n\n--- Данные для построения графиков (Пункт 4) ---" << std::endl;
    std::cout << "Скопируйте эти данные в инструмент для построения графиков (например, Python с Matplotlib, Excel, Gnuplot и т.д.)." << std::endl;
    std::cout << std::fixed << std::setprecision(7);

    std::cout << "\n# Точное решение (пример, шаг 0.05 для гладкости):" << std::endl;
    std::cout << "# x_exact, y_exact" << std::endl;
    for (double x_val = X0; x_val <= XN + 1e-9; x_val += 0.05) {
        // +1e-9 для корректной обработки XN из-за ошибок округления
         if (x_val > XN) x_val = XN;
        std::cout << x_val << ", " << exact_solution(x_val) << std::endl;
         if (x_val == XN) break;
    }

    std::cout << "\n# Данные для метода Эйлера-Коши:" << std::endl;
    std::cout << "# x_ec, y_ec" << std::endl;
    for(size_t i=0; i < x_euler_cauchy.size(); ++i) {
        std::cout << x_euler_cauchy[i] << ", " << y_euler_cauchy[i] << std::endl;
    }

    std::cout << "\n# Данные для метода Рунге-Кутты 4:" << std::endl;
    std::cout << "# x_rk4, y_rk4" << std::endl;
    for(size_t i=0; i < x_rk4.size(); ++i) {
        std::cout << x_rk4[i] << ", " << y_rk4[i] << std::endl;
    }

    return 0;
}