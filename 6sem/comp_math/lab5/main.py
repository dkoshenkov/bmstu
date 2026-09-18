import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

# Определяем точную функцию y(x)
def exact_solution_y(x):
    # Предотвращаем ошибку log(0) или отрицательного значения под корнем,
    # если (1 + e^x)/2 <= 0, хотя для x >= 0 это не должно произойти.
    # Также ln(z) требует z > 0. Для x >= 0, (1 + e^x)/2 всегда > 1, так что ln будет > 0.
    # И 2 * ln(...) + 1 должно быть >= 0.
    # y(0) = sqrt(2*ln((1+1)/2) + 1) = sqrt(2*ln(1) + 1) = sqrt(0+1) = 1. Это соответствует начальному условию.
    numerator = 1 + np.exp(x)
    denominator = 2
    term_inside_ln = numerator / denominator
    # Обработка случая, когда term_inside_ln может привести к ошибке в ln или sqrt
    if np.any(term_inside_ln <= 0):
        raise ValueError("Значение внутри ln должно быть положительным.")
    val_ln = np.log(term_inside_ln)
    term_inside_sqrt = 2 * val_ln + 1
    if np.any(term_inside_sqrt < 0):
        # Это может произойти, если 2*ln((1+e^x)/2) < -1, т.е. ln((1+e^x)/2) < -0.5
        # (1+e^x)/2 < e^(-0.5) ~ 0.606.
        # 1+e^x < 1.212
        # e^x < 0.212
        # x < ln(0.212) ~ -1.55. Это вне нашего диапазона [0, 1].
        raise ValueError("Значение под квадратным корнем не может быть отрицательным.")
    return np.sqrt(term_inside_sqrt)

# Определяем функцию f(x, y)
def f_xy(x, y):
    # Предотвращаем деление на ноль, если y = 0
    if np.any(y == 0):
        # Проверим, действительно ли y(x) может быть 0 на [0,1]
        # y(x) = sqrt(2*ln((1+e^x)/2)+1). y(x)=0 => 2*ln((1+e^x)/2)+1 = 0
        # ln((1+e^x)/2) = -0.5
        # (1+e^x)/2 = exp(-0.5) ~ 0.6065
        # 1+e^x = 1.213
        # e^x = 0.213
        # x = ln(0.213) ~ -1.54. Это вне нашего диапазона [0,1].
        # Значит, на [0,1] y(x) не обращается в ноль.
        pass # На нашем интервале и для нашего y(x) это не должно произойти
    return np.exp(x) / ((1 + np.exp(x)) * y)

# Задаем промежуток и шаг
x0 = 0.0
xn = 1.0
h = 0.1
x_values = np.arange(x0, xn + h, h) # +h чтобы включить xn

# Вычисляем значения y(x) для точного решения
y_values_exact = exact_solution_y(x_values)

# Вычисляем значения f(x, y(x))
f_values_on_exact_solution = f_xy(x_values, y_values_exact)

# Создаем DataFrame для таблицы значений y(x)
data_y = {'x': x_values, 'y(x)': y_values_exact}
df_y = pd.DataFrame(data_y)

# Создаем DataFrame для таблицы значений f(x, y(x))
data_f = {'x': x_values, 'f(x, y(x))': f_values_on_exact_solution}
df_f = pd.DataFrame(data_f)

# Вывод таблиц
print("Таблица значений для точного решения y(x):")
print(df_y.to_string(index=False))
print("\n" + "="*40 + "\n")
print("Таблица значений для f(x, y(x)) (правая часть ДУ на точном решении):")
print(df_f.to_string(index=False))

# Построение графиков
plt.figure(figsize=(12, 6))

# График y(x)
plt.subplot(1, 2, 1) # 1 строка, 2 столбца, 1-й график
plt.plot(x_values, y_values_exact, marker='o', linestyle='-', label='Точное решение y(x)')
plt.title('График точного решения y(x)')
plt.xlabel('x')
plt.ylabel('y(x)')
plt.grid(True)
plt.legend()

# График f(x, y(x))
plt.subplot(1, 2, 2) # 1 строка, 2 столбца, 2-й график
plt.plot(x_values, f_values_on_exact_solution, marker='s', linestyle='--', color='r', label='f(x, y(x))')
plt.title('График функции f(x, y(x))')
plt.xlabel('x')
plt.ylabel('f(x, y(x))')
plt.grid(True)
plt.legend()

plt.tight_layout() # Для лучшего расположения графиков
plt.show()