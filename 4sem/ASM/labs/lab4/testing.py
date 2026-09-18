import random

# Генерация случайной матрицы 5x5
matrix = [[random.randint(-100, 100) for _ in range(5)] for _ in range(5)]
print("Original matrix:")
for row in matrix:
    print(*row)


# Функция, которая находит все положительные элементы строки и помещает их на главную диагональ
def positive_diagonal(matrix):
    for i in range(len(matrix)):
        positive_elements = [x for x in matrix[i] if x > 0]
        if positive_elements:
            matrix[i][i] = max(positive_elements)
    return matrix


# Вызов функции и вывод результата
result_matrix = positive_diagonal(matrix)
print("\nResult matrix:")
for row in result_matrix:
    print(*row)
