# -*- coding: utf-8 -*-
import argparse

def process_file(input_path, output_path, max_decimals=None):
    """
    Обрабатывает входной файл, преобразуя строки с числами в заданный формат.

    Args:
        input_path (str): Путь к входному файлу.
        output_path (str): Путь к выходному файлу.
        max_decimals (int, optional): Максимальное количество знаков после запятой.
    """
    try:
        with open(input_path, 'r', encoding='utf-8') as infile, \
             open(output_path, 'w', encoding='utf-8') as outfile:
            for line in infile:
                parts = line.strip().split()
                if len(parts) == 4:
                    try:
                        # Извлекаем числовые значения
                        num1_str = parts[0]
                        num2_str = parts[1]
                        float3 = float(parts[2])
                        float4 = float(parts[3])

                        # Форматируем числа с плавающей точкой, если указан --max
                        if max_decimals is not None:
                            num3_str = f'{float3:.{max_decimals}f}'
                            num4_str = f'{float4:.{max_decimals}f}'
                        else:
                            num3_str = parts[2]
                            num4_str = parts[3]
                        
                        # Собираем и записываем отформатированную строку
                        output_line = f"[ {num1_str}  ], [ {num2_str}],[    {num3_str}],[    {num4_str}],\n"
                        outfile.write(output_line)
                    except (ValueError, IndexError):
                        # Если строка не содержит 4 числа, записываем ее как есть
                        outfile.write(line)
                else:
                    # Если в строке не 4 элемента, записываем ее как есть
                    outfile.write(line)
        print(f"Обработка завершена. Результат сохранен в файл: {output_path}")
    except FileNotFoundError:
        print(f"Ошибка: Не удалось найти входной файл '{input_path}'")
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")

if __name__ == "__main__":
    # Настройка парсера аргументов командной строки
    parser = argparse.ArgumentParser(
        description="Скрипт для форматирования числовых данных из файла.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        "input_filename",
        help="Имя исходного файла для чтения.\nПример: out.txt"
    )
    parser.add_argument(
        "output_filename",
        help="Имя файла для записи результата.\nПример: formatted_out.txt"
    )
    parser.add_argument(
        "--max",
        type=int,
        dest="max_decimals",
        default=None,
        help="Максимальное количество знаков после запятой.\nПример: --max=5"
    )

    # Запуск основной функции с переданными аргументами
    args = parser.parse_args()
    process_file(args.input_filename, args.output_filename, args.max_decimals)
