import os
import sys
import subprocess
from pathlib import Path


def convert_pptx_with_keynote(pptx_path, dpi=300, output_format="png"):
    """
    Конвертирует .pptx презентацию в папку с изображениями высокого качества
    с помощью Apple Keynote на macOS.

    :param pptx_path: Путь к файлу .pptx.
    :param dpi: Желаемое значение DPI (точек на дюйм) для изображений.
    :param output_format: Конечный формат изображений ('png' или 'jpeg').
    """
    if sys.platform != "darwin":
        print("Ошибка: Этот скрипт предназначен только для macOS и требует Apple Keynote.")
        sys.exit(1)

    pptx_file = Path(pptx_path).resolve()

    if not pptx_file.is_file():
        print(f"Ошибка: Файл не найден по пути: {pptx_path}")
        sys.exit(1)

    output_dir = Path(f"{pptx_file.stem}_images_hq").resolve()
    output_dir.mkdir(exist_ok=True)

    # Временная папка для исходных изображений TIFF
    temp_output_dir = output_dir / "temp_tiff"
    temp_output_dir.mkdir(exist_ok=True)

    print(f"Используем Keynote для конвертации '{pptx_file.name}' в TIFF для максимального качества...")

    # Экспортируем в TIFF для сохранения максимального качества
    applescript = f'''
    tell application "Keynote"
        set doc to open POSIX file "{str(pptx_file)}"
        export doc to POSIX file "{str(temp_output_dir)}" as slide images with properties {{image format:TIFF, all stages:false, skipped slides:false}}
        close doc without saving
    end tell
    '''

    try:
        subprocess.run(
            ['osascript', '-e', applescript],
            check=True,
            capture_output=True,
            text=True
        )
        print("\nКонвертация в Keynote (TIFF) завершена! Начинаю обработку изображений...")

        # Keynote может создавать дополнительную вложенную папку
        subfolders = [d for d in temp_output_dir.iterdir() if d.is_dir()]
        source_folder = temp_output_dir
        if subfolders:
            source_folder = subfolders[0]

        files_to_process = sorted(source_folder.glob("*.tiff"))

        if not files_to_process:
            print(f"!!! Ошибка: Не удалось найти TIFF файлы для обработки в '{source_folder}'.")
            sys.exit(1)

        print(f"Найдено {len(files_to_process)} файлов. Устанавливаю DPI и конвертирую в {output_format.upper()}...")

        for i, f in enumerate(files_to_process):
            new_name = output_dir / f"slide_{i + 1}.{output_format}"
            try:
                # Используем sips для установки DPI и конвертации
                subprocess.run(
                    [
                        'sips',
                        '--setProperty', 'dpiHeight', str(dpi),
                        '--setProperty', 'dpiWidth', str(dpi),
                        '--resampleHeightWidthMax', '4096', # Опционально: ограничиваем максимальный размер
                        '-s', 'format', output_format,
                        str(f),
                        '--out', str(new_name)
                    ],
                    check=True,
                    capture_output=True,
                    text=True
                )
            except subprocess.CalledProcessError as e:
                print(f"  Не удалось обработать файл '{f.name}' с помощью sips: {e.stderr}")
            except Exception as e:
                print(f"  Не удалось обработать '{f.name}': {e}")

        # Очистка временных файлов
        try:
            for temp_file in files_to_process:
                temp_file.unlink()
            if subfolders:
                subfolders[0].rmdir()
            temp_output_dir.rmdir()

        except OSError as e:
            print(f"Предупреждение: Не удалось полностью очистить временную папку: {e}")

        print("\n✅ Обработка изображений успешно завершена.")

    except subprocess.CalledProcessError as e:
        print("\n--- ОШИБКА ---")
        print("Не удалось выполнить AppleScript. Возможные причины:")
        print("1. Apple Keynote не установлен.")
        print("2. Проблемы с правами доступа (System Settings -> Privacy & Security -> Automation).")
        print(f"Детали ошибки: {e.stderr}")
        sys.exit(1)
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Использование: python {sys.argv[0]} <путь_к_файлу.pptx> [dpi] [format]")
        print("Пример: python convert.py my_presentation.pptx 300 png")
        sys.exit(1)

    pptx_file_path = sys.argv[1]
    user_dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 300
    user_format = sys.argv[3] if len(sys.argv) > 3 else "png"

    convert_pptx_with_keynote(pptx_file_path, dpi=user_dpi, output_format=user_format)