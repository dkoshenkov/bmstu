import os
import sys
import subprocess
import argparse
from pathlib import Path


def convert_docx_with_word(docx_path, dpi=300, output_format="png", output_dir=None, output_name=None):
    """
    Конвертирует .docx документ в изображения высокого качества
    с помощью Microsoft Word на macOS.

    :param docx_path: Путь к файлу .docx.
    :param dpi: Желаемое значение DPI (точек на дюйм) для изображений.
    :param output_format: Конечный формат изображений ('png' или 'jpeg').
    :param output_dir: Путь для сохранения результата (по умолчанию - папка со скриптом).
    :param output_name: Имя выходного файла (по умолчанию - имя исходного файла).
    """
    if sys.platform != "darwin":
        print("Ошибка: Этот скрипт предназначен только для macOS и требует Microsoft Word.")
        sys.exit(1)

    docx_file = Path(docx_path).resolve()

    if not docx_file.is_file():
        print(f"Ошибка: Файл не найден по пути: {docx_path}")
        sys.exit(1)

    # Определяем папку для сохранения (по умолчанию - папка со скриптом)
    if output_dir:
        save_dir = Path(output_dir).resolve()
    else:
        save_dir = Path(__file__).parent.resolve()
    
    save_dir.mkdir(parents=True, exist_ok=True)

    # Определяем имя выходного файла (по умолчанию - имя исходного файла)
    if output_name:
        base_name = output_name
    else:
        base_name = docx_file.stem

    # Временный PDF файл
    temp_pdf = save_dir / f"{base_name}_temp.pdf"

    print(f"Используем Microsoft Word для конвертации '{docx_file.name}' в PDF...")
    print(f"Сохранение в: {save_dir}")

    # Улучшенный AppleScript с дополнительной обработкой ошибок
    applescript = f'''
    tell application "Microsoft Word"
        activate
        delay 1
        try
            set theDoc to open file name "{str(docx_file)}" with read only
            delay 1
            save as theDoc file name "{str(temp_pdf)}" file format format PDF
            close theDoc saving no
            return "success"
        on error errMsg number errNum
            return "error: " & errMsg & " (" & errNum & ")"
        end try
    end tell
    '''

    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            check=True,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if "error:" in result.stdout:
            print(f"\n--- ОШИБКА Word ---")
            print(f"Word вернул ошибку: {result.stdout}")
            sys.exit(1)
        
        print("\nКонвертация в PDF завершена! Начинаю извлечение страниц...")

        if not temp_pdf.exists():
            print(f"!!! Ошибка: PDF файл не был создан по пути '{temp_pdf}'.")
            sys.exit(1)

        # Конвертируем PDF в изображения с высоким разрешением
        print(f"Конвертирую PDF в {output_format.upper()} с DPI={dpi}...")

        # Метод 1: Попытка использовать sips для первой страницы
        output_file = save_dir / f"{base_name}.{output_format}"
        
        try:
            subprocess.run(
                [
                    'sips',
                    '-s', 'format', output_format,
                    '--setProperty', 'dpiHeight', str(dpi),
                    '--setProperty', 'dpiWidth', str(dpi),
                    '--resampleHeightWidthMax', '8192',
                    str(temp_pdf),
                    '--out', str(output_file)
                ],
                check=True,
                capture_output=True,
                text=True
            )
            print(f"  ✓ Создано изображение: {output_file.name}")
            
        except subprocess.CalledProcessError:
            print("\nsips не справился с PDF. Пробую метод через qlmanage...")
            
            # Метод 2: QuickLook preview
            try:
                # Создаём preview с высоким разрешением
                preview_size = int(dpi * 11)  # Примерно размер для A4 при заданном DPI
                subprocess.run(
                    [
                        'qlmanage',
                        '-t',
                        '-s', str(preview_size),
                        '-o', str(save_dir),
                        str(temp_pdf)
                    ],
                    check=True,
                    capture_output=True,
                    text=True
                )
                
                # qlmanage создаёт файл с суффиксом .png
                generated_file = save_dir / f"{temp_pdf.stem}.png"
                
                if generated_file.exists():
                    if output_format == 'png':
                        generated_file.rename(output_file)
                    else:
                        # Конвертируем в нужный формат
                        subprocess.run(
                            ['sips', '-s', 'format', output_format, str(generated_file), '--out', str(output_file)],
                            check=True,
                            capture_output=True,
                            text=True
                        )
                        generated_file.unlink()
                    print(f"  ✓ Создано изображение: {output_file.name}")
                    
            except Exception as e:
                print(f"  Метод qlmanage также не сработал: {e}")
                print("\n⚠️  Попробуйте установить ImageMagick для лучшей обработки:")
                print("  brew install imagemagick")
                print(f"  Затем используйте: convert -density {dpi} input.pdf output.png")

        # Очистка временного PDF
        try:
            if temp_pdf.exists():
                temp_pdf.unlink()
        except OSError as e:
            print(f"Предупреждение: Не удалось удалить временный PDF: {e}")

        # Проверяем, было ли создано изображение
        if output_file.exists():
            print(f"\n✅ Обработка завершена!")
            print(f"📁 Результат: {output_file}")
        else:
            print("\n⚠️  Не удалось создать изображение. Проверьте PDF файл вручную.")

    except subprocess.CalledProcessError as e:
        print("\n--- ОШИБКА ---")
        print("Не удалось выполнить AppleScript. Возможные причины:")
        print("1. Microsoft Word не установлен.")
        print("2. Проблемы с правами доступа (System Settings -> Privacy & Security -> Automation).")
        print("3. Неверный путь к файлу или файл повреждён.")
        print(f"\nДетали ошибки: {e.stderr}")
        print(f"\nПолный путь к файлу: {docx_file}")
        print(f"Файл существует: {docx_file.exists()}")
        sys.exit(1)
    except subprocess.TimeoutExpired:
        print("\n--- ОШИБКА ---")
        print("Превышено время ожидания при конвертации документа.")
        print("Возможно, документ слишком большой или Word не отвечает.")
        sys.exit(1)
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Конвертирует .docx документ в изображение высокого качества через Microsoft Word',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Примеры использования:
  # Базовое использование (сохраняет в папку со скриптом с оригинальным именем)
  python %(prog)s document.docx
  
  # С указанием DPI и формата
  python %(prog)s document.docx --dpi 1000 --format png
  
  # Указать свою папку для сохранения
  python %(prog)s hw1/titul.docx --dir ./results
  
  # Указать своё имя файла
  python %(prog)s document.docx --name my_image
  
  # Полный контроль
  python %(prog)s hw1/titul.docx --dpi 600 --format jpeg --dir ./output --name title_page
        '''
    )
    
    parser.add_argument('docx_file', help='Путь к файлу .docx')
    parser.add_argument('--dpi', type=int, default=300, help='DPI для изображения (по умолчанию: 300)')
    parser.add_argument('--format', choices=['png', 'jpeg'], default='png', help='Формат изображения (по умолчанию: png)')
    parser.add_argument('--dir', '-d', dest='output_dir', help='Папка для сохранения (по умолчанию: папка со скриптом)')
    parser.add_argument('--name', '-n', dest='output_name', help='Имя выходного файла без расширения (по умолчанию: имя исходного файла)')
    
    args = parser.parse_args()
    
    convert_docx_with_word(
        args.docx_file,
        dpi=args.dpi,
        output_format=args.format,
        output_dir=args.output_dir,
        output_name=args.output_name
    )
