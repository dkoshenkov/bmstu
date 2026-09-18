import os
import sys
import subprocess
from pathlib import Path


def convert_pptx_with_keynote(pptx_path):
    # Новое имя файла будет .jpg, это стандарт
    output_format = "jpg"

    if sys.platform != "darwin":
        print("Ошибка: Этот скрипт предназначен только для macOS и требует Apple Keynote.")
        sys.exit(1)

    pptx_file = Path(pptx_path).resolve()
    if not pptx_file.is_file():
        print(f"Ошибка: Файл не найден по пути: {pptx_path}")
        sys.exit(1)

    output_dir = Path(f"{pptx_file.stem}_images").resolve()
    output_dir.mkdir(exist_ok=True)

    print(f"Используем Keynote для конвертации '{pptx_file.name}' в JPEG...")

    # Используем версию AppleScript без параметра 'quality' для максимальной совместимости
    applescript = f'''
       tell application "Keynote"
           set doc to open POSIX file "{str(pptx_file)}"
           export doc to POSIX file "{str(output_dir)}" as slide images with properties {{image format:JPEG, all stages:false, skipped slides:false}}
           close doc without saving
       end tell
       '''

    try:
        subprocess.run(['osascript', '-e', applescript], check=True, capture_output=True)
        print(f"\nКонвертация в Keynote завершена! Начинаю переименование файлов...")

        # --- НАЧАЛО ФИНАЛЬНОГО ИСПРАВЛЕНИЯ ---

        subfolders = [d for d in output_dir.iterdir() if d.is_dir()]

        if not subfolders:
            # Ищем и .jpg, и .jpeg файлы в основной папке
            files_to_rename = sorted(
                list(output_dir.glob("*.jpg")) + list(output_dir.glob("*.jpeg"))
            )
            source_folder_to_delete = None
        else:
            keynote_output_folder = subfolders[0]
            # Ищем и .jpg, и .jpeg файлы в подпапке
            files_to_rename = sorted(
                list(keynote_output_folder.glob("*.jpg")) + list(keynote_output_folder.glob("*.jpeg"))
            )
            source_folder_to_delete = keynote_output_folder

        if not files_to_rename:
            # Исправленный текст ошибки
            print(f"!!! Ошибка: Не удалось найти .jpg или .jpeg файлы для переименования.")
            sys.exit(1)

        print(f"Найдено {len(files_to_rename)} файлов. Переименовываю...")
        for i, f in enumerate(files_to_rename):
            new_name = output_dir / f"slide_{i + 1}.{output_format}"
            try:
                f.rename(new_name)
            except Exception as e:
                print(f"  Не удалось переименовать '{f.name}': {e}")

        if source_folder_to_delete:
            try:
                source_folder_to_delete.rmdir()
            except OSError:
                print(f"Предупреждение: Не удалось удалить папку '{source_folder_to_delete.name}'.")

        print("\n✅ Переименование успешно завершено.")
        # --- КОНЕЦ ФИНАЛЬНОГО ИСПРАВЛЕНИЯ ---

    except subprocess.CalledProcessError as e:
        print("\n--- ОШИБКА ---")
        print("Не удалось выполнить AppleScript.")
        print(f"Детали ошибки: {e.stderr.decode('utf-8')}")
        sys.exit(1)
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Использование: python {sys.argv[0]} <путь_к_файлу.pptx>")
        sys.exit(1)
    pptx_file_path = sys.argv[1]
    convert_pptx_with_keynote(pptx_file_path)