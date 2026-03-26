import os
import sys
import subprocess
import argparse
import shutil
from pathlib import Path


def run_cmd(cmd, timeout=120, check=True):
    return subprocess.run(
        cmd,
        check=check,
        capture_output=True,
        text=True,
        timeout=timeout
    )


def ensure_tool(name: str) -> bool:
    return shutil.which(name) is not None


def export_docx_to_pdf_with_word(docx_file: Path, temp_pdf: Path):
    applescript = '''
    on run argv
        set docPath to item 1 of argv
        set outPath to item 2 of argv
        tell application "Microsoft Word"
            activate
            delay 1
            try
                set theDoc to open (POSIX file docPath)
                delay 1
                save as theDoc file name outPath file format format PDF
                close theDoc saving no
                return "success"
            on error errMsg number errNum
                try
                    if exists theDoc then close theDoc saving no
                end try
                return "error: " & errMsg & " (" & errNum & ")"
            end try
        end tell
    end run
    '''

    result = run_cmd(
        ['osascript', '-e', applescript, str(docx_file), str(temp_pdf)],
        timeout=180
    )
    if "error:" in result.stdout:
        raise RuntimeError(f"Word error: {result.stdout.strip()}")

    if not temp_pdf.exists():
        raise RuntimeError(f"PDF was not created: {temp_pdf}")


def convert_pdf_all_pages_pdftoppm(temp_pdf: Path, save_dir: Path, base_name: str, dpi: int, output_format: str):
    """
    pdftoppm writes one file per page:
      base_name-1.png, base_name-2.png, ...
    """
    if output_format == "png":
        cmd = [
            "pdftoppm",
            "-r", str(dpi),
            "-png",
            str(temp_pdf),
            str(save_dir / base_name)
        ]
    else:
        # jpeg
        cmd = [
            "pdftoppm",
            "-r", str(dpi),
            "-jpeg",
            str(temp_pdf),
            str(save_dir / base_name)
        ]

    run_cmd(cmd, timeout=300)

    ext = "png" if output_format == "png" else "jpg"
    files = sorted(save_dir.glob(f"{base_name}-*.{ext}"))
    if not files:
        raise RuntimeError("pdftoppm completed but no image files were produced")
    return files


def convert_pdf_all_pages_magick(temp_pdf: Path, save_dir: Path, base_name: str, dpi: int, output_format: str):
    """
    ImageMagick fallback.
    Produces:
      base_name_page_000.png, base_name_page_001.png, ...
    """
    ext = "png" if output_format == "png" else "jpg"
    pattern = str(save_dir / f"{base_name}_page_%03d.{ext}")

    # -density must be before input PDF
    cmd = ["magick", "-density", str(dpi), str(temp_pdf)]

    # JPEG quality only for jpeg
    if output_format == "jpeg":
        cmd += ["-quality", "95"]

    # white background for pages with transparency
    cmd += ["-background", "white", "-alpha", "remove", "-alpha", "off", pattern]

    run_cmd(cmd, timeout=300)

    files = sorted(save_dir.glob(f"{base_name}_page_*.{ext}"))
    if not files:
        raise RuntimeError("ImageMagick completed but no image files were produced")
    return files


def convert_pdf_first_page_qlmanage(temp_pdf: Path, save_dir: Path, base_name: str, dpi: int, output_format: str):
    """
    Last-resort fallback: QuickLook thumbnail/preview (usually first page only).
    """
    preview_size = int(dpi * 11)  # rough scale
    run_cmd(
        ["qlmanage", "-t", "-s", str(preview_size), "-o", str(save_dir), str(temp_pdf)],
        timeout=120
    )

    generated_file = save_dir / f"{temp_pdf.name}.png"
    # qlmanage naming differs across macOS versions; try both variants
    if not generated_file.exists():
        generated_file = save_dir / f"{temp_pdf.stem}.png"

    if not generated_file.exists():
        raise RuntimeError("qlmanage did not produce a preview image")

    if output_format == "png":
        out = save_dir / f"{base_name}_page_001.png"
        generated_file.rename(out)
        return [out]

    out = save_dir / f"{base_name}_page_001.jpeg"
    run_cmd(["sips", "-s", "format", "jpeg", str(generated_file), "--out", str(out)], timeout=120)
    generated_file.unlink(missing_ok=True)
    return [out]


def convert_docx_first_page_qlmanage(docx_file: Path, save_dir: Path, base_name: str, dpi: int, output_format: str):
    """
    Fallback when Word automation fails: render first page from DOCX via QuickLook.
    """
    preview_size = int(dpi * 11)  # rough scale
    run_cmd(
        ["qlmanage", "-t", "-s", str(preview_size), "-o", str(save_dir), str(docx_file)],
        timeout=120
    )

    generated_file = save_dir / f"{docx_file.name}.png"
    if not generated_file.exists():
        generated_file = save_dir / f"{docx_file.stem}.png"

    if not generated_file.exists():
        raise RuntimeError("qlmanage did not produce a DOCX preview image")

    if output_format == "png":
        out = save_dir / f"{base_name}_page_001.png"
        generated_file.rename(out)
        return [out]

    out = save_dir / f"{base_name}_page_001.jpeg"
    run_cmd(["sips", "-s", "format", "jpeg", str(generated_file), "--out", str(out)], timeout=120)
    generated_file.unlink(missing_ok=True)
    return [out]


def normalize_filenames(files, save_dir: Path, base_name: str, output_format: str):
    """
    Normalize to:
      base_name_page_001.png
      base_name_page_002.png
      ...
    """
    final_ext = "png" if output_format == "png" else "jpeg"
    normalized = []

    for idx, src in enumerate(sorted(files), start=1):
        dst = save_dir / f"{base_name}_page_{idx:03d}.{final_ext}"

        if src.suffix.lower() == ".jpg" and final_ext == "jpeg":
            src.rename(dst)
        elif src.suffix.lower() == ".jpeg" and final_ext == "jpeg":
            src.rename(dst)
        elif src.suffix.lower() == ".png" and final_ext == "png":
            src.rename(dst)
        else:
            # format mismatch or extension mismatch -> convert via sips
            run_cmd(["sips", "-s", "format", final_ext, str(src), "--out", str(dst)], timeout=120)
            src.unlink(missing_ok=True)

        normalized.append(dst)

    return normalized


def convert_docx_with_word(docx_path, dpi=300, output_format="png", output_dir=None, output_name=None, keep_pdf=False):
    if sys.platform != "darwin":
        print("Ошибка: Этот скрипт предназначен только для macOS и требует Microsoft Word.")
        sys.exit(1)

    docx_file = Path(docx_path).resolve()
    if not docx_file.is_file():
        print(f"Ошибка: Файл не найден по пути: {docx_path}")
        sys.exit(1)

    save_dir = Path(output_dir).resolve() if output_dir else Path(__file__).parent.resolve()
    save_dir.mkdir(parents=True, exist_ok=True)

    base_name = output_name if output_name else docx_file.stem
    temp_pdf = save_dir / f"{base_name}_temp.pdf"

    print(f"Используем Microsoft Word для конвертации '{docx_file.name}' в PDF...")
    print(f"Сохранение в: {save_dir}")

    try:
        produced_files = None
        word_error = None

        try:
            export_docx_to_pdf_with_word(docx_file, temp_pdf)
            print("✓ Конвертация в PDF завершена.")
        except Exception as e:
            word_error = e
            print(f"Word не сработал: {e}")

        if temp_pdf.exists():
            # 1) Best path for multi-page: pdftoppm (Poppler)
            if ensure_tool("pdftoppm"):
                print(f"Конвертирую ВСЕ страницы через pdftoppm в {output_format.upper()} (DPI={dpi})...")
                try:
                    produced_files = convert_pdf_all_pages_pdftoppm(
                        temp_pdf, save_dir, base_name, dpi, output_format
                    )
                except Exception as e:
                    print(f"pdftoppm не сработал: {e}")

            # 2) Fallback: ImageMagick
            if produced_files is None and ensure_tool("magick"):
                print(f"Конвертирую ВСЕ страницы через ImageMagick в {output_format.upper()} (DPI={dpi})...")
                try:
                    produced_files = convert_pdf_all_pages_magick(
                        temp_pdf, save_dir, base_name, dpi, output_format
                    )
                except Exception as e:
                    print(f"ImageMagick не сработал: {e}")

            # 3) Fallback: qlmanage (first page only)
            if produced_files is None:
                print("⚠️  Не найден pdftoppm/magick. Пробую qlmanage (только первая страница).")
                produced_files = convert_pdf_first_page_qlmanage(
                    temp_pdf, save_dir, base_name, dpi, output_format
                )

        if produced_files is None:
            print("⚠️  Переключаюсь на qlmanage по DOCX (только первая страница).")
            produced_files = convert_docx_first_page_qlmanage(
                docx_file, save_dir, base_name, dpi, output_format
            )
            if word_error is not None:
                print("Примечание: Word-конвертация не удалась, используется запасной путь без Word.")

        final_files = normalize_filenames(produced_files, save_dir, base_name, output_format)

        if not keep_pdf and temp_pdf.exists():
            temp_pdf.unlink(missing_ok=True)

        print("\n✅ Обработка завершена!")
        print(f"Создано файлов: {len(final_files)}")
        for f in final_files:
            print(f"  - {f.name}")

        if len(final_files) == 1:
            print("\nПримечание: получена только 1 страница.")
            print("Для многостраничного результата установите Poppler или ImageMagick:")
            print("  brew install poppler")
            print("или")
            print("  brew install imagemagick")

    except subprocess.CalledProcessError as e:
        print("\n--- ОШИБКА ---")
        print("Не удалось выполнить команду.")
        if e.stderr:
            print(e.stderr.strip())
        sys.exit(1)
    except subprocess.TimeoutExpired:
        print("\n--- ОШИБКА ---")
        print("Превышено время ожидания при конвертации документа.")
        sys.exit(1)
    except Exception as e:
        print("\n--- ОШИБКА ---")
        print(str(e))
        sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Конвертирует .docx в изображения (все страницы) через Microsoft Word + PDF rasterizer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python %(prog)s document.docx
  python %(prog)s document.docx --dpi 600 --format png
  python %(prog)s document.docx --dir ./results --name title
  python %(prog)s document.docx --keep-pdf

Результат:
  title_page_001.png
  title_page_002.png
  ...
        """
    )

    parser.add_argument("docx_file", help="Путь к .docx файлу")
    parser.add_argument("--dpi", type=int, default=300, help="DPI (по умолчанию: 300)")
    parser.add_argument("--format", choices=["png", "jpeg"], default="png", help="Формат изображения")
    parser.add_argument("--dir", "-d", dest="output_dir", help="Папка для сохранения")
    parser.add_argument("--name", "-n", dest="output_name", help="Базовое имя выходных файлов")
    parser.add_argument("--keep-pdf", action="store_true", help="Не удалять временный PDF")

    args = parser.parse_args()

    convert_docx_with_word(
        args.docx_file,
        dpi=args.dpi,
        output_format=args.format,
        output_dir=args.output_dir,
        output_name=args.output_name,
        keep_pdf=args.keep_pdf,
    )
