import os
import time
import argparse
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

WD_FORMAT_PDF = 17
CALL_REJECTED = -2147418111  # "Call was rejected by callee."


def _retry(func, retries=12, delay=0.5):
    last_exc = None
    for _ in range(retries):
        try:
            return func()
        except pywintypes.com_error as e:
            last_exc = e
            if e.args and e.args[0] == CALL_REJECTED:
                time.sleep(delay)
                continue
            raise
    raise RuntimeError(
        "Word kept rejecting automation calls. "
        "Close all Word dialogs (Protected View, recovery, sign-in prompts) and try again."
    ) from last_exc


def convert_folder_to_pdf(input_dir, output_dir, recursive=False, skip_existing=True):
    input_path = Path(input_dir).resolve()
    output_path = Path(output_dir).resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise ValueError(f"Input directory does not exist: {input_path}")

    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.DisplayAlerts = 0

    try:
        files = input_path.rglob("*") if recursive else input_path.glob("*")

        for file in files:
            if file.suffix.lower() not in [".docx", ".doc"]:
                continue

            pdf_file = output_path / (file.stem + ".pdf")

            if skip_existing and pdf_file.exists():
                print(f"Skipping (exists): {pdf_file.name}")
                continue

            print(f"Converting: {file.name}")

            doc = _retry(lambda: word.Documents.Open(str(file), ReadOnly=True))
            _retry(lambda: doc.SaveAs(str(pdf_file), FileFormat=WD_FORMAT_PDF))
            _retry(lambda: doc.Close(False))

        print("✅ Conversion complete.")

    finally:
        try:
            _retry(lambda: word.Quit())
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(description="Batch convert Word files to PDF.")
    parser.add_argument("input_dir", help="Directory containing Word files")
    parser.add_argument("output_dir", help="Directory to save PDFs")
    parser.add_argument("-r", "--recursive", action="store_true", help="Scan subfolders")
    parser.add_argument("--no-skip", action="store_true", help="Do not skip existing PDFs")

    args = parser.parse_args()

    convert_folder_to_pdf(
        args.input_dir,
        args.output_dir,
        recursive=args.recursive,
        skip_existing=not args.no_skip,
    )


if __name__ == "__main__":
    main()
