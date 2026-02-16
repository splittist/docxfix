"""Export .docx files to PDF via Word COM automation (Windows-only).

Usage:
    python scripts/export_pdf.py [PATHS...]

PATHS can be individual .docx files or directories (searched non-recursively).
Defaults to scratch_out/*.docx if no arguments given.
PDFs are written to a pdf/ subdirectory next to each input file.

Requires: pywin32 (pip install pywin32) and Microsoft Word installed.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

WD_EXPORT_FORMAT_PDF = 17


def find_docx_files(args: list[str]) -> list[Path]:
    if not args:
        args = ["scratch_out"]

    files: list[Path] = []
    for arg in args:
        p = Path(arg)
        if p.is_file() and p.suffix.lower() == ".docx":
            files.append(p)
        elif p.is_dir():
            files.extend(sorted(p.glob("*.docx")))
        else:
            print(f"warning: skipping {arg} (not a .docx file or directory)")
    return files


def export_to_pdf(docx_paths: list[Path]) -> tuple[list[Path], list[tuple[Path, str]]]:
    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError:
        print("error: pywin32 is not installed. Run: uv pip install -e '.[win]'")
        sys.exit(1)

    word = None
    successes: list[Path] = []
    failures: list[tuple[Path, str]] = []

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        for docx_path in docx_paths:
            docx_path = docx_path.resolve()
            pdf_dir = docx_path.parent / "pdf"
            pdf_dir.mkdir(exist_ok=True)
            pdf_path = pdf_dir / docx_path.with_suffix(".pdf").name

            print(f"  {docx_path.name} -> pdf/{pdf_path.name} ... ", end="", flush=True)

            doc = None
            try:
                doc = word.Documents.Open(str(docx_path), ReadOnly=True)
                doc.ExportAsFixedFormat(
                    str(pdf_path),
                    WD_EXPORT_FORMAT_PDF,
                    OpenAfterExport=False,
                    OptimizeFor=0,  # wdExportOptimizeForPrint
                )
                successes.append(pdf_path)
                print("ok")
            except Exception as exc:
                failures.append((docx_path, str(exc)))
                print(f"FAILED: {exc}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=0)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
            # Give Word a moment to shut down
            time.sleep(0.5)

    return successes, failures


def main() -> None:
    docx_files = find_docx_files(sys.argv[1:])
    if not docx_files:
        print("No .docx files found.")
        sys.exit(1)

    print(f"Exporting {len(docx_files)} file(s) to PDF via Word COM...\n")
    successes, failures = export_to_pdf(docx_files)

    print(f"\nDone: {len(successes)} succeeded, {len(failures)} failed.")
    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()
