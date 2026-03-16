import os
import shutil
import subprocess
from pathlib import Path


SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def find_soffice():
    found = shutil.which("soffice")
    if found:
        return found

    for candidate in SOFFICE_CANDIDATES:
        if os.path.exists(candidate):
            return candidate

    return None


def needs_conversion(path, target_extension):
    return Path(path).suffix.lower() != target_extension.lower()


def convert_with_libreoffice(src, output_dir, target_extension):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice was not found. Install LibreOffice or make sure soffice.exe is available in PATH."
        )

    target_filter = "docx" if target_extension.lower() == ".docx" else "odt"
    command = [
        soffice,
        "--headless",
        "--convert-to",
        target_filter,
        "--outdir",
        str(output_dir),
        str(src),
    ]
    completed = subprocess.run(command, capture_output=True, text=True, check=False)
    if completed.returncode != 0:
        stderr = completed.stderr.strip()
        stdout = completed.stdout.strip()
        details = stderr or stdout or "Unknown conversion error"
        raise RuntimeError(f"LibreOffice conversion failed: {details}")

    converted_path = Path(output_dir) / f"{Path(src).stem}{target_extension}"
    if not converted_path.exists():
        raise RuntimeError(f"Converted file was not created: {converted_path}")

    return str(converted_path)
