"""Quick-start guide
===================
Install dependencies:
    pip install -r requirements.txt

Usage:
    python unlock_excel.py /path/to/file.xlsx

Example output:
    $ python unlock_excel.py sample.xlsx
    Processing: 100%|##########| 5/5 [00:00<00:00, ?step/s]
    File unlocked: sample_unlocked.xlsx
"""

from __future__ import annotations

import argparse
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

import msoffcrypto
from msoffcrypto.exceptions import FileFormatError
import openpyxl
from openpyxl.workbook.workbook import Workbook
from tqdm import tqdm


def decrypt_file(input_path: Path, temp_dir: Path) -> Path:
    """Remove file-open password using msoffcrypto."""
    output = temp_dir / input_path.name
    with open(input_path, "rb") as f:
        try:
            office_file = msoffcrypto.OfficeFile(f)
        except FileFormatError:
            return input_path
        if not office_file.is_encrypted():
            return input_path
        office_file.load_key(password="")
        with open(output, "wb") as decrypted:
            office_file.decrypt(decrypted)
    return output


def unlock_workbook(workbook: Workbook) -> None:
    """Remove workbook structure protection completely."""
    workbook.security = None


def unlock_worksheets(workbook: Workbook) -> None:
    """Remove worksheet protection and unhide sheets."""
    for sheet in workbook.worksheets:
        sheet.protection = None
        sheet.sheet_state = "visible"


def unlock_vba(zip_path: Path, temp_dir: Path) -> None:
    """Remove VBA project password if present."""
    tmp_zip = temp_dir / "patched.zip"
    with zipfile.ZipFile(zip_path, "r") as zin, zipfile.ZipFile(tmp_zip, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.lower() == "xl/vbaproject.bin":
                data = data.replace(b"DPB=", b"DPx=")
            zout.writestr(item, data)
    shutil.move(tmp_zip, zip_path)


def save_workbook(wb: Workbook, original: Path) -> Path:
    """Save workbook with *_unlocked suffix."""
    new_path = original.with_name(original.stem + "_unlocked" + original.suffix)
    wb.save(new_path)
    return new_path


def process_xlsx(path: Path) -> Path:
    """Process xlsx, xlsm, xltx, xltm files."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        with tqdm(total=5, desc="Processing", unit="step") as bar:
            decrypted = decrypt_file(path, tmp_dir)
            bar.update(1)
            wb = openpyxl.load_workbook(decrypted, keep_vba=True)
            bar.update(1)
            unlock_workbook(wb)
            unlock_worksheets(wb)
            bar.update(1)
            if wb.vba_archive:
                unlock_vba(decrypted, tmp_dir)
            bar.update(1)
            new_path = save_workbook(wb, path)
            bar.update(1)
            return new_path


def process_ods(path: Path) -> Path:
    """Remove protection from ODS file."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        out = tmp_dir / path.name
        with tqdm(total=3, desc="Processing", unit="step") as bar:
            with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(out, "w") as zout:
                bar.update(1)
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "content.xml":
                        tree = ET.fromstring(data)
                        for elem in tree.iter():
                            elem.attrib.pop("protected", None)
                            elem.attrib.pop("protection-key", None)
                        data = ET.tostring(tree)
                    zout.writestr(item, data)
                bar.update(1)
            new_path = path.with_name(path.stem + "_unlocked" + path.suffix)
            shutil.move(out, new_path)
            bar.update(1)
            return new_path


def process_file(path: Path) -> Optional[Path]:
    """Dispatch processing based on file extension."""
    ext = path.suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return process_xlsx(path)
    if ext == ".ods":
        return process_ods(path)
    if ext in {".xls", ".xlsb"}:
        raise NotImplementedError("Binary Excel formats not supported")
    raise ValueError(f"Unsupported file type: {ext}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Unlock protected Excel/ODS files")
    parser.add_argument("path", type=Path, help="Path to spreadsheet")
    args = parser.parse_args()

    try:
        result = process_file(args.path)
        if result:
            print(f"File unlocked: {result}")
    except Exception:
        print("Failed to unlock file")
        raise


if __name__ == "__main__":
    main()
