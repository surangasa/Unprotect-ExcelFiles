import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl
from unlock_excel import unlock_workbook, unlock_worksheets


def test_unlock_workbook():
    wb = openpyxl.Workbook()
    wb.security.lockStructure = True
    unlock_workbook(wb)
    assert wb.security is None


def test_unlock_worksheets():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.protection.sheet = True
    ws.sheet_state = 'hidden'
    unlock_worksheets(wb)
    assert ws.sheet_state == 'visible'
    assert ws.protection is None
