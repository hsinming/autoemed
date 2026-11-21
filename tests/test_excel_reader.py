from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def _build_xlsx(path, rows, header_col=1):
    """rows: list of (value, style); style is None, "fill", "red_font" or "theme_font".

    Plain cells get an explicit black RGB font: the reader rejects the theme default (see xfail below).
    """
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="Name")
    ws.cell(row=1, column=header_col, value="eMedical No.")
    for i, (value, style) in enumerate(rows, start=2):
        cell = ws.cell(row=i, column=header_col, value=value)
        if style != "theme_font":
            cell.font = Font(color="FF000000")
        if style == "fill":
            cell.fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")
        elif style == "red_font":
            cell.font = Font(color="FFFF0000")
    wb.save(path)
    return str(path)


def test_plain_rows_are_returned_as_strings(main_module, tmp_path):
    xlsx = _build_xlsx(tmp_path / "cases.xlsx", [("HAP 47994319", None), (12345678, None)])

    assert main_module.ExcelProcessor().extract_emedical_no(xlsx) == ["HAP 47994319", "12345678"]


def test_rows_with_fill_colour_are_skipped(main_module, tmp_path):
    xlsx = _build_xlsx(tmp_path / "cases.xlsx", [("HAP 1", None), ("HAP 2", "fill"), ("HAP 3", None)])

    assert main_module.ExcelProcessor().extract_emedical_no(xlsx) == ["HAP 1", "HAP 3"]


def test_rows_with_non_black_font_are_skipped(main_module, tmp_path):
    xlsx = _build_xlsx(tmp_path / "cases.xlsx", [("HAP 1", "red_font"), ("HAP 2", None)])

    assert main_module.ExcelProcessor().extract_emedical_no(xlsx) == ["HAP 2"]


def test_column_is_found_by_header_not_position(main_module, tmp_path):
    xlsx = _build_xlsx(tmp_path / "cases.xlsx", [("CEAC 1", None)], header_col=4)

    assert main_module.ExcelProcessor().extract_emedical_no(xlsx) == ["CEAC 1"]


@pytest.mark.xfail(strict=True, reason="bug: theme font colour (Excel default, theme=1 Text 1) is treated as non-black")
def test_default_theme_font_counts_as_black(main_module, tmp_path):
    xlsx = _build_xlsx(tmp_path / "cases.xlsx", [("HAP 1", "theme_font")])

    assert main_module.ExcelProcessor().extract_emedical_no(xlsx) == ["HAP 1"]


def test_runs_in_temporary_working_directory(main_module, tmp_path):
    assert Path.cwd() == tmp_path
