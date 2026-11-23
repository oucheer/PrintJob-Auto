from pathlib import Path

from printcase.excel_reader import read_cases
from printcase.models import CaseStatus, PrintCase
from printcase.validator import validate_case


def test_page_selection_parsing():
    from printcase.direct_print import parse_page_selection, page_order

    assert parse_page_selection("", 10) == list(range(10))
    assert parse_page_selection("1-3,5", 10) == [0, 1, 2, 4]
    assert parse_page_selection("last", 10) == [9]
    assert parse_page_selection("8-last", 10) == [7, 8, 9]
    assert parse_page_selection("odd", 6) == [0, 2, 4]
    assert parse_page_selection("even", 6) == [1, 3, 5]
    assert parse_page_selection("99", 10) == []

    # 逐份：整份重复；非逐份：每页连打
    assert page_order([0, 1, 2], 2, "yes") == [0, 1, 2, 0, 1, 2]
    assert page_order([0, 1, 2], 2, "no") == [0, 0, 1, 1, 2, 2]
    assert page_order([0, 1], 1, "no") == [0, 1]


def test_all_explicit_paper_sizes_resolve_to_driver_order():
    """所有显式纸张尺寸都必须能在驱动界面上定位到具体选项。"""
    from printcase.driver_ui import plan_actions

    caps = {"papers": ["A4 (210 x 297mm)"], "paper_ids": [9]}
    case = PrintCase(case_id="TC-01", enabled=True, file_path="a.pdf", paper="A4", print_quality="600", collate="yes")
    actions = {item.label: item for item in plan_actions(case, caps)}
    assert actions["纸张尺寸"].value == 0
    assert actions["纸张尺寸"].expect_value == 9      # 驱动自报的 DMPAPER 代号，用于回读校验
    assert actions["打印质量"].expect_value == 600
    assert actions["逐份"].expect_value == 1


def test_excel_reader_supports_chinese_headers_and_relative_paths(tmp_path):
    from openpyxl import Workbook

    doc = tmp_path / "sample.pdf"
    doc.write_bytes(b"%PDF-sample")
    xlsx = tmp_path / "cases.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "用例"
    ws.append(["用例编号", "启用", "文件路径", "打印机", "页码", "份数", "纸张", "方向", "颜色", "单双面", "缩放", "纸盒", "超时秒", "备注"])
    ws.append(["TC-001", "是", "sample.pdf", "Office Printer", "1-2", 2, "A4", "landscape", "monochrome", "duplexlong", "fit", "2", 60, "smoke test"])
    wb.save(xlsx)

    cases, warnings = read_cases(xlsx)
    assert warnings == []
    assert len(cases) == 1
    assert cases[0].case_id == "TC-001"
    assert cases[0].resolved_path == doc
    assert cases[0].copies == 2
    assert cases[0].duplex == "duplexlong"


def test_validation_rejects_unknown_printer_and_bad_page_range(tmp_path):
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-ERR", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Missing", pages="1--2")
    errors = validate_case(case, {"Known Printer"})
    assert any("打印机" in e for e in errors)
    assert any("页码格式" in e for e in errors)


def test_validation_accepts_valid_case(tmp_path):
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-OK", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Known Printer")
    assert validate_case(case, {"Known Printer"}) == []


def test_excel_reader_flags_invalid_boolean_and_integer_values(tmp_path):
    from openpyxl import Workbook

    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    xlsx = tmp_path / "bad-values.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "用例"
    ws.append(["case_id", "enabled", "file_path", "copies", "timeout_seconds"])
    ws.append(["TC-BAD", "maybe", "sample.pdf", "many", 90])
    wb.save(xlsx)
    cases, _ = read_cases(xlsx)
    assert len(cases) == 1
    assert len(cases[0].input_errors) == 2
    errors = validate_case(cases[0], {"Known Printer"}, "Known Printer")
    assert any("启用字段值" in e for e in errors)
    assert any("份数不是有效整数" in e for e in errors)
