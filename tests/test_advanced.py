from pathlib import Path

from openpyxl import Workbook

from printcase.excel_reader import read_cases
from printcase.models import CaseStatus, PrintCase
from printcase.office import DocumentConversionError, convert_document_to_pdf
from printcase.audit import AuditStore
from printcase.runner import PrintRunner
from printcase.spooler import QueueJob
from printcase.validator import validate_case

from test_runner import FakeSpooler, fake_driver_applier


def _make_case(tmp_path, **kwargs) -> PrintCase:
    defaults = dict(case_id="TC-ADV", enabled=True, file_path="a.pdf", resolved_path=tmp_path / "a.pdf", printer="Test Printer")
    defaults.update(kwargs)
    return PrintCase(**defaults)


def test_mixed_paper_omits_paper_token_for_mixed_size_printing(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, paper="mixed", resolved_path=pdf)
    assert validate_case(case, {"Test Printer"}, "Test Printer") == []
    # 混合尺寸：驱动界面里不指定纸张尺寸，由驱动按每页尺寸走纸
    from printcase.driver_ui import plan_actions
    assert not [a for a in plan_actions(case) if a.label == "纸张尺寸"]


def test_paper_auto_and_mixed_are_accepted(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    for value in ("auto", "mixed"):
        case = _make_case(tmp_path, paper=value)
        errors = [e for e in validate_case(case, {"Test Printer"}) if "纸张" in e]
        assert errors == [], (value, errors)


def test_advanced_print_attributes_validate(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(
        tmp_path,
        borderless="yes", separator_page="job", bin="纸盒1",
        binding_location="left", binding_method="staple", duplex="simplex",
    )
    assert validate_case(case, {"Test Printer"}) == []

    bad_borderless = _make_case(tmp_path, borderless="开")
    assert any("无边距" in e for e in validate_case(bad_borderless, set()))
    bad_separator = _make_case(tmp_path, separator_page="中间")
    assert any("分隔页" in e for e in validate_case(bad_separator, set()))
    bad_binding = _make_case(tmp_path, binding_location="中缝")
    assert any("装订位置" in e for e in validate_case(bad_binding, set()))
    bad_method = _make_case(tmp_path, binding_method="打钉")
    assert any("装订方式" in e for e in validate_case(bad_method, set()))


def test_binding_edge_requires_duplex(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, binding_location="longedge", duplex="simplex")
    assert any("双面" in e for e in validate_case(case, set()))
    case.duplex = "duplexlong"
    assert not any("双面" in e for e in validate_case(case, set()))


def test_docx_and_doc_are_accepted_as_print_sources(tmp_path):
    doc = tmp_path / "a.docx"
    doc.write_bytes(b"PK\x03\x04 placeholder")
    case = _make_case(tmp_path, file_path=str(doc), resolved_path=doc)
    assert validate_case(case, {"Test Printer"}, "Test Printer") == []
    doc2 = tmp_path / "b.doc"
    doc2.write_bytes(b"\xd0\xcf\x11\xe0 legacy")
    case2 = _make_case(tmp_path, file_path=str(doc2), resolved_path=doc2)
    assert validate_case(case2, {"Test Printer"}, "Test Printer") == []


def test_chinese_excel_values_map_to_canonical_settings(tmp_path):
    xlsx = tmp_path / "cases.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "用例"
    ws.append(["用例编号", "是否启用", "文件路径", "打印机", "页码", "份数", "纸张", "方向",
               "颜色", "单双面", "缩放", "纸盒", "无边距", "分隔页", "装订位置", "装订方式",
               "超时秒", "备注"])
    ws.append(["TC-CN", "是", "sample.docx", "", "1-3", 2, "混合", "横向",
               "黑白", "双面短边", "不缩放", "2", "是", "作业之间", "长边", "骑马钉", 60, "中文模板回环"])
    doc = tmp_path / "sample.docx"
    doc.write_bytes(b"PK")
    wb.save(xlsx)

    cases, warnings = read_cases(xlsx)
    assert warnings == []
    case = cases[0]
    assert case.paper == "mixed"
    assert case.orientation == "landscape"
    assert case.color == "monochrome"
    assert case.duplex == "duplexshort"
    assert case.scaling == "noscale"
    assert case.bin == "2"
    assert case.borderless == "yes"
    assert case.separator_page == "job"
    assert case.binding_location == "longedge"
    assert case.binding_method == "saddle"
    assert case.settings_summary().count("装订") == 1


def test_repository_template_is_fully_chinese_and_parses():
    template = Path(__file__).resolve().parents[1] / "templates" / "PrintCase_Template.xlsx"
    cases, warnings = read_cases(template)
    assert warnings == []
    # 模板会被用户填入真实用例，用例数量不固定；只要有可解析的中文用例即可。
    assert len(cases) >= 1
    assert all(case.case_id.startswith("TC-") for case in cases)
    # 启用且填写完整的用例必须能通过校验；禁用/无文件的用例（如打印机面板内部页）不参与执行。
    for case in cases:
        if case.enabled and case.file_path:
            assert validate_case(case, {"Test Printer"}, "Test Printer") == [], case.case_id


class JobThenClearSpooler(FakeSpooler):
    """第一次查询为空（基线），第二次出现作业，之后从队列消失。"""
    def __init__(self, document="sample.pdf"):
        super().__init__()
        self.calls = 0
        self.document = document

    def get_printer_status(self, name):
        return 0, "就绪"

    def get_jobs(self, name):
        self.calls += 1
        if self.calls == 1:
            return []
        if self.calls == 2:
            self.jobs = [QueueJob(303, self.document, 0, "无状态标记", 0, 1)]
            return list(self.jobs)
        self.jobs = []
        return []


def test_runner_prints_docx_through_ui_automation_without_conversion(tmp_path):
    """docx 属于 COM 直打格式：直接把源文档交给打印端，不做 PDF 转换。"""
    docx = tmp_path / "报告.docx"
    docx.write_bytes(b"PK placeholder")

    def converter(*_a, **_k):
        raise AssertionError("docx 不应触发 PDF 转换")

    submitted = []

    def submit(printer, case, source_path, log, temp_dir):
        submitted.append((printer, Path(source_path)))
        return "office-com", Path(source_path).stem

    case = PrintCase(case_id="TC-DOCX", enabled=True, file_path=str(docx), resolved_path=docx, printer="Test Printer")
    runner = PrintRunner(
        JobThenClearSpooler("报告.pdf"), AuditStore(tmp_path / "a1.sqlite3"),
        queue_timeout=1, ui_submitter=submit, driver_applier=fake_driver_applier(),
        document_converter=converter,
    )
    result = runner.run_cases([case], "Test Printer", "R1", lambda *_: None)[0]
    assert result.status == CaseStatus.QUEUE_CLEARED
    assert len(submitted) == 1
    assert submitted[0][0] == "Test Printer"          # 下发到用例/界面选择的驱动
    assert submitted[0][1].suffix == ".docx"          # 直接打开原文档
    assert result.settings["backend"] == "office-com"


def test_runner_reports_failure_when_conversion_is_unavailable(tmp_path):
    from printcase.office import DocumentConversionError

    pptx = tmp_path / "bad.pptx"
    pptx.write_bytes(b"PK")

    def converter(doc_path, **kwargs):
        raise DocumentConversionError("未找到 LibreOffice")

    # 不注入 submitter：走默认路由，转换失败应在下发前被拦截
    case = PrintCase(case_id="TC-NOCONV", enabled=True, file_path=str(pptx), resolved_path=pptx, printer="Test Printer")
    runner = PrintRunner(
        JobThenClearSpooler(), AuditStore(tmp_path / "a2.sqlite3"),
        queue_timeout=1, driver_applier=fake_driver_applier(),
        document_converter=converter,
    )
    result = runner.run_cases([case], "Test Printer", "R2", lambda *_: None)[0]
    assert result.status == CaseStatus.FAILED
    assert "转 PDF" in result.message


def test_runner_uses_driver_separator_without_extra_jobs(tmp_path):
    """分隔页不再是“软件另打一张 PDF 充数”：只下发一个作业，设置交给驱动。"""
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"%PDF")
    launched = []

    def submit(printer, case, source_path, log, temp_dir):
        launched.append((case.case_id, Path(source_path).name))
        return "gdi-pdf", Path(source_path).stem

    case = PrintCase(
        case_id="TC-SEP", enabled=True, file_path=str(pdf), resolved_path=pdf,
        printer="Test Printer", separator_page="job", bin="纸盒1",
    )
    runner = PrintRunner(JobThenClearSpooler(), AuditStore(tmp_path / "a3.sqlite3"), queue_timeout=1, submitter=submit, driver_applier=fake_driver_applier())
    result = runner.run_cases([case], "Test Printer", "R3", lambda *_: None)[0]
    assert result.status == CaseStatus.QUEUE_CLEARED
    assert [case_id for case_id, _name in launched] == ["TC-SEP"]   # 只有主作业，没有额外分隔页作业


class FakeWordDoc:
    def ExportAsFixedFormat(self, path, fmt):
        Path(path).write_bytes(b"%PDF-fake-word")
    def Close(self, save_changes):
        pass


class FakeWord:
    def __init__(self):
        self.Documents = self
        self.quit_called = False
    def Open(self, path, **kwargs):
        return FakeWordDoc()
    def Quit(self):
        self.quit_called = True


def test_office_conversion_with_fake_word(tmp_path):
    doc = tmp_path / "memo.docx"
    doc.write_bytes(b"PK")
    out = convert_document_to_pdf(doc, work_dir=tmp_path / "out", word_factory=FakeWord, prefer="word")
    assert out.name == "memo.pdf"
    assert out.read_bytes().startswith(b"%PDF")


def test_office_conversion_reports_missing_backends(tmp_path):
    doc = tmp_path / "memo.docx"
    doc.write_bytes(b"PK")
    try:
        convert_document_to_pdf(
            doc, work_dir=tmp_path / "out", prefer="soffice",
            soffice_path=str(tmp_path / "no_soffice.exe"),
        )
    except DocumentConversionError as exc:
        assert "soffice" in str(exc)
    else:
        raise AssertionError("应抛出 DocumentConversionError")


# ---------- Pantum CM806ADN 驱动属性对照补充 ----------

def test_collate_quality_media_skipblank_validate(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    ok = _make_case(
        tmp_path, resolved_path=pdf,
        collate="yes", print_quality="1200",
        media_type="普通纸(60-90gsm)", skip_blank="yes", copies=9999,
    )
    assert validate_case(ok, {"Test Printer"}, "Test Printer") == []

    bad_quality = _make_case(tmp_path, print_quality="高清")
    assert any("打印质量" in e for e in validate_case(bad_quality, set()))
    bad_collate = _make_case(tmp_path, collate="乱序")
    assert any("打印顺序" in e for e in validate_case(bad_collate, set()))
    bad_media = _make_case(tmp_path, media_type="普通纸,厚纸")
    assert any("纸张类型" in e for e in validate_case(bad_media, set()))
    bad_skip = _make_case(tmp_path, skip_blank="自动")
    assert any("跳过空白页" in e for e in validate_case(bad_skip, set()))
    too_many = _make_case(tmp_path, copies=10000)
    assert any("9999" in e for e in validate_case(too_many, set()))


def test_mixed_sizes_switch_rules(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    # 开启混合尺寸 + 自动选纸盒：通过
    ok = _make_case(tmp_path, resolved_path=pdf, mixed_sizes="yes", paper="auto", bin="")
    assert validate_case(ok, {"Test Printer"}, "Test Printer") == []
    from printcase.driver_ui import plan_actions
    assert not [a for a in plan_actions(ok) if a.label == "纸张尺寸"]

    # 开启混合尺寸但指定了纸盒：拦截
    bad_bin = _make_case(tmp_path, resolved_path=pdf, mixed_sizes="yes", bin="纸盒2")
    assert any("纸盒只能为“自动选择”" in e for e in validate_case(bad_bin, set()))

    # 开启混合尺寸但指定了具体纸张：拦截
    bad_paper = _make_case(tmp_path, resolved_path=pdf, mixed_sizes="yes", paper="A4 (210 x 297mm)")
    assert any("纸张列请留空" in e for e in validate_case(bad_paper, set()))

    # 取值非法
    bad_value = _make_case(tmp_path, mixed_sizes="开")
    assert any("混合尺寸必须" in e for e in validate_case(bad_value, set()))


def test_separator_copies_mode_accepted(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, resolved_path=pdf, separator_page="copies", copies=5, bin="纸盒1")
    assert validate_case(case, {"Test Printer"}, "Test Printer") == []
    assert "份数之间" in case.settings_summary()


def test_separator_old_values_rejected_and_tray_required(tmp_path):
    """模板里只剩 无/作业之间/份数之间；旧的前/后/前后 直接判错，不再假装支持。"""
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    for old in ("before", "after", "both", "each_copy"):
        case = _make_case(tmp_path, resolved_path=pdf, separator_page=old, bin="纸盒1")
        errors = validate_case(case, {"Test Printer"}, "Test Printer")
        assert any("分隔页必须为" in e for e in errors), old
    # 驱动在“纸张来源=自动选择”时会把“输出/分隔页”整组灰化：必须指定具体纸盒
    need_bin = _make_case(tmp_path, resolved_path=pdf, separator_page="job", bin="")
    assert any("具体纸盒" in e for e in validate_case(need_bin, {"Test Printer"}, "Test Printer"))


def test_chinese_excel_new_columns_map(tmp_path):
    xlsx = tmp_path / "cases.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "用例"
    ws.append(["用例编号", "文件路径", "份数", "纸盒", "打印顺序", "混合尺寸", "打印质量",
               "纸张类型", "跳过空白页", "分隔页", "分隔页纸张来源"])
    ws.append(["TC-PNT", "sample.pdf", 2, "多功能进纸器", "非逐份", "否", "600",
               "普通纸", "是", "份数之间", "多功能进纸器"])
    doc = tmp_path / "sample.pdf"
    doc.write_bytes(b"%PDF")
    wb.save(xlsx)
    cases, _ = read_cases(xlsx)
    case = cases[0]
    assert case.collate == "no"
    assert case.mixed_sizes == "no"
    assert case.print_quality == "600"
    assert case.media_type == "普通纸"
    assert case.skip_blank == "yes"
    assert case.separator_page == "copies"
    assert case.separator_bin == "多功能进纸器"


def test_runner_keeps_copies_in_one_job_for_copies_separator(tmp_path):
    """分隔页=份数之间 由驱动在份与份之间插入，份数仍按一个作业提交。"""
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"%PDF")
    launched = []

    def submit(printer, case, source_path, log, temp_dir):
        launched.append((case.case_id, case.copies, Path(source_path).name))
        return "gdi-pdf", Path(source_path).stem

    case = PrintCase(
        case_id="TC-BETW", enabled=True, file_path=str(pdf), resolved_path=pdf,
        printer="Test Printer", copies=3, separator_page="copies", bin="纸盒1",
    )
    runner = PrintRunner(
        JobThenClearSpooler(), AuditStore(tmp_path / "a4.sqlite3"),
        queue_timeout=1, submitter=submit, driver_applier=fake_driver_applier(),
    )
    result = runner.run_cases([case], "Test Printer", "R4", lambda *_: None)[0]
    assert result.status == CaseStatus.QUEUE_CLEARED
    assert launched == [("TC-BETW", 3, "sample.pdf")]


# ---------- 驱动属性页操作计划（driver_ui.plan_actions） ----------

def _action(actions, label):
    return next((item for item in actions if item.label == label), None)


def test_plan_actions_translation(tmp_path):
    from printcase.driver_ui import plan_actions

    case = _make_case(tmp_path, collate="yes", print_quality="1200", media_type="普通纸(60-90gsm)")
    actions = plan_actions(case)
    assert _action(actions, "逐份").value is True
    assert _action(actions, "逐份").expect_value == 1
    assert _action(actions, "打印质量").value[0] == "1200 dpi"
    assert _action(actions, "打印质量").expect_value == 1200
    # 无驱动能力数据时按名称在下拉里选（驱动自己负责编码）
    assert _action(actions, "纸张类型").kind == "combo_text"
    assert _action(actions, "纸张类型").value == "普通纸(60-90gsm)"

    case2 = _make_case(tmp_path, media_type="铜版纸", print_quality="high")
    actions2 = plan_actions(case2)
    assert _action(actions2, "纸张类型").value == "铜版纸"
    assert _action(actions2, "打印质量").value[0] == "high"


def test_plan_actions_uses_driver_order_for_private_media(tmp_path):
    from printcase.driver_ui import plan_actions

    # 驱动上报的私有介质代号不可靠，但下拉顺序稳定：按序号设置即可
    caps = {"media_names": ["普通纸(60-90gsm)", "厚纸1(106-120gsm)", "厚纸1+(121-157gsm)"]}
    case = _make_case(tmp_path, media_type="厚纸1+(121-157gsm)")
    action = _action(plan_actions(case, caps), "纸张类型")
    assert action.kind == "combo_index"
    assert action.value == 2
    # 私有介质没有可预期的标准代号：只要求回读值发生变化
    assert action.expect_field == "MediaType" and action.expect_value is None


def test_verify_reports_media_unchanged(tmp_path, monkeypatch):
    from printcase import driver_ui

    snapshots = [{"MediaType": 266, "PaperSize": 9}, {"MediaType": 266, "PaperSize": 9}]
    monkeypatch.setattr(driver_ui, "snapshot", lambda printer, w=None: snapshots.pop(0))
    actions = [driver_ui.Action("纸张", "纸张类型", "combo_index", 3, "MediaType")]
    problems = driver_ui.verify("Test Printer", actions, {"MediaType": 266, "PaperSize": 9})
    assert problems and "纸张类型" in problems[0]


def test_verify_accepts_driver_private_media_code(tmp_path, monkeypatch):
    from printcase import driver_ui

    monkeypatch.setattr(driver_ui, "snapshot", lambda printer, w=None: {"MediaType": 263, "PaperSize": 9})
    actions = [driver_ui.Action("纸张", "纸张类型", "combo_index", 3, "MediaType")]
    assert driver_ui.verify("Test Printer", actions, {"MediaType": 266, "PaperSize": 9}) == []


def test_extended_office_formats_are_accepted(tmp_path):
    for name in ("a.rtf", "b.odt", "c.wps", "d.pptx", "e.xls", "f.xlsx"):
        doc = tmp_path / name
        doc.write_bytes(b"stub-content")
        case = _make_case(tmp_path, file_path=str(doc), resolved_path=doc)
        errors = validate_case(case, {"Test Printer"}, "Test Printer")
        assert errors == [], (name, errors)


def test_backend_order_prefers_word_only_for_word_formats():
    from printcase.office import backend_order

    assert backend_order(".docx") == ["word", "wps", "soffice"]
    assert backend_order(".rtf") == ["word", "wps", "soffice"]
    assert backend_order(".odt") == ["soffice", "wps", "word"]
    assert backend_order(".pptx") == ["soffice"]
    assert backend_order(".xls") == ["soffice"]


def test_pptx_conversion_reports_libreoffice_requirement(tmp_path):
    from printcase.office import DocumentConversionError, convert_document_to_pdf

    doc = tmp_path / "slides.pptx"
    doc.write_bytes(b"PK")
    try:
        convert_document_to_pdf(doc, work_dir=tmp_path / "out", soffice_path=str(tmp_path / "missing.exe"))
    except DocumentConversionError as exc:
        assert "LibreOffice" in str(exc)
        assert ".pptx" in str(exc)
    else:
        raise AssertionError("无 LibreOffice 时应报错")


def test_probe_is_extension_aware():
    from printcase.office import probe_document_converter

    ok_word, detail_word = probe_document_converter(".docx")
    ok_pptx, detail_pptx = probe_document_converter(".pptx")
    assert isinstance(ok_word, bool) and detail_word
    assert isinstance(ok_pptx, bool) and detail_pptx
    if not ok_pptx:
        assert "LibreOffice" in detail_pptx


# ---------- 模板下拉框（数据有效性） ----------

from xlsx_dv import read_validations as _read_validations  # noqa: E402  数据验证读取（含 x14 扩展）


def _options_from_formula(wb, formula: str) -> list[str]:
    """把 DV 的 formula1（选项表范围或 IF 公式）解析成可选值列表。"""
    import re as _re

    values: list[str] = []
    for sheet_name, col, low, end_col, high in _re.findall(r"(选项)!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", formula):
        ws = wb[sheet_name]
        if col != end_col:
            continue
        for row in range(int(low), int(high) + 1):
            cell = ws[f"{col}{row}"]
            if cell.value is not None:
                values.append(str(cell.value))
    return values


def test_template_has_dropdowns_for_enum_columns():
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter as gcl

    template = Path(__file__).resolve().parents[1] / "templates" / "PrintCase_Template.xlsx"
    wb = load_workbook(template)
    ws = wb["用例"]
    headers = [cell.value for cell in ws[1]]
    dv_by_col = {}
    for dv in _read_validations(template, "用例"):
        for cell_range in str(dv.sqref).split():
            column = "".join(ch for ch in cell_range.split(":")[0] if ch.isalpha())
            dv_by_col[column] = dv

    enums = ["是否启用", "方向", "颜色", "单双面", "打印顺序", "混合尺寸",
             "跳过空白页", "无边距", "分隔页", "装订位置", "装订方式", "纸张", "纸张类型", "纸盒", "打印质量"]
    for header in enums:
        column = gcl(headers.index(header) + 1)
        dv = dv_by_col.get(column)
        assert dv is not None, f"{header} 列缺少下拉"
        assert dv.type == "list"
        assert not dv.showDropDown, f"{header} 的 showDropDown 必须为 False，否则 Excel 会隐藏箭头"
        assert len(_options_from_formula(wb, str(dv.formula1))) >= 2, header

    # 纸张类型允许手填驱动介质代号（下拉仅作推荐）
    dv_media = dv_by_col[gcl(headers.index("纸张类型") + 1)]
    assert dv_media.type == "list" and not dv_media.showErrorMessage

    # 已移除的列不能再出现
    for removed in ("自定义纸张", "逆序", "缩放", "超时秒", "打印机"):
        assert removed not in headers, f"{removed} 应已移除"

    # 无边距用 Y/N 表示
    assert _options_from_formula(wb, str(dv_by_col[gcl(headers.index("无边距") + 1)].formula1)) == ["Y", "N"]

    # 数字范围校验（超时秒列已删除，仅剩份数）
    for header, low, high in (("份数", "1", "9999"),):
        dv = dv_by_col[gcl(headers.index(header) + 1)]
        assert dv.type == "whole" and dv.formula1 == low and dv.formula2 == high

    # 示例数据必须落在下拉取值内（否则用户照抄示例会立刻被 Excel 标红）
    columns = {header: idx for idx, header in enumerate(headers)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[columns["用例编号"]]:
            continue
        for header in enums:
            value = row[columns[header]]
            if value in (None, ""):
                continue
            dv = dv_by_col[gcl(columns[header] + 1)]
            allowed = _options_from_formula(wb, str(dv.formula1))
            assert str(value) in allowed, (row[columns["用例编号"]], header, value, allowed)


def get_column_letter_for(headers, header):
    from openpyxl.utils import get_column_letter
    return get_column_letter(headers.index(header) + 1)


# ---------- 模板下拉框（数据有效性） ----------

def get_column_letter_for(headers, header):
    from openpyxl.utils import get_column_letter
    return get_column_letter(headers.index(header) + 1)


# ---------- 驱动下拉选项对齐 + 联动规则 ----------

def test_paper_labels_normalize_to_driver_tokens():
    from printcase.validator import normalize_paper

    cases = {
        "A4 (210 x 297mm)": "a4",
        "JIS B4 (257 x 364mm)": "jisb4",
        "Ledger (11 x 17 inch)": "ledger",
        "8K (270 x 390mm)": "8k",
        "Big16K (195 x 270mm)": "big16k",
        "A3Wide (12 x 18 inch)": "a3wide",
        "SRA3 (320 x 450mm)": "sra3",
        "自动": "auto",
    }
    for label, token in cases.items():
        assert normalize_paper(label) == token


def test_driver_specific_paper_uses_driver_order(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, resolved_path=pdf, paper="SRA3 (320 x 450mm)")
    assert validate_case(case, {"Test Printer"}, "Test Printer") == []
    from printcase.driver_ui import plan_actions
    # 专有尺寸按驱动上报顺序在界面里选中，并用驱动代号回读校验
    caps = {"papers": ["SRA3 (320 x 450mm)"], "paper_ids": [130]}
    action = _action(plan_actions(case, caps), "纸张尺寸")
    assert action.kind == "combo_index" and action.value == 0
    assert action.expect_value == 130


def test_named_tray_and_auto_source(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, resolved_path=pdf, bin="纸盒1")
    assert validate_case(case, {"Test Printer"}, "Test Printer") == []
    from printcase.driver_ui import plan_actions
    caps = {"bin_names": ["自动选择", "纸盒1"], "bin_ids": [7, 1]}
    action = _action(plan_actions(case, caps), "纸张来源")
    assert action.kind == "combo_index" and action.value == 1
    assert action.expect_value == 1
    assert "纸盒:纸盒1" in case.settings_summary()


def test_simplex_only_media_rejects_duplex(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    for media in ("信封纸", "明信片", "标签纸", "透明胶片"):
        case = _make_case(tmp_path, resolved_path=pdf, media_type=media, duplex="duplexlong")
        assert any("不支持双面" in e for e in validate_case(case, set())), media
    ok = _make_case(tmp_path, resolved_path=pdf, media_type="透明胶片", duplex="simplex")
    assert validate_case(ok, {"Test Printer"}, "Test Printer") == []


def test_separator_source_requires_separator(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF")
    case = _make_case(tmp_path, resolved_path=pdf, separator_page="none", separator_bin="多功能进纸器")
    assert any("灰化" in e for e in validate_case(case, {"Test Printer"}, "Test Printer"))


def test_plan_resolves_paper_and_tray_from_driver_caps(tmp_path):
    from printcase.driver_ui import plan_actions

    caps = {
        "papers": ["A4 (210 x 297mm)", "SRA3 (320 x 450mm)"],
        "paper_ids": [9, 130],
        "bin_names": ["自动选择", "多功能进纸器", "纸盒1"],
        "bin_ids": [7, 4, 1],
    }
    case = _make_case(tmp_path, paper="SRA3 (320 x 450mm)", bin="纸盒1")
    actions = plan_actions(case, caps)
    assert _action(actions, "纸张尺寸").value == 1
    assert _action(actions, "纸张尺寸").expect_value == 130
    assert _action(actions, "纸张来源").value == 2
    assert _action(actions, "纸张来源").expect_value == 1


def test_tray_missing_from_driver_table_falls_back_to_name_match(tmp_path):
    """驱动上报的纸盒表与用例取值对不上时按名称选，不能靠序号误配到别的盒。"""
    from printcase.driver_ui import plan_actions

    caps = {"bin_names": ["自动选择", "多功能进纸器", "纸盒3", "纸盒4"], "bin_ids": [7, 262, 259, 258]}
    case = _make_case(tmp_path, bin="纸盒1")
    action = _action(plan_actions(case, caps), "纸张来源")
    assert action.kind == "combo_text" and action.value == "纸盒1"
    assert action.expect_value is None


def test_template_matches_driver_dropdown_options():
    from openpyxl import load_workbook

    template = Path(__file__).resolve().parents[1] / "templates" / "PrintCase_Template.xlsx"
    wb = load_workbook(template)
    assert "选项" in wb.sheetnames
    options = wb["选项"]
    assert options.sheet_state == "hidden"

    papers = [c.value for c in options["A"] if c.value]
    assert papers[0] == "纸张"
    assert "SRA3 (320 x 450mm)" in papers and "JIS B4 (257 x 364mm)" in papers
    assert "自定义" not in papers  # 自定义纸型已移除
    assert "混合尺寸" not in papers  # 改为独立开关列
    assert "自动" in papers
    media = [c.value for c in options["B"] if c.value]
    assert "透明胶片" in media and "厚纸3(210-256gsm)" in media
    bins = [c.value for c in options["C"] if c.value]
    assert bins[1:] == ["自动选择", "多功能进纸器", "纸盒1", "纸盒2"]

    ws = wb["用例"]
    headers = [c.value for c in ws[1]]
    formulas = {str(dv.sqref): dv.formula1 for dv in _read_validations(template, "用例")}
    # 所有列表校验都引用隐藏的“选项”表（避免 255 字符上限）
    assert all("选项!" in str(f) for f in formulas.values() if str(f) != "1" and str(f) != "10")
    # 混合尺寸是独立开关列，紧跟在纸张之后
    assert headers[headers.index("纸张") + 1] == "混合尺寸"
    # 打印质量只有 600 / 1200
    from openpyxl.utils import get_column_letter as _gcl
    quality_letter = _gcl(headers.index("打印质量") + 1)
    quality_ref = next(ref for ref in formulas if ref.startswith(f"{quality_letter}2"))
    assert _options_from_formula(wb, str(formulas[quality_ref])) == ["600", "1200"]
    # 按用户要求：不做灰度联动，表格里不应有条件格式规则
    assert list(ws.conditional_formatting) == []


# ---------- WPS 转换后端 ----------

class FakeWPSDoc:
    def ExportAsFixedFormat(self, path, fmt):
        Path(path).write_bytes(b"%PDF-fake-wps")
    def Close(self, save_changes):
        pass


class FakeWPS:
    def __init__(self):
        self.Documents = self
    def Open(self, path, **kwargs):
        return FakeWPSDoc()
    def Quit(self):
        pass


def test_backend_order_includes_wps_fallback():
    from printcase.office import backend_order

    assert backend_order(".docx") == ["word", "wps", "soffice"]
    assert backend_order(".odt") == ["soffice", "wps", "word"]
    assert backend_order(".wps") == ["soffice", "wps", "word"]
    assert backend_order(".pptx") == ["soffice"]


def test_wps_conversion_backend(tmp_path):
    from printcase.office import convert_document_to_pdf

    doc = tmp_path / "memo.docx"
    doc.write_bytes(b"PK")
    out = convert_document_to_pdf(doc, work_dir=tmp_path / "out", prefer="wps", wps_factory=FakeWPS)
    assert out.name == "memo.pdf"
    assert out.read_bytes().startswith(b"%PDF")


def test_word_falls_back_to_wps_then_reports(tmp_path):
    from printcase.office import DocumentConversionError, convert_document_to_pdf

    doc = tmp_path / "memo.rtf"
    doc.write_bytes(b"rtf")
    # word_factory 抛错模拟无 Word；wps 正常 → 应走 wps 后端成功
    def broken_word():
        raise RuntimeError("Word not installed")
    out = convert_document_to_pdf(
        doc, work_dir=tmp_path / "out", prefer="word",
        word_factory=broken_word, wps_factory=FakeWPS,
    )
    assert out.read_bytes().startswith(b"%PDF")

    # 无任何组件 → 明确报错
    doc2 = tmp_path / "memo2.rtf"
    doc2.write_bytes(b"rtf")
    try:
        convert_document_to_pdf(
            doc2, work_dir=tmp_path / "out2", prefer="word",
            word_factory=broken_word,
            wps_factory=lambda: (_ for _ in ()).throw(RuntimeError("no wps")),
        )
    except DocumentConversionError as exc:
        assert "wps" in str(exc)
    else:
        raise AssertionError("应抛出 DocumentConversionError")


# ---------- 属性生效保证（严格模式） ----------

def test_driver_ui_verify_blocks_when_driver_ignores_value(monkeypatch):
    """驱动回读值与用例要求不一致时，用例必须被拦截（不下发）。"""
    from printcase import driver_ui

    monkeypatch.setattr(driver_ui, "snapshot", lambda printer, w=None: {"PrintQuality": 600, "PaperSize": 9, "MediaType": 266})
    actions = [driver_ui.Action("基本", "打印质量", "radio", ("1200 dpi", "1200"), "PrintQuality", 1200)]
    problems = driver_ui.verify("Test Printer", actions, {"PrintQuality": 600, "PaperSize": 9, "MediaType": 266})
    assert any("PrintQuality" in item for item in problems)


def test_unreported_media_blocks_case_in_preflight():
    from printcase.capabilities import validate_driver_capabilities
    from printcase.models import PrintCase

    case = PrintCase(case_id="T", enabled=True, file_path="x.pdf", media_type="厚纸9(300-350gsm)")
    caps = {"papers": None, "media_names": ["普通纸(60-90gsm)", "厚纸2(158-209gsm)"]}
    errors = validate_driver_capabilities(case, caps)
    assert any("厚纸9" in e and "驱动能力检查" in e for e in errors), errors


def test_private_media_reported_by_driver_passes_preflight():
    """Pantum 私有介质（厚纸1+）由驱动上报 → 不再被拦截。"""
    from printcase.capabilities import validate_driver_capabilities
    from printcase.models import PrintCase

    case = PrintCase(case_id="TC-003", enabled=True, file_path="x.pdf", media_type="厚纸1+(121-157gsm)")
    caps = {"papers": [], "media_names": ["普通纸(60-90gsm)", "厚纸1(106-120gsm)", "厚纸1+(121-157gsm)"]}
    assert validate_driver_capabilities(case, caps) == []


def test_runner_blocks_case_when_driver_ui_rejects(tmp_path):
    from printcase.driver_ui import DriverUiError
    from printcase.models import PrintCase
    from printcase.runner import PrintRunner

    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"%PDF")

    def submit(*_args, **_kwargs):
        raise AssertionError("属性无法保证生效时不应下发打印")

    def rejecting_applier(printer, case, log=None, caps=None):
        raise DriverUiError("纸张类型“厚纸1+(121-157gsm)”在驱动界面里找不到")

    case = PrintCase(
        case_id="TC-MEDIA", enabled=True, file_path=str(pdf), resolved_path=pdf,
        printer="Test Printer", media_type="厚纸1+(121-157gsm)",
    )
    runner = PrintRunner(
        JobThenClearSpooler(), AuditStore(tmp_path / "a5.sqlite3"),
        queue_timeout=1, submitter=submit, driver_applier=rejecting_applier,
        document_converter=lambda *a, **k: (_ for _ in ()).throw(AssertionError("无文档转换")),
    )
    result = runner.run_cases([case], "Test Printer", "R5", lambda *_: None)[0]
    assert result.status == CaseStatus.FAILED
    assert "本次不下发" in result.message
    assert "厚纸1+" in result.message


def test_plan_matches_driver_name_with_full_label(tmp_path):
    from printcase.driver_ui import plan_actions

    # 驱动上报带克重的完整名称，用例填完整标签 → 精确落在同一下拉项
    caps = {"media_names": ["自动选择", "普通纸", "厚纸2(158-209gsm)"]}
    case = _make_case(tmp_path, media_type="厚纸2(158-209gsm)")
    action = _action(plan_actions(case, caps), "纸张类型")
    assert action.kind == "combo_index" and action.value == 2


# ---------- 内置直打引擎（GDI / Office COM） ----------

class FakeSession:
    """替身：记录 GDI 打印会话的关键动作（创建 DC / 逐页绘制 / 关闭）。"""

    def __init__(self, printer, doc_name):
        self.calls = [("open", printer, doc_name)]
        self.pages = []

    def draw_image(self, image, box=None):
        self.calls.append(("page", image.size))
        self.pages.append(image.size)

    def close(self):
        self.calls.append(("close",))


def _session_factory(store: list):
    def factory(printer, doc_name):
        session = FakeSession(printer, doc_name)
        store.append(session)
        return session
    return factory


def _tiny_pdf(path: Path, pages: int = 3) -> Path:
    import pymupdf

    doc = pymupdf.open()
    for i in range(pages):
        page = doc.new_page()
        page.insert_text((72, 100), f"Page {i + 1}")
    doc.save(str(path))
    doc.close()
    return path


def test_print_pdf_gdi_sends_selected_pages_in_order(tmp_path):
    from printcase.direct_print import print_pdf_gdi

    pdf = _tiny_pdf(tmp_path / "sample.pdf", pages=3)
    store: list = []
    case = _make_case(tmp_path, resolved_path=pdf, pages="1-2", copies=2, collate="yes")
    logs = []
    print_pdf_gdi("Test Printer", case, pdf, logs.append, session_factory=_session_factory(store))
    assert len(store) == 1
    session = store[0]
    assert session.calls[0] == ("open", "Test Printer", "sample")  # 选择驱动：以打印机名创建 DC
    assert len(session.pages) == 4                    # 2 页 × 2 份
    assert session.calls[-1] == ("close",)            # 结束后释放 DC
    assert any("1/4" in text for text in logs)        # 前端可见的逐页进度


def test_print_image_gdi_single_page(tmp_path):
    from PIL import Image

    from printcase.direct_print import print_image_gdi

    img_path = tmp_path / "sample.png"
    Image.new("RGB", (400, 300), "white").save(img_path)
    store: list = []
    case = _make_case(tmp_path, resolved_path=img_path, copies=1)
    print_image_gdi("Test Printer", case, img_path, lambda text: None, session_factory=_session_factory(store))
    assert len(store) == 1
    assert store[0].calls[0] == ("open", "Test Printer", "sample")
    assert len(store[0].pages) == 1


def test_print_document_com_selects_printer_and_printout(tmp_path):
    from printcase.direct_print import print_document_com

    docx = tmp_path / "报告.docx"
    docx.write_bytes(b"PK")
    events = []

    class FakeDoc:
        def PrintOut(self, **kwargs):
            events.append(("printout", kwargs))
        def Close(self, save_changes):
            events.append(("close", save_changes))

    class FakeWordApp:
        ActivePrinter = ""
        DisplayAlerts = -1
        def __init__(self):
            self.Documents = self
        def Open(self, path, **kwargs):
            events.append(("open", Path(path).name, kwargs.get("ReadOnly")))
            return FakeDoc()
        def Quit(self):
            events.append(("quit",))

    case = _make_case(tmp_path, resolved_path=docx, copies=2, collate="yes", pages="1-2")
    print_document_com("Pantum CM806ADN", case, docx, lambda text: events.append(("log", text)), word_factory=FakeWordApp)

    kinds = [e[0] for e in events]
    assert "open" in kinds and "printout" in kinds and "quit" in kinds
    opened = next(e for e in events if e[0] == "open")
    assert opened[1] == "报告.docx" and opened[2] is True
    printout = next(e for e in events if e[0] == "printout")
    assert printout[1]["Copies"] == 2 and printout[1]["Collate"] is True
    assert printout[1]["Pages"] == "1-2"
    # 前端可见的驱动选择步骤
    logs = [e[1] for e in events if e[0] == "log"]
    assert any("选择打印机驱动：Pantum CM806ADN" in text for text in logs)


def test_com_word_failure_falls_back_to_wps(tmp_path):
    from printcase.direct_print import print_document_com

    docx = tmp_path / "memo.docx"
    docx.write_bytes(b"PK")

    def broken_word():
        raise RuntimeError("Word not installed")

    class FakeWPSApp:
        ActivePrinter = ""
        def __init__(self):
            self.Documents = self
        def Open(self, path, **kwargs):
            class D:
                def PrintOut(self, **kwargs):
                    pass
                def Close(self, save_changes):
                    pass
            return D()
        def Quit(self):
            pass

    events = []
    print_document_com("P", _make_case(tmp_path, resolved_path=docx), docx, events.append,
                       word_factory=broken_word, wps_factory=FakeWPSApp)
    assert any("WPS" in text for text in events)  # 回退到 WPS 文字
