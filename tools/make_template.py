"""重新生成 templates/PrintCase_Template.xlsx（全中文模板）。

选项值对齐 Pantum CM806ADN Series 驱动“纸张”属性页的下拉项。
按用户要求：不做灰度联动（互斥关系由执行前检查负责），
混合尺寸为独立开关，打印质量仅 600/1200。

运行：.venv/Scripts/python.exe tools/make_template.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.attr_catalog import ATTRIBUTES, PAGES, page_display  # noqa: E402
from printcase.extended_attrs import default_examples  # noqa: E402
OUT = ROOT / "templates" / "PrintCase_Template.xlsx"
# 允许通过命令行参数指定输出路径（模板在 Excel 中打开时先写临时文件再替换）
if len(sys.argv) > 1:
    OUT = Path(sys.argv[1]).resolve()

HEADERS = [
    "用例编号", "是否启用", "文件路径", "页码", "份数", "纸张", "混合尺寸",
    "方向", "颜色", "单双面", "纸盒", "打印顺序", "打印质量", "纸张类型",
    "跳过空白页", "无边距", "分隔页", "分隔页纸张来源", "装订位置", "装订方式",
    "备注/预期结果",
]

# ---- 驱动下拉项（严格按驱动截图录入）----
PAPERS = [
    "A4 (210 x 297mm)", "A3 (297 x 420mm)", "A5 (148 x 210mm)",
    "JIS B4 (257 x 364mm)", "JIS B5 (182 x 257mm)", "JIS B6 (128 x 182mm)",
    "Letter (8.5 x 11inch)", "Legal (8.5 x 14inch)", "Folio (8.5 x 13inch)",
    "Ledger (11 x 17 inch)", "Executive (7.25 x 10.5inch)", "Statement (5.5 x 8.5inch)",
    "8K (270 x 390mm)", "Big16K (195 x 270mm)", "A6 (105 x 148mm)",
    "A3Wide (12 x 18 inch)", "SRA3 (320 x 450mm)",
    "自动",
]
MEDIA_TYPES = [
    "普通纸(60-90gsm)", "普通纸+(91-105gsm)",
    "厚纸1(106-120gsm)", "厚纸1+(121-157gsm)", "厚纸2(158-209gsm)", "厚纸3(210-256gsm)",
    "信封纸", "明信片", "标签纸", "透明胶片",
]
PAPER_SOURCES = ["自动选择", "多功能进纸器", "纸盒1", "纸盒2"]
SEPARATORS = ["无", "前", "后", "前后", "份数之间"]
BIND_LOCATIONS = ["无", "左", "右", "上", "下", "长边", "短边"]
PRINT_ORDERS = ["逐份", "非逐份"]
DUPLEX_MODES = ["单面", "双面长边", "双面短边"]
YES_NO = ["是", "否"]
COLORS = ["彩色", "黑白"]
ORIENTS = ["纵向", "横向"]
Y_N = ["Y", "N"]
BIND_METHODS = ["无", "订书钉", "骑马钉", "打孔", "胶装"]
QUALITIES = ["600", "1200"]

# 选项表列顺序（A..L），供数据有效性引用
OPTION_COLUMNS: list[tuple[str, list[str]]] = [
    ("纸张", PAPERS),
    ("纸张类型", MEDIA_TYPES),
    ("纸张来源", PAPER_SOURCES),
    ("分隔页", SEPARATORS),
    ("装订位置", BIND_LOCATIONS),
    ("打印顺序", PRINT_ORDERS),
    ("是/否", YES_NO),
    ("颜色", COLORS),
    ("方向", ORIENTS),
    ("Y/N", Y_N),
    ("装订方式", BIND_METHODS),
    ("打印质量", QUALITIES),
    ("单双面", DUPLEX_MODES),
]
COL = {name: get_column_letter(idx + 1) for idx, (name, _) in enumerate(OPTION_COLUMNS)}
RANGE = {name: f"选项!${COL[name]}$2:${COL[name]}${len(values) + 1}" for name, values in OPTION_COLUMNS}


def _col(header: str) -> str:
    return get_column_letter(HEADERS.index(header) + 1)


ROWS = [
    ["TC-001", "是", "docs/A4-check.pdf", "", 1, "A4 (210 x 297mm)", "否", "纵向", "彩色", "单面",
     "自动选择", "逐份", "600", "普通纸(60-90gsm)", "否", "N", "无", "", "无", "无",
     "A4 单面基础打印（建议先试打这一条）"],
    ["TC-002", "是", "docs/duplex-check.pdf", "1-2", 2, "A4 (210 x 297mm)", "否", "纵向", "黑白", "双面长边",
     "自动选择", "逐份", "600", "普通纸(60-90gsm)", "否", "N", "份数之间", "", "长边", "骑马钉",
     "双面+装订示例：共 2 份，份与份之间自动插入分隔页"],
    ["TC-003", "否", "docs/tray-check.pdf", "1", 2, "自动", "是", "横向", "彩色", "单面",
     "自动选择", "逐份", "1200", "厚纸2(158-209gsm)", "否", "Y", "前", "多功能进纸器", "左", "订书钉",
     "混合尺寸示例：开启后纸盒固定为自动选择，由驱动按每页尺寸走纸"],
    ["TC-004", "否", "docs/A4-check.pdf", "", 1, "SRA3 (320 x 450mm)", "否", "横向", "彩色", "单面",
     "自动选择", "非逐份", "600", "厚纸1(106-120gsm)", "否", "N", "无", "", "无", "无",
     "SRA3 大幅面示例：该尺寸由 DEVMODE 精确指定"],
]

GUIDE_HEADER = ["字段（中文列名）", "英文原名 / 别名", "必填", "说明与可选值"]

GUIDE_ROWS = [
    ["用例编号", "case_id / 编号", "是", "每条用例必须唯一，例如 TC-001。"],
    ["是否启用", "enabled / 启用 / 执行", "否", "是/TRUE/1 执行；否/FALSE/0 跳过。默认“是”。"],
    ["文件路径", "file_path / 源文件", "是", "打印源文件；相对路径按本 Excel 所在目录解析。支持 PDF、XPS/OXPS、DjVu、CBZ/CBR、常见图片，以及 Office 文档：docx/doc/rtf（Word 优先）与 odt/wps/pptx/xls 等（需 LibreOffice），均自动转换为 PDF 后打印。"],
    ["页码", "pages / 页范围", "否", "留空表示全部；示例：1-3,5 或 last。"],
    ["份数", "copies / 打印份数", "否", "1–9999 的整数（对齐驱动份数上限）。"],
    ["纸张", "paper / 纸张尺寸", "否", "下拉项与驱动“纸张尺寸”一致：A4/A3/A5/JIS B4/JIS B5/JIS B6/Letter/Legal/Folio/Ledger/Executive/Statement/8K/Big16K/A6/A3Wide/SRA3，另有“自动”（跟随驱动默认）。开启“混合尺寸”时本列被忽略。"],
    ["混合尺寸", "mixed_sizes / 混合尺寸打印", "否", "是/否，对应驱动“混合尺寸”勾选框。选“是”时由驱动按文档每页尺寸走纸，纸张列忽略，且纸盒只能为“自动选择”。"],
    ["方向", "orientation / 横纵向", "否", "纵向 / 横向。最终方向还受驱动配置影响。"],
    ["颜色", "color / 彩色模式", "否", "彩色 / 黑白。"],
    ["单双面", "duplex / 双面模式", "否", "单面 / 双面长边 / 双面短边。纸张类型为信封纸/明信片/标签纸/透明胶片 时只能选单面（执行前检查会拦截双面）。"],
    ["纸盒", "bin / 纸张来源", "否", "下拉项与驱动“纸张来源”一致：自动选择 / 多功能进纸器 / 纸盒1 / 纸盒2。开启混合尺寸时必须为“自动选择”。"],
    ["打印顺序", "collate / 逐份", "否", "逐份 / 非逐份，对应驱动“打印顺序”选择；非逐份即按页堆叠输出。份数为 1 时无实际意义。"],
    ["打印质量", "print_quality / dpi", "否", "600 / 1200，对应驱动“打印质量”下拉（无“自动”项）。留空则沿用驱动当前默认。"],
    ["纸张类型", "media_type / 纸型", "否", "下拉项与驱动“纸张类型”一致：普通纸(60-90gsm)、普通纸+(91-105gsm)、厚纸1/1+/2/3、信封纸、明信片、标签纸、透明胶片。通过 DEVMODE 写入驱动；驱动未上报对应介质代号时按默认执行并记日志。"],
    ["跳过空白页", "skip_blank", "否", "是 / 否。驱动私有设置，按驱动默认处理并记录测试意图。"],
    ["无边距", "borderless / 满版打印", "否", "Y / N（默认 N）。无边距依赖驱动首选项中的无边距/满版设置，命令行无法强制开启；填 Y 前请先在驱动中配置好。"],
    ["分隔页", "separator_page / 间隔页", "否", "无 / 前 / 后 / 前后 / 份数之间：在主作业前/后或每份之间插入含用例 ID、文件名、时间的间隔页。"],
    ["分隔页纸张来源", "separator_bin / 分隔页纸盒", "否", "分隔页所用纸盒，取值同纸张来源；分隔页为“无”时请留空。"],
    ["装订位置", "binding_location / 装订边", "否", "无 / 左 / 右 / 上 / 下 / 长边 / 短边。长边、短边要求双面打印，否则执行前检查不通过。"],
    ["装订方式", "binding_method / 装订", "否", "无 / 订书钉 / 骑马钉 / 打孔 / 胶装。实际效果取决于驱动与选购装订器。"],
    ["备注/预期结果", "notes / 备注", "否", "测试备注或预期，不自动检测纸面质量。"],
]

GUIDE_FOOTER = [
    ["执行建议", "", "", "先运行“打印前检查”；先选一条用例试打验证，再批量执行。表格不做灰度提示，互斥关系由检查阶段统一拦截。"],
    ["打印机", "", "", "模板不含打印机列：目标打印机统一在软件界面顶部选择；旧模板如带“打印机”列仍兼容（逐条覆盖）。"],
    ["驱动扩展属性", "工作簿同名表", "否", "一行一个属性（用例编号/页面/属性/值）：纸张、版面、水印、高级、密码打印、封页等驱动专属项。页面与属性均为下拉框，取值见“属性速查”表。仅在“UI 自动化（pywinauto）”下发方式下生效。"],
    ["旧项目用例导入", "", "", "既有 print_panel 项目的 print.xlsx 可直接导入：软件会识别 sheet “print” 的 &属性名=值 文本块并转换为扩展属性。"],
    ["互斥规则", "", "", "混合尺寸=是 → 纸张列忽略、纸盒必须“自动选择”；打印顺序二选一（逐份/非逐份）；纸张类型=信封纸/明信片/标签纸/透明胶片 → 只能单面；分隔页=无 → 分隔页纸张来源必须为空；装订位置=长边/短边 → 必须双面。"],
    ["下拉选项", "", "", "固定取值列均为下拉框（选项维护在隐藏的“选项”工作表）。纸张、纸张类型、纸盒、装订位置、分隔页选项与驱动属性页一致；无边距用 Y/N 表示。"],
    ["状态说明", "", "", "队列完成/队列清除不能证明纸张已实际出纸；不确定状态不要直接重复打印。"],
]

VALIDATION_MAX_ROW = 500


def _style_header(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill("solid", fgColor="2367D1")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(*[Side(style="thin", color="B9C6D6")] * 4)


def _add_list_dv(ws, header: str, source: str, prompt: str = "", strict: bool = True) -> None:
    column = _col(header)
    dv = DataValidation(
        type="list",
        formula1=source,
        allow_blank=True,
        showDropDown=False,  # True 会隐藏下拉箭头，必须保持 False
        showErrorMessage=strict,
        showInputMessage=True,
    )
    dv.errorTitle = "取值不在允许范围"
    dv.error = "请从下拉列表中选择合法取值"
    dv.promptTitle = header
    dv.prompt = prompt or "点击单元格右侧下拉箭头选择"
    ws.add_data_validation(dv)
    dv.add(f"{column}2:{column}{VALIDATION_MAX_ROW}")


def _add_numeric_dv(ws, header: str, low: str, high: str) -> None:
    column = _col(header)
    dv = DataValidation(
        type="whole", operator="between", formula1=low, formula2=high,
        allow_blank=True, showErrorMessage=True, showInputMessage=True,
    )
    dv.errorTitle = "数值超出范围"
    dv.error = f"请输入 {low}–{high} 之间的整数"
    dv.promptTitle = header
    dv.prompt = f"请输入 {low}–{high} 之间的整数"
    ws.add_data_validation(dv)
    dv.add(f"{column}2:{column}{VALIDATION_MAX_ROW}")


EXT_SHEET_HEADERS = ["用例编号", "页面", "属性", "值", "说明"]


def _build_extended_sheet(wb) -> None:
    """驱动扩展属性表（长表）：一行一个属性，替代既有项目的 &属性=值 文本块。"""
    ws = wb.create_sheet("驱动扩展属性")
    ws.append(EXT_SHEET_HEADERS)
    rows = default_examples("TC-001")
    rows += [
        ("TC-002", page_display("driver_basic"), "逆序(0或1)", "1", "逐份与逆序由驱动决定，注意与“打印顺序”的语义差异"),
        ("TC-002", page_display("driver_paper"), "装订设置-订钉", "双钉", "需要装订器支持"),
        ("TC-003", page_display("driver_watermark"), "新建文本水印内容", "测试专用", "先选“新建文本水印”，再填内容/角度等"),
        ("TC-003", page_display("driver_password"), "样本打印(1)", "1", "样本打印与普通/延时/密码打印互斥"),
    ]
    for row in rows:
        ws.append(list(row))
    for idx, width in enumerate([12, 18, 40, 22, 46], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        _style_header(cell)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 24


def _build_attribute_catalog_sheet(wb) -> None:
    """属性速查表：列出全部扩展属性及取值类型/示例（只读参考）。"""
    ws = wb.create_sheet("属性速查")
    ws.append(["页面", "属性名", "取值类型", "示例", "说明"])
    for spec in ATTRIBUTES:
        ws.append([page_display(spec.page), spec.label, spec.kind, spec.example, spec.hint or "、".join(spec.options[:4])])
    for idx, width in enumerate([18, 44, 12, 24, 48], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        _style_header(cell)


def _build_options_sheet(wb) -> None:
    ws = wb.create_sheet("选项")
    columns = list(OPTION_COLUMNS) + [
        ("扩展-页面", [page_display(key) for key in PAGES]),
        ("扩展-属性", [spec.label for spec in ATTRIBUTES]),
    ]
    for idx, (name, values) in enumerate(columns, start=1):
        ws.cell(row=1, column=idx, value=name)
        for row, value in enumerate(values, start=2):
            ws.cell(row=row, column=idx, value=value)
        ws.column_dimensions[get_column_letter(idx)].width = 22
    _style_header(ws.cell(row=1, column=1))
    ws.sheet_state = "hidden"


def main() -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "用例"
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(row)
    widths = [10, 9, 24, 8, 6, 22, 10, 7, 7, 11, 14, 10, 9, 20, 10, 8, 10, 14, 10, 10, 44]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "C2"
    for cell in ws[1]:
        _style_header(cell)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=(cell.column == len(HEADERS)))
    ws.row_dimensions[1].height = 24

    guide = wb.create_sheet("填写说明")
    guide.append(GUIDE_HEADER)
    for row in GUIDE_ROWS + GUIDE_FOOTER:
        guide.append(row)
    for idx, width in enumerate([18, 26, 6, 84], start=1):
        guide.column_dimensions[get_column_letter(idx)].width = width
    guide.freeze_panes = "A2"
    for cell in guide[1]:
        _style_header(cell)
    for row in guide.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    _build_options_sheet(wb)
    _build_extended_sheet(wb)
    _build_attribute_catalog_sheet(wb)

    # 固定取值列 → 下拉（引用“选项”表，绕开 255 字符上限）
    _add_list_dv(ws, "是否启用", RANGE["是/否"])
    _add_list_dv(ws, "纸张", RANGE["纸张"])
    _add_list_dv(ws, "混合尺寸", RANGE["是/否"], prompt="是：按文档每页尺寸走纸，纸盒固定为自动选择")
    _add_list_dv(ws, "方向", RANGE["方向"])
    _add_list_dv(ws, "颜色", RANGE["颜色"])
    _add_list_dv(ws, "单双面", RANGE["单双面"])
    _add_list_dv(ws, "纸盒", RANGE["纸张来源"])
    _add_list_dv(ws, "打印顺序", RANGE["打印顺序"])
    _add_list_dv(ws, "打印质量", RANGE["打印质量"])
    _add_list_dv(ws, "纸张类型", RANGE["纸张类型"], strict=False,
                 prompt="可从下拉选择，也可输入驱动介质代号数字")
    _add_list_dv(ws, "跳过空白页", RANGE["是/否"])
    _add_list_dv(ws, "无边距", RANGE["Y/N"], prompt="Y 开启无边距、N 关闭（默认 N）")
    _add_list_dv(ws, "分隔页", RANGE["分隔页"])
    _add_list_dv(ws, "分隔页纸张来源", RANGE["纸张来源"])
    _add_list_dv(ws, "装订位置", RANGE["装订位置"])
    _add_list_dv(ws, "装订方式", RANGE["装订方式"])

    _add_numeric_dv(ws, "份数", "1", "9999")

    # 扩展属性表：页面 / 属性 下拉（引用“选项”表）
    ext = wb["驱动扩展属性"]
    options = wb["选项"]
    page_col = get_column_letter(len(OPTION_COLUMNS) + 1)
    label_col = get_column_letter(len(OPTION_COLUMNS) + 2)
    ext_rows = options.max_row
    for column, source_col, title in (("B", page_col, "页面"), ("C", label_col, "属性")):
        dv = DataValidation(
            type="list",
            formula1=f"选项!${source_col}$2:${source_col}${ext_rows}",
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=True,
            showInputMessage=True,
        )
        dv.errorTitle = "取值不在允许范围"
        dv.error = f"请从下拉列表中选择{title}"
        dv.promptTitle = title
        dv.prompt = "点击单元格右侧下拉箭头选择"
        ext.add_data_validation(dv)
        dv.add(f"{column}2:{column}{VALIDATION_MAX_ROW}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成：{OUT}")


if __name__ == "__main__":
    main()
