from __future__ import annotations

import re
from pathlib import Path

from .models import PrintCase
from .office import OFFICE_EXTENSIONS

_DIRECT_EXTENSIONS = {".pdf", ".xps", ".oxps", ".djvu", ".djv", ".cbz", ".cbr", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
# 内置直打支持的格式 + 需先转 PDF 的 Office 系列格式
SUPPORTED_EXTENSIONS = _DIRECT_EXTENSIONS | OFFICE_EXTENSIONS
# 纸张标准 token（与 Pantum CM806ADN 驱动“纸张尺寸”下拉一一对应）
PAPER_TOKENS = {
    "a4", "a3", "a5", "a2", "a6",
    "jisb4", "jisb5", "jisb6",
    "letter", "legal", "folio", "ledger", "executive", "statement",
    "8k", "big16k", "a3wide", "sra3",
    "auto", "mixed",
}
# 纸张名称归一化后由 DEVMODE PaperSize 精确下发
_PAPER_ALIASES = {
    "自动": "auto", "自动选择": "auto",
    "混合": "mixed", "混合尺寸": "mixed", "混合尺寸（按页面）": "mixed",

    "tabloid": "ledger",
}


def normalize_paper(value: str) -> str:
    """把驱动下拉标签（如 "SRA3 (320 x 450mm)"）或各种别名归一为纸张 token。"""
    text = str(value or "").strip()
    if not text:
        return ""
    if text in _PAPER_ALIASES:
        return _PAPER_ALIASES[text]
    lowered = text.casefold()
    if lowered in PAPER_TOKENS or lowered in _PAPER_ALIASES:
        return _PAPER_ALIASES.get(lowered, lowered)
    head = text.split("(")[0].strip().casefold().replace(" ", "").replace("-", "").replace("_", "")
    return _PAPER_ALIASES.get(head, head)


# 混合尺寸开关（对应驱动“混合尺寸”勾选框）
MIXED_SIZE_MODES = {"yes", "no"}
# 这些介质不支持双面（与驱动联动：选中后单双面被锁定为单面）
SIMPLEX_ONLY_MEDIA = {"信封纸", "明信片", "标签纸", "透明胶片"}
ORIENTATIONS = {"portrait", "landscape"}
COLORS = {"color", "monochrome"}
DUPLEX_MODES = {"simplex", "duplex", "duplexlong", "duplexshort"}
SCALINGS = {"fit", "shrink", "noscale"}
BORDERLESS_MODES = {"yes", "no"}
SEPARATOR_MODES = {"none", "before", "after", "both", "each_copy"}
BINDING_LOCATIONS = {"none", "left", "right", "top", "bottom", "longedge", "shortedge"}
BINDING_METHODS = {"none", "staple", "saddle", "punch", "glue"}
COLLATE_MODES = {"", "yes", "no"}
SKIP_BLANK_MODES = {"yes", "no"}
MAX_COPIES = 9999
_QUALITY_RE = re.compile(r"^(?:auto|\d{3,4})$", re.IGNORECASE)
# mixed/auto 不指定纸张：由驱动按文档每页尺寸输出
NO_EXPLICIT_PAPER = {"", "auto", "mixed"}
_PAGE_RE = re.compile(r"^(?:last|\d+(?:-\d+)?)(?:\s*,\s*(?:last|\d+(?:-\d+)?|odd|even))*$", re.IGNORECASE)


def validate_case(case: PrintCase, available_printers: set[str], fallback_printer: str = "") -> list[str]:
    errors: list[str] = list(case.input_errors)
    path = case.resolved_path or (Path(case.file_path).expanduser() if case.file_path else None)
    if path is None or not str(path).strip():
        errors.append("未填写打印文件路径")
    elif not path.exists() or not path.is_file():
        errors.append(f"打印文件不存在：{path}")
    elif path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        errors.append(f"暂不支持该文件格式：{path.suffix or '无扩展名'}")
    elif path.stat().st_size == 0:
        errors.append("打印文件大小为 0")

    printer = case.printer.strip() or fallback_printer.strip()
    if not printer:
        errors.append("未指定打印机，也没有选择默认打印机")
    elif printer not in available_printers:
        errors.append(f"系统中找不到打印机：{printer}")

    if case.pages:
        normalized_pages = case.pages.replace(" ", "")
        if not _PAGE_RE.fullmatch(normalized_pages):
            errors.append("页码格式错误，示例：1-3,5 或 last")
        elif len(normalized_pages.split(",")) > 10:
            errors.append("页码最多支持 10 段，请合并打印范围")
    if not 1 <= case.copies <= MAX_COPIES:
        errors.append(f"份数必须在 1–{MAX_COPIES} 之间")
    paper_token = normalize_paper(case.paper)
    if paper_token not in PAPER_TOKENS:
        errors.append(f"不支持的纸张：{case.paper}（请从模板下拉列表中选择纸张尺寸）")
    if case.mixed_sizes.lower() not in MIXED_SIZE_MODES:
        errors.append("混合尺寸必须为 yes / no（是/否）")
    elif case.mixed_sizes.lower() == "yes":
        if paper_token not in {"", "auto", "mixed"}:
            errors.append("开启混合尺寸时纸张由驱动按每页尺寸决定，纸张列请留空或选“自动”")
        if case.bin.strip():
            errors.append("开启混合尺寸时纸盒只能为“自动选择”，不能指定具体纸盒")
    if case.orientation.lower() not in ORIENTATIONS:
        errors.append("方向必须为 portrait 或 landscape")
    if case.color.lower() not in COLORS:
        errors.append("颜色必须为 color 或 monochrome")
    if case.duplex.lower() not in DUPLEX_MODES:
        errors.append("双面模式必须为 simplex / duplexlong / duplexshort")
    if case.scaling.lower() not in SCALINGS:
        errors.append("缩放必须为 fit / shrink / noscale")
    if case.bin and ("," in case.bin or "\n" in case.bin or "\r" in case.bin):
        errors.append("纸盒名称不能包含逗号或换行")
    if case.collate.lower() not in COLLATE_MODES:
        errors.append("打印顺序必须为 yes（逐份）/ no（逐页），留空使用驱动默认")
    if case.media_type.strip() in SIMPLEX_ONLY_MEDIA and case.duplex.lower() != "simplex":
        errors.append(f"纸张类型“{case.media_type}”不支持双面打印；请将单双面改为单面（驱动会灰化双面选项）")
    if not _QUALITY_RE.fullmatch(str(case.print_quality).lower()):
        errors.append("打印质量必须为 auto 或 DPI 数值（如 600、1200）")
    if case.media_type and ("," in case.media_type or "\n" in case.media_type or "\r" in case.media_type or len(case.media_type) > 30):
        errors.append("纸张类型不能包含逗号或换行，且不超过 30 个字符")
    if case.skip_blank.lower() not in SKIP_BLANK_MODES:
        errors.append("跳过空白页必须为 yes / no")
    if case.borderless.lower() not in BORDERLESS_MODES:
        errors.append("无边距必须为 yes / no")
    if case.separator_page.lower() not in SEPARATOR_MODES:
        errors.append("分隔页必须为 none / before / after / both / each_copy（份数之间）")
    if case.separator_bin and ("," in case.separator_bin or "\n" in case.separator_bin or "\r" in case.separator_bin):
        errors.append("分隔页纸张来源不能包含逗号或换行")
    if case.separator_bin and case.separator_page.lower() == "none":
        errors.append("未启用分隔页时不能指定“分隔页纸张来源”（该字段在分隔页=无 时为灰化状态）")
    if case.binding_location.lower() not in BINDING_LOCATIONS:
        errors.append("装订位置必须为 none / left / right / top / bottom / longedge / shortedge")
    if case.binding_method.lower() not in BINDING_METHODS:
        errors.append("装订方式必须为 none / staple / saddle / punch / glue")
    if case.binding_location.lower() in {"longedge", "shortedge"} and case.duplex.lower() == "simplex":
        errors.append("装订位置为长边/短边时需要双面打印；请将单双面改为 duplexlong / duplexshort（骑马钉等折页装订同样需要双面）")
    if not 10 <= case.timeout_seconds <= 600:
        errors.append("超时秒数必须在 10–600 之间")
    return errors
