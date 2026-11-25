from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .models import PrintCase
from .extended_attrs import ExtendedAttr, normalize as normalize_ext_attr, parse_legacy_row
from .validator import normalize_paper


HEADER_ALIASES = {
    "case_id": {"case_id", "caseid", "用例id", "用例编号", "测试用例id", "编号"},
    "enabled": {"enabled", "enable", "是否启用", "启用", "执行"},
    "file_path": {"file_path", "filepath", "文件路径", "打印文件", "源文件", "文档路径"},
    "printer": {"printer", "printer_name", "打印机", "打印机名称", "目标打印机"},
    "pages": {"pages", "page_range", "页码", "打印页码", "页范围"},
    "copies": {"copies", "copy", "份数", "打印份数"},
    "paper": {"paper", "paper_size", "纸张", "纸张大小", "纸张尺寸"},
    "mixed_sizes": {"mixed_sizes", "mixed_size", "混合尺寸", "混合尺寸打印", "混合纸张"},
    "orientation": {"orientation", "方向", "纸张方向", "横纵向"},
    "color": {"color", "colour", "颜色", "彩色模式"},
    "duplex": {"duplex", "sides", "单双面", "单双面打印", "双面模式"},
    "scaling": {"scaling", "scale", "缩放", "适应页面"},
    "bin": {"bin", "tray", "paper_source", "进纸来源", "纸盒", "纸盘", "进纸器", "纸盒/进纸来源", "纸张来源"},
    "collate": {"collate", "collate_order", "print_order", "逐份", "打印顺序", "逐份/非逐份", "逐份打印"},
    "print_quality": {"print_quality", "dpi", "resolution", "打印质量", "分辨率"},
    "media_type": {"media_type", "paper_type", "纸型", "纸张类型", "纸张克重"},
    "skip_blank": {"skip_blank", "skip_blank_pages", "跳过空白页", "跳过空白"},
    "borderless": {"borderless", "borderless_print", "无边距", "无边距打印", "满版打印"},
    "separator_page": {"separator_page", "separator", "分隔页", "间隔页", "插入分隔页"},
    "separator_bin": {"separator_bin", "separator_paper_source", "分隔页纸张来源", "分隔页纸盒", "间隔页纸盒"},
    "binding_location": {"binding_location", "binding_edge", "装订位置", "装订边", "装订边位置"},
    "binding_method": {"binding_method", "binding", "staple", "装订方式", "装订"},
    "timeout_seconds": {"timeout_seconds", "timeout", "超时秒", "超时时间", "超时秒数"},
    "notes": {"notes", "备注", "说明", "预期结果", "备注/预期结果"},
}

# 中文单元格取值 → 内部标准值；未命中的值保持原样，交给校验器报错。
VALUE_ALIASES: dict[str, dict[str, str]] = {
    "paper": {"自动": "auto", "混合": "mixed", "混合尺寸": "mixed", "tabloid": "ledger"},
    "orientation": {"纵向": "portrait", "横向": "landscape"},
    "color": {"彩色": "color", "黑白": "monochrome", "灰度": "monochrome", "单色": "monochrome"},
    "duplex": {
        "单面": "simplex", "双面": "duplex",
        "长边双面": "duplexlong", "双面长边": "duplexlong", "长边": "duplexlong",
        "短边双面": "duplexshort", "双面短边": "duplexshort", "短边": "duplexshort",
    },
    "scaling": {"适应": "fit", "适应页面": "fit", "缩小": "shrink", "不缩放": "noscale", "原始大小": "noscale"},
    "borderless": {
        "是": "yes", "开启": "yes", "启用": "yes", "否": "no", "关闭": "no", "停用": "no", "无": "no",
        "Y": "yes", "y": "yes", "N": "no", "n": "no",
    },
    "collate": {"逐份": "yes", "是": "yes", "按份": "yes", "非逐份": "no", "逐页": "no", "否": "no", "按页": "no", "默认": ""},
    "mixed_sizes": {"是": "yes", "否": "no", "开": "yes", "关": "no", "Y": "yes", "N": "no"},
    "skip_blank": {"是": "yes", "开": "yes", "开启": "yes", "否": "no", "关": "no", "关闭": "no"},
    "print_quality": {"自动": "auto", "默认": "auto"},
    "separator_page": {
        "无": "none", "否": "none", "不插入": "none", "关闭": "none", "正常": "none",
        "作业之间": "job", "作业与作业之间": "job", "作业间": "job",
        "份数之间": "copies", "每份之间": "copies", "份间": "copies", "逐份之间": "copies",
    },
    "binding_location": {
        "无": "none", "不装订": "none",
        "左": "left", "左侧": "left", "右": "right", "右侧": "right",
        "上": "top", "顶部": "top", "下": "bottom", "底部": "bottom",
        "长边": "longedge", "短边": "shortedge",
    },
    "binding_method": {
        "无": "none", "不装订": "none",
        "订书钉": "staple", "钉装": "staple", "staple": "staple",
        "骑马钉": "saddle", "骑马订": "saddle", "折页装订": "saddle",
        "打孔": "punch", "胶装": "glue",
    },
}


def _norm(value: Any) -> str:
    return str(value or "").strip().lower().replace(" ", "").replace("\n", "")


def _normalize_bin(value: Any) -> str:
    """纸盒“自动选择”等价于不指定纸盒，交给驱动自动判断。"""
    text = str(value or "").strip()
    return "" if text in {"自动选择", "自动", "auto"} else text


def _map_choice(field: str, value: Any) -> str:
    """按字段取值别名表归一化；未命中时保留原始小写文本，交给校验器处理。"""
    text = str(value or "").strip()
    if not text:
        return ""
    mapping = VALUE_ALIASES.get(field, {})
    if text in mapping:
        return mapping[text]
    return text.lower()


def _parse_bool(value: Any, default: bool = True) -> tuple[bool, str | None]:
    if value is None or str(value).strip() == "":
        return default, None
    if isinstance(value, bool):
        return value, None
    key = _norm(value)
    if key in {"1", "true", "yes", "y", "是", "启用", "执行", "√"}:
        return True, None
    if key in {"0", "false", "no", "n", "否", "停用", "不执行", "×"}:
        return False, None
    return default, f"启用字段值无法识别：{value!r}；请填写 TRUE/FALSE 或 是/否。"


def _parse_int(value: Any, default: int, label: str) -> tuple[int, str | None]:
    if value is None or str(value).strip() == "":
        return default, None
    try:
        numeric = float(value)
        if not numeric.is_integer():
            raise ValueError
        return int(numeric), None
    except (TypeError, ValueError):
        return default, f"{label}不是有效整数：{value!r}。"


EXT_SHEET_NAMES = ("驱动扩展属性", "扩展属性", "driver_attrs")
EXT_HEADER_ALIASES = {
    "case_id": {"用例编号", "用例id", "case_id", "编号"},
    "page": {"页面", "驱动页面", "page", "页签"},
    "label": {"属性", "属性名", "参数", "label", "attribute"},
    "value": {"值", "取值", "value", "设置值"},
}


def _read_extended_sheet(wb) -> tuple[dict[str, list[ExtendedAttr]], list[str]]:
    """读取“驱动扩展属性”长表（用例编号 / 页面 / 属性 / 值）。"""
    sheet_name = next((name for name in EXT_SHEET_NAMES if name in wb.sheetnames), None)
    if sheet_name is None:
        return {}, []
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return {}, []
    header_map: dict[str, int] = {}
    for index, raw in enumerate(header):
        key = _norm(raw)
        for canonical, aliases in EXT_HEADER_ALIASES.items():
            if key in {_norm(alias) for alias in aliases}:
                header_map[canonical] = index
                break

    warnings: list[str] = []
    result: dict[str, list[ExtendedAttr]] = {}
    if not {"case_id", "label"} <= set(header_map):
        warnings.append(f"“{sheet_name}”表缺少‘用例编号’或‘属性’列，扩展属性未导入。")
        return result, warnings

    for excel_row, row in enumerate(rows, start=2):
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue

        def cell(name: str, default: Any = "") -> Any:
            index = header_map.get(name)
            return row[index] if index is not None and index < len(row) and row[index] is not None else default

        case_id = str(cell("case_id")).strip()
        label = str(cell("label")).strip()
        if not case_id or not label:
            continue
        attr = normalize_ext_attr(str(cell("page")).strip(), label, cell("value"))
        result.setdefault(case_id, []).append(attr)
    return result, warnings


def _looks_like_legacy_workbook(wb) -> bool:
    """判断是否是既有 print_panel 项目的工作簿（sheet “print” + &属性=值 文本块）。"""
    if "print" not in wb.sheetnames:
        return False
    ws = wb["print"]
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        text = " ".join(str(value or "") for value in row)
        return "&" in text or "参数" in text or "测试次数" in text
    return False


def _read_legacy_print_sheet(wb, workbook_path: Path) -> tuple[list[PrintCase], list[str]]:
    """读取既有 print_panel 工作簿（sheet “print”）：列 3–11 为 &属性名=值 文本块。"""
    ws = wb["print"]
    rows = ws.iter_rows(values_only=True)
    try:
        next(rows)  # 表头
    except StopIteration:
        raise ValueError("print 工作表为空。")

    cases: list[PrintCase] = []
    warnings: list[str] = ["已按“既有 print_panel 工作表”格式导入；扩展属性需在“驱动扩展属性”表中维护。"]
    for excel_row, row in enumerate(rows, start=2):
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        case_id = str(row[0] or "").strip() or f"ROW-{excel_row}"
        title = str(row[1] or "").strip()
        file_ref = str(row[15] if len(row) > 15 and row[15] else "").strip()
        if not file_ref:
            warnings.append(f"{case_id}：旧表未填“打印文档路径”，请在用例表补充。")
        resolved = None
        if file_ref:
            candidate = Path(file_ref).expanduser()
            resolved = candidate if candidate.is_absolute() else (workbook_path.parent / candidate).resolve()
        attrs = parse_legacy_row(row, case_id)
        cases.append(PrintCase(
            case_id=case_id,
            enabled=True,
            file_path=file_ref,
            notes=title,
            source_row=excel_row,
            resolved_path=resolved,
            extended_attrs=attrs,
        ))
    if not cases:
        raise ValueError("print 工作表中没有识别到用例数据。")
    return cases, warnings


def read_cases(file_path: str | Path, sheet_name: str | None = None) -> tuple[list[PrintCase], list[str]]:
    """Read a case manifest. Relative document paths resolve against the workbook folder."""
    workbook_path = Path(file_path).expanduser().resolve()
    if workbook_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("仅支持 .xlsx / .xlsm 用例清单；旧版 .xls 请先另存为 .xlsx。")
    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到 Excel 文件：{workbook_path}")

    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        # 兼容既有 print_panel 项目的工作簿：sheet 名为 print、属性塞在 &属性名=值 文本块里
        if not sheet_name and _looks_like_legacy_workbook(wb):
            return _read_legacy_print_sheet(wb, workbook_path)

        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"工作表不存在：{sheet_name}")
            ws = wb[sheet_name]
        elif "用例" in wb.sheetnames:
            ws = wb["用例"]
        else:
            ws = wb[wb.sheetnames[0]]

        rows = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration:
            raise ValueError("Excel 工作表为空。")

        header_map: dict[str, int] = {}
        for idx, raw in enumerate(header_row):
            key = _norm(raw)
            for canonical, aliases in HEADER_ALIASES.items():
                if key in aliases:
                    header_map[canonical] = idx
                    break

        missing = [key for key in ("case_id", "file_path") if key not in header_map]
        if missing:
            raise ValueError("缺少必需列：" + ", ".join(missing) + "。请使用随软件提供的 Excel 模板。")

        cases: list[PrintCase] = []
        warnings: list[str] = []
        for excel_row, row in enumerate(rows, start=2):
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue

            def get(name: str, default: Any = "") -> Any:
                ix = header_map.get(name)
                return row[ix] if ix is not None and ix < len(row) and row[ix] is not None else default

            case_id = str(get("case_id", "")).strip()
            file_ref = str(get("file_path", "")).strip()
            if not case_id and not file_ref:
                continue
            if not case_id:
                case_id = f"ROW-{excel_row}"
                warnings.append(f"第 {excel_row} 行没有用例 ID，已临时使用 {case_id}。")

            enabled, enabled_error = _parse_bool(get("enabled", True), True)
            # 禁用的用例不会下发（如打印机面板生成的内部页，本就没有文件），
            # 对其报告“文件路径为空”是误导性警告；只对启用且缺文件的用例告警。
            if not file_ref and enabled:
                warnings.append(f"{case_id}：文件路径为空。")
            copies, copies_error = _parse_int(get("copies", 1), 1, "份数")
            timeout_seconds, timeout_error = _parse_int(get("timeout_seconds", 90), 90, "超时秒数")
            input_errors = [e for e in (enabled_error, copies_error, timeout_error) if e]
            paper_raw = str(get("paper", "A4")).strip() or "A4"
            mixed_raw = _map_choice("mixed_sizes", get("mixed_sizes", "no")) or "no"
            if mixed_raw == "yes":
                paper_raw = "mixed"
            cases.append(PrintCase(
                case_id=case_id,
                enabled=enabled,
                file_path=file_ref,
                printer=str(get("printer", "")).strip(),
                pages=str(get("pages", "")).strip(),
                copies=copies,
                paper=normalize_paper(_map_choice("paper", paper_raw)) or "a4",
                mixed_sizes=_map_choice("mixed_sizes", get("mixed_sizes", "no")) or "no",
                orientation=_map_choice("orientation", get("orientation", "portrait")) or "portrait",
                color=_map_choice("color", get("color", "color")) or "color",
                duplex=_map_choice("duplex", get("duplex", "simplex")) or "simplex",
                scaling=_map_choice("scaling", get("scaling", "fit")) or "fit",
                bin=_normalize_bin(get("bin", "")),
                collate=_map_choice("collate", get("collate", "")),
                print_quality=_map_choice("print_quality", get("print_quality", "auto")) or "auto",
                media_type=str(get("media_type", "")).strip(),
                skip_blank=_map_choice("skip_blank", get("skip_blank", "no")) or "no",
                borderless=_map_choice("borderless", get("borderless", "no")) or "no",
                separator_page=_map_choice("separator_page", get("separator_page", "none")) or "none",
                separator_bin=_normalize_bin(get("separator_bin", "")),
                binding_location=_map_choice("binding_location", get("binding_location", "none")) or "none",
                binding_method=_map_choice("binding_method", get("binding_method", "none")) or "none",
                timeout_seconds=timeout_seconds,
                notes=str(get("notes", "")).strip(),
                source_row=excel_row,
                resolved_path=(workbook_path.parent / file_ref).resolve() if file_ref and not Path(file_ref).is_absolute() else Path(file_ref).expanduser() if file_ref else None,
                input_errors=input_errors,
            ))
        # 驱动扩展属性（长表：用例编号 / 页面 / 属性 / 值）
        ext_map, ext_warnings = _read_extended_sheet(wb)
        warnings.extend(ext_warnings)
        if ext_map:
            for case in cases:
                case.extended_attrs = ext_map.get(case.case_id, [])

        if not cases:
            raise ValueError("没有识别到任何用例数据。请检查表头和数据行。")
        return cases, warnings
    finally:
        wb.close()
