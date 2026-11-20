from __future__ import annotations

import platform
import re
from typing import Callable

from .models import PrintCase
from .validator import normalize_paper

DMCOLLATE_TRUE = 1
DMCOLLATE_FALSE = 0

# 常用纸张的 DMPAPER_* 代号；驱动上报 DC_PAPERS 时优先用驱动返回值
FALLBACK_PAPER_CODES = {
    "letter": 1, "legal": 5, "statement": 6, "executive": 7,
    "a3": 8, "a4": 9, "a5": 11, "folio": 14, "ledger": 15,
    "jisb4": 12, "jisb5": 13, "jisb6": 88, "a6": 70,
}

# 标准 DMMEDIA_* 代号（普通纸/铜版纸/胶片），其余介质依赖驱动自定义代号
FALLBACK_MEDIA_CODES = {"普通纸": 1, "铜版纸": 2, "胶片": 3}


class DriverSettingsError(RuntimeError):
    """写入打印机驱动 DEVMODE 失败（权限不足、字段不受支持等）。"""


def _flag(win32module, name: str) -> int:
    """取 DEVMODE 相关常量。

    pywin32 的 win32print 不导出任何 DM_* 常量（只有 win32con 有），
    因此按 win32print → win32con → 内置值 的顺序解析。
    """
    flag = getattr(win32module, name, None)
    if flag is None:
        try:
            import win32con  # noqa: WPS433

            flag = getattr(win32con, name, None)
        except ImportError:
            flag = None
    if flag is None:
        flag = _FALLBACK_FLAGS.get(name)
    if flag is None:
        raise DriverSettingsError(f"当前环境缺少 DEVMODE 标志 {name}，无法设置驱动属性。")
    return int(flag)


_FALLBACK_FLAGS = {
    "DM_IN_BUFFER": 0x8, "DM_OUT_BUFFER": 0x2,
    "DM_ORIENTATION": 0x1, "DM_PAPERSIZE": 0x2, "DM_PAPERLENGTH": 0x4, "DM_PAPERWIDTH": 0x8,
    "DM_COPIES": 0x100, "DM_DEFAULTSOURCE": 0x200, "DM_PRINTQUALITY": 0x400,
    "DM_COLOR": 0x800, "DM_DUPLEX": 0x1000, "DM_COLLATE": 0x8000, "DM_MEDIATYPE": 0x02000000,
}


def _norm(value: str) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", str(value or "").casefold())


def _resolve_from_caps(token: str, names: list[str] | None, ids: list[int] | None) -> int | None:
    """在驱动上报的名称/代号列表里按名称双向匹配，返回对应代号。

    驱动上报的名称通常不含克重后缀（如“厚纸2”），而用例里填的是完整标签
    （如“厚纸2(158-209gsm)”），因此需要双向包含匹配。
    """
    if not names or not ids or len(names) != len(ids):
        return None
    wanted = _norm(token)
    if not wanted:
        return None
    for name, code in zip(names, ids):
        normalized = _norm(name)
        if not normalized:
            continue
        if normalized == wanted or wanted in normalized or normalized in wanted:
            return int(code)
    return None


def devmode_updates_for_case(
    case: PrintCase,
    caps: dict | None = None,
    errors: list[str] | None = None,
) -> list[tuple[str, str, int]]:
    """把用例属性翻译成 DEVMODE 字段更新。

    覆盖直接打印通道无法从文档携带的属性：逐份/非逐份、打印质量、纸张类型、
    纸张尺寸（全部尺寸，含 A4 等标准尺寸与 SRA3/8K 等驱动专有尺寸）、
    纸张来源（名称类纸盒）。
    任何“用户显式填写但无法保证生效”的属性都会写入 errors（而非静默跳过），
    调用方必须据此拦截用例，避免属性悄悄失效。
    """
    errors = errors if errors is not None else []
    updates: list[tuple[str, str, int]] = []
    caps = caps or {}

    if case.collate:
        updates.append(("Collate", "DM_COLLATE", DMCOLLATE_TRUE if case.collate == "yes" else DMCOLLATE_FALSE))

    quality = str(case.print_quality or "auto").lower()
    if quality and quality != "auto":
        quality_map = {"high": -4, "medium": -3, "low": -2, "draft": -1}
        if quality in quality_map:
            updates.append(("PrintQuality", "DM_PRINTQUALITY", quality_map[quality]))
        elif quality.isdigit():
            updates.append(("PrintQuality", "DM_PRINTQUALITY", int(quality)))
        else:
            errors.append(f"打印质量“{case.print_quality}”无法识别（可用 600/1200 或 DPI 数字）")

    media = str(case.media_type or "").strip()
    if media:
        if media.isdigit():
            updates.append(("MediaType", "DM_MEDIATYPE", int(media)))
        else:
            # 标准 DMMEDIA_* 代号优先：多数驱动只接受 1/2/3 这类标准值，
            # 而 DC_MEDIATYPENAMES 上报的“代号”常是厂商私有不透明 DWORD。
            code = None
            for key, fallback in FALLBACK_MEDIA_CODES.items():
                if key in media:
                    code = fallback
                    break
            if code is None:
                candidate = _resolve_from_caps(media, caps.get("media_names"), caps.get("media_ids"))
                if candidate is not None and 0 < candidate <= 0xFFFF:
                    code = candidate
            if code is None:
                errors.append(
                    f"纸张类型“{media}”为驱动私有介质，标准 DEVMODE 无法可靠下发；"
                    "请在“UI 自动化”模式下由驱动界面设置，或改填标准介质（普通纸/铜版纸/胶片）。"
                )
            else:
                updates.append(("MediaType", "DM_MEDIATYPE", code))

    paper_token = "mixed" if case.mixed_sizes.lower() == "yes" else normalize_paper(case.paper)
    # 所有显式纸张尺寸都通过 DEVMODE 精确写入驱动（标准尺寸用内置代号表，
    # 驱动专有尺寸优先从驱动上报的 DC_PAPERS 解析），保证属性真实生效。
    if paper_token and paper_token not in {"", "auto", "mixed"}:
        code = _resolve_from_caps(paper_token, caps.get("papers"), caps.get("paper_ids"))
        if code is None:
            code = FALLBACK_PAPER_CODES.get(paper_token)
        if code is None:
            errors.append(f"纸张尺寸 {case.paper} 无法在驱动中生效（驱动未上报该尺寸代号）。请确认驱动支持该尺寸，或在驱动中配置自定义纸张。")
        else:
            updates.append(("PaperSize", "DM_PAPERSIZE", code))

    bin_value = str(case.bin or "").strip()
    if bin_value and not bin_value.isdigit():
        code = _resolve_from_caps(bin_value, caps.get("bin_names"), caps.get("bin_ids"))
        if code is None:
            errors.append(f"纸盒“{bin_value}”无法在驱动中生效（驱动未上报该纸盒编号）。请改填纸盒编号数字。")
        else:
            updates.append(("DefaultSource", "DM_DEFAULTSOURCE", code))

    return updates


def _import_win32(win32module=None):
    if win32module is not None:
        return win32module
    if platform.system() != "Windows":
        raise DriverSettingsError("驱动 DEVMODE 设置仅支持 Windows。")
    try:
        import win32print
    except ImportError as exc:
        raise DriverSettingsError("缺少 pywin32，无法写入驱动 DEVMODE。") from exc
    return win32print


def _read_default_devmode(w, handle, printer_name: str):
    """读取驱动当前默认 DEVMODE（PyDEVMODEW 副本）。"""
    info = w.GetPrinter(handle, 2)
    dm = info.get("pDevMode")
    if dm is None:
        raise DriverSettingsError(f"打印机 {printer_name} 的驱动未提供可修改的 DEVMODE。")
    return dm, info


def _merge_devmode(w, handle, printer_name: str, dm, flags: int):
    """交给驱动校验/合并 DEVMODE（DocumentProperties 真实签名为 6 参数）。"""
    return w.DocumentProperties(None, handle, printer_name, dm, dm, flags)


def _write_devmode(w, printer_name: str, changes: list[tuple[str, str, int]]) -> None:
    """读取驱动当前默认 DEVMODE → 覆盖指定字段 → 驱动校验 → 写回打印机默认设置。"""
    handle = w.OpenPrinter(printer_name, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
    try:
        dm, info = _read_default_devmode(w, handle, printer_name)
        for attr, flag_name, value in changes:
            flag = _flag(w, flag_name)
            setattr(dm, attr, value)
            dm.Fields = int(dm.Fields) | flag
        # 驱动校验并合并：驱动可拒绝/调整不支持的组合
        _merge_devmode(w, handle, printer_name, dm, _flag(w, "DM_IN_BUFFER") | _flag(w, "DM_OUT_BUFFER"))
        for attr, flag_name, value in changes:
            flag = _flag(w, flag_name)
            if not (int(dm.Fields) & flag):
                raise DriverSettingsError(f"驱动不接受 DEVMODE 字段 {attr}（dmFields 未包含 {flag_name}）。")
        info["pDevMode"] = dm
        w.SetPrinter(handle, 2, info, 0)
    finally:
        w.ClosePrinter(handle)


def apply_case_to_driver(
    printer_name: str,
    case: PrintCase,
    log: Callable[[str], None] | None = None,
    win32module=None,
    caps: dict | None = None,
) -> Callable[[], None]:
    """把用例中命令行无法传递的属性写入驱动默认 DEVMODE，返回恢复函数。

    严格模式：任何显式填写但无法解析/写入/回读确认的属性都会抛出
    DriverSettingsError，由调用方拦截用例——保证不会出现“属性静默失效
    但作业照常下发”的情况。
    注意：这是打印机的全局默认设置，写入后到恢复前的窗口内，本机其他
    程序提交的打印作业也会受影响；本软件串行执行并在用例结束后立即恢复。
    """
    w = _import_win32(win32module)
    errors: list[str] = []
    updates = devmode_updates_for_case(case, caps=caps, errors=errors)
    if errors:
        raise DriverSettingsError("；".join(errors))
    if not updates:
        return lambda: None

    handle = w.OpenPrinter(printer_name, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
    try:
        dm, _info = _read_default_devmode(w, handle, printer_name)
        original_fields = int(dm.Fields)
        original_values = {attr: int(getattr(dm, attr)) for attr, _, _ in updates}
    finally:
        w.ClosePrinter(handle)

    _write_devmode(w, printer_name, updates)

    # 回读校验：重新读取驱动默认 DEVMODE，确认每个字段确实等于写入值。
    # 驱动可能接受字段标志却在内部改写数值，只有回读一致才算“生效有保证”。
    handle = w.OpenPrinter(printer_name, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
    try:
        dm, _info = _read_default_devmode(w, handle, printer_name)
        mismatches = [
            f"{attr}（期望 {value}，实际 {int(getattr(dm, attr))}）"
            for attr, _flag, value in updates
            if int(getattr(dm, attr)) != value
        ]
    finally:
        w.ClosePrinter(handle)
    if mismatches:
        # 介质类型常被驱动“归一化”为自己的私有编码（非 0 即为有效值）
        tolerated = []
        blocking = []
        for item in mismatches:
            if item.startswith("MediaType") and "实际 " in item:
                actual = item.rsplit("实际 ", 1)[-1].rstrip("）)")
                try:
                    if int(actual) > 0:
                        tolerated.append(item)
                        continue
                except ValueError:
                    pass
            blocking.append(item)
        if tolerated and log:
            log("驱动将介质类型归一化为其私有编码（" + "；".join(tolerated) + "），已按驱动值执行。")
        if blocking:
            raise DriverSettingsError(
                "驱动未接受以下设置（写入后回读不一致）：" + "、".join(blocking)
            )

    if log:
        summary = "、".join(f"{attr}={value}" for attr, _, value in updates)
        log(f"已写入驱动默认设置（DEVMODE）并回读确认：{summary}；用例执行结束后自动恢复。")

    def restore() -> None:
        changes = [(attr, flag_name, original_values[attr]) for attr, flag_name, _ in updates]
        handle = w.OpenPrinter(printer_name, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
        try:
            dm, info = _read_default_devmode(w, handle, printer_name)
            for attr, flag_name, value in changes:
                flag = _flag(w, flag_name)
                setattr(dm, attr, value)
                dm.Fields = int(dm.Fields) | flag
            dm.Fields = int(dm.Fields) | original_fields
            info["pDevMode"] = dm
            w.SetPrinter(handle, 2, info, 0)
        finally:
            w.ClosePrinter(handle)
        if log:
            log("驱动默认设置已恢复为用例执行前的状态。")

    return restore
