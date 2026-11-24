"""真实驱动验证：纸张来源=纸盒1 后能否设置“分隔页”，以及“分隔页=无”能否切回正常。

只开合驱动属性页，不打印、不出纸。
用法：.venv\\Scripts\\python.exe tools\\verify_separator_attr.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase.driver_ui import (  # noqa: E402
    DriverUiError,
    _combos_ready,
    _desktop,
    _switch_tab,
    apply_case,
    close_stale_preferences,
    open_preferences,
)
from printcase.extended_attrs import ExtendedAttr  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.ui_print import _label_before, _text_of  # noqa: E402

PRINTER = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else "Pantum M9705DN Plus-3700 Series"
checks: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    checks.append((name, ok, detail))
    print(f"  [{'通过' if ok else '未通过'}] {name}" + (f" — {detail}" if detail else ""))


def read_paper_page(desktop) -> dict:
    """读纸张页：来源、分隔页下拉是否可用与当前值。"""
    close_stale_preferences(PRINTER, desktop=desktop)
    time.sleep(0.5)
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=lambda _t: None)
    state = {"source": "", "separator": "", "separator_enabled": None}
    try:
        _switch_tab(window, "纸张", lambda _t: None)
        time.sleep(0.8)
        for index, combo in enumerate(_combos_ready(window, 3, timeout=6)):
            label = str(_label_before(combo)).strip()
            if label == "来源":
                state["source"] = str(combo.selected_text() or "").strip()
            elif label == "分隔页":
                state["separator_enabled"] = bool(combo.is_enabled())
                try:
                    state["separator"] = str(combo.selected_text() or "").strip()
                except Exception:  # noqa: BLE001
                    state["separator"] = ""
    finally:
        try:
            window.close()
        except Exception:  # noqa: BLE001
            pass
        close_stale_preferences(PRINTER, desktop=desktop)
        time.sleep(0.4)
    return state


def case_with(bin_value: str, separator: str) -> PrintCase:
    case = PrintCase(
        case_id="TC-SEP", enabled=True, file_path="x.docx",
        paper="A4 (210 x 297mm)", bin=bin_value,
        extended_attrs=[ExtendedAttr("driver_paper", "分隔页", separator)],
    )
    return case


def main() -> int:
    desktop = _desktop()
    print("== 0. 当前状态 ==")
    before = read_paper_page(desktop)
    print(f"   来源={before['source']!r} 分隔页={before['separator']!r} 可用={before['separator_enabled']}")

    print("== 1. 纸盒=纸盒1 + 分隔页=作业之间 ==")
    logs: list[str] = []
    try:
        restore = apply_case(PRINTER, case_with("纸盒1", "作业之间"), caps={"color": False},
                             desktop=desktop, log=lambda t: (logs.append(t), print("   ·", t))[0])
    except DriverUiError as exc:
        check("设置分隔页", False, str(exc))
        restore = None
    else:
        check("设置分隔页", True)
    state = read_paper_page(desktop)
    print(f"   回读：来源={state['source']!r} 分隔页={state['separator']!r} 可用={state['separator_enabled']}")
    check("纸张来源已切到纸盒1", state["source"] == "纸盒1", state["source"])
    check("分隔页下拉已解禁", state["separator_enabled"] is True)
    check("分隔页=作业之间", state["separator"] == "作业之间", state["separator"])

    print("== 2. 分隔页=无（输出切回正常） ==")
    try:
        apply_case(PRINTER, case_with("纸盒1", "无"), caps={"color": False},
                   desktop=desktop, log=lambda t: print("   ·", t))
    except DriverUiError as exc:
        check("关闭分隔页", False, str(exc))
    state = read_paper_page(desktop)
    print(f"   回读：来源={state['source']!r} 分隔页={state['separator']!r} 可用={state['separator_enabled']}")
    check("关闭后分隔页下拉重新灰化", state["separator_enabled"] is False)

    print("== 3. 恢复原状 ==")
    try:
        apply_case(
            PRINTER,
            PrintCase(case_id="TC-RESTORE", enabled=True, file_path="x.docx",
                      paper="A4 (210 x 297mm)", bin=before["source"] or "自动选择"),
            caps={"color": False}, desktop=desktop, log=lambda t: print("   ·", t),
        )
    except DriverUiError as exc:
        print(f"   恢复失败：{exc}")
    final = read_paper_page(desktop)
    print(f"   最终：来源={final['source']!r} 分隔页={final['separator']!r} 可用={final['separator_enabled']}")

    failed = [name for name, ok, _ in checks if not ok]
    print("\n结果：" + ("全部通过" if not failed else f"未通过 {len(failed)} 项：{'、'.join(failed)}"))
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
