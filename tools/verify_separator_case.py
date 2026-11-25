"""真实驱动验证：用例“分隔页”列（无/作业之间/份数之间）是否真的写进驱动。

走的是应用自己的路径 driver_ui.apply_case（不是手工点界面）：
1. 用例 分隔页=作业之间 + 纸盒=纸盒1 → 设置 → 重开属性页回读“输出/分隔页/分隔页纸张来源”；
2. 分隔页=无 → 恢复后“输出”回到“正常”、来源回到原值；
3. 分隔页=作业之间 但纸盒留空 → 必须在打开属性页之前就报错。

只开合驱动属性页，不打印、不出纸。
用法：.venv\\Scripts\\python.exe tools\\verify_separator_case.py
"""
from __future__ import annotations

import sys
import time
from dataclasses import replace
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.driver_ui import (  # noqa: E402
    DriverUiError,
    _combos_ready,
    _desktop,
    _switch_tab,
    apply_case,
    close_stale_preferences,
    open_preferences,
)
from printcase.excel_reader import read_cases  # noqa: E402
from printcase.extended_attrs import ExtendedAttr  # noqa: E402
from printcase.ui_print import (  # noqa: E402
    _combo_current_text,
    _descendants,
    _label_before,
    _text_of,
)

PRINTER = "Pantum M9705DN Plus-3700 Series"
checks: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    checks.append((name, bool(ok), detail))
    print(f"  [{'通过' if ok else '未通过'}] {name}" + (f" — {detail}" if detail else ""), flush=True)


def read_paper_page(desktop) -> dict:
    """读纸张页：主纸张来源、输出单选、分隔页与分隔页纸张来源。"""
    close_stale_preferences(PRINTER, desktop=desktop)
    time.sleep(1.5)          # 驱动点“确定”后异步落盘，读太早会拿到旧值
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=lambda _t: None)
    state = {
        "source": "", "output_mode": "", "separator": "",
        "separator_enabled": None, "separator_source": "", "separator_source_enabled": None,
    }
    try:
        _switch_tab(window, "纸张", lambda _t: None)
        time.sleep(0.8)
        for combo in _combos_ready(window, 6, timeout=6):
            label = str(_label_before(combo)).strip()
            if label == "来源":
                state["source"] = _combo_current_text(combo)
            elif label == "分隔页":
                state["separator_enabled"] = bool(combo.is_enabled())
                state["separator"] = _combo_current_text(combo)
            elif label == "纸张来源":
                state["separator_source_enabled"] = bool(combo.is_enabled())
                state["separator_source"] = _combo_current_text(combo)
        for radio in _descendants(window, control_type="RadioButton"):
            text = _text_of(radio).strip()
            try:
                if text and radio.is_selected():
                    state["output_mode"] = text
            except Exception:  # noqa: BLE001
                pass
    finally:
        try:
            window.close()
        except Exception:  # noqa: BLE001
            pass
        close_stale_preferences(PRINTER, desktop=desktop)
        time.sleep(0.4)
    return state


def main() -> int:
    desktop = _desktop()
    cases, _warnings = read_cases(ROOT / "templates" / "PrintCase_Template.xlsx")
    base = next(case for case in cases if case.case_id == "TC-010")

    print("== 0. 当前状态 ==")
    before = read_paper_page(desktop)
    print(f"   文章来源={before['source']!r} 输出={before['output_mode']!r} "
          f"分隔页={before['separator']!r} 分隔页来源={before['separator_source']!r}")

    # 注意：本机账户没有写回打印机默认 DEVMODE 的权限，apply_case 结束时的 restore()
    # 会失败（日志里会有提示）。所以每一步都不依赖“还原”，而是显式把驱动设成期望值。

    print("== 1. 起点：分隔页=无 + 纸盒=纸盒1 → 输出必须是“正常” ==")
    apply_case(PRINTER, replace(base, separator_page="none"), caps={"color": False},
               log=lambda t: print("   ·", t))
    time.sleep(1.5)
    state = read_paper_page(desktop)
    print(f"   回读：来源={state['source']!r} 输出={state['output_mode']!r} 分隔页={state['separator']!r}")
    check("分隔页=无 → 输出=正常", state["output_mode"] == "正常", state["output_mode"])

    print("== 2. 分隔页=作业之间 + 纸盒=纸盒1 + 分隔页纸张来源=纸盒2 ==")
    case = replace(base, separator_page="job", bin="纸盒1", separator_bin="纸盒2")
    try:
        apply_case(PRINTER, case, caps={"color": False}, log=lambda t: print("   ·", t))
        check("apply_case 接受分隔页=作业之间（真实驱动）", True)
    except DriverUiError as exc:
        check("apply_case 接受分隔页=作业之间（真实驱动）", False, str(exc))
    time.sleep(1.5)
    after = read_paper_page(desktop)
    print(f"   回读：来源={after['source']!r} 输出={after['output_mode']!r} "
          f"分隔页={after['separator']!r} 分隔页来源={after['separator_source']!r}")
    check("纸张来源=纸盒1", after["source"] == "纸盒1", after["source"])
    check("输出=分隔页（单选已被切过去）", after["output_mode"] == "分隔页", after["output_mode"])
    check("分隔页=作业之间（下拉已解禁并生效）", after["separator"] == "作业之间", after["separator"])
    check("分隔页纸张来源=纸盒2", after["separator_source"] == "纸盒2", after["separator_source"])

    print("== 3. 再切回 分隔页=无：输出必须回到正常（证明能关掉，不是只进不出） ==")
    apply_case(PRINTER, replace(base, separator_page="none"), caps={"color": False},
               log=lambda t: print("   ·", t))
    time.sleep(1.5)
    back = read_paper_page(desktop)
    print(f"   回读：输出={back['output_mode']!r} 分隔页={back['separator']!r}")
    check("分隔页=无 → 输出=正常（可关）", back["output_mode"] == "正常", back["output_mode"])
    check("输出正常时 分隔页下拉回到灰化", back["separator_enabled"] is False, str(back["separator_enabled"]))

    print("== 4. 分隔页=作业之间 但纸盒留空：必须在开窗口之前报错 ==")
    started = time.monotonic()
    try:
        apply_case(PRINTER, replace(base, separator_page="job", bin="", separator_bin=""),
                   caps={"color": False}, log=lambda _t: None)
        check("纸盒留空时拒绝下发", False, "竟然没报错")
    except DriverUiError as exc:
        elapsed = time.monotonic() - started
        check("纸盒留空时拒绝下发（且不空等）", "具体纸盒" in str(exc) and elapsed < 3.0,
              f"{elapsed:.2f}s，{str(exc)[:46]}")

    print("== 5. 收尾：纸张来源回到自动选择、输出回到正常 ==")
    apply_case(
        PRINTER, replace(base, bin="", separator_page="none",
                         extended_attrs=[ExtendedAttr("driver_paper", "纸张来源", "自动选择")]),
        caps={"color": False}, log=lambda t: print("   ·", t),
    )
    time.sleep(1.5)
    tail = read_paper_page(desktop)
    print(f"   回读：来源={tail['source']!r} 输出={tail['output_mode']!r} 分隔页={tail['separator']!r}")
    check("纸张来源回到自动选择", tail["source"] == "自动选择", tail["source"])
    check("输出仍为正常", tail["output_mode"] == "正常", tail["output_mode"])

    passed = sum(1 for _n, ok, _d in checks if ok)
    print(f"\n结果：{passed}/{len(checks)} 项通过")
    for name, ok, detail in checks:
        if not ok:
            print(f"  未通过：{name} — {detail}")
    return 0 if passed == len(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
