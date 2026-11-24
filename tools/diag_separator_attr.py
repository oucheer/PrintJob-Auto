"""诊断：驱动“纸张”页里的“分隔页”下拉到底有哪些项、设了是否真的保存。

只开合驱动属性页（会临时改一次“分隔页”，结束前恢复原值），不打印任何东西。

用法：
    .venv\\Scripts\\python.exe tools\\diag_separator_attr.py [打印机名] [目标值]
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase.driver_ui import (  # noqa: E402
    _combos_ready,
    _confirm,
    _desktop,
    _select_combo_index,
    _switch_tab,
    _visible_combos,
    close_stale_preferences,
    open_preferences,
    snapshot,
)
from printcase.ui_print import (  # noqa: E402
    _combo_current_text,
    _combo_items,
    _descendants,
    _label_before,
    _norm,
    _text_of,
)

PRINTER = sys.argv[1] if len(sys.argv) > 1 else "Pantum M9705DN Plus-3700 Series"
TARGET = sys.argv[2] if len(sys.argv) > 2 else "份数之间"


def dump_combos(window, log):
    combos = _visible_combos(window)
    log(f"纸张页可见下拉框 {len(combos)} 个：")
    for index, combo in enumerate(combos):
        name = _text_of(combo)
        label = _label_before(combo)
        try:
            current = combo.selected_text()
        except Exception as exc:  # noqa: BLE001
            current = f"<读不到: {type(exc).__name__}>"
        try:
            items = [str(item) for item in (combo.texts() or [])]
        except Exception as exc:  # noqa: BLE001
            items = [f"<读不到: {type(exc).__name__}>"]
        log(f"  [{index}] UIA名={name!r} 前置标签={label!r} 当前={current!r} 选项={items}")
    return combos


def main() -> int:
    log = lambda text: print("   ·", text)  # noqa: E731
    desktop = _desktop()

    print("== 0. 打印服务侧的分隔页文件（pSepFile） ==")
    try:
        info = snapshot(PRINTER)
        raw = info["info"]
        print(f"   pSepFile={raw.get('pSepFile')!r} pDatatype={raw.get('pDatatype')!r}")
        print(f"   DEVMODE 里没有“分隔页”字段，这点下面会验证")
    except Exception as exc:  # noqa: BLE001
        print(f"   读取失败：{exc}")

    print("== 1. 打开“纸张”页并枚举下拉项 ==")
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=log)
    try:
        if not _switch_tab(window, "纸张", log):
            print("   切不到“纸张”页")
            return 1
        combos = _combos_ready(window, 3, timeout=6)
        combos = dump_combos(window, log)

        print("== 2. 页面上的勾选框 / 文本 ==")
        for check in _descendants(window, control_type="CheckBox"):
            if _text_of(check).strip():
                print(f"   CheckBox: {_text_of(check)!r}")
        for text in _descendants(window, control_type="Text"):
            value = _text_of(text).strip()
            if value and ("分隔" in value or "插页" in value):
                print(f"   Text: {value!r}")
    finally:
        try:
            window.close()
        except Exception:  # noqa: BLE001
            pass

    print("== 3. 尝试设置“分隔页”并回读 ==")
    close_stale_preferences(PRINTER, desktop=desktop)
    time.sleep(0.6)
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=log)
    original = None
    try:
        _switch_tab(window, "纸张", log)
        combos = _combos_ready(window, 3, timeout=6)
        # 找“前置标签正好是分隔页”的那个（避免匹配到“分隔页（OHP插页）纸张来源”）
        target_combo = None
        for index, combo in enumerate(combos):
            label = _norm(_label_before(combo))
            name = _norm(_text_of(combo))
            if label == _norm("分隔页") or name == _norm("分隔页"):
                target_combo = combo
                print(f"   命中“分隔页”下拉：[{index}] 标签={_label_before(combo)!r}")
                break
        if target_combo is None:
            print("   没有前置标签正好等于“分隔页”的下拉；按“包含”匹配会命中的是：")
            for index, combo in enumerate(combos):
                if "分隔" in _norm(_label_before(combo)) or "分隔" in _norm(_text_of(combo)):
                    print(f"     [{index}] 标签={_label_before(combo)!r} 名={_text_of(combo)!r}")
            target_combo = combos[0] if combos else None
        if target_combo is None:
            return 1
        try:
            enabled = bool(target_combo.is_enabled())
        except Exception:  # noqa: BLE001
            enabled = True
        original = _combo_current_text(target_combo)
        if not enabled:
            source_text = ""
            for combo in combos:
                if str(_label_before(combo)).strip() == "来源":
                    source_text = _combo_current_text(combo)
            print(f"   分隔页下拉当前是灰化状态（来源={source_text!r}）。")
            print("   驱动规则：纸张来源必须是具体纸盒，并先选中“输出→分隔页”，该项才可选。")
            print("   要验证“能否设置”请运行 tools/verify_separator_attr.py。")
            return 0
        print(f"   设置前：{original!r}")
        items = _combo_items(target_combo)
        print(f"   可选项：{items}")
        if TARGET not in items:
            print(f"   目标 {TARGET!r} 不在可选项里，改用可选项里的第一项做写入验证")
            TARGET_USED = items[0] if items else TARGET
        else:
            TARGET_USED = TARGET
        index = items.index(TARGET_USED)
        _select_combo_index(target_combo, index)
        time.sleep(0.4)
        print(f"   点击后界面显示：{_combo_current_text(target_combo)!r}")
        _confirm(window, log)
        window = None
        time.sleep(1.0)
    finally:
        if window is not None:
            try:
                window.close()
            except Exception:  # noqa: BLE001
                pass

    close_stale_preferences(PRINTER, desktop=desktop)
    time.sleep(0.6)
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=log)
    try:
        _switch_tab(window, "纸张", log)
        combos = _combos_ready(window, 3, timeout=6)
        for combo in combos:
            if _norm(_label_before(combo)) == _norm("分隔页") or _norm(_text_of(combo)) == _norm("分隔页"):
                print(f"   重新打开后回读：{_combo_current_text(combo)!r}  ← 与设置值对比")
                break
    finally:
        try:
            window.close()
        except Exception:  # noqa: BLE001
            pass
    print("== 4. 恢复原值 ==")
    close_stale_preferences(PRINTER, desktop=desktop)
    time.sleep(0.6)
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=log)
    try:
        _switch_tab(window, "纸张", log)
        combos = _combos_ready(window, 3, timeout=6)
        for combo in combos:
            if _norm(_label_before(combo)) == _norm("分隔页") or _norm(_text_of(combo)) == _norm("分隔页"):
                items = _combo_items(combo)
                if original in items:
                    _select_combo_index(combo, items.index(original))
                    _confirm(window, log)
                    window = None
                    print(f"   已恢复为 {original!r}")
                else:
                    print(f"   原值 {original!r} 不在可选项 {items} 里，请手工核对")
                break
    finally:
        if window is not None:
            try:
                window.close()
            except Exception:  # noqa: BLE001
                pass
    close_stale_preferences(PRINTER, desktop=desktop)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
