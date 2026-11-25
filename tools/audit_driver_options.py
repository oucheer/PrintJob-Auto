"""把驱动属性页每个页签的真实控件与条目全部导出（只读，只开合窗口，不打印）。

用途：模板里的下拉选项必须与驱动完全一致。以前凭文档抄的选项（如“分隔页”的
无/前/后/前后）在真机上根本不存在，导致执行时才报“找不到该项”。这里把驱动实际
暴露的下拉条目、单选、勾选框逐条打印出来，供模板与 attr_catalog 对照。

用法：
    .venv\\Scripts\\python.exe tools\\audit_driver_options.py [打印机名] [--json 输出.json]
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from printcase.driver_ui import (  # noqa: E402
    _close,
    _combos_ready,
    _desktop,
    _switch_tab,
    close_stale_preferences,
    open_preferences,
)
from printcase.ui_print import (  # noqa: E402
    _combo_current_text,
    _combo_items,
    _descendants,
    _label_before,
    _text_of,
)

DEFAULT_PRINTER = "Pantum M9705DN Plus-3700 Series"
TABS = ("基本", "纸张", "版面", "水印", "高级")


def _tab_names(window) -> list[str]:
    names = []
    for tab in _descendants(window, control_type="TabItem"):
        text = _text_of(tab).strip()
        if text:
            names.append(text)
    return names


def dump_tab(window, name: str) -> dict:
    _switch_tab(window, name, lambda _t: None)
    time.sleep(0.6)
    combos = _combos_ready(window, 12, timeout=6)
    data = {"combos": [], "radios": [], "checkboxes": []}
    for index, combo in enumerate(combos):
        try:
            enabled = bool(combo.is_enabled())
        except Exception:  # noqa: BLE001
            enabled = True
        entry = {
            "index": index,
            "label": str(_label_before(combo)).strip(),
            "name": _text_of(combo).strip(),
            "current": _combo_current_text(combo),
            "enabled": enabled,
        }
        if enabled:
            try:
                entry["items"] = _combo_items(combo)
            except Exception as exc:  # noqa: BLE001
                entry["items"] = []
                entry["error"] = f"{type(exc).__name__}: {exc}"
        else:
            entry["items"] = []
            entry["note"] = "灰化：控件当前不可用"
        data["combos"].append(entry)
    for radio in _descendants(window, control_type="RadioButton"):
        text = _text_of(radio).strip()
        if not text:
            continue
        try:
            checked = radio.is_selected()
        except Exception:  # noqa: BLE001
            checked = None
        data["radios"].append({"text": text, "checked": checked})
    for check in _descendants(window, control_type="CheckBox"):
        text = _text_of(check).strip()
        if not text:
            continue
        try:
            checked = check.is_selected()
        except Exception:  # noqa: BLE001
            checked = None
        data["checkboxes"].append({"text": text, "checked": checked})
    return data


def main() -> int:
    args = [item for item in sys.argv[1:] if not item.startswith("--")]
    printer = args[0] if args else DEFAULT_PRINTER
    out_path = None
    for token in sys.argv[1:]:
        if token.startswith("--json"):
            out_path = token.split("=", 1)[1] if "=" in token else "dist/driver-options.json"

    desktop = _desktop()
    close_stale_preferences(printer, desktop=desktop)
    window = open_preferences(printer, timeout=25, desktop=desktop, log=lambda t: print("  ·", t))
    report: dict = {"printer": printer, "tabs": {}}
    try:
        tabs = _tab_names(window)
        report["tab_names"] = tabs
        print("驱动页签：", "、".join(tabs))
        for name in (tabs or TABS):
            key = name.split("(")[0].strip()
            matched = next((t for t in tabs if key and key in t), None)
            if not matched:
                continue
            try:
                report["tabs"][matched] = dump_tab(window, matched)
            except Exception as exc:  # noqa: BLE001
                report["tabs"][matched] = {"error": f"{type(exc).__name__}: {exc}"}
    finally:
        try:
            _close(window, log=lambda _t: None)
        except Exception:  # noqa: BLE001
            pass
        close_stale_preferences(printer, desktop=desktop)

    for tab, data in report["tabs"].items():
        print(f"\n===== {tab} =====")
        if "error" in data:
            print("  读取失败：", data["error"])
            continue
        for combo in data["combos"]:
            flag = "" if combo["enabled"] else "  [灰化]"
            print(f"  [{combo['index']}] 标签={combo['label']!r} 当前={combo['current']!r}{flag}")
            if combo["items"]:
                print(f"      条目({len(combo['items'])})：{'、'.join(combo['items'])}")
            elif combo.get("error"):
                print(f"      条目读取失败：{combo['error']}")
        radios = [f"{r['text']}{'(*)' if r['checked'] else ''}" for r in data["radios"]]
        if radios:
            print("  单选：", "、".join(radios))
        checks = [f"{c['text']}{'(*)' if c['checked'] else ''}" for c in data["checkboxes"]]
        if checks:
            print("  勾选：", "、".join(checks))

    if out_path:
        dest = Path(out_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n已写入 {dest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
