"""诊断：把驱动属性页每个页签上的当前选中项打印出来（只读，不做修改）。"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase.driver_ui import (  # noqa: E402
    _combos_ready,
    _desktop,
    _switch_tab,
    _tab_names,
    close_stale_preferences,
    open_preferences,
)
from printcase.ui_print import _descendants, _label_before, _text_of  # noqa: E402

ARGS = [item for item in sys.argv[1:] if not item.startswith("--")]
PRINTER = ARGS[0] if ARGS else "Pantum M9705DN Plus-3700 Series"


def dump_tab(window, name: str) -> None:
    print(f"\n===== {name} =====")
    for radio in _descendants(window, control_type="RadioButton"):
        text = _text_of(radio).strip()
        if not text:
            continue
        try:
            toggle = radio.get_toggle_state()
        except Exception:  # noqa: BLE001
            toggle = None
        mark = {1: "●", 0: "○"}.get(toggle, "?")
        print(f"  {mark} {text}")
    for check in _descendants(window, control_type="CheckBox"):
        text = _text_of(check).strip()
        if not text:
            continue
        try:
            toggle = check.get_toggle_state()
        except Exception:  # noqa: BLE001
            toggle = None
        print(f"  [{'x' if toggle == 1 else ' '}] {text}")
    for index, combo in enumerate(_combos_ready(window, 0, timeout=1.5)):
        try:
            current = str(combo.selected_text() or "")
        except Exception:  # noqa: BLE001
            current = "?"
        try:
            enabled = "启用" if combo.is_enabled() else "禁用"
        except Exception:  # noqa: BLE001
            enabled = "?"
        print(f"  Combo[{index}] {_label_before(combo)!r} = {current!r} ({enabled})")


def main() -> int:
    log = lambda text: print("   ·", text)  # noqa: E731
    desktop = _desktop()
    close_stale_preferences(PRINTER, desktop=desktop)
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=log)
    try:
        for name in [n for n in _tab_names(window) if n]:
            if not _switch_tab(window, name, log):
                continue
            time.sleep(0.9)
            dump_tab(window, name)
    finally:
        try:
            window.close()
        except Exception:  # noqa: BLE001
            pass
        close_stale_preferences(PRINTER, desktop=desktop)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
