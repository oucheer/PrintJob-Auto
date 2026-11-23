"""驱动属性页验证工具：设置 → 只读回读校验 → 还原（不下发打印）。

用法：
    .venv/Scripts/python.exe tools/verify_driver_ui.py --printer "Pantum CM605ADN Series" \
        --media "厚纸1+(121-157gsm)" --paper "A4 (210 x 297mm)"

说明：
- 会临时修改打印机的全局默认设置，脚本结束（含异常）后自动还原。
- 只做属性设置与回读校验，不向队列提交任何作业。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.driver_ui import DriverUiError, apply_case, available, snapshot  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402

FIELDS = ("PaperSize", "MediaType", "DefaultSource", "Duplex", "Color", "Copies", "PrintQuality")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--printer", default="", help="缺省用系统默认打印机")
    parser.add_argument("--paper", default="A4 (210 x 297mm)")
    parser.add_argument("--media", default="", help='纸张类型，如 "厚纸1+(121-157gsm)"')
    parser.add_argument("--bin", default="")
    parser.add_argument("--duplex", default="simplex", choices=("simplex", "duplexlong", "duplexshort"))
    parser.add_argument("--color", default="monochrome", choices=("color", "monochrome"))
    parser.add_argument("--quality", default="600")
    parser.add_argument("--copies", type=int, default=1)
    args = parser.parse_args()

    spooler = WindowsSpooler()
    printer = args.printer or spooler.win32print.GetDefaultPrinter()
    print(f"打印机：{printer}")

    ok, detail = available(printer)
    print(f"属性页自检：{'通过' if ok else '未通过'} — {detail}")
    if not ok:
        return 2

    caps = spooler.get_capabilities(printer)
    before = snapshot(printer)
    print("设置前：" + "、".join(f"{name}={before[name]}" for name in FIELDS))

    case = PrintCase(
        case_id="VERIFY", enabled=True, file_path="", printer=printer,
        paper=args.paper, media_type=args.media, bin=args.bin,
        duplex=args.duplex, color=args.color, print_quality=args.quality, copies=args.copies,
    )

    restore = None
    try:
        restore = apply_case(printer, case, caps=caps, log=lambda text: print(f"[log] {text}"))
        after = snapshot(printer)
        print("设置后：" + "、".join(f"{name}={after[name]}" for name in FIELDS))
        changed = {name: (before[name], after[name]) for name in FIELDS if before[name] != after[name]}
        print("变化项：" + ("、".join(f"{k} {v[0]}→{v[1]}" for k, v in changed.items()) or "无"))
        return 0
    except DriverUiError as exc:
        print(f"[FAIL] {exc}")
        return 1
    finally:
        if restore is not None:
            restore()


if __name__ == "__main__":
    raise SystemExit(main())
