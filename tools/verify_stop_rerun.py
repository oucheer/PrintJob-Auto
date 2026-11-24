"""端到端复现并验证用户场景：一批用例中途点停止 → 立刻再勾选另一批下发。

真实链路：真实打印机 + 真实驱动属性页（apply_case）；只把打印流重定向到
dist/stop-rerun/*.prn，不消耗纸张。

用法：.venv/Scripts/python.exe tools/verify_stop_rerun.py
"""
from __future__ import annotations

import sys
import tempfile
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase import direct_print  # noqa: E402
from printcase.audit import AuditStore  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402

STOP_AFTER = float(sys.argv[1]) if len(sys.argv) > 1 else 9.0
DOCS = ["A4-check.pdf", "duplex-check.pdf", "tray-check.pdf"]


def main() -> int:
    spooler = WindowsSpooler()
    printer = spooler.win32print.GetDefaultPrinter()
    out_dir = ROOT / "dist" / "stop-rerun"
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.prn"):
        stale.unlink()
    direct_print.SESSION_FACTORY = lambda p, d: direct_print.GdiSession(p, d, output_file=out_dir / f"{d}.prn")

    audit = AuditStore(Path(tempfile.mkdtemp(prefix="verify_audit_")) / "h.sqlite3")

    def cases(prefix: str) -> list[PrintCase]:
        return [
            PrintCase(
                case_id=f"{prefix}-{index + 1:02d}", enabled=True,
                file_path=str(ROOT / "templates/docs" / name),
                resolved_path=ROOT / "templates/docs" / name,
                printer=printer, copies=1, paper="A4", duplex="simplex", print_quality="600",
            )
            for index, name in enumerate(DOCS)
        ]

    events: list[str] = []

    def on_event(kind, payload) -> None:
        if kind == "log":
            events.append(str(payload))
            print("   ", str(payload)[:150])

    print(f"== 第一批（3 条），{STOP_AFTER:.0f} 秒后点停止 ==")
    runner = PrintRunner(spooler, audit, queue_timeout=20)
    holder = {"runner": runner}

    def stopper() -> None:
        time.sleep(STOP_AFTER)
        print(">>> 用户点击“停止”")
        holder["runner"].stop()

    threading.Thread(target=stopper, daemon=True).start()
    first = runner.run_cases(cases("B1"), printer, "STOP1", on_event)
    print("第一批：" + "；".join(f"{r.case_id}={r.status.value}" for r in first))

    print("\n== 第二批（另外 3 条，不停止） ==")
    runner2 = PrintRunner(spooler, audit, queue_timeout=20)
    started = time.monotonic()
    second = runner2.run_cases(cases("B2"), printer, "STOP2", on_event)
    print(f"第二批耗时 {time.monotonic() - started:.1f}s")
    for result in second:
        print(f"  {result.case_id}: {result.status.value} — {result.message[:120]}")

    failed = [r for r in second if r.status.value in {"下发失败", "打印前检查未通过"}]
    print("\n结论：" + ("第二批全部完成，无“属性页窗口未出现”类失败" if not failed else f"仍有失败 {len(failed)} 条"))
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
