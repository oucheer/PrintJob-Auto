"""注入式验证：让 printui 第一次启动就“静默退出”（复现用户遇到的故障），
看整轮下发能否靠重试自愈。真实打印机 + 真实驱动属性页，打印流重定向到文件。
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase import direct_print  # noqa: E402
from printcase import driver_ui  # noqa: E402
from printcase.audit import AuditStore  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402

real_launcher = driver_ui._default_launcher
state = {"dead_launches": int(sys.argv[1]) if len(sys.argv) > 1 else 1}


class DeadProc:
    """模拟 printui 秒退：poll() 永远是“已退出”，也没有任何窗口。"""

    pid = -1

    def poll(self):
        return 1

    def terminate(self):
        pass

    def wait(self, timeout=None):  # noqa: ARG002
        return 1

    def kill(self):
        pass


def flaky_launcher(printer: str):
    if state["dead_launches"] > 0:
        state["dead_launches"] -= 1
        print("   [注入] 本次 printui 静默退出，不创建窗口")
        return DeadProc()
    return real_launcher(printer)


def main() -> int:
    spooler = WindowsSpooler()
    printer = spooler.win32print.GetDefaultPrinter()
    out_dir = ROOT / "dist" / "flaky-launcher"
    out_dir.mkdir(parents=True, exist_ok=True)
    direct_print.SESSION_FACTORY = lambda p, d: direct_print.GdiSession(p, d, output_file=out_dir / f"{d}.prn")

    driver_ui._default_launcher = flaky_launcher  # noqa: SLF001 - 故意注入故障

    doc = ROOT / "templates/docs" / "A4-check.pdf"
    cases = [
        PrintCase(
            case_id=f"FLAKY-{index + 1}", enabled=True, file_path=str(doc), resolved_path=doc,
            printer=printer, copies=1, paper="A4", duplex="simplex", print_quality="600",
        )
        for index in range(2)
    ]
    logs: list[str] = []
    runner = PrintRunner(spooler, AuditStore(Path(tempfile.mkdtemp(prefix="flaky_audit_")) / "h.sqlite3"), queue_timeout=20)
    results = runner.run_cases(cases, printer, "FLAKY", lambda kind, payload: logs.append(str(payload)) if kind == "log" else None)
    direct_print.SESSION_FACTORY = None

    for line in logs:
        if "重试（第" in line or "结果：" in line:
            print("   ", line)
    for result in results:
        print(f"  {result.case_id}: {result.status.value} — {result.message[:100]}")
    ok = all(result.status.value != "下发失败" for result in results)
    print("结论：" + ("故障注入后仍全部下发成功（重试生效）" if ok else "仍有失败"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
