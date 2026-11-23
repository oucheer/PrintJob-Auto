"""批量下发冒烟：模拟 GUI 勾选多个用例一起下发，验证串行编排是否完整。

为免浪费纸张：驱动属性页/作业号走真实默认打印机，但通过 StartDoc 的
lpszOutput 把打印流重定向到 dist/batch-smoke/*.prn（驱动原始输出，通常是
PCL/PS 等页面描述语言，不是 PDF）。每个用例产出独立文件，便于核对。

用法：
    .venv/Scripts/python.exe tools/smoke_batch.py
"""
from __future__ import annotations

import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase import direct_print  # noqa: E402
from printcase.audit import AuditStore  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402


def main() -> int:
    spooler = WindowsSpooler()
    printer = spooler.win32print.GetDefaultPrinter()
    if printer not in spooler.list_printers():
        print(f"[FATAL] 默认打印机不存在：{printer}")
        return 2

    out_dir = ROOT / "dist" / "batch-smoke"
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.prn"):
        stale.unlink()

    # 每个文档输出到独立文件（按文档主干命名），作业号仍由真实打印机分配
    def factory(p, d):
        return direct_print.GdiSession(p, d, output_file=out_dir / f"{d}.prn")

    direct_print.SESSION_FACTORY = factory

    sources = [
        ROOT / "templates/docs/A4-check.pdf",
        ROOT / "templates/docs/duplex-check.pdf",
        ROOT / "templates/docs/tray-check.pdf",
    ]
    cases = [
        PrintCase(
            case_id=f"BATCH-{i + 1:02d}", enabled=True, file_path=str(src), resolved_path=src,
            printer=printer, copies=1, paper="A4", duplex="simplex", print_quality="600",
        )
        for i, src in enumerate(sources)
    ]

    temp_dir = Path(tempfile.mkdtemp(prefix="printcase_batch_"))
    before_jobs = {j.job_id for j in spooler.get_jobs(printer)}
    print(f"=== 批量下发冒烟开始 {datetime.now():%H:%M:%S} ===")
    print(f"打印机   : {printer}")
    print(f"用例数量 : {len(cases)}（{', '.join(c.case_id for c in cases)}）")
    print(f"输出目录 : {out_dir}")
    print("-" * 72)

    def on_event(kind, payload):
        if kind == "log":
            print(f"[log ] {payload}")
        elif kind == "result":
            print(f"[stat] {payload.case_id} → {payload.status.value}"
                  f"{' / Job#' + str(payload.job_id) if payload.job_id else ''}")

    audit = AuditStore(temp_dir / "batch-audit.sqlite3")
    runner = PrintRunner(spooler, audit, queue_timeout=10)
    started = time.monotonic()
    results = runner.run_cases(cases, printer, f"BATCH-{int(time.time())}", on_event)
    elapsed = time.monotonic() - started

    after_jobs = [j for j in spooler.get_jobs(printer) if j.job_id not in before_jobs]
    print("-" * 72)
    print(f"=== 批量下发冒烟结束（用时 {elapsed:.1f}s）===")
    for r in results:
        print(f"  {r.case_id}: {r.status.value} / Job#{r.job_id or '未关联'}")
    print(f"队列新增作业: {[(j.job_id, j.document) for j in after_jobs] or '无'}")
    produced = [p for p in sorted(out_dir.glob("*.prn")) if p.stat().st_size > 0]
    print(f"输出文件 ({len(produced)}): {[p.name for p in produced]}")
    job_ids = [r.job_id for r in results if r.job_id]
    ok = (
        len(produced) == len(cases)
        and all(r.status.value != "下发失败" for r in results)
        and len(job_ids) == len(cases)
        and len(set(job_ids)) == len(cases)  # 每个用例获得独立作业号，未串作业
    )
    print(f"批量下发结果: {'PASS' if ok else 'FAIL'}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
