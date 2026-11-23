"""端到端冒烟：真实下发打印作业，验证整条链路（可接受失败，重点是流程跑通）。

用法：
    .venv/Scripts/python.exe tools/smoke_e2e.py
    .venv/Scripts/python.exe tools/smoke_e2e.py --media "厚纸1+(121-157gsm)"
    .venv/Scripts/python.exe tools/smoke_e2e.py --file templates/docs/A4-check.pdf
    .venv/Scripts/python.exe tools/smoke_e2e.py --attr "逆序(0或1)=1"

说明：
- 打印机会收到真实作业（1 份、1–2 页），请确认连的是测试机/测试打印机。
- 脚本把前端“执行日志”页会收到的每一行事件原样打印出来，便于核对流程。
"""
from __future__ import annotations

import argparse
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.audit import AuditStore  # noqa: E402
from printcase.extended_attrs import ExtendedAttr  # noqa: E402
from printcase.attr_catalog import ATTRIBUTE_BY_LABEL  # noqa: E402
from printcase.models import PrintCase  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402


def ensure_docx(path: Path) -> Path:
    """用 Word 生成一个两页的测试文档（不依赖已有文件）。"""
    if path.exists() and path.stat().st_size > 0:
        return path
    import win32com.client

    path.parent.mkdir(parents=True, exist_ok=True)
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        doc = app.Documents.Add()
        for page in range(2):
            paragraph = doc.Content.Paragraphs.Add()
            paragraph.Range.Text = f"PrintCase Studio 冒烟测试 第 {page + 1} 页 {datetime.now():%H:%M:%S}"
            paragraph.Range.InsertParagraphAfter()
            if page == 0:
                doc.Content.InsertBreak(7)  # wdPageBreak
        doc.SaveAs2(str(path))
        doc.Close(False)
    finally:
        app.Quit()
    return path


def parse_attr(text: str) -> ExtendedAttr:
    label, _, value = text.partition("=")
    label = label.strip()
    spec = ATTRIBUTE_BY_LABEL.get(label)
    if spec is None:
        raise SystemExit(f"未知属性：{label}（可用属性见 templates 的“属性速查”表）")
    return ExtendedAttr(page=spec.page, label=label, value=value.strip())


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--printer", default="")
    parser.add_argument("--file", default="", help="打印源文件；缺省用 Word 生成两页 docx")
    parser.add_argument("--paper", default="A4 (210 x 297mm)")
    parser.add_argument("--media", default="", help='纸张类型，如 "厚纸1+(121-157gsm)"')
    parser.add_argument("--attr", action="append", default=[], help='扩展属性，如 "逆序(0或1)=1"')
    parser.add_argument("--copies", type=int, default=1)
    parser.add_argument("--to-file", default="", help="打印到文件（验证用，配合 Microsoft Print to PDF）")
    args = parser.parse_args()

    spooler = WindowsSpooler()
    printers = spooler.list_printers()
    printer = args.printer or spooler.win32print.GetDefaultPrinter()
    if args.to_file:
        printer = args.printer or "Microsoft Print to PDF"
        from printcase import direct_print

        output_path = Path(args.to_file).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()
        direct_print.SESSION_FACTORY = lambda p, d: direct_print.GdiSession(p, d, output_file=output_path)
    if printer not in printers:
        print(f"[FATAL] 打印机不存在：{printer}")
        return 2

    temp_dir = Path(tempfile.mkdtemp(prefix="printcase_smoke_"))
    if args.file:
        source = Path(args.file)
        if not source.is_absolute():
            source = (ROOT / source).resolve()
    else:
        source = ensure_docx(temp_dir / "smoke-test.docx")
    case = PrintCase(
        case_id="SMOKE", enabled=True, file_path=str(source), resolved_path=source,
        printer=printer, copies=args.copies, paper=args.paper,
        media_type=args.media, duplex="simplex", print_quality="600",
    )
    case.extended_attrs = [parse_attr(item) for item in args.attr]

    before = {job.job_id for job in spooler.get_jobs(printer)}
    print(f"=== 冒烟开始 {datetime.now():%H:%M:%S} ===")
    print(f"打印机   : {printer}")
    print(f"源文件   : {source}")
    print(f"纸张类型 : {args.media or '驱动默认'}")
    print(f"扩展属性 : {[attr.describe() for attr in case.extended_attrs] or '无'}")
    print(f"队列作业数（下发前）: {len(before)}")
    print("-" * 72)

    def on_event(kind: str, payload) -> None:
        if kind == "log":
            print(f"[log ] {payload}")
        elif kind == "result":
            print(f"[stat] {payload.case_id} → {payload.status.value}"
                  f"{' / Job#' + str(payload.job_id) if payload.job_id else ''}")

    audit = AuditStore(temp_dir / "smoke-audit.sqlite3")
    runner = PrintRunner(spooler, audit, queue_timeout=30)
    started = time.monotonic()
    results = runner.run_cases([case], printer, f"SMOKE-{int(time.time())}", on_event)
    elapsed = time.monotonic() - started

    result = results[0]
    after = spooler.get_jobs(printer)
    new_jobs = [job for job in after if job.job_id not in before]

    print("-" * 72)
    print(f"=== 冒烟结束 {datetime.now():%H:%M:%S}（用时 {elapsed:.1f}s）===")
    print(f"状态     : {result.status.value}")
    print(f"关联 Job : {result.job_id if result.job_id else '未关联'}")
    print(f"诊断信息 : {result.message}")
    print(f"队列新增作业: {[(job.job_id, job.document, job.status_text) for job in new_jobs] or '无'}")
    if args.to_file:
        out = Path(args.to_file).expanduser().resolve()
        print(f"输出文件   : {out} → {'已生成 %d 字节' % out.stat().st_size if out.exists() else '未生成'}")
    if new_jobs and not result.job_id:
        print("[提示] 队列里有新作业但未关联到用例，请检查作业名匹配。")
    return 0 if result.status.value not in {"下发失败"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
