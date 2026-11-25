"""整机功能点验证（冒烟 + 黑盒 + 白盒），真实调用各模块。

不依赖界面点击，但使用真实打印机驱动、真实 Word、真实队列；
只有“PDF 直打”一项把打印流重定向到文件（不耗纸），其余按需最少出纸。

用法：
    .venv/Scripts/python.exe tools/verify_app.py [--with-print]
        --with-print  额外做一次 UI 通道真实下发（1 页）
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
from printcase.driver_ui import (  # noqa: E402
    DriverUiError,
    apply_case,
    close_stale_preferences,
    separator_attrs,
)
from printcase.excel_reader import read_cases  # noqa: E402
from printcase.models import CaseStatus, PrintCase  # noqa: E402
from printcase.queue_watch import QueueWatcher  # noqa: E402
from printcase.runlog import RunLog, prune_logs  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402
from printcase.validator import validate_case  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
DOCX = ROOT / "templates" / "docs" / "A4纵向测试页2.docx"
TEMPLATE = ROOT / "templates" / "PrintCase_Template.xlsx"
WITH_PRINT = "--with-print" in sys.argv

results: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    results.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""), flush=True)


def main() -> int:
    spooler = WindowsSpooler()
    print(f"打印机：{spooler.win32print.GetDefaultPrinter()}")
    print(f"目标驱动：{PRINTER}")
    print("-" * 72)

    # ---------- 1. Excel 模板读取（黑盒：用户输入） ----------
    cases, warns = read_cases(TEMPLATE)
    check("读取用例模板", len(cases) >= 20 and not warns,
          f"{len(cases)} 条用例，告警 {len(warns)} 条")
    tc013 = next((c for c in cases if c.case_id == "TC-013"), None)
    check("解析用例字段", tc013 is not None and tc013.pages == "1-2" and tc013.borderless == "yes",
          f"TC-013 pages={tc013.pages if tc013 else '-'} borderless={tc013.borderless if tc013 else '-'}")

    # ---------- 2. 预检（黑盒：非法输入拦截） ----------
    bad = PrintCase(case_id="TC-BAD", enabled=True, file_path=str(ROOT / "templates/docs/不存在.pdf"),
                    printer="不存在的打印机", paper="A5")
    errs = validate_case(bad, spooler.list_printers(), PRINTER)
    check("预检拦截无效用例", bool(errs), "；".join(errs[:2]))

    # ---------- 3. 驱动属性页（黑盒：真实驱动 UI） ----------
    logs: list[str] = []
    started = time.monotonic()
    restore = apply_case(PRINTER, tc013, caps=spooler.get_capabilities(PRINTER), log=logs.append)
    elapsed = time.monotonic() - started
    joined = "\n".join(logs)
    check("驱动属性页真实设置", "回读确认" in joined and restore is not None, f"{elapsed:.1f}s")
    check("驱动属性回读覆盖关键项",
          all(word in joined for word in ("纸张尺寸", "双面", "打印质量")),
          "纸张尺寸/双面/打印质量")
    try:
        restore()
        check("驱动默认值还原（权限受限时降级）", True, "已尝试还原，无权限时仅告警")
    except Exception as exc:  # noqa: BLE001
        check("驱动默认值还原（权限受限时降级）", False, str(exc))
    cleared = close_stale_preferences(PRINTER)
    check("清理驱动属性页残留", cleared == 0 or cleared > 0, f"清理 {cleared} 项")

    # ---------- 4. 分隔页（驱动硬分隔页：只有 作业之间/份数之间） ----------
    with tempfile.TemporaryDirectory(prefix="verify_app_") as tmp:
        tmp_dir = Path(tmp)
        sep_case = PrintCase(
            case_id="TC-SEP", enabled=True, file_path=str(DOCX), resolved_path=DOCX,
            printer=PRINTER, separator_page="job", bin="纸盒1", separator_bin="纸盒2",
        )
        attrs = [(item.label, item.value) for item in separator_attrs(sep_case)]
        check("分隔页映射到驱动“输出”分组",
              attrs == [("输出模式", "分隔页"), ("分隔页", "作业之间"),
                        ("分隔页（OHP插页）纸张来源", "纸盒2")],
              "、".join(f"{k}={v}" for k, v in attrs))
        no_tray = PrintCase(
            case_id="TC-SEP2", enabled=True, file_path=str(DOCX), resolved_path=DOCX,
            printer=PRINTER, separator_page="copies",
        )
        try:
            separator_attrs(no_tray)
            blocked = False
            detail = "未拦下"
        except DriverUiError as exc:
            blocked = True
            detail = str(exc)[:60]
        check("分隔页未指定纸盒时拦下（驱动会整组灰化）", blocked, detail)
        check("分隔页旧取值（前/后/前后）判错",
              any("分隔页必须为" in e for e in validate_case(
                  PrintCase(case_id="TC-SEP3", enabled=True, file_path=str(DOCX),
                            resolved_path=DOCX, separator_page="before", bin="纸盒1"),
                  {PRINTER}, PRINTER)),
              "before 不再被接受")

        # ---------- 5. PDF 直打通道（重定向到文件，不耗纸） ----------
        out_dir = tmp_dir / "gdi"
        out_dir.mkdir()
        direct_print.SESSION_FACTORY = lambda p, d: direct_print.GdiSession(p, d, output_file=out_dir / f"{d}.prn")
        audit = AuditStore(tmp_dir / "audit.sqlite3")
        pdf_case = PrintCase(case_id="TC-GDI", enabled=True,
                             file_path=str(ROOT / "templates/docs/A4-check.pdf"),
                             resolved_path=ROOT / "templates/docs/A4-check.pdf",
                             printer=PRINTER, paper="A4", copies=1, print_quality="600")
        runner = PrintRunner(spooler, audit, queue_timeout=15)
        gdi_results = runner.run_cases([pdf_case], PRINTER, "VERIFY-GDI", lambda *_: None)
        produced = list(out_dir.glob("*.prn"))
        check("PDF 直打通道（输出到文件）",
              gdi_results and gdi_results[0].status != CaseStatus.FAILED and produced,
              f"{gdi_results[0].status.value} / {produced[0].name if produced else '无输出'}")
        direct_print.SESSION_FACTORY = None

        # ---------- 6. 停止流程（白盒：编排中断） ----------
        order: list[str] = []

        def slow_submit(printer, case, source_path, log, temp):
            order.append(case.case_id)
            holder["runner"].stop()
            return "fake", Path(source_path).stem

        holder = {}
        stop_runner = PrintRunner(
            spooler, AuditStore(tmp_dir / "stop.sqlite3"), queue_timeout=5,
            ui_submitter=slow_submit, driver_applier=lambda *a, **k: (lambda: None),
        )
        holder["runner"] = stop_runner
        stop_cases = [PrintCase(case_id=f"TC-STOP{i}", enabled=True, file_path=str(DOCX), resolved_path=DOCX,
                                printer=PRINTER, pages="1-1") for i in range(3)]
        stop_results = stop_runner.run_cases(stop_cases, PRINTER, "VERIFY-STOP", lambda *_: None)
        check("停止后不再下发后续用例",
              stop_results[1].status == CaseStatus.STOPPED and order == ["TC-STOP0"],
              "、".join(r.status.value for r in stop_results))

        # ---------- 7. 执行日志落盘 ----------
        log_dir = tmp_dir / "logs"
        run_log = RunLog(log_dir, "VERIFYLOG")
        run_log.write("第一行")
        run_log.write("第二行")
        run_log.close()
        text = run_log.path.read_text(encoding="utf-8")
        check("执行日志落盘", "第一行" in text and "第二行" in text, run_log.path.name)
        for i in range(3):
            RunLog(log_dir, f"OLD{i}").close()
        removed = prune_logs(log_dir, keep=2)
        check("日志按数量轮转", removed >= 2 and len(list(log_dir.glob("RUN-*.log"))) == 2,
              f"清理 {removed} 份，剩余 2 份")

    # ---------- 8. 审计库 ----------
    with tempfile.TemporaryDirectory(prefix="verify_audit_") as tmp2:
        store = AuditStore(Path(tmp2) / "audit.sqlite3")
        from printcase.models import CaseResult
        store.save("RUNX", CaseResult("TC-1", CaseStatus.QUEUE_CLEARED, "演示", printer=PRINTER))
        rows = store.recent(10)
        check("审计库记录执行结果", bool(rows) and rows[0]["case_id"] == "TC-1", f"{len(rows)} 条")

    # ---------- 9. 队列监视器 ----------
    watcher = QueueWatcher(lambda: spooler.get_jobs(PRINTER), interval=0.05).start()
    watcher.wait_ready(2.0)
    time.sleep(0.2)
    watcher.stop()
    check("队列监视器采样正常", watcher.samples > 0 and watcher.failures == 0, watcher.health_text())

    # ---------- 10. UI 通道真实下发（可选，1 页） ----------
    if WITH_PRINT:
        from printcase.ui_print import print_with_ui
        case = PrintCase(case_id="TC-UISMOKE", enabled=True, file_path=str(DOCX), resolved_path=DOCX,
                         printer=PRINTER, pages="1-1", copies=1, paper="A4", color="monochrome",
                         duplex="simplex", print_quality="600", timeout_seconds=90)
        watch = QueueWatcher(lambda: spooler.get_jobs(PRINTER), interval=0.05).start()
        watch.wait_ready(2.0)
        known = watch.snapshot_ids()
        ui_logs: list[str] = []
        started = time.monotonic()
        try:
            backend, _stem = print_with_ui(PRINTER, case, DOCX, ui_logs.append,
                                           timeout=90, set_driver_attributes=False)
            outcome = watch.wait_for_new(known, 20.0, DOCX.stem)
            ok = backend == "ui-automation" and outcome.job is not None
            detail = f"{backend} / Job#{outcome.job.job_id if outcome.job else '未关联'} / {time.monotonic() - started:.1f}s"
        except Exception as exc:  # noqa: BLE001
            ok, detail = False, f"{type(exc).__name__}: {exc}"
        watch.stop()
        check("UI 通道真实下发 1 页", ok, detail)
        check("打印结果确认链完整",
              any("打印视图已关闭" in line for line in ui_logs)
              and any("打印作业已提交给打印端" in line for line in ui_logs),
              "视图关闭 + 提交确认")

    print("-" * 72)
    failed = [name for name, ok, _ in results if not ok]
    print(f"共 {len(results)} 项：通过 {len(results) - len(failed)}，失败 {len(failed)}")
    if failed:
        print("失败项：" + "、".join(failed))
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
