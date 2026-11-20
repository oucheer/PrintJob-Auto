from __future__ import annotations

import platform
from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class QueueJob:
    job_id: int
    document: str
    status: int
    status_text: str
    pages_printed: int = 0
    total_pages: int = 0


class WindowsSpooler:
    """Thin adapter over Win32 print spooler APIs. Only enabled on Windows."""

    def __init__(self) -> None:
        if platform.system() != "Windows":
            raise RuntimeError("打印队列监控仅支持 Windows 10/11。")
        try:
            import win32print  # type: ignore
        except ImportError as exc:
            raise RuntimeError("缺少 pywin32，请运行 setup_windows.ps1。") from exc
        self.win32print = win32print

    def list_printers(self) -> list[str]:
        w = self.win32print
        flags = w.PRINTER_ENUM_LOCAL | w.PRINTER_ENUM_CONNECTIONS
        items = w.EnumPrinters(flags, None, 2)
        names = sorted({str(x.get("pPrinterName", "")).strip() for x in items if x.get("pPrinterName")})
        return names

    def get_jobs(self, printer_name: str) -> list[QueueJob]:
        w = self.win32print
        handle = w.OpenPrinter(printer_name)
        try:
            raw_jobs = w.EnumJobs(handle, 0, 999, 1)
            jobs: list[QueueJob] = []
            for item in raw_jobs:
                status = int(item.get("Status", 0) or 0)
                jobs.append(QueueJob(
                    job_id=int(item.get("JobId", 0)),
                    document=str(item.get("pDocument", "") or ""),
                    status=status,
                    status_text=self.describe_job_status(status),
                    pages_printed=int(item.get("PagesPrinted", 0) or 0),
                    total_pages=int(item.get("TotalPages", 0) or 0),
                ))
            return jobs
        finally:
            w.ClosePrinter(handle)

    def get_printer_status(self, printer_name: str) -> tuple[int, str]:
        w = self.win32print
        handle = w.OpenPrinter(printer_name)
        try:
            info = w.GetPrinter(handle, 6)
            status = int(info.get("dwStatus", 0) or 0)
            return status, self.describe_printer_status(status)
        finally:
            w.ClosePrinter(handle)

    def get_capabilities(self, printer_name: str) -> dict:
        """Read standard paper, tray, duplex and color capabilities from the driver."""
        w = self.win32print
        handle = w.OpenPrinter(printer_name)
        try:
            info = w.GetPrinter(handle, 2)
            port = str(info.get("pPortName", "") or "")
        finally:
            w.ClosePrinter(handle)

        def query(name: str, fallback: int):
            constant = getattr(w, name, fallback)
            try:
                return w.DeviceCapabilities(printer_name, port, constant)
            except Exception:
                return None

        paper_names = query("DC_PAPERNAMES", 16)
        paper_ids = query("DC_PAPERS", 2)
        bins = query("DC_BINS", 6)
        bin_names = query("DC_BINNAMES", 12)
        media_names = query("DC_MEDIATYPENAMES", 34)
        media_ids = query("DC_MEDIATYPES", 35)
        duplex = query("DC_DUPLEX", 7)
        color = query("DC_COLORDEVICE", 32)
        return {
            "papers": [str(x).strip() for x in paper_names if str(x).strip()] if isinstance(paper_names, (list, tuple)) else None,
            "paper_ids": [int(x) for x in paper_ids] if isinstance(paper_ids, (list, tuple)) else None,
            "media_names": [str(x).strip() for x in media_names if str(x).strip()] if isinstance(media_names, (list, tuple)) else None,
            "media_ids": [int(x) for x in media_ids] if isinstance(media_ids, (list, tuple)) else None,
            "bin_ids": [int(x) for x in bins] if isinstance(bins, (list, tuple)) else None,
            "bin_names": [str(x).strip() for x in bin_names if str(x).strip()] if isinstance(bin_names, (list, tuple)) else None,
            "duplex": bool(duplex) if duplex is not None and int(duplex) >= 0 else None,
            "color": bool(color) if color is not None and int(color) >= 0 else None,
        }

    def describe_job_status(self, status: int) -> str:
        w = self.win32print
        flags = [
            ("JOB_STATUS_ERROR", "打印错误"),
            ("JOB_STATUS_OFFLINE", "打印机离线"),
            ("JOB_STATUS_PAPEROUT", "缺纸"),
            ("JOB_STATUS_PAUSED", "作业已暂停"),
            ("JOB_STATUS_BLOCKED_DEVQ", "驱动阻止作业"),
            ("JOB_STATUS_USER_INTERVENTION", "需要人工干预"),
            ("JOB_STATUS_PRINTING", "打印中"),
            ("JOB_STATUS_SPOOLING", "正在进入队列"),
            ("JOB_STATUS_PRINTED", "驱动报告已打印"),
            ("JOB_STATUS_COMPLETE", "已发送到打印设备"),
            ("JOB_STATUS_DELETING", "正在删除"),
            ("JOB_STATUS_DELETED", "已删除"),
            ("JOB_STATUS_RETAINED", "作业保留"),
        ]
        names = [label for const, label in flags if status & getattr(w, const, 0)]
        return "、".join(names) if names else ("无状态标记" if status == 0 else f"状态位 0x{status:08X}")

    def describe_printer_status(self, status: int) -> str:
        w = self.win32print
        flags = [
            ("PRINTER_STATUS_PAUSED", "已暂停"),
            ("PRINTER_STATUS_ERROR", "错误"),
            ("PRINTER_STATUS_PENDING_DELETION", "等待删除"),
            ("PRINTER_STATUS_PAPER_JAM", "卡纸"),
            ("PRINTER_STATUS_PAPER_OUT", "缺纸"),
            ("PRINTER_STATUS_MANUAL_FEED", "等待手动进纸"),
            ("PRINTER_STATUS_PAPER_PROBLEM", "纸张异常"),
            ("PRINTER_STATUS_OFFLINE", "离线"),
            ("PRINTER_STATUS_IO_ACTIVE", "I/O 活动"),
            ("PRINTER_STATUS_BUSY", "忙碌"),
            ("PRINTER_STATUS_PRINTING", "打印中"),
            ("PRINTER_STATUS_OUTPUT_BIN_FULL", "出纸盒已满"),
            ("PRINTER_STATUS_NOT_AVAILABLE", "不可用"),
            ("PRINTER_STATUS_WAITING", "等待中"),
            ("PRINTER_STATUS_PROCESSING", "处理中"),
            ("PRINTER_STATUS_INITIALIZING", "初始化"),
            ("PRINTER_STATUS_WARMING_UP", "预热中"),
            ("PRINTER_STATUS_TONER_LOW", "碳粉不足"),
            ("PRINTER_STATUS_NO_TONER", "缺少碳粉"),
            ("PRINTER_STATUS_PAGE_PUNT", "页面无法处理"),
            ("PRINTER_STATUS_USER_INTERVENTION", "需要人工干预"),
            ("PRINTER_STATUS_OUT_OF_MEMORY", "内存不足"),
            ("PRINTER_STATUS_DOOR_OPEN", "舱门打开"),
            ("PRINTER_STATUS_SERVER_UNKNOWN", "打印服务器未知"),
        ]
        names = [label for const, label in flags if status & getattr(w, const, 0)]
        return "、".join(names) if names else ("就绪/无状态标记" if status == 0 else f"状态位 0x{status:08X}")
