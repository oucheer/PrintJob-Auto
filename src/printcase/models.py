from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any


class CaseStatus(str, Enum):
    PENDING = "待执行"
    PREFLIGHT_OK = "检查通过"
    PREFLIGHT_FAILED = "检查失败"
    SUBMITTING = "正在下发"
    QUEUED = "已进入队列"
    DRIVER_PRINTED = "驱动报告已打印"
    QUEUE_CLEARED = "队列已清除（未证实出纸）"
    SUBMITTED_UNVERIFIED = "已调用打印端（队列未关联）"
    TIMEOUT_QUEUED = "队列等待超时"
    FAILED = "下发失败"
    STOPPED = "已停止"


@dataclass(slots=True)
class PrintCase:
    case_id: str
    enabled: bool
    file_path: str
    printer: str = ""
    pages: str = ""
    copies: int = 1
    paper: str = "A4"
    mixed_sizes: str = "no"
    orientation: str = "portrait"
    color: str = "color"
    duplex: str = "simplex"
    scaling: str = "fit"
    bin: str = ""
    collate: str = ""
    print_quality: str = "auto"
    media_type: str = ""
    skip_blank: str = "no"
    borderless: str = "no"
    separator_page: str = "none"
    separator_bin: str = ""
    binding_location: str = "none"
    binding_method: str = "none"
    timeout_seconds: int = 90
    notes: str = ""
    source_row: int = 0
    resolved_path: Path | None = None
    input_errors: list[str] = field(default_factory=list)
    # 驱动扩展属性（移植自既有 print_panel 项目的属性词表），仅 UI 自动化模式生效
    extended_attrs: list = field(default_factory=list)

    def settings_summary(self) -> str:
        page_text = self.pages or "全部"
        paper_text = self.paper or "auto"
        parts = [f"{paper_text}/{self.orientation}/{self.duplex}/{self.color}", f"{self.copies}份", f"页码:{page_text}"]
        if self.bin:
            parts.append(f"纸盒:{self.bin}")
        if self.mixed_sizes == "yes":
            parts.append("混合尺寸")
        if self.collate == "no":
            parts.append("非逐份")
        elif self.collate == "yes":
            parts.append("逐份")
        if self.print_quality and self.print_quality.lower() != "auto":
            parts.append(f"{self.print_quality}dpi")
        if self.media_type:
            parts.append(f"纸型:{self.media_type}")
        if self.skip_blank == "yes":
            parts.append("跳过空白页")
        if self.borderless == "yes":
            parts.append("无边距")
        if self.separator_page != "none":
            label = {"before": "前", "after": "后", "both": "前后", "each_copy": "份间"}.get(self.separator_page, self.separator_page)
            parts.append(f"分隔页:{label}")
        if self.separator_bin:
            parts.append(f"分隔页纸盒:{self.separator_bin}")
        if self.binding_location != "none" or self.binding_method != "none":
            parts.append(f"装订:{self.binding_location}/{self.binding_method}")
        if self.extended_attrs:
            parts.append(f"扩展属性:{len(self.extended_attrs)}项")
        return " · ".join(parts)


@dataclass(slots=True)
class CaseResult:
    case_id: str
    status: CaseStatus
    message: str = ""
    printer: str = ""
    job_id: int | None = None
    started_at: datetime | None = None
    ended_at: datetime | None = None
    elapsed_seconds: float | None = None
    settings: dict[str, Any] = field(default_factory=dict)
    source_file: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "case_id": self.case_id,
            "status": self.status.value,
            "message": self.message,
            "printer": self.printer,
            "job_id": self.job_id,
            "started_at": self.started_at.isoformat(timespec="seconds") if self.started_at else "",
            "ended_at": self.ended_at.isoformat(timespec="seconds") if self.ended_at else "",
            "elapsed_seconds": round(self.elapsed_seconds, 2) if self.elapsed_seconds is not None else "",
            "settings": self.settings,
            "source_file": self.source_file,
        }
