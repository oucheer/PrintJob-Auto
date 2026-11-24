"""执行日志落盘。

界面上的“执行日志”只在内存里，程序一关就没了——偏偏下发失败时最需要的
正是那份完整过程。这里把每次运行（RUN）写成一个独立文本文件，放在
``%LOCALAPPDATA%\\PrintCaseStudio\\logs\\`` 下，并按数量保留最近的若干份。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

DEFAULT_KEEP = 40


def prune_logs(directory: Path, keep: int = DEFAULT_KEEP) -> int:
    """只保留最近 keep 份运行日志（按文件名里的时间戳排序）。"""
    directory = Path(directory)
    files = sorted(directory.glob("RUN-*.log"))
    removed = 0
    for old in files[:-keep] if keep > 0 and len(files) > keep else []:
        try:
            old.unlink()
            removed += 1
        except OSError:
            continue
    return removed


class RunLog:
    """一次运行的日志文件；写入即刷新，进程异常退出也不丢已有内容。"""

    def __init__(self, directory: Path, run_id: str, keep: int = DEFAULT_KEEP) -> None:
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.run_id = run_id
        self.keep = keep
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        self.path = self.directory / f"RUN-{stamp}-{run_id}.log"
        self._file = self.path.open("w", encoding="utf-8")

    @property
    def closed(self) -> bool:
        return self._file is None

    def write(self, text: str) -> None:
        if self._file is None:
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        for line in str(text).splitlines() or [""]:
            self._file.write(f"{stamp} {line}\n")
        self._file.flush()

    def close(self) -> Path:
        if self._file is not None:
            try:
                self._file.close()
            finally:
                self._file = None
            prune_logs(self.directory, self.keep)
        return self.path

    def __enter__(self) -> "RunLog":
        return self

    def __exit__(self, *_exc) -> None:
        self.close()
