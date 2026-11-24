"""打印队列持续监视。

Windows 队列里的作业存活时间可以短到几百毫秒（驱动把数据交给打印机后作业
立刻从队列消失）。执行器如果在“下发之后才开始一次一次地读队列”，读取间隔
（读队列耗时 + sleep）很容易大于作业存活时间，于是作业已经打印完成，程序却
永远看不到它，只能报“队列未关联”并空等到超时——批量执行时每条用例都要白等
一轮。

这里用一个后台线程在下发**之前**就开始高频轮询，把每个出现过的作业留档
（作业号、文档名、状态、首次/末次出现时间），执行器随后只是查询这份留档，
因此不会再漏掉短命作业，也不需要长时间空转。
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Callable, Iterable

from .spooler import QueueJob

# 队列作业状态里代表“出错”的关键字（与执行器判定保持一致）
ERROR_JOB_MASKS = ("打印错误", "打印机离线", "缺纸", "作业已暂停", "驱动阻止作业", "需要人工干预")


@dataclass
class JobObservation:
    """单个队列作业的观测留档。"""

    job: QueueJob
    first_seen: float
    last_seen: float
    disappeared_at: float | None = None

    def update(self, job: QueueJob, now: float) -> None:
        self.job = job
        self.last_seen = now
        self.disappeared_at = None


@dataclass
class WatchResult:
    """wait_for_new 的结果。"""

    job: QueueJob | None = None
    ambiguous: bool = False
    note: str = ""


class QueueWatcher:
    """在后台持续采样打印队列，保留所有出现过的作业。"""

    def __init__(
        self,
        reader: Callable[[], Iterable[QueueJob]],
        interval: float = 0.08,
        stale_after: float = 300.0,
        printed_mask: int = 0,
    ) -> None:
        self._reader = reader
        self._interval = interval
        self._stale_after = stale_after
        self._printed_mask = printed_mask
        self._lock = threading.Lock()
        self._observed: dict[int, JobObservation] = {}
        self._present: set[int] = set()
        self._stop = threading.Event()
        self._ready = threading.Event()
        self._thread: threading.Thread | None = None
        self.error: str = ""
        # 采样质量：用于区分“队列里真没有作业”和“监视器根本没读成功”
        self.samples = 0
        self.failures = 0

    # ---------- 生命周期 ----------
    def start(self) -> "QueueWatcher":
        if self._thread is not None:
            return self
        self._thread = threading.Thread(target=self._loop, name="PrintQueueWatcher", daemon=True)
        self._thread.start()
        return self

    def stop(self) -> None:
        self._stop.set()
        thread = self._thread
        if thread is not None:
            thread.join(timeout=2.0)
            self._thread = None

    def __enter__(self) -> "QueueWatcher":
        return self.start()

    def __exit__(self, *exc_info) -> None:
        self.stop()

    def wait_ready(self, timeout: float = 1.0) -> bool:
        """等第一次采样完成，确保 snapshot 拿到的是真实基线。"""
        return self._ready.wait(timeout)

    # ---------- 后台采样 ----------
    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                jobs = list(self._reader())
                self.error = ""
            except Exception as exc:  # noqa: BLE001
                self.error = f"{type(exc).__name__}: {exc}"
                self.failures += 1
                jobs = None
            if jobs is not None:
                # 记录过程本身出错（例如脏数据）不能把线程带走：线程一死，
                # 后续所有用例都会“队列未关联”，还查不出原因。
                try:
                    self._record(jobs)
                    self.samples += 1
                    self._ready.set()
                except Exception as exc:  # noqa: BLE001
                    self.error = f"记录队列快照失败：{type(exc).__name__}: {exc}"
                    self.failures += 1
            self._stop.wait(self._interval)

    def health_text(self) -> str:
        """给日志用的一句话采样质量说明。"""
        detail = f"采样 {self.samples} 次"
        if self.failures:
            detail += f"，失败 {self.failures} 次"
        if self.error:
            detail += f"（最后一次错误：{self.error}）"
        return detail

    def _record(self, jobs: list[QueueJob]) -> None:
        now = time.monotonic()
        seen_ids = {job.job_id for job in jobs}
        with self._lock:
            for job in jobs:
                observation = self._observed.get(job.job_id)
                if observation is None:
                    self._observed[job.job_id] = JobObservation(job=job, first_seen=now, last_seen=now)
                else:
                    observation.update(job, now)
            # 上一轮在队列里、这一轮消失 → 记下消失时间（供“队列已清除”判定）
            for job_id in self._present - seen_ids:
                observation = self._observed.get(job_id)
                if observation is not None and observation.disappeared_at is None:
                    observation.disappeared_at = now
            self._present = seen_ids
            self._prune(now)

    def _prune(self, now: float) -> None:
        if not self._stale_after:
            return
        for job_id in [key for key, obs in self._observed.items() if now - obs.last_seen > self._stale_after]:
            self._observed.pop(job_id, None)

    # ---------- 查询接口 ----------
    def snapshot_ids(self, include_recent: float = 120.0) -> set[int]:
        """下发前调用：返回“这段时间内出现过”的作业号，作为新作业的排除基线。"""
        now = time.monotonic()
        with self._lock:
            return {
                job_id for job_id, obs in self._observed.items()
                if now - obs.last_seen <= include_recent or obs.disappeared_at is None
            }

    def observation(self, job_id: int) -> JobObservation | None:
        with self._lock:
            return self._observed.get(job_id)

    def wait_for_new(
        self,
        known_ids: set[int] | None,
        timeout: float,
        document: str = "",
        should_stop: Callable[[], bool] | None = None,
    ) -> WatchResult:
        """等待一个不属于 known_ids 的新作业出现。

        document 非空时优先按文档名匹配；只有唯一一个新作业才允许“名称不匹配
        也算命中”，多个新作业且无法用名称区分时返回 ambiguous，交由上层保守处理。
        """
        known = set(known_ids or ())
        document_key = str(document or "").strip().casefold()
        deadline = time.monotonic() + max(0.0, timeout)
        while True:
            with self._lock:
                candidates = [obs for job_id, obs in self._observed.items() if job_id not in known]
            if candidates:
                if document_key:
                    matched = [
                        obs for obs in candidates
                        if (obs.job.document or "").strip() and (
                            document_key in obs.job.document.casefold()
                            or obs.job.document.casefold() in document_key
                        )
                    ]
                    if len(matched) == 1:
                        return WatchResult(job=matched[0].job, note="name")
                    if len(matched) > 1:
                        return WatchResult(ambiguous=True, note="multiple-name-match")
                    if len(candidates) == 1:
                        # 名称对不上但只有一个新作业：仍可能是同一作业（驱动改名）
                        return WatchResult(job=candidates[0].job, note="single-new")
                    return WatchResult(ambiguous=True, note="multiple-new-jobs")
                if len(candidates) == 1:
                    return WatchResult(job=candidates[0].job, note="single-new")
                return WatchResult(ambiguous=True, note="multiple-new-jobs")
            if should_stop is not None and should_stop():
                return WatchResult(note="stopped")
            if time.monotonic() >= deadline:
                return WatchResult(note="timeout")
            time.sleep(0.05)

    def wait_terminal(
        self,
        job_id: int,
        timeout: float,
        should_stop: Callable[[], bool] | None = None,
    ) -> tuple[str, QueueJob | None]:
        """跟踪已关联作业直到可观察终态。

        返回 (kind, job)，kind ∈ printed / error / gone / timeout / stopped。
        """
        deadline = time.monotonic() + max(0.0, timeout)
        last_job: QueueJob | None = None
        while True:
            with self._lock:
                observation = self._observed.get(job_id)
                if observation is not None:
                    last_job = observation.job
                    disappeared = observation.disappeared_at is not None
                else:
                    disappeared = last_job is not None
            if last_job is not None:
                if any(flag in (last_job.status_text or "") for flag in ERROR_JOB_MASKS):
                    return "error", last_job
                if self._printed_mask and last_job.status & self._printed_mask:
                    return "printed", last_job
                if disappeared:
                    return "gone", last_job
            if should_stop is not None and should_stop():
                return "stopped", last_job
            if time.monotonic() >= deadline:
                return "timeout", last_job
            time.sleep(0.1)
