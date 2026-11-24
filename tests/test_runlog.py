"""执行日志落盘：停止/失败后还能翻到完整过程。"""
from __future__ import annotations

from printcase.runlog import RunLog, prune_logs


def test_run_log_writes_lines_with_time_stamp(tmp_path):
    log = RunLog(tmp_path / "logs", "RUN123")
    log.write("第一行")
    log.write("多行\n第二行")
    log.close()

    text = log.path.read_text(encoding="utf-8")
    assert "第一行" in text and "第二行" in text
    assert log.path.name.startswith("RUN-") and log.path.name.endswith("RUN123.log")
    # 每行前面都有 HH:MM:SS 时间戳，便于还原先后顺序
    assert all(len(line.split(" ")[0].split(":")) == 3 for line in text.splitlines())


def test_run_log_closes_idempotently(tmp_path):
    log = RunLog(tmp_path / "logs", "RUN456")
    log.write("内容")
    path = log.close()
    assert log.closed
    log.write("关闭后写入被忽略")      # 不抛异常
    log.close()
    assert path.read_text(encoding="utf-8").count("内容") == 1


def test_prune_keeps_newest_logs(tmp_path):
    directory = tmp_path / "logs"
    directory.mkdir()
    for index in range(6):
        (directory / f"RUN-2026010{index}-000000-A{index}.log").write_text("x", encoding="utf-8")
    removed = prune_logs(directory, keep=3)
    assert removed == 3
    remaining = sorted(path.name for path in directory.glob("RUN-*.log"))
    assert len(remaining) == 3
    assert remaining[0].startswith("RUN-20260103")


class _FakeText:
    def __init__(self) -> None:
        self.content = ""

    def configure(self, **_kwargs) -> None:
        pass

    def insert(self, _index, text: str) -> None:
        self.content += text

    def see(self, _index) -> None:
        pass


def test_app_writes_execution_log_to_disk(tmp_path, monkeypatch):
    """界面“执行日志”要同步落盘，停止/失败后仍能翻到完整过程。"""
    import printcase.app as app_module

    monkeypatch.setattr(app_module, "data_dir", lambda: tmp_path)

    class MiniApp(app_module.PrintCaseApp):
        def __init__(self) -> None:      # 不创建 Tk 窗口，只验证日志接线
            self.log_text = _FakeText()
            self.run_log = None
            self.log_path = None
            self.run_id = "APPTEST"

    app = MiniApp()
    app._open_run_log()
    assert app.run_log is not None
    log_path = app.run_log.path
    app._log("用例开始")
    app._log("下发失败：演示")
    app._close_run_log()

    assert app.run_log is None and app.log_path == log_path
    text = log_path.read_text(encoding="utf-8")
    assert "用例开始" in text and "下发失败：演示" in text
    assert "本次执行日志文件" in text
    assert "用例开始" in app.log_text.content
