"""端到端验证：注入“UIA 全量枚举失效”，跑一条真实用例，确认下发链路仍可用。

这模拟用户现场（程序运行一段时间后 UIA 枚举读不到属性页窗口）。
只打印 1 条用例（TC-010，页码 1-2）。
"""
from __future__ import annotations

import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import printcase.driver_ui as driver_ui  # noqa: E402
from printcase.audit import AuditStore  # noqa: E402
from printcase.excel_reader import read_cases  # noqa: E402
from printcase.models import CaseStatus  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
CASE_ID = sys.argv[1] if len(sys.argv) > 1 else "TC-010"

real_desktop = driver_ui._desktop()


class BlindDesktop:
    """UIA 全量枚举抛异常；按句柄取窗口仍可用（真实场景下的典型状态）。"""

    def __init__(self, real) -> None:
        self._real = real
        self.blocked = 0

    def windows(self):
        self.blocked += 1
        raise RuntimeError("模拟 UIA 全量枚举失效")

    def window(self, *args, **kwargs):
        return self._real.window(*args, **kwargs)


blind = BlindDesktop(real_desktop)
driver_ui._desktop = lambda: blind  # 让 open_preferences 用“瞎眼”桌面

cases, _ = read_cases(ROOT / "templates" / "PrintCase_Template.xlsx")
case = next(c for c in cases if c.case_id == CASE_ID)
print(f"=== 注入 UIA 枚举失效，真实下发 {CASE_ID}（pages={case.pages!r}）===")

spooler = WindowsSpooler()
temp_dir = Path(tempfile.mkdtemp(prefix="printcase_blind_"))
runner = PrintRunner(spooler, AuditStore(temp_dir / "blind.sqlite3"), queue_timeout=40)

started = time.monotonic()
results = runner.run_cases([case], PRINTER, f"BLIND-{int(time.time())}",
                           lambda kind, payload: print(f"{time.monotonic() - started:6.1f}s "
                                                       f"[{kind}] {payload}", flush=True))

print("-" * 72)
for result in results:
    print(f"  {result.case_id}: {result.status.value} / Job#{result.job_id or '未关联'}")
    print(f"  {result.message[:160]}")
print(f"  UIA 全量枚举被拒绝 {blind.blocked} 次；总耗时 {time.monotonic() - started:.1f}s")

ok = bool(results) and results[0].status not in {CaseStatus.FAILED, CaseStatus.PREFLIGHT_FAILED}
print("\n结论：", "通过（属性设置与下发均成功）" if ok else "失败")
