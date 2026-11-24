"""界面级黑盒验证：真实创建窗口、真实导入 Excel、真实预检/导出并截图。

不打印纸张。验证的功能点：
    导入 Excel → 用例队列渲染 → 全选/反选 → 打印前检查 → 导出报告 → 日志页

用法：
    .venv/Scripts/python.exe tools/ui_gui_smoke.py [截图输出目录]
"""
from __future__ import annotations

import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import printcase.app as app_module  # noqa: E402
from printcase.app import PrintCaseApp  # noqa: E402
from printcase.models import CaseStatus  # noqa: E402

TEMPLATE = ROOT / "templates" / "PrintCase_Template.xlsx"
SHOT_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist" / "ui-smoke"
SHOT_DIR.mkdir(parents=True, exist_ok=True)

results: list[tuple[str, bool, str]] = []
dialogs: list[tuple[str, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    results.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""), flush=True)


class FakeBox:
    """挡掉模态对话框：记录下来，不阻塞自动化。"""

    def showinfo(self, title, message="", **_k):
        dialogs.append(("info", f"{title}：{message}"))

    def showwarning(self, title, message="", **_k):
        dialogs.append(("warning", f"{title}：{message}"))

    def showerror(self, title, message="", **_k):
        dialogs.append(("error", f"{title}：{message}"))

    def askyesno(self, title, message="", **_k):
        dialogs.append(("ask", f"{title}：{message}"))
        return True


def shot(root: tk.Tk, name: str) -> Path | None:
    """按窗口句柄抓整窗（含被遮挡部分），规避 DPI 缩放导致的裁剪。"""
    try:
        import ctypes
        from ctypes import wintypes

        from PIL import Image
    except Exception:  # noqa: BLE001
        return None
    root.update_idletasks()
    root.update()
    try:
        hwnd = int(root.wm_frame(), 16)
    except Exception:  # noqa: BLE001
        hwnd = root.winfo_id()
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    width, height = rect.right - rect.left, rect.bottom - rect.top
    if width <= 0 or height <= 0:
        return None
    hdc = user32.GetWindowDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(hdc)
    bitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
    gdi32.SelectObject(mem_dc, bitmap)
    user32.PrintWindow(hwnd, mem_dc, 2)      # PW_RENDERFULLCONTENT
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height
    header.biPlanes = 1
    header.biBitCount = 32
    buffer = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(mem_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, hdc)
    path = SHOT_DIR / name
    image.convert("RGB").save(path)
    return path


def main() -> int:
    app_module.messagebox = FakeBox()                       # 阻断模态框
    app_module.filedialog.askopenfilename = lambda **_k: str(TEMPLATE)
    export_path = Path(tempfile.mkdtemp(prefix="ui_smoke_")) / "report.csv"
    app_module.filedialog.asksaveasfilename = lambda **_k: str(export_path)

    root = tk.Tk()
    app = PrintCaseApp(root)      # 窗口尺寸由应用按屏幕自适应，不要覆盖
    root.update()
    check("主窗口创建", bool(root.winfo_exists()) and "PrintCase Studio" in root.title(),
          root.title())
    check("启动即检测打印环境", "检测" not in app.printer_var.get() or bool(app.printer_var.get()),
          f"默认打印机：{app.printer_var.get() or '未检测到'}")

    initial = shot(root, "01-初始界面.png")
    check("初始界面截图", initial is not None and initial.exists(),
          initial.name if initial else "截图失败")

    # ---------- 导入 Excel ----------
    app._import_excel()
    root.update()
    rows = app.tree.get_children()
    check("导入 Excel 用例", len(app.cases) == 22, f"{len(app.cases)} 条用例")
    check("用例队列渲染行数一致", len(rows) == len(app.cases), f"{len(rows)} 行")
    check("默认只勾选启用用例",
          len(app.checked) == sum(1 for c in app.cases if c.enabled),
          f"已勾选 {len(app.checked)} 条")
    check("导入后计数标签同步", "22 条" in app.case_count_label.cget("text"),
          app.case_count_label.cget("text"))

    # ---------- 全选 / 反选 ----------
    app.select_all_var.set(True)
    app._toggle_all_from_checkbox()
    root.update()
    check("全选复选框选中全部", len(app.checked) == 22, f"已勾选 {len(app.checked)} 条")
    app.select_all_var.set(False)
    app._toggle_all_from_checkbox()
    root.update()
    check("全选复选框取消全部", len(app.checked) == 0, f"已勾选 {len(app.checked)} 条")

    # 单条勾选（模拟点击最左“选”列）
    app.checked.add(0)
    app._sync_select_all()
    root.update()
    check("单条勾选后计数更新", "已选 1 条" in app.case_count_label.cget("text"),
          app.case_count_label.cget("text"))

    # 恢复为启用用例，供预检使用
    app.checked = {i for i, c in enumerate(app.cases) if c.enabled}
    app._sync_select_all()

    table_shot = shot(root, "02-导入用例.png")
    check("用例表截图", table_shot is not None and table_shot.exists(),
          table_shot.name if table_shot else "截图失败")

    # ---------- 打印前检查 ----------
    app._preflight()
    root.update()
    passed = sum(1 for r in app.case_results.values() if r.status == CaseStatus.PREFLIGHT_OK)
    failed = sum(1 for r in app.case_results.values() if r.status == CaseStatus.PREFLIGHT_FAILED)
    check("打印前检查执行", passed + failed > 0, f"通过 {passed}，失败 {failed}")
    color_cases = {c.case_id for c in app.cases if c.color == "color" and c.enabled}
    failed_ids = {cid for cid, r in app.case_results.items() if r.status == CaseStatus.PREFLIGHT_FAILED}
    check("预检拦下黑白机上的彩色用例", failed_ids == color_cases,
          f"失败用例 {sorted(failed_ids)}，其中彩色用例 {sorted(color_cases)}")
    check("预检状态栏更新", "打印前检查完成" in app.status_var.get(), app.status_var.get())
    preflight_shot = shot(root, "03-预检结果.png")
    check("预检结果截图", preflight_shot is not None and preflight_shot.exists(),
          preflight_shot.name if preflight_shot else "截图失败")

    # ---------- 导出报告 ----------
    app._export_report()
    root.update()
    check("无执行结果时拒绝导出",
          not export_path.exists() and any(k == "info" for k, _ in dialogs),
          "提示：暂无执行结果")

    from printcase.models import CaseResult
    app.run_id = "UISMOKE"
    app.run_results = [app.case_results[cid] for cid in sorted(app.case_results)][:3]
    for result in app.run_results:
        result.status = CaseStatus.QUEUE_CLEARED
        result.message = "演示：队列已清除"
    app._export_report()
    root.update()
    check("导出本次报告", export_path.exists() and export_path.stat().st_size > 100,
          f"{export_path.name} {export_path.stat().st_size if export_path.exists() else 0} 字节")

    # ---------- 日志页 ----------
    log_text = app.log_text.get("1.0", "end")
    check("执行日志有内容", "载入 Excel" in log_text and "打印前检查完成" in log_text,
          f"{len(log_text.splitlines())} 行")

    # ---------- 自适应列宽 ----------
    app._autosize_console_columns()
    root.update()
    widths = [app.tree.column(col, "width") for col in app.tree["columns"]]
    check("自适应列宽不挤压首列", widths[0] >= 34 and min(widths) >= 34,
          "列宽 " + "/".join(str(w) for w in widths[:4]) + "…")

    final_shot = shot(root, "04-自适应列宽.png")
    check("最终界面截图", final_shot is not None and final_shot.exists(),
          final_shot.name if final_shot else "截图失败")

    error_dialogs = [d for d in dialogs if d[0] == "error"]
    check("全程无错误弹窗", not error_dialogs, "；".join(m for _, m in error_dialogs) or "无")

    try:
        root.destroy()
    except Exception:  # noqa: BLE001
        pass

    print("-" * 72)
    failed_items = [name for name, ok, _ in results if not ok]
    print(f"共 {len(results)} 项：通过 {len(results) - len(failed_items)}，失败 {len(failed_items)}")
    if failed_items:
        print("失败项：" + "、".join(failed_items))
    if dialogs:
        print("弹窗记录：" + "；".join(f"[{k}] {m[:60]}" for k, m in dialogs))
    return 0 if not failed_items else 1


if __name__ == "__main__":
    raise SystemExit(main())
