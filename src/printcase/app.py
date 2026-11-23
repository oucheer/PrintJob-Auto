from __future__ import annotations

import os
import queue
import threading
import uuid
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .audit import AuditStore
from .capabilities import validate_driver_capabilities
from .excel_reader import read_cases
from .models import CaseResult, CaseStatus, PrintCase
from .office import OFFICE_EXTENSIONS, probe_document_converter
from .runner import PrintRunner
from .spooler import WindowsSpooler
from .appinfo import app_base_dir
from .validator import validate_case


BG = "#F3F6FA"
PANEL = "#FFFFFF"
INK = "#17283D"
MUTED = "#65758B"
ACCENT = "#2367D1"
GREEN = "#16845B"
RED = "#C74747"
AMBER = "#B97817"


def data_dir() -> Path:
    root = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    path = Path(root) / "PrintCaseStudio" if root else Path.home() / ".printcase-studio"
    path.mkdir(parents=True, exist_ok=True)
    return path


class PrintCaseApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"PrintCase Studio · Excel 打印用例执行平台 v{__version__}")
        self.root.geometry("1380x860")
        self.root.minsize(1100, 720)
        self.root.configure(bg=BG)

        self.cases: list[PrintCase] = []
        self.checked: set[int] = set()
        self.case_results: dict[str, CaseResult] = {}
        self.run_results: list[CaseResult] = []
        self.run_id = ""
        self.case_file: Path | None = None
        self.printer: WindowsSpooler | None = None
        self.runner: PrintRunner | None = None
        self.event_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.running = False
        self.audit = AuditStore(data_dir() / "history.sqlite3")

        self._setup_style()
        self._build_ui()
        self._refresh_environment()
        self._poll_events()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", background="white", fieldbackground="white", foreground=INK, rowheight=34, borderwidth=0, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", background="#EAF0F7", foreground="#44566D", font=("Microsoft YaHei UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#DDEBFF")], foreground=[("selected", INK)])
        style.configure("TCombobox", padding=6)
        style.configure("TButton", padding=(12, 8), font=("Microsoft YaHei UI", 9))
        style.configure("Primary.TButton", background=ACCENT, foreground="white", bordercolor=ACCENT)
        style.map("Primary.TButton", background=[("active", "#1756B8"), ("disabled", "#A7B9D2")], foreground=[("disabled", "#EEF2F8")])
        style.configure("Stop.TButton", background="#FCEDEE", foreground=RED)

    # 用例队列表格列（与 Excel 用例表逐列对应，最左为勾选列）
    TREE_COLUMNS = [
        ("check", "选", 34, "center"),
        ("case_id", "用例 ID", 92, "w"),
        ("file", "打印文件", 150, "w"),
        ("pages", "页码", 62, "w"),
        ("copies", "份数", 48, "center"),
        ("paper", "纸张尺寸", 132, "w"),
        ("mixed", "混合尺寸", 62, "center"),
        ("orientation", "方向", 52, "center"),
        ("color", "颜色", 52, "center"),
        ("duplex", "单双面", 70, "center"),
        ("bin", "纸盒", 84, "w"),
        ("collate", "打印顺序", 66, "center"),
        ("quality", "打印质量", 66, "center"),
        ("media", "纸张类型", 128, "w"),
        ("skip_blank", "跳过空白页", 68, "center"),
        ("borderless", "无边距", 52, "center"),
        ("separator", "分隔页", 62, "center"),
        ("sep_bin", "分隔页来源", 84, "w"),
        ("binding_loc", "装订位置", 68, "center"),
        ("binding_method", "装订方式", 68, "center"),
        ("ext_attrs", "扩展属性", 72, "center"),
        ("status", "执行状态", 148, "w"),
        ("time", "耗时", 56, "w"),
        ("message", "结果 / 诊断", 230, "w"),
    ]


    NAV_ITEMS = [
        ("console", "🖨", "打印执行"),
        ("log", "📋", "执行日志"),
        ("history", "🗂", "执行历史"),
        ("help", "📖", "使用帮助"),
    ]

    def _build_ui(self) -> None:
        # ---------- 左侧 SaaS 导航栏 ----------
        sidebar = tk.Frame(self.root, bg="#0D1B2E", width=176)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="PRINTCASE", bg="#0D1B2E", fg="white", font=("Arial", 15, "bold")).pack(anchor="w", padx=18, pady=(18, 0))
        tk.Label(sidebar, text="STUDIO · 打印用例控制台", bg="#0D1B2E", fg="#9FB3CB", font=("Microsoft YaHei UI", 8)).pack(anchor="w", padx=18, pady=(2, 16))

        self.nav_items: dict[str, tk.Label] = {}
        for key, icon, text in self.NAV_ITEMS:
            item = tk.Label(
                sidebar, text=f"{icon}   {text}", bg="#0D1B2E", fg="#9FB3CB",
                font=("Microsoft YaHei UI", 10, "bold"), anchor="w", padx=22, pady=10, cursor="hand2",
            )
            item.pack(fill="x")
            item.bind("<Button-1>", lambda _event, k=key: self._show_page(k))
            item.bind("<Enter>", lambda _event, w=item: w.configure(fg="white") if w.cget("bg") == "#0D1B2E" else None)
            item.bind("<Leave>", lambda _event, w=item: w.configure(fg="#9FB3CB") if w.cget("bg") == "#0D1B2E" else None)
            self.nav_items[key] = item

        tk.Label(sidebar, text=f"v{__version__}", bg="#0D1B2E", fg="#5A6E86", font=("Microsoft YaHei UI", 8)).pack(side="bottom", anchor="w", padx=18, pady=10)

        # ---------- 右侧主区域 ----------
        main_area = tk.Frame(self.root, bg=BG)
        main_area.pack(side="right", fill="both", expand=True)

        header = tk.Frame(main_area, bg="#10243D", height=52)
        header.pack(fill="x")
        header.pack_propagate(False)
        self.page_title = tk.Label(header, text="打印执行", bg="#10243D", fg="white", font=("Microsoft YaHei UI", 12, "bold"))
        self.page_title.pack(side="left", padx=22)
        self.env_label = tk.Label(header, text="检查运行环境…", bg="#203A59", fg="#DDE9F6", padx=12, pady=6, font=("Microsoft YaHei UI", 9))
        self.env_label.pack(side="right", padx=(0, 16))
        self.engine_label = tk.Label(header, text="打印引擎：检测中…", bg="#203A59", fg="#DDE9F6", padx=12, pady=6, font=("Microsoft YaHei UI", 9))
        self.engine_label.pack(side="right", padx=(0, 8))
        self.converter_label = tk.Label(header, text="文档转换：检测中…", bg="#203A59", fg="#DDE9F6", padx=12, pady=6, font=("Microsoft YaHei UI", 9))
        self.converter_label.pack(side="right", padx=(0, 8))

        pages = tk.Frame(main_area, bg=BG)
        pages.pack(fill="both", expand=True)
        pages.rowconfigure(0, weight=1)
        pages.columnconfigure(0, weight=1)
        self.page_frames: dict[str, tk.Frame] = {}
        for key, _icon, _text in self.NAV_ITEMS:
            frame = tk.Frame(pages, bg=BG)
            frame.grid(row=0, column=0, sticky="nsew")
            self.page_frames[key] = frame

        self._build_console_page(self.page_frames["console"])
        self._build_log_page(self.page_frames["log"])
        self._build_history_page(self.page_frames["history"])
        self._build_help_page(self.page_frames["help"])
        self._show_page("console")

    def _show_page(self, key: str) -> None:
        """左侧导航切换右侧功能页，并同步高亮。"""
        self.page_frames[key].tkraise()
        titles = {"console": "打印执行", "log": "执行日志", "history": "执行历史", "help": "使用帮助"}
        self.page_title.configure(text=titles.get(key, key))
        for name, item in self.nav_items.items():
            active = name == key
            item.configure(bg=ACCENT if active else "#0D1B2E", fg="white" if active else "#9FB3CB")
        if key == "history":
            self._refresh_history()
        if key == "help":
            self._load_help_text()

    # ---------- 页面一：打印执行 ----------
    def _build_console_page(self, page: tk.Frame) -> None:
        body = tk.Frame(page, bg=BG, padx=16, pady=10)
        body.pack(fill="both", expand=True)

        self.kpi_frame = tk.Frame(body, bg=BG)
        self.kpi_frame.pack(fill="x", pady=(0, 8))
        self.kpi_labels: dict[str, tk.Label] = {}
        for key, title, value in [("total", "识别用例", "0"), ("ready", "检查通过", "0"), ("accepted", "队列/驱动反馈", "0"), ("attention", "失败 / 待核实", "0")]:
            card = tk.Frame(self.kpi_frame, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1, padx=14, pady=8)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(card, text=title, bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(anchor="w")
            value_label = tk.Label(card, text=value, bg=PANEL, fg=INK, font=("Arial", 18, "bold"))
            value_label.pack(anchor="w", pady=(2, 0))
            self.kpi_labels[key] = value_label

        # 工具行 1：用例文件操作 + 已载入文件
        bar1 = tk.Frame(body, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1)
        bar1.pack(fill="x", pady=(0, 6))
        self.btn_import = ttk.Button(bar1, text="导入 Excel 用例", command=self._import_excel)
        self.btn_import.pack(side="left", padx=(10, 6), pady=7)
        ttk.Button(bar1, text="打开用例模板", command=self._open_template).pack(side="left", pady=7)
        self.file_label = tk.Label(bar1, text="尚未载入 Excel 用例文件 · 请先导入模板或现有用例", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 8))
        self.file_label.pack(side="left", padx=(12, 0))

        # 工具行 2：先打包右侧执行按钮（保证任何宽度下都可见），再打包左侧打印机选择
        toolbar = tk.Frame(body, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1)
        toolbar.pack(fill="x", pady=(0, 8))
        self.btn_stop = ttk.Button(toolbar, text="停止", command=self._stop_run, style="Stop.TButton", state="disabled")
        self.btn_stop.pack(side="right", padx=(6, 10), pady=8)
        self.btn_preflight = ttk.Button(toolbar, text="执行打印前检查", command=self._preflight)
        self.btn_preflight.pack(side="right", padx=(6, 0), pady=8)
        self.btn_run_selected = ttk.Button(toolbar, text="执行勾选用例", command=lambda: self._start_run(selected_only=True), style="Primary.TButton")
        self.btn_run_selected.pack(side="right", padx=(6, 0), pady=8)
        self.btn_run_all = ttk.Button(toolbar, text="执行全部", command=lambda: self._start_run(selected_only=False), style="Primary.TButton")
        self.btn_run_all.pack(side="right", padx=(6, 0), pady=8)
        tk.Label(toolbar, text="默认打印机", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(10, 2), pady=8)
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(toolbar, textvariable=self.printer_var, width=24, state="readonly")
        self.printer_combo.pack(side="left", padx=(6, 6), pady=8)
        ttk.Button(toolbar, text="刷新", command=self._refresh_printers).pack(side="left")
        self.driver_label = tk.Label(toolbar, text="驱动属性页：未自检", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 8))
        self.driver_label.pack(side="left", padx=(12, 0), pady=8)
        ttk.Button(toolbar, text="自检驱动属性页", command=self._check_driver_ui).pack(side="left", padx=(8, 0))

        # 用例队列占满整页宽度
        left = body
        table_header = tk.Frame(left, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1, padx=12, pady=8)
        table_header.pack(fill="x")
        tk.Label(table_header, text="用例队列", bg=PANEL, fg=INK, font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
        self.case_count_label = tk.Label(table_header, text="0 条", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9))
        self.case_count_label.pack(side="left", padx=(8, 0))
        ttk.Button(table_header, text="导出本次报告", command=self._export_report).pack(side="right")

        tree_area = tk.Frame(left, bg=PANEL)
        tree_area.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        tree_area.columnconfigure(0, weight=1)
        tree_area.rowconfigure(0, weight=1)
        columns = tuple(col[0] for col in self.TREE_COLUMNS)
        self.tree = ttk.Treeview(tree_area, columns=columns, show="headings", selectmode="extended")
        for key, text, width, anchor in self.TREE_COLUMNS:
            if key == "check":
                self.tree.heading(key, text=text, command=self._toggle_all_checks)
            else:
                self.tree.heading(key, text=text)
            self.tree.column(key, width=width, minwidth=34, anchor=anchor, stretch=(key in {"file", "media", "message", "paper"}))
        yscroll = ttk.Scrollbar(tree_area, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(tree_area, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("bad", foreground=RED)
        self.tree.tag_configure("good", foreground=GREEN)
        self.tree.tag_configure("muted", foreground=MUTED)
        self.tree.bind("<Button-1>", self._on_tree_click)

        footer = tk.Frame(body, bg=BG)
        footer.pack(fill="x", pady=(8, 0))
        tk.Label(footer, text="安全策略：不自动重复提交不确定作业 · 队列完成不等于物理出纸", bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 8)).pack(side="left")
        self.status_var = tk.StringVar(value="就绪 · 请导入 Excel 用例")
        tk.Label(footer, textvariable=self.status_var, bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 8)).pack(side="right")

    # ---------- 页面二：执行日志 ----------
    def _build_log_page(self, page: tk.Frame) -> None:
        bar = tk.Frame(page, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1)
        bar.pack(fill="x")
        tk.Label(bar, text="实时事件 · 队列反馈 · 故障提示", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(12, 0), pady=8)
        ttk.Button(bar, text="清空日志", command=self._clear_log).pack(side="right", padx=10, pady=6)
        self.log_text = tk.Text(page, bg="#F7F9FC", fg="#44566D", wrap="word", relief="flat", font=("Consolas", 10), padx=14, pady=12, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ---------- 页面三：执行历史 ----------
    def _build_history_page(self, page: tk.Frame) -> None:
        bar = tk.Frame(page, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1)
        bar.pack(fill="x")
        tk.Label(bar, text="历史执行记录（按时间倒序，最多 500 条）", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(12, 0), pady=8)
        ttk.Button(bar, text="刷新历史", command=self._refresh_history).pack(side="right", padx=10, pady=6)

        area = tk.Frame(page, bg=PANEL)
        area.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        area.columnconfigure(0, weight=1)
        area.rowconfigure(0, weight=1)
        columns = ("time", "run", "case", "status", "printer", "job", "elapsed", "file", "message")
        headings = {"time": "开始时间", "run": "RUN ID", "case": "用例 ID", "status": "执行状态", "printer": "打印机", "job": "Job ID", "elapsed": "耗时(s)", "file": "源文件", "message": "结果 / 诊断"}
        widths = {"time": 130, "run": 90, "case": 92, "status": 140, "printer": 130, "job": 60, "elapsed": 60, "file": 150, "message": 260}
        self.history_tree = ttk.Treeview(area, columns=columns, show="headings", selectmode="extended")
        for key in columns:
            self.history_tree.heading(key, text=headings[key])
            self.history_tree.column(key, width=widths[key], minwidth=40, anchor="w", stretch=(key in {"message", "file"}))
        yscroll = ttk.Scrollbar(area, orient="vertical", command=self.history_tree.yview)
        xscroll = ttk.Scrollbar(area, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.history_tree.tag_configure("bad", foreground=RED)
        self.history_tree.tag_configure("good", foreground=GREEN)
        self.history_tree.tag_configure("muted", foreground=MUTED)

    def _refresh_history(self) -> None:
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        try:
            rows = self.audit.recent(500)
        except Exception as exc:
            self._log(f"读取执行历史失败：{exc}")
            return
        for row in rows:
            tag = ""
            if any(word in (row["status"] or "") for word in ("失败", "超时", "待核实")):
                tag = "bad"
            elif any(word in (row["status"] or "") for word in ("已打印", "队列")):
                tag = "good"
            self.history_tree.insert("", "end", values=(
                row.get("started_at") or "—",
                row.get("run_id") or "",
                row.get("case_id") or "",
                row.get("status") or "",
                row.get("printer") or "",
                row.get("job_id") if row.get("job_id") is not None else "—",
                row.get("elapsed_seconds") if row.get("elapsed_seconds") is not None else "—",
                Path(row.get("source_file") or "").name or "—",
                (row.get("message") or "")[:200],
            ), tags=(tag,))

    # ---------- 页面三：使用帮助 ----------
    def _build_help_page(self, page: tk.Frame) -> None:
        bar = tk.Frame(page, bg=PANEL, highlightbackground="#E2E9F1", highlightthickness=1)
        bar.pack(fill="x")
        tk.Label(bar, text="使用说明（README）", bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(12, 0), pady=8)
        ttk.Button(bar, text="打开完整 README", command=self._open_readme).pack(side="right", padx=10, pady=6)
        self.help_text = tk.Text(page, bg="#F7F9FC", fg="#44566D", wrap="word", relief="flat", font=("Microsoft YaHei UI", 9), padx=16, pady=12, state="disabled")
        self.help_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _load_help_text(self) -> None:
        path = app_base_dir() / "README_CN.md"
        if not path.exists():
            path = Path(__file__).resolve().parents[2] / "README_CN.md"
        try:
            content = path.read_text(encoding="utf-8")
        except Exception as exc:
            content = f"无法读取 README_CN.md：{exc}"
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", "end")
        self.help_text.insert("1.0", content)
        self.help_text.configure(state="disabled")
        self.help_text.mark_set("insert", "1.0")

    def _refresh_environment(self) -> None:
        self._update_engine_label()
        self._update_converter_label()
        self._refresh_printers()

    def _update_converter_label(self) -> None:
        """文档转换组件探测：Word → WPS → LibreOffice 自动回退。"""
        try:
            ok, detail = probe_document_converter(".docx")
            if ok:
                self.converter_label.configure(text=f"文档转换：{detail}", bg="#1D654B", fg="#DDF4E8")
            else:
                self.converter_label.configure(text=f"文档转换：{detail}", bg="#7F3030", fg="#FFE6E6")
        except Exception as exc:
            self.converter_label.configure(text=f"文档转换：检测失败（{exc}）", bg="#7F3030", fg="#FFE6E6")

    def _update_engine_label(self) -> None:
        """内置直打引擎（GDI + PDF 渲染）自检，无需外部打印组件。"""
        try:
            from .direct_print import engine_available

            ok, detail = engine_available()
        except Exception as exc:  # noqa: WPS421
            ok, detail = False, f"自检异常：{exc}"
        if ok:
            self.engine_label.configure(text=f"打印引擎：{detail}", bg="#1D654B", fg="#DDF4E8")
        else:
            self.engine_label.configure(text=f"打印引擎：{detail}", bg="#7F3030", fg="#FFE6E6")

    def _check_driver_ui(self) -> None:
        """自检：当前默认打印机的驱动“打印首选项”能否被自动打开（不修改任何设置）。"""
        printer = self.printer_var.get().strip()
        if not printer:
            messagebox.showinfo("未选择打印机", "请先选择默认打印机。")
            return
        self.driver_label.configure(text="驱动属性页：自检中…", fg=MUTED)
        self.status_var.set(f"正在自检 {printer} 的驱动属性页…")

        def probe() -> None:
            try:
                from .driver_ui import available

                ok, detail = available(printer)
            except Exception as exc:  # noqa: WPS421
                ok, detail = False, f"自检异常：{type(exc).__name__}: {exc}"
            self.event_queue.put(("driver_check", (ok, detail)))

        threading.Thread(target=probe, name="driver-ui-check", daemon=True).start()

    def _apply_driver_check(self, ok: bool, detail: str) -> None:
        if ok:
            self.driver_label.configure(text=f"驱动属性页：{detail}", fg=GREEN)
            self.status_var.set(f"驱动属性页自检通过 · {detail}")
            self._log(f"驱动属性页自检通过：{detail}")
            return
        self.driver_label.configure(text="驱动属性页：不可用", fg=RED)
        self.status_var.set("驱动属性页自检未通过")
        self._log(f"驱动属性页自检未通过：{detail}")
        messagebox.showwarning(
            "驱动属性页不可用",
            f"{detail}\n\n"
            "本软件通过驱动自己的“打印首选项”设置纸张/介质等属性；"
            "打不开属性页时这些用例会被判定失败且不下发。\n"
            "请确认驱动已正确安装，并能手工打开“打印首选项”。",
        )

    def _refresh_printers(self) -> None:
        try:
            if self.printer is None:
                self.printer = WindowsSpooler()
            names = self.printer.list_printers()
            self.printer_combo["values"] = names
            current = self.printer_var.get()
            if current not in names:
                # 优先选中 Windows 默认打印机，避免默认落到“Fax/PDF”这类虚拟打印机
                try:
                    default_printer = self.printer.win32print.GetDefaultPrinter()
                except Exception:  # noqa: WPS420
                    default_printer = ""
                self.printer_var.set(default_printer if default_printer in names else (names[0] if names else ""))
            self.env_label.configure(text=f"Windows 打印服务可用 · {len(names)} 台打印机", fg="#DDF4E8", bg="#1D654B")
            if not names:
                self._log("系统未发现本地或已连接打印机。")
        except Exception as exc:
            self.printer = None
            self.printer_combo["values"] = []
            self.env_label.configure(text="打印服务不可用", fg="#FFE6E6", bg="#7F3030")
            self._log(f"打印服务初始化失败：{exc}")

    def _import_excel(self) -> None:
        filename = filedialog.askopenfilename(title="选择 Excel 打印用例", filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not filename:
            return
        try:
            cases, warnings = read_cases(filename)
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        self.cases = cases
        self.checked = {i for i, c in enumerate(cases) if c.enabled}
        self.case_results.clear()
        self.run_results = []
        self.case_file = Path(filename)
        self.file_label.configure(text=str(self.case_file))
        self.case_count_label.configure(text=f"{len(cases)} 条")
        self._render_cases()
        self._update_kpis()
        self.status_var.set(f"已识别 {len(cases)} 条用例")
        self._log(f"载入 Excel：{self.case_file.name}，识别 {len(cases)} 条用例。")
        for warning in warnings:
            self._log("提示：" + warning)
        disabled = sum(1 for c in cases if not c.enabled)
        if disabled:
            self._log(f"其中 {disabled} 条用例未启用，将跳过。")

    _CASE_LABELS = {
        "orientation": {"portrait": "纵向", "landscape": "横向"},
        "color": {"color": "彩色", "monochrome": "黑白"},
        "duplex": {"simplex": "单面", "duplex": "双面", "duplexlong": "双面长边", "duplexshort": "双面短边"},
        "collate": {"yes": "逐份", "no": "非逐份", "": "驱动默认"},
        "separator": {"none": "无", "before": "前", "after": "后", "both": "前后", "each_copy": "份数之间"},
        "binding_loc": {"none": "无", "left": "左", "right": "右", "top": "上", "bottom": "下", "longedge": "长边", "shortedge": "短边"},
        "binding_method": {"none": "无", "staple": "订书钉", "saddle": "骑马钉", "punch": "打孔", "glue": "胶装"},
    }

    def _render_cases(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        labels = self._CASE_LABELS
        for index, case in enumerate(self.cases):
            result = self.case_results.get(case.case_id)
            status = "未启用" if not case.enabled else CaseStatus.PENDING.value
            elapsed = "—"
            message = case.notes
            tag = "muted" if not case.enabled else ""
            if result:
                status = result.status.value
                elapsed = f"{result.elapsed_seconds:.1f}s" if result.elapsed_seconds is not None else "—"
                message = result.message or case.notes
                if result.status in {CaseStatus.FAILED, CaseStatus.PREFLIGHT_FAILED, CaseStatus.TIMEOUT_QUEUED, CaseStatus.SUBMITTED_UNVERIFIED}:
                    tag = "bad"
                elif result.status in {CaseStatus.DRIVER_PRINTED, CaseStatus.QUEUE_CLEARED, CaseStatus.QUEUED}:
                    tag = "good"
            values = (
                "☑" if index in self.checked else "☐",
                case.case_id,
                Path(case.file_path).name if case.file_path else "（路径为空）",
                case.pages or "全部",
                case.copies,
                case.paper or "自动",
                "Y" if case.mixed_sizes == "yes" else "N",
                labels["orientation"].get(case.orientation, case.orientation),
                labels["color"].get(case.color, case.color),
                labels["duplex"].get(case.duplex, case.duplex),
                case.bin or "自动选择",
                labels["collate"].get(case.collate, case.collate or "驱动默认"),
                case.print_quality or "默认",
                case.media_type or "驱动默认",
                "Y" if case.skip_blank == "yes" else "N",
                "Y" if case.borderless == "yes" else "N",
                labels["separator"].get(case.separator_page, case.separator_page),
                case.separator_bin or "—",
                labels["binding_loc"].get(case.binding_location, case.binding_location),
                labels["binding_method"].get(case.binding_method, case.binding_method),
                f"{len(case.extended_attrs)} 项" if case.extended_attrs else "—",
                status,
                elapsed,
                message,
            )
            self.tree.insert("", "end", iid=str(index), values=values, tags=(tag,))

    def _on_tree_click(self, event) -> str | None:
        """点击最左“选”列切换勾选状态。"""
        if self.tree.identify("region", event.x, event.y) != "cell":
            return None
        if self.tree.identify_column(event.x) != "#1":
            return None
        row = self.tree.identify_row(event.y)
        if not row:
            return None
        index = int(row)
        if index in self.checked:
            self.checked.discard(index)
        else:
            self.checked.add(index)
        self._render_cases()
        return "break"

    def _toggle_all_checks(self) -> None:
        """点击“选”列表头：全部勾选 / 全部取消。"""
        if len(self.checked) == len(self.cases) and self.cases:
            self.checked.clear()
        else:
            self.checked = set(range(len(self.cases)))
        self._render_cases()

    def _preflight(self) -> None:
        if not self.cases:
            messagebox.showinfo("尚无用例", "请先导入 Excel 用例文件。")
            return
        if self.printer is None:
            messagebox.showerror("打印服务不可用", "请在 Windows 10/11 上运行，并确认打印服务可用。")
            return
        printers = set(self.printer.list_printers())
        fallback = self.printer_var.get().strip()
        caps_cache: dict[str, dict | None] = {}
        id_counts: dict[str, int] = {}
        for case in self.cases:
            id_counts[case.case_id] = id_counts.get(case.case_id, 0) + 1
        ok = 0
        bad = 0
        for case in self.cases:
            if not case.enabled:
                continue
            target_printer = case.printer.strip() or fallback
            errors = validate_case(case, printers, fallback)
            source = case.resolved_path or (Path(case.file_path) if case.file_path else None)
            if source is not None and source.suffix.lower() in OFFICE_EXTENSIONS:
                probe = probe_document_converter(source.suffix)
                if not probe[0]:
                    errors.append(f"打印前检查：{source.suffix} 文件缺少可用转换组件（{probe[1]}）。")
            if target_printer and target_printer in printers:
                if target_printer not in caps_cache:
                    try:
                        caps_cache[target_printer] = self.printer.get_capabilities(target_printer)
                    except Exception as exc:
                        caps_cache[target_printer] = None
                        self._log(f"无法读取 {target_printer} 的驱动能力，跳过该项能力核验：{exc}")
                errors.extend(validate_driver_capabilities(case, caps_cache[target_printer]))
            if id_counts.get(case.case_id, 0) > 1:
                errors.append("用例 ID 重复；请保证每条用例 ID 唯一")
            if errors:
                self.case_results[case.case_id] = CaseResult(case.case_id, CaseStatus.PREFLIGHT_FAILED, "；".join(errors), printer=case.printer or fallback, source_file=case.file_path)
                bad += 1
            else:
                self.case_results[case.case_id] = CaseResult(case.case_id, CaseStatus.PREFLIGHT_OK, "文件、打印机及参数检查通过。", printer=case.printer or fallback, source_file=case.file_path)
                ok += 1
        self._render_cases()
        self._update_kpis()
        self.status_var.set(f"打印前检查完成：{ok} 条通过，{bad} 条有问题")
        self._log(f"打印前检查完成：{ok} 条通过，{bad} 条失败。")
        if bad:
            messagebox.showwarning("检查完成", f"检查通过：{ok} 条\n检查失败：{bad} 条\n请查看用例队列中的诊断信息，修正后再执行。")
        else:
            messagebox.showinfo("检查通过", f"{ok} 条启用用例全部通过打印前检查。")

    def _start_run(self, selected_only: bool) -> None:
        if self.running:
            return
        if not self.cases:
            messagebox.showinfo("尚无用例", "请先导入 Excel 用例文件。")
            return
        if self.printer is None:
            messagebox.showerror("打印服务不可用", "请在 Windows 10/11 上运行，并确认打印服务可用。")
            return
        try:
            from .direct_print import engine_available

            engine_ok, engine_detail = engine_available()
        except Exception as exc:  # noqa: WPS421
            engine_ok, engine_detail = False, f"自检异常：{exc}"
        if not engine_ok:
            messagebox.showerror("打印引擎不可用", f"内置直打引擎自检失败：{engine_detail}。请重新运行安装脚本 setup_windows.ps1 修复依赖。")
            return
        if selected_only:
            indices = sorted(self.checked)
            if not indices:
                indices = [int(i) for i in self.tree.selection()]
            if not indices:
                messagebox.showinfo("未选择用例", "请先勾选用例最左侧的“选”框（或按住 Ctrl 选中行）。")
                return
            cases = [self.cases[i] for i in indices]
        else:
            cases = [c for c in self.cases if c.enabled]
        if not cases:
            messagebox.showinfo("没有可执行用例", "当前没有启用的用例。")
            return

        printers = set(self.printer.list_printers())
        fallback = self.printer_var.get().strip()
        caps_cache: dict[str, dict | None] = {}
        duplicate_ids = {c.case_id for c in cases if sum(1 for x in self.cases if x.case_id == c.case_id) > 1}
        ready_cases: list[PrintCase] = []
        invalid: list[str] = []
        preflight_results: list[CaseResult] = []
        for case in cases:
            if not case.enabled:
                continue
            target_printer = case.printer.strip() or fallback
            errors = validate_case(case, printers, fallback)
            source = case.resolved_path or (Path(case.file_path) if case.file_path else None)
            if source is not None and source.suffix.lower() in OFFICE_EXTENSIONS:
                probe = probe_document_converter(source.suffix)
                if not probe[0]:
                    errors.append(f"打印前检查：{source.suffix} 文件缺少可用转换组件（{probe[1]}）。")
            if target_printer and target_printer in printers:
                if target_printer not in caps_cache:
                    try:
                        caps_cache[target_printer] = self.printer.get_capabilities(target_printer)
                    except Exception as exc:
                        caps_cache[target_printer] = None
                        self._log(f"无法读取 {target_printer} 的驱动能力，跳过该项能力核验：{exc}")
                errors.extend(validate_driver_capabilities(case, caps_cache[target_printer]))
            if case.case_id in duplicate_ids:
                errors.append("用例 ID 重复")
            if errors:
                failed_result = CaseResult(case.case_id, CaseStatus.PREFLIGHT_FAILED, "；".join(errors), printer=case.printer or fallback, source_file=case.file_path)
                self.case_results[case.case_id] = failed_result
                preflight_results.append(failed_result)
                invalid.append(case.case_id + "：" + "；".join(errors))
            else:
                ready_cases.append(case)
        self._render_cases()
        self._update_kpis()
        if invalid:
            messagebox.showwarning("部分用例未通过检查", "以下用例不会下发：\n\n" + "\n".join(invalid[:8]) + ("\n……" if len(invalid) > 8 else ""))
        if not ready_cases:
            return
        confirmed = messagebox.askyesno("确认批量下发", f"即将向打印机下发 {len(ready_cases)} 条用例。\n\n打印机：按用例指定或默认打印机\n策略：串行执行；不确定状态不自动重试。\n\n确认继续？")
        if not confirmed:
            return

        self.run_id = uuid.uuid4().hex[:12].upper()
        self.run_results = list(preflight_results)
        self.runner = PrintRunner(self.printer, self.audit)
        self.running = True
        self._set_running_ui(True)
        self.status_var.set(f"执行中 · RUN {self.run_id}")
        self._log(f"===== 开始执行 RUN {self.run_id}，共 {len(ready_cases)} 条用例 =====")

        def worker() -> None:
            try:
                self.runner.run_cases(ready_cases, fallback, self.run_id, lambda kind, payload: self.event_queue.put((kind, payload)))
                self.event_queue.put(("finished", self.run_id))
            except Exception as exc:
                self.event_queue.put(("fatal", str(exc)))

        self.worker = threading.Thread(target=worker, name="PrintCaseRunner", daemon=True)
        self.worker.start()

    def _stop_run(self) -> None:
        if self.runner and self.running:
            self.runner.stop()
            self.status_var.set("正在停止：当前作业可能已经进入打印队列")
            self._log("收到停止请求：将停止后续用例，不会假定当前打印作业已撤回。")
            self.btn_stop.configure(state="disabled")

    def _set_running_ui(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.btn_import.configure(state=state)
        self.btn_preflight.configure(state=state)
        self.btn_run_selected.configure(state=state)
        self.btn_run_all.configure(state=state)
        self.btn_stop.configure(state="normal" if running else "disabled")

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.event_queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "result" and isinstance(payload, CaseResult):
                    self.case_results[payload.case_id] = payload
                    if payload.ended_at is not None:
                        self.run_results.append(payload)
                    self._render_cases()
                    self._update_kpis()
                elif kind == "driver_check":
                    self._apply_driver_check(*payload)
                elif kind == "finished":
                    self.running = False
                    self._set_running_ui(False)
                    self.status_var.set(f"执行完成 · RUN {payload}")
                    self._log(f"===== 执行结束 RUN {payload} =====")
                    self._update_kpis()
                elif kind == "fatal":
                    self.running = False
                    self._set_running_ui(False)
                    self.status_var.set("执行异常")
                    self._log("执行线程异常：" + str(payload))
                    messagebox.showerror("执行异常", str(payload))
        except queue.Empty:
            pass
        self.root.after(150, self._poll_events)

    def _update_kpis(self) -> None:
        enabled_cases = [c for c in self.cases if c.enabled]
        statuses = [r.status for r in self.case_results.values()]
        ready = sum(1 for c in enabled_cases if self.case_results.get(c.case_id) and self.case_results[c.case_id].status in {CaseStatus.PREFLIGHT_OK, CaseStatus.PENDING})
        accepted = sum(1 for s in statuses if s in {CaseStatus.QUEUED, CaseStatus.DRIVER_PRINTED, CaseStatus.QUEUE_CLEARED})
        attention = sum(1 for s in statuses if s in {CaseStatus.FAILED, CaseStatus.PREFLIGHT_FAILED, CaseStatus.TIMEOUT_QUEUED, CaseStatus.SUBMITTED_UNVERIFIED})
        self.kpi_labels["total"].configure(text=str(len(self.cases)))
        self.kpi_labels["ready"].configure(text=str(ready))
        self.kpi_labels["accepted"].configure(text=str(accepted))
        self.kpi_labels["attention"].configure(text=str(attention))

    def _export_report(self) -> None:
        if not self.run_results:
            messagebox.showinfo("暂无执行结果", "完成至少一条打印用例后才能导出执行报告。")
            return
        filename = filedialog.asksaveasfilename(title="导出执行报告", defaultextension=".csv", filetypes=[("CSV 报告", "*.csv"), ("JSON 报告", "*.json")], initialfile=f"PrintCase_Report_{self.run_id or 'latest'}.csv")
        if not filename:
            return
        try:
            self.audit.export(self.run_id or "PREFLIGHT", Path(filename), self.run_results)
            messagebox.showinfo("导出成功", f"报告已保存：\n{filename}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _open_template(self) -> None:
        path = app_base_dir() / "templates" / "PrintCase_Template.xlsx"
        if not path.exists():
            path = Path(__file__).resolve().parents[2] / "templates" / "PrintCase_Template.xlsx"
        if path.exists():
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except Exception as exc:
                messagebox.showerror("无法打开模板", str(exc))
        else:
            messagebox.showwarning("模板不存在", "未找到 templates/PrintCase_Template.xlsx，请重新解压完整软件包。")

    def _open_readme(self) -> None:
        path = app_base_dir() / "README_CN.md"
        if not path.exists():
            path = Path(__file__).resolve().parents[2] / "README_CN.md"
        if path.exists():
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except Exception:
                webbrowser.open(path.as_uri())
        else:
            webbrowser.open(path.as_uri())

    def _log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("确认退出", "打印任务正在执行。退出将停止后续下发，但无法保证已提交到打印机的作业被撤回。仍要退出？"):
                return
            self._stop_run()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    PrintCaseApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
