"""原生直接打印链路：不依赖任何外部打印组件。

两条通道：

1. **Office COM 直打**（docx/doc/rtf/odt/wps/wpt）：用 Microsoft Word 或 WPS 文字
   的 COM 接口打开文档 → 设置 ActivePrinter（选择驱动）→ 按用例属性调用
   PrintOut 下发。打印属性（纸张类型、DPI、逐份等）来自已写入并回读确认的
   驱动默认 DEVMODE。

2. **GDI 直打**（PDF/图片/转换后的 PDF）：PyMuPDF 渲染 PDF 页面、Pillow 读取
   图片，再通过 win32ui 的打印机设备上下文直接把位图写入 Windows 打印队列。
   CreatePrinterDC 继承驱动默认 DEVMODE，因此同样使用已校验的属性。

页码选择、份数与逐份/非逐份在 GDI 通道内按页精确控制。
"""
from __future__ import annotations

import threading
from pathlib import Path
from typing import Callable

from .models import PrintCase
from .office import WPS_PROG_IDS

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
# Word / WPS 文字组件可以直接打开并打印的文档（无需先转 PDF）
COM_PRINTABLE_EXTENSIONS = {".docx", ".doc", ".rtf", ".odt", ".wps", ".wpt"}

RENDER_DPI = 180  # PDF 渲染分辨率；打印测试足够清晰且内存可控

# GDI 设备能力常量
HORZRES = 8
VERTRES = 10


class DirectPrintError(RuntimeError):
    """直接打印链路失败（打开文档、选择打印机、下发任一步骤失败）。"""


class ComponentMissingError(DirectPrintError):
    """本机没有能打开该文档的 COM 打印组件（Word/WPS 均不可用）。"""


def parse_page_selection(spec: str, total: int) -> list[int]:
    """把页码选择（如 “1-3,5,last,odd,even”）解析成 0 起始的页索引列表。"""
    text = str(spec or "").strip()
    if not text:
        return list(range(total))
    selected: list[int] = []
    for token in [t.strip().casefold() for t in text.split(",") if t.strip()]:
        if token == "last":
            selected.append(total - 1)
        elif token == "odd":
            selected.extend(range(0, total, 2))
        elif token == "even":
            selected.extend(range(1, total, 2))
        elif "-" in token:
            left, right = token.split("-", 1)
            start = total if left == "last" else int(left)
            end = total if right == "last" else int(right)
            if start > end:
                start, end = end, start
            selected.extend(range(start - 1, end))
        else:
            selected.append(int(token) - 1)
    return sorted({i for i in selected if 0 <= i < total})


def page_order(selected: list[int], copies: int, collate: str) -> list[int]:
    """按逐份/非逐份展开页序。

    逐份（collate=yes）：[1..n] × copies —— 每一份按页序完整输出。
    非逐份：每页连打 copies 次 —— [1,1,2,2,...]。
    """
    if copies <= 1:
        return list(selected)
    if collate == "yes":
        return [index for _ in range(copies) for index in selected]
    return [index for index in selected for _ in range(copies)]


# ---------------------------------------------------------------------------
# GDI 通道
# ---------------------------------------------------------------------------

class GdiSession:
    """GDI 打印会话（ctypes 直调 gdi32）。

    pywin32 的 win32ui.CDC.StartDoc 在部分环境/驱动上会直接崩溃，
    因此这里用 ctypes 调用 gdi32，参数签名完整声明，异常可捕获。
    """

    def __init__(self, printer: str, doc_name: str, output_file: str | Path | None = None):
        import ctypes
        import ctypes.wintypes as wintypes

        self._ctypes = ctypes
        self._gdi = ctypes.WinDLL("gdi32", use_last_error=True)
        gdi = self._gdi

        class DOCINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_int),
                ("lpszDocName", wintypes.LPCWSTR),
                ("lpszOutput", wintypes.LPCWSTR),
                ("lpszDatatype", wintypes.LPCWSTR),
                ("fwType", wintypes.DWORD),
            ]

        self._docinfo_type = DOCINFO
        gdi.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
        gdi.CreateDCW.restype = wintypes.HDC
        gdi.StartDocW.argtypes = [wintypes.HDC, ctypes.POINTER(DOCINFO)]
        gdi.StartDocW.restype = ctypes.c_int
        gdi.StretchDIBits.argtypes = [
            wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT, wintypes.DWORD,
        ]
        gdi.StretchDIBits.restype = ctypes.c_int
        for name in ("StartPage", "EndPage", "EndDoc", "DeleteDC"):
            getattr(gdi, name).argtypes = [wintypes.HDC]
            getattr(gdi, name).restype = ctypes.c_int
        gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
        gdi.GetDeviceCaps.restype = ctypes.c_int

        self.printer = printer
        self.doc_name = doc_name
        self.output_file = str(output_file) if output_file else None
        self.hdc = gdi.CreateDCW("WINSPOOL", printer, None, None)
        if not self.hdc:
            raise DirectPrintError(
                f"无法为打印机“{printer}”创建打印设备上下文（错误码 {ctypes.get_last_error()}）。"
            )
        self._doc_started = False
        self._in_page = False
        self.job_id: int | None = None

    # ---- 页面信息 ----
    def page_area(self) -> tuple[int, int]:
        width = self._gdi.GetDeviceCaps(self.hdc, HORZRES) or 2100
        height = self._gdi.GetDeviceCaps(self.hdc, VERTRES) or 2900
        return width, height

    # ---- 文档/页面控制 ----
    def start_document(self) -> None:
        info = self._docinfo_type(
            self._ctypes.sizeof(self._docinfo_type), self.doc_name, self.output_file, None, 0,
        )
        job_id = self._gdi.StartDocW(self.hdc, self._ctypes.byref(info))
        if not job_id:
            raise DirectPrintError(
                f"打印机“{self.printer}”拒绝开始打印作业（StartDoc 失败，错误码 "
                f"{self._ctypes.get_last_error()}）。请检查驱动状态与打印队列。"
            )
        # StartDoc 成功时返回 Windows 打印队列分配的作业号
        self.job_id = int(job_id) if int(job_id) > 0 else None
        self._doc_started = True

    def start_page(self) -> None:
        if not self._doc_started:
            self.start_document()
        if not self._gdi.StartPage(self.hdc):
            raise DirectPrintError(f"StartPage 失败（错误码 {self._ctypes.get_last_error()}）。")
        self._in_page = True

    def end_page(self) -> None:
        if self._in_page:
            self._gdi.EndPage(self.hdc)
            self._in_page = False

    def draw_image(self, image, box=None) -> None:
        """把 PIL 图像绘制到页面（缺省按可打印区域等比居中）。"""
        ctypes = self._ctypes
        from PIL import Image  # noqa: WPS433

        if not isinstance(image, Image.Image):
            raise DirectPrintError("draw_image 需要 PIL.Image 对象。")
        rgb = image.convert("RGB")
        width, height = rgb.size
        if box is None:
            area_w, area_h = self.page_area()
            scale = min(area_w / width, area_h / height)
            draw_w, draw_h = max(1, int(width * scale)), max(1, int(height * scale))
            left, top = (area_w - draw_w) // 2, (area_h - draw_h) // 2
            box = (left, top, left + draw_w, top + draw_h)

        # 24bpp DIB，行按 4 字节对齐
        stride = ((width * 3 + 3) // 4) * 4
        raw = rgb.tobytes("raw", "BGR")
        if stride == width * 3:
            bits = raw
        else:
            padded = bytearray(stride * height)
            for row in range(height):
                start = row * width * 3
                padded[row * stride: row * stride + width * 3] = raw[start: start + width * 3]
            bits = bytes(padded)

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_int32), ("biHeight", ctypes.c_int32),
                ("biPlanes", ctypes.c_uint16), ("biBitCount", ctypes.c_uint16),
                ("biCompression", ctypes.c_uint32), ("biSizeImage", ctypes.c_uint32),
                ("biXPelsPerMeter", ctypes.c_int32), ("biYPelsPerMeter", ctypes.c_int32),
                ("biClrUsed", ctypes.c_uint32), ("biClrImportant", ctypes.c_uint32),
            ]

        header = BITMAPINFOHEADER(
            ctypes.sizeof(BITMAPINFOHEADER), width, -height, 1, 24, 0, len(bits), 0, 0, 0, 0,
        )
        left, top, right, bottom = (int(value) for value in box)
        self.start_page()
        self._gdi.StretchDIBits(
            self.hdc, left, top, right - left, bottom - top,
            0, 0, width, height,
            bits, ctypes.byref(header), 0, 0x00CC0020,  # DIB_RGB_COLORS, SRCCOPY
        )
        self.end_page()

    def close(self) -> None:
        try:
            self.end_page()
            if self._doc_started:
                self._gdi.EndDoc(self.hdc)
                self._doc_started = False
        finally:
            self._gdi.DeleteDC(self.hdc)


# 可被冒烟脚本/测试替换（例如输出到文件）
SESSION_FACTORY = None


def _gdi_send(printer: str, doc_name: str, order: list[int], render_page, log, stop_event=None, session_factory=None) -> int | None:
    """按页序把渲染结果写入打印机（CreateDC 继承已校验的驱动默认 DEVMODE）。"""
    factory = session_factory or SESSION_FACTORY or GdiSession
    session = factory(printer, doc_name)
    try:
        for position, page_index in enumerate(order, 1):
            if stop_event is not None and stop_event.is_set():
                raise DirectPrintError("收到停止请求，剩余页面未下发；请检查打印队列。")
            image = render_page(page_index)
            session.draw_image(image)
            log(f"已发送第 {position}/{len(order)} 页（源页 {page_index + 1}）到打印机队列")
        return getattr(session, "job_id", None)
    finally:
        session.close()


def print_pdf_gdi(printer: str, case: PrintCase, pdf_path: Path, log: Callable[[str], None], stop_event=None, fitz_module=None, session_factory=None) -> int | None:
    """把 PDF 渲染为位图后经 GDI 直打；页码/份数/逐份按用例控制。"""
    if fitz_module is None:
        import pymupdf as fitz_module  # noqa: WPS433
    from PIL import Image  # noqa: WPS433

    document = fitz_module.open(str(pdf_path))
    try:
        total = document.page_count
        selected = parse_page_selection(case.pages, total)
        if not selected:
            raise DirectPrintError(f"页码“{case.pages}”在文档（共 {total} 页）中没有可打印页面。")
        order = page_order(selected, case.copies, case.collate)
        log(f"PDF 共 {total} 页；本次按 {case.pages or '全部'} 选择 {len(selected)} 页，展开为 {len(order)} 页（{case.copies} 份）")
        zoom = RENDER_DPI / 72
        matrix = fitz_module.Matrix(zoom, zoom)

        def render_page(page_index: int):
            pixmap = document[page_index].get_pixmap(matrix=matrix, alpha=False)
            mode = "RGB" if pixmap.n < 4 else "RGBA"
            return Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples)

        job_id = _gdi_send(printer, pdf_path.stem, order, render_page, log, stop_event=stop_event, session_factory=session_factory)
        if job_id is not None:
            log(f"Windows 打印队列已分配作业号 #{job_id}")
        return job_id
    finally:
        document.close()


def print_image_gdi(printer: str, case: PrintCase, image_path: Path, log: Callable[[str], None], stop_event=None, session_factory=None, pil_module=None) -> int | None:
    """把图片经 GDI 直打（多帧 TIF 按多页处理）。"""
    if pil_module is None:
        from PIL import Image as pil_module  # noqa: WPS433

    image = pil_module.open(str(image_path))
    frames = getattr(image, "n_frames", 1)
    selected = parse_page_selection(case.pages, frames)
    if not selected:
        raise DirectPrintError(f"页码“{case.pages}”在图片（{frames} 帧）中没有可打印页面。")
    order = page_order(selected, case.copies, case.collate)
    log(f"图片 {frames} 帧；本次打印 {len(selected)} 帧，展开为 {len(order)} 帧（{case.copies} 份）")

    def render_page(page_index: int):
        frame = pil_module.open(str(image_path))
        frame.seek(page_index)
        return frame.convert("RGB")

    return _gdi_send(printer, image_path.stem, order, render_page, log, stop_event=stop_event, session_factory=session_factory)


# ---------------------------------------------------------------------------
# Office COM 通道
# ---------------------------------------------------------------------------

def print_document_com(
    printer: str,
    case: PrintCase,
    doc_path: Path,
    log: Callable[[str], None],
    word_factory=None,
    wps_factory=None,
) -> None:
    """用 Word（优先）或 WPS 文字打开文档 → 选择打印机 → 按属性下发打印。"""
    import pythoncom  # noqa: WPS433

    pythoncom.CoInitialize()
    app = None
    backend = ""
    errors: list[str] = []
    try:
        candidates = [("Microsoft Word", ["Word.Application"], word_factory)]
        # Word 打不开 WPS 私有格式（.wps/.wpt），这两种只交给 WPS
        if doc_path.suffix.lower() not in {".wps", ".wpt"}:
            pass
        else:
            candidates = []
        candidates += [("WPS 文字", list(WPS_PROG_IDS), wps_factory)]
        for backend_name, prog_ids, factory in candidates:
            try:
                if factory is not None:
                    app = factory()
                else:
                    import win32com.client  # noqa: WPS433

                    for prog_id in prog_ids:
                        try:
                            app = win32com.client.DispatchEx(prog_id)
                            break
                        except Exception as exc:  # noqa: WPS421
                            errors.append(f"{prog_id}: {exc}")
                if app is not None:
                    backend = backend_name
                    break
            except Exception as exc:  # noqa: WPS421
                errors.append(f"{backend_name}: {exc}")
        if app is None:
            raise ComponentMissingError(
                "本机没有能打开该文档的打印组件（Word / WPS 均不可用）。" + "；".join(errors)
            )

        log(f"已启动打印组件：{backend}")
        try:
            app.Visible = False
        except Exception:  # noqa: WPS420
            pass
        try:
            app.DisplayAlerts = 0
        except Exception:  # noqa: WPS420
            pass

        log(f"正在选择打印机驱动：{printer}")
        app.ActivePrinter = printer  # 驱动不存在时 Word/WPS 会抛错，直接拦截
        log(f"打印机已选定：{app.ActivePrinter}")

        document = app.Documents.Open(str(doc_path), ReadOnly=True, AddToRecentFiles=False)
        try:
            log(f"文档已打开：{doc_path.name}")
            options = {"Copies": int(case.copies)}
            if case.collate:
                options["Collate"] = case.collate == "yes"
            if str(case.pages or "").strip():
                options["Range"] = 4  # wdPrintRangeOfPages
                options["Pages"] = case.pages.replace(" ", "")
            document.PrintOut(**options)
            log(f"已按用例属性下发打印（份数 {case.copies}，逐份={'是' if case.collate == 'yes' else '否/默认'}）")
        finally:
            document.Close(False)
        log("文档已关闭")
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: WPS420
                pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 路由
# ---------------------------------------------------------------------------

def print_with_driver(
    printer: str,
    case: PrintCase,
    source_path: Path,
    log: Callable[[str], None],
    stop_event=None,
    word_factory=None,
    wps_factory=None,
    document_converter=None,
    temp_dir: Path | None = None,
) -> tuple[str, str, int | None]:
    """按文档类型路由到 COM 直打或 GDI 直打。

    返回 (通道标识, 队列作业名主干)。转换产生临时 PDF 时作业名仍用源文件主干，
    便于在打印队列里按用例文件关联作业。
    """
    suffix = source_path.suffix.lower()
    if suffix in COM_PRINTABLE_EXTENSIONS:
        try:
            print_document_com(printer, case, source_path, log, word_factory=word_factory, wps_factory=wps_factory)
            return "office-com", source_path.stem, None
        except ComponentMissingError as exc:
            if document_converter is None:
                raise
            log(f"本机无 Word/WPS 打印组件（{exc}），改为转换为 PDF 后由内置 GDI 引擎直打。")

    if suffix == ".pdf":
        return "gdi-pdf", source_path.stem, print_pdf_gdi(printer, case, source_path, log, stop_event=stop_event)
    if suffix in IMAGE_EXTENSIONS:
        return "gdi-image", source_path.stem, print_image_gdi(printer, case, source_path, log, stop_event=stop_event)

    if document_converter is None:
        raise DirectPrintError(f"{suffix} 格式没有可用的直接打印通道，且未配置 PDF 转换组件。")
    if temp_dir is None:
        import tempfile  # noqa: WPS433

        temp_dir = Path(tempfile.mkdtemp(prefix="printcase_direct_"))
    log(f"{suffix} 格式需先转换为 PDF（Word/WPS 打不开该格式）…")
    converted = document_converter(source_path, work_dir=temp_dir, timeout=180)
    log(f"转换完成：{converted.name}，开始内置 GDI 直打")
    return "gdi-pdf-converted", source_path.stem, print_pdf_gdi(printer, case, converted, log, stop_event=stop_event)


def engine_available() -> tuple[bool, str]:
    """启动自检：GDI 打印所需的 win32ui / PyMuPDF / Pillow 是否可导入。"""
    missing: list[str] = []
    try:
        import win32ui  # noqa: F401, WPS433
    except ImportError:
        missing.append("win32ui(pywin32)")
    try:
        import pymupdf  # noqa: F401, WPS433
    except ImportError:
        missing.append("PyMuPDF")
    try:
        import PIL  # noqa: F401, WPS433
    except ImportError:
        missing.append("Pillow")
    if missing:
        return False, "缺少：" + "、".join(missing)
    return True, "内置直打引擎（GDI + PDF 渲染）就绪"
