from __future__ import annotations

import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

WORD_EXTENSIONS = {".docx", ".doc", ".rtf"}
WPS_WRITER_EXTENSIONS = {".wps", ".wpt", ".odt", ".ott"}  # WPS 文字可直接打开的格式
# 这些格式 Word 打不开，优先 LibreOffice；WPS 文字可处理其中文档类格式
LIBREOFFICE_EXTENSIONS = {
    ".odt", ".ott", ".wps", ".wpt",
    ".pptx", ".ppsx", ".pptm",
    ".xls", ".xlsx", ".et", ".ett",
}
OFFICE_EXTENSIONS = WORD_EXTENSIONS | LIBREOFFICE_EXTENSIONS
WD_EXPORT_FORMAT_PDF = 17  # Word / WPS 的 wdExportFormatPDF 常量
WPS_PROG_IDS = ("KWPS.Application", "Kwps.Application", "WPS.Application")


def backend_order(suffix: str) -> list[str]:
    """按格式决定转换后端优先级：Word 文档 Word→WPS→LibreOffice；
    LibreOffice 专属格式 soffice→WPS（文字类）→Word 兜底。"""
    suffix = str(suffix or "").lower()
    if suffix in WORD_EXTENSIONS:
        return ["word", "wps", "soffice"]
    if suffix in WPS_WRITER_EXTENSIONS:
        return ["soffice", "wps", "word"]
    if suffix in LIBREOFFICE_EXTENSIONS:
        return ["soffice"]
    return ["word", "wps", "soffice"]


class DocumentConversionError(RuntimeError):
    """Office 文档转 PDF 失败（找不到可用转换组件，或转换过程出错）。"""


def _word_com_available() -> bool:
    return platform.system() == "Windows"


def _find_soffice() -> str | None:
    hit = shutil.which("soffice") or shutil.which("soffice.exe")
    if hit:
        return hit
    for root in ("C:/Program Files", "C:/Program Files (x86)"):
        candidate = Path(root) / "LibreOffice" / "program" / "soffice.exe"
        if candidate.is_file():
            return str(candidate)
    return None


def _convert_with_word(doc_path: Path, out_pdf: Path, word_factory=None) -> Path:
    import pythoncom  # pywin32；保证本线程 COM 已初始化
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = word_factory() if word_factory else win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass
        doc = word.Documents.Open(
            str(doc_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
        )
        try:
            doc.ExportAsFixedFormat(str(out_pdf), WD_EXPORT_FORMAT_PDF)
        finally:
            doc.Close(False)
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    if not out_pdf.exists() or out_pdf.stat().st_size == 0:
        raise DocumentConversionError("Word 未生成有效的 PDF 输出。")
    return out_pdf


def _convert_with_wps(doc_path: Path, out_pdf: Path, wps_factory=None) -> Path:
    """用 WPS 文字（COM，ProgID：KWPS.Application 等）把文档导出为 PDF。"""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    if wps_factory:
        app = wps_factory()
    else:
        last_error: Exception | None = None
        for prog_id in WPS_PROG_IDS:
            try:
                app = win32com.client.DispatchEx(prog_id)
                break
            except Exception as exc:  # 未安装 WPS 或 ProgID 变化
                last_error = exc
        if app is None:
            raise DocumentConversionError(
                f"未找到 WPS 文字 COM 组件（尝试：{'、'.join(WPS_PROG_IDS)}）。{last_error}"
            )
    try:
        app.Visible = False
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        doc = app.Documents.Open(str(doc_path), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(str(out_pdf), WD_EXPORT_FORMAT_PDF)
        finally:
            doc.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    if not out_pdf.exists() or out_pdf.stat().st_size == 0:
        raise DocumentConversionError("WPS 未生成有效的 PDF 输出。")
    return out_pdf


def _convert_with_soffice(doc_path: Path, out_dir: Path, soffice_path: str, timeout: int) -> Path:
    if not Path(soffice_path).is_file():
        raise DocumentConversionError(
            f"未找到 LibreOffice 可执行文件（{soffice_path}）；{doc_path.suffix} 格式只能通过 LibreOffice 转换。"
        )
    # 独立 UserInstallation 配置目录，避免与正在运行的 LibreOffice 实例冲突
    profile = (Path(tempfile.gettempdir()) / "printcase_soffice_profile").resolve()
    command = [
        soffice_path,
        "-env:UserInstallation=file:///" + str(profile).replace("\\", "/"),
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(doc_path),
    ]
    try:
        proc = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        raise DocumentConversionError(f"LibreOffice 转换超时（{timeout} 秒）。") from exc
    expected = out_dir / (doc_path.stem + ".pdf")
    if not expected.exists() or expected.stat().st_size == 0:
        detail = (proc.stderr or proc.stdout or "").strip()[:300]
        raise DocumentConversionError(f"LibreOffice 未生成 PDF。{detail}")
    return expected


def convert_document_to_pdf(
    doc_path: str | Path,
    work_dir: str | Path | None = None,
    timeout: int = 180,
    word_factory=None,
    wps_factory=None,
    soffice_path: str | None = None,
    prefer: str = "auto",
) -> Path:
    """把 Office 系列文档转成临时 PDF，返回 PDF 路径。

    转换组件按 Word → WPS → LibreOffice 顺序自动检测（WPS 文字可处理
    docx/doc/rtf/wps/wpt/odt），保证 Office 全家桶与 WPS 生态都能覆盖。
    打印属性（纸盒、双面、纸张等）由内置直打引擎对转换后的 PDF 下发，
    保证不同源格式走同一条打印链路。
    """
    doc_path = Path(doc_path)
    if doc_path.suffix.lower() not in OFFICE_EXTENSIONS:
        raise DocumentConversionError(f"不是受支持的 Office 文档：{doc_path.suffix or '无扩展名'}")
    if not doc_path.exists():
        raise DocumentConversionError(f"找不到文档：{doc_path}")

    out_dir = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix="printcase_doc_"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out_pdf = out_dir / (doc_path.stem + ".pdf")

    errors: list[str] = []
    order = backend_order(doc_path.suffix)
    if prefer in {"soffice", "word", "wps"}:
        order = [prefer] + [b for b in order if b != prefer]
    for backend in order:
        try:
            if backend == "word":
                if not _word_com_available() and word_factory is None:
                    raise DocumentConversionError("Word COM 仅在 Windows 上可用。")
                return _convert_with_word(doc_path, out_pdf, word_factory=word_factory)
            if backend == "wps":
                return _convert_with_wps(doc_path, out_pdf, wps_factory=wps_factory)
            if backend == "soffice":
                path = soffice_path or _find_soffice()
                if not path:
                    raise DocumentConversionError(
                        f"未找到 LibreOffice（soffice）；{doc_path.suffix} 格式只能通过 LibreOffice 转换。"
                    )
                return _convert_with_soffice(doc_path, out_dir, path, timeout)
        except DocumentConversionError as exc:
            errors.append(f"{backend}: {exc}")
        except Exception as exc:  # COM 或子进程的意外异常
            errors.append(f"{backend}: {type(exc).__name__}: {exc}")
    raise DocumentConversionError("文档转换为 PDF 失败。已尝试：" + "；".join(errors))


def _wps_com_available() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        try:
            for prog_id in WPS_PROG_IDS:
                try:
                    app = win32com.client.DispatchEx(prog_id)
                except Exception:
                    continue
                try:
                    app.Quit()
                except Exception:
                    pass
                return True
        finally:
            pythoncom.CoUninitialize()
    except ImportError:
        return False
    return False


def probe_document_converter(extension: str = "") -> tuple[bool, str]:
    """打印前检查用：判断本机是否有能转换该格式的组件（Word → WPS → LibreOffice）。"""
    suffix = str(extension or "").lower()
    if suffix and suffix in {".pptx", ".ppsx", ".pptm", ".xls", ".xlsx", ".et", ".ett"}:
        soffice = _find_soffice()
        if soffice:
            return True, f"LibreOffice（{soffice}）"
        return False, f"{suffix} 格式需要 LibreOffice 转换（未检测到 soffice.exe）"
    if _word_com_available():
        try:
            import win32com.client  # noqa: F401

            return True, "Microsoft Word（COM）"
        except ImportError:
            pass
    if _wps_com_available():
        return True, "WPS 文字（COM）"
    soffice = _find_soffice()
    if soffice:
        return True, f"LibreOffice（{soffice}）"
    return False, "未找到 Word / WPS / LibreOffice"
