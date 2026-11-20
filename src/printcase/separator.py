from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .models import PrintCase

_ASCII_STRIP = re.compile(r"[^ -~]")


def _ascii(text: str) -> str:
    """PDF 内置字体只支持 WinAnsi 编码；把不可见字符统一替换为问号。"""
    return _ASCII_STRIP.sub("?", str(text))


def build_separator_pdf(dest: Path, lines: Iterable[str]) -> Path:
    """Generate a minimal single-page PDF with one line of text per entry.

    使用内置 Helvetica 字体的极简 PDF 生成器，不依赖第三方库。
    文本仅支持 ASCII；中文等字符会被替换为 '?'，请在 lines 里预先准备
    ASCII 内容（如用例 ID、时间戳、打印机英文名）。
    """
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    content_parts = ["BT", "/F1 28 Tf", "72 740 Td", "26 TL"]
    first = True
    for line in lines:
        text = _ascii(line) or " "
        escaped = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        if first:
            content_parts.append(f"({escaped}) Tj")
            first = False
        else:
            content_parts.append("T*")
            content_parts.append(f"({escaped}) Tj")
    content_parts.append("ET")
    stream = "\n".join(content_parts).encode("latin-1", "replace")

    objects: list[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    objects.append(b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    objects.append(b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>")
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    objects.append(b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream")

    out = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{index} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_offset = len(out)
    out += f"xref\n0 {len(objects) + 1}\n".encode()
    out += b"0000000000 65535 f \n"
    for offset in offsets[1:]:
        out += f"{offset:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
        f"startxref\n{xref_offset}\n%%EOF\n"
    ).encode()
    dest.write_bytes(bytes(out))
    return dest


def build_case_separator_lines(case: PrintCase, printer: str) -> list[str]:
    """分隔页内容：ASCII 为主的作业间隔信息。"""
    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_stem = Path(case.resolved_path or case.file_path).stem or case.file_path
    return [
        "PrintCase Studio - Separator Page",
        "==================================",
        f"Case  : {_ascii(case.case_id)}",
        f"File  : {_ascii(source_stem)}",
        f"Printer: {_ascii(printer)}",
        f"Copies : {case.copies}",
        f"Time  : {started}",
        "",
        "(This page separates print jobs.)",
    ]


def create_case_separator_pdf(case: PrintCase, printer: str, out_dir: Path) -> Path:
    safe_id = _ascii(case.case_id).replace(" ", "_") or "case"
    dest = Path(out_dir) / f"separator_{safe_id}.pdf"
    return build_separator_pdf(dest, build_case_separator_lines(case, printer))
