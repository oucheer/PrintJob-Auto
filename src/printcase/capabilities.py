from __future__ import annotations

import re

from .models import PrintCase
from .validator import normalize_paper


def _norm(value: str) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", str(value).casefold())


def _media_reported(media: str, names: list[str] | None) -> bool:
    """驱动是否上报了该纸张类型（名称可比克重后缀宽松匹配）。"""
    if not names:
        return True
    wanted = _norm(media)
    if not wanted:
        return True
    for name in names:
        normalized = _norm(name)
        if normalized and (normalized == wanted or wanted in normalized or normalized in wanted):
            return True
    return False


def validate_driver_capabilities(case: PrintCase, caps: dict | None) -> list[str]:
    """只核验驱动是否具备用例要求的选项（属性本身由驱动属性页设置）。"""
    if not caps:
        return []
    errors: list[str] = []
    papers = caps.get("papers")
    paper_token = normalize_paper(case.paper)
    if papers and paper_token and paper_token not in {"auto", "mixed"} and case.mixed_sizes != "yes":
        normalized = [_norm(p) for p in papers]
        if not any(paper_token == name or paper_token in name for name in normalized if name):
            errors.append(
                f"驱动能力检查：打印机未报告支持纸张 {case.paper}（驱动可用：{', '.join(papers[:8])}）"
            )
    if case.duplex.lower() != "simplex" and caps.get("duplex") is False:
        errors.append("驱动能力检查：该打印机未报告双面打印能力")
    if case.color.lower() == "color" and caps.get("color") is False:
        errors.append("驱动能力检查：该打印机未报告彩色打印能力")
    if case.bin:
        bin_names = caps.get("bin_names")
        bin_ids = caps.get("bin_ids")
        if bin_names is not None or bin_ids is not None:
            wanted = case.bin.strip()
            name_match = bool(bin_names) and any(_norm(wanted) == _norm(x) or _norm(wanted) in _norm(x) for x in bin_names)
            id_match = bool(bin_ids) and wanted.isdigit() and int(wanted) in bin_ids
            if not name_match and not id_match:
                known = ", ".join(bin_names or []) or ", ".join(str(x) for x in (bin_ids or []))
                errors.append(f"驱动能力检查：未发现纸盒/进纸器 {wanted}（可用：{known or '驱动未提供列表'}）")
    media = str(case.media_type or "").strip()
    if media and not media.isdigit() and not _media_reported(media, caps.get("media_names")):
        available = "、".join((caps.get("media_names") or [])[:8])
        errors.append(
            f"驱动能力检查：打印机未报告纸张类型“{media}”（驱动可用：{available or '驱动未提供列表'}）"
        )
    return errors
