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
    # 纸盒不在这里核验：DC_BINNAMES 是驱动的“标准纸盒表”，常常与属性页“纸张→来源”
    # 下拉的真实条目对不上（本机实测：属性页只有 自动选择/多功能进纸器/纸盒1/纸盒2，
    # DC_BINNAMES 却报 纸盒3/纸盒4/内置大容量纸盒），据此拦截会误伤合法用例。
    # 纸盒最终是在驱动属性页里按名称点选的，因此改由设置属性时按下拉实际条目报错
    # （见 driver_ui._apply_action 的“下拉里没有…”提示）。
    media = str(case.media_type or "").strip()
    if media and not media.isdigit() and not _media_reported(media, caps.get("media_names")):
        available = "、".join((caps.get("media_names") or [])[:8])
        errors.append(
            f"驱动能力检查：打印机未报告纸张类型“{media}”（驱动可用：{available or '驱动未提供列表'}）"
        )
    return errors
