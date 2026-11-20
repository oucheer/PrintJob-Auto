"""驱动扩展属性：模型、解析与取值规范化。

数据来源有两条：
1. 本项目的“驱动扩展属性”工作表（长表：用例编号 / 页面 / 属性 / 值），一行一个属性；
2. 既有 print_panel 项目的 `print` 工作表（9 列里塞 `&属性名=值` 多行文本）——只读导入，
   便于把历史用例直接搬过来。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .attr_catalog import (
    ATTRIBUTE_BY_LABEL,
    LEGACY_COLUMN_PAGES,
    AttributeSpec,
    attributes_for_page,
    page_display,
    page_key_by_display,
)

_BLOB_LINE = re.compile(r"^&(?P<label>[^=]+)=(?P<value>.*)$")


@dataclass(frozen=True)
class ExtendedAttr:
    page: str          # 内部页面标识（attr_catalog.PAGES 的 key）
    label: str         # 驱动/应用界面上的属性名
    value: str

    def spec(self) -> AttributeSpec | None:
        return ATTRIBUTE_BY_LABEL.get(self.label)

    def describe(self) -> str:
        return f"{page_display(self.page)}｜{self.label}={self.value}"


def parse_blob(text: str, page: str) -> list[ExtendedAttr]:
    """解析既有项目的 `&属性名=值` 多行文本块。"""
    attrs: list[ExtendedAttr] = []
    for raw_line in str(text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = _BLOB_LINE.match(line)
        if not match:
            continue
        label = match.group("label").strip()
        value = match.group("value").strip()
        attrs.append(ExtendedAttr(page=page, label=label, value=value))
    return attrs


def parse_legacy_row(row: tuple, case_id: str) -> list[ExtendedAttr]:
    """把既有项目一行（列 3–11 为属性块）解析成扩展属性列表。"""
    attrs: list[ExtendedAttr] = []
    for column, page in LEGACY_COLUMN_PAGES.items():
        index = column - 1
        if index >= len(row):
            continue
        attrs.extend(parse_blob(row[index], page))
    return attrs


def normalize(page_text: str, label: str, value) -> ExtendedAttr:
    """把模板单元格里的（页面显示名、属性名、值）规范化为 ExtendedAttr。"""
    page_key = page_key_by_display(str(page_text or "").strip())
    label_text = str(label or "").strip()
    spec = ATTRIBUTE_BY_LABEL.get(label_text)
    if spec is not None:
        page_key = spec.page
    return ExtendedAttr(page=page_key, label=label_text, value=str(value if value is not None else "").strip())


def group_by_page(attrs: list[ExtendedAttr]) -> dict[str, list[ExtendedAttr]]:
    grouped: dict[str, list[ExtendedAttr]] = {}
    for attr in attrs:
        grouped.setdefault(attr.page, []).append(attr)
    return grouped


def base_label(label: str) -> str:
    """去掉括号里的取值说明：'逐份(0或1)' → '逐份'。"""
    return str(label or "").split("(")[0].strip()


def split_pair(value: str) -> tuple[str, str] | None:
    """拆分 '25*45' / '45*45' / '200x150' 这类两段数值。"""
    text = str(value or "").strip()
    for sep in ("*", "x", "X", "×", ","):
        if sep in text:
            left, right = text.split(sep, 1)
            if left.strip() and right.strip():
                return left.strip(), right.strip()
    return None


def split_composite(value: str) -> list[str]:
    """拆分 '15-20-30-徘徊' / '666666-徘徊-小黑子' 这类用 - 连接的复合值。"""
    text = str(value or "").strip()
    if not text or "-" not in text:
        return [text] if text else []
    return [part.strip() for part in text.split("-") if part.strip()]


def bool_value(value: str) -> bool | None:
    text = str(value or "").strip().casefold()
    if text in {"1", "是", "yes", "y", "true", "开", "开启", "on"}:
        return True
    if text in {"0", "否", "no", "n", "false", "关", "关闭", "off", ""}:
        return False
    return None


def catalog_for_template() -> tuple[list[str], list[str]]:
    """给模板下拉用：页面显示名列表、全部属性名列表。"""
    pages = [page_display(page) for page in LEGACY_COLUMN_PAGES.values()]
    labels = [spec.label for spec in ATTRIBUTE_BY_LABEL.values()]
    return pages, labels


def default_examples(case_id: str) -> list[tuple[str, str, str, str]]:
    """模板示例行：(用例编号, 页面, 属性, 值)。"""
    examples: list[tuple[str, str, str, str]] = []
    for attr in (
        ExtendedAttr("driver_layout", "无边距打印(0或1)", "1"),
        ExtendedAttr("driver_basic", "逆序(0或1)", "1"),
        ExtendedAttr("driver_paper", "分隔页", "作业之间"),
        ExtendedAttr("driver_layout", "多页合一(16-有边框-向下再向左)", "2页合一"),
    ):
        spec = ATTRIBUTE_BY_LABEL.get(attr.label)
        page = page_display(spec.page if spec else attr.page)
        examples.append((case_id, page, attr.label, attr.value))
    return examples


__all__ = [
    "ExtendedAttr",
    "attributes_for_page",
    "base_label",
    "bool_value",
    "catalog_for_template",
    "default_examples",
    "group_by_page",
    "normalize",
    "parse_blob",
    "parse_legacy_row",
    "split_composite",
    "split_pair",
]
