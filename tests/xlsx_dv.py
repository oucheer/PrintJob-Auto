"""测试辅助：读取 xlsx 工作表的数据验证（含 Excel 的 x14 扩展部分）。

用 Excel / WPS 保存过含“跨表引用下拉列表”的工作簿时，这些验证会写成
``x14:dataValidations`` 扩展，openpyxl 只会告警并丢弃，导致模板验收测试
把“好好的下拉”误判成“下拉丢失”。这里额外解析 XML 补齐，让断言只看真实状态。
"""

from __future__ import annotations

import re
import zipfile

from openpyxl import load_workbook

_TRUE_FLAGS = {"1", "true"}


class DvRecord:
    """数据验证记录：统一 openpyxl 标准 DV 与 x14 扩展 DV 的字段。"""

    __slots__ = ("type", "sqref", "formula1", "formula2", "showDropDown", "showErrorMessage")

    def __init__(self, type, sqref, formula1, showDropDown=False, showErrorMessage=False, formula2=""):
        self.type = type
        self.sqref = sqref
        self.formula1 = formula1
        self.formula2 = formula2
        self.showDropDown = showDropDown
        self.showErrorMessage = showErrorMessage

    def __repr__(self) -> str:  # pragma: no cover - 便于断言失败时排查
        return f"DvRecord({self.type!r}, {self.sqref!r}, {self.formula1!r}, f2={self.formula2!r})"


def _flag(head: str, name: str) -> bool:
    found = re.search(rf'\b{name}="([^"]*)"', head)
    return bool(found) and found.group(1).lower() in _TRUE_FLAGS


def _sheet_xml_path(zf: zipfile.ZipFile, sheet_title: str) -> str | None:
    workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    targets: dict[str, str] = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels_xml):
        rid = re.search(r'Id="([^"]+)"', rel)
        tgt = re.search(r'Target="([^"]+)"', rel)
        if rid and tgt:
            targets[rid.group(1)] = tgt.group(1)

    for name, rid in re.findall(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', workbook_xml):
        if name == sheet_title and rid in targets:
            target = targets[rid].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    return None


def read_validations(path, sheet_title: str) -> list[DvRecord]:
    """返回指定工作表的数据验证记录（标准 DV + x14 扩展 DV）。"""
    records: list[DvRecord] = []

    wb = load_workbook(path)
    if sheet_title in wb.sheetnames:
        for dv in wb[sheet_title].data_validations.dataValidation:
            records.append(
                DvRecord(
                    str(dv.type or "list"),
                    str(dv.sqref or ""),
                    str(dv.formula1 if dv.formula1 is not None else ""),
                    bool(dv.showDropDown),
                    bool(dv.showErrorMessage),
                    str(dv.formula2 if dv.formula2 is not None else ""),
                )
            )

    with zipfile.ZipFile(path) as zf:
        sheet_path = _sheet_xml_path(zf, sheet_title)
        if sheet_path is None or sheet_path not in zf.namelist():
            return records
        xml = zf.read(sheet_path).decode("utf-8")

    for block in re.findall(r"<x14:dataValidation\b.*?</x14:dataValidation>", xml, re.S):
        head = block[: block.find(">")]
        sqref = re.search(r"<xm:sqref>(.*?)</xm:sqref>", block, re.S)
        formula1 = re.search(r"<x14:formula1>\s*<xm:f>(.*?)</xm:f>", block, re.S)
        formula2 = re.search(r"<x14:formula2>\s*<xm:f>(.*?)</xm:f>", block, re.S)
        type_match = re.search(r'\btype="([^"]*)"', head)
        records.append(
            DvRecord(
                type_match.group(1) if type_match else "list",
                (sqref.group(1) if sqref else "").strip(),
                (formula1.group(1) if formula1 else "").strip(),
                _flag(head, "showDropDown"),
                _flag(head, "showErrorMessage"),
                (formula2.group(1) if formula2 else "").strip(),
            )
        )
    return records
