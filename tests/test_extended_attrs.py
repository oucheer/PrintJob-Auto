"""驱动扩展属性（移植自既有 print_panel 项目）测试：目录、旧表导入、模板、设置引擎。"""
from pathlib import Path

import pytest

from printcase.attr_catalog import ATTRIBUTE_BY_LABEL, ATTRIBUTES, LEGACY_COLUMN_PAGES, page_display
from printcase.extended_attrs import (
    ExtendedAttr,
    normalize,
    parse_blob,
    parse_legacy_row,
    split_composite,
    split_pair,
)


def test_catalog_covers_legacy_attribute_vocabulary():
    """既有项目的关键属性必须都在目录里，且页面/类型正确。"""
    expected = {
        "wps上的画像方向(纵向(P)/横向(S))": ("wps_page", "choice"),
        "页码范围": ("print_dialog", "text"),
        "份数": ("driver_basic", "int"),
        "彩色/黑白": ("driver_basic", "choice"),
        "单面/双面长边/双面短边": ("driver_basic", "choice"),
        "分辨率": ("driver_basic", "choice"),
        "逐份(0或1)": ("driver_basic", "bool"),
        "逆序(0或1)": ("driver_basic", "bool"),
        "纸张尺寸(需要和选项显示的一样)": ("driver_paper", "choice"),
        "纸张来源": ("driver_paper", "choice"),
        "纸张类型": ("driver_paper", "choice"),
        "分隔页": ("driver_paper", "choice"),
        "装订设置-订钉": ("driver_paper", "choice"),
        "多页合一(16-有边框-向下再向左)": ("driver_layout", "choice"),
        "无边距打印(0或1)": ("driver_layout", "bool"),
        "新建文本水印内容": ("driver_watermark", "text"),
        "选择水印(如果是自定义的需要先新建)": ("driver_advanced", "text"),
        "样本打印(1)": ("driver_password", "bool"),
        "延时打印(设定时间-文件名)": ("driver_password", "combo_text"),
        "封面": ("driver_cover", "bool"),
        "封面纸盒": ("driver_cover", "choice"),
    }
    for label, (page, kind) in expected.items():
        spec = ATTRIBUTE_BY_LABEL.get(label)
        assert spec is not None, label
        assert (spec.page, spec.kind) == (page, kind), label

    assert len(ATTRIBUTES) >= 55
    assert set(LEGACY_COLUMN_PAGES) == set(range(3, 12))


def test_parse_blob_and_legacy_row():
    blob = "&份数=3\n&逐份(0或1)=1\n\n&彩色/黑白=黑白\n无效行\n"
    attrs = parse_blob(blob, "driver_basic")
    assert [(a.label, a.value) for a in attrs] == [("份数", "3"), ("逐份(0或1)", "1"), ("彩色/黑白", "黑白")]

    row = (1, "标题", "", "&页码范围=1-5", "&份数=3", "&纸张来源=纸盒1", "", "", "", "", "", "1")
    parsed = parse_legacy_row(row, "1")
    assert [(a.page, a.label) for a in parsed] == [
        ("print_dialog", "页码范围"),
        ("driver_basic", "份数"),
        ("driver_paper", "纸张来源"),
    ]


def test_normalize_uses_catalog_page_when_page_missing():
    attr = normalize("", "装订设置-订钉", "双钉")
    assert attr.page == "driver_paper"
    assert page_display(attr.page) == "驱动-纸张"


def test_split_helpers():
    assert split_pair("45*45") == ("45", "45")
    assert split_pair("200x150") == ("200", "150")
    assert split_pair("毫米") is None
    assert split_composite("15-20-30-徘徊") == ["15", "20", "30", "徘徊"]
    assert split_composite("666666-徘徊-小黑子") == ["666666", "徘徊", "小黑子"]
    assert split_composite("") == []


def _write_legacy_workbook(path: Path) -> Path:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "print"
    ws.append(["序号", "用例标题", "wps软件页面布局页面-需要设置的参数", "ctrl+p打印页面要设置的参数",
               "驱动-基本页面-要设置的参数", "驱动-纸张页面-要设置的参数", "驱动-版面页面-要设置的参数",
               "驱动-新建/编辑文本水印页面的参数", "驱动-高级页面-要设置的参数",
               "驱动-密码打印-页面要设置的参数", "驱动-封页-页面要设置的参数",
               "测试次数", "成功次数", "失败次数", "作业log关键字", "打印文档路径", "延时打印prn"])
    ws.append(["1", "A4-彩色-1份", "&wps上的纸张大小=A5 (148 x 210mm)",
               "&页码范围=1-2", "&份数=1\n&彩色/黑白=黑白", "&纸张来源=纸盒1\n&纸张类型=普通纸(60-90gsm)",
               "", "", "", "&样本打印(1)=1", "&封面=1\n&类型1=空白页", 1, "", "", "status:5",
               "docs/Word1.doc", ""])
    wb.save(path)
    return path


def test_legacy_print_panel_workbook_imports(tmp_path):
    doc = tmp_path / "docs" / "Word1.doc"
    doc.parent.mkdir(parents=True)
    doc.write_bytes(b"doc")
    xlsx = _write_legacy_workbook(tmp_path / "print.xlsx")

    from printcase.excel_reader import read_cases

    cases, warnings = read_cases(xlsx)
    assert len(cases) == 1
    case = cases[0]
    assert case.case_id == "1"
    assert case.notes == "A4-彩色-1份"
    assert case.resolved_path == doc
    assert any("print_panel" in w for w in warnings)

    pairs = {(attr.page, attr.label): attr.value for attr in case.extended_attrs}
    assert pairs[("wps_page", "wps上的纸张大小")] == "A5 (148 x 210mm)"
    assert pairs[("print_dialog", "页码范围")] == "1-2"
    assert pairs[("driver_basic", "份数")] == "1"
    assert pairs[("driver_basic", "彩色/黑白")] == "黑白"
    assert pairs[("driver_paper", "纸张来源")] == "纸盒1"
    assert pairs[("driver_password", "样本打印(1)")] == "1"
    assert pairs[("driver_cover", "类型1")] == "空白页"


def test_template_ships_extended_attribute_sheet_with_dropdowns():
    from openpyxl import load_workbook

    template = Path(__file__).resolve().parents[1] / "templates" / "PrintCase_Template.xlsx"
    wb = load_workbook(template)
    assert "驱动扩展属性" in wb.sheetnames
    assert "属性速查" in wb.sheetnames

    ws = wb["驱动扩展属性"]
    headers = [cell.value for cell in ws[1]]
    assert headers == ["用例编号", "页面", "属性", "值", "说明"]
    dv_ranges = {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}
    assert "B2:B500" in dv_ranges and "C2:C500" in dv_ranges
    assert all(dv.type == "list" and not dv.showDropDown for dv in dv_ranges.values())

    # 属性下拉必须覆盖目录里的全部属性
    label_dv = dv_ranges["C2:C500"]
    source = str(label_dv.formula1)
    assert source.startswith("选项!$")
    option_col = source.split("$")[1]
    options = [cell.value for cell in wb["选项"][option_col] if cell.value]
    for spec in ATTRIBUTES:
        assert spec.label in options, spec.label

    # 示例行必须能被读回
    from printcase.excel_reader import read_cases

    cases, warns = read_cases(template)
    assert warns == []
    by_id = {case.case_id: case for case in cases}
    assert len(by_id["TC-001"].extended_attrs) >= 4
    assert any(attr.label == "无边距打印(0或1)" for attr in by_id["TC-001"].extended_attrs)


# ---------- 设置引擎（替身驱动窗口） ----------

class FakeControl:
    control_type = ""

    def __init__(self, text="", parent=None, control_type=""):
        self.text = text
        self.control_type = control_type
        self.children = []
        self._parent = parent
        if parent is not None:
            parent.children.append(self)

    def window_text(self):
        return self.text

    def parent(self):
        return self._parent

    def descendants(self, control_type=None, **_kwargs):
        found = []
        for child in self.children:
            if not control_type or child.control_type == control_type:
                found.append(child)
            found.extend(child.descendants(control_type=control_type))
        return found


class FakeCheck(FakeControl):
    def __init__(self, text, parent=None, checked=0):
        super().__init__(text, parent, "CheckBox")
        self.checked = checked

    def get_toggle_state(self):
        return self.checked

    def toggle(self):
        self.checked = 1 - self.checked

    def click_input(self):
        self.toggle()


class FakeCombo(FakeControl):
    def __init__(self, text, options, parent=None):
        super().__init__(text, parent, "ComboBox")
        self.options = list(options)
        self.selected = None

    def texts(self):
        return list(self.options)

    def select(self, value):
        if value not in self.options:
            raise ValueError(value)
        self.selected = value


class FakeEdit(FakeControl):
    def __init__(self, label, parent=None):
        FakeControl(label, parent, "Text")     # 标签必须排在输入框之前
        super().__init__("", parent, "Edit")
        self.value = ""

    def set_edit_text(self, text):
        self.value = text


class FakeTab(FakeControl):
    def __init__(self, text, parent=None):
        super().__init__(text, parent, "TabItem")
        self.clicked = 0

    def click_input(self):
        self.clicked += 1


def build_driver_window():
    window = FakeControl("打印首选项")
    FakeTab("基本", window)
    FakeTab("纸张", window)
    FakeTab("版面", window)
    FakeTab("水印", window)
    FakeTab("高级", window)
    FakeTab("密码打印", window)
    FakeTab("封页", window)

    FakeCheck("逆序", window)
    FakeCheck("无边距打印", window)
    FakeCombo("分隔页", ["无", "作业之间", "份数之间"], window)
    FakeEdit("水印标题", window)
    return window


def test_apply_driver_attrs_sets_checkbox_combo_and_text():
    from printcase.ui_print import apply_driver_attrs

    window = build_driver_window()
    logs = []
    attrs = [
        ExtendedAttr("driver_basic", "逆序(0或1)", "1"),
        ExtendedAttr("driver_layout", "无边距打印(0或1)", "1"),
        ExtendedAttr("driver_paper", "分隔页", "作业之间"),
        ExtendedAttr("driver_watermark", "水印标题", "冒烟测试"),
    ]
    unresolved = apply_driver_attrs(window, attrs, logs.append)
    assert unresolved == []

    checkboxes = {c.window_text(): c.checked for c in window.descendants(control_type="CheckBox")}
    assert checkboxes["逆序"] == 1 and checkboxes["无边距打印"] == 1
    combo = window.descendants(control_type="ComboBox")[0]
    assert combo.selected == "作业之间"
    edits = window.descendants(control_type="Edit")
    assert edits[0].value == "冒烟测试"
    assert any("切换到“纸张”页" in line for line in logs)


def test_apply_driver_attrs_reports_unresolved():
    from printcase.ui_print import apply_driver_attrs

    window = build_driver_window()
    attrs = [
        ExtendedAttr("driver_paper", "分隔页", "不存在的选项"),
        ExtendedAttr("driver_basic", "接纸架", "接纸架2"),
    ]
    unresolved = apply_driver_attrs(window, attrs, lambda _t: None)
    assert len(unresolved) == 2
    assert any("分隔页" in item for item in unresolved)


class FakePageSetup:
    def __init__(self):
        self.Orientation = 0
        self.PaperSize = 0
        self.PageWidth = 0
        self.PageHeight = 0


class FakeDocument:
    def __init__(self):
        self.PageSetup = FakePageSetup()


def test_apply_wps_page_attrs_via_com():
    from printcase.ui_print import apply_wps_page_attrs

    document = FakeDocument()
    logs = []
    attrs = [
        ExtendedAttr("wps_page", "wps上的画像方向(纵向(P)/横向(S))", "横向(S)"),
        ExtendedAttr("wps_page", "wps上的纸张大小", "A5 (148 x 210mm)"),
        ExtendedAttr("wps_page", "wps上的预览", "整篇文档"),
    ]
    unresolved = apply_wps_page_attrs(document, attrs, logs.append)
    assert unresolved == []
    assert document.PageSetup.Orientation == 1
    assert document.PageSetup.PaperSize == 11
    assert any("预览范围" in line for line in logs)

    bad = apply_wps_page_attrs(FakeDocument(), [ExtendedAttr("wps_page", "wps上的纸张大小", "未知尺寸")], logs.append)
    assert bad and "未知尺寸" in bad[0]


# ---------- 以既有项目“参数对照表”为准的取值回归 ----------

# 摘自既有 print_panel 项目 print.xlsx / 参数对照表：真实用例里出现过的属性取值
LEGACY_REFERENCE_VALUES = {
    "wps上的画像方向(纵向(P)/横向(S))": "横向(S)",
    "wps上的预览": "整篇文档",
    "wps上的纸张大小": "A4 (210 x 297mm)",
    "wps上的自定义纸大小(宽*高)": "25*45",
    "页码范围": "1-5",
    "反片打印": "1",
    "份数": "3",
    "单面/双面长边/双面短边": "双面短边",
    "分辨率": "1200 dpi",
    "逐份(0或1)": "1",
    "逆序(0或1)": "1",
    "跳过空白页(是或否)": "是",
    "混合尺寸(0或1)": "0",
    "接纸架": "接纸架2",
    "偏移(开启或关闭)": "关闭",
    "纸张尺寸(需要和选项显示的一样)": "A4 (210 x 297mm)",
    "纸张来源": "自动选择",
    "纸张类型": "普通纸(60-90gsm)",
    "分隔页": "作业之间",
    "OHP插页(开启/关闭)": "关闭",
    "分隔页（OHP插页）纸张来源": "纸盒1",
    "自定义纸张类型": "自定义3",
    "尺寸": "毫米",
    "自定义尺寸名称": "计划2",
    "长": "200",
    "宽": "300",
    "装订设置-翻页方式": "上翻页",
    "装订设置-订钉": "双钉",
    "装订设置-打孔": "4孔",
    "装订设置-折叠": "Z折",
    "装订设置-折叠张数": "9张",
    "驱动上面的画像方向(横向/纵向)": "横向",
    "多页合一(16-有边框-向下再向左)": "16-有边框-向下再向左",
    "小册子(右边_打印在封面外侧_维持原纸张尺寸大小_纸张类型_纸张来源)": "右边_打印在封面外侧_维持原纸张尺寸大小_普通纸(60-90gsm)_纸盒1",
    "海报(2x2)": "3x3",
    "省墨(开启/关闭)": "关闭",
    "浓度(1~11)": "4",
    "旋转180度(0或1)": "1",
    "缩放": "40",
    "镜像打印(0或1)": "1",
    "无边距打印(0或1)": "1",
    "点击新建文本水印/编辑水印/查看水印": "新建文本水印",
    "新建文本水印角度": "30",
    "新建文本水印居中(1)": "1",
    "新建文本水印内容": "阴天",
    "指定起点坐标(x*y)": "45*45",
    "包含动态水印(1-打印日期-打印用户名称)": "1-打印日期-打印用户名称",
    "新建文本水印加粗(0或1)": "1",
    "新建文本水印斜体字(0或1)": "1",
    "水印文字尺寸": "200",
    "水印文字浓度": "最浓",
    "水印文字体": "仿宋",
    "水印标题": "哈哈哈哈哈",
    "选择水印(如果是自定义的需要先新建)": "哈哈哈哈哈",
    "删除水印": "哈哈哈哈哈",
    "仅第一页显示(0或1)": "1",
    "以位图打印TrueType字体(0或1)": "1",
    "使用打印机字体代替TrueType字体(1)": "1",
    "PDF快捷打印(0或1)": "1",
    "普通打印(1)": "1",
    "样本打印(1)": "1",
    "延时打印(设定时间-文件名)": "15-20-30-徘徊",
    "密码打印(密码-作业名-用户名)": "666666-徘徊-小黑子",
    "封面": "1",
    "类型1": "空白页",
    "封面纸盒": "多功能进纸器",
    "封底": "1",
    "类型2": "空白页",
}


def test_catalog_selects_every_value_used_by_legacy_cases():
    """以既有项目真实用例为准：出现过的取值必须能在目录里选中或按文本写入。"""
    from printcase.extended_attrs import bool_value, split_composite, split_pair

    for label, value in LEGACY_REFERENCE_VALUES.items():
        spec = ATTRIBUTE_BY_LABEL.get(label)
        assert spec is not None, f"缺少属性：{label}"
        if spec.kind == "bool":
            assert bool_value(value) is not None, (label, value)
        elif spec.kind == "choice":
            assert value in spec.options, f"{label} 缺少选项 {value}（现有：{spec.options}）"
        elif spec.kind == "pair":
            assert split_pair(value) is not None, (label, value)
        elif spec.kind == "combo_text":
            assert len(split_composite(value)) >= 2, (label, value)
        else:  # text / int 自由输入
            assert value.strip(), (label, value)
