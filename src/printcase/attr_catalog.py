"""驱动扩展属性目录（移植自既有 print_panel 项目的属性词表）。

既有项目用 `&属性名=值` 的多行文本放在 9 列里（WPS 页面布局 / Ctrl+P 打印页 /
驱动基本 / 纸张 / 版面 / 水印 / 高级 / 密码打印 / 封页）。这里把它整理成结构化
目录：每个属性带所属页面、控件类型、取值范围与示例，供模板下拉、读取器校验和
UI 自动化设置共用。

约定：
- kind="bool"   值 0/1（勾选框）
- kind="choice" 从下拉列表选（值为选项文本）
- kind="int"    数字输入框
- kind="text"   文本框
- kind="pair"   形如 “宽*高”或“x*y”的两段数值
- kind="combo_text" 复合文本（如“密码-作业名-用户名”）
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class AttributeSpec:
    label: str            # 模板中显示的属性名（与驱动界面文字保持一致）
    page: str             # 所属页面
    kind: str             # bool / choice / int / text / pair / combo_text
    options: tuple[str, ...] = field(default_factory=tuple)
    example: str = ""
    hint: str = ""


# 页面定义：key = 内部标识，value = (模板显示名, 驱动/应用界面标签关键字)
PAGES: dict[str, tuple[str, tuple[str, ...]]] = {
    "wps_page": ("WPS 页面布局", ("页面布局", "纸张方向", "纸张大小")),
    "print_dialog": ("Ctrl+P 打印页", ("打印", "print")),
    "driver_basic": ("驱动-基本", ("基本", "basic", "常用")),
    "driver_paper": ("驱动-纸张", ("纸张", "paper")),
    "driver_layout": ("驱动-版面", ("版面", "layout", "布局")),
    "driver_watermark": ("驱动-水印", ("水印", "watermark")),
    "driver_advanced": ("驱动-高级", ("高级", "advanced")),
    "driver_password": ("驱动-密码/作业类型", ("密码打印", "作业类型", "job type", "secure")),
    "driver_cover": ("驱动-封页", ("封页", "封面", "cover")),
}

PAPER_OPTIONS = (
    "A4 (210 x 297mm)", "A3 (297 x 420mm)", "A5 (148 x 210mm)", "A6 (105 x 148mm)",
    "JIS B4 (257 x 364mm)", "JIS B5 (182 x 257mm)", "JIS B6 (128 x 182mm)",
    "Letter (8.5 x 11inch)", "Legal (8.5 x 14inch)", "Folio (8.5 x 13inch)",
    "Ledger (11 x 17inch)", "Executive (7.25 x 10.5inch)", "Statement (5.5 x 8.5inch)",
    "8K (270 x 390mm)", "Big16K (195 x 270mm)", "A3Wide (12 x 18inch)", "SRA3 (320 x 450mm)",
)
MEDIA_OPTIONS = (
    "普通纸(60-90gsm)", "普通纸+(91-105gsm)", "厚纸1(106-120gsm)", "厚纸1+(121-157gsm)",
    "厚纸2(158-209gsm)", "厚纸3(210-256gsm)", "信封纸", "明信片", "标签纸", "透明胶片",
)
SOURCE_OPTIONS = ("自动选择", "多功能进纸器", "纸盒1", "纸盒2")
DUPLEX_OPTIONS = ("单面", "双面长边", "双面短边")
ON_OFF = ("开启", "关闭")
YES_NO = ("是", "否")
BIND_TURN = ("上翻页", "下翻页", "左翻页", "右翻页")
BIND_STAPLE = ("无", "单钉", "双钉")
BIND_PUNCH = ("无", "2孔", "4孔")
BIND_FOLD = ("无", "Z折", "双折", "三折")
BOOKLET = (
    "关闭",
    "左边_打印在封面外侧_维持原纸张尺寸大小_普通纸(60-90gsm)_自动选择",
    "右边_打印在封面外侧_维持原纸张尺寸大小_普通纸(60-90gsm)_纸盒1",
)
MULTI_PAGE = ("关闭", "2页合一", "4页合一", "9页合一", "16页合一", "16-有边框-向下再向左")
POSTER = ("关闭", "2x2", "3x3", "4x4")
WATERMARK_DENSITY = ("最浓", "浓", "标准", "淡", "最淡")
WATERMARK_FONT = ("宋体", "仿宋", "黑体", "楷体", "微软雅黑", "Arial")
JOB_TYPE = ("普通打印", "样本打印", "延时打印", "密码打印")


def _specs() -> list[AttributeSpec]:
    return [
        # ---- WPS 页面布局 ----
        AttributeSpec("wps上的画像方向(纵向(P)/横向(S))", "wps_page", "choice", ("纵向(P)", "横向(S)"), "横向(S)", "WPS 页面布局里的纸张方向"),
        AttributeSpec("wps上的纸张大小", "wps_page", "choice", PAPER_OPTIONS, "A5 (148 x 210mm)", "与 WPS 下拉文本一致"),
        AttributeSpec("wps上的预览", "wps_page", "choice", ("整篇文档", "当前页", "选定内容"), "整篇文档", "打印预览范围"),
        AttributeSpec("wps上的自定义纸大小(宽*高)", "wps_page", "pair", (), "25*45", "单位毫米，宽*高"),

        # ---- Ctrl+P 打印页 ----
        AttributeSpec("页码范围", "print_dialog", "text", (), "1-5", "留空=全部；与用例“页码”列等价"),
        AttributeSpec("反片打印", "print_dialog", "bool", (), "1", "镜像底片输出"),
        AttributeSpec("打印到文件", "print_dialog", "bool", (), "0", "输出到文件而不是打印机"),

        # ---- 驱动-基本 ----
        AttributeSpec("份数", "driver_basic", "int", (), "3", "与用例“份数”列等价"),
        AttributeSpec("以黑色打印文本(0或1)", "driver_basic", "bool", (), "1", "彩色内容中的文字用黑色"),
        AttributeSpec("彩色/黑白", "driver_basic", "choice", ("彩色", "黑白"), "黑白", ""),
        AttributeSpec("单面/双面长边/双面短边", "driver_basic", "choice", DUPLEX_OPTIONS, "双面短边", "与用例“单双面”列等价"),
        AttributeSpec("分辨率", "driver_basic", "choice", ("600 dpi", "1200 dpi"), "1200 dpi", "与用例“打印质量”对应"),
        AttributeSpec("逐份(0或1)", "driver_basic", "bool", (), "1", "开启=逐份；关闭=非逐份"),
        AttributeSpec("逆序(0或1)", "driver_basic", "bool", (), "1", "从最后一页开始输出"),
        AttributeSpec("跳过空白页(是或否)", "driver_basic", "choice", YES_NO, "是", ""),

        # ---- 驱动-纸张 ----
        AttributeSpec("混合尺寸(0或1)", "driver_paper", "bool", (), "0", "同用例“混合尺寸”"),
        AttributeSpec("接纸架", "driver_paper", "choice", ("关闭", "接纸架1", "接纸架2", "接纸架3"), "接纸架2", ""),
        AttributeSpec("偏移(开启或关闭)", "driver_paper", "choice", ON_OFF, "关闭", "奇偶页错位输出"),
        AttributeSpec("纸张尺寸(需要和选项显示的一样)", "driver_paper", "choice", PAPER_OPTIONS, "A4 (210 x 297mm)", "必须与驱动选项文字完全一致"),
        AttributeSpec("纸张来源", "driver_paper", "choice", SOURCE_OPTIONS, "自动选择", "同用例“纸盒”"),
        AttributeSpec("纸张类型", "driver_paper", "choice", MEDIA_OPTIONS, "普通纸(60-90gsm)", "同用例“纸张类型”"),
        AttributeSpec("输出模式", "driver_paper", "choice", ("正常", "分隔页", "OHP插页"), "分隔页", "纸张页“输出”分组；选“分隔页”后下面的分隔页下拉才可选（纸张来源必须是具体纸盒）"),
        AttributeSpec("分隔页", "driver_paper", "choice", ("份数之间", "作业之间"), "作业之间", "驱动硬分隔页：驱动只有这两项，选“无”等于把“输出模式”切回“正常”；纸张来源=自动选择时该项灰化不可选"),
        AttributeSpec("分隔页（OHP插页）纸张来源", "driver_paper", "choice", SOURCE_OPTIONS, "纸盒1", "“分隔页”下面那行的纸张来源（来源不能是自动选择）"),
        AttributeSpec("OHP插页(开启/关闭)", "driver_paper", "choice", ON_OFF, "关闭", ""),
        AttributeSpec("自定义纸张类型", "driver_paper", "choice", ("自定义1", "自定义2", "自定义3"), "自定义3", ""),
        AttributeSpec("尺寸", "driver_paper", "choice", ("毫米", "英寸"), "毫米", "自定义尺寸单位"),
        AttributeSpec("自定义尺寸名称", "driver_paper", "text", (), "计划2", ""),
        AttributeSpec("长", "driver_paper", "int", (), "200", "自定义纸张长边（单位见“尺寸”）"),
        AttributeSpec("宽", "driver_paper", "int", (), "300", "自定义纸张宽度"),
        AttributeSpec("装订设置-翻页方式", "driver_paper", "choice", BIND_TURN, "上翻页", ""),
        AttributeSpec("装订设置-订钉", "driver_paper", "choice", BIND_STAPLE, "双钉", ""),
        AttributeSpec("装订设置-打孔", "driver_paper", "choice", BIND_PUNCH, "4孔", ""),
        AttributeSpec("装订设置-折叠", "driver_paper", "choice", BIND_FOLD, "Z折", ""),
        AttributeSpec("装订设置-折叠张数", "driver_paper", "text", (), "9张", "如 9张"),

        # ---- 驱动-版面 ----
        AttributeSpec("驱动上面的画像方向(横向/纵向)", "driver_layout", "choice", ("横向", "纵向"), "横向", ""),
        AttributeSpec("多页合一(16-有边框-向下再向左)", "driver_layout", "choice", MULTI_PAGE, "16-有边框-向下再向左", ""),
        AttributeSpec("小册子(右边_打印在封面外侧_维持原纸张尺寸大小_纸张类型_纸张来源)", "driver_layout", "choice", BOOKLET, BOOKLET[2], "复合选项，需与驱动一致"),
        AttributeSpec("海报(2x2)", "driver_layout", "choice", POSTER, "3x3", ""),
        AttributeSpec("省墨(开启/关闭)", "driver_layout", "choice", ON_OFF, "关闭", ""),
        AttributeSpec("浓度(1~11)", "driver_layout", "int", (), "4", "1–11"),
        AttributeSpec("旋转180度(0或1)", "driver_layout", "bool", (), "1", ""),
        AttributeSpec("缩放", "driver_layout", "int", (), "40", "百分比"),
        AttributeSpec("镜像打印(0或1)", "driver_layout", "bool", (), "1", ""),
        AttributeSpec("无边距打印(0或1)", "driver_layout", "bool", (), "1", "真实勾选驱动无边距（同用例“无边距”）"),

        # ---- 驱动-水印 ----
        AttributeSpec("点击新建文本水印/编辑水印/查看水印", "driver_watermark", "choice", ("新建文本水印", "编辑水印", "查看水印"), "新建文本水印", "先点击对应按钮再填后续水印属性"),
        AttributeSpec("新建文本水印角度", "driver_watermark", "int", (), "30", ""),
        AttributeSpec("新建文本水印居中(1)", "driver_watermark", "bool", (), "1", ""),
        AttributeSpec("新建文本水印内容", "driver_watermark", "text", (), "阴天", ""),
        AttributeSpec("指定起点坐标(x*y)", "driver_watermark", "pair", (), "45*45", "不居中时的起点坐标"),
        AttributeSpec("包含动态水印(1-打印日期-打印用户名称)", "driver_watermark", "choice", ("关闭", "1-打印日期-打印用户名称"), "1-打印日期-打印用户名称", ""),
        AttributeSpec("新建文本水印加粗(0或1)", "driver_watermark", "bool", (), "1", ""),
        AttributeSpec("新建文本水印斜体字(0或1)", "driver_watermark", "bool", (), "1", ""),
        AttributeSpec("水印文字尺寸", "driver_watermark", "int", (), "200", ""),
        AttributeSpec("水印文字浓度", "driver_watermark", "choice", WATERMARK_DENSITY, "最浓", ""),
        AttributeSpec("水印文字体", "driver_watermark", "choice", WATERMARK_FONT, "仿宋", ""),
        AttributeSpec("水印标题", "driver_watermark", "text", (), "哈哈哈哈哈", "已保存水印的名称"),

        # ---- 驱动-高级 ----
        AttributeSpec("选择水印(如果是自定义的需要先新建)", "driver_advanced", "text", (), "哈哈哈哈哈", "填已保存的水印名"),
        AttributeSpec("删除水印", "driver_advanced", "text", (), "哈哈哈哈哈", "填要删除的水印名"),
        AttributeSpec("仅第一页显示(0或1)", "driver_advanced", "bool", (), "1", ""),
        AttributeSpec("以位图打印TrueType字体(0或1)", "driver_advanced", "bool", (), "1", ""),
        AttributeSpec("使用打印机字体代替TrueType字体(1)", "driver_advanced", "bool", (), "1", ""),
        AttributeSpec("PDF快捷打印(0或1)", "driver_advanced", "bool", (), "1", ""),

        # ---- 驱动-密码/作业类型 ----
        AttributeSpec("普通打印(1)", "driver_password", "bool", (), "1", "与样本/延时/密码打印互斥"),
        AttributeSpec("样本打印(1)", "driver_password", "bool", (), "1", "与普通/延时/密码打印互斥"),
        AttributeSpec("延时打印(设定时间-文件名)", "driver_password", "combo_text", (), "15-20-30-徘徊", "时间-文件名，多段用 - 分隔"),
        AttributeSpec("密码打印(密码-作业名-用户名)", "driver_password", "combo_text", (), "666666-徘徊-小黑子", "密码-作业名-用户名"),

        # ---- 驱动-封页 ----
        AttributeSpec("封面", "driver_cover", "bool", (), "1", ""),
        AttributeSpec("类型1", "driver_cover", "choice", ("空白页", "当前页"), "空白页", "封面类型"),
        AttributeSpec("封面纸盒", "driver_cover", "choice", SOURCE_OPTIONS, "多功能进纸器", ""),
        AttributeSpec("封底", "driver_cover", "bool", (), "1", ""),
        AttributeSpec("类型2", "driver_cover", "choice", ("空白页", "当前页"), "空白页", "封底类型"),
        AttributeSpec("封底纸盒", "driver_cover", "choice", SOURCE_OPTIONS, "纸盒1", ""),
    ]


ATTRIBUTES: tuple[AttributeSpec, ...] = tuple(_specs())
ATTRIBUTE_BY_LABEL: dict[str, AttributeSpec] = {spec.label: spec for spec in ATTRIBUTES}

# 兼容旧模板的表头别名（旧项目 sheet “print” 的列 → 页面）
LEGACY_COLUMN_PAGES = {
    3: "wps_page",
    4: "print_dialog",
    5: "driver_basic",
    6: "driver_paper",
    7: "driver_layout",
    8: "driver_watermark",
    9: "driver_advanced",
    10: "driver_password",
    11: "driver_cover",
}


def page_display(page: str) -> str:
    return PAGES.get(page, (page, ()))[0]


def page_keywords(page: str) -> tuple[str, ...]:
    return PAGES.get(page, (page, ()))[1]


def page_options() -> list[str]:
    return [display for display, _ in PAGES.values()]


def page_key_by_display(display: str) -> str:
    for key, (name, _words) in PAGES.items():
        if name == display:
            return key
    return display


def attributes_for_page(page: str) -> list[AttributeSpec]:
    return [spec for spec in ATTRIBUTES if spec.page == page]


def is_known_label(label: str) -> bool:
    return str(label or "").strip() in ATTRIBUTE_BY_LABEL
