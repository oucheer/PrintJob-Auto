from printcase.capabilities import validate_driver_capabilities
from printcase.models import PrintCase


def test_driver_capability_rejects_duplex_color_and_paper():
    case = PrintCase(case_id="TC", enabled=True, file_path="x.pdf", paper="A4", color="color", duplex="duplexlong", bin="纸盒1")
    caps = {"papers": ["Letter"], "duplex": False, "color": False}
    errors = validate_driver_capabilities(case, caps)
    assert len(errors) == 3
    assert any("纸张" in e for e in errors)
    assert any("双面" in e for e in errors)
    assert any("彩色" in e for e in errors)


def test_bin_table_from_driver_is_not_authoritative():
    """DC_BINNAMES 与属性页“来源”下拉常常不一致，不能据此拦下用例。

    本机实测：属性页下拉是 自动选择/多功能进纸器/纸盒1/纸盒2，
    DC_BINNAMES 却报 自动选择/多功能进纸器/纸盒3/纸盒4/内置大容量纸盒。
    """
    caps = {
        "bin_names": ["自动选择", "多功能进纸器", "纸盒3", "纸盒4", "内置大容量纸盒"],
        "bin_ids": [7, 262, 259, 258, 257],
    }
    for wanted in ("纸盒1", "纸盒2", "自动选择"):
        case = PrintCase(case_id="TC", enabled=True, file_path="x.pdf", paper="A4", bin=wanted)
        assert validate_driver_capabilities(case, caps) == [], wanted


def test_unknown_driver_capability_does_not_block_case():
    case = PrintCase(case_id="TC", enabled=True, file_path="x.pdf", paper="A4", duplex="duplexlong", color="color")
    assert validate_driver_capabilities(case, {"papers": None, "duplex": None, "color": None}) == []
