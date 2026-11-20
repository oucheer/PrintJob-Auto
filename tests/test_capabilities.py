from printcase.capabilities import validate_driver_capabilities
from printcase.models import PrintCase


def test_driver_capability_rejects_duplex_color_and_missing_tray():
    case = PrintCase(case_id="TC", enabled=True, file_path="x.pdf", paper="A4", color="color", duplex="duplexlong", bin="2")
    caps = {"papers": ["Letter"], "duplex": False, "color": False, "bin_ids": [1], "bin_names": ["Tray 1"]}
    errors = validate_driver_capabilities(case, caps)
    assert len(errors) == 4


def test_unknown_driver_capability_does_not_block_case():
    case = PrintCase(case_id="TC", enabled=True, file_path="x.pdf", paper="A4", duplex="duplexlong", color="color")
    assert validate_driver_capabilities(case, {"papers": None, "duplex": None, "color": None}) == []
