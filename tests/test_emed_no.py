import pytest


@pytest.mark.parametrize("raw, expected", [
    ("HAP47994319", "HAP 47994319"),
    ("HAP 47994319", "HAP 47994319"),
    ("  NZER12345678 ", "NZER 12345678"),
    ("CEAC2024123456", "CEAC 2024123456"),
])
def test_normalize_emed_no(main_module, raw, expected):
    assert main_module.EmedicalWorkflowManager._normalize_emed_no(raw) == expected


@pytest.mark.parametrize("emed_no, country", [
    ("HAP 47994319", "澳大利亞"),
    ("TRN 12345678", "澳大利亞"),
    ("NZER 12345678", "紐西蘭"),
    ("NZHR 12345678", "紐西蘭"),
    ("IME 12345678", "加拿大"),
    ("UMI 12345678", "加拿大"),
    ("UCI 12345678", "加拿大"),
    ("CEAC 2024123456", "美國"),
    ("XYZ 12345678", "未知國家"),
])
def test_country_from_prefix(main_module, emed_no, country):
    manager = main_module.EmedicalWorkflowManager(None, None)
    assert manager._get_country(emed_no) == country
