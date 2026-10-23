from pathlib import Path


def test_legacy_core_copy_exists():
    root = Path(__file__).resolve().parents[2]
    assert (root / "legacy" / "original_src" / "检验报告模板转换" / "converter.py").is_file()


def test_compat_entry_exports_original_api():
    import importlib.util
    root = Path(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("special_format_converter_compat_test", root / "converter.py")
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    assert callable(module.convert_file)
