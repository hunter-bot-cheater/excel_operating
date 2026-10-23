from pathlib import Path


def test_legacy_core_copy_exists():
    root = Path(__file__).resolve().parents[2]
    assert (root / "legacy" / "original_src" / "extract_core" / "api.py").is_file()


def test_compat_package_imports_original_api():
    import sys
    root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(root))
    from extract_core.api import run_batch
    assert callable(run_batch)
