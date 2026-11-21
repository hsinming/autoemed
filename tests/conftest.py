import importlib

import pytest


@pytest.fixture
def main_module(tmp_path, monkeypatch):
    """Import main with a temp cwd, because main writes log.txt on import.

    Never touches the Live Portal (ADR-0002): importing main does not start a browser.
    """
    monkeypatch.chdir(tmp_path)
    return importlib.import_module("main")
