"""Ejecuta los escenarios de extremo a extremo originales (scripts) como pruebas."""
import os
import subprocess
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).parent


@pytest.mark.parametrize("script", ["smoke.py", "precision.py", "new_features.py", "mcp_client.py"])
def test_script(script):
    r = subprocess.run([sys.executable, str(HERE / script)], capture_output=True, text=True, encoding="utf-8", timeout=300,
                       env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8", "APA7_OFFLINE": "1"})
    assert r.returncode == 0, r.stdout[-1500:] + r.stderr[-1500:]
