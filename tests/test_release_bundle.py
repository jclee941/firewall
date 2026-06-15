from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PY = sys.executable


def test_release_bundle_contains_gui_cli_and_offline_dependency_layout(tmp_path: Path) -> None:
    subprocess.run([PY, str(ROOT / "scripts" / "build_xlsm.py")], cwd=ROOT, check=True)
    subprocess.run([PY, str(ROOT / "scripts" / "make_request_folder.py")], cwd=ROOT, check=True)
    output = tmp_path / "secui-cli-gui-release.zip"

    result = subprocess.run(
        [
            PY,
            str(ROOT / "scripts" / "build_release_bundle.py"),
            "--no-rebuild",
            "--skip-wheel-download",
            "--output",
            str(output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    with zipfile.ZipFile(output) as zf:
        names = set(zf.namelist())
    assert "secui-cli-gui-release/README.md" in names
    assert "secui-cli-gui-release/requirements.txt" in names
    assert "secui-cli-gui-release/requirements-gui.txt" in names
    assert "secui-cli-gui-release/run_gui.sh" in names
    assert "secui-cli-gui-release/run_gui.bat" in names
    assert "secui-cli-gui-release/vendor/wheels/" not in names
    assert "secui-cli-gui-release/app/scripts/secui_gui.py" in names
    assert "secui-cli-gui-release/app/scripts/secui_cli.py" in names
    assert "secui-cli-gui-release/app/firewall_policy/gui_export.py" in names
    assert "secui-cli-gui-release/app/dist/firewall-policy-automation.xlsm" in names
    assert not any("__pycache__" in name for name in names)
