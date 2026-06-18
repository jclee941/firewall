from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

from scripts.build_standalone_gui import _stage_runtime, _zip_dir

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


def test_standalone_gui_zip_stages_runtime_files_with_fake_executable(tmp_path: Path) -> None:
    subprocess.run([PY, str(ROOT / "scripts" / "build_xlsm.py")], cwd=ROOT, check=True)
    subprocess.run([PY, str(ROOT / "scripts" / "make_request_folder.py")], cwd=ROOT, check=True)

    stage = tmp_path / "secui-gui-standalone"
    dist = stage / "_pyinstaller_dist"
    build = stage / "_pyinstaller_build"
    spec = stage / "_pyinstaller_spec"
    dist.mkdir(parents=True)
    build.mkdir()
    spec.mkdir()
    exe = dist / "secui-gui"
    exe.write_text("#!/bin/sh\n", encoding="utf-8")
    output = tmp_path / "secui-gui-standalone.zip"

    _stage_runtime(stage, exe)
    _zip_dir(stage, output)

    with zipfile.ZipFile(output) as zf:
        names = set(zf.namelist())
        readme = zf.read("secui-gui/README.md").decode("utf-8")

    assert "secui-gui/secui-gui" in names
    assert "secui-gui/dist/firewall-policy-automation.xlsm" in names
    assert any(name.startswith("secui-gui/request-folder/") for name in names)
    assert "secui-gui/README.md" in names
    assert "Python 설치 없이 실행" in readme
    assert "`dist/firewall-policy-automation.xlsm`" in readme
    assert "`request-folder/`" in readme
    assert not any("_pyinstaller_" in name for name in names)
