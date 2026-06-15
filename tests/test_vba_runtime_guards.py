import re
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
VBA_POLICY = ROOT / "vba" / "FirewallPolicyAutomation.bas"
XLSM = ROOT / "dist" / "firewall-policy-automation.xlsm"


def _build_workbook() -> None:
    _ = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "build_xlsm.py")],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )


def _macro_body(src: str, name: str) -> str:
    start = src.find(f"Private Sub {name}(")
    if start == -1:
        start = src.find(f"Public Sub {name}(")
    assert start != -1, f"{name} not found"
    end = src.find("End Sub", start)
    assert end != -1, f"{name} End Sub not found"
    return src[start:end]


def test_auto_run_vba_does_not_select_or_freeze_active_window():
    src = VBA_POLICY.read_text(encoding="utf-8")
    for name in ("AutoRunWorkbookOutputs", "ConvertRequestsToSecuiCli", "FormatSecuiCliSheet"):
        body = _macro_body(src, name)
        assert ".Activate" not in body
        assert ".Select" not in body
        assert "ActiveWindow" not in body


def test_secui_cli_vba_assigns_dictionary_objects_with_set():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert "Set serviceFanoutIndex(sourceDestinationKey) = services" in src
    assert "Set cliGroups(groupKey) = group" in src
    assert not re.search(r"^\s*serviceFanoutIndex\(sourceDestinationKey\)\s*=\s*services\s*$", src, re.MULTILINE)
    assert not re.search(r"^\s*cliGroups\(groupKey\)\s*=\s*group\s*$", src, re.MULTILINE)


def test_auto_run_output_sheets_are_not_protected():
    _build_workbook()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True)
    try:
        macro_written = (
            "requests",
            "processing_log",
            "secui_cli",
        )
        for sheet_name in macro_written:
            assert wb[sheet_name].protection.sheet is False, (
                f"{sheet_name} is written by auto-run VBA and must not be protected"
            )
    finally:
        wb.close()


def test_generated_workbook_opens_with_libreoffice_headless(tmp_path: Path):
    soffice = shutil.which("soffice")
    if soffice is None:
        pytest.skip("LibreOffice soffice is not installed")

    _build_workbook()
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    profile_dir = tmp_path / "lo-profile"
    result = subprocess.run(
        [
            soffice,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--norestore",
            "--nofirststartwizard",
            "--nodefault",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(XLSM),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
        timeout=120,
    )
    converted_files = list(output_dir.glob("*.xlsx"))
    diagnostic = result.stdout + result.stderr + f"\nconverted_files={converted_files}"
    assert result.returncode == 0, diagnostic
    assert converted_files, diagnostic
