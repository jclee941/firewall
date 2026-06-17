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
    if start == -1:
        start = src.find(f"Private Function {name}(")
    if start == -1:
        start = src.find(f"Public Function {name}(")
    assert start != -1, f"{name} not found"
    end_sub = src.find("End Sub", start)
    end_function = src.find("End Function", start)
    candidates = [pos for pos in (end_sub, end_function) if pos != -1]
    assert candidates, f"{name} end marker not found"
    return src[start:min(candidates)]


def test_auto_run_vba_does_not_select_or_freeze_active_window():
    src = VBA_POLICY.read_text(encoding="utf-8")
    for name in ("AutoRunWorkbookOutputs", "ConvertRequestsToSecuiCli", "FormatSecuiCliSheet"):
        body = _macro_body(src, name)
        assert ".Activate" not in body
        assert not re.search(r"\.Select\b", body)
        assert "ActiveWindow" not in body
    assert ".Activate" not in src
    assert not re.search(r"\.Select\b", src)
    assert "ActiveWindow" not in src


def test_vba_does_not_raise_user_errors_as_excel_1004():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert "vbObjectError + 1004" not in src


def test_auto_run_suppresses_per_file_error_dialogs():
    src = VBA_POLICY.read_text(encoding="utf-8")
    body = _macro_body(src, "MergeWorkbookFile")
    assert "If Not mSuppressMessages Then MsgBox" in body


def test_generated_workbook_binds_f9_to_full_output_workflow():
    src = VBA_POLICY.read_text(encoding="utf-8")
    body = _macro_body(src, "RunFirewallAutomationOutputs")
    assert "AutoRunWorkbookOutputs" in body

    _build_workbook()
    from pyopenvba import excel as ex

    workbook = ex.ExcelFile(str(XLSM))
    try:
        this_workbook = workbook.get_module("ThisWorkbook")
    finally:
        workbook.close()

    assert 'Application.OnKey "{F9}", FirewallAutomationHotkeyTarget()' in this_workbook
    assert 'Application.OnKey "{F9}"' in this_workbook
    assert "'!RunFirewallAutomationOutputs" in this_workbook


def test_runtime_vba_does_not_touch_excel_autofilter_state():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert ".AutoFilter" not in src
    assert ".AutoFilterMode" not in src
    assert ".FilterMode" not in src
    assert ".ShowAllData" not in src


def test_vba_avoids_office_enum_compile_dependencies():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert "XlSheetVisibility" not in src
    assert "msoFileDialogFolderPicker" not in src
    assert "FILE_DIALOG_FOLDER_PICKER" in src
    assert "SHEET_VISIBLE" in src
    assert "SHEET_HIDDEN" in src


def test_vba_procedure_declarations_are_not_single_long_byval_lines():
    src = VBA_POLICY.read_text(encoding="utf-8")
    long_declarations = [
        line for line in src.splitlines()
        if "ByVal" in line and (line.startswith("Private ") or line.startswith("Public ")) and len(line) > 140
    ]
    assert long_declarations == []


def test_secui_cli_vba_assigns_dictionary_objects_with_set():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert "Set serviceFanoutIndex(sourceDestinationKey) = services" in src
    assert "Set cliGroups(groupKey) = group" in src
    assert not re.search(r"^\s*serviceFanoutIndex\(sourceDestinationKey\)\s*=\s*services\s*$", src, re.MULTILINE)
    assert not re.search(r"^\s*cliGroups\(groupKey\)\s*=\s*group\s*$", src, re.MULTILINE)


def test_secui_cli_vba_defines_any_policy_value_helper():
    src = VBA_POLICY.read_text(encoding="utf-8")
    assert "Private Function IsAnyPolicyValue(" in src
    for token in ('"ANY"', '"ALL"', '"*"', '"0.0.0.0/0"'):
        assert token in src


def test_auto_run_output_sheets_are_not_protected():
    _build_workbook()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True)
    try:
        macro_written = (
            "requests",
            "processing_log",
            "route_results",
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
