import os
import subprocess
import sys
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSM = os.path.join(ROOT, "dist", "firewall-policy-automation.xlsm")
PY = sys.executable
VBA_POLICY = os.path.join(ROOT, "vba", "FirewallPolicyAutomation.bas")


VISIBLE_POLICY_ANALYSIS_HEADERS = [
    "판정",
    "요청번호",
    "대상방화벽",
    "출발지",
    "목적지",
    "서비스",
    "기존정책",
    "기존정책상태",
    "근거",
    "조치",
]

HIDDEN_POLICY_ANALYSIS_HEADERS = [
    "요청원본행",
    "정책원본행",
    "raw_source",
    "raw_destination",
    "raw_service",
    "normalized_source",
    "normalized_destination",
    "normalized_protocol",
    "normalized_port",
    "debug_note",
]

SECUI_EXPORT_HEADERS = [
    "policy_id",
    "policy_name",
    "firewall_name",
    "source",
    "destination",
    "service",
    "action",
    "enabled",
    "comment",
]


def _build() -> None:
    subprocess.run(
        [PY, os.path.join(ROOT, "scripts", "build_xlsm.py")],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )


def test_secui_export_and_policy_analysis_sheets_are_built_for_operator_use() -> None:
    _build()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True)

    assert "secui_policy_export" in wb.sheetnames
    assert "policy_analysis" in wb.sheetnames

    export = wb["secui_policy_export"]
    assert [export.cell(1, c).value for c in range(1, len(SECUI_EXPORT_HEADERS) + 1)] == SECUI_EXPORT_HEADERS
    assert export.cell(2, 2).value == "ALLOW_WEB_TO_DMZ"

    analysis = wb["policy_analysis"]
    visible_count = len(VISIBLE_POLICY_ANALYSIS_HEADERS)
    hidden_count = len(HIDDEN_POLICY_ANALYSIS_HEADERS)
    assert [analysis.cell(1, c).value for c in range(1, visible_count + 1)] == VISIBLE_POLICY_ANALYSIS_HEADERS
    assert [
        analysis.cell(1, c).value
        for c in range(visible_count + 1, visible_count + hidden_count + 1)
    ] == HIDDEN_POLICY_ANALYSIS_HEADERS

    for col in range(1, visible_count + 1):
        letter = get_column_letter(col)
        assert analysis.column_dimensions[letter].hidden is False

    for col in range(visible_count + 1, visible_count + hidden_count + 1):
        letter = get_column_letter(col)
        assert analysis.column_dimensions[letter].hidden is True

    statuses = [analysis.cell(row, 1).value for row in range(2, 5)]
    assert statuses == ["EXISTING_ALLOW", "EXISTING_DENY", "NO_EXISTING_POLICY"]
    assert analysis.auto_filter.ref == "A1:T4"
    assert analysis.freeze_panes == "A2"
    assert analysis.column_dimensions["A"].width >= 16
    assert analysis.column_dimensions["J"].width >= 24


def test_secui_policy_analysis_preserves_existing_workbook_contracts() -> None:
    _build()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True)

    req = wb["requests"]
    assert req.max_column == 25
    assert req.cell(2, 7).value == "대상방화벽"

    ranges = wb["firewall_ranges"]
    assert [ranges.cell(1, c).value for c in range(1, 8)] == [
        "firewall_name",
        "source_cidr",
        "destination_cidr",
        "direction",
        "path_order",
        "enabled",
        "comment",
    ]

    with zipfile.ZipFile(XLSM) as archive:
        assert any(name.endswith("vbaProject.bin") for name in archive.namelist())


def test_vba_has_read_only_secui_policy_analysis_macro() -> None:
    src = open(VBA_POLICY, encoding="utf-8").read()

    assert 'Private Const SECUI_POLICY_EXPORT_SHEET As String = "secui_policy_export"' in src
    assert 'Private Const POLICY_ANALYSIS_SHEET As String = "policy_analysis"' in src
    assert "Public Sub AnalyzeSecuiPolicyExport" in src
    assert "WriteSecuiPolicyExportHeaders" in src
    assert "WritePolicyAnalysisHeaders" in src
    assert "CopySecuiCliRows" in src
    assert "AnalyzeSecuiPolicyExport" in src
    assert "Application.Run \"FirewallRouteAnalysis.AnalyzeRequestRoutes\"" in src
