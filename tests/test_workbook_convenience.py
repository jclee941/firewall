from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSM = ROOT / "dist" / "firewall-policy-automation.xlsm"


def _build() -> None:
    subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "build_xlsm.py")],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )


def _usage_links(ws) -> set[str]:
    links: set[str] = set()
    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row=row, column=3)
        if cell.hyperlink is None or cell.hyperlink.target is None:
            continue
        links.add(str(cell.value))
    return links


def test_policy_summary_guides_existing_policy_review() -> None:
    _build()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True, data_only=False)
    try:
        assert "policy_summary" in wb.sheetnames

        summary = wb["policy_summary"]
        headers = [summary.cell(1, column).value for column in range(1, 5)]
        assert headers == ["구분", "건수", "검토 기준", "다음 조치"]
        assert summary.freeze_panes == "A2"
        assert summary.auto_filter.ref == "A1:D7"

        labels = [summary.cell(row, 1).value for row in range(2, 8)]
        assert labels == ["전체", "기존 허용", "기존 차단", "검토 필요", "비활성 일치", "기존 정책 없음"]
        formulas = [summary.cell(row, 2).value for row in range(2, 8)]
        assert formulas[0] == '=COUNTA(policy_analysis!A2:A5000)'
        assert '=COUNTIF(policy_analysis!A2:A5000,"EXISTING_ALLOW")' in formulas
        assert '=COUNTIF(policy_analysis!A2:A5000,"NO_EXISTING_POLICY")' in formulas

        assert "policy_summary" in _usage_links(wb["usage"])
    finally:
        wb.close()


def test_secui_export_headers_explain_required_paste_shape() -> None:
    _build()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True, data_only=False)
    try:
        export = wb["secui_policy_export"]
        for column in range(1, 10):
            comment = export.cell(1, column).comment
            assert comment is not None
            assert "붙여넣" in comment.text or "필수" in comment.text

        validation_ranges = [str(dv.sqref) for dv in export.data_validations.dataValidation]
        assert any("G2:G5000" in sqref for sqref in validation_ranges)
        assert any("H2:H5000" in sqref for sqref in validation_ranges)
    finally:
        wb.close()


def test_convenience_surfaces_keep_core_contracts() -> None:
    _build()
    wb = openpyxl.load_workbook(XLSM, keep_vba=True)
    try:
        assert wb["requests"].max_column == 25
        assert [wb["firewall_ranges"].cell(1, column).value for column in range(1, 8)] == [
            "firewall_name",
            "source_cidr",
            "destination_cidr",
            "direction",
            "path_order",
            "enabled",
            "comment",
        ]
    finally:
        wb.close()
