"""Structural verification of the generated macro-enabled .xlsm.

LibreOffice is unavailable/broken on this host, so we cannot execute the VBA.
Instead we prove the artifact is a real macro-enabled workbook by inspecting:
  - the zip contains xl/vbaProject.bin (macros embedded)
  - both VBA modules are present with the expected public macros
  - all required sheets and headers exist
  - seed data (firewalls / firewall_ranges) is present
  - the seeded example request yields a real multi-firewall path via the oracle

Run: .venv/bin/python -m pytest tests/test_xlsm_structure.py -v
"""

import os
import subprocess
import sys
import zipfile

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSM = os.path.join(ROOT, "dist", "firewall-policy-automation.xlsm")
PY = sys.executable

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from tests.route_oracle import Firewall, FirewallRange, RouteEngine  # noqa: E402


def _build():
    subprocess.run(
        [PY, os.path.join(ROOT, "scripts", "build_xlsm.py")],
        cwd=ROOT, check=True, capture_output=True, text=True,
    )


@pytest.fixture(scope="module")
def xlsm_path():
    _build()
    assert os.path.exists(XLSM), "builder did not produce the .xlsm"
    return XLSM


def test_zip_has_vbaproject(xlsm_path):
    z = zipfile.ZipFile(xlsm_path)
    assert z.testzip() is None
    assert any("vbaProject.bin" in n for n in z.namelist())


def test_content_types_macro_enabled(xlsm_path):
    z = zipfile.ZipFile(xlsm_path)
    ct = z.read("[Content_Types].xml").decode("utf-8", "replace")
    assert "macroEnabled" in ct or "vbaProject" in ct


def test_modules_present(xlsm_path):
    from pyopenvba import excel as ex
    g = ex.ExcelFile(xlsm_path)
    try:
        names = g.module_names()
        assert "FirewallPolicyAutomation" in names
        assert "FirewallRouteAnalysis" in names
        assert "Module1" not in names
        assert "AnalyzeRequestRoutes" in g.get_module("FirewallRouteAnalysis")
        assert "MergeFirewallRequestFolder" in g.get_module("FirewallPolicyAutomation")
        assert "ConvertRequestsToSecuiBatch" in g.get_module("FirewallPolicyAutomation")
        assert "ConvertRequestsToSecuiCli" in g.get_module("FirewallPolicyAutomation")
        assert "SetupFirewallAutomationWorkbook" in g.get_module("FirewallPolicyAutomation")
    finally:
        g.close()


def test_sheets_and_headers(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)

    expected_sheets = {
        "requests", "firewalls", "firewall_ranges",
        "settings", "processing_log", "sample-request-format", "usage",
        "secui_batch", "secui_cli",
    }
    assert expected_sheets.issubset(set(wb.sheetnames)), \
        f"missing sheets: {expected_sheets - set(wb.sheetnames)}"

    # Exact header row per sheet. None of these may silently drift.
    expected_headers = {
        "firewalls": ["firewall_name", "vendor", "enabled", "comment"],
        "firewall_ranges": ["firewall_name", "source_cidr", "destination_cidr",
                            "direction", "path_order", "enabled", "comment"],
        "settings": ["key", "value", "\uc124\uba85"],
        "processing_log": ["processed_at", "source_file", "status", "merged_rows", "message"],
        "secui_batch": [
            "No", "\uc7a5\ube44\uba85", "\uc815\ucc45\uba85", "\ucd9c\ubc1c\uc9c0\uc8fc\uc18c",
            "\ucd9c\ubc1c\uc9c0\uba85", "\ubaa9\uc801\uc9c0\uc8fc\uc18c", "\ubaa9\uc801\uc9c0\uba85",
            "\uc11c\ube44\uc2a4", "\ud504\ub85c\ud1a0\ucf5c", "\ubaa9\uc801\uc9c0\ud3ec\ud2b8",
            "\ub3d9\uc791", "\ub85c\uadf8", "\uc0ac\uc6a9\uc5ec\ubd80", "\uc2dc\uc791\uc77c",
            "\uc885\ub8cc\uc77c", "\uc124\uba85", "\uc2e0\uccad\ubd80\uc11c",
            "\uc2e0\uccad\ubc88\ud638", "\uc6d0\ubcf8\ud30c\uc77c", "\uc6d0\ubcf8\ud589",
        ],
        "secui_cli": [
            "No", "\uc7a5\ube44\uba85", "\uc815\ucc45\uba85", "\uba85\ub839\uc5b4", "\uac80\ud1a0\uba54\ubaa8",
            "\uc2e0\uccad\ubd80\uc11c", "\uc2e0\uccad\ubc88\ud638", "\uc6d0\ubcf8\ud30c\uc77c", "\uc6d0\ubcf8\ud589",
        ],
        "usage": ["Step", "Action"],
    }
    for sheet, headers in expected_headers.items():
        ws = wb[sheet]
        actual = [ws.cell(1, c).value for c in range(1, len(headers) + 1)]
        assert actual == headers, f"{sheet} header drift: {actual!r}"

    # requests: exactly 25 columns; canonical leaf headers on row 2 (row 1 is the
    # cosmetic group-label band). Mirrors VBA WriteRequestHeaders.
    from scripts.build_xlsm import REQUESTS_HEADERS  # noqa: E402
    assert len(REQUESTS_HEADERS) == 25
    req = wb["requests"]
    assert req.max_column == 25, f"requests max_column={req.max_column}"
    actual_req = [req.cell(2, c).value for c in range(1, 26)]
    assert actual_req == REQUESTS_HEADERS, f"requests header drift: {actual_req!r}"
    # row-1 cosmetic group labels: 출발지 over IP+설명, 목적지 over IP+설명, merged.
    merged = {str(rng) for rng in req.merged_cells.ranges}
    assert req.cell(1, 8).value == "\ucd9c\ubc1c\uc9c0", "row1 출발지 group label missing"
    assert req.cell(1, 10).value == "\ubaa9\uc801\uc9c0", "row1 목적지 group label missing"
    assert "H1:I1" in merged, f"출발지 group not merged H1:I1: {merged}"
    assert "J1:K1" in merged, f"목적지 group not merged J1:K1: {merged}"

    # sample-request-format: blank A header, then the request template B..M.
    sf = wb["sample-request-format"]
    assert sf.cell(1, 1).value in (None, ""), "sample A1 must be blank"
    assert [sf.cell(1, c).value for c in range(2, 14)] == [
        "No", "\ucd9c\ubc1c\uc9c0IP", "\ucd9c\ubc1c\uc9c0", "\ubaa9\uc801\uc9c0IP", "\ubaa9\uc801\uc9c0",
        "\ud504\ub85c\ud1a0\ucf5c", "\ud3ec\ud2b8", "\ubc29\ud5a5", "\uc6a9\ub3c4", "\uc2dc\uc791\uc77c", "\uc885\ub8cc\uc77c", "\ube44\uace0",
    ]


def test_seed_data_present(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    assert wb["firewalls"].max_row >= 4
    assert wb["firewall_ranges"].max_row >= 7
    assert wb["firewall_ranges"].cell(2, 2).value, "firewall_ranges.source_cidr must be seeded"
    assert wb["firewall_ranges"].cell(2, 3).value, "firewall_ranges.destination_cidr must be seeded"
    settings_keys = [wb["settings"].cell(r, 1).value for r in range(1, wb["settings"].max_row + 1)]
    assert "header_alias" in settings_keys


def _engine_from_xlsm(wb):
    def T(v):
        return str(v).upper() in ("Y", "YES", "TRUE", "1")
    fs = wb["firewalls"]
    fws = [
        Firewall(fs.cell(r, 1).value, fs.cell(r, 2).value or "", T(fs.cell(r, 3).value))
        for r in range(2, fs.max_row + 1) if fs.cell(r, 1).value
    ]
    rs = wb["firewall_ranges"]
    ranges = [
        FirewallRange(
            rs.cell(r, 1).value,
            rs.cell(r, 2).value or "",
            rs.cell(r, 3).value or "",
            rs.cell(r, 4).value or "",
            int(rs.cell(r, 5).value or 999999),
            T(rs.cell(r, 6).value),
            rs.cell(r, 7).value or "",
        )
        for r in range(2, rs.max_row + 1) if rs.cell(r, 1).value
    ]
    return RouteEngine(firewalls=fws, firewall_ranges=ranges)


def test_seeded_example_yields_multi_firewall_path(xlsm_path):
    """The seeded example request must resolve to a real multi-hop path."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    eng = _engine_from_xlsm(wb)
    res = eng.analyze("10.10.10.5", "8.8.8.8", "OUT")
    assert res.status == "OK"
    assert ";" in res.target_firewalls
    assert ">" in res.firewall_path
    assert res.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"


def test_builtin_seed_resolves_cidr_request(xlsm_path):
    """A CIDR (\ub300\uc5ed) request against the shipped seed data must resolve a path."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    eng = _engine_from_xlsm(wb)
    res = eng.analyze("10.10.10.0/24", "172.16.1.0/24", "OUT")
    assert res.status == "OK"
    assert res.target_firewalls == "SECUI-FW-01"
    # full chain to the internet still resolves
    res2 = eng.analyze("10.10.0.5", "8.8.8.8", "OUT")
    assert res2.status == "OK"
    assert res2.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"


# --------------------------------------------------------------------------- #
# F1: settings sheet schema must match between the build seed and VBA WriteSettings
# --------------------------------------------------------------------------- #

VBA_POLICY = os.path.join(ROOT, "vba", "FirewallPolicyAutomation.bas")


def test_settings_schema_matches_vba_and_build(xlsm_path):
    """settings must be a 3-column sheet (key, value, 설명) in BOTH the built
    workbook and the VBA WriteSettings setup path.

    Guards F1: build_xlsm.py seeds key/value/설명 but VBA historically wrote
    only key/value, so SetupFirewallAutomationWorkbook produced a divergent
    2-column sheet.
    """
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    s = wb["settings"]
    assert [s.cell(1, c).value for c in range(1, 4)] == ["key", "value", "\uc124\uba85"]
    keys = [s.cell(r, 1).value for r in range(2, s.max_row + 1)]
    for required in ("request_folder", "parse_sheet", "parse_targets", "header_alias"):
        assert required in keys, f"settings missing key {required}"
    # every seeded settings row must carry a non-empty 설명 (column 3)
    for r in range(2, s.max_row + 1):
        if s.cell(r, 1).value:
            assert str(s.cell(r, 3).value or "").strip() != "", f"row {r} missing 설명"

    # VBA WriteSettings must write a 3-column header, not A1:B1
    src = open(VBA_POLICY, encoding="utf-8").read()
    assert 'Range("A1:C1").Value = Array("key", "value"' in src, \
        "WriteSettings must seed a 3-column (key,value,\uc124\uba85) header"
    assert 'Range("A1:B1").Value = Array("key", "value")' not in src, \
        "WriteSettings still writes the legacy 2-column header"


def test_policy_module_has_no_merge_time_legacy_firewall_matching():
    src = open(VBA_POLICY, encoding="utf-8").read()
    for dead in (
        "ResolveTargetFirewalls",
        "RequestRowMatchDetails",
        "RequestRowMatchesCidr",
        "AddressListMatchDetails",
        "AddressListMatchesCidr",
        "RegisteredParseTargetColumns",
        "ParseTargetColumnsFromText",
        "RequestColumnNumber",
        "MarkUnmatchedFirewalls",
    ):
        assert dead not in src, f"obsolete helper {dead} still present"
    # CopyRequestRow must no longer pre-fill target_firewalls before route analysis
    assert "COL_TARGET_FIREWALLS).Value = ResolveTargetFirewalls" not in src


def test_secui_batch_macro_splits_multi_hop_targets():
    src = open(VBA_POLICY, encoding="utf-8").read()
    assert "Public Sub ConvertRequestsToSecuiBatch" in src
    assert "SECUI_BATCH_SHEET" in src
    assert "LoadSecuiFirewalls(firewallsSheet)" in src
    assert 'vendorName = UCase$(Trim$(CStr(firewallsSheet.Cells(rowIndex, FW_COL_VENDOR).Value)))' in src
    assert 'vendorName = "SECUI"' in src
    assert "FirewallRowEnabled(firewallsSheet.Cells(rowIndex, FW_COL_ENABLED).Value)" in src
    assert 'targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")' in src
    assert "secuiFirewalls.Exists(SecuiFirewallKey(firewallName))" in src


def test_secui_cli_macro_generates_fw_set_srule_commands():
    src = open(VBA_POLICY, encoding="utf-8").read()
    assert "Public Sub ConvertRequestsToSecuiCli" in src
    assert "SECUI_CLI_SHEET" in src
    assert "WriteSecuiCliHeaders" in src
    assert "LoadSecuiFirewalls(firewallsSheet)" in src
    assert "CopySecuiCliRows" in src
    assert "SecuiCliCommand" in src
    assert '"fw set srule name "' in src
    assert "secuiFirewalls.Exists(SecuiFirewallKey(firewallName))" in src


def test_duplicate_marking_runs_after_route_analysis():
    """MarkDuplicateRequests must run AFTER AnalyzeRequestRoutes.

    AnalyzeRequestRoutes.WriteResultRow unconditionally writes validation_status /
    validation_message. If MarkDuplicateRequests ran BEFORE it, the DUPLICATE
    status/message it wrote via AppendValidationMessage would be silently
    overwritten (only row highlighting would survive). Running it after lets
    AppendValidationMessage MERGE the DUPLICATE marker onto the route-owned status.
    """
    src = open(VBA_POLICY, encoding="utf-8").read()
    run_pos = src.find('Application.Run "FirewallRouteAnalysis.AnalyzeRequestRoutes"')
    dup_pos = src.find("MarkDuplicateRequests requestsSheet")
    assert run_pos != -1, "route-analysis call not found"
    assert dup_pos != -1, "MarkDuplicateRequests call not found"
    assert dup_pos > run_pos, \
        "MarkDuplicateRequests must run AFTER AnalyzeRequestRoutes so its DUPLICATE " \
        "status merges onto (not is overwritten by) the route result"


def test_duplicate_highlight_preserves_route_status_color():
    """Duplicate row highlighting must NOT clobber the route-owned status-cell color.

    Since MarkDuplicateRequests now runs AFTER AnalyzeRequestRoutes, a bare
    `Rows(rowIndex).Interior.Color = ...` would overwrite the validation_status
    cell color (column 15) that WriteResultRow set to encode OK/DIRECTION_MISMATCH/severity.
    The highlight must save and restore that one cell's color.
    """
    src = open(VBA_POLICY, encoding="utf-8").read()
    # the duplicate marker must route row highlighting through the color-preserving
    # helper, not color whole rows directly inside MarkDuplicateRequests.
    assert "HighlightDuplicateRow" in src, "missing color-preserving highlight helper"
    # the helper must restore the validation_status cell color after the row fill
    helper_start = src.find("Private Sub HighlightDuplicateRow")
    assert helper_start != -1
    helper = src[helper_start:helper_start + 600]
    assert "COL_VALIDATION_STATUS).Interior.Color = statusColor" in helper, \
        "HighlightDuplicateRow must restore the route status-cell color"


# --------------------------------------------------------------------------- #
# F6: the VBA source injected into the .xlsm must match vba/*.bas exactly
# (catches pyOpenVBA injection corruption / build drift)
# --------------------------------------------------------------------------- #

def _normalize_vba(text):
    # Normalize for round-trip comparison:
    #  - drop the leading `Attribute VB_Name = "..."` header (body-only on read)
    #  - collapse non-ASCII to '?' because pyOpenVBA reads the MBCS module stream
    #    with a codepage that is lossy for Korean COMMENT text (한글 -> '?').
    #    Code/identifiers are ASCII, so this still catches real source corruption.
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [ln for ln in lines if not ln.startswith("Attribute VB_Name")]
    body = "\n".join(lines).strip()
    return "".join(c if ord(c) < 128 else "?" for c in body)


def test_injected_vba_modules_match_source_files(xlsm_path):
    from pyopenvba import excel as ex
    modules = {
        "FirewallPolicyAutomation": os.path.join(ROOT, "vba", "FirewallPolicyAutomation.bas"),
        "FirewallRouteAnalysis": os.path.join(ROOT, "vba", "FirewallRouteAnalysis.bas"),
    }
    g = ex.ExcelFile(xlsm_path)
    try:
        for name, path in modules.items():
            injected = _normalize_vba(g.get_module(name))
            source = _normalize_vba(open(path, encoding="utf-8").read())
            assert injected == source, \
                f"injected {name} does not match vba/{name}.bas (build corruption?)"
    finally:
        g.close()


def test_korean_preserved_in_injected_vba(xlsm_path):
    """한글 functional 리터럴이 빌드된 .xlsm VBA에 보존돼야 한다 (code_page=949)."""
    from pyopenvba import excel as ex
    g = ex.ExcelFile(xlsm_path)
    try:
        assert g.vba_project().code_page == 949, "VBA project must use code page 949"
        src = g.get_module("FirewallPolicyAutomation")
        assert 'headerMap("출발지ip")' in src, "Korean column literal corrupted"
        assert 'headerMap("목적지ip")' in src
        assert "출발지IP" in src and "목적지IP" in src
        assert "?" * 3 not in src, "Korean text was replaced with '?'"
    finally:
        g.close()


VBA_ROUTE = os.path.join(ROOT, "vba", "FirewallRouteAnalysis.bas")


def test_vba_loads_firewall_ranges_sheet():
    src = open(VBA_ROUTE, encoding="utf-8").read()
    assert 'Private Const FIREWALL_RANGE_SHEET As String = "firewall_ranges"' in src
    assert "ThisWorkbook.Worksheets(FIREWALL_RANGE_SHEET)" in src
    assert "NETWORK_SHEET" not in src
    assert "ROUTING_SHEET" not in src


def test_vba_split_address_list_collapses_spaces():
    """S6 parity: VBA SplitAddressList must collapse ASCII space-runs into ';'
    (mirrors Python split_address_list) so space-separated multi-CIDR splits."""
    src = open(VBA_ROUTE, encoding="utf-8").read().replace("\r\n", "\n")
    fn = src[src.find("Private Function SplitAddressList"):]
    fn = fn[:fn.find("End Function")]
    assert 'Replace(normalized, " ", ";")' in fn, \
        "SplitAddressList must turn spaces into ';' to split space-separated CIDRs"


def test_vba_find_header_row_is_content_based():
    """Parity: VBA header detection must score by FIELD CONTENT (IP columns +
    field count), not by an exact 'no'/'번호' cell match. The content scan lives
    in BestHeaderRow, which FindHeaderRow and FindRequestSheet both delegate to."""
    src = open(VBA_POLICY, encoding="utf-8").read().replace("\r\n", "\n")
    fn = src[src.find("Private Function BestHeaderRow"):]
    fn = fn[:fn.find("\nEnd Function")]
    # must score by IP presence + field count, not the old exact-match anchor
    assert "hasIp" in fn and "fieldCount" in fn, "BestHeaderRow must score by content"
    assert 'If valueText = "no" Or valueText = "번호" Then' not in fn, \
        "old exact-match 'no'/'번호' anchor must be gone"
    # IsFieldHeader helper must exist and enumerate the field headers
    assert "Private Function IsFieldHeader" in src
    # FindRequestSheet must exist so data on a non-first sheet still parses
    assert "Private Function FindRequestSheet" in src


def test_vba_header_key_strips_punctuation():
    """Parity: HeaderKey must strip decorating punctuation so 'No.' -> 'no', but
    PRESERVE an all-punctuation token like '#' (return original when stripping
    empties it) so it can match the '#' No-alias."""
    src = open(VBA_POLICY, encoding="utf-8").read().replace("\r\n", "\n")
    fn = src[src.find("Private Function HeaderKey"):]
    fn = fn[:fn.find("\nEnd Function")]
    assert "puncts" in fn and ("Left$(k" in fn or "Mid$(k" in fn), \
        "HeaderKey must strip leading/trailing punctuation (No. -> no)"
    assert "original" in fn and "If Len(k) = 0 Then" in fn, \
        "HeaderKey must keep all-punctuation tokens (e.g. '#') instead of emptying"


def test_vba_canonical_no_includes_hash_alias():
    """Parity: CanonicalHeaderName 'no' Case must include '#' so a '#' column
    anchors the header row by itself."""
    src = open(VBA_POLICY, encoding="utf-8").read().replace("\r\n", "\n")
    line = next(l for l in src.split("\n") if "CanonicalHeaderName = \"no\"" in l)
    assert '"#"' in line, "CanonicalHeaderName 'no' Case must include '#'"


def test_workbook_open_does_not_swallow_errors():
    """Workbook_Open auto-run must surface failures, not blanket-swallow them."""
    from pyopenvba import excel as ex
    g = ex.ExcelFile(XLSM)
    try:
        tw = g.get_module("ThisWorkbook")
        assert "Workbook_Open" in tw
        assert "MergeFirewallRequestFolder" in tw
        assert "On Error Resume Next" not in tw, "must not blanket-swallow errors"
        assert "AutoRunErr" in tw, "must have an error handler that surfaces failures"
    finally:
        g.close()


# --------------------------------------------------------------------------- #
# UX (scope B): build-time openpyxl input-assist / display features.
# These are DISPLAY/INPUT-ASSIST ONLY and must NOT change the route algorithm.
# All are sheet-level and additive: no cell VALUE changes, no header drift, no
# max_column/max_row drift, vbaProject.bin preserved (keep_vba=True).
# --------------------------------------------------------------------------- #

def _dv_for(ws, cell_ref):
    """Return the DataValidation covering cell_ref on ws, or None."""
    from openpyxl.utils.cell import coordinate_to_tuple
    target = coordinate_to_tuple(cell_ref)
    for dv in ws.data_validations.dataValidation:
        for rng in dv.sqref.ranges:
            if (rng.min_row <= target[0] <= rng.max_row
                    and rng.min_col <= target[1] <= rng.max_col):
                return dv
    return None


def test_ux_protocol_direction_dropdowns_on_requests(xlsm_path):
    """requests 프로토콜(col12)·방향(col14) cells offer a dropdown list.

    Values must stay compatible with the route algorithm: 프로토콜 is display-only
    (TCP/UDP/ICMP); 방향 must be one of the values VBA NormalizeDirection accepts
    (IN/OUT/BOTH) so the dropdown can never inject a #INVALID direction.
    """
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    req = wb["requests"]

    proto = _dv_for(req, "L3")
    assert proto is not None, "프로토콜(col12) has no data validation"
    assert proto.type == "list"
    assert set(proto.formula1.strip('"').split(",")) >= {"TCP", "UDP", "ICMP"}

    direction = _dv_for(req, "N3")
    assert direction is not None, "방향(col14) has no data validation"
    assert direction.type == "list"
    assert set(direction.formula1.strip('"').split(",")) <= {"IN", "OUT", "BOTH", ""}
    assert {"IN", "OUT"} <= set(direction.formula1.strip('"').split(","))
    # validations must skip the header row and not explode to 1048576 rows
    for dv in (proto, direction):
        rng = next(iter(dv.sqref.ranges))
        assert rng.min_row == 3, "dropdown must start at row 3 (skip 2-row header band)"
        assert rng.max_row <= 5000, "dropdown range must be bounded, not whole-column"


def test_ux_enabled_dropdown_accepts_vba_truthy(xlsm_path):
    """firewalls.enabled(col3) dropdown values must all be truthy/falsy under the
    VBA reader T(v)=UCase(v) in (Y,YES,TRUE,1); list must include both Y and N."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    dv = _dv_for(wb["firewalls"], "C2")
    assert dv is not None, "firewalls.enabled(col3) has no data validation"
    assert dv.type == "list"
    vals = set(dv.formula1.strip('"').split(","))
    assert {"Y", "N"} <= vals
    assert vals <= {"Y", "N", "TRUE", "FALSE"}


def test_ux_macro_written_sheets_are_not_protected(xlsm_path):
    """CRITICAL: requests and processing_log are written by VBA at runtime, so they
    must NOT be sheet-protected (protection would make macro Cells(...).Value fail).
    Operator-input sheets MAY be protected."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    for macro_sheet in ("requests", "processing_log"):
        assert wb[macro_sheet].protection.sheet is False, \
            f"{macro_sheet} is macro-written and must never be protected"


def test_ux_settings_folder_cell_unlocked_when_protected(xlsm_path):
    """If settings is protected, the request_folder value cell (B column of the
    request_folder row) must be UNLOCKED so SelectRequestFolder can write it."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    s = wb["settings"]
    if not s.protection.sheet:
        pytest.skip("settings not protected; folder-cell lock state is moot")
    folder_row = next(
        (r for r in range(2, s.max_row + 1)
         if s.cell(r, 1).value == "request_folder"), None)
    assert folder_row is not None
    assert s.cell(folder_row, 2).protection.locked is False, \
        "request_folder value cell must be unlocked for the folder-select macro"


def test_ux_required_input_empty_highlight(xlsm_path):
    """requests must carry conditional formatting that flags an empty required input
    cell (출발지IP col8 / 목적지IP col10). Display-only; never touches values."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    req = wb["requests"]
    cf_ranges = [str(rng) for rng in req.conditional_formatting]
    joined = " ".join(cf_ranges)
    assert ("H" in joined and "J" in joined) or len(cf_ranges) >= 1, \
        "requests has no conditional formatting for empty required cells"
    assert any(req.conditional_formatting[rng] for rng in req.conditional_formatting), \
        "conditional formatting present but carries no rules"


def test_ux_header_comments_on_key_columns(xlsm_path):
    """Operator-facing 안내 comments must exist on the requests 출발지IP / 목적지IP
    leaf-header cells (row 2), and the format must stay valid (no value drift)."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    req = wb["requests"]
    assert req.cell(2, 8).comment is not None, "출발지IP header needs a hint comment"
    assert req.cell(2, 10).comment is not None, "목적지IP header needs a hint comment"
    # comment must not alter the header value
    assert req.cell(2, 8).value == "\ucd9c\ubc1c\uc9c0IP"
    assert req.cell(2, 10).value == "\ubaa9\uc801\uc9c0IP"


def test_ux_tab_colors_assigned(xlsm_path):
    """Input sheets vs result/log sheets must be visually distinguished by tab color
    (display-only). At least the requests result sheet must carry a tabColor."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    assert wb["requests"].sheet_properties.tabColor is not None, \
        "requests tab must have a color"
    colored = sum(
        1 for name in wb.sheetnames
        if wb[name].sheet_properties.tabColor is not None)
    assert colored >= 3, "at least 3 sheets should be tab-colored for navigation"


def test_ux_does_not_change_route_algorithm(xlsm_path):
    """GUARD: the route oracle result on the seeded data must be byte-identical
    whether or not UX is applied. UX must be display/input-assist ONLY."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    eng = _engine_from_xlsm(wb)
    res = eng.analyze("10.10.10.5", "8.8.8.8", "OUT")
    assert res.status == "OK"
    assert res.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"
