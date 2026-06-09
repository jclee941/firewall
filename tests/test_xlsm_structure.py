"""Structural verification of the generated macro-enabled .xlsm.

LibreOffice is unavailable/broken on this host, so we cannot execute the VBA.
Instead we prove the artifact is a real macro-enabled workbook by inspecting:
  - the zip contains xl/vbaProject.bin (macros embedded)
  - both VBA modules are present with the expected public macros
  - all required sheets and headers exist
  - seed data (firewalls / network_definitions / routing_paths) is present
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
from route_oracle import Firewall, Network, RouteEngine, RoutingPath  # noqa: E402


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
        assert "SetupFirewallAutomationWorkbook" in g.get_module("FirewallPolicyAutomation")
    finally:
        g.close()


def test_sheets_and_headers(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)

    # All 8 sheets must exist (matches CI/release verification, which the
    # previous version of this test did not — it omitted sample-request-format).
    expected_sheets = {
        "requests", "firewalls", "network_definitions", "routing_paths",
        "settings", "processing_log", "sample-request-format", "usage",
    }
    assert expected_sheets.issubset(set(wb.sheetnames)), \
        f"missing sheets: {expected_sheets - set(wb.sheetnames)}"

    # Exact header row per sheet. None of these may silently drift.
    expected_headers = {
        "firewalls": ["firewall_name", "vendor", "enabled", "inside_cidr", "outside_cidr", "comment"],
        "network_definitions": ["network_name", "network_cidr", "zone", "site", "enabled"],
        "routing_paths": ["firewall_name", "src_zone", "dst_zone",
                          "ingress_if", "egress_if", "path_order", "enabled"],
        "settings": ["key", "value", "\uc124\uba85"],
        "processing_log": ["processed_at", "source_file", "status", "merged_rows", "message"],
        "usage": ["Step", "Action"],
    }
    for sheet, headers in expected_headers.items():
        ws = wb[sheet]
        actual = [ws.cell(1, c).value for c in range(1, len(headers) + 1)]
        assert actual == headers, f"{sheet} header drift: {actual!r}"

    # requests: exactly 24 columns matching the build seed constant.
    from build_xlsm import REQUESTS_HEADERS  # noqa: E402
    assert len(REQUESTS_HEADERS) == 24
    req = wb["requests"]
    assert req.max_column == 24, f"requests max_column={req.max_column}"
    actual_req = [req.cell(1, c).value for c in range(1, 25)]
    assert actual_req == REQUESTS_HEADERS, f"requests header drift: {actual_req!r}"

    # sample-request-format: blank A header, then the request template B..M.
    sf = wb["sample-request-format"]
    assert sf.cell(1, 1).value in (None, ""), "sample A1 must be blank"
    assert [sf.cell(1, c).value for c in range(2, 14)] == [
        "No", "\ucd9c\ubc1c\uc9c0IP", "\ucd9c\ubc1c\uc9c0", "\ubaa9\uc801\uc9c0IP", "\ubaa9\uc801\uc9c0",
        "\ud504\ub85c\ud1a0\ucf5c", "\ud3ec\ud2b8", "\ubc29\ud5a5", "\uc6a9\ub3c4", "\uc2dc\uc791\uc77c", "\uc885\ub8cc\uc77c", "\ube44\uace0",
    ]


def test_seed_data_present(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    assert wb["firewalls"].max_row >= 4          # header + 3 firewalls
    assert wb["network_definitions"].max_row >= 6
    # firewalls.inside_cidr / outside_cidr seeded (auto-route mode); routing_paths
    # is header-only by default (auto-derive from inside/outside CIDRs).
    assert wb["firewalls"].cell(2, 4).value, "firewalls.inside_cidr must be seeded"
    assert wb["firewalls"].cell(2, 5).value, "firewalls.outside_cidr must be seeded"
    assert wb["routing_paths"].max_row >= 1
    # settings has the fallback toggle key
    settings_keys = [wb["settings"].cell(r, 1).value for r in range(1, wb["settings"].max_row + 1)]
    assert "route_legacy_fallback" in settings_keys


def _engine_from_xlsm(wb):
    """Build a RouteEngine from the seeded workbook (auto inside/outside mode)."""
    def T(v):
        return str(v).upper() in ("Y", "YES", "TRUE", "1")
    fs = wb["firewalls"]
    fws = [
        Firewall(fs.cell(r, 1).value, fs.cell(r, 2).value or "", T(fs.cell(r, 3).value),
                 inside_cidr=fs.cell(r, 4).value or "",
                 outside_cidr=fs.cell(r, 5).value or "")
        for r in range(2, fs.max_row + 1) if fs.cell(r, 1).value
    ]
    return RouteEngine(networks=[], firewalls=fws, routing_paths=[])


def test_seeded_example_yields_multi_firewall_path(xlsm_path):
    """The seeded example request must resolve to a real multi-hop path."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    eng = _engine_from_xlsm(wb)
    # 10.10.10.5 (FW-01 inside) -> 8.8.8.8 (FW-03 outside 0.0.0.0/0): 3 firewalls
    res = eng.analyze("10.10.10.5", "8.8.8.8", "OUT")
    assert res.status == "OK"
    assert ";" in res.target_firewalls
    assert ">" in res.firewall_path
    assert res.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"


def test_builtin_seed_resolves_cidr_request(xlsm_path):
    """A CIDR (\ub300\uc5ed) request against the shipped seed data must resolve a path."""
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    eng = _engine_from_xlsm(wb)
    # CIDR request inside FW-01.inside -> inside FW-01.outside: one firewall
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
    for required in ("request_folder", "parse_targets", "route_legacy_fallback", "header_alias"):
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


# --------------------------------------------------------------------------- #
# F7: the dead merge-time legacy CIDR matcher (and its now-orphaned helpers) must
# be gone. AnalyzeRequestRoutes solely owns target_firewalls / validation_* and the
# empty-target semantics for INTRA_ZONE / NO_PATH / ZONE_UNRESOLVED, so the old
# MarkUnmatchedFirewalls post-processor (which overwrote blanks with UNMATCHED) is
# --------------------------------------------------------------------------- #

def test_policy_module_has_no_merge_time_legacy_firewall_matching():
    src = open(VBA_POLICY, encoding="utf-8").read()
    # legacy CIDR-overlap matcher, its private helpers, the parse-target column
    # plumbing that only the matcher used, and the obsolete unmatched-marker that
    # corrupted route-owned output — all must be removed.
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
    cell color (column 15) that WriteResultRow set to encode OK/MULTI_PATH/severity.
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


def test_auto_zone_ordinal_counts_all_rows():
    """firewalls ordinal must increment for every row BEFORE the IsEnabled check,
    so disabled firewalls consume an ordinal (Python parity). Guards the single-
    line-If trap that previously ran fwOrdinal/fwZones unconditionally."""
    src = open(VBA_ROUTE, encoding="utf-8").read()
    norm = src.replace("\r\n", "\n")
    inc = norm.find("fwOrdinal = fwOrdinal + 1")
    gate = norm.find("If fwEnabled Then\n")
    inside = norm.find('fwInside(fwName) = Trim$(CStr(ws.Cells(i, 4).Value))')
    assert inc != -1, "ordinal increment line missing"
    assert gate != -1, "enabled gate must be a block 'If fwEnabled Then' on its own line"
    assert inside != -1, "fwInside assignment missing"
    assert inc < gate, "fwOrdinal must increment before the enabled gate"
    assert inside > gate, "fwInside must be assigned inside the enabled block"
    # no single-line If trap (a statement after Then)
    assert "If IsEnabled(ws.Cells(i, 3).Value) Then mEnabledFw" not in norm


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
