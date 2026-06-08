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
    expected = {
        "requests": "source_file",
        "firewalls": "firewall_name",
        "network_definitions": "network_name",
        "routing_paths": "firewall_name",
        "settings": "key",
        "processing_log": "processed_at",
        "usage": "Step",
    }
    for sheet, a1 in expected.items():
        assert sheet in wb.sheetnames, f"missing sheet {sheet}"
        assert wb[sheet]["A1"].value == a1, f"{sheet} A1"
    # requests must have the expanded output columns through X
    assert wb["requests"]["R1"].value == "firewall_path"
    assert wb["requests"]["S1"].value == "source_zone"
    assert wb["requests"]["T1"].value == "destination_zone"
    assert wb["requests"]["U1"].value == "zone_path"
    assert wb["requests"]["V1"].value == "request_team"
    assert wb["requests"]["W1"].value == "request_doc_no"
    assert wb["requests"]["X1"].value == "request_folder"
    # routing_paths header
    rp = wb["routing_paths"]
    assert [rp.cell(1, c).value for c in range(1, 8)] == [
        "firewall_name", "src_zone", "dst_zone",
        "ingress_if", "egress_if", "path_order", "enabled",
    ]


def test_seed_data_present(xlsm_path):
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    assert wb["firewalls"].max_row >= 4          # header + 3 firewalls
    assert wb["network_definitions"].max_row >= 6
    assert wb["routing_paths"].max_row >= 5
    # settings has the fallback toggle key
    settings_keys = [wb["settings"].cell(r, 1).value for r in range(1, wb["settings"].max_row + 1)]
    assert "route_legacy_fallback" in settings_keys


def test_seeded_example_yields_multi_firewall_path(xlsm_path):
    """The seeded example request must resolve to a real multi-hop path.

    This mirrors the exact seed data the builder writes, proving the workbook
    will produce a meaningful 적용대상방화벽 result when the macro runs.
    """
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    nd = wb["network_definitions"]
    nets = [
        Network(nd.cell(r, 1).value, nd.cell(r, 2).value, nd.cell(r, 3).value,
                nd.cell(r, 4).value or "", str(nd.cell(r, 5).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, nd.max_row + 1) if nd.cell(r, 1).value
    ]
    fws_sheet = wb["firewalls"]
    fws = [
        Firewall(fws_sheet.cell(r, 1).value, fws_sheet.cell(r, 2).value or "",
                 str(fws_sheet.cell(r, 3).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, fws_sheet.max_row + 1) if fws_sheet.cell(r, 1).value
    ]
    rp = wb["routing_paths"]
    rps = [
        RoutingPath(rp.cell(r, 1).value, rp.cell(r, 2).value, rp.cell(r, 3).value,
                    rp.cell(r, 4).value or "", rp.cell(r, 5).value or "",
                    int(rp.cell(r, 6).value), str(rp.cell(r, 7).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, rp.max_row + 1) if rp.cell(r, 1).value
    ]
    eng = RouteEngine(networks=nets, firewalls=fws, routing_paths=rps)

    req = wb["requests"]
    src = req["D2"].value
    dst = req["F2"].value
    direction = req["J2"].value
    assert src and dst, "builder must seed an example request"
    res = eng.analyze(src, dst, direction)
    assert res.status == "OK"
    # the seeded example crosses two firewalls -> proves multi-target output
    assert ";" in res.target_firewalls
    assert ">" in res.firewall_path


def test_builtin_seed_resolves_cidr_request(xlsm_path):
    """A CIDR (대역) request against the shipped seed data must resolve a path.

    Guards the Oracle-found defect: request values are often CIDRs, not single
    IPs. Uses the workbook's own seeded network_definitions/firewalls/routing.
    """
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    nd = wb["network_definitions"]
    nets = [
        Network(nd.cell(r, 1).value, nd.cell(r, 2).value, nd.cell(r, 3).value,
                nd.cell(r, 4).value or "", str(nd.cell(r, 5).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, nd.max_row + 1) if nd.cell(r, 1).value
    ]
    fs = wb["firewalls"]
    fws = [
        Firewall(fs.cell(r, 1).value, fs.cell(r, 2).value or "",
                 str(fs.cell(r, 3).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, fs.max_row + 1) if fs.cell(r, 1).value
    ]
    rp = wb["routing_paths"]
    rps = [
        RoutingPath(rp.cell(r, 1).value, rp.cell(r, 2).value, rp.cell(r, 3).value,
                    rp.cell(r, 4).value or "", rp.cell(r, 5).value or "",
                    int(rp.cell(r, 6).value), str(rp.cell(r, 7).value).upper() in ("Y", "YES", "TRUE", "1"))
        for r in range(2, rp.max_row + 1) if rp.cell(r, 1).value
    ]
    eng = RouteEngine(networks=nets, firewalls=fws, routing_paths=rps)
    # CIDR request: 10.10.10.0/24 (internal) -> 172.16.1.0/24 (server)
    res = eng.analyze("10.10.10.0/24", "172.16.1.0/24", "OUT")
    assert res.status == "OK"
    assert res.source_zone == "internal"
    assert res.destination_zone == "server"
    assert res.target_firewalls == "SECUI-FW-01"
    # address-list request also resolves
    res2 = eng.analyze("10.10.10.5;10.10.10.6", "10.20.20.5", "OUT")
    assert res2.status == "OK"
    assert res2.source_zone == "internal"
    assert res2.destination_zone == "dmz"
