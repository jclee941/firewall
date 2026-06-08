"""Verify scripts/make_request_folder.py produces a valid, analyzable folder tree.

LibreOffice macro execution is unavailable, so we prove the generated request
files parse with the same logic the VBA uses (request_parser_oracle) and resolve
to real firewall paths through the route engine seeded from build_xlsm constants,
including multi-firewall (multi-hop) paths.

Run: .venv/bin/python -m pytest tests/test_request_folder.py -v
"""

import glob
import os
import subprocess
import sys

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REQ_DIR = os.path.join(ROOT, "request-folder")
PY = sys.executable

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from request_parser_oracle import parse_request_sheet, sheet_to_filled_rows  # noqa: E402
from route_oracle import Firewall, Network, RouteEngine, RoutingPath  # noqa: E402


@pytest.fixture(scope="module")
def request_tree():
    subprocess.run(
        [PY, os.path.join(ROOT, "scripts", "make_request_folder.py")],
        cwd=ROOT, check=True, capture_output=True, text=True,
    )
    assert os.path.isdir(REQ_DIR)
    return REQ_DIR


def _seed_engine():
    from build_xlsm import FIREWALLS, NETWORK_DEFS, ROUTING_PATHS

    def truthy(v):
        return str(v).upper() in ("Y", "YES", "TRUE", "1")

    nets = [Network(r[0], r[1], r[2], r[3], truthy(r[4])) for r in NETWORK_DEFS[1:]]
    fws = [Firewall(r[0], r[1], truthy(r[2])) for r in FIREWALLS[1:]]
    rps = [RoutingPath(r[0], r[1], r[2], r[3], r[4], int(r[5]), truthy(r[6]))
           for r in ROUTING_PATHS[1:]]
    return RouteEngine(networks=nets, firewalls=fws, routing_paths=rps)


def test_tree_structure(request_tree):
    assert os.path.isfile(os.path.join(request_tree, "README.txt"))
    xlsx = glob.glob(os.path.join(request_tree, "**", "*.xlsx"), recursive=True)
    assert len(xlsx) >= 3, xlsx
    folders = {os.path.basename(os.path.dirname(f)) for f in xlsx}
    # team folders must follow <team>_<docno> so folder parsing yields team/doc_no
    assert any("_" in f and not f.startswith("_") for f in folders), folders
    # a header-only template must exist for operators to copy
    assert any("빈양식" in os.path.basename(f) for f in xlsx), xlsx


def test_empty_template_parses_to_zero_rows(request_tree):
    tmpl = glob.glob(os.path.join(request_tree, "**", "*빈양식*.xlsx"), recursive=True)
    assert tmpl, "blank template missing"
    ws = openpyxl.load_workbook(tmpl[0]).active
    parsed = parse_request_sheet(sheet_to_filled_rows(ws))
    assert parsed == [], "blank template must yield no data rows (header only)"


def test_each_request_parses_and_resolves(request_tree):
    eng = _seed_engine()
    xlsx = sorted(glob.glob(os.path.join(request_tree, "**", "*.xlsx"), recursive=True))
    data_files = [f for f in xlsx if "빈양식" not in os.path.basename(f)]
    assert data_files
    total_rows = 0
    multi_fw = 0
    for f in data_files:
        ws = openpyxl.load_workbook(f).active
        parsed = parse_request_sheet(sheet_to_filled_rows(ws))
        assert parsed, f"no rows parsed from {f}"
        total_rows += len(parsed)
        for req in parsed:
            res = eng.analyze(req["source_ip"], req["dest_ip"], req["direction"])
            assert res.status == "OK", f"{f}: {req} -> {res.status}"
            assert res.target_firewalls, f"{f}: {req} produced no firewalls"
            if len([x for x in res.target_firewalls.split(";") if x]) >= 2:
                multi_fw += 1
    assert total_rows >= 4
    # the folder must exercise multi-firewall (multi-hop) routing, not just 1-hop
    assert multi_fw >= 2, f"expected >=2 multi-firewall requests, got {multi_fw}"
