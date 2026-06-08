#!/usr/bin/env python3
"""Build a REAL macro-enabled firewall-policy-automation.xlsm on Linux.

No Windows Excel / PowerShell / COM required. Uses:
  - pyOpenVBA  : inject the two VBA modules into a real vbaProject.bin
  - openpyxl   : pre-create and seed all worksheets while preserving the VBA

Run:
  ./.venv/bin/python scripts/build_xlsm.py
Output:
  dist/firewall-policy-automation.xlsm
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parent.parent
VBA_DIR = ROOT / "vba"
DIST = ROOT / "dist"
OUT = DIST / "firewall-policy-automation.xlsm"

MODULES = [
    ("FirewallPolicyAutomation", VBA_DIR / "FirewallPolicyAutomation.bas"),
    ("FirewallRouteAnalysis", VBA_DIR / "FirewallRouteAnalysis.bas"),
]

# ---- seed data (kept in sync with the VBA Write*Headers seeds) ------------- #

REQUESTS_HEADERS = [
    "source_file", "source_row", "target_firewalls",
    "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트", "방향",
    "용도", "시작일", "종료일", "비고",
    "validation_status", "validation_message", "match_details",
    "firewall_path", "source_zone", "destination_zone", "zone_path",
    "request_team", "request_doc_no", "request_folder",
]

FIREWALLS = [
    ["firewall_name", "vendor", "enabled", "comment"],
    ["SECUI-FW-01", "SECUI", "Y", "내부-서버 구간"],
    ["SECUI-FW-02", "SECUI", "Y", "중간-DMZ 구간"],
    ["SECUI-FW-03", "SECUI", "Y", "DMZ-외부 구간"],
]

NETWORK_DEFS = [
    ["network_name", "network_cidr", "zone", "site", "enabled"],
    ["업무PC망", "10.10.0.0/16", "internal", "본사", "Y"],
    ["서버망", "172.16.1.0/24", "server", "IDC", "Y"],
    ["중간망", "10.30.0.0/16", "transit", "IDC", "Y"],
    ["DMZ망", "10.20.0.0/16", "dmz", "IDC", "Y"],
    ["외부", "0.0.0.0/0", "outside", "공통", "Y"],
    ["서버DMZ", "172.16.20.0/24", "dmz", "IDC", "Y"],
]

ROUTING_PATHS = [
    ["firewall_name", "src_zone", "dst_zone", "ingress_if", "egress_if", "path_order", "enabled"],
    ["SECUI-FW-01", "internal", "server", "eth1", "eth2", 10, "Y"],
    ["SECUI-FW-01", "internal", "transit", "eth1", "eth3", 20, "Y"],
    ["SECUI-FW-02", "transit", "dmz", "eth1", "eth2", 30, "Y"],
    ["SECUI-FW-03", "dmz", "outside", "eth1", "eth2", 40, "Y"],
]

SETTINGS = [
    ["key", "value"],
    ["request_folder", ""],
    ["parse_targets", "출발지IP;목적지IP"],
    ["route_legacy_fallback", "FALSE"],
    ["header_alias", ""],
]

PROCESSING_LOG = [["processed_at", "source_file", "status", "merged_rows", "message"]]

SAMPLE_FORMAT = [
    [None, "No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
     "방향", "용도", "시작일", "종료일", "비고"],
    [None, 1, "10.10.10.5", "업무PC", "172.16.1.10", "업무시스템", "TCP", "443",
     "IN", "HTTPS 업무 연동", "2026-01-01", "2026-12-31", "정기 신청"],
]

USAGE = [
    ["Step", "Action"],
    ["1", "firewalls 시트에 방화벽 장비를 등록한다"],
    ["2", "network_definitions 시트에 IP 대역과 zone을 등록한다 (대역정의)"],
    ["3", "routing_paths 시트에 zone-to-zone 경로를 등록한다 (라우팅경로)"],
    ["4", "requests 시트에 신청서를 입력하거나 MergeFirewallRequestFolder로 폴더를 통합한다"],
    ["5", "AnalyzeRequestRoutes 매크로를 실행해 적용대상방화벽 경로를 계산한다"],
]

# Pre-filled example requests so AnalyzeRequestRoutes has something to compute.
# Row 1: single-IP multi-hop (internal->transit->dmz => SECUI-FW-01;SECUI-FW-02)
# Row 2: CIDR (대역) request, also multi-hop, proving CIDR input with zero setup.
EXAMPLE_REQUEST_ROWS = [
    {
        "source_file": "example.xlsx",
        "source_row": 2,
        "출발지IP": "10.10.10.5",
        "출발지": "업무PC",
        "목적지IP": "10.20.20.5",
        "목적지": "DMZ서버",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
        "용도": "단일 IP 경로 예시",
        "시작일": "2026-01-01",
        "종료일": "2026-12-31",
        "비고": "build seed (single IP)",
    },
    {
        "source_file": "example.xlsx",
        "source_row": 3,
        "출발지IP": "10.10.10.0/24",
        "출발지": "업무PC대역",
        "목적지IP": "10.20.20.0/24",
        "목적지": "DMZ대역",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
        "용도": "CIDR 대역 경로 예시",
        "시작일": "2026-01-01",
        "종료일": "2026-12-31",
        "비고": "build seed (CIDR)",
    },
]


def _write_rows(ws, rows):
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is not None:
                ws.cell(row=r, column=c, value=val)


def main() -> int:
    for name, path in MODULES:
        if not path.exists():
            print(f"ERROR: missing VBA module {path}", file=sys.stderr)
            return 1

    DIST.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()

    # 1) create a real macro-enabled workbook and add both modules
    xf = ExcelFile.create_new(str(OUT))
    proj = xf.vba_project()
    for name, path in MODULES:
        src = path.read_text(encoding="utf-8")
        proj.add_module(name, src)
    # remove the default empty Module1 created by create_new
    if "Module1" in proj.module_names():
        proj.delete_module("Module1")
    xf.save()
    xf.close()

    # 2) open with openpyxl preserving VBA, build/seed all sheets
    wb = openpyxl.load_workbook(str(OUT), keep_vba=True)

    # repurpose the default sheet as 'requests'
    base = wb[wb.sheetnames[0]]
    base.title = "requests"
    for c, h in enumerate(REQUESTS_HEADERS, start=1):
        base.cell(row=1, column=c, value=h)
    # seeded example requests (single IP + CIDR), one per row starting at row 2
    col_index = {h: i + 1 for i, h in enumerate(REQUESTS_HEADERS)}
    for r, example in enumerate(EXAMPLE_REQUEST_ROWS, start=2):
        for key, val in example.items():
            base.cell(row=r, column=col_index[key], value=val)
    for cell in base[1]:
        cell.font = Font(bold=True)

    def add(title, rows):
        ws = wb.create_sheet(title)
        _write_rows(ws, rows)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        return ws

    add("firewalls", FIREWALLS)
    add("network_definitions", NETWORK_DEFS)
    add("routing_paths", ROUTING_PATHS)
    add("settings", SETTINGS)
    add("processing_log", PROCESSING_LOG)
    # sample-request-format has a blank A column header
    sf = wb.create_sheet("sample-request-format")
    _write_rows(sf, SAMPLE_FORMAT)
    add("usage", USAGE)

    wb.save(str(OUT))
    wb.close()

    # 3) report
    size = os.path.getsize(OUT)
    print(f"Built {OUT} ({size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
