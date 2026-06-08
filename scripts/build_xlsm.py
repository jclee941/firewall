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
import struct
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
    ["key", "value", "설명"],
    ["request_folder", "", "신청서 엑셀이 모여 있는 폴더 경로. 하위 폴더(예: 정보보호센터_1234)까지 재귀 탐색합니다."],
    ["parse_targets", "출발지IP;목적지IP", "적용대상방화벽 산정에 쓸 IP 컬럼(세미콜론 구분). IP 컬럼만 등록."],
    ["route_legacy_fallback", "FALSE", "라우팅 경로를 못 찾을 때 기존 CIDR 겹침 방식으로 대체할지(TRUE/FALSE)."],
    ["header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"],
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


# --------------------------------------------------------------------------- #
# Visual styling (first-open readability; VBA macros may re-format at runtime)
# --------------------------------------------------------------------------- #

_HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
_HEADER_FONT = Font(bold=True, color="1F3864")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
_THIN_BOTTOM = Border(bottom=Side(style="thin", color="9DB2CE"))

_WIDTHS = {
    "requests": {
        "A": 18, "B": 9, "C": 26, "D": 16, "E": 12, "F": 16, "G": 12,
        "H": 10, "I": 8, "J": 8, "K": 18, "L": 12, "M": 12, "N": 14,
        "O": 16, "P": 30, "Q": 34, "R": 24, "S": 14, "T": 16, "U": 22,
        "V": 16, "W": 14, "X": 20,
    },
    "firewalls": {"A": 16, "B": 10, "C": 9, "D": 28},
    "network_definitions": {"A": 14, "B": 18, "C": 12, "D": 10, "E": 9},
    "routing_paths": {"A": 16, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 9},
    "settings": {"A": 22, "B": 26, "C": 60},
    "processing_log": {"A": 20, "B": 22, "C": 10, "D": 12, "E": 40},
    "sample-request-format": {"A": 4, "B": 6, "C": 16, "D": 12, "E": 16,
                              "F": 12, "G": 10, "H": 8, "I": 8, "J": 18,
                              "K": 12, "L": 12, "M": 14},
    "usage": {"A": 8, "B": 70},
}

_FILTER_SHEETS = {"requests", "firewalls", "network_definitions",
                  "routing_paths", "processing_log"}

_FREEZE = {"requests": "E2"}


def _style_sheet(ws, header_row: int = 1):
    from openpyxl.utils import get_column_letter
    for col_letter, width in _WIDTHS.get(ws.title, {}).items():
        ws.column_dimensions[col_letter].width = width
    last_col = ws.max_column
    for c in range(1, last_col + 1):
        cell = ws.cell(row=header_row, column=c)
        if cell.value is None:
            continue
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
        cell.border = _THIN_BOTTOM
    ws.freeze_panes = _FREEZE.get(ws.title, "A2")
    if ws.title in _FILTER_SHEETS:
        ws.auto_filter.ref = (
            f"A{header_row}:{get_column_letter(last_col)}"
            f"{max(ws.max_row, header_row)}"
        )


def _write_rows(ws, rows):
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is not None:
                ws.cell(row=r, column=c, value=val)


def _force_codepage_949(proj) -> None:
    """Patch the parsed VBA project so it serializes with PROJECTCODEPAGE=949.

    pyOpenVBA's serialize_dir_stream reuses the PROJECTINFORMATION prefix from
    proj.dir_raw verbatim (which carries PROJECTCODEPAGE=1252 from the create_new
    template), and rebuild_module_stream MBCS-encodes each module with
    proj.code_page. Setting proj.code_page alone is ignored because the dir prefix
    is spliced unchanged. So we patch the 0x0003 (PROJECTCODEPAGE) record inside
    dir_raw to 949 AND set proj.code_page=949, then mark dir + modules dirty so the
    streams are actually re-encoded. cp949 covers 한글, so the source survives."""
    raw = bytearray(proj.dir_raw)
    i = 0
    patched = False
    while i + 6 <= len(raw):
        rid = struct.unpack_from("<H", raw, i)[0]
        size = struct.unpack_from("<I", raw, i + 2)[0]
        if rid == 0x0003 and size == 2:  # PROJECTCODEPAGE
            struct.pack_into("<H", raw, i + 6, 949)
            patched = True
            break
        i += 6 + size
    if not patched:
        raise RuntimeError("PROJECTCODEPAGE record not found in dir stream")
    proj.dir_raw = bytes(raw)
    proj.code_page = 949
    proj.dir_structure_dirty = True
    for m in proj.modules:
        m.dirty = True


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
    # CRITICAL: force the VBA project code page to 949 (Korean) so 한글 in the
    # module source—including FUNCTIONAL literals like headerMap("\ucd9c\ubc1c\uc9c0ip")—
    # survives. pyOpenVBA's create_new template hardcodes PROJECTCODEPAGE=1252 in
    # dir_raw, and serialize_dir_stream reuses that prefix verbatim; modules are
    # then MBCS-encoded with project.code_page. Left at 1252, every 한글 becomes '?'
    # and the parser cannot match Korean headers at runtime in Excel.
    _force_codepage_949(proj)
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
    _style_sheet(base)

    def add(title, rows):
        ws = wb.create_sheet(title)
        _write_rows(ws, rows)
        _style_sheet(ws)
        return ws

    add("firewalls", FIREWALLS)
    add("network_definitions", NETWORK_DEFS)
    add("routing_paths", ROUTING_PATHS)
    add("settings", SETTINGS)
    add("processing_log", PROCESSING_LOG)
    # sample-request-format has a blank A column header
    sf = wb.create_sheet("sample-request-format")
    _write_rows(sf, SAMPLE_FORMAT)
    _style_sheet(sf)
    add("usage", USAGE)

    wb.save(str(OUT))
    wb.close()

    # 3) report
    size = os.path.getsize(OUT)
    print(f"Built {OUT} ({size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
