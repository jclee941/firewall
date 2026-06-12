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
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
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
    "요청부서", "요청번호", "제목", "원본파일", "원본행", "검증상태",
    "대상방화벽",
    "출발지IP", "출발지설명", "목적지IP", "목적지설명", "프로토콜", "포트", "방향",
    "용도", "시작일", "종료일", "비고",
    "검증메시지", "방화벽경로", "출발매칭대역", "목적매칭대역", "대역경로",
    "매칭근거", "요청폴더",
]

FIREWALLS = [
    ["firewall_name", "vendor", "enabled", "comment"],
    ["SECUI-FW-01", "SECUI", "Y", "\ub0b4\ubd80-\uc11c\ubc84 \uad6c\uac04"],
    ["SECUI-FW-02", "SECUI", "Y", "\uc11c\ubc84-DMZ \uad6c\uac04"],
    ["SECUI-FW-03", "SECUI", "Y", "DMZ-\uc678\ubd80 \uad6c\uac04"],
]

FIREWALL_RANGES = [
    ["firewall_name", "source_cidr", "destination_cidr", "direction", "path_order", "enabled", "comment"],
    ["SECUI-FW-01", "10.10.0.0/16", "172.16.0.0/16", "OUT", 10, "Y", "\uc5c5\ubb34PC -> \uc11c\ubc84"],
    ["SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10, "Y", "\uc5c5\ubb34PC -> DMZ"],
    ["SECUI-FW-02", "10.10.0.0/16", "10.20.0.0/16", "OUT", 20, "Y", "\uc5c5\ubb34PC -> DMZ"],
    ["SECUI-FW-01", "10.10.0.0/16", "8.8.8.0/24", "OUT", 10, "Y", "\uc5c5\ubb34PC -> \uc678\ubd80 DNS"],
    ["SECUI-FW-02", "10.10.0.0/16", "8.8.8.0/24", "OUT", 20, "Y", "\uc5c5\ubb34PC -> \uc678\ubd80 DNS"],
    ["SECUI-FW-03", "10.10.0.0/16", "8.8.8.0/24", "OUT", 30, "Y", "\uc5c5\ubb34PC -> \uc678\ubd80 DNS"],
]

SETTINGS = [
    ["key", "value", "설명"],
    ["request_folder", "", "신청서 엑셀이 모여 있는 폴더 경로. 하위 폴더(예: 정보보호센터_1234)까지 재귀 탐색합니다."],
    ["parse_sheet", "", "파싱할 시트 이름(정확히 일치). 비워두면 헤더로 자동 감지합니다."],
    ["parse_targets", "출발지IP;목적지IP", "(사용 안 함/예약) 현재 동작에 영향 없음. 출발지IP와 목적지IP는 항상 필수입니다."],
    ["header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"],
]
PROCESSING_LOG = [["processed_at", "source_file", "status", "merged_rows", "message"]]

SECUI_BATCH_HEADERS = [
    "No", "장비명", "정책명", "출발지주소", "출발지명", "목적지주소", "목적지명",
    "서비스", "프로토콜", "목적지포트", "동작", "로그", "사용여부", "시작일",
    "종료일", "설명", "신청부서", "신청번호", "원본파일", "원본행",
]

SECUI_CLI_HEADERS = [
    "No", "장비명", "정책명", "명령어", "검토메모", "신청부서", "신청번호",
    "원본파일", "원본행",
]

# header_aliases: table-form alias mapping. Operators add one row per
# non-standard header in their request files. standard = the canonical column
# (or any built-in alias of it); your_column = the actual header in the request.
HEADER_ALIASES = [
    ["standard", "your_column", "\uc124\uba85"],
    ["\ucd9c\ubc1c\uc9c0", "", "\uc2e0\uccad\uc11c\uc758 \ucd9c\ubc1c\uc9c0 \uc774\ub984/\uc124\uba85 \uceec\ub7fc\uba85 (\uc608: \ucd9c\ubc1c\uc9c0ip\uc124\uba85)"],
    ["\ubaa9\uc801\uc9c0", "", "\uc2e0\uccad\uc11c\uc758 \ubaa9\uc801\uc9c0 \uc774\ub984/\uc124\uba85 \uceec\ub7fc\uba85"],
    ["\ud504\ub85c\ud1a0\ucf5c", "", "\uc608: tcp/udp"],
    ["\ubc29\ud5a5", "", "\uc608: \uad6c\ubd84"],
    ["\uc2dc\uc791\uc77c", "", "\uc608: \uc2dc\uc791\uc77c\uc790"],
    ["\uc885\ub8cc\uc77c", "", "\uc608: \uc885\ub8cc\uc77c\uc790"],
]

SAMPLE_FORMAT = [
    [None, "No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
     "방향", "용도", "시작일", "종료일", "비고"],
    [None, 1, "10.10.10.5", "업무PC", "172.16.1.10", "업무시스템", "TCP", "443",
     "IN", "HTTPS 업무 연동", "2026-01-01", "2026-12-31", "정기 신청"],
]

USAGE = [
    ["Step", "Action"],
    ["1", "firewalls 시트에 방화벽 장비명, 벤더, 사용여부를 등록한다"],
    ["2", "firewall_ranges 시트에 출발지대역, 목적지대역, 방향, 순서를 등록한다"],
    ["3", "대역은 IP/CIDR/ANY를 쓸 수 있고 여러 값은 세미콜론·콤마·줄바꿈·공백으로 구분한다"],
    ["4", "settings 시트의 request_folder에 신청서 폴더 경로를 적거나 SelectRequestFolder 매크로로 폴더를 선택한다"],
    ["5", "requests 시트에 직접 입력하거나 MergeFirewallRequestFolder 매크로로 폴더 안 신청서를 통합한다 (Alt+F8)"],
    ["6", "AnalyzeRequestRoutes 매크로를 실행해 대상방화벽과 검증 상태를 계산한다"],
    ["7", "ConvertRequestsToSecuiBatch 매크로로 requests 결과를 secui_batch 장비별 배치 양식으로 변환한다"],
    ["8", "ConvertRequestsToSecuiCli 매크로로 requests 결과를 secui_cli 장비별 CLI 명령 초안으로 변환한다"],
    ["⚠", "입력 시트(녹색·황색 탭)는 보호되어 있다. 헤더는 수정 불가, 데이터 입력 영역만 타이핑 가능"],
    ["ℹ", "requests·processing_log(파랑·회색 탭)은 매크로가 자동으로 채운다. 직접 수정하지 않는다"],
    ["💡", "프로토콜·방향 셀은 드롭다운 목록에서 선택. 출발지IP·목적지IP가 비면 빨간색으로 표시된다"],
]

# Pre-filled example requests so AnalyzeRequestRoutes has something to compute.
# Row 1: single-IP multi-hop (internal->transit->dmz => SECUI-FW-01;SECUI-FW-02)
# Row 2: CIDR (대역) request, also multi-hop, proving CIDR input with zero setup.
EXAMPLE_REQUEST_ROWS = [
    {
        "요청부서": "정보보호센터",
        "요청번호": "1234",
        "제목": "웹서비스 연동",
        "원본파일": "example.xlsx",
        "원본행": 2,
        "출발지IP": "10.10.10.5",
        "출발지설명": "업무PC",
        "목적지IP": "10.20.20.5",
        "목적지설명": "DMZ서버",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
        "용도": "단일 IP 경로 예시",
        "시작일": "2026-01-01",
        "종료일": "2026-12-31",
        "비고": "build seed (single IP)",
    },
    {
        "요청부서": "정보보호센터",
        "요청번호": "1234",
        "제목": "웹서비스 연동",
        "원본파일": "example.xlsx",
        "원본행": 3,
        "출발지IP": "10.10.10.0/24",
        "출발지설명": "업무PC대역",
        "목적지IP": "10.20.20.0/24",
        "목적지설명": "DMZ대역",
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
        "A": 16, "B": 12, "C": 20, "D": 28, "E": 8, "F": 18, "G": 32,
        "H": 18, "I": 16, "J": 18, "K": 16, "L": 10, "M": 14, "N": 10,
        "O": 28, "P": 12, "Q": 12, "R": 24, "S": 40, "T": 34, "U": 18,
        "V": 18, "W": 30, "X": 60, "Y": 24,
    },
    "firewalls": {"A": 16, "B": 10, "C": 9, "D": 28},
    "firewall_ranges": {"A": 16, "B": 18, "C": 18, "D": 10, "E": 12, "F": 9, "G": 36},
    "settings": {"A": 22, "B": 26, "C": 60},
    "header_aliases": {"A": 16, "B": 22, "C": 44},
    "processing_log": {"A": 20, "B": 22, "C": 10, "D": 12, "E": 40},
    "secui_batch": {
        "A": 6, "B": 18, "C": 36, "D": 18, "E": 18, "F": 18, "G": 18,
        "H": 16, "I": 10, "J": 12, "K": 10, "L": 8, "M": 10, "N": 12,
        "O": 12, "P": 42, "Q": 16, "R": 14, "S": 24, "T": 10,
    },
    "secui_cli": {"A": 6, "B": 18, "C": 36, "D": 120, "E": 60,
                  "F": 16, "G": 14, "H": 24, "I": 10},
    "sample-request-format": {"A": 4, "B": 6, "C": 16, "D": 12, "E": 16,
                              "F": 12, "G": 10, "H": 8, "I": 8, "J": 18,
                              "K": 12, "L": 12, "M": 14},
    "usage": {"A": 8, "B": 70},
}

_FILTER_SHEETS = {"requests", "firewalls", "firewall_ranges", "processing_log"}

# requests output layout: row 1 = cosmetic group labels, row 2 = leaf headers,
# data from row 3. Mirror of VBA REQ_HEADER_GROUP_ROW / REQ_HEADER_ROW / REQ_DATA_START_ROW.
_REQ_HEADER_GROUP_ROW = 1
_REQ_HEADER_ROW = 2
_REQ_DATA_START_ROW = 3
_FREEZE = {"requests": "H3"}


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

# --------------------------------------------------------------------------- #
# Build-time UX (scope B): input-assist + display only. Never changes a cell
# value, header, column count, or the route algorithm. All openpyxl-native;
# vbaProject.bin is preserved by keep_vba=True on save.
# --------------------------------------------------------------------------- #

_UX_LAST_ROW = 5000  # bound validations/CF so we never explode to 1048576 rows

# requests column ordinals (1-indexed) we attach UX to
_REQ_PROTOCOL_COL = 12    # 프로토콜
_REQ_PORT_COL = 13        # 포트
_REQ_DIRECTION_COL = 14   # 방향
_REQ_TARGET_COL = 7
_REQ_SRC_IP_COL = 8       # 출발지IP (required)
_REQ_DST_IP_COL = 10      # 목적지IP (required)

# tab colors: input sheets vs result/log sheets (display navigation only)
_TAB_COLORS = {
    "requests": "FF4472C4",            # result (blue)
    "firewalls": "FF70AD47",           # input (green)
    "firewall_ranges": "FFFFC000",
    "settings": "FFFFC000",            # config (amber)
    "header_aliases": "FFFFC000",
    "processing_log": "FFA6A6A6",      # log (grey)
    "secui_batch": "FF4472C4",
    "secui_cli": "FF4472C4",
    "sample-request-format": "FFED7D31",  # sample (orange)
    "usage": "FFED7D31",
}

_EMPTY_REQUIRED_FILL = PatternFill("solid", fgColor="FFC7CE")  # light red
_TARGET_FIREWALL_FILL = PatternFill("solid", fgColor="D9EAD3")

# operator-input sheets that MAY be protected. requests/processing_log are
# written by VBA at runtime and must NEVER be protected.
_PROTECT_SHEETS = {
    "firewalls", "firewall_ranges", "settings", "header_aliases",
}
_NO_PROTECT_SHEETS = {"requests", "processing_log", "secui_batch", "secui_cli"}


def _add_list_dv(ws, col_letter, values, *, allow_blank=True, start_row=2):
    """Attach a bounded list dropdown to col_letter rows start_row.._UX_LAST_ROW."""
    formula = '"' + ",".join(values) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
    dv.error = "목록에서 값을 선택하세요."
    dv.errorTitle = "잘못된 입력"
    dv.prompt = "드롭다운에서 선택"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{start_row}:{col_letter}{_UX_LAST_ROW}")
    return dv


def _apply_ux(wb) -> None:
    from openpyxl.utils import get_column_letter

    # 1) tab colors --------------------------------------------------------- #
    for name, color in _TAB_COLORS.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color

    req = wb["requests"]

    # 2) dropdowns on requests (프로토콜 / 방향) data starts at row 3 ------------- #
    _add_list_dv(req, get_column_letter(_REQ_PROTOCOL_COL),
                 ["TCP", "UDP", "ICMP"], start_row=_REQ_DATA_START_ROW)
    # 방향 values must all be accepted by VBA NormalizeDirection (IN/OUT/BOTH)
    _add_list_dv(req, get_column_letter(_REQ_DIRECTION_COL),
                 ["IN", "OUT", "BOTH"], start_row=_REQ_DATA_START_ROW)

    # mirror the demo dropdowns onto sample-request-format (cols G=7,I=9)
    if "sample-request-format" in wb.sheetnames:
        sf = wb["sample-request-format"]
        _add_list_dv(sf, "G", ["TCP", "UDP", "ICMP"])
        _add_list_dv(sf, "I", ["IN", "OUT", "BOTH"])

    # 3) firewalls.enabled dropdown (VBA-truthy values only) ----------------- #
    fw = wb["firewalls"]
    _add_list_dv(fw, "C", ["Y", "N"])
    ranges = wb["firewall_ranges"]
    _add_list_dv(ranges, "D", ["OUT", "IN", "BOTH"])
    _add_list_dv(ranges, "F", ["Y", "N"])

    # 4) conditional format: flag empty required IP cells -------------------- #
    src_letter = get_column_letter(_REQ_SRC_IP_COL)  # H
    dst_letter = get_column_letter(_REQ_DST_IP_COL)  # J
    for letter in (src_letter, dst_letter):
        rng = f"{letter}{_REQ_DATA_START_ROW}:{letter}{_UX_LAST_ROW}"
        # highlight when the cell is blank (display hint, never edits a value)
        rule = FormulaRule(
            formula=[f'ISBLANK({letter}{_REQ_DATA_START_ROW})'],
            fill=_EMPTY_REQUIRED_FILL,
            stopIfTrue=False,
        )
        req.conditional_formatting.add(rng, rule)
    target_letter = get_column_letter(_REQ_TARGET_COL)
    target_rng = f"{target_letter}{_REQ_DATA_START_ROW}:{target_letter}{_UX_LAST_ROW}"
    target_rule = FormulaRule(
        formula=[f'LEN(TRIM({target_letter}{_REQ_DATA_START_ROW}))>0'],
        fill=_TARGET_FIREWALL_FILL,
        stopIfTrue=False,
    )
    req.conditional_formatting.add(target_rng, target_rule)

    # 5) header hint comments on key input columns -------------------------- #
    _src_hint = (
        "출발지 IP 또는 CIDR\n"
        "예: 10.10.10.0/24 또는 10.10.10.5\n"
        "여러 개는 ; 로 구분"
    )
    _dst_hint = (
        "목적지 IP 또는 CIDR\n"
        "예: 172.16.1.10 또는 172.16.1.0/24\n"
        "여러 개는 ; 로 구분"
    )
    req.cell(_REQ_HEADER_ROW, _REQ_SRC_IP_COL).comment = Comment(_src_hint, "firewall-automation")
    req.cell(_REQ_HEADER_ROW, _REQ_DST_IP_COL).comment = Comment(_dst_hint, "firewall-automation")

    # 6) sheet protection (operator-input sheets only) ---------------------- #
    #    requests / processing_log are macro-written -> never protected.
    #    We UNLOCK the data-entry area rows 2.._UX_LAST_ROW across all columns so
    #    the editable span EXACTLY matches the validation/CF range advertised to
    #    Excel (no footgun where a dropdown row is silently un-typeable). The row
    #    limit is bounded (_UX_LAST_ROW) so per-cell <protection> nodes stay small.
    #    Headers (row 1) and everything below the limit stay locked.
    from openpyxl.styles import Protection
    unlocked = Protection(locked=False)
    for name in _PROTECT_SHEETS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        last_col = ws.max_column
        for r in range(2, _UX_LAST_ROW + 1):
            for c in range(1, last_col + 1):
                ws.cell(row=r, column=c).protection = unlocked
        # enable() flips protection.sheet=True and applies the option flags below
        ws.protection.enable()
        # allow common operator gestures even while protected
        ws.protection.autoFilter = False
        ws.protection.sort = False
        ws.protection.formatCells = True

_AUTO_RUN_BODY = (
    "\r\n"
    "Private Sub Workbook_Open()\r\n"
    "    ' Auto-run on open: integrate + analyze the request folder.\r\n"
    "    On Error GoTo AutoRunErr\r\n"
    "    Application.Run \"FirewallPolicyAutomation.MergeFirewallRequestFolder\"\r\n"
    "    Exit Sub\r\n"
    "AutoRunErr:\r\n"
    "    MsgBox \"\uc790\ub3d9 \ud1b5\ud569 \uc2e4\ud589 \uc911 \uc624\ub958: \" & Err.Description, vbExclamation\r\n"
    "End Sub\r\n"
)


def _inject_auto_run(proj) -> None:
    """Append Workbook_Open to the ThisWorkbook document module so the workbook
    auto-integrates the request folder when opened (macros must be enabled)."""
    m = proj.get_module("ThisWorkbook")
    if "Workbook_Open" in m.source:
        return
    m.source = m.source.rstrip("\r\n") + _AUTO_RUN_BODY
    m.dirty = True


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
    # auto-run on open: inject Workbook_Open into the ThisWorkbook document module
    _inject_auto_run(proj)
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
    # Row 1: cosmetic group labels (출발지 over IP+설명, 목적지 over IP+설명).
    # Row 2: canonical leaf headers. Data from row 3. Mirrors VBA WriteRequestHeaders.
    col_index = {h: i + 1 for i, h in enumerate(REQUESTS_HEADERS)}
    src_ip_col = col_index["출발지IP"]
    src_desc_col = col_index["출발지설명"]
    dst_ip_col = col_index["목적지IP"]
    dst_desc_col = col_index["목적지설명"]
    base.cell(row=_REQ_HEADER_GROUP_ROW, column=src_ip_col, value="출발지")
    base.cell(row=_REQ_HEADER_GROUP_ROW, column=dst_ip_col, value="목적지")
    from openpyxl.utils import get_column_letter as _gcl
    base.merge_cells(f"{_gcl(src_ip_col)}{_REQ_HEADER_GROUP_ROW}:{_gcl(src_desc_col)}{_REQ_HEADER_GROUP_ROW}")
    base.merge_cells(f"{_gcl(dst_ip_col)}{_REQ_HEADER_GROUP_ROW}:{_gcl(dst_desc_col)}{_REQ_HEADER_GROUP_ROW}")
    for c, h in enumerate(REQUESTS_HEADERS, start=1):
        base.cell(row=_REQ_HEADER_ROW, column=c, value=h)
    # seeded example requests (single IP + CIDR), one per row starting at data row
    for r, example in enumerate(EXAMPLE_REQUEST_ROWS, start=_REQ_DATA_START_ROW):
        for key, val in example.items():
            base.cell(row=r, column=col_index[key], value=val)
    # style the group label row too, then the leaf header row
    for c in (src_ip_col, dst_ip_col):
        gc = base.cell(row=_REQ_HEADER_GROUP_ROW, column=c)
        gc.font = _HEADER_FONT
        gc.fill = _HEADER_FILL
        gc.alignment = _HEADER_ALIGN
    _style_sheet(base, header_row=_REQ_HEADER_ROW)

    def add(title, rows):
        ws = wb.create_sheet(title)
        _write_rows(ws, rows)
        _style_sheet(ws)
        return ws

    add("firewalls", FIREWALLS)
    add("firewall_ranges", FIREWALL_RANGES)
    add("settings", SETTINGS)
    add("header_aliases", HEADER_ALIASES)
    add("processing_log", PROCESSING_LOG)
    add("secui_batch", [SECUI_BATCH_HEADERS])
    add("secui_cli", [SECUI_CLI_HEADERS])
    # sample-request-format has a blank A column header
    sf = wb.create_sheet("sample-request-format")
    _write_rows(sf, SAMPLE_FORMAT)
    _style_sheet(sf)
    add("usage", USAGE)

    # build-time UX (input-assist / display only; after all sheets are seeded)
    _apply_ux(wb)

    wb.save(str(OUT))
    wb.close()

    # 3) report
    size = os.path.getsize(OUT)
    print(f"Built {OUT} ({size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
