"""Pure-Python mirror of the VBA request-sheet parsing logic.

Mirrors, line-for-line in behavior:
  - HeaderKey                 (FirewallPolicyAutomation.bas:686)  lower + strip spaces
  - CanonicalHeaderName       (FirewallPolicyAutomation.bas:475)  alias map
  - FindHeaderRow             (FirewallPolicyAutomation.bas:440)  scan rows 1..30 for no/번호
  - BuildHeaderMap            (FirewallPolicyAutomation.bas:460)  canonical name -> column index
  - ValidateRequiredHeaders   (FirewallPolicyAutomation.bas:493)  11 required headers
  - CopyRequestRow extraction (FirewallPolicyAutomation.bas:197)  per-column value pull

This lets us verify, without Excel, that an arbitrary 신청서 layout
(varied column order, alias headers, header not on row 1, B-column No)
parses into the canonical request fields the route engine expects.
"""

from __future__ import annotations


# --------------------------------------------------------------------------- #
# Header normalization (mirror HeaderKey)
# --------------------------------------------------------------------------- #

def header_key(text) -> str:
    # LCase$(Replace(Trim$(headerText), " ", ""))
    if text is None:
        return ""
    return str(text).strip().replace(" ", "").lower()


# --------------------------------------------------------------------------- #
# Alias map (mirror CanonicalHeaderName)
# --------------------------------------------------------------------------- #

_ALIASES = {
    "no": ["no", "번호"],
    "출발지ip": ["출발지ip", "출발ip", "sourceip", "srcip", "src"],
    "출발지": ["출발지", "출발지명", "출발", "source", "srcname"],
    "목적지ip": ["목적지ip", "목적ip", "destinationip", "dstip", "dst"],
    "목적지": ["목적지", "목적지명", "목적", "destination", "dstname"],
    "프로토콜": ["프로토콜", "protocol", "proto"],
    "포트": ["포트", "port", "dport", "목적지포트"],
    "방향": ["방향", "direction"],
    "용도": ["용도", "목적", "usage", "purpose"],
    "시작일": ["시작일", "시작", "startdate", "start"],
    "종료일": ["종료일", "종료", "enddate", "end"],
    "비고": ["비고", "메모", "remark", "remarks", "note"],
}

# build reverse lookup; order mirrors VBA Select Case (first match wins).
# Note: "목적" maps to 용도 in VBA (the 용도 case comes after 목적지; but
# VBA Select Case for "목적" hits the 용도 case at line 485 only if it did
# not already match 목적지 at 481 — 목적지 case is "목적지","목적지명","목적"
# so "목적" actually canonicalizes to 목적지). We mirror VBA exactly:
# iterate cases in source order and return the first whose alias list contains
# the key.
_CASE_ORDER = [
    "no", "출발지ip", "출발지", "목적지ip", "목적지",
    "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고",
]


def canonical_header_name(key: str) -> str:
    for canon in _CASE_ORDER:
        if key in _ALIASES[canon]:
            return canon
    return key  # Case Else: return input unchanged


# --------------------------------------------------------------------------- #
# Sheet model: a sheet is a list[list[cell]] (1-based logic, 0-based storage)
# --------------------------------------------------------------------------- #

REQUIRED = ["출발지ip", "출발지", "목적지ip", "목적지", "프로토콜", "포트",
            "방향", "용도", "시작일", "종료일", "비고"]


class RequestParseError(Exception):
    pass


def _cell(rows: list[list], r1: int, c1: int):
    """1-based access; returns '' when out of range (mirrors empty cell)."""
    if r1 - 1 < 0 or r1 - 1 >= len(rows):
        return ""
    row = rows[r1 - 1]
    if c1 - 1 < 0 or c1 - 1 >= len(row):
        return ""
    v = row[c1 - 1]
    return "" if v is None else v


def _last_column(rows: list[list], r1: int) -> int:
    # mirror End(xlToLeft): index (1-based) of the last non-empty cell in the row
    if r1 - 1 >= len(rows):
        return 0
    row = rows[r1 - 1]
    last = 0
    for i, v in enumerate(row, start=1):
        if v is not None and str(v).strip() != "":
            last = i
    return last


def find_header_row(rows: list[list]) -> int:
    # scan rows 1..30 for a cell whose HeaderKey is "no" or "번호"
    for r1 in range(1, 31):
        last_col = _last_column(rows, r1)
        for c1 in range(1, last_col + 1):
            key = header_key(_cell(rows, r1, c1))
            if key == "no" or key == "번호":
                return r1
    raise RequestParseError("No/번호 헤더 행을 찾을 수 없습니다.")


def build_header_map(rows: list[list], header_row: int) -> dict[str, int]:
    hmap: dict[str, int] = {}
    last_col = _last_column(rows, header_row)
    for c1 in range(1, last_col + 1):
        name = canonical_header_name(header_key(_cell(rows, header_row, c1)))
        if name:
            hmap[name] = c1  # later columns overwrite earlier (mirror VBA dict assign)
    return hmap


def validate_required_headers(hmap: dict[str, int]) -> None:
    for req in REQUIRED:
        if req not in hmap:
            raise RequestParseError(f"필수 컬럼 누락: {req}")


def _source_last_row(rows: list[list], hmap: dict[str, int]) -> int:
    # mirror SourceLastRow: max of last non-empty row over 출발지ip and 목적지ip
    def last_nonempty(col):
        last = 0
        for r1 in range(1, len(rows) + 1):
            if str(_cell(rows, r1, col)).strip() != "":
                last = r1
        return last
    return max(last_nonempty(hmap["출발지ip"]), last_nonempty(hmap["목적지ip"]))


def _row_has_data(rows, r1, hmap) -> bool:
    return (str(_cell(rows, r1, hmap["출발지ip"])).strip() != "" or
            str(_cell(rows, r1, hmap["목적지ip"])).strip() != "")


def parse_request_sheet(rows: list[list]) -> list[dict]:
    """Parse a raw sheet into canonical request dicts.

    Returns one dict per data row with keys mirroring CopyRequestRow:
    source_ip, source_name, dest_ip, dest_name, protocol, port, direction,
    purpose, start_date, end_date, note, source_row.
    Raises RequestParseError on missing header row / required columns.
    """
    header_row = find_header_row(rows)
    hmap = build_header_map(rows, header_row)
    validate_required_headers(hmap)

    last_row = _source_last_row(rows, hmap)
    out: list[dict] = []
    for r1 in range(header_row + 1, last_row + 1):
        if not _row_has_data(rows, r1, hmap):
            continue
        out.append({
            "source_row": r1,
            "source_ip": str(_cell(rows, r1, hmap["출발지ip"])).strip(),
            "source_name": str(_cell(rows, r1, hmap["출발지"])).strip(),
            "dest_ip": str(_cell(rows, r1, hmap["목적지ip"])).strip(),
            "dest_name": str(_cell(rows, r1, hmap["목적지"])).strip(),
            "protocol": str(_cell(rows, r1, hmap["프로토콜"])).strip().upper(),
            "port": str(_cell(rows, r1, hmap["포트"])).strip(),
            "direction": str(_cell(rows, r1, hmap["방향"])).strip(),
            "purpose": str(_cell(rows, r1, hmap["용도"])).strip(),
            "start_date": str(_cell(rows, r1, hmap["시작일"])).strip(),
            "end_date": str(_cell(rows, r1, hmap["종료일"])).strip(),
            "note": str(_cell(rows, r1, hmap["비고"])).strip(),
        })
    return out
