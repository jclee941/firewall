"""Pure-Python mirror of the VBA request-sheet parsing logic.

Mirrors, line-for-line in behavior:
  - HeaderKey                 (FirewallPolicyAutomation.bas)  lower + strip spaces
  - CanonicalHeaderName       (FirewallPolicyAutomation.bas)  alias map
  - FindHeaderRow             (FirewallPolicyAutomation.bas)  scan rows 1..30 for no/번호
  - BuildHeaderMap            (FirewallPolicyAutomation.bas)  canonical name -> column index
  - ValidateRequiredHeaders   (FirewallPolicyAutomation.bas)  11 required headers
  - CopyRequestRow extraction (FirewallPolicyAutomation.bas)  per-column value pull

This lets us verify, without Excel, that an arbitrary 신청서 layout
(varied column order, alias headers, header not on row 1, B-column No)
parses into the canonical request fields the route engine expects.
"""

from __future__ import annotations

from datetime import date, datetime

type CellValue = str | int | float | bool | date | datetime | None
type SheetRows = list[list[CellValue]]
type RequestRow = dict[str, CellValue]
type NamedSheetRows = list[tuple[str, SheetRows]]


def _format_metadata_date(value) -> str:
    """Format a date-typed cell as yyyy-mm-dd (locale-independent).

    Mirrors VBA: 시작일/종료일 columns format real Date cells with
    Format$(cell.Value, "yyyy-mm-dd"); string cells stay as-is.
    Only applied to the start/end date columns.
    """
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value).strip() if value is not None else ""


# --------------------------------------------------------------------------- #
# Header normalization (mirror HeaderKey)
# --------------------------------------------------------------------------- #

def header_key(text) -> str:
    # LCase$(Replace(Trim$(headerText), " ", "")) then drop decorating
    # punctuation commonly wrapping headers (No. / No# / (No) / No: ).
    if text is None:
        return ""
    s = str(text).strip().replace(" ", "").lower()
    # drop decorating punctuation (No. / No# / (No) / No:), but only if something
    # remains — a token that is ALL punctuation (e.g. '#') is kept so it can match
    # a symbolic alias instead of collapsing to an empty (ignored) key.
    stripped = s.strip(".#:/()[]-．。")
    return stripped if stripped else s


# --------------------------------------------------------------------------- #
# Alias map (mirror CanonicalHeaderName)
# --------------------------------------------------------------------------- #

_ALIASES = {
    "no": ["no", "\ubc88\ud638", "\uc21c\ubc88", "\uc5f0\ubc88", "seq", "\uc21c\uc11c", "\ud56d\ubc88", "\uc77c\ub828\ubc88\ud638", "\ubc88", "#"],
    "\ub300\uc0c1\ubc29\ud654\ubcbd": ["\ub300\uc0c1\ubc29\ud654\ubcbd", "\ub300\uc0c1fw", "targetfirewalls", "targetfirewall",
               "firewall", "firewalls", "fw", "\uc7a5\ube44\uba85", "\ubc29\ud654\ubcbd\uba85"],
    "\ucd9c\ubc1c\uc9c0ip": ["\ucd9c\ubc1c\uc9c0ip", "\ucd9c\ubc1cip", "sourceip", "source", "srcip", "src",
                "\ucd9c\ubc1c\uc9c0\uc8fc\uc18c", "\uc1a1\uc2e0ip", "\uc6d0\ubcf8ip"],
    "\ucd9c\ubc1c\uc9c0": ["\ucd9c\ubc1c\uc9c0", "\ucd9c\ubc1c\uc9c0\uba85", "\ucd9c\ubc1c", "srcname",
              "\ucd9c\ubc1c\uc9c0\uc124\uba85", "\ucd9c\ubc1c\uc9c0ip\uc124\uba85", "\ucd9c\ubc1c\uc9c0ip\uc124\uba78",
              "\ucd9c\ubc1c\uc9c0\uc124\uba78", "\ucd9c\ubc1cip\uc124\uba85", "\ucd9c\ubc1cip\uc124\uba78",
              "\ucd9c\ubc1c\uc9c0\ub0b4\uc6a9", "\uc1a1\uc2e0\uc790", "src\uc124\uba85"],
    "\ubaa9\uc801\uc9c0ip": ["\ubaa9\uc801\uc9c0ip", "\ubaa9\uc801ip", "destinationip", "destination", "destiation", "dstip", "dst",
                "\ubaa9\uc801\uc9c0\uc8fc\uc18c", "\uc218\uc2e0ip"],
    "\ubaa9\uc801\uc9c0": ["\ubaa9\uc801\uc9c0", "\ubaa9\uc801\uc9c0\uba85", "\ubaa9\uc801", "dstname",
              "\ubaa9\uc801\uc9c0\uc124\uba85", "\ubaa9\uc801\uc9c0ip\uc124\uba85", "\ubaa9\uc801\uc9c0ip\uc124\uba78",
              "\ubaa9\uc801\uc9c0\uc124\uba78", "\ubaa9\uc801ip\uc124\uba85", "\ubaa9\uc801ip\uc124\uba78",
              "\ubaa9\uc801\uc9c0\ub0b4\uc6a9", "\uc218\uc2e0\uc790", "dst\uc124\uba85"],
    "\ud504\ub85c\ud1a0\ucf5c": ["\ud504\ub85c\ud1a0\ucf5c", "protocol", "proto", "tcp/udp", "tcpudp", "\ud504\ub85c\ud1a0",
                "\uc11c\ube44\uc2a4", "\ud504\ub85c\ud1a0\ucf5c\uad6c\ubd84", "l4"],
    "\ud3ec\ud2b8": ["\ud3ec\ud2b8", "port", "dport", "\ubaa9\uc801\uc9c0\ud3ec\ud2b8", "\uc11c\ube44\uc2a4\ud3ec\ud2b8",
             "\ud3ec\ud2b8\ubc88\ud638", "dstport", "service"],
    "\ubc29\ud5a5": ["\ubc29\ud5a5", "direction", "\uad6c\ubd84", "\ubc29\ud5a5\uad6c\ubd84", "inout", "in/out",
             "\uc1a1\uc218\uc2e0", "\uc1a1\uc218\uc2e0\uad6c\ubd84"],
    "\uc6a9\ub3c4": ["\uc6a9\ub3c4", "\ubaa9\uc801", "usage", "purpose", "\uc0ac\uc6a9\uc6a9\ub3c4", "\uc2e0\uccad\uc0ac\uc720", "\uc124\uba85"],
    "\uc2dc\uc791\uc77c": ["\uc2dc\uc791\uc77c", "\uc2dc\uc791", "startdate", "start", "\uc2dc\uc791\uc77c\uc790", "\uc2dc\uc791\ub0a0\uc9dc",
              "\uc801\uc6a9\uc77c", "\uc801\uc6a9\uc2dc\uc791\uc77c", "\uc0ac\uc6a9\uc2dc\uc791\uc77c"],
    "\uc885\ub8cc\uc77c": ["\uc885\ub8cc\uc77c", "\uc885\ub8cc", "enddate", "end", "\uc885\ub8cc\uc77c\uc790", "\uc885\ub8cc\ub0a0\uc9dc",
              "\ub9cc\ub8cc\uc77c", "\uc801\uc6a9\uc885\ub8cc\uc77c", "\uc0ac\uc6a9\uc885\ub8cc\uc77c"],
    "\ube44\uace0": ["\ube44\uace0", "\uba54\ubaa8", "remark", "remarks", "note", "\ucc38\uace0"],
}

# build reverse lookup; order mirrors VBA Select Case (first match wins).
# Note: "목적" maps to 용도 in VBA (the 용도 case comes after 목적지; but
# VBA Select Case for "목적" hits the 용도 case at line 485 only if it did
# not already match 목적지 at 481 — 목적지 case is "목적지","목적지명","목적"
# so "목적" actually canonicalizes to 목적지). We mirror VBA exactly:
# iterate cases in source order and return the first whose alias list contains
# the key.
_CASE_ORDER = [
    "no", "대상방화벽", "출발지ip", "출발지", "목적지ip", "목적지",
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

# Only the IP columns are truly required (the route analysis needs them). All
# other columns are optional metadata: read if present, blank if absent. This
# lets request forms that omit 비고/용도/날짜 etc. still merge.
REQUIRED = ["출발지ip", "목적지ip"]


class RequestParseError(Exception):
    pass


def sheet_to_filled_rows(ws, user_aliases: dict[str, str] | None = None) -> SheetRows:
    """Read an openpyxl worksheet into list[list], propagating merged-cell
    values ONLY across DATA rows (rows after the header row).

    Rationale (Option C, confirmed against Excel/VBA semantics):
      - Header row: a merged header's value lives only in the top-left cell;
        filling it would create duplicate header names and (last-write-wins)
        register the WRONG column. So the header row is left RAW.
      - Data rows: a vertically merged request field (e.g. 출발지IP merged
        B2:B3) must apply to every logical row, so we fill those down.
    The VBA mirror reads headers with raw Cells(headerRow,c).Value and reads
    data cells via .MergeArea.Cells(1,1).Value when .MergeCells is True.
    """
    # snapshot raw values (merged non-top-left cells are None here)
    raw_rows = [[c.value for c in row] for row in ws.iter_rows()]
    rows = [list(row) for row in raw_rows]
    # locate header row on the RAW grid (No/번호 lives in the merge top-left)
    try:
        header_row = find_header_row(rows, user_aliases)
    except RequestParseError:
        header_row = 0  # no header found; fill nothing, let caller raise
    # propagate top-left value across merged ranges, DATA rows only
    for rng in list(ws.merged_cells.ranges):
        top = ws.cell(rng.min_row, rng.min_col).value
        if top is None:
            continue
        for r in range(rng.min_row, rng.max_row + 1):
            if header_row and r <= header_row:
                continue  # never fill the header row (Option C)
            for c in range(rng.min_col, rng.max_col + 1):
                ri, ci = r - 1, c - 1
                while ri >= len(rows):
                    rows.append([])
                while ci >= len(rows[ri]):
                    rows[ri].append(None)
                if rows[ri][ci] is None:
                    rows[ri][ci] = top
    try:
        hmap = build_header_map(raw_rows, header_row, user_aliases)
    except RequestParseError:
        return rows
    for r1 in range(header_row + 1, len(rows) + 1):
        if _is_merged_continuation_row(ws, r1) and not _row_has_own_request_value(raw_rows, r1, hmap):
            rows[r1 - 1] = [None for _ in rows[r1 - 1]]
    return rows


def _is_merged_continuation_row(ws, r1: int) -> bool:
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row < r1 <= rng.max_row:
            return True
    return False


def _row_has_own_request_value(rows: SheetRows, r1: int, hmap: dict[str, int]) -> bool:
    for name, col in hmap.items():
        if name == "no":
            continue
        value = _cell(rows, r1, col)
        if str(value).strip() != "":
            return True
    return False


def _cell(rows: SheetRows, r1: int, c1: int) -> CellValue | str:
    """1-based access; returns '' when out of range (mirrors empty cell)."""
    if r1 - 1 < 0 or r1 - 1 >= len(rows):
        return ""
    row = rows[r1 - 1]
    if c1 - 1 < 0 or c1 - 1 >= len(row):
        return ""
    v = row[c1 - 1]
    return "" if v is None else v


def _opt(rows: SheetRows, r1: int, hmap: dict[str, int], name: str) -> CellValue | str:
    """Read an OPTIONAL column by canonical name: '' if the column is absent."""
    col = hmap.get(name)
    if not col:
        return ""
    return _cell(rows, r1, col)


import re as _re


def _norm_list(value) -> str:
    """Normalize a LIST-like cell (IP/port/protocol): newlines/tabs/commas/runs of
    whitespace become a single ';' so two values stay distinct and don't visually
    merge (e.g. '80\n443' -> '80;443'). An already-concatenated '80443' (no
    separator) cannot be un-merged and is left as-is."""
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ")
    # any of CRLF/CR/LF/TAB/comma/fullwidth comma·semicolon -> ';'
    s = _re.sub(r"[\r\n\t,\uff0c\uff1b]+", ";", s)
    # runs of spaces between tokens -> ';' (list fields: space is a delimiter)
    s = _re.sub(r" +", ";", s.strip())
    # collapse repeated ';' and trim
    s = _re.sub(r";+", ";", s).strip(";")
    return s


def _norm_text(value) -> str:
    """Normalize a PROSE cell (name/purpose/note/direction): only newlines/tabs
    become '; ' so descriptions don't visually concatenate; internal spaces are
    preserved (e.g. 'HTTPS 업무 연동' stays intact)."""
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ")
    s = _re.sub(r"[\r\n\t]+", "; ", s)
    s = _re.sub(r" +", " ", s).strip()
    s = _re.sub(r"(; )+", "; ", s).strip("; ").strip()
    return s


def _last_column(rows: SheetRows, r1: int) -> int:
    # mirror End(xlToLeft): index (1-based) of the last non-empty cell in the row
    if r1 - 1 >= len(rows):
        return 0
    row = rows[r1 - 1]
    last = 0
    for i, v in enumerate(row, start=1):
        if v is not None and str(v).strip() != "":
            last = i
    return last


# Field headers (besides 'no') whose presence identifies a header row.
_FIELD_HEADERS = frozenset(_CASE_ORDER) - {"no"}
_IP_HEADERS = frozenset({"\ucd9c\ubc1c\uc9c0ip", "\ubaa9\uc801\uc9c0ip"})


def _row_header_signature(rows: SheetRows, r1: int,
                          user_aliases: dict[str, str] | None = None) -> tuple[int, bool, bool]:
    """Return (field_count, has_no, has_ip) describing how header-like a row is.

    A cell is a recognized field header if it canonicalizes (built-in alias, then
    user alias) to a known canonical name. This is the same mapping used to build
    the column map, so a row that scores high here IS the row whose columns we can
    actually read."""
    last_col = _last_column(rows, r1)
    canon_seen: set[str] = set()
    has_no = False
    for c1 in range(1, last_col + 1):
        key = header_key(_cell(rows, r1, c1))
        if not key:
            continue
        name = canonical_header_name(key)
        if name == key and user_aliases and key in user_aliases:
            name = user_aliases[key]
        if name == "no":
            has_no = True
        elif name in _FIELD_HEADERS:
            canon_seen.add(name)
    has_ip = bool(canon_seen & _IP_HEADERS)
    return len(canon_seen), has_no, has_ip


def find_header_row(rows: SheetRows,
                    user_aliases: dict[str, str] | None = None) -> int:
    """Locate the header row by its HEADER CONTENT (not by requiring a 'No' cell).

    Scans rows 1..30 and picks the row with the most recognized field headers,
    requiring at least one IP column (출발지IP / 목적지IP) so a stray metadata
    row never wins. A 'No'/번호 cell is an optional tie-breaker, no longer
    mandatory — forms that omit it still parse."""
    best_row = 0
    best_key = (0, False)  # (field_count, has_no)
    for r1 in range(1, 31):
        field_count, has_no, has_ip = _row_header_signature(rows, r1, user_aliases)
        if not has_ip:
            continue
        # need a real header row: an IP column plus at least one more field, OR a
        # 'No' anchor alongside the IP column.
        if field_count < 2 and not has_no:
            continue
        key = (field_count, has_no)
        if key > best_key:
            best_key = key
            best_row = r1
    if best_row:
        return best_row
    raise RequestParseError(
        "\ud5e4\ub354 \ud589\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: \ucd9c\ubc1c\uc9c0IP/\ubaa9\uc801\uc9c0IP \uc5f4\uc774 \uc788\ub294 \ud589\uc774 \ud544\uc694\ud569\ub2c8\ub2e4.")

def sheet_header_score(rows: SheetRows,
                       user_aliases: dict[str, str] | None = None) -> int:
    """Score how header-like a sheet is, for cross-sheet comparison.

    Returns the field count of the sheet's best header row, or -1 when no row
    qualifies. Mirrors VBA SheetHeaderScore: same row scan as find_header_row but
    keyed on field_count only (no has_no tie-break) for picking among sheets."""
    best = -1
    for r1 in range(1, 31):
        field_count, has_no, has_ip = _row_header_signature(rows, r1, user_aliases)
        if not has_ip:
            continue
        if field_count < 2 and not has_no:
            continue
        if field_count > best:
            best = field_count
    return best


def find_request_sheet(sheets: list[SheetRows],
                       user_aliases: dict[str, str] | None = None) -> int:
    """Pick the index of the sheet whose best header row scores highest.

    `sheets` is a list of sheets, each a list[list] of cell values (a workbook).
    Ties keep the leftmost (lowest index) sheet, so single-sheet and
    first-sheet-wins forms behave exactly as before. Mirrors VBA FindRequestSheet.
    Raises RequestParseError when NO sheet has a recognizable header row."""
    best_index = -1
    best_score = -1
    for i, rows in enumerate(sheets):
        score = sheet_header_score(rows, user_aliases)
        if score > best_score:
            best_score = score
            best_index = i
    if best_score >= 0 and best_index >= 0:
        return best_index
    raise RequestParseError(
        "\ud5e4\ub354 \ud589\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: \ucd9c\ubc1c\uc9c0IP/\ubaa9\uc801\uc9c0IP \uc5f4\uc774 \uc788\ub294 \uc2dc\ud2b8\uac00 \ud544\uc694\ud569\ub2c8\ub2e4.")


def select_request_sheet(named_sheets: NamedSheetRows,
                         parse_sheet_name: str = "",
                         user_aliases: dict[str, str] | None = None) -> int:
    """Pick the sheet index honoring an explicit `parse_sheet` setting.

    `named_sheets` is a list of (name, rows) tuples (a workbook with sheet
    names). Mirrors VBA SelectRequestSheet:
      - blank parse_sheet_name -> auto-detect via find_request_sheet (leftmost
        tie, highest header score).
      - non-empty -> select the FIRST sheet whose name equals parse_sheet_name
        EXACTLY (no trim/case-fold/alias). Raise if no sheet has that name.
        Raise if the named sheet exists but has no recognizable header row
        (sheet_header_score < 0) -- an explicit choice never silently
        falls back to another sheet."""
    if not str(parse_sheet_name).strip():
        return find_request_sheet([rows for _name, rows in named_sheets], user_aliases)
    for i, (name, rows) in enumerate(named_sheets):
        if name == parse_sheet_name:
            if sheet_header_score(rows, user_aliases) < 0:
                raise RequestParseError(
                    "\ud30c\uc2f1 \ub300\uc0c1 \uc2dc\ud2b8\uc5d0\uc11c \ud5e4\ub354 \ud589\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: " + parse_sheet_name)
            return i
    raise RequestParseError(
        "\ud30c\uc2f1 \ub300\uc0c1 \uc2dc\ud2b8\ub97c \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: " + parse_sheet_name)

def build_header_map(rows: SheetRows, header_row: int,
                     user_aliases: dict[str, str] | None = None) -> dict[str, int]:
    hmap: dict[str, int] = {}
    last_col = _last_column(rows, header_row)
    for c1 in range(1, last_col + 1):
        raw = header_key(_cell(rows, header_row, c1))
        name = canonical_header_name(raw)
        # mirror VBA BuildHeaderMap: built-in first, then user alias fallback
        if name == raw and user_aliases and raw in user_aliases:
            name = user_aliases[raw]
        if name:
            hmap[name] = c1  # later columns overwrite earlier (mirror VBA dict assign)
    return hmap


def validate_required_headers(hmap: dict[str, int]) -> None:
    for req in REQUIRED:
        if req not in hmap:
            raise RequestParseError(f"필수 컬럼 누락: {req}")


def _source_last_row(rows: SheetRows, hmap: dict[str, int]) -> int:
    # mirror SourceLastRow: max of last non-empty row over 출발지ip and 목적지ip
    def last_nonempty(col: int) -> int:
        last = 0
        for r1 in range(1, len(rows) + 1):
            if str(_cell(rows, r1, col)).strip() != "":
                last = r1
        return last
    return max(last_nonempty(hmap["출발지ip"]), last_nonempty(hmap["목적지ip"]))


def _row_has_data(rows: SheetRows, r1: int, hmap: dict[str, int]) -> bool:
    return (str(_cell(rows, r1, hmap["출발지ip"])).strip() != "" or
            str(_cell(rows, r1, hmap["목적지ip"])).strip() != "")


def parse_request_sheet(rows: SheetRows,
                        user_aliases: dict[str, str] | None = None) -> list[RequestRow]:
    """Parse a raw sheet into canonical request dicts.

    Returns one dict per data row with keys mirroring CopyRequestRow:
    source_ip, source_name, dest_ip, dest_name, protocol, port, direction,
    purpose, start_date, end_date, note, source_row.
    Raises RequestParseError on missing header row / required columns.
    user_aliases mirrors settings header_alias (built-in canonical wins first).
    """
    header_row = find_header_row(rows, user_aliases)
    hmap = build_header_map(rows, header_row, user_aliases)
    validate_required_headers(hmap)

    last_row = _source_last_row(rows, hmap)
    out: list[RequestRow] = []
    for r1 in range(header_row + 1, last_row + 1):
        if not _row_has_data(rows, r1, hmap):
            continue
        out.append({
            "source_row": r1,
            "target_firewalls": _norm_list(_opt(rows, r1, hmap, "대상방화벽")),
            "source_ip": _norm_list(_opt(rows, r1, hmap, "출발지ip")),
            "source_name": _norm_text(_opt(rows, r1, hmap, "출발지")),
            "dest_ip": _norm_list(_opt(rows, r1, hmap, "목적지ip")),
            "dest_name": _norm_text(_opt(rows, r1, hmap, "목적지")),
            "protocol": _norm_list(_opt(rows, r1, hmap, "프로토콜")).upper(),
            "port": _norm_list(_opt(rows, r1, hmap, "포트")),
            "direction": _norm_text(_opt(rows, r1, hmap, "방향")),
            "purpose": _norm_text(_opt(rows, r1, hmap, "용도")),
            "start_date": _format_metadata_date(_opt(rows, r1, hmap, "시작일")),
            "end_date": _format_metadata_date(_opt(rows, r1, hmap, "종료일")),
            "note": _norm_text(_opt(rows, r1, hmap, "비고")),
        })
    return out


# --------------------------------------------------------------------------- #
# Cartesian explode: one request row with multi-valued 출발지IP / 목적지IP / 포트
# becomes the FULL product (출발지IP × 목적지IP × 포트), one firewall rule per row.
# Mirrors VBA SplitNormalizedList + the explode loop in CopyRequestRow.
# --------------------------------------------------------------------------- #

def split_list(normalized: str) -> list[str]:
    """Split a ';'-normalized list into tokens. An empty value yields [''] so a
    field with no value still produces exactly one (blank) explode element and
    the source row is never dropped. Mirrors VBA SplitNormalizedList."""
    s = (normalized or "").strip(";")
    if s == "":
        return [""]
    return [tok for tok in s.split(";") if tok != ""] or [""]


def explode_request_row(row: RequestRow) -> list[RequestRow]:
    srcs = split_list(str(row.get("source_ip", "") or ""))
    dsts = split_list(str(row.get("dest_ip", "") or ""))
    ports = split_list(str(row.get("port", "") or ""))
    out: list[RequestRow] = []
    for s in srcs:
        for d in dsts:
            for p in ports:
                r = dict(row)
                r["source_ip"] = s
                r["dest_ip"] = d
                r["port"] = p
                out.append(r)
    return out


def parse_request_sheet_exploded(rows: SheetRows,
                                 user_aliases: dict[str, str] | None = None) -> list[RequestRow]:
    """parse_request_sheet then explode each row into its cartesian product."""
    out: list[RequestRow] = []
    for row in parse_request_sheet(rows, user_aliases):
        out.extend(explode_request_row(row))
    return out
