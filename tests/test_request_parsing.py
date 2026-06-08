"""End-to-end 신청서 엑셀 파싱 검증.

Builds REAL .xlsx request files with varied layouts (column order, alias
headers, No on column B, header not on row 1), parses them with the VBA-mirror
parser, then routes each parsed row through the route engine — proving the full
pipeline (엑셀 파싱 -> zone 해석 -> 경로 산정) works on realistic forms.

Run: .venv/bin/python -m pytest tests/test_request_parsing.py -v
"""

import os
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.dirname(__file__))

from request_parser_oracle import (  # noqa: E402
    RequestParseError,
    canonical_header_name,
    find_header_row,
    header_key,
    parse_request_sheet,
)
from route_oracle import (  # noqa: E402
    Firewall,
    Network,
    RouteEngine,
    RoutingPath,
)


# --------------------------------------------------------------------------- #
# Shared topology for routing the parsed rows (matches shipped seed semantics)
# --------------------------------------------------------------------------- #

@pytest.fixture
def engine():
    nets = [
        Network("업무PC망", "10.10.0.0/16", "internal"),
        Network("서버망", "172.16.1.0/24", "server"),
        Network("중간망", "10.30.0.0/16", "transit"),
        Network("DMZ망", "10.20.0.0/16", "dmz"),
        Network("외부", "0.0.0.0/0", "outside"),
    ]
    fws = [Firewall("SECUI-FW-01"), Firewall("SECUI-FW-02"), Firewall("SECUI-FW-03")]
    rps = [
        RoutingPath("SECUI-FW-01", "internal", "server", "eth1", "eth2", 10),
        RoutingPath("SECUI-FW-01", "internal", "transit", "eth1", "eth3", 20),
        RoutingPath("SECUI-FW-02", "transit", "dmz", "eth1", "eth2", 30),
        RoutingPath("SECUI-FW-03", "dmz", "outside", "eth1", "eth2", 40),
    ]
    return RouteEngine(networks=nets, firewalls=fws, routing_paths=rps)


def _sheet_rows(ws):
    """Read an openpyxl worksheet into list[list] (the parser's input shape)."""
    return [[c.value for c in row] for row in ws.iter_rows()]


def _build_xlsx(tmp_path, name, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    p = tmp_path / name
    wb.save(p)
    return p


# --------------------------------------------------------------------------- #
# Unit: header normalization + alias + header-row detection
# --------------------------------------------------------------------------- #

def test_header_key_normalizes():
    assert header_key(" 출발지 IP ") == "출발지ip"
    assert header_key("Src IP") == "srcip"
    assert header_key(None) == ""


def test_canonical_aliases():
    assert canonical_header_name("srcip") == "출발지ip"
    assert canonical_header_name("출발ip") == "출발지ip"
    assert canonical_header_name("dstip") == "목적지ip"
    assert canonical_header_name("protocol") == "프로토콜"
    assert canonical_header_name("목적지포트") == "포트"
    assert canonical_header_name("unknowncol") == "unknowncol"


def test_find_header_row_scans_top_30():
    rows = [
        ["방화벽 정책 신청서"],          # row 1 title
        [],                              # row 2 blank
        ["No", "출발지IP", "목적지IP"],   # row 3 header
    ]
    assert find_header_row(rows) == 3


def test_find_header_row_missing_raises():
    rows = [["출발지IP", "목적지IP"]]  # no No/번호 anywhere
    with pytest.raises(RequestParseError):
        find_header_row(rows)


# --------------------------------------------------------------------------- #
# E2E #1: canonical layout, header on row 1
# --------------------------------------------------------------------------- #

def test_parse_canonical_layout(tmp_path, engine):
    rows = [
        ["No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
         "방향", "용도", "시작일", "종료일", "비고"],
        [1, "10.10.10.5", "업무PC", "172.16.1.10", "서버", "TCP", "443",
         "OUT", "업무", "2026-01-01", "2026-12-31", "정기"],
    ]
    p = _build_xlsx(tmp_path, "req1.xlsx", rows)
    parsed = parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
    assert len(parsed) == 1
    r = parsed[0]
    assert r["source_ip"] == "10.10.10.5"
    assert r["dest_ip"] == "172.16.1.10"
    assert r["protocol"] == "TCP"
    res = engine.analyze(r["source_ip"], r["dest_ip"], r["direction"])
    assert res.status == "OK"
    assert res.target_firewalls == "SECUI-FW-01"


# --------------------------------------------------------------------------- #
# E2E #2: alias headers + different column order + No on column B + title rows
# --------------------------------------------------------------------------- #

def test_parse_varied_layout(tmp_path, engine):
    rows = [
        ["방화벽 정책 신청서", None, None, None, None],       # row 1 title
        ["작성자: 홍길동", None, None, None, None],            # row 2 meta
        # row 3 header: No in column B, alias names, shuffled order
        [None, "No", "Dst IP", "Src IP", "Protocol", "Port", "Direction",
         "Usage", "Start", "End", "Note", "출발지", "목적지"],
        [None, 1, "10.20.20.5", "10.10.10.5", "tcp", "443", "OUT",
         "DMZ 연동", "2026-01-01", "2026-12-31", "비고1", "PC", "DMZ"],
        [None, 2, "172.16.1.10", "10.10.10.9", "udp", "53", "OUT",
         "DNS", "2026-01-01", "2026-12-31", "비고2", "PC", "DNS"],
    ]
    p = _build_xlsx(tmp_path, "req2.xlsx", rows)
    parsed = parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
    assert len(parsed) == 2

    # row 1: 10.10.10.5 (internal) -> 10.20.20.5 (dmz) => FW-01;FW-02
    res1 = engine.analyze(parsed[0]["source_ip"], parsed[0]["dest_ip"], parsed[0]["direction"])
    assert res1.status == "OK"
    assert res1.target_firewalls == "SECUI-FW-01;SECUI-FW-02"
    assert res1.zone_path == "internal>transit>dmz"

    # row 2: 10.10.10.9 (internal) -> 172.16.1.10 (server) => FW-01
    res2 = engine.analyze(parsed[1]["source_ip"], parsed[1]["dest_ip"], parsed[1]["direction"])
    assert res2.status == "OK"
    assert res2.target_firewalls == "SECUI-FW-01"
    # protocol uppercased by parser
    assert parsed[0]["protocol"] == "TCP"
    assert parsed[1]["protocol"] == "UDP"


# --------------------------------------------------------------------------- #
# E2E #3: CIDR + address-list request values survive parsing and route
# --------------------------------------------------------------------------- #

def test_parse_cidr_and_list(tmp_path, engine):
    rows = [
        ["No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
         "방향", "용도", "시작일", "종료일", "비고"],
        [1, "10.10.10.0/24", "PC대역", "10.20.20.0/24", "DMZ대역", "TCP", "443",
         "OUT", "대역 신청", "2026-01-01", "2026-12-31", ""],
        [2, "10.10.10.5;10.10.10.6", "PC들", "172.16.1.10", "서버", "TCP", "22",
         "OUT", "리스트 신청", "2026-01-01", "2026-12-31", ""],
    ]
    p = _build_xlsx(tmp_path, "req3.xlsx", rows)
    parsed = parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
    assert len(parsed) == 2

    r1 = engine.analyze(parsed[0]["source_ip"], parsed[0]["dest_ip"], parsed[0]["direction"])
    assert r1.status == "OK"
    assert r1.target_firewalls == "SECUI-FW-01;SECUI-FW-02"

    r2 = engine.analyze(parsed[1]["source_ip"], parsed[1]["dest_ip"], parsed[1]["direction"])
    assert r2.status == "OK"
    assert r2.target_firewalls == "SECUI-FW-01"


# --------------------------------------------------------------------------- #
# E2E #4: missing required column raises (matches ValidateRequiredHeaders)
# --------------------------------------------------------------------------- #

def test_parse_missing_required_column(tmp_path):
    rows = [
        # 비고 column omitted
        ["No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
         "방향", "용도", "시작일", "종료일"],
        [1, "10.10.10.5", "PC", "172.16.1.10", "서버", "TCP", "443",
         "OUT", "업무", "2026-01-01", "2026-12-31"],
    ]
    p = _build_xlsx(tmp_path, "req4.xlsx", rows)
    with pytest.raises(RequestParseError) as exc:
        parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
    assert "비고" in str(exc.value)


# --------------------------------------------------------------------------- #
# E2E #5: blank data rows skipped; trailing blanks ignored
# --------------------------------------------------------------------------- #

def test_parse_skips_blank_rows(tmp_path, engine):
    rows = [
        ["No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
         "방향", "용도", "시작일", "종료일", "비고"],
        [1, "10.10.10.5", "PC", "172.16.1.10", "서버", "TCP", "443",
         "OUT", "업무", "2026-01-01", "2026-12-31", ""],
        [None, None, None, None, None, None, None, None, None, None, None, None],
        [3, "", "PC", "", "서버", "TCP", "443", "OUT", "빈IP", "", "", ""],
    ]
    p = _build_xlsx(tmp_path, "req5.xlsx", rows)
    parsed = parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
    # only the one row with IP data is kept (row with empty src+dst IP skipped)
    assert len(parsed) == 1
    assert parsed[0]["source_ip"] == "10.10.10.5"


# --------------------------------------------------------------------------- #
# E2E #6: 비표준 헤더 + settings header_alias -> 파싱 -> 경로추적 (통합)
# --------------------------------------------------------------------------- #

def test_parse_with_user_aliases_then_route(tmp_path, engine):
    from user_alias_oracle import parse_user_aliases

    # settings header_alias: 비표준 헤더를 표준 컬럼으로 매핑
    user_aliases = parse_user_aliases(
        "출발지IP=출발지주소; 목적지IP=목적지주소; 출발지=출발지설명; "
        "목적지=목적지설명; 용도=신청사유"
    )
    # 신청서: 표준 별칭으로는 안 잡히는 헤더(출발지주소/목적지주소 등)를
    # settings header_alias 로 해석해야 파싱 성공.
    rows = [
        ["No", "출발지주소", "출발지설명", "목적지주소", "목적지설명",
         "프로토콜", "포트", "방향", "신청사유", "시작일", "종료일", "비고"],
        [1, "10.10.10.5", "PC", "10.20.20.5", "DMZ", "TCP", "443",
         "OUT", "DMZ 연동", "2026-01-01", "2026-12-31", "비고"],
    ]
    p = _build_xlsx(tmp_path, "req6.xlsx", rows)
    parsed = parse_request_sheet(
        _sheet_rows(openpyxl.load_workbook(p).active),
        user_aliases=user_aliases,
    )
    assert len(parsed) == 1
    r = parsed[0]
    assert r["source_ip"] == "10.10.10.5"
    assert r["dest_ip"] == "10.20.20.5"
    assert r["purpose"] == "DMZ 연동"
    # 경로추적: internal -> transit -> dmz => 두 방화벽
    res = engine.analyze(r["source_ip"], r["dest_ip"], r["direction"])
    assert res.status == "OK"
    assert res.target_firewalls == "SECUI-FW-01;SECUI-FW-02"
    assert res.zone_path == "internal>transit>dmz"


def test_user_alias_missing_without_setting_raises(tmp_path):
    # 사용자 별칭을 등록하지 않으면 비표준 헤더는 필수컬럼 누락으로 실패
    rows = [
        ["No", "출발지주소", "출발지설명", "목적지주소", "목적지설명",
         "프로토콜", "포트", "방향", "신청사유", "시작일", "종료일", "비고"],
        [1, "10.10.10.5", "PC", "10.20.20.5", "DMZ", "TCP", "443",
         "OUT", "연동", "2026-01-01", "2026-12-31", ""],
    ]
    p = _build_xlsx(tmp_path, "req7.xlsx", rows)
    with pytest.raises(RequestParseError):
        parse_request_sheet(_sheet_rows(openpyxl.load_workbook(p).active))
