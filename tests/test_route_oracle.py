import pytest

from tests.route_oracle import (
    Firewall,
    FirewallRange,
    RouteEngine,
    cidr_end,
    cidr_start,
    ip_to_number,
    ranges_overlap,
    split_address_list,
)


@pytest.fixture
def engine():
    firewalls = [
        Firewall("SECUI-FW-01", vendor="SECUI"),
        Firewall("SECUI-FW-02", vendor="SECUI"),
        Firewall("SECUI-FW-03", vendor="SECUI"),
        Firewall("FW-OFF", enabled=False),
    ]
    ranges = [
        FirewallRange("SECUI-FW-01", "10.10.0.0/16", "172.16.0.0/16", "OUT", 10),
        FirewallRange("SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10),
        FirewallRange("SECUI-FW-02", "10.10.0.0/16", "10.20.0.0/16", "OUT", 20),
        FirewallRange("SECUI-FW-01", "10.10.0.0/16", "8.8.8.0/24", "OUT", 10),
        FirewallRange("SECUI-FW-02", "10.10.0.0/16", "8.8.8.0/24", "OUT", 20),
        FirewallRange("SECUI-FW-03", "10.10.0.0/16", "8.8.8.0/24", "OUT", 30),
        FirewallRange("FW-OFF", "10.10.0.0/16", "10.30.0.0/16", "OUT", 1),
        FirewallRange("SECUI-FW-01", "10.10.0.0/16", "10.40.0.0/16", "OUT", 40, False),
    ]
    return RouteEngine(firewalls=firewalls, firewall_ranges=ranges)


def test_ip_to_number():
    assert ip_to_number("0.0.0.0") == 0
    assert ip_to_number("255.255.255.255") == 4294967295
    assert ip_to_number("10.10.10.5") == 168430085


def test_cidr_bounds():
    assert cidr_start("10.10.0.0/16") == ip_to_number("10.10.0.0")
    assert cidr_end("10.10.0.0/16") == ip_to_number("10.10.255.255")
    assert cidr_start("10.10.5.7/24") == ip_to_number("10.10.5.0")


def test_ranges_overlap():
    assert ranges_overlap("10.10.10.5", "10.10.0.0/16") is True
    assert ranges_overlap("10.10.0.0/16", "10.10.10.0/24") is True
    assert ranges_overlap("10.10.10.5", "192.168.0.0/16") is False
    assert ranges_overlap("bad", "10.0.0.0/8") is False


def test_split_address_list_space_separated():
    assert split_address_list("10.10.0.0/16 172.16.0.0/16") == [
        "10.10.0.0/16",
        "172.16.0.0/16",
    ]


def test_single_firewall_match(engine):
    result = engine.analyze("10.10.10.10", "172.16.1.10", "OUT")

    assert result.status == "OK"
    assert result.target_firewalls == "SECUI-FW-01"
    assert result.firewall_path == "SECUI-FW-01"
    assert result.source_zone == "10.10.0.0/16"
    assert result.destination_zone == "172.16.0.0/16"
    assert result.zone_path == "10.10.0.0/16>172.16.0.0/16"


def test_multiple_matching_firewalls_keep_user_order(engine):
    result = engine.analyze("10.10.10.10", "8.8.8.8", "OUT")

    assert result.status == "OK"
    assert result.path_count == 3
    assert result.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"
    assert result.firewall_path == "SECUI-FW-01>SECUI-FW-02>SECUI-FW-03"


def test_cidr_request_matches_by_overlap(engine):
    result = engine.analyze("10.10.10.0/24", "10.20.20.0/24", "OUT")

    assert result.status == "OK"
    assert result.target_firewalls == "SECUI-FW-01;SECUI-FW-02"


def test_blank_direction_matches_enabled_definitions(engine):
    result = engine.analyze("10.10.10.10", "172.16.1.10", "")

    assert result.status == "OK"
    assert result.target_firewalls == "SECUI-FW-01"


def test_direction_mismatch_when_reverse_definition_exists():
    eng = RouteEngine(
        firewalls=[Firewall("FW-IN")],
        firewall_ranges=[FirewallRange("FW-IN", "192.0.2.0/24", "10.10.0.0/16", "IN", 1)],
    )

    result = eng.analyze("10.10.10.10", "192.0.2.10", "OUT")

    assert result.status == "DIRECTION_MISMATCH"
    assert "opposite direction" in result.validation_message
    assert result.target_firewalls == ""


def test_no_match_when_no_definition_covers_flow(engine):
    result = engine.analyze("192.0.2.10", "198.51.100.10", "OUT")

    assert result.status == "NO_MATCH"
    assert result.target_firewalls == ""


def test_disabled_firewall_and_disabled_range_are_ignored(engine):
    disabled_fw = engine.analyze("10.10.10.10", "10.30.1.10", "OUT")
    disabled_range = engine.analyze("10.10.10.10", "10.40.1.10", "OUT")

    assert disabled_fw.status == "NO_MATCH"
    assert disabled_range.status == "NO_MATCH"


def test_any_destination_definition():
    eng = RouteEngine(
        firewalls=[Firewall("FW-ANY")],
        firewall_ranges=[FirewallRange("FW-ANY", "10.10.0.0/16", "ANY", "BOTH", 1)],
    )

    result = eng.analyze("10.10.10.10", "203.0.113.10", "OUT")

    assert result.status == "OK"
    assert result.target_firewalls == "FW-ANY"


def test_invalid_direction(engine):
    result = engine.analyze("10.10.10.10", "172.16.1.10", "SIDEWAYS")

    assert result.status == "DIRECTION_MISMATCH"
    assert "Invalid direction" in result.validation_message

# ---------------------------------------------------------------------------
# Direction synonym normalization spec
# ---------------------------------------------------------------------------
# normalize_direction must accept Korean / business direction labels, not just
# IN/OUT/BOTH. These tests are expected to FAIL RED until the synonym mapping is
# implemented in tests/route_oracle.py. Regression cases (IN/OUT/BOTH canonical,
# garbage) must continue to pass.


IN_DIRECTION_SYNONYMS = [
    "IN",
    "in",
    "INBOUND",
    "inbound",
    "인바운드",
    "수신",
    "외부->내부",
    "외부→내부",
    "외부-내부",
    "outside->inside",
    "external->internal",
]

OUT_DIRECTION_SYNONYMS = [
    "OUT",
    "out",
    "OUTBOUND",
    "outbound",
    "아웃바운드",
    "송신",
    "내부->외부",
    "내부→외부",
    "내부-외부",
    "inside->outside",
    "internal->external",
]

BOTH_DIRECTION_SYNONYMS = [
    "",  # blank
    "BOTH",
    "ANY",
    "ALL",
    "양방향",
    "양방",
    "쌍방향",
    "bidirectional",
    "bi-directional",
]

INVALID_DIRECTION_SYNONYMS = [
    "asdf",
    "좌우",
    "내부",  # standalone, not part of a directional phrase
    "외부",  # standalone, not part of a directional phrase
]


@pytest.mark.parametrize("text", IN_DIRECTION_SYNONYMS)
def test_normalize_direction_in_synonyms(text):
    assert RouteEngine.normalize_direction(text) == "IN"


@pytest.mark.parametrize("text", OUT_DIRECTION_SYNONYMS)
def test_normalize_direction_out_synonyms(text):
    assert RouteEngine.normalize_direction(text) == "OUT"


@pytest.mark.parametrize("text", BOTH_DIRECTION_SYNONYMS)
def test_normalize_direction_both_synonyms(text):
    assert RouteEngine.normalize_direction(text) == "BOTH"


@pytest.mark.parametrize("text", INVALID_DIRECTION_SYNONYMS)
def test_normalize_direction_garbage_remains_invalid(text):
    assert RouteEngine.normalize_direction(text) == "#INVALID"


def test_analyze_outbound_korean_resolves_against_out_ranges(engine):
    # Seeded engine has (10.10.0.0/16 -> 10.20.0.0/16) OUT for SECUI-FW-01.
    result = engine.analyze("10.10.10.5", "10.20.20.5", "아웃바운드")

    assert result.status == "OK"
    assert "SECUI-FW-01" in result.target_firewalls.split(";")


def test_analyze_inbound_korean_against_out_only_definition_is_direction_mismatch():
    # Only an OUT definition exists for this flow; an inbound (인바운드) request over
    # the same pair must be flagged DIRECTION_MISMATCH, not silently OK.
    out_only_engine = RouteEngine(
        firewalls=[Firewall("SECUI-FW-01", vendor="SECUI")],
        firewall_ranges=[
            FirewallRange("SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10),
        ],
    )

    result = out_only_engine.analyze("10.20.20.5", "10.10.10.5", "인바운드")

    assert result.status == "DIRECTION_MISMATCH"


def test_analyze_inbound_korean_resolves_against_in_range():
    # Seeded engine has no IN range, so use a small dedicated engine: traffic from
    # external 10.20 to internal 10.10, IN direction, must match SECUI-FW-01.
    inbound_engine = RouteEngine(
        firewalls=[Firewall("SECUI-FW-01", vendor="SECUI")],
        firewall_ranges=[
            FirewallRange("SECUI-FW-01", "10.20.0.0/16", "10.10.0.0/16", "IN", 10),
        ],
    )

    result = inbound_engine.analyze("10.20.20.5", "10.10.10.5", "인바운드")

    assert result.status == "OK"
    assert "SECUI-FW-01" in result.target_firewalls.split(";")
