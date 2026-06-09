"""Binding scenario tests for the firewall route-analysis algorithm.

These tests ARE the specification. vba/FirewallRouteAnalysis.bas must match.
Run: .venv/bin/python -m pytest tests/test_route_oracle.py -v
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(__file__))

from route_oracle import (  # noqa: E402
    Firewall,
    Network,
    RouteEngine,
    RoutingPath,
    cidr_end,
    cidr_start,
    ip_to_number,
    ranges_overlap,
)


# --------------------------------------------------------------------------- #
# Fixtures: a small but representative topology
#
#   zones: RED(10.10.0.0/16) GREEN(10.20.0.0/16) BLUE(10.30.0.0/16)
#          ALT(10.40.0.0/16) ISLAND_A(10.50.0.0/16) ISLAND_B(10.60.0.0/16)
#   firewalls: FW-A FW-B FW-C FW-D (FW-LEGACY enabled for fallback case)
#   edges:
#     RED -> GREEN via FW-A (order 10)
#     GREEN -> BLUE via FW-B (order 20)        (multi-hop RED>GREEN>BLUE)
#     RED -> ALT via FW-C (order 10)           (two equal one-hop paths to ALT)
#     RED -> ALT via FW-D (order 5)            -> tie-break picks FW-D (order 5)
# --------------------------------------------------------------------------- #


@pytest.fixture
def engine():
    networks = [
        Network("net-red", "10.10.0.0/16", "RED"),
        Network("net-green", "10.20.0.0/16", "GREEN"),
        Network("net-blue", "10.30.0.0/16", "BLUE"),
        Network("net-alt", "10.40.0.0/16", "ALT"),
        Network("net-isa", "10.50.0.0/16", "ISLAND_A"),
        Network("net-isb", "10.60.0.0/16", "ISLAND_B"),
        # more specific net inside RED to prove longest-prefix
        Network("net-red-dmz", "10.10.5.0/24", "RED_DMZ"),
    ]
    firewalls = [
        Firewall("FW-A"),
        Firewall("FW-B"),
        Firewall("FW-C"),
        Firewall("FW-D"),
        Firewall("FW-LEGACY"),
        Firewall("FW-OFF", enabled=False),
    ]
    routing = [
        RoutingPath("FW-A", "RED", "GREEN", "p1", "p2", 10),
        RoutingPath("FW-B", "GREEN", "BLUE", "p1", "p2", 20),
        RoutingPath("FW-C", "RED", "ALT", "p1", "p2", 10),
        RoutingPath("FW-D", "RED", "ALT", "p1", "p2", 5),
        # disabled edge must be ignored
        RoutingPath("FW-A", "RED", "BLUE", "p9", "p9", 1, enabled=False),
        # edge for a disabled firewall must be ignored
        RoutingPath("FW-OFF", "RED", "ISLAND_A", "p1", "p2", 1),
    ]
    return RouteEngine(networks=networks, firewalls=firewalls, routing_paths=routing)


# --------------------------------------------------------------------------- #
# Primitive sanity (RED-style anchors for CIDR math)
# --------------------------------------------------------------------------- #

def test_ip_to_number():
    assert ip_to_number("0.0.0.0") == 0
    assert ip_to_number("255.255.255.255") == 4294967295
    assert ip_to_number("10.10.10.5") == 168430085


def test_cidr_bounds():
    assert cidr_start("10.10.0.0/16") == ip_to_number("10.10.0.0")
    assert cidr_end("10.10.0.0/16") == ip_to_number("10.10.255.255")
    # non-aligned base normalizes down
    assert cidr_start("10.10.5.7/24") == ip_to_number("10.10.5.0")


def test_ranges_overlap():
    assert ranges_overlap("10.10.10.5", "10.10.0.0/16") is True
    assert ranges_overlap("10.10.0.0/16", "10.10.10.0/24") is True
    assert ranges_overlap("10.10.10.5", "192.168.0.0/16") is False
    assert ranges_overlap("bad", "10.0.0.0/8") is False


# --------------------------------------------------------------------------- #
# Scenario 1: zone resolve + single hop OK
# --------------------------------------------------------------------------- #

def test_scenario1_ok_single_hop(engine):
    r = engine.analyze("10.10.10.10", "10.20.20.20", "OUT")
    assert r.status == "OK"
    assert r.source_zone == "RED"
    assert r.destination_zone == "GREEN"
    assert r.target_firewalls == "FW-A"
    assert r.firewall_path == "FW-A"
    assert r.zone_path == "RED>GREEN"


# --------------------------------------------------------------------------- #
# Scenario 2: multi-hop path
# --------------------------------------------------------------------------- #

def test_scenario2_multi_hop(engine):
    r = engine.analyze("10.10.10.10", "10.30.30.30", "OUT")
    assert r.status == "OK"
    assert r.zone_path == "RED>GREEN>BLUE"
    assert r.firewall_path == "FW-A>FW-B"
    assert r.target_firewalls == "FW-A;FW-B"


# --------------------------------------------------------------------------- #
# Scenario 3: multi-path tie-break (two one-hop paths RED->ALT)
#   FW-D order=5 beats FW-C order=10 by edge key
# --------------------------------------------------------------------------- #

def test_scenario3_multipath_tiebreak(engine):
    r = engine.analyze("10.10.10.10", "10.40.40.40", "OUT")
    assert r.status == "MULTI_PATH"
    assert r.path_count == 2
    # deterministic winner = lower path_order -> FW-D
    assert r.firewall_path == "FW-D"
    assert r.target_firewalls == "FW-D"
    assert r.zone_path == "RED>ALT"


# --------------------------------------------------------------------------- #
# Scenario 4: intra-zone
# --------------------------------------------------------------------------- #

def test_scenario4_intra_zone(engine):
    r = engine.analyze("10.20.1.1", "10.20.2.2", "OUT")
    assert r.status == "INTRA_ZONE"
    assert r.source_zone == "GREEN"
    assert r.destination_zone == "GREEN"
    assert r.firewall_path == ""
    assert r.target_firewalls == ""


# --------------------------------------------------------------------------- #
# Scenario 5: source zone unresolved
# --------------------------------------------------------------------------- #

def test_scenario5_source_unresolved(engine):
    r = engine.analyze("192.0.2.10", "10.20.20.20", "OUT")
    assert r.status == "ZONE_UNRESOLVED"
    assert "Source zone unresolved" in r.validation_message


# --------------------------------------------------------------------------- #
# Scenario 6: destination zone unresolved
# --------------------------------------------------------------------------- #

def test_scenario6_dest_unresolved(engine):
    r = engine.analyze("10.10.10.10", "192.0.2.20", "OUT")
    assert r.status == "ZONE_UNRESOLVED"
    assert "Destination zone unresolved" in r.validation_message


# --------------------------------------------------------------------------- #
# Scenario 7: no path (both zones resolve, no edge connects, fallback off)
# --------------------------------------------------------------------------- #

def test_scenario7_no_path(engine):
    r = engine.analyze("10.50.50.10", "10.60.60.10", "OUT")
    assert r.status == "NO_PATH"
    assert r.target_firewalls == ""


# --------------------------------------------------------------------------- #
# Scenario 8: direction mismatch (only RED->GREEN exists, ask GREEN->BLUE? no.
#   ask the reverse of an existing edge: GREEN->RED with OUT, reverse exists)
# --------------------------------------------------------------------------- #

def test_scenario8_direction_mismatch(engine):
    r = engine.analyze("10.20.20.20", "10.10.10.10", "OUT")
    assert r.status == "DIRECTION_MISMATCH"
    assert "reverse path exists" in r.validation_message
    assert r.target_firewalls == ""


# --------------------------------------------------------------------------- #
# Scenario 9: blank/BOTH resolves using reverse when forward missing
# --------------------------------------------------------------------------- #

def test_scenario9_both_reverse(engine):
    r = engine.analyze("10.20.20.20", "10.10.10.10", "")  # blank => BOTH
    assert r.status in ("OK", "MULTI_PATH")
    assert r.target_firewalls == "FW-A"
    assert "reverse direction under BOTH" in r.validation_message


# --------------------------------------------------------------------------- #
# Scenario 10: legacy fallback (no path, toggle ON)
# --------------------------------------------------------------------------- #

def test_scenario10_legacy_fallback():
    networks = [
        Network("net-isa", "10.50.0.0/16", "ISLAND_A"),
        Network("net-isb", "10.60.0.0/16", "ISLAND_B"),
    ]
    firewalls = [Firewall("FW-LEGACY")]
    # an edge exists but not connecting ISLAND_A->ISLAND_B; its zones overlap
    # the request IPs so legacy fallback should still flag FW-LEGACY.
    routing = [RoutingPath("FW-LEGACY", "ISLAND_A", "ISLAND_B", "p1", "p2", 10)]
    eng = RouteEngine(
        networks=networks,
        firewalls=firewalls,
        routing_paths=routing,
        fallback_enabled=True,
    )
    # request from ISLAND_B -> ISLAND_A with OUT: forward has no edge
    # (edge is A->B), reverse (A->B) exists so it's DIRECTION_MISMATCH, not NO_PATH.
    # Use a genuine no-path: ISLAND_A -> ISLAND_A is intra; instead drop the
    # edge so there is no graph path at all.
    eng2 = RouteEngine(
        networks=networks,
        firewalls=firewalls,
        routing_paths=[],  # no edges -> NO_PATH -> fallback
        fallback_enabled=True,
    )
    r = eng2.analyze("10.50.50.10", "10.60.60.10", "OUT")
    # with no routing edges at all, fallback also finds nothing -> NO_PATH
    assert r.status == "NO_PATH"

    # now give fallback something to find: an edge whose zones overlap the IPs
    r2 = eng.analyze("10.50.50.10", "10.50.60.10", "OUT")
    # 10.50.x are both ISLAND_A -> intra-zone, not fallback. Use cross IPs:
    r3 = eng.analyze("10.60.0.5", "10.60.0.6", "OUT")
    assert r3.status == "INTRA_ZONE"

    # genuine fallback: zones resolve, no directed path forward or reverse,
    # but fallback edge zones overlap request IPs.
    networks3 = [
        Network("n1", "10.70.0.0/16", "Z1"),
        Network("n2", "10.80.0.0/16", "Z2"),
        Network("n3", "10.90.0.0/16", "Z3"),
    ]
    routing3 = [RoutingPath("FW-LEGACY", "Z1", "Z2", "p1", "p2", 10)]
    eng3 = RouteEngine(
        networks=networks3,
        firewalls=[Firewall("FW-LEGACY")],
        routing_paths=routing3,
        fallback_enabled=True,
    )
    # Z1 -> Z3: no path; fallback checks edges whose zones (Z1/Z2) overlap
    # src(10.70 in Z1) or dst(10.90 in Z3). Z1 overlaps src -> FW-LEGACY.
    r4 = eng3.analyze("10.70.0.5", "10.90.0.5", "OUT")
    assert r4.status == "LEGACY_FALLBACK"
    assert r4.target_firewalls == "FW-LEGACY"


# --------------------------------------------------------------------------- #
# Scenario 11: invalid direction
# --------------------------------------------------------------------------- #

def test_scenario11_invalid_direction(engine):
    r = engine.analyze("10.10.10.10", "10.20.20.20", "SIDEWAYS")
    assert r.status == "DIRECTION_MISMATCH"
    assert "Invalid direction" in r.validation_message


# --------------------------------------------------------------------------- #
# Scenario 12: longest-prefix wins (10.10.5.x is RED_DMZ not RED)
# --------------------------------------------------------------------------- #

def test_scenario12_longest_prefix(engine):
    z = engine.resolve_zone("10.10.5.50")
    assert z == "RED_DMZ"
    z2 = engine.resolve_zone("10.10.9.50")
    assert z2 == "RED"


# --------------------------------------------------------------------------- #
# Scenario 13: disabled firewall / disabled edge ignored
# --------------------------------------------------------------------------- #

def test_scenario13_disabled_ignored(engine):
    # FW-OFF edge RED->ISLAND_A must be ignored => no path
    r = engine.analyze("10.10.10.10", "10.50.50.50", "OUT")
    assert r.status == "NO_PATH"


# --------------------------------------------------------------------------- #
# Scenario 14: request value is a CIDR (대역) -> must resolve zone by overlap
# --------------------------------------------------------------------------- #

def test_scenario14_request_cidr(engine):
    # 10.10.10.0/24 is inside RED (10.10.0.0/16); 10.20.20.0/24 inside GREEN
    z = engine.resolve_zone("10.10.10.0/24")
    assert z == "RED"
    r = engine.analyze("10.10.10.0/24", "10.20.20.0/24", "OUT")
    assert r.status == "OK"
    assert r.source_zone == "RED"
    assert r.destination_zone == "GREEN"
    assert r.target_firewalls == "FW-A"


# --------------------------------------------------------------------------- #
# Scenario 15: request value is an address list -> first resolvable token used
# --------------------------------------------------------------------------- #

def test_scenario15_request_list(engine):
    r = engine.analyze("10.10.10.5;10.10.10.6", "10.20.20.5", "OUT")
    assert r.status == "OK"
    assert r.source_zone == "RED"
    assert r.destination_zone == "GREEN"
    assert r.target_firewalls == "FW-A"


# --------------------------------------------------------------------------- #
# Scenario 16: CIDR exactly equal to a defined network resolves to that zone
# --------------------------------------------------------------------------- #

def test_scenario16_cidr_exact_network(engine):
    # 10.10.5.0/24 is the RED_DMZ network exactly -> longest prefix wins
    z = engine.resolve_zone("10.10.5.0/24")
    assert z == "RED_DMZ"


# --------------------------------------------------------------------------- #
# Scenario 17: list with leading unresolvable token falls back to next token
# --------------------------------------------------------------------------- #

def test_scenario17_list_skips_unresolvable(engine):
    # 192.0.2.9 unresolved, 10.10.10.5 -> RED
    z = engine.resolve_zone("192.0.2.9;10.10.10.5")
    assert z == "RED"


# --------------------------------------------------------------------------- #
# Scenario 18: ambiguous zone (two equal-prefix networks, different zones)
# --------------------------------------------------------------------------- #

def test_scenario18_ambiguous_zone():
    networks = [
        Network("a", "10.10.0.0/16", "ZONE_A"),
        Network("b", "10.10.0.0/16", "ZONE_B"),  # same prefix, different zone
    ]
    eng = RouteEngine(networks=networks, firewalls=[], routing_paths=[])
    assert eng.resolve_zone("10.10.10.10") == "#AMBIGUOUS"
    r = eng.analyze("10.10.10.10", "10.10.20.20", "OUT")
    assert r.status == "ZONE_UNRESOLVED"
    assert "unresolved" in r.validation_message.lower()


# --------------------------------------------------------------------------- #
# Scenario 19: BOTH (blank) resolves via the FORWARD path with no reverse note
# --------------------------------------------------------------------------- #

def test_scenario19_both_forward_success(engine):
    r = engine.analyze("10.10.10.10", "10.20.20.20", "")  # blank => BOTH
    assert r.status == "OK"
    assert r.target_firewalls == "FW-A"
    assert r.zone_path == "RED>GREEN"
    assert "reverse direction under BOTH" not in r.validation_message


# --------------------------------------------------------------------------- #
# Scenario 20: IN direction succeeds by traversing destination -> source
# --------------------------------------------------------------------------- #

def test_scenario20_in_direction_success(engine):
    # graph has RED->GREEN; an IN request (src=GREEN, dst=RED) analyzes dst->src
    r = engine.analyze("10.20.20.20", "10.10.10.10", "IN")
    assert r.status == "OK"
    assert r.source_zone == "GREEN"
    assert r.destination_zone == "RED"
    assert r.target_firewalls == "FW-A"
    assert r.zone_path == "RED>GREEN"


# --------------------------------------------------------------------------- #
# Scenario 21-27: inside/outside CIDR auto-derivation (no routing_paths)
# Each firewall has inside_cidr + outside_cidr; a CIDR shared by two firewalls
# chains them. Request IPs resolve to the cidr:<base/prefix> zone of their range.
# --------------------------------------------------------------------------- #

def _auto_engine(firewalls):
    # no network_definitions needed: the inside/outside CIDRs ARE the zones
    return RouteEngine(networks=[], firewalls=firewalls, routing_paths=[])


def test_scenario21_auto_single_hop():
    eng = _auto_engine([Firewall("FW-1", inside_cidr="10.10.0.0/16",
                                 outside_cidr="172.16.0.0/16")])
    r = eng.analyze("10.10.10.5", "172.16.1.10", "OUT")
    assert r.status == "OK"
    assert r.target_firewalls == "FW-1"
    assert r.zone_path == "cidr:10.10.0.0/16>cidr:172.16.0.0/16"


def test_scenario22_auto_multi_hop_shared_transit():
    # 172.16/16 is FW-01.outside AND FW-02.inside -> one shared transit zone
    # 10.20/16 is FW-02.outside AND FW-03.inside -> chains all three
    eng = _auto_engine([
        Firewall("SECUI-FW-01", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
        Firewall("SECUI-FW-02", inside_cidr="172.16.0.0/16", outside_cidr="10.20.0.0/16"),
        Firewall("SECUI-FW-03", inside_cidr="10.20.0.0/16", outside_cidr="0.0.0.0/0"),
    ])
    r = eng.analyze("10.10.10.5", "8.8.8.8", "OUT")
    assert r.status == "OK"
    assert r.target_firewalls == "SECUI-FW-01;SECUI-FW-02;SECUI-FW-03"
    assert r.firewall_path == "SECUI-FW-01>SECUI-FW-02>SECUI-FW-03"


def test_scenario23_explicit_overrides_auto():
    # explicit routing_paths present -> authoritative, inside/outside ignored
    nets = [Network("n-int", "10.10.0.0/16", "internal"),
            Network("n-srv", "172.16.1.0/24", "server")]
    fws = [Firewall("FW-1", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16")]
    rps = [RoutingPath("FW-1", "internal", "transit", path_order=1)]
    eng = RouteEngine(networks=nets, firewalls=fws, routing_paths=rps)
    r = eng.analyze("10.10.10.5", "172.16.1.10", "OUT")
    assert r.status == "NO_PATH"  # explicit graph authoritative; ignores CIDRs


def test_scenario24_auto_cidr_canonicalized():
    # non-aligned base / bare IP still canonicalize to the same zone
    eng = _auto_engine([Firewall("FW-1", inside_cidr="10.10.5.7/16",
                                 outside_cidr="172.16.0.0/16")])
    r = eng.analyze("10.10.99.1", "172.16.1.10", "OUT")
    assert r.status == "OK"
    assert r.target_firewalls == "FW-1"
    assert r.zone_path.startswith("cidr:10.10.0.0/16>")


def test_scenario25_auto_one_side_only_no_path():
    eng = _auto_engine([Firewall("FW-1", inside_cidr="10.10.0.0/16", outside_cidr="")])
    r = eng.analyze("10.10.10.5", "172.16.1.10", "OUT")
    assert r.status in ("NO_PATH", "ZONE_UNRESOLVED")  # no outside -> no edge


def test_scenario26_auto_equal_paths_multipath():
    # two firewalls with identical inside/outside -> two equal one-hop paths
    eng = _auto_engine([
        Firewall("FW-A", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
        Firewall("FW-B", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
    ])
    r = eng.analyze("10.10.10.5", "172.16.1.10", "OUT")
    assert r.status == "MULTI_PATH"
    assert r.path_count == 2
    assert r.target_firewalls == "FW-A"  # winner = lower path_order (row 1)


def test_scenario27_auto_disabled_firewall_consumes_ordinal():
    """A disabled firewall consumes a path_order ordinal but produces no edges."""
    fws = [
        Firewall("FW-OFF", enabled=False, inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
        Firewall("FW-A", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
        Firewall("FW-B", inside_cidr="10.10.0.0/16", outside_cidr="172.16.0.0/16"),
    ]
    eng = RouteEngine(networks=[], firewalls=fws, routing_paths=[])
    r = eng.analyze("10.10.10.5", "172.16.1.10", "OUT")
    assert r.status == "MULTI_PATH"
    assert r.path_count == 2  # FW-OFF contributes nothing
    assert r.target_firewalls == "FW-A"  # ordinal 2 < FW-B ordinal 3
