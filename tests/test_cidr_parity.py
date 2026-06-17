"""Parity contract for the CIDR overlap engine.

This is the test that was MISSING and allowed three overlap implementations to
drift:

  - tests/route_oracle.py        ``ranges_overlap``      (route algorithm contract)
  - scripts/secui_cli_seed.py    SECUI match path        (CLI rule generator)
  - vba/FirewallRouteAnalysis.bas ``RangesOverlap``      (shipped Excel macro)

They MUST agree bit-for-bit. The authoritative behavior is strict IPv4 CIDR
(the shipped VBA macro can only do 32-bit math). IPv6, leading-zero octets and
malformed addresses are INVALID, not silently "no overlap" -- a firewall-rule
generator must never emit a rule for traffic it cannot verify, and the route
analyzer must never silently drop it either.

The single source of truth is ``firewall_policy.cidr``. Both the route oracle
and the SECUI seed import it, so they cannot drift again.
"""

from __future__ import annotations

import pytest

from firewall_policy import cidr
from scripts.secui_cli_seed import _network_overlaps as cli_overlaps
from tests.route_oracle import ranges_overlap as oracle_overlaps

# (left, right, expected_overlap) under the strict-IPv4 contract.
# Invalid inputs (IPv6 / leading-zero / malformed) overlap to False because the
# value is not a parseable IPv4 network -- but they must ALSO be classified
# invalid (see test_invalid_classification), which is what stops silent routing.
VALID_IPV4_CASES = [
    ("10.10.10.5", "10.10.0.0/16", True),
    ("10.10.0.0/16", "10.10.10.0/24", True),
    ("10.10.10.5", "192.168.0.0/16", False),
    ("10.0.0.5/24", "10.0.0.99", True),  # host bits set, strict=False
    ("0.0.0.0/0", "8.8.8.8", True),
    ("8.8.8.8", "8.8.4.4", False),
]

INVALID_INPUTS = [
    "2001:db8::/32",       # IPv6 network
    "2001:db8::1",         # IPv6 host
    "::1",                 # IPv6 loopback
    "010.000.000.001",     # leading-zero octets (ambiguous, banned)
    "10.0.0.256",          # octet > 255
    "10.0.0",              # too few octets
    "10.0.0.0.0",          # too many octets
    "10.0.0.1/33",         # prefix out of range
    "10.0.0.1/24/garbage", # extra slash segment (VBA used to accept this)
    "10.0.0.1/+24",        # non-digit prefix
    "10.0.0.-1",           # negative octet
    "garbage",             # not an address at all
]


@pytest.mark.parametrize("left,right,expected", VALID_IPV4_CASES)
def test_oracle_and_cli_overlap_agree_on_valid_ipv4(left, right, expected):
    assert oracle_overlaps(left, right) is expected
    assert cli_overlaps(left, right) is expected
    assert cidr.ipv4_ranges_overlap(left, right) is expected


@pytest.mark.parametrize("bad", INVALID_INPUTS)
def test_invalid_inputs_never_overlap_in_any_engine(bad):
    # An invalid value must NOT silently overlap a real network in EITHER engine.
    # This is the security-critical parity: before the fix, the CLI used
    # ipaddress.ip_network and matched IPv6, while the oracle/VBA did not.
    good = "10.0.0.0/8"
    assert oracle_overlaps(bad, good) is False
    assert cli_overlaps(bad, good) is False
    assert oracle_overlaps(good, bad) is False
    assert cli_overlaps(good, bad) is False


@pytest.mark.parametrize("bad", INVALID_INPUTS)
def test_invalid_inputs_are_classified_invalid(bad):
    # "invalid" is materially different from "no match": the shared module must
    # report it so the route analyzer can emit INVALID_ADDRESS and the CLI can
    # skip the row instead of silently producing/omitting a rule.
    assert cidr.parse_ipv4_network(bad) is None
    assert cidr.is_invalid_address(bad) is True


@pytest.mark.parametrize("good", ["10.0.0.1", "10.0.0.0/8", "0.0.0.0/0", "0.0.0.0"])
def test_valid_ipv4_is_not_classified_invalid(good):
    assert cidr.parse_ipv4_network(good) is not None
    assert cidr.is_invalid_address(good) is False


def test_any_tokens_are_not_invalid():
    # ANY/blank are list-level sentinels handled above the overlap layer, never
    # parsed as addresses, so they must not be flagged invalid.
    for token in ("ANY", "ALL", "*", "", "0.0.0.0/0"):
        assert cidr.is_invalid_address(token) is False


def test_both_engines_delegate_to_shared_module():
    # Guard against re-drift: the SECUI seed must route overlap through
    # firewall_policy.cidr, not ipaddress.ip_network directly.
    import inspect

    import scripts.secui_cli_seed as seed

    src = inspect.getsource(seed)
    assert "ip_network(" not in src, "secui_cli_seed must not call ip_network directly"
    assert "firewall_policy" in src and "cidr" in src
