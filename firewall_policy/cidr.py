"""Single source of truth for IPv4 CIDR parsing and overlap.

The route oracle (``tests/route_oracle.py``), the SECUI CLI generator
(``scripts/secui_cli_seed.py``) and the shipped Excel macro
(``vba/FirewallRouteAnalysis.bas``) all compute "do these two address ranges
overlap?". Historically each had its own implementation and they drifted: the
CLI used ``ipaddress.ip_network`` (which understands IPv6 and rejects
leading-zero octets) while the oracle/VBA used hand-rolled 32-bit IPv4 math
(which silently accepted leading zeros and could not see IPv6). The result was
that the CLI could emit a firewall rule for IPv6 traffic the route analyzer
declared unroutable -- the worst failure class for a firewall-rule generator.

This module is the one authoritative implementation. Both Python callers import
it; the VBA macro mirrors its rules. The contract is **strict IPv4 CIDR only**,
because the shipped macro can only do 32-bit arithmetic and partial IPv6 support
is exactly the unsafe divergence we are eliminating.

Strict IPv4 CIDR contract
-------------------------
Accepted:
  - ``A.B.C.D``            -> host, treated as ``/32``
  - ``A.B.C.D/N``          -> network, ``N`` in ``0..32`` (host bits ignored)
Rejected as INVALID (``parse_ipv4_network`` returns ``None``):
  - IPv6 in any form
  - a non-numeric octet
  - an octet outside ``0..255``
  - an empty octet, fewer/more than 4 octets
  - a negative octet
  - a leading-zero octet (e.g. ``010``); the literal octet ``"0"`` is fine
  - a prefix outside ``0..32`` or non-numeric

"Invalid" is deliberately distinct from "no overlap": a malformed value cannot
be reasoned about safely, so callers surface it (route status
``INVALID_ADDRESS``; the CLI skips the row) instead of silently treating it as a
non-match.
"""

from __future__ import annotations

from dataclasses import dataclass

# List-level sentinels meaning "everything". These are handled ABOVE the
# single-token overlap layer (see ``is_any_cidr`` in the callers); they are not
# IPv4 networks and must never be flagged invalid.
_ANY_TOKENS = frozenset({"", "*", "ANY", "ALL", "0.0.0.0/0"})


@dataclass(frozen=True)
class ParsedNetwork:
    """A validated IPv4 network as an inclusive ``[start, end]`` address range."""

    start: int
    end: int


def _parse_octet(text: str) -> int | None:
    clean = text.strip()
    if not clean or not clean.isdigit():
        return None
    # Reject leading-zero octets (ambiguous across tools, banned by modern
    # ipaddress) but allow the literal "0".
    if len(clean) > 1 and clean[0] == "0":
        return None
    value = int(clean)
    if value < 0 or value > 255:
        return None
    return value


def _ip_to_number(ip_text: str) -> int | None:
    parts = ip_text.strip().split(".")
    if len(parts) != 4:
        return None
    value = 0
    for part in parts:
        octet = _parse_octet(part)
        if octet is None:
            return None
        value = value * 256 + octet
    return value


def _prefix_length(cidr_text: str) -> int | None:
    if "/" not in cidr_text:
        return 32
    raw = cidr_text.split("/", 1)[1].strip()
    if not raw or not raw.isdigit():
        return None
    prefix = int(raw)
    if prefix < 0 or prefix > 32:
        return None
    return prefix


def cidr_base_ip(cidr_text: str) -> str:
    """The address part of a CIDR/host token (before any '/')."""
    return str(cidr_text).split("/", 1)[0].strip()


def cidr_prefix_length(cidr_text: str) -> int | None:
    """Public strict prefix parser: returns None for a malformed/out-of-range
    prefix or extra slash segments. Callers that need an exception (the route
    oracle) raise on None."""
    return _prefix_length(str(cidr_text))


def parse_ipv4_network(value: str | None) -> ParsedNetwork | None:
    """Parse a strict IPv4 address or CIDR into an inclusive address range.

    Returns ``None`` for any value that violates the strict-IPv4 contract
    (including IPv6, leading-zero octets and malformed input). Host bits are
    ignored, mirroring ``ipaddress.ip_network(..., strict=False)`` for the
    accepted IPv4 subset.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text or "/" not in text and text.count(".") != 3:
        # fast reject of obvious non-IPv4 (e.g. IPv6 "::1", bare words); the
        # detailed checks below still run for "A.B.C.D[/N]"-shaped input.
        if "/" not in text:
            return None
    base = text.split("/", 1)[0]
    prefix = _prefix_length(text)
    if prefix is None:
        return None
    base_num = _ip_to_number(base)
    if base_num is None:
        return None
    block = 1 << (32 - prefix)
    start = (base_num // block) * block
    return ParsedNetwork(start=start, end=start + block - 1)


def is_invalid_address(value: str | None) -> bool:
    """True when ``value`` is a concrete address token that fails to parse.

    ANY/blank sentinels are NOT invalid -- they are list-level wildcards handled
    above this layer. Only a non-sentinel token that ``parse_ipv4_network``
    rejects is invalid.
    """
    if value is None:
        return False
    text = str(value).strip()
    if text.upper() in _ANY_TOKENS:
        return False
    return parse_ipv4_network(text) is None


def ipv4_ranges_overlap(left: str | None, right: str | None) -> bool:
    """Do two single IPv4 address/CIDR tokens overlap?

    ANY sentinels overlap anything. Any token that is not a parseable strict
    IPv4 network yields ``False`` (no overlap); callers detect invalidity
    separately via :func:`is_invalid_address` and surface it.
    """
    if left is None or right is None:
        return False
    if str(left).strip().upper() in _ANY_TOKENS or str(right).strip().upper() in _ANY_TOKENS:
        return True
    left_net = parse_ipv4_network(left)
    right_net = parse_ipv4_network(right)
    if left_net is None or right_net is None:
        return False
    return left_net.start <= right_net.end and right_net.start <= left_net.end
