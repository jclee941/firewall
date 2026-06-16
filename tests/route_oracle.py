from __future__ import annotations

from dataclasses import dataclass, field


def ip_to_number(ip_text: str) -> int:
    parts = ip_text.strip().split(".")
    if len(parts) != 4:
        raise ValueError(f"Invalid IPv4 address: {ip_text}")
    value = 0
    for part in parts:
        clean = part.strip()
        if not clean.isdigit():
            raise ValueError(f"Invalid IPv4 address: {ip_text}")
        octet = int(clean)
        if octet < 0 or octet > 255:
            raise ValueError(f"Invalid IPv4 address: {ip_text}")
        value = value * 256 + octet
    return value


def cidr_prefix_length(cidr_text: str) -> int:
    if "/" not in cidr_text:
        return 32
    prefix = int(cidr_text.split("/", 1)[1].strip())
    if prefix < 0 or prefix > 32:
        raise ValueError(f"Invalid CIDR prefix: {cidr_text}")
    return prefix


def cidr_base_ip(cidr_text: str) -> str:
    return cidr_text.split("/", 1)[0].strip()


def cidr_block_size(cidr_text: str) -> int:
    return 2 ** (32 - cidr_prefix_length(cidr_text))


def cidr_start(cidr_text: str) -> int:
    base = ip_to_number(cidr_base_ip(cidr_text))
    block = cidr_block_size(cidr_text)
    return (base // block) * block


def cidr_end(cidr_text: str) -> int:
    return cidr_start(cidr_text) + cidr_block_size(cidr_text) - 1


def ranges_overlap(left_cidr: str, right_cidr: str) -> bool:
    try:
        if is_any_cidr(left_cidr) or is_any_cidr(right_cidr):
            return True
        if not left_cidr or not right_cidr:
            return False
        left_start, left_end = cidr_start(left_cidr), cidr_end(left_cidr)
        right_start, right_end = cidr_start(right_cidr), cidr_end(right_cidr)
        return left_start <= right_end and right_start <= left_end
    except ValueError:
        return False


def is_any_cidr(text: str) -> bool:
    return text.strip().upper() in ("", "*", "ANY", "ALL", "0.0.0.0/0")


def split_address_list(text: str | None) -> list[str]:
    if text is None:
        return []
    normalized = (
        text.replace("\u00a0", " ")
        .replace("\t", " ")
        .replace("\r\n", ";")
        .replace("\r", ";")
        .replace("\n", ";")
        .replace(",", ";")
        .replace("\uff0c", ";")
        .replace("\uff1b", ";")
        .replace(" ", ";")
    )
    return [part.strip() for part in normalized.split(";") if part.strip()]


_DIRECTION_IN_WORDS = ("IN", "INBOUND", "인바운드", "수신")
_DIRECTION_OUT_WORDS = ("OUT", "OUTBOUND", "아웃바운드", "송신")
_DIRECTION_BOTH_WORDS = (
    "BOTH", "ANY", "ALL", "양방향", "양방", "쌍방향",
    "BIDIRECTIONAL", "BI-DIRECTIONAL",
)
_DIRECTION_INSIDE_TOKENS = ("내부", "INSIDE", "INTERNAL")
_DIRECTION_OUTSIDE_TOKENS = ("외부", "OUTSIDE", "EXTERNAL")


def _normalize_direction_synonym(value: str) -> str:
    """Map Korean/business direction labels onto IN/OUT/BOTH or #INVALID.

    ``value`` is already uppercased and trimmed. Blank/IN/OUT/BOTH are handled
    by the caller. Arrow phrases (외부->내부, internal->external, 내부-외부, ...)
    resolve by flow direction; a standalone 내부/외부 stays #INVALID.
    """
    if value in _DIRECTION_IN_WORDS:
        return "IN"
    if value in _DIRECTION_OUT_WORDS:
        return "OUT"
    if value in _DIRECTION_BOTH_WORDS:
        return "BOTH"
    return _normalize_direction_arrow_phrase(value)


def _normalize_direction_arrow_phrase(value: str) -> str:
    canonical = (
        value.replace("\u2192", ">").replace("->", ">").replace("-", ">")
    )
    # Mirror VBA: split on the FIRST '>' only; reject if a separator remains in
    # the right side (3+ tokens / repeated separators). Empty tokens are NOT
    # filtered, so '내부-->외부', '->', '내부-' stay #INVALID.
    pos = canonical.find(">")
    if pos < 0:
        return "#INVALID"
    src = canonical[:pos].strip()
    dst = canonical[pos + 1:].strip()
    if ">" in dst:
        return "#INVALID"
    if not src or not dst:
        return "#INVALID"
    if src in _DIRECTION_OUTSIDE_TOKENS and dst in _DIRECTION_INSIDE_TOKENS:
        return "IN"
    if src in _DIRECTION_INSIDE_TOKENS and dst in _DIRECTION_OUTSIDE_TOKENS:
        return "OUT"
    return "#INVALID"


@dataclass(frozen=True)
class Firewall:
    firewall_name: str
    vendor: str = ""
    enabled: bool = True
    comment: str = ""


@dataclass(frozen=True)
class FirewallRange:
    firewall_name: str
    source_cidr: str
    destination_cidr: str
    direction: str = "BOTH"
    path_order: int = 999999
    enabled: bool = True
    comment: str = ""


@dataclass
class RouteResult:
    status: str = ""
    target_firewalls: str = ""
    firewall_path: str = ""
    zone_path: str = ""
    source_zone: str = ""
    destination_zone: str = ""
    validation_message: str = ""
    match_details: str = ""
    path_count: int = 0


@dataclass
class MatchedRange:
    rule: FirewallRange
    row_order: int


@dataclass
class RouteEngine:
    firewalls: list[Firewall] = field(default_factory=list)
    firewall_ranges: list[FirewallRange] = field(default_factory=list)
    _enabled_firewalls: set[str] = field(init=False, repr=False)
    _range_rows: list[MatchedRange] = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self._enabled_firewalls = {
            firewall.firewall_name for firewall in self.firewalls if firewall.enabled
        }
        self._range_rows = [
            MatchedRange(rule=rule, row_order=index)
            for index, rule in enumerate(self.firewall_ranges, start=1)
            if rule.enabled and rule.firewall_name in self._enabled_firewalls
        ]

    @staticmethod
    def normalize_direction(direction: str) -> str:
        value = (direction or "").strip().upper()
        if value == "":
            return "BOTH"
        if value in ("IN", "OUT", "BOTH"):
            return value
        return _normalize_direction_synonym(value)

    def analyze(self, src_ip: str, dst_ip: str, direction: str = "") -> RouteResult:
        flow_direction = self.normalize_direction(direction)
        if flow_direction == "#INVALID":
            return RouteResult(
                status="DIRECTION_MISMATCH",
                validation_message=f"Invalid direction: {direction}",
            )

        matches = self._select_matches(src_ip, dst_ip, flow_direction)
        if matches:
            return self._result_from_matches(matches)

        reverse_matches = self._select_matches(dst_ip, src_ip, "BOTH")
        if reverse_matches:
            return RouteResult(
                status="DIRECTION_MISMATCH",
                validation_message="No definition for requested direction; opposite direction exists",
                match_details=f"source_ip={src_ip}; destination_ip={dst_ip}",
            )

        return RouteResult(
            status="NO_MATCH",
            validation_message="No firewall range definition matched",
            match_details=f"source_ip={src_ip}; destination_ip={dst_ip}",
        )

    def _select_matches(
        self,
        src_ip: str,
        dst_ip: str,
        flow_direction: str,
    ) -> list[MatchedRange]:
        matches = [
            item for item in self._range_rows
            if self._direction_matches(item.rule.direction, flow_direction)
            and self._address_list_overlaps(src_ip, item.rule.source_cidr)
            and self._address_list_overlaps(dst_ip, item.rule.destination_cidr)
        ]
        return sorted(matches, key=self._match_key)

    @staticmethod
    def _direction_matches(rule_direction: str, flow_direction: str) -> bool:
        rule_value = RouteEngine.normalize_direction(rule_direction)
        if rule_value == "#INVALID":
            return False
        if flow_direction == "BOTH" or rule_value == "BOTH":
            return True
        return rule_value == flow_direction

    @staticmethod
    def _address_list_overlaps(request_value: str, definition_value: str) -> bool:
        if is_any_cidr(definition_value):
            return True
        request_tokens = split_address_list(request_value)
        definition_tokens = split_address_list(definition_value)
        if not request_tokens or not definition_tokens:
            return False
        return any(
            ranges_overlap(request_token, definition_token)
            for request_token in request_tokens
            for definition_token in definition_tokens
        )

    @staticmethod
    def _match_key(item: MatchedRange) -> tuple[int, int, str]:
        return (int(item.rule.path_order), item.row_order, item.rule.firewall_name)

    @staticmethod
    def _result_from_matches(matches: list[MatchedRange]) -> RouteResult:
        firewalls: list[str] = []
        for item in matches:
            if item.rule.firewall_name not in firewalls:
                firewalls.append(item.rule.firewall_name)
        first = matches[0].rule
        return RouteResult(
            status="OK",
            target_firewalls=";".join(firewalls),
            firewall_path=">".join(item.rule.firewall_name for item in matches),
            zone_path=f"{first.source_cidr}>{first.destination_cidr}",
            source_zone=first.source_cidr,
            destination_zone=first.destination_cidr,
            validation_message="Firewall range definition matched",
            match_details="; ".join(
                f"{item.rule.firewall_name}: {item.rule.source_cidr} -> "
                f"{item.rule.destination_cidr} ({item.rule.direction})"
                for item in matches
            ),
            path_count=len(matches),
        )
