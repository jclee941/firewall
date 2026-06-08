"""Pure-Python reference implementation of the firewall route-analysis algorithm.

This module is the BINDING CONTRACT for vba/FirewallRouteAnalysis.bas.
Every function here must mirror the VBA implementation exactly so that the
Python tests double as the specification.

Algorithm (confirmed by Oracle review):
  1. source_ip / destination_ip -> zone via longest-prefix match in
     network_definitions (enabled rows only).
  2. Build a directed zone graph from routing_paths (enabled rows only);
     each edge carries firewall_name and path_order.
  3. Deterministic BFS shortest zone path. firewall_path preserves hop order;
     target_firewalls = unique first-seen firewalls on the chosen path.
  4. Tie-break: per-edge key "{path_order:06d}|{firewall_name}|{to_zone}",
     path key = ";".join(edge keys); smallest string wins.
  5. Statuses: OK, MULTI_PATH, INTRA_ZONE, ZONE_UNRESOLVED, NO_PATH,
     DIRECTION_MISMATCH, LEGACY_FALLBACK.
  6. Direction: blank/BOTH = forward then reverse; OUT = src->dst; IN = dst->src.
  7. Legacy CIDR-overlap fallback only when graph fails and toggle on.

IPv4 numbers use plain integers (VBA uses Double, exact for 0..2^32-1).
"""

from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field


# --------------------------------------------------------------------------- #
# IPv4 / CIDR primitives (mirror VBA IpToNumber / CidrStart / CidrEnd)
# --------------------------------------------------------------------------- #

def ip_to_number(ip_text: str) -> int:
    parts = ip_text.strip().split(".")
    if len(parts) != 4:
        raise ValueError(f"Invalid IPv4 address: {ip_text}")
    value = 0
    for p in parts:
        p = p.strip()
        if not p.isdigit():
            raise ValueError(f"Invalid IPv4 address: {ip_text}")
        octet = int(p)
        if octet < 0 or octet > 255:
            raise ValueError(f"Invalid IPv4 address: {ip_text}")
        value = value * 256 + octet
    return value


def cidr_prefix_length(cidr_text: str) -> int:
    parts = cidr_text.split("/")
    if len(parts) == 1:
        return 32
    prefix = int(parts[1].strip())
    if prefix < 0 or prefix > 32:
        raise ValueError(f"Invalid CIDR prefix: {cidr_text}")
    return prefix


def cidr_base_ip(cidr_text: str) -> str:
    return cidr_text.split("/")[0].strip()


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
        if not left_cidr or not right_cidr:
            return False
        ls, le = cidr_start(left_cidr), cidr_end(left_cidr)
        rs, re = cidr_start(right_cidr), cidr_end(right_cidr)
        return ls <= re and rs <= le
    except ValueError:
        return False


def split_address_list(text: str) -> list[str]:
    if text is None:
        return []
    norm = (
        text.replace("\u00a0", " ")
        .replace("\t", " ")
        .replace("\r\n", ";")
        .replace("\r", ";")
        .replace("\n", ";")
        .replace(",", ";")
        .replace("\uff0c", ";")
        .replace("\uff1b", ";")
    )
    return [p.strip() for p in norm.split(";") if p.strip()]


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #

@dataclass(frozen=True)
class Network:
    network_name: str
    network_cidr: str
    zone: str
    site: str = ""
    enabled: bool = True


@dataclass(frozen=True)
class Firewall:
    firewall_name: str
    vendor: str = ""
    enabled: bool = True
    comment: str = ""


@dataclass(frozen=True)
class RoutingPath:
    firewall_name: str
    src_zone: str
    dst_zone: str
    ingress_if: str = ""
    egress_if: str = ""
    path_order: int = 999999
    enabled: bool = True


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


# --------------------------------------------------------------------------- #
# Engine
# --------------------------------------------------------------------------- #

@dataclass
class RouteEngine:
    networks: list[Network] = field(default_factory=list)
    firewalls: list[Firewall] = field(default_factory=list)
    routing_paths: list[RoutingPath] = field(default_factory=list)
    fallback_enabled: bool = False

    def __post_init__(self) -> None:
        self._enabled_fw = {
            f.firewall_name for f in self.firewalls if f.enabled
        }
        # directed adjacency: src_zone -> list of edges
        self._graph: dict[str, list[RoutingPath]] = {}
        for rp in self.routing_paths:
            if not rp.enabled:
                continue
            if rp.firewall_name not in self._enabled_fw:
                continue
            if not rp.src_zone or not rp.dst_zone:
                continue
            self._graph.setdefault(rp.src_zone, []).append(rp)
        for zone in self._graph:
            self._graph[zone].sort(key=self._edge_key)
        self._zone_cache: dict[str, str] = {}
        self._path_cache: dict[str, RouteResult] = {}

    # -- zone resolution (longest-prefix match) ----------------------------- #
    def resolve_zone(self, ip_text: str) -> str:
        """Resolve a request value to a zone.

        The value may be a single IP, a CIDR (대역), or a delimited list of
        either. Each token is resolved by longest-prefix overlap against the
        network_definitions; the first resolvable token wins.
        """
        if ip_text in self._zone_cache:
            return self._zone_cache[ip_text]
        result = "#UNRESOLVED"
        for token in split_address_list(ip_text):
            z = self._resolve_token_zone(token)
            if z not in ("#UNRESOLVED",):
                result = z
                break
        self._zone_cache[ip_text] = result
        return result

    def _resolve_token_zone(self, token: str) -> str:
        """Resolve one IP-or-CIDR token to a zone by longest-prefix overlap."""
        try:
            tok_start = cidr_start(token)
            tok_end = cidr_end(token)
        except ValueError:
            return "#UNRESOLVED"
        best_prefix = -1
        best_zone = "#UNRESOLVED"
        ambiguous = False
        for n in self.networks:
            if not n.enabled:
                continue
            try:
                start = cidr_start(n.network_cidr)
                end = cidr_end(n.network_cidr)
                prefix = cidr_prefix_length(n.network_cidr)
            except ValueError:
                continue
            # token overlaps this network (single IP is a /32 range)
            if start <= tok_end and tok_start <= end:
                if prefix > best_prefix:
                    best_prefix = prefix
                    best_zone = n.zone
                    ambiguous = False
                elif prefix == best_prefix and n.zone != best_zone:
                    ambiguous = True
        if ambiguous:
            return "#AMBIGUOUS"
        return best_zone

    # -- edge / path keys --------------------------------------------------- #
    @staticmethod
    def _edge_key(edge: RoutingPath) -> str:
        return f"{int(edge.path_order):06d}|{edge.firewall_name}|{edge.dst_zone}"

    def _path_key(self, edges: list[RoutingPath]) -> str:
        return ";".join(self._edge_key(e) for e in edges)

    # -- BFS shortest zone paths ------------------------------------------- #
    def _find_shortest_paths(self, start_zone: str, end_zone: str) -> list[list[RoutingPath]]:
        results: list[list[RoutingPath]] = []
        if start_zone == end_zone:
            return results
        queue: deque[tuple[str, list[RoutingPath], frozenset[str]]] = deque()
        queue.append((start_zone, [], frozenset({start_zone})))
        shortest_len = -1
        while queue:
            zone, edges, visited = queue.popleft()
            if shortest_len != -1 and len(edges) >= shortest_len:
                continue
            for edge in self._graph.get(zone, []):
                nxt = edge.dst_zone
                if nxt in visited:
                    continue
                new_edges = edges + [edge]
                if nxt == end_zone:
                    if shortest_len == -1:
                        shortest_len = len(new_edges)
                    if len(new_edges) == shortest_len:
                        results.append(new_edges)
                else:
                    if shortest_len == -1 or len(new_edges) < shortest_len:
                        queue.append((nxt, new_edges, visited | {nxt}))
        return results

    def _choose_best(self, paths: list[list[RoutingPath]]) -> list[RoutingPath]:
        return min(paths, key=self._path_key)

    def _directed(self, start_zone: str, end_zone: str) -> RouteResult:
        cache_key = f"{start_zone}|{end_zone}"
        if cache_key in self._path_cache:
            return self._path_cache[cache_key]
        paths = self._find_shortest_paths(start_zone, end_zone)
        res = RouteResult()
        if not paths:
            res.status = "NO_PATH"
            res.validation_message = "No routing path found"
            res.match_details = f"from={start_zone}; to={end_zone}"
        else:
            best = self._choose_best(paths)
            res.firewall_path = self._build_fw_path(best)
            res.target_firewalls = self._build_target_set(best)
            res.zone_path = self._build_zone_path(start_zone, best)
            res.match_details = self._build_match_details(start_zone, best)
            res.path_count = len(paths)
            if len(paths) == 1:
                res.status = "OK"
                res.validation_message = "Path resolved"
            else:
                res.status = "MULTI_PATH"
                res.validation_message = (
                    "Multiple equal shortest paths; selected by path_order/firewall_name"
                )
        self._path_cache[cache_key] = res
        return res

    # -- output builders ---------------------------------------------------- #
    @staticmethod
    def _build_fw_path(edges: list[RoutingPath]) -> str:
        return ">".join(e.firewall_name for e in edges)

    @staticmethod
    def _build_target_set(edges: list[RoutingPath]) -> str:
        seen: list[str] = []
        for e in edges:
            if e.firewall_name not in seen:
                seen.append(e.firewall_name)
        return ";".join(seen)

    @staticmethod
    def _build_zone_path(start_zone: str, edges: list[RoutingPath]) -> str:
        zones = [start_zone] + [e.dst_zone for e in edges]
        return ">".join(zones)

    @staticmethod
    def _build_match_details(start_zone: str, edges: list[RoutingPath]) -> str:
        parts = []
        cur = start_zone
        for e in edges:
            parts.append(
                f"{cur} -> {e.dst_zone} via {e.firewall_name}"
                f"(order={int(e.path_order)},in={e.ingress_if},out={e.egress_if})"
            )
            cur = e.dst_zone
        return "; ".join(parts)

    # -- legacy fallback (CIDR overlap against network_definitions) --------- #
    def _legacy_fallback(self, src_ip: str, dst_ip: str) -> RouteResult:
        # Mirror legacy behaviour: a firewall is a target if any routing_paths
        # row for it touches a zone whose network overlaps src or dst.
        # For the oracle we treat fallback as: firewalls whose any enabled
        # routing edge's src/dst zone contains src or dst IP.
        res = RouteResult()
        targets: list[str] = []
        for rp in self.routing_paths:
            if not rp.enabled or rp.firewall_name not in self._enabled_fw:
                continue
            zones = {rp.src_zone, rp.dst_zone}
            hit = False
            for n in self.networks:
                if not n.enabled or n.zone not in zones:
                    continue
                if ranges_overlap(src_ip, n.network_cidr) or ranges_overlap(dst_ip, n.network_cidr):
                    hit = True
                    break
            if hit and rp.firewall_name not in targets:
                targets.append(rp.firewall_name)
        if targets:
            res.status = "LEGACY_FALLBACK"
            res.target_firewalls = ";".join(targets)
            res.firewall_path = ">".join(targets)
            res.validation_message = "Resolved using legacy CIDR-overlap fallback"
            res.match_details = "fallback by network overlap"
        else:
            res.status = "NO_PATH"
            res.validation_message = "No routing path found"
        return res

    # -- main entry --------------------------------------------------------- #
    @staticmethod
    def normalize_direction(direction: str) -> str:
        d = (direction or "").strip().upper()
        if d == "":
            return "BOTH"
        if d in ("IN", "OUT", "BOTH"):
            return d
        return "#INVALID"

    def analyze(self, src_ip: str, dst_ip: str, direction: str = "") -> RouteResult:
        d = self.normalize_direction(direction)
        if d == "#INVALID":
            return RouteResult(
                status="DIRECTION_MISMATCH",
                validation_message=f"Invalid direction: {direction}",
            )

        src_zone = self.resolve_zone(src_ip)
        dst_zone = self.resolve_zone(dst_ip)

        if src_zone in ("#UNRESOLVED", "#AMBIGUOUS") or dst_zone in ("#UNRESOLVED", "#AMBIGUOUS"):
            msg = []
            if src_zone in ("#UNRESOLVED", "#AMBIGUOUS"):
                msg.append(f"Source zone unresolved ({src_ip})")
            if dst_zone in ("#UNRESOLVED", "#AMBIGUOUS"):
                msg.append(f"Destination zone unresolved ({dst_ip})")
            return RouteResult(
                status="ZONE_UNRESOLVED",
                source_zone="" if src_zone.startswith("#") else src_zone,
                destination_zone="" if dst_zone.startswith("#") else dst_zone,
                validation_message="; ".join(msg),
                match_details=f"source_ip={src_ip}; destination_ip={dst_ip}",
            )

        if src_zone == dst_zone:
            return RouteResult(
                status="INTRA_ZONE",
                source_zone=src_zone,
                destination_zone=dst_zone,
                validation_message="Source and destination in same zone; no firewall path required",
                match_details=f"source_zone={src_zone}; destination_zone={dst_zone}",
            )

        if d == "OUT":
            res = self._copy(self._directed(src_zone, dst_zone))
            if res.status == "NO_PATH":
                rev = self._directed(dst_zone, src_zone)
                if rev.status in ("OK", "MULTI_PATH"):
                    res.status = "DIRECTION_MISMATCH"
                    res.validation_message = "Requested OUT path not found, but reverse path exists"
                    res.match_details = rev.match_details
                    res.firewall_path = ""
                    res.target_firewalls = ""
                    res.zone_path = ""
        elif d == "IN":
            res = self._copy(self._directed(dst_zone, src_zone))
            if res.status == "NO_PATH":
                fwd = self._directed(src_zone, dst_zone)
                if fwd.status in ("OK", "MULTI_PATH"):
                    res.status = "DIRECTION_MISMATCH"
                    res.validation_message = "Requested IN path not found, but opposite path exists"
                    res.match_details = fwd.match_details
                    res.firewall_path = ""
                    res.target_firewalls = ""
                    res.zone_path = ""
        else:  # BOTH
            fwd = self._directed(src_zone, dst_zone)
            if fwd.status in ("OK", "MULTI_PATH"):
                res = self._copy(fwd)
            else:
                rev = self._directed(dst_zone, src_zone)
                if rev.status in ("OK", "MULTI_PATH"):
                    res = self._copy(rev)
                    res.validation_message += "; resolved using reverse direction under BOTH"
                else:
                    res = self._copy(fwd)

        if res.status == "NO_PATH" and self.fallback_enabled:
            fb = self._legacy_fallback(src_ip, dst_ip)
            if fb.status == "LEGACY_FALLBACK":
                res = fb

        res.source_zone = src_zone
        res.destination_zone = dst_zone
        return res

    @staticmethod
    def _copy(r: RouteResult) -> RouteResult:
        return RouteResult(
            status=r.status,
            target_firewalls=r.target_firewalls,
            firewall_path=r.firewall_path,
            zone_path=r.zone_path,
            source_zone=r.source_zone,
            destination_zone=r.destination_zone,
            validation_message=r.validation_message,
            match_details=r.match_details,
            path_count=r.path_count,
        )
