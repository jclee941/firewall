from __future__ import annotations

from ipaddress import ip_network
from typing import Final

from scripts.workbook_contract import (
    EXAMPLE_REQUEST_ROWS,
    FIREWALLS,
    FIREWALL_RANGES,
    SECUI_CLI_HEADERS,
    VENDOR_CLI_TEMPLATES,
)

REQ_TEAM: Final = "요청부서"
REQ_DOC: Final = "요청번호"
REQ_FILE: Final = "원본파일"
REQ_ROW: Final = "원본행"
REQ_TARGET: Final = "대상방화벽"
REQ_SOURCE_IP: Final = "출발지IP"
REQ_SOURCE_NAME: Final = "출발지설명"
REQ_DESTINATION_IP: Final = "목적지IP"
REQ_DESTINATION_NAME: Final = "목적지설명"
REQ_PROTOCOL: Final = "프로토콜"
REQ_PORT: Final = "포트"
REQ_DIRECTION: Final = "방향"
REQ_PURPOSE: Final = "용도"
REQ_NOTE: Final = "비고"


def secui_cli_seed_rows() -> list[list[str | int]]:
    groups: dict[str, dict[str, str | dict[str, str]]] = {}
    enabled_firewalls = {
        str(row[0]).strip()
        for row in FIREWALLS[1:]
        if str(row[1]).strip().upper() == "SECUI" and _enabled(str(row[2]))
    }
    command_template = str(VENDOR_CLI_TEMPLATES[1][3])
    service_index: dict[str, set[str]] = {}

    for request in EXAMPLE_REQUEST_ROWS:
        source_object = _object_or_value(str(request.get(REQ_SOURCE_NAME, "")), str(request[REQ_SOURCE_IP]))
        destination_object = _object_or_value(str(request.get(REQ_DESTINATION_NAME, "")), str(request[REQ_DESTINATION_IP]))
        service_object = _service_text(str(request[REQ_PROTOCOL]), str(request[REQ_PORT]))
        for firewall_name in _split_targets(str(request.get(REQ_TARGET, ""))):
            if firewall_name not in enabled_firewalls:
                continue
            key = _source_destination_key(
                firewall_name, source_object, str(request[REQ_SOURCE_IP]), destination_object, str(request[REQ_DESTINATION_IP])
            )
            service_index.setdefault(key, set()).add(_clean(service_object))

    for request in EXAMPLE_REQUEST_ROWS:
        source_object = _object_or_value(str(request.get(REQ_SOURCE_NAME, "")), str(request[REQ_SOURCE_IP]))
        destination_object = _object_or_value(str(request.get(REQ_DESTINATION_NAME, "")), str(request[REQ_DESTINATION_IP]))
        service_object = _service_text(str(request[REQ_PROTOCOL]), str(request[REQ_PORT]))
        for firewall_name in _split_targets(str(request.get(REQ_TARGET, ""))):
            if firewall_name not in enabled_firewalls:
                continue
            destination_address = str(request[REQ_DESTINATION_IP])
            source_destination_key = _source_destination_key(
                firewall_name, source_object, str(request[REQ_SOURCE_IP]), destination_object, destination_address
            )
            if len(service_index.get(source_destination_key, set())) > 1:
                group_key = "SRC_DST\x1e" + source_destination_key
            else:
                group_key = "DST_SVC\x1e" + _destination_service_key(
                    firewall_name, destination_object, destination_address, service_object
                )
            if group_key not in groups:
                range_info = _range_info(
                    firewall_name,
                    str(request[REQ_SOURCE_IP]),
                    str(request[REQ_DESTINATION_IP]),
                    str(request.get(REQ_DIRECTION, "")),
                )
                groups[group_key] = {
                    "firewall": firewall_name,
                    "request_row": str(request[REQ_ROW]),
                    "team": str(request.get(REQ_TEAM, "")),
                    "doc": str(request.get(REQ_DOC, "")),
                    "file": str(request.get(REQ_FILE, "")),
                    "destination_object": destination_object,
                    "destination_address": destination_address,
                    "service_object": service_object,
                    "source_interface": range_info["source_interface"],
                    "destination_interface": range_info["destination_interface"],
                    "network_scope": range_info["network_scope"],
                    "protocol": str(request[REQ_PROTOCOL]).lower().strip(),
                    "port": str(request[REQ_PORT]).strip(),
                    "purpose": str(request.get(REQ_PURPOSE, "")),
                    "note": str(request.get(REQ_NOTE, "")),
                    "sources": {},
                    "rows": {},
                    "services": {},
                }
            sources = groups[group_key]["sources"]
            rows = groups[group_key]["rows"]
            assert isinstance(sources, dict)
            assert isinstance(rows, dict)
            sources[source_object] = str(request[REQ_SOURCE_IP])
            rows[str(request[REQ_ROW])] = "Y"
            _add_service(groups[group_key], service_object)

    out: list[list[str | int]] = [list(SECUI_CLI_HEADERS)]
    for index, group in enumerate(groups.values(), start=1):
        out.append(_seed_output_row(index, group, command_template))
    return out


def _seed_output_row(index: int, group: dict[str, str | dict[str, str]], command_template: str) -> list[str | int]:
    sources = group["sources"]
    rows = group["rows"]
    assert isinstance(sources, dict)
    assert isinstance(rows, dict)
    policy_name = _policy_name(group)
    source_group = _group_name("SRC", group, policy_name)
    destination_group = _group_name("DST", group, policy_name)
    service_group = _group_name("SVC", group, policy_name)
    source_members = ";".join(sources.keys())
    destination_member = str(group["destination_object"])
    service_member = str(group["service_object"])

    commands: list[str] = []
    if not _is_any(source_members):
        commands.append(_group_command("addrgrp", source_group, source_members, str(group["firewall"])))
    if not _is_any(destination_member):
        commands.append(_group_command("addrgrp", destination_group, destination_member, str(group["firewall"])))
    if not _is_any(service_member):
        commands.append(_group_command("svcgrp", service_group, service_member, str(group["firewall"])))
    commands.append(_render_policy_command(command_template, group, policy_name, source_group, destination_group, service_group))

    return [
        index,
        str(group["firewall"]),
        policy_name,
        "\n".join(commands),
        "장비 CLI에서 'fw set srule help'로 옵션명 확인 후 적용 / 룰별 그룹객체 생성 후 정책 생성",
        str(group["team"]),
        str(group["doc"]),
        str(group["file"]),
        ";".join(rows.keys()),
    ]


def _render_policy_command(
    template: str,
    group: dict[str, str | dict[str, str]],
    policy_name: str,
    source_group: str,
    destination_group: str,
    service_group: str,
) -> str:
    replacements = {
        "firewall_name": str(group["firewall"]),
        "policy_name": policy_name,
        "source_interface": str(group["source_interface"]),
        "destination_interface": str(group["destination_interface"]),
        "source_object": _object_reference(source_group, group["sources"]),
        "destination_object": _object_reference(destination_group, str(group["destination_object"])),
        "service_object": _object_reference(service_group, str(group["service_object"])),
        "description": _description(str(group["purpose"]), str(group["note"])),
    }
    rendered = template
    for key, value in replacements.items():
        rendered = rendered.replace("{" + key + "}", _clean(value))
        rendered = rendered.replace("{" + key + "_q}", _quote(value))
    return rendered


def _object_reference(group_name: str, member: str | dict[str, str]) -> str:
    if isinstance(member, dict):
        member_text = ";".join(member.keys())
    else:
        member_text = member
    if _is_any(member_text):
        return "ANY"
    return group_name


def _source_destination_key(firewall_name: str, source_object: str, source_address: str, destination_object: str, destination_address: str) -> str:
    return _key(firewall_name, source_object, source_address, destination_object, destination_address)


def _destination_service_key(firewall_name: str, destination_object: str, destination_address: str, service_object: str) -> str:
    return _key(firewall_name, destination_object, destination_address, service_object)


def _key(*parts: str) -> str:
    return "\x1e".join(_clean(part).lower() for part in parts)


def _add_service(group: dict[str, str | dict[str, str]], service_object: str) -> None:
    services = group["services"]
    assert isinstance(services, dict)
    clean_service = _clean(service_object)
    if clean_service:
        services[clean_service] = "Y"
        group["service_object"] = ";".join(services.keys())


def _policy_name(group: dict[str, str | dict[str, str]]) -> str:
    parts = ("doc", "firewall", "network_scope", "service_object", "destination_object", "destination_address")
    return _object_token("_".join(str(group[part]) for part in parts))[:120]


def _group_name(prefix: str, group: dict[str, str | dict[str, str]], policy_name: str) -> str:
    return _object_token(f"GRP_{prefix}_{group['network_scope']}_{policy_name}")[:120]


def _group_command(group_type: str, group_name: str, members: str, firewall_name: str) -> str:
    return f"fw set {group_type} name {_quote(group_name)} member {_quote(members)} # device={firewall_name}"


def _range_info(firewall_name: str, source: str, destination: str, direction: str) -> dict[str, str]:
    for row in FIREWALL_RANGES[1:]:
        if str(row[0]).strip() != firewall_name or not _enabled(str(row[5])):
            continue
        if not _direction_matches(str(row[3]), direction):
            continue
        if _overlaps(source, str(row[1])) and _overlaps(destination, str(row[2])):
            scope = f"{row[9]}_TO_{row[10]}" if len(row) >= 11 and (row[9] or row[10]) else f"{row[1]}_TO_{row[2]}"
            return {
                "network_scope": str(scope),
                "source_interface": str(row[7] or "ANY") if len(row) >= 8 else "ANY",
                "destination_interface": str(row[8] or "ANY") if len(row) >= 9 else "ANY",
            }
    return {"network_scope": "ANY_TO_ANY", "source_interface": "ANY", "destination_interface": "ANY"}


def _overlaps(request_value: str, definition_value: str) -> bool:
    if _is_any(definition_value):
        return True
    request_parts = _split_values(request_value)
    definition_parts = _split_values(definition_value)
    return any(_network_overlaps(left, right) for left in request_parts for right in definition_parts)


def _network_overlaps(left: str, right: str) -> bool:
    try:
        return ip_network(left, strict=False).overlaps(ip_network(right, strict=False))
    except ValueError:
        return False


def _direction_matches(rule_direction: str, request_direction: str) -> bool:
    rule_value = _normalize_direction(rule_direction)
    request_value = _normalize_direction(request_direction)
    if rule_value == "#INVALID" or request_value == "#INVALID":
        return False
    return rule_value == "BOTH" or request_value == "BOTH" or rule_value == request_value


_DIR_IN = ("IN", "INBOUND", "인바운드", "수신")
_DIR_OUT = ("OUT", "OUTBOUND", "아웃바운드", "송신")
_DIR_BOTH = ("", "BOTH", "ANY", "ALL", "양방향", "양방", "쌍방향", "BIDIRECTIONAL", "BI-DIRECTIONAL")
_DIR_INSIDE = ("내부", "INSIDE", "INTERNAL")
_DIR_OUTSIDE = ("외부", "OUTSIDE", "EXTERNAL")


def _normalize_direction(direction: str) -> str:
    value = (direction or "").strip().upper()
    if value in _DIR_IN:
        return "IN"
    if value in _DIR_OUT:
        return "OUT"
    if value in _DIR_BOTH:
        return "BOTH"
    canonical = value.replace("\u2192", ">").replace("->", ">").replace("-", ">")
    parts = [part.strip() for part in canonical.split(">") if part.strip()]
    if len(parts) != 2:
        return "#INVALID"
    src, dst = parts
    if src in _DIR_OUTSIDE and dst in _DIR_INSIDE:
        return "IN"
    if src in _DIR_INSIDE and dst in _DIR_OUTSIDE:
        return "OUT"
    return "#INVALID"


def _service_text(protocol: str, port: str) -> str:
    port_text = _clean(port)
    if port_text == "" or port_text.isnumeric():
        return _clean(f"{protocol.lower().strip()}/{port_text}")
    return port_text


def _object_or_value(object_name: str, fallback: str) -> str:
    object_text = _clean(object_name)
    if object_text:
        return object_text
    return _clean(fallback)


def _split_targets(value: str) -> list[str]:
    return [part.strip() for part in value.split(";") if part.strip()]


def _split_values(value: str) -> list[str]:
    normalized = value.replace("\r\n", ";").replace("\r", ";").replace("\n", ";").replace(",", ";").replace(" ", ";")
    return [part.strip() for part in normalized.split(";") if part.strip()]


def _enabled(value: str) -> bool:
    return value.strip().upper() in ("", "Y", "YES", "TRUE", "1")


def _is_any(value: str) -> bool:
    return value.strip().upper() in ("", "ANY", "ALL", "*", "0.0.0.0/0")


def _description(purpose: str, note: str) -> str:
    return _clean(" / ".join(part for part in (purpose, note) if part.strip()))[:255]


def _quote(value: str) -> str:
    return '"' + _clean(value).replace('"', "'") + '"'


def _clean(value: str) -> str:
    return " ".join(str(value).strip().split())


def _object_token(value: str) -> str:
    chars = [ch if ch.isascii() and ch.isalnum() else "_" for ch in _clean(value)]
    result = "".join(chars)
    while "__" in result:
        result = result.replace("__", "_")
    return result.strip("_").upper() or "ANY"
