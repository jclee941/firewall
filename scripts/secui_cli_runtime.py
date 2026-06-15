from __future__ import annotations

from collections.abc import Sequence
from datetime import date, datetime

from scripts.secui_cli_seed import (
    REQ_DESTINATION_IP,
    REQ_DESTINATION_NAME,
    REQ_DIRECTION,
    REQ_DOC,
    REQ_FILE,
    REQ_NOTE,
    REQ_PORT,
    REQ_PROTOCOL,
    REQ_PURPOSE,
    REQ_ROW,
    REQ_SOURCE_IP,
    REQ_SOURCE_NAME,
    REQ_TARGET,
    REQ_TEAM,
    _add_service,
    _clean,
    _destination_service_key,
    _direction_matches,
    _enabled,
    _object_or_value,
    _overlaps,
    _seed_output_row,
    _service_text,
    _source_destination_key,
    _split_targets,
)
from scripts.workbook_contract import SECUI_CLI_HEADERS, VENDOR_CLI_TEMPLATES

type Cell = str | int | float | bool | date | datetime | None
type RequestRecord = dict[str, Cell]
type Table = Sequence[Sequence[Cell]]
type CliGroup = dict[str, str | dict[str, str]]


def secui_cli_rows(
    requests: Sequence[RequestRecord],
    firewalls: Table,
    firewall_ranges: Table,
    vendor_cli_templates: Table = VENDOR_CLI_TEMPLATES,
) -> list[list[str | int]]:
    groups: dict[str, CliGroup] = {}
    enabled_firewalls = _enabled_secui_firewalls(firewalls)
    command_template = _command_template(vendor_cli_templates)
    service_index: dict[str, set[str]] = {}

    for request in requests:
        source_object = _request_source_object(request)
        destination_object = _request_destination_object(request)
        service_object = _request_service_object(request)
        for firewall_name in _split_targets(str(request.get(REQ_TARGET, "") or "")):
            if firewall_name not in enabled_firewalls:
                continue
            key = _source_destination_key(
                firewall_name,
                source_object,
                str(request.get(REQ_SOURCE_IP, "") or ""),
                destination_object,
                str(request.get(REQ_DESTINATION_IP, "") or ""),
            )
            service_index.setdefault(key, set()).add(_clean(service_object))

    for request in requests:
        source_object = _request_source_object(request)
        destination_object = _request_destination_object(request)
        service_object = _request_service_object(request)
        for firewall_name in _split_targets(str(request.get(REQ_TARGET, "") or "")):
            if firewall_name not in enabled_firewalls:
                continue
            destination_address = str(request.get(REQ_DESTINATION_IP, "") or "")
            source_destination_key = _source_destination_key(
                firewall_name,
                source_object,
                str(request.get(REQ_SOURCE_IP, "") or ""),
                destination_object,
                destination_address,
            )
            if len(service_index.get(source_destination_key, set())) > 1:
                group_key = "SRC_DST\x1e" + source_destination_key
            else:
                group_key = "DST_SVC\x1e" + _destination_service_key(
                    firewall_name,
                    destination_object,
                    destination_address,
                    service_object,
                )
            if group_key not in groups:
                groups[group_key] = _new_group(
                    request,
                    firewall_name,
                    destination_object,
                    destination_address,
                    service_object,
                    firewall_ranges,
                )
            sources = groups[group_key]["sources"]
            rows = groups[group_key]["rows"]
            assert isinstance(sources, dict)
            assert isinstance(rows, dict)
            sources[source_object] = str(request.get(REQ_SOURCE_IP, "") or "")
            rows[str(request.get(REQ_ROW, "") or "")] = "Y"
            _add_service(groups[group_key], service_object)

    out: list[list[str | int]] = [list(SECUI_CLI_HEADERS)]
    for index, group in enumerate(groups.values(), start=1):
        out.append(_seed_output_row(index, group, command_template))
    return out


def _enabled_secui_firewalls(firewalls: Table) -> set[str]:
    return {
        str(row[0]).strip()
        for row in firewalls[1:]
        if len(row) >= 3 and str(row[1]).strip().upper() == "SECUI" and _enabled(str(row[2]))
    }


def _command_template(vendor_cli_templates: Table) -> str:
    for row in vendor_cli_templates[1:]:
        if len(row) >= 4 and str(row[0]).strip().upper() == "SECUI" and _enabled(str(row[2])):
            return str(row[3])
    return str(VENDOR_CLI_TEMPLATES[1][3])


def _request_source_object(request: RequestRecord) -> str:
    return _object_or_value(str(request.get(REQ_SOURCE_NAME, "") or ""), str(request.get(REQ_SOURCE_IP, "") or ""))


def _request_destination_object(request: RequestRecord) -> str:
    return _object_or_value(
        str(request.get(REQ_DESTINATION_NAME, "") or ""),
        str(request.get(REQ_DESTINATION_IP, "") or ""),
    )


def _request_service_object(request: RequestRecord) -> str:
    return _service_text(str(request.get(REQ_PROTOCOL, "") or ""), str(request.get(REQ_PORT, "") or ""))


def _new_group(
    request: RequestRecord,
    firewall_name: str,
    destination_object: str,
    destination_address: str,
    service_object: str,
    firewall_ranges: Table,
) -> CliGroup:
    range_info = _range_info(
        firewall_name,
        str(request.get(REQ_SOURCE_IP, "") or ""),
        str(request.get(REQ_DESTINATION_IP, "") or ""),
        str(request.get(REQ_DIRECTION, "") or ""),
        firewall_ranges,
    )
    return {
        "firewall": firewall_name,
        "request_row": str(request.get(REQ_ROW, "") or ""),
        "team": str(request.get(REQ_TEAM, "") or ""),
        "doc": str(request.get(REQ_DOC, "") or ""),
        "file": str(request.get(REQ_FILE, "") or ""),
        "destination_object": destination_object,
        "destination_address": destination_address,
        "service_object": service_object,
        "source_interface": range_info["source_interface"],
        "destination_interface": range_info["destination_interface"],
        "network_scope": range_info["network_scope"],
        "protocol": str(request.get(REQ_PROTOCOL, "") or "").lower().strip(),
        "port": str(request.get(REQ_PORT, "") or "").strip(),
        "purpose": str(request.get(REQ_PURPOSE, "") or ""),
        "note": str(request.get(REQ_NOTE, "") or ""),
        "sources": {},
        "rows": {},
        "services": {},
    }


def _range_info(
    firewall_name: str,
    source: str,
    destination: str,
    direction: str,
    firewall_ranges: Table,
) -> dict[str, str]:
    for row in firewall_ranges[1:]:
        if len(row) < 6 or str(row[0]).strip() != firewall_name or not _enabled(str(row[5])):
            continue
        if not _direction_matches(str(row[3]), direction):
            continue
        if _overlaps(source, str(row[1])) and _overlaps(destination, str(row[2])):
            has_zone = len(row) >= 11 and (row[9] or row[10])
            scope = f"{row[9]}_TO_{row[10]}" if has_zone else f"{row[1]}_TO_{row[2]}"
            return {
                "network_scope": str(scope),
                "source_interface": str(row[7] or "ANY") if len(row) >= 8 else "ANY",
                "destination_interface": str(row[8] or "ANY") if len(row) >= 9 else "ANY",
            }
    return {"network_scope": "ANY_TO_ANY", "source_interface": "ANY", "destination_interface": "ANY"}
