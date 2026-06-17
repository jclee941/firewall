from __future__ import annotations

from collections.abc import Mapping

from tests.route_oracle import Firewall, FirewallRange, RouteEngine, RouteResult

from scripts.secui_cli_runtime import Cell, RequestRecord
from scripts.workbook_contract import (
    EXAMPLE_REQUEST_ROWS,
    FIREWALLS,
    FIREWALL_RANGES,
    ROUTE_RESULTS_HEADERS,
)


def request_seed_records() -> list[RequestRecord]:
    engine = _seed_route_engine()
    records: list[RequestRecord] = []
    for request in EXAMPLE_REQUEST_ROWS:
        result = _analyze_request(engine, request)
        record: RequestRecord = {}
        for key, value in request.items():
            record[key] = value
        record["대상방화벽"] = result.target_firewalls or str(request.get("대상방화벽", "") or "")
        records.append(record)
    return records


def route_result_seed_rows() -> list[list[str | int]]:
    engine = _seed_route_engine()
    rows: list[list[str | int]] = [list(ROUTE_RESULTS_HEADERS)]
    for request in request_seed_records():
        result = _analyze_request(engine, request)
        rows.append([
            str(request.get("요청부서", "") or ""),
            str(request.get("요청번호", "") or ""),
            str(request.get("출발지IP", "") or ""),
            str(request.get("출발지설명", "") or ""),
            str(request.get("목적지IP", "") or ""),
            str(request.get("목적지설명", "") or ""),
            str(request.get("프로토콜", "") or ""),
            str(request.get("포트", "") or ""),
            str(request.get("방향", "") or ""),
            str(request.get("대상방화벽", "") or result.target_firewalls),
            result.status,
            result.validation_message,
            result.firewall_path,
            result.source_zone,
            result.destination_zone,
            result.zone_path,
            result.match_details,
            str(request.get("원본파일", "") or ""),
            _cell_int(request.get("원본행", 0), 0),
        ])
    return rows


def _analyze_request(engine: RouteEngine, request: Mapping[str, Cell]) -> RouteResult:
    return engine.analyze(
        str(request.get("출발지IP", "") or ""),
        str(request.get("목적지IP", "") or ""),
        str(request.get("방향", "") or ""),
    )


def _seed_route_engine() -> RouteEngine:
    firewalls = [
        Firewall(str(row[0]), str(row[1] or ""), _enabled_flag(row[2]))
        for row in FIREWALLS[1:]
        if row[0]
    ]
    ranges = [
        FirewallRange(
            str(row[0]),
            str(row[1] or ""),
            str(row[2] or ""),
            str(row[3] or ""),
            _cell_int(row[4], 999999),
            _enabled_flag(row[5]),
            str(row[6] or ""),
        )
        for row in FIREWALL_RANGES[1:]
        if row[0]
    ]
    return RouteEngine(firewalls=firewalls, firewall_ranges=ranges)


def _enabled_flag(value: str | int | bool | None) -> bool:
    return str(value or "").strip().upper() in ("Y", "YES", "TRUE", "1")


def _cell_int(value: Cell, default: int) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return default
        try:
            return int(text)
        except ValueError:
            return default
    return default
