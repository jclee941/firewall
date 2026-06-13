# VBA KNOWLEDGE BASE

## Overview

Shipped Excel macro runtime. Source here is injected into `vbaProject.bin` by `scripts/build_xlsm.py`; tests inspect the injected modules but do not execute Excel macros.

## Where To Look

| Task | Location | Notes |
| --- | --- | --- |
| Merge request workbooks | `FirewallPolicyAutomation.bas` | Folder recursion, sheet/header detection, merged-cell fill-down |
| Compute `대상방화벽` | `FirewallRouteAnalysis.bas` | Must mirror `tests/route_oracle.py` |
| Change parser aliases | `FirewallPolicyAutomation.bas`, `../tests/request_parser_oracle.py`, parser tests | Built-in aliases and `header_aliases` sheet are both active |
| Change workbook setup macros | `FirewallPolicyAutomation.bas`, `../scripts/build_xlsm.py`, `../tests/test_xlsm_structure.py` | VBA setup seed and Python build seed must match |
| Change SECUI conversion | `FirewallPolicyAutomation.bas`, `../tests/test_xlsm_structure.py` | Split `대상방화벽` by `;`; emit only enabled SECUI firewalls |
| Change SECUI service examples | `FirewallPolicyAutomation.bas`, `../scripts/workbook_contract.py`, `../tests/test_secui_service_catalog.py` | Reference sheet only; conversion still derives service from protocol/port |

## Commands

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python -m pytest tests/test_route_oracle.py -v
./.venv/bin/python -m pytest tests/test_xlsm_structure.py -v
```

LibreOffice/Excel macro execution is unavailable here. Verification is oracle parity plus generated workbook inspection.

## Runtime Contract

- `FirewallRouteAnalysis.bas` uses `firewalls` and `firewall_ranges` only. Do not reintroduce `network_definitions`, `routing_paths`, zone graph BFS, or legacy fallback.
- Direction values normalize to `IN`, `OUT`, `BOTH`, or `#INVALID`; blanks mean `BOTH`.
- Matched range rows sort by `path_order`, row order, firewall name.
- `대상방화벽` is column 7 and joins unique firewall names with `;`.
- `방화벽경로` preserves matched range order and joins with `>`.
- Duplicate marking runs after route analysis and must preserve the route-owned status-cell color.
- `service_catalog` is an operator reference sheet for SECUI service text examples, not a route-analysis input.

## Parser Contract

- Header detection must score by field content, not by an exact `No`/`번호` anchor alone.
- Explicit `parse_sheet` with an unrecognizable header raises/logs failure; never silently falls back.
- Merged data cells fill down across logical rows; header rows do not get fill-down pollution.
- IP and port cells are never merged by formatting because route and duplicate logic need per-row values.
- Always read the `header_aliases` table sheet, even when `settings.header_alias` is blank.

## Anti-Patterns

- Do not change VBA route/parser behavior without changing the matching Python oracle and tests.
- Do not blanket-swallow `Workbook_Open` or auto-run errors.
- Do not pre-fill `target_firewalls` before route analysis.
- Do not make display-only UX helpers change algorithm inputs or outputs.
