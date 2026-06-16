# TESTS KNOWLEDGE BASE

## Overview

Tests are the binding specification for the workbook and VBA behavior. They do not execute Excel macros; they verify Python mirrors, generated workbook structure, and injected VBA source.

## Where To Look

| Task | Location | Notes |
| --- | --- | --- |
| Range matching behavior | `route_oracle.py`, `test_route_oracle.py` | Must mirror `vba/FirewallRouteAnalysis.bas` |
| Request parsing behavior | `request_parser_oracle.py`, `test_request_parsing.py`, `test_merged_cells.py` | Mirrors parser logic in `vba/FirewallPolicyAutomation.bas` |
| Generated `.xlsm` invariants | `test_xlsm_structure.py` | Builds the workbook and inspects sheets, VBA modules, UX formatting |
| Workbook usage navigation | `test_workbook_usage_links.py` | Verifies `usage` opens first and links to key workbook sheets |
| Folder/sample parsing | `test_request_folder.py`, `test_folder_parse.py` | Uses `scripts/make_request_folder.py` output |
| Header alias behavior | `user_alias_oracle.py`, `test_user_alias.py` | Keep built-in aliases and settings/header_alias behavior aligned |
| SECUI CLI/GUI behavior | `test_secui_cli_script.py`, `test_secui_gui.py`, `test_secui_service_catalog.py`, `test_secui_policy_analysis.py` | Mirror in-workbook SECUI conversion; legacy analysis surfaces are explicitly removed |
| VBA runtime guards | `test_vba_runtime_guards.py` | Asserts no blanket `On Error Resume Next`; runtime sheets stay unprotected |
| Release bundle / weekly report | `test_release_bundle.py`, `test_weekly_report_sheet.py`, `test_workbook_convenience.py`, `test_date_cells.py` | Artifact shape and added workbook sheets |

## Commands

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python -m pytest tests/test_route_oracle.py -v
./.venv/bin/python -m pytest tests/test_xlsm_structure.py -v
```

## Invariants

- `tests/route_oracle.py` is the route/range algorithm contract.
- `tests/test_route_oracle.py` must cover every status and edge case the VBA route module implements.
- `requests` is a clean 14-column user sheet; header row 2 must match `scripts.workbook_contract.REQUESTS_HEADERS`. Source tracking is on hidden `_request_tracking`; route diagnostics on `route_results`.
- `대상방화벽` is column 3 in the built `requests` workbook; internal test/oracle field remains `target_firewalls`.
- Generated workbook must contain `firewall_ranges`; it must not require `network_definitions` or `routing_paths`.
- Injected VBA modules must match `vba/*.bas`; `Module1` must be absent.
- Release workflow sheet checks must track the same workbook contract as `test_sheets_and_headers`.
- `usage` quick links are UX-only and must not change route/SECUI behavior.

## Gotchas

- Tests import some local oracle modules through `sys.path`; keep existing import style unless doing a deliberate test package cleanup.
- `test_xlsm_structure.py` rebuilds `dist/firewall-policy-automation.xlsm`; do not commit `dist/`.
- Running folder tests may rewrite sample `.xlsx` files under `request-folder/`; avoid committing that binary churn.
- If a VBA behavior changes, add or update a Python oracle test first, then update VBA and workbook structure tests.
