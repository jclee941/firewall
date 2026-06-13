# PROJECT KNOWLEDGE BASE

Generated: 2026-06-13
Commit: 8860d14
Branch: master

## Overview

Excel-native VBA tool for merging firewall-policy request spreadsheets and computing `대상방화벽` from user-defined firewall range rows. No PowerQuery; Linux build/test tooling produces the shipped macro-enabled `.xlsm`.

## Structure

```text
.
├── vba/                  # shipped VBA modules injected into vbaProject.bin
├── scripts/              # Linux builders/scaffolders for xlsm and request-folder
├── tests/                # Python oracle/spec tests plus workbook structure checks
├── docs/                 # operator schema and Excel-native runbook
├── request-folder/       # sample request tree; xlsx files are generated/churn-prone
└── .github/workflows/    # CI and tag-driven release automation
```

## Where To Look

| Task | Location | Notes |
| --- | --- | --- |
| Change route/range calculation | `tests/route_oracle.py`, `vba/FirewallRouteAnalysis.bas`, `tests/test_route_oracle.py` | Python oracle and VBA must stay behavior-identical |
| Change workbook sheets/headers/seeds | `scripts/workbook_contract.py`, `scripts/build_xlsm.py`, `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py` | Build seed, VBA setup seed, and structure tests must stay in sync |
| Change workbook UX/navigation | `scripts/workbook_ux.py`, `tests/test_xlsm_structure.py`, `tests/test_workbook_usage_links.py` | Display/input-assist only; never alter route/parser values |
| Change request parser behavior | `tests/request_parser_oracle.py`, `vba/FirewallPolicyAutomation.bas`, parser tests | Header aliases and merged-cell behavior are mirrored |
| Change SECUI output | `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py`, docs | Batch and CLI split `대상방화벽` by `;` |
| Change build/release artifact shape | `scripts/build_xlsm.py`, `scripts/make_request_folder.py`, `.github/workflows/`, `tests/test_xlsm_structure.py` | CI/release are artifact-driven, not package-driven |
| Release | `.github/workflows/release.yml` | Tags `v*` build, test, verify, bundle, publish |

## Commands

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python -m pytest tests/test_route_oracle.py -v
./.venv/bin/python scripts/build_xlsm.py
./.venv/bin/python scripts/make_request_folder.py
```

Use committed `./.venv/bin/python`; dependency pins live in `requirements.txt` (`pyOpenVBA==2.0.0`, `openpyxl==3.1.5`, `pytest==9.0.3`). Build runs on Linux without Windows, Excel, PowerShell, or COM.

## Current Workbook Contract

- `requests` has exactly 25 columns. Row 1 is the cosmetic group band; row 2 is canonical headers; data starts row 3.
- Column 7 display header is `대상방화벽`; internal result field remains `target_firewalls`; values join with `;`.
- `방화벽경로` joins matched firewall rows with `>`.
- `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거` are display/debug outputs from the first/selected matched range rows.
- Operator input/reference sheets are `firewalls`, `firewall_ranges`, `settings`, `header_aliases`, `vendor_cli_templates`, `service_catalog`.
- `network_definitions`, `routing_paths`, zone graph BFS, and legacy fallback are removed. Do not reintroduce them.

## Route Algorithm

`firewall_ranges` is authoritative. A request row matches enabled range rows when:

1. `firewalls.enabled` and `firewall_ranges.enabled` are truthy (`Y`, `YES`, `TRUE`, `1`).
2. request source overlaps `source_cidr`.
3. request destination overlaps `destination_cidr`.
4. request direction and range direction match, or either side is `BOTH`/blank.
5. matches sort by `path_order`, row order, firewall name.

Statuses are `OK`, `NO_MATCH`, `DIRECTION_MISMATCH`, plus duplicate markers from merge post-processing.

## Build Pipeline

`scripts/build_xlsm.py` injects `vba/FirewallPolicyAutomation.bas` and `vba/FirewallRouteAnalysis.bas`, removes default `Module1`, seeds workbook sheets, applies UX formatting from `scripts/workbook_ux.py`, and preserves VBA. `tests/test_xlsm_structure.py` verifies module presence, sheet headers, Korean codepage handling, target-cell highlight, release workflow sheet expectations, and source-to-injected module parity. `tests/test_workbook_usage_links.py` verifies first-open usage-sheet quick links.

`dist/*.xlsm` artifacts are non-deterministic and not committed. `dist/README.txt` is tracked release-bundle text. CI/release rebuild binary artifacts from source.

## CI / Release

- Branch is `master`; there is no `main`.
- CI runs on push/PR to `master`: pytest, build, request-folder scaffold, artifact smoke verification.
- Release runs on `v*` tags: pytest, build, artifact verification, bundle zip, GitHub Release upload.
- Commit style in history is Conventional Commits (`feat:`, `fix:`, `docs:`, `ci:`, `chore:`).

## Gotchas

- VBA comments may be codepage-lossy in `vbaProject.bin`; user-visible workbook Korean cells are preserved. Structure tests compare normalized module source.
- LibreOffice macro execution is unavailable here. Tests validate Python oracle and inspect generated `.xlsm`; they do not execute VBA macros.
- `request-folder/*.xlsx` files are generated/sample binaries and often change by one byte after scaffold/test runs. Do not include that churn unless intentionally updating samples.
- `firewall-policy-automation.xlsx` and `request-folder/**/*.xlsx` are tracked operator/sample binaries; treat changes as artifact updates, not source edits.
- If route logic changes, update oracle, VBA, and scenario tests in one change. Never touch only one side.
