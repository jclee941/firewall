# PROJECT KNOWLEDGE BASE

Generated: 2026-06-16
Commit: fc149aa
Branch: master

## Overview

Excel-native VBA tool for merging firewall-policy request spreadsheets and computing `대상방화벽` from user-defined firewall range rows. No PowerQuery; Linux build/test tooling produces the shipped macro-enabled `.xlsm`.

## Structure

```text
.
├── vba/                  # shipped VBA modules injected into vbaProject.bin
├── firewall_policy/       # pure-Python mirrors of VBA parsing; backs secui_cli/gui
├── scripts/              # Linux builders, SECUI CLI/GUI, release bundlers
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
| Generate SECUI CLI outside Excel | `scripts/secui_cli.py`, `scripts/secui_cli_runtime.py`, `scripts/secui_cli_seed.py`, `firewall_policy/request_parser.py` | Linux CLI mirrors the in-workbook SECUI conversion; grouped rule generation |
| Change SECUI GUI / standalone build | `scripts/secui_gui.py`, `firewall_policy/gui_export.py`, `scripts/build_standalone_gui.py` | PySide6 GUI; packaged native bundle is a release artifact |
| Change SECUI output | `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py`, docs | Batch and CLI split `대상방화벽` by `;` |
| Change build/release artifact shape | `scripts/build_xlsm.py`, `scripts/make_request_folder.py`, `.github/workflows/`, `tests/test_xlsm_structure.py` | CI/release are artifact-driven, not package-driven |
| Release | `.github/workflows/release.yml` | Tags `v*` build, test, verify, bundle, publish |

## Commands

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python -m pytest tests/test_route_oracle.py -v
./.venv/bin/python scripts/build_xlsm.py
./.venv/bin/python scripts/make_request_folder.py
./.venv/bin/python scripts/secui_cli.py --workbook dist/firewall-policy-automation.xlsm --format text
./.venv/bin/python scripts/secui_gui.py   # PySide6 GUI
```

Use committed `./.venv/bin/python`; pins in `requirements.txt` (`pyOpenVBA==2.0.0`, `openpyxl==3.1.5`, `pytest==9.0.3`) and `requirements-gui.txt` (`PySide6==6.11.1`). Build runs on Linux without Windows, Excel, PowerShell, or COM.

## Current Workbook Contract

- `requests` is a CLEAN 14-column user sheet: 요청부서, 요청번호, 대상방화벽, 출발지, 출발지설명, 목적지, 목적지설명, 프로토콜, 포트, 방향, 용도, 시작일, 종료일, 비고. Headers on row 2; data starts row 3.
- `대상방화벽` is column 3 in `requests`; internal result field remains `target_firewalls`; values join with `;`.
- Source tracking (`원본파일`, `원본행`, `요청폴더`, `제목`) lives in the hidden `_request_tracking` sheet (keyed by `request_row`), NOT on `requests`.
- Route diagnostics (`검증상태`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거`, `원본파일`, `원본행`) live on `route_results`, not `requests`.
- Operator input/reference sheets are `firewalls`, `firewall_ranges`, `settings`, `header_aliases`, `vendor_cli_templates`, `service_catalog`.
- `network_definitions`, `routing_paths`, zone graph BFS, and legacy fallback are removed. Do not reintroduce them.
- Direction values normalize to `IN`/`OUT`/`BOTH`/`#INVALID`; blanks mean `BOTH`. Korean/business synonyms (인바운드/수신/외부->내부 = IN; 아웃바운드/송신/내부->외부 = OUT; 양방향/ANY = BOTH) are accepted; standalone 내부/외부 stay `#INVALID`. Mirror any change in `tests/route_oracle.py` + `vba/FirewallRouteAnalysis.bas` + `scripts/secui_cli_seed.py`.

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
- `request-folder/**/*.xlsx` files are generated/sample binaries and often change by one byte after scaffold/test runs. Do not include that churn unless intentionally updating samples.
- If route logic changes, update oracle, VBA, and scenario tests in one change. Never touch only one side.
