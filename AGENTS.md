# AGENTS.md

Excel-native (VBA) tool that merges firewall-policy request spreadsheets and computes which firewalls a src→dst flow traverses, via zone resolution + routing-graph BFS. No PowerQuery (DRM constraint). Build/test tooling is Python; the shipped artifact is a macro-enabled `.xlsm`.

## Commands

```bash
./.venv/bin/python -m pytest tests/ -q     # run the algorithm spec tests
./.venv/bin/python -m pytest tests/test_route_oracle.py::test_scenario1_ok_single_hop -v   # single test
./.venv/bin/python scripts/build_xlsm.py   # build dist/firewall-policy-automation.xlsm
```

Use the committed venv `./.venv/bin/python` (pins match `requirements.txt`: pyOpenVBA 2.0.0, openpyxl 3.1.5, pytest 9.0.3). Build runs on Linux — no Windows/Excel/PowerShell/COM.

## Architecture: the two sources of truth must stay in sync

- `tests/route_oracle.py` is the **binding contract / specification** for the route algorithm. Its module docstring states the exact rules (longest-prefix zone match, deterministic BFS, tie-break key `"{path_order:06d}|{firewall_name}|{to_zone}"`, status set, direction handling, legacy fallback).
- `vba/FirewallRouteAnalysis.bas` **must mirror `route_oracle.py` exactly**. The Python tests double as the VBA spec.
- `vba/FirewallPolicyAutomation.bas` is the second module: folder merge, header parsing, merged-cell handling, sheet setup macros.
- `tests/test_route_oracle.py` are scenario tests *against the oracle* — they do not execute VBA. LibreOffice macro execution is not available here, so changing algorithm behavior means editing oracle + `.bas` together and updating these tests.

When you change route logic: edit `route_oracle.py`, `FirewallRouteAnalysis.bas`, and the scenario tests in one change. Don't touch only one.

## Build pipeline (scripts/build_xlsm.py)

pyOpenVBA injects both `.bas` modules into a real `vbaProject.bin` and deletes the default `Module1`; openpyxl then seeds all sheets while preserving the VBA. Seed data (FIREWALLS, NETWORK_DEFS, ROUTING_PATHS, SETTINGS, REQUESTS_HEADERS, etc.) lives as Python constants in `build_xlsm.py` and **must stay in sync with the VBA `Write*Headers` seeds** (incl. the `settings` sheet, which is `key,value,설명` in BOTH `build_xlsm.py` SETTINGS and VBA `WriteSettings`). `tests/test_xlsm_structure.py` locks per-sheet schemas and round-trips the injected VBA back against `vba/*.bas` (Korean comments are codepage-lossy on read, so that test compares ASCII-normalized source).

CI/release assert structural invariants you must not silently break:
- `requests` sheet has exactly **24 columns** (`wb["requests"].max_column == 24`).
- Modules `FirewallPolicyAutomation` and `FirewallRouteAnalysis` present, `Module1` absent.
- All 8 sheets present with exact header rows (asserted by `test_sheets_and_headers`).
- Injected VBA module source must match `vba/*.bas` (`test_injected_vba_modules_match_source_files`).

## CI / release

- `.github/workflows/ci.yml`: triggers on **`master`** only (push + PR). Runs pytest → build → artifact-structure verification.
- `.github/workflows/release.yml`: triggers on `v*` tags. Build + test + verify, then `gh release` with the `.xlsm` attached. Release: `git tag v1.0.2 && git push origin v1.0.2`.
- Single-branch repo: `master` (no `main`). Commits follow Conventional Commits (`fix:`, `chore:`, `ci:`, `docs:`, `review:`).

## Conventions / gotchas

- Zone path separator inside the algorithm/oracle is `>` (e.g. `RED>GREEN>BLUE`), and `target_firewalls` joins with `;`. (README prose sometimes shows `->` for readability — trust the oracle/tests.)
- `dist/` artifact is **not committed** (non-deterministic timestamp); CI rebuilds from source. Don't commit built `.xlsm` files.
- VBA **source comments** in Korean are codepage-lossy in the built `vbaProject.bin` (pyOpenVBA template uses a Western code page; `한글` comments may show as `?` in the VBE). Cosmetic only — VBA code/identifiers are ASCII and run correctly; user-visible **sheet** Korean (seeded cells) is proper UTF-8 and intact. The round-trip test compares ASCII-normalized source for this reason.
- `route_legacy_fallback` is always `FALSE` (default and standing policy — never enable it). Unresolved routes must stay `NO_PATH`/`ZONE_UNRESOLVED`; the CIDR-overlap fallback and `LEGACY_FALLBACK` status are kept only as dormant legacy code, not used in operation.
- Detailed schema/operations docs: `docs/excel-schema.md`, `docs/excel-native.md`. Full data-input rules and `validation_status` meanings: `README.md`.
