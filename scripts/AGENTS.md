# SCRIPTS KNOWLEDGE BASE

## Overview

Linux-only builders for the shipped Excel workbook and sample request-folder tree. These scripts are executable project surfaces, not a Python package.

## Where To Look

| Task | Location | Notes |
| --- | --- | --- |
| Build `.xlsm` workbook | `build_xlsm.py` | Injects `vba/*.bas`, seeds sheets, formats UX, preserves macros |
| Change workbook headers/sheets | `build_xlsm.py`, `../vba/FirewallPolicyAutomation.bas`, `../tests/test_xlsm_structure.py` | Python constants and VBA `Write*` routines must agree |
| Change route seed examples | `build_xlsm.py`, `../tests/test_xlsm_structure.py`, `../tests/route_oracle.py` | Seeded `requests` rows must resolve through the oracle |
| Change request-folder samples | `make_request_folder.py`, `../tests/test_request_folder.py`, `../tests/test_folder_parse.py` | Folder names encode team/doc number |
| Change UX/protection behavior | `build_xlsm.py`, `../tests/test_xlsm_structure.py` | Display/input-assist only; must not alter route semantics |
| Change SECUI service examples | `workbook_contract.py`, `build_xlsm.py`, `../vba/FirewallPolicyAutomation.bas`, `../tests/test_secui_service_catalog.py` | Reference data only; do not change route logic |

## Commands

```bash
./.venv/bin/python scripts/build_xlsm.py
./.venv/bin/python scripts/make_request_folder.py
./.venv/bin/python -m pytest tests/test_xlsm_structure.py -q
```

Use the repo venv. There is no `pyproject.toml`; dependencies are pinned in `requirements.txt`.

## Build Contract

- `build_xlsm.py` writes `dist/firewall-policy-automation.xlsm`.
- It injects `FirewallPolicyAutomation` and `FirewallRouteAnalysis`, removes `Module1`, and expects exact source parity with `vba/*.bas`.
- It patches/preserves VBA code page 949 so Korean literals survive in the macro project.
- `requests` row 1 is a cosmetic group band, row 2 is canonical headers, row 3 starts data.
- `service_catalog` is a SECUI operator reference sheet; it is not an algorithm input.
- `_UX_LAST_ROW` bounds validation and conditional formatting; never expand UX ranges to whole columns.

## Anti-Patterns

- Do not add Windows Excel, COM, PowerShell, or LibreOffice runtime requirements.
- Do not protect `requests` or `processing_log`; VBA writes them at runtime.
- Do not let UX formatting mutate request values or route results.
- Do not update sheet/header seeds without matching VBA setup and structure tests.
- Do not commit generated `.xlsm` churn from local builds.

## Gotchas

- `make_request_folder.py` deletes and rebuilds `request-folder/`; tracked `.xlsx` samples can churn after tests.
- Sample request workbooks intentionally put the first labeled header, `No`, in column B.
- CI and release both rely on these scripts, so CLI behavior and output paths are part of the contract.
