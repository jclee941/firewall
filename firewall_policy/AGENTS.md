# FIREWALL_POLICY KNOWLEDGE BASE

## Overview

Pure-Python package mirroring VBA request-parsing/folder-parsing behavior, plus the GUI export bridge. Backs the `scripts/secui_cli.py` and `scripts/secui_gui.py` runtimes; the workbook still parses inside VBA at runtime.

## Where To Look

| Task | Location | Notes |
| --- | --- | --- |
| Change request-sheet parsing | `request_parser.py` | Line-for-line mirror of `../vba/FirewallPolicyAutomation.bas` (`HeaderKey`, `CanonicalHeaderName`, `FindHeaderRow`, `BuildHeaderMap`, `ValidateRequiredHeaders`, `CopyRequestRow`) |
| Change folder-name parsing | `folder_parse.py` | Mirror of VBA `ParseRequestFolderName`; split on first two `_` into team/doc_no/title |
| Change CLI export wiring | `gui_export.py` | Builds/runs `secui_cli` export commands; dispatches by `SourceMode`/`ExportFormat` |
| Verify parser parity | `../tests/request_parser_oracle.py`, `../tests/test_request_parsing.py`, `../tests/test_merged_cells.py` | Tests treat this package as the Python oracle |

## Conventions

- Python 3.12+ only: `type` aliases, `match`/`case` with `assert_never`, `from __future__ import annotations`.
- Public functions take raw `SheetRows` (1-based logical rows handled internally); no openpyxl objects except `sheet_to_filled_rows(ws, ...)`.
- `gui_export.py` re-enters via `scripts.secui_cli._run` / `_export_rows`; keep that contract when changing args.

## Anti-Patterns

- Do not let this package diverge from VBA behavior. Any parser/folder change here requires the matching VBA + test change in the same commit.
- Do not fill the header row or drop source rows during fill-down; merged data cells fill down, headers do not.
- Do not add `__main__`/CLI here. Entry points live in `../scripts/`; this stays a library (`__init__.py` is empty by design).
