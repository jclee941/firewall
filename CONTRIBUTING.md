# Contributing

This project ships an Excel-native firewall policy automation workbook and
matching Python CLI/GUI tooling. Keep changes small, evidence-backed, and
compatible with the workbook contract in `AGENTS.md`.

## Development Setup

Use the committed virtual environment when it is available:

```bash
./.venv/bin/python -m pytest tests/ -q
```

If you need to recreate it:

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
```

Install GUI dependencies only when working on the PySide GUI or native GUI
release bundle:

```bash
./.venv/bin/pip install -r requirements-gui.txt
```

## Common Checks

Run the narrowest relevant checks first, then the full suite before submitting:

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python scripts/build_xlsm.py
./.venv/bin/python scripts/make_request_folder.py
./.venv/bin/python scripts/secui_cli.py --workbook dist/firewall-policy-automation.xlsm --format text
```

Generated `dist/*.xlsm` files are release artifacts and are not committed.
`request-folder/**/*.xlsx` files may churn after scaffold/tests; do not include
that churn unless the sample files are intentionally being updated.

## Change Guidelines

- Keep the Python oracle, VBA macro behavior, and workbook tests in sync.
- Do not reintroduce removed legacy sheets such as `secui_batch`,
  `secui_policy_export`, `policy_analysis`, or `policy_summary`.
- Preserve the clean 14-column `requests` sheet contract.
- Put route diagnostics in `route_results`, not on `requests`.
- For route or SECUI output changes, include focused regression tests.
- Use Conventional Commit messages, matching the existing history
  (`feat:`, `fix:`, `docs:`, `test:`, `ci:`, `chore:`).

## Pull Requests

Before opening a pull request, include:

- A short summary of the user-visible behavior change.
- The verification commands you ran and their result.
- Any generated files intentionally changed.
- Screenshots only when the GUI/workbook visual behavior changed.
