import glob
import os
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT_PATH = Path(ROOT)
PY = sys.executable

sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, ROOT)
from tests.request_parser_oracle import parse_request_sheet, sheet_to_filled_rows  # noqa: E402


@pytest.fixture(scope="module")
def request_tree(tmp_path_factory):
    out_dir = tmp_path_factory.mktemp("request-folder-output") / "request-folder"
    subprocess.run(
        [PY, os.path.join(ROOT, "scripts", "make_request_folder.py"), "--output", str(out_dir)],
        cwd=ROOT, check=True, capture_output=True, text=True,
    )
    assert out_dir.is_dir()
    return str(out_dir)


def test_help_exposes_output_without_writing_default_folder():
    default_xlsx = sorted((ROOT_PATH / "request-folder").glob("**/*.xlsx"))
    assert default_xlsx, "default request-folder workbooks missing"
    before = {path: path.stat().st_mtime_ns for path in default_xlsx}

    result = subprocess.run(
        [PY, os.path.join(ROOT, "scripts", "make_request_folder.py"), "--help"],
        cwd=ROOT, check=True, capture_output=True, text=True,
    )

    assert "--output" in result.stdout
    assert {path: path.stat().st_mtime_ns for path in default_xlsx} == before


def test_tree_structure(request_tree):
    assert os.path.isfile(os.path.join(request_tree, "README.txt"))
    xlsx = glob.glob(os.path.join(request_tree, "**", "*.xlsx"), recursive=True)
    assert len(xlsx) >= 3, xlsx
    folders = {os.path.basename(os.path.dirname(f)) for f in xlsx}
    # team folders must follow <team>_<docno> so folder parsing yields team/doc_no
    assert any("_" in f and not f.startswith("_") for f in folders), folders
    # a header-only template must exist for operators to copy
    assert any("빈양식" in os.path.basename(f) for f in xlsx), xlsx


def test_empty_template_parses_to_zero_rows(request_tree):
    tmpl = glob.glob(os.path.join(request_tree, "**", "*빈양식*.xlsx"), recursive=True)
    assert tmpl, "blank template missing"
    ws = openpyxl.load_workbook(tmpl[0]).active
    parsed = parse_request_sheet(sheet_to_filled_rows(ws))
    assert parsed == [], "blank template must yield no data rows (header only)"


def test_each_request_parses_with_cli_targets(request_tree):
    xlsx = sorted(glob.glob(os.path.join(request_tree, "**", "*.xlsx"), recursive=True))
    data_files = [f for f in xlsx if "빈양식" not in os.path.basename(f)]
    assert data_files
    total_rows = 0
    multi_fw = 0
    for f in data_files:
        ws = openpyxl.load_workbook(f).active
        parsed = parse_request_sheet(sheet_to_filled_rows(ws))
        assert parsed, f"no rows parsed from {f}"
        total_rows += len(parsed)
        for req in parsed:
            targets = [x for x in str(req["target_firewalls"] or "").split(";") if x]
            assert targets, f"{f}: {req} has no target_firewalls"
            assert all(target.startswith("SECUI-FW-") for target in targets)
            if len(targets) >= 2:
                multi_fw += 1
    assert total_rows >= 4
    assert multi_fw >= 2, f"expected >=2 multi-firewall requests, got {multi_fw}"
