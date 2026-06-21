from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

from scripts.secui_cli_runtime import RequestRecord, secui_cli_rows
from scripts.secui_cli_seed import _overlaps, _split_values
from scripts.workbook_contract import FIREWALLS, FIREWALL_RANGES, VENDOR_CLI_TEMPLATES
from tests.route_oracle import split_address_list

ROOT = Path(__file__).resolve().parents[1]
PY = sys.executable


@pytest.mark.parametrize("any_def", ["ANY", "", "*", "ALL", "0.0.0.0/0"])
def test_secui_overlaps_blank_request_never_matches_any_definition(any_def: str) -> None:
    assert _overlaps("", any_def) is False
    assert _overlaps("   ", any_def) is False
    assert _overlaps("10.10.10.5", any_def) is True


def test_secui_matched_firewalls_skips_request_with_blank_source_against_any_range() -> None:
    firewalls = [
        ["firewall_name", "vendor", "enabled", "comment"],
        ["SECUI-FW-01", "SECUI", "Y", "any-source"],
    ]
    firewall_ranges = [
        ["firewall_name", "source_cidr", "destination_cidr", "direction", "path_order", "enabled", "comment"],
        ["SECUI-FW-01", "ANY", "10.20.0.0/16", "OUT", 10, "Y", "any -> dmz"],
    ]
    blank_source: list[RequestRecord] = [
        {
            "요청부서": "정보보호센터",
            "요청번호": "AUTO",
            "원본파일": "blank-source.xlsx",
            "원본행": 2,
            "대상방화벽": "",
            "출발지IP": "",
            "목적지IP": "10.20.20.5",
            "프로토콜": "TCP",
            "포트": "443",
            "방향": "OUT",
        },
    ]

    rows = secui_cli_rows(blank_source, firewalls, firewall_ranges, VENDOR_CLI_TEMPLATES)

    assert rows[1:] == []


@pytest.mark.parametrize(
    "text",
    [
        "10.0.0.1\u00a010.0.0.2",
        "10.0.0.1\t10.0.0.2",
        "10.0.0.1\uff0c10.0.0.2",
        "10.0.0.1\uff1b10.0.0.2",
    ],
)
def test_secui_split_values_matches_route_oracle_splitter(text: str) -> None:
    assert _split_values(text) == split_address_list(text)
    assert _split_values(text) == ["10.0.0.1", "10.0.0.2"]


@pytest.mark.parametrize("bad_ip", ["2001:db8::1", "010.000.000.001", "10.0.0.256"])
def test_secui_cli_skips_invalid_request_even_with_prefilled_target(bad_ip: str) -> None:
    requests: list[RequestRecord] = [
        {
            "요청부서": "정보보호센터",
            "요청번호": "AUTO",
            "원본파일": "prefilled.xlsx",
            "원본행": 2,
            "대상방화벽": "SECUI-FW-01;SECUI-FW-02",
            "출발지IP": bad_ip,
            "목적지IP": "10.20.20.5",
            "프로토콜": "TCP",
            "포트": "443",
            "방향": "OUT",
        },
    ]
    rows = secui_cli_rows(requests, FIREWALLS, FIREWALL_RANGES, VENDOR_CLI_TEMPLATES)
    assert rows[1:] == [], f"invalid request {bad_ip} must emit no SECUI rule"


@pytest.mark.parametrize("blank_field", ["출발지IP", "목적지IP"])
def test_secui_cli_skips_blank_address_request_even_with_prefilled_target(blank_field: str) -> None:
    base: RequestRecord = {
        "요청부서": "정보보호센터",
        "요청번호": "AUTO",
        "원본파일": "prefilled.xlsx",
        "원본행": 2,
        "대상방화벽": "SECUI-FW-01;SECUI-FW-02",
        "출발지IP": "10.10.10.5",
        "목적지IP": "10.20.20.5",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
    }
    base[blank_field] = ""
    rows = secui_cli_rows([base], FIREWALLS, FIREWALL_RANGES, VENDOR_CLI_TEMPLATES)
    assert rows[1:] == [], f"blank {blank_field} must emit no SECUI rule even with prefilled target"


def test_vba_secui_converter_skips_invalid_request_addresses() -> None:
    src = (ROOT / "vba" / "FirewallPolicyAutomation.bas").read_text(encoding="utf-8")
    assert "SecuiIsInvalidAddress" in src or "SecuiFirstInvalidToken" in src
    assert "SecuiParseOctet" in src
    assert "CDbl(parts(index))" not in src

    def body(start_marker: str, end_marker: str) -> str:
        segment = src[src.find(start_marker):]
        return segment[:segment.find(end_marker)]

    fanout = body("Private Function BuildSecuiCliServiceFanoutIndex", "End Function")
    collect = body("Private Sub CollectSecuiCliRows", "End Sub")
    copy = body("Private Function CopySecuiCliRows", "End Function")
    assert "SecuiRequestHasInvalidAddress" in fanout
    assert "SecuiRequestHasInvalidAddress" in collect
    assert "SecuiRequestHasInvalidAddress" in copy

    guard = body("Private Function SecuiRequestHasInvalidAddress", "End Function")
    assert "NormalizeListCell(sourceAddress)" in guard
    assert "NormalizeListCell(destinationAddress)" in guard
    assert "Len(Trim$(NormalizeListCell(sourceAddress))) = 0" in guard


def test_secui_cli_corrupt_workbook_fails_cleanly(tmp_path: Path) -> None:
    bad = tmp_path / "garbage.xlsm"
    bad.write_bytes(b"this is not a zip file")
    result = subprocess.run(
        [PY, str(ROOT / "scripts" / "secui_cli.py"), "--workbook", str(bad), "--format", "text"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 2, result.stderr
    assert "Traceback" not in result.stderr
    assert "ERROR" in result.stderr


def test_secui_cli_folder_skips_one_bad_file_and_keeps_good(tmp_path: Path) -> None:
    good_dir = tmp_path / "정보보호센터_1234"
    good_dir.mkdir(parents=True)
    write_minimal_request(good_dir / "신청.xlsx")

    bad_dir = tmp_path / "의미없는팀_9999"
    bad_dir.mkdir(parents=True)
    (bad_dir / "깨진신청.xlsx").write_bytes(b"definitely not a zip")

    text_output = tmp_path / "secui.txt"
    result = run_folder_export(tmp_path, text_output)
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
    assert "fw set srule" in text_output.read_text(encoding="utf-8")
    assert "깨진신청.xlsx" in result.stderr or "WARNING" in result.stderr


def test_gui_export_catches_broad_failure_surface() -> None:
    src = (ROOT / "scripts" / "secui_gui.py").read_text(encoding="utf-8")
    export = src[src.find("def _export"):]
    export = export[:export.find("\n    def ", 1)]
    assert "RequestParseError" in export
    assert "BadZipFile" in export or "InvalidFileException" in export
    assert "KeyError" in export and "ParseError" in export
    assert "OSError" in export
    assert "except Exception" in export
    assert export.find("except FileNotFoundError") < export.find("except OSError")
    assert export.find("except OSError") < export.find("except Exception")


def test_secui_cli_zip_valid_but_bad_xlsx_fails_cleanly(tmp_path: Path) -> None:
    bad = tmp_path / "bad-ooxml.xlsx"
    write_zip_valid_bad_xlsx(bad)
    result = subprocess.run(
        [PY, str(ROOT / "scripts" / "secui_cli.py"), "--workbook", str(bad), "--format", "text"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 2, result.stderr
    assert "Traceback" not in result.stderr
    assert "ERROR" in result.stderr


def test_secui_cli_folder_skips_zip_valid_but_bad_xlsx(tmp_path: Path) -> None:
    good_dir = tmp_path / "정보보호센터_1234"
    good_dir.mkdir(parents=True)
    write_minimal_request(good_dir / "신청.xlsx")

    bad_dir = tmp_path / "깨진팀_9999"
    bad_dir.mkdir(parents=True)
    write_zip_valid_bad_xlsx(bad_dir / "깨진.xlsx")

    text_output = tmp_path / "secui.txt"
    result = run_folder_export(tmp_path, text_output)
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
    assert "fw set srule" in text_output.read_text(encoding="utf-8")
    assert "WARNING" in result.stderr


def write_minimal_request(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["No", "출발지IP", "목적지IP", "프로토콜", "포트", "방향"])
    ws.append([1, "10.10.10.5", "10.20.20.5", "TCP", "443", "OUT"])
    wb.save(path)
    wb.close()


def write_zip_valid_bad_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("hello.txt", "not an xlsx")


def run_folder_export(folder: Path, text_output: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            PY, str(ROOT / "scripts" / "secui_cli.py"),
            "--request-folder", str(folder),
            "--output", str(text_output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
