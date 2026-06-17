from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import openpyxl

from scripts.secui_cli_runtime import RequestRecord, secui_cli_rows
from scripts.workbook_contract import FIREWALLS, FIREWALL_RANGES, VENDOR_CLI_TEMPLATES

ROOT = Path(__file__).resolve().parents[1]
PY = sys.executable
XLSM = ROOT / "dist" / "firewall-policy-automation.xlsm"


def test_secui_cli_script_generates_text_from_workbook(tmp_path: Path) -> None:
    subprocess.run([PY, str(ROOT / "scripts" / "build_xlsm.py")], cwd=ROOT, check=True)
    output = tmp_path / "secui.txt"

    result = subprocess.run(
        [PY, str(ROOT / "scripts" / "secui_cli.py"), "--workbook", str(XLSM), "--output", str(output)],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    text = output.read_text(encoding="utf-8")
    assert "SECUI-FW-01" in text
    assert "fw set srule" in text
    assert "fw set addrgrp" in text


def test_secui_cli_script_generates_xlsx_from_request_folder(tmp_path: Path) -> None:
    subprocess.run([PY, str(ROOT / "scripts" / "make_request_folder.py")], cwd=ROOT, check=True)
    text_output = tmp_path / "secui.txt"
    xlsx_output = tmp_path / "secui.xlsx"

    result = subprocess.run(
        [
            PY,
            str(ROOT / "scripts" / "secui_cli.py"),
            "--request-folder",
            str(ROOT / "request-folder"),
            "--output",
            str(text_output),
            "--output-xlsx",
            str(xlsx_output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert text_output.read_text(encoding="utf-8").count("fw set srule") >= 3
    wb = openpyxl.load_workbook(xlsx_output)
    try:
        assert "secui_cli" in wb.sheetnames
        assert wb["secui_cli"].max_row >= 4
    finally:
        wb.close()


def test_secui_cli_export_command_writes_xlsx_from_workbook(tmp_path: Path) -> None:
    subprocess.run([PY, str(ROOT / "scripts" / "build_xlsm.py")], cwd=ROOT, check=True)
    output = tmp_path / "secui-export.xlsx"

    result = subprocess.run(
        [
            PY,
            str(ROOT / "scripts" / "secui_cli.py"),
            "export",
            "--workbook",
            str(XLSM),
            "--format",
            "xlsx",
            "--output",
            str(output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    wb = openpyxl.load_workbook(output)
    try:
        ws = wb["secui_cli"]
        commands = [str(row[3].value or "") for row in ws.iter_rows(min_row=2)]
        assert any("fw set srule" in command for command in commands)
        assert any("fw set addrgrp" in command for command in commands)
    finally:
        wb.close()


def test_secui_cli_rows_derive_target_firewalls_when_parsed_request_has_blank_target() -> None:
    requests: list[RequestRecord] = [
        {
            "요청부서": "정보보호센터",
            "요청번호": "AUTO",
            "원본파일": "blank-target.xlsx",
            "원본행": 2,
            "대상방화벽": "",
            "출발지IP": "10.10.10.5",
            "출발지설명": "업무PC",
            "목적지IP": "10.20.20.5",
            "목적지설명": "DMZ서버",
            "프로토콜": "TCP",
            "포트": "443",
            "방향": "OUT",
            "용도": "자동 경로 탐색",
            "비고": "대상방화벽 공란",
        },
    ]

    rows = secui_cli_rows(requests, FIREWALLS, FIREWALL_RANGES, VENDOR_CLI_TEMPLATES)

    firewalls = [str(row[1]) for row in rows[1:]]
    assert firewalls == ["SECUI-FW-01", "SECUI-FW-02"]
    commands = "\n".join(str(row[3]) for row in rows[1:])
    assert "# device=SECUI-FW-01" in commands
    assert "# device=SECUI-FW-02" in commands
    assert "fw set srule" in commands


# ---------------------------------------------------------------------------
# SECUI address-overlap parity with route oracle + VBA SecuiAddressListOverlaps
# ---------------------------------------------------------------------------
# The Python SECUI mirror must match VBA SecuiAddressListOverlaps and the route
# oracle _address_list_overlaps: a BLANK request matches NOTHING (even against an
# ANY/blank definition), and the splitter must normalize NBSP / tab / fullwidth
# comma+semicolon exactly like route_oracle.split_address_list.

import pytest

from scripts.secui_cli_seed import _overlaps, _split_values
from tests.route_oracle import split_address_list


@pytest.mark.parametrize("any_def", ["ANY", "", "*", "ALL", "0.0.0.0/0"])
def test_secui_overlaps_blank_request_never_matches_any_definition(any_def: str) -> None:
    # Mirrors VBA SecuiAddressListOverlaps (blank guard BEFORE the ANY short-circuit)
    # and route_oracle._address_list_overlaps. An incomplete request must not route.
    assert _overlaps("", any_def) is False
    assert _overlaps("   ", any_def) is False
    # a complete request still matches an ANY definition
    assert _overlaps("10.10.10.5", any_def) is True


def test_secui_matched_firewalls_skips_request_with_blank_source_against_any_range() -> None:
    # A range whose source is ANY would (buggily) match a blank 출발지IP. VBA and the
    # route oracle reject the blank request, so the SECUI CLI must emit no firewall.
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
        "10.0.0.1\u00a010.0.0.2",   # non-breaking space
        "10.0.0.1\t10.0.0.2",         # tab
        "10.0.0.1\uff0c10.0.0.2",     # fullwidth comma
        "10.0.0.1\uff1b10.0.0.2",     # fullwidth semicolon
    ],
)
def test_secui_split_values_matches_route_oracle_splitter(text: str) -> None:
    # _split_values must normalize the same separators as VBA NormalizeListCell and
    # route_oracle.split_address_list so SECUI matching does not silently miss tokens.
    assert _split_values(text) == split_address_list(text)
    assert _split_values(text) == ["10.0.0.1", "10.0.0.2"]


@pytest.mark.parametrize("bad_ip", ["2001:db8::1", "010.000.000.001", "10.0.0.256"])
def test_secui_cli_skips_invalid_request_even_with_prefilled_target(bad_ip: str) -> None:
    # SECURITY: a request whose source/destination is not strict IPv4 must emit
    # NO SECUI rule, EVEN when 대상방화벽 is prefilled/stale. The route analyzer
    # would flag it INVALID_ADDRESS; the CLI must not emit unverifiable traffic.
    requests: list[RequestRecord] = [
        {
            "요청부서": "정보보호센터",
            "요청번호": "AUTO",
            "원본파일": "prefilled.xlsx",
            "원본행": 2,
            "대상방화벽": "SECUI-FW-01;SECUI-FW-02",  # prefilled / stale target
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
    # SECURITY: a request with a BLANK source or destination is non-routable (the
    # route analyzer returns NO_MATCH). The SECUI CLI must emit NO rule even when
    # 대상방화벽 is prefilled/stale -- otherwise it emits a bogus src/dst "ANY" rule
    # for an incomplete request. Same failure class as the invalid-address skip.
    base: RequestRecord = {
        "요청부서": "정보보호센터",
        "요청번호": "AUTO",
        "원본파일": "prefilled.xlsx",
        "원본행": 2,
        "대상방화벽": "SECUI-FW-01;SECUI-FW-02",  # prefilled / stale target
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
    # S6 parity: the in-workbook SECUI converter (FirewallPolicyAutomation.bas)
    # must also skip rows whose source/destination is not strict IPv4, mirroring
    # secui_cli_runtime._request_has_invalid_address. Otherwise a stale 대상방화벽
    # lets ConvertRequestsToSecuiCli emit a rule the route analyzer rejected.
    import os

    root = Path(__file__).resolve().parents[1]
    src = open(os.path.join(root, "vba", "FirewallPolicyAutomation.bas"), encoding="utf-8").read()
    assert "SecuiIsInvalidAddress" in src or "SecuiFirstInvalidToken" in src, \
        "SECUI converter must have a strict-IPv4 invalid-address helper"
    # the strict octet parser must be reused, not the old lax CDbl path
    assert "SecuiParseOctet" in src, "SecuiIpToNumber must use a strict octet parser"
    assert "CDbl(parts(index))" not in src, "old lax SecuiIpToNumber CDbl path must be gone"
    # EVERY SECUI emission path that reads the prefilled 대상방화벽 must guard
    # invalid-address requests before emitting: the live fanout/collect path AND
    # the tracked CopySecuiCliRows surface.
    def _body(start_marker: str, end_marker: str) -> str:
        seg = src[src.find(start_marker):]
        return seg[:seg.find(end_marker)]

    fanout = _body("Private Function BuildSecuiCliServiceFanoutIndex", "End Function")
    collect = _body("Private Sub CollectSecuiCliRows", "End Sub")
    copy = _body("Private Function CopySecuiCliRows", "End Function")
    assert "SecuiRequestHasInvalidAddress" in fanout, \
        "BuildSecuiCliServiceFanoutIndex must skip invalid-address requests"
    assert "SecuiRequestHasInvalidAddress" in collect, \
        "CollectSecuiCliRows must skip invalid-address requests"
    assert "SecuiRequestHasInvalidAddress" in copy, \
        "CopySecuiCliRows must skip invalid-address requests"
    # The guard must reject BLANK addresses too, not just malformed ones, so a
    # stale/prefilled target on an incomplete request cannot emit a bogus ANY rule.
    guard = _body("Private Function SecuiRequestHasInvalidAddress", "End Function")
    assert "NormalizeListCell(sourceAddress)" in guard and "NormalizeListCell(destinationAddress)" in guard, \
        "SecuiRequestHasInvalidAddress must treat a blank source/destination as non-routable"
    assert "Len(Trim$(NormalizeListCell(sourceAddress))) = 0" in guard, \
        "SecuiRequestHasInvalidAddress must short-circuit True on a blank address"


# ---------------------------------------------------------------------------
# Operational robustness: a corrupt workbook or one bad file in a batch must not
# crash the CLI with a stack trace or abort the whole run.
# ---------------------------------------------------------------------------


def test_secui_cli_corrupt_workbook_fails_cleanly(tmp_path: Path) -> None:
    # F1: a corrupt/non-xlsx workbook must produce a clean rc=2 error message,
    # not an uncaught zipfile.BadZipFile / InvalidFileException stack trace.
    bad = tmp_path / "garbage.xlsm"
    bad.write_bytes(b"this is not a zip file")
    result = subprocess.run(
        [PY, str(ROOT / "scripts" / "secui_cli.py"), "--workbook", str(bad), "--format", "text"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 2, result.stderr
    assert "Traceback" not in result.stderr, f"must not leak a stack trace: {result.stderr}"
    assert "ERROR" in result.stderr


def test_secui_cli_folder_skips_one_bad_file_and_keeps_good(tmp_path: Path) -> None:
    # F2: a folder with one unparseable .xlsx and one good one must NOT abort the
    # whole batch. The good file's rules must still be emitted; the bad file is
    # skipped with a warning. (Single bad file aborting the batch was the bug.)
    good_dir = tmp_path / "정보보호센터_1234"
    good_dir.mkdir(parents=True)
    good = good_dir / "신청.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["No", "출발지IP", "목적지IP", "프로토콜", "포트", "방향"])
    ws.append([1, "10.10.10.5", "10.20.20.5", "TCP", "443", "OUT"])
    wb.save(good)
    wb.close()

    bad_dir = tmp_path / "의미없는팀_9999"
    bad_dir.mkdir(parents=True)
    bad = bad_dir / "깨진신청.xlsx"
    bad.write_bytes(b"definitely not a zip")

    text_output = tmp_path / "secui.txt"
    result = subprocess.run(
        [
            PY, str(ROOT / "scripts" / "secui_cli.py"),
            "--request-folder", str(tmp_path),
            "--output", str(text_output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
    text = text_output.read_text(encoding="utf-8")
    assert "fw set srule" in text, "the good request file's rule must still be emitted"
    assert "깨진신청.xlsx" in result.stderr or "WARNING" in result.stderr, \
        "the skipped bad file should be reported as a warning"


def test_gui_export_catches_broad_failure_surface() -> None:
    # F3: the GUI _export handler must not crash the window on the common error
    # classes that are NOT ValueError/FileNotFoundError: RequestParseError (a bad
    # request file), corrupt-workbook (BadZipFile / InvalidFileException), and OSError.
    import os

    root = Path(__file__).resolve().parents[1]
    src = open(os.path.join(root, "scripts", "secui_gui.py"), encoding="utf-8").read()
    export = src[src.find("def _export"):]
    export = export[:export.find("\n    def ", 1)]
    assert "RequestParseError" in export, "_export must catch RequestParseError (bad request file)"
    assert "BadZipFile" in export or "InvalidFileException" in export, \
        "_export must catch corrupt-workbook errors"
    # zip-valid-but-bad xlsx raises KeyError / xml ParseError -- these must be in
    # the intentional 'unreadable' path, not only the last-resort Exception net.
    assert "KeyError" in export and "ParseError" in export, \
        "_export must catch zip-valid-but-malformed workbook errors (KeyError/ParseError)"
    assert "OSError" in export, "_export must catch I/O errors"
    assert "except Exception" in export, "_export must have a last-resort safety net"
    # ordering: FileNotFoundError (an OSError subclass) must be caught BEFORE the
    # broad OSError clause, else it would be shadowed.
    assert export.find("except FileNotFoundError") < export.find("except OSError"), \
        "FileNotFoundError must be handled before the broad OSError clause"
    # and the last-resort Exception must come last.
    assert export.find("except OSError") < export.find("except Exception"), \
        "the last-resort Exception net must be the final handler"


import zipfile as _zipfile


def _make_zip_valid_but_bad_xlsx(path: Path) -> None:
    # A valid zip that is NOT a valid xlsx (missing OOXML parts). openpyxl raises
    # KeyError ('no item named [Content_Types].xml') -- NOT BadZipFile.
    with _zipfile.ZipFile(path, "w") as z:
        z.writestr("hello.txt", "not an xlsx")


def test_secui_cli_zip_valid_but_bad_xlsx_fails_cleanly(tmp_path: Path) -> None:
    # F1 (broadened): a zip-valid-but-structurally-invalid .xlsx must still exit
    # rc=2 with a clean message, not an uncaught KeyError/ParseError traceback.
    bad = tmp_path / "bad-ooxml.xlsx"
    _make_zip_valid_but_bad_xlsx(bad)
    result = subprocess.run(
        [PY, str(ROOT / "scripts" / "secui_cli.py"), "--workbook", str(bad), "--format", "text"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 2, result.stderr
    assert "Traceback" not in result.stderr, f"must not leak a stack trace: {result.stderr}"
    assert "ERROR" in result.stderr


def test_secui_cli_folder_skips_zip_valid_but_bad_xlsx(tmp_path: Path) -> None:
    # F2 (broadened): a zip-valid-but-bad .xlsx in the batch must be skipped, not
    # abort the whole run with an uncaught KeyError.
    good_dir = tmp_path / "정보보호센터_1234"
    good_dir.mkdir(parents=True)
    good = good_dir / "신청.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["No", "출발지IP", "목적지IP", "프로토콜", "포트", "방향"])
    ws.append([1, "10.10.10.5", "10.20.20.5", "TCP", "443", "OUT"])
    wb.save(good)
    wb.close()
    bad_dir = tmp_path / "깨진팀_9999"
    bad_dir.mkdir(parents=True)
    _make_zip_valid_but_bad_xlsx(bad_dir / "깨진.xlsx")
    text_output = tmp_path / "secui.txt"
    result = subprocess.run(
        [
            PY, str(ROOT / "scripts" / "secui_cli.py"),
            "--request-folder", str(tmp_path),
            "--output", str(text_output),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
    assert "fw set srule" in text_output.read_text(encoding="utf-8")
    assert "WARNING" in result.stderr
