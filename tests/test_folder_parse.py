"""신청서 상위 폴더명 파싱 (정보보호센터_1234 -> 팀 / 문서번호).

Mirrors the VBA ParseRequestFolderName helper. Rule:
  - split on the LAST underscore: left = team/center, right = doc_no
  - no underscore: whole = team, doc_no = ""
  - trailing/leading spaces trimmed
  - the full original folder name is also kept (request_folder)

Run: .venv/bin/python -m pytest tests/test_folder_parse.py -v
"""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from folder_parse_oracle import parse_request_folder_name  # noqa: E402


def test_simple_team_docno():
    t, d = parse_request_folder_name("정보보호센터_1234")
    assert t == "정보보호센터"
    assert d == "1234"


def test_multi_underscore_uses_last():
    # team itself contains an underscore; doc no is the trailing token
    t, d = parse_request_folder_name("정보보호_2팀_5678")
    assert t == "정보보호_2팀"
    assert d == "5678"


def test_no_underscore():
    t, d = parse_request_folder_name("인프라팀")
    assert t == "인프라팀"
    assert d == ""


def test_trims_spaces():
    t, d = parse_request_folder_name("  정보보호센터 _ 1234  ")
    assert t == "정보보호센터"
    assert d == "1234"


def test_english_center():
    t, d = parse_request_folder_name("SecurityCenter_A-2026-001")
    assert t == "SecurityCenter"
    assert d == "A-2026-001"


def test_empty():
    t, d = parse_request_folder_name("")
    assert t == ""
    assert d == ""


def test_trailing_underscore():
    # "정보보호센터_" -> team=정보보호센터, doc=""
    t, d = parse_request_folder_name("정보보호센터_")
    assert t == "정보보호센터"
    assert d == ""
