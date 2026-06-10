"""신청서 상위 폴더명 파싱 (인프라시너지셀_2026-782_제목 -> 부서 / 번호 / 제목).

Mirrors the VBA ParseRequestFolderName helper. Rule:
  - split on the FIRST two underscores into three parts:
      team   = before the 1st underscore
      doc_no = between the 1st and 2nd underscore
      title  = everything after the 2nd underscore (may contain underscores)
  - one underscore: team + doc_no, title=""
  - no underscore: whole = team, doc_no="", title=""
  - trailing/leading spaces trimmed
  - the full original folder name is also kept (request_folder)

Run: .venv/bin/python -m pytest tests/test_folder_parse.py -v
"""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from folder_parse_oracle import parse_request_folder_name  # noqa: E402


def test_three_part_team_docno_title():
    t, d, ti = parse_request_folder_name("인프라시너지셀_2026-782_제목")
    assert t == "인프라시너지셀"
    assert d == "2026-782"
    assert ti == "제목"


def test_title_keeps_later_underscores():
    # everything after the 2nd underscore is the title, underscores preserved
    t, d, ti = parse_request_folder_name("인프라시너지셀_2026-782_제목_추가설명")
    assert t == "인프라시너지셀"
    assert d == "2026-782"
    assert ti == "제목_추가설명"


def test_two_part_no_title():
    t, d, ti = parse_request_folder_name("정보보호센터_1234")
    assert t == "정보보호센터"
    assert d == "1234"
    assert ti == ""


def test_no_underscore():
    t, d, ti = parse_request_folder_name("인프라팀")
    assert t == "인프라팀"
    assert d == ""
    assert ti == ""


def test_trims_spaces():
    t, d, ti = parse_request_folder_name("  정보보호센터 _ 1234 _ 웹서비스 ")
    assert t == "정보보호센터"
    assert d == "1234"
    assert ti == "웹서비스"


def test_english_center():
    t, d, ti = parse_request_folder_name("SecurityCenter_A-2026-001_TitleX")
    assert t == "SecurityCenter"
    assert d == "A-2026-001"
    assert ti == "TitleX"


def test_empty():
    t, d, ti = parse_request_folder_name("")
    assert t == ""
    assert d == ""
    assert ti == ""


def test_trailing_underscore():
    # "정보보호센터_" -> team=정보보호센터, doc="", title=""
    t, d, ti = parse_request_folder_name("정보보호센터_")
    assert t == "정보보호센터"
    assert d == ""
    assert ti == ""
