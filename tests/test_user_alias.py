"""사용자 정의 헤더 alias (settings의 header_alias 키).

Users add aliases without editing VBA, e.g. settings header_alias value:
  "출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"

Rule:
  - entries separated by ';'
  - each entry "canonical=alias1,alias2"
  - canonical is normalized via header_key (so 출발지IP -> 출발지ip)
  - each alias normalized via header_key
  - result: {normalized_alias: canonical_normalized}
  - canonical_header_name(key) checks built-in first, then user aliases

Run: .venv/bin/python -m pytest tests/test_user_alias.py -v
"""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from request_parser_oracle import (  # noqa: E402
    canonical_header_name,
    header_key,
)
from user_alias_oracle import (  # noqa: E402
    canonical_with_user_aliases,
    parse_user_aliases,
)


def test_parse_basic():
    m = parse_user_aliases("출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소")
    # canonical key is normalized; alias keys normalized
    assert m[header_key("출발지주소")] == header_key("출발지IP")
    assert m[header_key("Source Addr")] == header_key("출발지IP")
    assert m[header_key("목적지주소")] == header_key("목적지IP")


def test_parse_empty():
    assert parse_user_aliases("") == {}
    assert parse_user_aliases(None) == {}


def test_parse_skips_malformed():
    # entry without '=' is ignored; empty aliases ignored
    m = parse_user_aliases("출발지IP=; =alias; 목적지IP=목적지주소")
    assert header_key("목적지주소") in m
    assert m[header_key("목적지주소")] == header_key("목적지IP")
    # no crash, malformed dropped
    assert len(m) == 1


def test_user_alias_resolves():
    aliases = parse_user_aliases("출발지IP=출발지주소; 목적지IP=목적지주소")
    # a header the built-in map does NOT know
    assert canonical_with_user_aliases("출발지주소", aliases) == "출발지ip"
    assert canonical_with_user_aliases("목적지주소", aliases) == "목적지ip"


def test_builtin_wins_over_user():
    # built-in alias must still resolve even if user redefines weirdly
    aliases = parse_user_aliases("출발지IP=srcip")  # srcip is already built-in -> 출발지ip
    # built-in already maps srcip -> 출발지ip; user agrees, result stable
    assert canonical_with_user_aliases("srcip", aliases) == "출발지ip"


def test_unknown_still_passthrough():
    aliases = parse_user_aliases("출발지IP=출발지주소")
    # a totally unknown header returns itself (normalized)
    assert canonical_with_user_aliases("randomcol", aliases) == "randomcol"


def test_english_alias():
    aliases = parse_user_aliases("프로토콜=ip protocol; 포트=dest port")
    assert canonical_with_user_aliases("IP Protocol", aliases) == "프로토콜"
    assert canonical_with_user_aliases("Dest Port", aliases) == "포트"
