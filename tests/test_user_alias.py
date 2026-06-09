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
    aliases_from_rows,
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


def test_alias_sheet_rows():
    # header_aliases sheet: (standard, your_column) rows
    rows = [
        ["\ucd9c\ubc1c\uc9c0", "\ucd9c\ubc1c\uc9c0ip\uc124\uba85"],
        ["\ubaa9\uc801\uc9c0", "\ubaa9\uc801\uc9c0ip\uc124\uba85"],
        ["\ud504\ub85c\ud1a0\ucf5c", "tcp/udp"],
        ["\ubc29\ud5a5", "\uad6c\ubd84"],
        ["\uc2dc\uc791\uc77c", "\uc2dc\uc791\uc77c\uc790"],
        ["\uc885\ub8cc\uc77c", "\uc885\ub8cc\uc77c\uc790"],
    ]
    aliases = aliases_from_rows(rows)
    assert canonical_with_user_aliases("\ucd9c\ubc1c\uc9c0ip\uc124\uba85", aliases) == "\ucd9c\ubc1c\uc9c0"
    assert canonical_with_user_aliases("\ubaa9\uc801\uc9c0ip\uc124\uba85", aliases) == "\ubaa9\uc801\uc9c0"
    assert canonical_with_user_aliases("tcp/udp", aliases) == "\ud504\ub85c\ud1a0\ucf5c"
    assert canonical_with_user_aliases("\uad6c\ubd84", aliases) == "\ubc29\ud5a5"
    assert canonical_with_user_aliases("\uc2dc\uc791\uc77c\uc790", aliases) == "\uc2dc\uc791\uc77c"
    assert canonical_with_user_aliases("\uc885\ub8cc\uc77c\uc790", aliases) == "\uc885\ub8cc\uc77c"


def test_alias_sheet_standard_via_builtin():
    # `standard` side may itself be an alias (Source -> \ucd9c\ubc1c\uc9c0) and still resolve.
    # your_column uses names the built-in map does NOT know.
    rows = [["Source", "zzfromcol"], ["Destination", "zztocol"]]
    aliases = aliases_from_rows(rows)
    assert canonical_with_user_aliases("zzfromcol", aliases) == "\ucd9c\ubc1c\uc9c0"
    assert canonical_with_user_aliases("zztocol", aliases) == "\ubaa9\uc801\uc9c0"


def test_alias_sheet_skips_blank_rows():
    rows = [["", ""], ["\ubc29\ud5a5"], ["\ubc29\ud5a5", "\uad6c\ubd84"], None]
    aliases = aliases_from_rows(rows)
    assert aliases == {"\uad6c\ubd84": "\ubc29\ud5a5"}
