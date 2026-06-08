"""Pure-Python mirror of the VBA user-defined header alias logic.

settings 시트의 header_alias 값을 파싱해 {정규화별칭: canonical} 맵을 만들고,
built-in CanonicalHeaderName이 못 잡은 헤더를 사용자 별칭으로 보정한다.

format: "출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"
  - ';' separates entries
  - 'canonical=alias1,alias2'
  - both canonical and aliases normalized via header_key
"""

from __future__ import annotations

from request_parser_oracle import canonical_header_name, header_key


def parse_user_aliases(text) -> dict[str, str]:
    result: dict[str, str] = {}
    if not text:
        return result
    for entry in str(text).split(";"):
        entry = entry.strip()
        if "=" not in entry:
            continue
        left, right = entry.split("=", 1)
        canon = header_key(left)
        if not canon:
            continue
        for alias in right.split(","):
            a = header_key(alias)
            if a:
                result[a] = canon
    return result


def canonical_with_user_aliases(raw_header: str, user_aliases: dict[str, str]) -> str:
    key = header_key(raw_header)
    # built-in first
    canon = canonical_header_name(key)
    if canon != key:
        return canon  # built-in recognized it
    # user-defined fallback
    if key in user_aliases:
        return user_aliases[key]
    return key  # passthrough (unknown)
