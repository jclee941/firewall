"""Pure-Python mirror of the VBA ParseRequestFolderName helper.

Splits a request's parent folder name into (team/center, doc_no, title).
Rule: split on the FIRST two underscores into three parts.
  "인프라시너지셀_2026-782_제목" -> ("인프라시너지셀", "2026-782", "제목")
  "인프라시너지셀_2026-782_제목_추가" -> ("인프라시너지셀", "2026-782", "제목_추가")
  "정보보호센터_1234"      -> ("정보보호센터", "1234", "")
  "인프라팀"               -> ("인프라팀", "", "")
  "정보보호센터_"          -> ("정보보호센터", "", "")
Spaces around the parts are trimmed.
"""

from __future__ import annotations


def parse_request_folder_name(name: str) -> tuple[str, str, str]:
    s = (name or "").strip()
    if s == "":
        return "", "", ""
    first = s.find("_")
    if first < 0:
        return s.strip(), "", ""
    team = s[:first].strip()
    rest = s[first + 1:]
    second = rest.find("_")
    if second < 0:
        return team, rest.strip(), ""
    doc_no = rest[:second].strip()
    title = rest[second + 1:].strip()
    return team, doc_no, title
