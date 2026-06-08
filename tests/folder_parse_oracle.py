"""Pure-Python mirror of the VBA ParseRequestFolderName helper.

Splits a request's parent folder name into (team/center, doc_no).
Rule: split on the LAST underscore.
  "정보보호센터_1234"      -> ("정보보호센터", "1234")
  "정보보호_2팀_5678"      -> ("정보보호_2팀", "5678")
  "인프라팀"               -> ("인프라팀", "")
  "정보보호센터_"          -> ("정보보호센터", "")
Spaces around the parts are trimmed.
"""

from __future__ import annotations


def parse_request_folder_name(name: str) -> tuple[str, str]:
    s = (name or "").strip()
    if s == "":
        return "", ""
    idx = s.rfind("_")
    if idx < 0:
        return s.strip(), ""
    team = s[:idx].strip()
    doc_no = s[idx + 1:].strip()
    return team, doc_no
