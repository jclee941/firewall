# 방화벽 정책 자동화 (Firewall Policy Automation)

[![Build](https://img.shields.io/badge/build-Linux-blue.svg)](#quickstart)
[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB.svg)](#quickstart)
[![License](https://img.shields.io/badge/license-Internal-lightgrey.svg)](#license)
[![Macro](https://img.shields.io/badge/macro-.xlsm-217346.svg)](#package-contents)

## 한국어 요약

신청자가 작성한 방화벽 정책 요청 엑셀을 받아, 출발지/목적지/포트/방향/용도를
파싱하고, 등록된 방화벽 대역과 비교하여 `대상방화벽` 컬럼과 SECUI CLI 명령을
생성하는 도구입니다. 작업은 Excel-native VBA 워크북에서 수행하며, 동일한 로직을
리눅스에서 검증·실행하기 위한 Python CLI/GUI와 빌드 스크립트를 함께 제공합니다.

## English Summary

Linux-buildable Excel-native tool that ingests firewall-policy request workbooks,
parses source/destination/port/direction/usage fields, resolves them against a
firewall-range registry, and writes `대상방화벽` plus SECUI CLI commands. VBA
inside the shipped `.xlsm` is mirrored in Python so the same workflow runs from
the Linux CLI or PySide6 GUI.

## 빠른 상태 (Status)

| 항목 | 값 |
| --- | --- |
| 릴리스 상태 | 운영 배포 (Production) |
| 빌드 호스트 | Linux (Excel/Windows/COM 불필요) |
| 매크로 출력 | `.xlsm` (VBA 내장) |
| Python 동등성 | VBA ↔ Python 결과 동일성 검증 |
| 테스트 도구 | `pytest` 9.0.3, oracle 비교 |
| GUI 스택 | PySide6 6.11.1 |
| 핀 버전 | `pyOpenVBA==2.0.0`, `openpyxl==3.1.5` |

## 한눈에 보는 흐름 (Flow)

1. 신청자가 `request-folder/` 아래 팀별 폴더에 `신청서_*.xlsx`를 둔다.
2. VBA 워크북의 `requests` 시트(헤더 행 2, 데이터 행 3–)에 14개 표준 컬럼을 채운다.
3. `대상방화벽`은 `firewall_ranges` 시트와 비교하여 자동으로 결정한다.
4. SECUI CLI 출력은 `vendor_cli_templates`/`service_catalog`로 생성하고
   `route_results` 시트에 진단 정보까지 기록한다.
5. 리눅스에서는 동일 로직을 `secui_cli.py`/`secui_gui.py`로 실행한다.

## 목차 (Table of Contents)

- [패키지 구성 / Package Contents](#package-contents)
- [상태 / Status](#status)
- [먼저 읽을 파일 / First Files to Read](#first-files-to-read)
- [진입점 / API and Entry Points](#api-and-entry-points)
- [빠른 시작 / Quickstart](#quickstart)
- [명령어 / Commands Reference](#commands-reference)
- [워크북 계약 / Workbook Contract](#workbook-contract)
- [로컬 개발 / Local Development](#local-development)
- [테스트 / Testing](#testing)
- [기여 / Contributing](#contributing)
- [유지보수 / Maintainers](#maintainers)
- [추가 문서 / Further Documentation](#further-documentation)
- [라이선스 / License](#license)

## 패키지 구성 / Package Contents

| 경로 | 역할 |
| --- | --- |
| `vba/` | `.xlsm`에 주입되는 매크로 모듈 (요청 병합, 라우트 계산) |
| `firewall_policy/` | VBA 로직의 순수 Python 미러 (`cidr.py`, `folder_parse.py`, `request_parser.py`, `gui_export.py`) |
| `scripts/` | 빌더, SECUI CLI/GUI, 요청 폴더 생성기 |
| `tests/` | Python 오라클, 워크북 구조 회귀 테스트 |
| `docs/` | 운영자 스키마, Excel-native 런북, 벤치마크 |
| `request-folder/` | 샘플 신청 트리 (`정보보호센터_1234/`, `인프라팀_5678/` 등) |
| `requirements.txt` | 빌드/테스트 핀 (pyOpenVBA, openpyxl, pytest) |
| `requirements-gui.txt` | PySide6 GUI 핀 |

## 상태 / Status

- **운영 준비 (Production-ready)**: 운영 부서에서 사용 중이며, `v*` 태그로
  빌드/검증/릴리스가 자동화되어 있습니다.
- **교차 플랫폼 로직**: VBA와 Python 구현이 동일 동작을 보장하도록 오라클 테스트로
  고정되어 있습니다.
- **지원 종료/비추천 항목**: 없음. 기존 `network_definitions`/`routing_paths` 폴백은
  코드상 유지되지만 신규 워크플로에서는 사용하지 않습니다.

## 먼저 읽을 파일 / First Files to Read

| 목적 | 권장 파일 |
| --- | --- |
| 제품 동작 이해 | `vba/FirewallRouteAnalysis.bas`, `vba/FirewallPolicyAutomation.bas` |
| 시트 스키마 확인 | `docs/excel-schema.md`, `docs/excel-native.md` |
| CLI 사용법 | `scripts/secui_cli.py`, `scripts/secui_cli_runtime.py` |
| GUI 사용법 | `scripts/secui_gui.py`, `firewall_policy/gui_export.py` |
| 빌드 산출물 | `scripts/build_xlsm.py`, `scripts/build_release_bundle.py` |
| 동작 변경 영향도 | `AGENTS.md`의 *Where To Look* 표 |

## 진입점 / API and Entry Points

| 진입점 | 형태 | 주요 인자 |
| --- | --- | --- |
| `.xlsm` 워크북 | 매크로 실행 | 신청 폴더 경로, 출력 폴더 |
| `scripts/secui_cli.py` | Linux CLI | `--workbook`, `--format text` 등 |
| `scripts/secui_gui.py` | PySide6 GUI | 워크북 로드 후 대화형 실행 |
| `firewall_policy.request_parser` | Python 모듈 | 신청 행 → SECUI 룰 묶음 |
| `firewall_policy.folder_parse` | Python 모듈 | 신청 폴더 스캔/병합 |
| `firewall_policy.cidr` | Python 모듈 | CIDR 매칭/병합 |

## 빠른 시작 / Quickstart

### 1. 의존성 설치

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
# GUI 빌드까지 검증하려면:
./.venv/bin/pip install -r requirements-gui.txt
```

### 2. 워크북 빌드

```bash
./.venv/bin/python scripts/build_xlsm.py
# 결과: dist/firewall-policy-automation.xlsm
```

### 3. SECUI CLI 실행 (Excel 없이)

```bash
./.venv/bin/python scripts/secui_cli.py \
    --workbook dist/firewall-policy-automation.xlsm \
    --format text
```

### 4. SECUI GUI 실행

```bash
./.venv/bin/python scripts/secui_gui.py
```

### 5. 신청 폴더 시드 생성

```bash
./.venv/bin/python scripts/make_request_folder.py
```

## 명령어 / Commands Reference

| 명령 | 용도 |
| --- | --- |
| `scripts/build_xlsm.py` | VBA를 주입한 매크로 워크북 빌드 |
| `scripts/build_standalone_gui.py` | PySide6 GUI 번들 빌드 |
| `scripts/build_release_bundle.py` | 릴리스 산출물 패키징 |
| `scripts/make_request_folder.py` | 샘플 신청 트리 재생성 |
| `scripts/secui_cli.py` | 워크북 → SECUI CLI 변환 |
| `scripts/secui_cli_runtime.py` | CLI 런타임 헬퍼 |
| `scripts/secui_cli_seed.py` | CLI용 시드 데이터 |
| `scripts/secui_gui.py` | 대화형 GUI |
| `scripts/route_seed.py` | 라우트 계산용 시드 |
| `scripts/workbook_contract.py` | 시트/헤더 계약 정의 |
| `scripts/workbook_ux.py` | 워크북 UX 보조 모듈 |

## 워크북 계약 / Workbook Contract

| 시트 | 종류 | 비고 |
| --- | --- | --- |
| `requests` | 사용자 입력(14열) | 헤더 행 2, 데이터 행 3– |
| `route_results` | 진단 결과 | 검증상태/메시지/경로/매칭근거 |
| `_request_tracking` | 숨김 | 원본파일·원본행·요청폴더·제목 |
| `firewalls` | 참조 | 방화벽 메타데이터 |
| `firewall_ranges` | 참조 | 사용자 정의 방화벽 대역 |
| `settings` | 참조 | 동작 파라미터 |
| `header_aliases` | 참조 | 부서별 헤더 별칭 |
| `vendor_cli_templates` | 참조 | 벤더 CLI 템플릿 |
| `service_catalog` | 참조 | 서비스 카탈로그 |

### `requests` 14개 컬럼

`요청부서`, `요청번호`, `대상방화벽`, `출발지`, `출발지설명`, `목적지`,
`목적지설명`, `프로토콜`, `포트`, `방향`, `용도`, `시작일`, `종료일`, `비고`.

`대상방화벽`은 여러 값일 때 `;`로 결합하며, 내부 결과 필드명은
`target_firewalls`입니다.

## 로컬 개발 / Local Development

- 커밋된 `./.venv/bin/python` 사용을 권장합니다.
- 핀은 `requirements.txt`(`pyOpenVBA==2.0.0`, `openpyxl==3.1.5`,
  `pytest==9.0.3`)와 `requirements-gui.txt`(`PySide6==6.11.1`)를 그대로
  따릅니다.
- 빌드는 Linux에서만 검증되며 Windows/Excel/PowerShell/COM 없이 동작합니다.
- 라우트 계산 로직을 변경할 때 `tests/route_oracle.py`와
  `vba/FirewallRouteAnalysis.bas`가 항상 동일 동작을 유지하도록 동기화합니다.
- 워크북 UX/내비게이션 변경은 `scripts/workbook_ux.py`에서만 수행하고, 라우트나
  파서 결과 값은 건드리지 않습니다.

## 테스트 / Testing

```bash
# 전체 회귀
./.venv/bin/python -m pytest tests/ -q

# 라우트 계산 단일 검증 (Python 오라클 ↔ VBA 동작 확인)
./.venv/bin/python -m pytest tests/test_route_oracle.py -v

# 워크북 구조/UX 회귀
./.venv/bin/python -m pytest tests/test_xlsm_structure.py \
    tests/test_workbook_usage_links.py -v
```

테스트는 오라클 비교(`route_oracle.py`, `request_parser_oracle.py`,
`folder_parse_oracle.py`, `user_alias_oracle.py`)와 워크북 구조 검사,
SECUI CLI/GUI 가드 검사, 빌드 결과물 검사로 구성됩니다.

## 기여 / Contributing

- 워크플로는 `CONTRIBUTING.md`를 따릅니다.
- `AGENTS.md`의 *Where To Look* 표를 기준으로 변경 영향 범위를 먼저 파악합니다.
- VBA와 Python을 동시에 손대야 하는 변경은 두 측이 같은 동작임을 테스트로
  입증한 뒤에 머지합니다.

## 유지보수 / Maintainers

- 운영팀 내방화벽 자동화 담당 (저장소 OWNERS 기준)
- 회신이 늦은 경우 저장소 `CODEOWNERS` 또는 조직 위키의 *Points of Contact*를
  참고하세요.

## 추가 문서 / Further Documentation

- [`docs/excel-schema.md`](docs/excel-schema.md) — 시트/헤더 스키마
- [`docs/excel-native.md`](docs/excel-native.md) — Excel-native 런북
- [`docs/firewall-excel-benchmark.md`](docs/firewall-excel-benchmark.md) — 벤치마크
- [`docs/research-notes.md`](docs/research-notes.md) — 설계 메모
- [`request-folder/README.txt`](request-folder/README.txt) — 신청 폴더 규칙
- [`AGENTS.md`](AGENTS.md) — 저장소 지식 베이스(개발자용)

## 라이선스 / License

저장소 내 [`LICENSE`](LICENSE) 파일의 조항을 따릅니다.