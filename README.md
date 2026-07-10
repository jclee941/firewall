# 방화벽 정책 자동화 / Firewall Policy Automation

[![Python](https://img.shields.io/badge/python-3.x-blue.svg)](requirements.txt)
[![VBA](https://img.shields.io/badge/VBA-Excel--native-success.svg)](vba/)
[![License](https://img.shields.io/badge/license-See%20LICENSE-lightgrey.svg)](LICENSE)
[![Build](https://img.shields.io/badge/build-Linux%20%2B%20pyOpenVBA-orange.svg)](scripts/build_xlsm.py)

## 한 줄 요약 / Summary

**한국어** — Excel 네이티브 VBA로 방화벽 정책 신청서를 병합하고 `대상방화벽`을 자동 산출하는 도구입니다. Linux 빌드 도구로 매크로 포함 `.xlsm`을 생성하며, 동일 로직의 SECUI CLI/GUI(PySide6)를 함께 제공합니다.

**English** — Excel-native VBA automation that merges firewall policy request spreadsheets and computes `대상방화벽` from user-defined firewall ranges. Ships a macro-enabled `.xlsm` built on Linux, plus a behavior-identical SECUI CLI/GUI (PySide6).

## 상태 / Status

| 영역 | 상태 | 근거 |
| --- | --- | --- |
| 빌드 산출물 | 프로덕션 | `scripts/build_xlsm.py` + `pyOpenVBA==2.0.0` |
| 테스트 | 안정 | pytest 9.0.3, `tests/` 전체 통과 |
| VBA ↔ Python 정합 | 강제 | `tests/route_oracle.py`, `tests/request_parser_oracle.py` |
| SECUI CLI | 안정 | `scripts/secui_cli.py` |
| SECUI GUI | 안정 | PySide6 6.11.1, `scripts/secui_gui.py` |
| 의존성 | 고정 | `requirements.txt`, `requirements-gui.txt` 핀 사용 |

## 운영 흐름 / Operator Flow

1. 신청서를 `request-folder/` 하위 부서 폴더에 둔다.
2. `./.venv/bin/python scripts/build_xlsm.py` 로 `dist/firewall-policy-automation.xlsm` 생성.
3. 운영자가 `.xlsm` 을 열어 `requests` 시트에 행을 추가하고 `대상방화벽` 매크로를 실행.
4. `route_results` 시트에서 검증 상태와 경로 진단을 확인.
5. 외부에서는 `./.venv/bin/python scripts/secui_cli.py --workbook ...` 또는 SECUI GUI 사용.

---

## 목차 / Table of Contents

1. [목적 / Purpose](#목적--purpose)
2. [패키지 구성 / Package Contents](#패키지-구성--package-contents)
3. [먼저 읽을 파일 / First Files to Read](#먼저-읽을-파일--first-files-to-read)
4. [워크북 계약 / Workbook Contract](#워크북-계약--workbook-contract)
5. [빠른 시작 / Quickstart](#빠른-시작--quickstart)
6. [명령어 / Commands](#명령어--commands)
7. [아키텍처 / Architecture](#아키텍처--architecture)
8. [설정 / Configuration](#설정--configuration)
9. [로컬 개발 / Local Development](#로컬-개발--local-development)
10. [테스트 / Testing](#테스트--testing)
11. [기여 / Contributing](#기여--contributing)
12. [문서 / Further Documentation](#문서--further-documentation)
13. [유지보수 / Maintainers](#유지보수--maintainers)

---

## 목적 / Purpose

이 저장소는 다음을 목표로 합니다.

- **신청서 병합**: 부서별 폴더에 흩어진 `.xlsx` 신청서를 단일 `requests` 시트로 통합.
- **대상 방화벽 산출**: 운영자가 정의한 `firewall_ranges` 의 대역 매칭으로 `대상방화벽` 을 자동 계산.
- **검증 진단**: 출발지/목적지 IP가 어느 방화벽과 매칭되는지 근거와 함께 제시.
- **Excel 네이티브**: PowerQuery 없이 VBA만으로 동작, 매크로 포함 통합 문서로 배포.
- **Linux 빌드**: Windows/Excel/COM 없이 `pyOpenVBA` 로 `.xlsm` 생성.
- **외부 도구**: 동일 로직의 SECUI CLI/GUI 로 워크북 외부에서도 룰셋 산출.

핵심 사용자: 정보보호센터·인프라팀의 방화벽 정책 검토자, SECUI 장비 룰셋 작업자.

## 패키지 구성 / Package Contents

| 경로 | 역할 |
| --- | --- |
| `vba/` | `.xlsm` 에 주입되는 VBA 모듈. 정책 병합, `대상방화벽` 계산, SECUI 변환. |
| `firewall_policy/` | VBA 동작의 Python 미러(`cidr`, `request_parser`, `folder_parse`, `gui_export`). |
| `scripts/` | Linux 빌더, SECUI CLI/GUI, 릴리스 번들러. |
| `tests/` | Oracle 명세, VBA 동작 정합성, 워크북 구조 테스트. |
| `docs/` | 운영자용 스키마 문서, Excel 네이티브 런북, 벤치마크 메모. |
| `request-folder/` | 샘플 신청 트리. `.xlsx` 는 생성/갱신 대상. |

## 먼저 읽을 파일 / First Files to Read

| 우선순위 | 파일 | 이유 |
| --- | --- | --- |
| 1 | [`AGENTS.md`](AGENTS.md) | 프로젝트 지식 베이스, 워크북 계약, 명령어. |
| 2 | [`scripts/workbook_contract.py`](scripts/workbook_contract.py) | 시트/헤더/필드 정의의 단일 출처. |
| 3 | [`vba/FirewallPolicyAutomation.bas`](vba/FirewallPolicyAutomation.bas) | VBA 본체, `대상방화벽` 계산 로직. |
| 4 | [`vba/FirewallRouteAnalysis.bas`](vba/FirewallRouteAnalysis.bas) | 방화벽 범위/대역 분석. |
| 5 | [`tests/route_oracle.py`](tests/route_oracle.py) | 라우트 산출 명세, VBA 정합 기준. |
| 6 | [`docs/excel-schema.md`](docs/excel-schema.md) | 운영자가 시트 구조를 이해하기 위한 문서. |
| 7 | [`docs/excel-native.md`](docs/excel-native.md) | Excel 네이티브 런북. |

## 워크북 계약 / Workbook Contract

| 시트 | 분류 | 설명 |
| --- | --- | --- |
| `requests` | 사용자 입력 | 14컬럼, 헤더 행 2, 데이터 행 3 시작. |
| `_request_tracking` | 숨김 | `request_row` 키 기준 원본 추적 메타. |
| `route_results` | 결과 | 검증·매칭 진단, 근거 컬럼 포함. |
| `firewalls`, `firewall_ranges` | 운영자 참조 | 방화벽 목록과 사용자 정의 대역. |
| `settings`, `header_aliases` | 운영자 참조 | 동작 옵션과 헤더 별칭 매핑. |
| `vendor_cli_templates`, `service_catalog` | 운영자 참조 | 벤더 CLI 템플릿과 서비스 카탈로그. |

핵심 컬럼(요청서): `요청부서`, `요청번호`, `대상방화벽`, `출발지`, `출발지설명`, `목적지`, `목적지설명`, `프로토콜`, `포트`, `방향`, `용도`, `시작일`, `종료일`, `비고`.

규칙:

- `대상방화벽` 은 `requests` 3열, 내부 필드는 `target_firewalls`. 다중 값은 `;` 로 결합.
- 원본 추적(`원본파일`, `원본행`, `요청폴더`, `제목`)은 `requests` 가 아닌 `_request_tracking` 에 보관.
- 라우트 진단(`검증상태`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거`, `원본파일`, `원본행`)은 `route_results` 에 기록.

## 빠른 시작 / Quickstart

```bash
# 1) 가상환경 준비 (저장소 루트)
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
./.venv/bin/pip install -r requirements-gui.txt   # SECUI GUI 사용 시

# 2) 매크로 포함 통합 문서 빌드
./.venv/bin/python scripts/build_xlsm.py

# 3) 샘플 신청 트리 생성
./.venv/bin/python scripts/make_request_folder.py

# 4) 워크북 외부에서 SECUI 룰셋 생성
./.venv/bin/python scripts/secui_cli.py --workbook dist/firewall-policy-automation.xlsm --format text

# 5) GUI 진입
./.venv/bin/python scripts/secui_gui.py
```

빌드는 Windows, Excel, PowerShell, COM 없이 Linux 에서 완료됩니다.

## 명령어 / Commands

| 명령 | 용도 |
| --- | --- |
| `./.venv/bin/python -m pytest tests/ -q` | 전체 테스트. |
| `./.venv/bin/python -m pytest tests/test_route_oracle.py -v` | 라우트 산출 명세 테스트. |
| `./.venv/bin/python scripts/build_xlsm.py` | `.xlsm` 빌드. |
| `./.venv/bin/python scripts/make_request_folder.py` | 신청 폴더 트리 재생성. |
| `./.venv/bin/python scripts/secui_cli.py --workbook ... --format text` | SECUI 룰셋 산출. |
| `./.venv/bin/python scripts/secui_gui.py` | PySide6 GUI 진입. |
| `./.venv/bin/python scripts/build_standalone_gui.py` | SECUI GUI 네이티브 번들 빌드. |
| `./.venv/bin/python scripts/build_release_bundle.py` | 릴리스 번들 생성. |

`.venv/bin/python` 사용을 권장하며, 핀은 `pyOpenVBA==2.0.0`, `openpyxl==3.1.5`, `pytest==9.0.3`, `PySide6==6.11.1` 입니다.

## 아키텍처 / Architecture

| 계층 | 모듈 | 책임 |
| --- | --- | --- |
| VBA 본체 | `vba/FirewallPolicyAutomation.bas` | 신청 병합, `대상방화벽` 산출, SECUI 변환. |
| VBA 라우팅 | `vba/FirewallRouteAnalysis.bas` | 방화벽/대역 매칭, 경로 분석. |
| Python 미러 | `firewall_policy/cidr.py`, `request_parser.py`, `folder_parse.py` | 동일 로직의 파이썬 구현, SECUI 백엔드. |
| 워크북 정의 | `scripts/workbook_contract.py`, `workbook_ux.py` | 시트/헤더/UX 보조 정의. |
| 빌더 | `scripts/build_xlsm.py` | `pyOpenVBA` 로 `vbaProject.bin` 주입. |
| CLI | `scripts/secui_cli.py`, `secui_cli_runtime.py`, `secui_cli_seed.py` | 워크북 외부에서 룰셋 산출. |
| GUI | `scripts/secui_gui.py`, `firewall_policy/gui_export.py` | PySide6 UI와 워크북 익스포트. |
| 테스트 | `tests/` | Oracle 명세와 VBA 정합 검증. |

요청 처리 흐름:

1. `request-folder/` 트리에서 부서 폴더 단위로 `.xlsx` 수집.
2. `request_parser` 가 헤더 별칭과 병합 셀 규칙으로 행 파싱.
3. `firewall_ranges` 와 출발지/목적지를 CIDR 매칭하여 `대상방화벽` 결정.
4. `_request_tracking` 에 원본 추적 정보 기록, `route_results` 에 진단 결과 기록.
5. SECUI 변환 시 다중 `대상방화벽` 을 `;` 분리하여 그룹 룰 생성.

## 설정 / Configuration

| 항목 | 위치 | 비고 |
| --- | --- | --- |
| 시트 계약 | `scripts/workbook_contract.py` | 시트/컬럼 정의의 단일 출처. |
| 헤더 별칭 | 워크북 `header_aliases` 시트 | 부서별 상이한 표기 흡수. |
| 방화벽/대역 | 워크북 `firewalls`, `firewall_ranges` 시트 | 운영자가 직접 편집. |
| 동작 옵션 | 워크북 `settings` 시트 | 매크로 동작 토글 등. |
| 벤더 템플릿 | 워크북 `vendor_cli_templates` 시트 | SECUI CLI 출력 형식. |
| 서비스 카탈로그 | 워크북 `service_catalog` 시트 | 포트/프로토콜 명명 규칙. |
| 의존성 핀 | `requirements.txt`, `requirements-gui.txt` | 빌드/테스트 환경 고정. |

## 로컬 개발 / Local Development

- Python은 저장소 내 `./.venv/bin/python` 만 사용. 시스템 python 사용 금지.
- VBA 변경 시 `tests/route_oracle.py`, `tests/request_parser_oracle.py` 와 정합 유지.
- 워크북 UX(`scripts/workbook_ux.py`) 변경은 표시/입력 보조에만 영향. 라우트/파서 값은 변경 금지.
- 워크북 구조 변경 시 `tests/test_xlsm_structure.py` 와 `scripts/build_xlsm.py` 를 함께 갱신.
- `firewall_policy/` 의 Python 모듈은 VBA 동작의 정합 미러. 불일치 발생 시 Oracle 테스트가 실패함.

## 테스트 / Testing

| 테스트 | 대상 | 검증 범위 |
| --- | --- | --- |
| `tests/test_route_oracle.py` | `route_oracle.py` | 라우트 산출 명세. |
| `tests/test_route_oracle.py` ↔ VBA | `FirewallRouteAnalysis.bas` | VBA 동작 정합. |
| `tests/test_request_parser_oracle.py` ↔ VBA | `FirewallPolicyAutomation.bas` | 파서 정합. |
| `tests/test_xlsm_structure.py` | `build_xlsm.py` 산출물 | 시트/시드 구조. |
| `tests/test_cidr_parity.py` | `firewall_policy/cidr.py` | CIDR 계산. |
| `tests/test_folder_parse.py` | `folder_parse.py` | 신청 트리 파싱. |
| `tests/test_date_cells.py` | 워크북 | 날짜 셀 처리. |
| `tests/test_merged_cells.py` | 워크북 | 병합 셀 처리. |
| `tests/test_release_bundle.py` | `build_release_bundle.py` | 릴리스 번들. |
| `tests/test_release_workflows.py` | `.github/workflows/release.yml` | 릴리스 워크플로. |
| `tests/test_secui_cli_*.py` | `secui_cli.py` | CLI 가드/스크립트. |
| `tests/test_secui_gui.py` | `secui_gui.py` | GUI 동작. |
| `tests/test_secui_policy_analysis.py` | 정책 분석 | 정책 분석 경로. |
| `tests/test_secui_service_catalog.py` | 서비스 카탈로그 | 카탈로그 매칭. |
| `tests/test_user_alias.py` | 사용자 별칭 | 별칭 매핑. |
| `tests/test_vba_runtime_guards.py` | VBA 런타임 가드 | 매크로 안전 가드. |
| `tests/test_weekly_report_sheet.py` | 워크북 | 주간 리포트 시트. |
| `tests/test_workbook_convenience.py` | 워크북 | 편의 기능. |
| `tests/test_workbook_usage_links.py` | 워크북 | 사용 안내 링크. |
| `tests/test_build_route_results.py` | 빌드 | `route_results` 빌드 경로. |
| `tests/test_request_folder.py` | 신청 폴더 | 트리 생성/처리. |
| `tests/test_request_parsing.py` | 신청서 파싱 | 파서 동작. |

```bash
./.venv/bin/python -m pytest tests/ -q
```

## 기여 / Contributing

1. 이슈/요청을 먼저 등록하고 변경 범위를 합의합니다.
2. 분기는 `master` 에서 분기, 작업 후 PR 을 열고 Oracle 테스트 통과를 확인합니다.
3. VBA ↔ Python 정합이 깨지면 PR 이 거부될 수 있습니다.
4. 상세 규칙은 [`CONTRIBUTING.md`](CONTRIBUTING.md) 를 따릅니다.

## 문서 / Further Documentation

| 문서 | 용도 |
| --- | --- |
| [`docs/excel-schema.md`](docs/excel-schema.md) | 시트 스키마와 컬럼 정의. |
| [`docs/excel-native.md`](docs/excel-native.md) | Excel 네이티브 런북. |
| [`docs/firewall-excel-benchmark.md`](docs/firewall-excel-benchmark.md) | 방화벽 표 기반 벤치마크 메모. |
| [`docs/research-notes.md`](docs/research-notes.md) | 설계/탐구 메모. |
| [`request-folder/README.txt`](request-folder/README.txt) | 샘플 신청 트리 안내. |
| [`AGENTS.md`](AGENTS.md) | 프로젝트 지식 베이스. |

## 유지보수 / Maintainers

- 코드 소유: 방화벽 정책 자동화 팀 (정보보호/인프라 협업).
- 이슈 트래커: 저장소 내 이슈 탭을 사용.
- 릴리스: `.github/workflows/release.yml` — `v*` 태그 푸시 시 빌드·테스트·검증·번들·퍼블리시 자동화.

## 라이선스 / License

[`LICENSE`](LICENSE) 파일 참조.