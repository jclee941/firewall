# 방화벽 정책 자동화 (Firewall Policy Automation)

[![Build](https://img.shields.io/badge/build-Linux%20%2B%20pyOpenVBA%202.0.0-blue)](#빠른-시작-quickstart)
[![Tests](https://img.shields.io/badge/tests-pytest%209.0.3-green)](#테스트-testing)
[![Excel](https://img.shields.io/badge/Excel-.xlsm%20native-orange)](#워크북-계약-workbook-contract)
[![GUI](https://img.shields.io/badge/GUI-PySide6%206.11.1-blue)](#api--진입점-entry-points)

## 한국어 요약 (Summary)

Excel 네이티브 VBA 도구로, 방화벽 정책 신청서를 병합하고 사용자 정의 방화벽 대역 행으로부터 `대상방화벽`을 계산합니다. PowerQuery 없이 동작하며, Linux 환경에서 pyOpenVBA로 매크로 포함 `.xlsm`을 직접 빌드합니다. 동일 파싱·라우팅 로직이 `firewall_policy/` 패키지에 Python으로 미러링되어, 워크북 외부(CLI, GUI, 회귀 테스트)에서 동일한 결과를 보장합니다.

### 빠른 상태 (Quick Status)

| 항목 | 값 |
| --- | --- |
| 제품 성숙도 | 운영 배포 가능 (Excel/Windows/COM 없이 Linux 빌드) |
| 핵심 산출물 | 매크로 포함 `.xlsm` + SECUI CLI/GUI |
| 빌드 의존성 | `pyOpenVBA==2.0.0`, `openpyxl==3.1.5` |
| GUI 의존성 | `PySide6==6.11.1` (`requirements-gui.txt`) |
| 테스트 | `pytest==9.0.3` |
| 진입점 | `scripts/build_xlsm.py`, `scripts/secui_cli.py`, `scripts/secui_gui.py` |
| 테스트 명령 | `./.venv/bin/python -m pytest tests/ -q` |
| 데이터 모델 | `requests` 14컬럼, `route_results` 진단, `_request_tracking` 추적 |
| 라이선스 | 저장소 `LICENSE` 참조 |

### 운영 흐름 (Operator Flow)

1. `request-folder/<부서>_<번호>/` 에 신청서(xlsx)를 배치합니다.
2. `.xlsm`을 열어 요청 폴더를 동기화하고 `대상방화벽` 계산을 실행합니다.
3. `route_results` 시트에서 검증 상태와 매칭 근거를 확인합니다.
4. SECUI CLI 또는 GUI로 `대상방화벽`(세미콜론 결합)을 추출합니다.

## English (Reference)

| Item | Value |
| --- | --- |
| Tool type | Excel-native VBA, macro-enabled `.xlsm` |
| Computes | `대상방화벽` from firewall range rows |
| Build host | Linux + pyOpenVBA 2.0.0 (no Excel/COM) |
| External runner | SECUI CLI / PySide6 GUI |
| Test runner | pytest 9.0.3 |

## 목차 (Table of Contents)

- [구성 요소 (Package Contents)](#구성-요소-package-contents)
- [현재 상태 (Status)](#현재-상태-status)
- [먼저 읽을 파일 (First Files to Read)](#먼저-읽을-파일-first-files-to-read)
- [API / 진입점 (Entry Points)](#api--진입점-entry-points)
- [빠른 시작 (Quickstart)](#빠른-시작-quickstart)
- [워크북 계약 (Workbook Contract)](#워크북-계약-workbook-contract)
- [명령어 레퍼런스 (Commands Reference)](#명령어-레퍼런스-commands-reference)
- [테스트 (Testing)](#테스트-testing)
- [로컬 개발 (Local Development)](#로컬-개발-local-development)
- [기여 (Contribution)](#기여-contribution)
- [유지보수 (Maintainers)](#유지보수-maintainers)
- [참고 문서 (Further Documentation)](#참고-문서-further-documentation)

## 구성 요소 (Package Contents)

| 경로 | 역할 |
| --- | --- |
| `vba/` | `.xlsm`에 주입되는 VBA 모듈 (`FirewallPolicyAutomation.bas`, `FirewallRouteAnalysis.bas`) |
| `firewall_policy/` | VBA 파싱 로직의 Python 미러 (CLI/GUI 백엔드) |
| `scripts/` | Linux 빌더, SECUI CLI/GUI, 릴리스 번들러 |
| `tests/` | Python 오라클, 워크북 구조 테스트, 회귀 테스트 |
| `docs/` | 운영자 스키마, Excel 네이티브 런북, 벤치마크 |
| `request-folder/` | 부서별 신청서 샘플 트리 (xlsx는 생성·변동) |

## 현재 상태 (Status)

- Linux 빌드/테스트 통과: `tests/` 전체 pytest 실행으로 회귀 검증.
- 매크로 포함 `.xlsm` 빌드: Windows, Excel, PowerShell, COM 없이 가능.
- VBA ↔ Python 동등성: `tests/route_oracle.py`, `tests/request_parser_oracle.py`로 강제.
- SECUI 출력: `대상방화벽` 컬럼 값을 `;`로 결합한 배치/CLI 분리.
- 본 저장소는 별도 운영 인스턴스나 외부 API를 노출하지 않으며, 워크북과 CLI/GUI가 동일한 결과를 내는지를 회귀 테스트로 보장합니다.

## 먼저 읽을 파일 (First Files to Read)

| 작업 | 파일 |
| --- | --- |
| 라우트 계산 동작 이해 | `vba/FirewallRouteAnalysis.bas`, `tests/route_oracle.py`, `tests/test_route_oracle.py` |
| 워크북 시트/헤더/시드 구조 파악 | `scripts/workbook_contract.py`, `scripts/build_xlsm.py`, `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py` |
| 워크북 UX/내비게이션 변경 | `scripts/workbook_ux.py`, `tests/test_workbook_usage_links.py` |
| 신청서 파서 변경 | `tests/request_parser_oracle.py`, `vba/FirewallPolicyAutomation.bas`, `firewall_policy/request_parser.py` |
| 외부 SECUI 변환 | `scripts/secui_cli.py`, `scripts/secui_cli_runtime.py`, `scripts/secui_cli_seed.py` |
| GUI / 독립 실행 번들 | `scripts/secui_gui.py`, `firewall_policy/gui_export.py`, `scripts/build_standalone_gui.py` |
| 릴리스 산출물 변경 | `scripts/build_xlsm.py`, `scripts/build_release_bundle.py`, `scripts/make_request_folder.py`, `tests/test_release_bundle.py` |

## API / 진입점 (Entry Points)

### Excel 워크북 시트

| 시트 | 가시성 | 역할 |
| --- | --- | --- |
| `requests` | 표시 | 사용자 입력 14컬럼 (헤더 row 2, 데이터 row 3부터) |
| `route_results` | 표시 | 라우팅 진단 (검증상태, 매칭근거 등) |
| `_request_tracking` | 숨김 | 원본 파일/행 추적 (`request_row` 키) |
| `firewalls` | 표시 | 방화벽 정의 입력 |
| `firewall_ranges` | 표시 | 방화벽 대역 행 입력 |
| `settings` | 표시 | 운영 설정 |
| `header_aliases` | 표시 | 부서별 헤더 별칭 |
| `vendor_cli_templates` | 표시 | 벤더 CLI 템플릿 |
| `service_catalog` | 표시 | 서비스 카탈로그 |

### 스크립트 진입점

| 진입점 | 용도 |
| --- | --- |
| `scripts/build_xlsm.py` | 매크로 포함 `.xlsm` 빌드 |
| `scripts/secui_cli.py` | 외부 SECUI 변환 CLI |
| `scripts/secui_gui.py` | PySide6 GUI 실행 |
| `scripts/make_request_folder.py` | `request-folder/` 트리 생성 |
| `scripts/build_release_bundle.py` | 릴리스 산출물 패키징 |
| `scripts/build_standalone_gui.py` | GUI 독립 실행 번들 생성 |

## 빠른 시작 (Quickstart)

```bash
# 1) 의존성 설치 (저장소 루트)
python -m venv .venv
./.venv/bin/pip install -r requirements.txt -r requirements-gui.txt

# 2) 워크북 빌드
./.venv/bin/python scripts/build_xlsm.py

# 3) 회귀 테스트
./.venv/bin/python -m pytest tests/ -q

# 4) 외부 SECUI 변환 (텍스트)
./.venv/bin/python scripts/secui_cli.py \
    --workbook dist/firewall-policy-automation.xlsm \
    --format text

# 5) GUI 실행
./.venv/bin/python scripts/secui_gui.py
```

Excel에서는 `.xlsm`을 열고 `요청폴더 동기화` → `대상방화벽 계산` → `route_results` 검토 순서로 사용합니다.

## 워크북 계약 (Workbook Contract)

`requests` 시트는 다음 14컬럼의 클린 입력 시트입니다 (헤더 row 2, 데이터 row 3부터).

| # | 컬럼 | 용도 |
| --- | --- | --- |
| 1 | 요청부서 | 부서 식별자 |
| 2 | 요청번호 | 부서 내 신청 번호 |
| 3 | 대상방화벽 | 계산 결과 (내부 필드 `target_firewalls`, `;` 결합) |
| 4 | 출발지 | 출발지 주소/CIDR |
| 5 | 출발지설명 | 출발지 설명 |
| 6 | 목적지 | 목적지 주소/CIDR |
| 7 | 목적지설명 | 목적지 설명 |
| 8 | 프로토콜 | TCP/UDP 등 |
| 9 | 포트 | 포트/범위 |
| 10 | 방향 | in/out 등 |
| 11 | 용도 | 용도 분류 |
| 12 | 시작일 | 적용 시작일 |
| 13 | 종료일 | 적용 종료일 |
| 14 | 비고 | 비고 |

추가 규칙:

- 원본 추적 (`원본파일`, `원본행`, `요청폴더`, `제목`)은 숨김 시트 `_request_tracking`의 `request_row` 키로 보관됩니다. `requests` 시트에 직접 두지 않습니다.
- 라우팅 진단 (`검증상태`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거`, `원본파일`, `원본행`)은 `route_results` 시트에 기록됩니다.
- SECUI 출력(배치/CLI)은 `대상방화벽`을 `;`로 결합합니다.

## 명령어 레퍼런스 (Commands Reference)

| 명령 | 설명 |
| --- | --- |
| `./.venv/bin/python scripts/build_xlsm.py` | 매크로 포함 `.xlsm` 빌드 |
| `./.venv/bin/python scripts/make_request_folder.py` | `request-folder/` 트리 생성 |
| `./.venv/bin/python scripts/secui_cli.py --workbook <path> --format text` | SECUI 텍스트 변환 |
| `./.venv/bin/python scripts/secui_cli.py --workbook <path> --format batch` | SECUI 배치 변환 |
| `./.venv/bin/python scripts/secui_gui.py` | PySide6 GUI 실행 |
| `./.venv/bin/python scripts/build_standalone_gui.py` | GUI 독립 실행 번들 생성 |
| `./.venv/bin/python scripts/build_release_bundle.py` | 릴리스 산출물 패키징 |
| `./.venv/bin/python -m pytest tests/ -q` | 전체 회귀 테스트 |
| `./.venv/bin/python -m pytest tests/test_route_oracle.py -v` | 라우트 오라클 검증 |

## 테스트 (Testing)

| 테스트 종류 | 위치 |
| --- | --- |
| 라우트 계산 오라클 | `tests/route_oracle.py`, `tests/test_route_oracle.py` |
| 신청서 파서 오라클 | `tests/request_parser_oracle.py`, `tests/test_request_parsing.py` |
| 워크북 구조 | `tests/test_xlsm_structure.py`, `tests/test_workbook_usage_links.py` |
| SECUI CLI 동작 | `tests/test_secui_cli_script.py`, `tests/test_secui_cli_guards.py` |
| SECUI GUI 동작 | `tests/test_secui_gui.py`, `tests/test_secui_policy_analysis.py`, `tests/test_secui_service_catalog.py` |
| 릴리스 번들 | `tests/test_release_bundle.py`, `tests/test_release_workflows.py` |
| VBA 런타임 가드 | `tests/test_vba_runtime_guards.py` |
| 보조 유틸리티 | `tests/test_cidr_parity.py`, `tests/test_date_cells.py`, `tests/test_folder_parse.py`, `tests/test_merged_cells.py`, `tests/test_request_folder.py`, `tests/test_user_alias.py`, `tests/test_weekly_report_sheet.py`, `tests/test_workbook_convenience.py` |

## 로컬 개발 (Local Development)

- 의존성 파일: `requirements.txt`(빌드/테스트), `requirements-gui.txt`(PySide6 GUI).
- Python 인터프리터: 저장소 커밋된 `./.venv/bin/python`을 우선 사용합니다.
- 빌드 가능 환경: Linux. Windows/Excel/COM/PowerShell 의존성은 필요 없습니다.
- 동기화 규칙:
  - 라우트/파서 로직을 바꾸면 VBA(`vba/*.bas`)와 Python 오라클(`tests/*_oracle.py`)이 동일한 결과를 내야 합니다.
  - 시트/헤더를 바꾸면 `scripts/workbook_contract.py`, VBA 시드, `tests/test_xlsm_structure.py`를 함께 갱신합니다.
  - UX/내비게이션 변경은 표시와 입력 보조에만 영향을 줍니다. 라우트/파서 결과는 절대 변경하지 않습니다.

## 기여 (Contribution)

1. 변경 의도(라우트/파서/워크북/릴리스)를 명확히 한 뒤 관련 오라클 테스트를 먼저 갱신합니다.
2. VBA와 Python 미러가 행동 동등을 유지하도록 오라클 테스트를 통과시킵니다.
3. 워크북 시트/헤더 변경 시 `tests/test_xlsm_structure.py`를 통과시킵니다.
4. PR 전 `./.venv/bin/python -m pytest tests/ -q`로 전체 회귀를 확인합니다.
5. 자세한 절차는 `CONTRIBUTING.md`를 참조합니다.

## 유지보수 (Maintainers)

저장소 운영/변경 책임자는 저장소 메타데이터(`AGENTS.md`, `CONTRIBUTING.md`)와 커밋 기록을 참조합니다. 본 저장소는 외부 운영 채널이나 컨트롤 플레인을 노출하지 않습니다.

## 참고 문서 (Further Documentation)

| 문서 | 위치 |
| --- | --- |
| Excel 네이티브 런북 | `docs/excel-native.md` |
| Excel 스키마 | `docs/excel-schema.md` |
| 방화벽 Excel 벤치마크 | `docs/firewall-excel-benchmark.md` |
| 연구 노트 | `docs/research-notes.md` |
| 저장소 지식 베이스 | `AGENTS.md` |
| VBA 모듈 안내 | `vba/AGENTS.md` |
| 패키지 안내 | `firewall_policy/AGENTS.md` |
| 스크립트 안내 | `scripts/AGENTS.md` |
| 테스트 안내 | `tests/AGENTS.md` |
| 신청 폴더 안내 | `request-folder/README.txt` |
| 기여 가이드 | `CONTRIBUTING.md` |
| 라이선스 | `LICENSE` |