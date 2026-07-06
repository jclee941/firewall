# 방화벽 정책 자동화 / Firewall Policy Automation

![Build: xlsm on Linux](https://img.shields.io/badge/build-Linux%20%2B%20PyOpenVBA-2.0.0-2e7d32)
![Runtime: Excel 2016+](https://img.shields.io/badge/runtime-Excel%202016%2B%20VBA-1f6feb)
![Tests: pytest 9.0.3](https://img.shields.io/badge/tests-pytest%209.0.3-0a7c2f)
![License: see LICENSE](https://img.shields.io/badge/license-see%20LICENSE-555555)

## 한 줄 요약

DRM 환경에서 PowerQuery 없이 Excel VBA만으로 방화벽 정책 신청서를 통합하고,
사용자가 정의한 `firewall_ranges` 통과 대역 기준으로 `대상방화벽`을 계산하는 도구입니다.
빌드는 Linux에서, 실행은 Excel 매크로 워크북에서 이루어집니다.

## One-line Summary

A Linux-built, Excel-native VBA tool that merges firewall policy request workbooks
and computes the target firewall from user-defined `firewall_ranges` rows — without
PowerQuery, without an installed Excel, and without an internet connection.

## 상태 / Status

| 항목 / Item | 상태 / Status | 비고 / Notes |
| --- | --- | --- |
| 릴리스 / Release | 운영 가능 / Production-ready | `v*` 태그 → GitHub Actions가 `.xlsm`과 GUI 번들을 게시 |
| 런타임 / Runtime | Excel 2016 이상 + 매크로 허용 | DRM 오프라인 환경에서 동작 |
| 빌드 / Build | Linux + Python 3.x | PowerShell·COM·Office 설치 불필요 |
| 핵심 의존성 / Pinned deps | `pyOpenVBA==2.0.0`, `openpyxl==3.1.5`, `pytest==9.0.3` | `requirements.txt` 참고 |
| GUI 의존성 / GUI deps | `PySide6==6.11.1` | `requirements-gui.txt` 참고 |
| 네트워크 / Network | 불필요 / None | 모든 산출물은 워크북 내부에 저장 |
| 지원 / Support | 유지보수 중 / Actively maintained | 이슈·PR 라벨로 분류 운영 |

## 3분 사용 흐름 / 3-minute operator flow

1. `dist/firewall-policy-automation.xlsm`을 열고 매크로를 허용합니다.
2. `firewalls` 시트에 장비명·벤더·사용여부를 등록합니다.
3. `firewall_ranges` 시트에 출발지/목적지 대역·방향·표시 순서를 등록합니다.
4. `settings.request_folder`에 신청서 폴더를 입력하거나 `SelectRequestFolder` 매크로로 선택합니다.
5. 통합 문서를 다시 열면 `Workbook_Open`이 신청서 통합 → 경로탐색 → SECUI CLI 생성을 순서대로 자동 실행합니다.
6. `requests`에서 최종 신청 목록을, `route_results`에서 대상방화벽과 검증 결과를 확인합니다.
7. 수동 재실행이 필요하면 `F9` (전체) 또는 시트 우클릭 컨텍스트 메뉴를 사용합니다.

## 목차 / Contents

- [시작하기 전에 / Before You Start](#시작하기 전에--before-you-start)
- [패키지 구성 / Package Contents](#패키지-구성--package-contents)
- [워크북 컨트랙트 / Workbook Contract](#워크북-컨트랙트--workbook-contract)
- [입력 시트 / Input Sheets](#입력-시트--input-sheets)
- [결과 시트 / Result Sheets](#결과-시트--result-sheets)
- [빠른 시작 / Quickstart](#빠른-시작--quickstart)
- [명령어 / Commands](#명령어--commands)
- [설정 / Configuration](#설정--configuration)
- [테스트 / Testing](#테스트--testing)
- [로컬 개발 / Local Development](#로컬-개발--local-development)
- [기여 / Contributing](#기여--contributing)
- [도움말 / Getting Help](#도움말--getting-help)
- [유지보수 / Maintainers](#유지보수--maintainers)
- [라이선스 / License](#라이선스--license)
- [추가 문서 / Further Documentation](#추가-문서--further-documentation)

## 시작하기 전에 / Before You Start

- Excel 2016 이상(Windows/macOS)에서 매크로를 실행할 수 있어야 합니다.
- DRM/보안 PC에서는 매크로 보안 설정을 운영 정책에 맞게 조정해야 합니다.
- 신청서 폴더는 Excel 프로세스가 읽을 수 있는 로컬/매핑된 경로여야 합니다.
- 빌드 환경(Linux)에는 Python 3.x와 커밋된 `./.venv`가 필요합니다.

## 패키지 구성 / Package Contents

| 경로 / Path | 역할 / Role | 비고 / Notes |
| --- | --- | --- |
| `vba/` | 워크북에 삽입되는 VBA 모듈 원본 | `FirewallPolicyAutomation.bas`, `FirewallRouteAnalysis.bas` |
| `firewall_policy/` | VBA 동작의 순수 Python 미러 | 파서·CIDR·폴더·GUI 익스포트 |
| `scripts/` | Linux 빌더, CLI, GUI, 릴리스 번들러 | PySide6 GUI 포함 |
| `tests/` | Python 오라클 + 워크북 구조 테스트 | pytest 9.0.3 |
| `docs/` | 운영자 스키마/런북 | `excel-native.md`, `excel-schema.md` 등 |
| `request-folder/` | 신청서 샘플 트리 | `_빈양식/`이 빈 양식 보관 위치 |
| `dist/` | 빌드 산출물 | GitHub Actions가 생성 |
| `requirements.txt` | 빌드·테스트 의존성 | PyOpenVBA, openpyxl, pytest |
| `requirements-gui.txt` | GUI 의존성 | PySide6 |
| `AGENTS.md` | 프로젝트 지식 베이스 | 에이전트/기여자 컨텍스트 |
| `CONTRIBUTING.md` | 기여 가이드 | PR·이슈 규약 |
| `LICENSE` | 라이선스 전문 | 사내 라이선스 |

## 워크북 컨트랙트 / Workbook Contract

- `requests`는 14컬럼 사용자 시트입니다. 헤더는 2행, 데이터는 3행부터 시작합니다.
  컬럼 순서: `요청부서, 요청번호, 대상방화벽, 출발지, 출발지설명, 목적지, 목적지설명, 프로토콜, 포트, 방향, 용도, 시작일, 종료일, 비고`
- 내부 결과 필드명은 `target_firewalls`이며 `requests`에는 `대상방화벽`으로 표시됩니다.
  다중 매칭은 `;`로 결합됩니다.
- 원본 추적(`원본파일`, `원본행`, `요청폴더`, `제목`)은 숨김 시트 `_request_tracking`에 보관되며
  `request_row`로 키잉됩니다.
- 경로 진단(`검증상태`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`,
  `대역경로`, `매칭근거`, `원본파일`, `원본행`)은 `route_results`에 기록됩니다.
- 운영자가 직접 관리하는 입력/참조 시트는 `firewalls`, `firewall_ranges`, `settings`,
  `header_aliases`, `vendor_cli_templates`, `service_catalog`입니다.
- `network_definitions`, `routing_paths`, zone 그래프 BFS, 레거시 폴백은 더 이상 사용하지 않습니다.
  `대상방화벽` 계산의 유일한 기준은 `firewall_ranges`입니다.

## 입력 시트 / Input Sheets

| 시트 / Sheet | 역할 / Role | 운영자 작업 / Operator action |
| --- | --- | --- |
| `firewalls` | 방화벽 장비 메타데이터 | 장비 추가/비활성화 |
| `firewall_ranges` | 출발지/목적지 통과 대역 | 통과 대역 정의·순서 지정 |
| `settings` | 신청서 폴더, 파싱 대상 시트, 헤더 별칭 | 경로·별칭 관리 |
| `header_aliases` | 비표준 신청서 헤더 매핑 | 양식이 다른 부서 대응 |
| `vendor_cli_templates` | 벤더별 CLI 템플릿 | SECUI 외 벤더 확장 시 |
| `service_catalog` | SECUI 서비스 표기 예시 | `tcp/443`, `udp/53` 등 |

### `firewalls` 컬럼

| 컬럼 | 설명 | 예시 |
| --- | --- | --- |
| `firewall_name` | 장비 식별자. `firewall_ranges.firewall_name`과 일치해야 함 | `SECUI-FW-01` |
| `vendor` | 벤더명. SECUI 출력은 `SECUI` 장비만 사용 | `SECUI` |
| `enabled` | 사용 여부. `Y`, `YES`, `TRUE`, `1` 중 하나면 사용 | `Y` |
| `comment` | 설명 메모 | `내부-서버 구간` |

### `firewall_ranges` 컬럼

| 컬럼 | 설명 | 예시 |
| --- | --- | --- |
| `firewall_name` | 적용 방화벽. `firewalls.firewall_name`과 일치 | `SECUI-FW-01` |
| `source_cidr` | 출발지 IP/CIDR/목록. `ANY` 허용 | `<homelab-host>/16` |
| `destination_cidr` | 목적지 IP/CIDR/목록. `ANY` 허용 | `<homelab-host>/16` |
| `direction` | `OUT`, `IN`, `BOTH`, 빈값은 `BOTH` | `OUT` |
| `path_order` | 동일 매칭에서 우선순위. 정수, 오름차순 | `1` |
| `enabled` | 사용 여부. 미사용은 매칭에서 제외 | `Y` |
| `comment` | 설명 메모 | `DMZ → 내부` |

### `settings` 핵심 키

| 키 / Key | 설명 | 기본값 / Default |
| --- | --- | --- |
| `request_folder` | 신청서 루트 폴더. 비어 있으면 `SelectRequestFolder` 호출 | 빈 값 |
| `parse_sheet_name` | 신청서에서 읽을 시트 이름 후보 | 자동 감지 |
| `header_aliases` | 표준 외 헤더 별칭 참조 | `header_aliases` 시트 |

## 결과 시트 / Result Sheets

| 시트 / Sheet | 역할 / Role | 갱신 시점 / Refresh |
| --- | --- | --- |
| `requests` | 최종 통합 신청 목록 (14컬럼) | `Workbook_Open`, `F9` |
| `processing_log` | 통합 처리 로그 | 매 파이프라인 실행 |
| `route_results` | 대상방화벽·경로탐색·검증 결과 | 파이프라인 실행 시 |
| `secui_cli` | SECUI CLI 명령 초안 | 파이프라인 실행 시 |

## 빠른 시작 / Quickstart

### 운영자 (Excel 사용자) / Operator path

```text
1. dist/firewall-policy-automation.xlsm 더블클릭
2. 매크로 사용 허용
3. firewalls, firewall_ranges 시트에 데이터 입력
4. settings.request_folder 지정 (또는 SelectRequestFolder 매크로)
5. 통합 문서 저장 후 다시 열기 → Workbook_Open 자동 실행
6. requests / route_results / secui_cli 확인
```

### 빌드 담당자 (Linux) / Build path

```bash
# 1) 의존성 설치
python -m venv .venv
./.venv/bin/pip install -r requirements.txt
./.venv/bin/pip install -r requirements-gui.txt   # GUI 번들까지 빌드할 경우

# 2) 테스트
./.venv/bin/python -m pytest tests/ -q

# 3) 매크로 워크북 빌드
./.venv/bin/python scripts/build_xlsm.py
# 산출물: dist/firewall-policy-automation.xlsm

# 4) 샘플 신청서 폴더 만들기
./.venv/bin/python scripts/make_request_folder.py

# 5) 릴리스 번들 (선택)
./.venv/bin/python scripts/build_release_bundle.py
```

## 명령어 / Commands

| 명령 / Command | 용도 / Purpose |
| --- | --- |
| `./.venv/bin/python -m pytest tests/ -q` | 전체 테스트 실행 |
| `./.venv/bin/python -m pytest tests/test_route_oracle.py -v` | 경로 오라클만 실행 |
| `./.venv/bin/python scripts/build_xlsm.py` | `.xlsm` 빌드 |
| `./.venv/bin/python scripts/make_request_folder.py` | 신청서 샘플 트리 생성 |
| `./.venv/bin/python scripts/secui_cli.py --workbook dist/firewall-policy-automation.xlsm --format text` | Excel 없이 SECUI CLI 생성 |
| `./.venv/bin/python scripts/secui_cli.py --workbook … --format json` | JSON 출력 |
| `./.venv/bin/python scripts/secui_gui.py` | PySide6 GUI 실행 |
| `./.venv/bin/python scripts/build_standalone_gui.py` | 단일 실행 GUI 번들 빌드 |
| `./.venv/bin/python scripts/build_release_bundle.py` | 릴리스 산출물 묶음 |
| `./.venv/bin/python scripts/route_seed.py` | 라우트 시드 데이터 생성 |

## 설정 / Configuration

| 설정 / Setting | 위치 / Location | 비고 / Notes |
| --- | --- | --- |
| 신청서 폴더 | `settings.request_folder` | 미설정 시 `SelectRequestFolder` 매크로 사용 |
| 파싱 대상 시트 | `settings.parse_sheet_name` | 빈 값이면 워크북별 자동 감지 |
| 헤더 별칭 | `header_aliases` 시트 | 부서별 비표준 양식 대응 |
| 벤더 CLI 템플릿 | `vendor_cli_templates` 시트 | SECUI 외 벤더 확장 시 |
| 서비스 표기 | `service_catalog` 시트 | `tcp/443`, `udp/53` 등 SECUI 표기 매핑 |
| 매크로 보안 | Excel Trust Center | 운영 정책에 따라 조정 |
| 빌드 핀 | `requirements.txt`, `requirements-gui.txt` | 버전 고정, 임의 업그레이드 금지 |

## 테스트 / Testing

| 테스트 / Test | 책임 / Concern | 핵심 픽스처 / Key fixture |
| --- | --- | --- |
| `tests/test_route_oracle.py` | VBA와 Python 경로 계산 동치성 | `tests/route_oracle.py` |
| `tests/test_request_parsing.py` | 신청서 파서 동작 | `tests/request_parser_oracle.py` |
| `tests/test_folder_parse.py` | 신청서 폴더 트리 파싱 | `tests/folder_parse_oracle.py` |
| `tests/test_user_alias.py` | 사용자 별칭 처리 | `tests/user_alias_oracle.py` |
| `tests/test_cidr_parity.py` | CIDR/대역 일치성 | `firewall_policy/cidr.py` |
| `tests/test_xlsm_structure.py` | 산출 워크북 시트·헤더 구조 | `scripts/workbook_contract.py` |
| `tests/test_merged_cells.py` | 병합 셀 처리 | `FirewallPolicyAutomation.bas` |
| `tests/test_date_cells.py` | 날짜 셀 직렬화 | Excel 날짜 직렬화 |
| `tests/test_weekly_report_sheet.py` | 주간 리포트 시트 | `scripts/workbook_ux.py` |
| `tests/test_workbook_convenience.py` | 워크북 편의 동작 | `scripts/workbook_ux.py` |
| `tests/test_workbook_usage_links.py` | `usage` 시트 링크 검증 | `scripts/workbook_ux.py` |
| `tests/test_secui_cli_guards.py` | SECUI CLI 입력 검증 | `scripts/secui_cli.py` |
| `tests/test_secui_cli_script.py` | SECUI CLI 스크립트 동작 | `scripts/secui_cli.py` |
| `tests/test_secui_policy_analysis.py` | 정책 분석 | `scripts/secui_cli_runtime.py` |
| `tests/test_secui_service_catalog.py` | 서비스 카탈로그 | `service_catalog` 시트 |
| `tests/test_secui_gui.py` | PySide6 GUI | `scripts/secui_gui.py` |
| `tests/test_vba_runtime_guards.py` | VBA 런타임 가드 | `vba/` |
| `tests/test_build_route_results.py` | `route_results` 빌드 | `scripts/build_xlsm.py` |
| `tests/test_release_bundle.py` | 릴리스 번들 형태 | `scripts/build_release_bundle.py` |
| `tests/test_release_workflows.py` | 릴리스 워크플로 계약 | `.github/workflows/` |

테스트 실행 전 `./.venv`를 활성화하거나 `requirements.txt`를 설치하세요.
VBA와 Python은 `route_oracle.py` / `request_parser_oracle.py`를 기준으로 행동 동치성을 유지합니다.

## 로컬 개발 / Local Development

1. 저장소를 클론하고 `./.venv`에 의존성을 설치합니다.
2. VBA 모듈을 수정한 경우 `scripts/build_xlsm.py`로 `.xlsm`을 재생성합니다.
3. Python 미러(`firewall_policy/`)를 수정한 경우 동일 동작이 VBA에 반영되어야 합니다.
   오라클 테스트가 두 구현의 차이를 감지합니다.
4. GUI를 수정한 경우 `requirements-gui.txt`로 PySide6를 설치한 뒤
   `scripts/build_standalone_gui.py`로 단일 실행 번들을 생성합니다.
5. `scripts/workbook_contract.py`와 `scripts/workbook_ux.py`는 워크북 시트/UX의 단일 출처입니다.
   새 시트를 추가하면 `tests/test_xlsm_structure.py`도 함께 갱신해야 합니다.

권장 워크플로:

```text
코드 수정 → ./venv pytest → build_xlsm → 산출 .xlsm 검증 → PR
```

## 기여 / Contributing

`CONTRIBUTING.md`의 PR·이슈 규약을 따릅니다. 핵심 원칙은 다음과 같습니다.

- VBA(`vba/`)와 Python(`firewall_policy/`, `scripts/`) 동작은 오라클 테스트로 묶여 있습니다.
  한쪽만 변경하면 테스트가 실패합니다.
- 시트 구조·헤더 변경은 `scripts/workbook_contract.py`와 `tests/test_xlsm_structure.py`를
  함께 수정해야 합니다.
- UX/네비게이션 변경은 `scripts/workbook_ux.py`만 수정하고,
  라우트·파서 값을 변경해서는 안 됩니다.

## 도움말 / Getting Help

| 채널 / Channel | 용도 / Use for |
| --- | --- |
| `docs/excel-native.md` | Excel 워크북 운영 런북 |
| `docs/excel-schema.md` | 시트·컬럼 스키마 |
| `docs/firewall-excel-benchmark.md` | 벤치마크/성능 메모 |
| `docs/research-notes.md` | 설계 메모, 트레이드오프 |
| `AGENTS.md` | 프로젝트 지식 베이스, 변경 위치 가이드 |
| `request-folder/README.txt` | 신청서 폴더 규약 |
| GitHub Issues | 버그 리포트, 개선 제안 |
| GitHub Discussions | 사용 패턴·질문 |

## 유지보수 / Maintainers

이 저장소는 사내 운영팀이 유지보수합니다. 일상 운영, 빌드, 릴리스는
`.github/workflows/`의 CI와 `v*` 태그 트리거로 자동화됩니다.

- 코드 오너십: `vba/`, `firewall_policy/`, `scripts/`, `tests/` — 저장소 Contributors
- 릴리스 / 빌드: GitHub Actions (`v*` 태그 시 자동 게시)
- 사내 담당자: 저장소 CODEOWNERS 참고

## 라이선스 / License

`LICENSE` 파일의 내용을 따릅니다. 사내 사용 조건이 명시되어 있을 수 있으므로
배포·수정 전에 전문을 확인하세요.

## 추가 문서 / Further Documentation

- 워크북 운영: [`docs/excel-native.md`](docs/excel-native.md)
- 시트/컬럼 스키마: [`docs/excel-schema.md`](docs/excel-schema.md)
- 벤치마크: [`docs/firewall-excel-benchmark.md`](docs/firewall-excel-benchmark.md)
- 설계 메모: [`docs/research-notes.md`](docs/research-notes.md)
- 신청서 폴더 규약: [`request-folder/README.txt`](request-folder/README.txt)
- 에이전트/기여자 컨텍스트: [`AGENTS.md`](AGENTS.md)
- 기여 규약: [`CONTRIBUTING.md`](CONTRIBUTING.md)
- VBA 모듈: [`vba/FirewallPolicyAutomation.bas`](vba/FirewallPolicyAutomation.bas),
  [`vba/FirewallRouteAnalysis.bas`](vba/FirewallRouteAnalysis.bas)
- Python 미러: [`firewall_policy/`](firewall_policy)
- 빌더·CLI·GUI: [`scripts/`](scripts)