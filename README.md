# Firewall Policy Automation | 방화벽 정책 자동화

DRM 환경에서 PowerQuery 없이 Excel VBA만으로 방화벽 정책 신청서를 통합하고, 사용자 정의 방화벽 통과 대역 기준으로 대상방화벽을 계산하는 Excel 네이티브 도구입니다.

![platform](https://img.shields.io/badge/platform-Excel%20%2B%20Linux-1f6feb)
![runtime](https://img.shields.io/badge/runtime-VBA%20%2B%20Python%203-3776AB)
![build](https://img.shields.io/badge/build-Linux%20only-3DDC84)
![license](https://img.shields.io/badge/license-see%20LICENSE-lightgrey)

## 한국어 요약

`firewalls`와 `firewall_ranges` 두 시트만 관리하면, 신청서 폴더의 모든 xlsx를 통합하고 대상방화벽을 계산해 SECUI CLI 초안까지 만들어 주는 매크로 포함 Excel 워크북입니다. Linux 빌드 파이프라인이 `.xlsm`을 생성하므로 Windows/Excel/COM 없이도 빌드/테스트가 가능합니다.

## Summary

A macro-enabled Excel workbook that merges firewall-policy request spreadsheets and computes target firewalls from operator-defined range rows, with no PowerQuery. A Linux pipeline produces the shipped `.xlsm`, and a Python CLI/GUI mirrors the same logic for operators who prefer working outside Excel.

## Status

| 항목 | 상태 | 비고 |
| --- | --- | --- |
| 제품 성숙도 | 운영 가능 | 매크로 활성 Excel 워크북 |
| 빌드 환경 | Linux 전용 | Windows/Excel/COM/PowerShell 불필요 |
| 데이터 입력 시트 | 2개 핵심 + 4개 보조 | `firewalls`, `firewall_ranges`가 유일한 경로 계산 기준 |
| 결과 시트 | 4개 | `requests`, `processing_log`, `route_results`, `secui_cli` |
| 외부 인터페이스 | CLI + GUI | `secui_cli.py`, `secui_gui.py` (PySide6) |
| 테스트 | pytest 오라클 + 구조 테스트 | VBA 동작과 Python 오라클이 동일해야 통과 |
| 의존성 | 핀 버전 | `pyOpenVBA 2.0.0`, `openpyxl 3.1.5`, `pytest 9.0.3`, `PySide6 6.11.1` |

## 빠른 흐름

1. **빌드**: `./.venv/bin/python scripts/build_xlsm.py` 실행 → `dist/firewall-policy-automation.xlsm` 생성
2. **데이터 입력**: `firewalls`, `firewall_ranges`, `settings` 시트를 운영자가 채움
3. **자동 실행**: 워크북 열기 시 `Workbook_Open`이 신청서 통합 → 경로탐색 → SECUI CLI 생성을 순차 실행
4. **CLI/GUI**: 워크북 외부에서는 `scripts/secui_cli.py` 또는 `scripts/secui_gui.py`로 동일 로직 사용

## 목차

- [목적 및 구성](#목적-및-구성)
- [먼저 읽을 파일](#먼저-읽을-파일)
- [진입점](#진입점)
- [빠른 시작](#빠른-시작)
- [워크북 시트 한눈에 보기](#워크북-시트-한눈에-보기)
- [입력 시트 명세](#입력-시트-명세)
- [아키텍처](#아키텍처)
- [명령어 레퍼런스](#명령어-레퍼런스)
- [로컬 개발](#로컬-개발)
- [테스트](#테스트)
- [기여](#기여)
- [메인테이너](#메인테이너)
- [라이선스](#라이선스)
- [추가 문서](#추가-문서)

## 목적 및 구성

방화벽 정책 신청은 보통 (1) 여러 부서에서 다양한 양식의 xlsx가 들어오고, (2) 출발지/목적지 대역이 어떤 방화벽을 거쳐야 하는지 계산이 필요하고, (3) 결과를 SECUI CLI로 변환해야 하는 단계로 구성됩니다. 이 도구는 그 세 단계를 Excel 워크북 한 파일 안에서 처리합니다.

핵심 설계:

- **Excel 네이티브**: PowerQuery, 외부 데이터 원본, COM 호출 없이 VBA만으로 동작하므로 DRM/외부연결 차단 환경에서도 동작합니다.
- **데이터 입력 최소화**: 운영자가 직접 편집할 시트는 `firewalls`와 `firewall_ranges` 두 개입니다. 기존 zone/routing_paths/BFS 모델은 제거되었고, 대상방화벽 계산의 유일한 기준은 `firewall_ranges`입니다.
- **이중 구현 + 오라클 테스트**: `vba/`의 VBA 로직과 `firewall_policy/`의 Python 미러가 동일 동작을 내도록 `tests/`의 오라클 테스트로 묶여 있습니다.
- **외부 인터페이스**: 워크북 밖에서도 동일한 로직을 쓸 수 있도록 `secui_cli.py`(CLI)와 `secui_gui.py`(PySide6 GUI)를 제공합니다.

### 패키지 구성

| 경로 | 역할 |
| --- | --- |
| `vba/` | `.xlsm`에 주입되는 VBA 모듈 (`FirewallPolicyAutomation.bas`, `FirewallRouteAnalysis.bas`) |
| `firewall_policy/` | VBA 동작의 Python 미러. SECUI CLI/GUI의 백엔드 |
| `scripts/` | Linux 빌더, SECUI CLI/GUI, 릴리스 번들러 |
| `tests/` | pytest 오라클/스펙 테스트, 워크북 구조 테스트 |
| `docs/` | 운영자용 스키마 및 Excel 네이티브 런북 |
| `request-folder/` | 샘플 신청서 트리 (`_빈양식/`은 빈 양식) |

## 먼저 읽을 파일

| 작업 | 위치 | 메모 |
| --- | --- | --- |
| 경로/대역 계산 로직 변경 | `tests/route_oracle.py`, `vba/FirewallRouteAnalysis.bas`, `tests/test_route_oracle.py` | Python 오라클과 VBA가 동일 동작 유지 |
| 워크북 시트/헤더/시드 변경 | `scripts/workbook_contract.py`, `scripts/build_xlsm.py`, `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py` | 빌드 시드와 구조 테스트 동기화 필요 |
| 워크북 UX/네비게이션 변경 | `scripts/workbook_ux.py`, `tests/test_xlsm_structure.py`, `tests/test_workbook_usage_links.py` | 표시/입력 보조 전용. 경로/파서 값은 변경 금지 |
| 신청서 파서 동작 변경 | `tests/request_parser_oracle.py`, `vba/FirewallPolicyAutomation.bas`, `tests/test_request_parsing.py` | 헤더 별칭과 병합 셀 동작 미러링 |
| 워크북 외부 SECUI CLI 생성 | `scripts/secui_cli.py`, `scripts/secui_cli_runtime.py`, `scripts/secui_cli_seed.py`, `firewall_policy/request_parser.py` | 워크북과 동일한 변환 로직, 그룹 규칙 생성 |
| SECUI GUI / 스탠드얼론 빌드 | `scripts/secui_gui.py`, `firewall_policy/gui_export.py`, `scripts/build_standalone_gui.py` | PySide6, 스탠드얼론 번들은 릴리스 산출물 |
| SECUI 출력 변경 | `vba/FirewallPolicyAutomation.bas`, `tests/test_xlsm_structure.py`, `docs/` | 배치/CLI는 `대상방화벽`을 `;`로 분리 |
| 빌드/릴리스 산출물 변경 | `scripts/build_xlsm.py`, `scripts/make_request_folder.py`, `tests/test_xlsm_structure.py` | CI/릴리스는 산출물 기반 |

## 진입점

| 진입점 | 호출자 | 동작 |
| --- | --- | --- |
| `Workbook_Open` 이벤트 | Excel | 워크북 열림 시 자동 실행: 신청서 통합 → 경로탐색 → SECUI CLI 생성 |
| `RunAll` 매크로 (`F9`) | 운영자 | 위 파이프라인 전체를 수동으로 한 번 실행 |
| `SelectRequestFolder` 매크로 | 운영자 | `settings.request_folder`를 폴더 선택 대화상자로 갱신 |
| `scripts/secui_cli.py` | CLI 사용자 | 워크북 외부에서 SECUI CLI 텍스트/배치 생성 |
| `scripts/secui_gui.py` | GUI 사용자 | PySide6 기반 단일 창 도구 |

## 빠른 시작

1. `dist/firewall-policy-automation.xlsm`을 열고 매크로를 허용합니다.
2. `firewalls` 시트에 장비명, 벤더, 사용여부를 등록합니다.
3. `firewall_ranges` 시트에 출발지/목적지 CIDR, 방향, 표시 순서를 등록합니다.
4. `settings` 시트의 `request_folder` 값에 신청서 폴더 경로를 입력하거나 `SelectRequestFolder` 매크로로 선택합니다.
5. 워크북을 다시 열면 `Workbook_Open`이 신청서 통합, 경로탐색, SECUI CLI 생성을 순서대로 자동 실행합니다.
6. `requests`에서 최종 신청 목록을, `route_results`에서 대상방화벽과 경로탐색 결과를 확인합니다.
7. 수동 실행이 필요하면 `F9`를 누릅니다. 신청서 파싱, 경로분석, SECUI CLI 생성을 한 번에 실행합니다.

## 워크북 시트 한눈에 보기

### 입력 시트

| 시트 | 역할 |
| --- | --- |
| `firewalls` | 방화벽 장비 메타데이터 |
| `firewall_ranges` | 방화벽별 출발지/목적지 통과 대역 정의 |
| `settings` | 신청서 폴더, 파싱 대상 시트, 헤더 별칭 참조 |
| `header_aliases` | 표준이 아닌 신청서 헤더 매핑 |
| `vendor_cli_templates` | 벤더별 CLI 명령 템플릿 |
| `service_catalog` | SECUI 서비스 표기 예시 (`tcp/443`, `udp/53` 등) |

### 결과 시트

| 시트 | 역할 |
| --- | --- |
| `requests` | 14열 사용자용 통합 신청 목록 |
| `processing_log` | 통합 처리 로그 |
| `route_results` | 경로탐색/대상방화벽/검증 결과 |
| `secui_cli` | SECUI CLI 명령 초안 |

### 참고 시트

| 시트 | 역할 |
| --- | --- |
| `sample-request-format` | 신청서 양식 예시 |
| `usage` | 워크북 내 사용 순서 안내 |

내부 추적 시트:

| 시트 | 가시성 | 역할 |
| --- | --- | --- |
| `_request_tracking` | 숨김 | `원본파일`, `원본행`, `요청폴더`, `제목`을 `request_row` 키로 보관 |

## 입력 시트 명세

### `requests` 시트 (14열 사용자 시트)

헤더는 2행, 데이터는 3행부터 시작합니다.

| # | 컬럼 | 설명 |
| --- | --- | --- |
| 1 | 요청부서 | 신청 부서명 |
| 2 | 요청번호 | 부서 내부 식별 번호 |
| 3 | 대상방화벽 | 계산 결과. 내부 필드명은 `target_firewalls`, 값은 `;`로 결합 |
| 4 | 출발지 | 출발지 IP/CIDR/목록 |
| 5 | 출발지설명 | 출발지 설명 |
| 6 | 목적지 | 목적지 IP/CIDR/목록 |
| 7 | 목적지설명 | 목적지 설명 |
| 8 | 프로토콜 | TCP/UDP/ICMP 등 |
| 9 | 포트 | 포트 번호 또는 범위 |
| 10 | 방향 | OUT/IN/BOTH |
| 11 | 용도 | 사용 목적 |
| 12 | 시작일 | 정책 시작일 |
| 13 | 종료일 | 정책 종료일 |
| 14 | 비고 | 비고 |

검증/추적 정보는 `requests`에 두지 않고 다음 위치에 보관됩니다.

| 정보 | 위치 |
| --- | --- |
| `원본파일`, `원본행`, `요청폴더`, `제목` | 숨김 시트 `_request_tracking` (`request_row` 키) |
| `검증상태`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거`, `원본파일`, `원본행` | `route_results` 시트 |

### `firewalls` 시트

| 컬럼 | 설명 | 예시 |
| --- | --- | --- |
| `firewall_name` | 장비 식별자. `firewall_ranges.firewall_name`과 일치해야 함 | `SECUI-FW-01` |
| `vendor` | 벤더. SECUI 출력은 `vendor == "