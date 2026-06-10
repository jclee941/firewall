# Firewall Policy Excel Automation

DRM 환경에서 PowerQuery 없이 Excel 네이티브 VBA만으로 방화벽 정책 신청서를 통합 관리하는 저장소입니다. 신청서 대역과 방화벽 대역의 단순 CIDR 겹침만 보지 않고, IP 대역으로 결정되는 zone과 zone 사이 라우팅 경로를 따라 어떤 방화벽을 통과하는지(적용대상방화벽 = 통과 방화벽 경로)를 분석합니다.

## 목표

- 신청자가 작성한 Excel 파일을 한 폴더에 모읍니다.
- `network_definitions` 시트의 IP 대역과 zone 매핑으로 출발지/목적지의 zone을 결정합니다.
- `routing_paths` 시트의 zone 간 라우팅 경로로 통과하는 방화벽을 순서대로 결정합니다.
- 매크로 사용 통합 문서(`.xlsm`)에서 버튼/매크로로 폴더를 선택하고 통합을 실행합니다.
- `requests` 시트에 통합 결과와 검증 상태, 통과 경로를 남깁니다.
- 원본 파일명과 원본 행 번호를 남겨 추적 가능하게 합니다.

## 저장소 구조

```text
.
├── vba/
│   └── FirewallPolicyAutomation.bas   # Excel VBA 매크로 모듈
├── docs/
│   ├── excel-native.md                # 설치/운영 절차
│   ├── excel-schema.md                # 시트/컬럼/검증 정의
│   └── research-notes.md
├── scripts/
│   └── build_xlsm.py                  # Linux에서 .xlsm 빌드 (pyOpenVBA + openpyxl)
├── tests/
│   ├── route_oracle.py                # Python 오라클: 라우팅 알고리즘
│   └── test_route_oracle.py           # 알고리즘 회귀 테스트
├── firewall-policy-automation.xlsx    # 베이스 템플릿
├── dist/                              # 빌드된 .xlsm 산출물 위치
├── .gitignore
└── README.md
```

## 빠른 시작

Linux/macOS/Windows 어디서나 아래 한 줄로 매크로가 내장된 `dist/firewall-policy-automation.xlsm`이 생성됩니다. Excel을 띄우거나 VBA 편집기를 열 필요가 없습니다.

```bash
./.venv/bin/python scripts/build_xlsm.py
```

산출물은 `dist/firewall-policy-automation.xlsm`입니다. 이 파일을 열면 `requests`, `network_definitions`, `firewalls`, `routing_paths`, `settings` 시트가 미리 생성되어 있습니다. 매크로는 `Alt+F8`(매크로 대화상자)로 실행하거나, 원하면 Excel에서 도형/버튼에 매크로를 직접 연결해 사용합니다. 이후 운영은 이 `.xlsm` 파일 하나로 진행합니다.

> 빌드 스크립트는 `pyOpenVBA`로 `vbaProject.bin`을 만들고 `openpyxl`로 매크로/시트/스타일을 합칩니다. Windows Excel이나 PowerShell은 필요 없습니다.

## 시트 구성

기본 템플릿 `firewall-policy-automation.xlsx`에는 아래 시트가 미리 구성되어 있습니다.

| 시트 | 역할 |
|---|---|
| `requests` | 신청서 통합 결과 (방화벽 경로/검증 포함) |
| `network_definitions` | IP 대역 -> zone 매핑 (대역정의) |
| `firewalls` | 방화벽 장비 목록 (vendor, comment) |
| `routing_paths` | zone 간 라우팅 경로 (존-존 간선) |
| `settings` | 신청서 폴더, 파싱 대상, 라우팅 옵션 등록 |
| `sample-request-format` | 신청서 양식 예시 |
| `processing_log` | 파일별 처리 결과/오류 로그 |
| `usage` | 사용 순서 |

진행 순서는 다음과 같습니다.

1. `dist/firewall-policy-automation.xlsm`을 엽니다.
2. `network_definitions` 시트에 IP 대역과 zone을 등록합니다.
3. `firewalls` 시트에 방화벽명, vendor, 비고를 등록합니다.
4. `routing_paths` 시트에 zone 사이 라우팅 경로(ingress/egress 인터페이스 포함)를 등록합니다.
5. `settings` 시트의 `request_folder`에 신청서 폴더 경로를 등록하거나 매크로 `SelectRequestFolder`를 실행합니다.
6. `settings` 시트의 `parse_sheet`를 확인합니다. 기본값은 비어 있고(헤더로 시트 자동 감지), 특정 시트만 파싱하려면 그 시트 이름을 정확히 적습니다.
7. `settings` 시트의 `route_legacy_fallback` 값을 정합니다. 기본값은 `FALSE`입니다.
8. 매크로 `MergeFirewallRequestFolder`를 실행합니다 (`Alt+F8`).
9. 라우트 분석만 다시 돌리려면 매크로 `AnalyzeRequestRoutes`를 실행합니다.

## 주요 매크로

| 매크로 | 역할 |
|---|---|
| `SetupFirewallAutomationWorkbook` | 운영용 시트(`requests`, `network_definitions`, `firewalls`, `routing_paths`, `settings`) 생성 및 초기화 |
| `CreateSampleRequestWorkbook` | 신청자가 사용할 샘플 신청서 Excel 생성 |
| `SelectRequestFolder` | 신청서 폴더를 선택해서 `settings` 시트에 등록 |
| `MergeFirewallRequestFolder` | 폴더 안 신청서를 통합하고 zone/라우팅 기반으로 적용 대상 방화벽과 검증 상태 자동 입력 |
| `AnalyzeRequestRoutes` | `requests` 시트에 이미 통합된 행을 다시 라우트 분석 (방화벽/zone 데이터 갱신 시 유용) |

## 데이터 입력 규칙

### network_definitions (대역정의)

IP 대역을 zone에 매핑합니다. 여러 IP/CIDR이 같은 zone에 속할 수 있습니다.

| 컬럼 | 설명 | 예시 |
|---|---|---|
| `network_name` | 대역 이름 | `internal-hq` |
| `network_cidr` | CIDR 또는 단일 IP | `10.10.0.0/16` |
| `zone` | 소속 zone | `INTERNAL` |
| `site` | 사이트/위치 (선택) | `본사` |
| `enabled` | 사용 여부 (TRUE/FALSE) | `TRUE` |

### firewalls (방화벽 장비)

장비 단위로 등록합니다. 컬럼은 `firewall_name`, `vendor`, `enabled`, `inside_cidr`, `outside_cidr`, `comment` 여섯 개입니다. **각 방화벽의 내부대역(`inside_cidr`)과 외부대역(`outside_cidr`)만 적으면**, `routing_paths`를 비워도 경로가 자동 생성됩니다. 한 방화벽의 외부대역과 다른 방화벽의 내부대역이 같은 CIDR이면(공유 전송대역) 두 방화벽이 자동으로 연결되어 멀티홉 경로가 됩니다. 신청서 IP는 가장 좌은(구체적) CIDR에 매칭되어 zone이 정해지고, BFS는 목적지 zone에서 멈춰 그것이 종단입니다. `routing_paths`에 행을 넣으면 그게 우선됩니다(고급).

| 컬럼 | 설명 | 예시 |
|---|---|
| `firewall_name` | 방화벽 식별자 | `FW-INTERNAL-01` |
| `vendor` | 제조사 (참고용) | `Fortinet` |
| `enabled` | 사용 여부 | `TRUE` |
| `inside_cidr` | 내부 대역 (CIDR/IP, `;`로 여러 개) | `10.10.0.0/16` |
| `outside_cidr` | 외부 대역 (CIDR/IP, `;`로 여러 개) | `172.16.0.0/16` |
| `comment` | 비고 | `본사 코어` |

### routing_paths (라우팅경로)

zone 사이의 라우팅 간선입니다. 한 행은 한 방향의 한 hop입니다.

| 컬럼 | 설명 | 예시 |
|---|---|---|
| `firewall_name` | 이 hop을 처리하는 방화벽 | `FW-INTERNAL-01` |
| `src_zone` | 출발 zone | `INTERNAL` |
| `dst_zone` | 도착 zone | `DMZ` |
| `ingress_if` | 인바운드 인터페이스 | `port1` |
| `egress_if` | 아웃바운드 인터페이스 | `port2` |
| `path_order` | 동일 zone-쌍에 hop이 여럿일 때 우선순위 | `10` |
| `enabled` | 사용 여부 | `TRUE` |

## requests 시트 자동 입력 컬럼

`MergeFirewallRequestFolder`와 `AnalyzeRequestRoutes`는 `requests` 시트에 아래 컬럼을 자동으로 채웁니다. `requests` 시트는 운영자가 읽기 쉬운 순서로 24개 컬럼(셀병합 없음)을 배치하며, 모든 헤더는 한글 표시명입니다. 표시 순서: `요청부서`, `요청번호`, `원본파일`, `원본행`, `검증상태`, `적용대상방화벽`, `출발지IP`, `출발지`, `목적지IP`, `목적지`, `프로토콜`, `포트`, `방향`, `용도`, `시작일`, `종료일`, `비고`, `검증메시지`, `방화벽경로`, `출발Zone`, `목적Zone`, `Zone경로`, `매칭근거`, `요청폴더`.

| 컬럼 | 설명 |
|---|---|
| `원본파일` | 원본 신청서 파일명 |
| `원본행` | 원본 신청서 행 번호 |
| `적용대상방화벽` | 통과 방화벽 경로. hop 순서를 세미콜론(`;`)으로 연결한 값. `MULTI_PATH`일 때는 우선 경로의 hop 순서를 그대로 둠 |
| `방화벽경로` | 통과한 방화벽을 등장 순서대로 보존한 경로 문자열. hop은 `>`로 연결 (예: `FW-INTERNAL-01>FW-DMZ-01`) |
| `검증상태` | 검증 상태 값 (아래 표 참고) |
| `검증메시지` | 검증 사유 또는 참고 메시지 |
| `매칭근거` | zone 결정과 라우팅 hop 선택 근거를 사람이 읽을 수 있게 풀어 쓴 문자열 |
| `출발Zone` | 출발지 IP/CIDR에서 longest-prefix match로 결정된 출발 zone |
| `목적Zone` | 목적지 IP/CIDR에서 longest-prefix match로 결정된 도착 zone |
| `Zone경로` | zone 그래프 BFS 최단경로의 zone 열. zone은 `>`로 연결 (예: `INTERNAL>DMZ>EXTERNAL`) |
| `요청부서` | 신청서 상위 폴더명의 팀/센터 부분 (예: `정보보호센터`) |
| `요청번호` | 신청서 상위 폴더명의 문서번호 부분 (예: `1234`) |
| `요청폴더` | 신청서가 들어 있던 원본 폴더명 전체 |

## 매칭 규칙 (적용대상방화벽 = 통과 방화벽 경로)

기존 버전은 출발지IP/목적지IP와 방화벽 `cidr_list`의 겹침으로 적용 대상을 골랐습니다. 새 버전은 다음과 같이 동작합니다.

1. `network_definitions`에서 출발지 IP/CIDR과 가장 길게 겹치는 대역을 찾습니다. 그 대역의 zone이 `source_zone`이 됩니다. 목적지도 같은 방식으로 `destination_zone`을 결정합니다.
2. `routing_paths`를 zone 그래프(방향 가중치 없음)로 보고, `source_zone`에서 `destination_zone`까지의 결정적 BFS 최단경로를 찾습니다. 결정적이란 같은 zone-쌍에 `path_order`가 작은 hop부터 시도한다는 뜻입니다.
3. 경로 위의 각 hop에서 `firewall_name`을 모읍니다. hop 순서는 그대로 보존해서 `firewall_path`와 `zone_path`에 들어갑니다.
4. `target_firewalls`는 hop에서 처음 등장한 방화벽만 남깁니다. 같은 방화벽이 경로에 두 번 나와도 한 번만 기록합니다.
5. 출발/도착 zone이 같으면 `INTRA_ZONE`이 됩니다. 같은 zone 안 트래픽은 별도 방화벽이 개입하지 않으므로 `target_firewalls`는 비어 있습니다.
6. zone 결정은 IP 단위가 아니라 CIDR 단위 longest-prefix match입니다. 신청서 값이 `10.10.0.0/16`이고 `network_definitions`에 `10.10.0.0/16`과 `10.10.0.0/15`가 둘 다 있으면 더 긴 `10.10.0.0/16`이 이깁니다.

여러 방화벽이 정상입니다. `MULTI_PATH`로 표시되더라도 우선 경로의 hop 순서를 따라 한 줄로 결과가 나옵니다.

## validation_status 값

| 값 | 의미 |
|---|---|
| `OK` | 출발/도착 zone 결정, 라우팅 경로 결정, 모든 hop 정상. 통과 방화벽이 한 줄로 채워짐 |
| `MULTI_PATH` | zone 그래프에서 최단 hop 수의 경로가 둘 이상. 우선 경로가 적용되며 `validation_message`에 후보 수를 남김 |
| `INTRA_ZONE` | 출발 zone과 도착 zone이 동일. 방화벽이 개입하지 않으므로 `target_firewalls`는 비어 있음 |
| `ZONE_UNRESOLVED` | 출발 또는 도착 IP/CIDR이 `network_definitions` 어느 대역과도 겹치지 않음. zone을 결정할 수 없음 |
| `NO_PATH` | zone 그래프에 출발에서 도착으로 가는 경로가 없음. 라우팅 데이터 보강 필요 |
| `DIRECTION_MISMATCH` | 신청서 방향(`IN`/`OUT`/빈값)과 그래프 방향이 맞지 않음. `validation_message`에 시도한 방향을 남김 |
| `LEGACY_FALLBACK` | 라우팅 데이터가 비어 있거나 경로를 못 찾아서 기존 CIDR 겹침 방식으로 폴백함. `settings.route_legacy_fallback=TRUE`일 때만 발생 |

## direction (방향) 의미

신청서 `방향` 컬럼은 CIDR 계산의 입력이 아니라 라우팅 방향 제약입니다.

| 값 | 의미 |
|---|---|
| 빈값 또는 `BOTH` | 정방향 경로를 먼저 시도하고, 없으면 출발/도착을 뒤집어 역방향 경로를 다시 찾음 |
| `OUT` | 출발 -> 도착 한 방향만 시도 |
| `IN` | 도착 -> 출발 한 방향만 시도 (외부에서 내부로 들어오는 트래픽 표현) |

## settings 시트 키

| key | 설명 | 기본값 |
|---|---|---|
| `request_folder` | 신청서 폴더 경로 | 비어 있음 (실행 시 폴더 선택창) |
| `parse_sheet` | 파싱할 시트 이름(정확히 일치). 비우면 헤더로 자동 감지 | 비어 있음 |
| `parse_targets` | (사용 안 함/예약) 현재 동작에 영향 없음. 출발지IP·목적지IP는 항상 필수 | `출발지IP;목적지IP` |
| `route_legacy_fallback` | 라우팅 실패 시 기존 CIDR 겹침 방식으로 폴백할지 여부 (TRUE/FALSE) | `FALSE` |
| `header_alias` | 비표준 헤더 사용자 별칭. 형식: `출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소` | 비어 있음 |

`route_legacy_fallback`이 `FALSE`이면 라우팅 데이터가 없거나 경로를 못 찾을 때 그 행은 `NO_PATH` 또는 `ZONE_UNRESOLVED`로 남고 CIDR 폴백을 하지 않습니다. 기존 양식을 임시로 받아야 할 때만 `TRUE`로 두세요.

## 신청서 기본 컬럼

| 컬럼 | 설명 | 예시 |
|---|---|---|
| 출발지IP | 출발지 IP 또는 CIDR | 10.10.10.0/24 |
| 출발지 | 출발지 이름/설명 | 업무PC |
| 목적지IP | 목적지 IP 또는 CIDR | 172.16.1.10 |
| 목적지 | 목적지 이름/설명 | 업무시스템 |
| 프로토콜 | TCP/UDP/ICMP 등 | TCP |
| 포트 | 포트 또는 포트 범위 | 443 |
| 방향 | 정책 방향 (OUT/IN/빈값) | OUT |
| 용도 | 신청 목적 | HTTPS 업무 연동 |
| 시작일 | 시작일 | 2026-01-01 |
| 종료일 | 종료일 | 2026-12-31 |
| 비고 | 비고 | 정기 신청 |

## 신청서 폴더의 xlsx 구조

현재 매크로는 폴더 안 각 `.xlsx/.xlsm/.xls` 파일에서 **헤더 내용으로 헤더 행을 자동 인식**합니다. 즉 `출발지IP`/`목적지IP` 등 알아볼 수 있는 컬럼명이 가장 많은 행(최소한 출발지IP 또는 목적지IP 열 포함)을 헤더 행으로 봅니다. 데이터가 **첫 번째 시트가 아니어도** 됩니다. 매크로가 모든 시트를 훑어 헤더 점수가 가장 높은 시트를 자동으로 골라 파싱하므로, 결재/표지 시트가 앞에 있고 신청 내역이 뒤 시트에 있는 양식도 처리합니다(동점이면 가장 왼쪽 시트). `No`/`번호` 열은 있으면 도움이 되지만 **필수는 아닙니다**. `No.`·`순번`·`연번`·`순 번` 같은 변형이나 `No` 열이 아예 없는 양식도 처리하며, B열이 `No`이고 C/D열부터 출발지 정보가 시작되는 양식도 지원합니다.

헤더 행에는 아래 의미의 컬럼이 있어야 합니다.

```text
출발지IP | 출발지 | 목적지IP | 목적지 | 프로토콜 | 포트 | 방향 | 용도 | 시작일 | 종료일 | 비고
```

열 순서는 달라도 됩니다. 매크로가 헤더 행의 컬럼명을 별칭 맵으로 찾아서 읽습니다. 공백과 `.`·`#`·`( )` 같은 장식 구두점은 무시합니다. 예: `출발지 IP`·`No.` 모두 인식합니다.

일부 다른 표현도 흡수합니다.

| 표준 컬럼 | 허용 예시 |
|---|---|
| 출발지IP | 출발지 IP, 출발IP, Src IP |
| 출발지 | 출발지명, 출발 |
| 목적지IP | 목적지 IP, 목적IP, Dst IP |
| 목적지 | 목적지명, 목적 |
| 프로토콜 | Protocol, Proto |
| 포트 | Port, 목적지포트 |
| 용도 | 목적, Usage, Purpose |

`settings` 시트의 `header_alias`에 등록하면 내장 별칭으로는 안 잡히는 회사별 비표준 헤더도 표준 컬럼으로 매핑합니다. (내장 별칭이 먼저 적용되고, 못 잡은 헤더만 사용자 별칭으로 보완됩니다.)

## 병합 셀 / 하위폴더 / 폴더명 파싱

### 세로 병합 셀

신청서에서 한 신청 건이 여러 행에 걸쳐 `출발지IP`/`목적지IP` 셀을 **세로 병합**한 양식도 지원합니다. 병합 영역의 값이 아래 행에도 적용되어 행이 누락되지 않습니다. (데이터 셀은 병합 좌상단 값을 `MergeArea`로 읽고, 헤더 행은 좌상단만 인식합니다.) 제목/그룹 헤더가 헤더 위에 병합되어 있어도 헤더 내용 기반 인식으로 실제 헤더 행을 찾아 우회합니다.

> 참고: 가로 병합된 **헤더 셀**(예: `비고`를 L1:M1로 병합)은 Excel이 병합 좌상단 셀에만 값을 저장하므로 좌상단 컬럼만 헤더로 인식됩니다. 실무에서 헤더 가로병합은 권장하지 않습니다. (신청 데이터의 세로 병합은 완전히 지원됩니다.)

### 하위폴더 재귀

`request_folder` 아래의 하위폴더까지 재귀적으로 탐색해 신청서를 모읍니다. 팀/센터별로 폴더를 나눠 제출받는 운영을 그대로 처리합니다.

### 폴더명 파싱 (팀/문서번호)

신청서의 상위 폴더명을 마지막 `_` 기준으로 나누어 `request_team`/`request_doc_no` 컬럼에 기록합니다. 결재문서 내용은 파싱하지 않고 폴더명만 사용합니다.

| 폴더명 | request_team | request_doc_no |
|---|---|---|
| `정보보호센터_1234` | 정보보호센터 | 1234 |
| `정보보호_2팀_5678` | 정보보호_2팀 | 5678 |
| `인프라팀` | 인프라팀 | (빈값) |

## 검증과 QA

이 저장소의 LibreOffice는 매크로 실행 검증이 불가능한 상태입니다. 그래서 다음 두 단계로 알고리즘과 산출물을 검증합니다.

- 알고리즘: `tests/test_route_oracle.py`가 `tests/route_oracle.py`의 Python 오라클을 회귀 테스트합니다. 매크로와 같은 결정적 BFS/longest-prefix 로직을 Python에서 다시 구현해 결과를 비교합니다. 변경 후 `pytest`를 한 번 돌려 통과를 확인합니다.
- 산출물: 빌드된 `dist/firewall-policy-automation.xlsm`은 zip 안의 `xl/vbaProject.bin`, 모듈명(`FirewallPolicyAutomation`), 시트 목록, 헤더 행으로 구조를 검증합니다. Excel을 실제로 띄우지 않아도 매크로가 내장됐는지와 시트가 맞게 들어갔는지를 확인할 수 있습니다.

### 빌드 / 릴리즈 자동화

- 로컬 빌드: `./.venv/bin/python scripts/build_xlsm.py` → `dist/firewall-policy-automation.xlsm` 생성 (Excel/PowerShell 불필요, Linux에서 동작).
- CI(`.github/workflows/ci.yml`): master/main push와 PR마다 의존성 설치 → `pytest` → 빌드 → 산출물 구조 검증.
- 릴리즈(`.github/workflows/release.yml`): `v*` 태그를 push하면 자동으로 빌드·테스트·검증 후 GitHub 릴리즈를 생성하고 `.xlsm`을 첨부합니다.

```bash
# 새 버전 배포 (빌드·테스트·릴리즈 전부 자동)
git tag v1.0.2 && git push origin v1.0.2
```

의존성 버전은 `requirements.txt`에 고정되어 있어 CI 재현성이 보장됩니다.

## DRM 관련 전제

Excel이 DRM 파일을 열 수 있어야 VBA도 읽을 수 있습니다. DRM이 파일 열기 자체를 차단하면 PowerQuery뿐 아니라 VBA도 내용을 읽을 수 없습니다.
