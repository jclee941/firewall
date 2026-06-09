# Excel 네이티브 자동화 방식

DRM 때문에 PowerQuery가 막혀 있으면 Excel 내부 VBA 매크로 방식이 가장 현실적입니다. 이 저장소는 Excel 네이티브 방식만 대상으로 하며, 별도 Python 실행 없이 매크로 사용 통합 문서(`.xlsm`) 안에서 폴더의 신청서를 열고 통합합니다.

## 빌드 산출물

`dist/firewall-policy-automation.xlsm`이 매크로 사용 통합 문서입니다. 이 파일 안에서 폴더 선택, 통합, 라우팅 분석이 모두 돌아갑니다. `.xlsm` 자체는 운영용 워크북이지 사람이 편집할 원본이 아닙니다. 원본 소스는 `vba/FirewallPolicyAutomation.bas`이고, 빌드 스크립트는 `scripts/build_xlsm.py`입니다.

## 구성

운영 워크북에는 다음 시트가 생성됩니다.

| 시트 | 역할 |
|---|---|
| `requests` | 신청서 통합 결과. 자동 입력 컬럼 포함 |
| `network_definitions` | IP 대역과 zone 매핑 (대역정의) |
| `firewalls` | 방화벽 장비 목록 (vendor, comment) |
| `routing_paths` | zone 사이 라우팅 경로 (존-존 간선) |
| `settings` | 운영 설정 (폴더, 파싱 대상, 라우팅 옵션) |
| `processing_log` | 파일별 처리 결과/오류 로그 |

## 매크로 목록

| 매크로 | 역할 |
|---|---|
| `SetupFirewallAutomationWorkbook` | 운영용 시트 생성 및 초기화 |
| `CreateSampleRequestWorkbook` | 신청자 배포용 샘플 Excel 생성 |
| `SelectRequestFolder` | 신청서 폴더 등록 |
| `MergeFirewallRequestFolder` | 신청서 폴더 통합, zone/라우팅 분석, 검증 상태와 처리 로그 작성 |
| `AnalyzeRequestRoutes` | 통합된 행을 라우팅 분석만 다시 수행 (방화벽/zone 데이터 갱신 후 호출) |

## 설치 방법

1. 저장소를 클론합니다.
2. Python 가상환경을 준비하고 의존성을 설치합니다.
   ```bash
   python3 -m venv .venv
   ./.venv/bin/pip install pyOpenVBA openpyxl
   ```
3. 빌드 스크립트를 실행해 매크로 내장 `.xlsm`을 만듭니다.
   ```bash
   ./.venv/bin/python scripts/build_xlsm.py
   ```
   산출물은 `dist/firewall-policy-automation.xlsm`입니다.
4. `dist/firewall-policy-automation.xlsm`을 엽니다.
5. 매크로 `SetupFirewallAutomationWorkbook`를 한 번 실행해 운영용 시트를 초기화합니다 (`Alt+F8` 매크로 대화상자). 빌드된 `.xlsm`은 시트가 이미 구성되어 있으므로 이 단계는 선택입니다.
6. `network_definitions`, `firewalls`, `routing_paths` 시트에 실제 데이터를 입력합니다.
7. `settings` 시트의 `request_folder`에 신청서 폴더 경로를 등록합니다.
8. `settings` 시트의 `parse_sheet` 값을 확인합니다. 비어 두면 헤더로 시트를 자동 감지하고, 특정 시트만 파싱하려면 그 시트 이름을 적습니다.
9. `settings` 시트의 `route_legacy_fallback` 값을 정합니다. 기본값은 `FALSE`입니다.

Windows Excel과 PowerShell은 빌드에 필요하지 않습니다. VBA 편집기에서 모듈을 직접 가져오는 절차도 필요 없습니다.

## 사용 방법

1. 신청 Excel 파일들을 `settings` 시트의 `request_folder` 폴더에 모읍니다.
2. 매크로 `MergeFirewallRequestFolder`를 실행합니다.
3. `requests` 시트에 통합 결과가 생성됩니다.
4. `validation_status`, `validation_message`, `match_details`, `firewall_path`, `zone_path` 컬럼을 확인합니다.
5. 방화벽/zone 데이터를 갱신한 뒤 다시 분석해야 하면 `AnalyzeRequestRoutes`만 실행합니다.

`request_folder`가 비어 있거나 잘못된 경로이면 폴더 선택창이 뜹니다. 선택한 경로는 `settings` 시트에 저장되어 다음 실행부터 재사용됩니다.

신청서 샘플 파일이 필요하면 매크로 `CreateSampleRequestWorkbook`를 실행합니다.

## 자동 입력 컬럼

`MergeFirewallRequestFolder`와 `AnalyzeRequestRoutes`는 `requests` 시트에 아래 컬럼을 자동으로 채웁니다.

| 컬럼 | 설명 |
|---|---|
| `source_file` | 원본 신청서 파일명 |
| `source_row` | 원본 신청서 행 번호 |
| `source_zone` | 출발 IP/CIDR의 longest-prefix 매칭으로 결정된 출발 zone |
| `destination_zone` | 목적지 IP/CIDR의 longest-prefix 매칭으로 결정된 도착 zone |
| `zone_path` | zone 그래프 BFS 최단경로의 zone 열. zone은 `>`로 연결 (예: `INTERNAL>DMZ>EXTERNAL`) |
| `firewall_path` | 통과한 방화벽을 등장 순서대로 보존한 경로 문자열. hop은 `>`로 연결 (예: `FW-INTERNAL-01>FW-DMZ-01`) |
| `target_firewalls` | 통과 방화벽 경로. hop 순서를 세미콜론(`;`)으로 연결. 같은 방화벽이 두 번 나와도 한 번만 기록 |
| `validation_status` | 검증 상태 (OK/MULTI_PATH/INTRA_ZONE/ZONE_UNRESOLVED/NO_PATH/DIRECTION_MISMATCH/LEGACY_FALLBACK) |
| `validation_message` | 검증 사유 또는 참고 메시지 |
| `match_details` | zone 결정과 hop 선택 근거를 사람이 읽을 수 있게 풀어 쓴 문자열 |

## 라우팅 분석 방식

이 통합 문서는 CIDR 겹침만 보던 단순 매칭을 라우팅 경로 분석으로 바꿉니다.

1. `network_definitions`에서 출발/목적지 IP/CIDR과 가장 길게 겹치는 대역을 찾습니다. `network_cidr`이 더 좁은 쪽이 이깁니다.
2. 각 매칭 대역의 `zone`을 출발/도착 zone으로 잡습니다.
3. `routing_paths`의 `(src_zone, dst_zone, path_order, firewall_name, ingress_if, egress_if, enabled)` 행을 zone 그래프의 간선으로 봅니다. `enabled=FALSE`인 행은 그래프에서 제외합니다.
4. `source_zone`에서 `destination_zone`까지의 결정적 BFS 최단경로를 찾습니다. 같은 zone-쌍에 hop이 여럿이면 `path_order`가 작은 행부터 시도합니다.
5. 경로 위의 hop 순서대로 `firewall_name`을 모읍니다. `firewall_path`와 `zone_path`는 hop 순서를 그대로 보존합니다. `target_firewalls`는 hop에서 처음 등장한 방화벽만 남깁니다.
6. 신청서 `방향`이 빈값/`BOTH`면 정방향 먼저, 없으면 역방향을 시도합니다. `OUT`은 정방향만, `IN`은 역방향만 봅니다.
7. 경로를 못 찾으면 `NO_PATH`. 출발/도착 zone이 같으면 `INTRA_ZONE`. zone 자체를 못 정하면 `ZONE_UNRESOLVED`. 방향 불일치면 `DIRECTION_MISMATCH`로 표시합니다.
8. `settings.route_legacy_fallback=TRUE`이고 라우팅 데이터가 없거나 경로를 못 찾으면, 그때만 예외적으로 기존 CIDR 겹침 방식으로 폴백해 `LEGACY_FALLBACK`을 남깁니다.

## network_definitions 시트

| 컬럼 | 설명 | 예시 |
|---|---|---|
| `network_name` | 대역 이름 | `internal-hq` |
| `network_cidr` | CIDR 또는 단일 IP | `10.10.0.0/16` |
| `zone` | 소속 zone | `INTERNAL` |
| `site` | 사이트/위치 (선택) | `본사` |
| `enabled` | 사용 여부 (TRUE/FALSE) | `TRUE` |

## firewalls 시트

`cidr_list`는 더 이상 메인 매칭 기준이 아닙니다. 매칭은 `routing_paths`의 hop이 결정합니다. 이 시트는 장비 메타데이터만 보관합니다.

| 컬럼 | 설명 | 예시 |
|---|---|
| `firewall_name` | 방화벽 식별자 | `FW-INTERNAL-01` |
| `vendor` | 제조사 (참고용) | `Fortinet` |
| `enabled` | 사용 여부 | `TRUE` |
| `comment` | 비고 | `본사 코어` |

## routing_paths 시트

| 컬럼 | 설명 | 예시 |
|---|---|---|
| `firewall_name` | 이 hop을 처리하는 방화벽 | `FW-INTERNAL-01` |
| `src_zone` | 출발 zone | `INTERNAL` |
| `dst_zone` | 도착 zone | `DMZ` |
| `ingress_if` | 인바운드 인터페이스 | `port1` |
| `egress_if` | 아웃바운드 인터페이스 | `port2` |
| `path_order` | 동일 zone-쌍 우선순위 (작을수록 우선) | `10` |
| `enabled` | 사용 여부 | `TRUE` |

## settings 시트

| key | 설명 | 기본값 |
|---|---|---|
| `request_folder` | 신청서 폴더 경로 | 비어 있음 (실행 시 폴더 선택창) |
| `parse_sheet` | 파싱할 시트 이름(정확히 일치). 비우면 헤더로 자동 감지 | 비어 있음 |
| `parse_targets` | (사용 안 함/예약) 현재 동작에 영향 없음 | `출발지IP;목적지IP` |
| `route_legacy_fallback` | 라우팅 실패 시 CIDR 폴백 여부 (TRUE/FALSE) | `FALSE` |

## validation_status 값

| 값 | 의미 |
|---|---|
| `OK` | 출발/도착 zone 결정, 라우팅 경로 결정, 모든 hop 정상 |
| `MULTI_PATH` | 같은 hop 수의 경로가 둘 이상. 우선 경로가 적용되며 `validation_message`에 후보 수 기록 |
| `INTRA_ZONE` | 출발/도착 zone이 동일. 방화벽이 개입하지 않으므로 `target_firewalls`는 비어 있음 |
| `ZONE_UNRESOLVED` | 출발 또는 도착 IP/CIDR이 `network_definitions` 어느 대역과도 겹치지 않음 |
| `NO_PATH` | zone 그래프에 출발에서 도착으로 가는 경로가 없음 |
| `DIRECTION_MISMATCH` | 신청서 방향과 그래프 방향이 맞지 않음 |
| `LEGACY_FALLBACK` | 라우팅 데이터 부족/경로 실패로 CIDR 겹침 방식으로 폴백함. `route_legacy_fallback=TRUE`일 때만 발생 |

## 파싱 대상 시트 선택

`settings` 시트의 `parse_sheet`로 파싱할 시트를 지정합니다.

| key | value |
|---|---|
| request_folder | `C:\신청서\방화벽정책` |
| parse_sheet | (비움=자동 감지) 또는 `신청내역` |
| route_legacy_fallback | `FALSE` |

`parse_sheet`가 비어 있으면 매크로가 모든 시트를 훑어 `출발지IP`/`목적지IP` 헤더가 있는 시트를 자동으로 고릅니다(동점이면 가장 왼쪽). 시트 이름을 적으면 그 시트만 파싱하고, 그 이름의 시트가 없거나 헤더가 없으면 명확히 오류로 처리합니다(다른 시트로 조용히 넘어가지 않음).

> `parse_targets` 키는 현재 동작에 영향이 없는 예약 항목입니다. 신청서는 항상 `출발지IP`와 `목적지IP`를 모두 요구합니다.

## 신청서 파일 구조

각 신청서 파일에는 `No` 또는 `번호`가 있는 헤더 행이 권장됩니다(필수는 아님). 매크로는 모든 시트의 상단 30행을 훑어 `출발지IP`/`목적지IP` 등 인식 가능한 컬럼이 가장 많은 행을 헤더로 보고, 그런 헤더를 가진 시트를 자동으로 선택합니다. 데이터가 첫 번째 시트가 아니어도 됩니다(표지/결재 시트가 앞에 있는 양식 지원, 동점이면 가장 왼쪽 시트).

```text
출발지IP | 출발지 | 목적지IP | 목적지 | 프로토콜 | 포트 | 방향 | 용도 | 시작일 | 종료일 | 비고
```

열 순서는 달라도 됩니다. B열이 `No`이고 C/D열부터 출발지 정보가 시작되는 양식도 지원합니다. 매크로는 헤더명을 기준으로 값을 읽습니다.

## 검증과 QA

이 저장소의 LibreOffice는 매크로 실행 검증이 불가능합니다. 다음 두 단계로 대신 검증합니다.

- 알고리즘 회귀 테스트: `tests/test_route_oracle.py`가 `tests/route_oracle.py`의 Python 오라클과 매크로 로직이 같은 결과를 내는지 확인합니다. zone 결정과 BFS 라우팅을 Python으로 다시 구현해 비교합니다. 변경 후 `pytest`를 한 번 돌려 통과를 확인합니다.
- 산출물 구조 검증: 빌드된 `dist/firewall-policy-automation.xlsm`은 zip의 `xl/vbaProject.bin` 존재 여부, 모듈명(`FirewallPolicyAutomation`), 시트 목록, 헤더 행으로 검증합니다. Excel을 띄우지 않아도 매크로 내장과 시트 정합성을 확인할 수 있습니다.

## 한계

- IPv4 CIDR 기준입니다. IPv6 대역은 같은 longest-prefix 매칭으로 확장 가능하지만 현재 빌드 산출물은 IPv4만 검증합니다.
- 폴더 안의 `.xls`, `.xlsx`, `.xlsm` 파일을 대상으로 합니다.
- Excel이 DRM 파일 열기를 허용해야 합니다. DRM이 파일 열기 자체를 막으면 VBA도 읽을 수 없습니다.
- CSV 파일은 대상이 아닙니다. 신청서는 Excel 파일로 관리합니다.
- LibreOffice로 매크로 실행을 검증할 수 없습니다. 알고리즘과 산출물 구조 검증은 위 QA 절차를 따릅니다.
