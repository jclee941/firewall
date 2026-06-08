# Excel 신청서/운영 스키마

이 문서는 운영 워크북에 들어가는 시트들의 컬럼 정의와 검증 규칙, 라우팅 분석 규칙을 정리합니다.

## 통합 관리 방식

- 신청 파일은 한 폴더에 모읍니다. 예: `incoming-requests/`
- Excel VBA 매크로는 폴더 안의 `.xls`, `.xlsx`, `.xlsm` 파일을 읽습니다.
- 임시 파일(`~$`로 시작하는 Excel 잠금 파일)은 무시합니다.
- 각 행에는 `source_file`, `source_row`, `target_firewalls`, `firewall_path`, `validation_status`, `validation_message`, `match_details`, `source_zone`, `destination_zone`, `zone_path`, `request_team`, `request_doc_no`, `request_folder` 컬럼이 자동으로 채워집니다. (`requests` 시트는 총 24개 컬럼입니다.)
- `firewall_path`/`zone_path`는 hop을 `>`로 잇고(예: `FW-INTERNAL-01>FW-DMZ-01`, `INTERNAL>DMZ`), `target_firewalls`는 방화벽을 `;`로 잇습니다.
- 결과는 매크로 통합 문서의 `requests` 시트에 저장됩니다.
- 방화벽 매칭은 출발/도착 IP의 CIDR 겹침이 아니라 `network_definitions`의 zone과 `routing_paths`의 zone 그래프 BFS 최단경로로 결정합니다.

## 시트 목록

| 시트 | 역할 |
|---|---|
| `requests` | 신청서 통합 결과 |
| `network_definitions` | IP 대역과 zone 매핑 |
| `firewalls` | 방화벽 장비 목록 |
| `routing_paths` | zone 사이 라우팅 경로 |
| `settings` | 운영 설정 |
| `sample-request-format` | 신청서 양식 예시 |
| `processing_log` | 파일별 처리 결과/오류 로그 |
| `usage` | 사용 순서 |

## 신청서 xlsx 구조

신청서 폴더에 들어가는 각 `.xlsx/.xlsm/.xls` 파일은 첫 번째 시트에 `No` 또는 `번호` 컬럼이 있는 헤더 행을 가져야 합니다. 매크로는 상단 30행 안에서 `No`/`번호`를 찾아 그 행을 헤더로 사용합니다.

```text
출발지IP | 출발지 | 목적지IP | 목적지 | 프로토콜 | 포트 | 방향 | 용도 | 시작일 | 종료일 | 비고
```

열 순서는 달라도 됩니다. B열이 `No`이고 C/D열부터 출발지 정보가 시작되는 양식도 지원합니다. 매크로가 헤더명을 기준으로 읽습니다.

## requests 시트 권장 컬럼

| 컬럼명 | 필수 | 형식 | 설명 |
|---|---:|---|---|
| 출발지IP | 예 | IP 또는 CIDR | 출발지 주소. 여러 개는 세미콜론(`;`)으로 구분 |
| 출발지 | 예 | 문자열 | 출발지 이름 또는 설명 |
| 목적지IP | 예 | IP 또는 CIDR | 목적지 주소. 여러 개는 세미콜론(`;`)으로 구분 |
| 목적지 | 예 | 문자열 | 목적지 이름 또는 설명 |
| 프로토콜 | 예 | 문자열 | TCP, UDP, ICMP 등 |
| 포트 | 예 | 숫자 또는 범위 | 443, 80;443, 1000-2000 등 |
| 방향 | 예 | 문자열 | 정책 방향. 빈값/BOTH/OUT/IN (라우팅 방향 제약) |
| 용도 | 예 | 문자열 | 신청 목적 |
| 시작일 | 예 | 날짜 | 정책 시작일 |
| 종료일 | 예 | 날짜 | 정책 종료일 |
| 비고 | 아니오 | 문자열 | 참고 사항 |
| source_file | 자동 | 문자열 | 원본 신청서 파일명 |
| source_row | 자동 | 숫자 | 원본 신청서 행 번호 |
| source_zone | 자동 | 문자열 | 출발 IP/CIDR의 longest-prefix 매칭으로 결정된 출발 zone |
| destination_zone | 자동 | 문자열 | 목적지 IP/CIDR의 longest-prefix 매칭으로 결정된 도착 zone |
| zone_path | 자동 | 문자열 | zone 그래프 BFS 최단경로의 zone 열 |
| firewall_path | 자동 | 문자열 | 통과한 방화벽을 등장 순서대로 보존한 경로 |
| target_firewalls | 자동 | 문자열 | 통과 방화벽 경로. hop 순서, 세미콜론 구분, 중복 제거 |
| validation_status | 자동 | 문자열 | OK/MULTI_PATH/INTRA_ZONE/ZONE_UNRESOLVED/NO_PATH/DIRECTION_MISMATCH/LEGACY_FALLBACK |
| validation_message | 자동 | 문자열 | 검증 메시지/사유 |
| match_details | 자동 | 문자열 | zone 결정과 hop 선택 근거 |

## network_definitions 시트

| 컬럼 | 필수 | 형식 | 설명 |
|---|---:|---|---|
| `network_name` | 예 | 문자열 | 대역 이름 |
| `network_cidr` | 예 | IP 또는 CIDR | 매칭 단위 CIDR 또는 단일 IP |
| `zone` | 예 | 문자열 | 소속 zone |
| `site` | 아니오 | 문자열 | 사이트/위치 (참고용) |
| `enabled` | 예 | `Y`/`TRUE` 등 | 사용 여부. `Y`,`YES`,`TRUE`,`1`,빈값=사용; 그 외=제외 (시드는 `Y` 사용) |

같은 zone을 여러 행에 나눠 등록해도 됩니다. zone 결정은 IP 단위가 아니라 CIDR 단위 longest-prefix match입니다. 신청서 값이 `10.10.0.0/16`이고 `network_definitions`에 `10.10.0.0/16`과 `10.10.0.0/15`가 둘 다 있으면 더 좁은 `10.10.0.0/16`이 이깁니다.

## firewalls 시트

방화벽 매칭은 `routing_paths`의 zone 그래프 경로로만 결정합니다. firewalls 시트는 장비 식별과 라우팅 hop 표시에만 쓰이며, CIDR 컬럼(`cidr_list`)은 더 이상 존재하지 않습니다. 2번째 컬럼은 `vendor`입니다.

| 컬럼 | 필수 | 형식 | 설명 |
|---|---:|---|---|
| `firewall_name` | 예 | 문자열 | 방화벽 식별자. `routing_paths.firewall_name`과 같은 값 |
| `vendor` | 아니오 | 문자열 | 제조사 (참고용) |
| `enabled` | 예 | `Y`/`TRUE` 등 | 사용 여부. FALSE면 라우팅 hop에서도 제외 (`Y`,`YES`,`TRUE`,`1`,빈값=사용) |
| `comment` | 아니오 | 문자열 | 비고 |

## routing_paths 시트

| 컬럼 | 필수 | 형식 | 설명 |
|---|---:|---|---|
| `firewall_name` | 예 | 문자열 | 이 hop을 처리하는 방화벽 |
| `src_zone` | 예 | 문자열 | 출발 zone |
| `dst_zone` | 예 | 문자열 | 도착 zone |
| `ingress_if` | 아니오 | 문자열 | 인바운드 인터페이스 |
| `egress_if` | 아니오 | 문자열 | 아웃바운드 인터페이스 |
| `path_order` | 예 | 숫자 | 동일 zone-쌍 우선순위. 작을수록 우선 |
| `enabled` | 예 | `Y`/`TRUE` 등 | 사용 여부. FALSE면 그래프에서 제외 (`Y`,`YES`,`TRUE`,`1`,빈값=사용) |

한 zone-쌍에 hop이 여럿이면 `path_order`로 우선순위를 정합니다. BFS는 `path_order` 오름차순으로 첫 번째 경로를 우선 경로로 봅니다. 같은 hop 수의 다른 경로가 있으면 `MULTI_PATH`로 표시하고 우선 경로의 hop 순서를 결과에 남깁니다.

## settings 시트

| key | value 형식 | 설명 | 기본값 |
|---|---|---|---|
| `request_folder` | 경로 | 신청서 폴더 | 비어 있음 (실행 시 폴더 선택창) |
| `parse_targets` | 세미콜론 구분 컬럼명 | 파싱 대상 컬럼 | `출발지IP;목적지IP` |
| `route_legacy_fallback` | TRUE/FALSE | 라우팅 실패 시 CIDR 겹침 폴백 사용 여부. 운영 표준값은 항상 `FALSE` | `FALSE` |
| `header_alias` | `정규헤더=별칭1,별칭2; ...` | 비표준 헤더를 표준 컬럼으로 매핑 (내장 별칭 우선, 못 잡은 것만 보완) | 비어 있음 |

## 검증 규칙

- `출발지IP`, `출발지`, `목적지IP`, `목적지`, `프로토콜`, `포트`, `방향`, `용도`, `시작일`, `종료일`, `비고` 헤더가 있어야 합니다.
- `출발지IP`, `목적지IP`는 단일 IP 또는 CIDR이어야 합니다.
- 여러 주소는 `10.0.0.1;10.0.1.0/24`처럼 세미콜론으로 구분합니다. 쉼표, 전각 세미콜론/쉼표, 줄바꿈도 구분자로 처리합니다.
- `프로토콜`은 TCP, UDP, ICMP 등 사내 표준값을 권장합니다.
- `포트`는 단일 포트, 세미콜론 구분, 범위 표기를 허용합니다.
- `시작일`, `종료일`은 Excel 날짜 또는 `YYYY-MM-DD` 형식을 권장합니다.
- 자동 입력 컬럼은 사용자가 직접 입력하지 않아도 됩니다. 통합 시 `network_definitions`/`routing_paths`/`firewalls` 데이터와 검증 결과 기준으로 자동 입력됩니다.

## direction (방향) 의미

방향은 CIDR 계산 입력이 아니라 라우팅 방향 제약입니다.

| 값 | 동작 |
|---|---|
| 빈값 또는 `BOTH` | 정방향 경로를 먼저 시도하고, 없으면 출발/도착을 뒤집어 역방향 경로를 다시 찾음 |
| `OUT` | 출발 -> 도착 한 방향만 시도 |
| `IN` | 도착 -> 출발 한 방향만 시도 (외부에서 내부로 들어오는 트래픽 표현) |

## 적용 대상 방화벽 매칭 규칙 (라우팅 기반)

기존 버전은 출발/목적지 IP와 방화벽 `cidr_list`의 겹침으로 적용 대상을 골랐습니다. 새 버전은 다음 순서로 동작합니다.

1. `network_definitions`에서 출발/목적지 IP/CIDR과 가장 길게 겹치는 대역을 찾습니다. 그 대역의 `zone`이 출발/도착 zone이 됩니다. 겹치는 대역이 없으면 `ZONE_UNRESOLVED`입니다.
2. `routing_paths`의 `enabled=TRUE`인 행만 모아 zone 그래프(방향 가중치 없음)를 만듭니다. 노드는 zone, 간선은 `(src_zone, dst_zone, path_order)`로 정렬된 hop 목록입니다.
3. 출발 zone에서 도착 zone까지 결정적 BFS 최단경로를 찾습니다. 같은 zone-쌍에 hop이 여럿이면 `path_order` 오름차순으로 첫 번째 행을 우선 사용합니다.
4. `방향` 컬럼이 빈값/`BOTH`면 정방향부터, 없으면 출발/도착을 뒤집어 역방향을 시도합니다. `OUT`은 정방향만, `IN`은 역방향만 봅니다. 방향이 맞지 않으면 `DIRECTION_MISMATCH`입니다.
5. 경로 위의 hop을 순서대로 따라가며 `firewall_name`을 모읍니다. `zone_path`와 `firewall_path`는 hop 순서를 그대로 보존합니다. `target_firewalls`는 hop에서 처음 등장한 방화벽만 남깁니다. 같은 방화벽이 두 번 나와도 한 번만 기록합니다.
6. 같은 hop 수의 다른 경로가 있으면 `MULTI_PATH`이고 우선 경로의 결과만 적용합니다.
7. 출발/도착 zone이 같으면 `INTRA_ZONE`입니다. 방화벽이 개입하지 않으므로 `target_firewalls`는 비어 있습니다.
8. `settings.route_legacy_fallback=TRUE`이고 위 단계에서 경로를 못 찾았거나 라우팅 데이터가 비어 있으면 그때만 CIDR 겹침 방식으로 폴백해 `LEGACY_FALLBACK`으로 남깁니다. `FALSE`면 `NO_PATH` 또는 `ZONE_UNRESOLVED`로 남습니다.

## validation_status 값

| 값 | 의미 |
|---|---|
| `OK` | 출발/도착 zone 결정, 라우팅 경로 결정, 모든 hop 정상. 통과 방화벽이 한 줄로 채워짐 |
| `MULTI_PATH` | 같은 hop 수의 경로가 둘 이상. 우선 경로가 적용되며 `validation_message`에 후보 수를 기록 |
| `INTRA_ZONE` | 출발/도착 zone이 동일. 방화벽이 개입하지 않으므로 `target_firewalls`는 비어 있음 |
| `ZONE_UNRESOLVED` | 출발 또는 도착 IP/CIDR이 `network_definitions` 어느 대역과도 겹치지 않음 |
| `NO_PATH` | zone 그래프에 출발에서 도착으로 가는 경로가 없음. 라우팅 데이터 보강 필요 |
| `DIRECTION_MISMATCH` | 신청서 방향과 그래프 방향이 맞지 않음. `validation_message`에 시도한 방향을 남김 |
| `LEGACY_FALLBACK` | 라우팅 데이터가 비어 있거나 경로를 못 찾아서 기존 CIDR 겹침 방식으로 폴백함. `route_legacy_fallback=TRUE`일 때만 발생 |

## 운영 메모

- 실제 Excel 원본 파일은 개인정보나 내부망 정보가 포함될 수 있으므로 Git에 커밋하지 않습니다.
- 샘플 파일은 민감하지 않은 RFC1918 예시 주소만 사용합니다.
- DRM 때문에 PowerQuery를 사용할 수 없어도 Excel에서 파일을 열 수 있으면 VBA로 통합 가능합니다.
- LibreOffice는 이 저장소에서 매크로 실행 검증이 불가능합니다. 변경 후 `pytest tests/`로 알고리즘을 검증하고, 빌드된 `.xlsm`은 zip의 `xl/vbaProject.bin`/모듈명/시트/헤더로 구조를 검증합니다.
