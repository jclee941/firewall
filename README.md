# Firewall Policy Automation

DRM 환경에서 PowerQuery 없이 Excel VBA만으로 방화벽 정책 신청서를 통합하고, 사용자가 정의한 방화벽 통과 대역 기준으로 대상방화벽을 계산하는 도구입니다.

기존 zone/routing_paths/BFS 모델은 제거했습니다. 이제 운영자는 Excel에서 두 가지만 관리합니다.

- `firewalls`: 방화벽 장비 정의
- `firewall_ranges`: 어떤 출발지 대역에서 어떤 목적지 대역으로 갈 때 어떤 방화벽을 타는지

## 빠른 시작

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python scripts/build_xlsm.py
```

산출물은 `dist/firewall-policy-automation.xlsm`입니다.

## 운영 흐름

1. `firewalls` 시트에 장비명, 벤더, 사용여부를 등록합니다.
2. `firewall_ranges` 시트에 통과 대역을 등록합니다.
3. `settings.request_folder`에 신청서 폴더를 입력하거나 `SelectRequestFolder` 매크로로 선택합니다.
4. `MergeFirewallRequestFolder`로 신청서를 통합합니다.
5. `AnalyzeRequestRoutes`로 대상방화벽을 다시 계산할 수 있습니다.
6. SECUI 장비는 `ConvertRequestsToSecuiBatch` 또는 `ConvertRequestsToSecuiCli`로 배치/CLI 초안을 만듭니다.

## 시트

| 시트 | 역할 |
| --- | --- |
| `requests` | 통합 신청서와 분석 결과 |
| `firewalls` | 방화벽 장비 메타데이터 |
| `firewall_ranges` | 방화벽별 출발지/목적지 통과 대역 정의 |
| `settings` | 신청서 폴더, 파싱 대상 시트, 헤더 별칭 |
| `header_aliases` | 표준이 아닌 신청서 헤더 매핑 |
| `processing_log` | 통합 처리 로그 |
| `secui_batch` | SECUI 배치 입력용 행 |
| `secui_cli` | SECUI CLI 명령 초안 |
| `vendor_cli_templates` | 벤더별 CLI 명령 템플릿 |
| `sample-request-format` | 신청서 양식 예시 |
| `usage` | 워크북 안 사용 순서 |

## firewalls

| 컬럼 | 설명 | 예시 |
| --- | --- | --- |
| `firewall_name` | 장비 식별자. `firewall_ranges.firewall_name`과 일치해야 함 | `SECUI-FW-01` |
| `vendor` | 벤더. SECUI 출력은 `SECUI` 장비만 사용 | `SECUI` |
| `enabled` | 사용 여부. `Y`, `YES`, `TRUE`, `1`이면 사용 | `Y` |
| `comment` | 설명 | `내부-서버 구간` |

## firewall_ranges

| 컬럼 | 설명 | 예시 |
| --- | --- | --- |
| `firewall_name` | 적용 방화벽 | `SECUI-FW-01` |
| `source_cidr` | 출발지 IP/CIDR/목록. `ANY` 가능 | `10.10.0.0/16` |
| `destination_cidr` | 목적지 IP/CIDR/목록. `ANY` 가능 | `172.16.0.0/16` |
| `direction` | `OUT`, `IN`, `BOTH`, 빈값은 `BOTH` | `OUT` |
| `path_order` | 여러 방화벽이 매칭될 때 표시 순서. 작을수록 먼저 | `10` |
| `enabled` | 사용 여부 | `Y` |
| `comment` | 매칭 설명 | `업무PC -> 서버` |

대역 값은 IP, CIDR, `ANY`를 사용할 수 있습니다. 여러 값은 세미콜론, 콤마, 줄바꿈, 공백으로 나눕니다.

## 매칭 규칙

1. 신청서의 `출발지IP`와 `목적지IP`를 행 단위로 읽습니다.
2. `firewall_ranges.enabled=Y`이고 `firewalls.enabled=Y`인 행만 사용합니다.
3. 신청서 출발지와 `source_cidr`, 신청서 목적지와 `destination_cidr`가 CIDR 겹침이면 매칭입니다.
4. 방향은 `direction`과 신청서 `방향`이 맞아야 합니다. 빈값이나 `BOTH`는 양방향으로 봅니다.
5. 매칭된 행을 `path_order`, 행 순서, 방화벽명 순으로 정렬합니다.
6. `대상방화벽`은 정렬된 방화벽명을 중복 제거해 `;`로 연결합니다.
7. `방화벽경로`는 매칭 행 순서를 그대로 `>`로 연결합니다.

## 상태값

| 상태 | 의미 |
| --- | --- |
| `OK` | 하나 이상의 `firewall_ranges` 행이 매칭됨 |
| `NO_MATCH` | 매칭되는 방화벽 대역 정의가 없음 |
| `DIRECTION_MISMATCH` | 반대 방향 정의는 있으나 신청 방향과 맞지 않거나, 방향 값이 잘못됨 |
| `DUPLICATE` | 같은 요청이 중복 후보로 표시됨. 라우트 상태 뒤에 병합될 수 있음 |

## requests 결과 컬럼

`requests` 시트는 25개 컬럼을 유지합니다. 데이터는 3행부터입니다.

`요청부서`, `요청번호`, `제목`, `원본파일`, `원본행`, `검증상태`, `대상방화벽`, `출발지IP`, `출발지설명`, `목적지IP`, `목적지설명`, `프로토콜`, `포트`, `방향`, `용도`, `시작일`, `종료일`, `비고`, `검증메시지`, `방화벽경로`, `출발매칭대역`, `목적매칭대역`, `대역경로`, `매칭근거`, `요청폴더`

## SECUI 출력

`ConvertRequestsToSecuiBatch`와 `ConvertRequestsToSecuiCli`는 `대상방화벽`을 `;`로 분리해 장비별 행을 만듭니다. `firewalls.vendor=SECUI`이고 사용 중인 장비만 출력합니다.

CLI 명령은 `vendor_cli_templates` 시트의 `vendor=SECUI`, `enabled=Y` 행에 있는 `command_template`을 신청서 데이터로 치환해 만듭니다. 기본 placeholder는 `{policy_name_q}`, `{source_ip_q}`, `{destination_ip_q}`, `{service_q}`, `{description_q}`, `{firewall_name}`이며, `_q`가 붙은 값은 CLI용 따옴표로 감싼 값입니다. 실제 장비 적용 전 대상 장비에서 `fw set srule help`로 옵션명을 확인하세요.

## 개발

- 알고리즘 계약: `tests/route_oracle.py`
- VBA 구현: `vba/FirewallRouteAnalysis.bas`
- 워크북 빌드: `scripts/build_xlsm.py`
- 구조 테스트: `tests/test_xlsm_structure.py`

라우트 로직을 바꿀 때는 Python 오라클, VBA 모듈, 테스트를 함께 수정해야 합니다.
