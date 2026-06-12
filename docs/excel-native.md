# Excel Native Operation

`dist/firewall-policy-automation.xlsm`은 운영자가 직접 사용하는 매크로 통합 문서입니다. Windows/Excel에서 실행하고, Linux 빌드는 `scripts/build_xlsm.py`가 담당합니다.

## 매크로

| 매크로 | 역할 |
| --- | --- |
| `SetupFirewallAutomationWorkbook` | 운영 시트 생성/초기화 |
| `SelectRequestFolder` | 신청서 폴더 선택 |
| `MergeFirewallRequestFolder` | 신청서 폴더 통합 후 방화벽 대역 분석 |
| `AnalyzeRequestRoutes` | 이미 통합된 `requests` 행 재분석 |
| `ConvertRequestsToSecuiBatch` | SECUI 배치 양식 생성 |
| `ConvertRequestsToSecuiCli` | SECUI CLI 초안 생성 |
| `CreateSampleRequestWorkbook` | 신청서 샘플 파일 생성 |

## 운영 입력

### firewalls

장비 메타데이터만 적습니다.

| 컬럼 | 예시 |
| --- | --- |
| `firewall_name` | `SECUI-FW-01` |
| `vendor` | `SECUI` |
| `enabled` | `Y` |
| `comment` | `내부-서버 구간` |

### firewall_ranges

실제 계산 기준입니다. 한 행은 “이 출발지 대역에서 이 목적지 대역으로 갈 때 이 방화벽을 탄다”를 의미합니다.

| 컬럼 | 예시 |
| --- | --- |
| `firewall_name` | `SECUI-FW-01` |
| `source_cidr` | `10.10.0.0/16` |
| `destination_cidr` | `172.16.0.0/16` |
| `direction` | `OUT` |
| `path_order` | `10` |
| `enabled` | `Y` |
| `comment` | `업무PC -> 서버` |

대역 목록은 `;`, 콤마, 줄바꿈, 공백으로 나눌 수 있습니다. 전체 대역은 `ANY`를 권장합니다.

## 분석 결과

`requests`의 자동 결과 컬럼은 다음을 봅니다.

| 컬럼 | 의미 |
| --- | --- |
| `검증상태` | `OK`, `NO_MATCH`, `DIRECTION_MISMATCH` 등 |
| `대상방화벽` | 중복 제거된 방화벽 목록. `;` 구분 |
| `방화벽경로` | 매칭 행 순서. `>` 구분 |
| `출발매칭대역` | 첫 번째 매칭 정의의 출발지 대역 |
| `목적매칭대역` | 첫 번째 매칭 정의의 목적지 대역 |
| `대역경로` | `출발매칭대역>목적매칭대역` |
| `매칭근거` | 매칭된 정의 행 설명 |

## 상태값

| 상태 | 의미 |
| --- | --- |
| `OK` | 방화벽 대역 정의가 하나 이상 매칭됨 |
| `NO_MATCH` | 맞는 대역 정의가 없음 |
| `DIRECTION_MISMATCH` | 반대 방향 정의만 있거나 방향 값이 잘못됨 |
| `DUPLICATE` | 중복 후보. 기존 상태 뒤에 병합될 수 있음 |

## SECUI

SECUI 변환 매크로는 `firewalls.vendor=SECUI`이고 사용 중인 장비만 출력합니다. `대상방화벽`이 `SECUI-FW-01;SECUI-FW-02`이면 장비별로 행이 나뉩니다.

`secui_cli`의 명령은 검토용 초안입니다. 장비 펌웨어별 옵션명이 다를 수 있으므로 실제 반영 전 CLI 도움말에서 확인해야 합니다.

CLI 명령 형식은 `vendor_cli_templates` 시트의 `vendor=SECUI`, `enabled=Y` 행에 있는 `command_template`에서 수정합니다. 기본 템플릿은 `{policy_name_q}`, `{source_ip_q}`, `{destination_ip_q}`, `{service_q}`, `{description_q}`, `{firewall_name}` placeholder를 사용합니다.

## 개발 검증

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python scripts/build_xlsm.py
```

라우트 계산 로직 변경 시 `tests/route_oracle.py`, `vba/FirewallRouteAnalysis.bas`, 관련 테스트를 함께 수정합니다.
