# Excel Native Operation

`dist/firewall-policy-automation.xlsm`은 운영자가 직접 사용하는 매크로 통합 문서입니다. Windows/Excel에서 실행하고, Linux 빌드는 `scripts/build_xlsm.py`가 담당합니다.

## 첫 실행

1. 통합 문서를 열고 매크로를 사용하도록 허용합니다.
2. `firewalls`에 SECUI 장비명, 벤더, 사용여부를 먼저 채웁니다.
3. `settings.request_folder`에 신청서 폴더를 입력하거나 `SelectRequestFolder`로 선택합니다.
4. 신청서 원본 또는 `requests`의 `대상방화벽`에 SECUI 장비명을 입력할 수 있습니다. 비워 두면 `firewall_ranges` 기준으로 자동 산출합니다.
5. 통합 문서를 다시 열면 `Workbook_Open`이 신청서 통합, 경로탐색, `secui_cli` 출력을 자동 생성합니다.
6. 자동 실행을 기다리지 않으려면 Excel의 매크로 목록에서 `MergeFirewallRequestFolder`, `AnalyzeRequestRoutes`, `ConvertRequestsToSecuiCli`를 순서대로 수동 실행합니다.

## 매크로

| 매크로 | 역할 |
| --- | --- |
| `SetupFirewallAutomationWorkbook` | 운영 시트 생성/초기화 |
| `SelectRequestFolder` | 신청서 폴더 선택 |
| `MergeFirewallRequestFolder` | 신청서 폴더 통합 |
| `AnalyzeRequestRoutes` | 추후 라우팅 검증용 수동 재분석 |
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

## 최종 목록과 경로 결과

`requests`는 최종 신청 목록에 필요한 `요청부서`, `요청번호`, `출발지`, `출발지설명`, `목적지`, `목적지설명`, `프로토콜`, `포트`, `시작일`, `종료일`, `비고`만 표시합니다. 원본파일/원본행과 경로탐색 관련 컬럼은 내부 처리용으로 숨깁니다.

경로탐색 결과는 `route_results` 시트에서 확인합니다. CLI 생성은 사용자가 입력한 `대상방화벽`을 우선 사용하고, 비어 있으면 경로탐색이 채운 값을 기준으로 합니다.

| 컬럼 | 의미 |
| --- | --- |
| `검증상태` | `OK`, `NO_MATCH`, `DIRECTION_MISMATCH` 등 |
| `대상방화벽` | CLI 생성 대상 SECUI 장비 목록. `;` 구분 |
| `방화벽경로` | 매칭 행 순서. `>` 구분 |
| `출발매칭대역` | 첫 번째 매칭 정의의 출발지 대역 |
| `목적매칭대역` | 첫 번째 매칭 정의의 목적지 대역 |
| `대역경로` | `출발매칭대역>목적매칭대역` |
| `매칭근거` | 매칭된 정의 행 설명 |

## SECUI

SECUI 변환 매크로는 `대상방화벽`에 적힌 장비 중 `firewalls.vendor=SECUI`이고 사용 중인 장비만 출력합니다. `대상방화벽`이 `SECUI-FW-01;SECUI-FW-02`이면 장비별로 처리됩니다. 자동 실행 흐름에서는 신청서 통합 후 `AnalyzeRequestRoutes`가 먼저 실행되어 비어 있는 `대상방화벽`을 채운 다음 CLI를 생성합니다.

`secui_cli`는 합쳐도 권한이 넓어지지 않는 행을 한 룰로 묶습니다. 같은 장비·목적지·서비스인 행은 출발지를 그룹 멤버로 합치고, 같은 장비·출발지·목적지인 행은 서비스를 그룹 멤버로 합칩니다. 출발지와 서비스가 동시에 서로 다른 행은 과허용을 막기 위해 별도 룰로 분리합니다.

`DUPLICATE` 표시는 예전 라우팅 검토용 중복 후보 표시입니다. 기본 CLI 생성 흐름에서는 중복 검증을 실행하지 않고, 합칠 수 있는 행은 위 기준으로 묶어 `secui_cli`에 출력합니다.

`ANY`, `ALL`, `*`, `0.0.0.0/0`은 객체 생성 없이 정책에 `ANY`로 직접 들어갑니다. 이때 `firewall_ranges`의 `source_interface`, `destination_interface` 매칭값으로 인터페이스를 제한합니다. 대역 매칭이 없으면 인터페이스도 `ANY`로 표시되므로 반영 전 반드시 확인해야 합니다.

CLI 명령 형식은 `vendor_cli_templates` 시트의 `vendor=SECUI`, `enabled=Y` 행에 있는 `command_template`에서 수정합니다. 기본 템플릿은 `{policy_name_q}`, `{source_interface_q}`, `{destination_interface_q}`, `{source_object_q}`, `{destination_object_q}`, `{service_object_q}`, `{description_q}`, `{firewall_name}` placeholder를 사용합니다.

`service_catalog` 시트에서 SECUI 서비스 표기 예시(`tcp/443`, `udp/53`, `icmp/`)를 확인할 수 있습니다. 이 시트는 입력 편의용 참고표이며 `requests`의 포트 입력을 제한하지 않습니다.

## 개발 검증

```bash
./.venv/bin/python -m pytest tests/ -q
./.venv/bin/python scripts/build_xlsm.py
```

라우트 계산 로직 변경 시 `tests/route_oracle.py`, `vba/FirewallRouteAnalysis.bas`, 관련 테스트를 함께 수정합니다.
