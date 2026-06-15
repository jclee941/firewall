# Excel Schema

운영 워크북은 zone 또는 라우팅 그래프를 쓰지 않습니다. 방화벽 장비는 `firewalls`, 통과 대역은 `firewall_ranges`에 사용자가 직접 정의합니다.

## sheets

| 시트 | 역할 |
| --- | --- |
| `requests` | 신청서 통합 결과와 CLI 생성 입력 |
| `firewalls` | 방화벽 장비 정의 |
| `firewall_ranges` | 출발지/목적지 대역별 통과 방화벽 정의 |
| `settings` | 폴더, 파싱 대상 시트, 헤더 별칭 |
| `header_aliases` | 비표준 신청서 헤더 매핑 |
| `processing_log` | 처리 로그 |
| `secui_cli` | SECUI CLI 초안 |
| `vendor_cli_templates` | 벤더별 CLI 명령 템플릿 |
| `service_catalog` | SECUI 서비스 표기 예시 |
| `sample-request-format` | 신청서 예시 |
| `usage` | 사용 순서 |

## requests

`requests`는 25개 컬럼을 고정으로 사용합니다. 1행은 출발지/목적지 그룹 라벨, 2행은 실제 헤더, 데이터는 3행부터입니다.

| 컬럼 | 입력/자동 | 설명 |
| --- | --- | --- |
| 요청부서 | 자동/입력 | 신청 부서 |
| 요청번호 | 자동/입력 | 신청 번호 |
| 제목 | 자동/입력 | 신청 제목 |
| 원본파일 | 자동 | 원본 파일명 |
| 원본행 | 자동 | 원본 행 번호 |
| 검증상태 | 숨김/수동 | 라우팅 검토 수동 실행 시 `OK`, `NO_MATCH`, `DIRECTION_MISMATCH` |
| 대상방화벽 | 입력/자동 | CLI 생성 대상 SECUI 장비 목록. `;` 구분. 값이 있으면 셀 배경색으로 강조 |
| 출발지IP | 입력 | IP, CIDR, 목록 |
| 출발지설명 | 입력 | 출발지 이름 |
| 목적지IP | 입력 | IP, CIDR, 목록 |
| 목적지설명 | 입력 | 목적지 이름 |
| 프로토콜 | 입력 | TCP/UDP/ICMP 등 |
| 포트 | 입력 | 포트 또는 포트 목록 |
| 방향 | 입력 | OUT/IN/BOTH. 빈값은 BOTH |
| 용도 | 입력 | 신청 사유 |
| 시작일 | 입력 | 시작일 |
| 종료일 | 입력 | 종료일 |
| 비고 | 입력 | 비고 |
| 검증메시지 | 숨김/수동 | 라우팅 검토 상태 설명 |
| 방화벽경로 | 숨김/수동 | 매칭 행 순서의 방화벽. `>` 구분 |
| 출발매칭대역 | 숨김/수동 | 첫 번째 매칭 정의의 `source_cidr` |
| 목적매칭대역 | 숨김/수동 | 첫 번째 매칭 정의의 `destination_cidr` |
| 대역경로 | 숨김/수동 | `출발매칭대역>목적매칭대역` |
| 매칭근거 | 숨김/수동 | 매칭된 `firewall_ranges` 행 요약 |
| 요청폴더 | 자동 | 원본 폴더 |

## firewalls

| 컬럼 | 필수 | 설명 |
| --- | --- | --- |
| `firewall_name` | 예 | 방화벽 식별자 |
| `vendor` | 아니오 | 벤더. SECUI 출력 필터에 사용 |
| `enabled` | 예 | `Y`, `YES`, `TRUE`, `1`이면 사용 |
| `comment` | 아니오 | 설명 |

## firewall_ranges

| 컬럼 | 필수 | 설명 |
| --- | --- | --- |
| `firewall_name` | 예 | `firewalls.firewall_name`과 일치 |
| `source_cidr` | 예 | 출발지 IP/CIDR/목록. `ANY` 가능 |
| `destination_cidr` | 예 | 목적지 IP/CIDR/목록. `ANY` 가능 |
| `direction` | 아니오 | OUT/IN/BOTH. 빈값은 BOTH |
| `path_order` | 아니오 | 표시 순서. 기본 999999 |
| `enabled` | 예 | 사용 여부 |
| `comment` | 아니오 | 매칭 설명 |

## matching

1. 사용 중인 방화벽과 사용 중인 대역 정의만 대상으로 합니다.
2. 신청서 출발지와 `source_cidr`, 신청서 목적지와 `destination_cidr`가 CIDR 범위로 겹치면 매칭입니다.
3. `ANY`, `*`, `ALL`, 빈 정의값, `0.0.0.0/0`은 전체 대역으로 봅니다.
4. 방향은 신청서 `방향`과 `firewall_ranges.direction`이 같거나 둘 중 하나가 `BOTH`이면 통과합니다.
5. 매칭 결과는 `path_order`, 행 순서, 방화벽명 순으로 정렬합니다.
6. 매칭이 없고 반대 방향 정의만 있으면 `DIRECTION_MISMATCH`, 아무 정의도 없으면 `NO_MATCH`입니다.

## settings

| key | 설명 |
| --- | --- |
| `request_folder` | 신청서 폴더 |
| `parse_sheet` | 파싱할 시트명. 빈값이면 자동 감지 |
| `header_alias` | 비표준 헤더 별칭 |

## SECUI sheets

`secui_cli`는 `requests.대상방화벽`을 장비별로 분리해 생성합니다. `firewalls.vendor`가 `SECUI`이고 사용 중인 장비만 포함합니다. 라우팅 자동 탐색은 추후 기능이며, 현재 CLI 생성은 사용자가 입력한 `대상방화벽`을 기준으로 합니다. 통합 문서를 열면 `request_folder`가 존재할 때 신청서 통합 후 CLI 출력 시트를 자동 갱신합니다.

`secui_cli`는 합쳐도 권한이 넓어지지 않는 신청을 한 정책으로 묶습니다. 같은 장비·목적지 객체/주소·서비스인 신청은 출발지를 그룹 멤버로 합치고, 같은 장비·출발지 객체/주소·목적지 객체/주소인 신청은 서비스를 그룹 멤버로 합칩니다. 출발지와 서비스가 동시에 서로 다른 조합은 과허용을 막기 위해 별도 정책으로 분리합니다. `ANY`, `ALL`, `*`, `0.0.0.0/0`은 그룹 객체를 만들지 않고 정책 값에 `ANY`를 직접 넣습니다.

`firewall_ranges`의 `source_interface`, `destination_interface`, `source_zone`, `destination_zone`은 CLI 그룹명과 인터페이스 placeholder 산출에 쓰는 보조 기준입니다. 매칭되는 대역 정의가 없으면 `source_interface`/`destination_interface`는 `ANY`가 됩니다.

`vendor_cli_templates`는 CLI 명령어를 엑셀 데이터로 관리합니다. `vendor=SECUI`, `enabled=Y`인 첫 번째 행의 `command_template`을 사용하며, `{policy_name_q}`, `{source_interface_q}`, `{destination_interface_q}`, `{source_object_q}`, `{destination_object_q}`, `{service_object_q}`, `{description_q}`, `{firewall_name}` 같은 placeholder를 신청서 값과 룰별 그룹 객체명으로 치환합니다. `_q`가 붙은 placeholder는 CLI용 따옴표로 감싼 값입니다.

`service_catalog`는 SECUI CLI에서 쓰는 서비스 표기 예시를 제공합니다. `secui_service`는 `프로토콜/포트` 조합을 보여 주는 참고값이며, 라우트 계산에는 사용하지 않습니다. 목록에 없는 포트는 `requests.포트`에 직접 입력합니다.
