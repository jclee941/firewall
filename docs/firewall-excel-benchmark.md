# Firewall Excel Structure Benchmark

이 문서는 방화벽 정책 자동화 엑셀 구조를 설계할 때 참고할 공개 도구와 권장 시트 구조를 정리한다.

## Benchmark Sources

SECUI-specific public management repositories are scarce. Current public GitHub search did not surface a maintained SECUI-focused policy-management or CLI-generation repository. The most useful benchmark data therefore comes from spreadsheet/CSV-driven firewall policy tools and open-source NSPM-style projects.

| 도구 | 성격 | 구조에서 참고할 점 | URL |
| --- | --- | --- | --- |
| Tufin Firewall Access Request | Google Sheet 기반 방화벽 접근 요청 자동화 | 신청, 상태, 룰, 설정/디버그 탭을 분리하는 workflow 구조 | https://github.com/Tufin/Firewall-Access-Request |
| pfSense Firewall Rules Manager | pfSense XML을 CSV/Google Sheet로 변환하고 다시 XML로 되돌림 | 방화벽 룰을 한 테이블에서 편집하고 장비 포맷으로 export하는 흐름 | https://github.com/imthenachoman/pfSense-Firewall-Rules-Manager |
| fgpoliciestocsv | FortiGate 정책을 CSV/Excel 친화 포맷으로 export | 정책, 주소, 주소그룹, 서비스, 서비스그룹을 분리 export하는 구조 | https://github.com/maaaaz/fgpoliciestocsv |
| Azure Firewall Policy Export/Import | Azure Firewall Policy를 CSV로 export/import | 안정적인 CSV 스키마와 export-edit-import loop | https://github.com/WillyMoselhy/AzureFirewallPolicyExportImport |
| Firewall Orchestrator | 오픈소스 NSPM/방화벽 정책관리 | 장비 inventory, 정책 workflow, 보고/컴플라이언스 데이터 분리 | https://github.com/CactuseSecurity/firewall-orchestrator |
| firewall_policy_analyzer | 범용 방화벽 정책 충돌 분석 | source/destination/service/action 기반 룰 모델과 conflict 검출 | https://github.com/martimy/firewall_policy_analyzer |
| netclaw fwrule-analyzer | 멀티벤더 방화벽 룰 분석 | zones, addresses, ports, protocols, actions, applications 차원의 overlap/shadow/conflict/duplicate 분석 | https://github.com/automateyournetwork/netclaw |
| parsoalto | Palo Alto 룰을 사람이 읽기 쉬운 CSV로 변환 | 벤더 config를 flat CSV로 풀어 검토하는 export 구조 | https://github.com/olafhartong/parsoalto |

## Common Columns

공개 도구와 상용 NSPM 제품군에서 반복적으로 등장하는 핵심 데이터 축은 다음이다.

| 축 | 권장 컬럼 | 설명 |
| --- | --- | --- |
| 장비 | `firewall_name`, `vendor`, `enabled`, `comment` | 정책이 적용되는 장비와 벤더 필터 |
| 통과 대역 | `source_cidr`, `destination_cidr`, `direction`, `path_order` | 신청 트래픽이 어떤 방화벽을 타는지 계산하는 기준 |
| 신청 원문 | `request_team`, `request_doc_no`, `title`, `source_file`, `source_row` | 감사 추적과 원본 추적 |
| 트래픽 | `source_ip`, `source_name`, `destination_ip`, `destination_name`, `protocol`, `port`, `direction` | 정책 생성의 최소 단위 |
| 정책 메타 | `purpose`, `start_date`, `end_date`, `note` | 정책명/설명/만료 검토에 사용 |
| 분석 결과 | `validation_status`, `target_firewalls`, `validation_message`, `match_details` | 자동 검증과 후속 처리 |
| 출력 | `policy_name`, `command_template`, `review_note` | 배치/CLI 변환 결과 |

## Recommended Workbook Shape

현재 워크북은 아래 구조를 기준으로 유지한다.

| 시트 | 역할 | 입력 주체 |
| --- | --- | --- |
| `requests` | 신청서 통합 결과와 분석 결과 | 매크로/사용자 |
| `firewalls` | 방화벽 장비 inventory | 사용자 |
| `firewall_ranges` | 출발지/목적지 대역별 대상 방화벽 정의 | 사용자 |
| `settings` | 폴더/파싱 설정 | 사용자 |
| `header_aliases` | 신청서 헤더 별칭 | 사용자 |
| `vendor_cli_templates` | 벤더별 CLI 명령 템플릿 | 사용자 |
| `service_catalog` | SECUI 서비스 표기 예시 | 사용자/배포 산출물 |
| `secui_batch` | 장비별 배치 양식 | 매크로 |
| `secui_cli` | 장비별 CLI 명령 초안 | 매크로 |
| `processing_log` | 처리 로그 | 매크로 |
| `sample-request-format` | 신청서 예시 | 배포 산출물 |
| `usage` | 사용 순서 | 배포 산출물 |

## Excel Data Rules

1. 장비 정의와 통과 대역 정의를 분리한다.
2. 신청서 원문 컬럼과 분석 결과 컬럼을 같은 `requests` 행에 둔다.
3. 대상 방화벽은 `target_firewalls` 내부 필드, 화면 헤더는 `대상방화벽`으로 유지한다.
4. 다중 대상 방화벽은 `;`로 구분한다.
5. 경로 표시는 `>`로 구분한다.
6. 벤더별 출력 형식은 코드가 아니라 템플릿 시트 데이터로 관리한다.
7. CLI 출력은 바로 적용용이 아니라 검토용 초안으로 표시한다.

## Gap Backlog

| 기능 | 벤치마킹 근거 | 현재 상태 | 권장 방향 |
| --- | --- | --- | --- |
| 중복 정책 탐지 | netclaw, firewall_policy_analyzer | 기본 중복 후보 표시 | source/destination/protocol/port/direction 기준 상세 중복 리포트 |
| shadow/overlap 분석 | netclaw, Firewall Orchestrator | 미구현 | 기존 정책 테이블을 추가한 뒤 신규 신청과 비교 |
| 만료/재인증 | Firewall Orchestrator | 시작일/종료일 보유 | 종료일 기준 만료 예정 리포트 |
| 벤더별 export | pfSense Rules Manager, parsoalto, fgpoliciestocsv | SECUI batch/CLI | `vendor_cli_templates`로 벤더별 명령 형식을 데이터화 |
| 감사 추적 | NSPM 공통 | 원본파일/원본행/요청번호 보유 | 변경 전후 diff와 생성자/검토자 컬럼 추가 |
