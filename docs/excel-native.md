# Excel 네이티브 자동화 방식

DRM 때문에 PowerQuery가 막혀 있으면 Excel 내부 VBA 매크로 방식이 가장 현실적입니다. 이 저장소는 Excel 네이티브 방식만 대상으로 하며, 별도 Python 실행 없이 매크로 사용 통합 문서(`.xlsm`) 안에서 폴더의 신청서를 열고 통합합니다.

## 구성

`vba/FirewallPolicyAutomation.bas`를 Excel VBA 편집기에 가져와서 사용합니다.

통합 문서에는 다음 시트가 생성됩니다.

| 시트 | 역할 |
|---|---|
| `requests` | 통합 결과 저장 |
| `firewalls` | 방화벽명과 담당 CIDR 대역 관리 |
| `settings` | 운영 설정 기록 |

## 매크로 목록

| 매크로 | 역할 |
|---|---|
| `SetupFirewallAutomationWorkbook` | 운영용 시트 생성 |
| `CreateSampleRequestWorkbook` | 신청자 배포용 샘플 Excel 생성 |
| `SelectRequestFolder` | 신청서 폴더 등록 |
| `MergeFirewallRequestFolder` | 신청서 폴더 통합 |

## 설치 방법

1. Excel에서 새 통합 문서를 만들고 `.xlsm`으로 저장합니다.
2. `Alt + F11`로 VBA 편집기를 엽니다.
3. `File > Import File...`에서 `vba/FirewallPolicyAutomation.bas`를 가져옵니다.
4. Excel로 돌아와 매크로 `SetupFirewallAutomationWorkbook`를 한 번 실행합니다.
5. `firewalls` 시트에 실제 방화벽명과 CIDR 대역을 입력합니다.
6. `settings` 시트의 `request_folder` 값에 신청서 폴더 경로를 등록합니다.
7. `settings` 시트의 `parse_targets` 값에 파싱 대상 컬럼을 등록합니다. 기본값은 `출발지IP;목적지IP`입니다.

## 사용 방법

1. 신청 Excel 파일들을 `settings` 시트의 `request_folder` 폴더에 모읍니다.
2. 매크로 `MergeFirewallRequestFolder`를 실행합니다.
3. `requests` 시트에 통합 결과가 생성됩니다.

`request_folder`가 비어 있거나 잘못된 경로이면 폴더 선택창이 뜹니다. 선택한 경로는 `settings` 시트에 저장되어 다음 실행부터 재사용됩니다.

신청서 샘플 파일이 필요하면 매크로 `CreateSampleRequestWorkbook`를 실행합니다.

중복 신청 번호는 노란색으로 표시됩니다. 적용 대상 방화벽이 매칭되지 않으면 `target_firewalls`에 `UNMATCHED`가 입력되고 빨간색으로 표시됩니다.

## 자동 입력 컬럼

| 컬럼 | 설명 |
|---|---|
| `source_file` | 원본 신청서 파일명 |
| `source_row` | 원본 신청서 행 번호 |
| `target_firewalls` | 출발지/목적지 대역 기준 적용 대상 방화벽 |

## 방화벽 매칭 방식

`firewalls` 시트의 `cidr_list`에 방화벽 담당 대역을 입력합니다.

예시:

| firewall_name | cidr_list |
|---|---|
| FW-INTERNAL-01 | `10.10.0.0/16;172.16.1.0/24` |
| FW-DMZ-01 | `10.20.0.0/16;172.16.20.0/24` |

신청서의 `출발지IP` 또는 `목적지IP`가 담당 대역과 겹치면 `target_firewalls`에 방화벽명이 들어갑니다. 여러 개가 매칭되면 세미콜론(`;`)으로 연결됩니다.

신청서 값이 단일 IP가 아니라 CIDR 대역이어도 대역 겹침으로 계산합니다. 여러 값은 세미콜론, 쉼표, 줄바꿈으로 구분할 수 있습니다.

## 파싱 대상 등록

`settings` 시트에서 재사용할 파싱 대상 컬럼을 등록합니다.

| key | value |
|---|---|
| request_folder | `C:\신청서\방화벽정책` |
| parse_targets | `출발지IP;목적지IP` |

예를 들어 목적지IP만 기준으로 방화벽을 산정하려면 `목적지IP`만 입력합니다. 출발지IP와 목적지IP를 함께 보려면 `출발지IP;목적지IP`처럼 세미콜론으로 구분합니다.

## 신청서 파일 구조

각 신청서 파일은 첫 번째 시트에 `No` 또는 `번호`가 있는 헤더 행이 있어야 합니다. 매크로는 상단 30행에서 `No`/`번호`를 찾아 그 행을 헤더로 사용합니다.

```text
출발지IP | 출발지 | 목적지IP | 목적지 | 프로토콜 | 포트 | 방향 | 용도 | 시작일 | 종료일 | 비고
```

열 순서는 달라도 됩니다. B열이 `No`이고 C/D열부터 출발지 정보가 시작되는 양식도 지원합니다. 매크로는 헤더명을 기준으로 값을 읽습니다.

## 한계

- VBA 버전은 IPv4 CIDR 기준입니다.
- 폴더 안의 `.xls`, `.xlsx`, `.xlsm` 파일을 대상으로 합니다.
- Excel이 DRM 파일 열기를 허용해야 합니다. DRM이 파일 열기 자체를 막으면 VBA도 읽을 수 없습니다.
- CSV 파일은 대상이 아닙니다. 신청서는 Excel 파일로 관리합니다.
