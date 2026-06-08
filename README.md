# Firewall Policy Excel Automation

DRM 환경에서 PowerQuery 없이 Excel 네이티브 VBA만으로 방화벽 정책 신청서를 통합 관리하는 저장소입니다.

## 목표

- 신청자가 작성한 Excel 파일을 특정 폴더에 모읍니다.
- 매크로 사용 통합 문서(`.xlsm`)에서 버튼/매크로로 폴더를 선택합니다.
- 폴더 안의 신청 Excel 파일을 하나의 `requests` 시트로 통합합니다.
- `firewalls` 시트에 등록한 방화벽 대역을 재사용합니다.
- `settings` 시트에 등록한 파싱 대상 컬럼으로 적용 대상 방화벽을 자동 산정합니다.
- 원본 파일명과 원본 행 번호를 남겨 추적 가능하게 합니다.

## 저장소 구조

```text
.
├── vba/
│   └── FirewallPolicyAutomation.bas   # Excel VBA 매크로 모듈
├── docs/
│   ├── excel-native.md                # 설치/운영 절차
│   └── excel-schema.md                # 신청서 컬럼 정의
├── .gitignore
└── README.md
```

## 빠른 시작

1. Excel에서 새 통합 문서를 만들고 `.xlsm`으로 저장합니다.
2. `Alt + F11`로 VBA 편집기를 엽니다.
3. `File > Import File...`에서 `vba/FirewallPolicyAutomation.bas`를 가져옵니다.
4. Excel로 돌아와 매크로 `SetupFirewallAutomationWorkbook`를 실행합니다.
5. `firewalls` 시트에 방화벽명과 CIDR 대역을 등록합니다.
6. `settings` 시트의 `parse_targets`에 파싱 대상 컬럼을 등록합니다. 기본값은 `출발지IP;목적지IP`입니다.
7. `settings` 시트의 `request_folder`에 신청서 폴더 경로를 등록합니다.
8. 매크로 `MergeFirewallRequestFolder`를 실행합니다.

## 주요 매크로

| 매크로 | 역할 |
|---|---|
| `SetupFirewallAutomationWorkbook` | `requests`, `firewalls`, `settings` 시트 생성 및 초기화 |
| `CreateSampleRequestWorkbook` | 신청자가 사용할 샘플 신청서 Excel 생성 |
| `SelectRequestFolder` | 신청서 폴더를 선택해서 `settings` 시트에 등록 |
| `MergeFirewallRequestFolder` | 폴더 안 신청서를 통합하고 적용 대상 방화벽 자동 입력 |

## 재사용 등록 방식

### 방화벽 대역 등록

`firewalls` 시트에 등록합니다.

| firewall_name | cidr_list |
|---|---|
| FW-INTERNAL-01 | `10.10.0.0/16;172.16.1.0/24` |
| FW-DMZ-01 | `10.20.0.0/16;172.16.20.0/24` |

### 파싱 대상 등록

`settings` 시트에 등록합니다.

| key | value |
|---|---|
| request_folder | `C:\신청서\방화벽정책` |
| parse_targets | `출발지IP;목적지IP` |

`request_folder`가 비어 있거나 실제 폴더가 아니면 `MergeFirewallRequestFolder` 실행 시 폴더 선택창이 뜨고, 선택한 경로가 `settings` 시트에 저장됩니다.

예시:

- 출발지IP와 목적지IP를 모두 기준으로 산정: `출발지IP;목적지IP`
- 목적지IP만 기준으로 산정: `목적지IP`
- 출발지IP/목적지IP/방향을 파싱 대상으로 등록: `출발지IP;목적지IP;방향`

## 신청서 기본 컬럼

| 컬럼 | 설명 | 예시 |
|---|---|---|
| 출발지IP | 출발지 IP 또는 CIDR | 10.10.10.0/24 |
| 출발지 | 출발지 이름/설명 | 업무PC |
| 목적지IP | 목적지 IP 또는 CIDR | 172.16.1.10 |
| 목적지 | 목적지 이름/설명 | 업무시스템 |
| 프로토콜 | TCP/UDP/ICMP 등 | TCP |
| 포트 | 포트 또는 포트 범위 | 443 |
| 방향 | 정책 방향 | IN |
| 용도 | 신청 목적 | HTTPS 업무 연동 |
| 시작일 | 시작일 | 2026-01-01 |
| 종료일 | 종료일 | 2026-12-31 |
| 비고 | 비고 | 정기 신청 |

통합 결과에는 매크로가 아래 컬럼을 자동으로 추가/입력합니다.

| 컬럼 | 설명 |
|---|---|
| source_file | 원본 신청서 파일명 |
| source_row | 원본 신청서 행 번호 |
| target_firewalls | 적용 대상 방화벽 |

중복 후보는 `출발지IP + 목적지IP + 프로토콜 + 포트 + 방향 + 용도` 조합으로 판단해 노란색으로 표시합니다. 매칭되는 방화벽이 없으면 `target_firewalls`에 `UNMATCHED`가 입력되며 빨간색으로 표시됩니다.

## 신청서 폴더의 xlsx 구조

현재 매크로는 폴더 안 각 `.xlsx/.xlsm/.xls` 파일의 첫 번째 시트에서 `No` 또는 `번호`가 있는 행을 헤더 행으로 봅니다. 예를 들어 B열이 `No`이고 C/D열부터 출발지 정보가 시작되는 양식도 지원합니다.

헤더 행에는 아래 의미의 컬럼이 있어야 합니다.

```text
출발지IP | 출발지 | 목적지IP | 목적지 | 프로토콜 | 포트 | 방향 | 용도 | 시작일 | 종료일 | 비고
```

열 순서는 달라도 됩니다. 매크로가 `No` 기준 헤더 행의 컬럼명을 찾아서 읽습니다. 공백은 일부 무시합니다. 예: `출발지 IP`도 `출발지IP`로 인식합니다.

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

## DRM 관련 전제

Excel이 DRM 파일을 열 수 있어야 VBA도 읽을 수 있습니다. DRM이 파일 열기 자체를 차단하면 PowerQuery뿐 아니라 VBA도 내용을 읽을 수 없습니다.
