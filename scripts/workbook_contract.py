from __future__ import annotations

from typing import Final

REQUESTS_HEADERS: Final = [
    "요청부서", "요청번호", "제목", "원본파일", "원본행", "검증상태",
    "대상방화벽",
    "출발지IP", "출발지설명", "목적지IP", "목적지설명", "프로토콜", "포트", "방향",
    "용도", "시작일", "종료일", "비고",
    "검증메시지", "방화벽경로", "출발매칭대역", "목적매칭대역", "대역경로",
    "매칭근거", "요청폴더",
]

FIREWALLS: Final = [
    ["firewall_name", "vendor", "enabled", "comment"],
    ["SECUI-FW-01", "SECUI", "Y", "내부-서버 구간"],
    ["SECUI-FW-02", "SECUI", "Y", "서버-DMZ 구간"],
    ["SECUI-FW-03", "SECUI", "Y", "DMZ-외부 구간"],
]

FIREWALL_RANGES: Final = [
    ["firewall_name", "source_cidr", "destination_cidr", "direction", "path_order", "enabled", "comment", "source_interface", "destination_interface", "source_zone", "destination_zone"],
    ["SECUI-FW-01", "10.10.0.0/16", "172.16.0.0/16", "OUT", 10, "Y", "업무PC -> 서버", "inside", "server", "INTERNAL", "SERVER"],
    ["SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10, "Y", "업무PC -> DMZ", "inside", "dmz", "INTERNAL", "DMZ"],
    ["SECUI-FW-02", "10.10.0.0/16", "10.20.0.0/16", "OUT", 20, "Y", "업무PC -> DMZ", "server", "dmz", "SERVER", "DMZ"],
    ["SECUI-FW-01", "10.10.0.0/16", "8.8.8.0/24", "OUT", 10, "Y", "업무PC -> 외부 DNS", "inside", "outside", "INTERNAL", "EXTERNAL"],
    ["SECUI-FW-02", "10.10.0.0/16", "8.8.8.0/24", "OUT", 20, "Y", "업무PC -> 외부 DNS", "server", "outside", "SERVER", "EXTERNAL"],
    ["SECUI-FW-03", "10.10.0.0/16", "8.8.8.0/24", "OUT", 30, "Y", "업무PC -> 외부 DNS", "dmz", "outside", "DMZ", "EXTERNAL"],
]

SETTINGS: Final = [
    ["key", "value", "설명"],
    ["request_folder", "", "신청서 엑셀이 모여 있는 폴더 경로. 하위 폴더(예: 정보보호센터_1234)까지 재귀 탐색합니다."],
    ["parse_sheet", "", "파싱할 시트 이름(정확히 일치). 비워두면 헤더로 자동 감지합니다."],
    ["header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"],
]

PROCESSING_LOG: Final = [["processed_at", "source_file", "status", "merged_rows", "message"]]

SECUI_BATCH_HEADERS: Final = [
    "No", "장비명", "정책명", "출발지주소", "출발지명", "목적지주소", "목적지명",
    "서비스", "프로토콜", "목적지포트", "동작", "로그", "사용여부", "시작일",
    "종료일", "설명", "신청부서", "신청번호", "원본파일", "원본행",
]

SECUI_CLI_HEADERS: Final = [
    "No", "장비명", "정책명", "명령어", "검토메모", "신청부서", "신청번호",
    "원본파일", "원본행",
]

SECUI_POLICY_EXPORT: Final = [
    ["policy_id", "policy_name", "firewall_name", "source", "destination", "service", "action", "enabled", "comment"],
    ["1001", "ALLOW_WEB_TO_DMZ", "SECUI-FW-01", "10.10.10.5", "10.20.20.5", "tcp/443", "allow", "Y", "기존 허용 정책"],
    ["1002", "DENY_DNS_TO_INTERNET", "SECUI-FW-03", "10.10.10.5", "8.8.8.8", "udp/53", "deny", "Y", "차단 검토 필요"],
    ["1003", "DISABLED_DMZ_TEST", "SECUI-FW-02", "10.10.10.5", "10.20.20.5", "tcp/8443", "allow", "N", "비활성 정책"],
]

POLICY_ANALYSIS_HEADERS: Final = [
    "판정", "요청번호", "대상방화벽", "출발지", "목적지", "서비스", "기존정책",
    "기존정책상태", "근거", "조치",
    "요청원본행", "정책원본행", "raw_source", "raw_destination", "raw_service",
    "normalized_source", "normalized_destination", "normalized_protocol", "normalized_port",
    "debug_note",
]

POLICY_ANALYSIS_ROWS: Final = [
    [
        "EXISTING_ALLOW", "1234", "SECUI-FW-01", "10.10.10.5", "10.20.20.5", "TCP/443",
        "ALLOW_WEB_TO_DMZ", "사용", "출발지/목적지/서비스/허용 정책 일치", "기존 정책 확인 후 신청 처리 생략 검토",
        3, 2, "10.10.10.5", "10.20.20.5", "tcp/443", "10.10.10.5", "10.20.20.5", "TCP", "443",
        "seed: exact allow match",
    ],
    [
        "EXISTING_DENY", "1234", "SECUI-FW-03", "10.10.10.5", "8.8.8.8", "UDP/53",
        "DENY_DNS_TO_INTERNET", "사용", "차단 정책이 같은 트래픽과 일치", "방화벽 담당자 검토 필요",
        4, 3, "10.10.10.5", "8.8.8.8", "udp/53", "10.10.10.5", "8.8.8.8", "UDP", "53",
        "seed: deny match",
    ],
    [
        "NO_EXISTING_POLICY", "5678", "", "10.10.50.5", "172.16.9.9", "TCP/1521",
        "", "없음", "일치하는 기존 SECUI export 정책 없음", "신규 정책 생성 검토",
        5, "", "", "", "", "10.10.50.5", "172.16.9.9", "TCP", "1521",
        "seed: no match",
    ],
]

POLICY_SUMMARY_HEADERS: Final = ["구분", "건수", "검토 기준", "다음 조치"]

POLICY_SUMMARY_ROWS: Final = [
    ["전체", '=COUNTA(policy_analysis!A2:A5000)', "분석 대상 전체", "상태별 건수를 먼저 확인"],
    ["기존 허용", '=COUNTIF(policy_analysis!A2:A5000,"EXISTING_ALLOW")', "기존 허용 정책이 신청을 커버", "중복 신청 생략 또는 근거 첨부"],
    ["기존 차단", '=COUNTIF(policy_analysis!A2:A5000,"EXISTING_DENY")', "차단 정책과 일치", "방화벽 담당자 검토"],
    ["검토 필요", '=COUNTIF(policy_analysis!A2:A5000,"PARTIAL_MATCH")+COUNTIF(policy_analysis!A2:A5000,"OBJECT_UNRESOLVED")', "부분 일치 또는 객체명 미해석", "SECUI 객체/서비스 확인"],
    ["비활성 일치", '=COUNTIF(policy_analysis!A2:A5000,"DISABLED_MATCH")', "비활성 정책과 일치", "활성화 가능 여부 검토"],
    ["기존 정책 없음", '=COUNTIF(policy_analysis!A2:A5000,"NO_EXISTING_POLICY")', "일치 정책 없음", "신규 정책 생성 검토"],
]

VENDOR_CLI_TEMPLATES: Final = [
    ["vendor", "template_name", "enabled", "command_template", "review_note"],
    [
        "SECUI",
        "default_allow_srule",
        "Y",
        "fw set srule name {policy_name_q} action allow srcif {source_interface_q} dstif {destination_interface_q} src {source_object_q} dst {destination_object_q} service {service_object_q} log enable enable yes description {description_q} # device={firewall_name}",
        "장비 CLI에서 'fw set srule help'로 옵션명 확인 후 적용",
    ],
]

SERVICE_CATALOG: Final = [
    ["service_name", "protocol", "port", "secui_service", "description"],
    ["HTTPS", "TCP", "443", "tcp/443", "웹 HTTPS"],
    ["HTTP", "TCP", "80", "tcp/80", "웹 HTTP"],
    ["SSH", "TCP", "22", "tcp/22", "SSH 관리"],
    ["DNS-UDP", "UDP", "53", "udp/53", "DNS 조회"],
    ["DNS-TCP", "TCP", "53", "tcp/53", "DNS zone transfer 등"],
    ["ICMP", "ICMP", "", "icmp/", "ICMP/Ping"],
    ["CUSTOM", "TCP", "직접입력", "tcp/<port>", "목록에 없는 서비스는 포트 칸에 직접 입력"],
]

HEADER_ALIASES: Final = [
    ["standard", "your_column", "설명"],
    ["출발지", "", "신청서의 출발지 이름/설명 컬럼명 (예: 출발지ip설명)"],
    ["목적지", "", "신청서의 목적지 이름/설명 컬럼명"],
    ["프로토콜", "", "예: tcp/udp"],
    ["방향", "", "예: 구분"],
    ["시작일", "", "예: 시작일자"],
    ["종료일", "", "예: 종료일자"],
]

SAMPLE_FORMAT: Final = [
    [None, "No", "대상방화벽", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트",
     "방향", "용도", "시작일", "종료일", "비고"],
    [None, 1, "SECUI-FW-01", "10.10.10.5", "업무PC", "172.16.1.10", "업무시스템", "TCP", "443",
     "IN", "HTTPS 업무 연동", "2026-01-01", "2026-12-31", "정기 신청"],
]

USAGE: Final = [
    ["Step", "Action"],
    ["1", "firewalls 시트에 방화벽 장비명, 벤더, 사용여부를 등록한다"],
    ["2", "requests 또는 신청서 원본의 대상방화벽에 SECUI 장비명을 입력한다. 여러 장비는 ; 로 구분한다"],
    ["3", "vendor_cli_templates 시트에서 SECUI CLI 템플릿과 검토 메모를 확인한다"],
    ["4", "IP/CIDR/포트 여러 값은 세미콜론·콤마·줄바꿈·공백으로 구분한다"],
    ["5", "settings 시트의 request_folder에 신청서 폴더 경로를 적거나 SelectRequestFolder 매크로로 선택한다"],
    ["6", "requests 시트에 직접 입력하거나 MergeFirewallRequestFolder 매크로로 폴더 안 신청서를 통합한다 (Alt+F8)"],
    ["7", "라우팅 자동 탐색은 추후 기능이다. 현재 CLI 생성은 대상방화벽 입력값을 기준으로 한다"],
    ["8", "ConvertRequestsToSecuiBatch 매크로로 requests 결과를 secui_batch 장비별 배치 양식으로 변환한다"],
    ["9", "필요하면 vendor_cli_templates의 command_template을 장비 CLI 형식에 맞게 수정한다"],
    ["10", "ConvertRequestsToSecuiCli 매크로로 같은 장비·목적지·서비스 기준 룰별 그룹객체와 정책 CLI를 만든다"],
    ["11", "ANY는 객체를 만들지 않고 정책에 직접 들어간다. 인터페이스는 숨김 firewall_ranges의 source_interface/destination_interface 기준이다"],
    ["12", "기존 정책 검토가 필요하면 secui_policy_export에 export 정책을 붙여넣고 AnalyzeSecuiPolicyExport를 실행한다"],
    ["13", "policy_summary 시트에서 기존 허용·차단·검토필요 건수를 먼저 확인한다"],
    ["⚠", "입력 시트(녹색·황색 탭)는 보호되어 있다. 헤더는 수정 불가, 데이터 입력 영역만 타이핑 가능"],
    ["ℹ", "requests·policy_analysis·policy_summary·processing_log(파랑·회색 탭)은 매크로가 자동으로 채운다. 직접 수정하지 않는다"],
    ["💡", "프로토콜·포트는 service_catalog 예시를 참고하고, 방향은 드롭다운에서 선택한다"],
]

EXAMPLE_REQUEST_ROWS: Final = [
    {
        "요청부서": "정보보호센터",
        "요청번호": "1234",
        "제목": "웹서비스 연동",
        "원본파일": "example.xlsx",
        "원본행": 2,
        "대상방화벽": "SECUI-FW-01",
        "출발지IP": "10.10.10.5",
        "출발지설명": "업무PC",
        "목적지IP": "10.20.20.5",
        "목적지설명": "DMZ서버",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
        "용도": "단일 IP 경로 예시",
        "시작일": "2026-01-01",
        "종료일": "2026-12-31",
        "비고": "build seed (single IP)",
    },
    {
        "요청부서": "정보보호센터",
        "요청번호": "1234",
        "제목": "웹서비스 연동",
        "원본파일": "example.xlsx",
        "원본행": 3,
        "대상방화벽": "SECUI-FW-01;SECUI-FW-02",
        "출발지IP": "10.10.10.0/24",
        "출발지설명": "업무PC대역",
        "목적지IP": "10.20.20.0/24",
        "목적지설명": "DMZ대역",
        "프로토콜": "TCP",
        "포트": "443",
        "방향": "OUT",
        "용도": "CIDR 대역 경로 예시",
        "시작일": "2026-01-01",
        "종료일": "2026-12-31",
        "비고": "build seed (CIDR)",
    },
]

WIDTHS: Final = {
    "requests": {
        "A": 16, "B": 12, "C": 20, "D": 28, "E": 8, "F": 18, "G": 32,
        "H": 18, "I": 16, "J": 18, "K": 16, "L": 10, "M": 14, "N": 10,
        "O": 28, "P": 12, "Q": 12, "R": 24, "S": 40, "T": 34, "U": 18,
        "V": 18, "W": 30, "X": 60, "Y": 24,
    },
    "firewalls": {"A": 16, "B": 10, "C": 9, "D": 28},
    "firewall_ranges": {"A": 16, "B": 18, "C": 18, "D": 10, "E": 12, "F": 9, "G": 36,
                         "H": 18, "I": 20, "J": 16, "K": 18},
    "settings": {"A": 22, "B": 26, "C": 60},
    "header_aliases": {"A": 16, "B": 22, "C": 44},
    "processing_log": {"A": 20, "B": 22, "C": 10, "D": 12, "E": 40},
    "secui_batch": {
        "A": 6, "B": 18, "C": 36, "D": 18, "E": 18, "F": 18, "G": 18,
        "H": 16, "I": 10, "J": 12, "K": 10, "L": 8, "M": 10, "N": 12,
        "O": 12, "P": 42, "Q": 16, "R": 14, "S": 24, "T": 10,
    },
    "secui_cli": {"A": 6, "B": 18, "C": 36, "D": 120, "E": 60,
                  "F": 16, "G": 14, "H": 24, "I": 10},
    "secui_policy_export": {"A": 12, "B": 28, "C": 18, "D": 22, "E": 22,
                             "F": 16, "G": 12, "H": 10, "I": 36},
    "policy_analysis": {"A": 18, "B": 12, "C": 24, "D": 18, "E": 18,
                         "F": 14, "G": 28, "H": 14, "I": 42, "J": 28,
                         "K": 10, "L": 10, "M": 22, "N": 22, "O": 16,
                         "P": 22, "Q": 22, "R": 12, "S": 12, "T": 36},
    "policy_summary": {"A": 16, "B": 10, "C": 34, "D": 34},
    "vendor_cli_templates": {"A": 10, "B": 24, "C": 9, "D": 150, "E": 60},
    "service_catalog": {"A": 18, "B": 10, "C": 12, "D": 18, "E": 44},
    "sample-request-format": {"A": 4, "B": 6, "C": 22, "D": 16, "E": 12, "F": 16,
                              "G": 12, "H": 10, "I": 8, "J": 8, "K": 18,
                              "L": 12, "M": 12, "N": 14},
    "usage": {"A": 8, "B": 76},
}

FILTER_SHEETS: Final = {
    "requests", "firewalls", "firewall_ranges", "processing_log", "service_catalog",
    "secui_policy_export", "policy_analysis", "policy_summary",
}

FREEZE_PANES: Final = {"requests": "H3", "policy_analysis": "A2", "policy_summary": "A2"}

TAB_COLORS: Final = {
    "requests": "FF4472C4",
    "firewalls": "FF70AD47",
    "firewall_ranges": "FFFFC000",
    "settings": "FFFFC000",
    "header_aliases": "FFFFC000",
    "processing_log": "FFA6A6A6",
    "secui_batch": "FF4472C4",
    "secui_cli": "FF4472C4",
    "secui_policy_export": "FFFFC000",
    "policy_analysis": "FF4472C4",
    "policy_summary": "FF4472C4",
    "vendor_cli_templates": "FFFFC000",
    "service_catalog": "FFED7D31",
    "sample-request-format": "FFED7D31",
    "usage": "FFED7D31",
}

OPERATOR_VISIBLE_SHEETS: Final = {
    "usage",
    "requests",
    "settings",
    "firewalls",
    "secui_policy_export",
    "policy_summary",
    "policy_analysis",
    "secui_batch",
    "secui_cli",
    "vendor_cli_templates",
}

SUPPORT_DATA_SHEETS: Final = {
    "header_aliases",
    "firewall_ranges",
    "processing_log",
    "service_catalog",
    "sample-request-format",
}

PROTECT_SHEETS: Final = {
    "firewalls",
    "firewall_ranges",
    "settings",
    "header_aliases",
    "secui_policy_export",
    "vendor_cli_templates",
    "service_catalog",
}
