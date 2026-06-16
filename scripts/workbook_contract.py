from __future__ import annotations

from typing import Final

REQUESTS_HEADERS: Final = [
    "요청부서", "요청번호", "대상방화벽", "출발지", "출발지설명", "목적지", "목적지설명",
    "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고",
]

REQUEST_TRACKING_SHEET: Final = "_request_tracking"
REQUEST_TRACKING_HEADERS: Final = ["request_row", "원본파일", "원본행", "요청폴더", "제목"]

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

SECUI_CLI_HEADERS: Final = [
    "No", "장비명", "정책명", "명령어", "검토메모", "신청부서", "신청번호",
    "원본파일", "원본행",
]

ROUTE_RESULTS_HEADERS: Final = [
    "요청부서", "요청번호", "출발지", "출발지설명", "목적지", "목적지설명",
    "프로토콜", "포트", "방향", "대상방화벽", "검증상태", "검증메시지",
    "방화벽경로", "출발매칭대역", "목적매칭대역", "대역경로", "매칭근거",
    "원본파일", "원본행",
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
    ["2", "신청서 원본의 대상방화벽은 선택 입력이다. 비어 있으면 firewall_ranges 기준으로 자동 산출한다"],
    ["3", "vendor_cli_templates 시트에서 SECUI CLI 템플릿과 검토 메모를 확인한다"],
    ["4", "IP/CIDR/포트 여러 값은 세미콜론·콤마·줄바꿈·공백으로 구분한다"],
    ["5", "settings 시트의 request_folder에 신청서 폴더 경로를 적거나 SelectRequestFolder 매크로로 선택한다"],
    ["6", "통합 문서를 열면 request_folder가 있을 때 신청서 통합, 경로탐색, secui_cli 생성을 자동 실행한다"],
    ["7", "대상방화벽을 직접 입력하면 그 값을 우선 사용하고, 공란이면 출발지/목적지 대역으로 경로를 찾는다"],
    ["8", "경로탐색 결과는 route_results 시트에서 확인한다"],
    ["9", "필요하면 vendor_cli_templates의 command_template을 장비 CLI 형식에 맞게 수정한다"],
    ["10", "ConvertRequestsToSecuiCli 매크로로 합쳐도 권한이 넓어지지 않는 신청을 룰별 그룹객체와 정책 CLI로 묶는다"],
    ["11", "ANY는 객체를 만들지 않고 정책에 직접 들어간다. 인터페이스는 firewall_ranges의 source_interface/destination_interface 기준이다"],
    ["⚠", "입력 시트(녹색·황색 탭)는 보호되어 있다. 헤더는 수정 불가, 데이터 입력 영역만 타이핑 가능"],
    ["ℹ", "requests·secui_cli·processing_log(파랑·회색 탭)은 매크로가 자동으로 채운다. 직접 수정하지 않는다"],
    ["💡", "프로토콜·포트는 service_catalog 예시를 참고하고, 방향은 드롭다운에서 선택한다"],
]

WEEKLY_REPORT: Final = [
    ["항목", "내용"],
    ["업무개선", "방화벽 정책 요청서 취합 및 대상방화벽 산출 자동화 도구 개발 진행"],
    ["진행내용", "수작업으로 취합하던 방화벽 정책 요청서를 Excel 기반으로 자동 병합하는 기능 개발"],
    ["진행내용", "출발지/목적지 대역 기준으로 대상방화벽을 자동 산출하는 로직 구현"],
    ["진행내용", "SECUI 정책 생성 및 검증 기능을 추가하여 정책 작성 반복 작업을 줄이는 방향으로 개선 중"],
    ["진행내용", "테스트 중 확인된 매크로 런타임 오류를 수정하며 Windows Excel 환경 기준으로 안정화 진행"],
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
        "A": 16, "B": 12, "C": 32, "D": 18, "E": 16, "F": 18, "G": 16,
        "H": 10, "I": 14, "J": 10, "K": 28, "L": 12, "M": 12, "N": 24,
    },
    "firewalls": {"A": 16, "B": 10, "C": 9, "D": 28},
    "firewall_ranges": {"A": 16, "B": 18, "C": 18, "D": 10, "E": 12, "F": 9, "G": 36,
                         "H": 18, "I": 20, "J": 16, "K": 18},
    "settings": {"A": 22, "B": 26, "C": 60},
    "header_aliases": {"A": 16, "B": 22, "C": 44},
    "processing_log": {"A": 20, "B": 22, "C": 10, "D": 12, "E": 40},
    "route_results": {"A": 16, "B": 12, "C": 18, "D": 16, "E": 18, "F": 16,
                      "G": 10, "H": 14, "I": 10, "J": 32, "K": 16, "L": 40,
                      "M": 34, "N": 18, "O": 18, "P": 30, "Q": 60, "R": 24, "S": 10},
    "secui_cli": {"A": 6, "B": 18, "C": 36, "D": 120, "E": 60,
                  "F": 16, "G": 14, "H": 24, "I": 10},
    "vendor_cli_templates": {"A": 10, "B": 24, "C": 9, "D": 150, "E": 60},
    "service_catalog": {"A": 18, "B": 10, "C": 12, "D": 18, "E": 44},
    REQUEST_TRACKING_SHEET: {"A": 12, "B": 24, "C": 10, "D": 24, "E": 32},
    "sample-request-format": {"A": 4, "B": 6, "C": 22, "D": 16, "E": 12, "F": 16,
                              "G": 12, "H": 10, "I": 8, "J": 8, "K": 18,
                              "L": 12, "M": 12, "N": 14},
    "usage": {"A": 8, "B": 76},
    "주간보고": {"A": 14, "B": 96},
}

FILTER_SHEETS: Final = {
    "requests", "firewalls", "firewall_ranges", "route_results", "processing_log", "service_catalog",
}

FREEZE_PANES: Final = {"requests": "D3"}

TAB_COLORS: Final = {
    "requests": "FF4472C4",
    "firewalls": "FF70AD47",
    "firewall_ranges": "FFFFC000",
    "settings": "FFFFC000",
    "header_aliases": "FFFFC000",
    "processing_log": "FFA6A6A6",
    "route_results": "FF4472C4",
    "secui_cli": "FF4472C4",
    "vendor_cli_templates": "FFFFC000",
    "service_catalog": "FFED7D31",
    "sample-request-format": "FFED7D31",
    "usage": "FFED7D31",
    "주간보고": "FFED7D31",
}

OPERATOR_VISIBLE_SHEETS: Final = {
    "usage",
    "주간보고",
    "requests",
    "settings",
    "firewalls",
    "firewall_ranges",
    "route_results",
    "secui_cli",
    "vendor_cli_templates",
}

SUPPORT_DATA_SHEETS: Final = {
    "header_aliases",
    "processing_log",
    "service_catalog",
    "sample-request-format",
    REQUEST_TRACKING_SHEET,
}

PROTECT_SHEETS: Final = {
    "firewalls",
    "firewall_ranges",
    "settings",
    "header_aliases",
    "vendor_cli_templates",
    "service_catalog",
}
