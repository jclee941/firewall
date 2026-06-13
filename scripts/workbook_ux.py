from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from scripts.workbook_contract import PROTECT_SHEETS, TAB_COLORS

UX_LAST_ROW = 5000
REQ_PROTOCOL_COL = 12
REQ_PORT_COL = 13
REQ_DIRECTION_COL = 14
REQ_TARGET_COL = 7
REQ_SRC_IP_COL = 8
REQ_DST_IP_COL = 10
REQ_HEADER_ROW = 2
REQ_DATA_START_ROW = 3
EMPTY_REQUIRED_FILL = PatternFill("solid", fgColor="FFC7CE")
TARGET_FIREWALL_FILL = PatternFill("solid", fgColor="D9EAD3")
LINK_FONT = Font(color="0563C1", underline="single")
USAGE_LINKS = (
    ("requests", "신청 통합/분석 결과"),
    ("firewalls", "방화벽 장비 목록"),
    ("firewall_ranges", "출발지/목적지 통과 대역"),
    ("settings", "신청서 폴더와 파싱 설정"),
    ("secui_batch", "SECUI 배치 입력용 출력"),
    ("secui_cli", "SECUI CLI 명령 초안"),
    ("vendor_cli_templates", "벤더별 CLI 템플릿"),
    ("service_catalog", "SECUI 서비스 표기 예시"),
)


def add_list_validation(ws, col_letter, values, *, allow_blank=True, start_row=2):
    formula = '"' + ",".join(values) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
    dv.error = "목록에서 값을 선택하세요."
    dv.errorTitle = "잘못된 입력"
    dv.prompt = "드롭다운에서 선택"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{start_row}:{col_letter}{UX_LAST_ROW}")
    return dv


def apply_ux(wb) -> None:
    for name, color in TAB_COLORS.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color

    req = wb["requests"]
    add_list_validation(req, get_column_letter(REQ_PROTOCOL_COL), ["TCP", "UDP", "ICMP"], start_row=REQ_DATA_START_ROW)
    add_list_validation(req, get_column_letter(REQ_DIRECTION_COL), ["IN", "OUT", "BOTH"], start_row=REQ_DATA_START_ROW)

    if "sample-request-format" in wb.sheetnames:
        sf = wb["sample-request-format"]
        add_list_validation(sf, "G", ["TCP", "UDP", "ICMP"])
        add_list_validation(sf, "I", ["IN", "OUT", "BOTH"])

    if "usage" in wb.sheetnames:
        usage = wb["usage"]
        usage.cell(1, 3, "바로가기")
        usage.cell(1, 4, "설명")
        usage.column_dimensions["C"].width = 24
        usage.column_dimensions["D"].width = 34
        for row, (sheet_name, description) in enumerate(USAGE_LINKS, start=2):
            cell = usage.cell(row, 3, sheet_name)
            cell.hyperlink = f"#'{sheet_name}'!A1"
            cell.style = "Hyperlink"
            cell.font = LINK_FONT
            usage.cell(row, 4, description)
        wb.active = wb.sheetnames.index("usage")

    fw = wb["firewalls"]
    add_list_validation(fw, "C", ["Y", "N"])
    ranges = wb["firewall_ranges"]
    add_list_validation(ranges, "D", ["OUT", "IN", "BOTH"])
    add_list_validation(ranges, "F", ["Y", "N"])

    src_letter = get_column_letter(REQ_SRC_IP_COL)
    dst_letter = get_column_letter(REQ_DST_IP_COL)
    for letter in (src_letter, dst_letter):
        rng = f"{letter}{REQ_DATA_START_ROW}:{letter}{UX_LAST_ROW}"
        rule = FormulaRule(
            formula=[f"ISBLANK({letter}{REQ_DATA_START_ROW})"],
            fill=EMPTY_REQUIRED_FILL,
            stopIfTrue=False,
        )
        req.conditional_formatting.add(rng, rule)

    target_letter = get_column_letter(REQ_TARGET_COL)
    target_rng = f"{target_letter}{REQ_DATA_START_ROW}:{target_letter}{UX_LAST_ROW}"
    target_rule = FormulaRule(
        formula=[f"LEN(TRIM({target_letter}{REQ_DATA_START_ROW}))>0"],
        fill=TARGET_FIREWALL_FILL,
        stopIfTrue=False,
    )
    req.conditional_formatting.add(target_rng, target_rule)

    req.cell(REQ_HEADER_ROW, REQ_SRC_IP_COL).comment = Comment(
        "출발지 IP 또는 CIDR\n예: 10.10.10.0/24 또는 10.10.10.5\n여러 개는 ; 로 구분",
        "firewall-automation",
    )
    req.cell(REQ_HEADER_ROW, REQ_DST_IP_COL).comment = Comment(
        "목적지 IP 또는 CIDR\n예: 172.16.1.10 또는 172.16.1.0/24\n여러 개는 ; 로 구분",
        "firewall-automation",
    )
    req.cell(REQ_HEADER_ROW, REQ_PROTOCOL_COL).comment = Comment(
        "SECUI 서비스 표기용 프로토콜\n예: TCP, UDP, ICMP\nservice_catalog 시트의 secui_service 예시 참고",
        "firewall-automation",
    )
    req.cell(REQ_HEADER_ROW, REQ_PORT_COL).comment = Comment(
        "SECUI 서비스 표기용 포트\n예: 443 -> tcp/443, 53 -> udp/53\n목록에 없는 포트도 직접 입력 가능",
        "firewall-automation",
    )

    unlocked = Protection(locked=False)
    for name in PROTECT_SHEETS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        last_col = ws.max_column
        for row in range(2, UX_LAST_ROW + 1):
            for col in range(1, last_col + 1):
                ws.cell(row=row, column=col).protection = unlocked
        ws.protection.enable()
        ws.protection.autoFilter = False
        ws.protection.sort = False
        ws.protection.formatCells = True
