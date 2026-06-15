Option Explicit

Private Const SETTINGS_SHEET As String = "settings"
Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const REQUESTS_SHEET As String = "requests"
Private Const SECUI_BATCH_SHEET As String = "secui_batch"
Private Const SECUI_CLI_SHEET As String = "secui_cli"
Private Const SECUI_POLICY_EXPORT_SHEET As String = "secui_policy_export"
Private Const POLICY_ANALYSIS_SHEET As String = "policy_analysis"
Private Const POLICY_SUMMARY_SHEET As String = "policy_summary"
Private Const VENDOR_CLI_TEMPLATE_SHEET As String = "vendor_cli_templates"
Private Const SERVICE_CATALOG_SHEET As String = "service_catalog"
Private Const LOG_SHEET As String = "processing_log"
Private Const FIREWALL_RANGE_SHEET As String = "firewall_ranges"

' requests output layout: row 1 = cosmetic group labels, row 2 = leaf headers,
' data starts at row 3. Keep these in sync with build_xlsm.py constants.
Private Const REQ_HEADER_GROUP_ROW As Long = 1
Private Const REQ_HEADER_ROW As Long = 2
Private Const REQ_DATA_START_ROW As Long = 3
Private Const REQ_LAST_COL As Long = 25
Private Const COL_REQUEST_TEAM As Long = 1
Private Const COL_REQUEST_DOC_NO As Long = 2
Private Const COL_REQUEST_TITLE As Long = 3
Private Const COL_SOURCE_FILE As Long = 4
Private Const COL_SOURCE_ROW As Long = 5
Private Const COL_VALIDATION_STATUS As Long = 6
Private Const COL_TARGET_FIREWALLS As Long = 7
Private Const COL_SOURCE_IP As Long = 8
Private Const COL_SOURCE_NAME As Long = 9
Private Const COL_DESTINATION_IP As Long = 10
Private Const COL_DESTINATION_NAME As Long = 11
Private Const COL_PROTOCOL As Long = 12
Private Const COL_PORT As Long = 13
Private Const COL_DIRECTION As Long = 14
Private Const COL_PURPOSE As Long = 15
Private Const COL_START_DATE As Long = 16
Private Const COL_END_DATE As Long = 17
Private Const COL_NOTE As Long = 18
Private Const COL_VALIDATION_MESSAGE As Long = 19
Private Const COL_FIREWALL_PATH As Long = 20
Private Const COL_SOURCE_ZONE As Long = 21
Private Const COL_DESTINATION_ZONE As Long = 22
Private Const COL_ZONE_PATH As Long = 23
Private Const COL_MATCH_DETAILS As Long = 24
Private Const COL_REQUEST_FOLDER As Long = 25
Private Const SECUI_LAST_COL As Long = 20
Private Const SECUI_CLI_LAST_COL As Long = 9
Private Const SECUI_EXPORT_LAST_COL As Long = 9
Private Const POLICY_ANALYSIS_LAST_COL As Long = 20
Private Const POLICY_SUMMARY_LAST_COL As Long = 4
Private Const FW_COL_NAME As Long = 1
Private Const FW_COL_VENDOR As Long = 2
Private Const FW_COL_ENABLED As Long = 3
Private Const CLI_TEMPLATE_COL_VENDOR As Long = 1
Private Const CLI_TEMPLATE_COL_ENABLED As Long = 3
Private Const CLI_TEMPLATE_COL_COMMAND As Long = 4
Private Const CLI_TEMPLATE_COL_NOTE As Long = 5

Private mUserAliases As Object
Private mParseSheetName As String
Private mSuppressMessages As Boolean

Public Sub AutoRunWorkbookOutputs()
    Dim oldSuppress As Boolean
    Dim settingsSheet As Worksheet

    oldSuppress = mSuppressMessages
    mSuppressMessages = True
    On Error GoTo AutoRunFailed

    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    If FolderExists(SettingsValue(settingsSheet, "request_folder")) Then
        MergeFirewallRequestFolder
    End If
    ConvertRequestsToSecuiBatch
    ConvertRequestsToSecuiCli
    AnalyzeSecuiPolicyExport

    mSuppressMessages = oldSuppress
    Exit Sub

AutoRunFailed:
    mSuppressMessages = oldSuppress
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SetupFirewallAutomationWorkbook()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim firewallRangeSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim logSheet As Worksheet
    Dim secuiBatchSheet As Worksheet
    Dim secuiCliSheet As Worksheet
    Dim secuiPolicyExportSheet As Worksheet
    Dim policyAnalysisSheet As Worksheet
    Dim policySummarySheet As Worksheet
    Dim vendorCliTemplateSheet As Worksheet
    Dim serviceCatalogSheet As Worksheet

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set firewallRangeSheet = EnsureSheet(FIREWALL_RANGE_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    Set logSheet = EnsureSheet(LOG_SHEET)
    Set secuiBatchSheet = EnsureSheet(SECUI_BATCH_SHEET)
    Set secuiCliSheet = EnsureSheet(SECUI_CLI_SHEET)
    Set secuiPolicyExportSheet = EnsureSheet(SECUI_POLICY_EXPORT_SHEET)
    Set policyAnalysisSheet = EnsureSheet(POLICY_ANALYSIS_SHEET)
    Set policySummarySheet = EnsureSheet(POLICY_SUMMARY_SHEET)
    Set vendorCliTemplateSheet = EnsureSheet(VENDOR_CLI_TEMPLATE_SHEET)
    Set serviceCatalogSheet = EnsureSheet(SERVICE_CATALOG_SHEET)

    WriteRequestHeaders requestsSheet
    WriteFirewallHeaders firewallsSheet
    WriteFirewallRangeHeaders firewallRangeSheet
    WriteSettings settingsSheet
    WriteLogHeaders logSheet
    WriteSecuiBatchHeaders secuiBatchSheet
    WriteSecuiCliHeaders secuiCliSheet
    WriteSecuiPolicyExportHeaders secuiPolicyExportSheet
    WritePolicyAnalysisHeaders policyAnalysisSheet
    WritePolicySummarySheet policySummarySheet
    WriteVendorCliTemplateHeaders vendorCliTemplateSheet
    WriteServiceCatalogHeaders serviceCatalogSheet
    FormatRequestsSheet requestsSheet
    FormatFirewallsSheet firewallsSheet
    FormatGenericSheet firewallRangeSheet, "A:K"
    FormatLogSheet logSheet
    FormatSecuiBatchSheet secuiBatchSheet
    FormatSecuiCliSheet secuiCliSheet
    FormatGenericSheet secuiPolicyExportSheet, "A:I"
    FormatPolicyAnalysisSheet policyAnalysisSheet
    FormatPolicySummarySheet policySummarySheet
    FormatGenericSheet vendorCliTemplateSheet, "A:E"
    FormatGenericSheet serviceCatalogSheet, "A:E"
    ApplyOperatorSheetVisibility

    MsgBox "방화벽 정책 자동화 시트 구성이 완료되었습니다.", vbInformation
End Sub

Public Sub AnalyzeSecuiPolicyExport()
    Dim requestsSheet As Worksheet
    Dim exportSheet As Worksheet
    Dim analysisSheet As Worksheet
    Dim summarySheet As Worksheet
    Dim lastRow As Long
    Dim requestRow As Long
    Dim outputRow As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set exportSheet = EnsureSheet(SECUI_POLICY_EXPORT_SHEET)
    Set analysisSheet = EnsureSheet(POLICY_ANALYSIS_SHEET)
    Set summarySheet = EnsureSheet(POLICY_SUMMARY_SHEET)
    WriteSecuiPolicyExportHeaders exportSheet
    WritePolicyAnalysisHeaders analysisSheet
    WritePolicySummarySheet summarySheet
    FormatPolicyAnalysisSheet analysisSheet
    FormatPolicySummarySheet summarySheet

    analysisSheet.Rows("2:" & analysisSheet.Rows.Count).ClearContents
    analysisSheet.Rows("2:" & analysisSheet.Rows.Count).Interior.Pattern = xlNone
    outputRow = 2

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    For requestRow = REQ_DATA_START_ROW To lastRow
        If Len(Trim$(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value))) > 0 Then
            WritePolicyAnalysisRow requestsSheet, exportSheet, analysisSheet, requestRow, outputRow
            outputRow = outputRow + 1
        End If
    Next requestRow

    FormatPolicyAnalysisSheet analysisSheet
    WritePolicySummarySheet summarySheet
    FormatPolicySummarySheet summarySheet
    ShowInfo CStr(outputRow - 2) & "건의 SECUI 기존 정책 분석 결과를 만들었습니다."
End Sub

Private Sub WritePolicyAnalysisRow(ByVal requestsSheet As Worksheet, ByVal exportSheet As Worksheet, ByVal analysisSheet As Worksheet, ByVal requestRow As Long, ByVal outputRow As Long)
    Dim result As Object
    Dim serviceText As String
    Set result = AnalyzeExistingPolicyForRequest(requestsSheet, exportSheet, requestRow)
    serviceText = RequestServiceText(requestsSheet, requestRow)

    analysisSheet.Cells(outputRow, 1).Value = result("status")
    analysisSheet.Cells(outputRow, 2).Value = requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value
    analysisSheet.Cells(outputRow, 3).Value = requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value
    analysisSheet.Cells(outputRow, 4).Value = requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value
    analysisSheet.Cells(outputRow, 5).Value = requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value
    analysisSheet.Cells(outputRow, 6).Value = serviceText
    analysisSheet.Cells(outputRow, 7).Value = result("policy_name")
    analysisSheet.Cells(outputRow, 8).Value = result("policy_state")
    analysisSheet.Cells(outputRow, 9).Value = result("reason")
    analysisSheet.Cells(outputRow, 10).Value = result("action_note")
    analysisSheet.Cells(outputRow, 11).Value = requestRow
    analysisSheet.Cells(outputRow, 12).Value = result("policy_row")
    analysisSheet.Cells(outputRow, 13).Value = result("raw_source")
    analysisSheet.Cells(outputRow, 14).Value = result("raw_destination")
    analysisSheet.Cells(outputRow, 15).Value = result("raw_service")
    analysisSheet.Cells(outputRow, 16).Value = NormalizePolicyAddress(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value))
    analysisSheet.Cells(outputRow, 17).Value = NormalizePolicyAddress(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value))
    analysisSheet.Cells(outputRow, 18).Value = UCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    analysisSheet.Cells(outputRow, 19).Value = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    analysisSheet.Cells(outputRow, 20).Value = result("debug_note")

    ColorPolicyAnalysisRow analysisSheet, outputRow, CStr(result("status"))
End Sub

Private Function AnalyzeExistingPolicyForRequest(ByVal requestsSheet As Worksheet, ByVal exportSheet As Worksheet, ByVal requestRow As Long) As Object
    Dim best As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim reqSource As String
    Dim reqDestination As String
    Dim reqProtocol As String
    Dim reqPort As String
    Dim reqFirewalls As String
    Dim sawUnresolved As Boolean

    Set best = NewPolicyAnalysisResult("NO_EXISTING_POLICY", "", "", "없음", "일치하는 기존 SECUI export 정책 없음", "신규 정책 생성 검토", "", "", "", "", "", "")
    reqSource = NormalizePolicyAddress(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value))
    reqDestination = NormalizePolicyAddress(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value))
    reqProtocol = UCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    reqPort = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    reqFirewalls = Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value))

    lastRow = exportSheet.Cells(exportSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        Dim policyName As String
        Dim policyFirewall As String
        Dim policySource As String
        Dim policyDestination As String
        Dim policyService As String
        Dim policyAction As String
        Dim policyEnabled As Boolean
        Dim addressMatch As Boolean
        Dim serviceMatch As Boolean
        Dim firewallMatch As Boolean

        policyName = Trim$(CStr(exportSheet.Cells(rowIndex, 2).Value))
        policyFirewall = Trim$(CStr(exportSheet.Cells(rowIndex, 3).Value))
        policySource = CStr(exportSheet.Cells(rowIndex, 4).Value)
        policyDestination = CStr(exportSheet.Cells(rowIndex, 5).Value)
        policyService = CStr(exportSheet.Cells(rowIndex, 6).Value)
        policyAction = NormalizePolicyAction(CStr(exportSheet.Cells(rowIndex, 7).Value))
        policyEnabled = FirewallRowEnabled(exportSheet.Cells(rowIndex, 8).Value)
        firewallMatch = PolicyFirewallMatches(reqFirewalls, policyFirewall)
        addressMatch = PolicyValueMatches(reqSource, NormalizePolicyAddress(policySource)) And _
            PolicyValueMatches(reqDestination, NormalizePolicyAddress(policyDestination))
        serviceMatch = PolicyServiceMatches(reqProtocol, reqPort, policyService)

        If HasUnresolvedPolicyObject(policySource) Or HasUnresolvedPolicyObject(policyDestination) Or HasUnresolvedPolicyObject(policyService) Then
            sawUnresolved = True
        End If

        If firewallMatch And addressMatch And serviceMatch Then
            If policyEnabled Then
                If policyAction = "DENY" Then
                    Set AnalyzeExistingPolicyForRequest = NewPolicyAnalysisResult("EXISTING_DENY", policyName, CStr(rowIndex), "사용", "차단 정책이 같은 트래픽과 일치", "방화벽 담당자 검토 필요", policySource, policyDestination, policyService, reqSource, reqDestination, "deny match")
                    Exit Function
                ElseIf policyAction = "ALLOW" Then
                    Set AnalyzeExistingPolicyForRequest = NewPolicyAnalysisResult("EXISTING_ALLOW", policyName, CStr(rowIndex), "사용", "출발지/목적지/서비스/허용 정책 일치", "기존 정책 확인 후 신청 처리 생략 검토", policySource, policyDestination, policyService, reqSource, reqDestination, "allow match")
                    Exit Function
                End If
            Else
                Set best = NewPolicyAnalysisResult("DISABLED_MATCH", policyName, CStr(rowIndex), "비활성", "비활성 정책이 같은 트래픽과 일치", "정책 활성 여부 검토", policySource, policyDestination, policyService, reqSource, reqDestination, "disabled match")
            End If
        ElseIf firewallMatch And addressMatch And Len(CStr(best("policy_name"))) = 0 Then
            Set best = NewPolicyAnalysisResult("PARTIAL_MATCH", policyName, CStr(rowIndex), IIf(policyEnabled, "사용", "비활성"), "출발지/목적지는 같으나 서비스 또는 동작이 다름", "기존 정책 수정/신규 정책 여부 검토", policySource, policyDestination, policyService, reqSource, reqDestination, "partial address match")
        End If
    Next rowIndex

    If CStr(best("status")) = "NO_EXISTING_POLICY" And sawUnresolved Then
        Set best = NewPolicyAnalysisResult("OBJECT_UNRESOLVED", "", "", "검토필요", "SECUI export에 객체명/서비스명이 있어 CIDR 또는 포트로 확정할 수 없음", "객체 원본을 확인해 분석값 보정", "", "", "", reqSource, reqDestination, "unresolved object")
    End If
    Set AnalyzeExistingPolicyForRequest = best
End Function

Private Function NewPolicyAnalysisResult(ByVal statusText As String, ByVal policyName As String, ByVal policyRow As String, ByVal policyState As String, ByVal reasonText As String, ByVal actionNote As String, ByVal rawSource As String, ByVal rawDestination As String, ByVal rawService As String, ByVal normalizedSource As String, ByVal normalizedDestination As String, ByVal debugNote As String) As Object
    Set NewPolicyAnalysisResult = CreateObject("Scripting.Dictionary")
    NewPolicyAnalysisResult("status") = statusText
    NewPolicyAnalysisResult("policy_name") = policyName
    NewPolicyAnalysisResult("policy_row") = policyRow
    NewPolicyAnalysisResult("policy_state") = policyState
    NewPolicyAnalysisResult("reason") = reasonText
    NewPolicyAnalysisResult("action_note") = actionNote
    NewPolicyAnalysisResult("raw_source") = rawSource
    NewPolicyAnalysisResult("raw_destination") = rawDestination
    NewPolicyAnalysisResult("raw_service") = rawService
    NewPolicyAnalysisResult("normalized_source") = normalizedSource
    NewPolicyAnalysisResult("normalized_destination") = normalizedDestination
    NewPolicyAnalysisResult("debug_note") = debugNote
End Function

Private Function RequestServiceText(ByVal requestsSheet As Worksheet, ByVal requestRow As Long) As String
    Dim proto As String
    Dim portText As String
    proto = UCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    portText = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    If Len(proto) = 0 And Len(portText) = 0 Then
        RequestServiceText = ""
    ElseIf Len(portText) = 0 Then
        RequestServiceText = proto
    Else
        RequestServiceText = proto & "/" & portText
    End If
End Function

Private Function NormalizePolicyAction(ByVal actionText As String) As String
    Dim value As String
    value = LCase$(Trim$(actionText))
    If value = "allow" Or value = "accept" Or value = "permit" Or value = "pass" Or value = "허용" Then
        NormalizePolicyAction = "ALLOW"
    ElseIf value = "deny" Or value = "drop" Or value = "reject" Or value = "차단" Then
        NormalizePolicyAction = "DENY"
    Else
        NormalizePolicyAction = UCase$(value)
    End If
End Function

Private Function NormalizePolicyAddress(ByVal value As String) As String
    NormalizePolicyAddress = Trim$(Replace(Replace(value, vbCr, ";"), vbLf, ";"))
End Function

Private Function PolicyFirewallMatches(ByVal requestFirewalls As String, ByVal policyFirewall As String) As Boolean
    If Len(Trim$(policyFirewall)) = 0 Or IsAnyPolicyValue(policyFirewall) Then
        PolicyFirewallMatches = True
    ElseIf Len(Trim$(requestFirewalls)) = 0 Then
        PolicyFirewallMatches = True
    Else
        PolicyFirewallMatches = InStr(1, ";" & requestFirewalls & ";", ";" & policyFirewall & ";", vbTextCompare) > 0
    End If
End Function

Private Function PolicyValueMatches(ByVal requestValue As String, ByVal policyValue As String) As Boolean
    If IsAnyPolicyValue(policyValue) Then
        PolicyValueMatches = True
    Else
        PolicyValueMatches = InStr(1, ";" & policyValue & ";", ";" & requestValue & ";", vbTextCompare) > 0 Or _
            StrComp(Trim$(requestValue), Trim$(policyValue), vbTextCompare) = 0
    End If
End Function

Private Function PolicyServiceMatches(ByVal requestProtocol As String, ByVal requestPort As String, ByVal policyService As String) As Boolean
    Dim value As String
    value = LCase$(Trim$(policyService))
    If IsAnyPolicyValue(value) Then
        PolicyServiceMatches = True
    Else
        PolicyServiceMatches = InStr(1, value, LCase$(requestProtocol), vbTextCompare) > 0 And _
            (Len(requestPort) = 0 Or InStr(1, value, LCase$(requestPort), vbTextCompare) > 0)
    End If
End Function

Private Function IsAnyPolicyValue(ByVal value As String) As Boolean
    Dim text As String
    text = UCase$(Trim$(value))
    IsAnyPolicyValue = (Len(text) = 0 Or text = "ANY" Or text = "ALL" Or text = "*" Or text = "0.0.0.0/0")
End Function

Private Function HasUnresolvedPolicyObject(ByVal value As String) As Boolean
    Dim text As String
    text = Trim$(value)
    If Len(text) = 0 Or IsAnyPolicyValue(text) Then Exit Function
    HasUnresolvedPolicyObject = (InStr(1, text, ".", vbTextCompare) = 0 And _
        InStr(1, text, "/", vbTextCompare) = 0 And _
        InStr(1, text, "tcp", vbTextCompare) = 0 And _
        InStr(1, text, "udp", vbTextCompare) = 0 And _
        InStr(1, text, "icmp", vbTextCompare) = 0 And _
        Not IsNumeric(text))
End Function

Private Sub ColorPolicyAnalysisRow(ByVal worksheet As Worksheet, ByVal rowIndex As Long, ByVal statusText As String)
    Select Case statusText
        Case "EXISTING_ALLOW"
            worksheet.Range("A" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(217, 234, 211)
        Case "EXISTING_DENY", "NO_EXISTING_POLICY"
            worksheet.Range("A" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(244, 204, 204)
        Case Else
            worksheet.Range("A" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(255, 242, 204)
    End Select
End Sub

Public Sub MergeFirewallRequestFolder()
    Dim folderPath As String
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim logSheet As Worksheet
    Dim nextRow As Long
    Dim mergedCount As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    Set logSheet = EnsureSheet(LOG_SHEET)
    WriteRequestHeaders requestsSheet
    WriteFirewallHeaders firewallsSheet
    WriteSettings settingsSheet
    WriteLogHeaders logSheet
    LoadUserAliases settingsSheet
    mParseSheetName = SettingsValue(settingsSheet, "parse_sheet")
    folderPath = RequestFolderPath(settingsSheet)
    If Len(folderPath) = 0 Then Exit Sub

    requestsSheet.Rows(REQ_DATA_START_ROW & ":" & requestsSheet.Rows.Count).Clear
    logSheet.Rows("2:" & logSheet.Rows.Count).ClearContents
    nextRow = REQ_DATA_START_ROW
    mergedCount = MergeFolderFiles(folderPath, requestsSheet, firewallsSheet, logSheet, nextRow)
    FormatRequestsSheet requestsSheet
    FormatLogSheet logSheet

    ShowInfo CStr(mergedCount) & "건의 신청서를 통합했습니다."
End Sub

Public Sub SelectRequestFolder()
    Dim settingsSheet As Worksheet
    Dim folderPath As String

    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    WriteSettings settingsSheet
    folderPath = PickFolder(Trim$(CStr(settingsSheet.Range("B2").Value)))
    If Len(folderPath) = 0 Then Exit Sub
    settingsSheet.Range("B2").Value = folderPath
    MsgBox "신청서 폴더를 등록했습니다: " & folderPath, vbInformation
End Sub

Public Sub CreateSampleRequestWorkbook()
    Dim sampleBook As Workbook
    Dim sampleSheet As Worksheet
    Dim outputPath As Variant

    Set sampleBook = Workbooks.Add
    Set sampleSheet = sampleBook.Worksheets(1)
    sampleSheet.Name = "firewall_requests"
    sampleSheet.Range("B1:N1").Value = Array("No", "대상방화벽", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고")
    sampleSheet.Range("B2:N2").Value = Array(1, "SECUI-FW-01", "10.10.10.0/24", "업무PC", "172.16.1.10", "업무시스템", "TCP", "443", "IN", "HTTPS 업무 연동", "2026-01-01", "2026-12-31", "정기 신청")
    sampleSheet.Columns("A:N").AutoFit
    sampleSheet.Rows(1).Font.Bold = True
    sampleSheet.Range("B1:N1").AutoFilter

    outputPath = Application.GetSaveAsFilename(InitialFileName:="firewall-request-template.xlsx", FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If outputPath = False Then
        sampleBook.Close SaveChanges:=False
        Exit Sub
    End If

    sampleBook.SaveAs Filename:=CStr(outputPath), FileFormat:=xlOpenXMLWorkbook
    sampleBook.Close SaveChanges:=False
    MsgBox "샘플 신청서를 생성했습니다: " & CStr(outputPath), vbInformation
End Sub

Public Sub ConvertRequestsToSecuiBatch()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim secuiBatchSheet As Worksheet
    Dim secuiFirewalls As Object
    Dim requestRow As Long
    Dim secuiRow As Long
    Dim lastRow As Long
    Dim convertedRows As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set secuiBatchSheet = EnsureSheet(SECUI_BATCH_SHEET)
    WriteFirewallHeaders firewallsSheet
    WriteSecuiBatchHeaders secuiBatchSheet
    Set secuiFirewalls = LoadSecuiFirewalls(firewallsSheet)
    secuiBatchSheet.Rows("2:" & secuiBatchSheet.Rows.Count).Clear

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    secuiRow = 2
    For requestRow = REQ_DATA_START_ROW To lastRow
        convertedRows = convertedRows + CopySecuiBatchRows(requestsSheet, secuiBatchSheet, secuiFirewalls, requestRow, secuiRow)
    Next requestRow

    FormatSecuiBatchSheet secuiBatchSheet
    ShowInfo CStr(convertedRows) & "건의 SECUI 배치 행을 생성했습니다."
End Sub

Public Sub ConvertRequestsToSecuiCli()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim secuiCliSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim secuiFirewalls As Object
    Dim cliTemplate As Object
    Dim cliGroups As Object
    Dim serviceFanoutIndex As Object
    Dim groupKey As Variant
    Dim requestRow As Long
    Dim cliRow As Long
    Dim lastRow As Long
    Dim convertedRows As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set secuiCliSheet = EnsureSheet(SECUI_CLI_SHEET)
    Set templateSheet = EnsureSheet(VENDOR_CLI_TEMPLATE_SHEET)
    WriteFirewallHeaders firewallsSheet
    WriteSecuiCliHeaders secuiCliSheet
    WriteVendorCliTemplateHeaders templateSheet
    Set secuiFirewalls = LoadSecuiFirewalls(firewallsSheet)
    Set cliTemplate = LoadVendorCliTemplate(templateSheet, "SECUI")
    Set cliGroups = CreateObject("Scripting.Dictionary")

    secuiCliSheet.Rows("2:" & secuiCliSheet.Rows.Count).Clear

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    Set serviceFanoutIndex = BuildSecuiCliServiceFanoutIndex(requestsSheet, secuiFirewalls, lastRow)
    For requestRow = REQ_DATA_START_ROW To lastRow
        CollectSecuiCliRows requestsSheet, secuiFirewalls, serviceFanoutIndex, requestRow, cliGroups
    Next requestRow

    cliRow = 2
    For Each groupKey In cliGroups.Keys
        WriteSecuiCliGroup requestsSheet, secuiCliSheet, cliGroups(groupKey), cliRow, cliTemplate
        cliRow = cliRow + 1
        convertedRows = convertedRows + 1
    Next groupKey

    FormatSecuiCliSheet secuiCliSheet
    ShowInfo CStr(convertedRows) & "건의 SECUI CLI 명령 초안을 생성했습니다."
End Sub

Private Function CopySecuiBatchRows(ByVal requestsSheet As Worksheet, ByVal secuiBatchSheet As Worksheet, ByVal secuiFirewalls As Object, ByVal requestRow As Long, ByRef secuiRow As Long) As Long
    Dim targetFirewalls As Variant
    Dim firewallValue As Variant
    Dim firewallName As String
    Dim written As Long

    targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")
    For Each firewallValue In targetFirewalls
        firewallName = Trim$(CStr(firewallValue))
        If Len(firewallName) > 0 And secuiFirewalls.Exists(SecuiFirewallKey(firewallName)) Then
            WriteSecuiBatchRow requestsSheet, secuiBatchSheet, requestRow, secuiRow, firewallName
            secuiRow = secuiRow + 1
            written = written + 1
        End If
    Next firewallValue

    CopySecuiBatchRows = written
End Function

Private Function CopySecuiCliRows(ByVal requestsSheet As Worksheet, ByVal secuiCliSheet As Worksheet, ByVal secuiFirewalls As Object, ByVal cliTemplate As Object, ByVal requestRow As Long, ByRef cliRow As Long) As Long
    Dim targetFirewalls As Variant
    Dim firewallValue As Variant
    Dim firewallName As String
    Dim written As Long

    targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")
    For Each firewallValue In targetFirewalls
        firewallName = Trim$(CStr(firewallValue))
        If Len(firewallName) > 0 And secuiFirewalls.Exists(SecuiFirewallKey(firewallName)) Then
            WriteSecuiCliRow requestsSheet, secuiCliSheet, requestRow, cliRow, firewallName, cliTemplate
            cliRow = cliRow + 1
            written = written + 1
        End If
    Next firewallValue

    CopySecuiCliRows = written
End Function

Private Function BuildSecuiCliServiceFanoutIndex(ByVal requestsSheet As Worksheet, ByVal secuiFirewalls As Object, ByVal lastRow As Long) As Object
    Dim serviceFanoutIndex As Object
    Dim services As Object
    Dim requestRow As Long
    Dim targetFirewalls As Variant
    Dim firewallValue As Variant
    Dim firewallName As String
    Dim sourceObject As String
    Dim sourceAddress As String
    Dim destinationObject As String
    Dim destinationAddress As String
    Dim serviceObject As String
    Dim sourceDestinationKey As String

    Set serviceFanoutIndex = CreateObject("Scripting.Dictionary")
    For requestRow = REQ_DATA_START_ROW To lastRow
        sourceAddress = CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value)
        destinationAddress = CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value)
        sourceObject = SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_NAME).Value), sourceAddress)
        destinationObject = SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value), destinationAddress)
        serviceObject = SecuiCliServiceText(LCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value))), Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value)))

        targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")
        For Each firewallValue In targetFirewalls
            firewallName = Trim$(CStr(firewallValue))
            If Len(firewallName) > 0 And secuiFirewalls.Exists(SecuiFirewallKey(firewallName)) Then
                sourceDestinationKey = SecuiCliSourceDestinationKey(firewallName, sourceObject, sourceAddress, destinationObject, destinationAddress)
                If Not serviceFanoutIndex.Exists(sourceDestinationKey) Then
                    Set services = CreateObject("Scripting.Dictionary")
                    Set serviceFanoutIndex(sourceDestinationKey) = services
                Else
                    Set services = serviceFanoutIndex(sourceDestinationKey)
                End If
                services(CleanSecuiText(serviceObject)) = True
            End If
        Next firewallValue
    Next requestRow
    Set BuildSecuiCliServiceFanoutIndex = serviceFanoutIndex
End Function

Private Sub CollectSecuiCliRows(ByVal requestsSheet As Worksheet, ByVal secuiFirewalls As Object, ByVal serviceFanoutIndex As Object, ByVal requestRow As Long, ByVal cliGroups As Object)
    Dim targetFirewalls As Variant
    Dim firewallValue As Variant
    Dim firewallName As String
    Dim sourceObject As String
    Dim sourceAddress As String
    Dim destinationObject As String
    Dim destinationAddress As String
    Dim serviceObject As String
    Dim sourceDestinationKey As String
    Dim groupKey As String
    Dim group As Object

    sourceAddress = CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value)
    destinationAddress = CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value)
    sourceObject = SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_NAME).Value), sourceAddress)
    destinationObject = SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value), destinationAddress)
    serviceObject = SecuiCliServiceText(LCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value))), Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value)))

    targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")
    For Each firewallValue In targetFirewalls
        firewallName = Trim$(CStr(firewallValue))
        If Len(firewallName) > 0 And secuiFirewalls.Exists(SecuiFirewallKey(firewallName)) Then
            sourceDestinationKey = SecuiCliSourceDestinationKey(firewallName, sourceObject, sourceAddress, destinationObject, destinationAddress)
            If serviceFanoutIndex.Exists(sourceDestinationKey) And serviceFanoutIndex(sourceDestinationKey).Count > 1 Then
                groupKey = "SRC_DST" & Chr$(30) & sourceDestinationKey
            Else
                groupKey = "DST_SVC" & Chr$(30) & SecuiCliDestinationServiceKey(firewallName, destinationObject, destinationAddress, serviceObject)
            End If
            If Not cliGroups.Exists(groupKey) Then
                Set group = NewSecuiCliGroup(requestsSheet, requestRow, firewallName, destinationObject, serviceObject)
                Set cliGroups(groupKey) = group
            Else
                Set group = cliGroups(groupKey)
            End If
            AddSecuiCliGroupSource group, sourceObject, sourceAddress, requestRow
            AddSecuiCliGroupService group, serviceObject
        End If
    Next firewallValue
End Sub

Private Function NewSecuiCliGroup(ByVal requestsSheet As Worksheet, ByVal requestRow As Long, ByVal firewallName As String, ByVal destinationObject As String, ByVal serviceObject As String) As Object
    Dim group As Object
    Set group = CreateObject("Scripting.Dictionary")
    group("firewall_name") = firewallName
    group("request_row") = requestRow
    Dim rangeInfo As Object
    Set rangeInfo = FirstMatchingFirewallRangeInfo( _
        firewallName, _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_DIRECTION).Value))

    group("destination_address") = CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value)
    group("destination_name") = CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value)
    group("destination_object") = destinationObject
    group("protocol") = LCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    group("port") = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    group("service_object") = serviceObject
    group("source_interface") = CStr(rangeInfo("source_interface"))
    group("destination_interface") = CStr(rangeInfo("destination_interface"))
    group("network_scope") = CStr(rangeInfo("network_scope"))
    Set group("sources") = CreateObject("Scripting.Dictionary")
    Set group("source_rows") = CreateObject("Scripting.Dictionary")
    Set group("services") = CreateObject("Scripting.Dictionary")
    AddSecuiCliGroupService group, serviceObject
    Set NewSecuiCliGroup = group
End Function

Private Sub AddSecuiCliGroupSource(ByVal group As Object, ByVal sourceObject As String, ByVal sourceAddress As String, ByVal requestRow As Long)
    Dim sources As Object
    Dim sourceRows As Object
    Dim cleanSourceObject As String

    Set sources = group("sources")
    Set sourceRows = group("source_rows")
    cleanSourceObject = CleanSecuiText(sourceObject)
    If Len(cleanSourceObject) = 0 Then Exit Sub
    sources(cleanSourceObject) = CleanSecuiText(sourceAddress)
    sourceRows(CStr(requestRow)) = True
End Sub

Private Sub AddSecuiCliGroupService(ByVal group As Object, ByVal serviceObject As String)
    Dim services As Object
    Dim cleanServiceObject As String

    Set services = group("services")
    cleanServiceObject = CleanSecuiText(serviceObject)
    If Len(cleanServiceObject) = 0 Then Exit Sub
    services(cleanServiceObject) = True
    group("service_object") = JoinedDictionaryKeys(services, ";")
End Sub

Private Function SecuiCliSourceDestinationKey(ByVal firewallName As String, ByVal sourceObject As String, ByVal sourceAddress As String, ByVal destinationObject As String, ByVal destinationAddress As String) As String
    SecuiCliSourceDestinationKey = LCase$(CleanSecuiText(firewallName)) & Chr$(30) & _
        LCase$(CleanSecuiText(sourceObject)) & Chr$(30) & _
        LCase$(CleanSecuiText(sourceAddress)) & Chr$(30) & _
        LCase$(CleanSecuiText(destinationObject)) & Chr$(30) & _
        LCase$(CleanSecuiText(destinationAddress))
End Function

Private Function SecuiCliDestinationServiceKey(ByVal firewallName As String, ByVal destinationObject As String, ByVal destinationAddress As String, ByVal serviceObject As String) As String
    SecuiCliDestinationServiceKey = LCase$(CleanSecuiText(firewallName)) & Chr$(30) & _
        LCase$(CleanSecuiText(destinationObject)) & Chr$(30) & _
        LCase$(CleanSecuiText(destinationAddress)) & Chr$(30) & _
        LCase$(CleanSecuiText(serviceObject))
End Function

Private Function LoadSecuiFirewalls(ByVal firewallsSheet As Worksheet) As Object
    Dim secuiFirewalls As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim firewallName As String
    Dim vendorName As String

    Set secuiFirewalls = CreateObject("Scripting.Dictionary")
    lastRow = firewallsSheet.Cells(firewallsSheet.Rows.Count, FW_COL_NAME).End(xlUp).Row
    For rowIndex = 2 To lastRow
        firewallName = Trim$(CStr(firewallsSheet.Cells(rowIndex, FW_COL_NAME).Value))
        vendorName = UCase$(Trim$(CStr(firewallsSheet.Cells(rowIndex, FW_COL_VENDOR).Value)))
        If Len(firewallName) > 0 _
                And vendorName = "SECUI" _
                And FirewallRowEnabled(firewallsSheet.Cells(rowIndex, FW_COL_ENABLED).Value) Then
            secuiFirewalls(SecuiFirewallKey(firewallName)) = True
        End If
    Next rowIndex

    Set LoadSecuiFirewalls = secuiFirewalls
End Function

Private Function FirewallRowEnabled(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = UCase$(Trim$(CStr(value)))
    FirewallRowEnabled = (Len(textValue) = 0 Or textValue = "Y" Or textValue = "YES" Or textValue = "TRUE" Or textValue = "1")
End Function

Private Function SecuiFirewallKey(ByVal firewallName As String) As String
    SecuiFirewallKey = LCase$(Trim$(firewallName))
End Function

Private Function LoadVendorCliTemplate(ByVal templateSheet As Worksheet, ByVal vendorFilter As String) As Object
    Dim result As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim vendorName As String
    Dim commandTemplate As String

    Set result = CreateObject("Scripting.Dictionary")
    result("command_template") = DefaultVendorCliTemplate()
    result("review_note") = DefaultVendorCliReviewNote()

    lastRow = templateSheet.Cells(templateSheet.Rows.Count, CLI_TEMPLATE_COL_VENDOR).End(xlUp).Row
    For rowIndex = 2 To lastRow
        vendorName = UCase$(Trim$(CStr(templateSheet.Cells(rowIndex, CLI_TEMPLATE_COL_VENDOR).Value)))
        commandTemplate = Trim$(CStr(templateSheet.Cells(rowIndex, CLI_TEMPLATE_COL_COMMAND).Value))
        If vendorName = UCase$(Trim$(vendorFilter)) _
                And FirewallRowEnabled(templateSheet.Cells(rowIndex, CLI_TEMPLATE_COL_ENABLED).Value) _
                And Len(commandTemplate) > 0 Then
            result("command_template") = commandTemplate
            result("review_note") = CStr(templateSheet.Cells(rowIndex, CLI_TEMPLATE_COL_NOTE).Value)
            Exit For
        End If
    Next rowIndex

    Set LoadVendorCliTemplate = result
End Function

Private Sub WriteSecuiBatchRow(ByVal requestsSheet As Worksheet, ByVal secuiBatchSheet As Worksheet, ByVal requestRow As Long, ByVal secuiRow As Long, ByVal firewallName As String)
    Dim proto As String
    Dim portText As String
    Dim serviceText As String

    proto = UCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    portText = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    serviceText = proto
    If Len(portText) > 0 Then serviceText = AppendToken(serviceText, portText, "/")

    secuiBatchSheet.Cells(secuiRow, 1).Value = secuiRow - 1
    secuiBatchSheet.Cells(secuiRow, 2).Value = firewallName
    secuiBatchSheet.Cells(secuiRow, 3).Value = SecuiPolicyName(requestsSheet, requestRow, firewallName)
    secuiBatchSheet.Cells(secuiRow, 4).Value = requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value
    secuiBatchSheet.Cells(secuiRow, 5).Value = requestsSheet.Cells(requestRow, COL_SOURCE_NAME).Value
    secuiBatchSheet.Cells(secuiRow, 6).Value = requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value
    secuiBatchSheet.Cells(secuiRow, 7).Value = requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value
    secuiBatchSheet.Cells(secuiRow, 8).Value = serviceText
    secuiBatchSheet.Cells(secuiRow, 9).Value = proto
    secuiBatchSheet.Cells(secuiRow, 10).Value = portText
    secuiBatchSheet.Cells(secuiRow, 11).Value = "Allow"
    secuiBatchSheet.Cells(secuiRow, 12).Value = "Y"
    secuiBatchSheet.Cells(secuiRow, 13).Value = "Y"
    secuiBatchSheet.Cells(secuiRow, 14).Value = requestsSheet.Cells(requestRow, COL_START_DATE).Value
    secuiBatchSheet.Cells(secuiRow, 15).Value = requestsSheet.Cells(requestRow, COL_END_DATE).Value
    secuiBatchSheet.Cells(secuiRow, 16).Value = SecuiDescription(requestsSheet, requestRow)
    secuiBatchSheet.Cells(secuiRow, 17).Value = requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value
    secuiBatchSheet.Cells(secuiRow, 18).Value = requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value
    secuiBatchSheet.Cells(secuiRow, 19).Value = requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value
    secuiBatchSheet.Cells(secuiRow, 20).Value = requestsSheet.Cells(requestRow, COL_SOURCE_ROW).Value
End Sub

Private Sub WriteSecuiCliRow(ByVal requestsSheet As Worksheet, ByVal secuiCliSheet As Worksheet, ByVal requestRow As Long, ByVal cliRow As Long, ByVal firewallName As String, ByVal cliTemplate As Object)
    Dim policyName As String
    Dim reviewNote As String
    policyName = SecuiPolicyName(requestsSheet, requestRow, firewallName)
    reviewNote = Trim$(CStr(cliTemplate("review_note")))
    If Len(reviewNote) = 0 Then reviewNote = DefaultVendorCliReviewNote()

    secuiCliSheet.Cells(cliRow, 1).Value = cliRow - 1
    secuiCliSheet.Cells(cliRow, 2).Value = firewallName
    secuiCliSheet.Cells(cliRow, 3).Value = policyName
    secuiCliSheet.Cells(cliRow, 4).Value = SecuiCliCommand(requestsSheet, requestRow, firewallName, policyName, CStr(cliTemplate("command_template")))
    secuiCliSheet.Cells(cliRow, 5).Value = reviewNote
    secuiCliSheet.Cells(cliRow, 6).Value = requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value
    secuiCliSheet.Cells(cliRow, 7).Value = requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value
    secuiCliSheet.Cells(cliRow, 8).Value = requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value
    secuiCliSheet.Cells(cliRow, 9).Value = requestsSheet.Cells(requestRow, COL_SOURCE_ROW).Value
End Sub

Private Sub WriteSecuiCliGroup(ByVal requestsSheet As Worksheet, ByVal secuiCliSheet As Worksheet, ByVal group As Object, ByVal cliRow As Long, ByVal cliTemplate As Object)
    Dim requestRow As Long
    Dim firewallName As String
    Dim policyName As String
    Dim sourceGroupName As String
    Dim destinationGroupName As String
    Dim serviceGroupName As String
    Dim reviewNote As String
    Dim commandText As String

    requestRow = CLng(group("request_row"))
    firewallName = CStr(group("firewall_name"))
    policyName = SecuiGroupPolicyName(requestsSheet, group)
    sourceGroupName = SecuiGroupObjectName("SRC", policyName, requestsSheet, group)
    destinationGroupName = SecuiGroupObjectName("DST", policyName, requestsSheet, group)
    serviceGroupName = SecuiGroupObjectName("SVC", policyName, requestsSheet, group)
    reviewNote = Trim$(CStr(cliTemplate("review_note")))
    If Len(reviewNote) = 0 Then reviewNote = DefaultVendorCliReviewNote()
    reviewNote = AppendToken(reviewNote, "룰별 그룹객체 생성 후 정책 생성", " / ")

    commandText = SecuiGroupObjectCommands(group, sourceGroupName, destinationGroupName, serviceGroupName)
    commandText = AppendCliLine(commandText, RenderVendorCliTemplate( _
        CStr(cliTemplate("command_template")), _
        firewallName, _
        policyName, _
        JoinedDictionaryKeys(group("sources"), ";"), _
        JoinedDictionaryKeys(group("sources"), ";"), _
        SecuiPolicyObjectReference(sourceGroupName, JoinedDictionaryKeys(group("sources"), ";")), _
        CStr(group("destination_address")), _
        CStr(group("destination_name")), _
        SecuiPolicyObjectReference(destinationGroupName, CStr(group("destination_object"))), _
        CStr(group("protocol")), _
        CStr(group("port")), _
        SecuiPolicyObjectReference(serviceGroupName, CStr(group("service_object"))), _
        CStr(group("source_interface")), _
        CStr(group("destination_interface")), _
        SecuiDescription(requestsSheet, requestRow), _
        CStr(requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_ROW).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_START_DATE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_END_DATE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_PURPOSE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_NOTE).Value)))

    secuiCliSheet.Cells(cliRow, 1).Value = cliRow - 1
    secuiCliSheet.Cells(cliRow, 2).Value = firewallName
    secuiCliSheet.Cells(cliRow, 3).Value = policyName
    secuiCliSheet.Cells(cliRow, 4).Value = commandText
    secuiCliSheet.Cells(cliRow, 5).Value = reviewNote
    secuiCliSheet.Cells(cliRow, 6).Value = requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value
    secuiCliSheet.Cells(cliRow, 7).Value = requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value
    secuiCliSheet.Cells(cliRow, 8).Value = requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value
    secuiCliSheet.Cells(cliRow, 9).Value = JoinedDictionaryKeys(group("source_rows"), ";")
End Sub

Private Function SecuiGroupPolicyName(ByVal requestsSheet As Worksheet, ByVal group As Object) As String
    Dim requestRow As Long
    Dim nameText As String
    requestRow = CLng(group("request_row"))
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value), "_")
    nameText = AppendToken(nameText, CStr(group("firewall_name")), "_")
    nameText = AppendToken(nameText, SecuiNetworkScopeToken(requestsSheet, group), "_")
    nameText = AppendToken(nameText, CStr(group("service_object")), "_")
    nameText = AppendToken(nameText, CStr(group("destination_object")), "_")
    nameText = AppendToken(nameText, CStr(group("destination_address")), "_")
    SecuiGroupPolicyName = Left$(SecuiObjectNameToken(nameText), 120)
End Function

Private Function SecuiGroupObjectName(ByVal prefixText As String, ByVal policyName As String, ByVal requestsSheet As Worksheet, ByVal group As Object) As String
    Dim objectName As String
    objectName = "GRP_" & prefixText & "_" & SecuiNetworkScopeToken(requestsSheet, group) & "_" & policyName
    SecuiGroupObjectName = Left$(SecuiObjectNameToken(objectName), 120)
End Function

Private Function SecuiGroupObjectCommands(ByVal group As Object, ByVal sourceGroupName As String, ByVal destinationGroupName As String, ByVal serviceGroupName As String) As String
    Dim commandText As String
    If Not IsAnyPolicyValue(JoinedDictionaryKeys(group("sources"), ";")) Then
        commandText = AppendCliLine(commandText, "fw set addrgrp name " & SecuiCliQuote(sourceGroupName) & " member " & SecuiCliQuote(JoinedDictionaryKeys(group("sources"), ";")) & " # device=" & CStr(group("firewall_name")))
    End If
    If Not IsAnyPolicyValue(CStr(group("destination_object"))) Then
        commandText = AppendCliLine(commandText, "fw set addrgrp name " & SecuiCliQuote(destinationGroupName) & " member " & SecuiCliQuote(CStr(group("destination_object"))) & " # device=" & CStr(group("firewall_name")))
    End If
    If Not IsAnyPolicyValue(CStr(group("service_object"))) Then
        commandText = AppendCliLine(commandText, "fw set svcgrp name " & SecuiCliQuote(serviceGroupName) & " member " & SecuiCliQuote(CStr(group("service_object"))) & " # device=" & CStr(group("firewall_name")))
    End If
    SecuiGroupObjectCommands = commandText
End Function

Private Function SecuiPolicyObjectReference(ByVal groupName As String, ByVal memberText As String) As String
    If IsAnyPolicyValue(memberText) Then
        SecuiPolicyObjectReference = "ANY"
    Else
        SecuiPolicyObjectReference = groupName
    End If
End Function

Private Function JoinedDictionaryKeys(ByVal values As Object, ByVal delimiter As String) As String
    Dim key As Variant
    Dim result As String
    For Each key In values.Keys
        result = AppendToken(result, CStr(key), delimiter)
    Next key
    JoinedDictionaryKeys = result
End Function

Private Function SecuiNetworkScopeToken(ByVal requestsSheet As Worksheet, ByVal group As Object) As String
    Dim requestRow As Long
    Dim sourceZone As String
    Dim destinationZone As String
    Dim rangeScope As String

    requestRow = CLng(group("request_row"))
    sourceZone = CleanSecuiText(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_ZONE).Value))
    destinationZone = CleanSecuiText(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_ZONE).Value))
    If Len(sourceZone) > 0 Or Len(destinationZone) > 0 Then
        SecuiNetworkScopeToken = SecuiObjectNameToken(sourceZone & "_TO_" & destinationZone)
        Exit Function
    End If

    rangeScope = CleanSecuiText(CStr(group("network_scope")))
    If Len(rangeScope) > 0 Then
        SecuiNetworkScopeToken = SecuiObjectNameToken(rangeScope)
    Else
        SecuiNetworkScopeToken = "ANY_TO_ANY"
    End If
End Function

Private Function FirstMatchingFirewallRangeInfo(ByVal firewallName As String, ByVal sourceText As String, ByVal destinationText As String, ByVal directionText As String) As Object
    Dim info As Object
    Dim rangeSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set info = CreateObject("Scripting.Dictionary")
    info("network_scope") = ""
    info("source_interface") = "ANY"
    info("destination_interface") = "ANY"

    Set rangeSheet = EnsureSheet(FIREWALL_RANGE_SHEET)
    lastRow = rangeSheet.Cells(rangeSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(rangeSheet.Cells(rowIndex, 1).Value)), firewallName, vbTextCompare) = 0 _
                And FirewallRowEnabled(rangeSheet.Cells(rowIndex, 6).Value) _
                And SecuiDirectionMatches(CStr(rangeSheet.Cells(rowIndex, 4).Value), directionText) _
                And SecuiAddressListOverlaps(sourceText, CStr(rangeSheet.Cells(rowIndex, 2).Value)) _
                And SecuiAddressListOverlaps(destinationText, CStr(rangeSheet.Cells(rowIndex, 3).Value)) Then
            If Len(Trim$(CStr(rangeSheet.Cells(rowIndex, 10).Value))) > 0 Or Len(Trim$(CStr(rangeSheet.Cells(rowIndex, 11).Value))) > 0 Then
                info("network_scope") = CStr(rangeSheet.Cells(rowIndex, 10).Value) & "_TO_" & CStr(rangeSheet.Cells(rowIndex, 11).Value)
            Else
                info("network_scope") = CStr(rangeSheet.Cells(rowIndex, 2).Value) & "_TO_" & CStr(rangeSheet.Cells(rowIndex, 3).Value)
            End If
            If Len(Trim$(CStr(rangeSheet.Cells(rowIndex, 8).Value))) > 0 Then info("source_interface") = CStr(rangeSheet.Cells(rowIndex, 8).Value)
            If Len(Trim$(CStr(rangeSheet.Cells(rowIndex, 9).Value))) > 0 Then info("destination_interface") = CStr(rangeSheet.Cells(rowIndex, 9).Value)
            Set FirstMatchingFirewallRangeInfo = info
            Exit Function
        End If
    Next rowIndex
    Set FirstMatchingFirewallRangeInfo = info
End Function

Private Function SecuiDirectionMatches(ByVal ruleDirection As String, ByVal requestDirection As String) As Boolean
    Dim ruleValue As String
    Dim requestValue As String
    ruleValue = UCase$(Trim$(ruleDirection))
    requestValue = UCase$(Trim$(requestDirection))
    If Len(ruleValue) = 0 Then ruleValue = "BOTH"
    If Len(requestValue) = 0 Then requestValue = "BOTH"
    If ruleValue = "BOTH" Or requestValue = "BOTH" Then
        SecuiDirectionMatches = True
    Else
        SecuiDirectionMatches = (ruleValue = requestValue)
    End If
End Function

Private Function SecuiAddressListOverlaps(ByVal requestValue As String, ByVal definitionValue As String) As Boolean
    Dim requestParts As Variant
    Dim definitionParts As Variant
    Dim requestPart As Variant
    Dim definitionPart As Variant

    If IsAnyPolicyValue(definitionValue) Then
        SecuiAddressListOverlaps = True
        Exit Function
    End If

    requestParts = Split(NormalizeListCell(requestValue), ";")
    definitionParts = Split(NormalizeListCell(definitionValue), ";")
    For Each requestPart In requestParts
        For Each definitionPart In definitionParts
            If SecuiRangesOverlap(CStr(requestPart), CStr(definitionPart)) Then
                SecuiAddressListOverlaps = True
                Exit Function
            End If
        Next definitionPart
    Next requestPart
End Function

Private Function SecuiRangesOverlap(ByVal leftCidr As String, ByVal rightCidr As String) As Boolean
    On Error GoTo InvalidRange
    If IsAnyPolicyValue(leftCidr) Or IsAnyPolicyValue(rightCidr) Then
        SecuiRangesOverlap = True
        Exit Function
    End If
    If Len(Trim$(leftCidr)) = 0 Or Len(Trim$(rightCidr)) = 0 Then Exit Function

    Dim leftStart As Double
    Dim leftEnd As Double
    Dim rightStart As Double
    Dim rightEnd As Double
    leftStart = SecuiCidrStart(leftCidr)
    leftEnd = SecuiCidrEnd(leftCidr)
    rightStart = SecuiCidrStart(rightCidr)
    rightEnd = SecuiCidrEnd(rightCidr)
    SecuiRangesOverlap = (leftStart <= rightEnd And rightStart <= leftEnd)
    Exit Function

InvalidRange:
    SecuiRangesOverlap = False
End Function

Private Function SecuiCidrStart(ByVal cidrText As String) As Double
    Dim base As Double
    Dim block As Double
    base = SecuiIpToNumber(Split(Trim$(cidrText), "/")(0))
    block = 2 ^ (32 - SecuiCidrPrefix(cidrText))
    SecuiCidrStart = Int(base / block) * block
End Function

Private Function SecuiCidrEnd(ByVal cidrText As String) As Double
    SecuiCidrEnd = SecuiCidrStart(cidrText) + (2 ^ (32 - SecuiCidrPrefix(cidrText))) - 1
End Function

Private Function SecuiCidrPrefix(ByVal cidrText As String) As Long
    If InStr(cidrText, "/") = 0 Then
        SecuiCidrPrefix = 32
    Else
        SecuiCidrPrefix = CLng(Trim$(Split(cidrText, "/")(1)))
    End If
End Function

Private Function SecuiIpToNumber(ByVal ipText As String) As Double
    Dim parts As Variant
    Dim index As Long
    Dim value As Double
    parts = Split(Trim$(ipText), ".")
    If UBound(parts) <> 3 Then Err.Raise vbObjectError + 2001, , "Invalid IPv4 address"
    For index = 0 To 3
        value = value * 256 + CDbl(parts(index))
    Next index
    SecuiIpToNumber = value
End Function

Private Function SecuiObjectNameToken(ByVal value As String) As String
    Dim text As String
    Dim index As Long
    Dim ch As String
    Dim result As String

    text = CleanSecuiText(value)
    For index = 1 To Len(text)
        ch = Mid$(text, index, 1)
        If ch Like "[A-Za-z0-9]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next index
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    result = TrimChars(result, "_")
    If Len(result) = 0 Then result = "ANY"
    SecuiObjectNameToken = UCase$(result)
End Function

Private Function AppendCliLine(ByVal baseText As String, ByVal lineText As String) As String
    If Len(Trim$(lineText)) = 0 Then
        AppendCliLine = baseText
    ElseIf Len(baseText) = 0 Then
        AppendCliLine = lineText
    Else
        AppendCliLine = baseText & vbLf & lineText
    End If
End Function

Private Function SecuiCliCommand(ByVal requestsSheet As Worksheet, ByVal requestRow As Long, ByVal firewallName As String, ByVal policyName As String, ByVal commandTemplate As String) As String
    Dim proto As String
    Dim portText As String

    proto = LCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    portText = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))

    SecuiCliCommand = RenderVendorCliTemplate( _
        commandTemplate, _
        firewallName, _
        policyName, _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_NAME).Value), _
        SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_NAME).Value), CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value)), _
        CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value), _
        SecuiCliAddressText(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_NAME).Value), CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value)), _
        proto, _
        portText, _
        SecuiCliServiceText(proto, portText), _
        "", _
        "", _
        SecuiDescription(requestsSheet, requestRow), _
        CStr(requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_SOURCE_ROW).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_START_DATE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_END_DATE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_PURPOSE).Value), _
        CStr(requestsSheet.Cells(requestRow, COL_NOTE).Value))
End Function

Private Function RenderVendorCliTemplate(ByVal commandTemplate As String, ByVal firewallName As String, ByVal policyName As String, ByVal sourceIp As String, ByVal sourceName As String, ByVal sourceObject As String, ByVal destinationIp As String, ByVal destinationName As String, ByVal destinationObject As String, ByVal protocolText As String, ByVal portText As String, ByVal serviceText As String, ByVal sourceInterface As String, ByVal destinationInterface As String, ByVal descriptionText As String, ByVal requestTeam As String, ByVal requestDocNo As String, ByVal sourceFile As String, ByVal sourceRow As String, ByVal startDate As String, ByVal endDate As String, ByVal purposeText As String, ByVal noteText As String) As String
    Dim rendered As String
    rendered = commandTemplate
    If Len(Trim$(rendered)) = 0 Then rendered = DefaultVendorCliTemplate()

    rendered = ReplaceSecuiToken(rendered, "firewall_name", firewallName)
    rendered = ReplaceSecuiToken(rendered, "policy_name", policyName)
    rendered = ReplaceSecuiToken(rendered, "source_ip", sourceIp)
    rendered = ReplaceSecuiToken(rendered, "source_name", sourceName)
    rendered = ReplaceSecuiToken(rendered, "source_object", sourceObject)
    rendered = ReplaceSecuiToken(rendered, "destination_ip", destinationIp)
    rendered = ReplaceSecuiToken(rendered, "destination_name", destinationName)
    rendered = ReplaceSecuiToken(rendered, "destination_object", destinationObject)
    rendered = ReplaceSecuiToken(rendered, "protocol", protocolText)
    rendered = ReplaceSecuiToken(rendered, "port", portText)
    rendered = ReplaceSecuiToken(rendered, "service", serviceText)
    rendered = ReplaceSecuiToken(rendered, "service_object", serviceText)
    rendered = ReplaceSecuiToken(rendered, "source_interface", sourceInterface)
    rendered = ReplaceSecuiToken(rendered, "destination_interface", destinationInterface)
    rendered = ReplaceSecuiToken(rendered, "description", descriptionText)
    rendered = ReplaceSecuiToken(rendered, "request_team", requestTeam)
    rendered = ReplaceSecuiToken(rendered, "request_doc_no", requestDocNo)
    rendered = ReplaceSecuiToken(rendered, "source_file", sourceFile)
    rendered = ReplaceSecuiToken(rendered, "source_row", sourceRow)
    rendered = ReplaceSecuiToken(rendered, "start_date", startDate)
    rendered = ReplaceSecuiToken(rendered, "end_date", endDate)
    rendered = ReplaceSecuiToken(rendered, "purpose", purposeText)
    rendered = ReplaceSecuiToken(rendered, "note", noteText)
    RenderVendorCliTemplate = rendered
End Function

Private Function ReplaceSecuiToken(ByVal commandText As String, ByVal tokenName As String, ByVal tokenValue As String) As String
    Dim cleanedValue As String
    cleanedValue = CleanSecuiText(tokenValue)
    ReplaceSecuiToken = Replace(commandText, "{" & tokenName & "}", cleanedValue)
    ReplaceSecuiToken = Replace(ReplaceSecuiToken, "{" & tokenName & "_q}", SecuiCliQuote(cleanedValue))
End Function

Private Function SecuiCliServiceText(ByVal proto As String, ByVal portText As String) As String
    Dim serviceText As String
    serviceText = CleanSecuiText(portText)
    If Len(serviceText) = 0 Or IsNumeric(serviceText) Then
        SecuiCliServiceText = CleanSecuiText(proto & "/" & portText)
    Else
        SecuiCliServiceText = serviceText
    End If
End Function

Private Function SecuiCliAddressText(ByVal objectName As String, ByVal fallbackAddress As String) As String
    Dim objectText As String
    objectText = CleanSecuiText(objectName)
    If Len(objectText) > 0 Then
        SecuiCliAddressText = objectText
    Else
        SecuiCliAddressText = CleanSecuiText(fallbackAddress)
    End If
End Function

Private Function DefaultVendorCliTemplate() As String
    DefaultVendorCliTemplate = "fw set srule name {policy_name_q} action allow srcif {source_interface_q} dstif {destination_interface_q} src {source_object_q} dst {destination_object_q} service {service_object_q} log enable enable yes description {description_q} # device={firewall_name}"
End Function

Private Function DefaultVendorCliReviewNote() As String
    DefaultVendorCliReviewNote = "장비 CLI에서 'fw set srule help'로 옵션명 확인 후 적용"
End Function

Private Function SecuiCliQuote(ByVal value As String) As String
    SecuiCliQuote = """" & Replace(CleanSecuiText(value), """", "'") & """"
End Function

Private Function SecuiPolicyName(ByVal requestsSheet As Worksheet, ByVal requestRow As Long, ByVal firewallName As String) As String
    Dim nameText As String
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value), "_")
    nameText = AppendToken(nameText, firewallName, "_")
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value), "_")
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_PORT).Value), "_")
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value), "_")
    nameText = AppendToken(nameText, CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value), "_")
    SecuiPolicyName = Left$(CleanSecuiText(nameText), 120)
End Function

Private Function SecuiDescription(ByVal requestsSheet As Worksheet, ByVal requestRow As Long) As String
    Dim descText As String
    descText = AppendToken(descText, CStr(requestsSheet.Cells(requestRow, COL_PURPOSE).Value), " / ")
    descText = AppendToken(descText, CStr(requestsSheet.Cells(requestRow, COL_NOTE).Value), " / ")
    descText = AppendToken(descText, CStr(requestsSheet.Cells(requestRow, COL_VALIDATION_STATUS).Value), " / ")
    SecuiDescription = Left$(CleanSecuiText(descText), 255)
End Function

Private Function AppendToken(ByVal baseText As String, ByVal tokenText As String, ByVal delimiter As String) As String
    Dim t As String
    t = Trim$(CStr(tokenText))
    If Len(t) = 0 Then
        AppendToken = baseText
    ElseIf Len(baseText) = 0 Then
        AppendToken = t
    Else
        AppendToken = baseText & delimiter & t
    End If
End Function

Private Function CleanSecuiText(ByVal value As String) As String
    Dim s As String
    s = Trim$(CStr(value))
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanSecuiText = s
End Function

Private Function MergeFolderFiles(ByVal folderPath As String, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByVal logSheet As Worksheet, ByRef nextRow As Long) As Long
    ' Recurse subfolders so each team folder (e.g. 정보보호센터_1234) is scanned.
    Dim mergedCount As Long
    Dim folderName As String
    folderName = FolderLeafName(folderPath)
    mergedCount = MergeFolderTree(folderPath, folderName, requestsSheet, firewallsSheet, logSheet, nextRow)
    MergeFolderFiles = mergedCount
End Function

Private Function MergeFolderTree(ByVal folderPath As String, ByVal folderName As String, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByVal logSheet As Worksheet, ByRef nextRow As Long) As Long
    Dim fileName As String
    Dim mergedCount As Long
    Dim subFolders As Collection
    Dim subName As Variant

    ' 1) merge Excel files directly in this folder
    fileName = Dir(folderPath & Application.PathSeparator & "*.xls*")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            If StrComp(folderPath & Application.PathSeparator & fileName, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                mergedCount = mergedCount + MergeWorkbookFile(folderPath & Application.PathSeparator & fileName, fileName, folderName, requestsSheet, firewallsSheet, logSheet, nextRow)
            End If
        End If
        fileName = Dir()
    Loop

    ' 2) collect subfolder names first (Dir state is reused by MergeWorkbookFile)
    Set subFolders = CollectSubFolders(folderPath)
    For Each subName In subFolders
        ' Carry the request-folder context downward: a nested attachment folder
        ' (e.g. .../인프라시너지셀_2026-782_제목/첨부파일/) keeps the parsed
        ' team_doc_title name instead of being parsed as '첨부파일'.
        Dim childContext As String
        childContext = ChildFolderContext(folderName, CStr(subName))
        mergedCount = mergedCount + MergeFolderTree(folderPath & Application.PathSeparator & CStr(subName), childContext, requestsSheet, firewallsSheet, logSheet, nextRow)
    Next subName

    MergeFolderTree = mergedCount
End Function

Private Function ChildFolderContext(ByVal parentContext As String, ByVal childName As String) As String
    ' If the child folder name itself looks like a request folder (has a non-empty
    ' doc number after the first underscore), it becomes the new context; else the
    ' parent's request-folder context is inherited.
    Dim t As String, d As String, ti As String
    ParseRequestFolderName childName, t, d, ti
    If Len(d) > 0 Then
        ChildFolderContext = childName
    Else
        ChildFolderContext = parentContext
    End If
End Function

Private Function CollectSubFolders(ByVal folderPath As String) As Collection
    Dim result As Collection
    Dim entry As String
    Set result = New Collection
    entry = Dir(folderPath & Application.PathSeparator & "*", vbDirectory)
    Do While Len(entry) > 0
        If entry <> "." And entry <> ".." Then
            If (GetAttr(folderPath & Application.PathSeparator & entry) And vbDirectory) = vbDirectory Then
                result.Add entry
            End If
        End If
        entry = Dir()
    Loop
    Set CollectSubFolders = result
End Function

Private Function FolderLeafName(ByVal folderPath As String) As String
    Dim p As String
    p = folderPath
    Do While Len(p) > 0 And Right$(p, 1) = Application.PathSeparator
        p = Left$(p, Len(p) - 1)
    Loop
    Dim sepPos As Long
    sepPos = InStrRev(p, Application.PathSeparator)
    If sepPos > 0 Then
        FolderLeafName = Mid$(p, sepPos + 1)
    Else
        FolderLeafName = p
    End If
End Function

Private Sub ParseRequestFolderName(ByVal folderName As String, ByRef team As String, ByRef docNo As String, ByRef title As String)
    ' 인프라시너지셀_2026-782_제목 -> team=인프라시너지셀, docNo=2026-782, title=제목
    ' Split on the FIRST two underscores; title keeps any later underscores.
    Dim s As String
    s = Trim$(folderName)
    team = "" : docNo = "" : title = ""
    If Len(s) = 0 Then Exit Sub
    Dim first As Long, second As Long, rest As String
    first = InStr(s, "_")
    If first = 0 Then
        team = Trim$(s) : Exit Sub
    End If
    team = Trim$(Left$(s, first - 1))
    rest = Mid$(s, first + 1)
    second = InStr(rest, "_")
    If second = 0 Then
        docNo = Trim$(rest) : Exit Sub
    End If
    docNo = Trim$(Left$(rest, second - 1))
    title = Trim$(Mid$(rest, second + 1))
End Sub

Private Function MergeWorkbookFile(ByVal filePath As String, ByVal sourceFileName As String, ByVal folderName As String, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByVal logSheet As Worksheet, ByRef nextRow As Long) As Long
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim headerMap As Object
    Dim headerRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim mergedCount As Long
    Dim firstOutputRow As Long

    firstOutputRow = nextRow

    On Error GoTo OpenFailed
    Set sourceBook = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set sourceSheet = SelectRequestSheet(sourceBook, mParseSheetName)
    headerRow = FindHeaderRow(sourceSheet)
    Set headerMap = BuildHeaderMap(sourceSheet, headerRow)
    ValidateRequiredHeaders headerMap, sourceFileName

    lastRow = SourceLastRow(sourceSheet, headerMap)
    For rowIndex = headerRow + 1 To lastRow
        If RequestSourceRowHasData(sourceSheet, rowIndex, headerMap) Then
            ' CopyRequestRow explodes one source row into N rules and advances nextRow.
            mergedCount = mergedCount + CopyRequestRow(sourceSheet, rowIndex, headerMap, requestsSheet, firewallsSheet, nextRow, sourceFileName, folderName)
        End If
    Next rowIndex

    ' Vertically merge the identity cells (요청부서/요청번호/제목) for this file's
    ' contiguous output block so same-document rows are visually grouped.
    MergeIdentityBlock requestsSheet, firstOutputRow, nextRow - 1

    sourceBook.Close SaveChanges:=False
    AppendProcessingLog logSheet, sourceFileName, "OK", mergedCount, "헤더 " & CStr(headerRow) & "행, " & OutputRowMessage(firstOutputRow, nextRow)
    MergeWorkbookFile = mergedCount
    Exit Function

OpenFailed:
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    AppendProcessingLog logSheet, sourceFileName, "ERROR", 0, Err.Description
    MsgBox "파일을 처리할 수 없습니다: " & filePath & vbCrLf & Err.Description, vbExclamation
    MergeWorkbookFile = 0
End Function

Private Function CopyRequestRow(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByRef targetRow As Long, ByVal sourceFileName As String, ByVal folderName As String) As Long
    ' Explode one source request into the full 출발지IP × 목적지IP × 포트 product:
    ' one firewall rule per output row. Mirrors request_parser_oracle.explode_request_row.
    Dim srcs() As String, dsts() As String, ports() As String
    srcs = SplitNormalizedList(NormalizeListCell(ReadOpt(sourceSheet, sourceRow, headerMap, "출발지ip")))
    dsts = SplitNormalizedList(NormalizeListCell(ReadOpt(sourceSheet, sourceRow, headerMap, "목적지ip")))
    ports = SplitNormalizedList(NormalizeListCell(ReadOpt(sourceSheet, sourceRow, headerMap, "포트")))

    Dim targetFirewalls As String
    Dim srcName As String, dstName As String, proto As String, direction As String
    Dim purpose As String, startDate As String, endDate As String, note As String
    targetFirewalls = NormalizeListCell(ReadOpt(sourceSheet, sourceRow, headerMap, "대상방화벽"))
    srcName = NormalizeTextCell(ReadOpt(sourceSheet, sourceRow, headerMap, "출발지"))
    dstName = NormalizeTextCell(ReadOpt(sourceSheet, sourceRow, headerMap, "목적지"))
    proto = UCase$(NormalizeListCell(ReadOpt(sourceSheet, sourceRow, headerMap, "프로토콜")))
    direction = NormalizeTextCell(ReadOpt(sourceSheet, sourceRow, headerMap, "방향"))
    purpose = NormalizeTextCell(ReadOpt(sourceSheet, sourceRow, headerMap, "용도"))
    startDate = FormatMetadataDate(ReadOpt(sourceSheet, sourceRow, headerMap, "시작일"))
    endDate = FormatMetadataDate(ReadOpt(sourceSheet, sourceRow, headerMap, "종료일"))
    note = NormalizeTextCell(ReadOpt(sourceSheet, sourceRow, headerMap, "비고"))

    Dim reqTeam As String, reqDocNo As String, reqTitle As String
    ParseRequestFolderName folderName, reqTeam, reqDocNo, reqTitle

    Dim i As Long, j As Long, k As Long, written As Long
    For i = LBound(srcs) To UBound(srcs)
        For j = LBound(dsts) To UBound(dsts)
            For k = LBound(ports) To UBound(ports)
                requestsSheet.Cells(targetRow, COL_REQUEST_TEAM).Value = reqTeam
                requestsSheet.Cells(targetRow, COL_REQUEST_DOC_NO).Value = reqDocNo
                requestsSheet.Cells(targetRow, COL_REQUEST_TITLE).Value = reqTitle
                requestsSheet.Cells(targetRow, COL_SOURCE_FILE).Value = sourceFileName
                requestsSheet.Cells(targetRow, COL_SOURCE_ROW).Value = sourceRow
                requestsSheet.Cells(targetRow, COL_TARGET_FIREWALLS).Value = targetFirewalls
                requestsSheet.Cells(targetRow, COL_SOURCE_IP).Value = srcs(i)
                requestsSheet.Cells(targetRow, COL_SOURCE_NAME).Value = srcName
                requestsSheet.Cells(targetRow, COL_DESTINATION_IP).Value = dsts(j)
                requestsSheet.Cells(targetRow, COL_DESTINATION_NAME).Value = dstName
                requestsSheet.Cells(targetRow, COL_PROTOCOL).Value = proto
                requestsSheet.Cells(targetRow, COL_PORT).Value = ports(k)
                requestsSheet.Cells(targetRow, COL_DIRECTION).Value = direction
                requestsSheet.Cells(targetRow, COL_PURPOSE).Value = purpose
                requestsSheet.Cells(targetRow, COL_START_DATE).Value = startDate
                requestsSheet.Cells(targetRow, COL_END_DATE).Value = endDate
                requestsSheet.Cells(targetRow, COL_NOTE).Value = note
                requestsSheet.Cells(targetRow, COL_REQUEST_FOLDER).Value = folderName
                targetRow = targetRow + 1
                written = written + 1
            Next k
        Next j
    Next i
    CopyRequestRow = written
End Function

' Split a ';'-normalized list into an array; an empty value yields a single
' blank element so a missing field still produces one row. Mirrors
' request_parser_oracle.split_list.
Private Function SplitNormalizedList(ByVal normalized As String) As String()
    Dim s As String
    s = TrimChars(Trim$(CStr(normalized)), ";")
    If Len(s) = 0 Then
        Dim blank(0 To 0) As String
        blank(0) = ""
        SplitNormalizedList = blank
        Exit Function
    End If
    SplitNormalizedList = Split(s, ";")
End Function

' Vertically merge the identity cells (요청부서/요청번호/제목) across a contiguous
' output block [firstRow..lastRow] so same-document rows read as one group.
' IP/포트 cells are NEVER merged (route + duplicate detection need per-row values).
Private Sub MergeIdentityBlock(ByVal worksheet As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If lastRow <= firstRow Then Exit Sub
    Dim cols As Variant, c As Variant
    cols = Array(COL_REQUEST_TEAM, COL_REQUEST_DOC_NO, COL_REQUEST_TITLE)
    Application.DisplayAlerts = False
    For Each c In cols
        With worksheet.Range(worksheet.Cells(firstRow, CLng(c)), worksheet.Cells(lastRow, CLng(c)))
            .Merge
            .VerticalAlignment = xlCenter
        End With
    Next c
    Application.DisplayAlerts = True
End Sub

' True if a canonical name is a recognized FIELD header (everything but 'no').
Private Function IsFieldHeader(ByVal canon As String) As Boolean
    Select Case canon
        Case "대상방화벽", "출발지ip", "출발지", "목적지ip", "목적지", "프로토콜", _
             "포트", "방향", "용도", "시작일", "종료일", "비고"
            IsFieldHeader = True
    End Select
End Function

' Locate the header row by HEADER CONTENT (not by requiring a 'No' cell). Picks
' the row with the most recognized field headers, requiring an IP column, with a
' 'No' cell only as a tie-breaker. Mirrors tests/request_parser_oracle.py.
Private Function FindHeaderRow(ByVal worksheet As Worksheet) As Long
    Dim bestRow As Long
    bestRow = BestHeaderRow(worksheet)
    If bestRow > 0 Then
        FindHeaderRow = bestRow
        Exit Function
    End If
    Err.Raise vbObjectError + 1003, , "헤더 행을 찾을 수 없습니다: 출발지IP/목적지IP 열이 있는 행이 필요합니다."
End Function

' Non-raising header scan: returns the best header row, or 0 if none qualifies.
' Shared by FindHeaderRow (raises) and FindRequestSheet (compares sheets).
Private Function BestHeaderRow(ByVal worksheet As Worksheet) As Long
    Dim rowIndex As Long, columnIndex As Long, lastColumn As Long
    Dim key As String, canon As String
    Dim bestRow As Long, bestCount As Long, bestHasNo As Long
    bestRow = 0: bestCount = -1: bestHasNo = -1

    For rowIndex = 1 To 30
        lastColumn = worksheet.Cells(rowIndex, worksheet.Columns.Count).End(xlToLeft).Column
        Dim seen As Object
        Set seen = CreateObject("Scripting.Dictionary")
        Dim hasNo As Boolean, hasIp As Boolean
        hasNo = False: hasIp = False
        For columnIndex = 1 To lastColumn
            key = HeaderKey(CStr(worksheet.Cells(rowIndex, columnIndex).Value))
            If Len(key) > 0 Then
                canon = CanonicalHeaderName(key)
                If canon = key Then canon = UserAliasCanonical(key)
                If canon = "no" Then
                    hasNo = True
                ElseIf IsFieldHeader(canon) Then
                    If Not seen.Exists(canon) Then seen(canon) = True
                    If canon = "출발지ip" Or canon = "목적지ip" Then hasIp = True
                End If
            End If
        Next columnIndex

        If hasIp Then
            Dim fieldCount As Long
            fieldCount = seen.Count
            If fieldCount >= 2 Or hasNo Then
                Dim hasNoN As Long
                hasNoN = IIf(hasNo, 1, 0)
                If fieldCount > bestCount Or (fieldCount = bestCount And hasNoN > bestHasNo) Then
                    bestCount = fieldCount
                    bestHasNo = hasNoN
                    bestRow = rowIndex
                End If
            End If
        End If
    Next rowIndex

    BestHeaderRow = bestRow
End Function

' Score how header-like a sheet's best row is, for cross-sheet comparison.
' Returns -1 when the sheet has no qualifying header row; otherwise the field
' count of its best header row. Mirrors tests/request_parser_oracle.py.
Private Function SheetHeaderScore(ByVal worksheet As Worksheet) As Long
    Dim rowIndex As Long, columnIndex As Long, lastColumn As Long
    Dim key As String, canon As String
    Dim bestCount As Long
    bestCount = -1

    For rowIndex = 1 To 30
        lastColumn = worksheet.Cells(rowIndex, worksheet.Columns.Count).End(xlToLeft).Column
        Dim seen As Object
        Set seen = CreateObject("Scripting.Dictionary")
        Dim hasNo As Boolean, hasIp As Boolean
        hasNo = False: hasIp = False
        For columnIndex = 1 To lastColumn
            key = HeaderKey(CStr(worksheet.Cells(rowIndex, columnIndex).Value))
            If Len(key) > 0 Then
                canon = CanonicalHeaderName(key)
                If canon = key Then canon = UserAliasCanonical(key)
                If canon = "no" Then
                    hasNo = True
                ElseIf IsFieldHeader(canon) Then
                    If Not seen.Exists(canon) Then seen(canon) = True
                    If canon = "출발지ip" Or canon = "목적지ip" Then hasIp = True
                End If
            End If
        Next columnIndex
        If hasIp Then
            Dim fieldCount As Long
            fieldCount = seen.Count
            If fieldCount >= 2 Or hasNo Then
                If fieldCount > bestCount Then bestCount = fieldCount
            End If
        End If
    Next rowIndex

    SheetHeaderScore = bestCount
End Function

' Pick the sheet whose best header row scores highest. Ties keep the leftmost
' (lowest index) sheet so single-sheet forms behave exactly as before. Raises
' the same 1003 error when NO sheet has a recognizable header.
Private Function FindRequestSheet(ByVal sourceBook As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim bestSheet As Worksheet
    Dim bestScore As Long, score As Long
    bestScore = -1
    For Each ws In sourceBook.Worksheets
        score = SheetHeaderScore(ws)
        If score > bestScore Then
            bestScore = score
            Set bestSheet = ws
        End If
    Next ws
    If bestScore >= 0 And Not bestSheet Is Nothing Then
        Set FindRequestSheet = bestSheet
        Exit Function
    End If
    Err.Raise vbObjectError + 1003, , "헤더 행을 찾을 수 없습니다: 출발지IP/목적지IP 열이 있는 시트가 필요합니다."
End Function

' Honor an explicit settings.parse_sheet choice. Mirrors oracle select_request_sheet:
'   blank -> auto-detect (FindRequestSheet, leftmost tie / highest score)
'   non-empty -> EXACT Worksheet.Name match (no trim/case-fold); raise if absent;
'   raise if the named sheet has no recognizable header (never silently fall back).
Private Function SelectRequestSheet(ByVal sourceBook As Workbook, ByVal parseSheetName As String) As Worksheet
    Dim ws As Worksheet
    If Len(Trim$(parseSheetName)) = 0 Then
        Set SelectRequestSheet = FindRequestSheet(sourceBook)
        Exit Function
    End If
    For Each ws In sourceBook.Worksheets
        If ws.Name = parseSheetName Then
            If SheetHeaderScore(ws) < 0 Then
                Err.Raise vbObjectError + 1003, , "파싱 대상 시트에서 헤더 행을 찾을 수 없습니다: " & parseSheetName
            End If
            Set SelectRequestSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 1004, , "파싱 대상 시트를 찾을 수 없습니다: " & parseSheetName
End Function

Private Function BuildHeaderMap(ByVal worksheet As Worksheet, ByVal headerRow As Long) As Object
    Dim headerMap As Object
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerName As String

    Set headerMap = CreateObject("Scripting.Dictionary")
    lastColumn = worksheet.Cells(headerRow, worksheet.Columns.Count).End(xlToLeft).Column
    For columnIndex = 1 To lastColumn
        Dim rawKey As String
        rawKey = HeaderKey(CStr(worksheet.Cells(headerRow, columnIndex).Value))
        headerName = CanonicalHeaderName(rawKey)
        If headerName = rawKey Then headerName = UserAliasCanonical(rawKey)
        If Len(headerName) > 0 Then headerMap(headerName) = columnIndex
    Next columnIndex
    Set BuildHeaderMap = headerMap
End Function

Private Function CanonicalHeaderName(ByVal headerName As String) As String
    Select Case headerName
        Case "no", "번호", "순번", "연번", "seq", "순서", "항번", "일련번호", "번", "#": CanonicalHeaderName = "no"
        Case "대상방화벽", "대상fw", "targetfirewalls", "targetfirewall", "firewall", "firewalls", "fw", "장비명", "방화벽명": CanonicalHeaderName = "대상방화벽"
        Case "출발지ip", "출발ip", "sourceip", "source", "srcip", "src", "출발지주소", "송신ip", "원본ip": CanonicalHeaderName = "출발지ip"
        Case "출발지", "출발지명", "출발", "srcname", "출발지설명", "출발지ip설명", "출발지ip설멸", "출발지설멸", "출발ip설명", "출발ip설멸", "출발지내용", "송신자", "src설명": CanonicalHeaderName = "출발지"
        Case "목적지ip", "목적ip", "destinationip", "destination", "destiation", "dstip", "dst", "목적지주소", "수신ip": CanonicalHeaderName = "목적지ip"
        Case "목적지", "목적지명", "목적", "dstname", "목적지설명", "목적지ip설명", "목적지ip설멸", "목적지설멸", "목적ip설명", "목적ip설멸", "목적지내용", "수신자", "dst설명": CanonicalHeaderName = "목적지"
        Case "프로토콜", "protocol", "proto", "tcp/udp", "tcpudp", "프로토", "서비스", "프로토콜구분", "l4": CanonicalHeaderName = "프로토콜"
        Case "포트", "port", "dport", "목적지포트", "서비스포트", "포트번호", "dstport", "service": CanonicalHeaderName = "포트"
        Case "방향", "direction", "구분", "방향구분", "inout", "in/out", "송수신", "송수신구분": CanonicalHeaderName = "방향"
        Case "용도", "목적", "usage", "purpose", "사용용도", "신청사유", "설명": CanonicalHeaderName = "용도"
        Case "시작일", "시작", "startdate", "start", "시작일자", "시작날짜", "적용일", "적용시작일", "사용시작일": CanonicalHeaderName = "시작일"
        Case "종료일", "종료", "enddate", "end", "종료일자", "종료날짜", "만료일", "적용종료일", "사용종료일": CanonicalHeaderName = "종료일"
        Case "비고", "메모", "remark", "remarks", "note", "참고": CanonicalHeaderName = "비고"
        Case Else: CanonicalHeaderName = headerName
    End Select
End Function

Private Function UserAliasCanonical(ByVal rawKey As String) As String
    ' Look up a user-defined alias loaded from settings header_alias.
    If mUserAliases Is Nothing Then
        UserAliasCanonical = rawKey
        Exit Function
    End If
    If mUserAliases.Exists(rawKey) Then
        UserAliasCanonical = CStr(mUserAliases(rawKey))
    Else
        UserAliasCanonical = rawKey
    End If
End Function

Private Sub LoadUserAliases(ByVal settingsSheet As Worksheet)
    ' Parse settings header_alias value: "출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소"
    Set mUserAliases = CreateObject("Scripting.Dictionary")
    Dim raw As String
    raw = SettingsValue(settingsSheet, "header_alias")
    ' Parse the one-line settings.header_alias only if present...
    If Len(Trim$(raw)) > 0 Then
        Dim entries() As String, i As Long
        entries = Split(raw, ";")
        For i = LBound(entries) To UBound(entries)
            Dim entry As String
            entry = Trim$(entries(i))
            Dim eqPos As Long
            eqPos = InStr(entry, "=")
            If eqPos > 0 Then
                Dim canon As String
                canon = HeaderKey(Left$(entry, eqPos - 1))
                If Len(canon) > 0 Then
                    Dim aliasList() As String, j As Long
                    aliasList = Split(Mid$(entry, eqPos + 1), ",")
                    For j = LBound(aliasList) To UBound(aliasList)
                        Dim a As String
                        a = HeaderKey(aliasList(j))
                        If Len(a) > 0 Then mUserAliases(a) = canon
                    Next j
                End If
            End If
        Next i
    End If
    ' ...but ALWAYS read the header_aliases table sheet (the operator-friendly
    ' path), even when settings.header_alias is blank.
    LoadAliasSheet
End Sub

Private Sub LoadAliasSheet()
    ' header_aliases sheet: col1=standard (canonical or its built-in alias),
    ' col2=your_column (the operator's actual header). Easier than the settings
    ' one-line string — operators just add a row per non-standard column.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("header_aliases")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        Dim std As String, yourCol As String
        std = HeaderKey(CStr(ws.Cells(r, 1).Value))
        yourCol = HeaderKey(CStr(ws.Cells(r, 2).Value))
        If Len(std) > 0 And Len(yourCol) > 0 Then
            mUserAliases(yourCol) = CanonicalHeaderName(std)
        End If
    Next r
End Sub

Private Function SettingsValue(ByVal settingsSheet As Worksheet, ByVal key As String) As String
    Dim lastRow As Long, i As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If LCase$(Trim$(CStr(settingsSheet.Cells(i, 1).Value))) = LCase$(key) Then
            SettingsValue = Trim$(CStr(settingsSheet.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
End Function

Private Sub ShowInfo(ByVal messageText As String)
    If Not mSuppressMessages Then MsgBox messageText, vbInformation
End Sub

Private Sub ValidateRequiredHeaders(ByVal headerMap As Object, ByVal sourceFileName As String)
    Dim requiredHeaders As Variant
    Dim header As Variant

    ' Only the IP columns are truly required; the rest are optional metadata
    ' (read if present, blank if absent) so forms omitting 비고/용도/날짜 still merge.
    requiredHeaders = Array("출발지ip", "목적지ip")
    For Each header In requiredHeaders
        If Not headerMap.Exists(CStr(header)) Then Err.Raise vbObjectError + 1001, , sourceFileName & " 필수 컬럼 누락: " & CStr(header)
    Next header
End Sub



Private Function OutputRowMessage(ByVal firstOutputRow As Long, ByVal nextRow As Long) As String
    If nextRow <= firstOutputRow Then
        OutputRowMessage = "결과 행 없음"
    Else
        OutputRowMessage = "결과 " & CStr(firstOutputRow) & "-" & CStr(nextRow - 1) & "행"
    End If
End Function

Private Function SourceLastRow(ByVal sourceSheet As Worksheet, ByVal headerMap As Object) As Long
    ' Merge-aware last-row detection. Raw End(xlUp) on the IP columns can stop
    ' short when 출발지IP/목적지IP are vertically merged (Excel stores the value
    ' only in the merge top-left). So we start from the sheet's real used range
    ' and walk back to the last logical row, excluding pure merge-continuation
    ' rows whose values all come from top-left merged cells.
    Dim usedLast As Long
    Dim r As Long

    usedLast = sourceSheet.UsedRange.Row + sourceSheet.UsedRange.Rows.Count - 1
    For r = usedLast To 1 Step -1
        If RequestSourceRowHasData(sourceSheet, r, headerMap) Then
            SourceLastRow = r
            Exit Function
        End If
    Next r
    SourceLastRow = 1
End Function

Private Function RequestSourceRowHasData(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object) As Boolean
    Dim hasIp As Boolean
    hasIp = Len(Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("출발지ip"))))) > 0 Or _
        Len(Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("목적지ip"))))) > 0
    RequestSourceRowHasData = hasIp And RowHasOwnRequestData(sourceSheet, sourceRow, headerMap)
End Function

Private Function RowHasOwnRequestData(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object) As Boolean
    ' A lower row inside a vertically merged request block has visible values via
    ' ReadDataCell, but those values belong to the top-left row. Treat it as a
    ' continuation unless at least one request field has its own cell value.
    Dim key As Variant
    Dim cell As Range
    For Each key In headerMap.Keys
        If CStr(key) <> "no" Then
            Set cell = sourceSheet.Cells(sourceRow, CLng(headerMap(key)))
            If CellHasOwnRequestValue(cell) Then
                RowHasOwnRequestData = True
                Exit Function
            End If
        End If
    Next key
End Function

Private Function CellHasOwnRequestValue(ByVal cell As Range) As Boolean
    If cell.MergeCells Then
        If cell.MergeArea.Row <> cell.Row Or cell.MergeArea.Column <> cell.Column Then
            CellHasOwnRequestValue = False
            Exit Function
        End If
        CellHasOwnRequestValue = Len(Trim$(CStr(cell.MergeArea.Cells(1, 1).Value))) > 0
    Else
        CellHasOwnRequestValue = Len(Trim$(CStr(cell.Value))) > 0
    End If
End Function

Private Function ReadDataCell(ByVal sourceSheet As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    ' Read a request DATA cell, honoring merged ranges: a vertically/horizontally
    ' merged data cell exposes its value only at the merge top-left, so we read
    ' .MergeArea.Cells(1,1).Value. Header mapping stays raw (Option C).
    With sourceSheet.Cells(r, c)
        If .MergeCells Then
            ReadDataCell = .MergeArea.Cells(1, 1).Value
        Else
            ReadDataCell = .Value
        End If
    End With
End Function

Private Function ReadOpt(ByVal sourceSheet As Worksheet, ByVal r As Long, ByVal headerMap As Object, ByVal name As String) As Variant
    ' Read an OPTIONAL request column by canonical name: "" if the column is
    ' absent from headerMap (so missing metadata columns don't error out).
    If headerMap.Exists(name) Then
        ReadOpt = ReadDataCell(sourceSheet, r, CLng(headerMap(name)))
    Else
        ReadOpt = ""
    End If
End Function

' Normalize a LIST-like cell (IP/port/protocol): newlines/tabs/commas/space runs
' become a single ';' so two values stay distinct (e.g. "80" + LF + "443" ->
' "80;443"). Mirrors tests/request_parser_oracle.py _norm_list.
Private Function NormalizeListCell(ByVal value As Variant) As String
    Dim s As String
    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCrLf, ";")
    s = Replace(s, vbCr, ";")
    s = Replace(s, vbLf, ";")
    s = Replace(s, vbTab, ";")
    s = Replace(s, ",", ";")
    s = Replace(s, ChrW(65292), ";")   ' fullwidth comma
    s = Replace(s, ChrW(65307), ";")   ' fullwidth semicolon
    s = Replace(s, " ", ";")           ' list field: space is a delimiter
    s = CollapseChar(s, ";")
    NormalizeListCell = TrimChars(s, ";")
End Function

' Normalize a PROSE cell (name/purpose/note/direction): only newlines/tabs become
' '; ' so descriptions don't visually concatenate; internal spaces are preserved.
' Mirrors tests/request_parser_oracle.py _norm_text.
Private Function NormalizeTextCell(ByVal value As Variant) As String
    Dim s As String
    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCrLf, "; ")
    s = Replace(s, vbCr, "; ")
    s = Replace(s, vbLf, "; ")
    s = Replace(s, vbTab, "; ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Do While InStr(s, "; ; ") > 0
        s = Replace(s, "; ; ", "; ")
    Loop
    NormalizeTextCell = TrimChars(Trim$(s), "; ")
End Function

' Collapse runs of a single character to one occurrence.
Private Function CollapseChar(ByVal s As String, ByVal ch As String) As String
    Do While InStr(s, ch & ch) > 0
        s = Replace(s, ch & ch, ch)
    Loop
    CollapseChar = s
End Function

' Trim a leading/trailing delimiter token (single char or short string) from both ends.
Private Function TrimChars(ByVal s As String, ByVal token As String) As String
    Dim n As Long
    n = Len(token)
    Do While Len(s) >= n And Left$(s, n) = token
        s = Mid$(s, n + 1)
    Loop
    Do While Len(s) >= n And Right$(s, n) = token
        s = Left$(s, Len(s) - n)
    Loop
    TrimChars = s
End Function

Private Function FormatMetadataDate(ByVal value As Variant) As String
    ' 시작일/종료일: 실제 Date 셀은 yyyy-mm-dd 로 결정적 포맷(로케일 비의존).
    ' 문자열 날짜는 그대로 유지(VarType이 vbDate일 때만 포맷).
    ' 오라클 _format_metadata_date(isinstance datetime/date)와 동일.
    If VarType(value) = vbDate Then
        FormatMetadataDate = Format$(value, "yyyy-mm-dd")
    Else
        FormatMetadataDate = Trim$(CStr(value))
    End If
End Function

Private Sub WriteRowValidation(ByVal worksheet As Worksheet, ByVal rowIndex As Long)
    Dim messageText As String

    If Len(Trim$(CStr(worksheet.Cells(rowIndex, COL_SOURCE_IP).Value))) = 0 Then messageText = AppendMessageText(messageText, "출발지IP 비어 있음")
    If Len(Trim$(CStr(worksheet.Cells(rowIndex, COL_DESTINATION_IP).Value))) = 0 Then messageText = AppendMessageText(messageText, "목적지IP 비어 있음")
    If Len(Trim$(CStr(worksheet.Cells(rowIndex, COL_PROTOCOL).Value))) = 0 Then messageText = AppendMessageText(messageText, "프로토콜 비어 있음")
    If Len(Trim$(CStr(worksheet.Cells(rowIndex, COL_PORT).Value))) = 0 Then messageText = AppendMessageText(messageText, "포트 비어 있음")

    If Len(messageText) > 0 Then
        worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Value = "WARN"
        worksheet.Cells(rowIndex, COL_VALIDATION_MESSAGE).Value = messageText
        worksheet.Cells(rowIndex, COL_VALIDATION_MESSAGE).Interior.Color = RGB(255, 235, 156)
    Else
        worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Value = "OK"
    End If
End Sub

Private Sub AppendValidationMessage(ByVal worksheet As Worksheet, ByVal rowIndex As Long, ByVal statusText As String, ByVal messageText As String)
    Dim currentStatus As String
    Dim currentMessage As String

    currentStatus = Trim$(CStr(worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Value))
    currentMessage = Trim$(CStr(worksheet.Cells(rowIndex, COL_VALIDATION_MESSAGE).Value))

    If currentStatus = "OK" Or Len(currentStatus) = 0 Then
        worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Value = statusText
    ElseIf InStr(1, currentStatus, statusText, vbTextCompare) = 0 Then
        worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Value = currentStatus & ";" & statusText
    End If
    worksheet.Cells(rowIndex, COL_VALIDATION_MESSAGE).Value = AppendMessageText(currentMessage, messageText)
End Sub

Private Function AppendMessageText(ByVal currentText As String, ByVal newText As String) As String
    If Len(currentText) = 0 Then
        AppendMessageText = newText
    ElseIf Len(newText) = 0 Then
        AppendMessageText = currentText
    ElseIf InStr(1, currentText, newText, vbTextCompare) > 0 Then
        AppendMessageText = currentText
    Else
        AppendMessageText = currentText & "; " & newText
    End If
End Function

Private Sub WriteLogHeaders(ByVal worksheet As Worksheet)
    worksheet.Range("A1:E1").Value = Array("processed_at", "source_file", "status", "merged_rows", "message")
End Sub

Private Sub WriteSecuiBatchHeaders(ByVal worksheet As Worksheet)
    worksheet.Range("A1:T1").Value = Array("No", "장비명", "정책명", "출발지주소", "출발지명", "목적지주소", "목적지명", "서비스", "프로토콜", "목적지포트", "동작", "로그", "사용여부", "시작일", "종료일", "설명", "신청부서", "신청번호", "원본파일", "원본행")
End Sub

Private Sub WriteSecuiCliHeaders(ByVal worksheet As Worksheet)
    worksheet.Range("A1:I1").Value = Array("No", "장비명", "정책명", "명령어", "검토메모", "신청부서", "신청번호", "원본파일", "원본행")
End Sub

Private Sub WriteSecuiPolicyExportHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:I1").Value = Array("policy_id", "policy_name", "firewall_name", "source", "destination", "service", "action", "enabled", "comment")
        worksheet.Range("A2:I4").Value = Array( _
            Array("1001", "ALLOW_WEB_TO_DMZ", "SECUI-FW-01", "10.10.10.5", "10.20.20.5", "tcp/443", "allow", "Y", "기존 허용 정책"), _
            Array("1002", "DENY_DNS_TO_INTERNET", "SECUI-FW-03", "10.10.10.5", "8.8.8.8", "udp/53", "deny", "Y", "차단 검토 필요"), _
            Array("1003", "DISABLED_DMZ_TEST", "SECUI-FW-02", "10.10.10.5", "10.20.20.5", "tcp/8443", "allow", "N", "비활성 정책"))
    End If
    AnnotateSecuiPolicyExportHeaders worksheet
End Sub

Private Sub WritePolicyAnalysisHeaders(ByVal worksheet As Worksheet)
    worksheet.Range("A1:T1").Value = Array("판정", "요청번호", "대상방화벽", "출발지", "목적지", "서비스", "기존정책", "기존정책상태", "근거", "조치", "요청원본행", "정책원본행", "raw_source", "raw_destination", "raw_service", "normalized_source", "normalized_destination", "normalized_protocol", "normalized_port", "debug_note")
End Sub

Private Sub WritePolicySummarySheet(ByVal worksheet As Worksheet)
    worksheet.Range("A1:D1").Value = Array("구분", "건수", "검토 기준", "다음 조치")
    worksheet.Range("A2:D7").Value = Array( _
        Array("전체", "", "분석 대상 전체", "상태별 건수를 먼저 확인"), _
        Array("기존 허용", "", "기존 허용 정책이 신청을 커버", "중복 신청 생략 또는 근거 첨부"), _
        Array("기존 차단", "", "차단 정책과 일치", "방화벽 담당자 검토"), _
        Array("검토 필요", "", "부분 일치 또는 객체명 미해석", "SECUI 객체/서비스 확인"), _
        Array("비활성 일치", "", "비활성 정책과 일치", "활성화 가능 여부 검토"), _
        Array("기존 정책 없음", "", "일치 정책 없음", "신규 정책 생성 검토"))
    worksheet.Range("B2").Formula = "=COUNTA(policy_analysis!A2:A5000)"
    worksheet.Range("B3").Formula = "=COUNTIF(policy_analysis!A2:A5000,""EXISTING_ALLOW"")"
    worksheet.Range("B4").Formula = "=COUNTIF(policy_analysis!A2:A5000,""EXISTING_DENY"")"
    worksheet.Range("B5").Formula = "=COUNTIF(policy_analysis!A2:A5000,""PARTIAL_MATCH"")+COUNTIF(policy_analysis!A2:A5000,""OBJECT_UNRESOLVED"")"
    worksheet.Range("B6").Formula = "=COUNTIF(policy_analysis!A2:A5000,""DISABLED_MATCH"")"
    worksheet.Range("B7").Formula = "=COUNTIF(policy_analysis!A2:A5000,""NO_EXISTING_POLICY"")"
End Sub

Private Sub AnnotateSecuiPolicyExportHeaders(ByVal worksheet As Worksheet)
    Dim comments As Variant, c As Long
    comments = Array( _
        "필수: SECUI export 정책 ID를 붙여넣습니다.", _
        "필수: 기존 정책명을 붙여넣습니다.", _
        "필수: firewalls 시트의 방화벽명과 같은 이름을 붙여넣습니다.", _
        "필수: 출발지 주소/CIDR/ANY 또는 객체명을 붙여넣습니다.", _
        "필수: 목적지 주소/CIDR/ANY 또는 객체명을 붙여넣습니다.", _
        "필수: tcp/443, udp/53, ANY 같은 서비스 표기를 붙여넣습니다.", _
        "필수: allow/deny/drop/reject/accept/pass 중 하나를 붙여넣습니다.", _
        "필수: Y/N 또는 TRUE/FALSE 사용여부를 붙여넣습니다.", _
        "선택: SECUI export 설명이나 운영 메모를 붙여넣습니다.")
    For c = 1 To SECUI_EXPORT_LAST_COL
        worksheet.Cells(1, c).ClearComments
        worksheet.Cells(1, c).AddComment CStr(comments(c - 1))
    Next c
End Sub

Private Sub WriteVendorCliTemplateHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:E1").Value = Array("vendor", "template_name", "enabled", "command_template", "review_note")
        worksheet.Range("A2:E2").Value = Array( _
            "SECUI", _
            "default_allow_srule", _
            "Y", _
            DefaultVendorCliTemplate(), _
            DefaultVendorCliReviewNote())
    End If
End Sub

Private Sub WriteServiceCatalogHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:E1").Value = Array("service_name", "protocol", "port", "secui_service", "description")
        worksheet.Range("A2:E8").Value = Array( _
            Array("HTTPS", "TCP", "443", "tcp/443", "웹 HTTPS"), _
            Array("HTTP", "TCP", "80", "tcp/80", "웹 HTTP"), _
            Array("SSH", "TCP", "22", "tcp/22", "SSH 관리"), _
            Array("DNS-UDP", "UDP", "53", "udp/53", "DNS 조회"), _
            Array("DNS-TCP", "TCP", "53", "tcp/53", "DNS zone transfer 등"), _
            Array("ICMP", "ICMP", "", "icmp/", "ICMP/Ping"), _
            Array("CUSTOM", "TCP", "직접입력", "tcp/<port>", "목록에 없는 서비스는 포트 칸에 직접 입력"))
    End If
End Sub

Private Sub AppendProcessingLog(ByVal worksheet As Worksheet, ByVal sourceFileName As String, ByVal statusText As String, ByVal mergedRows As Long, ByVal messageText As String)
    Dim nextRow As Long

    nextRow = worksheet.Cells(worksheet.Rows.Count, 1).End(xlUp).Row + 1
    worksheet.Cells(nextRow, 1).Value = Now
    worksheet.Cells(nextRow, 2).Value = sourceFileName
    worksheet.Cells(nextRow, 3).Value = statusText
    worksheet.Cells(nextRow, 4).Value = mergedRows
    worksheet.Cells(nextRow, 5).Value = messageText
End Sub

Private Sub FormatLogSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:E").AutoFit
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:E1").AutoFilter
End Sub

Private Sub FormatSecuiBatchSheet(ByVal worksheet As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(6, 18, 36, 18, 18, 18, 18, 16, 10, 12, 10, 8, 10, 12, 12, 42, 16, 14, 24, 10)
    For c = 1 To SECUI_LAST_COL
        worksheet.Columns(c).ColumnWidth = widths(c - 1)
    Next c
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:T1").AutoFilter
End Sub

Private Sub FormatSecuiCliSheet(ByVal worksheet As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(6, 18, 36, 120, 60, 16, 14, 24, 10)
    For c = 1 To SECUI_CLI_LAST_COL
        worksheet.Columns(c).ColumnWidth = widths(c - 1)
    Next c
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:I1").AutoFilter
End Sub

Private Sub FormatPolicyAnalysisSheet(ByVal worksheet As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(18, 12, 24, 18, 18, 14, 28, 14, 42, 28, 10, 10, 22, 22, 16, 22, 22, 12, 12, 36)
    For c = 1 To POLICY_ANALYSIS_LAST_COL
        worksheet.Columns(c).ColumnWidth = widths(c - 1)
    Next c
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:T1").AutoFilter
    worksheet.Range("K:T").EntireColumn.Hidden = True
End Sub

Private Sub FormatPolicySummarySheet(ByVal worksheet As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(16, 10, 34, 34)
    For c = 1 To POLICY_SUMMARY_LAST_COL
        worksheet.Columns(c).ColumnWidth = widths(c - 1)
    Next c
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:D7").AutoFilter
End Sub

Private Sub MarkDuplicateRequests(ByVal worksheet As Worksheet)
    Dim seen As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim duplicateKey As String

    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = worksheet.Cells(worksheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    For rowIndex = REQ_DATA_START_ROW To lastRow
        duplicateKey = RequestDuplicateKey(worksheet, rowIndex)
        If Len(duplicateKey) > 0 Then
            If seen.Exists(duplicateKey) Then
                HighlightDuplicateRow worksheet, rowIndex
                HighlightDuplicateRow worksheet, CLng(seen(duplicateKey))
                AppendValidationMessage worksheet, rowIndex, "DUPLICATE", "중복 후보: " & CStr(seen(duplicateKey)) & "행"
                AppendValidationMessage worksheet, CLng(seen(duplicateKey)), "DUPLICATE", "중복 후보: " & CStr(rowIndex) & "행"
            Else
                seen.Add duplicateKey, rowIndex
            End If
        End If
    Next rowIndex
End Sub

Private Sub HighlightDuplicateRow(ByVal worksheet As Worksheet, ByVal rowIndex As Long)
    Dim statusColor As Variant
    statusColor = worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Interior.Color
    worksheet.Rows(rowIndex).Interior.Color = RGB(255, 230, 153)
    worksheet.Cells(rowIndex, COL_VALIDATION_STATUS).Interior.Color = statusColor
End Sub

Private Function RequestDuplicateKey(ByVal worksheet As Worksheet, ByVal rowIndex As Long) As String
    RequestDuplicateKey = Trim$(CStr(worksheet.Cells(rowIndex, COL_SOURCE_IP).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_DESTINATION_IP).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PROTOCOL).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PORT).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_DIRECTION).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PURPOSE).Value))
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Sub ApplyOperatorSheetVisibility()
    SetSheetVisibility REQUESTS_SHEET, xlSheetVisible
    SetSheetVisibility SETTINGS_SHEET, xlSheetVisible
    SetSheetVisibility FIREWALLS_SHEET, xlSheetVisible
    SetSheetVisibility SECUI_POLICY_EXPORT_SHEET, xlSheetVisible
    SetSheetVisibility POLICY_SUMMARY_SHEET, xlSheetVisible
    SetSheetVisibility POLICY_ANALYSIS_SHEET, xlSheetVisible
    SetSheetVisibility SECUI_BATCH_SHEET, xlSheetVisible
    SetSheetVisibility SECUI_CLI_SHEET, xlSheetVisible
    SetSheetVisibility VENDOR_CLI_TEMPLATE_SHEET, xlSheetVisible

    ThisWorkbook.Worksheets(REQUESTS_SHEET).Activate
    SetSheetVisibility "usage", xlSheetVisible
    SetSheetVisibility "header_aliases", xlSheetHidden
    SetSheetVisibility FIREWALL_RANGE_SHEET, xlSheetHidden
    SetSheetVisibility LOG_SHEET, xlSheetHidden
    SetSheetVisibility SERVICE_CATALOG_SHEET, xlSheetHidden
    SetSheetVisibility "sample-request-format", xlSheetHidden
End Sub

Private Sub SetSheetVisibility(ByVal sheetName As String, ByVal visibility As XlSheetVisibility)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Visible = visibility
    On Error GoTo 0
End Sub

Private Sub WriteRequestHeaders(ByVal worksheet As Worksheet)
    ' Row 1: cosmetic group labels (출발지 over IP+설명, 목적지 over IP+설명).
    ' Row 2: canonical leaf headers (25). Data starts at row 3.
    worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_SOURCE_IP).Value = "출발지"
    worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_DESTINATION_IP).Value = "목적지"
    Application.DisplayAlerts = False
    worksheet.Range(worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_SOURCE_IP), worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_SOURCE_NAME)).Merge
    worksheet.Range(worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_DESTINATION_IP), worksheet.Cells(REQ_HEADER_GROUP_ROW, COL_DESTINATION_NAME)).Merge
    Application.DisplayAlerts = True
    worksheet.Range("A" & REQ_HEADER_ROW & ":Y" & REQ_HEADER_ROW).Value = Array("요청부서", "요청번호", "제목", "원본파일", "원본행", "검증상태", "대상방화벽", "출발지IP", "출발지설명", "목적지IP", "목적지설명", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고", "검증메시지", "방화벽경로", "출발매칭대역", "목적매칭대역", "대역경로", "매칭근거", "요청폴더")
End Sub

Private Sub WriteFirewallHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:D1").Value = Array("firewall_name", "vendor", "enabled", "comment")
        worksheet.Range("A2:D4").Value = Array( _
            Array("SECUI-FW-01", "SECUI", "Y", "내부-서버 구간"), _
            Array("SECUI-FW-02", "SECUI", "Y", "서버-DMZ 구간"), _
            Array("SECUI-FW-03", "SECUI", "Y", "DMZ-외부 구간"))
    End If
End Sub

Private Sub WriteSettings(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:C1").Value = Array("key", "value", "설명")
        worksheet.Range("A2:C2").Value = Array("request_folder", "", "신청서 엑셀이 모여 있는 폴더 경로. 하위 폴더(예: 정보보호센터_1234)까지 재귀 탐색합니다.")
        worksheet.Range("A3:C3").Value = Array("parse_sheet", "", "파싱할 시트 이름(정확히 일치). 비워두면 헤더로 자동 감지합니다.")
        worksheet.Range("A4:C4").Value = Array("header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소")
    Else
        ' Upgrade an already-seeded settings sheet: add parse_sheet if it predates
        ' this version, without disturbing existing rows/values.
        EnsureSettingRow worksheet, "parse_sheet", "", "파싱할 시트 이름(정확히 일치). 비워두면 헤더로 자동 감지합니다."
    End If
End Sub

' Append a settings key (with value + 설명) only if it is not already present.
' Case-insensitive key match, mirroring SettingsValue lookups.
Private Sub EnsureSettingRow(ByVal worksheet As Worksheet, ByVal key As String, ByVal value As String, ByVal note As String)
    Dim lastRow As Long, i As Long
    lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If LCase$(Trim$(CStr(worksheet.Cells(i, 1).Value))) = LCase$(key) Then Exit Sub
    Next i
    worksheet.Range("A" & (lastRow + 1) & ":C" & (lastRow + 1)).Value = Array(key, value, note)
End Sub

Private Sub FormatRequestsSheet(ByVal worksheet As Worksheet)
    Dim widths As Variant, c As Long
    widths = Array(16, 12, 20, 28, 8, 18, 32, 18, 16, 18, 16, 10, 14, 10, 28, 12, 12, 24, 40, 34, 18, 18, 30, 60, 24)
    For c = 1 To REQ_LAST_COL
        worksheet.Columns(c).ColumnWidth = widths(c - 1)
    Next c
    worksheet.Rows(REQ_HEADER_GROUP_ROW).Font.Bold = True
    worksheet.Rows(REQ_HEADER_ROW).Font.Bold = True
    worksheet.Range("A" & REQ_HEADER_ROW & ":Y" & REQ_HEADER_ROW).AutoFilter
End Sub

Private Function HeaderKey(ByVal headerText As String) As String
    Dim k As String, original As String
    k = LCase$(Replace(Trim$(headerText), " ", ""))
    original = k
    ' strip decorating punctuation (No. / No# / (No) / No:)
    Dim puncts As Variant, p As Variant
    puncts = Array(".", "#", ":", "/", "(", ")", "[", "]", "-", ChrW(65294), ChrW(12290))
    Dim changed As Boolean
    Do
        changed = False
        For Each p In puncts
            If Len(k) > 0 Then
                If Left$(k, 1) = p Then k = Mid$(k, 2): changed = True
            End If
            If Len(k) > 0 Then
                If Right$(k, 1) = p Then k = Left$(k, Len(k) - 1): changed = True
            End If
        Next p
    Loop While changed
    ' a token that is ALL punctuation (e.g. '#') is kept so it can match a
    ' symbolic alias instead of collapsing to an empty (ignored) key.
    If Len(k) = 0 Then
        HeaderKey = original
    Else
        HeaderKey = k
    End If
End Function

Private Sub FormatFirewallsSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:D").AutoFit
    worksheet.Rows(1).Font.Bold = True
End Sub

Private Sub WriteFirewallRangeHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:K1").Value = Array("firewall_name", "source_cidr", "destination_cidr", "direction", "path_order", "enabled", "comment", "source_interface", "destination_interface", "source_zone", "destination_zone")
        worksheet.Range("A2:K7").Value = Array( _
            Array("SECUI-FW-01", "10.10.0.0/16", "172.16.0.0/16", "OUT", 10, "Y", "업무PC -> 서버", "inside", "server", "INTERNAL", "SERVER"), _
            Array("SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10, "Y", "업무PC -> DMZ", "inside", "dmz", "INTERNAL", "DMZ"), _
            Array("SECUI-FW-02", "10.10.0.0/16", "10.20.0.0/16", "OUT", 20, "Y", "업무PC -> DMZ", "server", "dmz", "SERVER", "DMZ"), _
            Array("SECUI-FW-01", "10.10.0.0/16", "8.8.8.0/24", "OUT", 10, "Y", "업무PC -> 외부 DNS", "inside", "outside", "INTERNAL", "EXTERNAL"), _
            Array("SECUI-FW-02", "10.10.0.0/16", "8.8.8.0/24", "OUT", 20, "Y", "업무PC -> 외부 DNS", "server", "outside", "SERVER", "EXTERNAL"), _
            Array("SECUI-FW-03", "10.10.0.0/16", "8.8.8.0/24", "OUT", 30, "Y", "업무PC -> 외부 DNS", "dmz", "outside", "DMZ", "EXTERNAL"))
    End If
End Sub

Private Sub FormatGenericSheet(ByVal worksheet As Worksheet, ByVal cols As String)
    worksheet.Columns(cols).AutoFit
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range(Replace(cols, ":", "1:") & "1").AutoFilter
End Sub

Private Function RequestFolderPath(ByVal settingsSheet As Worksheet) As String
    Dim configuredPath As String
    Dim selectedPath As String

    configuredPath = Trim$(CStr(settingsSheet.Range("B2").Value))
    If FolderExists(configuredPath) Then
        RequestFolderPath = configuredPath
        Exit Function
    End If

    selectedPath = PickFolder(configuredPath)
    If Len(selectedPath) > 0 Then
        settingsSheet.Range("B2").Value = selectedPath
        RequestFolderPath = selectedPath
    End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    FolderExists = Len(Dir(folderPath, vbDirectory)) > 0
End Function

Private Function PickFolder(Optional ByVal initialPath As String = "") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "방화벽 신청서 폴더 선택"
        If FolderExists(initialPath) Then .InitialFileName = initialPath & Application.PathSeparator
        If .Show <> -1 Then Exit Function
        PickFolder = .SelectedItems(1)
    End With
End Function
