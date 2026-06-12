Option Explicit

Private Const SETTINGS_SHEET As String = "settings"
Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const REQUESTS_SHEET As String = "requests"
Private Const SECUI_BATCH_SHEET As String = "secui_batch"
Private Const SECUI_CLI_SHEET As String = "secui_cli"
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
Private Const FW_COL_NAME As Long = 1
Private Const FW_COL_VENDOR As Long = 2
Private Const FW_COL_ENABLED As Long = 3

Private mUserAliases As Object
Private mParseSheetName As String

Public Sub SetupFirewallAutomationWorkbook()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim firewallRangeSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim logSheet As Worksheet
    Dim secuiBatchSheet As Worksheet
    Dim secuiCliSheet As Worksheet

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set firewallRangeSheet = EnsureSheet(FIREWALL_RANGE_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    Set logSheet = EnsureSheet(LOG_SHEET)
    Set secuiBatchSheet = EnsureSheet(SECUI_BATCH_SHEET)
    Set secuiCliSheet = EnsureSheet(SECUI_CLI_SHEET)

    WriteRequestHeaders requestsSheet
    WriteFirewallHeaders firewallsSheet
    WriteFirewallRangeHeaders firewallRangeSheet
    WriteSettings settingsSheet
    WriteLogHeaders logSheet
    WriteSecuiBatchHeaders secuiBatchSheet
    WriteSecuiCliHeaders secuiCliSheet
    FormatRequestsSheet requestsSheet
    FormatFirewallsSheet firewallsSheet
    FormatGenericSheet firewallRangeSheet, "A:G"
    FormatLogSheet logSheet
    FormatSecuiBatchSheet secuiBatchSheet
    FormatSecuiCliSheet secuiCliSheet

    MsgBox "방화벽 정책 자동화 시트 구성이 완료되었습니다.", vbInformation
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

    On Error GoTo RouteFailed
    Application.Run "FirewallRouteAnalysis.AnalyzeRequestRoutes"
    On Error GoTo 0
    GoTo RouteDone
RouteFailed:
    AppendProcessingLog logSheet, "(firewall range analysis)", "ERROR", 0, "방화벽 대역 분석 실패: " & Err.Description
    MsgBox "방화벽 대역 분석 중 오류가 발생했습니다: " & Err.Description, vbExclamation
    On Error GoTo 0
RouteDone:
    ' Append duplicate-candidate markers AFTER route analysis owns validation_*,
    ' so AppendValidationMessage merges onto the route status instead of being
    ' overwritten by WriteResultRow.
    MarkDuplicateRequests requestsSheet

    MsgBox CStr(mergedCount) & "건의 신청서를 통합했습니다.", vbInformation
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
    sampleSheet.Range("B1:M1").Value = Array("No", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고")
    sampleSheet.Range("B2:M2").Value = Array(1, "10.10.10.0/24", "업무PC", "172.16.1.10", "업무시스템", "TCP", "443", "IN", "HTTPS 업무 연동", "2026-01-01", "2026-12-31", "정기 신청")
    sampleSheet.Columns("A:M").AutoFit
    sampleSheet.Rows(1).Font.Bold = True
    sampleSheet.Range("B1:M1").AutoFilter

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
    MsgBox CStr(convertedRows) & "건의 SECUI 배치 행을 생성했습니다.", vbInformation
End Sub

Public Sub ConvertRequestsToSecuiCli()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim secuiCliSheet As Worksheet
    Dim secuiFirewalls As Object
    Dim requestRow As Long
    Dim cliRow As Long
    Dim lastRow As Long
    Dim convertedRows As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set secuiCliSheet = EnsureSheet(SECUI_CLI_SHEET)
    WriteFirewallHeaders firewallsSheet
    WriteSecuiCliHeaders secuiCliSheet
    Set secuiFirewalls = LoadSecuiFirewalls(firewallsSheet)
    secuiCliSheet.Rows("2:" & secuiCliSheet.Rows.Count).Clear

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    cliRow = 2
    For requestRow = REQ_DATA_START_ROW To lastRow
        convertedRows = convertedRows + CopySecuiCliRows(requestsSheet, secuiCliSheet, secuiFirewalls, requestRow, cliRow)
    Next requestRow

    FormatSecuiCliSheet secuiCliSheet
    MsgBox CStr(convertedRows) & "건의 SECUI CLI 명령 초안을 생성했습니다.", vbInformation
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

Private Function CopySecuiCliRows(ByVal requestsSheet As Worksheet, ByVal secuiCliSheet As Worksheet, ByVal secuiFirewalls As Object, ByVal requestRow As Long, ByRef cliRow As Long) As Long
    Dim targetFirewalls As Variant
    Dim firewallValue As Variant
    Dim firewallName As String
    Dim written As Long

    targetFirewalls = Split(Trim$(CStr(requestsSheet.Cells(requestRow, COL_TARGET_FIREWALLS).Value)), ";")
    For Each firewallValue In targetFirewalls
        firewallName = Trim$(CStr(firewallValue))
        If Len(firewallName) > 0 And secuiFirewalls.Exists(SecuiFirewallKey(firewallName)) Then
            WriteSecuiCliRow requestsSheet, secuiCliSheet, requestRow, cliRow, firewallName
            cliRow = cliRow + 1
            written = written + 1
        End If
    Next firewallValue

    CopySecuiCliRows = written
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

Private Sub WriteSecuiCliRow(ByVal requestsSheet As Worksheet, ByVal secuiCliSheet As Worksheet, ByVal requestRow As Long, ByVal cliRow As Long, ByVal firewallName As String)
    Dim policyName As String
    policyName = SecuiPolicyName(requestsSheet, requestRow, firewallName)

    secuiCliSheet.Cells(cliRow, 1).Value = cliRow - 1
    secuiCliSheet.Cells(cliRow, 2).Value = firewallName
    secuiCliSheet.Cells(cliRow, 3).Value = policyName
    secuiCliSheet.Cells(cliRow, 4).Value = SecuiCliCommand(requestsSheet, requestRow, firewallName, policyName)
    secuiCliSheet.Cells(cliRow, 5).Value = "장비 CLI에서 'fw set srule help'로 옵션명 확인 후 적용"
    secuiCliSheet.Cells(cliRow, 6).Value = requestsSheet.Cells(requestRow, COL_REQUEST_TEAM).Value
    secuiCliSheet.Cells(cliRow, 7).Value = requestsSheet.Cells(requestRow, COL_REQUEST_DOC_NO).Value
    secuiCliSheet.Cells(cliRow, 8).Value = requestsSheet.Cells(requestRow, COL_SOURCE_FILE).Value
    secuiCliSheet.Cells(cliRow, 9).Value = requestsSheet.Cells(requestRow, COL_SOURCE_ROW).Value
End Sub

Private Function SecuiCliCommand(ByVal requestsSheet As Worksheet, ByVal requestRow As Long, ByVal firewallName As String, ByVal policyName As String) As String
    Dim proto As String
    Dim portText As String
    Dim srcIp As String
    Dim dstIp As String
    Dim descText As String

    proto = LCase$(Trim$(CStr(requestsSheet.Cells(requestRow, COL_PROTOCOL).Value)))
    portText = Trim$(CStr(requestsSheet.Cells(requestRow, COL_PORT).Value))
    srcIp = Trim$(CStr(requestsSheet.Cells(requestRow, COL_SOURCE_IP).Value))
    dstIp = Trim$(CStr(requestsSheet.Cells(requestRow, COL_DESTINATION_IP).Value))
    descText = SecuiDescription(requestsSheet, requestRow)

    SecuiCliCommand = "fw set srule name " & SecuiCliQuote(policyName) & _
        " action allow src " & SecuiCliQuote(srcIp) & _
        " dst " & SecuiCliQuote(dstIp) & _
        " service " & SecuiCliQuote(proto & "/" & portText) & _
        " log enable enable yes description " & SecuiCliQuote(descText) & _
        " # device=" & CleanSecuiText(firewallName)
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

    Dim srcName As String, dstName As String, proto As String, direction As String
    Dim purpose As String, startDate As String, endDate As String, note As String
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
                WriteRowValidation requestsSheet, targetRow
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
        Case "출발지ip", "출발지", "목적지ip", "목적지", "프로토콜", _
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
    ' Highlight the duplicate row WITHOUT clobbering the route-owned
    ' validation_status cell color (column 15), which AnalyzeRequestRoutes set
    ' to encode OK/DIRECTION_MISMATCH/severity. Save and restore that one cell's color.
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
        worksheet.Range("A4:C4").Value = Array("parse_targets", "출발지IP;목적지IP", "(사용 안 함/예약) 현재 동작에 영향 없음. 출발지IP와 목적지IP는 항상 필수입니다.")
        worksheet.Range("A5:C5").Value = Array("header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소")
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
        worksheet.Range("A1:G1").Value = Array("firewall_name", "source_cidr", "destination_cidr", "direction", "path_order", "enabled", "comment")
        worksheet.Range("A2:G7").Value = Array( _
            Array("SECUI-FW-01", "10.10.0.0/16", "172.16.0.0/16", "OUT", 10, "Y", "업무PC -> 서버"), _
            Array("SECUI-FW-01", "10.10.0.0/16", "10.20.0.0/16", "OUT", 10, "Y", "업무PC -> DMZ"), _
            Array("SECUI-FW-02", "10.10.0.0/16", "10.20.0.0/16", "OUT", 20, "Y", "업무PC -> DMZ"), _
            Array("SECUI-FW-01", "10.10.0.0/16", "8.8.8.0/24", "OUT", 10, "Y", "업무PC -> 외부 DNS"), _
            Array("SECUI-FW-02", "10.10.0.0/16", "8.8.8.0/24", "OUT", 20, "Y", "업무PC -> 외부 DNS"), _
            Array("SECUI-FW-03", "10.10.0.0/16", "8.8.8.0/24", "OUT", 30, "Y", "업무PC -> 외부 DNS"))
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
