Option Explicit

Private Const SETTINGS_SHEET As String = "settings"
Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const REQUESTS_SHEET As String = "requests"
Private Const LOG_SHEET As String = "processing_log"
Private Const NETWORK_SHEET As String = "network_definitions"
Private Const ROUTING_SHEET As String = "routing_paths"

Private Const COL_SOURCE_FILE As Long = 1
Private Const COL_SOURCE_ROW As Long = 2
Private Const COL_TARGET_FIREWALLS As Long = 3
Private Const COL_SOURCE_IP As Long = 4
Private Const COL_SOURCE_NAME As Long = 5
Private Const COL_DESTINATION_IP As Long = 6
Private Const COL_DESTINATION_NAME As Long = 7
Private Const COL_PROTOCOL As Long = 8
Private Const COL_PORT As Long = 9
Private Const COL_DIRECTION As Long = 10
Private Const COL_PURPOSE As Long = 11
Private Const COL_START_DATE As Long = 12
Private Const COL_END_DATE As Long = 13
Private Const COL_NOTE As Long = 14
Private Const COL_VALIDATION_STATUS As Long = 15
Private Const COL_VALIDATION_MESSAGE As Long = 16
Private Const COL_MATCH_DETAILS As Long = 17
Private Const COL_FIREWALL_PATH As Long = 18
Private Const COL_SOURCE_ZONE As Long = 19
Private Const COL_DESTINATION_ZONE As Long = 20
Private Const COL_ZONE_PATH As Long = 21
Private Const COL_REQUEST_TEAM As Long = 22
Private Const COL_REQUEST_DOC_NO As Long = 23
Private Const COL_REQUEST_FOLDER As Long = 24

Private mUserAliases As Object

Public Sub SetupFirewallAutomationWorkbook()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim logSheet As Worksheet
    Dim networkSheet As Worksheet
    Dim routingSheet As Worksheet

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    Set logSheet = EnsureSheet(LOG_SHEET)
    Set networkSheet = EnsureSheet(NETWORK_SHEET)
    Set routingSheet = EnsureSheet(ROUTING_SHEET)

    WriteRequestHeaders requestsSheet
    WriteFirewallHeaders firewallsSheet
    WriteSettings settingsSheet
    WriteLogHeaders logSheet
    WriteNetworkHeaders networkSheet
    WriteRoutingHeaders routingSheet
    FormatRequestsSheet requestsSheet
    FormatFirewallsSheet firewallsSheet
    FormatLogSheet logSheet
    FormatGenericSheet networkSheet, "A:E"
    FormatGenericSheet routingSheet, "A:G"

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
    folderPath = RequestFolderPath(settingsSheet)
    If Len(folderPath) = 0 Then Exit Sub

    requestsSheet.Rows("2:" & requestsSheet.Rows.Count).ClearContents
    logSheet.Rows("2:" & logSheet.Rows.Count).ClearContents
    nextRow = 2
    mergedCount = MergeFolderFiles(folderPath, requestsSheet, firewallsSheet, logSheet, nextRow)
    FormatRequestsSheet requestsSheet
    FormatLogSheet logSheet

    On Error GoTo RouteFailed
    Application.Run "FirewallRouteAnalysis.AnalyzeRequestRoutes"
    On Error GoTo 0
    GoTo RouteDone
RouteFailed:
    AppendProcessingLog logSheet, "(route analysis)", "ERROR", 0, "라우트 분석 실패: " & Err.Description
    MsgBox "라우트 분석 중 오류가 발생했습니다: " & Err.Description, vbExclamation
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
        mergedCount = mergedCount + MergeFolderTree(folderPath & Application.PathSeparator & CStr(subName), CStr(subName), requestsSheet, firewallsSheet, logSheet, nextRow)
    Next subName

    MergeFolderTree = mergedCount
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

Private Sub ParseRequestFolderName(ByVal folderName As String, ByRef team As String, ByRef docNo As String)
    ' 정보보호센터_1234 -> team=정보보호센터, docNo=1234 (split on LAST underscore)
    Dim s As String
    s = Trim$(folderName)
    If Len(s) = 0 Then
        team = "" : docNo = "" : Exit Sub
    End If
    Dim idx As Long
    idx = InStrRev(s, "_")
    If idx = 0 Then
        team = Trim$(s) : docNo = ""
    Else
        team = Trim$(Left$(s, idx - 1))
        docNo = Trim$(Mid$(s, idx + 1))
    End If
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
    Set sourceSheet = sourceBook.Worksheets(1)
    headerRow = FindHeaderRow(sourceSheet)
    Set headerMap = BuildHeaderMap(sourceSheet, headerRow)
    ValidateRequiredHeaders headerMap, sourceFileName

    lastRow = SourceLastRow(sourceSheet, headerMap)
    For rowIndex = headerRow + 1 To lastRow
        If RequestSourceRowHasData(sourceSheet, rowIndex, headerMap) Then
            CopyRequestRow sourceSheet, rowIndex, headerMap, requestsSheet, firewallsSheet, nextRow, sourceFileName, folderName
            nextRow = nextRow + 1
            mergedCount = mergedCount + 1
        End If
    Next rowIndex

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

Private Sub CopyRequestRow(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByVal targetRow As Long, ByVal sourceFileName As String, ByVal folderName As String)
    requestsSheet.Cells(targetRow, COL_SOURCE_FILE).Value = sourceFileName
    requestsSheet.Cells(targetRow, COL_SOURCE_ROW).Value = sourceRow
    requestsSheet.Cells(targetRow, COL_SOURCE_IP).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("출발지ip"))))
    requestsSheet.Cells(targetRow, COL_SOURCE_NAME).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("출발지"))))
    requestsSheet.Cells(targetRow, COL_DESTINATION_IP).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("목적지ip"))))
    requestsSheet.Cells(targetRow, COL_DESTINATION_NAME).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("목적지"))))
    requestsSheet.Cells(targetRow, COL_PROTOCOL).Value = UCase$(Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("프로토콜")))))
    requestsSheet.Cells(targetRow, COL_PORT).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("포트"))))
    requestsSheet.Cells(targetRow, COL_DIRECTION).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("방향"))))
    requestsSheet.Cells(targetRow, COL_PURPOSE).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("용도"))))
    requestsSheet.Cells(targetRow, COL_START_DATE).Value = FormatMetadataDate(ReadDataCell(sourceSheet, sourceRow, headerMap("시작일")))
    requestsSheet.Cells(targetRow, COL_END_DATE).Value = FormatMetadataDate(ReadDataCell(sourceSheet, sourceRow, headerMap("종료일")))
    requestsSheet.Cells(targetRow, COL_NOTE).Value = Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("비고"))))
    Dim reqTeam As String, reqDocNo As String
    ParseRequestFolderName folderName, reqTeam, reqDocNo
    requestsSheet.Cells(targetRow, COL_REQUEST_TEAM).Value = reqTeam
    requestsSheet.Cells(targetRow, COL_REQUEST_DOC_NO).Value = reqDocNo
    requestsSheet.Cells(targetRow, COL_REQUEST_FOLDER).Value = folderName
    WriteRowValidation requestsSheet, targetRow
End Sub

Private Function FindHeaderRow(ByVal worksheet As Worksheet) As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim lastColumn As Long
    Dim valueText As String

    For rowIndex = 1 To 30
        lastColumn = worksheet.Cells(rowIndex, worksheet.Columns.Count).End(xlToLeft).Column
        For columnIndex = 1 To lastColumn
            valueText = HeaderKey(CStr(worksheet.Cells(rowIndex, columnIndex).Value))
            If valueText = "no" Or valueText = "번호" Then
                FindHeaderRow = rowIndex
                Exit Function
            End If
        Next columnIndex
    Next rowIndex

    Err.Raise vbObjectError + 1003, , "No/번호 헤더 행을 찾을 수 없습니다. B열 No 기준 신청서인지 확인하세요."
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
        Case "no", "번호": CanonicalHeaderName = "no"
        Case "출발지ip", "출발ip", "sourceip", "srcip", "src": CanonicalHeaderName = "출발지ip"
        Case "출발지", "출발지명", "출발", "source", "srcname": CanonicalHeaderName = "출발지"
        Case "목적지ip", "목적ip", "destinationip", "dstip", "dst": CanonicalHeaderName = "목적지ip"
        Case "목적지", "목적지명", "목적", "destination", "dstname": CanonicalHeaderName = "목적지"
        Case "프로토콜", "protocol", "proto": CanonicalHeaderName = "프로토콜"
        Case "포트", "port", "dport", "목적지포트": CanonicalHeaderName = "포트"
        Case "방향", "direction": CanonicalHeaderName = "방향"
        Case "용도", "목적", "usage", "purpose": CanonicalHeaderName = "용도"
        Case "시작일", "시작", "startdate", "start": CanonicalHeaderName = "시작일"
        Case "종료일", "종료", "enddate", "end": CanonicalHeaderName = "종료일"
        Case "비고", "메모", "remark", "remarks", "note": CanonicalHeaderName = "비고"
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
    If Len(Trim$(raw)) = 0 Then Exit Sub
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

    requiredHeaders = Array("출발지ip", "출발지", "목적지ip", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고")
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
    ' and walk back to the last row that ReadDataCell sees as having IP data.
    Dim usedLast As Long
    Dim r As Long

    usedLast = sourceSheet.UsedRange.Row + sourceSheet.UsedRange.Rows.Count - 1
    For r = usedLast To 1 Step -1
        If Len(Trim$(CStr(ReadDataCell(sourceSheet, r, headerMap("출발지ip"))))) > 0 _
           Or Len(Trim$(CStr(ReadDataCell(sourceSheet, r, headerMap("목적지ip"))))) > 0 Then
            SourceLastRow = r
            Exit Function
        End If
    Next r
    SourceLastRow = 1
End Function

Private Function RequestSourceRowHasData(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object) As Boolean
    RequestSourceRowHasData = Len(Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("출발지ip"))))) > 0 Or _
        Len(Trim$(CStr(ReadDataCell(sourceSheet, sourceRow, headerMap("목적지ip"))))) > 0
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

Private Sub MarkDuplicateRequests(ByVal worksheet As Worksheet)
    Dim seen As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim duplicateKey As String

    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = worksheet.Cells(worksheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    For rowIndex = 2 To lastRow
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
    ' to encode OK/MULTI_PATH/severity. Save and restore that one cell's color.
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
    worksheet.Range("A1:X1").Value = Array("source_file", "source_row", "target_firewalls", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고", "validation_status", "validation_message", "match_details", "firewall_path", "source_zone", "destination_zone", "zone_path", "request_team", "request_doc_no", "request_folder")
End Sub

Private Sub WriteFirewallHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:D1").Value = Array("firewall_name", "vendor", "enabled", "comment")
        worksheet.Range("A2:D4").Value = Array( _
            Array("SECUI-FW-01", "SECUI", "Y", "내부-서버 구간"), _
            Array("SECUI-FW-02", "SECUI", "Y", "중간-DMZ 구간"), _
            Array("SECUI-FW-03", "SECUI", "Y", "DMZ-외부 구간"))
    End If
End Sub

Private Sub WriteSettings(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:C1").Value = Array("key", "value", "설명")
        worksheet.Range("A2:C2").Value = Array("request_folder", "", "신청서 엑셀이 모여 있는 폴더 경로. 하위 폴더(예: 정보보호센터_1234)까지 재귀 탐색합니다.")
        worksheet.Range("A3:C3").Value = Array("parse_targets", "출발지IP;목적지IP", "적용대상방화벽 산정에 쓸 IP 컬럼(세미콜론 구분). IP 컬럼만 등록.")
        worksheet.Range("A4:C4").Value = Array("route_legacy_fallback", "FALSE", "라우팅 경로를 못 찾을 때 기존 CIDR 겹침 방식으로 대체할지(TRUE/FALSE).")
        worksheet.Range("A5:C5").Value = Array("header_alias", "", "비표준 헤더 별칭. 형식: 출발지IP=출발지주소,Source Addr; 목적지IP=목적지주소")
    End If
End Sub

Private Sub FormatRequestsSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:X").AutoFit
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:X1").AutoFilter
End Sub

Private Function HeaderKey(ByVal headerText As String) As String
    HeaderKey = LCase$(Replace(Trim$(headerText), " ", ""))
End Function

Private Sub FormatFirewallsSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:B").AutoFit
    worksheet.Rows(1).Font.Bold = True
End Sub

Private Sub WriteNetworkHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:E1").Value = Array("network_name", "network_cidr", "zone", "site", "enabled")
        worksheet.Range("A2:E7").Value = Array( _
            Array("업무PC망", "10.10.0.0/16", "internal", "본사", "Y"), _
            Array("서버망", "172.16.1.0/24", "server", "IDC", "Y"), _
            Array("중간망", "10.30.0.0/16", "transit", "IDC", "Y"), _
            Array("DMZ망", "10.20.0.0/16", "dmz", "IDC", "Y"), _
            Array("외부", "0.0.0.0/0", "outside", "공통", "Y"), _
            Array("서버DMZ", "172.16.20.0/24", "dmz", "IDC", "Y"))
    End If
End Sub

Private Sub WriteRoutingHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:G1").Value = Array("firewall_name", "src_zone", "dst_zone", "ingress_if", "egress_if", "path_order", "enabled")
        worksheet.Range("A2:G5").Value = Array( _
            Array("SECUI-FW-01", "internal", "server", "eth1", "eth2", 10, "Y"), _
            Array("SECUI-FW-01", "internal", "transit", "eth1", "eth3", 20, "Y"), _
            Array("SECUI-FW-02", "transit", "dmz", "eth1", "eth2", 30, "Y"), _
            Array("SECUI-FW-03", "dmz", "outside", "eth1", "eth2", 40, "Y"))
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
