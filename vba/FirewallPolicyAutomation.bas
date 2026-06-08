Option Explicit

Private Const SETTINGS_SHEET As String = "settings"
Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const REQUESTS_SHEET As String = "requests"

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

Public Sub SetupFirewallAutomationWorkbook()
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim settingsSheet As Worksheet

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)

    WriteRequestHeaders requestsSheet
    WriteFirewallHeaders firewallsSheet
    WriteSettings settingsSheet
    FormatRequestsSheet requestsSheet
    FormatFirewallsSheet firewallsSheet

    MsgBox "방화벽 정책 자동화 시트 구성이 완료되었습니다.", vbInformation
End Sub

Public Sub MergeFirewallRequestFolder()
    Dim folderPath As String
    Dim requestsSheet As Worksheet
    Dim firewallsSheet As Worksheet
    Dim settingsSheet As Worksheet
    Dim nextRow As Long
    Dim mergedCount As Long

    Set requestsSheet = EnsureSheet(REQUESTS_SHEET)
    Set firewallsSheet = EnsureSheet(FIREWALLS_SHEET)
    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    WriteRequestHeaders requestsSheet
    WriteSettings settingsSheet
    folderPath = RequestFolderPath(settingsSheet)
    If Len(folderPath) = 0 Then Exit Sub

    requestsSheet.Rows("2:" & requestsSheet.Rows.Count).ClearContents
    nextRow = 2
    mergedCount = MergeFolderFiles(folderPath, requestsSheet, firewallsSheet, nextRow)
    MarkDuplicateRequests requestsSheet
    MarkUnmatchedFirewalls requestsSheet
    FormatRequestsSheet requestsSheet

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

Private Function MergeFolderFiles(ByVal folderPath As String, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByRef nextRow As Long) As Long
    Dim fileName As String
    Dim mergedCount As Long

    fileName = Dir(folderPath & Application.PathSeparator & "*.xls*")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            mergedCount = mergedCount + MergeWorkbookFile(folderPath & Application.PathSeparator & fileName, fileName, requestsSheet, firewallsSheet, nextRow)
        End If
        fileName = Dir()
    Loop

    MergeFolderFiles = mergedCount
End Function

Private Function MergeWorkbookFile(ByVal filePath As String, ByVal sourceFileName As String, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByRef nextRow As Long) As Long
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim headerMap As Object
    Dim headerRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim mergedCount As Long

    On Error GoTo OpenFailed
    Set sourceBook = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set sourceSheet = sourceBook.Worksheets(1)
    headerRow = FindHeaderRow(sourceSheet)
    Set headerMap = BuildHeaderMap(sourceSheet, headerRow)
    ValidateRequiredHeaders headerMap, sourceFileName

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, headerMap("출발지ip")).End(xlUp).Row
    For rowIndex = headerRow + 1 To lastRow
        If Len(Trim$(CStr(sourceSheet.Cells(rowIndex, headerMap("출발지ip")).Value))) > 0 Then
            CopyRequestRow sourceSheet, rowIndex, headerMap, requestsSheet, firewallsSheet, nextRow, sourceFileName
            nextRow = nextRow + 1
            mergedCount = mergedCount + 1
        End If
    Next rowIndex

    sourceBook.Close SaveChanges:=False
    MergeWorkbookFile = mergedCount
    Exit Function

OpenFailed:
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    MsgBox "파일을 열 수 없습니다: " & filePath & vbCrLf & Err.Description, vbExclamation
    MergeWorkbookFile = 0
End Function

Private Sub CopyRequestRow(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long, ByVal headerMap As Object, ByVal requestsSheet As Worksheet, ByVal firewallsSheet As Worksheet, ByVal targetRow As Long, ByVal sourceFileName As String)
    requestsSheet.Cells(targetRow, COL_SOURCE_FILE).Value = sourceFileName
    requestsSheet.Cells(targetRow, COL_SOURCE_ROW).Value = sourceRow
    requestsSheet.Cells(targetRow, COL_SOURCE_IP).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("출발지ip")).Value))
    requestsSheet.Cells(targetRow, COL_SOURCE_NAME).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("출발지")).Value))
    requestsSheet.Cells(targetRow, COL_DESTINATION_IP).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("목적지ip")).Value))
    requestsSheet.Cells(targetRow, COL_DESTINATION_NAME).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("목적지")).Value))
    requestsSheet.Cells(targetRow, COL_PROTOCOL).Value = UCase$(Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("프로토콜")).Value)))
    requestsSheet.Cells(targetRow, COL_PORT).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("포트")).Value))
    requestsSheet.Cells(targetRow, COL_DIRECTION).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("방향")).Value))
    requestsSheet.Cells(targetRow, COL_PURPOSE).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("용도")).Value))
    requestsSheet.Cells(targetRow, COL_START_DATE).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("시작일")).Value))
    requestsSheet.Cells(targetRow, COL_END_DATE).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("종료일")).Value))
    requestsSheet.Cells(targetRow, COL_NOTE).Value = Trim$(CStr(sourceSheet.Cells(sourceRow, headerMap("비고")).Value))
    requestsSheet.Cells(targetRow, COL_TARGET_FIREWALLS).Value = ResolveTargetFirewalls(firewallsSheet, requestsSheet, targetRow)
End Sub

Private Function ResolveTargetFirewalls(ByVal firewallsSheet As Worksheet, ByVal requestsSheet As Worksheet, ByVal requestRow As Long) As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim firewallName As String
    Dim cidrText As String
    Dim matched As Object

    Set matched = CreateObject("Scripting.Dictionary")
    lastRow = firewallsSheet.Cells(firewallsSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 2 To lastRow
        firewallName = Trim$(CStr(firewallsSheet.Cells(rowIndex, 1).Value))
        cidrText = Trim$(CStr(firewallsSheet.Cells(rowIndex, 2).Value))
        If Len(firewallName) > 0 And Len(cidrText) > 0 Then
            If RequestRowMatchesCidr(requestsSheet, requestRow, cidrText) Then
                If Not matched.Exists(firewallName) Then matched.Add firewallName, firewallName
            End If
        End If
    Next rowIndex

    ResolveTargetFirewalls = JoinDictionaryKeys(matched)
End Function

Private Function RequestRowMatchesCidr(ByVal requestsSheet As Worksheet, ByVal requestRow As Long, ByVal cidrText As String) As Boolean
    Dim parseColumns As Variant
    Dim index As Long
    Dim columnNumber As Long

    parseColumns = RegisteredParseTargetColumns()
    For index = LBound(parseColumns) To UBound(parseColumns)
        columnNumber = CLng(parseColumns(index))
        If AddressListMatchesCidr(Trim$(CStr(requestsSheet.Cells(requestRow, columnNumber).Value)), cidrText) Then
            RequestRowMatchesCidr = True
            Exit Function
        End If
    Next index
End Function

Private Function RegisteredParseTargetColumns() As Variant
    Dim settingsSheet As Worksheet
    Dim targetText As String

    Set settingsSheet = EnsureSheet(SETTINGS_SHEET)
    targetText = Trim$(CStr(settingsSheet.Range("B3").Value))
    If Len(targetText) = 0 Then targetText = "출발지IP;목적지IP"
    RegisteredParseTargetColumns = ParseTargetColumnsFromText(targetText)
End Function

Private Function ParseTargetColumnsFromText(ByVal targetText As String) As Variant
    Dim names() As String
    Dim columns() As Long
    Dim index As Long

    names = Split(targetText, ";")
    ReDim columns(LBound(names) To UBound(names))
    For index = LBound(names) To UBound(names)
        columns(index) = RequestColumnNumber(Trim$(LCase$(names(index))))
    Next index
    ParseTargetColumnsFromText = columns
End Function

Private Function RequestColumnNumber(ByVal columnName As String) As Long
    Select Case columnName
        Case "출발지ip": RequestColumnNumber = COL_SOURCE_IP
        Case "출발지": RequestColumnNumber = COL_SOURCE_NAME
        Case "목적지ip": RequestColumnNumber = COL_DESTINATION_IP
        Case "목적지": RequestColumnNumber = COL_DESTINATION_NAME
        Case "프로토콜": RequestColumnNumber = COL_PROTOCOL
        Case "포트": RequestColumnNumber = COL_PORT
        Case "방향": RequestColumnNumber = COL_DIRECTION
        Case "용도": RequestColumnNumber = COL_PURPOSE
        Case "시작일": RequestColumnNumber = COL_START_DATE
        Case "종료일": RequestColumnNumber = COL_END_DATE
        Case "비고": RequestColumnNumber = COL_NOTE
        Case Else: Err.Raise vbObjectError + 1002, , "등록되지 않은 파싱 대상 컬럼: " & columnName
    End Select
End Function

Private Function AddressListMatchesCidr(ByVal addressList As String, ByVal cidrList As String) As Boolean
    Dim addresses() As String
    Dim cidrs() As String
    Dim addressIndex As Long
    Dim cidrIndex As Long

    addresses = Split(addressList, ";")
    cidrs = Split(cidrList, ";")
    For addressIndex = LBound(addresses) To UBound(addresses)
        For cidrIndex = LBound(cidrs) To UBound(cidrs)
            If IpMatchesCidr(Trim$(addresses(addressIndex)), Trim$(cidrs(cidrIndex))) Then
                AddressListMatchesCidr = True
                Exit Function
            End If
        Next cidrIndex
    Next addressIndex
End Function

Private Function IpMatchesCidr(ByVal ipText As String, ByVal cidrText As String) As Boolean
    Dim ipParts() As String
    Dim cidrParts() As String
    Dim baseIp As String
    Dim prefixLength As Long
    Dim ipValue As Double
    Dim baseValue As Double
    Dim blockSize As Double

    On Error GoTo NotMatched
    If Len(ipText) = 0 Or Len(cidrText) = 0 Then Exit Function
    ipParts = Split(ipText, "/")
    cidrParts = Split(cidrText, "/")
    baseIp = cidrParts(0)
    If UBound(cidrParts) = 0 Then
        prefixLength = 32
    Else
        prefixLength = CLng(cidrParts(1))
    End If

    ipValue = IpToNumber(ipParts(0))
    baseValue = IpToNumber(baseIp)
    blockSize = 2 ^ (32 - prefixLength)
    IpMatchesCidr = Fix(ipValue / blockSize) = Fix(baseValue / blockSize)
    Exit Function

NotMatched:
    IpMatchesCidr = False
End Function

Private Function IpToNumber(ByVal ipText As String) As Double
    Dim parts() As String

    parts = Split(ipText, ".")
    If UBound(parts) <> 3 Then Err.Raise vbObjectError + 1000, , "Invalid IPv4 address: " & ipText
    IpToNumber = CDbl(parts(0)) * 16777216# + CDbl(parts(1)) * 65536# + CDbl(parts(2)) * 256# + CDbl(parts(3))
End Function

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
        headerName = CanonicalHeaderName(HeaderKey(CStr(worksheet.Cells(headerRow, columnIndex).Value)))
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

Private Sub ValidateRequiredHeaders(ByVal headerMap As Object, ByVal sourceFileName As String)
    Dim requiredHeaders As Variant
    Dim header As Variant

    requiredHeaders = Array("출발지ip", "출발지", "목적지ip", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고")
    For Each header In requiredHeaders
        If Not headerMap.Exists(CStr(header)) Then Err.Raise vbObjectError + 1001, , sourceFileName & " 필수 컬럼 누락: " & CStr(header)
    Next header
End Sub

Private Function JoinDictionaryKeys(ByVal dictionary As Object) As String
    Dim key As Variant
    Dim result As String

    For Each key In dictionary.Keys
        If Len(result) > 0 Then result = result & ";"
        result = result & CStr(key)
    Next key
    JoinDictionaryKeys = result
End Function

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
                worksheet.Rows(rowIndex).Interior.Color = RGB(255, 230, 153)
                worksheet.Rows(CLng(seen(duplicateKey))).Interior.Color = RGB(255, 230, 153)
            Else
                seen.Add duplicateKey, rowIndex
            End If
        End If
    Next rowIndex
End Sub

Private Function RequestDuplicateKey(ByVal worksheet As Worksheet, ByVal rowIndex As Long) As String
    RequestDuplicateKey = Trim$(CStr(worksheet.Cells(rowIndex, COL_SOURCE_IP).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_DESTINATION_IP).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PROTOCOL).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PORT).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_DIRECTION).Value)) & "|" & _
        Trim$(CStr(worksheet.Cells(rowIndex, COL_PURPOSE).Value))
End Function

Private Sub MarkUnmatchedFirewalls(ByVal worksheet As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = worksheet.Cells(worksheet.Rows.Count, COL_SOURCE_IP).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If Len(Trim$(CStr(worksheet.Cells(rowIndex, COL_TARGET_FIREWALLS).Value))) = 0 Then
            worksheet.Cells(rowIndex, COL_TARGET_FIREWALLS).Value = "UNMATCHED"
            worksheet.Cells(rowIndex, COL_TARGET_FIREWALLS).Interior.Color = RGB(255, 199, 206)
        End If
    Next rowIndex
End Sub

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
    worksheet.Range("A1:N1").Value = Array("source_file", "source_row", "target_firewalls", "출발지IP", "출발지", "목적지IP", "목적지", "프로토콜", "포트", "방향", "용도", "시작일", "종료일", "비고")
End Sub

Private Sub WriteFirewallHeaders(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:B1").Value = Array("firewall_name", "cidr_list")
        worksheet.Range("A2:B4").Value = Array(Array("FW-INTERNAL-01", "10.10.0.0/16;172.16.1.0/24"), Array("FW-DMZ-01", "10.20.0.0/16;172.16.20.0/24"), Array("FW-SERVER-01", "10.30.0.0/16;172.16.30.0/24"))
    End If
End Sub

Private Sub WriteSettings(ByVal worksheet As Worksheet)
    If Len(CStr(worksheet.Cells(1, 1).Value)) = 0 Then
        worksheet.Range("A1:B1").Value = Array("key", "value")
        worksheet.Range("A2:B2").Value = Array("request_folder", "")
        worksheet.Range("A3:B3").Value = Array("parse_targets", "출발지IP;목적지IP")
    End If
End Sub

Private Sub FormatRequestsSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:N").AutoFit
    worksheet.Rows(1).Font.Bold = True
    worksheet.Range("A1:N1").AutoFilter
End Sub

Private Function HeaderKey(ByVal headerText As String) As String
    HeaderKey = LCase$(Replace(Trim$(headerText), " ", ""))
End Function

Private Sub FormatFirewallsSheet(ByVal worksheet As Worksheet)
    worksheet.Columns("A:B").AutoFit
    worksheet.Rows(1).Font.Bold = True
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
