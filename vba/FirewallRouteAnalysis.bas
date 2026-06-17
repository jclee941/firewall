Attribute VB_Name = "FirewallRouteAnalysis"
Option Explicit

Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const FIREWALL_RANGE_SHEET As String = "firewall_ranges"
Private Const REQUESTS_SHEET As String = "requests"
Private Const REQUEST_TRACKING_SHEET As String = "_request_tracking"
Private Const ROUTE_RESULTS_SHEET As String = "route_results"

Private Const REQ_DATA_START_ROW As Long = 3
Private Const RCOL_TEAM As Long = 1
Private Const RCOL_DOC_NO As Long = 2
Private Const RCOL_TARGET As Long = 3
Private Const RCOL_SOURCE_IP As Long = 4
Private Const RCOL_SOURCE_NAME As Long = 5
Private Const RCOL_DEST_IP As Long = 6
Private Const RCOL_DEST_NAME As Long = 7
Private Const RCOL_PROTOCOL As Long = 8
Private Const RCOL_PORT As Long = 9
Private Const RCOL_DIRECTION As Long = 10
Private Const TCOL_REQUEST_ROW As Long = 1
Private Const TCOL_SOURCE_FILE As Long = 2
Private Const TCOL_SOURCE_ROW As Long = 3

Private mEnabledFw As Object
Private mRanges As Collection

Public Sub AnalyzeRequestRoutes(Optional ByVal showMessage As Boolean = True)
    Dim requestsSheet As Worksheet
    Dim lastRow As Long
    Dim lastByDst As Long
    Dim rowIndex As Long
    Dim analyzed As Long
    Dim res As Object
    Dim routeResults As Object

    Set requestsSheet = ThisWorkbook.Worksheets(REQUESTS_SHEET)
    Set routeResults = CreateObject("Scripting.Dictionary")
    LoadRouteData

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, RCOL_SOURCE_IP).End(xlUp).Row
    lastByDst = requestsSheet.Cells(requestsSheet.Rows.Count, RCOL_DEST_IP).End(xlUp).Row
    If lastByDst > lastRow Then lastRow = lastByDst

    For rowIndex = REQ_DATA_START_ROW To lastRow
        Dim srcIp As String
        Dim dstIp As String
        Dim direction As String
        srcIp = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_SOURCE_IP).Value))
        dstIp = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_DEST_IP).Value))
        direction = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_DIRECTION).Value))
        If Len(srcIp) > 0 Or Len(dstIp) > 0 Then
            Set res = AnalyzeRoute(srcIp, dstIp, direction)
            Set routeResults(CStr(rowIndex)) = res
            WriteResultRow requestsSheet, rowIndex, res
            analyzed = analyzed + 1
        End If
    Next rowIndex

    RefreshRouteResults requestsSheet, lastRow, routeResults
    If showMessage Then MsgBox CStr(analyzed) & "건의 신청서 방화벽 대역을 분석했습니다.", vbInformation
End Sub

Private Sub RefreshRouteResults(ByVal requestsSheet As Worksheet, ByVal lastRow As Long, ByVal routeResults As Object)
    Dim routeSheet As Worksheet
    Dim requestRow As Long
    Dim outputRow As Long
    Dim res As Object

    Set routeSheet = EnsureRouteResultsSheet()
    WriteRouteResultsHeaders routeSheet
    routeSheet.Rows("2:" & routeSheet.Rows.Count).Clear

    outputRow = 2
    For requestRow = REQ_DATA_START_ROW To lastRow
        If Len(Trim$(CStr(requestsSheet.Cells(requestRow, RCOL_SOURCE_IP).Value))) > 0 _
                Or Len(Trim$(CStr(requestsSheet.Cells(requestRow, RCOL_DEST_IP).Value))) > 0 Then
            routeSheet.Cells(outputRow, 1).Value = requestsSheet.Cells(requestRow, RCOL_TEAM).Value
            routeSheet.Cells(outputRow, 2).Value = requestsSheet.Cells(requestRow, RCOL_DOC_NO).Value
            routeSheet.Cells(outputRow, 3).Value = requestsSheet.Cells(requestRow, RCOL_SOURCE_IP).Value
            routeSheet.Cells(outputRow, 4).Value = requestsSheet.Cells(requestRow, RCOL_SOURCE_NAME).Value
            routeSheet.Cells(outputRow, 5).Value = requestsSheet.Cells(requestRow, RCOL_DEST_IP).Value
            routeSheet.Cells(outputRow, 6).Value = requestsSheet.Cells(requestRow, RCOL_DEST_NAME).Value
            routeSheet.Cells(outputRow, 7).Value = requestsSheet.Cells(requestRow, RCOL_PROTOCOL).Value
            routeSheet.Cells(outputRow, 8).Value = requestsSheet.Cells(requestRow, RCOL_PORT).Value
            routeSheet.Cells(outputRow, 9).Value = requestsSheet.Cells(requestRow, RCOL_DIRECTION).Value
            routeSheet.Cells(outputRow, 10).Value = requestsSheet.Cells(requestRow, RCOL_TARGET).Value
            If routeResults.Exists(CStr(requestRow)) Then
                Set res = routeResults(CStr(requestRow))
                routeSheet.Cells(outputRow, 11).Value = res("status")
                routeSheet.Cells(outputRow, 12).Value = res("validation_message")
                routeSheet.Cells(outputRow, 13).Value = res("firewall_path")
                routeSheet.Cells(outputRow, 14).Value = res("source_zone")
                routeSheet.Cells(outputRow, 15).Value = res("destination_zone")
                routeSheet.Cells(outputRow, 16).Value = res("zone_path")
                routeSheet.Cells(outputRow, 17).Value = res("match_details")
            End If
            routeSheet.Cells(outputRow, 18).Value = RequestTrackingValue(requestRow, TCOL_SOURCE_FILE)
            routeSheet.Cells(outputRow, 19).Value = RequestTrackingValue(requestRow, TCOL_SOURCE_ROW)
            outputRow = outputRow + 1
        End If
    Next requestRow
    routeSheet.Rows(1).Font.Bold = True
    routeSheet.Columns("A:S").AutoFit
End Sub

Private Function EnsureRouteResultsSheet() As Worksheet
    On Error Resume Next
    Set EnsureRouteResultsSheet = ThisWorkbook.Worksheets(ROUTE_RESULTS_SHEET)
    On Error GoTo 0
    If EnsureRouteResultsSheet Is Nothing Then
        Set EnsureRouteResultsSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureRouteResultsSheet.Name = ROUTE_RESULTS_SHEET
    End If
End Function

Private Function RequestTrackingValue(ByVal requestRow As Long, ByVal trackingColumn As Long) As String
    Dim trackingSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Set trackingSheet = EnsureRequestTrackingSheet()
    lastRow = trackingSheet.Cells(trackingSheet.Rows.Count, TCOL_REQUEST_ROW).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If CLng(Val(CStr(trackingSheet.Cells(rowIndex, TCOL_REQUEST_ROW).Value))) = requestRow Then
            RequestTrackingValue = CStr(trackingSheet.Cells(rowIndex, trackingColumn).Value)
            Exit Function
        End If
    Next rowIndex
End Function

Private Function EnsureRequestTrackingSheet() As Worksheet
    On Error Resume Next
    Set EnsureRequestTrackingSheet = ThisWorkbook.Worksheets(REQUEST_TRACKING_SHEET)
    On Error GoTo 0
    If EnsureRequestTrackingSheet Is Nothing Then
        Set EnsureRequestTrackingSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureRequestTrackingSheet.Name = REQUEST_TRACKING_SHEET
    End If
End Function

Private Sub WriteRouteResultsHeaders(ByVal worksheet As Worksheet)
    worksheet.Range("A1:S1").Value = Array("요청부서", "요청번호", "출발지", "출발지설명", "목적지", "목적지설명", "프로토콜", "포트", "방향", "대상방화벽", "검증상태", "검증메시지", "방화벽경로", "출발매칭대역", "목적매칭대역", "대역경로", "매칭근거", "원본파일", "원본행")
End Sub

Private Sub WriteResultRow(ByVal sheet As Worksheet, ByVal rowIndex As Long, ByVal res As Object)
    Dim targetCell As Range
    Set targetCell = sheet.Cells(rowIndex, RCOL_TARGET)
    If Len(Trim$(CStr(targetCell.Value))) = 0 Then targetCell.Value = res("target_firewalls")
    If res("status") = "OK" Then
        targetCell.Interior.Color = RGB(217, 234, 211)
    ElseIf res("status") = "DIRECTION_MISMATCH" Then
        targetCell.Interior.Color = RGB(255, 242, 204)
    Else
        targetCell.Interior.Color = RGB(255, 230, 230)
    End If
End Sub

Private Sub LoadRouteData()
    Set mEnabledFw = CreateObject("Scripting.Dictionary")
    Set mRanges = New Collection

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Set ws = ThisWorkbook.Worksheets(FIREWALLS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        Dim firewallName As String
        firewallName = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        If Len(firewallName) > 0 Then
            If IsEnabled(ws.Cells(rowIndex, 3).Value) Then mEnabledFw(firewallName) = True
        End If
    Next rowIndex

    Set ws = ThisWorkbook.Worksheets(FIREWALL_RANGE_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        Dim fw As String
        fw = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        If Len(fw) > 0 And mEnabledFw.Exists(fw) Then
            If IsEnabled(ws.Cells(rowIndex, 6).Value) Then
                Dim item As Object
                Set item = CreateObject("Scripting.Dictionary")
                item("firewall_name") = fw
                item("source_cidr") = Trim$(CStr(ws.Cells(rowIndex, 2).Value))
                item("destination_cidr") = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
                item("direction") = Trim$(CStr(ws.Cells(rowIndex, 4).Value))
                item("path_order") = ParsePathOrder(ws.Cells(rowIndex, 5).Value)
                item("row_order") = rowIndex - 1
                mRanges.Add item
            End If
        End If
    Next rowIndex
End Sub

Private Function InvalidAddressResult(ByVal side As String, ByVal token As String, ByVal srcIp As String, ByVal dstIp As String) As Object
    Dim res As Object
    Set res = NewResult()
    res("status") = "INVALID_ADDRESS"
    res("validation_message") = "Invalid IPv4 address in request " & side & ": " & token
    res("match_details") = "source_ip=" & srcIp & "; destination_ip=" & dstIp
    Set InvalidAddressResult = res
End Function

Public Function AnalyzeRoute(ByVal srcIp As String, ByVal dstIp As String, ByVal direction As String) As Object
    ' Surface a malformed request IP (IPv6, leading-zero octets, garbage) as a
    ' visible INVALID_ADDRESS instead of a silent NO_MATCH, mirroring
    ' RouteEngine.analyze in tests/route_oracle.py. Blank cells are handled
    ' downstream (match nothing); ANY is a wildcard.
    Dim badToken As String
    badToken = FirstInvalidToken(srcIp)
    If Len(badToken) > 0 Then
        Set AnalyzeRoute = InvalidAddressResult("source", badToken, srcIp, dstIp)
        Exit Function
    End If
    badToken = FirstInvalidToken(dstIp)
    If Len(badToken) > 0 Then
        Set AnalyzeRoute = InvalidAddressResult("destination", badToken, srcIp, dstIp)
        Exit Function
    End If

    Dim flowDirection As String
    flowDirection = NormalizeDirection(direction)
    If flowDirection = "#INVALID" Then
        Dim invalid As Object
        Set invalid = NewResult()
        invalid("status") = "DIRECTION_MISMATCH"
        invalid("validation_message") = "Invalid direction: " & direction
        Set AnalyzeRoute = invalid
        Exit Function
    End If

    Dim matches As Collection
    Set matches = SelectMatches(srcIp, dstIp, flowDirection)
    If matches.Count > 0 Then
        Set AnalyzeRoute = ResultFromMatches(matches)
        Exit Function
    End If

    Dim reverseMatches As Collection
    Set reverseMatches = SelectMatches(dstIp, srcIp, "BOTH")
    If reverseMatches.Count > 0 Then
        Dim mismatch As Object
        Set mismatch = NewResult()
        mismatch("status") = "DIRECTION_MISMATCH"
        mismatch("validation_message") = "No definition for requested direction; opposite direction exists"
        mismatch("match_details") = "source_ip=" & srcIp & "; destination_ip=" & dstIp
        Set AnalyzeRoute = mismatch
        Exit Function
    End If

    Dim noMatch As Object
    Set noMatch = NewResult()
    noMatch("status") = "NO_MATCH"
    noMatch("validation_message") = "No firewall range definition matched"
    noMatch("match_details") = "source_ip=" & srcIp & "; destination_ip=" & dstIp
    Set AnalyzeRoute = noMatch
End Function

Private Function SelectMatches(ByVal srcIp As String, ByVal dstIp As String, ByVal flowDirection As String) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim item As Variant
    For Each item In mRanges
        If DirectionMatches(CStr(item("direction")), flowDirection) Then
            If AddressListOverlaps(srcIp, CStr(item("source_cidr"))) Then
                If AddressListOverlaps(dstIp, CStr(item("destination_cidr"))) Then
                    AddSortedMatch result, item
                End If
            End If
        End If
    Next item
    Set SelectMatches = result
End Function

Private Sub AddSortedMatch(ByVal matches As Collection, ByVal item As Object)
    Dim index As Long
    For index = 1 To matches.Count
        If MatchKey(item) < MatchKey(matches(index)) Then
            matches.Add item, , index
            Exit Sub
        End If
    Next index
    matches.Add item
End Sub

Private Function MatchKey(ByVal item As Object) As String
    MatchKey = Format$(CLng(item("path_order")), "000000") & "|" & _
               Format$(CLng(item("row_order")), "000000") & "|" & _
               CStr(item("firewall_name"))
End Function

Private Function DirectionMatches(ByVal ruleDirection As String, ByVal flowDirection As String) As Boolean
    Dim ruleValue As String
    ruleValue = NormalizeDirection(ruleDirection)
    If ruleValue = "#INVALID" Then Exit Function
    If flowDirection = "BOTH" Or ruleValue = "BOTH" Then
        DirectionMatches = True
    Else
        DirectionMatches = (ruleValue = flowDirection)
    End If
End Function

Public Function NormalizeDirection(ByVal direction As String) As String
    Dim value As String
    value = UCase$(Trim$(direction))
    If Len(value) = 0 Then
        NormalizeDirection = "BOTH"
    ElseIf value = "IN" Or value = "OUT" Or value = "BOTH" Then
        NormalizeDirection = value
    Else
        NormalizeDirection = NormalizeDirectionSynonym(value)
    End If
End Function

Private Function NormalizeDirectionSynonym(ByVal value As String) As String
    ' value is already UCase$ + Trim$. Blank/IN/OUT/BOTH handled by caller.
    Select Case value
        Case "INBOUND", "인바운드", "수신"
            NormalizeDirectionSynonym = "IN"
        Case "OUTBOUND", "아웃바운드", "송신"
            NormalizeDirectionSynonym = "OUT"
        Case "ANY", "ALL", "양방향", "양방", "쌍방향", "BIDIRECTIONAL", "BI-DIRECTIONAL"
            NormalizeDirectionSynonym = "BOTH"
        Case Else
            NormalizeDirectionSynonym = NormalizeDirectionArrowPhrase(value)
    End Select
End Function

Private Function NormalizeDirectionArrowPhrase(ByVal value As String) As String
    Dim canonical As String
    Dim src As String
    Dim dst As String
    Dim pos As Long
    canonical = Replace(value, Chr$(&H2192), ">")
    canonical = Replace(canonical, "->", ">")
    canonical = Replace(canonical, "-", ">")
    pos = InStr(canonical, ">")
    If pos = 0 Then
        NormalizeDirectionArrowPhrase = "#INVALID"
        Exit Function
    End If
    src = Trim$(Left$(canonical, pos - 1))
    dst = Trim$(Mid$(canonical, pos + 1))
    ' Reject 3+ token phrases (e.g. A>B>C) so only a clean pair resolves.
    If InStr(dst, ">") > 0 Then
        NormalizeDirectionArrowPhrase = "#INVALID"
        Exit Function
    End If
    If IsOutsideToken(src) And IsInsideToken(dst) Then
        NormalizeDirectionArrowPhrase = "IN"
    ElseIf IsInsideToken(src) And IsOutsideToken(dst) Then
        NormalizeDirectionArrowPhrase = "OUT"
    Else
        NormalizeDirectionArrowPhrase = "#INVALID"
    End If
End Function

Private Function IsInsideToken(ByVal token As String) As Boolean
    IsInsideToken = (token = "내부" Or token = "INSIDE" Or token = "INTERNAL")
End Function

Private Function IsOutsideToken(ByVal token As String) As Boolean
    IsOutsideToken = (token = "외부" Or token = "OUTSIDE" Or token = "EXTERNAL")
End Function

Private Function AddressListOverlaps(ByVal requestValue As String, ByVal definitionValue As String) As Boolean
    Dim requests() As String
    Dim definitions() As String
    ' A blank request token list matches NOTHING, even an ANY definition: an
    ' incomplete request (empty IP cell) must never resolve to a route. Check the
    ' request side BEFORE the ANY short-circuit (mirrors Python _address_list_overlaps).
    requests = SplitAddressList(requestValue)
    If UBound(requests) < LBound(requests) Then Exit Function

    If IsAnyCidr(definitionValue) Then
        AddressListOverlaps = True
        Exit Function
    End If

    definitions = SplitAddressList(definitionValue)
    If UBound(definitions) < LBound(definitions) Then Exit Function

    Dim i As Long
    Dim j As Long
    For i = LBound(requests) To UBound(requests)
        For j = LBound(definitions) To UBound(definitions)
            If RangesOverlap(requests(i), definitions(j)) Then
                AddressListOverlaps = True
                Exit Function
            End If
        Next j
    Next i
End Function

Private Function ResultFromMatches(ByVal matches As Collection) As Object
    Dim res As Object
    Set res = NewResult()
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim target As String
    Dim path As String
    Dim details As String
    Dim item As Variant

    For Each item In matches
        If Len(path) > 0 Then path = path & ">"
        path = path & CStr(item("firewall_name"))
        If Not seen.Exists(CStr(item("firewall_name"))) Then
            seen(CStr(item("firewall_name"))) = True
            If Len(target) > 0 Then target = target & ";"
            target = target & CStr(item("firewall_name"))
        End If
        If Len(details) > 0 Then details = details & "; "
        details = details & CStr(item("firewall_name")) & ": " & _
                  CStr(item("source_cidr")) & " -> " & _
                  CStr(item("destination_cidr")) & " (" & _
                  CStr(item("direction")) & ")"
    Next item

    Dim first As Object
    Set first = matches(1)
    res("status") = "OK"
    res("target_firewalls") = target
    res("firewall_path") = path
    res("source_zone") = CStr(first("source_cidr"))
    res("destination_zone") = CStr(first("destination_cidr"))
    res("zone_path") = CStr(first("source_cidr")) & ">" & CStr(first("destination_cidr"))
    res("validation_message") = "Firewall range definition matched"
    res("match_details") = details
    res("path_count") = matches.Count
    Set ResultFromMatches = res
End Function

Private Function NewResult() As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("status") = ""
    r("target_firewalls") = ""
    r("firewall_path") = ""
    r("zone_path") = ""
    r("source_zone") = ""
    r("destination_zone") = ""
    r("validation_message") = ""
    r("match_details") = ""
    r("path_count") = 0
    Set NewResult = r
End Function

Private Function SplitAddressList(ByVal text As String) As String()
    Dim normalized As String
    normalized = Replace(CStr(text), ChrW(160), " ")
    normalized = Replace(normalized, vbTab, " ")
    normalized = Replace(normalized, vbCrLf, ";")
    normalized = Replace(normalized, vbCr, ";")
    normalized = Replace(normalized, vbLf, ";")
    normalized = Replace(normalized, ",", ";")
    normalized = Replace(normalized, ChrW(65292), ";")
    normalized = Replace(normalized, ChrW(65307), ";")
    normalized = Replace(normalized, " ", ";")

    Dim raw() As String
    raw = Split(normalized, ";")
    Dim temp() As String
    Dim count As Long
    count = -1
    Dim i As Long
    For i = LBound(raw) To UBound(raw)
        Dim token As String
        token = Trim$(raw(i))
        If Len(token) > 0 Then
            count = count + 1
            ReDim Preserve temp(0 To count)
            temp(count) = token
        End If
    Next i
    If count < 0 Then
        ' No real tokens: return a genuinely empty array (UBound < LBound) so
        ' a blank IP cell matches nothing, mirroring Python split_address_list.
        SplitAddressList = Split(vbNullString, ";", 0)
    Else
        SplitAddressList = temp
    End If
End Function

Private Function IsAnyCidr(ByVal text As String) As Boolean
    Dim value As String
    value = UCase$(Trim$(text))
    IsAnyCidr = (Len(value) = 0 Or value = "*" Or value = "ANY" Or _
                 value = "ALL" Or value = "0.0.0.0/0")
End Function

Private Function IsStrictIpv4(ByVal text As String) As Boolean
    ' True when text parses as a strict IPv4 address or CIDR (the same contract
    ' as firewall_policy.cidr.parse_ipv4_network). IPv6, leading-zero octets and
    ' malformed values return False.
    On Error GoTo bad
    Dim base As String
    base = CidrBaseIp(text)
    Dim prefix As Long
    prefix = CidrPrefixLength(text)   ' raises on a bad prefix
    IpToNumber base                   ' raises on a bad address
    IsStrictIpv4 = True
    Exit Function
bad:
    IsStrictIpv4 = False
End Function

Private Function IsInvalidAddress(ByVal token As String) As Boolean
    ' A concrete address token that fails the strict-IPv4 contract is invalid.
    ' ANY/blank wildcards are NOT invalid (handled above the overlap layer).
    If IsAnyCidr(token) Then Exit Function
    IsInvalidAddress = Not IsStrictIpv4(token)
End Function

Private Function FirstInvalidToken(ByVal listValue As String) As String
    ' Return the first non-ANY token in a list cell that is not strict IPv4,
    ' else an empty string. Used to surface INVALID_ADDRESS.
    Dim toks() As String
    toks = SplitAddressList(listValue)
    If UBound(toks) < LBound(toks) Then Exit Function
    Dim i As Long
    For i = LBound(toks) To UBound(toks)
        If IsInvalidAddress(toks(i)) Then
            FirstInvalidToken = toks(i)
            Exit Function
        End If
    Next i
End Function

Private Function RangesOverlap(ByVal leftCidr As String, ByVal rightCidr As String) As Boolean
    On Error GoTo bad
    If IsAnyCidr(leftCidr) Or IsAnyCidr(rightCidr) Then
        RangesOverlap = True
        Exit Function
    End If
    Dim ls As Double
    Dim le As Double
    Dim rs As Double
    Dim re As Double
    ls = CidrStart(leftCidr)
    le = CidrEnd(leftCidr)
    rs = CidrStart(rightCidr)
    re = CidrEnd(rightCidr)
    RangesOverlap = (ls <= re And rs <= le)
    Exit Function
bad:
    RangesOverlap = False
End Function

Private Function IpToNumber(ByVal ipText As String) As Double
    Dim parts() As String
    parts = Split(Trim$(ipText), ".")
    If UBound(parts) - LBound(parts) + 1 <> 4 Then Err.Raise 5
    Dim i As Long
    Dim value As Double
    value = 0
    For i = 0 To 3
        value = value * 256# + ParseOctet(parts(i))
    Next i
    IpToNumber = value
End Function

Private Function ParseOctet(ByVal octetText As String) As Long
    ' Strict IPv4 octet: digits only, 0..255, and NO leading zero (the literal
    ' "0" is allowed). Mirrors firewall_policy.cidr._parse_octet so VBA, the
    ' route oracle and the SECUI CLI agree exactly. Reject everything else.
    Dim clean As String
    clean = Trim$(octetText)
    If Len(clean) = 0 Then Err.Raise 5
    Dim j As Long
    For j = 1 To Len(clean)
        If InStr("0123456789", Mid$(clean, j, 1)) = 0 Then Err.Raise 5
    Next j
    If Len(clean) > 1 And Left$(clean, 1) = "0" Then Err.Raise 5
    Dim octet As Long
    octet = CLng(clean)
    If octet < 0 Or octet > 255 Then Err.Raise 5
    ParseOctet = octet
End Function

Private Function CidrPrefixLength(ByVal cidrText As String) As Long
    Dim segs() As String
    segs = Split(cidrText, "/")
    If UBound(segs) = 0 Then
        CidrPrefixLength = 32
        Exit Function
    End If
    ' Exactly one '/' is allowed; '10.0.0.1/24/garbage' must be rejected so VBA
    ' agrees with firewall_policy.cidr._prefix_length.
    If UBound(segs) <> 1 Then Err.Raise 5
    Dim raw As String
    raw = Trim$(segs(1))
    If Len(raw) = 0 Then Err.Raise 5
    Dim k As Long
    For k = 1 To Len(raw)
        If InStr("0123456789", Mid$(raw, k, 1)) = 0 Then Err.Raise 5
    Next k
    CidrPrefixLength = CLng(raw)
    If CidrPrefixLength < 0 Or CidrPrefixLength > 32 Then Err.Raise 5
End Function

Private Function CidrBaseIp(ByVal cidrText As String) As String
    CidrBaseIp = Trim$(Split(cidrText, "/")(0))
End Function

Private Function CidrBlockSize(ByVal cidrText As String) As Double
    CidrBlockSize = 2 ^ (32 - CidrPrefixLength(cidrText))
End Function

Private Function CidrStart(ByVal cidrText As String) As Double
    Dim base As Double
    Dim block As Double
    base = IpToNumber(CidrBaseIp(cidrText))
    block = CidrBlockSize(cidrText)
    CidrStart = Fix(base / block) * block
End Function

Private Function CidrEnd(ByVal cidrText As String) As Double
    CidrEnd = CidrStart(cidrText) + CidrBlockSize(cidrText) - 1
End Function

Private Function IsEnabled(ByVal value As Variant) As Boolean
    Dim text As String
    text = UCase$(Trim$(CStr(value)))
    IsEnabled = (text = "Y" Or text = "YES" Or text = "TRUE" Or text = "1")
End Function

Private Function ParsePathOrder(ByVal value As Variant) As Long
    If IsNumeric(value) Then
        ParsePathOrder = CLng(value)
    Else
        ParsePathOrder = 999999
    End If
End Function
