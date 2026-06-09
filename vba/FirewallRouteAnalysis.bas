Attribute VB_Name = "FirewallRouteAnalysis"
Option Explicit

' Route analysis for firewall requests.
' Mirrors tests/route_oracle.py exactly (the binding contract).
'
' Sheets used:
'   requests             : source/destination IP rows + analysis output
'   network_definitions  : network_name, network_cidr, zone, site, enabled
'   firewalls            : firewall_name, vendor, enabled, comment
'   routing_paths        : firewall_name, src_zone, dst_zone, ingress_if, egress_if, path_order, enabled
'   settings             : key/value (route_legacy_fallback toggle)
'
' Algorithm:
'   src/dst IP -> zone via longest-prefix match in network_definitions
'   directed zone graph from routing_paths (enabled + enabled firewall only)
'   deterministic BFS shortest zone path
'   tie-break edge key = Format(path_order,"000000") & "|" & firewall_name & "|" & dst_zone
'   statuses: OK, MULTI_PATH, INTRA_ZONE, ZONE_UNRESOLVED, NO_PATH,
'             DIRECTION_MISMATCH, LEGACY_FALLBACK

Private Const NETWORK_SHEET As String = "network_definitions"
Private Const FIREWALLS_SHEET As String = "firewalls"
Private Const ROUTING_SHEET As String = "routing_paths"
Private Const REQUESTS_SHEET As String = "requests"
Private Const SETTINGS_SHEET As String = "settings"

' requests output columns (1-based). Inputs read by header name.
Private Const RCOL_TARGET As Long = 3          ' target_firewalls
Private Const RCOL_SOURCE_IP As Long = 4       ' 출발지IP
Private Const RCOL_DEST_IP As Long = 6         ' 목적지IP
Private Const RCOL_DIRECTION As Long = 10      ' 방향
Private Const RCOL_VALID_STATUS As Long = 15   ' validation_status
Private Const RCOL_VALID_MSG As Long = 16      ' validation_message
Private Const RCOL_MATCH As Long = 17          ' match_details
Private Const RCOL_FW_PATH As Long = 18        ' firewall_path
Private Const RCOL_SRC_ZONE As Long = 19       ' source_zone
Private Const RCOL_DST_ZONE As Long = 20       ' destination_zone
Private Const RCOL_ZONE_PATH As Long = 21      ' zone_path

' module-level caches and graph state (rebuilt each run)
Private mGraph As Object          ' src_zone -> Collection of edge dicts
Private mNetworks As Collection   ' network dicts
Private mEnabledFw As Object      ' firewall_name -> True
Private mZoneCache As Object      ' ip -> zone
Private mPathCache As Object      ' "start|end" -> result dict
Private mFallback As Boolean

' ===================================================================== '
' Public entry point
' ===================================================================== '
Public Sub AnalyzeRequestRoutes()
    Dim requestsSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim srcIp As String, dstIp As String, direction As String
    Dim res As Object
    Dim analyzed As Long

    Set requestsSheet = ThisWorkbook.Worksheets(REQUESTS_SHEET)
    LoadRouteData

    lastRow = requestsSheet.Cells(requestsSheet.Rows.Count, RCOL_SOURCE_IP).End(xlUp).Row
    Dim lastByDst As Long
    lastByDst = requestsSheet.Cells(requestsSheet.Rows.Count, RCOL_DEST_IP).End(xlUp).Row
    If lastByDst > lastRow Then lastRow = lastByDst

    For rowIndex = 2 To lastRow
        srcIp = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_SOURCE_IP).Value))
        dstIp = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_DEST_IP).Value))
        direction = Trim$(CStr(requestsSheet.Cells(rowIndex, RCOL_DIRECTION).Value))
        If Len(srcIp) > 0 Or Len(dstIp) > 0 Then
            Set res = AnalyzeRoute(srcIp, dstIp, direction)
            WriteResultRow requestsSheet, rowIndex, res
            analyzed = analyzed + 1
        End If
    Next rowIndex

    MsgBox CStr(analyzed) & "건의 신청서 경로를 분석했습니다.", vbInformation
End Sub

Private Sub WriteResultRow(ByVal sheet As Worksheet, ByVal rowIndex As Long, ByVal res As Object)
    sheet.Cells(rowIndex, RCOL_TARGET).Value = res("target_firewalls")
    sheet.Cells(rowIndex, RCOL_FW_PATH).Value = res("firewall_path")
    sheet.Cells(rowIndex, RCOL_SRC_ZONE).Value = res("source_zone")
    sheet.Cells(rowIndex, RCOL_DST_ZONE).Value = res("destination_zone")
    sheet.Cells(rowIndex, RCOL_ZONE_PATH).Value = res("zone_path")
    sheet.Cells(rowIndex, RCOL_VALID_STATUS).Value = res("status")
    sheet.Cells(rowIndex, RCOL_VALID_MSG).Value = res("validation_message")
    sheet.Cells(rowIndex, RCOL_MATCH).Value = res("match_details")

    Dim st As String
    st = res("status")
    Dim c As Range
    Set c = sheet.Cells(rowIndex, RCOL_VALID_STATUS)
    If st = "OK" Then
        c.Interior.Color = RGB(198, 239, 206)
    ElseIf st = "MULTI_PATH" Or st = "LEGACY_FALLBACK" Or st = "INTRA_ZONE" Then
        c.Interior.Color = RGB(255, 235, 156)
    Else
        c.Interior.Color = RGB(255, 199, 206)
    End If
End Sub

' ===================================================================== '
' Data loading
' ===================================================================== '
Private Sub LoadRouteData()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set mNetworks = New Collection
    Set mEnabledFw = CreateObject("Scripting.Dictionary")
    Set mGraph = CreateObject("Scripting.Dictionary")
    Set mZoneCache = CreateObject("Scripting.Dictionary")
    Set mPathCache = CreateObject("Scripting.Dictionary")
    mFallback = ReadFallbackToggle()

    ' networks
    Set ws = ThisWorkbook.Worksheets(NETWORK_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Len(Trim$(CStr(ws.Cells(i, 1).Value))) > 0 Then
            If IsEnabled(ws.Cells(i, 5).Value) Then
                Dim n As Object
                Set n = CreateObject("Scripting.Dictionary")
                n("network_cidr") = Trim$(CStr(ws.Cells(i, 2).Value))
                n("zone") = Trim$(CStr(ws.Cells(i, 3).Value))
                mNetworks.Add n
            End If
        End If
    Next i

    ' firewalls (collect enabled set + per-firewall inside/outside CIDRs)
    Dim fwInside As Object   ' firewall_name -> inside_cidr string
    Dim fwOutside As Object  ' firewall_name -> outside_cidr string
    Set fwInside = CreateObject("Scripting.Dictionary")
    Set fwOutside = CreateObject("Scripting.Dictionary")
    Dim fwOrder As Object    ' firewall_name -> 1-based row ordinal
    Set fwOrder = CreateObject("Scripting.Dictionary")
    Dim fwOrdinal As Long
    fwOrdinal = 0
    Set ws = ThisWorkbook.Worksheets(FIREWALLS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim fwName As String
        fwName = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(fwName) > 0 Then
            ' ordinal counts EVERY non-empty firewall row (enabled or not), to
            ' mirror Python enumerate(self.firewalls, start=1).
            fwOrdinal = fwOrdinal + 1
            Dim fwEnabled As Boolean
            fwEnabled = IsEnabled(ws.Cells(i, 3).Value)
            If fwEnabled Then
                mEnabledFw(fwName) = True
                If Not fwOrder.Exists(fwName) Then fwOrder(fwName) = fwOrdinal
                fwInside(fwName) = Trim$(CStr(ws.Cells(i, 4).Value))    ' inside_cidr
                fwOutside(fwName) = Trim$(CStr(ws.Cells(i, 5).Value))   ' outside_cidr
            End If
        End If
    Next i

    ' routing_paths -> graph (explicit model is authoritative when present)
    Dim explicitCount As Long
    explicitCount = 0
    Set ws = ThisWorkbook.Worksheets(ROUTING_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim rfw As String, sz As String, dz As String
        rfw = Trim$(CStr(ws.Cells(i, 1).Value))
        sz = Trim$(CStr(ws.Cells(i, 2).Value))
        dz = Trim$(CStr(ws.Cells(i, 3).Value))
        If Len(rfw) > 0 And Len(sz) > 0 And Len(dz) > 0 Then
            If IsEnabled(ws.Cells(i, 7).Value) Then
                If mEnabledFw.Exists(rfw) Then
                    Dim edge As Object
                    Set edge = CreateObject("Scripting.Dictionary")
                    edge("firewall_name") = rfw
                    edge("src_zone") = sz
                    edge("dst_zone") = dz
                    edge("ingress_if") = Trim$(CStr(ws.Cells(i, 4).Value))
                    edge("egress_if") = Trim$(CStr(ws.Cells(i, 5).Value))
                    edge("path_order") = ParsePathOrder(ws.Cells(i, 6).Value)
                    If Not mGraph.Exists(sz) Then mGraph.Add sz, New Collection
                    mGraph(sz).Add edge
                    explicitCount = explicitCount + 1
                End If
            End If
        End If
    Next i

    ' no explicit routing_paths -> auto-derive from firewalls inside/outside CIDRs:
    ' each distinct CIDR becomes a zone "cidr:<base/prefix>"; a firewall links its
    ' inside-zone(s) <-> outside-zone(s). A CIDR shared by two firewalls (transit)
    ' becomes ONE zone, chaining them into a multi-hop path. The CIDR is also added
    ' as a network so request IPs resolve to that zone.
    If explicitCount = 0 Then
        ' auto mode: derive zones from firewall inside/outside CIDRs into a temp
        ' network set. Only if any CIDR was derived do we replace mNetworks (so the
        ' legacy zones don't collide); pure zone-graph setups with no inside/outside
        ' keep network_definitions. Parity with oracle (auto and derived_nets).
        Dim derivedNets As Collection
        Set derivedNets = New Collection
        Dim seenNet As Object
        Set seenNet = CreateObject("Scripting.Dictionary")
        Dim fwKey As Variant
        For Each fwKey In fwInside.Keys
            Dim insArr() As String, outArr() As String
            insArr = SplitCidrList(CStr(fwInside(fwKey)))
            outArr = SplitCidrList(CStr(fwOutside(fwKey)))
            Dim ci As Long
            For ci = LBound(insArr) To UBound(insArr)
                AddCidrNetworkTo derivedNets, insArr(ci), seenNet
            Next ci
            For ci = LBound(outArr) To UBound(outArr)
                AddCidrNetworkTo derivedNets, outArr(ci), seenNet
            Next ci
            ' link inside <-> outside through this firewall (both directions)
            Dim a As Long, b As Long
            For a = LBound(insArr) To UBound(insArr)
                Dim sz2 As String
                sz2 = CanonZone(insArr(a))
                If Len(sz2) > 0 Then
                    For b = LBound(outArr) To UBound(outArr)
                        Dim dz2 As String
                        dz2 = CanonZone(outArr(b))
                        If Len(dz2) > 0 And sz2 <> dz2 Then
                            AddDerivedEdge CStr(fwKey), sz2, dz2, CLng(fwOrder(fwKey))
                            AddDerivedEdge CStr(fwKey), dz2, sz2, CLng(fwOrder(fwKey))
                        End If
                    Next b
                End If
            Next a
        Next fwKey
        If derivedNets.Count > 0 Then Set mNetworks = derivedNets
    End If

    ' sort each adjacency list by edge key
    Dim key As Variant
    For Each key In mGraph.Keys
        SortEdges mGraph(key)
    Next key
End Sub

Private Function ReadFallbackToggle() As Boolean
    Dim ws As Worksheet, lastRow As Long, i As Long
    On Error GoTo done
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If LCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = "route_legacy_fallback" Then
            ReadFallbackToggle = IsEnabled(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
done:
End Function

Private Function IsEnabled(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    If s = "" Then
        IsEnabled = True   ' default enabled when blank
    ElseIf s = "Y" Or s = "YES" Or s = "TRUE" Or s = "1" Then
        IsEnabled = True
    Else
        IsEnabled = False
    End If
End Function

Private Function ParsePathOrder(ByVal v As Variant) As Long
    If Len(Trim$(CStr(v))) = 0 Or Not IsNumeric(v) Then
        ParsePathOrder = 999999
    Else
        ParsePathOrder = CLng(v)
    End If
End Function

' Split an inside/outside CIDR cell into tokens (';' , ',' , whitespace).
Private Function SplitCidrList(ByVal raw As String) As String()
    SplitCidrList = SplitAddressList(raw)
End Function

' Canonical zone label for one CIDR/IP: "cidr:<base>/<prefix>". '' if invalid.
' Mirrors tests/route_oracle.py _canon_zone so derived zones match Python.
Private Function CanonZone(ByVal cidr As String) As String
    Dim c As String
    c = Trim$(cidr)
    If Len(c) = 0 Then Exit Function
    On Error GoTo bad
    Dim base As Double, prefix As Long
    base = CidrStart(c)
    prefix = CidrPrefixLength(c)
    On Error GoTo 0
    Dim o0 As Long, o1 As Long, o2 As Long, o3 As Long
    o0 = Int(base / 16777216#)
    o1 = Int((base - o0 * 16777216#) / 65536#)
    o2 = Int((base - o0 * 16777216# - o1 * 65536#) / 256#)
    o3 = base - o0 * 16777216# - o1 * 65536# - o2 * 256#
    CanonZone = "cidr:" & o0 & "." & o1 & "." & o2 & "." & o3 & "/" & prefix
    Exit Function
bad:
    On Error GoTo 0
    CanonZone = ""
End Function

' Register a CIDR as a network (zone = its canonical label) once.
Private Sub AddCidrNetworkTo(ByVal coll As Collection, ByVal cidr As String, ByVal seenNet As Object)
    Dim z As String
    z = CanonZone(cidr)
    If Len(z) = 0 Then Exit Sub
    If seenNet.Exists(z) Then Exit Sub
    seenNet(z) = True
    Dim n As Object
    Set n = CreateObject("Scripting.Dictionary")
    n("network_cidr") = Trim$(cidr)
    n("zone") = z
    coll.Add n
End Sub

' Add one directed derived edge to the graph.
Private Sub AddDerivedEdge(ByVal fwName As String, ByVal sz As String, ByVal dz As String, ByVal pathOrder As Long)
    Dim de As Object
    Set de = CreateObject("Scripting.Dictionary")
    de("firewall_name") = fwName
    de("src_zone") = sz
    de("dst_zone") = dz
    de("ingress_if") = ""
    de("egress_if") = ""
    de("path_order") = pathOrder
    If Not mGraph.Exists(sz) Then mGraph.Add sz, New Collection
    mGraph(sz).Add de
End Sub

' ===================================================================== '
' Zone resolution (longest-prefix match)
' ===================================================================== '
Private Function ResolveZone(ByVal ipText As String) As String
    ' Value may be a single IP, a CIDR (대역), or a delimited list of either.
    ' Each token is resolved by longest-prefix overlap; first resolvable wins.
    If mZoneCache.Exists(ipText) Then
        ResolveZone = mZoneCache(ipText)
        Exit Function
    End If

    Dim result As String
    Dim tokens() As String
    Dim i As Long
    Dim z As String
    result = "#UNRESOLVED"
    tokens = SplitAddressList(ipText)
    For i = LBound(tokens) To UBound(tokens)
        If Len(Trim$(tokens(i))) > 0 Then
            z = ResolveTokenZone(Trim$(tokens(i)))
            If z <> "#UNRESOLVED" Then
                result = z
                Exit For
            End If
        End If
    Next i

    mZoneCache(ipText) = result
    ResolveZone = result
End Function

Private Function ResolveTokenZone(ByVal token As String) As String
    Dim tokStart As Double, tokEnd As Double
    On Error GoTo unresolved
    tokStart = CidrStart(token)
    tokEnd = CidrEnd(token)
    On Error GoTo 0

    Dim bestPrefix As Long
    Dim bestZone As String
    Dim ambiguous As Boolean
    Dim i As Long
    bestPrefix = -1
    bestZone = "#UNRESOLVED"
    ambiguous = False

    For i = 1 To mNetworks.Count
        Dim n As Object
        Set n = mNetworks(i)
        Dim cidr As String
        cidr = n("network_cidr")
        Dim st As Double, en As Double, pfx As Long
        On Error GoTo skip
        st = CidrStart(cidr)
        en = CidrEnd(cidr)
        pfx = CidrPrefixLength(cidr)
        On Error GoTo 0
        If st <= tokEnd And tokStart <= en Then
            If pfx > bestPrefix Then
                bestPrefix = pfx
                bestZone = n("zone")
                ambiguous = False
            ElseIf pfx = bestPrefix And n("zone") <> bestZone Then
                ambiguous = True
            End If
        End If
        GoTo continue
skip:
        On Error GoTo 0
continue:
    Next i

    If ambiguous Then
        ResolveTokenZone = "#AMBIGUOUS"
    Else
        ResolveTokenZone = bestZone
    End If
    Exit Function

unresolved:
    On Error GoTo 0
    ResolveTokenZone = "#UNRESOLVED"
End Function

' ===================================================================== '
' Edge / path keys + sort
' ===================================================================== '
Private Function EdgeKey(ByVal edge As Object) As String
    EdgeKey = Format$(CLng(edge("path_order")), "000000") & "|" & _
              edge("firewall_name") & "|" & edge("dst_zone")
End Function

Private Function PathKey(ByVal edges As Collection) As String
    Dim s As String, i As Long
    For i = 1 To edges.Count
        If i > 1 Then s = s & ";"
        s = s & EdgeKey(edges(i))
    Next i
    PathKey = s
End Function

Private Sub SortEdges(ByVal edges As Collection)
    ' selection sort by EdgeKey into a new ordered collection (small lists)
    Dim count As Long
    count = edges.count
    If count < 2 Then Exit Sub

    Dim taken() As Boolean
    ReDim taken(1 To count)
    Dim ordered As Collection
    Set ordered = New Collection

    Dim picked As Long, i As Long
    Dim n As Long
    For n = 1 To count
        picked = 0
        Dim bestKey As String
        bestKey = ""
        For i = 1 To count
            If Not taken(i) Then
                Dim k As String
                k = EdgeKey(edges(i))
                If picked = 0 Or k < bestKey Then
                    picked = i
                    bestKey = k
                End If
            End If
        Next i
        taken(picked) = True
        ordered.Add edges(picked)
    Next n

    ' replace contents of edges with ordered
    Do While edges.count > 0
        edges.Remove 1
    Loop
    For i = 1 To ordered.count
        edges.Add ordered(i)
    Next i
End Sub

' ===================================================================== '
' BFS shortest zone paths
' ===================================================================== '
Private Function FindShortestPaths(ByVal startZone As String, ByVal endZone As String) As Collection
    Dim results As Collection
    Set results = New Collection
    If startZone = endZone Then
        Set FindShortestPaths = results
        Exit Function
    End If

    Dim queue As Collection
    Set queue = New Collection

    Dim initVisited As Object
    Set initVisited = CreateObject("Scripting.Dictionary")
    initVisited(startZone) = True

    Dim initState As Object
    Set initState = CreateObject("Scripting.Dictionary")
    initState("zone") = startZone
    Set initState("edges") = New Collection
    Set initState("visited") = initVisited
    queue.Add initState

    Dim shortestLen As Long
    shortestLen = -1

    Do While queue.Count > 0
        Dim state As Object
        Set state = queue(1)
        queue.Remove 1

        Dim curEdges As Collection
        Set curEdges = state("edges")
        If shortestLen <> -1 And curEdges.Count >= shortestLen Then GoTo continueLoop

        Dim zone As String
        zone = state("zone")
        If mGraph.Exists(zone) Then
            Dim adj As Collection
            Set adj = mGraph(zone)
            Dim i As Long
            For i = 1 To adj.Count
                Dim edge As Object
                Set edge = adj(i)
                Dim nxt As String
                nxt = edge("dst_zone")
                Dim visited As Object
                Set visited = state("visited")
                If Not visited.Exists(nxt) Then
                    Dim newEdges As Collection
                    Set newEdges = CloneEdges(curEdges)
                    newEdges.Add edge
                    If nxt = endZone Then
                        If shortestLen = -1 Then shortestLen = newEdges.Count
                        If newEdges.Count = shortestLen Then results.Add newEdges
                    Else
                        If shortestLen = -1 Or newEdges.Count < shortestLen Then
                            Dim newVisited As Object
                            Set newVisited = CloneDict(visited)
                            newVisited(nxt) = True
                            Dim newState As Object
                            Set newState = CreateObject("Scripting.Dictionary")
                            newState("zone") = nxt
                            Set newState("edges") = newEdges
                            Set newState("visited") = newVisited
                            queue.Add newState
                        End If
                    End If
                End If
            Next i
        End If
continueLoop:
    Loop

    Set FindShortestPaths = results
End Function

Private Function ChooseBest(ByVal paths As Collection) As Collection
    Dim best As Collection
    Set best = paths(1)
    Dim bestKey As String
    bestKey = PathKey(best)
    Dim i As Long
    For i = 2 To paths.Count
        Dim k As String
        k = PathKey(paths(i))
        If k < bestKey Then
            Set best = paths(i)
            bestKey = k
        End If
    Next i
    Set ChooseBest = best
End Function

Private Function CloneEdges(ByVal src As Collection) As Collection
    Dim r As Collection
    Set r = New Collection
    Dim i As Long
    For i = 1 To src.Count
        r.Add src(i)
    Next i
    Set CloneEdges = r
End Function

Private Function CloneDict(ByVal src As Object) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In src.Keys
        r(k) = src(k)
    Next k
    Set CloneDict = r
End Function

' ===================================================================== '
' Directed resolve + caching
' ===================================================================== '
Private Function DirectedResolve(ByVal startZone As String, ByVal endZone As String) As Object
    Dim cacheKey As String
    cacheKey = startZone & "|" & endZone
    If mPathCache.Exists(cacheKey) Then
        Set DirectedResolve = mPathCache(cacheKey)
        Exit Function
    End If

    Dim res As Object
    Set res = NewResult()
    Dim paths As Collection
    Set paths = FindShortestPaths(startZone, endZone)

    If paths.Count = 0 Then
        res("status") = "NO_PATH"
        res("validation_message") = "No routing path found"
        res("match_details") = "from=" & ZoneDisplay(startZone) & "; to=" & ZoneDisplay(endZone)
    Else
        Dim best As Collection
        Set best = ChooseBest(paths)
        res("firewall_path") = BuildFwPath(best)
        res("target_firewalls") = BuildTargetSet(best)
        res("zone_path") = BuildZonePath(startZone, best)
        res("match_details") = BuildMatchDetails(startZone, best)
        res("path_count") = paths.Count
        If paths.Count = 1 Then
            res("status") = "OK"
            res("validation_message") = "Path resolved"
        Else
            res("status") = "MULTI_PATH"
            res("validation_message") = "Multiple equal shortest paths; selected by path_order/firewall_name"
        End If
    End If

    mPathCache(cacheKey) = res
    Set DirectedResolve = res
End Function

Private Function BuildFwPath(ByVal edges As Collection) As String
    Dim s As String, i As Long
    For i = 1 To edges.Count
        If i > 1 Then s = s & ">"
        s = s & edges(i)("firewall_name")
    Next i
    BuildFwPath = s
End Function

Private Function BuildTargetSet(ByVal edges As Collection) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim s As String, i As Long
    For i = 1 To edges.Count
        Dim fw As String
        fw = edges(i)("firewall_name")
        If Not seen.Exists(fw) Then
            seen(fw) = True
            If Len(s) > 0 Then s = s & ";"
            s = s & fw
        End If
    Next i
    BuildTargetSet = s
End Function

Private Function BuildZonePath(ByVal startZone As String, ByVal edges As Collection) As String
    Dim s As String, i As Long
    s = ZoneDisplay(startZone)
    For i = 1 To edges.Count
        s = s & ">" & ZoneDisplay(edges(i)("dst_zone"))
    Next i
    BuildZonePath = s
End Function

Private Function BuildMatchDetails(ByVal startZone As String, ByVal edges As Collection) As String
    Dim s As String, cur As String, i As Long
    cur = startZone
    For i = 1 To edges.Count
        If i > 1 Then s = s & "; "
        Dim e As Object
        Set e = edges(i)
        s = s & ZoneDisplay(cur) & " -> " & ZoneDisplay(e("dst_zone")) & _
            " (" & e("firewall_name") & ")"
        cur = e("dst_zone")
    Next i
    BuildMatchDetails = s
End Function

' DISPLAY-ONLY zone label (mirror of tests/route_oracle.py _zone_display).
' Never used in graph/cache/tie-break/status; only when building output strings.
Private Function ZoneDisplay(ByVal zone As String) As String
    If zone = "cidr:0.0.0.0/0" Or zone = "0.0.0.0/0" Then
        ZoneDisplay = "외부"
    ElseIf Left$(zone, 5) = "cidr:" Then
        ZoneDisplay = Mid$(zone, 6)
    Else
        ZoneDisplay = zone
    End If
End Function

' ===================================================================== '
' Legacy fallback (CIDR overlap against network_definitions zones)
' ===================================================================== '
Private Function LegacyFallback(ByVal srcIp As String, ByVal dstIp As String) As Object
    Dim res As Object
    Set res = NewResult()
    Dim targets As Object
    Set targets = CreateObject("Scripting.Dictionary")
    Dim ordered As String

    Dim zone As Variant
    For Each zone In mGraph.Keys
        Dim adj As Collection
        Set adj = mGraph(zone)
        Dim i As Long
        For i = 1 To adj.Count
            Dim edge As Object
            Set edge = adj(i)
            Dim fw As String
            fw = edge("firewall_name")
            If Not targets.Exists(fw) Then
                If ZonesOverlapIp(edge("src_zone"), edge("dst_zone"), srcIp, dstIp) Then
                    targets(fw) = True
                    If Len(ordered) > 0 Then ordered = ordered & ";"
                    ordered = ordered & fw
                End If
            End If
        Next i
    Next zone

    If Len(ordered) > 0 Then
        res("status") = "LEGACY_FALLBACK"
        res("target_firewalls") = ordered
        res("firewall_path") = Replace(ordered, ";", ">")
        res("validation_message") = "Resolved using legacy CIDR-overlap fallback"
        res("match_details") = "fallback by network overlap"
    Else
        res("status") = "NO_PATH"
        res("validation_message") = "No routing path found"
    End If
    Set LegacyFallback = res
End Function

Private Function ZonesOverlapIp(ByVal zoneA As String, ByVal zoneB As String, ByVal srcIp As String, ByVal dstIp As String) As Boolean
    Dim i As Long
    For i = 1 To mNetworks.Count
        Dim n As Object
        Set n = mNetworks(i)
        If n("zone") = zoneA Or n("zone") = zoneB Then
            If NetworkRangesOverlap(srcIp, n("network_cidr")) Or NetworkRangesOverlap(dstIp, n("network_cidr")) Then
                ZonesOverlapIp = True
                Exit Function
            End If
        End If
    Next i
End Function

' ===================================================================== '
' Main route analysis
' ===================================================================== '
Public Function AnalyzeRoute(ByVal srcIp As String, ByVal dstIp As String, ByVal direction As String) As Object
    Dim d As String
    d = NormalizeDirection(direction)
    Dim res As Object

    If d = "#INVALID" Then
        Set res = NewResult()
        res("status") = "DIRECTION_MISMATCH"
        res("validation_message") = "Invalid direction: " & direction
        Set AnalyzeRoute = res
        Exit Function
    End If

    Dim srcZone As String, dstZone As String
    srcZone = ResolveZone(srcIp)
    dstZone = ResolveZone(dstIp)

    If srcZone = "#UNRESOLVED" Or srcZone = "#AMBIGUOUS" Or _
       dstZone = "#UNRESOLVED" Or dstZone = "#AMBIGUOUS" Then
        Set res = NewResult()
        res("status") = "ZONE_UNRESOLVED"
        Dim msg As String
        If srcZone = "#UNRESOLVED" Or srcZone = "#AMBIGUOUS" Then msg = "Source zone unresolved (" & srcIp & ")"
        If dstZone = "#UNRESOLVED" Or dstZone = "#AMBIGUOUS" Then
            If Len(msg) > 0 Then msg = msg & "; "
            msg = msg & "Destination zone unresolved (" & dstIp & ")"
        End If
        res("validation_message") = msg
        res("match_details") = "source_ip=" & srcIp & "; destination_ip=" & dstIp
        If Left$(srcZone, 1) <> "#" Then res("source_zone") = ZoneDisplay(srcZone)
        If Left$(dstZone, 1) <> "#" Then res("destination_zone") = ZoneDisplay(dstZone)
        Set AnalyzeRoute = res
        Exit Function
    End If

    If srcZone = dstZone Then
        Set res = NewResult()
        res("status") = "INTRA_ZONE"
        res("source_zone") = ZoneDisplay(srcZone)
        res("destination_zone") = ZoneDisplay(dstZone)
        res("validation_message") = "Source and destination in same zone; no firewall path required"
        res("match_details") = "source_zone=" & ZoneDisplay(srcZone) & "; destination_zone=" & ZoneDisplay(dstZone)
        Set AnalyzeRoute = res
        Exit Function
    End If

    If d = "OUT" Then
        Set res = CopyResult(DirectedResolve(srcZone, dstZone))
        If res("status") = "NO_PATH" Then
            Dim rev As Object
            Set rev = DirectedResolve(dstZone, srcZone)
            If rev("status") = "OK" Or rev("status") = "MULTI_PATH" Then
                res("status") = "DIRECTION_MISMATCH"
                res("validation_message") = "Requested OUT path not found, but reverse path exists"
                res("match_details") = rev("match_details")
                res("firewall_path") = ""
                res("target_firewalls") = ""
                res("zone_path") = ""
            End If
        End If
    ElseIf d = "IN" Then
        Set res = CopyResult(DirectedResolve(dstZone, srcZone))
        If res("status") = "NO_PATH" Then
            Dim fwd As Object
            Set fwd = DirectedResolve(srcZone, dstZone)
            If fwd("status") = "OK" Or fwd("status") = "MULTI_PATH" Then
                res("status") = "DIRECTION_MISMATCH"
                res("validation_message") = "Requested IN path not found, but opposite path exists"
                res("match_details") = fwd("match_details")
                res("firewall_path") = ""
                res("target_firewalls") = ""
                res("zone_path") = ""
            End If
        End If
    Else  ' BOTH
        Dim f2 As Object
        Set f2 = DirectedResolve(srcZone, dstZone)
        If f2("status") = "OK" Or f2("status") = "MULTI_PATH" Then
            Set res = CopyResult(f2)
        Else
            Dim r2 As Object
            Set r2 = DirectedResolve(dstZone, srcZone)
            If r2("status") = "OK" Or r2("status") = "MULTI_PATH" Then
                Set res = CopyResult(r2)
                res("validation_message") = res("validation_message") & "; resolved using reverse direction under BOTH"
            Else
                Set res = CopyResult(f2)
            End If
        End If
    End If

    If res("status") = "NO_PATH" And mFallback Then
        Dim fb As Object
        Set fb = LegacyFallback(srcIp, dstIp)
        If fb("status") = "LEGACY_FALLBACK" Then Set res = fb
    End If

    res("source_zone") = ZoneDisplay(srcZone)
    res("destination_zone") = ZoneDisplay(dstZone)
    Set AnalyzeRoute = res
End Function

Private Function NormalizeDirection(ByVal direction As String) As String
    Dim d As String
    d = UCase$(Trim$(direction))
    If d = "" Then
        NormalizeDirection = "BOTH"
    ElseIf d = "IN" Or d = "OUT" Or d = "BOTH" Then
        NormalizeDirection = d
    Else
        NormalizeDirection = "#INVALID"
    End If
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

Private Function CopyResult(ByVal src As Object) As Object
    Dim r As Object
    Set r = NewResult()
    Dim k As Variant
    For Each k In src.Keys
        r(k) = src(k)
    Next k
    Set CopyResult = r
End Function

' ===================================================================== '
' IPv4 / CIDR primitives (mirror FirewallPolicyAutomation; self-contained)
' ===================================================================== '
Private Function SplitAddressList(ByVal addressList As String) As String()
    Dim normalized As String
    normalized = Replace(addressList, ChrW(160), " ")
    normalized = Replace(normalized, vbTab, " ")
    normalized = Replace(normalized, vbCrLf, ";")
    normalized = Replace(normalized, vbCr, ";")
    normalized = Replace(normalized, vbLf, ";")
    normalized = Replace(normalized, ",", ";")
    normalized = Replace(normalized, ChrW(65292), ";")
    normalized = Replace(normalized, ChrW(65307), ";")
    ' runs of ASCII spaces between tokens are also a separator (mirror Python
    ' split_address_list / request _norm_list): CIDR/IP tokens never contain a
    ' space, so 'a b' -> 'a;b'. Empty tokens from repeated separators are dropped.
    normalized = Replace(normalized, " ", ";")
    Dim raw() As String
    raw = Split(normalized, ";")
    Dim outList() As String
    ReDim outList(0 To UBound(raw))
    Dim i As Long, n As Long
    n = 0
    For i = LBound(raw) To UBound(raw)
        If Len(Trim$(raw(i))) > 0 Then
            outList(n) = Trim$(raw(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        SplitAddressList = Split(vbNullString, ";")  ' empty array (UBound -1)
    Else
        ReDim Preserve outList(0 To n - 1)
        SplitAddressList = outList
    End If
End Function

Private Function IpToNumber(ByVal ipText As String) As Double
    Dim parts() As String, i As Long, octet As Long
    parts = Split(Trim$(ipText), ".")
    If UBound(parts) <> 3 Then Err.Raise vbObjectError + 1000, , "Invalid IPv4 address: " & ipText
    For i = 0 To 3
        If Not IsAllDigits(Trim$(parts(i))) Then Err.Raise vbObjectError + 1000, , "Invalid IPv4 address: " & ipText
        octet = CLng(Trim$(parts(i)))
        If octet < 0 Or octet > 255 Then Err.Raise vbObjectError + 1000, , "Invalid IPv4 address: " & ipText
    Next i
    IpToNumber = CDbl(Trim$(parts(0))) * 16777216# + CDbl(Trim$(parts(1))) * 65536# + CDbl(Trim$(parts(2))) * 256# + CDbl(Trim$(parts(3)))
End Function

' True only for a non-empty run of ASCII digits 0-9 (mirrors Python str.isdigit
' for the IPv4/prefix use: rejects '+1', '1e1', '1.0', '', whitespace).
Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function CidrPrefixLength(ByVal cidrText As String) As Long
    Dim p() As String
    p = Split(cidrText, "/")
    If UBound(p) = 0 Then
        CidrPrefixLength = 32
    Else
        If Not IsAllDigits(Trim$(p(1))) Then Err.Raise vbObjectError + 1004, , "Invalid CIDR prefix: " & cidrText
        CidrPrefixLength = CLng(Trim$(p(1)))
    End If
    If CidrPrefixLength < 0 Or CidrPrefixLength > 32 Then Err.Raise vbObjectError + 1004, , "Invalid CIDR prefix: " & cidrText
End Function

Private Function CidrBaseIp(ByVal cidrText As String) As String
    CidrBaseIp = Trim$(Split(cidrText, "/")(0))
End Function

Private Function CidrBlockSize(ByVal cidrText As String) As Double
    CidrBlockSize = 2 ^ (32 - CidrPrefixLength(cidrText))
End Function

Private Function CidrStart(ByVal cidrText As String) As Double
    Dim baseValue As Double, blockSize As Double
    baseValue = IpToNumber(CidrBaseIp(cidrText))
    blockSize = CidrBlockSize(cidrText)
    CidrStart = Fix(baseValue / blockSize) * blockSize
End Function

Private Function CidrEnd(ByVal cidrText As String) As Double
    CidrEnd = CidrStart(cidrText) + CidrBlockSize(cidrText) - 1
End Function

Private Function NetworkRangesOverlap(ByVal leftCidr As String, ByVal rightCidr As String) As Boolean
    Dim ls As Double, le As Double, rs As Double, re As Double
    On Error GoTo nomatch
    If Len(leftCidr) = 0 Or Len(rightCidr) = 0 Then Exit Function
    ls = CidrStart(leftCidr): le = CidrEnd(leftCidr)
    rs = CidrStart(rightCidr): re = CidrEnd(rightCidr)
    NetworkRangesOverlap = ls <= re And rs <= le
    Exit Function
nomatch:
    NetworkRangesOverlap = False
End Function
