Attribute VB_Name = "GET_JSON_DATA"
Option Explicit
Option Compare Text
Private Function XMLHTTPs(Optional Server As Boolean) As Object
    If Server Then
        Set XMLHTTPs = VBA.CreateObject("MSXML2.serverXMLHTTP.6.0")
        XMLHTTPs.SetTimeouts 7000, 7000, 9000, 9000
    Else
        Set XMLHTTPs = VBA.CreateObject("MSXML2.XMLHTTP.6.0")
    End If
End Function
Function ChangeType(temp As Variant) As Variant
    If IsNumeric(temp) Then
        ChangeType = temp
    Else
        ChangeType = "" & temp
    End If
End Function

' Du lieu gia vnindex - SSI
Function VNINDEX(Optional RangeTime = "6M") As Variant
    Dim tempArr, Value, Res As Variant
    Dim JSON As Object
    Dim Item As Dictionary
    Dim i, K As Integer
    Dim URL, UT As String
    Select Case RangeTime
        Case "6M":
        UT = "SixMonths"
        Case "9M":
        UT = "NineMonths"
        Case "3M":
        UT = "ThreeMonths"
        Case "1Y":
        UT = "OneYear"
        Case "3Y":
        UT = "ThreeYears"
        Case Else
        UT = "AllTime"
    End Select
    URL = "https://fiin-market.ssi.com.vn/MarketInDepth/GetValuationSeriesV2?language=vi&Code=VNINDEX&TimeRange=" & UT
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("items").Count, 1 To 4)
    
    tempArr(0, 1) = "Date"
    tempArr(0, 2) = "Price"
    tempArr(0, 3) = "P/E"
    tempArr(0, 4) = "P/B"
    i = 1
    For Each Item In JSON("items")
        tempArr(i, 1) = CDate(Left(Item("tradingDate"), 10))
        tempArr(i, 2) = Item("value")
        tempArr(i, 3) = Item("r21")
        tempArr(i, 4) = Item("r25")
        i = i + 1
    Next
    
    VNINDEX = tempArr
ends:
    Exit Function
End Function

' Du lieu gia co phieu - 24hmoney.com
Public Function STOCKDATA(MaCK As Variant, StartDate As Date, Optional EndDate = 0) As Variant
    Dim tempArr As Variant
    Dim JSON, Res$, RE As Object
    Dim i, K As Integer
    Dim MaSIC, URL As String
    Dim TimeJ, closeJ, openJ, highJ, lowJ, volumeJ, Value As Variant
    
    If EndDate = 0 Then EndDate = Now
    
    URL = "https://api-common-t19.24hmoney.vn/web-hook/open-api/tradingview/history?symbol=" & UCase(MaCK) & _
    "&resolution=D&from_ts=" & Date2Unix(StartDate, True) & "&to_ts=" & Date2Unix(EndDate, True)
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
    
    Set TimeJ = JSON("t")
    Set closeJ = JSON("c")
    Set openJ = JSON("o")
    Set highJ = JSON("h")
    Set lowJ = JSON("l")
    Set volumeJ = JSON("v")
    
    K = 0
    For Each Value In TimeJ
        K = K + 1
    Next Value
    
    ReDim tempArr(0 To K, 1 To 6)
    tempArr(0, 1) = "Date"
    tempArr(0, 2) = "Volume"
    tempArr(0, 3) = "Open"
    tempArr(0, 4) = "High"
    tempArr(0, 5) = "Low"
    tempArr(0, 6) = "Close"
    For i = 1 To K
        tempArr(i, 1) = Unix2Date(TimeJ(i), False)
        tempArr(i, 3) = openJ(i)
        tempArr(i, 4) = highJ(i)
        tempArr(i, 5) = lowJ(i)
        tempArr(i, 6) = closeJ(i)
        tempArr(i, 2) = volumeJ(i)
    Next i
    STOCKDATA = tempArr
ends:
    Exit Function
End Function

' Danh sach ma chung khoan  - VNDIRECT
Function STOCKLIST() As Variant
    Dim tempArr, Value, Res As Variant
    Dim JSON As Object
    Dim Item As Dictionary
    Dim i, K As Integer
    Dim URL As String
    
    URL = "https://finfo-api.vndirect.com.vn/v4/stocks?q=type:stock,ifc~floor:HOSE,HNX,UPCOM&size=10000"
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    ReDim tempArr(JSON("data").Count, 1 To 8)
    
    ''Title
    tempArr(0, 1) = "STT"
    tempArr(0, 2) = "MCK"
    tempArr(0, 3) = "Name"
    tempArr(0, 4) = "English Name"
    tempArr(0, 5) = "Floor"
    tempArr(0, 6) = "Type"
    tempArr(0, 7) = "Niem Yet"
    tempArr(0, 8) = "Huy Niem Yet"
    i = 1
    For Each Item In JSON("data")
        tempArr(i, 1) = i
        tempArr(i, 2) = ChangeType(Item("code"))
        tempArr(i, 3) = ChangeType(Item("companyName"))
        tempArr(i, 4) = ChangeType(Item("companyNameEng"))
        tempArr(i, 5) = ChangeType(Item("floor"))
        tempArr(i, 6) = ChangeType(Item("type"))
        tempArr(i, 7) = ChangeType(Item("listedDate"))
        tempArr(i, 8) = ChangeType(Item("delistedDate"))
        i = i + 1
    Next
    STOCKLIST = tempArr
ends:
    Exit Function
End Function

' Danh sach ma chung khoan theo ma nganh nghe - 24hmoney
Function STOCKICB(icbcode As Variant, Optional stt = 0) As Variant
    Dim tempArr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K As Integer
    Dim URL As String
    If stt <> 0 Then stt = -1
    URL = "https://api-finance-t19.24hmoney.vn/v2/ios/stock-recommend/business-all?group_id=" & icbcode
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("data")("data").Count, 1 To 6 + stt)
    
    If stt = 0 Then tempArr(0, 1) = "STT"
    tempArr(0, 2 + stt) = "MCK"
    tempArr(0, 3 + stt) = "Name"
    tempArr(0, 4 + stt) = "Price"
    tempArr(0, 5 + stt) = "change"
    tempArr(0, 6 + stt) = "change%"
    
    i = 1
    For Each Item In JSON("data")("data")
        If stt = 0 Then tempArr(i, 1) = i
        tempArr(i, 2 + stt) = Item("symbol")
        tempArr(i, 3 + stt) = Item("company_name")
        tempArr(i, 4 + stt) = Item("price")
        tempArr(i, 5 + stt) = Item("change")
        tempArr(i, 6 + stt) = Item("change_percent")
        i = i + 1
    Next
    
    STOCKICB = tempArr
ends:
    Exit Function
End Function

' Danh sach cong ty co BCKQKD theo quy/nam - 24hmoney
Function FINANCIALREPORT(Year As Variant, Optional quy = 0, Optional Sapxep = 1, Optional page = 1, Optional per_page = 2000) As Variant
    Dim tempArr, arr, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, j, K As Integer
    Dim URL, tmp2, sx As String
    On Error Resume Next
    
    Select Case Sapxep
        Case 1:
        sx = "profit_after_tax_tndn_percent"
        Case 2:
        sx = "roe"
        Case Else
        sx = "profit_after_tax_tndn"
    End Select
    
    URL = "https://api-finance-t19.24hmoney.vn/v1/web/company/financial-report-filter?year=" _
    & Year & "&quarter=" & quy & "&floor=all&symbol=&key=" & sx & "&sort=desc&page=" & page & "&per_page=" & per_page
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("data").Count, 1 To 10)
    
    tempArr(0, 1) = "MaCK"
    tempArr(0, 2) = "San"
    tempArr(0, 3) = "Ten Cong Ty - " & IIf(quy <> 0, "Quy " & quy & "/", "Nam ") & Year
    tempArr(0, 4) = "LNST (ty)"
    tempArr(0, 5) = "LNST (%YoY)"
    tempArr(0, 6) = "EPS 4Q"
    tempArr(0, 7) = "ROA (%)"
    tempArr(0, 8) = "ROE (%)"
    tempArr(0, 9) = "P/E "
    tempArr(0, 10) = "P/B "
    
    i = 1
    For Each tmp In JSON("data")
        tempArr(i, 1) = ChangeType(tmp("symbol"))
        tempArr(i, 2) = ChangeType(tmp("floor"))
        tempArr(i, 3) = ChangeType(tmp("short_name"))
        tempArr(i, 4) = ChangeType(tmp("profit_after_tax_tndn"))
        tempArr(i, 5) = ChangeType(tmp("profit_after_tax_tndn_percent"))
        tempArr(i, 6) = ChangeType(tmp("eps"))
        tempArr(i, 7) = CDec(Replace(tmp("roa"), ".", ","))
        tempArr(i, 8) = CDec(Replace(tmp("roe"), ".", ","))
        tempArr(i, 9) = CDec(Replace(tmp("pe"), ".", ","))
        tempArr(i, 10) = CDec(Replace(tmp("pb"), ".", ","))
        
        i = i + 1
    Next
    FINANCIALREPORT = tempArr
ends:
    Exit Function
End Function

' Thong tin co ban ve co phieu - 24hmoney.com
' https://api-finance-t19.24hmoney.vn/v2/ios/companies/index?locale=vi&browser_id=sweb1668255ycbln2baj85y18ac332867&symbol=HPG
Function STOCKNOTE(MCK As Variant) As Variant
    Dim tempArr, Value, Res As Variant
    Dim JSON As Object
    Dim Item As Dictionary
    Dim i, K As Integer
    Dim URL, MaSIC, tmp As String
    URL = "https://api-finance-t19.24hmoney.vn/v2/ios/companies/index?locale=vi&symbol=" & UCase(MCK)
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
    ReDim tempArr(1 To 15, 1 To 2)
    
    tempArr(15, 1) = "Group (" & JSON("data")("group_id") & ")"
    tempArr(15, 2) = JSON("data")("group_name")
    
    tempArr(1, 1) = "P/E"
    tempArr(1, 2) = JSON("data")("pe")
    
    tempArr(2, 1) = "P/B"
    tempArr(2, 2) = JSON("data")("pb")
    
    tempArr(3, 1) = "EPS " & JSON("data")("year")
    tempArr(3, 2) = CDec(Replace(JSON("data")("eps"), ".", ","))
    
    tempArr(12, 1) = "foreign_current_room (Million)"
    tempArr(12, 2) = JSON("data")("foreign_current_room") / 1000000
    
    tempArr(4, 1) = "EPS 4Q"
    tempArr(4, 2) = CDec(Replace(JSON("data")("eps4Q"), ".", ","))
    
    tempArr(13, 1) = "foreign_total_room (Million)"
    tempArr(13, 2) = JSON("data")("foreign_total_room") / 1000000
    
    tempArr(5, 1) = "SLCP niem yet (Million)"
    tempArr(5, 2) = JSON("data")("listed_share_vol") / 1000000
    
    tempArr(14, 1) = "foreign_current_room_percent"
    tempArr(14, 2) = JSON("data")("foreign_current_room_percent")
    
    tempArr(6, 1) = "SLCP luu hanh (Million)"
    tempArr(6, 2) = JSON("data")("circulation_vol") / 1000000
    
    tempArr(7, 1) = "ROE %"
    tmp = JSON("data")("roe")
    tempArr(7, 2) = CDec(Replace(tmp, ".", ","))
    
    tempArr(8, 1) = "ROA %"
    tmp = JSON("data")("roa")
    tempArr(8, 2) = CDec(Replace(tmp, ".", ","))
    
    tempArr(9, 1) = "Beta"
    tempArr(9, 2) = JSON("data")("the_beta")
    
    tempArr(10, 1) = "Free float"
    tempArr(10, 2) = JSON("data")("free_float")
    
    tempArr(11, 1) = "Free float %"
    tempArr(11, 2) = JSON("data")("free_float_rate")
    
    STOCKNOTE = tempArr
ends:
    Exit Function
End Function

' Thong ke EPS, PE - 24hmoney.com
Function STOCKEPS(MCK As Variant) As Variant
    Dim tempArr, arr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K As Integer
    Dim URL As String
    
    URL = "https://api-finance-t19.24hmoney.vn/v1/ios/company/financial-graph?symbol=" & MCK & "&graph_type=6"
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("data")("points").Count, 1 To 3)
    tempArr(0, 1) = "Lable"
    tempArr(0, 3) = "PE"
    tempArr(0, 2) = "EPS pha loang"
    
    i = 1
    For Each tmp In JSON("data")("x-axis")
        tempArr(i, 1) = ChangeType(tmp("name"))
        i = i + 1
    Next
    i = 1
    For Each tmp In JSON("data")("points")
        tempArr(i, 2) = ChangeType(tmp("y"))
        tempArr(i, 3) = ChangeType(tmp("y1"))
        i = i + 1
    Next
    STOCKEPS = tempArr
ends:
    Exit Function
End Function

' Lich su chi co tuc cua co phieu  - 24hmoney.com
Function STOCKDIVIDENT(MCK As Variant) As Variant
    Dim tempArr, arr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K, tmp2 As Integer
    Dim URL As String
    
    URL = "https://api-finance-t19.24hmoney.vn/v1/ios/company/dividend-schedule?symbol=" & MCK
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("data").Count, 1 To 2)
    tempArr(0, 1) = "Date"
    tempArr(0, 2) = "Divident Payout"
    
    i = 1
    For Each tmp In JSON("data")
        tempArr(i, 1) = ChangeType(tmp("end_date"))
        If ChangeType(tmp("type")) = 1 Then
            tempArr(i, 2) = "Tr" & ChrW(7843) & " c" & ChrW(7893) & " t" & ChrW(7913) & "c b" & ChrW(7857) & "ng ti" & ChrW(7873) & "n m" & ChrW(7863) & "t v" & ChrW(7899) & "i t" & ChrW(7927) & " l" & ChrW(7879) & " " & tmp("ratio") * 100 & "%"
        Else
            tempArr(i, 2) = "Tr" & ChrW(7843) & " c" & ChrW(7893) & " t" & ChrW(7913) & "c b" & ChrW(7857) & "ng c" & ChrW(7893) & " phi" & ChrW(7871) & "u v" & ChrW(7899) & "i t" & ChrW(7927) & " l" & ChrW(7879) & " 1:" & tmp("ratio")
        End If
        i = i + 1
    Next
    STOCKDIVIDENT = tempArr
ends:
    Exit Function
End Function

' Danh sach co dong - fialda.com
Function STOCKHOLDER(MCK As Variant, Optional Individual = False) As Variant
    Dim tempArr, arr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K As Integer
    Dim URL, da As String
    Dim indi As Boolean
    
    URL = "https://fwtapi3.fialda.com/api/services/app/StockInfo/GetMajorShareHolders?symbol=" & MCK
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
    
    ReDim tempArr(0 To JSON("result").Count, 1 To 5)
    
    tempArr(0, 1) = "STT"
    tempArr(0, 2) = "Name"
    tempArr(0, 3) = "So CP"
    tempArr(0, 4) = "Ty le%"
    tempArr(0, 5) = "Ngay update"
    
    i = 1
    For Each tmp In JSON("result")
        indi = tmp("isIndividual")
        
        If indi = Individual Then
            tempArr(i, 1) = i
            tempArr(i, 2) = ChangeType(tmp("name"))
            tempArr(i, 4) = ChangeType(tmp("ownership"))
            tempArr(i, 3) = ChangeType(tmp("shares"))
            da = ChangeType(tmp("updatedDate"))
            tempArr(i, 5) = Left(da, 10)
            i = i + 1
        End If
    Next
    ReDim arr(0 To i - 1, 1 To 5)
    For K = 0 To i - 1
        arr(K, 1) = tempArr(K, 1)
        arr(K, 2) = tempArr(K, 2)
        arr(K, 4) = tempArr(K, 3)
        arr(K, 3) = tempArr(K, 4)
        arr(K, 5) = tempArr(K, 5)
    Next
    STOCKHOLDER = arr
ends:
    Exit Function
End Function

' Giao dich noi bo va lien quan - fialda.com
Function STOCKEVENT(MCK As Variant, Optional EventType = 2, Optional pageSize = 30, Optional pageNumber = 1) As Variant
    Dim tempArr, arr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K As Integer
    Dim URL As String
    Dim da As Long
    Select Case EventType
        Case 1: ' GD co dong lon
        
        Case 3: ' GD lien quan
        
        Case Else
        EventType = 2 ' GD noi bo
    End Select
    
    URL = "https://fwtapi3.fialda.com/api/services/app/Event/GetAll?typeId=" & EventType & "&symbol=" & MCK & "&pageNumber=" & pageNumber & "&pageSize=" & pageSize & "&sortColumn=startDate&isDesc=true"
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    K = JSON("result")("totalCount")
    ReDim tempArr(0 To K, 1 To 10)
    
    tempArr(0, 1) = "STT"
    tempArr(0, 2) = "Ho va Ten"
    tempArr(0, 3) = "Chuc vu"
    tempArr(0, 4) = "Loai GD"
    tempArr(0, 5) = "SL dang ky"
    tempArr(0, 6) = "SL khop lenh"
    tempArr(0, 7) = "SL sau GD"
    tempArr(0, 8) = "Trang thai"
    tempArr(0, 9) = "Ngay ÐK"
    tempArr(0, 10) = "Ngay KT"
    
    i = 1
    For Each tmp In JSON("result")("items")
        tempArr(i, 1) = i
        tempArr(i, 2) = ChangeType(tmp("name"))
        tempArr(i, 3) = ChangeType(tmp("positionName"))
        tempArr(i, 4) = ChangeType(tmp("actionTypeName"))
        tempArr(i, 5) = ChangeType(tmp("shareRegister"))
        tempArr(i, 6) = ChangeType(tmp("shareAcquire"))
        tempArr(i, 7) = ChangeType(tmp("shareAfterTrade"))
        tempArr(i, 8) = ChangeType(tmp("tradeStatusName"))
        tempArr(i, 9) = Left(ChangeType(tmp("startDate")), 10)
        tempArr(i, 10) = Left(ChangeType(tmp("endDate")), 10)
        i = i + 1
    Next
    
    STOCKEVENT = tempArr
ends:
    Exit Function
End Function

'Cong ty con va lien ket -Fialda.com
Function STOCKSUBCOMPANY(MCK As Variant) As Variant
    Dim tempArr, arr, Value, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i, K As Integer
    Dim URL As String
    
    URL = "https://fwtapi4.fialda.com/api/services/app/StockInfo/GetSubCompanies?symbol=" & MCK
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    ReDim tempArr(0 To JSON("result").Count, 1 To 5)
    tempArr(0, 1) = "STT"
    tempArr(0, 2) = "Ten"
    tempArr(0, 3) = "So huu (%)"
    tempArr(0, 4) = "SL nam giu (CP)"
    tempArr(0, 5) = "Von Dieu le"
    i = 1
    For Each tmp In JSON("result")
        tempArr(i, 1) = i
        tempArr(i, 2) = ChangeType(tmp("name"))
        tempArr(i, 3) = ChangeType(tmp("ownership"))
        tempArr(i, 4) = ChangeType(tmp("shares"))
        tempArr(i, 5) = ChangeType(tmp("charterCapital"))
        If tempArr(i, 5) = "" Then tempArr(i, 5) = "Chua niem yet"
        i = i + 1
    Next
    STOCKSUBCOMPANY = tempArr
ends:
    Exit Function
End Function

' Ham ho tro lay du lieu co phieu bank
Private Function TimeLine(Optional quy = True) As String
    Dim y, m, q, i, K As Integer
    y = Year(Now())
    m = Month(Now())
    K = 0
    If m <= 3 Then
        q = 4
    Else
        If m <= 6 Then
            q = 1
        Else
            If m <= 9 Then
                q = 2
            Else
                q = 3
            End If
        End If
    End If
    If q = 4 Then K = -1
    If quy Then
        Select Case q
            Case 4:
            For i = 1 To 4
                TimeLine = TimeLine & "&Timeline=" & y - 2 & "_" & i
            Next i
            For i = 1 To 4
                TimeLine = TimeLine & "&Timeline=" & y - 1 & "_" & i
            Next i
            Case 3:
            TimeLine = "&Timeline=" & y - 2 & "_4"
            For i = 1 To 4
                TimeLine = TimeLine & "&Timeline=" & y - 1 & "_" & i
            Next i
            For i = 1 To 3
                TimeLine = TimeLine & "&Timeline=" & y & "_" & i
            Next i
            Case 2:
            TimeLine = "&Timeline=" & y - 2 & "_1" & "&Timeline=" & y - 2 & "_2"
            For i = 1 To 4
                TimeLine = TimeLine & "&Timeline=" & y - 1 & "_" & i
            Next i
            TimeLine = "&Timeline=" & y & "_1" & "&Timeline=" & y & "_2"
            Case 1:
            For i = 1 To 3
                TimeLine = TimeLine & "&Timeline=" & y - 2 & "_" & i
            Next i
            For i = 1 To 4
                TimeLine = TimeLine & "&Timeline=" & y - 1 & "_" & i
            Next i
            TimeLine = "&Timeline=" & y & "_1"
        End Select
    Else
        For i = 0 To 4
            TimeLine = TimeLine & "&Timeline=" & y - 5 + i & "_5"
        Next i
    End If
End Function
Private Function STRQUY(S As String) As String
    Dim tmp, tmp2 As String
    tmp = Left(S, 4)
    tmp2 = Right(S, 1)
    STRQUY = "Quy " & tmp2 & "/" & tmp
    If CLng(tmp2) = 5 Then STRQUY = "Nam " & tmp
End Function
' thong tin du lieu co phieu Bank - SSI
Function STOCKBANK(MCK As Variant, Optional quy = True) As Variant
    Dim tempArr, Res As Variant
    Dim JSON As Object
    Dim Item, tmp As Dictionary
    Dim i As Integer
    Dim URL, tmp2 As String
    On Error Resume Next
    
    URL = "https://fiin-fundamental.ssi.com.vn/FinancialAnalysis/GetFinancialRatioV2?language=vi&Type=Company&OrganCode=" & MCK & TimeLine(quy)
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    Set JSON = JsonConverter.ParseJson(Res)
    ReDim tempArr(0 To 28, 0 To JSON("totalCount"))
    
    tempArr(0, 0) = "Key"
    tempArr(1, 0) = "P/E" ' ryd21
    tempArr(2, 0) = "P/B" ' ryd25
    tempArr(3, 0) = "EPS"
    tempArr(4, 0) = "BVPS"
    tempArr(5, 0) = "Thu nhap lai thuan (ty)"
    tempArr(6, 0) = "Thu nhap lai thuan %QoQ"
    tempArr(7, 0) = "LNST (ty)"
    tempArr(8, 0) = "LNST %QoQ"
    tempArr(9, 0) = "NIM %"
    tempArr(10, 0) = "ROA %"
    tempArr(11, 0) = "ROE %"
    tempArr(12, 0) = "Tang truong tien gui %"
    tempArr(13, 0) = "Tang truong tin dung %"
    tempArr(14, 0) = "Chi phi/ Thu nhap %"
    tempArr(15, 0) = "Thu nhap ngoai lai/tu lai %"
    tempArr(16, 0) = "YOEA %"
    tempArr(17, 0) = "COF %"
    tempArr(18, 0) = "VCSH/ tong no %"
    tempArr(19, 0) = "VCSH/ tong cho vay %"
    tempArr(20, 0) = "VCSH/ tai san %"
    tempArr(21, 0) = "LDR %"
    tempArr(22, 0) = "CASA %"
    tempArr(23, 0) = "Ty le no xau %"
    tempArr(24, 0) = "Du phong RR tin dung/ No xau %"
    tempArr(25, 0) = "Du phong RR tin dung/ Cho vay %"
    tempArr(26, 0) = "Trich lap du phong/ Cho vay %"
    tempArr(27, 0) = "Von hoa (ty)"
    tempArr(28, 0) = "So luong CP LH (trieu)"
    
    i = 1
    For Each tmp In JSON("items")
        tmp2 = ChangeType(tmp("key"))
        tempArr(0, i) = STRQUY(tmp2)
        
        tmp2 = tmp("value")("organCode")
        If tmp2 <> "EndOfData" Then
            tempArr(1, i) = tmp("value")("ryd21")
            tempArr(2, i) = tmp("value")("ryd25")
            tempArr(3, i) = tmp("value")("ryd14")
            tempArr(4, i) = tmp("value")("ryd7")
            tempArr(5, i) = tmp("value")("rev") / 1000000000#
            tempArr(6, i) = CDec(Replace(tmp("value")("ryq67"), ".", ","))
            tempArr(7, i) = tmp("value")("isa22") / 10000000#
            tempArr(8, i) = CDec(Replace(tmp("value")("ryq39"), ".", ","))
            tempArr(9, i) = CDec(Replace(tmp("value")("ryq44"), ".", ","))
            tempArr(10, i) = CDec(Replace(tmp("value")("ryq14"), ".", ","))
            tempArr(11, i) = CDec(Replace(tmp("value")("ryq12"), ".", ","))
            tempArr(12, i) = CDec(Replace(tmp("value")("rtq51"), ".", ","))
            tempArr(13, i) = CDec(Replace(tmp("value")("rtq50"), ".", ","))
            tempArr(14, i) = CDec(Replace(tmp("value")("ryq48"), ".", ","))
            tempArr(15, i) = CDec(Replace(tmp("value")("ryq47"), ".", ","))
            tempArr(16, i) = CDec(Replace(tmp("value")("ryq45"), ".", ","))
            tempArr(17, i) = CDec(Replace(tmp("value")("ryq46"), ".", ","))
            tempArr(18, i) = CDec(Replace(tmp("value")("ryq54"), ".", ","))
            tempArr(19, i) = CDec(Replace(tmp("value")("ryq55"), ".", ","))
            tempArr(20, i) = CDec(Replace(tmp("value")("ryq56"), ".", ","))
            tempArr(21, i) = CDec(Replace(tmp("value")("ryq57"), ".", ","))
            tempArr(22, i) = CDec(Replace(tmp("value")("casa"), ".", ","))
            tempArr(23, i) = CDec(Replace(tmp("value")("ryq58"), ".", ","))
            tempArr(24, i) = CDec(Replace(tmp("value")("ryq59"), ".", ","))
            tempArr(25, i) = CDec(Replace(tmp("value")("ryq60"), ".", ","))
            tempArr(26, i) = CDec(Replace(tmp("value")("ryq61"), ".", ","))
            tempArr(27, i) = CDec(Replace(tmp("value")("ryd11"), ".", ",")) / 1000000000#
            tempArr(28, i) = CDec(Replace(tmp("value")("ryd3"), ".", ",")) / 1000000#
        End If
        i = i + 1
    Next
    
    STOCKBANK = tempArr
ends:
    Exit Function
End Function
Function WICHART(Name As String, Optional HangHoa = True, Optional LamMuot = True) As Variant
    Dim tempArr(), arr, Res, res2 As Variant
    Dim JSON, o As Object
    Dim dm, tmp  As Dictionary
    Dim i, j, K, w As Integer
    Dim URL, arrr() As String
    On Error Resume Next
    
    URL = "https://api.wichart.vn/vietnambiz/vi-mo?name=" & Name & IIf(HangHoa, "&key=hanghoa", "")
    
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/json"
        .Send ""
        If .Status <> 200 Then
          GoTo ends
        End If
        Res = .responseText
    End With
    
    Set JSON = JsonConverter.ParseJson(Res)
        
    arrr = Split(Res, "{""unit""")
    K = UBound(arrr) - LBound(arrr) + 1
    If K < 2 Then GoTo ends

    ReDim arr(0 To 10000, 0 To K - 1)
    arr(2, 0) = "Date"
        
    For i = 1 To K - 1
        If i > 1 Then
            arr(0, i) = ""
            arr(1, i) = ""
        End If
            
        j = 3
        Set o = JsonConverter.ParseJson("{""unit""" & arrr(i))
        arr(2, i) = o("name") & "(" & o("unit") & ")"
        For Each dm In o("data")
            If i = 1 Then arr(j, 0) = Unix2Date(dm(1))
                arr(j, i) = ChangeType(dm(2))
                j = j + 1
        Next
        
        'lam muot
        If LamMuot Then
            w = j - 2
            While w > 2
                If arr(w, i) = "" Then arr(w, i) = arr(w + 1, i)
                w = w - 1
            Wend
        End If
    Next i
    ' chuyen du lieu
    ReDim tempArr(0 To j - 1, 0 To K - 1)
        w = 0
        For i = 0 To K - 1
            For w = 0 To j - 1
                tempArr(w, i) = arr(w, i)
            Next w
        Next i
    'tieu de
    tempArr(0, 0) = "Title"
    tempArr(0, 1) = ChangeType(JSON("title"))
    tempArr(1, 0) = "Update"
    tempArr(1, 1) = ChangeType(JSON("timeUpdate"))
    WICHART = tempArr
ends:
    Exit Function
End Function
