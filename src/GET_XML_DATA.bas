Attribute VB_Name = "GET_XML_DATA"
Option Explicit
Public Const UTCTime% = 7
Public Const SPACE_INDENT As Byte = 4
Public Const REANYChar = "(?:[\u0000-\u007C\u007E-\uFFFF](?:\\\})?)*?"
Public Const IE_FireAnt = "https://www.fireant.vn/"
Public BCTC As Object
Private Function XMLHTTPs(Optional Server As Boolean) As Object
    If Server Then
        Set XMLHTTPs = VBA.CreateObject("MSXML2.serverXMLHTTP.6.0")
        XMLHTTPs.SetTimeouts 7000, 7000, 9000, 9000
    Else
        Set XMLHTTPs = VBA.CreateObject("MSXML2.XMLHTTP.6.0")
    End If
End Function
Private Function glbRegexs()
    Set glbRegexs = VBA.CreateObject("VBScript.RegExp")
    With glbRegexs
        .Global = 1
        .IgnoreCase = 1
        .MultiLine = 1
    End With
End Function
Private Function JValue(ByVal Text, ByVal Key, _
    Optional ByVal LockString$ = """", _
    Optional ByVal DefaultValueNull As Variant = "") As Variant
    If CStr(Text) = "" Or CStr(Key) = "" Then JValue = "": Exit Function
    On Error GoTo E
    Dim L%, i As Long, sp$()
    L = Len(LockString) + 1
    Key = LockString & Key & LockString & ":"
    sp = Split(Text, Key)
    If UBound(sp) <= 0 Then
        JValue = ""
    Else
        JValue = StringVB(sp(1), LockString, DefaultValueNull)
    End If
    Exit Function
E:
    Debug.Print JValue, Key
End Function

Public Function Unix2Date(ByVal vUnixDate$, Optional ByVal UnixMilliseconds As Boolean = True) As Date
    If vUnixDate Like "*[Dd]ate(*)*" Then '"/Date(1568086206000)/"
    Dim f As Long, L As Long
    f = InStr(vUnixDate, "(") + 1: L = InStr(vUnixDate, ")") - f
    vUnixDate = Mid(vUnixDate, f, L)
End If
If UnixMilliseconds Then vUnixDate = VBA.CDec(vUnixDate) / 1000
Unix2Date = VBA.DateAdd("s", vUnixDate, CDbl(VBA.DateSerial(1970, 1, 1))) + UTCTime / 24
End Function

Public Function Date2Unix(ByVal vDate, Optional ByVal UnixMilliseconds As Boolean = True) As Long
    Date2Unix = VBA.DateDiff("s", CDbl(VBA.DateSerial(1970, 1, 1)), vDate)
End Function
Private Function StringVB(ByVal Text$, _
    Optional ByVal LockString$ = """", _
    Optional ByVal DefaultValueNull As Variant = "") As Variant
    Dim i As Long, L As Long, E As Long, iLock As Boolean
    L = Len(Text)
    If Text Like LockString & "*" Then
        For i = 2 To L
            If Mid(Text, i, 1) = LockString Then
                iLock = Not iLock: If iLock Then StringVB = Mid(Text, 2, i - 2)
            Else
                If iLock Then Exit For
            End If
        Next
        If StringVB Like "*Date(*)*" Then
            StringVB = Unix2Date(StringVB, 1)
        End If
        ElseIf Text Like "null*" Then
        StringVB = DefaultValueNull
        ElseIf Text Like "true*" Then
        StringVB = True
        ElseIf Text Like "false*" Then
        StringVB = False
    Else
        For i = 1 To L
            If IsNumeric(Left(Text, i)) Then
                StringVB = CDec(Left(Text, i))
                ElseIf IsNumeric(Left(Text, i) & "0") Then
                StringVB = 0
            Else
                Exit For
            End If
        Next
    End If
End Function

'Form load BCTC
Sub formBCTC(Ctrl As IRibbonControl)
    Dim myBCTC As Object
    Set myBCTC = New BCTC
    With myBCTC
        If Len(ActiveCell.Value) = 3 Then
            .txtMaCP.Value = ActiveCell.Value
        Else
            .txtMaCP.Value = "MBB"
        End If
        .txtLuu.Value = ActiveCell.Address()
        .Show vbModeless
    End With
End Sub
Private Function shortTxt(txt As String) As Boolean
    Dim myarrTitle(), myarrSB(), myarrtxt() As String
    Dim i As Integer
    
    shortTxt = False
    
    myarrTitle = Array("T" & ChrW(7892) & "NG C" & ChrW(7896) & "NG TÀI S" & ChrW(7842) & "N", "TÀI S" & ChrW(7842) & "N", _
    "NGU" & ChrW(7890) & "N V" & ChrW(7888) & "N", "T" & ChrW(7892) & "NG N" & ChrW(7906) & " PH" & ChrW(7842) & "I TR" & ChrW(7842) & " VÀ V" & ChrW(7888) & "N CH" & ChrW(7910) & " S" & ChrW(7902) & " H" & ChrW(7918) & "U", _
    "T" & ChrW(7892) & "NG C" & ChrW(7896) & "NG NGU" & ChrW(7890) & "N V" & ChrW(7888) & "N", "T" & ChrW(7892) & "NG C" & ChrW(7896) & "NG N" & ChrW(7906) & " PH" & ChrW(7842) & "I TR" & ChrW(7842) & " VÀ V" & ChrW(7888) & "N CH" & ChrW(7910) & " S" & ChrW(7902) & " H" & ChrW(7918) & "U", _
    "L" & ChrW(7906) & "I NHU" & ChrW(7852) & "N " & ChrW(272) & "? PHÂN PH" & ChrW(7888) & "I CHO NHÀ " & ChrW(272) & ChrW(7846) & "U T" & ChrW(431))
    
    For i = LBound(myarrTitle) To UBound(myarrTitle)
        If txt = myarrTitle(i) Then
            shortTxt = True
            Exit Function
        End If
    Next i
    
    myarrSB = Array("A", "B", "C", "D", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X")
    myarrtxt = Split(Trim(txt), ".")
    For i = LBound(myarrSB) To UBound(myarrSB)
        If myarrtxt(0) = myarrSB(i) Then
            shortTxt = True
            Exit Function
        End If
    Next i
    
End Function
Sub deleRange()
    Dim deleteRange As Range
    Dim check As Integer
    Range("B4").Select
    check = MsgBoxUni(VNI("Baïn coù chaéc chaén muoán xoaù toaøn boä danh saùch khoâng?"), vbInformation + vbYesNo, VNI("Xoaù toaøn boä danh saùch?"))
    If (check = vbYes) Then
        Set deleteRange = Selection.CurrentRegion
        Call preTable(deleteRange, False)
    Else
        Exit Sub
    End If
End Sub
Function preTable(r As Excel.Range, Optional mau As Boolean = True)
    Dim R2 As Range
    r.Clear
    If mau Then
        With r(1, 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 12
        End With
        ' tieu de
        Set R2 = r.Resize(2, r.Columns.Count)
        With R2
            .Borders.LineStyle = 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Font
                .TintAndShade = 0
                .ThemeColor = xlThemeColorDark1
                
            End With
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .TintAndShade = -0.249977111117893
                .ThemeColor = xlThemeColorLight2
                '.Color = 12611584
            End With
        End With
    End If
End Function
Sub initFireAnt( _
    Optional ByVal MaSIC$ = "MBB", _
    Optional ByVal ReportType$ = 1, _
    Optional ByVal DataColumns = 4, _
    Optional ByVal ToQuarter = 0, _
    Optional ByVal Unit& = 1000000, _
    Optional ByVal So As Boolean, _
    Optional ByVal tar As Range)
    Dim ToYear As Long
    Dim URL As String
    ToYear = Year(Now) + 1
    
    Select Case ReportType
        Case "CDKT":       ReportType = 1
        Case "KQKD":       ReportType = 2
        Case "LCTTTT":    ReportType = 3
        Case "LCTTGT":    ReportType = 4
        Case Else
        ReportType = 1
    End Select
    
    URL = IE_FireAnt & "api/Data/Finance/LastestFinancialReports?symbol=" & MaSIC & _
    "&type=" & ReportType & _
    "&year=" & ToYear & _
    "&quarter=" & ToQuarter & _
    "&count=" & CStr(DataColumns)
    Dim c, r, R2 As Long
    Dim Res$, RE As Object
    Dim i As Integer, sp$(), T$, Levels&
    Dim m, ms, tmp0, tmp1, tmp2, tmp3
    Dim rowtostart As Integer
    rowtostart = tar.row
    
    tar(1, 1).Value = MaSIC
    tar(1, 2).Value = ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & " : x " & Unit
    tar(1, 3).Value = "Ngu" & ChrW(7891) & "n :"
    tar(1, 4).Value = "FireAnt.vn"
    tar(1, 5).Value = "Th" & ChrW(7901) & "i gian :"
    tar(1, 6).Value = Format(Now, "dd/MM/YYYY")
    
    Dim coltostart As Integer
    coltostart = tar.Column
    
    'On Error GoTo ends
    ' connect to web
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    ' xu ly chuoi
    Set RE = glbRegexs
    sp = Split(Res, "{""ID"":")
    i = UBound(sp)
    If i < 1 Then GoTo ends
    RE.Pattern = "\{""Period"":"".*?"",""Year"":(\d+),""Quarter"":(\d),""Value"":(?:null)?(" & REANYChar & ")\}"
    R2 = 1
    If ReportType <> 1 Then So = False
    For r = 1 To i
        Levels = JValue(sp(r), "Level") - 1: If Levels < 0 Then Levels = 0
        tmp0 = Space(Levels * SPACE_INDENT) & JValue(sp(r), "Name") ' tieu de
        If shortTxt(CStr(tmp0)) = True Or So = False Then
            R2 = R2 + 1
            Cells(rowtostart + R2, coltostart).Value = tmp0
            Set m = Nothing: Set m = RE.Execute(sp(r))
            If m.Count Then
                For c = 1 To m.Count
                    tmp1 = m(c - 1).submatches(0) ' year
                    tmp2 = m(c - 1).submatches(1) 'quater
                    tmp3 = m(c - 1).submatches(2) ' value
                    If VBA.IsNumeric(tmp3) Then tmp3 = CDec(tmp3)
                    If tmp3 = "" Then tmp3 = 0
                    Cells(rowtostart + R2, c + coltostart).Value = Round(tmp3 / (10 * Unit), 0)
                    Cells(rowtostart + R2, c + coltostart).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
                    If r = 1 Then
                        Cells(rowtostart + 1, c + coltostart).Value = IIf(tmp2 > 0, "Q" & tmp2 & "/", "") & tmp1
                        If tmp2 = 0 Then Cells(rowtostart + 1, c + coltostart).NumberFormat = "0"
                    End If
                Next
            End If
        End If
    Next
    Dim targ As Range
    Set targ = Range(Cells(rowtostart, coltostart), Cells(rowtostart + R2, coltostart + Val(DataColumns)))
    targ.Columns.AutoFit
    targ.Rows.AutoFit
    targ.Borders.LineStyle = 1
ends:
    'MsgBox "nothing to show"
End Sub

Private Function change(pre As Variant, cur As Variant) As String
    Dim c As String, d As String
    c = Application.WorksheetFunction.Text((cur - pre) / pre, "0.00%")
    d = Application.WorksheetFunction.Text((cur - pre), "#,###")
    change = d & " (" & c & ")"
End Function
' Get gia co phieu - fireant
Public Function FIREANTPRICE(MCK As Variant) As Variant
    On Error Resume Next
    Dim Res$, RE As Object
    Dim r%, i&, sp$(), T$
    Dim m, ms
    Dim tempArr As Variant
    ReDim tempArr(1 To 2)
    
    If Not IsMaCP(MCK) Then Exit Function
    With XMLHTTPs
        .Open "GET", "https://www.fireant.vn/api/Data/Markets/Quotes?symbols=" & MCK, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    sp = Split(Res, """Date"":")
    Set RE = glbRegexs
    RE.Pattern = """Symbol"":""(.+?)"",""Name"":""([^"",]*)"",.*""PriceBasic"":([^"",]*),.*""PriceClose"":([^"",]*)"
    For r = LBound(sp) + 1 To UBound(sp)
        Set m = RE.Execute(sp(r))
        If m.Count Then
            Set ms = m(0).submatches
            T = ms(0) ' Symbol
            If UCase(T) = UCase(MCK) Then
                tempArr(1) = Application.WorksheetFunction.Text(ms(3), "#,###") ' dong cua hien tai
                tempArr(2) = change(ms(2), ms(3)) ' close truoc
            End If
        End If
    Next
ends:
    Set RE = Nothing
    FIREANTPRICE = tempArr
End Function
' BCTC tu fireant
Public Function FIREANTBALANCESHEET(MaCK As Variant, Optional quy = True, Optional Donvi = 1000000, Optional Socot = 8) As Variant
    Dim tempArr As Variant
    Dim c, r As Long
    Dim ToYear, ToQuarter As Long
    Dim Res$, RE As Object
    Dim i As Integer, sp$(), T$, Levels&
    Dim m, ms, tmp0, tmp1, tmp2, tmp3
    Dim MaSIC, URL As String
    ToYear = Year(Now) + 1
    ToQuarter = 0
    If quy Then ToQuarter = 1
    
    URL = IE_FireAnt & "api/Data/Finance/LastestFinancialReports?symbol=" & UCase(MaCK) & _
    "&type=1&year=" & ToYear & "&quarter=" & ToQuarter & "&count=" & Socot
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    ' xu ly chuoi
    Set RE = glbRegexs
    sp = Split(Res, "{""ID"":")
    i = UBound(sp)
    If i < 1 Then GoTo ends
    ReDim tempArr(1 To i + 1, 1 To Socot + 1)
    tempArr(1, 1) = "B" & ChrW(7842) & "NG CÂN " & ChrW(272) & ChrW(7888) & "I K" & ChrW(7870) & " TOÁN"
    
    RE.Pattern = "\{""Period"":"".*?"",""Year"":(\d+),""Quarter"":(\d),""Value"":(?:null)?(" & REANYChar & ")\}"
    For r = 1 To i
        Levels = JValue(sp(r), "Level") - 1: If Levels < 0 Then Levels = 0
        tempArr(r + 1, 1) = Space(Levels * SPACE_INDENT) & JValue(sp(r), "Name") ' tieu de
        Set m = Nothing: Set m = RE.Execute(sp(r))
        If m.Count Then
            For c = 1 To m.Count
                tmp1 = m(c - 1).submatches(0) ' year
                tmp2 = m(c - 1).submatches(1) 'quater
                tmp3 = m(c - 1).submatches(2) ' value
                If VBA.IsNumeric(tmp3) Then tmp3 = CDec(tmp3)
                If tmp3 = "" Then tmp3 = 0
                tempArr(r + 1, c + 1) = Round(tmp3 / (10 * Donvi), 0)
                If r = 1 Then tempArr(1, c + 1) = IIf(tmp2 > 0, "Q" & tmp2 & "/", "") & tmp1
            Next
        End If
    Next r
    FIREANTBALANCESHEET = tempArr
ends:
    Exit Function
End Function

Public Function FIREANTINCOME(MaCK As Variant, Optional quy = True, Optional Donvi = 1000000, Optional Socot = 12) As Variant
    Dim tempArr As Variant
    Dim c, r As Long
    Dim ToYear, ToQuarter As Long
    Dim Res$, RE As Object
    Dim i As Integer, sp$(), T$, Levels&
    Dim m, ms, tmp0, tmp1, tmp2, tmp3
    Dim MaSIC, URL As String
    ToYear = Year(Now) + 1
    ToQuarter = 0
    If quy Then ToQuarter = 1
    
    URL = IE_FireAnt & "api/Data/Finance/LastestFinancialReports?symbol=" & UCase(MaCK) & _
    "&type=2&year=" & ToYear & "&quarter=" & ToQuarter & "&count=" & Socot
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    ' xu ly chuoi
    Set RE = glbRegexs
    sp = Split(Res, "{""ID"":")
    i = UBound(sp)
    If i < 1 Then GoTo ends
    ReDim tempArr(1 To i + 1, 1 To Socot + 1)
    tempArr(1, 1) = "K" & ChrW(7870) & "T QU" & ChrW(7842) & " KINH DOANH"
    
    RE.Pattern = "\{""Period"":"".*?"",""Year"":(\d+),""Quarter"":(\d),""Value"":(?:null)?(" & REANYChar & ")\}"
    For r = 1 To i
        Levels = JValue(sp(r), "Level") - 1: If Levels < 0 Then Levels = 0
        tempArr(r + 1, 1) = Space(Levels * SPACE_INDENT) & JValue(sp(r), "Name") ' tieu de
        Set m = Nothing: Set m = RE.Execute(sp(r))
        If m.Count Then
            For c = 1 To m.Count
                tmp1 = m(c - 1).submatches(0) ' year
                tmp2 = m(c - 1).submatches(1) 'quater
                tmp3 = m(c - 1).submatches(2) ' value
                If VBA.IsNumeric(tmp3) Then tmp3 = CDec(tmp3)
                If tmp3 = "" Then tmp3 = 0
                tempArr(r + 1, c + 1) = Round(tmp3 / (10 * Donvi), 0)
                If r = 1 Then tempArr(1, c + 1) = IIf(tmp2 > 0, "Q" & tmp2 & "/", "") & tmp1
            Next
        End If
    Next r
    FIREANTINCOME = tempArr
ends:
    Exit Function
End Function

Public Function FIREANTCASHFLOWDIRECT(MaCK As Variant, Optional quy = True, Optional Donvi = 1000000, Optional Socot = 8) As Variant
    Dim tempArr As Variant
    Dim c, r As Long
    Dim ToYear, ToQuarter As Long
    Dim Res$, RE As Object
    Dim i As Integer, sp$(), T$, Levels&
    Dim m, ms, tmp0, tmp1, tmp2, tmp3
    Dim MaSIC, URL As String
    
    ToYear = Year(Now) + 1
    ToQuarter = 0
    If quy Then ToQuarter = 1
    
    URL = IE_FireAnt & "api/Data/Finance/LastestFinancialReports?symbol=" & UCase(MaCK) & _
    "&type=3&year=" & ToYear & "&quarter=" & ToQuarter & "&count=" & Socot
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    ' xu ly chuoi
    Set RE = glbRegexs
    sp = Split(Res, "{""ID"":")
    i = UBound(sp)
    If i < 1 Then GoTo ends
    ReDim tempArr(1 To i + 1, 1 To Socot + 1)
    tempArr(1, 1) = "BÁO CÁO L" & ChrW(431) & "U CHUY" & ChrW(7874) & "N TI" & ChrW(7872) & "N T" & ChrW(7878) & " TR" & ChrW(7920) & "C TI" & ChrW(7870) & "P"
    
    RE.Pattern = "\{""Period"":"".*?"",""Year"":(\d+),""Quarter"":(\d),""Value"":(?:null)?(" & REANYChar & ")\}"
    For r = 1 To i
        Levels = JValue(sp(r), "Level") - 1: If Levels < 0 Then Levels = 0
        tempArr(r + 1, 1) = Space(Levels * SPACE_INDENT) & JValue(sp(r), "Name") ' tieu de
        Set m = Nothing: Set m = RE.Execute(sp(r))
        If m.Count Then
            For c = 1 To m.Count
                tmp1 = m(c - 1).submatches(0) ' year
                tmp2 = m(c - 1).submatches(1) 'quater
                tmp3 = m(c - 1).submatches(2) ' value
                If VBA.IsNumeric(tmp3) Then tmp3 = CDec(tmp3)
                If tmp3 = "" Then tmp3 = 0
                tempArr(r + 1, c + 1) = Round(tmp3 / (10 * Donvi), 0)
                If r = 1 Then tempArr(1, c + 1) = IIf(tmp2 > 0, "Q" & tmp2 & "/", "") & tmp1
            Next
        End If
    Next r
    FIREANTCASHFLOWDIRECT = tempArr
ends:
    Exit Function
End Function

Public Function FIREANTCASHFLOWINDIRECT(MaCK As Variant, Optional quy = True, Optional Donvi = 1000000, Optional Socot = 8) As Variant
    Dim tempArr As Variant
    Dim c, r As Long
    Dim ToYear, ToQuarter As Long
    Dim Res$, RE As Object
    Dim i As Integer, sp$(), T$, Levels&
    Dim m, ms, tmp0, tmp1, tmp2, tmp3
    Dim MaSIC, URL As String
    MaSIC = MaCK
    If Not IsMaCP(MaSIC) Then GoTo ends
    ToYear = Year(Now) + 1
    ToQuarter = 0
    If quy Then ToQuarter = 1
    
    URL = IE_FireAnt & "api/Data/Finance/LastestFinancialReports?symbol=" & UCase(MaCK) & _
    "&type=4&year=" & ToYear & "&quarter=" & ToQuarter & "&count=" & Socot
    With XMLHTTPs
        .Open "GET", URL, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    
    ' xu ly chuoi
    Set RE = glbRegexs
    sp = Split(Res, "{""ID"":")
    i = UBound(sp)
    If i < 1 Then GoTo ends
    ReDim tempArr(1 To i + 1, 1 To Socot + 1)
    tempArr(1, 1) = "BÁO CÁO L" & ChrW(431) & "U CHUY" & ChrW(7874) & "N TI" & ChrW(7872) & "N T" & ChrW(7878) & " GIÁN TI" & ChrW(7870) & "P"
    
    
    RE.Pattern = "\{""Period"":"".*?"",""Year"":(\d+),""Quarter"":(\d),""Value"":(?:null)?(" & REANYChar & ")\}"
    For r = 1 To i
        Levels = JValue(sp(r), "Level") - 1: If Levels < 0 Then Levels = 0
        tempArr(r + 1, 1) = Space(Levels * SPACE_INDENT) & JValue(sp(r), "Name") ' tieu de
        Set m = Nothing: Set m = RE.Execute(sp(r))
        If m.Count Then
            For c = 1 To m.Count
                tmp1 = m(c - 1).submatches(0) ' year
                tmp2 = m(c - 1).submatches(1) 'quater
                tmp3 = m(c - 1).submatches(2) ' value
                If VBA.IsNumeric(tmp3) Then tmp3 = CDec(tmp3)
                If tmp3 = "" Then tmp3 = 0
                tempArr(r + 1, c + 1) = Round(tmp3 / (10 * Donvi), 0)
                If r = 1 Then tempArr(1, c + 1) = IIf(tmp2 > 0, "Q" & tmp2 & "/", "") & tmp1
            Next
        End If
    Next r
    FIREANTCASHFLOWINDIRECT = tempArr
ends:
    Exit Function
End Function

' Thong tin gia giao dich co phieu
Function STOCKVN(Symbol As Variant, Optional Index As Integer = 0) As Variant
    On Error Resume Next
    Dim Res$, RE As Object
    Dim r%, i&, sp$(), T$
    Dim m, ms
    Dim SIC As String
    SIC = UCase(Symbol)
    With XMLHTTPs
        .Open "GET", "https://www.fireant.vn/api/Data/Markets/Quotes?symbols=" & SIC, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send ""
        If .Status <> 200 Then
            GoTo ends
        End If
        Res = .responseText
    End With
    sp = Split(Res, """Date"":")
    Set RE = glbRegexs
    RE.Pattern = """Symbol"":""(.+?)"",""Name"":""([^"",]*)"",.*""Exchange"":""([^"",]*)"",.*""PriceBasic"":([^"",]*),.*""PriceClose"":([^"",]*)"
    For r = LBound(sp) + 1 To UBound(sp)
        Set m = RE.Execute(sp(r))
        If m.Count Then
            Set ms = m(0).submatches
            Select Case Index
                Case 1:
                STOCKVN = ms(1) 'name
                Case 2:
                STOCKVN = ms(2) 'san
                Case 3:
                STOCKVN = Val(ms(4)) - Val(ms(3)) ' change
                Case 4:
                STOCKVN = (Val(ms(4)) - Val(ms(3))) / Val(ms(3)) ' change%
                Case Else
                STOCKVN = Val(ms(4)) ' dong cua hien tai
            End Select
        End If
    Next
ends:
    Set RE = Nothing
End Function

Public Function IsMaCP(ByVal SIC As String) As Boolean
    Dim i As Integer
    Dim arrMaCP(), arrMaCP2(), arrMaCP3(), arrMaCP4(), arrMaCP5()
    SIC = UCase(SIC)
    
    arrMaCP = Array("A32", "AAA", "AAM", "AAS", "AAT", "AAV", "ABB", "ABC", "ABI", "ABR", "ABS", "ABT", "AC4", "ACB", "ACC", "ACE", "ACG", "ACL", "ACM", "ACS", "ACV", _
    "ADC", "ADG", "ADP", "ADS", "AFC", "AFX", "AG1", "AGC", "AGD", "AGF", "AGG", "AGM", "AGP", "AGR", "AGX", "AIC", "ALP", "ALT", "ALV", "AMC", "AMD", "AME", "AMP", "AMS", "AMV", "ANT", "ANV", "APC", "APF", "APG", _
    "APH", "API", "APL", "APP", "APS", "APT", "AQN", "ARM", "ART", "ASA", "ASD", "ASG", "ASIAGF", "ASM", "ASP", "AST", "ATA", "ATB", "ATD", "ATG", "ATS", "AUM", "AVC", "AVF", "AVS", "B82", "BAB", "BAL", "BAM", "BAS", _
    "BAX", "BBC", "BBH", "BBM", "BBS", "BBT", "BCA", "BCB", "BCC", "BCE", "BCF", "BCG", "BCI", "BCM", "BCP", "BCV", "BDB", "BDC", "BDF", "BDG", "BDP", "BDT", "BDW", "BED", "BEL", "BFC", "BGM", "BGW", "BHA", "BHC", _
    "BHG", "BHK", "BHN", "BHP", "BHS", "BHT", "BHV", "BIC", "BID", "BII", "BIO", "BKC", "BKG", "BKH", "BLF", "BLI", "BLN", "BLT", "BLU", "BLW", "BM9", "BMC", "BMD", "BMF", "BMG", "BMI", "BMJ", "BMN", "BMP", "BMS", _
    "BMV", "BNA", "BNW", "BOT", "BPC", "BPW", "BQB", "BRC", "BRR", "BRS", "BSA", "BSC", "BSD", "BSG", "BSH", "BSI", "BSL", "BSP", "BSQ", "BSR", "BST", "BT1", "BT6", "BTB", "BTC", "BTD", "BTG", "BTH", "BTN", "BTP", _
    "BTR", "BTS", "BTT", "BTU", "BTV", "BTW", "BUD", "BVB", "BVG", "BVH", "BVL", "BVN", "BVS", "BWA", "BWE", "BWS", "BXD", "BXH", "BXT", "C12", "C21", "C22", "C32", "C36", "C47", "C4G", "C69", "C71", "C92", "CAB", "CAD", "CAG", "CAM", "CAN", _
    "CAP", "CAT", "CAV", "CBC", "CBI", "CBS", "CC1", "CC4", "CCA", "CCH", "CCI", "CCL", "CCM", "CCP", "CCR", "CCT", "CCV", "CDC", "CDG", "CDH", "CDN", "CDO", "CDP", "CDR", "CE1", "CEC", "CEE", "CEG", "CEN", "CEO", _
    "CER", "CET", "CFC", "CFM", "CFV", "CGL", "CGP", "CGV", "CH5", "CHC", "CHP", "CHS", "CI5", "CIA", "CIC", "CID", "CIG", "CII", "CIP", "CJC", "CKA", "CKD", "CKG", "CKH", "CKV", "CLC", "CLG", "CLH", "CLL", "CLM", _
    "CLP", "CLS", "CLW", "CLX", "CMC", "CMD", "CMF", "CMG", "CMI", "CMK", "CMN", "CMP", "CMS", "CMT", "CMV", "CMW", "CMX", "CNC", "CNG", "CNH", "CNN", "CNT", "COM", "CPA", "CPC", "CPH", "CPI", "CPW", "CQN", "CQT", "CRC")
    
    arrMaCP2 = Array("CRE", "CSC", "CSG", "CSI", "CSM", "CST", "CSV", "CT3", "CT5", "CT6", "CTA", "CTB", "CTC", "CTD", "CTF", "CTG", "CTI", "CTM", "CTN", "CTP", "CTR", "CTS", "CTT", "CTV", "CTW", "CTX", "CVC", "CVH", "CVN", "CVT", _
    "CX8", "CXH", "CYC", "CZC", "D11", "D26", "D2D", "DAC", "DAD", "DAE", "DAG", "DAH", "DAP", "DAR", "DAS", "DAT", "DBC", "DBD", "DBF", "DBH", "DBM", "DBT", "DBW", "DC1", "DC2", "DC4", "DCC", "DCD", "DCF", "DCG", _
    "DCH", "DCI", "DCL", "DCM", "DCR", "DCS", "DCT", "DDG", "DDH", "DDM", "DDN", "DDV", "DFC", "DFF", "DFS", "DGC", "DGL", "DGT", "DGW", "DHA", "DHB", "DHC", "DHD", "DHG", "DHI", "DHL", "DHM", "DHN", "DHP", "DHT", _
    "DIC", "DID", "DIG", "DIH", "DKC", "DKH", "DKP", "DL1", "DLC", "DLD", "DLG", "DLR", "DLT", "DLV", "DM7", "DMC", "DNA", "DNB", "DNC", "DND", "DNE", "DNF", "DNH", "DNL", "DNM", "DNN", "DNP", "DNR", "DNS", "DNT", _
    "DNW", "DNY", "DOC", "DOP", "DP1", "DP2", "DP3", "DPC", "DPD", "DPG", "DPH", "DPM", "DPP", "DPR", "DPS", "DQC", "DRC", "DRG", "DRH", "DRI", "DRL", "DS3", "DSC", "DSG", "DSN", "DSP", "DSS", "DST", "DSV", "DT4", _
    "DTA", "DTB", "DTC", "DTD", "DTE", "DTG", "DTI", "DTK", "DTL", "DTN", "DTP", "DTT", "DTV", "DUS", "DVC", "DVD", "DVG", "DVH", "DVN", "DVP", "DVW", "DWS", "DX2", "DXD", "DXG", "DXL", "DXP", "DXS", _
    "DXV", "DZM", "E12", "E1SSHN30", "E1VFVN30", "E29", "EAD", "EBA", "EBS", "ECI", "EFI", "EIB", "EIC", "EID", "EIN", "ELC", "EMC", "EME", "EMG", "EMS", "ENF", "EPC", "EPH", "EVE", "EVF", "EVG", "EVS", "FBA", "FBC", _
    "FBT", "FCC", "FCM", "FCN", "FCS", "FDC", "FDG", "FDT", "FGL", "FHN", "FHS", "FIC", "FID", "FIR", "FIT", "FLC", "FMC", "FOC", "FOX", "FPC", "FPT", "FRC", "FRM", "FRT", "FSO", "FT1", "FTI", "FTM", "FTS", "G20", "G36", _
    "GAB", "GAS", "GBS", "GCB", "GDT", "GDW", "GE2", "GEG", "GER", "GEX", "GFC", "GGG", "GGS", "GH3", "GHA", "GHC", "GIC", "GIL", "GKM", "GLC", "GLT", "GLW", "GMA", "GMC", "GMD", "GMX", "GND", "GQN", "GSM", "GSP", "GTA", _
    "GTC", "GTD", "GTH", "GTK", "GTN", "GTS", "GTT", "GVR", "GVT", "H11", "HAB", "HAC", "HAD", "HAF", "HAG", "HAH", "HAI", "HAM", "HAN", "HAP", "HAR", "HAS", "HAT", "HAV", "HAW", "HAX", "HBB", "HBC", "HBD", "HBE", _
    "HBH", "HBI", "HBS", "HBW", "HC1", "HC3", "HCB", "HCC", "HCD", "HCI", "HCM", "HCS", "HCT", "HD2", "HD3", "HD6", "HD8", "HDA", "HDB", "HDC", "HDG", "HDM", "HDO", "HDP", "HDW", "HEC", "HEJ", "HEM", "HEP", "HES", _
    "HEV", "HFB", "HFC", "HFS", "HFT", "HFX", "HGA", "HGC", "HGM", "HGR", "HGT", "HGW", "HHA", "HHC", "HHG", "HHL", "HHN", "HHP", "HHR", "HHS", "HHV", "HID", "HIG", "HII", "HIZ", "HJC", "HJS", "HKB", "HKC", "HKP", _
    "HKT", "HLA", "HLB", "HLC", "HLD", "HLE", "HLG", "HLR", "HLS", "HLT", "HLY", "HMC", "HMG", "HMH", "HMS")
    
    arrMaCP3 = Array("HNA", "HNB", "HND", "HNE", "HNF", "HNG", "HNI", "HNM", "HNP", "HNR", "HNT", "HOM", "HOT", "HPB", "HPD", "HPG", "HPH", "HPI", "HPL", "HPM", "HPP", "HPR", "HPS", "HPT", "HPU", "HPW", "HPX", "HQC", "HRB", _
    "HRC", "HRG", "HRT", "HSA", "HSC", "HSG", "HSI", "HSL", "HSM", "HSP", "HST", "HSV", "HT1", "HT2", "HTB", "HTC", "HTE", "HTG", "HTH", "HTI", "HTK", "HTL", "HTM", "HTN", "HTP", "HTR", _
    "HTT", "HTU", "HTV", "HTW", "HU1", "HU3", "HU4", "HU6", "HUB", "HUG", "HUT", "HUX", "HVA", "HVC", "HVG", "HVH", "HVN", "HVT", "HVX", "HWS", "I10", "IBC", "IBD", "IBN", "ICC", "ICF", "ICG", "ICI", "ICN", "ICT", "IDC", "IDI", _
    "IDJ", "IDN", "IDP", "IDV", "IFC", "IFS", "IHK", "IJC", "IKH", "ILA", "ILB", "ILC", "ILS", "IME", "IMP", "IMT", "IN4", "INC", "INN", "IPA", "IPH", "IRC", "ISG", "ISH", "IST", "ITA", "ITC", "ITD", "ITQ", "ITS", "IVS", "JOS", _
    "JSC", "JVC", "KAC", "KBC", "KBE", "KBT", "KCB", "KCE", "KDC", "KDF", "KDH", "KDM", "KGM", "KGU", "KHA", "KHB", "KHD", "KHG", "KHL", "KHP", "KHS", "KHW", "KIP", "KKC", "KLB", "KLF", "KLM", "KLS", "KMF", "KMR", "KMT", "KOS", _
    "KPF", "KSA", "KSB", "KSC", "KSD", "KSE", "KSF", "KSH", "KSK", "KSQ", "KSS", "KST", "KSV", "KTB", "KTC", "KTL", "KTS", "KTT", "KTU", "KVC", "L10", "L12", "L14", "L18", "L35", "L40", "L43", "L44", "L45", "L61", "L62", "L63", _
    "LAF", "LAI", "LAS", "LAW", "LBC", "LBE", "LBM", "LCC", "LCD", "LCG", "LCM", "LCS", "LCW", "LDG", "LDP", "LDW", "LEC", "LG9", "LGC", "LGL", _
    "LGM", "LHC", "LHG", "LIC", "LIG", "LIX", "LKW", "LLM", "LM3", "LM7", "LM8", "LMC", "LMH", "LMI", "LNC", "LO5", "LPB", "LPT", "LQN", "LSS", "LTC", "LTG", "LUT", "LWS", "LYF", "M10", "MA1", "MAC", "MAFPF1", "MAS", "MAX", "MBB", _
    "MBG", "MBN", "MBS", "MC3", "MCC", "MCF", "MCG", "MCH", "MCI", "MCL", "MCM", "MCO", "MCP", "MCT", "MCV", "MDA", "MDC", "MDF", "MDG", "MDN", "MDT", "MEC", "MED", "MEF", "MEG", "MEL", "MES", "MFS", "MGC", "MGG", "MH3", "MHC", "MHL", "MHP", _
    "MHY", "MIC", "MIE", "MIG", "MIH", "MIM", "MJC", "MKP", "MKV", "MLC", "MLN", "MLS", "MMC", "MML", "MNB", "MNC", "MND", "MPC", "MPT", "MPY", "MQB", "MQN", "MRF", "MSB", "MSC", "MSH", "MSN", "MSR", "MST", "MTA", "MTB", "MTC", "MTG", "MTH", "MTL", "MTM", "MTP", "MTS", "MTV", "MVB", "MVC", "MVN", "MVY", "MWG", "MXC", "NAB", "NAC", "NAF", "NAG", "NAP", "NAS", _
    "NAU", "NAV", "NAW", "NBB", "NBC", "NBE", "NBP", "NBR", "NBT", "NBW", "NCP", "NCS", "NCT", "ND2", "NDC", "NDF", "NDN", "NDP", "NDT", "NDW", "NDX", "NED", "NET", "NFC", "NGC", "NHA", "NHC", "NHH", "NHN", _
    "NHP", "NHS", "NHT", "NHV", "NHW", "NIS", "NJC", "NKD", "NKG", "NLC", "NLG", "NLS", "NMK", "NNB", "NNC", "NNG", "NNQ", "NNT", "NOS", "NPH", "NPS", "NQB", "NQN", "NQT", "NRC", "NS2", "NS3", "NSC", "NSG", "NSH", "NSL", "NSN", "NSP", "NSS", _
    "NST", "NT2", "NTB", "NTC", "NTF", "NTH", "NTL", "NTP", "NTR", "NTT", "NTW", "NUE", "NVB", "NVC", "NVL", "NVN", "NVP")
    
    arrMaCP4 = Array("NVT", "NWT", "OCB", "OCH", "OGC", "OIL", "ONE", "ONW", "OPC", "ORS", "PAC", "PAI", "PAN", "PAP", "PAS", "PBC", "PBK", "PBP", "PBT", "PC1", "PCC", "PCE", "PCF", "PCG", "PCM", "PCN", "PCT", "PDB", "PDC", "PDN", "PDR", "PDT", "PDV", "PEC", _
    "PEG", "PEN", "PEQ", "PET", "PFL", "PFV", "PGB", "PGC", "PGD", "PGI", "PGN", "PGS", "PGT", "PGV", "PHC", "PHH", "PHN", "PHP", "PHR", "PHS", "PHT", "PIA", "PIC", "PID", "PIS", "PIT", "PIV", "PJC", "PJS", "PJT", "PKR", "PLA", "PLC", "PLE", _
    "PLO", "PLP", "PLX", "PMB", "PMC", "PME", "PMG", "PMJ", "PMP", "PMS", "PMT", "PMW", "PNC", "PND", "PNG", "PNJ", "PNP", "PNT", "POB", "POM", "POS", "POT", "POV", "POW", "PPC", "PPE", "PPG", _
    "PPH", "PPP", "PPS", "PPY", "PQN", "PRC", "PRE", "PRO", "PRT", "PRUBF1", "PSB", "PSC", "PSD", "PSE", "PSG", "PSH", "PSI", "PSL", "PSN", "PSP", "PSW", "PTB", "PTC", "PTD", "PTE", "PTG", "PTH", "PTI", "PTK", "PTL", "PTM", "PTO", _
    "PTP", "PTS", "PTT", "PTV", "PTX", "PV2", "PVA", "PVB", "PVC", "PVD", "PVE", "PVF", "PVG", "PVH", "PVI", "PVL", "PVM", "PVO", "PVP", "PVR", "PVS", "PVT", "PVV", "PVX", "PVY", "PWA", "PWS", "PX1", "PXA", "PXC", "PXI", "PXL", "PXM", "PXS", _
    "PXT", "PYU", "QBR", "QBS", "QCC", "QCG", "QHD", "QHW", "QLD", "QLT", "QNC", "QNS", "QNT", "QNU", "QNW", "QPH", "QSP", "QST", "QTC", "QTP", "RAL", "RAT", "RBC", "RCC", "RCD", "RCL", "RDP", "REE", "REM", "RGC", "RHC", "RHN", "RIC", "RLC", _
    "ROS", "RTB", "RTH", "RTS", "S12", "S27", "S33", "S4A", "S55", "S64", "S72", "S74", "S91", "S96", "S99", "SAB", "SAC", "SAF", "SAL", "SAM", "SAP", "SAS", "SAV", "SB1", "SBA", "SBC", "SBD", "SBH", "SBL", "SBM", _
    "SBR", "SBS", "SBT", "SBV", "SC5", "SCA", "SCC", "SCD", "SCG", "SCH", "SCI", "SCJ", "SCL", "SCO", "SCR", "SCS", "SCV", "SCY", "SD1", "SD2", "SD3", "SD4", "SD5", "SD6", "SD7", "SD8", "SD9", "SDA", "SDB", "SDC", "SDD", "SDE", "SDF", _
    "SDG", "SDH", "SDI", "SDJ", "SDK", "SDN", "SDP", "SDS", "SDT", "SDU", "SDV", "SDX", "SDY", "SEA", "SEB", "SEC", "SED", "SEL", "SEP", "SFC", "SFG", "SFI", "SFN", "SFT", "SGB", "SGC", "SGD", "SGH", "SGI", "SGN", "SGO", "SGP", "SGR", "SGS", "SGT", "SHA", _
    "SHB", "SHC", "SHE", "SHG", "SHI", "SHN", "SHP", "SHS", "SHV", "SHX", "SIC", "SID", "SIG", "SII", "SIP", "SIV", "SJ1", "SJC", "SJD", "SJE", "SJF", "SJG", "SJM", "SJS", "SKG", "SKH", "SKN", "SKS", "SKV", "SLC", "SLS", "SMA", "SMB", "SMC", "SME", "SMN", _
    "SMT", "SNC", "SNG", "SNZ", "SON", "SOV", "SP2", "SPA", "SPB", "SPC", "SPD", "SPH", "SPI", "SPM", "SPP", "SPV", "SQC", "SRA", "SRB", "SRC", "SRF", "SRT", "SSB", "SSC", "SSF", "SSG", "SSH", "SSI", "SSM", "SSN", "SSS", "SSU", "ST8", "STB", "STC", "STG", _
    "STH", "STK", "STL", "STP", "STS", "STT", "STU", "STV", "STW", "SUM", "SVC", "SVD", "SVG", "SVH", "SVI", "SVL", "SVN", "SVS", "SVT", "SWC", "SZB", "SZC", "SZE", _
    "SZL", "T12", "TA3", "TA6", "TA9", "TAC", "TAG", "TAN", "TAP", "TAR", "TAS", "TAW", "TB8", "TBC", "TBD", "TBH", "TBN", "TBT", "TBX", "TC6", "TCB", "TCD", "TCH", "TCI", "TCJ", "TCK", "TCL", "TCM", "TCO", "TCR", "TCS", "TCT", "TCW", "TDB", "TDC", "TDF")
    
    arrMaCP5 = Array("TDG", "TDH", "TDM", "TDN", "TDP", "TDS", "TDT", "TDW", "TEC", "TEG", "TEL", "TET", "TFC", "TGG", "TGP", "TH1", "THB", "THD", _
    "THG", "THI", "THN", "THP", "THR", "THS", "THT", "THU", "THV", "THW", "TID", "TIE", "TIG", "TIP", "TIS", "TIX", "TJC", "TKA", "TKC", "TKG", "TKU", "TL4", "TLC", "TLD", "TLG", "TLH", "TLI", "TLP", "TLT", "TMB", "TMC", "TMG", "TMP", "TMS", "TMT", "TMW", _
    "TMX", "TN1", "TNA", "TNB", "TNC", "TND", "TNG", "TNH", "TNI", "TNM", "TNP", "TNS", "TNT", "TNW", "TNY", "TOP", "TOS", "TOT", "TOW", "TPB", "TPC", "TPH", "TPP", "TPS", "TQN", "TQW", "TR1", "TRA", "TRC", "TRS", "TRT", "TS3", "TS4", "TS5", "TSB", "TSC", "TSD", "TSG", _
    "TSJ", "TSM", "TST", "TTA", "TTB", "TTC", "TTD", "TTE", "TTF", "TTG", "TTH", "TTJ", "TTL", "TTN", "TTP", "TTR", "TTS", "TTT", "TTV", "TTZ", "TUG", "TV1", "TV2", "TV3", "TV4", "TV6", "TVA", "TVB", "TVC", "TVD", "TVG", "TVH", "TVM", "TVN", "TVP", "TVS", _
    "TVT", "TVU", "TVW", "TW3", "TXM", "TYA", "UCT", "UDC", "UDJ", "UDL", "UEM", "UIC", "UMC", "UNI", "UPC", "UPH", "USC", "USD", "V11", "V12", "V15", "V21", "VAB", "VAF", "VAT", "VAV", "VBB", "VBC", "VBG", "VBH", "VC1", "VC2", "VC3", "VC5", "VC6", "VC7", _
    "VC9", "VCA", "VCB", "VCC", "VCE", "VCF", "VCG", "VCH", "VCI", "VCM", "VCP", "VCR", "VCS", "VCT", "VCV", "VCW", "VCX", "VDB", "VDL", "VDM", "VDN", "VDP", "VDS", "VDT", "VE1", "VE2", "VE3", "VE4", "VE8", "VE9", "VEA", "VEC", "VEE", "VEF", "VES", "VET", _
    "VFC", "VFG", "VFMVF1", "VFMVF4", "VFMVFA", "VFR", "VFS", "VGC", "VGG", "VGI", "VGL", "VGP", "VGR", "VGS", "VGT", "VGV", "VHC", "VHD", "VHE", "VHF", "VHG", "VHH", "VHI", "VHL", "VHM", "VIA", "VIB", "VIC", "VID", "VIE", "VIF", "VIG", "VIH", "VIM", "VIN", "VIP", "VIR", _
    "VIS", "VIT", "VIW", "VIX", "VJC", "VKC", "VKD", "VKP", "VLA", "VLB", "VLC", "VLF", "VLG", "VLP", "VLW", "VMA", "VMC", "VMD", "VMG", "VMI", "VMS", "VNA", "VNB", "VNC", "VND", "VNE", "VNF", "VNG", "VNH", "VNI", "VNIndex", "VNL", "VNM", "VNN", "VNP", "VNR", _
    "VNS", "VNT", "VNX", "VNY", "VOC", "VOS", "VPA", "VPB", "VPC", "VPD", "VPG", "VPH", "VPI", "VPK", "VPL", "VPR", "VPS", "VPW", "VQC", "VRC", "VRE", "VRG", "VSA", "VSC", "VSE", "VSF", "VSG", "VSH", "VSI", "VSM", "VSN", "VSP", "VST", "VT1", "VT8", "VTA", "VTB", "VTC", _
    "VTD", "VTE", "VTG", "VTH", "VTI", "VTJ", "VTK", "VTL", "VTM", "VTO", "VTP", "VTQ", "VTR", "VTS", "VTV", "VTX", "VVN", "VW3", "VWS", "VXB", "VXP", "VXT", "WCS", "WSB", "WSS", "WTC", "WTN", "X18", "X20", "X26", "X77", "XDH", "XHC", "XLV", "XMC", "XMD", "XMP", "XPH", _
    "YBC", "YBM", "YEG", "YRC", "YSC", "YTC")
    IsMaCP = False
    If Len(SIC) <> 3 Then Exit Function
    For i = LBound(arrMaCP) To UBound(arrMaCP)
        If SIC = arrMaCP(i) Then
            IsMaCP = True
            Exit Function
        End If
    Next i
    For i = LBound(arrMaCP2) To UBound(arrMaCP2)
        If SIC = arrMaCP2(i) Then
            IsMaCP = True
            Exit Function
        End If
    Next i
    For i = LBound(arrMaCP3) To UBound(arrMaCP3)
        If SIC = arrMaCP3(i) Then
            IsMaCP = True
            Exit Function
        End If
    Next i
    For i = LBound(arrMaCP4) To UBound(arrMaCP4)
        If SIC = arrMaCP4(i) Then
            IsMaCP = True
            Exit Function
        End If
    Next i
    For i = LBound(arrMaCP5) To UBound(arrMaCP5)
        If SIC = arrMaCP5(i) Then
            IsMaCP = True
            Exit Function
        End If
    Next i
End Function
