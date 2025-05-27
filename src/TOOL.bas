Attribute VB_Name = "TOOL"
Option Explicit
Public SearchForm As Object
Sub search(Optional Ctrl As IRibbonControl)
    Dim myform As Object
    Set myform = New SearchForm
    myform.Show vbModeless
End Sub

Sub CreateCSV(Ctrl As IRibbonControl)
    Dim rng As Range
    Dim WorkRng As Range
    Dim X, y, cot, hang, i, j As Integer
    Dim csv, tmp As String
    On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(VNI("Choïn vuøng :"), "Select range", WorkRng.Address, Type:=8)
    If Err Then Exit Sub
    X = WorkRng.row
    y = WorkRng.Column
    cot = WorkRng.Columns.Count - 1
    hang = WorkRng.Rows.Count - 1
    
    For i = 0 To hang
        For j = 0 To cot
            If i > 0 And j > 0 Then
                csv = csv & """" & Round(Cells(X + i, y + j).Value, 2) & """" & IIf(j = cot, Chr(13), ";")
            Else
                tmp = Cells(X + i, y + j).Value
                If j = 0 Then
                    tmp = Replace(tmp, "Q1", "03")
                    tmp = Replace(tmp, "Q2", "06")
                    tmp = Replace(tmp, "Q3", "09")
                    tmp = Replace(tmp, "Q4", "12")
                End If
                csv = csv & tmp & IIf(j = cot, Chr(13), ";")
            End If
        Next j
    Next i
    'copy to clipboard
    Dim oData As DataObject
    Set oData = New DataObject
    With oData
        .SetText csv
        .PutInClipboard
    End With
    MsgBox "OK"
End Sub
Sub QuickMenu(Ctrl As IRibbonControl)
    Call ShowPopUpMenu
End Sub
Sub GroupCellsSameFormat(Ctrl As IRibbonControl)
    Dim startr As Integer, endr As Integer, lr As Integer, r As Integer
    Dim myarr
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.Selection
    If Ctrl.Tag = "clear" Then
        rng.ClearOutline
        Exit Sub
    End If
    
    Set rng = Application.InputBox(VNI("Choïn vuøng :"), "Select range", rng.Address, Type:=8)
    If Err.Number <> 0 Then Exit Sub
    startr = 0: endr = 0
    For r = rng.row To rng.row + rng.Rows.Count - 2
        myarr = Left(Trim(Range("A" & r).Value), 1)
        Select Case Ctrl.Tag
            Case "lama":
            If myarr = "I" Or myarr = "V" Or myarr = "X" Then 'dieu kien nhom theo chu la ma
            If startr = 0 Then
                startr = r + 1
            Else
                endr = r - 1
                Range(Cells(startr, 1), Cells(endr, 1)).Rows.Group
                startr = r + 1
            End If
        End If
        Case "latinh":
        If myarr = "A" Or myarr = "B" Or myarr = "C" Or myarr = "D" Or myarr = "E" Then 'dieu kien nhom theo chu la tin
        If startr = 0 Then
            startr = r + 1
        Else
            endr = r - 1
            Range(Cells(startr, 1), Cells(endr, 1)).Rows.Group
            startr = r + 1
        End If
    End If
    Case "arap":
    If myarr = 1 Or myarr = 2 Or myarr = 3 Or myarr = 4 Or myarr = 5 Or myarr = 6 Or myarr = 7 Or myarr = 8 Or myarr = 9 Then 'dieu kien nhom
    If startr = 0 Then
        startr = r + 1
    Else
        endr = r - 1
        Range(Cells(startr, 1), Cells(endr, 1)).Rows.Group
        startr = r + 1
    End If
End If
Case Else
Exit Sub
End Select
Next r
Range(Cells(startr, 1), Cells(r, 1)).Rows.Group
End Sub
Sub Caculator(Ctrl As IRibbonControl)
    Dim rng As Range
    Dim WorkRng As Range
    Dim X, y, cot, hang, i As Integer
    On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(VNI("Choïn vuøng :"), "Select range", WorkRng.Address, Type:=8)
    If Err Then Exit Sub
    X = WorkRng.row
    y = WorkRng.Column
    cot = WorkRng.Columns.Count
    hang = WorkRng.Rows.Count - 1
    If cot = 2 And X > 1 And hang > 0 Then
        '''' Ti trong
        Range(Cells(X - 1, y + 2), Cells(X - 1, y + 3)).Merge
        Range(Cells(X - 1, y + 2), Cells(X - 1, y + 3)).Value = "T" & ChrW(7927) & " tr" & ChrW(7885) & "ng"
        
        Cells(X, y + 2).Value = Cells(X, y).Value
        Cells(X, y + 3).Value = Cells(X, y + 1).Value
        Range(Cells(X - 1, y + 2), Cells(X, y + 3)).HorizontalAlignment = xlCenter
        
        For i = 1 To hang
            Cells(X + i, y + 2).Value = Cells(X + i, y).Value / Cells(X + hang, y).Value
            Cells(X + i, y + 2).NumberFormat = "0.0%"
            Cells(X + i, y + 3).Value = Cells(X + i, y + 1).Value / Cells(X + hang, y + 1).Value
            Cells(X + i, y + 3).NumberFormat = "0.0%"
        Next i
        
        ''' Thay doi
        Range(Cells(X - 1, y + 4), Cells(X - 1, y + 5)).Merge
        Range(Cells(X - 1, y + 4), Cells(X - 1, y + 5)).Value = "Thay " & ChrW(273) & ChrW(7893) & "i"
        
        Cells(X, y + 4).Value = "(2) - (1)"
        Cells(X, y + 5).Value = "%"
        Range(Cells(X - 1, y + 4), Cells(X, y + 5)).HorizontalAlignment = xlCenter
        
        For i = 1 To hang
            Cells(X + i, y + 4).Value = Cells(X + i, y + 1).Value - Cells(X + i, y).Value
            Cells(X + i, y + 4).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
            Cells(X + i, y + 5).Value = Cells(X + i, y + 4).Value / Cells(X + i, y).Value
            Cells(X + i, y + 5).NumberFormat = "0.0%"
        Next i
        Range(Cells(X - 1, y + 2), Cells(X + hang, y + 5)).Borders.LineStyle = 1
    Else
        MsgBoxUni VNI("Vuøng döõ lieäu khoâng phuø hôïp!"), vbInformation, _
        VNI("Thaát baïi!")
    End If
End Sub

Sub Show_Industry(Ctrl As IRibbonControl)
    On Error Resume Next
    Dim MyWB As Workbook
    Set MyWB = ActiveWorkbook
    
    ThisWorkbook.IsAddin = False
    Application.ScreenUpdating = False
    If MyWB Is Nothing Then
        Sheets(1).Copy
    Else
        Sheets(1).Copy Before:=MyWB.Sheets(1)
    End If
    ThisWorkbook.IsAddin = True
    ThisWorkbook.Save
    Application.ScreenUpdating = False
End Sub
Sub Show_Option(Ctrl As IRibbonControl)
    On Error Resume Next
    Dim MyWB As Workbook
    Set MyWB = ActiveWorkbook
    
    ThisWorkbook.IsAddin = False
    Application.ScreenUpdating = False
    If MyWB Is Nothing Then
        Sheets(2).Copy
    Else
        Sheets(2).Copy Before:=MyWB.Sheets(1)
    End If
    ThisWorkbook.IsAddin = True
    ThisWorkbook.Save
    Application.ScreenUpdating = False
End Sub
Sub Show_Help(Ctrl As IRibbonControl)
    On Error Resume Next
    Dim MyWB As Workbook
    Set MyWB = ActiveWorkbook
    
    ThisWorkbook.IsAddin = False
    Application.ScreenUpdating = False
    If MyWB Is Nothing Then
        Sheets(3).Copy
    Else
        Sheets(3).Copy Before:=MyWB.Sheets(1)
    End If
    ThisWorkbook.IsAddin = True
    ThisWorkbook.Save
    Application.ScreenUpdating = False
End Sub
Sub GotoTop(Ctrl As IRibbonControl)
    On Error Resume Next
    Calculate
    Application.ActiveWindow.ScrollRow = 1
    Application.ActiveWindow.ScrollColumn = 1
End Sub
Sub Show_Hide_Tabs(Ctrl As IRibbonControl)
    On Error Resume Next
    ActiveWindow.DisplayWorkbookTabs = Not ActiveWindow.DisplayWorkbookTabs
    Application.ScreenUpdating = True
End Sub
'
Sub ChangeToDot(Ctrl As IRibbonControl)
    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(VNI("Choïn vuøng :"), "Select range", WorkRng.Address, Type:=8)
    If Err Then Exit Sub
    If WorkRng.Rows.Count < 10000 Then
        For Each rng In WorkRng
            If Left(rng.Value, 3) = VNI("Quyù") Then
                rng.Value = Left(rng.Value, 10)
            Else
                rng.Value = Replace(rng.Value, ",", "")
                rng.Value = CDec(Replace(rng.Value, ".", ","))
            End If
        Next
        WorkRng.WrapText = False
        WorkRng.Columns.AutoFit
        WorkRng.Rows.AutoFit
    Else
        MsgBoxUni VNI("Vuøng döõ lieäu khoâng phuø hôïp hoaëc quaù lôùn!"), vbInformation, _
        VNI("Thaát baïi!")
    End If
End Sub
Sub RemoveWrapText()
    Cells.Select
    Selection.WrapText = False
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
End Sub
Sub Website(Ctrl As IRibbonControl)
    Dim sticker, san As String
    Dim Add
    On Error Resume Next
    
    If Ctrl.Tag = "IBOARD" Then
        ThisWorkbook.FollowHyperlink "https://iboard.ssi.com.vn/"
        Exit Sub
    End If
    
    If Not IsMaCP(ActiveCell.Value) And Len(ActiveCell.Value) <> 8 Then
        MsgBoxUni VNI("Choïn maõ coå phieáu, CW khoâng hôïp leä!"), vbInformation, VNI("Thaát baïi!")
        Exit Sub
    End If
    
    Select Case Ctrl.Tag
        Case "Vietstock":
        sticker = "https://finance.vietstock.vn/" & UCase(ActiveCell.Value) & "-abc.htm"
        Case "Fialda":
        sticker = "https://fwt.fialda.com/co-phieu/" & UCase(ActiveCell.Value) & "/tongquan"
        Case "Wichart":
        sticker = "https://wichart.vn/" & UCase(ActiveCell.Value)
        Case "24hmoney":
        sticker = "https://24hmoney.vn/stock/" & UCase(ActiveCell.Value)
        Case "TVSI":
        sticker = "https://finance.tvsi.com.vn/Enterprises/OverView?symbol=" & UCase(ActiveCell.Value)
        Case "VNDIRECT":
        sticker = "https://dstock.vndirect.com.vn/tong-quan/" & UCase(ActiveCell.Value)
        Case "Cafef":
        Add = ActiveCell.AddressLocal
        san = Application.Evaluate("=STOCKVN(" & Add & ", 2)")
        If san = "HOSTC" Then san = "HOSE"
        sticker = "http://s.cafef.vn/" & san & "/" & UCase(ActiveCell.Value) & "-abc.chn"
        Case "Tradingview":
        Add = ActiveCell.AddressLocal
        san = Application.Evaluate("=STOCKVN(" & Add & ", 2)")
        If san = "HOSTC" Then san = "HOSE"
        If san = "HASTC" Then san = "HNX"
        sticker = "https://vn.tradingview.com/chart/?symbol=" & san & "%3A" & UCase(ActiveCell.Value)
        Case "stockchart":
        sticker = "https://stockchart.vietstock.vn/?stockcode=" & UCase(ActiveCell.Value)
        Case "fireant":
        sticker = "https://fireant.vn/charts/content/symbols/" & UCase(ActiveCell.Value)
        Case "tcbs":
        sticker = "https://static.tcbs.com.vn/oneclick/" & UCase(ActiveCell.Value) & ".pdf"
    End Select
    ThisWorkbook.FollowHyperlink sticker
    ActiveCell.Value = UCase(ActiveCell.Value)
End Sub
Sub KQKD(Ctrl As IRibbonControl)
    If Ctrl.Tag = "Vietstock.vn" Then ThisWorkbook.FollowHyperlink "https://finance.vietstock.vn/ket-qua-kinh-doanh/"
    If Ctrl.Tag = "Wichart.vn" Then ThisWorkbook.FollowHyperlink "https://wichart.vn/kqkd"
End Sub
Sub Tim_Kiem_Trung_Lap(Ctrl As IRibbonControl)
    On Error GoTo ends
    Dim Range1 As Range, Range2 As Range, Rng1 As Range, Rng2 As Range, outRng As Range
    Dim xvalue, xTitleId
    xTitleId = VNI("Tìm döõ lieäu truøng laëp")
    Set Range1 = Application.Selection
    Set Range1 = Application.InputBox(VNI("Choïn vuøng 1 :"), xTitleId, Range1.Address, Type:=8)
    Set Range2 = Application.InputBox(VNI("Choïn vuøng 2 :"), xTitleId, Type:=8)
    Application.ScreenUpdating = False
    For Each Rng1 In Range1
        xvalue = Rng1.Value
        For Each Rng2 In Range2
            If xvalue = Rng2.Value Then
                If outRng Is Nothing Then
                    Set outRng = Rng1
                Else
                    Set outRng = Application.Union(outRng, Rng1)
                End If
            End If
        Next
    Next
    outRng.Font.Bold = True
    outRng.Select
    Application.ScreenUpdating = True
ends:
    MsgBoxUni VNI("Choïn vuøng döõ lieäu bò loãi!"), vbCritical, VNI("Thaát baïi")
    Exit Sub
End Sub
Sub TradingviewList(Ctrl As IRibbonControl)
    Dim rng As Range
    Dim WorkRng As Range
    Dim List As String
    Dim Add, san
    On Error Resume Next
    
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox(VNI("Choïn vuøng :"), "Select range", WorkRng.Address, Type:=8)
    If Err Then Exit Sub
    If WorkRng.Cells.Count < 100 Then
        For Each rng In WorkRng
            If IsMaCP(rng.Value) Then
                Add = rng.AddressLocal
                san = Application.Evaluate("=STOCKVN(" & Add & ", 2)")
                If san = "HOSTC" Then san = "HOSE"
                If san = "HASTC" Then san = "HNX"
                List = List & san & ":" & rng.Value & ","
            End If
        Next
    Else
        MsgBoxUni VNI("Vuøng döõ lieäu khoâng phuø hôïp hoaëc quaù lôùn!"), vbInformation, _
        VNI("Thaát baïi!")
        Exit Sub
    End If
    
    'copy to clipboard
    Dim oData As DataObject
    Set oData = New DataObject
    With oData
        .SetText List
        .PutInClipboard
    End With
    MsgBoxUni VNI("Ñaõ copy danh saùch maõ thaønh coâng, coù theå theâm vaøo danh saùch treân Tradingview.com !"), vbInformation, _
    VNI("Thaønh coâng!")
End Sub

Sub InsertLogoToCell(Ctrl As IRibbonControl)
    On Error GoTo ends
    Dim URL As String
    Dim pic As Picture
    Dim MCK, Destination As Range
    Dim xvalue, xTitleId
    
    xTitleId = VNI("Load Logo")
    Set MCK = Application.Selection
    Set MCK = Application.InputBox(VNI("Select MCK :"), xTitleId, MCK.Address, Type:=8)
    Set Destination = Application.InputBox(VNI("Select Range :"), xTitleId, Type:=8)
    

    For Each pic In ActiveSheet.Pictures
        If Not Application.Intersect(pic.TopLeftCell, Range(Destination.Offset(0, 0).Address)) Is Nothing Then pic.Delete
    Next pic
    
    URL = "https://wichart.vn/_next/image?url=%2Fimages%2Flogo-dn%2F" & MCK.Value & ".jpeg&w=1200&q=75"
    
    With ActiveSheet.Pictures.Insert(URL)
        .Top = Destination.Offset(0, 0).Top
        .Left = Destination.Offset(0, 0).Left
        .ShapeRange.LockAspectRatio = msoFalse
        .ShapeRange.Height = Destination.Height
        .ShapeRange.Width = Destination.Width
        .Cut
    End With
    Destination.PasteSpecial
       If Err.Number <> 0 Then
        MsgBoxUni VNI("Taûi logo thaát baïi!"), vbCritical, VNI("Thaát baïi")
        Exit Sub
    End If
ends:
    Exit Sub
End Sub
''''''''''''''''''''''''''''
Private Function IsDateString(ByVal str As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^([0-9]{1,2})\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s*([0-9]{4})$"
    IsDateString = regEx.Test(str)
End Function
Private Function RemoveTextAfterAt(originalText As String) As String
    Dim positionOfAt As Long
    positionOfAt = InStr(originalText, " at")
    If positionOfAt > 0 Then
        RemoveTextAfterAt = Left(originalText, positionOfAt - 1)
    Else
        RemoveTextAfterAt = originalText
    End If
End Function
Sub SplitData(Ctrl As IRibbonControl)
    On Error Resume Next
    Dim oData As DataObject
    Set oData = New DataObject
        oData.GetFromClipboard
        
    Dim myVariable As String
    myVariable = oData.GetText
    If Len(myVariable) < 5 Then
    MsgBox "No Data"
    Exit Sub
    End If
    
    Dim dataArray() As String
    dataArray = Split(Trim(myVariable), vbCrLf)
    

    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Chon vung:", "Select range", WorkRng.Address, Type:=8)
    If Err Then Exit Sub
    
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, WorkRng, , xlYes)
    tbl.Name = "Geocosmis"
    tbl.HeaderRowRange.Cells(1, 1).Value = "Date"
    tbl.HeaderRowRange.Cells(1, 2).Value = "Event"
    ' Ðua các ph?n t? c?a m?ng vào b?ng d? li?u
    Dim i As Long
    Dim eventDate As String
    Dim eventName As String
    
    For i = 0 To UBound(dataArray)
        
        
        If IsDateString(dataArray(i)) Then
            eventDate = dataArray(i)
        Else
            eventName = RemoveTextAfterAt(dataArray(i))
            If Trim(eventName) <> "" Then
            tbl.ListRows.Add
            tbl.Range(tbl.ListRows.Count + 1, 1).Value = eventDate
            tbl.Range(tbl.ListRows.Count + 1, 2).Value = eventName
            End If
        End If
    Next i

    tbl.ListColumns(1).DataBodyRange.NumberFormat = "dd/mm/yyyy"
End Sub
'Copyright by Phebungphe
