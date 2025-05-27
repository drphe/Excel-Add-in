Attribute VB_Name = "GET_HTML_DATA"
Option Explicit
Option Compare Text
Private Function ConnectHTTP(Optional ByVal LoadForServer As Boolean = True, Optional ByVal Timeout& = 0) As Object
    On Error Resume Next
    Dim o As Object
    Set o = VBA.CreateObject("MSXML2." & VBA.IIf(LoadForServer, "Server", "") & "XMLHTTP.6.0")
    If o Is Nothing Then
        Set o = VBA.CreateObject("Microsoft." & VBA.IIf(LoadForServer, "Server", "") & "XMLHTTP")
        If o Is Nothing Then
            Set o = VBA.CreateObject("MSXML2.XMLHTTP")
            If o Is Nothing Then
                Set o = VBA.CreateObject("WinHttp.WinHttpRequest.5.1")
            End If
            If Timeout > 0 And LoadForServer And Not o Is Nothing Then
                o.SetTimeouts Timeout, Timeout, Timeout, Timeout
            End If
        End If
    End If
    If o Is Nothing Then
        VBA.Err.Raise 11120, , "Error: could not load Microsoft and MSXML2 and WinHttp Library!"
    End If
    On Error GoTo 0
    Set ConnectHTTP = o
End Function
Private Function GetHTML(URL As String) As Object
    On Error Resume Next
    Dim HTML As Object
    Dim o As Object
    Set HTML = CreateObject("htmlfile")
    Set o = ConnectHTTP
    With o
        .Open "GET", URL, False
        .Send
        HTML.Body.Innerhtml = .responseText
    End With
    Set GetHTML = HTML
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise 50501, , "Error: Load Element fail!"
        Exit Function
    End If
End Function
' ID=1 lay thong tin giao dich, ID=2 lay thong tin co ban , source cophieu68
Function STOCKTRADING(MCK As Variant, Optional Id = 1) As Variant
    On Error Resume Next
    Dim Tr As Object
    Dim Td As Object
    Dim HTML As Object
    Dim irow As Integer
    Dim icol As Integer
    Dim tempArr As Variant
    
    Set HTML = GetHTML("https://www.cophieu68.vn/snapshot.php?id=" & MCK)
    
    If Id <> 1 Then Id = 2
    ReDim tempArr(0 To 7, 0 To 1)
    irow = 0
    icol = 0
    With HTML.getElementsByTagName("table")(Id)
        For Each Tr In .Rows
            For Each Td In Tr.Cells
                tempArr(irow, icol) = Td.InnerText
                icol = icol + 1
            Next Td
            icol = 0
            irow = irow + 1
        Next Tr
    End With
    Set HTML = Nothing
    STOCKTRADING = tempArr
End Function
'Dien thong tin tu web
Private Function GETINFOSTOCK(Optional Index As Integer = 1)
    On Error Resume Next
    Dim o, Tr, Td As Object
    Dim HTML As Object
    Dim i, irow, icol As Integer
    Dim URL, SIC As String
    
    SIC = ActiveCell.Value
    If Not IsMaCP(SIC) Then Exit Function
    URL = "https://www.cophieu68.vn/snapshot.php?id=" & SIC
    
    Set HTML = CreateObject("htmlfile")
    Set o = ConnectHTTP
    With o
        .Open "GET", URL, False
        .Send
        HTML.Body.Innerhtml = .responseText
    End With
    icol = 0
    Select Case Index
        Case 2:
        i = 2
        irow = 2
        'ActiveCell.Offset(2, 0).Value = "Ch" & ChrW(7881) & " s" & ChrW(7889) & " c" & ChrW(417) & " b" & ChrW(7843) & "n"
        Case Else
        i = 1
        irow = 2
        'ActiveCell.Offset(2, 1).Value = "Giao d" & ChrW(7883) & "ch trong ngày"
    End Select
    
    With HTML.getElementsByTagName("table")(i)
        For Each Tr In .Rows
            For Each Td In Tr.Cells
                ActiveCell.Offset(irow, icol).Value = Td.InnerText
                icol = icol + 1
            Next Td
            icol = 0
            If Index = 2 Then icol = 0
            irow = irow + 1
        Next Tr
    End With
    Set HTML = Nothing
End Function
' Thong tin giao dich
Sub TTGD(Optional Ctrl As IRibbonControl)
    On Error Resume Next
    Dim MCK As String
    MCK = ActiveCell.Value
    If IsMaCP(MCK) Then
        Call GETINFOSTOCK(2)
        ActiveCell.Value = UCase(ActiveCell.Value)
    Else
        MsgBoxUni VNI("Choïn maõ coå phieáu khoâng hôïp leä!"), vbInformation, _
        VNI("Thaát baïi!")
        Exit Sub
    End If
End Sub
' Thong tin co ban
Sub getSTOCKVN(Optional Ctrl As IRibbonControl)
    On Error Resume Next
    Dim MCK, Add, san As String
    Dim ad As Range
    
    MCK = ActiveCell.Value
    Add = ActiveCell.AddressLocal
    
    If IsMaCP(MCK) Then
        ActiveCell.Offset(0, 1).Formula = STOCKVN(ActiveCell, 1) & " (" & STOCKVN(ActiveCell, 2) & ")"   ' name
        'ActiveCell.Offset(1, 0).Formula = "=STOCKVN(" & add & ", 2)" ' san
        'ActiveCell.Offset(0, 1).Formula = "=STOCKVN(" & add & ", 1) & "" ("" & STOCKVN(" & add & ", 2) & "")""" 'name
        
        ActiveCell.Offset(1, 0).Formula = STOCKVN(ActiveCell, 0)
        'ActiveCell.Offset(1, 0).Formula = "=STOCKVN(" & add & ", 0)"
        ActiveCell.Offset(1, 0).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        
        ActiveCell.Offset(1, 1).Formula = STOCKVN(ActiveCell, 3)
        'ActiveCell.Offset(1, 1).Formula = "=STOCKVN(" & add & ", 3)"
        ActiveCell.Offset(1, 1).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        
        ActiveCell.Offset(1, 2).Formula = STOCKVN(ActiveCell, 4)
        'ActiveCell.Offset(1, 2).Formula = "=STOCKVN(" & add & ", 4)"
        ActiveCell.Offset(1, 2).NumberFormat = "0.00%"
        
        Call GETINFOSTOCK(1)
        ActiveCell.Value = UCase(ActiveCell.Value)
    Else
        MsgBoxUni VNI("Choïn maõ coå phieáu khoâng hôïp leä!"), vbInformation, _
        VNI("Thaát baïi!")
        Exit Sub
    End If
End Sub
' Country Default Spreads and Risk Premiums
Function RISKPREMIUMS() As Variant
    On Error Resume Next
    Dim Tr As Object
    Dim Td As Object
    Dim HTML As Object
    Dim irow As Integer
    Dim icol As Integer
    Dim tempArr As Variant
    
    Set HTML = GetHTML("https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html")
    
    ReDim tempArr(0 To 190, 0 To 5)
    irow = 0
    icol = 0
    With HTML.getElementsByTagName("table")(0)
        For Each Tr In .Rows
            For Each Td In Tr.Cells
                tempArr(irow, icol) = Td.InnerText
                icol = icol + 1
            Next Td
            icol = 0
            irow = irow + 1
        Next Tr
    End With
    Set HTML = Nothing
    RISKPREMIUMS = tempArr
End Function
