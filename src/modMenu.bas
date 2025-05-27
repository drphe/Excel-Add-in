Attribute VB_Name = "modMenu"
Option Explicit
Public Const Mname As String = "MyPopUpMenu"
Public Const NameKQKD = "\KQKD.xlsm"
Public Const NameTS = "\Ma tran tam soat.xlsm"
Public Const NameDM = "\Danh muc dau tu.xlsb"
' For right click menu
'
'''
'Private Sub Workbook_Activate()
'   Call AddToCellMenu
'End Sub
'Private Sub Workbook_Deactivate()
'   Call DeleteFromCellMenu
'End Sub
''''
' for left clik menu
'Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'If Target(1, 1).value = "menu" Then Call ShowPopUpMenu
'End Sub

'''' Right menu sub/function
Function GotoPage()
    On Error Resume Next
    Dim SIC As String
    Dim Keyword As Variant
    Dim wb As Workbook
    Dim homesheet
    homesheet = 1 ' name of home page
    
    SIC = ActiveCell.Value
    Keyword = CommandBars.ActionControl.Parameter
    
    Select Case Keyword
        Case "ThongTin":
        MsgBoxUni VNI("Moät saûn phaåm cuûa Anh Pheâ daønh cho nhöõng F0" _
        & vbNewLine & "vaø coøn nhieàu hôn theá."), vbInformation, VNI("Xin chaøo baïn")
        Case "HomePage":
        Sheets(homesheet).Activate
        Case "Web":
        If IsMaCP(SIC) = "" Then
            MsgBoxUni VNI("Maõ baïn choïn khoâng hôïp leä!"), vbInformation, VNI("Thoâng baùo")
        Else
            ThisWorkbook.FollowHyperlink "https://finance.vietstock.vn/" & SIC & "-abc.htm"
        End If
        Case "iboard":
        ThisWorkbook.FollowHyperlink "https://iboard.ssi.com.vn/"
        Case "LocKQKD":
        Set wb = Workbooks.Open(ThisWorkbook.Path & NameKQKD)
        Case "MaTran":
        Set wb = Workbooks.Open(ThisWorkbook.Path & NameTS)
        Case "DanhMuc":
        Set wb = Workbooks.Open(ThisWorkbook.Path & NameDM)
        Case Else
        Call SayIt
    End Select
End Function
Sub ShowPopUpMenu()
    'Delete PopUp menu if it exist
    Call DeletePopUpMenu
    'Create the PopUpmenu
    Call initPopUpMenu
    'Show the PopUp menu
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopup
    On Error GoTo 0
End Sub
Sub initPopUpMenu()
    Dim wsSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Integer
    Dim MenuItem As CommandBarPopup
    
    'Add PopUp menu
    With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, _
        MenuBar:=False, Temporary:=True)
        
        'add a button
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Short Menu"
            .FaceId = 578
            .OnAction = ""
            .Enabled = False
            .TooltipText = "Quick Menu"
        End With
        'Create table of sheet
        Set wsSheet = ActiveSheet
        i = 1
        For Each ws In Worksheets
            With .Controls.Add(Type:=msoControlButton)
                .Caption = ChrW(272) & ChrW(7871) & "n " & ws.Name
                If ws.Name = wsSheet.Name Then
                    .FaceId = 5
                    .Enabled = False
                Else
                    .FaceId = 186
                    .Enabled = True
                End If
                .Parameter = ws.Name
                .OnAction = "ChangeSheetOnWorkbook"
                If ws.Visible <> xlSheetVisible Then .Visible = False
                If i = 1 Then .BeginGroup = True
                i = i + 1
            End With
        Next ws
    End With
End Sub
Sub DeletePopUpMenu()
    'Delete PopUp menu if it exist
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    On Error GoTo 0
End Sub
Function ChangeSheetOnWorkbook()
    On Error Resume Next
    Dim SheetName As Variant
    SheetName = CommandBars.ActionControl.Parameter
    Sheets(SheetName).Activate
End Function

Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim mysubmenu As CommandBarControl
    
    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu
    
    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")
    
    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, Id:=3, Before:=1
    ''------------------------------------------------------------------------------------------------'''
    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=2)
        .OnAction = "GotoPage"
        .FaceId = 422
        .Caption = ChrW(272) & ChrW(7871) & "n VietStock"
        .Tag = "My_Cell_Control_Tag"
        .Parameter = "Web"
    End With
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=3)
        .OnAction = "GotoPage"
        .FaceId = 1084
        .Caption = "B" & ChrW(7843) & "ng " & ChrW(273) & "i" & ChrW(7879) & "n iBoard"
        .Tag = "My_Cell_Control_Tag"
        .Parameter = "iboard"
    End With
    ' Add a custom submenu with three buttons.
    Set mysubmenu = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=4)
    
    With mysubmenu
        .Caption = "T" & ChrW(7846) & "M SOÁT C" & ChrW(7892) & " PHI" & ChrW(7870) & "U"
        .Tag = "My_Cell_Control_Tag"
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "GotoPage"
            .FaceId = 602
            .Caption = "L" & ChrW(7885) & "c danh sách"
            .Parameter = "LocKQKD"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "GotoPage"
            .FaceId = 304
            .Caption = "Ma tr" & ChrW(7853) & "n t" & ChrW(7847) & "m soát"
            .Parameter = "MaTran"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "GotoPage"
            .FaceId = 125
            .Caption = "Danh m" & ChrW(7909) & "c " & ChrW(273) & ChrW(7847) & "u t" & ChrW(432)
            .Parameter = "DanhMuc"
        End With
    End With
    '''''------------------------------------------------------------------------------''
    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim Ctrl As CommandBarControl
    
    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")
    
    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each Ctrl In ContextMenu.Controls
        If Ctrl.Tag = "My_Cell_Control_Tag" Then
            Ctrl.Delete
        End If
    Next Ctrl
    
    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(Id:=3).Delete
    On Error GoTo 0
End Sub
Sub SayIt()
    Application.Speech.Speak Selection.Value
End Sub
