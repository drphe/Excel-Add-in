VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BCTC 
   Caption         =   "B�O C�O T�I CH�NH"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "BCTC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BCTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rng As Range
''''''''-----------------''''
Private Sub cmdTai_Click()
    'On Error Resume Next
    Dim deleteRange As Range
    Dim Socot, Donvi, ReportType, ToQuarter As Integer
    Dim MSIC, BC As String
    Dim typ As Boolean
    
    Application.ScreenUpdating = False
    MSIC = CStr(txtMaCP.Value)
    BC = txtBC.Value
    Socot = txtSocot.Value
    Donvi = txtDonvi.Value
    typ = short.Value
    Set rng = Range(txtLuu.Value)
    If nam.Value = True Then
        ToQuarter = 0
    Else
        ToQuarter = 1
    End If
    ' kiem tra ma cp
    If Not IsMaCP(MSIC) Then
        MsgBoxUni VNI("Ma� co� phie�u kho�ng �u�ng!"), vbInformation, VNI("Co� lo�i")
        Exit Sub
    End If
    
    ' kiem tra so cot
    If Socot < 3 Or Socot > 12 Then
        MsgBoxUni VNI("So� l���ng co�t kho�ng h��p le�!"), vbInformation, VNI("Co� lo�i")
        Exit Sub
    End If
    ' kiem tra don vi
    If Donvi < 0 Then
        MsgBoxUni VNI("Nha�p ��n v� kho�ng h��p le�"), vbInformation, VNI("Co� lo�i")
        Exit Sub
    End If
    'Chuan bi bang
    Set deleteRange = rng.Resize(200, Socot + 1)
    Call preTable(deleteRange)
    Call initFireAnt(MSIC, BC, Socot, ToQuarter, Donvi, typ, rng)
    
    Me.Hide
    rng.Select
    Application.ScreenUpdating = True
End Sub
Private Sub cmdHuy_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    With Me
        .txtMaCP.ControlTipText = VNI("Nha�p ma� co� phie�u")
        .txtDonvi.ControlTipText = VNI("Nha�p ��n v� x 1000, 1000000")
        .txtSocot.ControlTipText = VNI("Nha�p so� co�t d�� lie�u, ma�c ��nh la�86 co�t")
        .txtLuu.ControlTipText = VNI("Cho�n v� tr� l�u d�� lie�u")
    End With
End Sub
Public Sub UserForm_Initialize()
    'select range
    Me.txtLuu.DropButtonStyle = fmDropButtonStyleReduce
    Me.txtLuu.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    'Set Rng = Range("A1")
    'txtLuu.Value = Rng.Address(False, False)
    
    ' add combobox
    txtBC.ListStyle = fmListStyleOption
    txtBC.AddItem "CDKT"
    txtBC.AddItem "KQKD"
    txtBC.AddItem "LCTTTT"
    txtBC.AddItem "LCTTGT"
    txtBC.Value = "CDKT"
End Sub
Public Sub txtLuu_DropButtonClick()
    On Error Resume Next
    Me.Hide
    Set rng = Application.InputBox("Select the range", "Range Picker", txtLuu.Text, Type:=8)
    txtLuu.Value = rng.Address(False, False)
    Me.Show
End Sub
Private Sub quy_Click()
    nam.Value = Not quy.Value
End Sub
Private Sub nam_Click()
    quy.Value = Not nam.Value
End Sub
Private Sub full_Click()
    short.Value = Not full.Value
End Sub
Private Sub short_Click()
    full.Value = Not short.Value
End Sub
Private Sub txtDonvi_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub txtMaCP_Change()
    txtMaCP.Value = UCase(txtMaCP.Value)
End Sub

Private Sub txtSocot_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
        KeyAscii = 0
        Beep
    End Select
End Sub
