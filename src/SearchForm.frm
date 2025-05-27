VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "Search Forn"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "SearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private sArray
Dim Dic As Object

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim oName()
    Dim i As Integer
    Set Dic = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    ReDim oName(0 To ActiveSheet.ListObjects.Count - 1)
    For Each lo In ActiveSheet.ListObjects
        oName(i) = lo.Name
        i = i + 1
    Next
   
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)
    
    sArray = tbl.Range.Value
    LBDMHH.List() = sArray
    
    CBDMHH.List = oName
    CBDMHH.Value = oName(0)

End Sub

Private Sub CBDMHH_Change()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(CBDMHH.Value)
    
    sArray = tbl.Range.Value
    LBDMHH.List() = sArray
End Sub
Private Sub TXTFIND_Change()
    On Error Resume Next
    Call WaitFor(0.05)
    Dim arr, i
    
    arr = Filter2DArray(sArray, 1, TXTFIND.Text, True)
    
    If Not IsArray(arr) Then LBDMHH.Clear: Exit Sub
    
    LBDMHH.List() = IIf(Trim(TXTFIND.Text) = "", sArray, arr)
    
    For i = 0 To Me.LBDMHH.ListCount - 1
        
        If Dic.Exists(Me.LBDMHH.List(i, 0)) Then Me.LBDMHH.Selected(i) = True
        
    Next
    
End Sub

Private Sub LBDMHH_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub LBDMHH_Click()
    Call WaitFor(0.05)
    Dim Id, i
    
    Id = LBDMHH.ListIndex
    
    With Me.LBDMHH
        
        On Error Resume Next
        
        If Id > 0 Then ActiveCell.Value = .List(Id, 0)
        If .Selected(Id) Then

            If Not Dic.Exists(.List(Id, 0)) Then
                
                Dic.Add .List(Id, 0), .List(Id, 0) & ";" & .List(Id, 1)
                
            End If
            
        Else
            
            Dic.Remove (.List(Id, 0))
            
        End If
        
    End With
End Sub
Private Sub UserForm_Terminate()
    On Error Resume Next
    Set Dic = Nothing
    Erase sArray
End Sub
Function Filter2DArray(ByVal sArray, ByVal ColIndex As Long, ByVal FindStr As String, ByVal HasTitle As Boolean)
    
    Dim tmpArr, i As Long, j As Long, arr, Dic, TmpStr, tmp, Chk As Boolean, TmpVal As Double
    
    On Error Resume Next
    
    Set Dic = CreateObject("Scripting.Dictionary")
    
    tmpArr = sArray
    
    ColIndex = ColIndex + LBound(tmpArr, 2) - 1
    
    Chk = (InStr("><=", Left(FindStr, 1)) > 0)
    
    For i = LBound(tmpArr, 1) - HasTitle To UBound(tmpArr, 1)
        
        If Chk Then
            
            TmpVal = CDbl(tmpArr(i, ColIndex))
            
            If Evaluate(TmpVal & FindStr) Then Dic.Add i, ""
            
        Else
            
            If InStr(UCase(tmpArr(i, ColIndex)), UCase(FindStr)) Then Dic.Add i, ""
            
        End If
        
    Next
    
    If Dic.Count > 0 Then
        
        tmp = Dic.Keys
        
        ReDim arr(LBound(tmpArr, 1) To UBound(tmp) + LBound(tmpArr, 1) - HasTitle, LBound(tmpArr, 2) To UBound(tmpArr, 2))
        
        For i = LBound(tmpArr, 1) - HasTitle To UBound(tmp) + LBound(tmpArr, 1) - HasTitle
            
            For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
                
                arr(i, j) = tmpArr(tmp(i - LBound(tmpArr, 1) + HasTitle), j)
                
            Next
            
        Next
        
        If HasTitle Then
            
            For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
                
                arr(LBound(tmpArr, 1), j) = tmpArr(LBound(tmpArr, 1), j)
                
            Next
            
        End If
        
    End If
    
    Filter2DArray = arr
    
End Function
Sub WaitFor(NumOfSeconds As Single)
    Dim SngSec As Single
    SngSec = Timer + NumOfSeconds
  
    Do While Timer < SngSec
        DoEvents
    Loop
End Sub


