Attribute VB_Name = "MsgUnicode"
'Khai bao cac ham API trong thu vien User32.DLL. Ch?y trong m�i tru?ng Office 32 ho?c 64-bit
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#End If

Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant = vbNullString) As VbMsgBoxResult
    'Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant, Optional HelpFile, Optional Context) As VbMsgBoxResult
    'BStrMsg,BStrTitle : La chuoi Unicode
    Dim BStrMsg, BStrTitle
    'H�m StrConv Chuyen chuoi ve ma Unicode
    BStrMsg = StrConv(PromptUni, vbUnicode)
    BStrTitle = StrConv(TitleUni, vbUnicode)
    
    MsgBoxUni = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons)
End Function
Function TCVN3toUNICODE(vnstr As String)
    Dim c As String, i As Integer
    For i = 1 To Len(vnstr)
        c = Mid(vnstr, i, 1)
        Select Case c
            Case "a": c = ChrW$(97)
            Case "�": c = ChrW$(225)
            Case "�": c = ChrW$(224)
            Case "�": c = ChrW$(7843)
            Case "�": c = ChrW$(227)
            Case "�": c = ChrW$(7841)
            Case "�": c = ChrW$(259)
            Case "�": c = ChrW$(7855)
            Case "�": c = ChrW$(7857)
            Case "�": c = ChrW$(7859)
            Case "�": c = ChrW$(7861)
            Case "�": c = ChrW$(7863)
            Case "�": c = ChrW$(226)
            Case "�": c = ChrW$(7845)
            Case "�": c = ChrW$(7847)
            Case "�": c = ChrW$(7849)
            Case "�": c = ChrW$(7851)
            Case "�": c = ChrW$(7853)
            Case "e": c = ChrW$(101)
            Case "�": c = ChrW$(233)
            Case "�": c = ChrW$(232)
            Case "�": c = ChrW$(7867)
            Case "�": c = ChrW$(7869)
            Case "�": c = ChrW$(7865)
            Case "�": c = ChrW$(234)
            Case "�": c = ChrW$(7871)
            Case "�": c = ChrW$(7873)
            Case "�": c = ChrW$(7875)
            Case "�": c = ChrW$(7877)
            Case "�": c = ChrW$(7879)
            Case "o": c = ChrW$(111)
            Case "�": c = ChrW$(243)
            Case "�": c = ChrW$(242)
            Case "�": c = ChrW$(7887)
            Case "�": c = ChrW$(245)
            Case "�": c = ChrW$(7885)
            Case "�": c = ChrW$(244)
            Case "�": c = ChrW$(7889)
            Case "�": c = ChrW$(7891)
            Case "�": c = ChrW$(7893)
            Case "�": c = ChrW$(7895)
            Case "�": c = ChrW$(7897)
            Case "�": c = ChrW$(417)
            Case "�": c = ChrW$(7899)
            Case "�": c = ChrW$(7901)
            Case "�": c = ChrW$(7903)
            Case "�": c = ChrW$(7905)
            Case "�": c = ChrW$(7907)
            Case "i": c = ChrW$(105)
            Case "�": c = ChrW$(237)
            Case "�": c = ChrW$(236)
            Case "�": c = ChrW$(7881)
            Case "�": c = ChrW$(297)
            Case "�": c = ChrW$(7883)
            Case "u": c = ChrW$(117)
            Case "�": c = ChrW$(250)
            Case "�": c = ChrW$(249)
            Case "�": c = ChrW$(7911)
            Case "�": c = ChrW$(361)
            Case "�": c = ChrW$(7909)
            Case "�": c = ChrW$(432)
            Case "�": c = ChrW$(7913)
            Case "�": c = ChrW$(7915)
            Case "�": c = ChrW$(7917)
            Case "�": c = ChrW$(7919)
            Case "�": c = ChrW$(7921)
            Case "y": c = ChrW$(121)
            Case "�": c = ChrW$(253)
            Case "�": c = ChrW$(7923)
            Case "�": c = ChrW$(7927)
            Case "�": c = ChrW$(7929)
            Case "�": c = ChrW$(7925)
            Case "�": c = ChrW$(273)
            Case "A": c = ChrW$(65)
            Case "�": c = ChrW$(258)
            Case "�": c = ChrW$(194)
            Case "E": c = ChrW$(69)
            Case "�": c = ChrW$(202)
            Case "O": c = ChrW$(79)
            Case "�": c = ChrW$(212)
            Case "�": c = ChrW$(416)
            Case "I": c = ChrW$(73)
            Case "U": c = ChrW$(85)
            Case "�": c = ChrW$(431)
            Case "Y": c = ChrW$(89)
            Case "�": c = ChrW$(272)
        End Select
        TCVN3toUNICODE = TCVN3toUNICODE + c
    Next i
End Function
Function VNItoUNICODE(vnstr As String)
    Dim c As String, i As Integer
    Dim db As Boolean
    For i = 1 To Len(vnstr)
        db = False
        If i < Len(vnstr) Then
            c = Mid(vnstr, i + 1, 1)
            If c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or _
            c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or _
            c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or _
            c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or _
            c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or _
            c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Or c = "�" Then db = True
        End If
        If db Then
            c = Mid(vnstr, i, 2)
            Select Case c
                Case "a�": c = ChrW$(225)
                Case "a�": c = ChrW$(224)
                Case "a�": c = ChrW$(7843)
                Case "a�": c = ChrW$(227)
                Case "a�": c = ChrW$(7841)
                Case "a�": c = ChrW$(259)
                Case "a�": c = ChrW$(7855)
                Case "a�": c = ChrW$(7857)
                Case "a�": c = ChrW$(7859)
                Case "a�": c = ChrW$(7861)
                Case "a�": c = ChrW$(7863)
                Case "a�": c = ChrW$(226)
                Case "a�": c = ChrW$(7845)
                Case "a�": c = ChrW$(7847)
                Case "a�": c = ChrW$(7849)
                Case "a�": c = ChrW$(7851)
                Case "a�": c = ChrW$(7853)
                Case "e�": c = ChrW$(233)
                Case "e�": c = ChrW$(232)
                Case "e�": c = ChrW$(7867)
                Case "e�": c = ChrW$(7869)
                Case "e�": c = ChrW$(7865)
                Case "e�": c = ChrW$(234)
                Case "e�": c = ChrW$(7871)
                Case "e�": c = ChrW$(7873)
                Case "e�": c = ChrW$(7875)
                Case "e�": c = ChrW$(7877)
                Case "e�": c = ChrW$(7879)
                Case "o�": c = ChrW$(243)
                Case "o�": c = ChrW$(242)
                Case "o�": c = ChrW$(7887)
                Case "o�": c = ChrW$(245)
                Case "o�": c = ChrW$(7885)
                Case "o�": c = ChrW$(244)
                Case "o�": c = ChrW$(7889)
                Case "o�": c = ChrW$(7891)
                Case "o�": c = ChrW$(7893)
                Case "o�": c = ChrW$(7895)
                Case "o�": c = ChrW$(7897)
                Case "��": c = ChrW$(7899)
                Case "��": c = ChrW$(7901)
                Case "��": c = ChrW$(7903)
                Case "��": c = ChrW$(7905)
                Case "��": c = ChrW$(7907)
                Case "u�": c = ChrW$(250)
                Case "u�": c = ChrW$(249)
                Case "u�": c = ChrW$(7911)
                Case "u�": c = ChrW$(361)
                Case "u�": c = ChrW$(7909)
                Case "��": c = ChrW$(7913)
                Case "��": c = ChrW$(7915)
                Case "��": c = ChrW$(7917)
                Case "��": c = ChrW$(7919)
                Case "��": c = ChrW$(7921)
                Case "y�": c = ChrW$(253)
                Case "y�": c = ChrW$(7923)
                Case "y�": c = ChrW$(7927)
                Case "y�": c = ChrW$(7929)
                Case "A�": c = ChrW$(193)
                Case "A�": c = ChrW$(192)
                Case "A�": c = ChrW$(7842)
                Case "A�": c = ChrW$(195)
                Case "A�": c = ChrW$(7840)
                Case "A�": c = ChrW$(258)
                Case "A�": c = ChrW$(7854)
                Case "A�": c = ChrW$(7856)
                Case "A�": c = ChrW$(7858)
                Case "A�": c = ChrW$(7860)
                Case "A�": c = ChrW$(7862)
                Case "A�": c = ChrW$(194)
                Case "A�": c = ChrW$(7844)
                Case "A�": c = ChrW$(7846)
                Case "A�": c = ChrW$(7848)
                Case "A�": c = ChrW$(7850)
                Case "A�": c = ChrW$(7852)
                Case "E�": c = ChrW$(201)
                Case "E�": c = ChrW$(200)
                Case "E�": c = ChrW$(7866)
                Case "E�": c = ChrW$(7868)
                Case "E�": c = ChrW$(7864)
                Case "E�": c = ChrW$(202)
                Case "E�": c = ChrW$(7870)
                Case "E�": c = ChrW$(7872)
                Case "E�": c = ChrW$(7874)
                Case "E�": c = ChrW$(7876)
                Case "E�": c = ChrW$(7878)
                Case "O�": c = ChrW$(211)
                Case "O�": c = ChrW$(210)
                Case "O�": c = ChrW$(7886)
                Case "O�": c = ChrW$(213)
                Case "O�": c = ChrW$(7884)
                Case "O�": c = ChrW$(212)
                Case "O�": c = ChrW$(7888)
                Case "O�": c = ChrW$(7890)
                Case "O�": c = ChrW$(7892)
                Case "O�": c = ChrW$(7894)
                Case "O�": c = ChrW$(7896)
                Case "��": c = ChrW$(7898)
                Case "��": c = ChrW$(7900)
                Case "��": c = ChrW$(7902)
                Case "��": c = ChrW$(7904)
                Case "��": c = ChrW$(7906)
                Case "U�": c = ChrW$(218)
                Case "U�": c = ChrW$(217)
                Case "U�": c = ChrW$(7910)
                Case "U�": c = ChrW$(360)
                Case "U�": c = ChrW$(7908)
                Case "��": c = ChrW$(7912)
                Case "��": c = ChrW$(7914)
                Case "��": c = ChrW$(7916)
                Case "��": c = ChrW$(7918)
                Case "��": c = ChrW$(7920)
                Case "Y�": c = ChrW$(221)
                Case "Y�": c = ChrW$(7922)
                Case "Y�": c = ChrW$(7926)
                Case "Y�": c = ChrW$(7928)
            End Select
        Else
            c = Mid(vnstr, i, 1)
            Select Case c
                Case "�": c = ChrW$(417)
                Case "�": c = ChrW$(237)
                Case "�": c = ChrW$(236)
                Case "�": c = ChrW$(7881)
                Case "�": c = ChrW$(297)
                Case "�": c = ChrW$(7883)
                Case "�": c = ChrW$(432)
                Case "�": c = ChrW$(7925)
                Case "�": c = ChrW$(273)
                Case "�": c = ChrW$(416)
                Case "�": c = ChrW$(205)
                Case "�": c = ChrW$(204)
                Case "�": c = ChrW$(7880)
                Case "�": c = ChrW$(296)
                Case "�": c = ChrW$(7882)
                Case "�": c = ChrW$(431)
                Case "�": c = ChrW$(7924)
                Case "�": c = ChrW$(272)
            End Select
        End If
        VNItoUNICODE = VNItoUNICODE + c
        If db Then i = i + 1
    Next i
End Function

Function UNC(strTCVN3 As String)
    UNC = TCVN3toUNICODE(strTCVN3)
End Function
Function VNI(strVNI As String)
    VNI = VNItoUNICODE(strVNI)
End Function
Function UniVba(TxtUni As String) As String
    'Viet Tieng Viet Trong VBA
    Dim N, uni1 As String, uni2 As String
    If TxtUni = "" Then
        UniVba = """"""
    Else
        TxtUni = TxtUni & " "
        If AscW(Left(TxtUni, 1)) < 256 Then UniVba = """"
        For N = 1 To Len(TxtUni) - 1
            uni1 = Mid(TxtUni, N, 1)
            uni2 = AscW(Mid(TxtUni, N + 1, 1))
            If AscW(uni1) > 255 And uni2 > 255 Then
                UniVba = UniVba & "ChrW(" & AscW(uni1) & ") & "
                ElseIf AscW(uni1) > 255 And uni2 < 256 Then
                UniVba = UniVba & "ChrW(" & AscW(uni1) & ") & """
                ElseIf AscW(uni1) < 256 And uni2 > 255 Then
                UniVba = UniVba & uni1 & """ & "
            Else
                UniVba = UniVba & uni1
            End If
        Next
        If Right(UniVba, 4) = " & """ Then
            UniVba = Mid(UniVba, 1, Len(UniVba) - 4)
        Else
            UniVba = UniVba & """"
        End If
    End If
End Function
Function UniXmlCode(TxtUni As String) As String
    'Viet Tieng Viet Trong VBA
    Dim N, uni1 As String, uni2 As String
    If TxtUni = "" Then
        UniXmlCode = ""
    Else
        TxtUni = TxtUni & " "
        If AscW(Left(TxtUni, 1)) < 256 Then UniXmlCode = ""
        For N = 1 To Len(TxtUni) - 1
            uni1 = Mid(TxtUni, N, 1)
            uni2 = AscW(Mid(TxtUni, N + 1, 1))
            If AscW(uni1) > 128 And uni2 > 128 Then
                UniXmlCode = UniXmlCode & "&#" & AscW(uni1) & ";"
                ElseIf AscW(uni1) > 128 And uni2 < 129 Then
                UniXmlCode = UniXmlCode & "&#" & AscW(uni1) & ";"
                ElseIf AscW(uni1) < 129 And uni2 > 128 Then
                UniXmlCode = UniXmlCode & uni1 & ""
            Else
                UniXmlCode = UniXmlCode & uni1
            End If
        Next
        If Right(UniXmlCode, 4) = " " Then
            UniXmlCode = Mid(UniXmlCode, 1, Len(UniXmlCode) - 4)
        Else
            UniXmlCode = UniXmlCode & ""
        End If
    End If
End Function

'Neu muon thay the ham MsgBox cua VB, hay dung ham duoi day
'Function MsgBox(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant = vbNullString) As VbMsgBoxResult
'    Dim BStrMsg, BStrTitle
'    BStrMsg = StrConv(PromptUni, vbUnicode)
'    BStrTitle = StrConv(TitleUni, vbUnicode)
'
'    MsgBox = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons)
'End Function

Sub TestFontInRANGE()
    'Test trong Excel
    MsgBoxUni Range("B3").Value, vbInformation, Range("B4").Value
    'MsgBox Range("B3").Value, vbInformation, _
    Range("B4").Value & " - Dung ham MsgBox cua VB/VBA thi loi"
End Sub

Sub TestFontTCVN3()
    'UNC la ham chuyen tu ma TCVN3 sang Unicode
    MsgBoxUni UNC("xin th�ng b�o"), vbInformation, _
    UNC("Chu�i ��a v�o l� l� ki�u g� TCVN3")
End Sub

Sub TestFontVNI()
    'VNI la ham chuyen tu ma VNI sang Unicode
    MsgBoxUni VNI("Co�ng hoa� xa� ho�i chu� ngh�a Vie�t Nam"), vbInformation, _
    VNI("Chuo�i ��a va�o la� la� kie�u go� VNI")
    
End Sub
