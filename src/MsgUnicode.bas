Attribute VB_Name = "MsgUnicode"
'Khai bao cac ham API trong thu vien User32.DLL. Ch?y trong môi tru?ng Office 32 ho?c 64-bit
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
    'Hàm StrConv Chuyen chuoi ve ma Unicode
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
            Case "¸": c = ChrW$(225)
            Case "µ": c = ChrW$(224)
            Case "¶": c = ChrW$(7843)
            Case "·": c = ChrW$(227)
            Case "¹": c = ChrW$(7841)
            Case "¨": c = ChrW$(259)
            Case "¾": c = ChrW$(7855)
            Case "»": c = ChrW$(7857)
            Case "¼": c = ChrW$(7859)
            Case "½": c = ChrW$(7861)
            Case "Æ": c = ChrW$(7863)
            Case "©": c = ChrW$(226)
            Case "Ê": c = ChrW$(7845)
            Case "Ç": c = ChrW$(7847)
            Case "È": c = ChrW$(7849)
            Case "É": c = ChrW$(7851)
            Case "Ë": c = ChrW$(7853)
            Case "e": c = ChrW$(101)
            Case "Ð": c = ChrW$(233)
            Case "Ì": c = ChrW$(232)
            Case "Î": c = ChrW$(7867)
            Case "Ï": c = ChrW$(7869)
            Case "Ñ": c = ChrW$(7865)
            Case "ª": c = ChrW$(234)
            Case "Õ": c = ChrW$(7871)
            Case "Ò": c = ChrW$(7873)
            Case "Ó": c = ChrW$(7875)
            Case "Ô": c = ChrW$(7877)
            Case "Ö": c = ChrW$(7879)
            Case "o": c = ChrW$(111)
            Case "ã": c = ChrW$(243)
            Case "ß": c = ChrW$(242)
            Case "á": c = ChrW$(7887)
            Case "â": c = ChrW$(245)
            Case "ä": c = ChrW$(7885)
            Case "«": c = ChrW$(244)
            Case "è": c = ChrW$(7889)
            Case "å": c = ChrW$(7891)
            Case "æ": c = ChrW$(7893)
            Case "ç": c = ChrW$(7895)
            Case "é": c = ChrW$(7897)
            Case "¬": c = ChrW$(417)
            Case "í": c = ChrW$(7899)
            Case "ê": c = ChrW$(7901)
            Case "ë": c = ChrW$(7903)
            Case "ì": c = ChrW$(7905)
            Case "î": c = ChrW$(7907)
            Case "i": c = ChrW$(105)
            Case "Ý": c = ChrW$(237)
            Case "×": c = ChrW$(236)
            Case "Ø": c = ChrW$(7881)
            Case "Ü": c = ChrW$(297)
            Case "Þ": c = ChrW$(7883)
            Case "u": c = ChrW$(117)
            Case "ó": c = ChrW$(250)
            Case "ï": c = ChrW$(249)
            Case "ñ": c = ChrW$(7911)
            Case "ò": c = ChrW$(361)
            Case "ô": c = ChrW$(7909)
            Case "­": c = ChrW$(432)
            Case "ø": c = ChrW$(7913)
            Case "õ": c = ChrW$(7915)
            Case "ö": c = ChrW$(7917)
            Case "÷": c = ChrW$(7919)
            Case "ù": c = ChrW$(7921)
            Case "y": c = ChrW$(121)
            Case "ý": c = ChrW$(253)
            Case "ú": c = ChrW$(7923)
            Case "û": c = ChrW$(7927)
            Case "ü": c = ChrW$(7929)
            Case "þ": c = ChrW$(7925)
            Case "®": c = ChrW$(273)
            Case "A": c = ChrW$(65)
            Case "¡": c = ChrW$(258)
            Case "¢": c = ChrW$(194)
            Case "E": c = ChrW$(69)
            Case "£": c = ChrW$(202)
            Case "O": c = ChrW$(79)
            Case "¤": c = ChrW$(212)
            Case "¥": c = ChrW$(416)
            Case "I": c = ChrW$(73)
            Case "U": c = ChrW$(85)
            Case "¦": c = ChrW$(431)
            Case "Y": c = ChrW$(89)
            Case "§": c = ChrW$(272)
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
            If c = "ù" Or c = "ø" Or c = "û" Or c = "õ" Or c = "ï" Or _
            c = "ê" Or c = "é" Or c = "è" Or c = "ú" Or c = "ü" Or c = "ë" Or _
            c = "â" Or c = "á" Or c = "à" Or c = "å" Or c = "ã" Or c = "ä" Or _
            c = "Ù" Or c = "Ø" Or c = "Û" Or c = "Õ" Or c = "Ï" Or _
            c = "Ê" Or c = "É" Or c = "È" Or c = "Ú" Or c = "Ü" Or c = "Ë" Or _
            c = "Â" Or c = "Á" Or c = "À" Or c = "Å" Or c = "Ã" Or c = "Ä" Then db = True
        End If
        If db Then
            c = Mid(vnstr, i, 2)
            Select Case c
                Case "aù": c = ChrW$(225)
                Case "aø": c = ChrW$(224)
                Case "aû": c = ChrW$(7843)
                Case "aõ": c = ChrW$(227)
                Case "aï": c = ChrW$(7841)
                Case "aê": c = ChrW$(259)
                Case "aé": c = ChrW$(7855)
                Case "aè": c = ChrW$(7857)
                Case "aú": c = ChrW$(7859)
                Case "aü": c = ChrW$(7861)
                Case "aë": c = ChrW$(7863)
                Case "aâ": c = ChrW$(226)
                Case "aá": c = ChrW$(7845)
                Case "aà": c = ChrW$(7847)
                Case "aå": c = ChrW$(7849)
                Case "aã": c = ChrW$(7851)
                Case "aä": c = ChrW$(7853)
                Case "eù": c = ChrW$(233)
                Case "eø": c = ChrW$(232)
                Case "eû": c = ChrW$(7867)
                Case "eõ": c = ChrW$(7869)
                Case "eï": c = ChrW$(7865)
                Case "eâ": c = ChrW$(234)
                Case "eá": c = ChrW$(7871)
                Case "eà": c = ChrW$(7873)
                Case "eå": c = ChrW$(7875)
                Case "eã": c = ChrW$(7877)
                Case "eä": c = ChrW$(7879)
                Case "où": c = ChrW$(243)
                Case "oø": c = ChrW$(242)
                Case "oû": c = ChrW$(7887)
                Case "oõ": c = ChrW$(245)
                Case "oï": c = ChrW$(7885)
                Case "oâ": c = ChrW$(244)
                Case "oá": c = ChrW$(7889)
                Case "oà": c = ChrW$(7891)
                Case "oå": c = ChrW$(7893)
                Case "oã": c = ChrW$(7895)
                Case "oä": c = ChrW$(7897)
                Case "ôù": c = ChrW$(7899)
                Case "ôø": c = ChrW$(7901)
                Case "ôû": c = ChrW$(7903)
                Case "ôõ": c = ChrW$(7905)
                Case "ôï": c = ChrW$(7907)
                Case "uù": c = ChrW$(250)
                Case "uø": c = ChrW$(249)
                Case "uû": c = ChrW$(7911)
                Case "uõ": c = ChrW$(361)
                Case "uï": c = ChrW$(7909)
                Case "öù": c = ChrW$(7913)
                Case "öø": c = ChrW$(7915)
                Case "öû": c = ChrW$(7917)
                Case "öõ": c = ChrW$(7919)
                Case "öï": c = ChrW$(7921)
                Case "yù": c = ChrW$(253)
                Case "yø": c = ChrW$(7923)
                Case "yû": c = ChrW$(7927)
                Case "yõ": c = ChrW$(7929)
                Case "AÙ": c = ChrW$(193)
                Case "AØ": c = ChrW$(192)
                Case "AÛ": c = ChrW$(7842)
                Case "AÕ": c = ChrW$(195)
                Case "AÏ": c = ChrW$(7840)
                Case "AÊ": c = ChrW$(258)
                Case "AÉ": c = ChrW$(7854)
                Case "AÈ": c = ChrW$(7856)
                Case "AÚ": c = ChrW$(7858)
                Case "AÜ": c = ChrW$(7860)
                Case "AË": c = ChrW$(7862)
                Case "AÂ": c = ChrW$(194)
                Case "AÁ": c = ChrW$(7844)
                Case "AÀ": c = ChrW$(7846)
                Case "AÅ": c = ChrW$(7848)
                Case "AÃ": c = ChrW$(7850)
                Case "AÄ": c = ChrW$(7852)
                Case "EÙ": c = ChrW$(201)
                Case "EØ": c = ChrW$(200)
                Case "EÛ": c = ChrW$(7866)
                Case "EÕ": c = ChrW$(7868)
                Case "EÏ": c = ChrW$(7864)
                Case "EÂ": c = ChrW$(202)
                Case "EÁ": c = ChrW$(7870)
                Case "EÀ": c = ChrW$(7872)
                Case "EÅ": c = ChrW$(7874)
                Case "EÃ": c = ChrW$(7876)
                Case "EÄ": c = ChrW$(7878)
                Case "OÙ": c = ChrW$(211)
                Case "OØ": c = ChrW$(210)
                Case "OÛ": c = ChrW$(7886)
                Case "OÕ": c = ChrW$(213)
                Case "OÏ": c = ChrW$(7884)
                Case "OÂ": c = ChrW$(212)
                Case "OÁ": c = ChrW$(7888)
                Case "OÀ": c = ChrW$(7890)
                Case "OÅ": c = ChrW$(7892)
                Case "OÃ": c = ChrW$(7894)
                Case "OÄ": c = ChrW$(7896)
                Case "ÔÙ": c = ChrW$(7898)
                Case "ÔØ": c = ChrW$(7900)
                Case "ÔÛ": c = ChrW$(7902)
                Case "ÔÕ": c = ChrW$(7904)
                Case "ÔÏ": c = ChrW$(7906)
                Case "UÙ": c = ChrW$(218)
                Case "UØ": c = ChrW$(217)
                Case "UÛ": c = ChrW$(7910)
                Case "UÕ": c = ChrW$(360)
                Case "UÏ": c = ChrW$(7908)
                Case "ÖÙ": c = ChrW$(7912)
                Case "ÖØ": c = ChrW$(7914)
                Case "ÖÛ": c = ChrW$(7916)
                Case "ÖÕ": c = ChrW$(7918)
                Case "ÖÏ": c = ChrW$(7920)
                Case "YÙ": c = ChrW$(221)
                Case "YØ": c = ChrW$(7922)
                Case "YÛ": c = ChrW$(7926)
                Case "YÕ": c = ChrW$(7928)
            End Select
        Else
            c = Mid(vnstr, i, 1)
            Select Case c
                Case "ô": c = ChrW$(417)
                Case "í": c = ChrW$(237)
                Case "ì": c = ChrW$(236)
                Case "æ": c = ChrW$(7881)
                Case "ó": c = ChrW$(297)
                Case "ò": c = ChrW$(7883)
                Case "ö": c = ChrW$(432)
                Case "î": c = ChrW$(7925)
                Case "ñ": c = ChrW$(273)
                Case "Ô": c = ChrW$(416)
                Case "Í": c = ChrW$(205)
                Case "Ì": c = ChrW$(204)
                Case "Æ": c = ChrW$(7880)
                Case "Ó": c = ChrW$(296)
                Case "Ò": c = ChrW$(7882)
                Case "Ö": c = ChrW$(431)
                Case "Î": c = ChrW$(7924)
                Case "Ñ": c = ChrW$(272)
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
    MsgBoxUni UNC("xin th«ng b¸o"), vbInformation, _
    UNC("Chuçi ®­a vµo lµ lµ kiÓu gâ TCVN3")
End Sub

Sub TestFontVNI()
    'VNI la ham chuyen tu ma VNI sang Unicode
    MsgBoxUni VNI("Coäng hoaø xaõ hoäi chuû nghóa Vieät Nam"), vbInformation, _
    VNI("Chuoãi ñöa vaøo laø laø kieåu goõ VNI")
    
End Sub
