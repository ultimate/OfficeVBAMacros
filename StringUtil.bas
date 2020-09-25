Attribute VB_Name = "StringUtil"
'--------------------------------------------------------------------
' String-Utils (Hilfsfunktionen f�r Strings)
'--------------------------------------------------------------------

'--------------------------------------------------------------------
' Pr�fe ob ein String mit einem bestimmten anderen String endet
'--------------------------------------------------------------------
Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

'--------------------------------------------------------------------
' Pr�fe ob ein String mit einem bestimmten anderen String beginnt
'--------------------------------------------------------------------
Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function


'--------------------------------------------------------------------
' Finde das erste Zeichen aus einer Auswahl aus Zeichen in einem String
' und gibt dessen Index zur�ck
' z. B. StringUtil.FirstInStr("Was w�re, wenn...?", ".,?")  => 9
'       StringUtil.FirstInStr("Was w�re, wenn...?", ".?")   => 15
'       StringUtil.FirstInStr("Was w�re, wenn...?", "?")    => 18
'--------------------------------------------------------------------
Public Function FirstInStr(str As String, chars As String) As Long
    Dim index As Long
    Dim c As Integer
    
    FirstInStr = 0
    
    For c = 1 To Len(chars)
        index = InStr(str, Mid(chars, c, 1))
        'Debug.Print Mid(chars, c, 1) & " @ " & index
        If index <> 0 And (index < FirstInStr Or FirstInStr = 0) Then
            FirstInStr = index
        End If
    Next
End Function
