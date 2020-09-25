Attribute VB_Name = "StringUtil"
'--------------------------------------------------------------------
' String-Utils (Hilfsfunktionen für Strings)
'--------------------------------------------------------------------

'--------------------------------------------------------------------
' Prüfe ob ein String mit einem bestimmten anderen String endet
'--------------------------------------------------------------------
Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

'--------------------------------------------------------------------
' Prüfe ob ein String mit einem bestimmten anderen String beginnt
'--------------------------------------------------------------------
Public Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

