Attribute VB_Name = "UIExtension"
'--------------------------------------------------
' Zeige den Kategorie-Auswahl-Dialog an
' (Verwende diese Funktion, wenn eine Mail im separaten Fenster angezeigt wird)
'--------------------------------------------------
Public Sub ShowCatDialog()
    Dim mail As MailItem
    Set mail = Application.ActiveInspector.CurrentItem
    mail.ShowCategoriesDialog
End Sub

'--------------------------------------------------
' Zeige den Kategorie-Auswahl-Dialog an
' (Verwende diese Funktion, wenn eine oder mehrere Mails in einem Ordner selektiert sind
' und/oder im Lesebereich angezeigt werden)
'--------------------------------------------------
Public Sub ShowCatDialog2()
    Dim mail As MailItem
    Dim count As Integer
    Dim cats As String
    
    count = 0
    
    For Each mail In Application.ActiveExplorer.selection
        If count = 0 Then
            mail.ShowCategoriesDialog
            cats = mail.Categories
        Else
            mail.Categories = cats
            mail.Save
        End If
        count = count + 1
    Next mail
End Sub

'--------------------------------------------------
' Verschlüsselung entfernen
' (Verwende diese Funktion, wenn eine oder mehrere Mails in einem Ordner selektiert sind
' und/oder im Lesebereich angezeigt werden)
'--------------------------------------------------
Public Sub RemoveEncryption()
    If MsgBox("Verschlüsselung wirklich von allen ausgewählten Mails entfernen?", vbYesNo) = vbYes Then
        Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
                
        Dim mail As MailItem
        
        For Each mail In Application.ActiveExplorer.selection
            mail.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, 0
            mail.Save
        Next mail
   End If
End Sub
