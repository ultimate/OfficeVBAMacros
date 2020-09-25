Attribute VB_Name = "UIExtension"
'--------------------------------------------------
' Zeige den Kategorie-Auswahl-Dialog an
'--------------------------------------------------
Public Sub ShowCatDialog()
    Dim mail As Object
    Set mail = Application.ActiveInspector.CurrentItem
    mail.ShowCategoriesDialog
End Sub
