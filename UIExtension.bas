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


'--------------------------------------------------
' Aufgaben-Status aus Jira aktualisieren
' Sucht nach dem ersten Treffer eines Jira-Links und zieht sich per JSON-API den Assignee und den Status
' Anschließend werden die Informationen in folgende Aufgabenfelder geschrieben:
' - Assignee => Kontakte
' - Status   => Abrechnungsinformationen
'--------------------------------------------------
Public Sub UpdateTasksFromJira()
    Dim task As taskItem
    Dim handled As Boolean
    Dim handleAll As Boolean
    Dim i As Integer
    Dim total As Integer
    
    Dim ns As Outlook.NameSpace
    Dim folder As MAPIFolder
    Dim taskList As Object
    
    Set ns = Application.GetNamespace("MAPI")
    Set folder = ns.GetDefaultFolder(olFolderTasks)
    Set taskList = folder.Items
    
    Dim selection As Integer
    selection = MsgBox("Tasks werden aus Jira aktualisiert" & vbNewLine & "Erledigte auch aktualisieren?", vbYesNoCancel + vbDefaultButton2, "Jira-Task-Status")
        
    If selection = vbYes Then
        total = taskList.count
        handleAll = True
    ElseIf selection = vbNo Then
        total = 0
        For Each task In taskList
            If task.status <> olTaskComplete Then
                total = total + 1
            End If
        Next
        handleAll = False
    Else
        Exit Sub
    End If
    
    i = 0
    ProgressBar.Show (vbModeless)
    Call ProgressBar.Update(i, total)
              
    'For Each task In taskList
    For Each task In taskList
        handled = False
        If task.status <> olTaskComplete Or handleAll Then
        
            ' je zu prüfendem Jira einen If-Update-Block (nach Prio sortiert)
            If Not handled Then
                handled = UpdateTaskFromJira(task, "https://server1.com/sjira/")
            End If
            If Not handled Then
                handled = UpdateTaskFromJira(task, "https://server2.com/sjira")
            End If
        
            i = i + 1
            Call ProgressBar.Update(i, total)
        End If
    Next
    ProgressBar.Hide
End Sub

