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
' Zeige den Kategorie-Auswahl-Dialog an
' (Verwende diese Funktion, wenn eine oder mehrere Mails in einem Ordner selektiert sind
' und/oder im Lesebereich angezeigt werden)
'--------------------------------------------------
Public Sub CategorizeConversations()
    Dim item
    Dim cats As String
    Dim conv As Conversation
    Dim convMail As MailItem
    Dim convMeeting As MeetingItem
    
    For Each item In Application.ActiveExplorer.selection
        Set conv = item.GetConversation
        If Not IsNull(conv) Then
            ' get categories
            cats = GetCategories(conv)
            Debug.Print "setting categories to: "; cats
            ' apply categories
            Call SetCategories(cats, conv)
        End If
    Next item
End Sub

'--------------------------------------------------
' Verschl�sselung entfernen
' (Verwende diese Funktion, wenn eine oder mehrere Mails in einem Ordner selektiert sind
' und/oder im Lesebereich angezeigt werden)
'--------------------------------------------------
Public Sub RemoveEncryption()
    If MsgBox("Verschl�sselung wirklich von allen ausgew�hlten Mails entfernen?", vbYesNo) = vbYes Then
        Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
                
        Dim mail As MailItem
        
        For Each mail In Application.ActiveExplorer.selection
            mail.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, 0
            mail.Save
        Next mail
   End If
End Sub

'--------------------------------------------------
' Konvertiere einen nicht-ganzt�gigen Termin in einen ganzt�gigen Termin
' (Die Funktion ist erforderlich, wenn man Termine �ndern m�chte, bei denen man eingeladen
' wurde, weil die Funktion dann nicht per UI verf�gbar ist.)
'--------------------------------------------------
Public Sub ConvertAppointmentAllDay()
    Dim appointment As AppointmentItem
    
    For Each appointment In Application.ActiveExplorer.selection
        appointment.AllDayEvent = True
        appointment.Save
    Next appointment
End Sub

'--------------------------------------------------
' Konvertiere einen ganzt�gigen Termin in einen nicht-ganzt�gigen Termin
' (Die Funktion ist erforderlich, wenn man Termine �ndern m�chte, bei denen man eingeladen
' wurde, weil die Funktion dann nicht per UI verf�gbar ist.)
'--------------------------------------------------
Public Sub ConvertAppointmentNotAllDay()
    Dim appointment As AppointmentItem
    
    For Each appointment In Application.ActiveExplorer.selection
        appointment.AllDayEvent = False
        appointment.Save
    Next appointment
End Sub

