Attribute VB_Name = "Test"

'--------------------------------------------------
' Test-Mail-Funktionen
'--------------------------------------------------
Public Sub Test_Attachment_Config_And_Kuerzels()
    Debug.Print "Konstanten"
    Debug.Print "  KUERZEL_FILE       >" & AttachmentConfig.KUERZEL_FILE
    Debug.Print "  ARCHIVE_FOLDER     >" & AttachmentConfig.ARCHIVE_FOLDER
    Debug.Print "  FILENAME_PATTERN   >" & AttachmentConfig.FILENAME_PATTERN
    Debug.Print "  DIRECTION_FROM     >" & AttachmentConfig.DIRECTION_FROM
    Debug.Print "  DIRECTION_TO       >" & AttachmentConfig.DIRECTION_TO
End Sub


'--------------------------------------------------
' Test-String-Funktionen
'--------------------------------------------------
Public Sub Test_String_StartsWith_EndsWith()
    Debug.Print StringUtil.StartsWith("abcd", "a")
    Debug.Print StringUtil.StartsWith("abcd", "b")
    Debug.Print StringUtil.StartsWith("abcd", "c")
    Debug.Print StringUtil.StartsWith("abcd", "d")
    Debug.Print StringUtil.StartsWith("abcd", "ab")
    Debug.Print StringUtil.StartsWith("abcd", "cd")
    
    Debug.Print StringUtil.EndsWith("abcd", "a")
    Debug.Print StringUtil.EndsWith("abcd", "b")
    Debug.Print StringUtil.EndsWith("abcd", "c")
    Debug.Print StringUtil.EndsWith("abcd", "d")
    Debug.Print StringUtil.EndsWith("abcd", "ab")
    Debug.Print StringUtil.EndsWith("abcd", "cd")
    
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", ".,?")
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", ",.?")
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", "?.,")
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", "?.")
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", ".?")
    Debug.Print StringUtil.FirstInStr("Was wäre, wenn...?", "?")
End Sub

Public Sub Test_Something()

    Debug.Print ("Test_Something")
    Debug.Print ("------------------")
    
    Dim rec As Integer
    Dim mail As MailItem
    Dim address As String
    Dim name As String
    
    Set olA = New Outlook.Application
    Set olNS = olA.GetNamespace("MAPI")
    
    For Each mail In Application.ActiveExplorer.selection
    
        For rec = 1 To mail.Recipients.count
            If (mail.Recipients.Item(rec).type = 1) Then
                
                address = GetAddress(mail.Recipients.Item(rec).addressEntry)
                name = GetName(mail.Recipients.Item(rec).addressEntry)
                Debug.Print (address)
                Debug.Print (name)
                Debug.Print ("------------------")
            End If
        Next
    Next mail

End Sub

Public Sub Test_Cert()
    Dim url1 As String
    Dim url2 As String
    Dim request As MSXML2.XMLHTTP60
    Dim result As String
    
    url1 = "https://server1.com/login"
    url2 = "https://server1.com/sjira/rest/api/latest/issue/XX-1188"
    
    Set request = New MSXML2.XMLHTTP60
    Call request.Open("POST", url1, False)
    Call request.Send("login-form-type=cert")
    result = request.responseText
    
    Debug.Print result
    
    Set request = New MSXML2.XMLHTTP60
    Call request.Open("GET", url2, False)
    Call request.Send
    result = request.responseText
    
    Debug.Print result
End Sub

Public Sub Test_LoginForm()
    Dim username As String
    Dim password As String
    
    LoginForm.UrlLabel = "http://www.example.com"
    
    LoginForm.Show (vbModal)
    
    If (LoginForm.okAction) Then
        username = LoginForm.username
        password = LoginForm.password
        Debug.Print "confirmed " & username & ":" & password
    Else
        username = LoginForm.username
        password = LoginForm.password
        Debug.Print "canceled " & username & ":" & password
    End If
    
    LoginForm.Reset (False)
End Sub
