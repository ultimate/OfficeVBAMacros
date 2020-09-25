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
    Debug.Print "Kürzel-Test"
    Debug.Print "  K:" & AttachmentUtil.GetKuerzel("john.doe@company-a.com") & ":K"
    Debug.Print "  K:" & AttachmentUtil.GetKuerzel("foo.bar@company-b.com") & ":K"
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

'--------------------------------------------------
' Test-Categories-Funktionen
'--------------------------------------------------
Public Sub Test_MergeCategories()

    Debug.Print MergeCategories("a;b;c;d", "c;d;e;f")
    Debug.Print MergeCategories("a;b;c;d", "")
    Debug.Print MergeCategories("", "a;b;c;d")
    
    Debug.Print MergeCategories("a;b", "b")
    Debug.Print MergeCategories("a;b", "a")
    Debug.Print MergeCategories("a;b", "a;b")
    Debug.Print MergeCategories("a;b", "b;a")
    
    Debug.Print MergeCategories("a: b & c;1: 2 & 3", "1: 2 & 3")
    Debug.Print MergeCategories("a: b & c;1: 2 & 3", "a: b & c")
    Debug.Print MergeCategories("a: b & c;1: 2 & 3", "a: b & c;1: 2 & 3")
    Debug.Print MergeCategories("a: b & c;1: 2 & 3", "1: 2 & 3;a: b & c")
        
    Debug.Print MergeCategories("a;a", "a")
    
    Dim cats As String
    cats = "1;2;3"
    Debug.Print MergeCategories(cats, "3;4")
    Debug.Print cats
End Sub

Public Sub Print_Selection()

    Debug.Print ("Print_Selection")
    Debug.Print ("------------------")
    
    For Each item In Application.ActiveExplorer.selection
        Debug.Print TypeName(item); " --> "; item.Categories
    Next item
    Debug.Print ("------------------")

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

Public Sub Test_PrintTypeInfo()

    Debug.Print ("Test_PrintTypeInfo")
    Debug.Print ("------------------")
    
    Dim rec As Integer
    Dim item As Object
    
    Set olA = New Outlook.Application
    Set olNS = olA.GetNamespace("MAPI")
    
    For Each item In Application.ActiveExplorer.selection
    
        Debug.Print TypeName(item)
        
        Dim t As TLI.TLIApplication
        Set t = New TLI.TLIApplication
        
        Dim ti As TLI.TypeInfo
        Set ti = t.InterfaceInfoFromObject(item)
        
        Dim mi As TLI.MemberInfo, i As Long
        For Each mi In ti.Members
            Select Case mi.DescKind
            Case TLI.DESCKIND_VARDESC:
                Debug.Print "  VAR:   ", mi.name, mi.Value
            Case TLI.DESCKIND_FUNCDESC:
                Select Case mi.InvokeKind
                Case TLI.INVOKE_PROPERTYGET:
                    Debug.Print "  FUNC:  ", mi.name
                End Select
                    
            End Select
        Next mi
        
    Next item

End Sub
