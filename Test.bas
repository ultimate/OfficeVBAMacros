Attribute VB_Name = "Test"

'--------------------------------------------------
' Test-Mail-Funktionen
'--------------------------------------------------
Public Sub Test_Attachment_Config_And_Kuerzels()
    Debug.Print "Konstanten"
    Debug.Print "  C2035_SVN_FOLDER   >" & AttachmentConfig.C2035_SVN_FOLDER
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
End Sub

'--------------------------------------------------
' Debug
'--------------------------------------------------
Public Sub Test_Something()
    Dim mail As MailItem
    Set mail = Application.ActiveInspector.CurrentItem
    
    Debug.Print "anz=" & mail.Attachments.Count
    Debug.Print "pos=" & mail.Attachments.Item(1).position
    mail.Attachments.Item(1).position = mail.Attachments.Item(1).position + 10
    mail.Save
    
End Sub
