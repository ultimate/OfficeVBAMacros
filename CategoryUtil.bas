Attribute VB_Name = "CategoryUtil"


'--------------------------------------------------------------------
' Merge to category strings
' "a;b;c;d" + "c;d;e;f" => "a;b;c;d;e;f"
'--------------------------------------------------------------------
Public Function MergeCategories(cats1 As String, cats2 As String) As String
    Dim cats() As String
    cats = Split("", ";") ' create an empty array of 0 length
    
    Dim found As Boolean
      
    For Each c1 In Split(cats1, ";")
        found = False
        For Each c2 In cats
            If c2 = c1 Then
                found = True
            End If
        Next c2
        If Not found Then
            ReDim Preserve cats(0 To UBound(cats) + 1)
            cats(UBound(cats)) = c1
        End If
    Next c1
    
    For Each c2 In Split(cats2, ";")
        found = False
        For Each c1 In cats
            If c1 = c2 Then
                found = True
            End If
        Next c1
        If Not found Then
            ReDim Preserve cats(0 To UBound(cats) + 1)
            cats(UBound(cats)) = c2
        End If
    Next c2
    
    MergeCategories = Join(cats, ";")
    
End Function

'--------------------------------------------------------------------
' Get a merged string of all categories for the given Conversation
' Start calling with no item, item will then be determined automatically
' and method called recursively.
'--------------------------------------------------------------------
Public Function GetCategories(conv As Conversation, Optional item As Variant = Null) As String
    Dim items As SimpleItems
    Dim convMail As MailItem
    Dim convMeeting As MeetingItem
    Dim convAppointment As AppointmentItem
    Dim convPost As PostItem
    
    If Not IsNull(item) Then
        Set items = conv.GetChildren(item)
    Else
        Set items = conv.GetRootItems
    End If
    
    On Error Resume Next
     
    For Each convChild In items
        If TypeOf convChild Is MailItem Then
            Set convMail = convChild
            Debug.Print "MailItem:        "; convMail.ReceivedTime; " "; convMail.Categories
            GetCategories = MergeCategories(GetCategories, convMail.Categories)
        ElseIf TypeOf convChild Is MeetingItem Then
            Set convMeeting = convChild
            Debug.Print "MeetingItem:     "; convMeeting.SentOn; " "; convMeeting.Categories
            GetCategories = MergeCategories(GetCategories, convMeeting.Categories)
        ElseIf TypeOf convChild Is AppointmentItem Then
            Set convAppointment = convChild
            Debug.Print "AppointmentItem: "; convAppointment.ReplyTime; " "; convAppointment.Categories
            GetCategories = MergeCategories(GetCategories, convAppointment.Categories)
        ElseIf TypeOf convChild Is PostItem Then
            Set convPost = convChild
            Debug.Print "PostItem:        "; convPost.ReceivedTime; " "; convPost.Categories
            GetCategories = MergeCategories(GetCategories, convPost.Categories)
        End If
        GetCategories = MergeCategories(GetCategories, GetCategories(conv, convChild))
    Next convChild
End Function

'--------------------------------------------------------------------
' Set a categories for all items of the given Conversation
' Start calling with no item, item will then be determined automatically
' and method called recursively.
'--------------------------------------------------------------------
Public Sub SetCategories(cats As String, conv As Conversation, Optional item As Variant = Null)
    Dim items As SimpleItems
    Dim convMail As MailItem
    Dim convMeeting As MeetingItem
    Dim convAppointment As AppointmentItem
    Dim convPost As PostItem
    
    If Not IsNull(item) Then
        Set items = conv.GetChildren(item)
    Else
        Set items = conv.GetRootItems
    End If
    
    On Error Resume Next
    
    For Each convChild In items
        If TypeOf convChild Is MailItem Then
            Set convMail = convChild
            convMail.Categories = cats
            convMail.Save
        ElseIf TypeOf convChild Is MeetingItem Then
            Set convMeeting = convChild
            convMeeting.Categories = cats
            convMeeting.Save
        ElseIf TypeOf convChild Is AppointmentItem Then
            Set convAppointment = convChild
            convAppointment.Categories = cats
            convAppointment.Save
        ElseIf TypeOf convChild Is PostItem Then
            Set convPost = convChild
            convPost.Categories = cats
            convPost.Save
        End If
        Call SetCategories(cats, conv, convChild)
    Next convChild
End Sub
