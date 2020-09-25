Attribute VB_Name = "JiraUtil"
'--------------------------------------------------
' Definition der benötigen Konstanten und Konfiguration
'--------------------------------------------------
' JIRA API Sub-URL
Public Const JIRA_API As String = "rest/api/latest/issue/"
' JSON identifier
Public Const JSON_ASSIGNEE As String = "assignee"
Public Const JSON_ASSIGNEE_DISP As String = "displayName"
Public Const JSON_STATUS As String = "status"
Public Const JSON_STATUS_DISP As String = "name"
Public Const JSON_SPRINT As String = "customfield_10461"
Public Const JSON_SPRINT_DISP As String = "name"
Public Const JSON_SPRINT_NULL As String = "null"
Public Const JSON_ERROR As String = "errorMessages"

'--------------------------------------------------------------------
' Status für eine Aufgabe aus Jira aktualisieren
' Sucht nach dem ersten Treffer eines Jira-Links und zieht sich per JSON-API den Assignee und den Status
' Anschließend werden die Informationen in folgende Aufgabenfelder geschrieben:
' - Assignee => Kontakte
' - Status   => Abrechnungsinformationen
'--------------------------------------------------------------------
Public Function UpdateTaskFromJira(task As taskItem, jiraURL As String) As Boolean
    Dim startIndex As Long
    Dim endIndex As Long
    Dim url As String
    Dim json As String
    Dim errorMessage As String
    Dim assignee As String
    Dim status As String
    Dim sprint As String
    Dim request As MSXML2.XMLHTTP60
    
    ' Add "/" at end if not present
    If Not EndsWith(jiraURL, "/") Then
        jiraURL = jiraURL & "/"
    End If
    ' Remove "browse/" from end if present (otherwise appending API sub-url is not possible)
    If EndsWith(jiraURL, "browse/") Then
        jiraURL = Left(jiraURL, Len(jiraURL) - Len("browse/"))
    End If
    
    url = ""
    startIndex = 1
    ' check link occurrence
    If InStr(task.Body, jiraURL & "browse/") <> 0 Then
        Do
            startIndex = InStr(startIndex, task.Body, jiraURL & "browse/")
            If startIndex = 0 Then
                Exit Do
            End If
            ' extract ID & build API url (only look at first URL match)
            ' +7 required for "browse/" in URL
            startIndex = startIndex + Len(jiraURL) + 7
            url = Right(task.Body, Len(task.Body) - startIndex + 1)
            endIndex = FirstInStr(url, "> ,;?" & vbNewLine & vbCr)
            If endIndex = 0 Then
                endIndex = Len(task.Body) + 1
            End If
            url = Left(url, endIndex - 1)
            
            ' check if valid link (including "-" for an ID and not just project)
        Loop While InStr(url, "-") = 0
    End If
        
    If url <> "" Then
        url = jiraURL & JIRA_API & url
        Debug.Print url
        ' get status from Jira via HTTP
        Set request = New MSXML2.XMLHTTP60
        Call request.Open("GET", url, False)
        Call request.Send
        json = request.responseText
        ' check json
        If StartsWith(json, "{""" & JSON_ERROR) Then
            ' find error
            startIndex = InStr(json, """" & JSON_ERROR & """:")
            errorMessage = Right(json, Len(json) - startIndex - Len(JSON_ERROR) - 3)
            endIndex = InStr(errorMessage, "]")
            errorMessage = Left(errorMessage, endIndex - 1)
            Debug.Print "> ERROR: " & errorMessage
            UpdateTaskFromJira = False
        ElseIf (StartsWith(json, "{") And EndsWith(json, "}")) Then
            ' find assignee
            startIndex = InStr(json, """" & JSON_ASSIGNEE & """:")
            startIndex = InStr(startIndex, json, """" & JSON_ASSIGNEE_DISP & """:")
            assignee = Right(json, Len(json) - startIndex - Len(JSON_ASSIGNEE_DISP) - 3)
            endIndex = InStr(assignee, """")
            assignee = Left(assignee, endIndex - 1)
            Debug.Print "> Assignee = " & assignee
            ' find status (von hinten, da ggf. vorher Subtask & Link-Stati kommen)
            startIndex = InStrRev(json, """" & JSON_STATUS & """:")
            startIndex = InStr(startIndex, json, """" & JSON_STATUS_DISP & """:")
            status = Right(json, Len(json) - startIndex - Len(JSON_STATUS_DISP) - 3)
            endIndex = InStr(status, """")
            status = Left(status, endIndex - 1)
            Debug.Print "> Status   = " & status
            ' find sprint (ACHTUNG: custom_field, daher format ,name=ISX, statt "name":"ISX",)
            startIndex = InStr(json, """" & JSON_SPRINT & """:")
            endIndex = InStr(startIndex, json, "" & JSON_SPRINT_NULL & ",")
            startIndex = InStr(startIndex, json, "," & JSON_SPRINT_DISP & "=")
            Debug.Print url & " s=" & startIndex & " e=" & endIndex
            If (startIndex <> 0) Then
                sprint = Right(json, Len(json) - startIndex - Len(JSON_SPRINT_DISP) - 1)
                endIndex = InStr(sprint, ",")
                sprint = Left(sprint, endIndex - 1)
            Else
                sprint = ""
            End If
            Debug.Print "> Sprint   = " & sprint
            ' set fields
            task.BillingInformation = "[" & status & "]"
            task.ContactNames = "@ " & assignee
            task.Mileage = sprint
            task.Save
            UpdateTaskFromJira = True
        Else
            Debug.Print "> ERROR: Could not get JSON"
            UpdateTaskFromJira = False
        End If
    Else
        UpdateTaskFromJira = False
    End If
End Function

