Attribute VB_Name = "JiraUtil"
'--------------------------------------------------
' Definition der benötigen Konstanten und Konfiguration
'--------------------------------------------------
' JIRA API Sub-URL
Public Const JIRA_API As String = "rest/api/latest/issue/"
' JSON identifier
' Field 1
Public Const JSON_ASSIGNEE As String = "assignee"
Public Const JSON_ASSIGNEE_DISP As String = "displayName"
' Field 2
Public Const JSON_STATUS As String = "status"
Public Const JSON_STATUS_DISP As String = "name"
' Field 4
Public Const JSON_SPRINT As String = "customfield_10461"
Public Const JSON_SPRINT_DISP As String = "name"
Public Const JSON_SPRINT_NULL As String = "null"
' error message
Public Const JSON_ERROR As String = "errorMessages"
' Authentifizierungsvarianten
Private Const AUTH_AUTO As String = "auto"
Private Const AUTH_PW As String = "password"
Private Const AUTH_CERT As String = "cert"
' CERT Authentification Sub-URL & Parameter
Public Const CERT_URL As String = "login.form"
Public Const CERT_PARAM As String = "login-form-type=cert"

'--------------------------------------------------
' Custom Type für die Authentifizierung
'--------------------------------------------------
Private Type JiraConfiguration
     url As String          ' gibt die URL zur Jira-Instanz an
     fields As Integer      ' gibt an, welche conf.fields aus dem Ticket ausgelesen werden sollen (Bit-Maske)
     authType As String     ' gibt den Authentifizierungs-Typ an
     authInfo As String     ' Zwischenspeicher für Authentifizierungs-Informationen - wird automatisch befüllt
End Type

'--------------------------------------------------
' Aufgaben-Status aus Jira aktualisieren
' Sucht nach dem ersten Treffer eines Jira-Links und zieht sich per JSON-API den Assignee und den Status
' Anschließend werden die Informationen in folgende Aufgabenfelder geschrieben:
' - Assignee => Kontakte
' - Status   => Abrechnungsinformationen
' - Sprint   => Reisekilometer
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
    
    ' Prepare configurations
    Dim conf_1 As JiraConfiguration
    With conf_1
        .url = "https://server1.com/sjira/"
        .fields = 1 Or 2 Or 4
        .authType = AUTH_AUTO
    End With
    
    Dim conf_2 As JiraConfiguration
    With conf_2
        .url = "https://server2.com/sjira/"
        .fields = 1 Or 2
        .authType = AUTH_CERT
    End With
    
    Dim conf_3 As JiraConfiguration
    With conf_3
        .url = "https://server3.com/sjira"
        .fields = 1 Or 2
        .authType = AUTH_PW
    End With
              
    'For Each task In taskList
    For Each task In taskList
        handled = False
        If task.status <> olTaskComplete Or handleAll Then
        
            ' je zu prüfendem Jira einen If-Update-Block (nach Prio sortiert)
            If Not handled Then
                handled = UpdateTaskFromJira(task, conf_1)
            End If
            If Not handled Then
                handled = UpdateTaskFromJira(task, conf_2)
            End If
            If Not handled Then
                handled = UpdateTaskFromJira(task, conf_3)
            End If
        
            i = i + 1
            Call ProgressBar.Update(i, total)
        End If
    Next
    ProgressBar.Hide
    
    Debug.Print auth3
End Sub

'--------------------------------------------------------------------
' Status für eine Aufgabe aus Jira aktualisieren
' Sucht nach dem ersten Treffer eines Jira-Links und zieht sich per JSON-API den Assignee und den Status
' Anschließend werden die Informationen in folgende Aufgabenfelder geschrieben:
' - Assignee => Kontakte
' - Status   => Abrechnungsinformationen
' - Sprint   => Reisekilometer
'--------------------------------------------------------------------
Public Function UpdateTaskFromJira(task As taskItem, ByRef conf As JiraConfiguration) As Boolean
    Dim startIndex As Long
    Dim endIndex As Long
    Dim url As String
    Dim json As String
    Dim auth64 As String
    Dim errorMessage As String
    Dim assignee As String
    Dim status As String
    Dim sprint As String
    Dim request As MSXML2.XMLHTTP60
    
    ' Add "/" at end if not present
    If Not EndsWith(conf.url, "/") Then
        conf.url = conf.url & "/"
    End If
    ' Remove "browse/" from end if present (otherwise appending API sub-url is not possible)
    If EndsWith(conf.url, "browse/") Then
        conf.url = Left(conf.url, Len(conf.url) - Len("browse/"))
    End If
    
    url = ""
    startIndex = 1
    ' check link occurrence
    If InStr(task.Body, conf.url & "browse/") <> 0 Then
        Do
            startIndex = InStr(startIndex, task.Body, conf.url & "browse/")
            If startIndex = 0 Then
                Exit Do
            End If
            ' extract ID & build API url (only look at first URL match)
            ' +7 required for "browse/" in URL
            startIndex = startIndex + Len(conf.url) + 7
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
        url = conf.url & JIRA_API & url
        Debug.Print url
        
        ' prepare to get status from Jira via HTTP
        Set request = New MSXML2.XMLHTTP60
        Call request.Open("GET", url, False)
        
        ' Authenticate
        If (conf.authType = AUTH_AUTO) Then
            ' do nothing
        ElseIf (conf.authType = AUTH_PW) Then
            If (conf.authInfo = "") Then
                LoginForm.UrlLabel = conf.url
                LoginForm.Show (vbModal)
                If (LoginForm.okAction) Then
                    conf.authInfo = base64.Base64EncodeString(LoginForm.username & ":" & LoginForm.password)
                End If
                LoginForm.Reset (False)
            End If
            If (conf.authInfo <> "") Then
                Call request.setRequestHeader("Authorization", "Basic " & conf.authInfo)
            End If
        ElseIf (conf.authType = AUTH_CERT) Then
            Dim authUrl As String
            authUrl = Left(conf.url, InStr(Len("https://") + 1, conf.url, "/"))
            authUrl = authUrl & CERT_URL
            Dim loginRequest As MSXML2.XMLHTTP60
            Set loginRequest = New MSXML2.XMLHTTP60
            Call loginRequest.Open("POST", authUrl, False)
            Call loginRequest.Send(CERT_PARAM)
            result = loginRequest.responseText
        End If
        
        ' get status from Jira via HTTP
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
            If ((conf.fields And 1) <> 0) Then
                ' find assignee
                startIndex = InStr(json, """" & JSON_ASSIGNEE & """:")
                startIndex = InStr(startIndex, json, """" & JSON_ASSIGNEE_DISP & """:")
                assignee = Right(json, Len(json) - startIndex - Len(JSON_ASSIGNEE_DISP) - 3)
                endIndex = InStr(assignee, """")
                assignee = Left(assignee, endIndex - 1)
                Debug.Print "> Assignee = " & assignee
            End If
            If ((conf.fields And 2) <> 0) Then
                ' find status (von hinten, da ggf. vorher Subtask & Link-Stati kommen)
                startIndex = InStrRev(json, """" & JSON_STATUS & """:")
                startIndex = InStr(startIndex, json, """" & JSON_STATUS_DISP & """:")
                status = Right(json, Len(json) - startIndex - Len(JSON_STATUS_DISP) - 3)
                endIndex = InStr(status, """")
                status = Left(status, endIndex - 1)
                Debug.Print "> Status   = " & status
            End If
            If ((conf.fields And 4) <> 0) Then
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
            End If
            ' set conf.fields
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

