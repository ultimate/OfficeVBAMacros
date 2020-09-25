Attribute VB_Name = "AttachmentUtil"
'--------------------------------------------------
' FUNKTIONEN ZUM ARCHIVIEREN VON MAILS
'--------------------------------------------------

'--------------------------------------------------
' Custom Type für die Bearbeitung der Mails
'--------------------------------------------------
Private Type AttachmentUpdate
     text As String
     position As Integer
     attachmentItem As attachment
End Type

'--------------------------------------------------
' Verwende diese Funktion, wenn eine Mail im separaten Fenster angezeigt wird
'--------------------------------------------------
Public Sub ArchiveAttachmentsForCurrentMail()
    Dim mail As MailItem
    Dim archivedAttachments As Integer
    Dim del As Boolean
    Dim opt As Integer
    
    If (AttachmentConfig.SHOW_CONFIRM) Then
        opt = MsgBox(AttachmentConfig.MSG_CONFIRM, vbYesNoCancel)
        If (opt = vbYes) Then
            del = True
        ElseIf (opt = vbNo) Then
            del = False
        Else
            Return
        End If
    Else
        del = AttachmentConfig.DELETE_ATTACHMENTS
    End If
    
    Set mail = Application.ActiveInspector.CurrentItem
    archivedAttachments = ArchiveAttachments(mail, del)
    
    If (AttachmentConfig.SHOW_SUMMARY) Then
        If (archivedAttachments = 0) Then
            MsgBox AttachmentConfig.MSG_ARCHIVED_0
        ElseIf (archivedAttachments = 1) Then
            MsgBox archivedAttachments & AttachmentConfig.MSG_ARCHIVED_1
        Else
            MsgBox archivedAttachments & AttachmentConfig.MSG_ARCHIVED_N
        End If
    End If
End Sub

'--------------------------------------------------
' Verwende diese Funktion, wenn eine oder mehrere Mails in einem Ordner selektiert sind
' und/oder im Lesebereich angezeigt werden
'--------------------------------------------------
Public Sub ArchiveAttachmentsForSelectedMails()
    Dim mail As MailItem
    Dim archivedAttachments As Integer
    Dim totalArchivedAttachments As Integer
    Dim archivedMails As Integer
    Dim selectedMails As Integer
    Dim del As Boolean
    Dim opt As Integer
    
    If (AttachmentConfig.SHOW_CONFIRM) Then
        opt = MsgBox(AttachmentConfig.MSG_CONFIRM, vbYesNoCancel)
        If (opt = vbYes) Then
            del = True
        ElseIf (opt = vbNo) Then
            del = False
        Else
            Return
        End If
    Else
        del = AttachmentConfig.DELETE_ATTACHMENTS
    End If
    
    totalArchivedAttachments = 0
    selectedMails = 0
    archivedMails = 0
    
    For Each mail In Application.ActiveExplorer.Selection
        archivedAttachments = ArchiveAttachments(mail, del)
        totalArchivedAttachments = totalArchivedAttachments + archivedAttachments
        selectedMails = selectedMails + 1
        If (archivedAttachments > 0) Then
            archivedMails = archivedMails + 1
        End If
    Next mail
    
    If (AttachmentConfig.SHOW_SUMMARY) Then
        If (selectedMails = 0) Then
            MsgBox AttachmentConfig.MSG_MAILS_0
        ElseIf (selectedMails = 1) Then
            If (totalArchivedAttachments = 0) Then
                MsgBox AttachmentConfig.MSG_ARCHIVED_0
            ElseIf (totalArchivedAttachments = 1) Then
                MsgBox totalArchivedAttachments & AttachmentConfig.MSG_ARCHIVED_1
            Else
                MsgBox totalArchivedAttachments & AttachmentConfig.MSG_ARCHIVED_N
            End If
        Else
            If (totalArchivedAttachments = 0) Then
                MsgBox archivedMails & AttachmentConfig.MSG_MAILS_OF & selectedMails & AttachmentConfig.MSG_MAILS_N & vbCrLf & AttachmentConfig.MSG_ARCHIVED_0
            ElseIf (totalArchivedAttachments = 1) Then
                MsgBox archivedMails & AttachmentConfig.MSG_MAILS_OF & selectedMails & AttachmentConfig.MSG_MAILS_N & vbCrLf & totalArchivedAttachments & AttachmentConfig.MSG_ARCHIVED_1
            Else
                MsgBox archivedMails & AttachmentConfig.MSG_MAILS_OF & selectedMails & AttachmentConfig.MSG_MAILS_N & vbCrLf & totalArchivedAttachments & AttachmentConfig.MSG_ARCHIVED_N
            End If
        End If
    End If
End Sub

'--------------------------------------------------
' Archiviere und Entferne die Anhänge zur der übergebenen Mail
'--------------------------------------------------
Public Function ArchiveAttachments(mail As MailItem, del As Boolean) As Integer
    
    Dim fileName As String
    Dim fileFolder As String
    Dim fileNamePattern As String
    Dim address As String
    Dim name As String
    Dim rec As Integer
        
    fileNamePattern = AttachmentConfig.FILENAME_PATTERN
    
    ' generelle Platzhalter ersetzen
    fileNamePattern = Replace(fileNamePattern, "%DATETIME", Format(mail.ReceivedTime, AttachmentConfig.DATE_FORMAT))
    fileNamePattern = Replace(fileNamePattern, ":", ".")
    If (AttachmentConfig.REPLACE_SPACES) Then
        fileName = Replace(fileName, " ", "_")
    End If
    ' Mystischer Check, ob Email von einem selbst gesendet wurde
    If (mail.ReceivedByName = "") Then
        ' Message wurde von einem selbst gesendet
        fileNamePattern = Replace(fileNamePattern, "%DIRECTION", AttachmentConfig.DIRECTION_TO)
        For rec = 1 To mail.recipients.Count
            ' nur "An" Empfaenger (Type = 1) betrachten
            If (mail.recipients.Item(rec).Type = 1) Then
                address = GetAddress(mail.recipients.Item(rec).addressEntry)
                name = GetName(mail.recipients.Item(rec).addressEntry)
                ' Debug.Print name & ":" & address
                fileNamePattern = Replace(fileNamePattern, "%CONTACTMAIL", address & ",%CONTACTMAIL")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTNAME", mail.recipients.Item(rec).name & ",%CONTACTNAME")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTSYMBOL", GetKuerzel(address) & ",%CONTACTSYMBOL")
            End If
        Next
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTMAIL", "")
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTNAME", "")
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTSYMBOL", "")
    Else
        address = GetAddress(mail.Sender)
        name = GetName(mail.Sender)
        ' Debug.Print name & ":" & address
        fileNamePattern = Replace(fileNamePattern, "%DIRECTION", AttachmentConfig.DIRECTION_FROM)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTMAIL", address)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTNAME", mail.SenderName)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTSYMBOL", GetKuerzel(address))
    End If
    
    Debug.Print "Archiviere Anhänge im Pfad -> " & archiveFolder & "\" & fileNamePattern
    Debug.Print "  Anhänge werden entfernt? " & del
    
    Dim att As attachment
    Dim archivedAttachments As Integer
    archivedAttachments = 0
    
    Dim attachmentUpdates() As AttachmentUpdate
    ReDim attachmentUpdates(mail.Attachments.Count)
    
    Dim offset As Integer
    offset = 0
        
    If (mail.Attachments.Count > 0) Then
        For Each att In mail.Attachments
            If (att.Size >= AttachmentConfig.MIN_FILE_SIZE) Then
                fileName = fileNamePattern
                fileName = Replace(fileName, "%FILENAME", att.fileName)
                If (AttachmentConfig.REPLACE_SPACES) Then
                    fileName = Replace(fileName, " ", "_")
                End If
                fileName = AttachmentConfig.ARCHIVE_FOLDER & "\" & fileName
                
                Debug.Print "  Archiviere Anhang " & (archivedAttachments + 1) & ": " & att.fileName & " (Größe=" & att.Size & ") -> " & fileName
                
                fileFolder = Left(fileName, InStrRev(fileName, "\") - 1)
                If Dir(fileFolder, vbDirectory) = "" Then
                    ' Ordner existiert nicht und muss erstellt werden
                    MkDir (fileFolder)
                End If
                att.SaveAsFile (fileName)
                            
                If mail.BodyFormat = olFormatHTML Then
                    attachmentUpdates(archivedAttachments).text = Replace(AttachmentConfig.MSG_IN_MAIL_TEXT, "%I", archivedAttachments + 1) & "<a href=""" & fileName & """>" & fileName & "</a><br/>"
                    attachmentUpdates(archivedAttachments).position = 0
                ElseIf mail.BodyFormat = olFormatRichText Then
                    attachmentUpdates(archivedAttachments).text = Replace(AttachmentConfig.MSG_IN_MAIL_TEXT, "%I", archivedAttachments + 1) & """file://" & fileName & """"
                    attachmentUpdates(archivedAttachments).position = att.position + offset - 1
                    ' Ich weiss nicht genau warum, aber da ist ein offset in der
                    ' Position drin, der sich von Anhang zu Anhang aufsummiert
                    ' Und dann noch den neu eingefügten Link berücksichtigen
                    offset = offset - 31 + Len(attachmentUpdates(archivedAttachments).text)
                Else
                    attachmentUpdates(archivedAttachments).text = Replace(AttachmentConfig.MSG_IN_MAIL_TEXT, "%I", archivedAttachments + 1) & """file://" & fileName & """"
                    attachmentUpdates(archivedAttachments).position = 0
                End If
                
                Set attachmentUpdates(archivedAttachments).attachmentItem = att
                
                archivedAttachments = archivedAttachments + 1
            Else
                Debug.Print "  Überspringe Anhang " & (archivedAttachments + 1) & ": " & att.fileName & " (Größe=" & att.Size & ")"
                If (del) Then
                    If mail.BodyFormat = olFormatRichText Then
                        ' Anhang-Position korrigieren (da ggf. zuvor Anhänge entfernt
                        ' und Text eingefügt wurde, verschiebt sich die Position)
                        Debug.Print "position old=" & att.position & " offset=" & offset
                        att.position = att.position + offset
                        Debug.Print "position new=" & att.position
                    End If
                End If
            End If
        Next att
        
        If (del) Then
            Dim i As Integer
            Dim msgUpdate As String
            
            ' Löschen der Attachments darf erst ganz am Schluss erfolgen
            ' Sonst zerhaut man sich die Schleife
            For i = archivedAttachments - 1 To 0 Step -1
                attachmentUpdates(i).attachmentItem.Delete
            Next i
            
            mail.Save
            
            If mail.BodyFormat = olFormatHTML Then
                Debug.Print "  Aktualisiere HTML-Inhalt"
                msgUpdate = ""
                For i = 0 To archivedAttachments - 1
                    msgUpdate = msgUpdate & attachmentUpdates(i).text & vbCrLf
                Next i
                mail.HTMLBody = "<p><i>" & msgUpdate & "</i></p>" & mail.HTMLBody
            ElseIf mail.BodyFormat = olFormatRichText Then
                Debug.Print "  Aktualisiere RTF-Inhalt"
                'RTF-Word-Editor
                Dim mailInspector As Outlook.Inspector
                Dim mailEditor As Word.Document
                Dim mailProtection As WdProtectionType
                
                Set mailInspector = mail.GetInspector
                Set mailEditor = mailInspector.WordEditor
                
                mailProtection = mailEditor.ProtectionType
                If (mailProtection <> wdNoProtection) Then
                    'Debug.Print "  Mail ist gegen Bearbeitung geschützt: Hebe Schutz auf"
                    mailEditor.UnProtect
                End If
                                
                For i = 0 To archivedAttachments - 1
                    'Debug.Print "    inserting at " & attachmentUpdates(i).position & ": " & attachmentUpdates(i).text
                    mailEditor.Characters(attachmentUpdates(i).position).InsertAfter (attachmentUpdates(i).text)
                    ' Text kursiv machen
                    mailEditor.Range(attachmentUpdates(i).position, attachmentUpdates(i).position + Len(attachmentUpdates(i).text)).Italic = True
                Next i
            Else
                Debug.Print "  Aktualisiere Plain-Text-Inhalt"
                msgUpdate = ""
                For i = 0 To archivedAttachments - 1
                    msgUpdate = msgUpdate & attachmentUpdates(i).text & vbCrLf
                Next i
                mail.Body = msgUpdate & vbCrLf & mail.Body
            End If
        
            Debug.Print "  Speichere aktualisierte Mail"
            mail.Save
        End If
        
    End If
        
    ArchiveAttachments = archivedAttachments
End Function

'--------------------------------------------------
' Lese das Kürzel zu einer Mail-Adresse aus der Kürzel-Datei
'--------------------------------------------------
Public Function GetKuerzel(mailAdress As String) As String
            
    GetKuerzel = mailAdress
        
    Dim kuerzelTxt As String
    ''' Datei lesen (mit utf-8!)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (AttachmentConfig.KUERZEL_FILE)
    kuerzelTxt = objStream.ReadText()
    objStream.Close
     
    Dim index As Integer
    Dim line As String
    
    index = 0
    Do While Len(kuerzelTxt) > 0
        line = Left(kuerzelTxt, InStr(kuerzelTxt, vbCrLf) - 1)
        ' Debug.Print "L: " & line & " :L"
        If (StringUtil.StartsWith(line, "#")) Then
            ' Kommentar-Zeile
        ElseIf (StringUtil.EndsWith(line, mailAdress)) Then
            GetKuerzel = Left(line, InStr(line, "=") - 1)
        End If
        kuerzelTxt = Right(kuerzelTxt, Len(kuerzelTxt) - InStr(kuerzelTxt, vbCrLf) - 1)
    Loop
End Function

'--------------------------------------------------
' Extrahiere die Adresse aus einem Adress-Eintrag
' Wird benötigt, da das Vorgehen für Exchange-Nutzer anders ist
' als für externe Kontakte
'--------------------------------------------------
Public Function GetAddress(addressEntry As addressEntry) As String
    If (addressEntry.GetExchangeUser Is Nothing) Then
        GetAddress = addressEntry.address
    Else
        GetAddress = addressEntry.GetExchangeUser.PrimarySmtpAddress
    End If
End Function

'--------------------------------------------------
' Extrahiere den Namen aus einem Adress-Eintrag
' Wird benötigt, da das Vorgehen für Exchange-Nutzer anders ist
' als für externe Kontakte
'--------------------------------------------------
Public Function GetName(addressEntry As addressEntry) As String
    If (addressEntry.GetExchangeUser Is Nothing) Then
        GetName = addressEntry.name
    Else
        GetName = addressEntry.GetExchangeUser.name
    End If
End Function
