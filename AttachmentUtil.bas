Attribute VB_Name = "AttachmentUtil"
'--------------------------------------------------
' FUNKTIONEN ZUM ARCHIVIEREN VON MAILS
'--------------------------------------------------

' Choose the preferred language
#Const LANGUAGE = "DE"

'--------------------------------------------------
' Interne Konstanten
'--------------------------------------------------
Private Const OPTION_WITH_DELETE As Integer = 1
Private Const OPTION_WITHOUT_DELETE As Integer = 0
Private Const OPTION_CANCEL As Integer = -1
' Anzeige-Messages
#If LANGUAGE = "DE" Then
    Private Const MSG_TITLE As String = "AttachmentUtil V1.0 by J.Verkin - Nutzung auf eigene Gefahr"
    Private Const MSG_CONFIRM As String = "Anhänge archivieren und entfernen?" & vbCrLf _
                                & "  Ja = Archivieren UND Entfernen" & vbCrLf _
                                & "  Nein = Archivieren OHNE Entfernen" & vbCrLf _
                                & "  Abbrechen = Keine Aktion durchführen" & vbCrLf _
                                & vbCrLf _
                                & "Achtung: Nutzung auf eigene Gefahr!"
    Private Const MSG_OVERWRITE As String = "Datei bereits vorhanden! Überschreiben?" & vbCrLf _
                                & "  %FILENAME"
    Private Const MSG_NO_KUERZEL As String = "Kein Kontakt-Kürzel gefunden! Fortfahren?" & vbCrLf _
                                & "  %FILENAME"
    Private Const MSG_ARCHIVED_0 As String = "Keine Anhänge zum archivieren vorhanden"
    Private Const MSG_ARCHIVED_1 As String = " Anhang erfolgreich archiviert"
    Private Const MSG_ARCHIVED_N As String = " Anhänge erfolgreich archiviert"
    Private Const MSG_MAILS_0 As String = "Keine Mails ausgewählt"
    Private Const MSG_MAILS_OF As String = " von "
    Private Const MSG_MAILS_N As String = " ausgewählten Mails enthielten Anhänge"
    Private Const MSG_SIZE As String = "Postfach-Speicher freigegeben = ca. "
    Private Const MSG_ATT_IN_MAIL_TEXT As String = "Anhang %I entfernt und archivert unter: "
    Private Const MSG_IMG_IN_MAIL_TEXT As String = "Bild %I entfernt und archivert unter: "
#Else
    Private Const MSG_TITLE As String = "AttachmentUtil V1.0 by J.Verkin - Use at your own risk"
    Private Const MSG_CONFIRM As String = "Archive and remove attachments?" & vbCrLf _
                                & "  Yes = archive AND remove" & vbCrLf _
                                & "  No = archive WITHOUT remove" & vbCrLf _
                                & "  Cancel = do nothing" & vbCrLf _
                                & vbCrLf _
                                & "Attention: Use at your own risk!"
    Private Const MSG_OVERWRITE As String = "File exists! Overwrite?" & vbCrLf _
                                & "  %FILENAME"
    Private Const MSG_NO_KUERZEL As String = "No contact symbol found! Continue?" & vbCrLf _
                                & "  %FILENAME"
    Private Const MSG_ARCHIVED_0 As String = "No attachments available for archiving"
    Private Const MSG_ARCHIVED_1 As String = " attachment successfully archived"
    Private Const MSG_ARCHIVED_N As String = " attachments successfully archived"
    Private Const MSG_MAILS_0 As String = "No mail selected"
    Private Const MSG_MAILS_OF As String = " of "
    Private Const MSG_MAILS_N As String = " selected mails contained attachments"
    Private Const MSG_SIZE As String = "Postbox-storage freed = approx. "
    Private Const MSG_ATT_IN_MAIL_TEXT As String = "Attachment %I removed and archived at: "
    Private Const MSG_IMG_IN_MAIL_TEXT As String = "Image %I removed and archived at: "
#End If
' Factor für die Umrechnung von HIMETRIC in Pixel
Private Const FACTOR_HIMETRIC As Double = 26.45833

'--------------------------------------------------
' Custom Type für die Bearbeitung der Mails
'--------------------------------------------------
Private Type AttachmentUpdate
     filename As String
     position As Integer
     attachmentName As String
     attachmentItem As attachment
     shapeItem As Word.InlineShape
End Type

'--------------------------------------------------
' Verwende diese Funktion, wenn eine Mail im separaten Fenster angezeigt wird
'--------------------------------------------------
Public Sub ArchiveAttachmentsForCurrentMail()
    Dim mail As MailItem
    Dim archivedAttachments As Integer
    Dim sizeBefore As Long
    Dim sizeAfter As Long
    Dim opt As Integer
    
    opt = ShowConfirm
    If (opt = OPTION_CANCEL) Then Exit Sub
    
    Set mail = Application.ActiveInspector.CurrentItem
    sizeBefore = mail.Size
    archivedAttachments = ArchiveAttachments(mail, opt = OPTION_WITH_DELETE)
    sizeAfter = mail.Size
    
    Call ShowSummary(1, 1, archivedAttachments, sizeBefore - sizeAfter)
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
    Dim sizeBefore As Long
    Dim sizeAfter As Long
    Dim opt As Integer
    
    opt = ShowConfirm
    If (opt = OPTION_CANCEL) Then Exit Sub
    
    totalArchivedAttachments = 0
    selectedMails = 0
    archivedMails = 0
    sizeBefore = 0
    sizeAfter = 0
    
    For Each mail In Application.ActiveExplorer.Selection
        sizeBefore = sizeBefore + mail.Size
        archivedAttachments = ArchiveAttachments(mail, opt = OPTION_WITH_DELETE)
        totalArchivedAttachments = totalArchivedAttachments + archivedAttachments
        selectedMails = selectedMails + 1
        If (archivedAttachments > 0) Then
            archivedMails = archivedMails + 1
        End If
        sizeAfter = sizeAfter + mail.Size
    Next mail
    
    Call ShowSummary(selectedMails, archivedMails, totalArchivedAttachments, sizeBefore - sizeAfter)
End Sub

'--------------------------------------------------
' Archiviere und Entferne die Anhänge zur der übergebenen Mail
'--------------------------------------------------
Public Function ArchiveAttachments(mail As MailItem, del As Boolean) As Integer
    
    ' Ein paar Variablen
    Dim fileNamePattern As String
    Dim archivedAttachments As Integer: archivedAttachments = 0
    Dim archivedOLEs As Integer: archivedOLEs = 0
    Dim continueWithoutKuerzel As Boolean
            
    fileNamePattern = GetFileNamePattern(mail)
    
    ' Prüfe, ob Dateiname ein @ enthält = kein Kürzel gefunden
    If (InStr(fileNamePattern, "@") <> 0) Then
        continueWithoutKuerzel = ShowNoKuerzel(fileNamePattern)
    Else
        continueWithoutKuerzel = True
    End If
    
    If (continueWithoutKuerzel) Then
        Debug.Print "Archiviere Anhänge im Pfad -> " & ARCHIVE_FOLDER & "\" & fileNamePattern
        Debug.Print "  Anhänge werden entfernt? " & del
            
        If (mail.Attachments.count > 0 And mail.Size > AttachmentConfig.MIN_MAIL_SIZE) Then
            ' (normale) Anhänge behandeln
            archivedAttachments = HandleAttachments(mail, del, fileNamePattern)
            
            If (mail.BodyFormat = olFormatRichText) Then
                ' OLE Bilder behandeln
                archivedOLEs = HandleOLEImages(mail, del, fileNamePattern)
            End If
        End If
    Else
        Debug.Print "Archivieren abgebrochen, da kein Kürzel gefunden -> " & ARCHIVE_FOLDER & "\" & fileNamePattern
    End If
        
    ArchiveAttachments = archivedAttachments + archivedOLEs
End Function

'--------------------------------------------------
' Normale Anhaenge behandeln
'--------------------------------------------------
Private Function HandleAttachments(mail As MailItem, del As Boolean, fileNamePattern As String) As Integer
    If (mail.Attachments.count > 0) Then
    
        ' Ein paar Variablen
        Dim filename As String
        Dim fileFolder As String
        Dim overwrite As Boolean
        Dim att As attachment
        Dim position As Integer
        Dim offset As Integer: offset = 0
        Dim counter As Integer
        Dim i As Integer
        Dim text As String
        Dim imgToSmall As Boolean
            
        ' RTF-Word-Editor
        Dim mailInspector As Outlook.Inspector
        Dim mailEditor As Word.Document
        Dim mailProtection As WdProtectionType
        
        ' RTF-Variablen nur bei RFT-Mail initialisieren
        If (mail.BodyFormat = olFormatRichText) Then
            Set mailInspector = mail.GetInspector
            Set mailEditor = mailInspector.WordEditor
        End If
        
        Dim archivedAttachments As Integer: archivedAttachments = 0
        Dim attachmentUpdates() As AttachmentUpdate: ReDim attachmentUpdates(mail.Attachments.count)
        
        counter = 1
        For Each att In mail.Attachments
            If (att.Type = olOLE) Then
                ' OLE-Bilder müssen separat behandelt werden!
                ' Speichern der Bilder ist nur über RTF-Word-Editor möglich
                ' Normales Speichern resultiert in nicht lesbarem Bitmap
            Else
                ' Dateiname aus Pattern ermitteln
                filename = fileNamePattern
                filename = Replace(filename, "%FILENAME", att.filename)
                If (AttachmentConfig.REPLACE_SPACES) Then
                    filename = Replace(filename, " ", "_")
                End If
                filename = AttachmentConfig.ARCHIVE_FOLDER & "\" & filename
                
                ' Prüfen, ob es ein HTML-Bild ist
                imgToSmall = False
                If (mail.BodyFormat = olFormatHTML) Then
                    ' Suche nach HTML TAG für eingebettetes Bild in der Form
                    ' <img width=85 height=76 id="Bild_x0020_1" src="cid:image001.png@01D29976.A64E5DB0">
                    If (InStr(mail.HTMLBody, "src=""cid:" & att.filename & "@") <> 0) Then
                        If (att.Size < AttachmentConfig.MIN_IMAGE_SIZE) Then
                            imgToSmall = True
                        End If
                    End If
                End If
            
                If (att.Size >= AttachmentConfig.MIN_FILE_SIZE And Not imgToSmall) Then
                    
                    fileFolder = Left(filename, InStrRev(filename, "\") - 1)
                    If Dir(fileFolder, vbDirectory) = "" Then
                        ' Ordner existiert nicht und muss erstellt werden
                        MkDir (fileFolder)
                        overwrite = True
                    ElseIf Dir(filename) <> "" Then
                        ' Datei existiert bereits
                        overwrite = ShowOverwrite(filename)
                    Else
                        overwrite = True
                    End If
                    
                    If (overwrite) Then
                        Debug.Print "  Archiviere Anhang: " & att.filename & " (Größe=" & att.Size & ") -> " & filename
                    
                        att.SaveAsFile (filename)
                        
                        attachmentUpdates(archivedAttachments).filename = filename
                        attachmentUpdates(archivedAttachments).attachmentName = att.filename
                        attachmentUpdates(archivedAttachments).position = att.position
                        Set attachmentUpdates(archivedAttachments).attachmentItem = att
                        
                        archivedAttachments = archivedAttachments + 1
                    Else
                        Debug.Print "  Überspringe Anhang: " & att.filename & " (Größe=" & att.Size & ") -> DATEI EXISTIERT BEREITS"
                    End If
                Else
                    Debug.Print "  Überspringe Anhang: " & att.filename & " (Größe=" & att.Size & ")"
                End If
            End If
            counter = counter + 1
        Next att
        
        If (del And archivedAttachments > 0) Then
            Dim msgUpdate As String
            Dim startIndex As Long
            Dim endIndex As Long
            Dim htmlTag As String
            Dim maxLength As Long
            
            ' Löschen der Attachments darf erst ganz am Schluss erfolgen
            ' Sonst zerhaut man sich die Schleife
            For i = archivedAttachments - 1 To 0 Step -1
                attachmentUpdates(i).attachmentItem.Delete
            Next i
            
            If mail.BodyFormat = olFormatHTML Then
                Debug.Print "  Aktualisiere HTML-Inhalt"
                msgUpdate = ""
                Dim countImg As Integer: countImg = 1
                Dim countAtt As Integer: countAtt = 1
                For i = 0 To archivedAttachments - 1
                    ' Suche nach HTML TAG für eingebettetes Bild in der Form
                    ' <img width=85 height=76 id="Bild_x0020_1" src="cid:image001.png@01D29976.A64E5DB0">
                    startIndex = InStr(mail.HTMLBody, "src=""cid:" & attachmentUpdates(i).attachmentName & "@")
                    If (startIndex <> 0) Then
                        ' HTML-Tag gefunden
                        text = Replace(MSG_IMG_IN_MAIL_TEXT, "%I", countImg) & "<a href=""" & attachmentUpdates(i).filename & """>" & attachmentUpdates(i).filename & "</a><br/>"
                        ' --> Link an entsprechender Stelle einfügen
                        startIndex = InStrRev(mail.HTMLBody, "<img", startIndex)
                        endIndex = InStr(startIndex, mail.HTMLBody, """>")
                        htmlTag = Left(Right(mail.HTMLBody, Len(mail.HTMLBody) - startIndex + 1), endIndex - startIndex + 2)
                        mail.HTMLBody = Replace(mail.HTMLBody, htmlTag, "<i>" & text & "</i>")
                        countImg = countImg + 1
                    Else
                        ' Kein HTML-Tag vorhanden
                        text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", countAtt) & "<a href=""" & attachmentUpdates(i).filename & """>" & attachmentUpdates(i).filename & "</a><br/>"
                        ' --> Link wird am Anfang eingefügt
                        msgUpdate = msgUpdate & text & vbCrLf
                        countAtt = countAtt + 1
                    End If
                Next i
                startIndex = InStr(mail.HTMLBody, "<body")
                startIndex = InStr(startIndex + 1, mail.HTMLBody, ">")
                mail.HTMLBody = Left(mail.HTMLBody, startIndex) & "<p><i>" & msgUpdate & "</i></p><hr/><br/>" & Right(mail.HTMLBody, Len(mail.HTMLBody) - startIndex)
            ElseIf mail.BodyFormat = olFormatRichText Then
                Debug.Print "  Aktualisiere RTF-Inhalt"
                                
                mailProtection = mailEditor.ProtectionType
                If (mailProtection <> wdNoProtection) Then
                    mailEditor.UnProtect
                End If
                                
                For i = 0 To archivedAttachments - 1
                    text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", i + 1) & """file://" & attachmentUpdates(i).filename & """"
                    position = attachmentUpdates(i).position + offset - 1
                    
                    mailEditor.Characters(position).InsertAfter (text)
                    ' Text kursiv machen
                    mailEditor.Range(position, position + Len(text)).Italic = True
                    ' Ich weiss nicht genau warum, aber da ist ein offset in der Position drin, der sich von Anhang zu Anhang aufsummiert
                    ' Und dann noch den neu eingefügten Link berücksichtigen
                    offset = offset - 31 + Len(text)
                Next i
                
                ' Editor aktivieren, damit Änderungen korrekt gespeichert werden
                mailEditor.Activate
            Else
                Debug.Print "  Aktualisiere Plain-Text-Inhalt"
                msgUpdate = ""
                maxLength = 0
                For i = 0 To archivedAttachments - 1
                    text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", i + 1) & """file://" & attachmentUpdates(i).filename & """"
                    msgUpdate = msgUpdate & text & vbCrLf
                    If (Len(text) > maxLength) Then
                        maxLength = Len(text)
                    End If
                Next i
                mail.Body = msgUpdate & vbCrLf & String(maxLength * 1.5, "-") & vbCrLf & vbCrLf & mail.Body
            End If
        
            Debug.Print "  Speichere aktualisierte Mail"
            mail.Save
        End If
        
        HandleAttachments = archivedAttachments
    Else
        HandleAttachments = 0
    End If
End Function

'--------------------------------------------------
' OLE Bilder behandeln
'--------------------------------------------------
Private Function HandleOLEImages(mail As MailItem, del As Boolean, fileNamePattern As String) As Integer
    If (mail.BodyFormat = olFormatRichText) Then
    
        ' Ein paar Variablen
        Dim filename As String
        Dim fileFolder As String
        Dim overwrite As Boolean
        Dim att As attachment
        Dim position As Integer
        Dim offset As Integer: offset = 0
        Dim counter As Integer
        Dim i As Integer
        Dim text As String
        Dim w As Integer
        Dim h As Integer
        Dim estimatedSize As Long
        
        ' RTF-Word-Editor
        Dim mailInspector As Outlook.Inspector: Set mailInspector = mail.GetInspector
        Dim mailEditor As Word.Document: Set mailEditor = mailInspector.WordEditor
        Dim mailProtection As WdProtectionType
        Dim ishp As Word.InlineShape
        Dim ishpRng As Word.Range
        Dim pic As IPictureDisp
        Dim PicSave As PicSave: Set PicSave = New PicSave
        
        Dim archivedOLEs As Integer: archivedOLEs = 0
        Dim attachmentUpdates() As AttachmentUpdate: ReDim attachmentUpdates(mailEditor.InlineShapes.count)
                
        counter = 1
        For Each ishp In mailEditor.InlineShapes
            ' Dateiname aus Pattern ermitteln
            filename = fileNamePattern
            filename = Replace(filename, "%FILENAME", "image" & Format(counter, "000") & ".png")
            If (AttachmentConfig.REPLACE_SPACES) Then
                filename = Replace(filename, " ", "_")
            End If
            filename = AttachmentConfig.ARCHIVE_FOLDER & "\" & filename
            
            If ishp.Type = Word.WdInlineShapeType.wdInlineShapePicture Then
                Set ishpRng = ishp.Range
                ishpRng.CopyAsPicture
                Set pic = ClipboardUtil.PastePicture(xlBitmap)
                w = Round(pic.Width / FACTOR_HIMETRIC)
                h = Round(pic.Height / FACTOR_HIMETRIC)
                estimatedSize = CLng(3) * w * h
                
                If (estimatedSize > AttachmentConfig.MIN_IMAGE_SIZE) Then
                    Debug.Print "  Archiviere OLE " & counter & ": (Type=" & ishp.Type & ", Größe=" & estimatedSize & ") -> " & filename
                                                  
                    fileFolder = Left(filename, InStrRev(filename, "\") - 1)
                    If Dir(fileFolder, vbDirectory) = "" Then
                        ' Ordner existiert nicht und muss erstellt werden
                        MkDir (fileFolder)
                        overwrite = True
                    ElseIf Dir(filename) <> "" Then
                        ' Datei existiert bereits
                        overwrite = ShowOverwrite(filename)
                    Else
                        overwrite = True
                    End If
                    
                    If (overwrite) Then
                        Call PicSave.SavePicture(pic, filename, fmtPNG)
                        If (del) Then
                            attachmentUpdates(archivedOLEs).filename = filename
                            attachmentUpdates(archivedOLEs).position = ishpRng.End
                            Set attachmentUpdates(archivedOLEs).shapeItem = ishp
                            ' find matching attachment item for this OLE (so we can delete it later)
                            For Each att In mail.Attachments
                                If (att.position = ishpRng.End) Then
                                    Set attachmentUpdates(archivedOLEs).attachmentItem = att
                                End If
                            Next att
                        End If
                        archivedOLEs = archivedOLEs + 1
                    Else
                        Debug.Print "  Überspringe OLE " & counter & ": (Type=" & ishp.Type & ", Größe=" & estimatedSize & ") -> DATEI EXISTIERT BEREITS"
                    End If
                Else
                    Debug.Print "  Überspringe OLE " & counter & ": (Type=" & ishp.Type & ", Größe=" & estimatedSize & ")"
                End If
            Else
                Debug.Print "  Überspringe OLE " & counter & ": (Type=" & ishp.Type & ")"
            End If
            counter = counter + 1
        Next ishp
    
        If (del And archivedOLEs > 0) Then
              
            mailProtection = mailEditor.ProtectionType
            If (mailProtection <> wdNoProtection) Then
                mailEditor.UnProtect
            End If
            
            Dim countOLE As Integer: countOLE = 1
            offset = 0
            For i = archivedOLEs - 1 To 0 Step -1
                text = Replace(MSG_IMG_IN_MAIL_TEXT, "%I", i + 1) & """file://" & attachmentUpdates(i).filename & """"
                position = attachmentUpdates(i).position + offset - 1
                
                attachmentUpdates(i).shapeItem.Range.InsertAfter (text)
                position = attachmentUpdates(i).shapeItem.Range.End
                ' Text kursiv machen
                mailEditor.Range(position, position + Len(text)).Italic = True
                ' Shape entfernen
                attachmentUpdates(i).shapeItem.Delete
            Next i
                        
            ' Editor aktivieren, damit Änderungen korrekt gespeichert werden
            mailEditor.Activate
            Debug.Print "  Speichere aktualisierte Mail"
            mail.Save
        End If
        
        HandleOLEImages = archivedOLEs
    Else
        HandleOLEImages = 0
    End If
    
End Function

'--------------------------------------------------
' Ersetze alle Platzhalter im FILENAME_PATTERN für diese Mail
'--------------------------------------------------
Public Function GetFileNamePattern(mail As MailItem) As String
    Dim fileNamePattern As String
    Dim address As String
    Dim name As String
    Dim rec As Integer
    
    fileNamePattern = AttachmentConfig.FILENAME_PATTERN
    
    ' generelle Platzhalter ersetzen
    fileNamePattern = Replace(fileNamePattern, "%DATETIME", Format(mail.ReceivedTime, AttachmentConfig.DATE_FORMAT))
    fileNamePattern = Replace(fileNamePattern, ":", ".")
    ' Mystischer Check, ob Email von einem selbst gesendet wurde
    If (mail.ReceivedByName = "") Then
        ' Message wurde von einem selbst gesendet
        fileNamePattern = Replace(fileNamePattern, "%DIRECTION", AttachmentConfig.DIRECTION_TO)
        ' Für jeden Empfänger einen Eintrag in den Platzhaltern vornehmen
        For rec = 1 To mail.Recipients.count
            ' nur "An" Empfaenger (Type = 1) betrachten
            If (mail.Recipients.Item(rec).Type = 1) Then
                address = GetAddress(mail.Recipients.Item(rec).addressEntry)
                name = GetName(mail.Recipients.Item(rec).addressEntry)
                fileNamePattern = Replace(fileNamePattern, "%CONTACTMAIL", address & ",%CONTACTMAIL")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTNAME", mail.Recipients.Item(rec).name & ",%CONTACTNAME")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTSYMBOL", GetKuerzel(address) & ",%CONTACTSYMBOL")
            End If
        Next
        ' Am Ende der Aufzählung noch den letzten Platzhalter ersetzen
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTMAIL", "")
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTNAME", "")
        fileNamePattern = Replace(fileNamePattern, ",%CONTACTSYMBOL", "")
    Else
        address = GetAddress(mail.Sender)
        name = GetName(mail.Sender)
        fileNamePattern = Replace(fileNamePattern, "%DIRECTION", AttachmentConfig.DIRECTION_FROM)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTMAIL", address)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTNAME", mail.SenderName)
        fileNamePattern = Replace(fileNamePattern, "%CONTACTSYMBOL", GetKuerzel(address))
    End If
    
    GetFileNamePattern = fileNamePattern
End Function


'--------------------------------------------------
' Lese das Kürzel zu einer Mail-Adresse aus der Kürzel-Datei
' Kürzel-Datei ist in der Form Aufgebaut:
'  KRZL1=person1@company1.com
'  KRZL2=person2@company2.com
'--------------------------------------------------
Public Function GetKuerzel(mailAdress As String) As String
    ' Default-Wert setzen
    GetKuerzel = mailAdress
        
    Dim kuerzelTxt As String
    ' Datei lesen (mit utf-8!)
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
    ' Gehe alle Zeilen durch und suche Zeilen, in denen der Teil
    ' nach dem "=" der Mail-Adresse entspricht
    Do While Len(kuerzelTxt) > 0
        line = Left(kuerzelTxt, InStr(kuerzelTxt, vbCrLf) - 1)
        If (StringUtil.StartsWith(line, "#")) Then
            ' Kommentar-Zeile ignorieren
        ElseIf (StringUtil.EndsWith(line, mailAdress)) Then
            GetKuerzel = Left(line, InStr(line, "=") - 1)
        End If
        kuerzelTxt = Right(kuerzelTxt, Len(kuerzelTxt) - InStr(kuerzelTxt, vbCrLf) - 1)
    Loop
End Function

'--------------------------------------------------
' Finde zugehoerigen Exchange User
'--------------------------------------------------
Public Function ResolveExchangeUser(addressEntry As addressEntry) As ExchangeUser
    If Not (addressEntry.GetExchangeUser Is Nothing) Then
        Set ResolveExchangeUser = addressEntry.GetExchangeUser
    Else
        ' Try to resolve Exchange User
        Set olNS = Application.GetNamespace("MAPI")
        Set olRecip = olNS.CreateRecipient(addressEntry.address)
        olRecip.Resolve
        Set ResolveExchangeUser = olRecip.addressEntry.GetExchangeUser
    End If
End Function

'--------------------------------------------------
' Extrahiere die Adresse aus einem Adress-Eintrag
' Wird benötigt, da das Vorgehen für Exchange-Nutzer anders ist
' als für externe Kontakte
'--------------------------------------------------
Public Function GetAddress(addressEntry As addressEntry) As String
    Set exchUser = ResolveExchangeUser(addressEntry)
    If Not (exchUser Is Nothing) Then
        GetAddress = exchUser.PrimarySmtpAddress
    Else
        GetAddress = addressEntry.address
    End If
End Function

'--------------------------------------------------
' Extrahiere den Namen aus einem Adress-Eintrag
' Wird benötigt, da das Vorgehen für Exchange-Nutzer anders ist
' als für externe Kontakte
'--------------------------------------------------
Public Function GetName(addressEntry As addressEntry) As String
    Set exchUser = ResolveExchangeUser(addressEntry)
    If Not (exchUser Is Nothing) Then
        GetName = exchUser.name
    Else
        GetName = addressEntry.name
    End If
End Function


'--------------------------------------------------
' Zeige den Zusammenfassungsdialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
' Gibt folgende mögliche Optionen zurück (siehe interne Konstanten)
'   OPTION_WITH_DELETE
'   OPTION_WITHOUT_DELETE
'   OPTION_CANCEL
'--------------------------------------------------
Private Function ShowConfirm() As Integer
    Dim opt As Integer
    If (AttachmentConfig.SHOW_CONFIRM) Then
        opt = MsgBox(MSG_CONFIRM, vbYesNoCancel, MSG_TITLE)
        If (opt = vbYes) Then
            ShowConfirm = 1
        ElseIf (opt = vbNo) Then
            ShowConfirm = 0
        Else
            ShowConfirm = -1
        End If
    Else
        ShowConfirm = IIf(AttachmentConfig.DELETE_ATTACHMENTS, 1, 0)
    End If
End Function

'--------------------------------------------------
' Zeige den Zusammenfassungsdialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
'--------------------------------------------------
Private Sub ShowSummary(selectedMails As Integer, archivedMails As Integer, archivedAttachments As Integer, archivedBytes As Long)
    Dim msg As String
    Dim sizeUnit As String
    Dim sizeScaled As Double
    Dim FACTOR As Double: FACTOR = 1024#
        
    If (selectedMails = 0) Then
        msg = MSG_MAILS_0
    ElseIf (selectedMails = 1) Then
        If (archivedAttachments = 0) Then
            msg = MSG_ARCHIVED_0
        ElseIf (archivedAttachments = 1) Then
            msg = archivedAttachments & MSG_ARCHIVED_1
        Else
            msg = archivedAttachments & MSG_ARCHIVED_N
        End If
    Else
        If (archivedAttachments = 0) Then
            msg = archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & MSG_ARCHIVED_0
        ElseIf (archivedAttachments = 1) Then
            msg = archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & archivedAttachments & MSG_ARCHIVED_1
        Else
            msg = archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & archivedAttachments & MSG_ARCHIVED_N
        End If
    End If
    
    If (archivedBytes < FACTOR) Then
        sizeUnit = "B"
        sizeScaled = archivedBytes
    ElseIf (archivedBytes < FACTOR * FACTOR) Then
        sizeUnit = "KB"
        sizeScaled = Round(archivedBytes / FACTOR, 2)
    Else
        sizeUnit = "MB"
        sizeScaled = Round(archivedBytes / FACTOR / FACTOR, 2)
    End If
    
    msg = msg & vbCrLf & MSG_SIZE & sizeScaled & " " & sizeUnit
    
    Debug.Print msg
        
    If (AttachmentConfig.SHOW_SUMMARY) Then
        Call MsgBox(msg, vbOKOnly, MSG_TITLE)
    End If
End Sub

'--------------------------------------------------
' Zeige den Datei-Überschreiben?-Dialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
' Gibt TRUE oder FALSE zurück
'--------------------------------------------------
Private Function ShowOverwrite(filename As String) As Boolean
    Dim opt As Integer
    If (AttachmentConfig.SHOW_OVERWRITE) Then
        opt = MsgBox(Replace(MSG_OVERWRITE, "%FILENAME", filename), vbYesNo, MSG_TITLE)
        If (opt = vbYes) Then
            ShowOverwrite = True
        Else
            ShowOverwrite = False
        End If
    Else
        ShowOverwrite = AttachmentConfig.OVERWRITE_EXISTING_FILES
    End If
End Function

'--------------------------------------------------
' Zeige den Kein-Kürzel-Gefunden-Dialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
' Gibt TRUE oder FALSE zurück
'--------------------------------------------------
Private Function ShowNoKuerzel(filename As String) As Boolean
    Dim opt As Integer
    If (AttachmentConfig.SHOW_NO_KUERZEL) Then
        opt = MsgBox(Replace(MSG_NO_KUERZEL, "%FILENAME", filename), vbYesNo, MSG_TITLE)
        If (opt = vbYes) Then
            ShowNoKuerzel = True
        Else
            ShowNoKuerzel = False
        End If
    Else
        ShowNoKuerzel = True
    End If
End Function

