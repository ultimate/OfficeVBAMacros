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
    Private Const MSG_CONFIRM As String = "Anh�nge archivieren und entfernen?" & vbCrLf _
                                & "  Ja = Archivieren UND Entfernen" & vbCrLf _
                                & "  Nein = Archivieren OHNE Entfernen" & vbCrLf _
                                & "  Abbrechen = Keine Aktion durchf�hren" & vbCrLf _
                                & vbCrLf _
                                & "Achtung: Nutzung auf eigene Gefahr!"
    Private Const MSG_OVERWRITE As String = "Datei bereits vorhanden! �berschreiben?" & vbCrLf _
                                & "  %FILENAME"
    Private Const MSG_ARCHIVED_0 As String = "Keine Anh�nge zum archivieren vorhanden"
    Private Const MSG_ARCHIVED_1 As String = " Anhang erfolgreich archiviert"
    Private Const MSG_ARCHIVED_N As String = " Anh�nge erfolgreich archiviert"
    Private Const MSG_MAILS_0 As String = "Keine Mails ausgew�hlt"
    Private Const MSG_MAILS_OF As String = " von "
    Private Const MSG_MAILS_N As String = " ausgew�hlten Mails enthielten Anh�nge"
    Private Const MSG_SIZE As String = "Gesamtgr��e = %SIZE"
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
    Private Const MSG_ARCHIVED_0 As String = "No attachments available for archiving"
    Private Const MSG_ARCHIVED_1 As String = " attachment successfully archived"
    Private Const MSG_ARCHIVED_N As String = " attachments successfully archived"
    Private Const MSG_MAILS_0 As String = "No mail selected"
    Private Const MSG_MAILS_OF As String = " of "
    Private Const MSG_MAILS_N As String = " selected mails contained attachments"
    Private Const MSG_SIZE As String = "Total size = %SIZE"
    Private Const MSG_ATT_IN_MAIL_TEXT As String = "Attachment %I removed and archived at: "
    Private Const MSG_IMG_IN_MAIL_TEXT As String = "Image %I removed and archived at: "
#End If
' Factor f�r die Umrechnung von HIMETRIC in Pixel
Private Const FACTOR_HIMETRIC As Double = 26.45833

'--------------------------------------------------
' Custom Type f�r die Bearbeitung der Mails
'--------------------------------------------------
Private Type AttachmentUpdate
     fileName As String
     position As Integer
     attachmentName As String
     attachmentItem As attachment
     shapeItem As Word.InlineShape
End Type

'--------------------------------------------------
' Custom Type f�r die R�ckgabe des Archivierungsergebnisses
'--------------------------------------------------
Private Type ArchiveResult
     archivedAttachments As Integer
     archivedBytes As Long
End Type

'--------------------------------------------------
' Verwende diese Funktion, wenn eine Mail im separaten Fenster angezeigt wird
'--------------------------------------------------
Public Sub ArchiveAttachmentsForCurrentMail()
    Dim mail As MailItem
    Dim archivedAttachments As Integer
    Dim opt As Integer
    
    opt = ShowConfirm
    If (opt = OPTION_CANCEL) Then Exit Sub
    
    Set mail = Application.ActiveInspector.CurrentItem
    archivedAttachments = ArchiveAttachments(mail, opt = OPTION_WITH_DELETE)
    
    Call ShowSummary(1, 1, archivedAttachments, 0)
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
    Dim opt As Integer
    
    opt = ShowConfirm
    If (opt = OPTION_CANCEL) Then Exit Sub
    
    totalArchivedAttachments = 0
    selectedMails = 0
    archivedMails = 0
    
    For Each mail In Application.ActiveExplorer.Selection
        archivedAttachments = ArchiveAttachments(mail, opt = OPTION_WITH_DELETE)
        totalArchivedAttachments = totalArchivedAttachments + archivedAttachments
        selectedMails = selectedMails + 1
        If (archivedAttachments > 0) Then
            archivedMails = archivedMails + 1
        End If
    Next mail
    
    Call ShowSummary(selectedMails, archivedMails, totalArchivedAttachments, 0)
End Sub

'--------------------------------------------------
' Archiviere und Entferne die Anh�nge zur der �bergebenen Mail
'--------------------------------------------------
Public Function ArchiveAttachments(mail As MailItem, del As Boolean) As Integer
    
    ' Ein paar Variablen
    Dim fileName As String
    Dim fileFolder As String
    Dim fileNamePattern As String
    Dim saveFile As Boolean
    Dim att As attachment
    Dim position As Integer
    Dim offset As Integer: offset = 0
    Dim archivedAttachments As Integer: archivedAttachments = 0
    Dim attachmentUpdates() As AttachmentUpdate: ReDim attachmentUpdates(mail.Attachments.Count)
    Dim counter As Integer
    Dim i As Integer
    Dim text As String
        
    ' RTF-Word-Editor
    Dim mailInspector As Outlook.Inspector
    Dim mailEditor As Word.Document
    Dim mailProtection As WdProtectionType
    Dim ishp As Word.InlineShape
    Dim ishpRng As Word.Range
    Dim pic As IPictureDisp
    Dim PicSave As PicSave
    Dim archivedOLEs As Integer: archivedOLEs = 0
    
    ' RTF-Variablen nur bei RFT-Mail initialisieren
    If (mail.BodyFormat = olFormatRichText) Then
        Set mailInspector = mail.GetInspector
        Set mailEditor = mailInspector.WordEditor
        Set PicSave = New PicSave
        ReDim attachmentUpdates(mail.Attachments.Count + mailEditor.InlineShapes.Count)
    End If
            
    fileNamePattern = GetFileNamePattern(mail)
    
    Debug.Print "Archiviere Anh�nge im Pfad -> " & archiveFolder & "\" & fileNamePattern
    Debug.Print "  Anh�nge werden entfernt? " & del
        
    If (mail.Attachments.Count > 0) Then
        ' (normale) Anh�nge behandeln
        counter = 1
        For Each att In mail.Attachments
            'Debug.Print att.Type & " (ole=" & olOLE & ")"
            If (att.Type = olOLE) Then
                ' OLE-Bilder m�ssen separat behandelt werden!
                ' Speichern der Bilder ist nur �ber RTF-Word-Editor m�glich
                ' Normales Speichern resultiert in nicht lesbarem Bitmap
            ElseIf (att.Size >= AttachmentConfig.MIN_FILE_SIZE) Then
                
                ' Dateiname aus Pattern ermitteln
                fileName = fileNamePattern
                fileName = Replace(fileName, "%FILENAME", att.fileName)
                If (AttachmentConfig.REPLACE_SPACES) Then
                    fileName = Replace(fileName, " ", "_")
                End If
                fileName = AttachmentConfig.ARCHIVE_FOLDER & "\" & fileName
                
                fileFolder = Left(fileName, InStrRev(fileName, "\") - 1)
                If Dir(fileFolder, vbDirectory) = "" Then
                    ' Ordner existiert nicht und muss erstellt werden
                    MkDir (fileFolder)
                    overwrite = True
                ElseIf Dir(fileName) <> "" Then
                    ' Datei existiert bereits
                    overwrite = ShowOverwrite(fileName)
                Else
                    overwrite = True
                End If
                
                If (overwrite) Then
                    Debug.Print "  Archiviere Anhang " & counter & ": " & att.DisplayName & " (Gr��e=" & att.Size & ") -> " & fileName
                
                    att.SaveAsFile (fileName)
                    
                    attachmentUpdates(archivedAttachments).fileName = fileName
                    attachmentUpdates(archivedAttachments).attachmentName = att.fileName
                    attachmentUpdates(archivedAttachments).position = att.position
                    Set attachmentUpdates(archivedAttachments).attachmentItem = att
                    
                    archivedAttachments = archivedAttachments + 1
                Else
                    Debug.Print "  �berspringe Anhang " & counter & ": " & att.DisplayName & " (Gr��e=" & att.Size & ") -> DATEI EXISTIERT BEREITS"
                End If
            Else
                Debug.Print "  �berspringe Anhang " & counter & ": " & att.fileName & " (Gr��e=" & att.Size & ")"
            End If
            counter = counter + 1
        Next att
        
        If (del And archivedAttachments > 0) Then
            Dim msgUpdate As String
            Dim startIndex As Integer
            Dim endIndex As Integer
            Dim htmlTag As String
            Dim maxLength As Integer
            
            ' L�schen der Attachments darf erst ganz am Schluss erfolgen
            ' Sonst zerhaut man sich die Schleife
            For i = archivedAttachments - 1 To 0 Step -1
                attachmentUpdates(i).attachmentItem.Delete
            Next i
            
            'mail.Save
            
            If mail.BodyFormat = olFormatHTML Then
                Debug.Print "  Aktualisiere HTML-Inhalt"
                msgUpdate = ""
                Dim countImg As Integer: countImg = 0
                Dim countAtt As Integer: countAtt = 0
                For i = 0 To archivedAttachments - 1
                    ' Suche nach HTML TAG f�r eingebettetes Bild in der Form
                    ' <img width=85 height=76 id="Bild_x0020_1" src="cid:image001.png@01D29976.A64E5DB0">
                    startIndex = InStr(mail.HTMLBody, "src=""cid:" & attachmentUpdates(i).attachmentName & "@")
                    If (startIndex <> 0) Then
                        ' HTML-Tag gefunden
                        text = Replace(MSG_IMG_IN_MAIL_TEXT, "%I", countImg + 1) & "<a href=""" & fileName & """>" & attachmentUpdates(i).fileName & "</a><br/>"
                        ' --> Link an entsprechender Stelle einf�gen
                        startIndex = InStrRev(mail.HTMLBody, "<img", startIndex)
                        endIndex = InStr(startIndex, mail.HTMLBody, """>")
                        htmlTag = Left(Right(mail.HTMLBody, Len(mail.HTMLBody) - startIndex + 1), endIndex - startIndex + 2)
                        'Debug.Print "    Ersetze " & htmlTag
                        mail.HTMLBody = Replace(mail.HTMLBody, htmlTag, "<i>" & text & "</i>")
                        countImg = countImg + 1
                    Else
                        ' Kein HTML-Tag vorhanden
                        text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", countAtt + 1) & "<a href=""" & fileName & """>" & attachmentUpdates(i).fileName & "</a><br/>"
                        ' --> Link wird am Anfang eingef�gt
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
                    'Debug.Print "  Mail ist gegen Bearbeitung gesch�tzt: Hebe Schutz auf"
                    mailEditor.UnProtect
                End If
                                
                For i = 0 To archivedAttachments - 1
                    text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", i + 1) & """file://" & attachmentUpdates(i).fileName & """"
                    position = attachmentUpdates(i).position + offset - 1
                    Debug.Print "    inserting at " & position & ": " & text
                    
                    mailEditor.Characters(position).InsertAfter (text)
                    ' Text kursiv machen
                    mailEditor.Range(position, position + Len(text)).Italic = True
                    ' Ich weiss nicht genau warum, aber da ist ein offset in der
                    ' Position drin, der sich von Anhang zu Anhang aufsummiert
                    ' Und dann noch den neu eingef�gten Link ber�cksichtigen
                    offset = offset - 31 + Len(text)
                Next i
            Else
                Debug.Print "  Aktualisiere Plain-Text-Inhalt"
                msgUpdate = ""
                maxLength = 0
                For i = 0 To archivedAttachments - 1
                    text = Replace(MSG_ATT_IN_MAIL_TEXT, "%I", i + 1) & """file://" & attachmentUpdates(i).fileName & """"
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
        
        ' OLE Bilder behandeln
        counter = 1
        If (mail.BodyFormat = olFormatRichText) Then
            Dim w As Integer
            Dim h As Integer
            Dim estimatedSize As Long

            For Each ishp In mailEditor.InlineShapes
                ' Dateiname aus Pattern ermitteln
                fileName = fileNamePattern
                fileName = Replace(fileName, "%FILENAME", "image" & Format(counter, "000") & ".png")
                If (AttachmentConfig.REPLACE_SPACES) Then
                    fileName = Replace(fileName, " ", "_")
                End If
                fileName = AttachmentConfig.ARCHIVE_FOLDER & "\" & fileName
                
                If ishp.Type = Word.WdInlineShapeType.wdInlineShapePicture Then
                    Set ishpRng = ishp.Range
                    ishpRng.CopyAsPicture
                    Set pic = ClipboardUtil.PastePicture(xlBitmap)
                    w = Round(pic.Width / FACTOR_HIMETRIC)
                    h = Round(pic.Height / FACTOR_HIMETRIC)
                    Debug.Print "    Dimension~=" & w & "x" & h
                    estimatedSize = CLng(3) * w * h
                    Debug.Print "    Size~=" & estimatedSize
                    
                    If (estimatedSize > AttachmentConfig.MIN_IMAGE_SIZE) Then
                        Debug.Print "  Archiviere OLE " & counter & ": (Type=" & ishp.Type & ", Gr��e=" & estimatedSize & ") -> " & fileName
                              
                        fileFolder = Left(fileName, InStrRev(fileName, "\") - 1)
                        If Dir(fileFolder, vbDirectory) = "" Then
                            ' Ordner existiert nicht und muss erstellt werden
                            MkDir (fileFolder)
                            overwrite = True
                        ElseIf Dir(fileName) <> "" Then
                            ' Datei existiert bereits
                            overwrite = ShowOverwrite(fileName)
                        Else
                            overwrite = True
                        End If
                        
                        If (overwrite) Then
                            Call PicSave.SavePicture(pic, fileName, fmtPNG)
                            If (del) Then
                                attachmentUpdates(archivedAttachments + archivedOLEs).fileName = fileName
                                'attachmentUpdates(archivedAttachments + archivedOLEs).attachmentName = "image" & Format(counter, "000") & ".png"
                                Set attachmentUpdates(archivedAttachments + archivedOLEs).shapeItem = ishp
                            End If
                            archivedOLEs = archivedOLEs + 1
                        Else
                            Debug.Print "  �berspringe OLE " & counter & ": (Type=" & ishp.Type & ", Gr��e=" & estimatedSize & ") -> DATEI EXISTIERT BEREITS"
                        End If
                    Else
                        Debug.Print "  �berspringe OLE " & counter & ": (Type=" & ishp.Type & ", Gr��e=" & estimatedSize & ")"
                    End If
                Else
                    Debug.Print "  �berspringe OLE " & counter & ": (Type=" & ishp.Type & ")"
                End If
                counter = counter + 1
            Next ishp
        End If
        
        If (del And archivedOLEs > 0) Then
            ' L�schen der Attachments darf erst ganz am Schluss erfolgen
            ' Sonst zerhaut man sich die Schleife
            Dim countOLE As Integer: countOLE = archivedOLEs
            For i = archivedOLEs + archivedAttachments - 1 To archivedAttachments Step -1
                text = Replace(MSG_IMG_IN_MAIL_TEXT, "%I", countOLE) & """file://" & attachmentUpdates(i).fileName & """"
                attachmentUpdates(i).shapeItem.Range.InsertAfter (text)
                position = attachmentUpdates(i).shapeItem.Range.End
                ' Text kursiv machen
                mailEditor.Range(position, position + Len(text)).Italic = True
                ' Shape entfernen
                attachmentUpdates(i).shapeItem.Delete
                countOLE = countOLE - 1
            Next i
            
            Debug.Print "  Speichere aktualisierte Mail"
            mail.Save
        End If
        
    End If
        
    ArchiveAttachments = archivedAttachments
End Function

'--------------------------------------------------
' Ersetze alle Platzhalter im FILENAME_PATTERN f�r diese Mail
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
        ' F�r jeden Empf�nger einen Eintrag in den Platzhaltern vornehmen
        For rec = 1 To mail.Recipients.Count
            ' nur "An" Empfaenger (Type = 1) betrachten
            If (mail.Recipients.Item(rec).Type = 1) Then
                address = GetAddress(mail.Recipients.Item(rec).addressEntry)
                name = GetName(mail.Recipients.Item(rec).addressEntry)
                fileNamePattern = Replace(fileNamePattern, "%CONTACTMAIL", address & ",%CONTACTMAIL")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTNAME", mail.Recipients.Item(rec).name & ",%CONTACTNAME")
                fileNamePattern = Replace(fileNamePattern, "%CONTACTSYMBOL", GetKuerzel(address) & ",%CONTACTSYMBOL")
            End If
        Next
        ' Am Ende der Aufz�hlung noch den letzten Platzhalter ersetzen
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
' Lese das K�rzel zu einer Mail-Adresse aus der K�rzel-Datei
' K�rzel-Datei ist in der Form Aufgebaut:
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
' Extrahiere die Adresse aus einem Adress-Eintrag
' Wird ben�tigt, da das Vorgehen f�r Exchange-Nutzer anders ist
' als f�r externe Kontakte
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
' Wird ben�tigt, da das Vorgehen f�r Exchange-Nutzer anders ist
' als f�r externe Kontakte
'--------------------------------------------------
Public Function GetName(addressEntry As addressEntry) As String
    If (addressEntry.GetExchangeUser Is Nothing) Then
        GetName = addressEntry.name
    Else
        GetName = addressEntry.GetExchangeUser.name
    End If
End Function

'--------------------------------------------------
' Zeige den Zusammenfassungsdialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
' Gibt folgende m�gliche Optionen zur�ck (siehe interne Konstanten)
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
    If (AttachmentConfig.SHOW_SUMMARY) Then
        If (selectedMails = 0) Then
            Call MsgBox(MSG_MAILS_0, vbOKOnly, MSG_TITLE)
        ElseIf (selectedMails = 1) Then
            If (archivedAttachments = 0) Then
                Call MsgBox(MSG_ARCHIVED_0, vbOKOnly, MSG_TITLE)
            ElseIf (archivedAttachments = 1) Then
                Call MsgBox(archivedAttachments & MSG_ARCHIVED_1, vbOKOnly, MSG_TITLE)
            Else
                Call MsgBox(archivedAttachments & MSG_ARCHIVED_N, vbOKOnly, MSG_TITLE)
            End If
        Else
            If (archivedAttachments = 0) Then
                Call MsgBox(archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & MSG_ARCHIVED_0, vbOKOnly, MSG_TITLE)
            ElseIf (archivedAttachments = 1) Then
                Call MsgBox(archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & archivedAttachments & MSG_ARCHIVED_1, vbOKOnly, MSG_TITLE)
            Else
                Call MsgBox(archivedMails & MSG_MAILS_OF & selectedMails & MSG_MAILS_N & vbCrLf & archivedAttachments & MSG_ARCHIVED_N, vbOKOnly, MSG_TITLE)
            End If
        End If
    End If
End Sub

'--------------------------------------------------
' Zeige den Datei-�berschreiben?-Dialog
' Dialog wird nur angezeigt, wenn in AttachmentConfig aktiviert
' Gibt TRUE oder FALSE zur�ck
'--------------------------------------------------
Private Function ShowOverwrite(fileName As String) As Boolean
    Dim opt As Integer
    If (AttachmentConfig.SHOW_OVERWRITE) Then
        opt = MsgBox(Replace(MSG_OVERWRITE, "%FILENAME", fileName), vbYesNo, MSG_TITLE)
        If (opt = vbYes) Then
            ShowOverwrite = True
        Else
            ShowOverwrite = False
        End If
    Else
        ShowOverwrite = AttachmentConfig.OVERWRITE_EXISTING_FILES
    End If
End Function
