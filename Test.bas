Attribute VB_Name = "Test"


'--------------------------------------------------
' Test-Mail-Funktionen
'--------------------------------------------------
Public Sub Test_Attachment_Config_And_Kuerzels()
    Debug.Print "Konstanten"
    Debug.Print "  C2035_SVN_FOLDER   >" & AttachmentConfig.SVN_FOLDER
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
    
    Debug.Print "aaa"
    
    #If Win64 = 1 Then
        Debug.Print "Win64"
    #End If
    #If VBA7 = 1 Then
        Debug.Print "VBA7"
    #End If
End Sub

'--------------------------------------------------
' Debug
'--------------------------------------------------
Public Sub Debug_Attachments()
    Dim mail As MailItem
    Set mail = Application.ActiveInspector.CurrentItem
    
    Debug.Print "size=" & mail.Size
    Debug.Print "anz=" & mail.Attachments.count
    
    Dim att As attachment
    For Each att In mail.Attachments
        
        Debug.Print att.DisplayName
        Debug.Print "  type=" & att.Type
        If (att.Type = olOLE) Then
            Debug.Print "  file=embedded"
        Else
            Debug.Print "  file=" & att.fileName
        End If
        Debug.Print "  size=" & att.Size
        Debug.Print "  pos=" & att.position
    Next att
       
    'RTF-Word-Editor
    Dim mailInspector As Outlook.Inspector: Set mailInspector = mail.GetInspector
    Dim mailEditor As Word.Document: Set mailEditor = mailInspector.WordEditor
    
    Dim ishp As Word.InlineShape
    Dim sRng As Word.Range
    Dim sty As Word.Style
    Dim pic As IPictureDisp
    Dim nbr As Integer: nbr = 0
    
    Dim PicSave As PicSave: Set PicSave = New PicSave
    Dim w
    Dim h
    
    Debug.Print Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject & "=wdInlineShapeEmbeddedOLEObject"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject & "=wdInlineShapeLinkedOLEObject"
    Debug.Print Word.WdInlineShapeType.wdInlineShapePicture & "=wdInlineShapePicture"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeLinkedPicture & "=wdInlineShapeLinkedPicture"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeOLEControlObject & "=wdInlineShapeOLEControlObject"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeHorizontalLine & "=wdInlineShapeHorizontalLine"
    Debug.Print Word.WdInlineShapeType.wdInlineShapePictureHorizontalLine & "=wdInlineShapePictureHorizontalLine"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeLinkedPictureHorizontalLine & "=wdInlineShapeLinkedPictureHorizontalLine"
    Debug.Print Word.WdInlineShapeType.wdInlineShapePictureBullet & "=wdInlineShapePictureBullet"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeScriptAnchor & "=wdInlineShapeScriptAnchor"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeOWSAnchor & "=wdInlineShapeOWSAnchor"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeChart & "=wdInlineShapeChart"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeDiagram & "=wdInlineShapeDiagram"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeLockedCanvas & "=wdInlineShapeLockedCanvas"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeSmartArt & "=wdInlineShapeSmartArt"
    Debug.Print Word.WdInlineShapeType.wdInlineShapeWebVideo & "=wdInlineShapeWebVideo"
    Debug.Print "shapes=" & mailEditor.InlineShapes.count
                
    For Each ishp In mailEditor.InlineShapes
        Debug.Print nbr & " type=" & ishp.Type & " alt=" & ishp.AlternativeText
        If ishp.Type = Word.WdInlineShapeType.wdInlineShapePicture Then
            Set sRng = ishp.Range
            Set sty = sRng.Style
            Debug.Print sty.NameLocal
            If (StringUtil.StartsWith(sty.NameLocal, "zLGP")) Then
                'nothing (I'm bypassing certain inline graphics with known style names)
            Else
                sRng.CopyAsPicture
                Set pic = ClipboardUtil.PastePicture(xlBitmap)
                
                w = Round(pic.Width / AttachmentConfig.FACTOR_HIMETRIC)
                h = Round(pic.Height / AttachmentConfig.FACTOR_HIMETRIC)
                Debug.Print "    Dimension~=" & w & "x" & h
                Debug.Print "    Size~=" & (w * h * 3)
                'Debug.Print "    Size=" & Picture1.ScaleX(pic.Width, vbHimetric, vbPixels) & "x" & Picture1.ScaleY(pic.Height, vbHimetric, vbPixels)
                'PicSave.SavePicture pic, "Y:\Eigene Dateien\_Archiv\" & Format(nbr, "0000") & ".png", fmtPNG
            End If
        End If
        nbr = nbr + 1
    Next
End Sub


Public Sub Convert_to_HTML()
    
    Dim mail As MailItem
    Set mail = Application.ActiveInspector.CurrentItem
    
    'mail.BodyFormat = olFormatHTML
    
End Sub

Public Sub Test_Something()

    Dim mail As MailItem
    
    For Each mail In Application.ActiveExplorer.Selection
        Debug.Print mail.Categories
    Next mail

End Sub
