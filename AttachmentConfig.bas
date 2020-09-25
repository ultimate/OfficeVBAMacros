Attribute VB_Name = "AttachmentConfig"
'--------------------------------------------------
' Definition der benötigen Konstanten und Konfiguration
'--------------------------------------------------
' Pfad zum lokal ausgecheckten SVN
Public Const SVN_FOLDER As String = "D:\SVN\Project_X"
' Relativer Pfad zur Kürzel-Datei im SVN
' FEST! Nur ändern, wenn Kürzel-Datei außerhalb des SVNs abgelegt wird
Public Const KUERZEL_FILE As String = SVN_FOLDER & "\kuerzel.properties"
' Pfad zum Archiv für Email-Anhänge
Public Const ARCHIVE_FOLDER = "Y:\Eigene Dateien\_Archiv"
' Ablage-Namens-Struktur für Dateianhänge (kann Ordnerpfade enthalten)
' Hinweis: Darf keines der folgenden ungültigen Zeichen enthalten: /:*?"<>|
' Gültige Platzhalter:
' - %DATETIME           // Datum und Uhrzeit im definierten Format
'                       // siehe dazu DATE_FORMAT
' - %DIRECTION          // Senderichtung
'                       // siehe DIRECTION_FROM and DIRECTION_TO
' - %CONTACTMAIL        // Email-Adresse des Kommunikationspartners
' - %CONTACTNAME        // Name des Kommunikationspartners
' - %CONTACTSYMBOL      // Kürzel des Kommunikationspartners gemäß kuerzel.txt
'                       // Falls Kürzel nicht enthalten, wird die Email-Adresse ausgegeben
' - %FILENAME           // original Dateiname
Public Const FILENAME_PATTERN As String = "%DATETIME %DIRECTION %CONTACTSYMBOL\%FILENAME"
' Datums-Format
' siehe dazu https://msdn.microsoft.com/en-us/library/office/gg251755.aspx
' Hinweis: Darf keines der folgenden ungültigen Zeichen enthalten: \/:*?"<>|
Public Const DATE_FORMAT As String = "yyyy.mm.dd Hh.Nn.Ss"
' Text-Bausteine für %DIRECTION
Public Const DIRECTION_FROM As String = "von"
Public Const DIRECTION_TO As String = "an"
' Sollen Leerzeichen im Dateinamen ersetzt werden
' (Leerzeichem im Pfad zum Archiv-Ordner werden selbstverständlich nicht ersetzt)
Public Const REPLACE_SPACES As Boolean = False
' Minimale Dateigröße zum Archivieren in Bytes (verwende *1024 für KB; *1024*1024 für MB, etc.)
Public Const MIN_FILE_SIZE As Integer = 4500
' Soll eine Bestätigung/Zusammenfassung angezeigt werden?
Public Const SHOW_CONFIRM As Boolean = True
Public Const SHOW_SUMMARY As Boolean = True
' Sollen Anhänge entfernt werden?
' Option wird genutzt, wenn kein Bestätigungs-Dialog angezeigt werden soll
Public Const DELETE_ATTACHMENTS As Boolean = True
' Anzeige-Messages
Public Const MSG_CONFIRM As String = "Anhänge archivieren und entfernen?" & vbCrLf _
                            & "  Ja = Archivieren und Entfernen" & vbCrLf _
                            & "  Nein = Archivieren OHNE Entfernen" & vbCrLf _
                            & "  Abbrechen = Keine Aktion durchführen"
Public Const MSG_ARCHIVED_0 As String = "Keine Anhänge zum archivieren vorhanden"
Public Const MSG_ARCHIVED_1 As String = " Anhang erfolgreich archiviert"
Public Const MSG_ARCHIVED_N As String = " Anhänge erfolgreich archiviert"
Public Const MSG_MAILS_0 As String = "Keine Mails ausgewählt"
Public Const MSG_MAILS_OF As String = " von "
Public Const MSG_MAILS_N As String = " ausgewählten Mails enthielten Anhänge"
Public Const MSG_IN_MAIL_TEXT As String = "Anhang %I entfernt und archivert unter: "
