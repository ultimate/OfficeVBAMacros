Attribute VB_Name = "AttachmentConfig"
'--------------------------------------------------
' Definition der ben�tigen Konstanten und Konfiguration
'--------------------------------------------------
' Pfad zur K�rzel-Datei
Public Const KUERZEL_FILE As String = "D:\Project_X\kuerzel.properties"
' Pfad zum Archiv f�r Email-Anh�nge
Public Const ARCHIVE_FOLDER = "Y:\Eigene Dateien\_Archiv"
' Ablage-Namens-Struktur f�r Dateianh�nge (kann Ordnerpfade enthalten)
' Hinweis: Darf keines der folgenden ung�ltigen Zeichen enthalten: /:*?"<>|
' G�ltige Platzhalter:
' - %DATETIME           // Datum und Uhrzeit im definierten Format
'                       // siehe dazu DATE_FORMAT
' - %DIRECTION          // Senderichtung
'                       // siehe DIRECTION_FROM and DIRECTION_TO
' - %CONTACTMAIL        // Email-Adresse des Kommunikationspartners
' - %CONTACTNAME        // Name des Kommunikationspartners
' - %CONTACTSYMBOL      // K�rzel des Kommunikationspartners gem�� kuerzel.properties
'                       // Falls K�rzel nicht enthalten, wird die Email-Adresse ausgegeben
' - %FILENAME           // original Dateiname
Public Const FILENAME_PATTERN As String = "%DATETIME %DIRECTION %CONTACTSYMBOL\%FILENAME"
' Datums-Format
' siehe dazu https://msdn.microsoft.com/en-us/library/office/gg251755.aspx
' Hinweis: Darf keines der folgenden ung�ltigen Zeichen enthalten: \/:*?"<>|
Public Const DATE_FORMAT As String = "yyyy.mm.dd Hh.Nn.Ss"
' Text-Bausteine f�r %DIRECTION
Public Const DIRECTION_FROM As String = "von"
Public Const DIRECTION_TO As String = "an"
' Sollen Leerzeichen im original Dateinamen ersetzt werden
' (Leerzeichem im Pfad zum Archiv-Ordner werden selbstverst�ndlich nicht ersetzt)
Public Const REPLACE_SPACES As Boolean = False
' Minimale Gr��e zum Archivieren in Bytes
'   MIN_MAIL_SIZE = Mindestgr��e f�r die Gesamte E-Mail ab der archiviert wird
'   MIN_FILE_SIZE = Mindestgr��e f�r "normale" Dateien
'   MIN_IMAGE_SIZE = Mindestgr��e f�r eingebettete Bilder (RTF oder HTML)
' Hinweis: Die tats�chliche Gr��e auf der Festplatte kann bei OLE-Objekten und OLE-Bildern
'   von der Speichergr��e in der Mail abweichen (in der Mail sind die Datei i.d.R. Gr��er).
' Hinweis: 12000 reicht damit kleine Firmen-Logos nicht archiviert werden.
'   Wenn man das nicht will, einfach 0 oder anderen kleineren Wert eintragen
Public Const MIN_MAIL_SIZE As Long = 0
Public Const MIN_FILE_SIZE As Long = 0
Public Const MIN_IMAGE_SIZE As Long = 12000
' Soll eine Best�tigung/Zusammenfassung/�berschreiben-Dialog angezeigt werden?
Public Const SHOW_CONFIRM As Boolean = True
Public Const SHOW_SUMMARY As Boolean = True
Public Const SHOW_OVERWRITE As Boolean = False
Public Const SHOW_NO_KUERZEL As Boolean = True
' Sollen Anh�nge entfernt werden?
' Option wird genutzt, wenn kein Best�tigungs-Dialog angezeigt werden soll
Public Const DELETE_ATTACHMENTS As Boolean = True
' Sollen vorhandene Dateien �berschrieben werden?
' Option wird genutzt, wenn kein �berschreiben-Dialog angezeigt werden soll
Public Const OVERWRITE_EXISTING_FILES As Boolean = True
