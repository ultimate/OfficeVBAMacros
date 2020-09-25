Attribute VB_Name = "AttachmentConfig"
'--------------------------------------------------
' Definition der benötigen Konstanten und Konfiguration
'--------------------------------------------------
' Pfad zur Kürzel-Datei
Public Const KUERZEL_FILE As String = "D:\Project_X\kuerzel.properties"
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
' Sollen Leerzeichen im original Dateinamen ersetzt werden
' (Leerzeichem im Pfad zum Archiv-Ordner werden selbstverständlich nicht ersetzt)
Public Const REPLACE_SPACES As Boolean = False
' Minimale Größe zum Archivieren in Bytes
'   MIN_MAIL_SIZE = Mindestgröße für die Gesamte E-Mail ab der archiviert wird
'   MIN_FILE_SIZE = Mindestgröße für "normale" Dateien
'   MIN_IMAGE_SIZE = Mindestgröße für eingebettete Bilder (RTF oder HTML)
' Hinweis: Die tatsächliche Größe auf der Festplatte kann bei OLE-Objekten und OLE-Bildern
'   von der Speichergröße in der Mail abweichen (in der Mail sind die Datei i.d.R. Größer).
' Hinweis: 12000 reicht damit Firmen-Logo nicht archiviert werden.
'   Wenn man das nicht will, einfach 0 oder anderen kleineren Wert eintragen
Public Const MIN_MAIL_SIZE As Long = 0
Public Const MIN_FILE_SIZE As Long = 0
Public Const MIN_IMAGE_SIZE As Long = 12000
' Soll eine Bestätigung/Zusammenfassung/Überschreiben-Dialog angezeigt werden?
Public Const SHOW_CONFIRM As Boolean = True
Public Const SHOW_SUMMARY As Boolean = True
Public Const SHOW_OVERWRITE As Boolean = False
' Sollen Anhänge entfernt werden?
' Option wird genutzt, wenn kein Bestätigungs-Dialog angezeigt werden soll
Public Const DELETE_ATTACHMENTS As Boolean = True
' Sollen vorhandene Dateien überschrieben werden?
' Option wird genutzt, wenn kein Überschreiben-Dialog angezeigt werden soll
Public Const OVERWRITE_EXISTING_FILES As Boolean = True
