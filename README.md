# OfficeVBAMacros
My collection of VBA extensions for MS Office

# License Information
All of my code is free to use under MIT License.
But please respect integrated 3rd party code licenses stated in the respective files:
- https://github.com/ultimate/OfficeVBAMacros/blob/master/ClipboardUtil.bas
- https://github.com/ultimate/OfficeVBAMacros/blob/master/PicSave.cls
- https://github.com/ultimate/OfficeVBAMacros/blob/master/Base64.bas

# AttachmentUtil
## Informationen
Das AttachmentUtil for Outlook dient dazu automatisiert Anhänge aus E-Mails zu archivieren und anschließen zu entfernen um auf diese Weise wieder Speicher im Postfach freizugeben. Anschließend werden Links in die E-Mails zu den archivierten Anhängen eingefügt um den direkten Zugriff auf die Dateien weiterhin zu ermöglichen.
Dazu kann eine Konfiguration hinterlegt werden, wo (Ordnerpfad) und in welcher Form (relativer Pfad und Dateiname abhängig von der E-Mail und vom Anhang) die Anhänge archiviert werden sollen.
Der Archivierungsvorgang muss dafür manuell für eine oder mehrere ausgewählte Mails angestoßen werden. Es findet keine automatische Archivierung bei vollem Postfach oder anderen Ereignissen statt. Nach dem manuellen Anstoßen der Archivierung läuft das Makro ansonsten aber automatisch durch alle ausgewählten Mails durch und archiviert und entfernt automatisch alle Anhänge und eingebetteten Bilder gemäß Konfiguration. 
Achtung: Nutzung auf eigene Gefahr! :-)

## Benutzung
Nach der Installation (siehe unten) kann das AttachmentUtil wie folgt benutzt werden:

1.	Entsprechend benutzerdefinierten Button anklicken  
  Hinweis: Text und Symbol sind benutzerdefiniert änderbar (siehe Installation)
  
2.	Dialog bestätigen 
  Hinweis: Dialog kann in Konfiguration deaktiviert werden.
  
3.	Ggf. muss eine Bestätigung zum Überschreiben von Dateien erfolgen

4.	Anschließend wird eine Zusammenfassung angezeigt, wie viele E-Mails bearbeitet wurden und wieviel Speicher ungefähr freigegeben wurde.

5.	Das Ergebnis sieht dann abhängig vom Typ der E-Mail (Nur-Text, Rich-Text, HTML) wie folgt aus:

    a.	Sind die Dateien oder Bilder oben im E-Mail-Kopf angehängt, so werden die Links werden am Anfang der E-Mail eingefügt. 
    Gilt für:
    -	Datei-Anhänge von Nur-Text-Mails
    -	Datei-Anhänge von HTML-Mails
    -	Bild-Anhänge von HTML-Mails
    
    b.	Sind die Dateien oder Bilder im Text eingebettet, so werden die Links an der entsprechenden Stelle in den E-Mail-Text eingefügt:
    Gilt für:
    -	Datei-Anhänge (per Icon) in Rich-Text-Mails 
    -	Bild-Anhänge (per Icon) in Rich-Text-Mails
    -	Eingebettete Bilder in Rich-Text-Mails 
    -	Eingebettete Bilder in HTML –Mails

## Dateiliste
Folgende Dateien (Module/Klassenmodule) werden für die Funktion benötigt:
-	AttachmentConfig.bas
-	AttachmentUtil.bas
-	ClipboardUtil.bas
-	PicSave.cls
-	StringUtil.bas

## Voraussetzungen
1.	Outlook-Makro-Sicherheit prüfen
    a.	Datei >> Optionen >> Trust Center >> Einstellungen für das Trust Center
    b.	>> Makroeinstellungen >> Mindestens die Option „Benachrichtigungen für digital signierte Makros. Alle anderen Makros werden deaktiviert“ oder schlechter auswählen.
    
## Installation

### Dateien Einbinden:
1.	Visual Basic Editor öffnen 
2.	Alle erforderlichen Dateien importieren (s.o.)	

### Referenzen einbinden:
1.	Extras >> Verweise… 
2.	Folgende Referenzen (falls nicht vorhanden) aktivieren:
    - Visual Basic For Applications
    - Microsoft Outlook X.Y Object Library
    - OLE Automation
    - Microsoft Office X.Y Object Library
    - Microsoft Word X.Y Object Library
 
### Konfiguration:
1.	Die Konfiguration erfolgt über das Module „AttachmentConfig“ (AttachmentConfig.bas). Dort sind alle Einstellmöglichkeiten dokumentiert.
2.	Außerdem kann bei Bedarf die Sprache auf Englisch geändert werden. Dies muss im Module „AttachmentUtil“ (AttachmentUtil.bas) erfolgen. (Define ganz am Anfang)
    Hinweis: Es müssen mindestens folgende Eigenschaften angepasst werden:
    -	KUERZEL_FILE >> Pfad zur Datei zum Nachschlagen der Personen-Kürzel
    -	ARCHIVE_FOLDER >> Ordner für die Archivierung

### Einbindung in die GUI
1.	Es werden 2 unterschiedliche Funktionen für die Einbindung in die Outlook-GUI angeboten:
    - 	ArchiveAttachmentsForCurrentMail() >> Diese Funktion muss im E-Mail-Bearbeitungs-Fenster verwendet werden. Es werden immer nur die Anhänge der aktuell geöffneten E-Mail archiviert.
    - ArchiveAttachmentsForSelectedMails() >> Diese Funktion muss im Outlook-Hauptfenster und im Lesebereich verwendet werden. Es werden die Anhänge für alle im aktuellen Postfach oder Ordner markierten E-Mails archiviert.
2.	Hinzufügen der Makros wie folgt:

    a.	Rechtsklick auf entsprechenden Menübereich >> Menüband anpassen
    
    b.	Ggf. neue Gruppe und/oder Registerkarte anlegen
    
    c.	Unter „Makros“ entsprechendes Makro auswählen und hinzufügen
    
    d.	Name und Symbol vergeben
