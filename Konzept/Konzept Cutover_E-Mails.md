Konzepot für die Erstellung von E-Mails im Rahmen des Cutovers

# Zielsetzung
Dieses Dokument beschreibt das Konzept für die Erstellung und den Versand von E-Mails während des Cutover-Prozesses. Ziel ist es, eine klare und konsistente Kommunikation mit allen beteiligten Stakeholdern sicherzustellen. 
Für die umsetzung des Ziels ist es notwendig, aus dem bestehenden Cutoverplan (Excel-Datei) die Aktivitäten auszulesen und für jede Aktivität ein E-Mail via MS Outlook zu erzeugen.
# Anforderungen an Verarbeitung
1. Für die Erzeugung der E-MAils ist ein Python-Script zu erstellen.
2. Das Script soll eine Nutzeroberfläche enthalten in dem die Eingabevariablen für die Verarbeitung abgefragt werden
Als Eingabevariable im Script sollen:
2.1 der Windows-Pfad und die die Excel-Datei aus dem Windows-Verzeichnis ausgewählt werden, in der der Cutover-Plan mit den Aktivitäten abgelegt ist. 
2.2 Es soll das Tabellenblatt ausgewählt werden, in dem die Cutover-Aktivitäten enthalten sind.
2.3 Es soll eine Ident des Cutoverplans eingegeben werden, dass im Betreff des E-Mails und im Inhalt des Mails aufgeführt wird.
3. Als Ergebnis der Verarbeitung sollen E-Mails für Microsoft Outlook für jede Aktivität in der Cutoverplanung erstellt werden. 
Die E-Mails sollen wie folgt Strukturiert werden:
--->
An: <e-mail des Ausführenden>
Bcc: hendrik.max4@dhl.com
Betreff: <Ident Cutover> - <Ident der Aktivität> - <Kurzbeschreibung Aktivität>
Inhalt: 
Hallo,
bitte führe die folgende Cutover-Aktivität
<Ident Aktivität> - <Beschreibung der Aktivität>
am: <Plan-Start-Datum Aktivität> 
im System: <System/Mandant-Buchungskreis>
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status im Cutoverplan <Ident Cutoverplan> in der o.a. Cutover-Aktivität ein.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus.

Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com

<-----

# Prompt: Bitte plane die Umsetzung des Konzepts. Bitte stelle Fragen, wenn in der Aufgabnestellung etwas unklar ist. Ziel der Planung ist ein detalliertes Umsetzungskonzept. Das Konzept befindet sich in in @Konzept. Das Python-Script soll im Ordner @Script erstellt werden.

