# Anleitung: Cutover E-Mail Generator starten

## Methode 1: Per Doppelklick (Empfohlen - Am Einfachsten!)

### Start über Batch-Datei

**Die Batch-Datei ist bereits vorhanden!**

1. Navigiere zum Hauptordner:
   ```
   C:\Users\hendrik.max\Documents\DEV_LOCL\cutover_mail\
   ```
2. Doppelklick auf die Datei:
   ```
   Start_Cutover_Mail.bat
   ```
3. Das Programm-Fenster öffnet sich automatisch!

### Desktop-Verknüpfung erstellen (Optional)

Für noch schnelleren Zugriff:

1. Rechtsklick auf `Start_Cutover_Mail.bat` im Hauptordner
2. Wähle "Verknüpfung erstellen"
3. Ziehe die Verknüpfung auf den Desktop
4. Jetzt kannst Du das Programm direkt vom Desktop starten!

---

## Methode 2: Über Kommandozeile (Für Fortgeschrittene)

### Schritt 1: Kommandozeile öffnen
- Drücke `Windows-Taste + R`
- Tippe `cmd` ein und drücke Enter
- Oder: Suche nach "Eingabeaufforderung" im Startmenü

### Schritt 2: Zum Projektordner navigieren
```cmd
cd C:\Users\hendrik.max\Documents\DEV_LOCL\cutover_mail\Script
```

### Schritt 3: Programm starten
```cmd
.venv\Scripts\python.exe cutover_mail_generator.py
```

Das Programm-Fenster öffnet sich und Du kannst mit der E-Mail-Generierung beginnen.

---

## Fehlerbehebung

### "Python wurde nicht gefunden"
- Stelle sicher, dass Du im richtigen Ordner bist (`Script`)
- Prüfe, ob der `.venv` Ordner existiert

### "Modul nicht gefunden"
- Das Virtual Environment wurde möglicherweise gelöscht
- Führe folgende Befehle aus:
  ```cmd
  cd C:\Users\hendrik.max\Documents\DEV_LOCL\cutover_mail\Script
  python -m venv .venv
  .venv\Scripts\pip.exe install -r requirements.txt
  ```

### Programm startet nicht
- Prüfe, ob eine andere Instanz bereits läuft (Task-Manager)
- Starte den Computer neu und versuche es erneut

---

## Programmnutzung

Nach dem Start:
1. **Excel-Datei auswählen** - Klicke auf "Durchsuchen..." und wähle Deine Cutoverplan-Datei
2. **Tabellenblatt wählen** - Wähle das Blatt mit den Aktivitäten (z.B. "CuOv-Plan DPAG neu")
3. **Cutover-Ident eingeben** - z.B. "JOSEF" oder "DPN_ECH"
4. **Ausgabepfad wählen** - Ordner, wo die E-Mail-Dateien gespeichert werden sollen
5. **Filter setzen** (optional) - Nach IST-Status oder Aktivitäts-Ident filtern
6. **E-Mails generieren** - Klicke auf den Button und warte auf die Bestätigung

Die generierten EML-Dateien kannst Du per Doppelklick in Outlook öffnen.

---

**Version:** 1.0
**Datum:** 2025-11-18
**Autor:** Hendrik Max
