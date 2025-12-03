# Cutover E-Mail Generator

Automatische Erstellung von Cutover-E-Mails aus Excel-Cutoverplan für Microsoft Outlook.

## Überblick

Dieses Python-Tool liest einen Excel-basierten Cutoverplan und erstellt automatisch personalisierte E-Mails für jede Cutover-Aktivität. Die E-Mails können entweder als Outlook-Entwürfe erstellt oder als .eml-Dateien gespeichert werden.

## Features

- **Excel-Integration**: Liest Cutover-Aktivitäten direkt aus Excel-Dateien
- **GUI-Oberfläche**: Benutzerfreundliche grafische Oberfläche mit tkinter
- **Flexible E-Mail-Erstellung**:
  - Outlook-Entwürfe (zur Prüfung vor dem Versand)
  - .eml-Dateien (zum späteren Öffnen in jedem E-Mail-Client)
- **Mehrere Empfänger**: Unterstützt mehrere E-Mail-Adressen pro Aktivität (getrennt durch `;` oder `,`)
- **Filter-Optionen**:
  - Nach IST-Status filtern
  - Nach Aktivitäts-Ident filtern
- **Fortschrittsanzeige**: Live-Updates während der Verarbeitung
- **Robuste Fehlerbehandlung**: 
  - Validierung von E-Mail-Adressen
  - Automatische Bereinigung von Zeilenumbrüchen in Aktivitätsbeschreibungen
  - Aussagekräftige Fehlermeldungen

## Voraussetzungen

- **Python**: Version 3.8 oder höher
- **Microsoft Outlook**: Muss installiert sein (für Outlook-Integration)
- **Betriebssystem**: Windows (wegen COM-Schnittstelle zu Outlook)

## Installation

### 1. Python installieren

Falls noch nicht vorhanden, laden Sie Python von [python.org](https://www.python.org/downloads/) herunter und installieren Sie es.

### 2. Virtual Environment erstellen (empfohlen)

```bash
# In das Script-Verzeichnis wechseln
cd C:\Users\hendrik.max\Documents\DEV_LOCL\cutover_mail\Script

# Virtual Environment erstellen
python -m venv venv

# Virtual Environment aktivieren
venv\Scripts\activate
```

### 3. Dependencies installieren

```bash
pip install -r requirements.txt
```

## Verwendung

### Programm starten

```bash
python cutover_mail_generator.py
```

### Schritt-für-Schritt-Anleitung

1. **Excel-Datei auswählen**
   - Klicken Sie auf "Durchsuchen..." bei "Excel-Datei"
   - Wählen Sie Ihre Cutoverplan-Datei (.xlsx) aus

2. **Tabellenblatt auswählen**
   - Wählen Sie aus dem Dropdown das Tabellenblatt mit den Cutover-Aktivitäten

3. **Cutover-Ident eingeben**
   - Geben Sie die Identifikation für den Cutover ein (z.B. "JOSEF", "DPN_ECH")
   - Diese erscheint im Betreff jeder E-Mail

4. **Link Cutover-Plan eingeben** (optional)
   - SharePoint-Link zum Cutover-Plan
   - Wird in der E-Mail als anklickbarer Link eingefügt

5. **Ausgabepfad wählen**
   - Klicken Sie auf "Durchsuchen..." und wählen Sie einen Ordner für die .eml-Dateien

6. **Filter setzen** (optional)
   - **IST-Status filtern**: Nur Aktivitäten mit bestimmtem Status (z.B. leer)
   - **Aktivitäts-Ident filtern**: Nur Aktivitäten mit bestimmtem Ident (z.B. "3.1")

7. **E-Mails generieren**
   - Klicken Sie auf "E-Mails generieren"
   - Der Fortschritt wird angezeigt
   - Bei Outlook-Modus: Die E-Mail-Entwürfe öffnen sich automatisch

## Excel-Struktur

Das Tool erwartet folgende Spalten in der Excel-Datei:

### Pflichtfelder

| Spaltenname | Beschreibung |
|-------------|--------------|
| **Ident** | Eindeutige Aktivitäts-ID (wird als String gelesen, z.B. "1.10") |
| **Aktivität** | Beschreibung der Aktivität (Zeilenumbrüche werden automatisch bereinigt) |
| **E-Mail** | E-Mail-Adresse(n) des Ausführenden (mehrere Adressen mit `;` oder `,` trennen) |
| **PLAN-Start** | Geplantes Start-Datum |
| **PLAN-Ende** | Geplantes End-Datum |
| **System/Mandant-Buchungskreis** | System-Information |

### Optionale Felder

| Spaltenname | Verwendung |
|-------------|------------|
| **IST-Status** | Für Filterung nach Status |
| **Bereich** | Für Filterung nach Bereich |
| **PLAN-Ende** | Wird in E-Mail angezeigt |

## E-Mail-Format

Jede E-Mail wird automatisch wie folgt erstellt:

**An:** {E-Mail-Adresse(n) aus Excel}
**Bcc:** hendrik.max4@dhl.com
**Betreff:** {Cutover-Ident} - {Ident} - {Aktivität}

> **Hinweis**: Bei mehreren Empfängern werden alle im "An:"-Feld aufgeführt.

**Inhalt:**
```
Hallo,

bitte führe die folgende Cutover-Aktivität aus dem Cutover-Plan {Tabellenblatt}:

{Ident} - {Aktivität}

von: {PLAN-Start}
bis: {PLAN-Ende}
im System: {System/Mandant-Buchungskreis}
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status:

abgeschlossen

im Cutoverplan:

{Link Cutover-Plan}


in der o.a. Cutover-Aktivität ein und

sende mir die E-Mail mit "abgeschlossen" am Ende des Betreffs zurück.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus.

Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com
```

## Projektstruktur

```
Script/
├── cutover_mail_generator.py    # Hauptprogramm mit GUI
├── excel_parser.py               # Excel-Verarbeitung
├── email_generator.py            # E-Mail-Erstellung
├── config.py                     # Konfiguration
├── requirements.txt              # Python-Dependencies
└── README.md                     # Diese Datei
```

## Konfiguration anpassen

Um das E-Mail-Template oder andere Einstellungen anzupassen, bearbeiten Sie die Datei `config.py`:

- `BCC_EMAIL`: BCC-Empfänger-Adresse
- `EMAIL_TEMPLATE`: E-Mail-Textvorlage
- `SIGNATURE`: Signatur am Ende der E-Mail
- `EXCEL_COLUMNS`: Spalten-Mapping für Excel

## Fehlerbehebung

### "Outlook nicht verfügbar"
- Stellen Sie sicher, dass Microsoft Outlook installiert ist
- Versuchen Sie, Outlook einmal manuell zu starten

### "Excel-Datei kann nicht gelesen werden"
- Prüfen Sie, ob die Datei im .xlsx-Format vorliegt
- Stellen Sie sicher, dass die Datei nicht geöffnet ist
- Prüfen Sie die Dateiberechtigungen

### "Spalte 'XYZ' fehlt"
- Überprüfen Sie, ob alle Pflichtfelder in der Excel-Datei vorhanden sind
- Achten Sie auf korrekte Schreibweise der Spaltennamen

### "Ungültige E-Mail"
- Prüfen Sie die E-Mail-Adressen in der Excel-Datei
- E-Mails müssen das Format `name@domain.com` haben
- Mehrere E-Mails können mit `;` oder `,` getrennt werden (z.B. `max@dhl.com; anna@dhl.com`)
- Zeilen mit ungültigen E-Mails werden automatisch übersprungen

### "Fehler beim Speichern der EML-Datei"
- Prüfen Sie, ob der Ausgabepfad existiert und beschreibbar ist
- Stellen Sie sicher, dass keine Datei mit gleichem Namen bereits geöffnet ist

## Lizenz

Internes Tool für DHL JOSEF Cutover-Management.

## Kontakt

**Hendrik Max**
Cutover-Manager JOSEF
hendrik.max4@dhl.com

---

## Changelog

### Version 1.3 (2025-12-03)
- ✓ Unterstützung für mehrere E-Mail-Empfänger (getrennt durch `;` oder `,`)
- ✓ Automatische Bereinigung von Zeilenumbrüchen in Aktivitätsbeschreibungen
- ✓ Verbesserte Fehlerbehandlung

### Version 1.2 (2025-11-30)
- ✓ Ident-Trunkierung Fix ("1.10" bleibt "1.10")
- ✓ Link zum Cutover-Plan in E-Mails
- ✓ Tabellenblatt-Name in E-Mail-Text
- ✓ EML-Datei-Export statt MSG

### Version 1.1 (2025-11-23)
- ✓ Grundfunktionalität

---

**Version:** 1.3
**Datum:** 2025-12-03
