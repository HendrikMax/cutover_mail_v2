# Cutover E-Mail Generator

Automatische Erstellung von Cutover-E-Mails aus Excel-Cutoverplan für Microsoft Outlook.

## Überblick

Dieses Python-Tool liest einen Excel-basierten Cutoverplan und erstellt automatisch personalisierte E-Mails für jede Cutover-Aktivität. Die E-Mails können entweder als Outlook-Entwürfe erstellt oder als .msg-Dateien gespeichert werden.

## Features

- **Excel-Integration**: Liest Cutover-Aktivitäten direkt aus Excel-Dateien
- **GUI-Oberfläche**: Benutzerfreundliche grafische Oberfläche mit tkinter
- **Flexible E-Mail-Erstellung**:
  - Outlook-Entwürfe (zur Prüfung vor dem Versand)
  - .msg-Dateien (zum späteren Öffnen)
- **Filter-Optionen**:
  - Nach IST-Status filtern
  - Nach Bereich filtern
- **Fortschrittsanzeige**: Live-Updates während der Verarbeitung
- **Fehlerbehandlung**: Robuste Validierung und aussagekräftige Fehlermeldungen

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

4. **E-Mail-Modus wählen**
   - **Outlook-Entwürfe erstellen**: E-Mails werden in Outlook als Entwürfe geöffnet
   - **Als .msg-Dateien speichern**: E-Mails werden als Dateien gespeichert

5. **Ausgabepfad wählen** (nur bei .msg-Modus)
   - Klicken Sie auf "Durchsuchen..." und wählen Sie einen Ordner

6. **Filter setzen** (optional)
   - ☑ "Nur Aktivitäten mit leerem IST-Status": Nur unerledigte Aktivitäten
   - "Bereich filtern": Nur Aktivitäten eines bestimmten Bereichs

7. **E-Mails generieren**
   - Klicken Sie auf "E-Mails generieren"
   - Der Fortschritt wird angezeigt
   - Bei Outlook-Modus: Die E-Mail-Entwürfe öffnen sich automatisch

## Excel-Struktur

Das Tool erwartet folgende Spalten in der Excel-Datei:

### Pflichtfelder

| Spaltenname | Beschreibung |
|-------------|--------------|
| **Ident** | Eindeutige Aktivitäts-ID |
| **Aktivität** | Beschreibung der Aktivität |
| **E-Mail** | E-Mail-Adresse des Ausführenden |
| **PLAN-Start** | Geplantes Start-Datum |
| **System/Mandant-Buchungskreis** | System-Information |

### Optionale Felder

| Spaltenname | Verwendung |
|-------------|------------|
| **IST-Status** | Für Filterung nach Status |
| **Bereich** | Für Filterung nach Bereich |

## E-Mail-Format

Jede E-Mail wird automatisch wie folgt erstellt:

**An:** {E-Mail-Adresse aus Excel}
**Bcc:** hendrik.max4@dhl.com
**Betreff:** {Cutover-Ident} - {Ident} - {Aktivität}

**Inhalt:**
```
Hallo,

bitte führe die folgende Cutover-Aktivität
{Ident} - {Aktivität}
am: {PLAN-Start}
im System: {System/Mandant-Buchungskreis}
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status im
Cutoverplan {Cutover-Ident} in der o.a. Cutover-Aktivität ein.

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
- Zeilen mit ungültigen E-Mails werden automatisch übersprungen

## Lizenz

Internes Tool für DHL JOSEF Cutover-Management.

## Kontakt

**Hendrik Max**
Cutover-Manager JOSEF
hendrik.max4@dhl.com

---

**Version:** 1.0
**Datum:** 2025-11-18
