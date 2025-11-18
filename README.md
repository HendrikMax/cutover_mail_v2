# Cutover E-Mail Generator

**Automatische Erstellung von Cutover-E-Mails aus Excel-Cutoverplan**

Ein Python-Tool mit GUI zur automatisierten Generierung von personalisierten E-Mails für Cutover-Aktivitäten aus einem Excel-basierten Cutoverplan.

## Features

- 📊 **Excel-Integration**: Liest Cutover-Aktivitäten direkt aus Excel-Dateien
- 🖥️ **Benutzerfreundliche GUI**: Intuitive Oberfläche mit tkinter
- 📧 **Flexible E-Mail-Erstellung**:
  - Outlook-Entwürfe (zur Prüfung vor dem Versand)
  - EML-Dateien (zum späteren Öffnen in jedem E-Mail-Client)
- 🔍 **Filter-Optionen**:
  - Nach IST-Status filtern
  - Nach Bereich filtern
- 📈 **Fortschrittsanzeige**: Live-Updates während der Verarbeitung
- ✅ **Robuste Fehlerbehandlung**: Validierung und aussagekräftige Fehlermeldungen

## Voraussetzungen

- **Python**: Version 3.8 oder höher
- **Microsoft Outlook**: Für Outlook-Entwürfe (optional)
- **Betriebssystem**: Windows (wegen Outlook-Integration)

## Installation

### 1. Repository klonen

```bash
git clone https://github.com/[username]/cutover_mail.git
cd cutover_mail
```

### 2. Virtual Environment erstellen

```bash
cd Script
python -m venv .venv
```

### 3. Virtual Environment aktivieren

```bash
# Windows
.venv\Scripts\activate
```

### 4. Dependencies installieren

```bash
pip install -r requirements.txt
```

## Verwendung

### Programm starten

```bash
cd Script
python cutover_mail_generator.py
```

### Schritt-für-Schritt-Anleitung

1. **Excel-Datei auswählen**
   - Klicken Sie auf "Durchsuchen..." und wählen Sie Ihre Cutoverplan-Datei

2. **Tabellenblatt auswählen**
   - Wählen Sie das Blatt mit den Cutover-Aktivitäten

3. **Cutover-Ident eingeben**
   - Geben Sie die Identifikation ein (z.B. "JOSEF", "DPN_ECH")

4. **E-Mail-Modus wählen**
   - **Outlook-Entwürfe**: E-Mails werden zum Prüfen geöffnet
   - **E-Mail-Dateien (.eml)**: Werden im gewählten Ordner gespeichert

5. **Filter setzen** (optional)
   - IST-Status: z.B. nur "offen"
   - Bereich: Spezifischen Bereich auswählen

6. **E-Mails generieren**
   - Klicken Sie auf "E-Mails generieren"
   - Verfolgen Sie den Fortschritt im Log-Fenster

## Excel-Struktur

### Erforderliche Spalten

| Spaltenname | Beschreibung | Pflicht |
|-------------|--------------|---------|
| **Ident** | Eindeutige Aktivitäts-ID | Ja |
| **Aktivität** | Beschreibung der Aktivität | Ja |
| **E-Mail** | E-Mail-Adresse des Ausführenden | Ja |
| **PLAN-Start** | Geplantes Start-Datum | Ja |
| **System/Mandant-Buchungskreis** | System-Information | Ja |
| **IST-Status** | Status für Filterung | Nein |
| **Bereich** | Bereich für Filterung | Nein |

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
cutover_mail/
├── Konzept/
│   ├── Konzept Cutover_E-Mails.md
│   └── Umsetzungskonzept Cutover_E-Mails.md
├── Input_Datei/
│   └── (Excel-Dateien hier ablegen)
├── Script/
│   ├── cutover_mail_generator.py    # Hauptprogramm
│   ├── excel_parser.py               # Excel-Verarbeitung
│   ├── email_generator.py            # E-Mail-Erstellung
│   ├── config.py                     # Konfiguration
│   ├── requirements.txt              # Dependencies
│   └── README.md                     # Dokumentation
├── .gitignore
└── README.md                         # Diese Datei
```

## Konfiguration

Um das E-Mail-Template oder andere Einstellungen anzupassen, bearbeiten Sie `Script/config.py`:

- `BCC_EMAIL`: BCC-Empfänger-Adresse
- `EMAIL_TEMPLATE`: E-Mail-Textvorlage
- `SIGNATURE`: Signatur am Ende der E-Mail
- `EXCEL_COLUMNS`: Spalten-Mapping für Excel

## Fehlerbehebung

### "Excel-Datei kann nicht gelesen werden"
- Prüfen Sie, ob die Datei im .xlsx-Format vorliegt
- Stellen Sie sicher, dass die Datei nicht geöffnet ist

### "Spalte 'XYZ' fehlt"
- Überprüfen Sie, ob alle Pflichtfelder in der Excel-Datei vorhanden sind
- Die Header-Zeile muss in Zeile 3 sein

### "Ungültige E-Mail"
- E-Mails müssen das Format `name@domain.com` haben
- Zeilen mit ungültigen E-Mails werden automatisch übersprungen

### "Outlook nicht verfügbar"
- Nur für Outlook-Entwürfe-Modus relevant
- Verwenden Sie alternativ den EML-Dateien-Modus

## Technische Details

- **GUI**: tkinter (Standard Python)
- **Excel**: openpyxl, pandas
- **Outlook**: pywin32 (nur für Outlook-Entwürfe)
- **E-Mail**: Python email-Bibliothek (für EML-Dateien)

## Lizenz

Internes Tool für DHL JOSEF Cutover-Management.

## Autor

**Hendrik Max**
Cutover-Manager JOSEF
hendrik.max4@dhl.com

## Version

1.0 - Initial Release (2025-11-18)

---

🤖 Generated with [Claude Code](https://claude.com/claude-code)
