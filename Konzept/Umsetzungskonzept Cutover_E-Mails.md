# Umsetzungskonzept: Cutover E-Mail Generator

**Projekt:** Automatische Erstellung von Cutover-E-Mails aus Excel-Cutoverplan
**Datum:** 2025-11-18
**Version:** 1.0

---

## 1. Übersicht

Dieses Dokument beschreibt die technische Umsetzung eines Python-Scripts, das aus einem Excel-basierten Cutoverplan automatisch E-Mails für Microsoft Outlook generiert. Für jede Cutover-Aktivität wird eine personalisierte E-Mail an den zuständigen Ausführenden erstellt.

---

## 2. Projektstruktur

```
cutover_mail/
├── Konzept/
│   ├── Konzept Cutover_E-Mails.md
│   └── Umsetzungskonzept Cutover_E-Mails.md (diese Datei)
├── Input_Datei/
│   └── DHL_JOSEF CuOvPlan DPN_ECH V01 20250813 DRAFT.xlsx
└── Script/
    ├── cutover_mail_generator.py    # Hauptprogramm mit GUI
    ├── excel_parser.py               # Excel-Verarbeitung
    ├── email_generator.py            # E-Mail-Erstellung
    ├── config.py                     # Konfiguration & Konstanten
    └── requirements.txt              # Python-Dependencies
```

---

## 3. Technologie-Stack

| Komponente | Technologie | Begründung |
|------------|-------------|------------|
| **GUI** | tkinter | Standard in Python enthalten, keine zusätzlichen Dependencies |
| **Excel-Verarbeitung** | openpyxl | Robuste Bibliothek für .xlsx-Dateien |
| **Outlook-Integration** | pywin32 (win32com.client) | Direkte COM-Schnittstelle zu Outlook |
| **Datenverarbeitung** | pandas (optional) | Für erweiterte Filterung und Datenmanipulation |

---

## 4. Excel-Struktur (Input)

### 4.1 Relevante Spalten

Basierend auf der Beispieldatei `DHL_JOSEF CuOvPlan DPN_ECH V01 20250813 DRAFT.xlsx`:

| Spaltenname | Verwendung | Pflichtfeld |
|-------------|------------|-------------|
| **Ident** | Aktivitäts-ID für Betreff | Ja |
| **Aktivität** | Beschreibung der Aktivität | Ja |
| **E-Mail** | Empfänger-Adresse | Ja |
| **PLAN-Start** | Geplantes Start-Datum | Ja |
| **System/Mandant-Buchungskreis** | System-Information | Ja |
| **IST-Status** | Für optionale Filterung | Nein |
| **Bereich** | Für optionale Filterung | Nein |

### 4.2 Alle verfügbaren Spalten
```
Bereich | Ident | Vorgänger | PLAN-Start | PLAN-Ende | PLAN-Dauer (hh:mm) |
System/Mandant-Buchungskreis | Buchungsperiode (MM/JJJJ) | Buchungsdatum (TT.MM.JJJJ) |
Aktivität | technische Informationen | Ausführung durch | E-Mail | IST-Status |
IST-Start | IST-Ende | IST-Dauer (dd:hh:mm) | unterstützende Dokumentation |
Bemerkungen | Link auf CuOv-Plan
```

---

## 5. E-Mail-Struktur (Output)

### 5.1 E-Mail-Template

**An:** `{Spalte: E-Mail}`
**Bcc:** `hendrik.max4@dhl.com`
**Betreff:** `{Cutover-Ident} - {Ident} - {Aktivität}`

**Inhalt:**
```
Hallo,

bitte führe die folgende Cutover-Aktivität
{Ident} - {Aktivität}
am: {PLAN-Start}
im System: {System/Mandant-Buchungskreis}
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status im Cutoverplan {Cutover-Ident} in der o.a. Cutover-Aktivität ein.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus.

Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com
```

### 5.2 Beispiel-E-Mail

**An:** max.mustermann@dhl.com
**Bcc:** hendrik.max4@dhl.com
**Betreff:** JOSEF - A001 - Buchungsperiode öffnen in SAP FI

**Inhalt:**
```
Hallo,

bitte führe die folgende Cutover-Aktivität
A001 - Buchungsperiode öffnen in SAP FI
am: 13.08.2025 08:00
im System: SAP PRD/1000
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status im Cutoverplan JOSEF in der o.a. Cutover-Aktivität ein.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus.

Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com
```

---

## 6. GUI-Spezifikation

### 6.1 Eingabe-Elemente

| Element | Typ | Beschreibung |
|---------|-----|--------------|
| **Excel-Dateiauswahl** | Button + Label | Öffnet Datei-Dialog zur Auswahl der .xlsx-Datei |
| **Tabellenblatt** | Dropdown | Zeigt alle verfügbaren Tabellenblätter |
| **Cutover-Ident** | Eingabefeld | Freitext (z.B. "JOSEF", "DPN_ECH") |
| **E-Mail-Modus** | Radio-Buttons | "Outlook-Entwürfe erstellen" / "Als .msg-Dateien speichern" |
| **Filter-Optionen** | Checkboxen/Dropdown | Optional: Nach IST-Status oder Bereich filtern |
| **Ausgabepfad** | Button + Label | (nur bei .msg-Modus) Ordner für gespeicherte E-Mails |

### 6.2 Aktions-Elemente

| Element | Beschreibung |
|---------|--------------|
| **"E-Mails generieren"** | Startet die Verarbeitung |
| **Fortschrittsbalken** | Zeigt Anzahl verarbeiteter Aktivitäten |
| **Log-Ausgabe** | Textfeld mit Verarbeitungsprotokoll |

### 6.3 GUI-Layout (Mockup)

```
┌─────────────────────────────────────────────────────────┐
│  Cutover E-Mail Generator                               │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  Excel-Datei:  [____________________________] [Durchsuchen] │
│                                                         │
│  Tabellenblatt: [Dropdown ▼]                           │
│                                                         │
│  Cutover-Ident: [____________]  (z.B. "JOSEF")         │
│                                                         │
│  E-Mail-Modus:                                         │
│    ○ Outlook-Entwürfe erstellen                        │
│    ○ Als .msg-Dateien speichern                        │
│                                                         │
│  Ausgabepfad:  [____________________________] [Durchsuchen] │
│    (nur aktiv bei .msg-Modus)                          │
│                                                         │
│  Filter-Optionen:                                      │
│    □ Nur Aktivitäten mit leerem IST-Status             │
│    Bereich filtern: [Alle ▼]                           │
│                                                         │
│  [     E-Mails generieren     ]                        │
│                                                         │
│  Fortschritt: [████████░░░░░░░░░░] 8/20               │
│                                                         │
│  ┌─────────────────────────────────────────────┐       │
│  │ Log:                                        │       │
│  │ Excel-Datei erfolgreich geladen...          │       │
│  │ 20 Aktivitäten gefunden.                   │       │
│  │ E-Mail für A001 erstellt ✓                 │       │
│  │ E-Mail für A002 erstellt ✓                 │       │
│  └─────────────────────────────────────────────┘       │
└─────────────────────────────────────────────────────────┘
```

---

## 7. Modulbeschreibungen

### 7.1 config.py

**Zweck:** Zentrale Konfigurationsdatei für alle Konstanten

**Inhalt:**
```python
# E-Mail-Konfiguration
BCC_EMAIL = "hendrik.max4@dhl.com"
SIGNATURE = """Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com"""

# E-Mail-Template
EMAIL_TEMPLATE = """Hallo,

bitte führe die folgende Cutover-Aktivität
{ident} - {aktivitaet}
am: {plan_start}
im System: {system}
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status im Cutoverplan {cutover_ident} in der o.a. Cutover-Aktivität ein.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus.

{signature}"""

# Excel-Spalten-Mapping
EXCEL_COLUMNS = {
    'ident': 'Ident',
    'aktivitaet': 'Aktivität',
    'email': 'E-Mail',
    'plan_start': 'PLAN-Start',
    'system': 'System/Mandant-Buchungskreis',
    'ist_status': 'IST-Status',
    'bereich': 'Bereich'
}

# Validierung
REQUIRED_COLUMNS = ['Ident', 'Aktivität', 'E-Mail', 'PLAN-Start', 'System/Mandant-Buchungskreis']
```

---

### 7.2 excel_parser.py

**Zweck:** Excel-Datei einlesen und Aktivitäten extrahieren

**Hauptfunktionen:**

```python
def get_sheet_names(file_path: str) -> list:
    """
    Gibt alle Tabellenblatt-Namen der Excel-Datei zurück.

    Args:
        file_path: Pfad zur Excel-Datei

    Returns:
        Liste der Tabellenblatt-Namen
    """

def load_activities(file_path: str, sheet_name: str, filters: dict = None) -> list:
    """
    Lädt Aktivitäten aus Excel-Datei.

    Args:
        file_path: Pfad zur Excel-Datei
        sheet_name: Name des Tabellenblatts
        filters: Optional - Dictionary mit Filterkriterien
                 z.B. {'IST-Status': '', 'Bereich': 'SAP'}

    Returns:
        Liste von Dictionaries mit Aktivitätsdaten

    Raises:
        FileNotFoundError: Datei nicht gefunden
        ValueError: Pflichtfelder fehlen
    """

def validate_columns(df) -> bool:
    """
    Prüft, ob alle Pflichtfelder vorhanden sind.

    Args:
        df: pandas DataFrame

    Returns:
        True wenn alle Pflichtfelder vorhanden

    Raises:
        ValueError: Wenn Pflichtfelder fehlen
    """

def validate_email(email: str) -> bool:
    """
    Validiert E-Mail-Format.

    Args:
        email: E-Mail-Adresse

    Returns:
        True wenn gültig
    """
```

**Implementierungsdetails:**
- Verwendet `openpyxl` oder `pandas` für Excel-Verarbeitung
- Startet Daten-Lesen ab Zeile 2 (Zeile 1 = Header)
- Filtert leere Zeilen heraus
- Konvertiert Datum-Felder in lesbares Format
- Validiert E-Mail-Adressen mit regex

---

### 7.3 email_generator.py

**Zweck:** E-Mails in Outlook erstellen oder als .msg-Dateien speichern

**Hauptfunktionen:**

```python
def create_email_body(activity: dict, cutover_ident: str) -> str:
    """
    Erstellt E-Mail-Inhalt aus Template.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Returns:
        Formatierter E-Mail-Text
    """

def create_email_subject(activity: dict, cutover_ident: str) -> str:
    """
    Erstellt E-Mail-Betreff.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Returns:
        Formatierter Betreff (gekürzt bei Bedarf)
    """

def create_outlook_draft(activity: dict, cutover_ident: str):
    """
    Erstellt E-Mail-Entwurf in Outlook.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Raises:
        Exception: Outlook nicht verfügbar
    """

def save_as_msg(activity: dict, cutover_ident: str, output_path: str):
    """
    Speichert E-Mail als .msg-Datei.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation
        output_path: Pfad zum Ausgabeordner

    Raises:
        IOError: Speichern fehlgeschlagen
    """

def generate_emails(activities: list, cutover_ident: str, mode: str,
                   output_path: str = None, progress_callback=None):
    """
    Generiert E-Mails für alle Aktivitäten.

    Args:
        activities: Liste von Aktivitäten
        cutover_ident: Cutover-Identifikation
        mode: 'outlook' oder 'msg'
        output_path: Optional - Pfad für .msg-Dateien
        progress_callback: Optional - Callback für Fortschritt

    Returns:
        Dictionary mit Statistiken (erfolg, fehler)
    """
```

**Implementierungsdetails:**
- Verwendet `win32com.client` für Outlook-Integration
- Dateiname für .msg: `{Cutover-Ident}_{Ident}_{Timestamp}.msg`
- Fehlerbehandlung für ungültige E-Mail-Adressen
- Fortschritts-Callback für GUI-Update

---

### 7.4 cutover_mail_generator.py

**Zweck:** Hauptprogramm mit GUI

**Hauptklasse:**

```python
class CutoverMailGeneratorGUI:
    """
    Hauptfenster der Anwendung.
    """

    def __init__(self):
        """Initialisiert GUI-Elemente."""

    def browse_excel_file(self):
        """Öffnet Datei-Dialog für Excel-Auswahl."""

    def load_sheet_names(self):
        """Lädt Tabellenblatt-Namen aus gewählter Datei."""

    def browse_output_folder(self):
        """Öffnet Ordner-Dialog für Ausgabepfad."""

    def toggle_output_path(self):
        """Aktiviert/deaktiviert Ausgabepfad je nach Modus."""

    def validate_inputs(self) -> bool:
        """Validiert alle Eingabefelder."""

    def generate_emails(self):
        """Startet E-Mail-Generierung."""

    def update_progress(self, current: int, total: int):
        """Aktualisiert Fortschrittsbalken."""

    def log_message(self, message: str):
        """Fügt Nachricht zum Log hinzu."""

    def run(self):
        """Startet GUI-Hauptschleife."""
```

**Implementierungsdetails:**
- Threading für E-Mail-Generierung (GUI bleibt responsiv)
- Automatische Validierung bei Eingaben
- Fehler-Dialoge mit `messagebox`
- Tooltip-Hilfe für komplexe Felder

---

## 8. Implementierungsreihenfolge

### Phase 1: Grundgerüst (Priorität: Hoch)
1. **requirements.txt** erstellen
   - Alle Dependencies auflisten
   - Versions-Pins setzen

2. **config.py** implementieren
   - Alle Konstanten definieren
   - Template-Strings anlegen

### Phase 2: Core-Funktionalität (Priorität: Hoch)
3. **excel_parser.py** implementieren
   - Excel-Datei einlesen
   - Spalten validieren
   - Filterlogik umsetzen
   - Unit-Tests schreiben

4. **email_generator.py** implementieren
   - Template-Engine
   - Outlook-Integration (win32com)
   - MSG-Export
   - Error-Handling

### Phase 3: GUI (Priorität: Mittel)
5. **cutover_mail_generator.py** implementieren
   - GUI-Layout erstellen
   - Event-Handler verbinden
   - Threading implementieren
   - Fortschrittsanzeige

### Phase 4: Testing & Dokumentation (Priorität: Mittel)
6. **Integration-Testing**
   - Test mit Beispiel-Excel
   - Verschiedene Szenarien testen
   - Fehlerbehandlung prüfen

7. **README.md** erstellen
   - Installation
   - Nutzungsanleitung
   - Screenshots

---

## 9. Fehlerbehandlung

### 9.1 Validierungen

| Prüfung | Fehlermeldung | Aktion |
|---------|---------------|--------|
| Excel-Datei existiert | "Datei nicht gefunden" | Dialog anzeigen |
| Pflichtfelder vorhanden | "Spalte '{name}' fehlt" | Abbruch mit Hinweis |
| E-Mail-Format gültig | "Ungültige E-Mail: {email}" | Zeile überspringen, loggen |
| Outlook verfügbar | "Outlook nicht installiert" | Modus wechseln vorschlagen |
| Schreibrechte Ausgabepfad | "Keine Berechtigung" | Anderen Pfad wählen |

### 9.2 Exception-Handling

```python
try:
    # Excel-Verarbeitung
    activities = excel_parser.load_activities(file_path, sheet_name)
except FileNotFoundError:
    messagebox.showerror("Fehler", "Excel-Datei nicht gefunden")
except ValueError as e:
    messagebox.showerror("Fehler", f"Excel-Validierung: {e}")
except Exception as e:
    messagebox.showerror("Fehler", f"Unerwarteter Fehler: {e}")
    logging.exception("Fehler bei Excel-Verarbeitung")
```

---

## 10. Testing-Strategie

### 10.1 Unit-Tests

- **excel_parser.py**
  - Test: Korrekte Spalten erkannt
  - Test: Filter funktioniert
  - Test: Leere Zeilen ignoriert

- **email_generator.py**
  - Test: Template korrekt gefüllt
  - Test: Betreff gekürzt bei Überlänge
  - Test: E-Mail-Validierung

### 10.2 Integration-Tests

- **End-to-End-Test**
  - Excel laden → Aktivitäten filtern → E-Mails generieren
  - Verschiedene Modi (Outlook / MSG)
  - Mit verschiedenen Filter-Kombinationen

### 10.3 Manuelle Tests

- GUI-Bedienbarkeit
- Outlook-Integration
- Performance (100+ Aktivitäten)

---

## 11. Erweiterungsmöglichkeiten (Optional)

### 11.1 Kurzfristig
- **Excel-Export** der verarbeiteten Aktivitäten
- **Vorschau** einer Beispiel-E-Mail vor Generierung
- **Batch-Verarbeitung** mehrerer Cutoverplan-Dateien

### 11.2 Mittelfristig
- **E-Mail-Versand** direkt aus dem Tool
- **Status-Tracking** (E-Mails versendet markieren)
- **Template-Editor** in GUI

### 11.3 Langfristig
- **Web-Interface** statt Desktop-GUI
- **Datenbank-Integration** für Cutoverplan
- **Automatisierung** (zeitgesteuerter Versand)

---

## 12. Installation & Deployment

### 12.1 Voraussetzungen

- Python 3.8+
- Microsoft Outlook installiert (für Outlook-Modus)
- Windows (wegen win32com)

### 12.2 Installation

```bash
# 1. Repository klonen / Dateien kopieren
cd C:\Users\hendrik.max\Documents\DEV_LOCL\cutover_mail\Script

# 2. Virtual Environment erstellen
python -m venv venv

# 3. Virtual Environment aktivieren
venv\Scripts\activate

# 4. Dependencies installieren
pip install -r requirements.txt
```

### 12.3 Start

```bash
python cutover_mail_generator.py
```

---

## 13. Anhang

### 13.1 requirements.txt (Vorschau)

```txt
openpyxl==3.1.2
pywin32==306
pandas==2.2.0
python-dateutil==2.8.2
```

### 13.2 Beispiel-Ausgabe (Log)

```
[2025-11-18 14:30:15] Excel-Datei geladen: DHL_JOSEF CuOvPlan DPN_ECH V01 20250813 DRAFT.xlsx
[2025-11-18 14:30:15] Tabellenblatt: Cutover-Plan
[2025-11-18 14:30:15] 45 Aktivitäten gefunden
[2025-11-18 14:30:15] Filter angewendet: IST-Status = leer
[2025-11-18 14:30:15] 23 Aktivitäten nach Filterung
[2025-11-18 14:30:16] E-Mail erstellt: A001 - max.mustermann@dhl.com ✓
[2025-11-18 14:30:16] E-Mail erstellt: A002 - anna.schmidt@dhl.com ✓
[2025-11-18 14:30:17] Warnung: Ungültige E-Mail bei A005, übersprungen
[2025-11-18 14:30:18] E-Mail erstellt: A006 - tom.mueller@dhl.com ✓
...
[2025-11-18 14:30:45] Fertig! 22 E-Mails erfolgreich erstellt, 1 Fehler
```

---

## 14. Kontakt & Support

**Projekt-Owner:** Hendrik Max
**E-Mail:** hendrik.max4@dhl.com
**Rolle:** Cutover-Manager JOSEF

---

*Ende des Umsetzungskonzepts*
