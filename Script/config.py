"""
Konfigurationsdatei für Cutover E-Mail Generator
Enthält alle Konstanten, Templates und Spalten-Mappings
"""

# E-Mail-Konfiguration
BCC_EMAIL = "hendrik.max4@dhl.com"

SIGNATURE = """Beste Grüße
Hendrik

Hendrik Max
Cutover-Manager JOSEF
hendrik.max4@dhl.com"""

# E-Mail-Template
EMAIL_TEMPLATE = """Hallo,

bitte führe die folgende Cutover-Aktivität aus dem Cutover-Plan {sheet_name}:

{ident} - {aktivitaet}

von: {plan_start}
bis: {plan_ende}
im System: {system}
aus.

Bitte trage nach Ausführung der Cutover-Aktivität den Status:

abgeschlossen

im Cutoverplan:

{cutover_plan_link}


in der o.a. Cutover-Aktivität ein und

sende mir die E-Mail mit "abgeschlossen" am Ende des Betreffs zurück.

Für Rückfragen stehe ich Dir sehr gern zur Verfügung.

Vielen Dank im Voraus und viel Erfolg!

{signature}"""

# Excel-Spalten-Mapping (exakte Namen aus Excel-Datei)
EXCEL_COLUMNS = {
    'bereich': 'Bereich',
    'ident': 'Ident',
    'vorgaenger': 'Vorgänger',
    'plan_start': 'PLAN-Start',
    'plan_ende': 'PLAN-Ende',
    'plan_dauer': 'PLAN-Dauer (hh:mm)',
    'system': 'System/Mandant-Buchungskreis',
    'buchungsperiode': 'Buchungsperiode (MM/JJJJ)',
    'buchungsdatum': 'Buchungsdatum\n(TT.MM.JJJJ)',
    'aktivitaet': 'Aktivität',
    'technische_info': 'technische Informationen',
    'ausfuehrung_durch': 'Ausführung  durch',
    'email': 'E-Mail',
    'ist_status': 'IST-Status',
    'ist_start': 'IST-Start',
    'ist_ende': 'IST-Ende',
    'ist_dauer': 'IST-Dauer (dd:hh:mm)',
    'dokumentation': 'unterstützende Dokumentation \n',
    'bemerkungen': 'Bemerkungen',
    'link': 'Link auf CuOv-Plan'
}

# Pflichtfelder für Validierung
REQUIRED_COLUMNS = [
    'Ident',
    'Aktivität',
    'E-Mail',
    'PLAN-Start',
    'System/Mandant-Buchungskreis'
]

# GUI-Konfiguration
GUI_TITLE = "Cutover E-Mail Generator"
GUI_WIDTH = 700
GUI_HEIGHT = 650

# Ausgabe-Konfiguration
MSG_FILE_PREFIX = "CutoverMail"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# Filter-Optionen
FILTER_STATUS_EMPTY = "Nur Aktivitäten mit leerem IST-Status"
