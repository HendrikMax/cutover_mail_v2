"""
E-Mail-Generator für Cutover-Aktivitäten
Erstellt E-Mails als Outlook-Entwürfe oder EML-Dateien
"""

import win32com.client
from pathlib import Path
from typing import Dict, List, Optional, Callable
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
import config


def create_email_subject(activity: Dict[str, str], cutover_ident: str) -> str:
    """
    Erstellt E-Mail-Betreff.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Returns:
        Formatierter Betreff (gekürzt bei Bedarf)
    """
    ident = activity['ident']
    aktivitaet = activity['aktivitaet']

    # Betreff zusammensetzen
    subject = f"{cutover_ident} - {ident} - {aktivitaet}"

    # Optional: Betreff kürzen, falls zu lang (Outlook-Limit ~255 Zeichen)
    max_length = 200
    if len(subject) > max_length:
        subject = subject[:max_length - 3] + "..."

    return subject


def create_email_body(activity: Dict[str, str], cutover_ident: str) -> str:
    """
    Erstellt E-Mail-Inhalt aus Template.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Returns:
        Formatierter E-Mail-Text
    """
    body = config.EMAIL_TEMPLATE.format(
        ident=activity['ident'],
        aktivitaet=activity['aktivitaet'],
        plan_start=activity['plan_start'],
        system=activity['system'],
        cutover_ident=cutover_ident,
        signature=config.SIGNATURE
    )

    return body


def create_outlook_draft(activity: Dict[str, str], cutover_ident: str) -> None:
    """
    Erstellt E-Mail-Entwurf in Outlook.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation

    Raises:
        Exception: Outlook nicht verfügbar oder Fehler bei Erstellung
    """
    try:
        # Outlook-Instanz erstellen
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem

        # E-Mail-Eigenschaften setzen
        mail.To = activity['email']
        mail.BCC = config.BCC_EMAIL
        mail.Subject = create_email_subject(activity, cutover_ident)
        mail.Body = create_email_body(activity, cutover_ident)

        # Als Entwurf anzeigen (nicht senden!)
        mail.Display(False)

    except Exception as e:
        raise Exception(f"Fehler beim Erstellen der Outlook-E-Mail: {e}")


def save_as_eml(
    activity: Dict[str, str],
    cutover_ident: str,
    output_path: str
) -> str:
    """
    Speichert E-Mail als EML-Datei (Standard E-Mail-Format).

    EML-Dateien können in Outlook und allen anderen E-Mail-Clients
    geöffnet werden. Diese Methode ist stabil und benötigt kein COM.

    Args:
        activity: Dictionary mit Aktivitätsdaten
        cutover_ident: Cutover-Identifikation
        output_path: Pfad zum Ausgabeordner

    Returns:
        Pfad zur erstellten EML-Datei

    Raises:
        IOError: Speichern fehlgeschlagen
    """
    try:
        # Dateiname erstellen (sicher für Dateisystem)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        ident_safe = activity['ident']
        for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '.', ' ']:
            ident_safe = ident_safe.replace(char, '_')
        ident_safe = ident_safe.strip()[:30]

        filename = f"{config.MSG_FILE_PREFIX}_{cutover_ident}_{ident_safe}_{timestamp}.eml"

        # Vollständiger Pfad
        output_dir = Path(output_path)
        output_dir.mkdir(parents=True, exist_ok=True)
        file_path = output_dir / filename

        # E-Mail erstellen mit Python email-Bibliothek
        msg = MIMEMultipart()
        msg['From'] = config.BCC_EMAIL
        msg['To'] = activity['email']
        msg['Bcc'] = config.BCC_EMAIL
        msg['Subject'] = create_email_subject(activity, cutover_ident)
        msg['Date'] = formatdate(localtime=True)

        # E-Mail-Body
        body = create_email_body(activity, cutover_ident)
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # Als EML-Datei speichern
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(msg.as_string())

        return str(file_path)

    except Exception as e:
        raise IOError(f"Fehler beim Speichern der EML-Datei: {e}")


def generate_emails(
    activities: List[Dict[str, str]],
    cutover_ident: str,
    mode: str,
    output_path: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> Dict[str, int]:
    """
    Generiert E-Mails für alle Aktivitäten.

    Args:
        activities: Liste von Aktivitäten
        cutover_ident: Cutover-Identifikation
        mode: 'outlook' oder 'msg'
        output_path: Optional - Pfad für .msg-Dateien (erforderlich bei mode='msg')
        progress_callback: Optional - Callback-Funktion mit Signatur:
                          callback(current: int, total: int, message: str)

    Returns:
        Dictionary mit Statistiken:
        {
            'erfolg': int,    # Anzahl erfolgreich erstellter E-Mails
            'fehler': int,    # Anzahl fehlgeschlagener E-Mails
            'gesamt': int     # Gesamtanzahl
        }

    Raises:
        ValueError: Ungültiger Modus oder fehlender output_path
    """
    if mode not in ['outlook', 'msg']:
        raise ValueError(f"Ungültiger Modus: {mode}. Erlaubt: 'outlook' oder 'msg'")

    if mode == 'msg' and not output_path:
        raise ValueError("output_path ist erforderlich für Modus 'msg'")

    total = len(activities)
    erfolg = 0
    fehler = 0
    fehler_liste = []

    for i, activity in enumerate(activities, 1):
        try:
            if mode == 'outlook':
                create_outlook_draft(activity, cutover_ident)
                message = f"E-Mail erstellt: {activity['ident']} - {activity['email']}"
            else:  # mode == 'msg' (verwendet jetzt EML-Format)
                # Speichere als EML-Datei (stabiler als MSG)
                if progress_callback:
                    progress_callback(i, total, f"Speichere EML-Datei: {activity['ident']}...")

                file_path = save_as_eml(activity, cutover_ident, output_path)
                message = f"EML-Datei gespeichert: {activity['ident']} -> {Path(file_path).name}"

            erfolg += 1

            # Progress-Callback aufrufen
            if progress_callback:
                progress_callback(i, total, message)

        except Exception as e:
            fehler += 1
            import traceback
            fehler_detail = traceback.format_exc()
            fehler_liste.append({
                'ident': activity['ident'],
                'email': activity['email'],
                'fehler': str(e),
                'detail': fehler_detail
            })
            message = f"FEHLER bei {activity['ident']}: {e}"

            # Progress-Callback auch bei Fehler aufrufen
            if progress_callback:
                progress_callback(i, total, message)

    # Statistiken zurückgeben
    stats = {
        'erfolg': erfolg,
        'fehler': fehler,
        'gesamt': total,
        'fehler_liste': fehler_liste
    }

    return stats


def test_outlook_connection() -> bool:
    """
    Testet, ob Outlook verfügbar ist.

    Returns:
        True wenn Outlook verfügbar, False sonst
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return True
    except:
        return False
