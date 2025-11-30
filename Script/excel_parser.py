"""
Excel-Parser für Cutover-Pläne
Lädt und validiert Excel-Dateien mit Cutover-Aktivitäten
"""

import pandas as pd
import re
from typing import List, Dict, Optional
from pathlib import Path
import config


def get_sheet_names(file_path: str) -> List[str]:
    """
    Gibt alle Tabellenblatt-Namen der Excel-Datei zurück.

    Args:
        file_path: Pfad zur Excel-Datei

    Returns:
        Liste der Tabellenblatt-Namen

    Raises:
        FileNotFoundError: Wenn die Datei nicht gefunden wird
        ValueError: Wenn die Datei keine gültige Excel-Datei ist
    """
    if not Path(file_path).exists():
        raise FileNotFoundError(f"Excel-Datei nicht gefunden: {file_path}")

    try:
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        return excel_file.sheet_names
    except Exception as e:
        raise ValueError(f"Fehler beim Lesen der Excel-Datei: {e}")


def validate_columns(df: pd.DataFrame) -> None:
    """
    Prüft, ob alle Pflichtfelder vorhanden sind.

    Args:
        df: pandas DataFrame

    Raises:
        ValueError: Wenn Pflichtfelder fehlen
    """
    missing_columns = []
    for required_col in config.REQUIRED_COLUMNS:
        if required_col not in df.columns:
            missing_columns.append(required_col)

    if missing_columns:
        raise ValueError(
            f"Folgende Pflichtfelder fehlen in der Excel-Datei: {', '.join(missing_columns)}"
        )


def validate_email(email: str) -> bool:
    """
    Validiert E-Mail-Format.

    Args:
        email: E-Mail-Adresse

    Returns:
        True wenn gültig, False sonst
    """
    if pd.isna(email) or not email:
        return False

    # Einfache E-Mail-Validierung mit regex
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, str(email).strip()))


def format_date(date_value) -> str:
    """
    Formatiert Datumswerte in lesbares Format.

    Args:
        date_value: Datumswert (kann verschiedene Typen haben)

    Returns:
        Formatierter Datumsstring
    """
    if pd.isna(date_value):
        return ""

    # Wenn es bereits ein String ist
    if isinstance(date_value, str):
        return date_value.strip()

    # Wenn es ein datetime-Objekt ist
    try:
        if hasattr(date_value, 'strftime'):
            return date_value.strftime('%d.%m.%Y %H:%M')
        else:
            return str(date_value)
    except:
        return str(date_value)


def clean_value(value) -> str:
    """
    Bereinigt Werte (entfernt NaN, None, etc.).

    Args:
        value: Zu bereinigender Wert

    Returns:
        Bereinigter String
    """
    if pd.isna(value):
        return ""
    return str(value).strip()


def load_activities(
    file_path: str,
    sheet_name: str,
    filters: Optional[Dict[str, str]] = None
) -> List[Dict[str, str]]:
    """
    Lädt Aktivitäten aus Excel-Datei.

    Args:
        file_path: Pfad zur Excel-Datei
        sheet_name: Name des Tabellenblatts
        filters: Optional - Dictionary mit Filterkriterien
                 z.B. {'IST-Status': '', 'Bereich': 'SAP'}
                 Leerer String bedeutet: "Feld muss leer sein"

    Returns:
        Liste von Dictionaries mit Aktivitätsdaten

    Raises:
        FileNotFoundError: Datei nicht gefunden
        ValueError: Pflichtfelder fehlen oder Sheet nicht gefunden
    """
    if not Path(file_path).exists():
        raise FileNotFoundError(f"Excel-Datei nicht gefunden: {file_path}")

    try:
        # Excel-Datei einlesen (header=2 weil Spaltennamen in Zeile 3 sind)
        # Ident-Spalte explizit als String lesen, um "1.10" -> 1.1 zu verhindern
        ident_col = config.EXCEL_COLUMNS['ident']
        df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            header=2, 
            engine='openpyxl',
            dtype={ident_col: str}
        )

        # Spalten validieren
        validate_columns(df)

        # Leere Zeilen entfernen (wo Ident leer ist)
        df = df[df[config.EXCEL_COLUMNS['ident']].notna()]

        # Filter anwenden, falls vorhanden
        if filters:
            for col_name, filter_value in filters.items():
                if col_name in df.columns:
                    if filter_value == '':
                        # Filtern nach leeren Feldern
                        df = df[df[col_name].isna() | (df[col_name] == '')]
                    else:
                        # Filtern nach bestimmtem Wert (String-Vergleich für Robustheit)
                        df = df[df[col_name].astype(str) == str(filter_value)]

        # In Liste von Dictionaries konvertieren
        activities = []
        for _, row in df.iterrows():
            # E-Mail validieren
            email = clean_value(row[config.EXCEL_COLUMNS['email']])
            if not validate_email(email):
                # Überspringe Zeilen mit ungültiger E-Mail
                continue

            activity = {
                'bereich': clean_value(row.get(config.EXCEL_COLUMNS['bereich'], '')),
                'ident': clean_value(row[config.EXCEL_COLUMNS['ident']]),
                'aktivitaet': clean_value(row[config.EXCEL_COLUMNS['aktivitaet']]),
                'email': email,
                'plan_start': format_date(row[config.EXCEL_COLUMNS['plan_start']]),
                'plan_ende': format_date(row.get(config.EXCEL_COLUMNS['plan_ende'])),
                'system': clean_value(row[config.EXCEL_COLUMNS['system']]),
                'ist_status': clean_value(row.get(config.EXCEL_COLUMNS['ist_status'], '')),
            }

            activities.append(activity)

        return activities

    except FileNotFoundError:
        raise
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Fehler beim Verarbeiten der Excel-Datei: {e}")


def get_unique_values(file_path: str, sheet_name: str, column_name: str) -> List[str]:
    """
    Gibt alle eindeutigen Werte einer Spalte zurück (für Filter-Dropdown).

    Args:
        file_path: Pfad zur Excel-Datei
        sheet_name: Name des Tabellenblatts
        column_name: Name der Spalte

    Returns:
        Liste der eindeutigen Werte (sortiert)
    """
    try:
        # Ident-Spalte explizit als String lesen, um "1.10" -> 1.1 zu verhindern
        dtype_spec = {}
        if column_name == config.EXCEL_COLUMNS['ident']:
            dtype_spec = {column_name: str}
        
        df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            header=2, 
            engine='openpyxl',
            dtype=dtype_spec
        )
        if column_name in df.columns:
            unique_values = df[column_name].dropna().unique()
            return sorted([str(val) for val in unique_values])
        return []
    except:
        return []
