"""Zeigt alle Tabellenblätter und deren Spalten"""
import pandas as pd
import sys

if len(sys.argv) > 1:
    file_path = sys.argv[1]
else:
    print("Bitte Pfad zur Excel-Datei angeben")
    sys.exit(1)

try:
    excel_file = pd.ExcelFile(file_path)
    print(f"Datei: {file_path}\n")
    print(f"Gefundene Tabellenblätter: {len(excel_file.sheet_names)}\n")

    for i, sheet_name in enumerate(excel_file.sheet_names, 1):
        print(f"\n{'='*60}")
        print(f"Blatt {i}: {sheet_name}")
        print('='*60)

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
            print(f"Spalten: {list(df.columns)[:10]}")  # Erste 10 Spalten
            print(f"Anzahl Zeilen: {len(df)}")

            # Prüfe ob Pflichtfelder vorhanden sind
            required = ['Ident', 'Aktivität', 'E-Mail', 'PLAN-Start', 'System/Mandant-Buchungskreis']
            found = [col for col in required if col in df.columns]
            if found:
                print(f"✓ Gefundene Pflichtfelder: {found}")

        except Exception as e:
            print(f"Fehler beim Lesen: {e}")

except Exception as e:
    print(f"FEHLER: {e}")
    import traceback
    traceback.print_exc()
