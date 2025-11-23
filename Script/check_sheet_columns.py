"""Zeigt alle Spalten eines bestimmten Blatts"""
import pandas as pd
import sys

file_path = "C:\\Users\\hendrik.max\\Documents\\DEV_LOCL\\cutover_mail\\Input_Datei\\DHL_JOSEF CuOvPlan DPN_ECH V01 20250813 DRAFT.xlsx"
sheet_name = "CuOv-Plan DPAG neu"

try:
    print(f"Lese Blatt: {sheet_name}\n")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

    print("ALLE Spalten:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i:3}. '{col}'")

    print(f"\nAnzahl Spalten: {len(df.columns)}")
    print(f"Anzahl Zeilen: {len(df)}")

    print("\nErste Zeile der Daten:")
    if len(df) > 0:
        first_row = df.iloc[0]
        for col in df.columns[:15]:  # Erste 15 Spalten
            print(f"  {col}: {first_row[col]}")

except Exception as e:
    print(f"FEHLER: {e}")
    import traceback
    traceback.print_exc()
