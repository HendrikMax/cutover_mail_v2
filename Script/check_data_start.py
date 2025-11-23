"""Findet wo Daten anfangen"""
import pandas as pd

file_path = "C:\\Users\\hendrik.max\\Documents\\DEV_LOCL\\cutover_mail\\Input_Datei\\DHL_JOSEF CuOvPlan DPN_ECH V01 20250813 DRAFT.xlsx"
sheet_name = "CuOv-Plan DPAG neu"

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

    print("Erste 10 Zeilen - Spalte 'Ident':")
    for i in range(min(10, len(df))):
        ident = df.iloc[i]['Ident']
        aktivitaet = df.iloc[i]['Aktivität'] if 'Aktivität' in df.columns else 'N/A'
        email = df.iloc[i]['E-Mail'] if 'E-Mail' in df.columns else 'N/A'
        print(f"Zeile {i}: Ident={ident}, Aktivität={aktivitaet[:30] if pd.notna(aktivitaet) else 'nan'}, E-Mail={email}")

    # Entferne leere Zeilen
    df_clean = df[df['Ident'].notna()]
    print(f"\nNach Entfernen leerer Zeilen: {len(df_clean)} Zeilen")

    if len(df_clean) > 0:
        print("\nErste Zeile mit Daten:")
        first = df_clean.iloc[0]
        print(f"  Ident: {first['Ident']}")
        print(f"  Aktivität: {first['Aktivität']}")
        print(f"  E-Mail: {first['E-Mail']}")

except Exception as e:
    print(f"FEHLER: {e}")
    import traceback
    traceback.print_exc()
