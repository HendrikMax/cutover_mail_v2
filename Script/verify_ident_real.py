
import sys
import os
import pandas as pd
import config

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import excel_parser

def test_ident_preservation():
    filename = 'test_ident_preservation.xlsx'
    sheet_name = 'TestSheet'
    
    # Create a DataFrame with "1.10" as a string
    # We need to ensure it's written as string to Excel if possible, 
    # or at least see how pandas handles it.
    # If we write it as string, and read it back without dtype, pandas might infer float.
    
    df = pd.DataFrame({
        'Ident': ['1.10'],
        'Aktivität': ['Test Activity'],
        'E-Mail': ['test@example.com'],
        'PLAN-Start': ['2025-01-01'],
        'PLAN-Ende': ['2025-01-01'],
        'System/Mandant-Buchungskreis': ['SYS1'],
        'IST-Status': ['']
    })
    
    # Write to Excel
    # We write to row 3 (header=2 means header is at index 2, i.e., row 3)
    # So we need empty rows before.
    # Actually, to match the parser: header=2 means the header is the 3rd row (0-indexed 2).
    # So we need 2 empty rows, then the header, then data.
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write empty dataframe for first 2 rows? Or just start at startrow=2
        df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
        
    try:
        print(f"Reading back from {filename}...")
        activities = excel_parser.load_activities(filename, sheet_name)
        
        if len(activities) > 0:
            ident = activities[0]['ident']
            print(f"Ident loaded: '{ident}'")
            
            if ident == "1.10":
                print("SUCCESS: Ident '1.10' preserved.")
            elif ident == "1.1":
                print("FAILURE: Ident '1.10' became '1.1'.")
            else:
                print(f"Unexpected Ident: {ident}")
        else:
            print("FAILURE: No activities loaded.")
            
    except Exception as e:
        print(f"FAILURE: {e}")
    finally:
        if os.path.exists(filename):
            os.remove(filename)

if __name__ == "__main__":
    test_ident_preservation()
