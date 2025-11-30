
import sys
import os
import pandas as pd
from unittest.mock import MagicMock

# Mock config
sys.modules['config'] = MagicMock()
import config
config.EXCEL_COLUMNS = {
    'bereich': 'Bereich',
    'ident': 'Ident',
    'aktivitaet': 'Aktivität',
    'email': 'E-Mail',
    'plan_start': 'PLAN-Start',
    'plan_ende': 'PLAN-Ende',
    'system': 'System',
    'ist_status': 'IST-Status'
}
config.REQUIRED_COLUMNS = ['Ident', 'Aktivität', 'E-Mail', 'PLAN-Start', 'System']

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import excel_parser

def test_ident_truncation():
    # Create a DataFrame where 'Ident' is a float
    # This simulates what happens when Excel reads "1.10" as a number
    data = {
        'Ident': [1.10],
        'Aktivität': ['Test Activity'],
        'E-Mail': ['test@example.com'],
        'PLAN-Start': ['2025-01-01'],
        'System': ['SYS1'],
        'IST-Status': ['']
    }
    df = pd.DataFrame(data)
    
    # Mock read_excel to return this dataframe
    # Note: In reality, read_excel does the type conversion. 
    # Here we simulate the RESULT of read_excel being a float.
    pd.read_excel = MagicMock(return_value=df)
    pd.ExcelFile = MagicMock()
    pd.ExcelFile.return_value.sheet_names = ['Sheet1']
    
    # Mock path exists
    with open('dummy_ident.xlsx', 'w') as f:
        f.write('dummy')
        
    try:
        activities = excel_parser.load_activities('dummy_ident.xlsx', 'Sheet1')
        if len(activities) > 0:
            act = activities[0]
            ident = act['ident']
            print(f"Ident loaded: '{ident}'")
            if ident == "1.1":
                print("REPRODUCED: Ident '1.10' (float) became '1.1'")
            elif ident == "1.10":
                print("FIXED: Ident preserved as '1.10'")
            else:
                print(f"Unexpected Ident: {ident}")
    except Exception as e:
        print(f"FAILURE: Parser crashed with: {e}")
    finally:
        if os.path.exists('dummy_ident.xlsx'):
            os.remove('dummy_ident.xlsx')

if __name__ == "__main__":
    test_ident_truncation()
