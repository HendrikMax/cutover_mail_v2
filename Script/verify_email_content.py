"""
Verification script to test email content generation
"""
import sys
from pathlib import Path
from email import message_from_file

# Add Script directory to path
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

import email_generator

# Mock activity data
test_activity = {
    'ident': '1.1 - DE58',
    'aktivitaet': 'Livesetzen des neuen 925NK Prozesses, damit DHL AG Fahrzeuge korrekt korrigiert werden (Vorsteuerkorrektur)',
    'plan_start': '15.12.2025 00:00',
    'system': 'FP2/100-1000',
    'email': 'test@example.com'
}

test_cutover_ident = 'JOSEF'
test_sheet_name = 'DPN_ECH'
test_cutover_link = 'https://dpdhl.sharepoint.com/:x:/r/teams/Josef400-Finanz-Systeme/Shared%20Documents/Systeme%20Finance/00%20Cross%20WS%20Information/Cutover/03%20CuOv_DPN_ECH/01%20Plan/DHL_JOSEF%20CuOvPlan%20DPN_ECH%20V01%2020250813%20DRAFT.xlsx?d=wdd91fe4364544d2681336c99f0a7908c&csf=1&web=1&e=J8D7VO'
test_output_path = str(script_dir / 'test_output')

print("Testing email generation with new parameters...")
print(f"Sheet name: {test_sheet_name}")
print(f"Cutover link: {test_cutover_link[:50]}...")
print()

# Create test output directory
Path(test_output_path).mkdir(exist_ok=True)

# Generate EML file
try:
    file_path = email_generator.save_as_eml(
        test_activity,
        test_cutover_ident,
        test_output_path,
        test_sheet_name,
        test_cutover_link
    )
    
    print(f"✓ EML file created: {file_path}")
    
    # Parse EML file using email library
    with open(file_path, 'r', encoding='utf-8') as f:
        msg = message_from_file(f)
    
    # Extract text content from multipart message
    decoded_content = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type in ['text/plain', 'text/html']:
                payload = part.get_payload(decode=True)
                if payload:
                    decoded_content += payload.decode('utf-8', errors='ignore')
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            decoded_content = payload.decode('utf-8', errors='ignore')
    
    # Check for required strings
    checks = [
        (test_sheet_name, "Sheet name"),
        (test_cutover_link[:50], "Cutover plan link (first 50 chars)"),
        ("abgeschlossen", "Status instruction"),
        ("sende mir die E-Mail", "Reply instruction"),
        ("viel Erfolg", "Closing phrase")
    ]
    
    print("\nVerifying content:")
    all_passed = True
    for check_string, description in checks:
        if check_string in decoded_content:
            print(f"✓ {description} found")
        else:
            print(f"✗ {description} NOT found")
            all_passed = False
    
    if all_passed:
        print("\n✓ All verification checks passed!")
        print("\nSample of decoded content:")
        print(decoded_content[:500])
    else:
        print("\n✗ Some checks failed!")
        print("\nDecoded content preview:")
        print(decoded_content[:500])
        sys.exit(1)
        
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
