import zipfile
import xml.etree.ElementTree as ET

def verify_xlsx(filename):
    print(f"Verifying {filename}...")
    try:
        with zipfile.ZipFile(filename, 'r') as z:
            # Check for sheets
            workbook_xml = z.read('xl/workbook.xml')
            root = ET.fromstring(workbook_xml)
            sheets = [node.get('name') for node in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')]

            expected_sheets = [
                'Dashboard', 'Original PL', 'Data Whole', 'Partial',
                'Open Boxes', 'First Scan', 'SCAN WHOLE', 'Segregation',
                'Scan Partial', 'Data After Reception', 'Missing', 'Settings'
            ]

            for s in expected_sheets:
                if s in sheets:
                    print(f"[OK] Sheet '{s}' found.")
                else:
                    print(f"[ERROR] Sheet '{s}' MISSING.")

            # Check for tables
            table_files = [f for f in z.namelist() if f.startswith('xl/tables/table')]
            print(f"Found {len(table_files)} tables.")

    except Exception as e:
        print(f"Verification failed: {e}")

if __name__ == '__main__':
    verify_xlsx('SmartReceive_Excel_Pro.xlsx')
