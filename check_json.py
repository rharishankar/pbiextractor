import json
import os

def check_json_file(json_file='pbix_analysis/complete_analysis.json'):
    """Diagnose what's in your JSON file."""
    
    print("=" * 80)
    print("JSON FILE DIAGNOSTIC")
    print("=" * 80)
    print(f"Checking: {json_file}\n")
    
    if not os.path.exists(json_file):
        print(f"❌ File not found: {json_file}")
        print("\nPossible solutions:")
        print("1. Make sure you ran the PBIX parser first")
        print("2. Check the file path is correct")
        print("3. The parser creates: pbix_analysis/complete_analysis.json")
        return
    
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"❌ Error reading JSON: {e}")
        return
    
    print("✅ JSON file loaded successfully\n")
    print("=" * 80)
    print("WHAT'S IN YOUR FILE:")
    print("=" * 80)
    
    # Check top-level keys
    print("\nTop-level sections found:")
    for key in data.keys():
        print(f"  ✓ {key}")
    
    # Check for dataModel
    if 'dataModel' in data:
        print("\n✅ GOOD NEWS: 'dataModel' section EXISTS!")
        
        dm = data['dataModel']
        print("\nData Model contains:")
        for key in dm.keys():
            if key == 'tables':
                print(f"  ✓ {key}: {len(dm[key])} tables")
            elif key == 'relationships':
                print(f"  ✓ {key}: {len(dm[key])} relationships")
            elif key == 'roles':
                print(f"  ✓ {key}: {len(dm[key])} roles")
            else:
                print(f"  ✓ {key}")
        
        # Check if tables have data
        if dm.get('tables'):
            print(f"\n✅ You have {len(dm['tables'])} tables - READY TO ANALYZE!")
            print("\nSample table names:")
            for i, table in enumerate(dm['tables'][:5]):
                print(f"  • {table.get('name', 'Unknown')}")
            if len(dm['tables']) > 5:
                print(f"  ... and {len(dm['tables']) - 5} more")
        else:
            print("\n⚠️  WARNING: 'tables' list is empty")
    
    else:
        print("\n❌ PROBLEM: No 'dataModel' section found!")
        print("\nThis means:")
        print("  • Your PBIX is in an older format (pre-2020)")
        print("  • OR the DataModelSchema file was missing")
        
        print("\n" + "=" * 80)
        print("SOLUTIONS:")
        print("=" * 80)
        
        # Check what else is available
        if 'reportLayout' in data:
            print("\n✓ You DO have: Report Layout (pages and visuals)")
        if 'connections' in data:
            print("✓ You DO have: Data Connections")
        if 'metadata' in data:
            print("✓ You DO have: Metadata")
        
        print("\n📋 RECOMMENDED ACTIONS:")
        print("-" * 80)
        print("\nOption 1: Convert PBIX to PBIT")
        print("  1. Open your PBIX in Power BI Desktop")
        print("  2. File → Save As → Template (.pbit)")
        print("  3. Run the parser on the PBIT file instead")
        
        print("\nOption 2: Use pbi-tools (if available)")
        print("  1. Ask your developer to run: pbi-tools extract yourfile.pbix")
        print("  2. Send you the extracted folder")
        print("  3. Point the analysis script to that folder")
        
        print("\nOption 3: Analyze what you DO have")
        print("  I can create a script to analyze just the report layout,")
        print("  connections, and metadata (without the data model)")
    
    print("\n" + "=" * 80)

if __name__ == '__main__':
    import sys
    
    # Allow command line argument
    json_file = 'pbix_analysis/complete_analysis.json'
    if len(sys.argv) > 1:
        json_file = sys.argv[1]
    
    check_json_file(json_file)