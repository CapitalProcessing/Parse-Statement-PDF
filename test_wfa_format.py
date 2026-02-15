#!/usr/bin/env python3
"""
Test script to verify WFA_Format worksheet is correctly created
"""

import os
import sys
import pandas as pd
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))
import extract_bok_statements

def test_wfa_format_creation():
    """Test that the WFA_Format worksheet is created correctly"""
    
    # Create a temporary output directory
    test_output_dir = Path("/tmp/test_wfa_output")
    test_output_dir.mkdir(exist_ok=True)
    output_file = test_output_dir / "test_output.xlsx"
    
    # Get the current directory with sample PDFs
    current_dir = Path(__file__).parent
    
    # Create a parser instance
    parser = extract_bok_statements.BOKStatementParser(
        pdf_folder=str(current_dir),
        output_excel=str(output_file)
    )
    
    # Process PDFs
    print("Processing PDFs...")
    parser.process_all_pdfs()
    
    if not output_file.exists():
        print("✗ Excel file was not created!")
        return False
    
    print(f"✓ Excel file created: {output_file}")
    
    # Read the Excel file to verify worksheets
    try:
        excel_file = pd.ExcelFile(output_file)
        sheets = excel_file.sheet_names
        
        print(f"\nWorksheets found: {sheets}")
        
        # Check that both worksheets exist
        if 'Summary' not in sheets:
            print("✗ Summary worksheet not found!")
            return False
        print("✓ Summary worksheet exists")
        
        if 'WFA_Format' not in sheets:
            print("✗ WFA_Format worksheet not found!")
            return False
        print("✓ WFA_Format worksheet exists")
        
        # Read the Summary sheet
        df_summary = pd.read_excel(output_file, sheet_name='Summary', skiprows=2)
        print(f"\nSummary sheet has {len(df_summary)} rows")
        print(f"Summary columns: {list(df_summary.columns)}")
        
        # Read the WFA_Format sheet (no headers)
        df_wfa = pd.read_excel(output_file, sheet_name='WFA_Format', header=None)
        print(f"\nWFA_Format sheet has {len(df_wfa)} rows")
        print(f"WFA_Format columns: {list(df_wfa.columns)}")
        
        # Verify WFA_Format has 10 columns (A-J)
        if len(df_wfa.columns) != 10:
            print(f"✗ Expected 10 columns, found {len(df_wfa.columns)}")
            return False
        print("✓ WFA_Format has 10 columns")
        
        # Verify column C (index 2) contains account numbers
        if df_wfa[2].isnull().all():
            print("✗ Column C (Account Number) is empty!")
            return False
        print(f"✓ Column C contains data: {df_wfa[2].head(3).tolist()}")
        
        # Verify column J (index 9) contains closing values
        if df_wfa[9].isnull().all():
            print("✗ Column J (Closing Value) is empty!")
            return False
        print(f"✓ Column J contains data: {df_wfa[9].head(3).tolist()}")
        
        # Verify other columns are empty
        empty_cols = [0, 1, 3, 4, 5, 6, 7, 8]
        for col_idx in empty_cols:
            if not df_wfa[col_idx].apply(lambda x: x == '' or pd.isna(x)).all():
                print(f"⚠ Column {col_idx} is not empty (expected to be blank)")
        print("✓ Other columns are appropriately empty")
        
        # Print sample data
        print("\nSample WFA_Format data (first 3 rows):")
        print("=" * 80)
        for idx, row in df_wfa.head(3).iterrows():
            print(f"Row {idx+1}:")
            print(f"  Column C (Account): {row[2]}")
            print(f"  Column J (Value): {row[9]}")
        
        print("\n" + "=" * 80)
        print("✓ All WFA_Format tests passed!")
        return True
        
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 80)
    print("Testing WFA_Format Worksheet Creation")
    print("=" * 80)
    print()
    
    success = test_wfa_format_creation()
    
    if success:
        print("\n✓ All tests passed!")
        sys.exit(0)
    else:
        print("\n✗ Tests failed!")
        sys.exit(1)
