#!/usr/bin/env python3
"""
Simple structural validation for extract_bok_statements.py
Tests code structure without importing (no dependencies needed)
"""

import re
import os

def read_file():
    """Read the script file"""
    script_path = os.path.join(os.path.dirname(__file__), 'extract_bok_statements.py')
    with open(script_path, 'r') as f:
        return f.read()

def test_header():
    """Test that file has correct header"""
    content = read_file()
    if 'BOK Financial Statement Parser' in content:
        print("✓ File has correct header")
        return True
    else:
        print("✗ File missing correct header")
        return False

def test_no_wfa_parser():
    """Test that WFAParser class is not present"""
    content = read_file()
    if 'class WFAParser' not in content:
        print("✓ WFAParser class correctly removed")
        return True
    else:
        print("✗ WFAParser class should be removed")
        return False

def test_has_bok_parser():
    """Test that BOKFinancialParser exists"""
    content = read_file()
    if 'class BOKFinancialParser:' in content:
        print("✓ BOKFinancialParser class exists")
        return True
    else:
        print("✗ BOKFinancialParser class missing")
        return False

def test_has_bok_statement_parser():
    """Test that BOKStatementParser exists (not CombinedStatementParser)"""
    content = read_file()
    has_bok = 'class BOKStatementParser:' in content
    no_combined = 'class CombinedStatementParser' not in content
    
    if has_bok and no_combined:
        print("✓ BOKStatementParser class exists and CombinedStatementParser removed")
        return True
    else:
        if not has_bok:
            print("✗ BOKStatementParser class missing")
        if not no_combined:
            print("✗ CombinedStatementParser should be renamed to BOKStatementParser")
        return False

def test_is_bok_financial_method():
    """Test that is_bok_financial method exists"""
    content = read_file()
    if 'def is_bok_financial(' in content:
        print("✓ is_bok_financial method exists")
        return True
    else:
        print("✗ is_bok_financial method missing")
        return False

def test_extract_beneficiary_function():
    """Test that extract_beneficiary_and_account function exists"""
    content = read_file()
    if 'def extract_beneficiary_and_account(' in content:
        print("✓ extract_beneficiary_and_account function exists")
        return True
    else:
        print("✗ extract_beneficiary_and_account function missing")
        return False

def test_gui_functions():
    """Test that GUI functions exist"""
    content = read_file()
    has_select_source = 'def select_source_folder(' in content
    has_select_output = 'def select_output_file(' in content
    
    if has_select_source and has_select_output:
        print("✓ GUI selection functions exist")
        return True
    else:
        print("✗ GUI selection functions missing")
        return False

def test_main_function():
    """Test that main function exists and has correct title"""
    content = read_file()
    has_main = 'def main():' in content
    has_title = 'BOK Financial Statement Parser' in content and 'print("BOK Financial Statement Parser")' in content
    
    if has_main:
        print("✓ main() function exists")
        if has_title:
            print("  ✓ Has correct title")
        return has_main and has_title
    else:
        print("✗ main() function missing")
        return False

def test_excel_title():
    """Test that Excel title is correct"""
    content = read_file()
    # Check for dynamic title (should use datetime formatting now)
    if 'BOK Financial Statement Summary' in content and ('title_date' in content or 'strftime' in content):
        print("✓ Excel title is correct (using dynamic date)")
        return True
    elif 'BOK Financial Statement Summary - October 2025' in content:
        print("⚠ Excel title uses hardcoded date (consider making dynamic)")
        return True
    else:
        print("✗ Excel title incorrect or missing")
        return False

def test_simplified_output():
    """Test that output messages are simplified for BOK only"""
    content = read_file()
    
    # Should have BOK-specific messages
    has_bok_scanning = 'Scanning' in content and 'BOK Financial statements' in content
    has_skipping_msg = 'Skipping' in content and 'non-BOK' in content
    
    # Should NOT have WFA statistics
    no_wfa_stats = 'wfa_total' not in content and 'wfa_count' not in content
    
    if has_bok_scanning and has_skipping_msg and no_wfa_stats:
        print("✓ Output messages simplified for BOK only")
        return True
    else:
        if not has_bok_scanning:
            print("✗ Missing BOK scanning messages")
        if not has_skipping_msg:
            print("✗ Missing non-BOK skipping messages")
        if not no_wfa_stats:
            print("✗ WFA statistics variables should be removed")
        return False

def test_default_filename():
    """Test that default output filename is BOK_Statement_Summary"""
    content = read_file()
    if 'BOK_Statement_Summary' in content:
        print("✓ Default output filename is correct")
        return True
    else:
        print("✗ Default output filename should be BOK_Statement_Summary")
        return False

def test_syntax():
    """Test Python syntax by compiling"""
    try:
        script_path = os.path.join(os.path.dirname(__file__), 'extract_bok_statements.py')
        with open(script_path, 'r') as f:
            compile(f.read(), script_path, 'exec')
        print("✓ Python syntax is valid")
        return True
    except SyntaxError as e:
        print(f"✗ Syntax error: {e}")
        return False

def main():
    """Run all validation tests"""
    print("="*80)
    print("Structural Validation for extract_bok_statements.py")
    print("="*80)
    print()
    
    tests = [
        ("Syntax Check", test_syntax),
        ("File Header", test_header),
        ("No WFA Parser", test_no_wfa_parser),
        ("Has BOK Financial Parser", test_has_bok_parser),
        ("Has BOK Statement Parser", test_has_bok_statement_parser),
        ("is_bok_financial Method", test_is_bok_financial_method),
        ("extract_beneficiary_and_account Function", test_extract_beneficiary_function),
        ("GUI Functions", test_gui_functions),
        ("Main Function", test_main_function),
        ("Excel Title", test_excel_title),
        ("Simplified Output", test_simplified_output),
        ("Default Filename", test_default_filename),
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n{test_name}:")
        print("-" * 80)
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"✗ Test failed with exception: {e}")
            results.append((test_name, False))
    
    # Print summary
    print("\n" + "="*80)
    print("Test Summary:")
    print("="*80)
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status}: {test_name}")
    
    print()
    print(f"Total: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n✓ All validation tests passed!")
        return 0
    else:
        print(f"\n⚠ {total - passed} test(s) failed")
        return 1

if __name__ == "__main__":
    import sys
    sys.exit(main())
