#!/usr/bin/env python3
"""
Simple validation test for extract_bok_statements.py
Tests that the script has correct structure and key functions
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

def test_imports():
    """Test that the module imports without errors"""
    try:
        import extract_bok_statements
        print("✓ Module imports successfully")
        return True
    except Exception as e:
        print(f"✗ Failed to import module: {e}")
        return False

def test_functions_exist():
    """Test that required functions exist"""
    import extract_bok_statements
    
    required_functions = [
        'extract_beneficiary_and_account',
        'select_source_folder',
        'select_output_file',
        'main'
    ]
    
    all_exist = True
    for func_name in required_functions:
        if hasattr(extract_bok_statements, func_name):
            print(f"✓ Function '{func_name}' exists")
        else:
            print(f"✗ Function '{func_name}' missing")
            all_exist = False
    
    return all_exist

def test_classes_exist():
    """Test that required classes exist"""
    import extract_bok_statements
    
    required_classes = [
        'BOKFinancialParser',
        'BOKStatementParser'
    ]
    
    # Should NOT have WFAParser
    should_not_exist = ['WFAParser', 'CombinedStatementParser']
    
    all_correct = True
    
    for class_name in required_classes:
        if hasattr(extract_bok_statements, class_name):
            print(f"✓ Class '{class_name}' exists")
        else:
            print(f"✗ Class '{class_name}' missing")
            all_correct = False
    
    for class_name in should_not_exist:
        if not hasattr(extract_bok_statements, class_name):
            print(f"✓ Class '{class_name}' correctly removed")
        else:
            print(f"✗ Class '{class_name}' should not exist")
            all_correct = False
    
    return all_correct

def test_bok_parser_methods():
    """Test that BOKStatementParser has required methods"""
    import extract_bok_statements
    
    parser_class = extract_bok_statements.BOKStatementParser
    required_methods = [
        'is_bok_financial',
        'process_all_pdfs',
        'save_to_excel'
    ]
    
    all_exist = True
    for method_name in required_methods:
        if hasattr(parser_class, method_name):
            print(f"✓ Method 'BOKStatementParser.{method_name}' exists")
        else:
            print(f"✗ Method 'BOKStatementParser.{method_name}' missing")
            all_exist = False
    
    return all_exist

def test_beneficiary_extraction():
    """Test beneficiary and account extraction from filenames"""
    import extract_bok_statements
    
    test_cases = [
        # (filename, expected_beneficiary, expected_account_pattern)
        ("First Coverage Re BIC - 1150-0007374.1.pdf", "BIC", r"\d+-\d+\.\d+"),
        ("Kamal Alhajli WH BIC - 3719-3369.pdf", "BIC", r"\d+-\d+"),
        ("Auto Lane Re DAC - 3292-9150.pdf", "DAC", r"\d+-\d+"),
    ]
    
    import re
    all_passed = True
    
    for filename, expected_ben, account_pattern in test_cases:
        beneficiary, account = extract_bok_statements.extract_beneficiary_and_account(filename)
        
        if beneficiary == expected_ben:
            print(f"✓ Correctly extracted beneficiary '{beneficiary}' from '{filename[:30]}...'")
        else:
            print(f"✗ Expected beneficiary '{expected_ben}' but got '{beneficiary}' from '{filename[:30]}...'")
            all_passed = False
        
        if account and re.match(account_pattern, account):
            print(f"  ✓ Extracted valid account number '{account}'")
        else:
            print(f"  ⚠ Account number '{account}' may not match expected pattern")
    
    return all_passed

def main():
    """Run all validation tests"""
    print("="*80)
    print("Validating extract_bok_statements.py")
    print("="*80)
    print()
    
    tests = [
        ("Import Test", test_imports),
        ("Function Existence Test", test_functions_exist),
        ("Class Existence Test", test_classes_exist),
        ("BOKStatementParser Methods Test", test_bok_parser_methods),
        ("Beneficiary Extraction Test", test_beneficiary_extraction),
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
    sys.exit(main())
