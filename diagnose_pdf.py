"""
Diagnostic script to see what's actually in the PDF
"""

from pathlib import Path
import pdfplumber
import re

PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"

# Test the first PDF
pdf_files = list(Path(PDF_FOLDER).glob('*.pdf'))

if pdf_files:
    test_file = pdf_files[0]  # 5Y Re BIC - 7282-8588.pdf
    print(f"Analyzing: {test_file.name}\n")
    
    with pdfplumber.open(test_file) as pdf:
        print(f"Total pages: {len(pdf.pages)}\n")
        
        for page_num, page in enumerate(pdf.pages):
            print(f"\n{'='*80}")
            print(f"PAGE {page_num + 1}")
            print(f"{'='*80}")
            
            text = page.extract_text()
            
            if text:
                print(f"Text length: {len(text)} characters")
                print(f"\nFirst 800 characters:")
                print(text[:800])
                
                # Check for key terms
                print(f"\nKey term checks:")
                print(f"  - Contains 'SNAPSHOT': {'SNAPSHOT' in text}")
                print(f"  - Contains 'Page 1 of': {bool(re.search(r'Page\s+1\s+of', text))}")
                print(f"  - Contains 'Closing': {'Closing' in text or 'closing' in text}")
                print(f"  - Contains 'Progress summary': {'Progress summary' in text}")
                
                # Search for any dollar amounts
                dollar_amounts = re.findall(r'\$[\d,]+\.\d{2}', text)
                if dollar_amounts:
                    print(f"\n  Found dollar amounts: {dollar_amounts[:10]}")
                
                # Look for any line with "value" in it
                value_lines = [line for line in text.split('\n') if 'value' in line.lower()]
                if value_lines:
                    print(f"\n  Lines containing 'value':")
                    for line in value_lines[:5]:
                        print(f"    {line[:100]}")
            else:
                print("No text extracted from this page")