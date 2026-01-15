"""
Diagnostic script using PyPDF2 to see all pages
"""

from pathlib import Path
import PyPDF2
import re

PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"

# Test the first PDF
pdf_files = list(Path(PDF_FOLDER).glob('*.pdf'))

if pdf_files:
    test_file = pdf_files[0]  # 5Y Re BIC - 7282-8588.pdf
    print(f"Analyzing: {test_file.name}\n")
    
    with open(test_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        print(f"Total pages (PyPDF2): {len(pdf_reader.pages)}\n")
        
        for page_num, page in enumerate(pdf_reader.pages):
            print(f"\n{'='*80}")
            print(f"PAGE {page_num + 1} (Physical page {page_num + 1})")
            print(f"{'='*80}")
            
            try:
                text = page.extract_text()
                
                if text:
                    print(f"Text length: {len(text)} characters")
                    print(f"\nFirst 800 characters:")
                    print(text[:800])
                    
                    # Check for key terms
                    print(f"\nKey term checks:")
                    print(f"  - Contains 'SNAPSHOT': {'SNAPSHOT' in text}")
                    print(f"  - Contains 'Page 1 of': {bool(re.search(r'Page\s*1\s*of', text))}")
                    print(f"  - Contains 'Closing': {'Closing' in text or 'closing' in text.lower()}")
                    print(f"  - Contains 'Progress': {'Progress' in text}")
                    
                    # Search for any dollar amounts
                    dollar_amounts = re.findall(r'\$[\d,]+\.\d{2}', text)
                    if dollar_amounts:
                        print(f"\n  Found {len(dollar_amounts)} dollar amounts: {dollar_amounts[:5]}")
                    
                    # Look for "Closing value" specifically
                    if 'losing' in text.lower():
                        # Find context around "closing"
                        matches = re.finditer(r'.{0,40}[Cc]losing.{0,60}', text)
                        print(f"\n  'Closing' context:")
                        for match in matches:
                            print(f"    {match.group()}")
                    
                else:
                    print("No text extracted from this page")
                    
            except Exception as e:
                print(f"Error extracting text: {e}")