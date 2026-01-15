"""
Test pdfplumber on one PDF to see if it extracts the text properly
"""

from pathlib import Path
import pdfplumber
import re

PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"

pdf_files = list(Path(PDF_FOLDER).glob('*.pdf'))

if pdf_files:
    # Test the first PDF
    test_file = pdf_files[0]
    print(f"Testing: {test_file.name}\n")
    
    with pdfplumber.open(test_file) as pdf:
        print(f"Total pages in PDF: {len(pdf.pages)}\n")
        
        # Search through all pages for "Page 1 of [any number]"
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            # Check if this page has "Page 1 of XX" (where XX is any number)
            if re.search(r'Page\s+1\s+of\s+\d+', text):
                print(f"✓ Found 'Page 1 of XX' on PDF page {page_num + 1}")
                print(f"\n{'='*80}")
                print("FULL PAGE TEXT:")
                print(f"{'='*80}\n")
                print(text)
                print(f"\n{'='*80}")
                
                # Look for "Closing value"
                closing_matches = re.finditer(r'.{0,30}Closing\s+value.{0,80}', text, re.IGNORECASE)
                print("\nFound 'Closing value' references:")
                for match in closing_matches:
                    print(f"  '{match.group()}'")
                
                # Try to extract the value
                value_match = re.search(r'Closing\s+value\s+\$\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
                if value_match:
                    print(f"\n✓ Extracted value: ${value_match.group(1)}")
                else:
                    print("\n⚠ Could not extract value")
                
                break
        else:
            print("⚠ Did not find 'Page 1 of XX' on any page")
            print("\nShowing first 3 pages to find the pattern:")
            for i in range(min(3, len(pdf.pages))):
                print(f"\n--- PDF Page {i+1} (first 300 chars) ---")
                page_text = pdf.pages[i].extract_text()
                print(page_text[:300] if page_text else "No text extracted")