"""
Debug version 2 - Show more text to find where "Closing value" appears
"""

import re
from pathlib import Path
import PyPDF2

PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"

pdf_files = list(Path(PDF_FOLDER).glob('*.pdf'))

if pdf_files:
    # Test just the first PDF
    test_file = pdf_files[0]
    print(f"Testing: {test_file.name}\n")
    
    with open(test_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Find the SNAPSHOT page
        for page_num, page in enumerate(pdf_reader.pages):
            text = page.extract_text()
            if 'SNAPSHOT' in text:
                print(f"Found SNAPSHOT on page {page_num + 1}")
                print(f"\n{'='*80}")
                print("FULL PAGE TEXT:")
                print(f"{'='*80}\n")
                print(text)
                print(f"\n{'='*80}")
                
                # Look for "closing" (case insensitive)
                if 'losing' in text.lower():
                    print("\nFound 'losing' in text - searching for context:")
                    matches = re.finditer(r'.{0,50}[Ll]osing.{0,100}', text)
                    for i, match in enumerate(matches, 1):
                        print(f"\nMatch {i}: '{match.group()}'")
                break