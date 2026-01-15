"""
Bank Statement Parser - DEBUG VERSION
This version will show us the extracted text to help diagnose the issue
"""

import re
from pathlib import Path
import PyPDF2
import pandas as pd
from typing import Optional

class BankStatementParser:
    def __init__(self, pdf_folder: str, output_excel: str = "Bank_Statement_Summary.xlsx"):
        self.pdf_folder = Path(pdf_folder)
        self.output_excel = output_excel
        self.results = []
    
    def extract_account_number_from_filename(self, filename: str) -> Optional[str]:
        """Extract account number from filename"""
        name_part = filename.replace('.pdf', '')
        if ' - ' in name_part:
            parts = name_part.rsplit(' - ', 1)
            if len(parts) > 1:
                # Clean up the account number part
                account = parts[1].strip()
                # Remove any extra text like " BIC" at the end
                account = account.split()[0]
                return account
        return None
    
    def extract_page_one_text(self, pdf_path: Path) -> Optional[str]:
        """Extract text from the page labeled 'Page 1 of XXX'"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Search through pages to find "Page 1 of"
                for page_num, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    # Look for "Page 1 of" pattern
                    if re.search(r'Page\s+1\s+of\s+\d+', text, re.IGNORECASE):
                        return text
                
                # Fallback: if not found, try the first page
                if len(pdf_reader.pages) > 0:
                    return pdf_reader.pages[0].extract_text()
                    
        except Exception as e:
            print(f"Error reading {pdf_path.name}: {e}")
        
        return None
    
    def extract_closing_value(self, text: str) -> Optional[float]:
        """Extract closing value from page 1 text"""
        if not text:
            return None
        
        # Multiple patterns to try
        patterns = [
            r'Closing\s+value\s+\$\s*([\d,]+\.\d{2})',
            r'Closing\s+value[:\s]+\$?\s*([\d,]+\.\d{2})',
            r'CLOSING\s+VALUE[:\s]+\$?\s*([\d,]+\.\d{2})',
            # Try without space before $
            r'Closing\s+value\$\s*([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value_str = match.group(1).replace(',', '').replace(' ', '')
                try:
                    value = float(value_str)
                    if value >= 0:
                        return value
                except ValueError:
                    continue
        
        return None
    
    def process_pdf_debug(self, pdf_path: Path, show_text: bool = False) -> dict:
        """Process a single PDF with debug output"""
        print(f"\nProcessing: {pdf_path.name}")
        
        # Extract account number from filename
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        print(f"  Account Number: {account_number}")
        
        # Extract page 1 text
        page_text = self.extract_page_one_text(pdf_path)
        
        if show_text and page_text:
            print(f"\n  --- TEXT EXTRACT (first 500 chars) ---")
            print(page_text[:500])
            print(f"  --- END TEXT EXTRACT ---\n")
            
            # Look for "Closing" in the text
            closing_matches = re.finditer(r'.{0,30}[Cc]losing.{0,50}', page_text)
            print(f"  Found 'Closing' references:")
            for match in closing_matches:
                print(f"    '{match.group()}'")
        
        # Extract closing value
        closing_value = self.extract_closing_value(page_text)
        print(f"  Closing Value: {closing_value if closing_value else '⚠ NOT FOUND'}")
        
        status = "✓" if (account_number and closing_value is not None) else "⚠"
        
        return {
            'Filename': pdf_path.name,
            'Account Number': account_number,
            'Closing Value': closing_value,
            'Status': status
        }
    
    def test_first_few(self, count: int = 3):
        """Test the first few PDFs with detailed output"""
        if not self.pdf_folder.exists():
            print(f"Error: Folder '{self.pdf_folder}' does not exist!")
            return
        
        pdf_files = list(self.pdf_folder.glob('*.pdf'))[:count]
        
        print(f"\n{'='*80}")
        print(f"TESTING FIRST {count} PDFs WITH DETAILED OUTPUT")
        print(f"{'='*80}")
        
        for pdf_file in pdf_files:
            result = self.process_pdf_debug(pdf_file, show_text=True)
            self.results.append(result)
        
        print(f"\n{'='*80}")
        print("Test complete. Check the output above to see what text is being extracted.")
        print(f"{'='*80}")


def main():
    """Main execution function - DEBUG MODE"""
    # Configuration
    PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"
    
    print("="*80)
    print("Bank Statement Parser - DEBUG MODE")
    print("="*80)
    print(f"Source Folder: {PDF_FOLDER}")
    print("="*80)
    
    # Create parser and test first few files
    parser = BankStatementParser(PDF_FOLDER)
    parser.test_first_few(count=3)
    
    print("\n\nPlease review the text extracts above.")
    print("Look for how 'Closing value' appears in your PDFs.")
    print("Share this output so we can adjust the extraction pattern.")


if __name__ == "__main__":
    main()