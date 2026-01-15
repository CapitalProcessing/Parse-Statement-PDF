"""
Bank Statement Parser - Final Version
Extracts account numbers and closing values from Continuity Group statements
Handles multiple pages and finds the correct "Page 1 of XX" page
"""

import re
from pathlib import Path
import pandas as pd
from typing import Optional, Tuple
import sys

# Try to import both PDF libraries
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False

if not HAS_PDFPLUMBER and not HAS_PYPDF2:
    print("ERROR: Please install either pdfplumber or PyPDF2")
    print("Run: pip install pdfplumber")
    sys.exit(1)


class BankStatementParser:
    def __init__(self, pdf_folder: str, output_excel: str = "Bank_Statement_Summary.xlsx"):
        self.pdf_folder = Path(pdf_folder)
        self.output_excel = output_excel
        self.results = []
    
    def extract_account_number_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract account number from filename
        Format: "Entity Name - XXXX-XXXX.pdf" or "Entity Name - XXXX-XXXX BIC.pdf"
        """
        name_part = filename.replace('.pdf', '')
        
        if ' - ' in name_part:
            parts = name_part.rsplit(' - ', 1)
            if len(parts) > 1:
                # Get the part after the last dash
                account_part = parts[1].strip()
                # Extract just the account number (remove trailing text like "BIC")
                # Match pattern: digits-digits or digits-digits.digits
                match = re.search(r'(\d+[-\.]\d+(?:\.\d+)?)', account_part)
                if match:
                    return match.group(1)
                # Fallback: take first space-separated token
                return account_part.split()[0]
        
        return None
    
    def find_snapshot_page_pdfplumber(self, pdf_path: Path) -> Optional[str]:
        """Find and extract the SNAPSHOT page using pdfplumber"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    # Look for "Page 1 of" pattern
                    if re.search(r'Page\s+1\s+of\s+\d+', text):
                        return text
                    
                    # Also accept pages with SNAPSHOT and Closing value
                    if 'SNAPSHOT' in text and 'Closing' in text:
                        return text
        except Exception as e:
            print(f"  pdfplumber error: {e}")
        
        return None
    
    def find_snapshot_page_pypdf2(self, pdf_path: Path) -> Optional[str]:
        """Find and extract the SNAPSHOT page using PyPDF2"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page_num, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    # Look for "Page 1 of" pattern
                    if re.search(r'Page\s*1\s*of\s*\d+', text):
                        return text
                    
                    # Also accept pages with SNAPSHOT
                    if 'SNAPSHOT' in text or 'Closing' in text:
                        return text
        except Exception as e:
            print(f"  PyPDF2 error: {e}")
        
        return None
    
    def extract_page_text(self, pdf_path: Path) -> Optional[str]:
        """Extract text from the target page using best available method"""
        text = None
        
        # Try pdfplumber first (usually better)
        if HAS_PDFPLUMBER:
            text = self.find_snapshot_page_pdfplumber(pdf_path)
            if text:
                return text
        
        # Fallback to PyPDF2
        if HAS_PYPDF2:
            text = self.find_snapshot_page_pypdf2(pdf_path)
        
        return text
    
    def extract_closing_value(self, text: str) -> Optional[float]:
        """
        Extract closing value from page text
        Handles multiple formats and spacing variations
        """
        if not text:
            return None
        
        # Try multiple patterns to handle different text extraction results
        patterns = [
            # Standard format with spaces: "Closing value $108,250.83"
            r'Closing\s+value\s+\$\s*([\d,]+\.\d{2})',
            
            # No space before $: "Closing value$108,250.83"
            r'Closing\s+value\$\s*([\d,]+\.\d{2})',
            
            # Compressed text (no spaces): "Closingvalue$108,250.83"
            r'Closingvalue\s*\$\s*([\d,]+\.\d{2})',
            
            # With colon: "Closing value: $108,250.83"
            r'Closing\s+value:\s*\$?\s*([\d,]+\.\d{2})',
            
            # Case insensitive variations
            r'CLOSING\s+VALUE\s+\$\s*([\d,]+\.\d{2})',
            r'Closing\s+Value\s+\$\s*([\d,]+\.\d{2})',
            
            # Alternative: look for value pattern after "Closing value" label
            r'Closing\s+value[:\s]+\$?\s*([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                value_str = match.group(1).replace(',', '').replace(' ', '')
                try:
                    value = float(value_str)
                    if value >= 0:  # Sanity check
                        return value
                except ValueError:
                    continue
        
        return None
    
    def process_pdf(self, pdf_path: Path, verbose: bool = False) -> dict:
        """Process a single PDF and extract account number and closing value"""
        if verbose:
            print(f"Processing: {pdf_path.name}")
        else:
            print(".", end="", flush=True)
        
        # Extract account number from filename
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        
        # Extract page text
        page_text = self.extract_page_text(pdf_path)
        
        # Extract closing value
        closing_value = self.extract_closing_value(page_text)
        
        status = "✓" if (account_number and closing_value is not None) else "⚠"
        
        return {
            'Filename': pdf_path.name,
            'Account Number': account_number,
            'Closing Value': closing_value,
            'Status': status
        }
    
    def process_all_pdfs(self, verbose: bool = False):
        """Process all PDFs in the folder"""
        if not self.pdf_folder.exists():
            print(f"Error: Folder '{self.pdf_folder}' does not exist!")
            return
        
        pdf_files = sorted(list(self.pdf_folder.glob('*.pdf')))
        
        if not pdf_files:
            print(f"No PDF files found in '{self.pdf_folder}'")
            return
        
        print(f"\nFound {len(pdf_files)} PDF files to process")
        print("="*80)
        
        if not verbose:
            print("Processing", end=" ", flush=True)
        
        for pdf_file in pdf_files:
            result = self.process_pdf(pdf_file, verbose=verbose)
            self.results.append(result)
        
        if not verbose:
            print()  # New line after dots
        
        print("="*80)
        self.save_to_excel()
    
    def save_to_excel(self):
        """Save results to Excel file"""
        if not self.results:
            print("No results to save!")
            return
        
        df = pd.DataFrame(self.results)
        df = df.sort_values('Filename')
        
        total_value = df[df['Closing Value'].notna()]['Closing Value'].sum()
        success_count = len(df[df['Status'] == '✓'])
        
        with pd.ExcelWriter(self.output_excel, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            title_fmt = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center'
            })
            
            header_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center'
            })
            
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
            
            total_label_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#E2EFDA',
                'border': 1,
                'align': 'right'
            })
            
            total_value_fmt = workbook.add_format({
                'bold': True,
                'num_format': '$#,##0.00',
                'bg_color': '#E2EFDA',
                'border': 1
            })
            
            worksheet.merge_range('A1:D1', 'Bank Statement Summary - October 2025', title_fmt)
            
            worksheet.set_column('A:A', 55)
            worksheet.set_column('B:B', 18)
            worksheet.set_column('C:C', 18, money_fmt)
            worksheet.set_column('D:D', 8)
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            last_row = len(df) + 3
            worksheet.write(last_row, 1, 'TOTAL:', total_label_fmt)
            worksheet.write(last_row, 2, total_value, total_value_fmt)
        
        print(f"\n{'='*80}")
        print(f"✓ Results saved to: {self.output_excel}")
        print(f"  Total PDFs processed: {len(self.results)}")
        print(f"  Successfully extracted: {success_count} ({success_count/len(self.results)*100:.1f}%)")
        print(f"  Need review: {len(self.results) - success_count}")
        print(f"  Total Closing Value: ${total_value:,.2f}")
        print(f"{'='*80}")
        
        needs_review = [r for r in self.results if r['Status'] != '✓']
        if needs_review:
            print(f"\n⚠ The following {len(needs_review)} statements need manual review:")
            print("-" * 80)
            for item in needs_review[:15]:
                print(f"  • {item['Filename']}")
                issues = []
                if not item['Account Number']:
                    issues.append("no account #")
                if item['Closing Value'] is None:
                    issues.append("no closing value")
                if issues:
                    print(f"    ({', '.join(issues)})")
            if len(needs_review) > 15:
                print(f"  ... and {len(needs_review) - 15} more")
            print()


def main():
    """Main execution function"""
    PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"
    OUTPUT_FILE = "Bank_Statement_Summary_Oct_2025.xlsx"
    
    print("="*80)
    print("Bank Statement Parser - Continuity Group/Wells Fargo Statements")
    print("="*80)
    
    if HAS_PDFPLUMBER:
        print("Using: pdfplumber (recommended)")
    elif HAS_PYPDF2:
        print("Using: PyPDF2 (fallback)")
    
    print(f"Source Folder: {PDF_FOLDER}")
    print(f"Output File: {OUTPUT_FILE}")
    
    parser = BankStatementParser(PDF_FOLDER, OUTPUT_FILE)
    parser.process_all_pdfs(verbose=False)
    
    # Show sample of successful extractions
    successful = [r for r in parser.results if r['Status'] == '✓']
    if successful:
        print("\nSample successful extractions (first 3):")
        print("-" * 80)
        for result in successful[:3]:
            print(f"  {result['Filename']}")
            print(f"  Account: {result['Account Number']}, Value: ${result['Closing Value']:,.2f}\n")


if __name__ == "__main__":
    main()