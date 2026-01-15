"""
Bank Statement Parser - WORKING VERSION
Extracts account numbers and closing values from Continuity Group statements
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
                account_part = parts[1].strip()
                # Extract account number pattern
                match = re.search(r'(\d+[-\.]\d+(?:\.\d+)?)', account_part)
                if match:
                    return match.group(1)
                return account_part.split()[0]
        
        return None
    
    def find_snapshot_page(self, pdf_path: Path) -> Optional[str]:
        """Find the page with 'Page 1 of XX' which contains the closing value"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    # Look for "Page 1 of" (with or without spaces)
                    if re.search(r'Page\s*1\s*of\s*\d+', text):
                        return text
                
        except Exception as e:
            print(f"  Error reading PDF: {e}")
        
        return None
    
    def extract_closing_value(self, text: str) -> Optional[float]:
        """
        Extract closing value from page text
        The text has no spaces, so "Closing value" appears as "Closingvalue"
        Pattern: "Closingvalue $45,156.04 $45,156.04"
        """
        if not text:
            return None
        
        # Patterns for text with spaces removed
        patterns = [
            # No spaces: "Closingvalue $45,156.04" or "Closingvalue$45,156.04"
            r'Closingvalue\s*\$\s*([\d,]+\.\d{2})',
            
            # Just in case there are some spaces
            r'Closing\s*value\s*\$\s*([\d,]+\.\d{2})',
            
            # Case variations
            r'CLOSINGVALUE\s*\$\s*([\d,]+\.\d{2})',
            r'ClosingValue\s*\$\s*([\d,]+\.\d{2})',
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
    
    def process_pdf(self, pdf_path: Path) -> dict:
        """Process a single PDF"""
        print(".", end="", flush=True)
        
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        page_text = self.find_snapshot_page(pdf_path)
        closing_value = self.extract_closing_value(page_text)
        
        status = "✓" if (account_number and closing_value is not None) else "⚠"
        
        return {
            'Filename': pdf_path.name,
            'Account Number': account_number,
            'Closing Value': closing_value,
            'Status': status
        }
    
    def process_all_pdfs(self):
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
        print("Processing", end=" ", flush=True)
        
        for pdf_file in pdf_files:
            result = self.process_pdf(pdf_file)
            self.results.append(result)
        
        print()  # New line
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
            for item in needs_review[:10]:
                print(f"  • {item['Filename']}")
                issues = []
                if not item['Account Number']:
                    issues.append("no account #")
                if item['Closing Value'] is None:
                    issues.append("no closing value")
                if issues:
                    print(f"    ({', '.join(issues)})")
            if len(needs_review) > 10:
                print(f"  ... and {len(needs_review) - 10} more")
            print()
        
        # Show sample successful extractions
        successful = [r for r in self.results if r['Status'] == '✓']
        if successful:
            print("\nSample successful extractions (first 5):")
            print("-" * 80)
            for result in successful[:5]:
                print(f"  {result['Filename']}")
                print(f"    Account: {result['Account Number']}, Closing Value: ${result['Closing Value']:,.2f}")
            print()


def main():
    """Main execution function"""
    PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"
    OUTPUT_FILE = "Bank_Statement_Summary_Oct_2025.xlsx"
    
    print("="*80)
    print("Bank Statement Parser - Continuity Group/Wells Fargo Statements")
    print("="*80)
    print(f"Source Folder: {PDF_FOLDER}")
    print(f"Output File: {OUTPUT_FILE}")
    print("="*80)
    
    parser = BankStatementParser(PDF_FOLDER, OUTPUT_FILE)
    parser.process_all_pdfs()


if __name__ == "__main__":
    main()