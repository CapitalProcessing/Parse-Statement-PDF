"""
Bank Statement Parser - Extracts account numbers and closing values
Updated for WFA/Continuity Group statements with space-less text extraction
"""

import re
from pathlib import Path
import PyPDF2
import pandas as pd
from typing import Optional

class BankStatementParser:
    def __init__(self, pdf_folder: str, output_excel: str = "Bank_Statement_Summary.xlsx"):
        """
        Initialize the parser
        
        Args:
            pdf_folder: Path to folder containing PDF statements
            output_excel: Name of output Excel file
        """
        self.pdf_folder = Path(pdf_folder)
        self.output_excel = output_excel
        self.results = []
    
    def extract_account_number_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract account number from filename
        Format: "Entity Name - XXXX-XXXX.pdf"
        Example: "Baby Goat Re BIC Enterprise Risk - 2193-4125.pdf"
        Returns: "2193-4125"
        """
        # Remove .pdf extension
        name_part = filename.replace('.pdf', '')
        
        # Split on the last ' - ' to get the account number
        if ' - ' in name_part:
            parts = name_part.rsplit(' - ', 1)
            if len(parts) > 1:
                # Clean up the account number part (remove any trailing text like " BIC")
                account = parts[1].strip().split()[0]
                return account
        
        return None
    
    def extract_page_one_text(self, pdf_path: Path) -> Optional[str]:
        """
        Extract text from the page labeled 'Page 1 of XXX'
        This is the SNAPSHOT page with the closing value
        """
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Search through pages to find "Page 1 of" or "SNAPSHOT"
                for page_num, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    # Look for "Page 1 of" pattern or "SNAPSHOT"
                    if re.search(r'Page\s*1\s*of\s*\d+', text, re.IGNORECASE) or \
                       re.search(r'SNAPSHOT', text, re.IGNORECASE):
                        return text
                
                # Fallback: if not found, try the first page
                if len(pdf_reader.pages) > 0:
                    return pdf_reader.pages[0].extract_text()
                    
        except Exception as e:
            print(f"Error reading {pdf_path.name}: {e}")
        
        return None
    
    def extract_closing_value(self, text: str) -> Optional[float]:
        """
        Extract closing value from page 1 text
        The PDF text extraction removes spaces, so "Closing value" becomes "Closingvalue"
        Pattern: Closingvalue$5,727.59$5,727.59
        """
        if not text:
            return None
        
        # Patterns that account for missing spaces in extracted text
        patterns = [
            # No spaces: Closingvalue$5,727.59
            r'Closingvalue\s*\$\s*([\d,]+\.\d{2})',
            r'CLOSINGVALUE\s*\$\s*([\d,]+\.\d{2})',
            
            # With possible spaces
            r'Closing\s*value\s*\$\s*([\d,]+\.\d{2})',
            r'CLOSING\s*VALUE\s*\$\s*([\d,]+\.\d{2})',
            
            # Alternative format
            r'Closingvalue[:\s]*\$?\s*([\d,]+\.\d{2})',
            
            # Just in case there's a colon
            r'Closingvalue:\s*\$?\s*([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value_str = match.group(1).replace(',', '').replace(' ', '')
                try:
                    value = float(value_str)
                    if value >= 0:  # Accept zero or positive values
                        return value
                except ValueError:
                    continue
        
        return None
    
    def process_pdf(self, pdf_path: Path) -> dict:
        """Process a single PDF and extract account number and closing value"""
        print(f"Processing: {pdf_path.name}")
        
        # Extract account number from filename
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        
        # Extract page 1 text
        page_text = self.extract_page_one_text(pdf_path)
        
        # Extract closing value
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
        
        pdf_files = list(self.pdf_folder.glob('*.pdf'))
        
        if not pdf_files:
            print(f"No PDF files found in '{self.pdf_folder}'")
            return
        
        print(f"\nFound {len(pdf_files)} PDF files to process")
        print("="*80 + "\n")
        
        for pdf_file in pdf_files:
            result = self.process_pdf(pdf_file)
            self.results.append(result)
        
        print("\n" + "="*80)
        self.save_to_excel()
    
    def save_to_excel(self):
        """Save results to Excel file"""
        if not self.results:
            print("No results to save!")
            return
        
        df = pd.DataFrame(self.results)
        
        # Sort by filename
        df = df.sort_values('Filename')
        
        # Calculate total (only for successful extractions)
        total_value = df[df['Closing Value'].notna()]['Closing Value'].sum()
        success_count = len(df[df['Status'] == '✓'])
        
        # Create Excel writer
        with pd.ExcelWriter(self.output_excel, engine='xlsxwriter') as writer:
            # Write main data
            df.to_excel(writer, sheet_name='Summary', index=False, startrow=2)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            # Define formats
            title_fmt = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            header_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
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
            
            # Write title
            worksheet.merge_range('A1:D1', 'Bank Statement Summary - October 2025', title_fmt)
            
            # Format column widths
            worksheet.set_column('A:A', 55)  # Filename
            worksheet.set_column('B:B', 18)  # Account Number
            worksheet.set_column('C:C', 18, money_fmt)  # Closing Value
            worksheet.set_column('D:D', 8)   # Status
            
            # Apply header format manually
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            # Add total row
            last_row = len(df) + 3
            worksheet.write(last_row, 1, 'TOTAL:', total_label_fmt)
            worksheet.write(last_row, 2, total_value, total_value_fmt)
        
        # Print summary
        print(f"\n{'='*80}")
        print(f"✓ Results saved to: {self.output_excel}")
        print(f"  Total PDFs processed: {len(self.results)}")
        print(f"  Successfully extracted: {success_count}")
        print(f"  Need review: {len(self.results) - success_count}")
        print(f"  Total Closing Value: ${total_value:,.2f}")
        print(f"{'='*80}")
        
        # List items that need review
        needs_review = [r for r in self.results if r['Status'] != '✓']
        if needs_review:
            print(f"\n⚠ The following {len(needs_review)} statements need manual review:")
            print("-" * 80)
            for item in needs_review:
                print(f"  • {item['Filename']}")
                if not item['Account Number']:
                    print(f"    - Could not extract account number")
                if item['Closing Value'] is None:
                    print(f"    - Could not extract closing value")
            print()


def main():
    """Main execution function"""
    # Configuration
    PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"
    OUTPUT_FILE = "Bank_Statement_Summary_Oct_2025.xlsx"
    
    print("="*80)
    print("Bank Statement Parser - October 2025")
    print("Continuity Group Statements")
    print("="*80)
    print(f"Source Folder: {PDF_FOLDER}")
    print(f"Output File: {OUTPUT_FILE}")
    print("="*80)
    
    # Create parser and process files
    parser = BankStatementParser(PDF_FOLDER, OUTPUT_FILE)
    parser.process_all_pdfs()
    
    # Print sample results
    if parser.results:
        print("\nSample Results (first 5):")
        print("-" * 80)
        for result in parser.results[:5]:
            print(f"File: {result['Filename']}")
            print(f"Account: {result['Account Number']}")
            if result['Closing Value'] is not None:
                print(f"Closing Value: ${result['Closing Value']:,.2f}")
            else:
                print(f"Closing Value: ⚠ NOT FOUND")
            print(f"Status: {result['Status']}\n")


if __name__ == "__main__":
    main()