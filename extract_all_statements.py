"""
Combined Bank Statement Parser
Handles both Wells Fargo Advisors and BOK Financial statements
"""

import re
from pathlib import Path
import pandas as pd
import PyPDF2
from typing import Optional, Dict
from datetime import datetime


class WFAParser:
    """Parser for Wells Fargo Advisors / Continuity Group statements"""
    
    def __init__(self):
        self.name = "Wells Fargo Advisors"
    
    def can_parse(self, pdf_path: Path) -> bool:
        """Check if this PDF is a WFA statement"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                if len(pdf_reader.pages) > 0:
                    text = pdf_reader.pages[0].extract_text()
                    return 'CONTINUITY GROUP' in text or 'Wells Fargo Advisors' in text
        except:
            pass
        return False
    
    def extract_account_number_from_filename(self, filename: str) -> Optional[str]:
        """Extract account number from filename"""
        name_part = filename.replace('.pdf', '')
        
        if ' - ' in name_part:
            parts = name_part.rsplit(' - ', 1)
            if len(parts) > 1:
                account_part = parts[1].strip()
                match = re.search(r'(\d+[-\.]\d+(?:\.\d+)?)', account_part)
                if match:
                    return match.group(1)
                return account_part.split()[0]
        
        return None
    
    def find_snapshot_page(self, pdf_path: Path) -> Optional[str]:
        """Find the page with 'Page 1 of XX'"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if text and re.search(r'Page\s*1\s*of\s*\d+', text):
                        return text
        except:
            pass
        return None
    
    def extract_closing_value(self, text: str) -> Optional[float]:
        """Extract closing value from WFA format"""
        if not text:
            return None
        
        patterns = [
            r'Closingvalue\s*\$\s*([\d,]+\.\d{2})',
            r'Closing\s*value\s*\$\s*([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    value = float(match.group(1).replace(',', ''))
                    if value >= 0:
                        return value
                except:
                    pass
        
        return None
    
    def parse(self, pdf_path: Path) -> Dict:
        """Parse a WFA PDF"""
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        page_text = self.find_snapshot_page(pdf_path)
        closing_value = self.extract_closing_value(page_text)
        
        return {
            'Filename': pdf_path.name,
            'Account Number': account_number,
            'Closing Value': closing_value,
            'Institution': 'Wells Fargo Advisors',
            'Status': '✓' if (account_number and closing_value is not None) else '⚠'
        }


class BOKFinancialParser:
    """Parser for BOK Financial statements"""
    
    def __init__(self):
        self.name = "BOK Financial"
    
    def can_parse(self, pdf_path: Path) -> bool:
        """Check if this PDF is a BOK Financial statement"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                if len(pdf_reader.pages) > 0:
                    text = pdf_reader.pages[0].extract_text()
                    return 'BOK FINANCIAL' in text or 'BOK Financial' in text
        except:
            pass
        return False
    
    def extract_account_number_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract account number from filename
        BOK format: longer numbers with hyphen AND period (e.g., 1150-0007431.1)
        """
        name_part = filename.replace('.pdf', '')
        
        if ' - ' in name_part:
            parts = name_part.rsplit(' - ', 1)
            if len(parts) > 1:
                account_part = parts[1].strip()
                # Match pattern with both hyphen and period: XXXX-XXXXXXX.X
                match = re.search(r'(\d+[-]\d+\.\d+)', account_part)
                if match:
                    return match.group(1)
                # Fallback to simpler pattern
                match = re.search(r'(\d+[-\.]\d+(?:\.\d+)?)', account_part)
                if match:
                    return match.group(1)
                return account_part.split()[0]
        
        return None
    
    def find_page_2(self, pdf_path: Path) -> Optional[str]:
        """Find the page with 'Page 2 of XX'"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if text and re.search(r'Page\s*2\s*of\s*\d+', text):
                        return text
        except:
            pass
        return None
    
    def extract_total_value(self, text: str) -> Optional[float]:
        """
        Extract Total value from BOK Financial format
        Looking for the "Total" line in the Investment Summary section
        Pattern: "Total                705,122.36" or similar
        """
        if not text:
            return None
        
        # Patterns to find "Total" followed by a dollar amount
        patterns = [
            # "Total" with spaces and amount: "Total     705,122.36"
            r'Total\s+([\d,]+\.\d{2})',
            
            # "Total" with dollar sign: "Total $705,122.36"
            r'Total\s*\$\s*([\d,]+\.\d{2})',
            
            # More specific: look for it in Investment Summary context
            r'Investment\s+Summary.*?Total\s+([\d,]+\.\d{2})',
            
            # Case variations
            r'TOTAL\s+([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            # Try with DOTALL flag to match across lines
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    value = float(match.group(1).replace(',', ''))
                    # Sanity check: BOK statements should have reasonable values
                    if value >= 0 and value < 1000000000:  # Less than 1 billion
                        return value
                except:
                    pass
        
        return None
    
    def parse(self, pdf_path: Path) -> Dict:
        """Parse a BOK Financial PDF"""
        account_number = self.extract_account_number_from_filename(pdf_path.name)
        page_text = self.find_page_2(pdf_path)
        total_value = self.extract_total_value(page_text)
        
        return {
            'Filename': pdf_path.name,
            'Account Number': account_number,
            'Closing Value': total_value,
            'Institution': 'BOK Financial',
            'Status': '✓' if (account_number and total_value is not None) else '⚠'
        }


class CombinedStatementParser:
    """Main parser that combines results from multiple institutions"""
    
    def __init__(self, pdf_folder: str, output_excel: str = "Bank_Statement_Summary.xlsx"):
        self.pdf_folder = Path(pdf_folder)
        self.output_excel = output_excel
        self.results = []
        
        # Initialize parsers
        self.wfa_parser = WFAParser()
        self.bok_parser = BOKFinancialParser()
    
    def select_parser(self, pdf_path: Path):
        """Select the appropriate parser for a PDF"""
        # Try BOK parser first (more specific)
        if self.bok_parser.can_parse(pdf_path):
            return self.bok_parser
        # Try WFA parser
        elif self.wfa_parser.can_parse(pdf_path):
            return self.wfa_parser
        # Default to WFA (most common)
        else:
            return self.wfa_parser
    
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
            parser = self.select_parser(pdf_file)
            result = parser.parse(pdf_file)
            self.results.append(result)
            print(".", end="", flush=True)
        
        print()
        print("="*80)
        self.save_to_excel()
    
    def save_to_excel(self):
        """Save combined results to Excel"""
        if not self.results:
            print("No results to save!")
            return
        
        df = pd.DataFrame(self.results)
        df = df.sort_values('Filename')
        
        # Calculate totals by institution
        wfa_df = df[df['Institution'] == 'Wells Fargo Advisors']
        bok_df = df[df['Institution'] == 'BOK Financial']
        
        wfa_total = wfa_df['Closing Value'].sum()
        bok_total = bok_df['Closing Value'].sum()
        total_value = df['Closing Value'].sum()
        
        success_count = len(df[df['Status'] == '✓'])
        wfa_count = len(wfa_df)
        bok_count = len(bok_df)
        wfa_success = len(wfa_df[wfa_df['Status'] == '✓'])
        bok_success = len(bok_df[bok_df['Status'] == '✓'])
        
        with pd.ExcelWriter(self.output_excel, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            # Formats
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
            
            # Title
            worksheet.merge_range('A1:E1', 'Bank Statement Summary - October 2025', title_fmt)
            
            # Column widths
            worksheet.set_column('A:A', 55)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 18, money_fmt)
            worksheet.set_column('D:D', 25)
            worksheet.set_column('E:E', 10)
            
            # Headers
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            # Total row
            last_row = len(df) + 3
            worksheet.write(last_row, 1, 'GRAND TOTAL:', total_label_fmt)
            worksheet.write(last_row, 2, total_value, total_value_fmt)
        
        # Print summary
        print(f"\n{'='*80}")
        print(f"✓ Results saved to: {self.output_excel}")
        print(f"{'='*80}")
        print(f"  Total PDFs processed: {len(self.results)}")
        print(f"  Successfully extracted: {success_count} ({success_count/len(self.results)*100:.1f}%)")
        print(f"  Need review: {len(self.results) - success_count}")
        print(f"\n  By Institution:")
        print(f"    Wells Fargo Advisors:")
        print(f"      Files: {wfa_count} ({wfa_success} successful)")
        print(f"      Total Value: ${wfa_total:,.2f}")
        print(f"    BOK Financial:")
        print(f"      Files: {bok_count} ({bok_success} successful)")
        print(f"      Total Value: ${bok_total:,.2f}")
        print(f"\n  GRAND TOTAL: ${total_value:,.2f}")
        print(f"{'='*80}")
        
        # List files needing review
        needs_review = df[df['Status'] != '✓']
        if len(needs_review) > 0:
            print(f"\n⚠ {len(needs_review)} statements need manual review:")
            print("-" * 80)
            for _, row in needs_review.head(15).iterrows():
                print(f"  • {row['Filename']}")
                print(f"    Institution: {row['Institution']}", end="")
                issues = []
                if pd.isna(row['Account Number']):
                    issues.append("no account #")
                if pd.isna(row['Closing Value']):
                    issues.append("no closing value")
                if issues:
                    print(f" - ({', '.join(issues)})")
                else:
                    print()
            if len(needs_review) > 15:
                print(f"  ... and {len(needs_review) - 15} more")
            print()
        
        # Show sample successful extractions from each institution
        if wfa_success > 0:
            print("\nSample WFA extractions (first 3):")
            print("-" * 80)
            for _, row in wfa_df[wfa_df['Status'] == '✓'].head(3).iterrows():
                print(f"  {row['Filename']}")
                print(f"    Account: {row['Account Number']}, Value: ${row['Closing Value']:,.2f}")
        
        if bok_success > 0:
            print("\nSample BOK Financial extractions:")
            print("-" * 80)
            for _, row in bok_df[bok_df['Status'] == '✓'].head(3).iterrows():
                print(f"  {row['Filename']}")
                print(f"    Account: {row['Account Number']}, Value: ${row['Closing Value']:,.2f}")
        print()


def main():
    """Main execution"""
    PDF_FOLDER = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\10 - Oct 2025"
    OUTPUT_FILE = "Bank_Statement_Summary_Oct_2025.xlsx"
    
    print("="*80)
    print("Combined Bank Statement Parser")
    print("Wells Fargo Advisors + BOK Financial")
    print("="*80)
    print(f"Source Folder: {PDF_FOLDER}")
    print(f"Output File: {OUTPUT_FILE}")
    print(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    parser = CombinedStatementParser(PDF_FOLDER, OUTPUT_FILE)
    parser.process_all_pdfs()


if __name__ == "__main__":
    main()