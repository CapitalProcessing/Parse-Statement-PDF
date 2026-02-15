"""
BOK Financial Statement Parser
Extracts account numbers and total values from BOK Financial bank statements
Automatically identifies and processes only BOK statements from mixed PDF folders
GUI version with folder and file selection dialogs
"""

import re
from pathlib import Path
import pandas as pd
import PyPDF2
from typing import Optional, Dict, Tuple
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


def extract_beneficiary_and_account(filename: str) -> Tuple[Optional[str], Optional[str]]: 
    """
    Extract beneficiary nickname and account number from filename
    Pattern: "Name [WAREHOUSE] BEN - Account.pdf" where BEN is DAC, BIC, FOR, etc.
    WAREHOUSE can be:  WH, Warehouse, Whse, Whouse, Warehse (case insensitive)
    
    Returns:
        Tuple of (beneficiary, account_number)
    """
    name_part = filename.replace('.pdf', '')
    beneficiary = None
    account_number = None
    
    # Warehouse variations (case insensitive)
    warehouse_pattern = r'\b(WH|Warehouse|Whse|Whouse|Warehse)\b'
    
    if ' - ' in name_part:
        # Split on ' - ' to separate name part from account part
        parts = name_part.rsplit(' - ', 1)
        
        if len(parts) > 1:
            name_section = parts[0].strip()
            account_section = parts[1].strip()
            
            # Extract beneficiary (last word before ' - ', skipping warehouse indicators)
            name_words = name_section.split()
            
            # Look backwards through words to find beneficiary code
            for word in reversed(name_words):
                # Skip warehouse indicators (case insensitive)
                if re.match(warehouse_pattern, word, re.IGNORECASE):
                    continue
                
                # Check if this word is likely a beneficiary code (2-4 uppercase letters)
                if re.match(r'^[A-Z]{2,4}$', word, re.IGNORECASE):
                    beneficiary = word.upper()  # Normalize to uppercase
                    break
            
            # Extract account number from account section
            # Try BOK pattern first (has both hyphen and period)
            match = re.search(r'(\d+[-]\d+\.\d+)', account_section)
            if match: 
                account_number = match.group(1)
            else:
                # Try WFA pattern with optional period
                match = re.search(r'(\d+[-]\d+(?:\.\d+)?)', account_section)
                if match: 
                    account_number = match.group(1)
                else:
                    # Fallback: first word in account section
                    account_number = account_section.split()[0]
    
    return beneficiary, account_number


class BOKFinancialParser:
    """Parser for BOK Financial statements"""
    
    def __init__(self):
        self.name = "BOK Financial"
    
    def find_account_overview_page(self, pdf_path: Path) -> Optional[str]:
        """
        Find the Account Overview page (labeled as Page 2 of XX, physical page 4)
        """
        try: 
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    # Look for "Page 2 of" (may have no spaces:  "Page2of")
                    if re.search(r'Page\s*2\s*of\s*\d+', text) or 'Page2of' in text:
                        # Also verify it has Account Overview
                        if 'AccountOverview' in text or 'Account Overview' in text: 
                            return text
        except:
            pass
        return None
    
    def extract_total_value(self, text: str) -> Optional[float]:
        """
        Extract Total value from BOK Financial format
        Looking for:  "Total  705,122.36" (after Accrued Income in Investment Summary)
        """
        if not text:
            return None
        
        # Pattern to match "Total" that comes after "AccruedIncome"
        patterns = [
            # Look for "AccruedIncome" followed by amount, then "Total" with amount
            r'AccruedIncome\s+[\d,]+\.\d{2}\s+Total\s+([\d,]+\.\d{2})',
            
            # More general:  find "Total" followed by amount, avoiding compound words
            r'\bTotal\s+([\d,]+\.\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
            if match:
                try: 
                    value = float(match.group(1).replace(',', ''))
                    # Sanity check
                    if 0 < value < 10000000000: 
                        return value
                except: 
                    pass
        
        return None
    
    def parse(self, pdf_path: Path) -> Dict:
        """Parse a BOK Financial PDF"""
        beneficiary, account_number = extract_beneficiary_and_account(pdf_path.name)
        page_text = self.find_account_overview_page(pdf_path)
        total_value = self.extract_total_value(page_text)
        
        return {
            'Filename': pdf_path.name,
            'Beneficiary': beneficiary,
            'Account Number': account_number,
            'Closing Value': total_value,
            'Institution': 'BOK Financial',
            'Status': '✓' if (account_number and total_value is not None) else '⚠'
        }


class BOKStatementParser:
    """Parser for BOK Financial statements from mixed PDF folders"""
    
    def __init__(self, pdf_folder: str, output_excel: str = "BOK_Statement_Summary.xlsx"):
        self.pdf_folder = Path(pdf_folder)
        self.output_excel = output_excel
        self.results = []
        self.bok_parser = BOKFinancialParser()
    
    def is_bok_financial(self, pdf_path: Path) -> bool:
        """
        Check if PDF is from BOK Financial
        BOK FINANCIAL appears on physical pages 1, 3, 4+ (most pages except blank page 2)
        """
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                # Check first page and page 3 for BOK identifier
                pages_to_check = [0, 2, 3] if len(pdf_reader.pages) > 3 else range(len(pdf_reader.pages))
                
                for page_num in pages_to_check: 
                    if page_num < len(pdf_reader.pages):
                        text = pdf_reader.pages[page_num].extract_text()
                        if text and ('BOKFINANCIAL' in text or 'BOK FINANCIAL' in text or 'BOKF' in text):
                            return True
        except:
            pass
        return False
    
    def process_all_pdfs(self):
        """Process all BOK Financial PDFs in the folder"""
        if not self.pdf_folder.exists():
            print(f"Error: Folder '{self.pdf_folder}' does not exist!")
            return
        
        pdf_files = sorted(list(self.pdf_folder.glob('*.pdf')))
        
        if not pdf_files:
            print(f"No PDF files found in '{self.pdf_folder}'")
            return
        
        # Filter to only BOK statements
        print(f"\nScanning {len(pdf_files)} PDF files...")
        bok_files = [f for f in pdf_files if self.is_bok_financial(f)]
        
        print(f"Found {len(bok_files)} BOK Financial statements to process")
        print(f"Skipping {len(pdf_files) - len(bok_files)} non-BOK statements")
        print("="*80)
        print("Processing BOK statements", end=" ", flush=True)
        
        for pdf_file in bok_files:
            result = self.bok_parser.parse(pdf_file)
            self.results.append(result)
            print(".", end="", flush=True)
        
        print()
        print("="*80)
        self.save_to_excel()
    
    def save_to_excel(self):
        """Save results to Excel"""
        if not self.results:
            print("No results to save!")
            return
        
        df = pd.DataFrame(self.results)
        df = df.sort_values('Filename')
        
        # Calculate totals
        total_value = df['Closing Value'].sum()
        success_count = len(df[df['Status'] == '✓'])
        
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
            worksheet.merge_range('A1:F1', 'BOK Financial Statement Summary - October 2025', title_fmt)
            
            # Column widths
            worksheet.set_column('A:A', 55)  # Filename
            worksheet.set_column('B:B', 12)  # Beneficiary
            worksheet.set_column('C:C', 20)  # Account Number
            worksheet.set_column('D:D', 18, money_fmt)  # Closing Value
            worksheet.set_column('E:E', 25)  # Institution
            worksheet.set_column('F:F', 10)  # Status
            
            # Headers
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(2, col_num, value, header_fmt)
            
            # Total row
            last_row = len(df) + 3
            worksheet.write(last_row, 2, 'TOTAL:', total_label_fmt)
            worksheet.write(last_row, 3, total_value, total_value_fmt)
        
        # Print summary
        print(f"\n{'='*80}")
        print(f"✓ Results saved to: {self.output_excel}")
        print(f"{'='*80}")
        print(f"  Total BOK PDFs processed: {len(self.results)}")
        print(f"  Successfully extracted: {success_count} ({success_count/len(self.results)*100:.1f}%)")
        print(f"  Need review: {len(self.results) - success_count}")
        print(f"\n  TOTAL VALUE: ${total_value:,.2f}")
        print(f"{'='*80}")
        
        # Show beneficiary breakdown
        beneficiary_summary = df.groupby('Beneficiary', dropna=False)['Closing Value'].agg(['sum', 'count'])
        if len(beneficiary_summary) > 0:
            print(f"\n  By Beneficiary:")
            for beneficiary, row in beneficiary_summary.iterrows():
                ben_name = beneficiary if pd.notna(beneficiary) else "Unknown"
                print(f"    {ben_name}: {int(row['count'])} files, ${row['sum']:,.2f}")
        
        # List files needing review
        needs_review = df[df['Status'] != '✓']
        if len(needs_review) > 0:
            print(f"\n⚠ {len(needs_review)} statements need manual review:")
            print("-" * 80)
            for _, row in needs_review.iterrows():
                print(f"  • {row['Filename']}")
                issues = []
                if pd.isna(row['Beneficiary']):
                    issues.append("no beneficiary")
                if pd.isna(row['Account Number']):
                    issues.append("no account #")
                if pd.isna(row['Closing Value']):
                    issues.append("no closing value")
                if issues:
                    print(f"    ({', '.join(issues)})")
            print()
        
        # Show sample successful extractions
        if success_count > 0:
            print("\nSample extractions:")
            print("-" * 80)
            for _, row in df[df['Status'] == '✓'].head(3).iterrows():
                print(f"  {row['Filename']}")
                print(f"    Beneficiary: {row['Beneficiary']}, Account: {row['Account Number']}, Value: ${row['Closing Value']:,.2f}")
        print()


def select_source_folder():
    """Open a dialog to select the source PDF folder"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    folder_path = filedialog.askdirectory(
        title="Select Folder Containing PDF Bank Statements",
        initialdir=r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\Holding"
    )
    
    root.destroy()
    return folder_path


def select_output_file():
    """Open a dialog to select where to save the output Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Generate default filename with current date
    default_name = f"BOK_Statement_Summary_{datetime.now().strftime('%b_%Y')}.xlsx"
    
    file_path = filedialog.asksaveasfilename(
        title="Save Excel Report As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name
    )
    
    root.destroy()
    return file_path


def main():
    """Main execution with GUI dialogs"""
    print("="*80)
    print("BOK Financial Statement Parser")
    print("="*80)
    print()
    
    # Step 1: Select source folder
    print("Step 1: Select the folder containing PDF bank statements...")
    pdf_folder = select_source_folder()
    
    if not pdf_folder:
        print("No folder selected. Exiting.")
        return
    
    print(f"✓ Selected source folder: {pdf_folder}")
    print()
    
    # Verify folder has PDFs
    pdf_count = len(list(Path(pdf_folder).glob('*.pdf')))
    if pdf_count == 0:
        messagebox.showerror(
            "No PDFs Found",
            f"No PDF files found in the selected folder:\n{pdf_folder}"
        )
        print(f"Error: No PDF files found in '{pdf_folder}'")
        return
    
    print(f"Found {pdf_count} PDF files in folder")
    print()
    
    # Step 2: Select output file location
    print("Step 2: Choose where to save the Excel report...")
    output_file = select_output_file()
    
    if not output_file:
        print("No output file selected. Exiting.")
        return
    
    print(f"✓ Selected output file: {output_file}")
    print()
    
    # Step 3: Process the files
    print("="*80)
    print(f"Processing started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    try:
        parser = BOKStatementParser(pdf_folder, output_file)
        parser.process_all_pdfs()
        
        # Show success message
        if parser.results:
            messagebox.showinfo(
                "Processing Complete",
                f"Successfully processed {len(parser.results)} BOK Financial PDF files!\n\n"
                f"Report saved to:\n{output_file}\n\n"
                "Check the console for detailed results."
            )
        else:
            messagebox.showwarning(
                "No BOK PDFs Found",
                f"No BOK Financial statements found in the selected folder.\n\n"
                f"Scanned {pdf_count} PDF files total."
            )
        
    except Exception as e:
        error_msg = f"An error occurred during processing:\n{str(e)}"
        messagebox.showerror("Error", error_msg)
        print(f"\nError: {error_msg}")


if __name__ == "__main__":
    main()
