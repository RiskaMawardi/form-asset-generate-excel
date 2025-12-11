import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os
import glob

class SimpleExcelGenerator:
    def __init__(self, csv_file, template_file, output_folder='generated_excel'):
       
        self.csv_file = csv_file
        self.template_file = template_file
        self.output_folder = output_folder
        
        # Create output folder if doesn't exist
        os.makedirs(output_folder, exist_ok=True)
    
    def read_csv_responses(self):
       
        try:
           
            df = pd.read_csv(self.csv_file, encoding='utf-8')
            
            df.columns = df.columns.str.strip()
            
            print(f"✓ Successfully read {len(df)} responses from CSV")
            print(f"  Columns found: {', '.join(df.columns.tolist())}")
            
            return df
        except Exception as e:
            print(f"✗ Error reading CSV: {str(e)}")
            return None
    
    def map_columns(self, df):
        """
        Map CSV columns to template fields
        Adjust this based on your actual form column names
        """
       
        column_mapping = {
            'Timestamp': 'Timestamp',
            'Nama': 'Nama',
            'Divisi': 'Divisi',
            'PIC': 'PIC',
            'Area': 'Area',
            'No. Asset': 'No. Asset',
            'Jenis Inventaris': 'Jenis Inventaris',
            'Upload Foto No. Asset': 'Upload Foto No. Asset'
        }
        
        detected_mapping = {}
        for key, value in column_mapping.items():
            for col in df.columns:
                if key.lower() in col.lower() or value.lower() in col.lower():
                    detected_mapping[col] = key
                    break
        
        return detected_mapping
    
    def fill_excel_template(self, template_path, output_path, data):
        try:
          
            wb = load_workbook(template_path)
            ws = wb.active
            
            # G1 = Area
            if 'Area' in data and pd.notna(data['Area']):
                ws['G1'] = data['Area']
            
            # G2 = Divisi
            if 'Divisi' in data and pd.notna(data['Divisi']):
                ws['G2'] = data['Divisi']
            
            # B9 = Jenis Inventaris
            if 'Jenis Inventaris' in data and pd.notna(data['Jenis Inventaris']):
                ws['B9'] = data['Jenis Inventaris']
            
            # C9 = No Inventaris (No. Asset)
            if 'No. Asset' in data and pd.notna(data['No. Asset']):
                ws['C9'] = data['No. Asset']
            
            # E9 = Nama
            if 'Nama' in data and pd.notna(data['Nama']):
                ws['E9'] = data['Nama']
            
            # D27 = Nama
            if 'Nama' in data and pd.notna(data['Nama']):
                ws['D27'] = data['Nama']
            
            # J28 = PIC
            if 'PIC' in data and pd.notna(data['PIC']):
                ws['J28'] = data['PIC']
            
            # Save the file
            wb.save(output_path)
            return True
            
        except Exception as e:
            print(f"✗ Error filling template: {str(e)}")
            return False
    
    def generate_excel_for_all_responses(self):
        """Generate Excel files for all CSV responses"""
        
        # Read CSV
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            print("No data to process!")
            return
        
        # Map columns
        column_mapping = self.map_columns(df)
        print(f"\n📋 Column mapping:")
        for csv_col, mapped_col in column_mapping.items():
            print(f"   {csv_col} → {mapped_col}")
        
        # Rename columns based on mapping
        df_mapped = df.rename(columns=column_mapping)
        
        print(f"\n🔄 Processing {len(df_mapped)} responses...\n")
        
        # Generate Excel for each response
        success_count = 0
        for idx, row in df_mapped.iterrows():
            try:
                # Create filename
                timestamp = str(row.get('Timestamp', '')).replace('/', '-').replace(':', '-').replace(' ', '_')
                nama = str(row.get('Nama', 'Unknown')).replace(' ', '_')
                area = str(row.get('Area', 'Unknown')).replace(' ', '_')
                
                filename = f"{idx+1}_{area}_{nama}_{timestamp}.xlsx"
                output_path = os.path.join(self.output_folder, filename)
                
                # Fill template and save
                if self.fill_excel_template(self.template_file, output_path, row):
                    print(f"✓ [{idx+1}/{len(df_mapped)}] Created: {filename}")
                    success_count += 1
                else:
                    print(f"✗ [{idx+1}/{len(df_mapped)}] Failed: {filename}")
                
            except Exception as e:
                print(f"✗ [{idx+1}/{len(df_mapped)}] Error: {str(e)}")
        
        print(f"\n{'='*60}")
        print(f"✓ Completed! Successfully generated {success_count}/{len(df_mapped)} Excel files")
        print(f"📁 Files saved in: {os.path.abspath(self.output_folder)}")
        print(f"{'='*60}")
    
    def generate_single_excel(self, row_number):
        """Generate Excel for a specific row number"""
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            return
        
        if row_number < 1 or row_number > len(df):
            print(f"✗ Error: Row number {row_number} is out of range (1-{len(df)})")
            return
        
        column_mapping = self.map_columns(df)
        df_mapped = df.rename(columns=column_mapping)
        
        row = df_mapped.iloc[row_number - 1]
        
        timestamp = str(row.get('Timestamp', '')).replace('/', '-').replace(':', '-').replace(' ', '_')
        nama = str(row.get('Nama', 'Unknown')).replace(' ', '_')
        area = str(row.get('Area', 'Unknown')).replace(' ', '_')
        
        filename = f"{row_number}_{area}_{nama}_{timestamp}.xlsx"
        output_path = os.path.join(self.output_folder, filename)
        
        if self.fill_excel_template(self.template_file, output_path, row):
            print(f"✓ Generated Excel file for row {row_number}: {filename}")
            print(f"📁 Saved to: {os.path.abspath(output_path)}")
        else:
            print(f"✗ Failed to generate Excel file for row {row_number}")


def auto_detect_csv():
    """Auto-detect CSV file in current directory"""
    csv_files = glob.glob("*.csv")
    if csv_files:
        # Get the most recently modified CSV
        latest_csv = max(csv_files, key=os.path.getmtime)
        print(f"📄 Auto-detected CSV file: {latest_csv}")
        return latest_csv
    return None


# ==================== MAIN PROGRAM ====================

if __name__ == "__main__":
    print("="*60)
    print("  FORM IT ASSET - AUTO EXCEL GENERATOR")
    print("  (No Google Cloud Setup Required!)")
    print("="*60)
    
    # Configuration
    TEMPLATE_FILE = 'template_inventaris.xlsx'
    OUTPUT_FOLDER = 'generated_excel'
    
    # Auto-detect CSV or specify manually
    CSV_FILE = auto_detect_csv()  # Auto-detect
    # CSV_FILE = 'Form Asset IT PT. Interbat (Jawaban).csv'  # Or specify manually
    
    if CSV_FILE is None:
        print("\n❌ No CSV file found!")
        print("\n📝 Instructions:")
        print("1. Go to your Google Form responses")
        print("2. Click the Google Sheets icon")
        print("3. In Sheets: File → Download → CSV (.csv)")
        print("4. Place the CSV file in the same folder as this script")
        print("5. Run this script again")
        exit()
    
    # Check if template exists
    if not os.path.exists(TEMPLATE_FILE):
        print(f"\n❌ Template file not found: {TEMPLATE_FILE}")
        print("Please place your Excel template in the same folder!")
        exit()
    
    # Initialize generator
    generator = SimpleExcelGenerator(
        csv_file=CSV_FILE,
        template_file=TEMPLATE_FILE,
        output_folder=OUTPUT_FOLDER
    )
    
    print("\n")
    
    # Generate Excel for ALL responses
    generator.generate_excel_for_all_responses()
    
    # Or generate for specific row only
    # generator.generate_single_excel(row_number=1)
    
    print("\n✅ Done! Press Enter to exit...")
    input()