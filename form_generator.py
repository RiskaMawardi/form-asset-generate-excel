import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
import os
import glob
import re
import requests
from io import BytesIO
from PIL import Image as PILImage

class SimpleExcelGenerator:
    def __init__(self, csv_file, template_file, output_folder='generated_excel', insert_images=True):
        self.csv_file = csv_file
        self.template_file = template_file
        self.output_folder = output_folder
        self.insert_images = insert_images
        self.temp_image_folder = os.path.join(output_folder, 'temp_images')
        
        os.makedirs(output_folder, exist_ok=True)
        if insert_images:
            os.makedirs(self.temp_image_folder, exist_ok=True)
    
    def download_image_from_gdrive(self, url):
        """Download image from Google Drive URL"""
        try:
            # Extract file ID from Google Drive URL
            # Format: https://drive.google.com/open?id=FILE_ID
            # or: https://drive.google.com/file/d/FILE_ID/view
            file_id = None
            
            if 'id=' in url:
                file_id = url.split('id=')[1].split('&')[0]
            elif '/d/' in url:
                file_id = url.split('/d/')[1].split('/')[0]
            
            if not file_id:
                print(f"   WARNING: Could not extract file ID from URL: {url}")
                return None
            
            # Download URL for Google Drive
            download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            
            # Download image
            response = requests.get(download_url, timeout=10)
            if response.status_code == 200:
                # Open image with PIL to validate and resize
                img = PILImage.open(BytesIO(response.content))
                
                # Resize image if too large (max width 200px)
                max_width = 200
                if img.width > max_width:
                    ratio = max_width / img.width
                    new_height = int(img.height * ratio)
                    img = img.resize((max_width, new_height), PILImage.Resampling.LANCZOS)
                
                # Save to temp file
                temp_filename = f"{file_id}.png"
                temp_path = os.path.join(self.temp_image_folder, temp_filename)
                img.save(temp_path, 'PNG')
                
                return temp_path
            else:
                print(f"   WARNING: Failed to download image (status {response.status_code})")
                return None
                
        except Exception as e:
            print(f"   WARNING: Error downloading image: {str(e)}")
            return None
    
    def read_csv_responses(self):
        """Read CSV file from Google Forms"""
        try:
            df = pd.read_csv(self.csv_file, encoding='utf-8')
            
            # Handle duplicate columns by renaming
            cols = pd.Series(df.columns)
            for dup in cols[cols.duplicated()].unique():
                indices = cols[cols == dup].index.values.tolist()
                for i, idx in enumerate(indices):
                    if i != 0:
                        cols[idx] = f"{dup}_duplicate{i}"
            df.columns = cols.str.strip()
            
            print(f"✓ Successfully read {len(df)} responses from CSV")
            return df
        except Exception as e:
            print(f"✗ Error reading CSV: {str(e)}")
            return None
    
    def extract_assets_from_row(self, row):
        """Extract all assets from a single row"""
        assets = []
        columns = row.index.tolist()
        
        # Find all asset numbers (1, 2, 3, etc.)
        asset_numbers = set()
        for col in columns:
            # Look for "Asset 1", "Asset 2", etc. in column names
            # But skip "Upload Foto" columns
            if 'upload' not in col.lower() and 'foto' not in col.lower():
                match = re.search(r'Asset\s+(\d+)', col, re.IGNORECASE)
                if match:
                    asset_numbers.add(int(match.group(1)))
        
        # Process each asset number found
        for num in sorted(asset_numbers):
            no_asset = None
            jenis = None
            foto = None
            
            # Find all columns for this specific asset number
            for col in columns:
                # Check if this column is for asset number {num}
                if re.search(rf'\bAsset\s+{num}\b', col, re.IGNORECASE):
                    col_lower = col.lower()
                    
                    # Match "No. Asset 1" but NOT "Upload Foto No. asset 1"
                    if 'no' in col_lower and 'upload' not in col_lower and 'foto' not in col_lower:
                        value = row[col]
                        # Only take if not empty and not a URL
                        if pd.notna(value) and str(value).strip() and not str(value).startswith('http'):
                            no_asset = str(value).strip()
                    
                    # Match "Jenis Asset 1"
                    elif 'jenis' in col_lower:
                        value = row[col]
                        if pd.notna(value) and str(value).strip():
                            jenis = str(value).strip()
                    
                    # Match "Upload Foto No. asset 1"
                    elif 'upload' in col_lower or 'foto' in col_lower:
                        value = row[col]
                        if pd.notna(value) and str(value).strip():
                            foto = str(value).strip()
            
            # Add asset if No. Asset exists
            if no_asset:
                assets.append({
                    'no': no_asset,
                    'jenis': jenis if jenis else '',
                    'foto': foto if foto else ''
                })
        
        # Fallback: check for non-numbered columns (for single asset or duplicate columns)
        # This handles columns like "No. Asset 1_duplicate1", "Jenis Asset_duplicate1"
        if not assets:
            no_asset = None
            jenis = None
            foto = None
            
            for col in columns:
                col_lower = col.lower()
                
                # Look for any "No. Asset" column (numbered or not)
                if 'no' in col_lower and 'asset' in col_lower and 'upload' not in col_lower and 'foto' not in col_lower:
                    value = row[col]
                    if pd.notna(value) and str(value).strip() and not str(value).startswith('http'):
                        no_asset = str(value).strip()
                        break  # Take first match
            
            if no_asset:
                # Find corresponding Jenis and Foto
                for col in columns:
                    col_lower = col.lower()
                    if 'jenis' in col_lower and 'asset' in col_lower:
                        value = row[col]
                        if pd.notna(value):
                            jenis = str(value).strip()
                            break
                
                for col in columns:
                    col_lower = col.lower()
                    if ('upload' in col_lower or 'foto' in col_lower) and 'asset' in col_lower:
                        value = row[col]
                        if pd.notna(value):
                            foto = str(value).strip()
                            break
                
                assets.append({
                    'no': no_asset,
                    'jenis': jenis if jenis else '',
                    'foto': foto if foto else ''
                })
        
        return assets
    
    def fill_excel_template(self, template_path, output_path, person_info, assets):
        """Fill Excel template with data"""
        try:
            wb = load_workbook(template_path)
            ws = wb.active
            
            # Fill header (G1, G2, D27, J28)
            ws['H1'] = person_info.get('Area', '')
            ws['H2'] = person_info.get('Divisi', '')
            ws['E27'] = person_info.get('Nama', '')
            ws['K28'] = person_info.get('PIC', '')
            
            # Fill assets starting from row 9
            for idx, asset in enumerate(assets):
                row_num = 9 + idx
                ws[f'B{row_num}'] = asset['jenis']  # Jenis Inventaris
                ws[f'C{row_num}'] = asset['no']     # No. Asset
                ws[f'F{row_num}'] = person_info.get('Nama', '')  # Nama
                
                # Insert image if URL exists and insert_images is enabled
                if self.insert_images and asset['foto'] and asset['foto'].startswith('http'):
                    print(f"      -> Downloading image for asset {idx+1}...")
                    image_path = self.download_image_from_gdrive(asset['foto'])
                    
                    if image_path and os.path.exists(image_path):
                        try:
                            # Insert image in column D (foto)
                            img = XLImage(image_path)
                            
                            # Resize image to fit in cell (optional)
                            # Excel cell default height ~15, width ~64 pixels
                            img.width = 100  # pixels
                            img.height = 100  # pixels
                            
                            # Anchor to cell D{row_num}
                            ws.add_image(img, f'D{row_num}')
                            
                            # Set row height to accommodate image
                            ws.row_dimensions[row_num].height = 75  # points (~100px)
                            
                            print(f"      -> Image inserted successfully")
                        except Exception as e:
                            print(f"      -> WARNING: Could not insert image: {str(e)}")
            
            wb.save(output_path)
            return True
        except Exception as e:
            print(f"ERROR filling template: {str(e)}")
            return False
    
    def generate_excel_consolidated(self):
        """Generate one Excel per person with all their assets"""
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            print("No data to process!")
            return
        
        # Normalize basic column names
        for col in df.columns:
            if 'nama' in col.lower() and 'jenis' not in col.lower():
                df = df.rename(columns={col: 'Nama'})
            elif 'divisi' in col.lower():
                df = df.rename(columns={col: 'Divisi'})
            elif 'area' in col.lower():
                df = df.rename(columns={col: 'Area'})
            elif col.lower().strip() == 'pic':
                df = df.rename(columns={col: 'PIC'})
        
        # Group by person
        df['person_key'] = (
            df.get('Nama', 'Unknown').fillna('Unknown') + '_' +
            df.get('Divisi', 'Unknown').fillna('Unknown') + '_' +
            df.get('Area', 'Unknown').fillna('Unknown')
        )
        
        print(f"\nProcessing {len(df)} responses...\n")
        
        grouped = {}
        for _, row in df.iterrows():
            person_key = row['person_key']
            
            if person_key not in grouped:
                grouped[person_key] = {
                    'info': {
                        'Nama': row.get('Nama', 'Unknown'),
                        'Divisi': row.get('Divisi', 'Unknown'),
                        'Area': row.get('Area', 'Unknown'),
                        'PIC': row.get('PIC', 'Unknown')
                    },
                    'assets': []
                }
            
            # Extract assets from this row
            assets = self.extract_assets_from_row(row)
            grouped[person_key]['assets'].extend(assets)
        
        # Generate Excel files
        success = 0
        for idx, (person_key, data) in enumerate(grouped.items(), 1):
            info = data['info']
            assets = data['assets']
            
            # Create filename
            nama = str(info['Nama']).replace(' ', '_')
            area = str(info['Area']).replace(' ', '_').replace('(', '').replace(')', '').replace('.', '')
            divisi = str(info['Divisi']).replace(' ', '_')
            
            filename = f"{idx}_{area}_{divisi}_{nama}_{len(assets)}assets.xlsx"
            output_path = os.path.join(self.output_folder, filename)
            
            # Fill template
            if self.fill_excel_template(self.template_file, output_path, info, assets):
                print(f"OK [{idx}/{len(grouped)}] Created: {filename}")
                print(f"   -> {len(assets)} asset(s)")
                for i, asset in enumerate(assets, 1):
                    print(f"      {i}. {asset['jenis']} | {asset['no']}")
                success += 1
            else:
                print(f"ERROR [{idx}/{len(grouped)}] Failed: {filename}")
        
        print(f"\n{'='*60}")
        print(f"OK Completed! {success}/{len(grouped)} files generated")
        print(f"Files saved in: {os.path.abspath(self.output_folder)}")
        print(f"{'='*60}")
    
    def generate_excel_separate(self):
        """Generate one Excel per response"""
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            print("No data to process!")
            return
        
        # Normalize column names
        for col in df.columns:
            if 'nama' in col.lower() and 'jenis' not in col.lower():
                df = df.rename(columns={col: 'Nama'})
            elif 'divisi' in col.lower():
                df = df.rename(columns={col: 'Divisi'})
            elif 'area' in col.lower():
                df = df.rename(columns={col: 'Area'})
            elif col.lower().strip() == 'pic':
                df = df.rename(columns={col: 'PIC'})
        
        print(f"\nProcessing {len(df)} responses...\n")
        
        success = 0
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            info = {
                'Nama': row.get('Nama', 'Unknown'),
                'Divisi': row.get('Divisi', 'Unknown'),
                'Area': row.get('Area', 'Unknown'),
                'PIC': row.get('PIC', 'Unknown')
            }
            
            assets = self.extract_assets_from_row(row)
            
            # Create filename
            nama = str(info['Nama']).replace(' ', '_')
            area = str(info['Area']).replace(' ', '_').replace('(', '').replace(')', '').replace('.', '')
            
            filename = f"{idx}_{area}_{nama}_{len(assets)}assets.xlsx"
            output_path = os.path.join(self.output_folder, filename)
            
            if self.fill_excel_template(self.template_file, output_path, info, assets):
                print(f"OK [{idx}/{len(df)}] Created: {filename}")
                print(f"   -> {len(assets)} asset(s)")
                for i, asset in enumerate(assets, 1):
                    print(f"      {i}. {asset['jenis']} | {asset['no']}")
                success += 1
            else:
                print(f"ERROR [{idx}/{len(df)}] Failed: {filename}")
        
        print(f"\n{'='*60}")
        print(f"OK Completed! {success}/{len(df)} files generated")
        print(f"Files saved in: {os.path.abspath(self.output_folder)}")
        print(f"{'='*60}")


def auto_detect_csv():
    """Auto-detect CSV file"""
    csv_files = glob.glob("*.csv")
    if csv_files:
        latest = max(csv_files, key=os.path.getmtime)
        print(f"Auto-detected: {latest}")
        return latest
    return None


if __name__ == "__main__":
    print("="*60)
    print("  FORM IT ASSET - EXCEL GENERATOR")
    print("="*60)
    
    TEMPLATE_FILE = 'template_inventaris.xlsx'
    OUTPUT_FOLDER = 'generated_excel'
    
    print("\nSelect Mode:")
    print("  1. CONSOLIDATED - 1 file per person (all assets merged)")
    print("  2. SEPARATE - 1 file per response")
    
    mode = input("\nEnter mode (1 or 2) [default=1]: ").strip() or "1"
    
    CSV_FILE = auto_detect_csv()
    
    if CSV_FILE is None:
        print("\nERROR: No CSV file found!")
        print("Place CSV file in the same folder and run again.")
        input("\nPress Enter to exit...")
        exit()
    
    if not os.path.exists(TEMPLATE_FILE):
        print(f"\nERROR: Template not found: {TEMPLATE_FILE}")
        input("\nPress Enter to exit...")
        exit()
    
    generator = SimpleExcelGenerator(CSV_FILE, TEMPLATE_FILE, OUTPUT_FOLDER)
    
    # Ask if want to insert images
    print("\nInsert images from Google Drive links?")
    print("  Y - Yes, download and insert images (slower)")
    print("  N - No, skip images (faster)")
    insert_img = input("\nInsert images? (Y/N) [default=Y]: ").strip().upper() or "Y"
    
    generator.insert_images = (insert_img == "Y")
    
    print()
    if mode == "1":
        print("Running CONSOLIDATED mode...\n")
        generator.generate_excel_consolidated()
    else:
        print("Running SEPARATE mode...\n")
        generator.generate_excel_separate()
    
    print("\nDone! Press Enter to exit...")
    input()