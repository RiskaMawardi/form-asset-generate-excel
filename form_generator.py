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
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class SimpleExcelGenerator:
    def __init__(self, csv_file, template_file, output_folder='generated_excel', insert_images=True, generate_pdf=True):
        self.csv_file = csv_file
        self.template_file = template_file
        self.output_folder = output_folder
        self.insert_images = insert_images
        self.enable_pdf = generate_pdf  # Changed from self.generate_pdf to avoid conflict
        self.temp_image_folder = os.path.join(output_folder, 'temp_images')
        self.pdf_folder = os.path.join(output_folder, 'pdf_output')
        
        os.makedirs(output_folder, exist_ok=True)
        if insert_images:
            os.makedirs(self.temp_image_folder, exist_ok=True)
        if self.enable_pdf:
            os.makedirs(self.pdf_folder, exist_ok=True)
    
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
            
            # Check if image already downloaded
            temp_filename = f"{file_id}.png"
            temp_path = os.path.join(self.temp_image_folder, temp_filename)
            
            if os.path.exists(temp_path):
                print(f"      -> Using cached image")
                return temp_path
            
            # Headers to mimic browser request
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1'
            }
            
            # Try multiple download URLs
            download_urls = [
                f"https://drive.google.com/uc?export=download&id={file_id}",
                f"https://drive.usercontent.google.com/download?id={file_id}&export=download",
                f"https://lh3.googleusercontent.com/d/{file_id}",
                f"https://drive.google.com/thumbnail?id={file_id}&sz=w1000"
            ]
            
            # Create session with retry
            session = requests.Session()
            session.headers.update(headers)
            
            # Try each URL with retry logic
            for attempt, download_url in enumerate(download_urls, 1):
                try:
                    print(f"      -> Attempt {attempt}/{len(download_urls)}: {download_url.split('?')[0]}...")
                    
                    response = session.get(download_url, timeout=20, allow_redirects=True, verify=True)
                    
                    # Check if response is valid
                    if response.status_code == 200 and len(response.content) > 1000:
                        # Check content type
                        content_type = response.headers.get('content-type', '').lower()
                        
                        # Skip if HTML (usually error page)
                        if 'text/html' in content_type:
                            print(f"      -> Got HTML response, trying next URL...")
                            continue
                        
                        # Try to open as image
                        try:
                            img = PILImage.open(BytesIO(response.content))
                            
                            # Resize image if too large (max width 200px)
                            max_width = 200
                            if img.width > max_width:
                                ratio = max_width / img.width
                                new_height = int(img.height * ratio)
                                img = img.resize((max_width, new_height), PILImage.Resampling.LANCZOS)
                            
                            # Save to temp file
                            img.save(temp_path, 'PNG')
                            print(f"      -> ✓ Image downloaded successfully!")
                            return temp_path
                        
                        except Exception as img_error:
                            print(f"      -> Invalid image data, trying next URL...")
                            continue
                    
                    elif response.status_code == 403:
                        print(f"      -> Access denied (403), trying next URL...")
                        continue
                    elif response.status_code == 404:
                        print(f"      -> File not found (404)")
                        break  # No point trying other URLs
                    else:
                        print(f"      -> HTTP {response.status_code}, trying next URL...")
                        continue
                    
                except requests.exceptions.Timeout:
                    print(f"      -> Timeout, trying next URL...")
                    continue
                except requests.exceptions.ConnectionError as e:
                    print(f"      -> Connection error: {str(e)[:50]}...")
                    continue
                except requests.exceptions.RequestException as e:
                    print(f"      -> Request failed: {str(e)[:50]}...")
                    continue
                except Exception as e:
                    print(f"      -> Error: {str(e)[:50]}...")
                    continue
            
            print(f"   ✗ All download attempts failed for this image")
            return None
                
        except Exception as e:
            print(f"   ✗ Unexpected error: {str(e)[:100]}")
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

            # Fill year from timestamp in A1
            timestamp = person_info.get('Timestamp', '')
            if timestamp:
                try:
                    # Try to parse timestamp and extract year
                    # Format could be: "12/19/2024 10:30:00" or "2024-12-19 10:30:00"
                    if '/' in timestamp:
                        year = timestamp.split('/')[-1].split(' ')[0]  # Get year from MM/DD/YYYY
                    elif '-' in timestamp:
                        year = timestamp.split('-')[0]  # Get year from YYYY-MM-DD
                    else:
                        year = datetime.now().year
                    
                    # Get current value in A1 (e.g., "TAHUN ...")
                    current_a1 = ws['E1'].value or "Tahun ..."
                    # Replace "..." with the actual year
                    ws['E1'] = current_a1.replace('...', str(year))
                except:
                    # If parsing fails, use current year
                    current_a1 = ws['E1'].value or "Tahun ..."
                    ws['E1'] = current_a1.replace('...', str(datetime.now().year))
            
            ws['H1'] = person_info.get('Area', '')
            ws['H2'] = person_info.get('Divisi', '')
            ws['E27'] = person_info.get('Dibuat Oleh', '')
            ws['K28'] = person_info.get('PIC', '')
            
            # Fill assets starting from row 9
            for idx, asset in enumerate(assets):
                row_num = 9 + idx
                ws[f'A{row_num}'] = idx + 1  # Nomor urut (1, 2, 3, ...)
                ws[f'B{row_num}'] = asset['jenis']  # Jenis Inventaris
                ws[f'C{row_num}'] = asset['no']     # No. Asset
                ws[f'F{row_num}'] = person_info.get('Nama', '')  # Nama
                ws[f'G{row_num}'] = person_info.get('Jabatan', '')  # Jabatan
                
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
    
    def generate_pdf(self, pdf_path, person_info, assets):
        """Generate professional PDF document"""
        try:
            doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4),  # Changed to landscape
                                   rightMargin=2*cm, leftMargin=2*cm,
                                   topMargin=2*cm, bottomMargin=2*cm)
            
            story = []
            styles = getSampleStyleSheet()
            
            # Custom styles
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#1a1a1a'),
                spaceAfter=30,
                alignment=1  # Center
            )
            
            header_style = ParagraphStyle(
                'CustomHeader',
                parent=styles['Normal'],
                fontSize=11,
                textColor=colors.HexColor('#333333'),
                spaceAfter=12
            )
            
            # Title
            timestamp = person_info.get('Timestamp', '')
            year = datetime.now().year
            if timestamp:
                try:
                    if '/' in timestamp:
                        year = timestamp.split('/')[-1].split(' ')[0]
                    elif '-' in timestamp:
                        year = timestamp.split('-')[0]
                except:
                    pass
            
            title = Paragraph(f"<b>DAFTAR INVENTARIS TAHUN {year}</b>", title_style)
            story.append(title)
            story.append(Spacer(1, 0.5*cm))
            
            # Header Information
            header_data = [
                ['Area', ':', person_info.get('Area', '')],
                ['Divisi', ':', person_info.get('Divisi', '')],
                ['Dibuat Oleh', ':', person_info.get('Dibuat Oleh', '')],
                ['PIC', ':', person_info.get('PIC', '')]
            ]
            
            header_table = Table(header_data, colWidths=[4*cm, 0.5*cm, 14*cm])  # Wider for landscape
            header_table.setStyle(TableStyle([
                ('FONT', (0, 0), (-1, -1), 'Helvetica', 10),
                ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#333333')),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ]))
            
            story.append(header_table)
            story.append(Spacer(1, 1*cm))
            
            # Assets Table Header
            table_data = [['No', 'Foto', 'Jenis Inventaris', 'No. Asset', 'Nama', 'Jabatan']]
            
            # Assets Data
            for idx, asset in enumerate(assets):
                foto_cell = ''
                
                # Try to add image if available
                if self.insert_images and asset['foto'] and asset['foto'].startswith('http'):
                    image_path = self.download_image_from_gdrive(asset['foto'])
                    if image_path and os.path.exists(image_path):
                        try:
                            img = RLImage(image_path, width=2*cm, height=2*cm)
                            foto_cell = img
                        except:
                            foto_cell = 'N/A'
                    else:
                        foto_cell = 'N/A'
                else:
                    foto_cell = 'N/A'
                
                row = [
                    str(idx + 1),
                    foto_cell,
                    asset['jenis'],
                    asset['no'],
                    person_info.get('Nama', ''),
                    person_info.get('Jabatan', '')
                ]
                table_data.append(row)
            
            # Create table
            col_widths = [1*cm, 2.5*cm, 4*cm, 3*cm, 3.5*cm, 3*cm]
            assets_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            
            # Table styling
            table_style = TableStyle([
                # Header
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                
                # Body
                ('FONT', (0, 1), (-1, -1), 'Helvetica', 9),
                ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # No column
                ('ALIGN', (1, 1), (1, -1), 'CENTER'),  # Foto column
                ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
                
                # Grid
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                
                # Alternating rows
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')]),
                
                # Padding
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ])
            
            assets_table.setStyle(table_style)
            story.append(assets_table)
            
            # Build PDF
            doc.build(story)
            return True
            
        except Exception as e:
            print(f"      -> ERROR generating PDF: {str(e)}")
            return False
    
    def generate_excel_consolidated(self):
        """Generate one Excel per person with all their assets"""
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            print("No data to process!")
            return
        
        # Normalize basic column names
        for col in df.columns:
            if 'timestamp' in col.lower():
                df = df.rename(columns={col: 'Timestamp'})
            elif 'nama' in col.lower() and 'jenis' not in col.lower():
                df = df.rename(columns={col: 'Nama'})
            elif 'divisi' in col.lower():
                df = df.rename(columns={col: 'Divisi'})
            elif 'area' in col.lower():
                df = df.rename(columns={col: 'Area'})
            elif col.lower().strip() == 'pic':
                df = df.rename(columns={col: 'PIC'})
            elif 'dibuat' in col.lower() and 'oleh' in col.lower():
                df = df.rename(columns={col: 'Dibuat Oleh'})
            elif 'jabatan' in col.lower():
                df = df.rename(columns={col: 'Jabatan'})
        
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
                        'Timestamp': row.get('Timestamp', ''),
                        'Nama': row.get('Nama', 'Unknown'),
                        'Divisi': row.get('Divisi', 'Unknown'),
                        'Area': row.get('Area', 'Unknown'),
                        'PIC': row.get('PIC', 'Unknown'),
                        'Dibuat Oleh': row.get('Dibuat Oleh', ''),
                        'Jabatan': row.get('Jabatan', '')
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
                
                # Generate PDF if enabled
                if self.enable_pdf:
                    pdf_filename = filename.replace('.xlsx', '.pdf')
                    pdf_path = os.path.join(self.pdf_folder, pdf_filename)
                    print(f"   -> Generating PDF...")
                    if self.generate_pdf(pdf_path, info, assets):
                        print(f"   -> PDF created: {pdf_filename}")
                    else:
                        print(f"   -> PDF generation failed")
                
                success += 1
            else:
                print(f"ERROR [{idx}/{len(grouped)}] Failed: {filename}")
        
        print(f"\n{'='*60}")
        print(f"OK Completed! {success}/{len(grouped)} files generated")
        print(f"Excel files saved in: {os.path.abspath(self.output_folder)}")
        if self.enable_pdf:
            print(f"PDF files saved in: {os.path.abspath(self.pdf_folder)}")
        print(f"{'='*60}")
    
    def generate_excel_separate(self):
        """Generate one Excel per response"""
        df = self.read_csv_responses()
        if df is None or len(df) == 0:
            print("No data to process!")
            return
        
        # Normalize column names
        for col in df.columns:
            if 'timestamp' in col.lower():
                df = df.rename(columns={col: 'Timestamp'})
            elif 'nama' in col.lower() and 'jenis' not in col.lower():
                df = df.rename(columns={col: 'Nama'})
            elif 'divisi' in col.lower():
                df = df.rename(columns={col: 'Divisi'})
            elif 'area' in col.lower():
                df = df.rename(columns={col: 'Area'})
            elif col.lower().strip() == 'pic':
                df = df.rename(columns={col: 'PIC'})
            elif 'dibuat' in col.lower() and 'oleh' in col.lower():
                df = df.rename(columns={col: 'Dibuat Oleh'})
            elif 'jabatan' in col.lower():
                df = df.rename(columns={col: 'Jabatan'})
        
        print(f"\nProcessing {len(df)} responses...\n")
        
        success = 0
        for idx, (_, row) in enumerate(df.iterrows(), 1):
            info = {
                'Timestamp': row.get('Timestamp', ''),
                'Nama': row.get('Nama', 'Unknown'),
                'Divisi': row.get('Divisi', 'Unknown'),
                'Area': row.get('Area', 'Unknown'),
                'PIC': row.get('PIC', 'Unknown'),
                'Dibuat Oleh': row.get('Dibuat Oleh', ''),
                'Jabatan': row.get('Jabatan', '')
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
                
                # Generate PDF if enabled
                if self.enable_pdf:
                    pdf_filename = filename.replace('.xlsx', '.pdf')
                    pdf_path = os.path.join(self.pdf_folder, pdf_filename)
                    print(f"   -> Generating PDF...")
                    if self.generate_pdf(pdf_path, info, assets):
                        print(f"   -> PDF created: {pdf_filename}")
                    else:
                        print(f"   -> PDF generation failed")
                
                success += 1
            else:
                print(f"ERROR [{idx}/{len(df)}] Failed: {filename}")
        
        print(f"\n{'='*60}")
        print(f"OK Completed! {success}/{len(df)} files generated")
        print(f"Excel files saved in: {os.path.abspath(self.output_folder)}")
        if self.generate_pdf:
            print(f"PDF files saved in: {os.path.abspath(self.pdf_folder)}")
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
    
    # Ask if want to generate PDF
    print("\nGenerate PDF files?")
    print("  Y - Yes, create PDF alongside Excel (recommended)")
    print("  N - No, Excel only")
    gen_pdf = input("\nGenerate PDF? (Y/N) [default=Y]: ").strip().upper() or "Y"
    
    generator.enable_pdf = (gen_pdf == "Y")
    
    print()
    if mode == "1":
        print("Running CONSOLIDATED mode...\n")
        generator.generate_excel_consolidated()
    else:
        print("Running SEPARATE mode...\n")
        generator.generate_excel_separate()
    
    print("\nDone! Press Enter to exit...")
    input()