"""DD1750 core - Robust BOM extraction."""

import io
import math
import re
from dataclasses import dataclass
from typing import List

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

# Column positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP = 616.0
Y_TABLE_BOTTOM = 89.5
ROWS_PER_PAGE = 18
ROW_H = (Y_TABLE_TOP - Y_TABLE_BOTTOM) / ROWS_PER_PAGE
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str) -> List[BomItem]:
    items = []
    
    try:
        print(f"DEBUG: Attempting to extract from {pdf_path}")
        
        with pdfplumber.open(pdf_path) as pdf:
            print(f"DEBUG: PDF has {len(pdf.pages)} pages")
            
            for page_num, page in enumerate(pdf.pages):
                print(f"DEBUG: Processing page {page_num}")
                tables = page.extract_tables()
                
                for table_num, table in enumerate(tables):
                    print(f"DEBUG: Page {page_num}, Table {table_num}")
                    
                    if len(table) < 2:
                        continue
                    
                    header = table[0]
                    lv_idx = None
                    desc_idx = None
                    mat_idx = None
                    auth_idx = None
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            print(f"DEBUG:   Column {i}: '{text}'")
                            
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text or 'NOMENCLATURE' in text or 'PART NO.' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_idx = i
                    
                    print(f"DEBUG: Found columns - LV:{lv_idx}, DESC:{desc_idx}, MAT:{mat_idx}, AUTH:{auth_idx}")
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    for row_num, row in enumerate(table[1:]):
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Get LV cell
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if lv_cell:
                            lv_text = str(lv_cell).strip().upper()
                            print(f"DEBUG: Row {row_num}, LV: '{lv_text}'")
                        
                            # Relaxed check - just look for 'B' or 'B9'
                            if not ('B' in lv_text or lv_text.startswith('B')):
                                continue
                        else:
                            continue
                        
                        # Get description cell
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            print(f"DEBUG: Row {row_num}, Desc lines: {lines}")
                            
                            # SMART LOGIC: Pick the best line from description cell
                            # Filter out empty or very short lines
                            description_lines = []
                            for ln in lines:
                                ln = ln.strip()
                                if len(ln) > 3:  # Minimum length check
                                    description_lines.append(ln)
                            
                            # If no good lines found, use the raw lines
                            if not description_lines:
                                description_lines = [ln for ln in lines if ln.strip()]
                            
                            # Select description
                            if description_lines:
                                # Prefer longer lines
                                description_lines.sort(key=len, reverse=True)
                                description = description_lines[0].strip()
                            else:
                                # Fallback to first line if reasonable
                                if lines:
                                    description = lines[0].strip()
                            
                            print(f"DEBUG: Row {row_num}, Selected description: '{description[:50]}...'")
                        
                            # Clean description
                            if description:
                                # Remove trailing codes like (WTY ARC, etc)
                                description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                                # Remove multiple spaces
                                description = re.sub(r'\s+', ' ', description).strip()
                                # Remove parentheses and content
                                description = description.split('(')[0].strip()
                            
                            # Double check
                            if len(description) < 3:
                                print(f"DEBUG: Row {row_num}, Description too short after cleaning: '{description}'")
                        
                        if not description:
                            print(f"DEBUG: Row {row_num}, No valid description found")
                            continue
                        
                        # Get NSN
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Get Quantity
                        qty = 1
                        if auth_idx is not None and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                try:
                                    qty_str = str(qty_cell).strip()
                                    print(f"DEBUG: Row {row_num}, Qty raw: '{qty_str}'")
                                    match = re.search(r'(\d+)', qty_str)
                                    if match:
                                        qty = int(match.group(1))
                                        print(f"DEBUG: Row {row_num}, Qty extracted: {qty}")
                                    else:
                                        print(f"DEBUG: Row {row_num}, No number in qty")
                                except:
                                    print(f"DEBUG: Row {row_num}, Failed to parse qty")
                        
                        # Add item
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description[:100],
                            nsn=nsn,
                            qty=qty
                        ))
                        print(f"DEBUG: Added item {len(items)}")
    
    except Exception as e:
        print(f"CRITICAL ERROR in extraction: {e}")
        import traceback
        traceback.print_exc()
        return []
    
    print(f"DEBUG: Total items extracted: {len(items)}")
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str):
    items = extract_items_from_pdf(bom_path)
    
    print(f"\nItems found: {len(items)}")
    
    if not items:
        # Fallback: Write template only
        print("DEBUG: No items found, writing template")
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_path, 'wb') as f:
                writer.write(f)
        except Exception as e:
            print(f"ERROR writing template: {e}")
        return out_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    template_reader = PdfReader(template_path)
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        first_row = Y_TABLE_TOP - 5.0
        
        for i, item in enumerate(page_items):
            y = first_row - (i * ROW_H)
            
            c.setFont("Helvetica", 8)
            c.drawCentredString(66, y - 7, str(item.line_no))
            
            c.setFont("Helvetica", 7)
            c.drawString(92, y - 7, item.description[:50])
            
            if item.nsn:
                c.setFont("Helvetica", 6)
                c.drawString(92, y - 12, f"NSN: {item.nsn}")
            
            c.setFont("Helvetica", 8)
            c.drawCentredString(386, y - 7, "EA")
            c.drawCentredString(431, y - 7, str(item.qty))
            c.drawCentredString(484, y - 7, "0")
            c.drawCentredString(540, y - 7, str(item.qty))
        
        c.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        
        # Get template page (use first page for all to ensure consistent background)
        if page_num < len(template_reader.pages):
            page = template_reader.pages[page_num]
        else:
            page = template_reader.pages[0]
        
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    try:
        with open(out_path, 'wb') as f:
            writer.write(f)
    except Exception as e:
        print(f"ERROR writing PDF: {e}")

    return out_path, len(items)
