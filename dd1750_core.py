"""DD1750 core - Full inspection-ready version with admin fields."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict
from datetime import datetime

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# Try to register Arial font
try:
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    DEFAULT_FONT = 'Arial'
except:
    DEFAULT_FONT = 'Helvetica'


# Constants
ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = 612.0, 792.0

# Column positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE
PAD_X = 3.0

# Admin field positions (typical DD1750 layout)
Y_ADMIN_TOP = 720  # Top section for admin info
Y_REQUISITION = 700
Y_ORDER = 680
Y_PACKED_BY = 660
Y_DATE = 640
Y_PAGE_INFO = 100  # Bottom for page numbering
Y_TOTAL_BOXES = 115  # For "TOTAL NO. OF BOXES"


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract items from BOM PDF."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2:
                        continue
                    
                    # Find columns
                    header = table[0]
                    lv_idx = None
                    desc_idx = None
                    mat_idx = None
                    qty_idx = None
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'QTY' in text or 'AUTH' in text:
                                qty_idx = i
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    # Process rows
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        # Check LV = 'B'
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Get description (second line)
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            if len(lines) >= 2:
                                description = lines[1].strip()
                            else:
                                description = lines[0].strip()
                            
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description:
                            continue
                        
                        # Get NSN
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        # Get quantity
                        qty = 1
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                try:
                                    qty = int(str(qty_cell).strip())
                                except:
                                    qty = 1
                        
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description[:100],
                            nsn=nsn,
                            qty=qty
                        ))
    
    except Exception as e:
        print(f"ERROR: {e}")
        return []
    
    return items


def _draw_admin_fields(c, admin_data: Dict, page_num: int, total_pages: int):
    """Draw admin fields on the PDF overlay."""
    font = DEFAULT_FONT
    
    # REQUISITION NO. (typically around x=300, y=700)
    if admin_data.get('requisition_no'):
        c.setFont(font, 10)
        c.drawString(300, 700, f"REQUISITION NO.: {admin_data['requisition_no']}")
    
    # ORDER NO. (typically around x=300, y=680)
    if admin_data.get('order_no'):
        c.setFont(font, 10)
        c.drawString(300, 680, f"ORDER NO.: {admin_data['order_no']}")
    
    # PACKED BY (typically around x=44, y=660)
    if admin_data.get('packed_by'):
        c.setFont(font, 10)
        c.drawString(44, 660, f"PACKED BY: {admin_data['packed_by']}")
    
    # DATE (typically around x=300, y=660)
    if admin_data.get('date'):
        c.setFont(font, 10)
        c.drawString(300, 660, f"DATE: {admin_data['date']}")
    
    # TOTAL NO. OF BOXES (typically around x=44, y=115)
    if admin_data.get('num_boxes'):
        c.setFont(font, 10)
        c.drawString(44, 115, f"TOTAL NO. OF BOXES: {admin_data['num_boxes']}")
    
    # PAGE NUMBERING (bottom of each page)
    if total_pages > 1:
        c.setFont(font, 10)
        page_text = f"PAGE {page_num} OF {total_pages}"
        # Center the page number
        page_width = pdfmetrics.stringWidth(page_text, font, 10)
        x_pos = (PAGE_W - page_width) / 2
        c.drawString(x_pos, 100, page_text)


def generate_dd1750_from_pdf(
    bom_pdf_path: str,
    template_pdf_path: str,
    out_pdf_path: str,
    start_page: int = 0,
    admin_data: Dict = None
):
    """Generate DD1750 with admin fields."""
    
    if admin_data is None:
        admin_data = {}
    
    try:
        items = extract_items_from_pdf(bom_pdf_path, start_page)
        
        print(f"\nItems found: {len(items)}")
        for i, item in enumerate(items, 1):
            print(f"{i}. '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        
        if not items:
            # Return empty template
            reader = PdfReader(template_pdf_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_pdf_path, 'wb') as f:
                writer.write(f)
            return out_pdf_path, 0
        
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create overlay
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
            
            first_row_top = Y_TABLE_TOP_LINE - 5.0
            
            # Draw item rows
            for i, item in enumerate(page_items):
                y = first_row_top - (i * ROW_H)
                y_desc = y - 7.0
                y_nsn = y - 12.2
                
                # Box number
                can.setFont(DEFAULT_FONT, 8)
                can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
                
                # Description
                can.setFont(DEFAULT_FONT, 7)
                desc = item.description[:50] if len(item.description) > 50 else item.description
                can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
                
                # NSN
                if item.nsn:
                    can.setFont(DEFAULT_FONT, 6)
                    can.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
                
                # Quantities
                can.setFont(DEFAULT_FONT, 8)
                can.drawCentredString((X_UOI_L + X_UOI_R)/2, y_desc, "EA")
                can.drawCentredString((X_INIT_L + X_INIT_R)/2, y_desc, str(item.qty))
                can.drawCentredString((X_SPARES_L + X_SPARES_R)/2, y_desc, "0")
                can.drawCentredString((X_TOTAL_L + X_TOTAL_R)/2, y_desc, str(item.qty))
            
            # Draw admin fields on FIRST page only
            if page_num == 0:
                _draw_admin_fields(can, admin_data, page_num + 1, total_pages)
            
            # Draw page numbering on ALL pages
            _draw_page_numbering(can, page_num + 1, total_pages)
            
            can.save()
            packet.seek(0)
            
            # Merge with template
            overlay = PdfReader(packet)
            page = PdfReader(template_pdf_path).pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
        
        with open(out_pdf_path, 'wb') as f:
            writer.write(f)
        
        return out_pdf_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        try:
            reader = PdfReader(template_pdf_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(out_pdf_path, 'wb') as f:
                writer.write(f)
        except:
            pass
        return out_pdf_path, 0


def _draw_page_numbering(c, page_num: int, total_pages: int):
    """Draw page numbers at bottom."""
    font = DEFAULT_FONT
    
    if total_pages > 1:
        c.setFont(font, 10)
        page_text = f"PAGE {page_num} OF {total_pages}"
        page_width = pdfmetrics.stringWidth(page_text, font, 10)
        x_pos = (PAGE_W - page_width) / 2
        c.drawString(x_pos, 100, page_text)
