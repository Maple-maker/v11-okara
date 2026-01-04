"""DD1750 core - With proper box positioning."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = 612.0, 792.0

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


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


# Admin field positions (from DD1750 template boxes)
ADMIN_POSITIONS = {
    # Top section boxes
    'unit': {'x': 44, 'y': 745, 'width': 200},  # UNIT box
    'requisition': {'x': 300, 'y': 745, 'width': 150},  # REQUISITION NO.
    'page': {'x': 500, 'y': 745, 'width': 60},  # PAGE
    'date': {'x': 44, 'y': 695, 'width': 120},  # DATE box
    'order_no': {'x': 250, 'y': 695, 'width': 150},  # ORDER NO.
    'total_boxes': {'x': 450, 'y': 695, 'width': 100},  # TOTAL NO. OF BOXES
    # Bottom section (appears on every page)
    'packed_by': {'x': 44, 'y': 130, 'width': 200},  # PACKED BY signature line
    'received_by': {'x': 300, 'y': 130, 'width': 200},  # RECEIVED BY
}


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2:
                        continue
                    
                    header = table[0]
                    lv_idx = desc_idx = mat_idx = qty_idx = None
                    
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
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            description = re.sub(r'\s+(WTY|ARC|CIIC|UI|SCMC)$', '', description, flags=re.IGNORECASE)
                            description = re.sub(r'\s+', ' ', description).strip()
                        
                        if not description:
                            continue
                        
                        nsn = ""
                        if mat_idx is not None and mat_idx < len(row):
                            mat_cell = row[mat_idx]
                            if mat_cell:
                                match = re.search(r'\b(\d{9})\b', str(mat_cell))
                                if match:
                                    nsn = match.group(1)
                        
                        qty = 1
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                try:
                                    qty = int(str(qty_cell).strip())
                                except:
                                    qty = 1
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
    
    return items


def _truncate_to_fit(text: str, font: str, size: float, max_width: float) -> str:
    """Truncate text to fit in a specific width."""
    if not text:
        return ""
    
    while pdfmetrics.stringWidth(text, font, size) > max_width and len(text) > 3:
        text = text[:-1]
    
    return text.strip()


def _draw_admin_on_page(c, admin_data: Dict, page_num: int, total_pages: int):
    """Draw admin fields in their proper boxes on EVERY page."""
    font = "Helvetica"
    
    # UNIT (top left)
    if admin_data.get('unit'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['unit'], font, 8, ADMIN_POSITIONS['unit']['width'])
        c.drawString(ADMIN_POSITIONS['unit']['x'], ADMIN_POSITIONS['unit']['y'], text)
    
    # REQUISITION NO.
    if admin_data.get('requisition_no'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['requisition_no'], font, 8, ADMIN_POSITIONS['requisition']['width'])
        c.drawString(ADMIN_POSITIONS['requisition']['x'], ADMIN_POSITIONS['requisition']['y'], f"REQ: {text}")
    
    # PAGE (top right)
    if total_pages > 1:
        c.setFont(font, 8)
        page_text = f"{page_num}/{total_pages}"
        c.drawString(ADMIN_POSITIONS['page']['x'], ADMIN_POSITIONS['page']['y'], page_text)
    
    # DATE
    if admin_data.get('date'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['date'], font, 8, ADMIN_POSITIONS['date']['width'])
        c.drawString(ADMIN_POSITIONS['date']['x'], ADMIN_POSITIONS['date']['y'], text)
    
    # ORDER NO.
    if admin_data.get('order_no'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['order_no'], font, 8, ADMIN_POSITIONS['order_no']['width'])
        c.drawString(ADMIN_POSITIONS['order_no']['x'], ADMIN_POSITIONS['order_no']['y'], text)
    
    # TOTAL NO. OF BOXES
    if admin_data.get('num_boxes'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['num_boxes'], font, 8, ADMIN_POSITIONS['total_boxes']['width'])
        c.drawString(ADMIN_POSITIONS['total_boxes']['x'], ADMIN_POSITIONS['total_boxes']['y'], text)
    
    # PACKED BY (bottom section - on EVERY page)
    if admin_data.get('packed_by'):
        c.setFont(font, 8)
        text = _truncate_to_fit(admin_data['packed_by'], font, 8, ADMIN_POSITIONS['packed_by']['width'])
        c.drawString(ADMIN_POSITIONS['packed_by']['x'], ADMIN_POSITIONS['packed_by']['y'], f"PACKED BY: {text}")
        
        # Draw signature line
        c.setFont(font, 6)
        c.drawString(ADMIN_POSITIONS['packed_by']['x'], ADMIN_POSITIONS['packed_by']['y'] - 10, "(Signature)")
    
    # RECEIVED BY (bottom section - on EVERY page)
    c.setFont(font, 8)
    c.drawString(ADMIN_POSITIONS['received_by']['x'], ADMIN_POSITIONS['received_by']['y'], "RECEIVED BY:")
    c.setFont(font, 6)
    c.drawString(ADMIN_POSITIONS['received_by']['x'], ADMIN_POSITIONS['received_by']['y'] - 10, "(Signature)")


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0, admin_data=None):
    if admin_data is None:
        admin_data = {}
    
    items = extract_items_from_pdf(bom_path, start_page)
    print(f"Items found: {len(items)}")
    
    if not items:
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
        first_row_top = Y_TABLE_TOP_LINE - 5.0
        
        # Draw item rows
        for i, item in enumerate(page_items):
            y = first_row_top - (i * ROW_H)
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            can.setFont("Helvetica", 8)
            can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
            
            can.setFont("Helvetica", 7)
            desc = item.description[:50] if len(item.description) > 50 else item.description
            can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
            
            if item.nsn:
                can.setFont("Helvetica", 6)
                can.drawString(X_CONTENT_L + PAD_X, y_nsn, f"NSN: {item.nsn}")
            
            can.setFont("Helvetica", 8)
            can.drawCentredString((X_UOI_L + X_UOI_R)/2, y_desc, "EA")
            can.drawCentredString((X_INIT_L + X_INIT_R)/2, y_desc, str(item.qty))
            can.drawCentredString((X_SPARES_L + X_SPARES_R)/2, y_desc, "0")
            can.drawCentredString((X_TOTAL_L + X_TOTAL_R)/2, y_desc, str(item.qty))
        
        # Draw admin fields on EVERY page
        _draw_admin_on_page(can, admin_data, page_num + 1, total_pages)
        
        can.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        page = PdfReader(template_path).pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)
