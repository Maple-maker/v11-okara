"""DD1750 core - Template-aware positioning."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict, Optional

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

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


def detect_template_positions(template_path: str) -> Dict:
    """Detect actual text box positions from the template."""
    positions = {}
    
    try:
        with pdfplumber.open(template_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text() or ""
            
            # Look for text near the top (admin fields)
            words = page.extract_words()
            
            # Find positions of common labels
            for word in words:
                word_text = word['text'].upper()
                x0, y0, x1, y1 = word['x0'], word['top'], word['x1'], word['top']
                
                # Find labels and detect their positions
                if 'REQUISITION' in word_text:
                    positions['requisition'] = {'x': x1 + 5, 'y': y0}
                elif 'ORDER' in word_text and 'NO' in word_text:
                    positions['order'] = {'x': x1 + 5, 'y': y0}
                elif word_text == 'DATE':
                    positions['date'] = {'x': x1 + 5, 'y': y0}
                elif 'BOXES' in word_text:
                    positions['boxes'] = {'x': x1 + 5, 'y': y0}
                elif word_text == 'UNIT':
                    positions['unit'] = {'x': x1 + 5, 'y': y0}
                elif word_text == 'PAGE':
                    positions['page'] = {'x': x1 + 5, 'y': y0}
                elif 'PACKED' in word_text and 'BY' in word_text:
                    positions['packed_by'] = {'x': x0, 'y': y0}
                elif 'RECEIVED' in word_text and 'BY' in word_text:
                    positions['received_by'] = {'x': x0, 'y': y0}
            
            print(f"Detected positions: {positions}")
    
    except Exception as e:
        print(f"Error detecting positions: {e}")
    
    return positions


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
                    lv_idx = desc_idx = mat_idx = auth_qty_idx = None
                    
                    for i, cell in enumerate(header):
                        if cell:
                            text = str(cell).upper()
                            if 'LV' in text or 'LEVEL' in text:
                                lv_idx = i
                            elif 'DESC' in text or 'NOMENCLATURE' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text and 'QTY' in text:
                                auth_qty_idx = i
                                print(f"Found AUTH QTY column at index {i}")
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Get description - second line
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                        
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
                        
                        # Get QTY from AUTH QTY column
                        qty = 1
                        if auth_qty_idx is not None and auth_qty_idx < len(row):
                            qty_cell = row[auth_qty_idx]
                            if qty_cell:
                                qty_text = str(qty_cell).strip()
                                print(f"AUTH QTY cell: '{qty_text}'")
                                match = re.search(r'(\d+)', qty_text)
                                if match:
                                    qty = int(match.group(1))
                                    print(f"  -> Extracted qty: {qty}")
                        
                        items.append(BomItem(len(items) + 1, description.strip(), nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    return items


def generate_dd1750_from_pdf(bom_path: str, template_path: str, output_path: str, 
                            start_page: int = 0, admin_data: Dict = None) -> tuple:
    if admin_data is None:
        admin_data = {}
    
    # Detect template positions
    positions = detect_template_positions(template_path)
    
    items = extract_items_from_pdf(bom_path, start_page)
    print(f"\nItems extracted: {len(items)}")
    
    if not items:
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
        except:
            pass
        return output_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        first_row = Y_TABLE_TOP - 5.0
        
        # Draw items
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
        
        # Draw admin fields using detected positions
        c.setFont("Helvetica", 10)
        
        if positions.get('unit') and admin_data.get('unit'):
            pos = positions['unit']
            c.drawString(pos['x'], pos['y'], admin_data['unit'][:30])
        
        if positions.get('requisition') and admin_data.get('requisition_no'):
            pos = positions['requisition']
            c.drawString(pos['x'], pos['y'], f"REQ: {admin_data['requisition_no']}")
        
        if positions.get('date') and admin_data.get('date'):
            pos = positions['date']
            c.drawString(pos['x'], pos['y'], admin_data['date'])
        
        if positions.get('order') and admin_data.get('order_no'):
            pos = positions['order']
            c.drawString(pos['x'], pos['y'], f"ORDER: {admin_data['order_no']}")
        
        if positions.get('boxes') and admin_data.get('num_boxes'):
            pos = positions['boxes']
            c.drawString(pos['x'], pos['y'], admin_data['num_boxes'])
        
        if positions.get('page') and total_pages > 1:
            pos = positions['page']
            c.setFont("Helvetica", 8)
            c.drawString(pos['x'], pos['y'], f"{page_num + 1}/{total_pages}")
        
        # Bottom section
        if positions.get('packed_by') and admin_data.get('packed_by'):
            pos = positions['packed_by']
            c.setFont("Helvetica", 10)
            c.drawString(pos['x'], pos['y'], admin_data['packed_by'])
            c.setFont("Helvetica", 8)
            c.drawString(pos['x'], pos['y'] - 10, "(Signature)")
        
        if positions.get('received_by'):
            pos = positions['received_by']
            c.setFont("Helvetica", 10)
            c.drawString(pos['x'], pos['y'], "RECEIVED BY:")
            c.setFont("Helvetica", 8)
            c.drawString(pos['x'], pos['y'] - 10, "(Signature)")
        
        c.save()
        packet.seek(0)
        
        overlay = PdfReader(packet)
        page = PdfReader(template_path).pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)
